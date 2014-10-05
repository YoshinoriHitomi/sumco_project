Attribute VB_Name = "s_XSDCA_SQL"
'分割結晶(品番) (XSDCA) ｱｸｾｽ関数


'***テーブル「XSDCA」へのデータアクセス関数***
'＊注意 ﾊﾟﾗﾒｰﾀに値をｾｯﾄする時、まず全て初期化すること

Option Explicit

'●分割結晶(品番)
Public Type typ_XSDCA
    CRYNUMCA As String * 12      ' ﾌﾞﾛｯｸID･結晶番号
    HINBCA As String * 8
    INPOSCA As Integer         ' 結晶内開始位置
    REVNUMCA As Integer
    FACTORYCA As String * 1
    OPECA As String * 1
    KCKNTCA As Integer            ' 工程連番
    SXLIDCA As String * 13
    XTALCA As String * 12
    NEKKNTCA As String * 5       ' 最終通過管理工程
    NEWKNTCA As String * 5       ' 最終通過工程
    NEWKKBCA As String * 2       ' 最終通過作業区分
    NEMACOCA As Integer          ' 最終通過処理回数
    GNKKNTCA As String * 5       ' 現在管理工程
    GNWKNTCA As String * 5       ' 現在工程
    GNWKKBCA As String * 2       ' 現在作業区分
    GNMACOCA As Integer          ' 現在処理回数
    GNDAYCA As Date              ' 現在処理日付
    GNLCA As Integer             ' 現在長さ
    GNWCA As Long                ' 現在重量
    GNMCA As Integer             ' 現在枚数
    SUMITLCA As Integer          ' SUMMIT長さ
    SUMITWCA As Long             ' SUMMIT重量
    SUMITMCA As Integer          ' SUMMIT枚数
    CHGCA As Long                ' ﾁｬｰｼﾞ量
    KAKOUBCA As String * 1       ' 加工区分
    KEIDAYCA As Date             ' 計上日付
    GNTKUBCA As String * 3       ' 棚区分
    GNTNOCA As String * 4        ' 棚番号
    XTWORKCA As String * 2       ' 製造工場
    WFWORKCA As String * 2       ' ｳｪｰﾊ製造
    LSTATBCA As String * 1       ' 最終状態区分
    RSTATBCA As String * 1       ' 流動状態区分
    LUFRCCA As String * 3        ' 格上ｺｰﾄﾞ
    LUFRBCA As String * 1        ' 格上区分
    LDFRCCA As String * 3        ' 格下ｺｰﾄﾞ
    LDFRBCA As String * 1        ' 格下区分
    HOLDCCA As String * 3        ' ﾎｰﾙﾄﾞｺｰﾄﾞ
    HOLDBCA As String * 1        ' ホールド区分
    EXKUBCA As String * 1        ' 例外区分
    HENPKCA As String * 1        ' 返品区分
    LIVKCA As String * 1         ' 生死区分
    KANKCA As String * 1         ' 完了区分
    NFCA As String * 1           ' 入庫区分
    SAKJCA As String * 1         ' 削除区分
    TDAYCA As Date               ' 登録日付
    KDAYCA As Date               ' 更新日付
    SUMITBCA As String * 1       ' SUMMIT送信フラグ
    SNDKCA As String * 1         ' 送信フラグ
    SNDDAYCA As Date             ' 送信日付
    '2003.06.11 (SPK)Y.Katabami tuika
    CUTCNTCA As String * 1       ' 新規／再切区分 '1':再切
    HINBFLGCA As String * 1      ' 代表品番フラグ '1'：代表品番
    WFHOLDFLGCA As String * 1    ' ホールド区分(WF) 09/02/13 ooba
    HOLDKTCA As String * 5
    RPCRYNUMCA As String * 12    ' 親ﾌﾞﾛｯｸID　05/09/20 ooba
    KBLKFLGCA As String * 1      ' 関連ﾌﾞﾛｯｸﾌﾗｸﾞ　06/10/31 ooba
    BLKPOST As Integer            ' ブロック内位置(XSDCA補助項目) 07/07/25 shindo
    BLKPOSB As Integer            ' ブロック内位置(XSDCA補助項目) 07/07/25 shindo
    PLANTCATCA As String         ' 向先　2007/08/15 SPK Tsutsumi
End Type

'更新用
Public Type typ_XSDCA_Update
    CRYNUMCA As String      ' ﾌﾞﾛｯｸID･結晶番号
    HINBCA As String
    INPOSCA As String         ' 結晶内開始位置
    REVNUMCA As String
    FACTORYCA As String
    OPECA As String
    KCKNTCA As String            ' 工程連番
    SXLIDCA As String
    XTALCA As String
    NEKKNTCA As String       ' 最終通過管理工程
    NEWKNTCA As String      ' 最終通過工程
    NEWKKBCA As String       ' 最終通過作業区分
    NEMACOCA As String          ' 最終通過処理回数
    GNKKNTCA As String      ' 現在管理工程
    GNWKNTCA As String        ' 現在工程
    GNWKKBCA As String        ' 現在作業区分
    GNMACOCA As String        ' 現在処理回数
    GNDAYCA As String              ' 現在処理日付
    GNLCA As String             ' 現在長さ
    GNWCA As String                ' 現在重量
    GNMCA As String             ' 現在枚数
    SUMITLCA As String          ' SUMMIT長さ
    SUMITWCA As String             ' SUMMIT重量
    SUMITMCA As String          ' SUMMIT枚数
    CHGCA As String              ' ﾁｬｰｼﾞ量
    KAKOUBCA As String        ' 加工区分
    KEIDAYCA As String             ' 計上日付
    GNTKUBCA As String        ' 棚区分
    GNTNOCA As String        ' 棚番号
    XTWORKCA As String        ' 製造工場
    WFWORKCA As String        ' ｳｪｰﾊ製造
    LSTATBCA As String       ' 最終状態区分
    RSTATBCA As String        ' 流動状態区分
    LUFRCCA As String         ' 格上ｺｰﾄﾞ
    LUFRBCA As String         ' 格上区分
    LDFRCCA As String         ' 格下ｺｰﾄﾞ
    LDFRBCA As String         ' 格下区分
    HOLDCCA As String        ' ﾎｰﾙﾄﾞｺｰﾄﾞ
    HOLDBCA As String         ' ホールド区分
    EXKUBCA As String         ' 例外区分
    HENPKCA As String         ' 返品区分
    LIVKCA As String          ' 生死区分
    KANKCA As String          ' 完了区分
    NFCA As String            ' 入庫区分
    SAKJCA As String          ' 削除区分
    TDAYCA As String               ' 登録日付
    KDAYCA As String               ' 更新日付
    SUMITBCA As String        ' SUMMIT送信フラグ
    SNDKCA As String         ' 送信フラグ
    SNDDAYCA As String             ' 送信日付
    '2003.06.11 (SPK)Y.Katabami tuika
    CUTCNTCA As String * 1       ' 新規／再切区分 '1':再切
    HINBFLGCA As String * 1      ' 代表品番フラグ '1'：代表品番
    HOLDKTCA As String * 5
    RPCRYNUMCA As String * 12    ' 親ﾌﾞﾛｯｸID　05/09/20 ooba
    KBLKFLGCA As String * 1      ' 関連ﾌﾞﾛｯｸﾌﾗｸﾞ　06/10/31 ooba
    PLANTCATCA As String         ' 向先　2007/08/15 SPK Tsutsumi
End Type

'●SELECT●

'概要      :テーブル「XSDCA」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型               ,説明
'          :records()     ,O    ,typ_XSDCA     ,抽出レコード
'          :sqlWhere      ,I    ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I    ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :

Public Function DBDRV_GetXSDCA(records() As typ_XSDCA, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN

    Dim sql As String       'SQL全体
    Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select * From XSDCA"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDCA = FUNCTION_RETURN_FAILURE
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
            '2003.06.11 (SPK)Y.katabami tuika
            If IsNull(rs.Fields("CUTCNTCA")) = False Then .CUTCNTCA = rs.Fields("CUTCNTCA")
            If IsNull(rs.Fields("HINBFLGCA")) = False Then .HINBFLGCA = rs.Fields("HINBFLGCA")
            '2005/07
            If IsNull(rs.Fields("HOLDKTCA")) = False Then .HOLDKTCA = rs.Fields("HOLDKTCA")
            If IsNull(rs.Fields("RPCRYNUMCA")) = False Then .RPCRYNUMCA = rs.Fields("RPCRYNUMCA")   '05/09/20 ooba
            If IsNull(rs.Fields("KBLKFLGCA")) = False Then .KBLKFLGCA = rs.Fields("KBLKFLGCA")      '06/10/31 ooba
            If IsNull(rs.Fields("PLANTCATCA")) = False Then .PLANTCATCA = rs.Fields("PLANTCATCA")      '07/08/22 SPK Tsutsumi Add
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDCA = FUNCTION_RETURN_SUCCESS
End Function

'●UPDATE●

'●更新項目を構造体にセットして引き渡す

'概要      :テーブル「XSDCA」を更新する ptrn1
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型               ,説明
'          :records()     ,O   ,typ_XSDCA     ,更新レコード
'          :sqlWhere      ,I   ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I   ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O   ,FUNCTION_RETURN  ,抽出の成否
'説明      :

Public Function UpdateXSDCA(records As typ_XSDCA_Update, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
On Error GoTo proc_err
    gErr.Push "s_XSDCA_SQL.bas -- Function UpdateXSDCA"

    Dim sql As String       'SQL全体
    Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim nowtime As Date
    Dim nowtime_sql As String   'サーバ時間(SQL文)

    nowtime = getSvrTime()    'サーバーの時間を取得するように変更 2003/6/4 tuku

    ''SQLを組み立てる
    sqlBase = "Select * From XSDCA"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs Is Nothing Then
        UpdateXSDCA = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''指定箇所を更新する
    recCnt = rs.RecordCount

    If recCnt = 0 Then
        UpdateXSDCA = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    'XSDCAのUPDATE前に呼び出し
    #If Y3_CREATE = 1 Then
        ''XODY3作成  add 2009/01/08 SETmiyatake
        Call CreateOrUpdateXODY3(records.CRYNUMCA, records.SXLIDCA, records.LIVKCA, sqlWhere)
    #End If

'>>>>> .EditをSQL(UPDATE)文に変更　2009/06/29 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"

    With records
        
        ''SQLを組み立てる
        sql = "UPDATE XSDCA SET" & vbLf
        
        ''更新日付
        sql = sql & " KDAYCA = " & nowtime_sql & vbLf
        
        ''ブロックID・結晶番号
        If .CRYNUMCA <> "" And left(.CRYNUMCA, 1) <> vbNullChar Then
            sql = sql & ",CRYNUMCA = '" & .CRYNUMCA & "'" & vbLf
        End If
        
        ''品番
        If .HINBCA <> "" And left(.HINBCA, 1) <> vbNullChar Then
            sql = sql & ",HINBCA = '" & .HINBCA & "'" & vbLf
        End If
        
        ''結晶内開始位置
        If .INPOSCA <> "" Then
            sql = sql & ",INPOSCA = '" & CStr(CInt(.INPOSCA)) & "'" & vbLf
        End If
        
        ''製品番号改訂番号
        If .REVNUMCA <> "" Then
            sql = sql & ",REVNUMCA = '" & CStr(CInt(.REVNUMCA)) & "'" & vbLf
        End If
        
        ''工場
        If .FACTORYCA <> "" And left(.FACTORYCA, 1) <> vbNullChar Then
            sql = sql & ",FACTORYCA = '" & .FACTORYCA & "'" & vbLf
        End If
        
        ''操業条件
        If .OPECA <> "" And left(.OPECA, 1) <> vbNullChar Then
            sql = sql & ",OPECA = '" & .OPECA & "'" & vbLf
        End If
        
        ''工程連番
        If .KCKNTCA <> "" Then
            sql = sql & ",KCKNTCA = '" & CStr(CInt(.KCKNTCA)) & "'" & vbLf
        End If
        
        ''SXLID
        If .SXLIDCA <> "" And left(.SXLIDCA, 1) <> vbNullChar Then
            sql = sql & ",SXLIDCA = '" & .SXLIDCA & "'" & vbLf
        End If
        
        ''結晶番号
        If .XTALCA <> "" And left(.XTALCA, 1) <> vbNullChar Then
            sql = sql & ",XTALCA = '" & .XTALCA & "'" & vbLf
        End If
        
        ''最終通過管理工程
        If .NEKKNTCA <> "" And left(.NEKKNTCA, 1) <> vbNullChar Then
            sql = sql & ",NEKKNTCA = '" & .NEKKNTCA & "'" & vbLf
        End If
        
        ''最終通過工程
        If .NEWKNTCA <> "" And left(.NEWKNTCA, 1) <> vbNullChar Then
            sql = sql & ",NEWKNTCA = '" & .NEWKNTCA & "'" & vbLf
        End If
        
        ''最終通過作業区分
        If .NEWKKBCA <> "" And left(.NEWKKBCA, 1) <> vbNullChar Then
            sql = sql & ",NEWKKBCA = '" & .NEWKKBCA & "'" & vbLf
        End If
        
        ''最終通過処理回数
        If .NEMACOCA <> "" Then
            sql = sql & ",NEMACOCA = '" & CStr(CInt(.NEMACOCA)) & "'" & vbLf
        End If
        
        ''現在管理工程
        If .GNKKNTCA <> "" And left(.GNKKNTCA, 1) <> vbNullChar Then
            sql = sql & ",GNKKNTCA = '" & .GNKKNTCA & "'" & vbLf
        End If
        
        ''現在工程
        If .GNWKNTCA <> "" And left(.GNWKNTCA, 1) <> vbNullChar Then
            sql = sql & ",GNWKNTCA = '" & .GNWKNTCA & "'" & vbLf
        End If
        
        ''現在作業区分
        If .GNWKKBCA <> "" And left(.GNWKKBCA, 1) <> vbNullChar Then
            sql = sql & ",GNWKKBCA = '" & .GNWKKBCA & "'" & vbLf
        End If
        
        ''現在処理回数
        If .GNMACOCA <> "" Then
            sql = sql & ",GNMACOCA = '" & CStr(CInt(.GNMACOCA)) & "'" & vbLf
        End If
        
        ''現在処理日付
        If .GNDAYCA <> "" Then
            sql = sql & ",GNDAYCA = TO_DATE('" & Format$(CDate(.GNDAYCA), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''現在長さ
        If .GNLCA <> "" Then
            sql = sql & ",GNLCA = '" & CStr(CInt(.GNLCA)) & "'" & vbLf
        End If
        
        ''現在重量
        If .GNWCA <> "" Then
            sql = sql & ",GNWCA = '" & CStr(CLng(.GNWCA)) & "'" & vbLf
        End If
        
        ''現在枚数
        If .GNMCA <> "" Then
            sql = sql & ",GNMCA = '" & CStr(CInt(.GNMCA)) & "'" & vbLf
        End If
        
        ''SUMIT長さ
        If .SUMITLCA <> "" Then
            sql = sql & ",SUMITLCA = '" & CStr(CInt(.SUMITLCA)) & "'" & vbLf
        End If
        
        ''SUMIT重量
        If .SUMITWCA <> "" Then
            sql = sql & ",SUMITWCA = '" & CStr(CLng(.SUMITWCA)) & "'" & vbLf
        End If
        
        ''SUMIT枚数
        If .SUMITMCA <> "" Then
            sql = sql & ",SUMITMCA = '" & CStr(CInt(.SUMITMCA)) & "'" & vbLf
        End If
        
        ''チャージ量
        If .CHGCA <> "" Then
            sql = sql & ",CHGCA = '" & CStr(CLng(.CHGCA)) & "'" & vbLf
        End If
        
        ''加工区分
        If .KAKOUBCA <> "" And left(.KAKOUBCA, 1) <> vbNullChar Then
            sql = sql & ",KAKOUBCA = '" & .KAKOUBCA & "'" & vbLf
        End If
        
        ''計上日付
        If .KEIDAYCA <> "" Then
            sql = sql & ",KEIDAYCA = TO_DATE('" & Format$(CDate(.KEIDAYCA), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''棚区分
        If .GNTKUBCA <> "" And left(.GNTKUBCA, 1) <> vbNullChar Then
            sql = sql & ",GNTKUBCA = '" & .GNTKUBCA & "'" & vbLf
        End If
        
        ''棚番号
        If .GNTNOCA <> "" And left(.GNTNOCA, 1) <> vbNullChar Then
            sql = sql & ",GNTNOCA = '" & .GNTNOCA & "'" & vbLf
        End If
        
        ''製造工場
        If .XTWORKCA <> "" And left(.XTWORKCA, 1) <> vbNullChar Then
            sql = sql & ",XTWORKCA = '" & .XTWORKCA & "'" & vbLf
        End If
        
        ''ウェーハ製造
        If .WFWORKCA <> "" And left(.WFWORKCA, 1) <> vbNullChar Then
            sql = sql & ",WFWORKCA = '" & .WFWORKCA & "'" & vbLf
        End If
        
        ''最終状態区分
        If .LSTATBCA <> "" And left(.LSTATBCA, 1) <> vbNullChar Then
            sql = sql & ",LSTATBCA = '" & .LSTATBCA & "'" & vbLf
        End If
        
        ''流動状態区分
        If .RSTATBCA <> "" And left(.RSTATBCA, 1) <> vbNullChar Then
            sql = sql & ",RSTATBCA = '" & .RSTATBCA & "'" & vbLf
        End If
        
        ''格上コード
        If .LUFRCCA <> "" And left(.LUFRCCA, 1) <> vbNullChar Then
            sql = sql & ",LUFRCCA = '" & .LUFRCCA & "'" & vbLf
        End If
        
        ''格上区分
        If .LUFRBCA <> "" And left(.LUFRBCA, 1) <> vbNullChar Then
            sql = sql & ",LUFRBCA = '" & .LUFRBCA & "'" & vbLf
        End If
        
        ''格下コード
        If .LDFRCCA <> "" And left(.LDFRCCA, 1) <> vbNullChar Then
            sql = sql & ",LDFRCCA = '" & .LDFRCCA & "'" & vbLf
        End If
        
        ''格下区分
        If .LDFRBCA <> "" And left(.LDFRBCA, 1) <> vbNullChar Then
            sql = sql & ",LDFRBCA = '" & .LDFRBCA & "'" & vbLf
        End If
        
        ''ホールドコード
        If .HOLDCCA <> "" And left(.HOLDCCA, 1) <> vbNullChar Then
            sql = sql & ",HOLDCCA = '" & .HOLDCCA & "'" & vbLf
        End If
        
        ''ホールド区分
        If .HOLDBCA <> "" And left(.HOLDBCA, 1) <> vbNullChar Then
            sql = sql & ",HOLDBCA = '" & .HOLDBCA & "'" & vbLf
        End If
        
        ''例外区分
        If .EXKUBCA <> "" And left(.EXKUBCA, 1) <> vbNullChar Then
            sql = sql & ",EXKUBCA = '" & .EXKUBCA & "'" & vbLf
        End If
        
        ''返品区分
        If .HENPKCA <> "" And left(.HENPKCA, 1) <> vbNullChar Then
            sql = sql & ",HENPKCA = '" & .HENPKCA & "'" & vbLf
        End If
        
        ''生死区分
        If .LIVKCA <> "" And left(.LIVKCA, 1) <> vbNullChar Then
            sql = sql & ",LIVKCA = '" & .LIVKCA & "'" & vbLf
        End If
        
        ''完了区分
        If .KANKCA <> "" And left(.KANKCA, 1) <> vbNullChar Then
            sql = sql & ",KANKCA = '" & .KANKCA & "'" & vbLf
        End If
        
        ''入庫区分
        If .NFCA <> "" And left(.NFCA, 1) <> vbNullChar Then
            sql = sql & ",NFCA = '" & .NFCA & "'" & vbLf
        End If
        
        ''削除区分
        If .SAKJCA <> "" And left(.SAKJCA, 1) <> vbNullChar Then
            sql = sql & ",SAKJCA = '" & .SAKJCA & "'" & vbLf
        End If
        
        ''登録日付
        If .TDAYCA <> "" Then
            sql = sql & ",TDAYCA = TO_DATE('" & Format$(CDate(.TDAYCA), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''SUMIT送信フラグ
        If .SUMITBCA <> "" And left(.SUMITBCA, 1) <> vbNullChar Then
            sql = sql & ",SUMITBCA = '" & .SUMITBCA & "'" & vbLf
        End If
        
        ''送信フラグ
        If .SNDKCA <> "" And left(.SNDKCA, 1) <> vbNullChar Then
            sql = sql & ",SNDKCA = '" & .SNDKCA & "'" & vbLf
        End If
        
        ''送信日付
        If .SNDDAYCA <> "" Then
            sql = sql & ",SNDDAYCA = TO_DATE('" & Format$(CDate(.SNDDAYCA), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''切断処理区分
        If .CUTCNTCA <> "" And left(.CUTCNTCA, 1) <> vbNullChar Then
            sql = sql & ",CUTCNTCA = '" & .CUTCNTCA & "'" & vbLf
        End If
        
        ''代表品番フラグ
        If .HINBFLGCA <> "" And left(.HINBFLGCA, 1) <> vbNullChar Then
            sql = sql & ",HINBFLGCA = '" & .HINBFLGCA & "'" & vbLf
        End If
        
        ''ホールド工程
        If .HOLDKTCA <> "" And left(.HOLDKTCA, 1) <> vbNullChar Then
            sql = sql & ",HOLDKTCA = '" & .HOLDKTCA & "'" & vbLf
        End If
        
        ''親ブロックID
        If .RPCRYNUMCA <> "" And left(.RPCRYNUMCA, 1) <> vbNullChar Then
            sql = sql & ",RPCRYNUMCA = '" & .RPCRYNUMCA & "'" & vbLf
        End If
        
        ''関連ブロックフラグ
        If .KBLKFLGCA <> "" And left(.KBLKFLGCA, 1) <> vbNullChar Then
            sql = sql & ",KBLKFLGCA = '" & .KBLKFLGCA & "'" & vbLf
        End If
        
        ''向先
        If .PLANTCATCA <> "" And left(.PLANTCATCA, 1) <> vbNullChar Then
            sql = sql & ",PLANTCATCA = '" & .PLANTCATCA & "'" & vbLf
        End If
        
        sql = sql & " " & sqlWhere & vbLf
    
        'SQLを実行
        recCnt = OraDB.ExecuteSQL(sql)
        
        '返り値が1以外はエラー
        If recCnt < 0 Then
            GoTo proc_err
        ElseIf recCnt = 0 Then
            '0件更新…エラー(既存通り)
            UpdateXSDCA = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf recCnt > 1 Then
            '複数件更新…エラー(既存は複数SELECTした最初の一件のみ更新)
            UpdateXSDCA = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        
    End With
'<<<<< .EditをSQL(UPDATE)文に変更　2009/06/29 SETsw kubota ------------------

    UpdateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    UpdateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'●INSERT●

'概要      :テーブル「XSDCA」にレコードを挿入する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pXSDCA 　　  ,I  ,typ_XSDCA_Update   ,XSDCA更新用ﾃﾞｰﾀ
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function CreateXSDCA(pXSDCA As typ_XSDCA_Update, sErrMsg As String) As FUNCTION_RETURN


    Dim sql As String
    Dim sDBName As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim rs2 As OraDynaset    'RecordSet
'    Dim recCnt As Long      'レコード数
    Dim nowtime As Date
    Dim nowtime_sql As String   'サーバ時間(SQL文)

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDCA_SQL.bas -- Function CreateXSDCA"
    sErrMsg = ""
    sDBName = "XSDCA"
    'nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    nowtime = getSvrTime()    'サーバーの時間を取得するように変更 2003/6/4 tuku

'>>>>> .AddNewをSQL(INSERT)文に変更　2009/06/29 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With pXSDCA
        
        sql = "INSERT INTO XSDCA ("
        sql = sql & " CRYNUMCA"         ' 1:ﾌﾞﾛｯｸID・結晶番号
        sql = sql & ",HINBCA"           ' 2:品番
        sql = sql & ",INPOSCA"          ' 3:結晶内開始位置
        sql = sql & ",REVNUMCA"         ' 4:製品番号改訂番号
        sql = sql & ",FACTORYCA"        ' 5:工場
        sql = sql & ",OPECA"            ' 6:操業条件
        sql = sql & ",KCKNTCA"          ' 7:工程連番
        sql = sql & ",SXLIDCA"          ' 8:SXLID
        sql = sql & ",XTALCA"           ' 9:結晶番号
        sql = sql & ",NEKKNTCA"         '10:最終通過管理工程
        sql = sql & ",NEWKNTCA"         '11:最終通過工程
        sql = sql & ",NEWKKBCA"         '12:最終通過作業区分
        sql = sql & ",NEMACOCA"         '13:最終通過処理回数
        sql = sql & ",GNKKNTCA"         '14:現在管理工程
        sql = sql & ",GNWKNTCA"         '15:現在工程
        sql = sql & ",GNWKKBCA"         '16:現在作業区分
        sql = sql & ",GNMACOCA"         '17:現在処理回数
        sql = sql & ",GNDAYCA"          '18:現在処理日付
        sql = sql & ",GNLCA"            '19:現在長さ
        sql = sql & ",GNWCA"            '20:現在重量
        sql = sql & ",GNMCA"            '21:現在枚数
        sql = sql & ",SUMITLCA"         '22:SUMMIT長さ
        sql = sql & ",SUMITWCA"         '23:SUMMIT重量
        sql = sql & ",SUMITMCA"         '24:SUMMIT枚数
        sql = sql & ",CHGCA"            '25:ﾁｬｰｼﾞ量
        sql = sql & ",KAKOUBCA"         '26:加工区分
        If .KEIDAYCA <> "" Then
            sql = sql & ",KEIDAYCA"         '27:計上日付
        End If
        sql = sql & ",GNTKUBCA"         '28:棚区分
        sql = sql & ",GNTNOCA"          '29:棚番号
        sql = sql & ",XTWORKCA"         '30:製造工場
        sql = sql & ",WFWORKCA"         '31:ｳｪｰﾊ製造
        sql = sql & ",LSTATBCA"         '32:最終状態区分
        sql = sql & ",RSTATBCA"         '33:流動状態区分
        sql = sql & ",LUFRCCA"          '34:格上ｺｰﾄﾞ
        sql = sql & ",LUFRBCA"          '35:格上区分
        sql = sql & ",LDFRCCA"          '36:格下ｺｰﾄﾞ
        sql = sql & ",LDFRBCA"          '37:格下区分
        sql = sql & ",HOLDCCA"          '38:ﾎｰﾙﾄﾞｺｰﾄﾞ
        sql = sql & ",HOLDBCA"          '39:ﾎｰﾙﾄﾞ区分
        sql = sql & ",EXKUBCA"          '40:例外区分
        sql = sql & ",HENPKCA"          '41:返品区分
        sql = sql & ",LIVKCA"           '42:生死区分
        sql = sql & ",KANKCA"           '43:完了区分
        sql = sql & ",NFCA"             '44:入庫区分
        sql = sql & ",SAKJCA"           '45:削除区分
        sql = sql & ",TDAYCA"           '46:登録日付
        sql = sql & ",KDAYCA"           '47:更新日付
        sql = sql & ",SUMITBCA"         '48:SUMMIT送信ﾌﾗｸﾞ
        sql = sql & ",SNDKCA"           '49:送信ﾌﾗｸﾞ
        sql = sql & ",SNDDAYCA"         '50:送信日付
        sql = sql & ",CUTCNTCA"         '51:新規／再切区分
        sql = sql & ",HINBFLGCA"        '52:代表品番フラグ
        sql = sql & ",HOLDKTCA"         '53:ﾎｰﾙﾄﾞ工程
        sql = sql & ",RPCRYNUMCA"       '54:親ﾌﾞﾛｯｸID
        sql = sql & ",KBLKFLGCA"        '55:関連ﾌﾞﾛｯｸﾌﾗｸﾞ
        sql = sql & ",PLANTCATCA"       '56:向先
        sql = sql & ")"
        sql = sql & "VALUES (" & vbLf

        ' 1:ﾌﾞﾛｯｸID・結晶番号
        If .CRYNUMCA <> "" And left(.CRYNUMCA, 1) <> vbNullChar Then
            sql = sql & " '" & .CRYNUMCA & "'" & vbLf
        Else
            sql = sql & " '" & Space(12) & "'" & vbLf
        End If

        ' 2:品番
        If .HINBCA <> "" And left(.HINBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HINBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        ' 3:結晶内開始位置
        If .INPOSCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.INPOSCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 4:製品番号改訂番号
        If .REVNUMCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.REVNUMCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 5:工場
        If .FACTORYCA <> "" And left(.FACTORYCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .FACTORYCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 6:操業条件
        If .OPECA <> "" And left(.OPECA, 1) <> vbNullChar Then
            sql = sql & ",'" & .OPECA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 7:工程連番
        If .KCKNTCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.KCKNTCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 8:SXLID
        If .SXLIDCA <> "" And left(.SXLIDCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .SXLIDCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(13) & "'" & vbLf
        End If

        ' 9:結晶番号
        If .XTALCA <> "" And left(.XTALCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .XTALCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '10:最終通過管理工程
        If .NEKKNTCA <> "" And left(.NEKKNTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEKKNTCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '11:最終通過工程
        If .NEWKNTCA <> "" And left(.NEWKNTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEWKNTCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '12:最終通過作業区分
        If .NEWKKBCA <> "" And left(.NEWKKBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEWKKBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '13:最終通過処理回数
        If .NEMACOCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.NEMACOCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '14:現在管理工程
        If .GNKKNTCA <> "" And left(.GNKKNTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNKKNTCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '15:現在工程
        If .GNWKNTCA <> "" And left(.GNWKNTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNWKNTCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '16:現在作業区分
        If .GNWKKBCA <> "" And left(.GNWKKBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNWKKBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '17:現在処理回数
        If .GNMACOCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNMACOCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '18:現在処理日付
        sql = sql & "," & nowtime_sql & vbLf

        '19:現在長さ
        If .GNLCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNLCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '20:現在重量
        If .GNWCA <> "" Then
            sql = sql & ",'" & CStr(CLng(.GNWCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '21:現在枚数
        If .GNMCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNMCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '22:SUMMIT長さ
        If .SUMITLCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.SUMITLCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '23:SUMMIT重量
        If .SUMITWCA <> "" Then
            sql = sql & ",'" & CStr(CLng(.SUMITWCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '24:SUMMIT枚数
        If .SUMITMCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.SUMITMCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '25:ﾁｬｰｼﾞ量
        If .CHGCA <> "" Then
            sql = sql & ",'" & CStr(CLng(.CHGCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '26:加工区分
        If .KAKOUBCA <> "" And left(.KAKOUBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .KAKOUBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '27:計上日付
        If .KEIDAYCA <> "" Then
            sql = sql & ",TO_DATE('" & Format$(CDate(.KEIDAYCA), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If

        '28:棚区分
        If .GNTKUBCA <> "" And left(.GNTKUBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNTKUBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '29:棚番号
        If .GNTNOCA <> "" And left(.GNTNOCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNTNOCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(4) & "'" & vbLf
        End If

        '30:製造工場
        sql = sql & ",'" & FACTORYCD & "'" & vbLf

        '31:ｳｪｰﾊ製造
        If .WFWORKCA <> "" And left(.WFWORKCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .WFWORKCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '32:最終状態区分
        If .LSTATBCA <> "" And left(.LSTATBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LSTATBCA & "'" & vbLf
        Else
            sql = sql & ",'T'" & vbLf
        End If

        '33:流動状態区分
        If .RSTATBCA <> "" And left(.RSTATBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .RSTATBCA & "'" & vbLf
        Else
            sql = sql & ",'T'" & vbLf
        End If

        '34:格上ｺｰﾄﾞ
        If .LUFRCCA <> "" And left(.LUFRCCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LUFRCCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '35:格上区分
        If .LUFRBCA <> "" And left(.LUFRBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LUFRBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '36:格下ｺｰﾄﾞ
        If .LDFRCCA <> "" And left(.LDFRCCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LDFRCCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '37:格下区分
        If .LDFRBCA <> "" And left(.LDFRBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LDFRBCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '38:ﾎｰﾙﾄﾞｺｰﾄﾞ
        If .HOLDCCA <> "" And left(.HOLDCCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDCCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '39:ﾎｰﾙﾄﾞ区分
        If .HOLDBCA <> "" And left(.HOLDBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDBCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '40:例外区分
        If .EXKUBCA <> "" And left(.EXKUBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .EXKUBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '41:返品区分
        If .HENPKCA <> "" And left(.HENPKCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HENPKCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '42:生死区分
        If .LIVKCA <> "" And left(.LIVKCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LIVKCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '43:完了区分
        If .KANKCA <> "" And left(.KANKCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .KANKCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '44:入庫区分
        If .NFCA <> "" And left(.NFCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .NFCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '45:削除区分
        If .SAKJCA <> "" And left(.SAKJCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .SAKJCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '46:登録日付
        sql = sql & "," & nowtime_sql & vbLf

        '47:更新日付
        sql = sql & "," & nowtime_sql & vbLf

        '48:SUMMIT送信ﾌﾗｸﾞ
        sql = sql & ",'0'" & vbLf

        '49:送信ﾌﾗｸﾞ
        sql = sql & ",'0'" & vbLf

        '50:送信日付
        sql = sql & ",NULL" & vbLf

        '51:新規／再切区分
        If .CUTCNTCA <> "" And left(.CUTCNTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .CUTCNTCA & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '52:代表品番フラグ
        If .HINBFLGCA <> "" And left(.HINBFLGCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HINBFLGCA & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '53:ﾎｰﾙﾄﾞ工程
        If .HOLDKTCA <> "" And left(.HOLDKTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDKTCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '54:親ﾌﾞﾛｯｸID
        If .RPCRYNUMCA <> "" And left(.RPCRYNUMCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .RPCRYNUMCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '55:関連ﾌﾞﾛｯｸﾌﾗｸﾞ
        If .KBLKFLGCA <> "" And left(.KBLKFLGCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .KBLKFLGCA & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '56:向先
        If .PLANTCATCA <> "" And left(.PLANTCATCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .PLANTCATCA & "'" & vbLf
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
    
    #If Y3_CREATE = 1 Then
        ''XODY3作成  add 2009/01/08 SETmiyatake
        Call CreateOrUpdateXODY3(pXSDCA.CRYNUMCA, pXSDCA.SXLIDCA, pXSDCA.LIVKCA)  'upd SETkimizuka
    #End If

    CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", "DB", sDBName)
    CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :該当するﾚｺｰﾄﾞ有無をﾁｪｯｸ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_BlockID    ,I  ,String           ,SXLID
'      　　:p_Hinban     ,O  ,String           ,長さ
'      　　:p_Inpos      ,O  ,Integer          ,長さ
'      　　:戻り値       ,O  ,Boolean        　,ﾚｺｰﾄﾞなし(TRUE)/あり(FALSE)
'説明　　　：品番振替、ｸﾘｽﾀﾙ格上など死ﾚｺｰﾄﾞと同品番への変更に対応
'履歴　　　：2002/08/29 ohno
Public Function CheckUniqueRecord(p_BlockID As String, p_Hinban As String, p_Inpos As Integer) As Boolean
    Dim sql As String
    Dim rs As OraDynaset

    sql = "SELECT * FROM XSDCA WHERE CRYNUMCA = '" & p_BlockID
    sql = sql & "' AND HINBCA = '" & p_Hinban
    sql = sql & "' AND INPOSCA = " & p_Inpos

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        CheckUniqueRecord = True
    Else
        CheckUniqueRecord = False
    End If

End Function


'概要      :最終通過処理回数を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_sCrynum    ,I  ,String           ,ブロックID
'      　　:p_iInpos     ,I  ,Integer          ,開始位置
'      　　:戻り値       ,O  ,Integer        　,処理回数
Public Function GetNEMACOC(p_sCrynum As String, p_iInpos As Integer) As Integer
    Dim sql As String
    Dim rs As OraDynaset


    sql = "SELECT GNMACOCA FROM XSDCA WHERE CRYNUMCA = '" & p_sCrynum
    sql = sql & "' AND INPOSCA = " & p_iInpos
'    sql = sql & " AND KCKNTCA = (SELECT MAX(KCKNTCA) FROM XSDCA WHERE CRYNUMCA = '" & p_sCrynum
'    sql = sql & "' AND INPOSCA = " & p_iInpos & ")"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        GetNEMACOC = 1
    Else
        GetNEMACOC = CInt(rs.Fields("GNMACOCA"))
    End If

End Function


'概要      :テーブル「XSDCA」の工程をチェックする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型               ,説明
'          :tXSDCA()      ,I    ,typ_XSDCA        ,チェックレコード
'          :sNowCode      ,I    ,String           ,チェック工程
'          :sErrMsg       ,O    ,String           ,エラーメッセージ
'          :戻り値        ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :指定データの工程が引数で渡された工程と同じかチェックする
'          :複数ブロックのチェックは対応していない。
'           2006/03/10 新規作成　仕掛工程再チェック機能追加

Public Function DBDRV_CheckCodeXSDCA(tXSDCA() As typ_XSDCA, sNowCode As String, sErrMsg As String) As FUNCTION_RETURN

    Dim lsSql As String             'SQL全体
    Dim rs As OraDynaset            'RecordSet
    Dim tReadXSDCA() As typ_XSDCA   '取得データ
    Dim llLoopCnt   As Long
    Dim llBlockSt  As Long          'ブロックの配列内開始位置
    Dim llBlockEd  As Long          'ブロックの配列内終了位置
    Dim i As Long
    Dim j As Long
    Dim sDBName As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDCA_SQL.bas -- Function DBDRV_CheckCodeXSDCA"
    sErrMsg = ""
    sDBName = "XSDCA"
    llBlockSt = 1
    llBlockEd = UBound(tXSDCA)

'    For llLoopCnt = 1 To UBound(tXSDCA)
'        'ブロックの配列内開始位置を保持
'        llBlockSt = llLoopCnt
'        'ブロックの配列内終了位置を保持
'        llBlockEd = llLoopCnt
'        For i = llLoopCnt + 1 To UBound(tXSDCA)
'            If tXSDCA(i).CRYNUMCA <> tXSDCA(i - 1).CRYNUMCA Then
'                llBlockEd = i - 1
'                llLoopCnt = llBlockEd
'                Exit For
'            End If
'        Next i

    i = 0
    ReDim tReadXSDCA(0) As typ_XSDCA
    For llLoopCnt = 1 To llBlockEd
        If llLoopCnt = 1 Or _
           (llLoopCnt > 1 And tXSDCA(llLoopCnt).CRYNUMCA <> tXSDCA(llLoopCnt - 1).CRYNUMCA) Then
            ''SQLを組み立てる
            lsSql = ""
            lsSql = lsSql & " SELECT"
            lsSql = lsSql & "   CRYNUMCA"
            lsSql = lsSql & "  ,HINBCA"
            lsSql = lsSql & "  ,INPOSCA"
            lsSql = lsSql & "  ,GNWKNTCA"
            lsSql = lsSql & " FROM"
            lsSql = lsSql & "   XSDCA"
            lsSql = lsSql & " WHERE CRYNUMCA = '" & tXSDCA(llLoopCnt).CRYNUMCA & "'"
            lsSql = lsSql & "   AND LIVKCA = '0' "


            ''データを抽出する
            Set rs = OraDB.DBCreateDynaset(lsSql, ORADYN_DEFAULT)
            If rs Is Nothing Then
'                ReDim records(0)
                DBDRV_CheckCodeXSDCA = FUNCTION_RETURN_FAILURE
                Exit Function
            End If

            ''抽出結果を格納する
'            i = 0
'            ReDim tReadXSDCA(0) As typ_XSDCA
            Do Until rs.EOF 'データがなくなるまで取得
                i = i + 1
                ReDim Preserve tReadXSDCA(i) As typ_XSDCA
                With tReadXSDCA(i)
                    If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
                    If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
                    If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
                    If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
                End With
                rs.MoveNext
            Loop
            rs.Close
        End If
    Next llLoopCnt

        '同じブロックの範囲でループする
        For i = llBlockSt To llBlockEd
            For j = 1 To UBound(tReadXSDCA)
                'ブロック、品番、結晶内開始位置が同じ物を探す
                If Trim(tXSDCA(i).CRYNUMCA) = Trim(tReadXSDCA(j).CRYNUMCA) And _
                   Trim(tXSDCA(i).HINBCA) = Trim(tReadXSDCA(j).HINBCA) And _
                   Trim(tXSDCA(i).INPOSCA) = Trim(tReadXSDCA(j).INPOSCA) Then

                    '現在工程が同じかチェックする
                    If Trim(tReadXSDCA(j).GNWKNTCA) <> Trim(sNowCode) Then
                        '現在工程が違う = ブロックがすでに動いている場合、エラー終了
                        sErrMsg = GetMsgStr("EBLK6")
                        DBDRV_CheckCodeXSDCA = FUNCTION_RETURN_FAILURE
                        Exit Function
                    Else
                        '同じ場合、次の品番へ
                        Exit For
                    End If

                End If
            Next j
        Next i

'    Next llLoopCnt

    DBDRV_CheckCodeXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print lsSql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", "DB", sDBName)
    DBDRV_CheckCodeXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit


End Function


