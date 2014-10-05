Attribute VB_Name = "s_XSDCB_SQL"
'分割結晶(SXL) (XSDCB) ｱｸｾｽ関数

'●TEST用

'***テーブル「XSDCB」へのデータアクセス関数***
'＊注意 ﾊﾟﾗﾒｰﾀに値をｾｯﾄする時、まず全て初期化すること

Option Explicit

'●分割結晶(SXL)
Public Type typ_XSDCB
    SXLIDCB As String * 1        ' SXLID
    KCNTCB As Integer            ' 工程連番
    XTALCB As String * 12        ' 結晶番号
    INPOSCB As Integer           ' 結晶内開始位置
    LENCB As Integer             ' 長さ
    HINBCB As String * 8         ' 品番
    REVNUMCB As Integer          ' 電話番号改訂番号
    FACTORYCB As String * 1      ' 工場
    OPECB As String * 1          ' 操業条件
    MAICB As Integer             ' 実枚数
    WSRMAICB As Integer          ' WS洗後枚数
    WSNMAICB As Integer          ' WS洗浄欠落枚数
    WFCMAICB As Integer          ' 受入枚数
    SXLRMAICB As Integer         ' SXL指示(良品)
    SXLNMAICB As Integer         ' SXL指示(不良)
    WFCNMAICB As Integer         ' WFC内欠落枚数
    SXLEMAICB As Integer         ' SXL確定枚数
    SRMAICB As Integer           ' サンプル抜指示(良品)
    SNMAICB As Integer           ' サンプル抜指示(不良)
    STMAICB As Integer           ' サンプル枚数
    FURIMAICB As Integer         ' 振替枚数
    XTWORKCB As String * 2       ' 製造工場
    WFWORKCB As String * 2       ' ウェーハ製造
    FURYCCB As String * 3        ' 不良理由
    LSTCCB As String * 1         ' 採取状態区分
    LUFRCCB As String * 3        ' 格上コード
    LUFRBCB As String * 1        ' 格上区分
    LDERCCB As String * 3        ' 格下コード
    LDFRBCB As String * 1        ' 格下区分
    HOLDCCB As String * 3        ' ホールドコード
    HOLDBCB As String * 1        ' ホールド区分
    EXKUBCB As String * 1        ' 例外区分
    HENPKCB As String * 1        ' 返品区分
    LIVKCB As String * 1         ' 生死区分
    KANKCB As String * 1         ' 完了区分
    NFCB As String * 1           ' 入庫区分
    SAKJCB As String * 1         ' 削除区分
    TDAYCB As Date               ' 登録日付
    KDAYCB As Date               ' 更新日付
    SUMITCB As String * 1        ' SUMIT送信フラグ
    SNDKCB As String * 1         ' 返品区分
    SNDAYCB As Date              ' 送信日付
    'add start 2003/03/25 hitec)matsumoto ----
    NEWKNTCB As String           ' 最終通過工程
    GNWKNTCB As String           ' 現在工程
    MOTHINCB As String           ' 元品番
    'add end 2003/03/25 hitec)matsumoto ----
    PLANTCATCB As String         ' 向先 2007/08/30 SPK Tsutsumi Add
End Type

'更新用
Public Type typ_XSDCB_Update
    SXLIDCB As String            ' SXLID
    KCNTCB As String             ' 工程連番
    XTALCB As String             ' 結晶番号
    INPOSCB As String            ' 結晶内開始位置
    LENCB As String              ' 長さ
    HINBCB As String             ' 品番
    REVNUMCB As String           ' 電話番号改訂番号
    FACTORYCB As String          ' 工場
    OPECB As String              ' 操業条件
    MAICB As String              ' 実枚数
    WSRMAICB As String           ' WS洗後枚数
    WSNMAICB As String           ' WS洗浄欠落枚数
    WFCMAICB As String           ' 受入枚数
    SXLRMAICB As String          ' SXL指示(良品)
    SXLNMAICB As String          ' SXL指示(不良)
    WFCNMAICB As String          ' WFC内欠落枚数
    SXLEMAICB As String          ' SXL確定枚数
    SRMAICB As String            ' サンプル抜指示(良品)
    SNMAICB As String            ' サンプル抜指示(不良)
    STMAICB As String            ' サンプル枚数
    FURIMAICB As String          ' 振替枚数
    XTWORKCB As String           ' 製造工場
    WFWORKCB As String           ' ウェーハ製造
    FURYCCB As String            ' 不良理由
    LSTCCB As String             ' 採取状態区分
    LUFRCCB As String            ' 格上コード
    LUFRBCB As String            ' 格上区分
    LDERCCB As String            ' 格下コード
    LDFRBCB As String            ' 格下区分
    HOLDCCB As String            ' ホールドコード
    HOLDBCB As String            ' ホールド区分
    EXKUBCB As String            ' 例外区分
    HENPKCB As String            ' 返品区分
    LIVKCB As String             ' 生死区分
    KANKCB As String             ' 完了区分
    NFCB As String               ' 入庫区分
    SAKJCB As String             ' 削除区分
    TDAYCB As String             ' 登録日付
    KDAYCB As String             ' 更新日付
    SUMITCB As String            ' SUMIT送信フラグ
    SNDKCB As String             ' 返品区分
    SNDAYCB As String            ' 送信日付
    'add start 2003/03/25 hitec)matsumoto ----
    NEWKNTCB As String           ' 最終通過工程
    GNWKNTCB As String           ' 現在工程
    MOTHINCB As String           ' 元品番
    'add end 2003/03/25 hitec)matsumoto ----
    PLANTCATCB As String         ' 向先 2007/08/30 SPK Tsutsumi Add
End Type

'●SELECT●

'概要      :テーブル「XSDCB」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型               ,説明
'          :records()     ,O    ,typ_XSDCB     ,抽出レコード
'          :sqlWhere      ,I    ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I    ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :

Public Function DBDRV_GetXSDCB(records() As typ_XSDCB, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim sql As String       'SQL全体
    Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long


    ''SQLを組み立てる
    sqlBase = "Select * From XSDCB"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDCB = FUNCTION_RETURN_FAILURE
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
            If IsNull(rs.Fields("SXLIDCB")) = False Then .SXLIDCB = rs.Fields("SXLIDCB")
            If IsNull(rs.Fields("KCNTCB")) = False Then .KCNTCB = rs.Fields("KCNTCB")
            If IsNull(rs.Fields("XTALCB")) = False Then .XTALCB = rs.Fields("XTALCB")
            If IsNull(rs.Fields("INPOSCB")) = False Then .INPOSCB = rs.Fields("INPOSCB")
            If IsNull(rs.Fields("LENCB")) = False Then .LENCB = rs.Fields("LENCB")
            If IsNull(rs.Fields("HINBCB")) = False Then .HINBCB = rs.Fields("HINBCB")
            If IsNull(rs.Fields("REVNUMCB")) = False Then .REVNUMCB = rs.Fields("REVNUMCB")
            If IsNull(rs.Fields("FACTORYCB")) = False Then .FACTORYCB = rs.Fields("FACTORYCB")
            If IsNull(rs.Fields("OPECB")) = False Then .OPECB = rs.Fields("OPECB")
            If IsNull(rs.Fields("MAICB")) = False Then .MAICB = rs.Fields("MAICB")
            If IsNull(rs.Fields("WSRMAICB")) = False Then .WSRMAICB = rs.Fields("WSRMAICB")
            If IsNull(rs.Fields("WSNMAICB")) = False Then .WSNMAICB = rs.Fields("WSNMAICB")
            If IsNull(rs.Fields("WFCMAICB")) = False Then .WFCMAICB = rs.Fields("WFCMAICB")
            If IsNull(rs.Fields("SXLRMAICB")) = False Then .SXLRMAICB = rs.Fields("SXLRMAICB")
            If IsNull(rs.Fields("SXLNMAICB")) = False Then .SXLNMAICB = rs.Fields("SXLNMAICB")
            If IsNull(rs.Fields("WFCNMAICB")) = False Then .WFCNMAICB = rs.Fields("WFCNMAICB")
            If IsNull(rs.Fields("SXLEMAICB")) = False Then .SXLEMAICB = rs.Fields("SXLEMAICB")
            If IsNull(rs.Fields("SRMAICB")) = False Then .SRMAICB = rs.Fields("SRMAICB")
            If IsNull(rs.Fields("SNMAICB")) = False Then .SNMAICB = rs.Fields("SNMAICB")
            If IsNull(rs.Fields("STMAICB")) = False Then .STMAICB = rs.Fields("STMAICB")
            If IsNull(rs.Fields("FURIMAICB")) = False Then .FURIMAICB = rs.Fields("FURIMAICB")
            If IsNull(rs.Fields("XTWORKCB")) = False Then .XTWORKCB = rs.Fields("XTWORKCB")
            If IsNull(rs.Fields("WFWORKCB")) = False Then .WFWORKCB = rs.Fields("WFWORKCB")
            If IsNull(rs.Fields("FURYCCB")) = False Then .FURYCCB = rs.Fields("FURYCCB")
            If IsNull(rs.Fields("LSTCCB")) = False Then .LSTCCB = rs.Fields("LSTCCB")
            If IsNull(rs.Fields("LUFRCCB")) = False Then .LUFRCCB = rs.Fields("LUFRCCB")
            If IsNull(rs.Fields("LUFRBCB")) = False Then .LUFRBCB = rs.Fields("LUFRBCB")
            If IsNull(rs.Fields("LDERCCB")) = False Then .LDERCCB = rs.Fields("LDERCCB")
            If IsNull(rs.Fields("LDFRBCB")) = False Then .LDFRBCB = rs.Fields("LDFRBCB")
            If IsNull(rs.Fields("HOLDCCB")) = False Then .HOLDCCB = rs.Fields("HOLDCCB")
            If IsNull(rs.Fields("HOLDBCB")) = False Then .HOLDBCB = rs.Fields("HOLDBCB")
            If IsNull(rs.Fields("EXKUBCB")) = False Then .EXKUBCB = rs.Fields("EXKUBCB")
            If IsNull(rs.Fields("HENPKCB")) = False Then .HENPKCB = rs.Fields("HENPKCB")
            If IsNull(rs.Fields("LIVKCB")) = False Then .LIVKCB = rs.Fields("LIVKCB")
            If IsNull(rs.Fields("KANKCB")) = False Then .KANKCB = rs.Fields("KANKCB")
            If IsNull(rs.Fields("NFCB")) = False Then .NFCB = rs.Fields("NFCB")
            If IsNull(rs.Fields("SAKJCB")) = False Then .SAKJCB = rs.Fields("SAKJCB")
            If IsNull(rs.Fields("TDAYCB")) = False Then .TDAYCB = rs.Fields("TDAYCB")
            If IsNull(rs.Fields("KDAYCB")) = False Then .KDAYCB = rs.Fields("KDAYCB")
            If IsNull(rs.Fields("SUMITCB")) = False Then .SUMITCB = rs.Fields("SUMITCB")
            If IsNull(rs.Fields("SNDKCB")) = False Then .SNDKCB = rs.Fields("SNDKCB")
            If IsNull(rs.Fields("SNDAYCB")) = False Then .SNDAYCB = rs.Fields("SNDAYCB")
            'add start 2003/03/25 hitec)matsumoto ------
            If IsNull(rs.Fields("NEWKNTCB")) = False Then .NEWKNTCB = rs.Fields("NEWKNTCB")
            If IsNull(rs.Fields("GNWKNTCB")) = False Then .GNWKNTCB = rs.Fields("GNWKNTCB")
            If IsNull(rs.Fields("MOTHINCB")) = False Then .MOTHINCB = rs.Fields("MOTHINCB")
            'add edn   2003/03/25 hitec)matsumoto ------
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDCB = FUNCTION_RETURN_SUCCESS
End Function

'●UPDATE●

'●更新項目を構造体にセットして引き渡す

'概要      :テーブル「XSDCB」を更新する ptrn1
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型               ,説明
'          :records()     ,O   ,typ_XSDCB     ,更新レコード
'          :sqlWhere      ,I   ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I   ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O   ,FUNCTION_RETURN  ,抽出の成否
'説明      :

Public Function UpdateXSDCB(records As typ_XSDCB_Update, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
On Error GoTo proc_err
    gErr.Push "s_XSDCB_SQL.bas -- Function UpdateXSDCB"

    Dim sql As String       'SQL全体
'    Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
'    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
'    Dim i As Long
    Dim nowtime As Date
    Dim nowtime_sql As String   'サーバ時間(SQL文)
    
    nowtime = getSvrTime()    'サーバーの時間を取得するように変更 2003/6/4 tuku

'>>>>> .EditをSQL(UPDATE)文に変更　2009/06/18 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"

    With records
        
        ''SQLを組み立てる
        sql = "UPDATE XSDCB SET" & vbLf
        
        ''更新日付
        sql = sql & " KDAYCB = " & nowtime_sql & vbLf
        
        ''SXLID
        If .SXLIDCB <> "" And Left(.SXLIDCB, 1) <> vbNullChar Then
            sql = sql & ",SXLIDCB = '" & .SXLIDCB & "'" & vbLf
        End If
        
        ''工程連番
        If .KCNTCB <> "" Then
            sql = sql & ",KCNTCB = '" & CStr(CInt(.KCNTCB)) & "'" & vbLf
        End If
        
        ''結晶番号
        If .XTALCB <> "" And Left(.XTALCB, 1) <> vbNullChar Then
            sql = sql & ",XTALCB = '" & .XTALCB & "'" & vbLf
        End If
        
        ''結晶内開始位置
        If .INPOSCB <> "" Then
            sql = sql & ",INPOSCB = '" & CStr(CInt(.INPOSCB)) & "'" & vbLf
        End If
        
        ''長さ
        If .LENCB <> "" Then
            sql = sql & ",LENCB = '" & CStr(CInt(.LENCB)) & "'" & vbLf
        End If
        
        ''品番
        If .HINBCB <> "" And Left(.HINBCB, 1) <> vbNullChar Then
            sql = sql & ",HINBCB = '" & .HINBCB & "'" & vbLf
        End If
        
        ''製品番号改訂番号
        If .REVNUMCB <> "" Then
            sql = sql & ",REVNUMCB = '" & CStr(CInt(.REVNUMCB)) & "'" & vbLf
        End If
        
        ''工場
        If .FACTORYCB <> "" And Left(.FACTORYCB, 1) <> vbNullChar Then
            sql = sql & ",FACTORYCB = '" & .FACTORYCB & "'" & vbLf
        End If
        
        ''操業条件
        If .OPECB <> "" And Left(.OPECB, 1) <> vbNullChar Then
            sql = sql & ",OPECB = '" & .OPECB & "'" & vbLf
        End If
        
        ''実枚数
        If .MAICB <> "" Then
            sql = sql & ",MAICB = '" & CStr(CInt(.MAICB)) & "'" & vbLf
        End If
        
        ''WS洗後枚数
        If .WSRMAICB <> "" Then
            sql = sql & ",WSRMAICB = '" & CStr(CInt(.WSRMAICB)) & "'" & vbLf
        End If
        
        ''WS洗後欠落枚数
        If .WSNMAICB <> "" Then
            sql = sql & ",WSNMAICB = '" & CStr(CInt(.WSNMAICB)) & "'" & vbLf
        End If
        
        ''WFC受入枚数
        If .WFCMAICB <> "" Then
            sql = sql & ",WFCMAICB = '" & CStr(CInt(.WFCMAICB)) & "'" & vbLf
        End If
        
        ''SXL指示（良品）
        If .SXLRMAICB <> "" Then
            sql = sql & ",SXLRMAICB = '" & CStr(CInt(.SXLRMAICB)) & "'" & vbLf
        End If
        
        ''SXL指示（不良）
        If .SXLNMAICB <> "" Then
            sql = sql & ",SXLNMAICB = '" & CStr(CInt(.SXLNMAICB)) & "'" & vbLf
        End If
        
        ''WFC内欠落枚数
        If .WFCNMAICB <> "" Then
            sql = sql & ",WFCNMAICB = '" & CStr(CInt(.WFCNMAICB)) & "'" & vbLf
        End If
        
        ''SXL確定枚数
        If .SXLEMAICB <> "" Then
            sql = sql & ",SXLEMAICB = '" & CStr(CInt(.SXLEMAICB)) & "'" & vbLf
        End If
        
        ''サンプル抜指示（良品）
        If .SRMAICB <> "" Then
            sql = sql & ",SRMAICB = '" & CStr(CInt(.SRMAICB)) & "'" & vbLf
        End If
        
        ''サンプル抜指示（不良）
        If .SNMAICB <> "" Then
            sql = sql & ",SNMAICB = '" & CStr(CInt(.SNMAICB)) & "'" & vbLf
        End If
        
        ''サンプル枚数
        If .STMAICB <> "" Then
            sql = sql & ",STMAICB = '" & CStr(CInt(.STMAICB)) & "'" & vbLf
        End If
        
        ''振替枚数
        If .FURIMAICB <> "" Then
            sql = sql & ",FURIMAICB = '" & CStr(CInt(.FURIMAICB)) & "'" & vbLf
        End If
        
        ''製造工場
        If .XTWORKCB <> "" And Left(.XTWORKCB, 1) <> vbNullChar Then
            sql = sql & ",XTWORKCB = '" & .XTWORKCB & "'" & vbLf
        End If
        
        ''ウェーハ製造
        If .WFWORKCB <> "" And Left(.WFWORKCB, 1) <> vbNullChar Then
            sql = sql & ",WFWORKCB = '" & .WFWORKCB & "'" & vbLf
        End If
        
        ''不良理由
        If .FURYCCB <> "" And Left(.FURYCCB, 1) <> vbNullChar Then
            sql = sql & ",FURYCCB = '" & .FURYCCB & "'" & vbLf
        End If
        
        ''最終状態区分
        If .LSTCCB <> "" And Left(.LSTCCB, 1) <> vbNullChar Then
            sql = sql & ",LSTCCB = '" & .LSTCCB & "'" & vbLf
        End If
        
        ''格上コード
        If .LUFRCCB <> "" And Left(.LUFRCCB, 1) <> vbNullChar Then
            sql = sql & ",LUFRCCB = '" & .LUFRCCB & "'" & vbLf
        End If
        
        ''格上区分
        If .LUFRBCB <> "" And Left(.LUFRBCB, 1) <> vbNullChar Then
            sql = sql & ",LUFRBCB = '" & .LUFRBCB & "'" & vbLf
        End If
        
        ''格下コード
        If .LDERCCB <> "" And Left(.LDERCCB, 1) <> vbNullChar Then
            sql = sql & ",LDERCCB = '" & .LDERCCB & "'" & vbLf
        End If
        
        ''格下区分
        If .LDFRBCB <> "" And Left(.LDFRBCB, 1) <> vbNullChar Then
            sql = sql & ",LDFRBCB = '" & .LDFRBCB & "'" & vbLf
        End If
        
        ''ホールドコード
        If .HOLDCCB <> "" And Left(.HOLDCCB, 1) <> vbNullChar Then
            sql = sql & ",HOLDCCB = '" & .HOLDCCB & "'" & vbLf
        End If
        
        ''ホールド区分
        If .HOLDBCB <> "" And Left(.HOLDBCB, 1) <> vbNullChar Then
            sql = sql & ",HOLDBCB = '" & .HOLDBCB & "'" & vbLf
        End If
        
        ''例外区分
        If .EXKUBCB <> "" And Left(.EXKUBCB, 1) <> vbNullChar Then
            sql = sql & ",EXKUBCB = '" & .EXKUBCB & "'" & vbLf
        End If
        
        ''返品区分
        If .HENPKCB <> "" And Left(.HENPKCB, 1) <> vbNullChar Then
            sql = sql & ",HENPKCB = '" & .HENPKCB & "'" & vbLf
        End If
        
        ''生死区分
        If .LIVKCB <> "" And Left(.LIVKCB, 1) <> vbNullChar Then
            sql = sql & ",LIVKCB = '" & .LIVKCB & "'" & vbLf
        End If
        
        ''完了区分
        If .KANKCB <> "" And Left(.KANKCB, 1) <> vbNullChar Then
            sql = sql & ",KANKCB = '" & .KANKCB & "'" & vbLf
        End If
        
        ''入庫区分
        If .NFCB <> "" And Left(.NFCB, 1) <> vbNullChar Then
            sql = sql & ",NFCB = '" & .NFCB & "'" & vbLf
        End If
        
        ''削除区分
        If .SAKJCB <> "" And Left(.SAKJCB, 1) <> vbNullChar Then
            sql = sql & ",SAKJCB = '" & .SAKJCB & "'" & vbLf
        End If
        
        ''登録日付
        If .TDAYCB <> "" Then
            sql = sql & ",TDAYCB = TO_DATE('" & Format$(CDate(.TDAYCB), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''SUMIT送信フラグ
        If .SUMITCB <> "" And Left(.SUMITCB, 1) <> vbNullChar Then
            sql = sql & ",SUMITCB = '" & .SUMITCB & "'" & vbLf
        End If
        
        ''送信フラグ
        If .SNDKCB <> "" And Left(.SNDKCB, 1) <> vbNullChar Then
            sql = sql & ",SNDKCB = '" & .SNDKCB & "'" & vbLf
        End If
        
        ''送信日付
        If .SNDAYCB <> "" Then
            sql = sql & ",SNDAYCB = TO_DATE('" & Format$(CDate(.SNDAYCB), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''最終通過工程
        If .NEWKNTCB <> "" And Left(.NEWKNTCB, 1) <> vbNullChar Then
            sql = sql & ",NEWKNTCB = '" & .NEWKNTCB & "'" & vbLf
        End If
        
        ''現在工程
        If .GNWKNTCB <> "" And Left(.GNWKNTCB, 1) <> vbNullChar Then
            sql = sql & ",GNWKNTCB = '" & .GNWKNTCB & "'" & vbLf
        End If
        
        ''振替品番(元）
        If .MOTHINCB <> "" And Left(.MOTHINCB, 1) <> vbNullChar Then
            sql = sql & ",MOTHINCB = '" & .MOTHINCB & "'" & vbLf
        End If
        
        ''事業所区分
        If .PLANTCATCB <> "" And Left(.PLANTCATCB, 2) <> vbNullChar Then
            sql = sql & ",PLANTCATCB = '" & .PLANTCATCB & "'" & vbLf
        End If

        sql = sql & " " & sqlWhere & vbLf
    
        'SQLを実行
        recCnt = OraDB.ExecuteSQL(sql)
        
        '返り値が1以外はエラー
        If recCnt < 0 Then
            GoTo proc_err
        ElseIf recCnt = 0 Then
            '0件更新…エラー(既存通り)
            UpdateXSDCB = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf recCnt > 1 Then
            '複数件更新…エラー(既存は複数SELECTした最初の一件のみ更新)
            UpdateXSDCB = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    
    End With
'<<<<< .EditをSQL(UPDATE)文に変更　2009/06/18 SETsw kubota ------------------

    UpdateXSDCB = FUNCTION_RETURN_SUCCESS


proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    UpdateXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'●INSERT●  NULLの場合、charならスペース、NumberならNULLを入れる

'概要      :テーブル「XSDCB」にレコードを挿入する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pXSDCB 　　  ,I  ,typ_XSDCB_Update   ,XSDCB更新用ﾃﾞｰﾀ
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function CreateXSDCB(pXSDCB As typ_XSDCB_Update, sErrMsg As String) As FUNCTION_RETURN


    Dim sql As String
    Dim sDbName As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim recCnt As Long      'レコード数
    Dim nowtime As Date
    Dim nowtime_sql As String   'サーバ時間(SQL文)
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDCB_SQL.bas -- Function CreateXSDCB"
    sErrMsg = ""
    sDbName = "XSDCB"
    'nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    nowtime = getSvrTime()    'サーバーの時間を取得するように変更 2003/6/4 tuku

'>>>>> .AddNewをSQL(INSERT)文に変更　2009/06/18 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With pXSDCB
        sql = "INSERT INTO XSDCB ("
        sql = sql & " SXLIDCB"      ' 1:SXLID
        sql = sql & ",KCNTCB"       ' 2:工程連番
        sql = sql & ",XTALCB"       ' 3:結晶番号
        sql = sql & ",INPOSCB"      ' 4:結晶内開始位置
        sql = sql & ",LENCB"        ' 5:長さ
        sql = sql & ",HINBCB"       ' 6:品番
        sql = sql & ",REVNUMCB"     ' 7:製品番号改訂番号
        sql = sql & ",FACTORYCB"    ' 8:工場
        sql = sql & ",OPECB"        ' 9:操業条件
        sql = sql & ",MAICB"        '10:実枚数
        sql = sql & ",WSRMAICB"     '11:WS洗後枚数
        sql = sql & ",WSNMAICB"     '12:WS洗後欠落枚数
        sql = sql & ",WFCMAICB"     '13:WFC受入枚数
        sql = sql & ",SXLRMAICB"    '14:SXL指示（良品）
        sql = sql & ",SXLNMAICB"    '15:SXL指示（不良）
        sql = sql & ",WFCNMAICB"    '16:WFC内欠落枚数
        sql = sql & ",SXLEMAICB"    '17:SXL確定枚数
        sql = sql & ",SRMAICB"      '18:サンプル抜指示（良品）
        sql = sql & ",SNMAICB"      '19:サンプル抜指示（不良）
        sql = sql & ",STMAICB"      '20:サンプル枚数
        sql = sql & ",FURIMAICB"    '21:振替枚数
        sql = sql & ",XTWORKCB"     '22:製造工場
        sql = sql & ",WFWORKCB"     '23:ウェーハ製造
        sql = sql & ",FURYCCB"      '24:不良理由
        sql = sql & ",LSTCCB"       '25:最終状態区分
        sql = sql & ",LUFRCCB"      '26:格上コード
        sql = sql & ",LUFRBCB"      '27:格上区分
        sql = sql & ",LDERCCB"      '28:格下コード
        sql = sql & ",LDFRBCB"      '29:格下区分
        sql = sql & ",HOLDCCB"      '30:ホールドコード
        sql = sql & ",HOLDBCB"      '31:ホールド区分
        sql = sql & ",EXKUBCB"      '32:例外区分
        sql = sql & ",HENPKCB"      '33:返品区分
        sql = sql & ",LIVKCB"       '34:生死区分
        sql = sql & ",KANKCB"       '35:完了区分
        sql = sql & ",NFCB"         '36:入庫区分
        sql = sql & ",SAKJCB"       '37:削除区分
        sql = sql & ",TDAYCB"       '38:登録日付
        sql = sql & ",KDAYCB"       '39:更新日付
        sql = sql & ",SUMITCB"      '40:SUMIT送信フラグ
        sql = sql & ",SNDKCB"       '41:送信フラグ
        sql = sql & ",SNDAYCB"      '42:送信日付
        sql = sql & ",NEWKNTCB"     '43:最終通過工程
        sql = sql & ",GNWKNTCB"     '44:現在工程
        sql = sql & ",MOTHINCB"     '45:振替品番(元）
        sql = sql & ",PLANTCATCB"   '46:事業所区分
        sql = sql & ")"
        sql = sql & "VALUES ("

        ' 1:SXLID
        If .SXLIDCB <> "" Then
            sql = sql & " '" & .SXLIDCB & "'" & vbLf
        Else
            sql = sql & " '" & Space(13) & "'" & vbLf
        End If

        ' 2:工程連番
        If .KCNTCB <> "" Then
            sql = sql & ",'" & CStr(CInt(.KCNTCB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 3:結晶番号
        If .XTALCB <> "" Then
            sql = sql & ",'" & .XTALCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If
        
        ' 4:結晶内開始位置
        If .INPOSCB <> "" Then
            sql = sql & ",'" & .INPOSCB & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 5:長さ
        If .LENCB <> "" Then
            sql = sql & ",'" & .LENCB & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 6:品番
        If .HINBCB <> "" Then
            sql = sql & ",'" & .HINBCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        ' 7:製品番号改訂番号
        If .REVNUMCB <> "" Then
            sql = sql & ",'" & .REVNUMCB & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 8:工場
        If .FACTORYCB <> "" Then
            sql = sql & ",'" & .FACTORYCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 9:操業条件
        If .OPECB <> "" Then
            sql = sql & ",'" & .OPECB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '10:実枚数
        If .MAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.MAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '11:WS洗後枚数
        If .WSRMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.WSRMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '12:WS洗後欠落枚数
        If .WSNMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.WSNMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '13:WFC受入枚数
        If .WFCMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.WFCMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '14:SXL指示（良品）
        If .SXLRMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.SXLRMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '15:SXL指示（不良）
        If .SXLNMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.SXLNMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '16:WFC内欠落枚数
        If .WFCNMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.WFCNMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '17:SXL確定枚数
        If .SXLEMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.SXLEMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '18:サンプル抜指示（良品）
        If .SRMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.SRMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '19:サンプル抜指示（不良）
        If .SNMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.SNMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '20:サンプル枚数
        If .STMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.STMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '21:振替枚数
        If .FURIMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.FURIMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '22:製造工場
        sql = sql & ",'" & FACTORYCD & "'" & vbLf

        '23:ウェーハ製造
        If .WFWORKCB <> "" Then
            sql = sql & ",'" & .WFWORKCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '24:不良理由
        If .FURYCCB <> "" Then
            sql = sql & ",'" & .FURYCCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '25:最終状態区分
        If .LSTCCB <> "" Then
            sql = sql & ",'" & .LSTCCB & "'" & vbLf
        Else
            sql = sql & ",'T'" & vbLf           '通常
        End If

        '26:格上コード
        If .LUFRCCB <> "" Then
            sql = sql & ",'" & .LUFRCCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '27:格上区分
        If .LUFRBCB <> "" Then
            sql = sql & ",'" & .LUFRBCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '28:格下コード
        If .LDERCCB <> "" Then
            sql = sql & ",'" & .LDERCCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '29:格下区分
        If .LDFRBCB <> "" Then
            sql = sql & ",'" & .LDFRBCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '30:ホールドコード
        If .HOLDCCB <> "" Then
            sql = sql & ",'" & .HOLDCCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '31:ホールド区分
        If .HOLDBCB <> "" Then
            sql = sql & ",'" & .HOLDBCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '32:例外区分
        If .EXKUBCB <> "" Then
            sql = sql & ",'" & .EXKUBCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '33:返品区分
        If .HENPKCB <> "" Then
            sql = sql & ",'" & .HENPKCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '34:生死区分
        If .LIVKCB <> "" Then
            sql = sql & ",'" & .LIVKCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '35:完了区分
        If .KANKCB <> "" Then
            sql = sql & ",'" & .KANKCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '36:入庫区分
        If .NFCB <> "" Then
            sql = sql & ",'" & .NFCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '37:削除区分
        If .SAKJCB <> "" Then
            sql = sql & ",'" & .SAKJCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '38:登録日付
        sql = sql & "," & nowtime_sql & vbLf
        
        '39:更新日付
        sql = sql & "," & nowtime_sql & vbLf

        '40:SUMIT送信フラグ
        If .SUMITCB <> "" Then
            sql = sql & ",'" & .SUMITCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '41:送信フラグ
        If .SNDKCB <> "" Then
            sql = sql & ",'" & .SNDKCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '42:送信日付
        sql = sql & ",NULL" & vbLf

        '43:最終通過工程
        If .NEWKNTCB <> "" Then
            sql = sql & ",'" & .NEWKNTCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If
        
        '44:現在工程
        If .GNWKNTCB <> "" Then
            sql = sql & ",'" & .GNWKNTCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If
        
        '45:振替品番(元）
        If .MOTHINCB <> "" Then
            sql = sql & ",'" & .MOTHINCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        '46:事業所区分
        If .PLANTCATCB <> "" Then
            sql = sql & ",'" & .PLANTCATCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        sql = sql & ")" & vbLf
    
        'SQLを実行
        If OraDB.ExecuteSQL(sql) < 1 Then
            GoTo proc_err
        End If

    End With
'<<<<< .AddNewをSQL(INSERT)文に変更　2009/06/18 SETsw kubota ------------------

    CreateXSDCB = FUNCTION_RETURN_SUCCESS

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
    CreateXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :工程連番を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_SXLID      ,I  ,String           ,SXLID
'      　　:戻り値       ,O  ,Integer        　,工程連番
Public Function GetKCNTCB(p_SXLID As String) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT KCNTCB FROM XSDCB WHERE SXLIDCB = '" & p_SXLID & "'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("KCNTCB")) Then
        GetKCNTCB = 1
    Else
        GetKCNTCB = CInt(rs.Fields("KCNTCB")) + 1
    End If
    
End Function


'概要      :該当するﾚｺｰﾄﾞ有無をﾁｪｯｸ(あれば長さを取得する)
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_SXLID      ,I  ,String           ,SXLID
'      　　:p_Length     ,O  ,Integer          ,長さ
'      　　:戻り値       ,O  ,Integer        　,ﾚｺｰﾄﾞ数
Public Function CheckSXLrecord(p_SXLID As String, p_Length As Integer) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT LENCB FROM XSDCB WHERE SXLIDCB = '" & p_SXLID & "'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("LENCB")) Then
        CheckSXLrecord = 0
    Else
        CheckSXLrecord = 1
        p_Length = CInt(rs.Fields("LENCB"))
    End If
    
End Function


