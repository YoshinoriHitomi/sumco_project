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
    NEWKNTCB As String           ' 最終通過工程
    GNWKNTCB As String           ' 現在工程
    MOTHINCB As String           ' 元品番
    RLENCB As Integer            ' 理論長さ
    SHOLDCLSCB As String         ' ホールド区分(SXL確定)
    PLANTCATCB As String         ' 向先
    KBLKFLGCB As String * 1      ' 関連ﾌﾞﾛｯｸﾌﾗｸﾞ　08/01/31 ooba
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
    NEWKNTCB As String           ' 最終通過工程
    GNWKNTCB As String           ' 現在工程
    MOTHINCB As String           ' 元品番
    RLENCB As String             ' 理論長さ
    SHOLDCLSCB As String         ' ホールド区分(SXL確定)
    PLANTCATCB As String         ' 向先
    KBLKFLGCB As String          ' 関連ﾌﾞﾛｯｸﾌﾗｸﾞ　08/01/31 ooba
End Type

'●SELECT●
'*******************************************************************************************
'*    関数名        : DBDRV_GetXSDCB
'*
'*    処理概要      : 1.テーブル「XSDCB」から条件にあったレコードを抽出する
'*
'*    パラメータ    : 変数名       ,IO   ,型            ,説明
'*                   records()     ,O    ,typ_XSDCB     ,抽出レコード
'*                   sqlWhere      ,I    ,String        ,抽出条件(SQLのWhere節:省略可能)
'*                   sqlOrder      ,I    ,String        ,抽出順序(SQLのOrder by節:省略可能)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function DBDRV_GetXSDCB(records() As typ_XSDCB, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN

    Dim sSql        As String       ' SQL全体
    Dim sSqlBase    As String       ' SQL基本部(WHERE節の前まで)
    Dim rs          As OraDynaset   ' RecordSet
    Dim intRecCnt   As Long         ' レコード数
    Dim i           As Long

    ' SQLを組み立てる
    sSqlBase = "Select * From XSDCB"
    sSql = sSqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sSql = sSql & " " & sqlWhere & " " & sqlOrder
    End If

    ' データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDCB = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ' 抽出結果を格納する
    intRecCnt = rs.RecordCount
    ReDim records(intRecCnt)
    If intRecCnt = 0 Then
        Exit Function
    End If
    For i = 1 To intRecCnt
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
            If IsNull(rs.Fields("NEWKNTCB")) = False Then .NEWKNTCB = rs.Fields("NEWKNTCB")
            If IsNull(rs.Fields("GNWKNTCB")) = False Then .GNWKNTCB = rs.Fields("GNWKNTCB")
            If IsNull(rs.Fields("MOTHINCB")) = False Then .MOTHINCB = rs.Fields("MOTHINCB")
            If IsNull(rs.Fields("RLENCB")) = False Then .RLENCB = rs.Fields("RLENCB")
            If IsNull(rs.Fields("SHOLDCLSCB")) = False Then .SHOLDCLSCB = rs.Fields("SHOLDCLSCB")
            If IsNull(rs.Fields("PLANTCATCB")) = False Then .PLANTCATCB = rs.Fields("PLANTCATCB")   ' 向先 2007/09/04 SPK Tsutsumi Add
            If IsNull(rs.Fields("KBLKFLGCB")) = False Then .KBLKFLGCB = rs.Fields("KBLKFLGCB")      '08/01/31 ooba
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDCB = FUNCTION_RETURN_SUCCESS
End Function

'●UPDATE●
'*******************************************************************************************
'*    関数名        : UpdateXSDCB
'*
'*    処理概要      : 1.テーブル「XSDCB」を更新する ptrn1
'*                    (更新項目を構造体にセットして引き渡す)
'*
'*    パラメータ    : 変数名       ,IO   ,型            ,説明
'*                   records()     ,O    ,typ_XSDCB     ,更新レコード
'*                   sqlWhere      ,I    ,String        ,抽出条件(SQLのWhere節:省略可能)
'*                   sqlOrder      ,I    ,String        ,抽出順序(SQLのOrder by節:省略可能)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function UpdateXSDCB(records As typ_XSDCB_Update, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    On Error GoTo proc_err
'    gErr.Push "s_XSDCB_SQL.bas -- Function CreateXSDCB"

    Dim sSql        As String       ' SQL全体
    Dim sSqlBase    As String       ' SQL基本部(WHERE節の前まで)
    Dim rs          As OraDynaset   ' RecordSet
    Dim intRecCnt   As Long         ' レコード数
    Dim i           As Long
    Dim dtmNowtime  As Date

    dtmNowtime = getSvrTime()       ' サーバーの時間を取得するように変更 2003/6/4 tuku

'>>>>> Edit-->UPDATEに変更　2009/07/21　SSS.Marushita
    
    With records
        
        ''SQLを組み立てる
        sSql = "UPDATE XSDCB SET" & vbLf
        
        ''更新日付
        sSql = sSql & " KDAYCB = TO_DATE('" & Format$(dtmNowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
        
        ''SXLID
        If .SXLIDCB <> "" And left(.SXLIDCB, 1) <> vbNullChar Then
            sSql = sSql & ",SXLIDCB = '" & .SXLIDCB & "'" & vbLf
        End If
        
        ''工程連番
        If .KCNTCB <> "" Then
            sSql = sSql & ",KCNTCB = '" & CStr(CInt(.KCNTCB)) & "'" & vbLf
        End If
        
        ''結晶番号
        If .XTALCB <> "" And left(.XTALCB, 1) <> vbNullChar Then
            sSql = sSql & ",XTALCB = '" & .XTALCB & "'" & vbLf
        End If
        
        ''結晶内開始位置
        If .INPOSCB <> "" Then
            sSql = sSql & ",INPOSCB = '" & CStr(CInt(.INPOSCB)) & "'" & vbLf
        End If
        
        ''長さ
        If .LENCB <> "" Then
            sSql = sSql & ",LENCB = '" & CStr(CInt(.LENCB)) & "'" & vbLf
        End If
        
        ''品番
        If .HINBCB <> "" And left(.HINBCB, 1) <> vbNullChar Then
            sSql = sSql & ",HINBCB = '" & .HINBCB & "'" & vbLf
        End If
        
        ''製品番号改訂番号
        If .REVNUMCB <> "" Then
            sSql = sSql & ",REVNUMCB = '" & CStr(CInt(.REVNUMCB)) & "'" & vbLf
        End If
        
        ''工場
        If .FACTORYCB <> "" And left(.FACTORYCB, 1) <> vbNullChar Then
            sSql = sSql & ",FACTORYCB = '" & .FACTORYCB & "'" & vbLf
        End If
        
        ''操業条件
        If .OPECB <> "" And left(.OPECB, 1) <> vbNullChar Then
            sSql = sSql & ",OPECB = '" & .OPECB & "'" & vbLf
        End If
        
        ''実枚数
        If .MAICB <> "" Then
            sSql = sSql & ",MAICB = '" & CStr(CInt(.MAICB)) & "'" & vbLf
        End If
        
        ''WS洗後枚数
        If .WSRMAICB <> "" Then
            sSql = sSql & ",WSRMAICB = '" & CStr(CInt(.WSRMAICB)) & "'" & vbLf
        End If
        
        ''WS洗後欠落枚数
        If .WSNMAICB <> "" Then
            sSql = sSql & ",WSNMAICB = '" & CStr(CInt(.WSNMAICB)) & "'" & vbLf
        End If
        
        ''WFC受入枚数
        If .WFCMAICB <> "" Then
            sSql = sSql & ",WFCMAICB = '" & CStr(CInt(.WFCMAICB)) & "'" & vbLf
        End If
        
        ''SXL指示（良品）
        If .SXLRMAICB <> "" Then
            sSql = sSql & ",SXLRMAICB = '" & CStr(CInt(.SXLRMAICB)) & "'" & vbLf
        End If
        
        ''SXL指示（不良）
        If .SXLNMAICB <> "" Then
            sSql = sSql & ",SXLNMAICB = '" & CStr(CInt(.SXLNMAICB)) & "'" & vbLf
        End If
        
        ''WFC内欠落枚数
        If .WFCNMAICB <> "" Then
            sSql = sSql & ",WFCNMAICB = '" & CStr(CInt(.WFCNMAICB)) & "'" & vbLf
        End If
        
        ''SXL確定枚数
        If .SXLEMAICB <> "" Then
            sSql = sSql & ",SXLEMAICB = '" & CStr(CInt(.SXLEMAICB)) & "'" & vbLf
        End If
        
        ''サンプル抜指示（良品）
        If .SRMAICB <> "" Then
            sSql = sSql & ",SRMAICB = '" & CStr(CInt(.SRMAICB)) & "'" & vbLf
        End If
        
        ''サンプル抜指示（不良）
        If .SNMAICB <> "" Then
            sSql = sSql & ",SNMAICB = '" & CStr(CInt(.SNMAICB)) & "'" & vbLf
        End If
        
        ''サンプル枚数
        If .STMAICB <> "" Then
            sSql = sSql & ",STMAICB = '" & CStr(CInt(.STMAICB)) & "'" & vbLf
        End If
        
        ''振替枚数
        If .FURIMAICB <> "" Then
            sSql = sSql & ",FURIMAICB = '" & CStr(CInt(.FURIMAICB)) & "'" & vbLf
        End If
        
        ''製造工場
        If .XTWORKCB <> "" And left(.XTWORKCB, 1) <> vbNullChar Then
            sSql = sSql & ",XTWORKCB = '" & .XTWORKCB & "'" & vbLf
        End If
        
        ''ウェーハ製造
        If .WFWORKCB <> "" And left(.WFWORKCB, 1) <> vbNullChar Then
            sSql = sSql & ",WFWORKCB = '" & .WFWORKCB & "'" & vbLf
        End If
        
        ''不良理由
        If .FURYCCB <> "" And left(.FURYCCB, 1) <> vbNullChar Then
            sSql = sSql & ",FURYCCB = '" & .FURYCCB & "'" & vbLf
        End If
        
        ''最終状態区分
        If .LSTCCB <> "" And left(.LSTCCB, 1) <> vbNullChar Then
            sSql = sSql & ",LSTCCB = '" & .LSTCCB & "'" & vbLf
        End If
        
        ''格上コード
        If .LUFRCCB <> "" And left(.LUFRCCB, 1) <> vbNullChar Then
            sSql = sSql & ",LUFRCCB = '" & .LUFRCCB & "'" & vbLf
        End If
        
        ''格上区分
        If .LUFRBCB <> "" And left(.LUFRBCB, 1) <> vbNullChar Then
            sSql = sSql & ",LUFRBCB = '" & .LUFRBCB & "'" & vbLf
        End If
        
        ''格下コード
        If .LDERCCB <> "" And left(.LDERCCB, 1) <> vbNullChar Then
            sSql = sSql & ",LDERCCB = '" & .LDERCCB & "'" & vbLf
        End If
        
        ''格下区分
        If .LDFRBCB <> "" And left(.LDFRBCB, 1) <> vbNullChar Then
            sSql = sSql & ",LDFRBCB = '" & .LDFRBCB & "'" & vbLf
        End If
        
        ''ホールドコード
        If .HOLDCCB <> "" And left(.HOLDCCB, 1) <> vbNullChar Then
            sSql = sSql & ",HOLDCCB = '" & .HOLDCCB & "'" & vbLf
        End If
        
        ''ホールド区分
        If .HOLDBCB <> "" And left(.HOLDBCB, 1) <> vbNullChar Then
            sSql = sSql & ",HOLDBCB = '" & .HOLDBCB & "'" & vbLf
        End If
        
        ''例外区分
        If .EXKUBCB <> "" And left(.EXKUBCB, 1) <> vbNullChar Then
            sSql = sSql & ",EXKUBCB = '" & .EXKUBCB & "'" & vbLf
        End If
        
        ''返品区分
        If .HENPKCB <> "" And left(.HENPKCB, 1) <> vbNullChar Then
            sSql = sSql & ",HENPKCB = '" & .HENPKCB & "'" & vbLf
        End If
        
        ''生死区分
        If .LIVKCB <> "" And left(.LIVKCB, 1) <> vbNullChar Then
            sSql = sSql & ",LIVKCB = '" & .LIVKCB & "'" & vbLf
        End If
        
        ''完了区分
        If .KANKCB <> "" And left(.KANKCB, 1) <> vbNullChar Then
            sSql = sSql & ",KANKCB = '" & .KANKCB & "'" & vbLf
        End If
        
        ''入庫区分
        If .NFCB <> "" And left(.NFCB, 1) <> vbNullChar Then
            sSql = sSql & ",NFCB = '" & .NFCB & "'" & vbLf
        End If
        
        ''削除区分
        If .SAKJCB <> "" And left(.SAKJCB, 1) <> vbNullChar Then
            sSql = sSql & ",SAKJCB = '" & .SAKJCB & "'" & vbLf
        End If
        
        ''登録日付
        If .TDAYCB <> "" Then
            sSql = sSql & ",TDAYCB = TO_DATE('" & Format$(CDate(.TDAYCB), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If

        ''SUMIT送信フラグ
        If .SUMITCB <> "" And left(.SUMITCB, 1) <> vbNullChar Then
            sSql = sSql & ",SUMITCB = '" & .SUMITCB & "'" & vbLf
        End If
        
        ''送信フラグ
        If .SNDKCB <> "" And left(.SNDKCB, 1) <> vbNullChar Then
            sSql = sSql & ",SNDKCB = '" & .SNDKCB & "'" & vbLf
        End If
        
        ''送信日付
        If .SNDAYCB <> "" Then
            sSql = sSql & ",SNDAYCB = TO_DATE('" & Format$(CDate(.SNDAYCB), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''最終通過工程
        If .NEWKNTCB <> "" And left(.NEWKNTCB, 1) <> vbNullChar Then
            sSql = sSql & ",NEWKNTCB = '" & .NEWKNTCB & "'" & vbLf
        End If
        
        ''現在工程
        If .GNWKNTCB <> "" And left(.GNWKNTCB, 1) <> vbNullChar Then
            sSql = sSql & ",GNWKNTCB = '" & .GNWKNTCB & "'" & vbLf
        End If
        
        ''振替品番(元）
        If .MOTHINCB <> "" And left(.MOTHINCB, 1) <> vbNullChar Then
            sSql = sSql & ",MOTHINCB = '" & .MOTHINCB & "'" & vbLf
        End If

        ''理論長さ
        If .RLENCB <> "" And left(.RLENCB, 1) <> vbNullChar Then
            sSql = sSql & ",RLENCB = '" & .RLENCB & "'" & vbLf
        End If
        
        ''ホールド区分(SXL確定)
        If .SHOLDCLSCB <> "" And left(.SHOLDCLSCB, 1) <> vbNullChar Then
            sSql = sSql & ",SHOLDCLSCB = '" & .SHOLDCLSCB & "'" & vbLf
        End If

        ''向先
        If .PLANTCATCB <> "" And left(.PLANTCATCB, 2) <> vbNullChar Then
            sSql = sSql & ",PLANTCATCB = '" & .PLANTCATCB & "'" & vbLf
        End If
            
        ''関連ﾌﾞﾛｯｸﾌﾗｸﾞ
        If .KBLKFLGCB <> "" And left(.KBLKFLGCB, 1) <> vbNullChar Then
            sSql = sSql & ",KBLKFLGCB = '" & .KBLKFLGCB & "'" & vbLf
        End If
        
        sSql = sSql & " " & sqlWhere & vbLf
    
        'SQLを実行
        intRecCnt = OraDB.ExecuteSQL(sSql)
        
        '返り値が1以外はエラー
        If intRecCnt < 0 Then
            GoTo proc_err
        ElseIf intRecCnt = 0 Then
            '0件更新…エラー(既存通り)
            UpdateXSDCB = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf intRecCnt > 1 Then
            '複数件更新…エラー(既存は複数SELECTした最初の一件のみ更新)
            UpdateXSDCB = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    
    End With

'<<<<< Edit-->UPDATEに変更　2009/07/21　SSS.Marushita
    
    UpdateXSDCB = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
'    gErr.Pop
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    UpdateXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'●INSERT●  NULLの場合、charならスペース、NumberならNULLを入れる
'*******************************************************************************************
'*    関数名        : CreateXSDCB
'*
'*    処理概要      : 1.テーブル「XSDCB」にレコードを挿入する
'*                      (NULLの場合、charならスペース、NumberならNULLを入れる)
'*
'*    パラメータ    : 変数名      ,IO  ,型                ,説明
'*      　         　udtXSDCB 　　  ,I   ,typ_XSDCB_Update  ,XSDCB更新用ﾃﾞｰﾀ
'*               　　sErrMsg　　　,O   ,String         　 ,エラーメッセージ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function CreateXSDCB(udtXSDCB As typ_XSDCB_Update, sErrMsg As String) As FUNCTION_RETURN
    Dim sSql        As String
    Dim sDBName     As String
    Dim rs          As OraDynaset   ' RecordSet
    Dim lngRecCnt   As Long         ' レコード数
    Dim dtmNowtime  As Date

    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_XSDCB_SQL.bas -- Function CreateXSDCB"
    sErrMsg = ""
    sDBName = "XSDCB"
    dtmNowtime = getSvrTime()       ' サーバーの時間を取得するように変更

'>>>>> AddNew-->INSERTに変更　2009/07/21　SSS.Marushita
    
    With udtXSDCB
        sSql = "INSERT INTO XSDCB ("
        sSql = sSql & " SXLIDCB" & vbLf      ' 1:SXLID
        sSql = sSql & ",KCNTCB" & vbLf       ' 2:工程連番
        sSql = sSql & ",XTALCB" & vbLf       ' 3:結晶番号
        sSql = sSql & ",INPOSCB" & vbLf      ' 4:結晶内開始位置
        sSql = sSql & ",LENCB" & vbLf        ' 5:長さ
        sSql = sSql & ",HINBCB" & vbLf       ' 6:品番
        sSql = sSql & ",REVNUMCB" & vbLf     ' 7:製品番号改訂番号
        sSql = sSql & ",FACTORYCB" & vbLf    ' 8:工場
        sSql = sSql & ",OPECB" & vbLf        ' 9:操業条件
        sSql = sSql & ",MAICB" & vbLf        '10:実枚数
        sSql = sSql & ",WSRMAICB" & vbLf     '11:WS洗後枚数
        sSql = sSql & ",WSNMAICB" & vbLf     '12:WS洗後欠落枚数
        sSql = sSql & ",WFCMAICB" & vbLf     '13:WFC受入枚数
        sSql = sSql & ",SXLRMAICB" & vbLf    '14:SXL指示（良品）
        sSql = sSql & ",SXLNMAICB" & vbLf    '15:SXL指示（不良）
        sSql = sSql & ",WFCNMAICB" & vbLf    '16:WFC内欠落枚数
        sSql = sSql & ",SXLEMAICB" & vbLf    '17:SXL確定枚数
        sSql = sSql & ",SRMAICB" & vbLf      '18:サンプル抜指示（良品）
        sSql = sSql & ",SNMAICB" & vbLf      '19:サンプル抜指示（不良）
        sSql = sSql & ",STMAICB" & vbLf      '20:サンプル枚数
        sSql = sSql & ",FURIMAICB" & vbLf    '21:振替枚数
        sSql = sSql & ",XTWORKCB" & vbLf     '22:製造工場
        sSql = sSql & ",WFWORKCB" & vbLf     '23:ウェーハ製造
        sSql = sSql & ",FURYCCB" & vbLf      '24:不良理由
        sSql = sSql & ",LSTCCB" & vbLf       '25:最終状態区分
        sSql = sSql & ",LUFRCCB" & vbLf      '26:格上コード
        sSql = sSql & ",LUFRBCB" & vbLf      '27:格上区分
        sSql = sSql & ",LDERCCB" & vbLf      '28:格下コード
        sSql = sSql & ",LDFRBCB" & vbLf      '29:格下区分
        sSql = sSql & ",HOLDCCB" & vbLf      '30:ホールドコード
        sSql = sSql & ",HOLDBCB" & vbLf      '31:ホールド区分
        sSql = sSql & ",EXKUBCB" & vbLf      '32:例外区分
        sSql = sSql & ",HENPKCB" & vbLf      '33:返品区分
        sSql = sSql & ",LIVKCB" & vbLf       '34:生死区分
        sSql = sSql & ",KANKCB" & vbLf       '35:完了区分
        sSql = sSql & ",NFCB" & vbLf         '36:入庫区分
        sSql = sSql & ",SAKJCB" & vbLf       '37:削除区分
        sSql = sSql & ",TDAYCB" & vbLf       '38:登録日付
        sSql = sSql & ",KDAYCB" & vbLf       '39:更新日付
        sSql = sSql & ",SUMITCB" & vbLf      '40:SUMIT送信フラグ
        sSql = sSql & ",SNDKCB" & vbLf       '41:送信フラグ
        sSql = sSql & ",SNDAYCB" & vbLf      '42:送信日付
        sSql = sSql & ",NEWKNTCB" & vbLf     '43:最終通過工程
        sSql = sSql & ",GNWKNTCB" & vbLf     '44:現在工程
        sSql = sSql & ",MOTHINCB" & vbLf     '45:振替品番(元）
        sSql = sSql & ",SHOLDCLSCB" & vbLf   '46:ホールド区分
        sSql = sSql & ",RLENCB" & vbLf       '47:理論長さ
        sSql = sSql & ",KBLKFLGCB" & vbLf    '48:関連ﾌﾞﾛｯｸﾌﾗｸﾞ
        sSql = sSql & ")"
        sSql = sSql & "VALUES ("

        ' 1:SXLID
        If .SXLIDCB <> "" Then
            sSql = sSql & " '" & .SXLIDCB & "'" & vbLf
        Else
            sSql = sSql & " '" & Space(13) & "'" & vbLf
        End If
               
        ' 2:工程連番
        If .KCNTCB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.KCNTCB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        ' 3:結晶番号
        If .XTALCB <> "" Then
            sSql = sSql & ",'" & .XTALCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(12) & "'" & vbLf
        End If
        
        ' 4:結晶内開始位置
        If .INPOSCB <> "" Then
            sSql = sSql & ",'" & .INPOSCB & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If
        
        ' 5:長さ
        If .LENCB <> "" Then
            sSql = sSql & ",'" & .LENCB & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        ' 6:品番
        If .HINBCB <> "" Then
            sSql = sSql & ",'" & .HINBCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(8) & "'" & vbLf
        End If
        
        ' 7:製品番号改訂番号
        If .REVNUMCB <> "" Then
            sSql = sSql & ",'" & .REVNUMCB & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        ' 8:工場
        If .FACTORYCB <> "" Then
            sSql = sSql & ",'" & .FACTORYCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 9:操業条件
        If .OPECB <> "" Then
            sSql = sSql & ",'" & .OPECB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(1) & "'" & vbLf
        End If

        '10:実枚数
        If .MAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.MAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '11:WS洗後枚数
        If .WSRMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.WSRMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '12:WS洗後欠落枚数
        If .WSNMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.WSNMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '13:WFC受入枚数
        If .WFCMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.WFCMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '14:SXL指示（良品）
        If .SXLRMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.SXLRMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '15:SXL指示（不良）
        If .SXLNMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.SXLNMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '16:WFC内欠落枚数
        If .WFCNMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.WFCNMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '17:SXL確定枚数
        If .SXLEMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.SXLEMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '18:サンプル抜指示（良品）
        If .SRMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.SRMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '19:サンプル抜指示（不良）
        If .SNMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.SNMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '20:サンプル枚数
        If .STMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.STMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '21:振替枚数
        If .FURIMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.FURIMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf  '未使用(NULL)
        End If

        '22:製造工場
        sSql = sSql & ",'" & FACTORYCD & "'" & vbLf     '42 固定

        '23:ウェーハ製造
        If .WFWORKCB <> "" Then
            sSql = sSql & ",'" & .WFWORKCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(2) & "'" & vbLf '未使用(ｽﾍﾟｰｽ)
        End If

        '24:不良理由
        If .FURYCCB <> "" Then
            sSql = sSql & ",'" & .FURYCCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(3) & "'" & vbLf '未使用(ｽﾍﾟｰｽ)
        End If

        '25:最終状態区分
        If .LSTCCB <> "" Then
            sSql = sSql & ",'" & .LSTCCB & "'" & vbLf
        Else
            sSql = sSql & ",'T'" & vbLf       '通常
        End If

        '26:格上コード
        If .LUFRCCB <> "" Then
            sSql = sSql & ",'" & .LUFRCCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(3) & "'" & vbLf '未使用(ｽﾍﾟｰｽ)
        End If

        '27:格上区分
        If .LUFRBCB <> "" Then
            sSql = sSql & ",'" & .LUFRBCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(1) & "'" & vbLf '未使用(ｽﾍﾟｰｽ)
        End If

        '28:格下コード
        If .LDERCCB <> "" Then
            sSql = sSql & ",'" & .LDERCCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(3) & "'" & vbLf '未使用(ｽﾍﾟｰｽ)
        End If

        '29:格下区分
        If .LDFRBCB <> "" Then
            sSql = sSql & ",'" & .LDFRBCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf        '通常
        End If

        '30:ホールドコード
        If .HOLDCCB <> "" Then
            sSql = sSql & ",'" & .HOLDCCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(3) & "'" & vbLf '未使用(ｽﾍﾟｰｽ)
        End If

        '31:ホールド区分
        If .HOLDBCB <> "" Then
            sSql = sSql & ",'" & .HOLDBCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf        '通常
        End If

        '32:例外区分
        If .EXKUBCB <> "" Then
            sSql = sSql & ",'" & .EXKUBCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(1) & "'" & vbLf
        End If

        '33:返品区分
        If .HENPKCB <> "" Then
            sSql = sSql & ",'" & .HENPKCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(1) & "'" & vbLf '未使用(ｽﾍﾟｰｽ)
        End If

        '34:生死区分
        If .LIVKCB <> "" Then
            sSql = sSql & ",'" & .LIVKCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf        '生ロット
        End If

        '35:完了区分
        If .KANKCB <> "" Then
            sSql = sSql & ",'" & .KANKCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf        '通常
        End If

        '36:入庫区分
        If .NFCB <> "" Then
            sSql = sSql & ",'" & .NFCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf        '固定
        End If

        '37:削除区分
        If .SAKJCB <> "" Then
            sSql = sSql & ",'" & .SAKJCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf         '固定
        End If

        '38:登録日付
        sSql = sSql & ",TO_DATE('" & Format$(dtmNowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf

        '39:更新日付
        sSql = sSql & ",TO_DATE('" & Format$(dtmNowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf

        '40:SUMIT送信フラグ
        If .SUMITCB <> "" Then
            sSql = sSql & ",'" & .SUMITCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf    '初期値(今回使用しない)
        End If

        '41:送信フラグ
        If .SNDKCB <> "" Then
            sSql = sSql & ",'" & .SNDKCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf    '初期値(今回使用しない)
        End If

        '42:送信日付
        sSql = sSql & ",NULL" & vbLf

        '43:最終通過工程
        If .NEWKNTCB <> "" Then
            sSql = sSql & ",'" & .NEWKNTCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(5) & "'" & vbLf
        End If
        
        '44:現在工程
        If .GNWKNTCB <> "" Then
            sSql = sSql & ",'" & .GNWKNTCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(5) & "'" & vbLf
        End If
        
        '45:振替品番(元）
        If .MOTHINCB <> "" Then
            sSql = sSql & ",'" & .MOTHINCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(8) & "'" & vbLf
        End If

        '46:ホールド区分
        If .SHOLDCLSCB <> "" Then
            sSql = sSql & ",'" & .SHOLDCLSCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf
        End If
        
        '47:理論長さ
        If .RLENCB <> "" Then
            sSql = sSql & ",'" & .RLENCB & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If
        
        '48:関連ブロックフラグ
        If .KBLKFLGCB <> "" And left(.KBLKFLGCB, 1) <> vbNullChar Then
            sSql = sSql & ",'" & .KBLKFLGCB & "'" & vbLf
        Else
            sSql = sSql & ",NULL" & vbLf
        End If

        sSql = sSql & ")" & vbLf
    
        'SQLを実行
        If OraDB.ExecuteSQL(sSql) < 1 Then
            GoTo proc_err
        End If

    End With

'<<<<< AddNew-->INSERTに変更　2009/07/21　SSS.Marushita

    Debug.Print sSql

    CreateXSDCB = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
'    gErr.Pop
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    sErrMsg = GetMsgStr("ENG11", "DB", sDBName)
    CreateXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************************
'*    関数名        : GetKCNTCB
'*
'*    処理概要      : 1.工程連番を取得する
'*
'*    パラメータ    : 変数名      ,IO  ,型               ,説明
'*      　         　sP_SXLID      ,I   ,String           ,SXLID
'*
'*    戻り値        : 工程連番
'*
'*******************************************************************************************
Public Function GetKCNTCB(sP_SXLID As String) As Integer
    Dim sSql    As String
    Dim rs      As OraDynaset

    sSql = "SELECT KCNTCB FROM XSDCB WHERE SXLIDCB = '" & sP_SXLID & "'"

    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("KCNTCB")) Then
        GetKCNTCB = 1
    Else
        GetKCNTCB = CInt(rs.Fields("KCNTCB")) + 1
    End If
End Function

'*******************************************************************************************
'*    関数名        : CheckSXLrecord
'*
'*    処理概要      : 1.該当するﾚｺｰﾄﾞ有無をﾁｪｯｸ(あれば長さを取得する)
'*
'*    パラメータ    : 変数名      ,IO  ,型               ,説明
'*      　         　sP_SXLID      ,I   ,String           ,SXLID
'*　　　　　　　　　 intP_Length     ,O   ,Integer          ,長さ
'*
'*    戻り値        : レコード数
'*
'*******************************************************************************************
Public Function CheckSXLrecord(sP_SXLID As String, intP_Length As Integer) As Integer
    Dim sSql    As String
    Dim rs      As OraDynaset

    sSql = "SELECT LENCB FROM XSDCB WHERE SXLIDCB = '" & sP_SXLID & "'"

    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("LENCB")) Then
        CheckSXLrecord = 0
    Else
        CheckSXLrecord = 1
        intP_Length = CInt(rs.Fields("LENCB"))
    End If
End Function

'*******************************************************************************************
'*    関数名        : DBDRV_CheckCodeXSDCB
'*
'*    処理概要      : 1.テーブル「XSDCB」の工程をチェックする(CW740/CW750/CW760/CW800)
'*                      (指定データの工程が引数で渡された工程と同じかチェックする)
'*
'*    パラメータ    : 変数名      ,IO  ,型               ,説明
'*                   sChkSXLID()  ,I   ,String           ,チェックレコード
'*                   sNowCode     ,I   ,String           ,チェック工程
'*                   sErrMsg      ,O   ,String           ,エラーメッセージ
'*
'*    戻り値        : レコード数
'*
'*******************************************************************************************
Public Function DBDRV_CheckCodeXSDCB(sChkSXLID() As String, sNowCode As String, sErrMsg As String) As FUNCTION_RETURN
    Dim sSql            As String             ' SQL全体
    Dim rs              As OraDynaset         ' RecordSet
    Dim udtReadXSDCB()  As typ_XSDCB_Update   ' 取得データ
    Dim lngLoopCnt      As Long
    Dim i               As Long
    Dim j               As Long
    Dim sDBName         As String

    ' エラーハンドラの設定
    On Error GoTo proc_err
'    gErr.Push "s_XSDCB_SQL.bas -- Function DBDRV_CheckCodeXSDCB"
    sErrMsg = ""
    sDBName = "XSDCB"

    i = 0
    ReDim udtReadXSDCB(0)
    For lngLoopCnt = 1 To UBound(sChkSXLID)
        If lngLoopCnt = 1 Or _
           (lngLoopCnt > 1 And sChkSXLID(lngLoopCnt) <> sChkSXLID(lngLoopCnt - 1)) Then

            ' SQLを組み立てる
            sSql = ""
            sSql = sSql & " SELECT"
            sSql = sSql & "   SXLIDCB"
            sSql = sSql & "  ,GNWKNTCB"
            sSql = sSql & " FROM"
            sSql = sSql & "   XSDCB"
            sSql = sSql & " WHERE SXLIDCB = '" & sChkSXLID(lngLoopCnt) & "'"
            sSql = sSql & "   AND LIVKCB = '0' "

            ' データを抽出する
            Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
            If rs Is Nothing Then
                DBDRV_CheckCodeXSDCB = FUNCTION_RETURN_FAILURE
                Exit Function
            End If

            i = i + 1
            ReDim Preserve udtReadXSDCB(i)

            ' 抽出結果を格納する
            If IsNull(rs.Fields("SXLIDCB")) = False Then udtReadXSDCB(i).SXLIDCB = rs.Fields("SXLIDCB")
            If IsNull(rs.Fields("GNWKNTCB")) = False Then udtReadXSDCB(i).GNWKNTCB = rs.Fields("GNWKNTCB")
            rs.Close
        End If
    Next lngLoopCnt

    For j = 1 To i
        ' 現在工程が同じかチェックする(CST02は除く)
        If Trim(udtReadXSDCB(j).GNWKNTCB) <> Trim(sNowCode) And _
           Trim(udtReadXSDCB(j).GNWKNTCB) <> "CST02" Then
            ' 現在工程が違う = SXLがすでに動いている場合、エラー終了
            sErrMsg = GetMsgStr("EBLK6")
            DBDRV_CheckCodeXSDCB = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    Next j

    DBDRV_CheckCodeXSDCB = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
'    gErr.Pop
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    sErrMsg = GetMsgStr("ENG11", "DB", sDBName)
    DBDRV_CheckCodeXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function
