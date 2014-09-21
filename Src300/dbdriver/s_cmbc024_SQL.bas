Attribute VB_Name = "s_cmbc024_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DBアクセス関数
' 定義内容: TBCMJ008 (ＢＭＤ実績)
' 参照　　: 060211_結晶検査
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmjc001e_Disp
   ' CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
   ' TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    factory As String * 1           ' 工場
    opecond As String * 1           ' 操業条件
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    GOUKI As String * 3             ' 号機
    MEASMETH As String * 1          ' 測定方法
    MEASSPOT As Integer             ' 測定点
    MAG As String * 4               ' 倍率
    HTPRC As String * 2             ' 熱処理方法
    KKSP As String * 3              ' 結晶欠陥測定位置
    KKSET As String * 3             ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    MEAS1 As Double                 ' 測定値１
    MEAS2 As Double                 ' 測定値２
    MEAS3 As Double                 ' 測定値３
    MEAS4 As Double                 ' 測定値４
    MEAS5 As Double                 ' 測定値５
    MEASMIN As Double               ' MIN
    MEASMAX As Double               ' MAX
    MEASAVE As Double               ' AVE
    BMDMNBUNP As Double             ' BMD面内分布
   ' TSTAFFID As String * 8          ' 登録社員ID
   ' KSTAFFID As String * 8          ' 更新社員ID
   ' UPDDATE As Date                 ' 更新日付
   ' SENDFLAG As String * 1          ' 送信フラグ
   ' SENDDATE As Date                ' 送信日付

End Type



'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMJ008」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_cmjc001e_Disp ,抽出レコード
'          :SPLNUMs()     ,I  ,Integer      ,抽出条件配列(サンプルNo)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/20作成　長野
'          :2003/04/02    　yakimura  項目追加対応
Public Function DBDRV_Getcmjc001e_Disp(records() As typ_cmjc001e_Disp, SMPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim sqlWhere As String  'SQLのWHERE部分
Dim sqlGroup As String  'SQLのGROUP部分
Dim sqlOrder As String  'SQLのOrder部分
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    DBDRV_Getcmjc001e_Disp = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001e_SQL.bas -- Function DBDRV_Getcmjc001e_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN, MEASMAX, MEASAVE, BMDMNBUNP "
    sqlBase = sqlBase & "From TBCMJ008"
    ''抽出条件(ｻﾝﾌﾟﾙNO)の取り出し
    sqlWhere = "Where SMPLNO in ("
    For i = 1 To UBound(SMPLNUMs)
        sqlWhere = sqlWhere & "'" & SMPLNUMs(i) & "'"
        If i < UBound(SMPLNUMs) Then
            sqlWhere = sqlWhere & ", "
        End If
    Next
    sqlWhere = sqlWhere & ") "
    sqlGroup = "GROUP BY CRYNUM, POSITION, SMPKBN, TRANCOND "
    sqlOrder = "ORDER BY POSITION"
    sql = sqlBase & sqlWhere & sqlGroup & sqlOrder
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_Getcmjc001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .POSITION = rs("POSITION")      ' 位置
            .SMPKBN = rs("SMPKBN")          ' サンプル区分
            .TRANCOND = rs("TRANCOND")      ' 処理条件
            .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
            .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
            .hinban = rs("HINBAN")          ' 品番
            .REVNUM = rs("REVNUM")          ' 製品番号改訂番号
            .factory = rs("FACTORY")        ' 工場
            .opecond = rs("OPECOND")        ' 操業条件
            .KRPROCCD = rs("KRPROCCD")      ' 管理工程コード
            .PROCCODE = rs("PROCCODE")      ' 工程コード
            .GOUKI = rs("GOUKI")            ' 号機
            .MEASMETH = rs("MEASMETH")      ' 測定方法
            .MEASSPOT = rs("MEASSPOT")      ' 測定点
            .MAG = rs("MAG")                ' 倍率
            .HTPRC = rs("HTPRC")            ' 熱処理方法
            .KKSP = rs("KKSP")              ' 結晶欠陥測定位置
            .KKSET = rs("KKSET")            ' 結晶欠陥測定条件＋選択ET代
            .MEAS1 = rs("MEAS1")            ' 測定値１
            .MEAS2 = rs("MEAS2")            ' 測定値２
            .MEAS3 = rs("MEAS3")            ' 測定値３
            .MEAS4 = rs("MEAS4")            ' 測定値４
            .MEAS5 = rs("MEAS5")            ' 測定値５
            .MEASMIN = rs("MEASMIN")            ' MIN
            .MEASMAX = rs("MEASMAX")            ' MAX
            .MEASAVE = rs("MEASAVE")            ' AVE
'OSF，BMD項目追加対応  2002.04.02 yakimura
            If rs("BMDMNBUNP") <> vbNullString Then
               .BMDMNBUNP = rs("BMDMNBUNP")     ' ＢＭＤ面内分布
            Else
               .BMDMNBUNP = 0
            End If
'OSF，BMD項目追加対応  2002.04.02 yakimura
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001e_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :引数で渡されたレコードをTBCMJ008に追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :record        ,I  ,typ_cmjc001e_Disp ,抽出レコード
'          :CRYNUM        ,I  ,String       ,結晶番号
'          :TSTAFFID      ,I  ,String       ,登録社員ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :処理回数はテーブル上の最大値+1とする。
'履歴      :2001/06/22(Fri)作成　長野
'          :2003/04/02    　yakimura  項目追加対応

Public Function DBDRV_Getcmjc001e_Exec(record As typ_cmjc001e_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQLベース部分
Dim sqlWhere As String  'SQLWhere部分
Dim sqlGroup As String  'SQLGroup部分

'    CRYNUM             結晶番号　⇒引数
'    TRANCNT         　 処理回数　⇒最大
'   TSTAFFID            登録社員ID　⇒引数
 '   REGDATE 　　　     登録日付　⇒SYSDATE
 '   KSTAFFID           更新社員ID　⇒" "
 '   UPDDATE            更新日付　⇒SYSDATE
 '   SENDFLAG           送信フラグ　⇒"0"
 '   SENDDATE           送信日付　⇒SYSDATE

    DBDRV_Getcmjc001e_Exec = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001e_SQL.bas -- Function DBDRV_Getcmjc001e_Exec"

    sqlBase = "Insert into TBCMJ008 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, " & _
              "KRPROCCD, PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN, MEASMAX, MEASAVE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, BMDMNBUNP) " & vbCrLf
    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "',  '" & record.hinban & "', " & record.REVNUM & ",'" & record.factory & "', '" & record.opecond & "', '" & _
               record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', '" & record.MEASMETH & "', " & record.MEASSPOT & ", '" & record.MAG & "', '" & _
               record.HTPRC & "', '" & record.KKSP & "', '" & record.KKSET & "', " & record.MEAS1 & ", " & record.MEAS2 & ", " & record.MEAS3 & ", " & record.MEAS4 & ", " & _
               record.MEAS5 & ", " & record.MEASMIN & ", " & record.MEASMAX & ", " & record.MEASAVE & ", '" & TSTAFFID & "', SYSDATE,' ', SYSDATE, '0', SYSDATE, " & record.BMDMNBUNP & " from TBCMJ008 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') " & vbCrLf
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
            
Debug.Print sql
    ''SQLの実行
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001e_Exec = FUNCTION_RETURN_SUCCESS



proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function



'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMJ008」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ008 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
'          :2003/04/02    　yakimura  項目追加対応
Public Function DBDRV_GetTBCMJ008(records() As typ_TBCMJ008, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN," & _
              " MEASMAX, MEASAVE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, BMDMNBUNP "
    sqlBase = sqlBase & "From TBCMJ008"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ008 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .POSITION = rs("POSITION")       ' 位置
            .SMPKBN = rs("SMPKBN")           ' サンプル区分
            .TRANCOND = rs("TRANCOND")       ' 処理条件
            .TRANCNT = rs("TRANCNT")         ' 処理回数
            .SMPLNO = rs("SMPLNO")           ' サンプルＮｏ
            .SMPLUMU = rs("SMPLUMU")         ' サンプル有無
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .GOUKI = rs("GOUKI")             ' 号機
            .MEASMETH = rs("MEASMETH")       ' 測定方法
            .MEASSPOT = rs("MEASSPOT")       ' 測定点
            .MAG = rs("MAG")                 ' 倍率
            .HTPRC = rs("HTPRC")             ' 熱処理方法
            .KKSP = rs("KKSP")               ' 結晶欠陥測定位置
            .KKSET = rs("KKSET")             ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
            .MEAS1 = rs("MEAS1")             ' 測定値１
            .MEAS2 = rs("MEAS2")             ' 測定値２
            .MEAS3 = rs("MEAS3")             ' 測定値３
            .MEAS4 = rs("MEAS4")             ' 測定値４
            .MEAS5 = rs("MEAS5")             ' 測定値５
            .MEASMIN = rs("MEASMIN")         ' MIN
            .MEASMAX = rs("MEASMAX")         ' MAX
            .MEASAVE = rs("MEASAVE")         ' AVE
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
'OSF，BMD項目追加対応  2002.04.02 yakimura
            If rs("BMDMNBUNP") <> vbNullString Then
               .BMDMNBUNP = rs("BMDMNBUNP")  ' ＢＭＤ面内分布
            Else
               .BMDMNBUNP = 0
            End If
'OSF，BMD項目追加対応  2002.04.02 yakimura
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMJ008 = FUNCTION_RETURN_SUCCESS
End Function



