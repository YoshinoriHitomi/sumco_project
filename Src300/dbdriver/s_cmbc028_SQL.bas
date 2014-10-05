Attribute VB_Name = "s_cmbc028_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DBアクセス関数
' 定義内容: TBCMJ001 (EPD実績)
' 参照　　: 060211_結晶検査
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmjc001i_Disp
  '  CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
  '  TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    factory As String * 1           ' 工場
    opecond As String * 1           ' 操業条件
    GOUKI As String * 3             ' 号機
    MEASURE As Integer              ' 測定値
  '  TSTAFFID As String * 8          ' 登録社員ID
  '  REGDATE As Date                 ' 登録日付
  '  KSTAFFID As String * 8          ' 更新社員ID
  '  UPDDATE As Date                 ' 更新日付
  '  SENDFLAG As String * 1          ' 送信フラグ
  '  SENDDATE As Date                ' 送信日付
End Type

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMJ001」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_cmjc001i_Disp ,抽出レコード
'          :SPLNUMs()     ,I  ,Integer      ,抽出条件配列(サンプルNo)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/20作成　長野
Public Function DBDRV_Getcmjc001i_Disp(records() As typ_cmjc001i_Disp, SMPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim sqlWhere As String  'SQLのWHERE部分
Dim sqlGroup As String  'SQLのGROUP部分
Dim sqlOrder As String  'SQLのOrder部分
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    DBDRV_Getcmjc001i_Disp = FUNCTION_RETURN_FAILURE
    
    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001i_SQL.bas -- Function DBDRV_Getcmjc001i_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
              " OPECOND, GOUKI, MEASURE "
    sqlBase = sqlBase & "From TBCMJ001"
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
        DBDRV_Getcmjc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .POSITION = rs("POSITION")       ' 位置
            .SMPKBN = rs("SMPKBN")           ' サンプル区分
            .TRANCOND = rs("TRANCOND")       ' 処理条件
            .SMPLNO = rs("SMPLNO")           ' サンプルＮｏ
            .SMPLUMU = rs("SMPLUMU")         ' サンプル有無
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .GOUKI = rs("GOUKI")             ' 号機
            .MEASURE = rs("MEASURE")         ' 測定値
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001i_Disp = FUNCTION_RETURN_SUCCESS

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

'概要      :引数で渡されたレコードをTBCMJ001に追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :record        ,I  ,typ_cmjc001i_Disp ,抽出レコード
'          :CRYNUM        ,I  ,String       ,結晶番号
'          :TSTAFFID      ,I  ,String       ,登録社員ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :処理回数はテーブル上の最大値+1とする。
'履歴      :2001/06/25(Mon)作成　長野

Public Function DBDRV_Getcmjc001i_Exec(record As typ_cmjc001i_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
Dim sql         As String   'SQL
Dim sqlBase     As String   'SQLベース部分
Dim sqlWhere    As String  'SQLWhere部分
Dim sqlGroup    As String  'SQLGroup部分


'    CRYNUM             結晶番号　⇒引数
'    TRANCNT         　 処理回数　⇒最大
'   TSTAFFID            登録社員ID　⇒引数
 '   REGDATE 　　　     登録日付　⇒SYSDATE
 '   KSTAFFID           更新社員ID　⇒" "
 '   UPDDATE            更新日付　⇒SYSDATE
 '   SENDFLAG           送信フラグ　⇒"0"
 '   SENDDATE           送信日付　⇒SYSDATE

    DBDRV_Getcmjc001i_Exec = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001i_SQL.bas -- Function DBDRV_Getcmjc001i_Exec"

    sqlBase = "Insert into TBCMJ001 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, " & _
              "HINBAN, REVNUM, FACTORY, OPECOND, GOUKI, MEASURE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.hinban & "', " & record.REVNUM & ", '" & _
               record.factory & "', '" & record.opecond & "', '" & record.GOUKI & "', " & record.MEASURE & ", '" & _
               TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ001 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
            
    ''SQLの実行
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001i_Exec = FUNCTION_RETURN_SUCCESS


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

'概要      :テーブル「TBCMJ001」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ001 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCMJ001(records() As typ_TBCMJ001, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
              " OPECOND, GOUKI, MEASURE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ001"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ001 = FUNCTION_RETURN_FAILURE
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
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .GOUKI = rs("GOUKI")             ' 号機
            .MEASURE = rs("MEASURE")         ' 測定値
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMJ001 = FUNCTION_RETURN_SUCCESS
End Function

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME036」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME036 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
'          :06/04/11 ooba　関数名変更 <DBDRV_GetTBCME036> ⇒ <DBDRV_GetTBCME036_cmbc028>
Public Function DBDRV_GetTBCME036_cmbc028(records() As typ_TBCME036, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, EPDSETCH, EPDUP, CUTUNIT, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO," & _
              " STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME036"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME036_cmbc028 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .hinban = rs("HINBAN")           ' 品番
            .mnorevno = rs("MNOREVNO")       ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .EPDSETCH = rs("EPDSETCH")       ' EPD　選択エッチ
            .EPDUP = fncNullCheck(rs("EPDUP"))             ' EPD　上限
            .CUTUNIT = fncNullCheck(rs("CUTUNIT"))         ' カット単位
            .IFKBN = rs("IFKBN")             ' Ｉ／Ｆ区分
            .SYORIKBN = rs("SYORIKBN")       ' 処理区分
            .SPECRRNO = rs("SPECRRNO")       ' 仕様登録依頼番号
            .SXLMCNO = rs("SXLMCNO")         ' ＳＸＬ製作条件番号
            .WFMCNO = rs("WFMCNO")           ' ＷＦ製作条件番号
            .StaffID = rs("STAFFID")         ' 社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME036_cmbc028 = FUNCTION_RETURN_SUCCESS
End Function

