Attribute VB_Name = "s_cmbc021_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DBアクセス関数
' 定義内容: TBCMJ003 (Ｏｉ実績)
' 参照　　: 060211_結晶検査
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmjc001b_Disp
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
    OIMEAS1 As Double               ' Ｏｉ測定値１
    OIMEAS2 As Double               ' Ｏｉ測定値２
    OIMEAS3 As Double               ' Ｏｉ測定値３
    OIMEAS4 As Double               ' Ｏｉ測定値４
    OIMEAS5 As Double               ' Ｏｉ測定値５
    ORGRES As Double                ' ＯＲＧ結果
    SETDTM As Date                  ' 設定日時
    EFFECTTM As Integer             ' 有効時間
    FTIRMETH As String              ' ＦＴＩＲ相関式
    YCOEF As Double                 ' ＦＴＩＲ換算式（Ｙ切片）
    XCOEF As Double                 ' ＦＴＩＲ換算式（Ｘ係数）
    AVE As Double                   ' ＡＶＥ
    SIGMA As Double                 ' σ（シグマ）
    FTIRCONV As Double              ' ＦＴＩＲ換算
    INSPECTWAY As String * 2        ' 検査方法
    JudgData As Double
   ' TSTAFFID As String * 8          ' 登録社員ID
   ' REGDATE As Date                 ' 登録日付
   ' KSTAFFID As String * 8          ' 更新社員ID
   ' UPDDATE As Date                 ' 更新日付
   ' SENDFLAG As String * 1          ' 送信フラグ
   ' SENDDATE As Date                ' 送信日付
End Type


'                                     2001/06/20
'================================================
' DBアクセス関数
' 定義内容: TBCMJ004 (Cs実績)
' 参照　　: 060211_結晶検査
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmjc001b_Disp2
'    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
'    TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    factory As String * 1           ' 工場
    opecond As String * 1           ' 操業条件
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    GOUKI As String * 3             ' 号機
    CSMEAS As Double                ' Cs実測値
    PRE70P As Double                ' ７０％推定値
    INSPECTWAY As String * 2        ' 検査方法
 '   TSTAFFID As String * 8          ' 登録社員ID
 '   REGDATE As Date                 ' 登録日付
 '   KSTAFFID As String * 8          ' 更新社員ID
 '   UPDDATE As Date                 ' 更新日付
 '   SENDFLAG As String * 1          ' 送信フラグ
 '   SENDDATE As Date                ' 送信日付
End Type
'                                     2001/06/20
'================================================
' DBアクセス関数
' 定義内容: TBCMJ003 (Ｏｉ実績)
' 参照　　: 060211_結晶検査
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmjc001c_Disp
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
    OIMEAS1 As Double               ' Ｏｉ測定値１
    OIMEAS2 As Double               ' Ｏｉ測定値２
    OIMEAS3 As Double               ' Ｏｉ測定値３
    OIMEAS4 As Double               ' Ｏｉ測定値４
    OIMEAS5 As Double               ' Ｏｉ測定値５
    ORGRES As Double                ' ＯＲＧ結果
    SETDTM As Date                  ' 設定日時
    EFFECTTM As Integer             ' 有効時間
    FTIRMETH As String              ' ＦＴＩＲ相関式
    YCOEF As Double                 ' ＦＴＩＲ換算式（Ｙ切片）
    XCOEF As Double                 ' ＦＴＩＲ換算式（Ｘ係数）
    AVE As Double                   ' ＡＶＥ
    SIGMA As Double                 ' σ（シグマ）
    FTIRCONV As Double              ' ＦＴＩＲ換算
    INSPECTWAY As String * 2        ' 検査方法
    JudgData As Double              ' 検索対象値
   ' TSTAFFID As String * 8          ' 登録社員ID
   ' REGDATE As Date                 ' 登録日付
   ' KSTAFFID As String * 8          ' 更新社員ID
   ' UPDDATE As Date                 ' 更新日付
   ' SENDFLAG As String * 1          ' 送信フラグ
   ' SENDDATE As Date                ' 送信日付
End Type


'(2002/07 s_cmzcF_TBCME019_SQL.basより移動)
'フィールド名検索用
Dim fldNames() As String    '現rsに含まれるフィールド名保持配列
Dim fldCnt As Integer       '現rsに含まれるフィールド数



'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMJ003」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_cmjc001b_Disp ,抽出レコード
'          :SPLNUMs()     ,I  ,Integer      ,抽出条件配列(サンプルNo)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/20(Wed)作成　長野
Public Function DBDRV_Getcmjc001b_Disp(records() As typ_cmjc001b_Disp, SPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim sqlWhere As String  'SQLのWHERE部分
Dim sqlGroup As String  'SQLのGROUP部分
Dim sqlOrder As String  'SQLのOrder部分
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long           'ループカウント

    DBDRV_Getcmjc001b_Disp = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001b_SQL.bas -- Function DBDRV_Getcmjc001b_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, MAX(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA "
    sqlBase = sqlBase & "From TBCMJ003"
    ''抽出条件(ｻﾝﾌﾟﾙNO)の取り出し
    sqlWhere = "Where SMPLNO in ("
    For i = 1 To UBound(SPLNUMs)
        sqlWhere = sqlWhere & "'" & SPLNUMs(i) & "'"
        If i < UBound(SPLNUMs) Then
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
        DBDRV_Getcmjc001b_Disp = FUNCTION_RETURN_FAILURE
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
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .GOUKI = rs("GOUKI")             ' 号機
            .OIMEAS1 = rs("OIMEAS1")         ' Ｏｉ測定値１
            .OIMEAS2 = rs("OIMEAS2")         ' Ｏｉ測定値２
            .OIMEAS3 = rs("OIMEAS3")         ' Ｏｉ測定値３
            .OIMEAS4 = rs("OIMEAS4")         ' Ｏｉ測定値４
            .OIMEAS5 = rs("OIMEAS5")         ' Ｏｉ測定値５
            .ORGRES = rs("ORGRES")           ' ＯＲＧ結果
            .SETDTM = rs("SETDTM")           ' 設定日時
            .EFFECTTM = rs("EFFECTTM")       ' 有効時間
            .FTIRMETH = rs("FTIRMETH")       ' ＦＴＩＲ相関式
            .YCOEF = rs("YCOEF")             ' ＦＴＩＲ換算式（Ｙ切片）
            .XCOEF = rs("XCOEF")             ' ＦＴＩＲ換算式（Ｘ係数）
            .AVE = rs("AVE")                 ' ＡＶＥ
            .SIGMA = rs("SIGMA")             ' σ（シグマ）
            .FTIRCONV = rs("FTIRCONV")       ' ＦＴＩＲ換算
            .INSPECTWAY = rs("INSPECTWAY")   ' 検査方法
            .JudgData = rs("JUDGDATA")       ' 検索対象値
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001b_Disp = FUNCTION_RETURN_SUCCESS


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
'概要      :テーブル「TBCMJ004」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_cmjc001b_Disp2 ,抽出レコード
'          :SPLNUMs()     ,I  ,Integer      ,抽出条件配列(サンプルNo)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/20作成　長野
Public Function DBDRV_Getcmjc001b_Disp2(records() As typ_cmjc001b_Disp2, SPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim sqlWhere As String  'SQLのWHERE部分
Dim sqlGroup As String  'SQLのGROUP部分
Dim sqlOrder As String  'SQLのOrder部分
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    DBDRV_Getcmjc001b_Disp2 = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001b_SQL.bas -- Function DBDRV_Getcmjc001b_Disp2"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, CSMEAS, PRE70P, INSPECTWAY "
    sqlBase = sqlBase & "From TBCMJ004"
    ''抽出条件(ｻﾝﾌﾟﾙNO)の取り出し
    sqlWhere = "Where SMPLNO in ("
    For i = 1 To UBound(SPLNUMs)
        sqlWhere = sqlWhere & "'" & SPLNUMs(i) & "'"
        If i < UBound(SPLNUMs) Then
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
        DBDRV_Getcmjc001b_Disp2 = FUNCTION_RETURN_FAILURE
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
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .GOUKI = rs("GOUKI")             ' 号機
            .CSMEAS = rs("CSMEAS")           ' Cs実測値
            .PRE70P = rs("PRE70P")           ' ７０％推定値
            .INSPECTWAY = rs("INSPECTWAY")   ' 検査方法
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001b_Disp2 = FUNCTION_RETURN_SUCCESS

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

'概要      :引数で渡されたレコードをTBCMJ003に追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :record        ,I  ,typ_TBCMJ003 ,抽出レコード
'          :CRYNUM        ,I  ,String       ,結晶番号
'          :TSTAFFID      ,I  ,String       ,登録社員ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :新規追加の際、処理回数はテーブル上の最大値+1とする。
'履歴      :2001/06/20(Wed)作成　長野

Public Function DBDRV_Getcmjc001b_Exec(record As typ_cmjc001b_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQLベース部分
Dim sqlWhere As String  'SQLWhere部分
Dim sqlGroup As String  'SQLGroup部分
Dim SetDate As Variant  '設定日時
Dim rs As OraDynaset    'OracleDynaset

'    CRYNUM             結晶番号　⇒引数
'    TRANCNT         　 処理回数　⇒最大
'   TSTAFFID            登録社員ID　⇒引数
 '   REGDATE 　　　     登録日付　⇒SYSDATE
 '   KSTAFFID           更新社員ID　⇒" "
 '   UPDDATE            更新日付　⇒SYSDATE
 '   SENDFLAG           送信フラグ　⇒"0"
 '   SENDDATE           送信日付　⇒SYSDATE
    
    DBDRV_Getcmjc001b_Exec = FUNCTION_RETURN_FAILURE
    
    ''最大カウント取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001b_SQL.bas -- Function DBDRV_Getcmjc001b_Exec"

    sqlBase = "select nvl(MAX(TRANCNT),0) + 1 as w_TRANCNT from TBCMJ003 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
    ''設定日時フォーマット処理
    SetDate = Format$(record.SETDTM, "yyyy-mm-dd hh:mm:ss")

    ''SQLを組み立てる
    sql = "Insert into TBCMJ003 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
    sql = sql & "Values( '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', " & rs!w_TRANCNT & ", " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', " & record.OIMEAS1 & ", " & _
               record.OIMEAS2 & ", " & record.OIMEAS3 & ", " & record.OIMEAS4 & ", " & record.OIMEAS5 & ", " & record.ORGRES & ", " & _
               "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss'), " & record.EFFECTTM & ", '" & record.FTIRMETH & "', " & record.YCOEF & ", " & record.XCOEF & ", " & _
               record.AVE & ", " & record.SIGMA & ", " & record.FTIRCONV & ", '" & record.INSPECTWAY & "', " & record.JudgData & ", '" & TSTAFFID & "', " & _
               "SYSDATE, ' ', SYSDATE, '0', SYSDATE) "
''''    '' OI_NULL対応　2005/03/07 TUKU START 山口殿からの依頼により変更中止--------------------------------------------------------------------
''''    sql = "Insert into TBCMJ003 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
''''              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
''''              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
''''    sql = sql & "Values( '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', " & rs!w_TRANCNT & ", " & _
''''               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
''''               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', "
''''                If (record.OIMEAS1 = -1) Then sql = sql & " NULL , " Else sql = sql & record.OIMEAS1 & ", "    'OI測定値1
''''                If (record.OIMEAS2 = -1) Then sql = sql & " NULL , " Else sql = sql & record.OIMEAS2 & ", "    'OI測定値2
''''                If (record.OIMEAS3 = -1) Then sql = sql & " NULL , " Else sql = sql & record.OIMEAS3 & ", "    'OI測定値2
''''                If (record.OIMEAS4 = -1) Then sql = sql & " NULL , " Else sql = sql & record.OIMEAS4 & ", "    'OI測定値3
''''                If (record.OIMEAS5 = -1) Then sql = sql & " NULL , " Else sql = sql & record.OIMEAS5 & ", "    'OI測定値4
''''                If (record.ORGRES = -1) Then sql = sql & " NULL , " Else sql = sql & record.ORGRES & ", "      'ORG
''''    sql = sql & "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss'), " & record.EFFECTTM & ", '" & record.FTIRMETH & "', " & record.YCOEF & ", " & record.XCOEF & ", " & _
''''               record.AVE & ", " & record.SIGMA & ", " & record.FTIRCONV & ", '" & record.INSPECTWAY & "', " & record.JudgData & ", '" & TSTAFFID & "', " & _
''''               "SYSDATE, ' ', SYSDATE, '0', SYSDATE) "
''''    '' OI_NULL対応　2005/03/07 TUKU END   --------------------------------------------------------------------

    ''SQLの実行
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001b_Exec = FUNCTION_RETURN_SUCCESS
    

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

'概要      :引数で渡されたレコードをTBCMJ004に追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :record        ,I  ,typ_TBCMJ004 ,抽出レコード
'          :CRYNUM        ,I  ,String       ,結晶番号
'          :TSTAFFID      ,I  ,String       ,登録社員ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :新規追加の際、処理回数はテーブル上の最大値+1とする。
'履歴      :2001/06/22(Fri)作成　長野

Public Function DBDRV_Getcmjc001b_Exec2(record As typ_cmjc001b_Disp2, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQLベース部分
Dim sqlWhere As String  'SQLWhere部分
Dim sqlGroup As String  'SQLGroup部分

'    CRYNUM             結晶番号　⇒引数
'    TRANCNT         　 処理回数　⇒最大
'    TSTAFFID           登録社員ID　⇒引数
'    REGDATE 　　　     登録日付　⇒SYSDATE
'    KSTAFFID           更新社員ID　⇒””
'    UPDDATE            更新日付　⇒SYSDATE
'    SENDFLAG           送信フラグ　⇒"０"
'    SENDDATE           送信日付　⇒SYSDATE
    
    DBDRV_Getcmjc001b_Exec2 = FUNCTION_RETURN_FAILURE
    
    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001b_SQL.bas -- Function DBDRV_Getcmjc001b_Exec2"

    sqlBase = "Insert into TBCMJ004 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, CSMEAS, PRE70P, INSPECTWAY, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', " & record.CSMEAS & ", " & _
               record.PRE70P & ", '" & record.INSPECTWAY & "', '" & TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ004 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
''''    '' OI_NULL対応　2005/03/07 TUKU START 山口殿からの依頼で変更中止--------------------------------------------------------------------
''''    sqlBase = "Insert into TBCMJ004 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
''''              " PROCCODE, GOUKI, CSMEAS, PRE70P, INSPECTWAY, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
''''    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
''''               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
''''               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', "
''''                If (record.CSMEAS = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.CSMEAS & ", "    'CS測定値
''''                If (record.PRE70P = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.PRE70P & ", "    'CS70%
''''    sqlBase = sqlBase & " '" & record.INSPECTWAY & "', '" & TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ004 "
''''    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'''''    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
''''    sql = sqlBase & sqlWhere & sqlGroup
''''    '' OI_NULL対応　2005/03/07 TUKU END   --------------------------------------------------------------------
            
    ''SQLの実行
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001b_Exec2 = FUNCTION_RETURN_SUCCESS



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





'概要      :引数のフィールド名がfldNames()配列に含まれているかどうかの判定。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :fldName       ,I  ,typ_TBCME018 ,抽出レコード
'          :戻り値        ,O  ,Boolean      ,True:在り／False：無し
'説明      :
'履歴      :2001/06/27作成　野村 (2002/07 s_cmzcF_TBCME019_SQL.basより移動)

Private Function fldNameExist(fldName As String) As Boolean
    Dim sql         As String           'SQL全体
    Dim i As Integer                    'ﾙｰﾌﾟｶｳﾝﾄ


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME019_SQL.bas -- Function fldNameExist"

    fldNameExist = False                'ｴﾗｰｽﾃｰﾀｽ（初期値）ｾｯﾄ
    
    For i = 1 To fldCnt                 'ﾌｨｰﾙﾄﾞ数分ﾙｰﾌﾟ
        If fldName = fldNames(i) Then   '引数のﾌｨｰﾙﾄﾞ名と一致するものがあった場合
            fldNameExist = True         '正常ｽﾃｰﾀｽｾｯﾄ
            Exit For                    'ﾙｰﾌﾟを抜ける
        End If
    Next
    

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

'概要      :テーブル「TBCMJ003」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ003 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCMJ003(records() As typ_TBCMJ003, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ003"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ003 = FUNCTION_RETURN_FAILURE
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
''''            .OIMEAS1 = rs("OIMEAS1")         ' Ｏｉ測定値１
''''            .OIMEAS2 = rs("OIMEAS2")         ' Ｏｉ測定値２
''''            .OIMEAS3 = rs("OIMEAS3")         ' Ｏｉ測定値３
''''            .OIMEAS4 = rs("OIMEAS4")         ' Ｏｉ測定値４
''''            .OIMEAS5 = rs("OIMEAS5")         ' Ｏｉ測定値５
''''            .ORGRES = rs("ORGRES")           ' ＯＲＧ結果
'OI_NULL対応　2005/03/03 TUKU START --------------------------------------------------
            If IsNull(rs("OIMEAS1")) = False Then .OIMEAS1 = rs("OIMEAS1") Else .OIMEAS1 = -1  'Ｏｉ測定値1
            If IsNull(rs("OIMEAS2")) = False Then .OIMEAS2 = rs("OIMEAS2") Else .OIMEAS2 = -1  'Ｏｉ測定値2
            If IsNull(rs("OIMEAS3")) = False Then .OIMEAS3 = rs("OIMEAS3") Else .OIMEAS3 = -1  'Ｏｉ測定値3
            If IsNull(rs("OIMEAS4")) = False Then .OIMEAS4 = rs("OIMEAS4") Else .OIMEAS4 = -1  'Ｏｉ測定値4
            If IsNull(rs("OIMEAS5")) = False Then .OIMEAS5 = rs("OIMEAS5") Else .OIMEAS5 = -1  'Ｏｉ測定値5
            If IsNull(rs("ORGRES")) = False Then .ORGRES = rs("ORGRES") Else .ORGRES = -1    ' ＯＲＧ結果
'OI_NULL対応　2005/03/03 TUKU END   --------------------------------------------------
            
            .SETDTM = rs("SETDTM")           ' 設定日時
            .EFFECTTM = rs("EFFECTTM")       ' 有効時間
            .FTIRMETH = rs("FTIRMETH")       ' ＦＴＩＲ相関式
            .YCOEF = rs("YCOEF")             ' ＦＴＩＲ換算式（Ｙ切片）
            .XCOEF = rs("XCOEF")             ' ＦＴＩＲ換算式（Ｘ係数）
            .AVE = rs("AVE")                 ' ＡＶＥ
            .SIGMA = rs("SIGMA")             ' σ（シグマ）
            .FTIRCONV = rs("FTIRCONV")       ' ＦＴＩＲ換算
            .INSPECTWAY = rs("INSPECTWAY")   ' 検査方法
            .JudgData = rs("JUDGDATA")       ' 検索対象値
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

    DBDRV_GetTBCMJ003 = FUNCTION_RETURN_SUCCESS
End Function

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMJ004」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ004 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCMJ004(records() As typ_TBCMJ004, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, CSMEAS, PRE70P, INSPECTWAY, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ004"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ004 = FUNCTION_RETURN_FAILURE
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
''''            .CSMEAS = rs("CSMEAS")           ' Cs実測値
''''            .PRE70P = rs("PRE70P")           ' ７０％推定値
'OI_NULL対応　2005/03/03 TUKU START --------------------------------------------------
            If IsNull(rs("CSMEAS")) = False Then .CSMEAS = rs("CSMEAS") Else .CSMEAS = -1  ' Cs実測値
            If IsNull(rs("PRE70P")) = False Then .PRE70P = rs("PRE70P") Else .PRE70P = -1  ' ７０％推定値
'OI_NULL対応　2005/03/03 TUKU START --------------------------------------------------
            .INSPECTWAY = rs("INSPECTWAY")   ' 検査方法
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

    DBDRV_GetTBCMJ004 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME037」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME037 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCME037(records() As typ_TBCME037, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME037"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME037 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .DELCLS = rs("DELCLS")           ' 削除区分
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCD = rs("PROCCD")           ' 工程コード
            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
            .RPHINBAN = rs("RPHINBAN")       ' ねらい品番
            .RPREVNUM = rs("RPREVNUM")       ' ねらい品番製品番号改訂番号
            .RPFACT = rs("RPFACT")           ' ねらい品番工場
            .RPOPCOND = rs("RPOPCOND")       ' ねらい品番操業条件
            .PRODCOND = rs("PRODCOND")       ' 製作条件
            .PGID = rs("PGID")               ' ＰＧ－ＩＤ
            .UPLENGTH = rs("UPLENGTH")       ' 引上げ長さ
            .TOPLENG = rs("TOPLENG")         ' ＴＯＰ長さ
            .BODYLENG = rs("BODYLENG")       ' 直胴長さ
            .BOTLENG = rs("BOTLENG")         ' ＢＯＴ長さ
            .FREELENG = rs("FREELENG")       ' フリー長
            .DIAMETER = rs("DIAMETER")       ' 直径
            .CHARGE = rs("CHARGE")           ' チャージ量
            .SEED = rs("SEED")               ' シード
            .ADDDPCLS = rs("ADDDPCLS")       ' 追加ドープ種類
            .ADDDPPOS = rs("ADDDPPOS")       ' 追加ドープ位置
            .ADDDPVAL = rs("ADDDPVAL")       ' 追加ドープ量
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME037 = FUNCTION_RETURN_SUCCESS
End Function
'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :引数で渡されたレコードをTBCMJ003に追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :record        ,I  ,typ_cmjc001c_Disp ,抽出レコード
'          :CRYNUM        ,I  ,String       ,結晶番号
'          :TSTAFFID      ,I  ,String       ,登録社員ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :処理回数はテーブル上の最大値+1とする。
'履歴      :2001/06/22(Fri)作成　長野

Public Function DBDRV_Getcmjc001c_Exec(record As typ_cmjc001c_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN

Dim sql As String       'SQL全体
Dim sqlBase As String   'SQLベース部分
Dim sqlWhere As String  'SQLWhere部分
Dim sqlGroup As String  'SQLGroup部分
Dim SetDate As Variant  '設定日時

'    CRYNUM             結晶番号　⇒引数
'    TRANCNT         　 処理回数　⇒最大
'   TSTAFFID            登録社員ID　⇒引数
 '   REGDATE 　　　     登録日付　⇒SYSDATE
 '   KSTAFFID           更新社員ID　⇒" "
 '   UPDDATE            更新日付　⇒SYSDATE
 '   SENDFLAG           送信フラグ　⇒"0"
 '   SENDDATE           送信日付　⇒SYSDATE
    
    DBDRV_Getcmjc001c_Exec = FUNCTION_RETURN_FAILURE

    ''設定日時

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001c_SQL.bas -- Function DBDRV_Getcmjc001c_Exec"

    SetDate = Format$(record.SETDTM, "yyyy-mm-dd hh:mm:ss")

    ''SQLを組み立てる
    sqlBase = "Insert into TBCMJ003 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', " & record.OIMEAS1 & ", " & _
               record.OIMEAS2 & ", " & record.OIMEAS3 & ", " & record.OIMEAS4 & ", " & record.OIMEAS5 & ", " & record.ORGRES & ", " & _
               "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss'), " & record.EFFECTTM & ", '" & record.FTIRMETH & "', " & record.YCOEF & ", " & record.XCOEF & ", " & _
               record.AVE & ", " & record.SIGMA & ", " & record.FTIRCONV & ", '" & record.INSPECTWAY & "', " & record.JudgData & ", '" & TSTAFFID & "', " & _
               "SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ003 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
    
    ''SQLの実行
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001c_Exec = FUNCTION_RETURN_SUCCESS
    

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
