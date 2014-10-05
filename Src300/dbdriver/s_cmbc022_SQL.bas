Attribute VB_Name = "s_cmbc022_SQL"
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
Public Type typ_cmjc001c_Disp
   ' CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
   ' TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long 6桁対応 2007/05/28 SETsw kubota
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

'(2002/07 DBDRV_GetTBCME018より移動)
'フィールド名検索用
Dim fldNames() As String    '現rsに含まれるフィールド名保持配列
Dim fldCnt As Integer       '現rsに含まれるフィールド数

'使用していないため削除 2011/08/23 SETsw kubota
''------------------------------------------------
'' DBアクセス関数
''------------------------------------------------
'
''概要      :テーブル「TBCMJ003」から条件にあったレコードを抽出する
''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''          :records()     ,O  ,typ_cmjc001c_Disp ,抽出レコード
''          :SPLNUMs()     ,I  ,Integer      ,抽出条件配列(サンプルNo)
''          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
''説明      :
''履歴      :2001/06/20(Wed)作成　長野
'Public Function DBDRV_Getcmjc001c_Disp(records() As typ_cmjc001c_Disp, SPLNUMs() As Integer) As FUNCTION_RETURN
'Dim sql As String       'SQL全体
'Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
'Dim sqlWhere As String  'SQLのWHERE部分
'Dim sqlGroup As String  'SQLのGROUP部分
'Dim sqlOrder As String  'SQLのOrder部分
'Dim rs As OraDynaset    'RecordSet
'Dim recCnt As Long      'レコード数
'Dim i As Long           'ループカウント
'
'    DBDRV_Getcmjc001c_Disp = FUNCTION_RETURN_FAILURE
'
'    ''SQLを組み立てる
'
'    'エラーハンドラの設定
'    On Error GoTo proc_err
'    gErr.Push "s_cmzcF_cmjc001c_SQL.bas -- Function DBDRV_Getcmjc001c_Disp"
'
'    sqlBase = "Select POSITION, SMPKBN, TRANCOND, MAX(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
'              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
'              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA "
'    sqlBase = sqlBase & "From TBCMJ003"
'    ''抽出条件(ｻﾝﾌﾟﾙNO)の取り出し
'     sqlWhere = "Where SMPLNO in ("
'    For i = 1 To UBound(SPLNUMs)
'        sqlWhere = sqlWhere & "'" & SPLNUMs(i) & "'"
'        If i < UBound(SPLNUMs) Then
'            sqlWhere = sqlWhere & ", "
'        End If
'    Next
'    sqlWhere = sqlWhere & ") "
'    sqlGroup = "GROUP BY CRYNUM, POSITION, SMPKBN, TRANCOND "
'    sqlOrder = "ORDER BY POSITION"
'    sql = sqlBase & sqlWhere & sqlGroup & sqlOrder
'
'    ''データを抽出する
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'    If rs Is Nothing Then
'        ReDim records(0)
'        DBDRV_Getcmjc001c_Disp = FUNCTION_RETURN_FAILURE
'        GoTo proc_exit
'    End If
'
'    ''抽出結果を格納する
'    recCnt = rs.RecordCount
'    ReDim records(recCnt)
'    For i = 1 To recCnt
'        With records(i)
'            .POSITION = rs("POSITION")       ' 位置
'            .SMPKBN = rs("SMPKBN")           ' サンプル区分
'            .TRANCOND = rs("TRANCOND")       ' 処理条件
'            .SMPLNO = rs("SMPLNO")           ' サンプルＮｏ
'            .SMPLUMU = rs("SMPLUMU")         ' サンプル有無
'            .hinban = rs("HINBAN")           ' 品番
'            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
'            .factory = rs("FACTORY")         ' 工場
'            .opecond = rs("OPECOND")         ' 操業条件
'            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
'            .PROCCODE = rs("PROCCODE")       ' 工程コード
'            .GOUKI = rs("GOUKI")             ' 号機
'            .OIMEAS1 = rs("OIMEAS1")         ' Ｏｉ測定値１
'            .OIMEAS2 = rs("OIMEAS2")         ' Ｏｉ測定値２
'            .OIMEAS3 = rs("OIMEAS3")         ' Ｏｉ測定値３
'            .OIMEAS4 = rs("OIMEAS4")         ' Ｏｉ測定値４
'            .OIMEAS5 = rs("OIMEAS5")         ' Ｏｉ測定値５
'            .ORGRES = rs("ORGRES")           ' ＯＲＧ結果
'            .SETDTM = rs("SETDTM")           ' 設定日時
'            .EFFECTTM = rs("EFFECTTM")       ' 有効時間
'            .FTIRMETH = rs("FTIRMETH")       ' ＦＴＩＲ相関式
'            .YCOEF = rs("YCOEF")             ' ＦＴＩＲ換算式（Ｙ切片）
'            .XCOEF = rs("XCOEF")             ' ＦＴＩＲ換算式（Ｘ係数）
'            .AVE = rs("AVE")                 ' ＡＶＥ
'            .SIGMA = rs("SIGMA")             ' σ（シグマ）
'            .FTIRCONV = rs("FTIRCONV")       ' ＦＴＩＲ換算
'            .INSPECTWAY = rs("INSPECTWAY")   ' 検査方法
'            .JudgData = rs("JUDGDATA")       ' 検索対象値
'        End With
'        rs.MoveNext
'    Next
'    rs.Close
'
'    DBDRV_Getcmjc001c_Disp = FUNCTION_RETURN_SUCCESS
'
'proc_exit:
'    '終了
'    gErr.Pop
'    Exit Function
'
'proc_err:
'    'エラーハンドラ
'    Debug.Print "====== Error SQL ======"
'    Debug.Print sql
'    gErr.HandleError
'    Resume proc_exit
'End Function

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
    '' OI_NULL対応　2005/03/07 TUKU START 山口殿からの依頼で変更中止--------------------------------------------------------------------
''''    sqlBase = "Insert into TBCMJ003 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
''''              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
''''              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
''''    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
''''               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
''''               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', "
''''                If (record.OIMEAS1 = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.OIMEAS1 & ", "    'OI測定値1
''''                If (record.OIMEAS2 = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.OIMEAS2 & ", "    'OI測定値2
''''                If (record.OIMEAS3 = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.OIMEAS3 & ", "    'OI測定値2
''''                If (record.OIMEAS4 = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.OIMEAS4 & ", "    'OI測定値3
''''                If (record.OIMEAS5 = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.OIMEAS5 & ", "    'OI測定値4
''''                If (record.ORGRES = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.ORGRES & ", "      'ORG
''''    sqlBase = sqlBase & "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss'), " & record.EFFECTTM & ", '" & record.FTIRMETH & "', " & record.YCOEF & ", " & record.XCOEF & ", " & _
''''               record.AVE & ", " & record.SIGMA & ", " & record.FTIRCONV & ", '" & record.INSPECTWAY & "', " & record.JudgData & ", '" & TSTAFFID & "', " & _
''''               "SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ003 "
''''    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'''''    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
''''    sql = sqlBase & sqlWhere & sqlGroup
    '' OI_NULL対応　2005/03/07 TUKU END   --------------------------------------------------------------------
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


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMB014」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB014 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMB014_SQL.basより移動)
Public Function DBDRV_GetTBCMB014(records() As typ_TBCMB014, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select GOUKI, INPDATE, FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG," & _
              " SENDDATE "
    sqlBase = sqlBase & "From TBCMB014"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB014 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .GOUKI = rs("GOUKI")             ' 号機
            .INPDATE = rs("INPDATE")         ' 日付
            .FTIRFZI = rs("FTIRFZI")         ' FTIR（FZ)
            .FTIRCZH = rs("FTIRCZH")         ' FTIR（CZ高）
            .FTIRCZC = rs("FTIRCZC")         ' FTIR（CZ中）
            .MS1FZ = rs("MS1FZ")             ' 測定サンプル1（FZ)
            .MS1CZ1 = rs("MS1CZ1")           ' 測定サンプル1（CZ-1)
            .MS1CZ2 = rs("MS1CZ2")           ' 測定サンプル1（CZ-2)
            .MS2FZ = rs("MS2FZ")             ' 測定サンプル2（FZ)
            .MS2CZ1 = rs("MS2CZ1")           ' 測定サンプル2（CZ-1)
            .MS2CZ2 = rs("MS2CZ2")           ' 測定サンプル2（CZ-2)
            .MS3FZ = rs("MS3FZ")             ' 測定サンプル3（FZ)
            .MS3CZ1 = rs("MS3CZ1")           ' 測定サンプル3（CZ-1)
            .MS3CZ2 = rs("MS3CZ2")           ' 測定サンプル3（CZ-2)
            .MS4FZ = rs("MS4FZ")             ' 測定サンプル4（FZ)
            .MS4CZ1 = rs("MS4CZ1")           ' 測定サンプル4（CZ-1)
            .MS4CZ2 = rs("MS4CZ2")           ' 測定サンプル4（CZ-2)
            .MS5FZ = rs("MS5FZ")             ' 測定サンプル5（FZ)
            .MS5CZ1 = rs("MS5CZ1")           ' 測定サンプル5（CZ-1)
            .MS5CZ2 = rs("MS5CZ2")           ' 測定サンプル5（CZ-2)
            .MSAVEFZ = rs("MSAVEFZ")         ' 測定平均（FZ）
            .MSAVECZ1 = rs("MSAVECZ1")       ' 測定平均（CZ-1）
            .MSAVECZ2 = rs("MSAVECZ2")       ' 測定平均（CZ-2）
            .MSSGFZ = rs("MSSGFZ")           ' 測定σ（FZ）
            .MSSGCZ1 = rs("MSSGCZ1")         ' 測定σ（CZ-1）
            .MSSGCZ2 = rs("MSSGCZ2")         ' 測定σ（CZ-2）
            .MSPSGFZ = rs("MSPSGFZ")         ' 測定AVE+σ（FZ）
            .MSPSGCZ1 = rs("MSPSGCZ1")       ' 測定AVE+σ（CZ-1）
            .MSPSGCZ2 = rs("MSPSGCZ2")       ' 測定AVE+σ（CZ-2）
            .MSNSGFZ = rs("MSNSGFZ")         ' 測定AVE-σ（FZ）
            .MSNSGCZ1 = rs("MSNSGCZ1")       ' 測定AVE-σ（CZ-1）
            .MSNSGCZ2 = rs("MSNSGCZ2")       ' 測定AVE-σ（CZ-2）
            .MINFZ = rs("MINFZ")             ' MIN（FZ）
            .MINCZ1 = rs("MINCZ1")           ' MIN（CZ-1）
            .MINCZ2 = rs("MINCZ2")           ' MIN（CZ-2）
            .MAXFZ = rs("MAXFZ")             ' MAX（FZ）
            .MAXCZ1 = rs("MAXCZ1")           ' MAX（CZ-1）
            .MAXCZ2 = rs("MAXCZ2")           ' MAX（CZ-2）
            .SGCK1FZ = rs("SGCK1FZ")         ' σckサンプル1（FZ)
            .SGCK1CZ1 = rs("SGCK1CZ1")       ' σckサンプル1（CZ-1)
            .SGCK1CZ2 = rs("SGCK1CZ2")       ' σckサンプル1（CZ-2)
            .SGCK2FZ = rs("SGCK2FZ")         ' σckサンプル2（FZ)
            .SGCK2CZ1 = rs("SGCK2CZ1")       ' σckサンプル2（CZ-1)
            .SGCK2CZ2 = rs("SGCK2CZ2")       ' σckサンプル2（CZ-2)
            .SGCK3FZ = rs("SGCK3FZ")         ' σckサンプル3（FZ)
            .SGCK3CZ1 = rs("SGCK3CZ1")       ' σckサンプル3（CZ-1)
            .SGCK3CZ2 = rs("SGCK3CZ2")       ' σckサンプル3（CZ-2)
            .SGCK4FZ = rs("SGCK4FZ")         ' σckサンプル4（FZ)
            .SGCK4CZ1 = rs("SGCK4CZ1")       ' σckサンプル4（CZ-1)
            .SGCK4CZ2 = rs("SGCK4CZ2")       ' σckサンプル4（CZ-2)
            .SGCK5FZ = rs("SGCK5FZ")         ' σckサンプル5（FZ)
            .SGCK5CZ1 = rs("SGCK5CZ1")       ' σckサンプル5（CZ-1)
            .SGCK5CZ2 = rs("SGCK5CZ2")       ' σckサンプル5（CZ-2)
            .SGCKDFZ = rs("SGCKDFZ")         ' σckデータ数（FZ）
            .SGCKDCZ1 = rs("SGCKDCZ1")       ' σckデータ数（CZ-1）
            .SGCKDCZ2 = rs("SGCKDCZ2")       ' σckデータ数（CZ-2）
            .SGCKAFZ = rs("SGCKAFZ")         ' σck平均（FZ）
            .SGCKAACZ1 = rs("SGCKAACZ1")     ' σck平均（CZ-1）
            .SGCKACZ2 = rs("SGCKACZ2")       ' σck平均（CZ-2）
            .SGNFZ = rs("SGNFZ")             ' σckσ（FZ）
            .SGNCZ1 = rs("SGNCZ1")           ' σckσ CZ-1）
            .SGNCZ2 = rs("SGNCZ2")           ' σckσ（CZ-2）
            .FTIRFZ = rs("FTIRFZ")           ' FTIR換算（FZ）
            .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR換算（CZ-1）
            .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR換算（CZ-2）
            .EFFECTTM = rs("EFFECTTM")       ' 有効時間
            .YCOEF = rs("YCOEF")             ' ＦＴＩＲ換算式（Ｙ切片）
            .XCOEF = rs("XCOEF")             ' ＦＴＩＲ換算式（Ｘ係数）
            .RSQUARE = rs("RSQUARE")         ' Ｒ２乗
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

    DBDRV_GetTBCMB014 = FUNCTION_RETURN_SUCCESS
End Function


'概要      :引数のフィールド名がfldNames()配列に含まれているかどうかの判定。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :fldName       ,I  ,typ_TBCME018 ,抽出レコード
'          :戻り値        ,O  ,Boolean      ,True:在り／False：無し
'説明      :
'履歴      :2001/06/27作成　野村  (2002/07 DBDRV_GetTBCME018より移動)

Private Function fldNameExist(fldName As String) As Boolean
    Dim sql         As String           'SQL全体
    Dim i As Integer                    'ﾙｰﾌﾟｶｳﾝﾄ


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME018_SQL.bas -- Function fldNameExist"

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

'概要      :テーブル「TBCME036」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME036 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCME036_SQL.basより移動)
'          :06/04/11 ooba　関数名変更 <DBDRV_GetTBCME036> ⇒ <DBDRV_GetTBCME036_cmbc022>
Public Function DBDRV_GetTBCME036_cmbc022(records() As typ_TBCME036, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
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
        DBDRV_GetTBCME036_cmbc022 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
'NULL対応 ----- START ----- 2003/12/10
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
'NULL対応 -----  END  ----- 2003/12/10

    DBDRV_GetTBCME036_cmbc022 = FUNCTION_RETURN_SUCCESS
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
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcF_TBCME037_SQL.basより移動)
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

'概要      :テーブル「TBCMJ005」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ005 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMJ005_SQL.basより移動)
Public Function DBDRV_GetTBCMJ005(records() As typ_TBCMJ005, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, CALCMAX, CALCAVE, MEAS1, MEAS2, MEAS3, MEAS4," & _
              " MEAS5, MEAS6, MEAS7, MEAS8, MEAS9, MEAS10, MEAS11, MEAS12, MEAS13, MEAS14, MEAS15, MEAS16, MEAS17, MEAS18," & _
              " MEAS19, MEAS20, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ005"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ005 = FUNCTION_RETURN_FAILURE
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
            .CALCMAX = rs("CALCMAX")         ' 計算結果 Max
            .CALCAVE = rs("CALCAVE")         ' 計算結果 Ave
            .MEAS1 = rs("MEAS1")             ' 測定値１
            .MEAS2 = rs("MEAS2")             ' 測定値２
            .MEAS3 = rs("MEAS3")             ' 測定値３
            .MEAS4 = rs("MEAS4")             ' 測定値４
            .MEAS5 = rs("MEAS5")             ' 測定値５
            .MEAS6 = rs("MEAS6")             ' 測定値６
            .MEAS7 = rs("MEAS7")             ' 測定値７
            .MEAS8 = rs("MEAS8")             ' 測定値８
            .MEAS9 = rs("MEAS9")             ' 測定値９
            .MEAS10 = rs("MEAS10")           ' 測定値１０
            .MEAS11 = rs("MEAS11")           ' 測定値１１
            .MEAS12 = rs("MEAS12")           ' 測定値１２
            .MEAS13 = rs("MEAS13")           ' 測定値１３
            .MEAS14 = rs("MEAS14")           ' 測定値１４
            .MEAS15 = rs("MEAS15")           ' 測定値１５
            .MEAS16 = rs("MEAS16")           ' 測定値１６
            .MEAS17 = rs("MEAS17")           ' 測定値１７
            .MEAS18 = rs("MEAS18")           ' 測定値１８
            .MEAS19 = rs("MEAS19")           ' 測定値１９
            .MEAS20 = rs("MEAS20")           ' 測定値２０
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

    DBDRV_GetTBCMJ005 = FUNCTION_RETURN_SUCCESS
End Function


'概要      :テーブル「TBCMJ006」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ006 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCMJ006_SQL.basより移動)
Public Function DBDRV_GetTBCMJ006(records() As typ_TBCMJ006, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MSRSDEN, MSRSLDL, MSRSDVD2, MS01LDL1, MS01LDL2, MS01LDL3, MS01LDL4, MS01LDL5, MS01DEN1, MS01DEN2," & _
              " MS01DEN3, MS01DEN4, MS01DEN5, MS02LDL1, MS02LDL2, MS02LDL3, MS02LDL4, MS02LDL5, MS02DEN1, MS02DEN2, MS02DEN3," & _
              " MS02DEN4, MS02DEN5, MS03LDL1, MS03LDL2, MS03LDL3, MS03LDL4, MS03LDL5, MS03DEN1, MS03DEN2, MS03DEN3, MS03DEN4," & _
              " MS03DEN5, MS04LDL1, MS04LDL2, MS04LDL3, MS04LDL4, MS04LDL5, MS04DEN1, MS04DEN2, MS04DEN3, MS04DEN4, MS04DEN5," & _
              " MS05LDL1, MS05LDL2, MS05LDL3, MS05LDL4, MS05LDL5, MS05DEN1, MS05DEN2, MS05DEN3, MS05DEN4, MS05DEN5, MS06LDL1," & _
              " MS06LDL2, MS06LDL3, MS06LDL4, MS06LDL5, MS06DEN1, MS06DEN2, MS06DEN3, MS06DEN4, MS06DEN5, MS07LDL1, MS07LDL2," & _
              " MS07LDL3, MS07LDL4, MS07LDL5, MS07DEN1, MS07DEN2, MS07DEN3, MS07DEN4, MS07DEN5, MS08LDL1, MS08LDL2, MS08LDL3," & _
              " MS08LDL4, MS08LDL5, MS08DEN1, MS08DEN2, MS08DEN3, MS08DEN4, MS08DEN5, MS09LDL1, MS09LDL2, MS09LDL3, MS09LDL4," & _
              " MS09LDL5, MS09DEN1, MS09DEN2, MS09DEN3, MS09DEN4, MS09DEN5, MS10LDL1, MS10LDL2, MS10LDL3, MS10LDL4, MS10LDL5," & _
              " MS10DEN1, MS10DEN2, MS10DEN3, MS10DEN4, MS10DEN5, MS11LDL1, MS11LDL2, MS11LDL3, MS11LDL4, MS11LDL5, MS11DEN1," & _
              " MS11DEN2, MS11DEN3, MS11DEN4, MS11DEN5, MS12LDL1, MS12LDL2, MS12LDL3, MS12LDL4, MS12LDL5, MS12DEN1, MS12DEN2," & _
              " MS12DEN3, MS12DEN4, MS12DEN5, MS13LDL1, MS13LDL2, MS13LDL3, MS13LDL4, MS13LDL5, MS13DEN1, MS13DEN2, MS13DEN3," & _
              " MS13DEN4, MS13DEN5, MS14LDL1, MS14LDL2, MS14LDL3, MS14LDL4, MS14LDL5, MS14DEN1, MS14DEN2, MS14DEN3, MS14DEN4," & _
              " MS14DEN5, MS15LDL1, MS15LDL2, MS15LDL3, MS15LDL4, MS15LDL5, MS15DEN1, MS15DEN2, MS15DEN3, MS15DEN4, MS15DEN5," & _
              " MS01DVD2,  MS02DVD2 , MS03DVD2 , MS04DVD2 , MS05DVD2 , TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ006"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ006 = FUNCTION_RETURN_FAILURE
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
            .MSRSDEN = rs("MSRSDEN")         ' 測定結果 Den
            .MSRSLDL = rs("MSRSLDL")         ' 測定結果 L/DL
            .MSRSDVD2 = rs("MSRSDVD2")       ' 測定結果 DVD2
            .MS01LDL1 = rs("MS01LDL1")       ' 測定値01 L/DL1
            .MS01LDL2 = rs("MS01LDL2")       ' 測定値01 L/DL2
            .MS01LDL3 = rs("MS01LDL3")       ' 測定値01 L/DL3
            .MS01LDL4 = rs("MS01LDL4")       ' 測定値01 L/DL4
            .MS01LDL5 = rs("MS01LDL5")       ' 測定値01 L/DL5
            .MS01DEN1 = rs("MS01DEN1")       ' 測定値01 Den1
            .MS01DEN2 = rs("MS01DEN2")       ' 測定値01 Den2
            .MS01DEN3 = rs("MS01DEN3")       ' 測定値01 Den3
            .MS01DEN4 = rs("MS01DEN4")       ' 測定値01 Den4
            .MS01DEN5 = rs("MS01DEN5")       ' 測定値01 Den5
            .MS02LDL1 = rs("MS02LDL1")       ' 測定値02 L/DL1
            .MS02LDL2 = rs("MS02LDL2")       ' 測定値02 L/DL2
            .MS02LDL3 = rs("MS02LDL3")       ' 測定値02 L/DL3
            .MS02LDL4 = rs("MS02LDL4")       ' 測定値02 L/DL4
            .MS02LDL5 = rs("MS02LDL5")       ' 測定値02 L/DL5
            .MS02DEN1 = rs("MS02DEN1")       ' 測定値02 Den1
            .MS02DEN2 = rs("MS02DEN2")       ' 測定値02 Den2
            .MS02DEN3 = rs("MS02DEN3")       ' 測定値02 Den3
            .MS02DEN4 = rs("MS02DEN4")       ' 測定値02 Den4
            .MS02DEN5 = rs("MS02DEN5")       ' 測定値02 Den5
            .MS03LDL1 = rs("MS03LDL1")       ' 測定値03 L/DL1
            .MS03LDL2 = rs("MS03LDL2")       ' 測定値03 L/DL2
            .MS03LDL3 = rs("MS03LDL3")       ' 測定値03 L/DL3
            .MS03LDL4 = rs("MS03LDL4")       ' 測定値03 L/DL4
            .MS03LDL5 = rs("MS03LDL5")       ' 測定値03 L/DL5
            .MS03DEN1 = rs("MS03DEN1")       ' 測定値03 Den1
            .MS03DEN2 = rs("MS03DEN2")       ' 測定値03 Den2
            .MS03DEN3 = rs("MS03DEN3")       ' 測定値03 Den3
            .MS03DEN4 = rs("MS03DEN4")       ' 測定値03 Den4
            .MS03DEN5 = rs("MS03DEN5")       ' 測定値03 Den5
            .MS04LDL1 = rs("MS04LDL1")       ' 測定値04 L/DL1
            .MS04LDL2 = rs("MS04LDL2")       ' 測定値04 L/DL2
            .MS04LDL3 = rs("MS04LDL3")       ' 測定値04 L/DL3
            .MS04LDL4 = rs("MS04LDL4")       ' 測定値04 L/DL4
            .MS04LDL5 = rs("MS04LDL5")       ' 測定値04 L/DL5
            .MS04DEN1 = rs("MS04DEN1")       ' 測定値04 Den1
            .MS04DEN2 = rs("MS04DEN2")       ' 測定値04 Den2
            .MS04DEN3 = rs("MS04DEN3")       ' 測定値04 Den3
            .MS04DEN4 = rs("MS04DEN4")       ' 測定値04 Den4
            .MS04DEN5 = rs("MS04DEN5")       ' 測定値04 Den5
            .MS05LDL1 = rs("MS05LDL1")       ' 測定値05 L/DL1
            .MS05LDL2 = rs("MS05LDL2")       ' 測定値05 L/DL2
            .MS05LDL3 = rs("MS05LDL3")       ' 測定値05 L/DL3
            .MS05LDL4 = rs("MS05LDL4")       ' 測定値05 L/DL4
            .MS05LDL5 = rs("MS05LDL5")       ' 測定値05 L/DL5
            .MS05DEN1 = rs("MS05DEN1")       ' 測定値05 Den1
            .MS05DEN2 = rs("MS05DEN2")       ' 測定値05 Den2
            .MS05DEN3 = rs("MS05DEN3")       ' 測定値05 Den3
            .MS05DEN4 = rs("MS05DEN4")       ' 測定値05 Den4
            .MS05DEN5 = rs("MS05DEN5")       ' 測定値05 Den5
            .MS06LDL1 = rs("MS06LDL1")       ' 測定値06 L/DL1
            .MS06LDL2 = rs("MS06LDL2")       ' 測定値06 L/DL2
            .MS06LDL3 = rs("MS06LDL3")       ' 測定値06 L/DL3
            .MS06LDL4 = rs("MS06LDL4")       ' 測定値06 L/DL4
            .MS06LDL5 = rs("MS06LDL5")       ' 測定値06 L/DL5
            .MS06DEN1 = rs("MS06DEN1")       ' 測定値06 Den1
            .MS06DEN2 = rs("MS06DEN2")       ' 測定値06 Den2
            .MS06DEN3 = rs("MS06DEN3")       ' 測定値06 Den3
            .MS06DEN4 = rs("MS06DEN4")       ' 測定値06 Den4
            .MS06DEN5 = rs("MS06DEN5")       ' 測定値06 Den5
            .MS07LDL1 = rs("MS07LDL1")       ' 測定値07 L/DL1
            .MS07LDL2 = rs("MS07LDL2")       ' 測定値07 L/DL2
            .MS07LDL3 = rs("MS07LDL3")       ' 測定値07 L/DL3
            .MS07LDL4 = rs("MS07LDL4")       ' 測定値07 L/DL4
            .MS07LDL5 = rs("MS07LDL5")       ' 測定値07 L/DL5
            .MS07DEN1 = rs("MS07DEN1")       ' 測定値07 Den1
            .MS07DEN2 = rs("MS07DEN2")       ' 測定値07 Den2
            .MS07DEN3 = rs("MS07DEN3")       ' 測定値07 Den3
            .MS07DEN4 = rs("MS07DEN4")       ' 測定値07 Den4
            .MS07DEN5 = rs("MS07DEN5")       ' 測定値07 Den5
            .MS08LDL1 = rs("MS08LDL1")       ' 測定値08 L/DL1
            .MS08LDL2 = rs("MS08LDL2")       ' 測定値08 L/DL2
            .MS08LDL3 = rs("MS08LDL3")       ' 測定値08 L/DL3
            .MS08LDL4 = rs("MS08LDL4")       ' 測定値08 L/DL4
            .MS08LDL5 = rs("MS08LDL5")       ' 測定値08 L/DL5
            .MS08DEN1 = rs("MS08DEN1")       ' 測定値08 Den1
            .MS08DEN2 = rs("MS08DEN2")       ' 測定値08 Den2
            .MS08DEN3 = rs("MS08DEN3")       ' 測定値08 Den3
            .MS08DEN4 = rs("MS08DEN4")       ' 測定値08 Den4
            .MS08DEN5 = rs("MS08DEN5")       ' 測定値08 Den5
            .MS09LDL1 = rs("MS09LDL1")       ' 測定値09 L/DL1
            .MS09LDL2 = rs("MS09LDL2")       ' 測定値09 L/DL2
            .MS09LDL3 = rs("MS09LDL3")       ' 測定値09 L/DL3
            .MS09LDL4 = rs("MS09LDL4")       ' 測定値09 L/DL4
            .MS09LDL5 = rs("MS09LDL5")       ' 測定値09 L/DL5
            .MS09DEN1 = rs("MS09DEN1")       ' 測定値09 Den1
            .MS09DEN2 = rs("MS09DEN2")       ' 測定値09 Den2
            .MS09DEN3 = rs("MS09DEN3")       ' 測定値09 Den3
            .MS09DEN4 = rs("MS09DEN4")       ' 測定値09 Den4
            .MS09DEN5 = rs("MS09DEN5")       ' 測定値09 Den5
            .MS10LDL1 = rs("MS10LDL1")       ' 測定値10 L/DL1
            .MS10LDL2 = rs("MS10LDL2")       ' 測定値10 L/DL2
            .MS10LDL3 = rs("MS10LDL3")       ' 測定値10 L/DL3
            .MS10LDL4 = rs("MS10LDL4")       ' 測定値10 L/DL4
            .MS10LDL5 = rs("MS10LDL5")       ' 測定値10 L/DL5
            .MS10DEN1 = rs("MS10DEN1")       ' 測定値10 Den1
            .MS10DEN2 = rs("MS10DEN2")       ' 測定値10 Den2
            .MS10DEN3 = rs("MS10DEN3")       ' 測定値10 Den3
            .MS10DEN4 = rs("MS10DEN4")       ' 測定値10 Den4
            .MS10DEN5 = rs("MS10DEN5")       ' 測定値10 Den5
            .MS11LDL1 = rs("MS11LDL1")       ' 測定値11 L/DL1
            .MS11LDL2 = rs("MS11LDL2")       ' 測定値11 L/DL2
            .MS11LDL3 = rs("MS11LDL3")       ' 測定値11 L/DL3
            .MS11LDL4 = rs("MS11LDL4")       ' 測定値11 L/DL4
            .MS11LDL5 = rs("MS11LDL5")       ' 測定値11 L/DL5
            .MS11DEN1 = rs("MS11DEN1")       ' 測定値11 Den1
            .MS11DEN2 = rs("MS11DEN2")       ' 測定値11 Den2
            .MS11DEN3 = rs("MS11DEN3")       ' 測定値11 Den3
            .MS11DEN4 = rs("MS11DEN4")       ' 測定値11 Den4
            .MS11DEN5 = rs("MS11DEN5")       ' 測定値11 Den5
            .MS12LDL1 = rs("MS12LDL1")       ' 測定値12 L/DL1
            .MS12LDL2 = rs("MS12LDL2")       ' 測定値12 L/DL2
            .MS12LDL3 = rs("MS12LDL3")       ' 測定値12 L/DL3
            .MS12LDL4 = rs("MS12LDL4")       ' 測定値12 L/DL4
            .MS12LDL5 = rs("MS12LDL5")       ' 測定値12 L/DL5
            .MS12DEN1 = rs("MS12DEN1")       ' 測定値12 Den1
            .MS12DEN2 = rs("MS12DEN2")       ' 測定値12 Den2
            .MS12DEN3 = rs("MS12DEN3")       ' 測定値12 Den3
            .MS12DEN4 = rs("MS12DEN4")       ' 測定値12 Den4
            .MS12DEN5 = rs("MS12DEN5")       ' 測定値12 Den5
            .MS13LDL1 = rs("MS13LDL1")       ' 測定値13 L/DL1
            .MS13LDL2 = rs("MS13LDL2")       ' 測定値13 L/DL2
            .MS13LDL3 = rs("MS13LDL3")       ' 測定値13 L/DL3
            .MS13LDL4 = rs("MS13LDL4")       ' 測定値13 L/DL4
            .MS13LDL5 = rs("MS13LDL5")       ' 測定値13 L/DL5
            .MS13DEN1 = rs("MS13DEN1")       ' 測定値13 Den1
            .MS13DEN2 = rs("MS13DEN2")       ' 測定値13 Den2
            .MS13DEN3 = rs("MS13DEN3")       ' 測定値13 Den3
            .MS13DEN4 = rs("MS13DEN4")       ' 測定値13 Den4
            .MS13DEN5 = rs("MS13DEN5")       ' 測定値13 Den5
            .MS14LDL1 = rs("MS14LDL1")       ' 測定値14 L/DL1
            .MS14LDL2 = rs("MS14LDL2")       ' 測定値14 L/DL2
            .MS14LDL3 = rs("MS14LDL3")       ' 測定値14 L/DL3
            .MS14LDL4 = rs("MS14LDL4")       ' 測定値14 L/DL4
            .MS14LDL5 = rs("MS14LDL5")       ' 測定値14 L/DL5
            .MS14DEN1 = rs("MS14DEN1")       ' 測定値14 Den1
            .MS14DEN2 = rs("MS14DEN2")       ' 測定値14 Den2
            .MS14DEN3 = rs("MS14DEN3")       ' 測定値14 Den3
            .MS14DEN4 = rs("MS14DEN4")       ' 測定値14 Den4
            .MS14DEN5 = rs("MS14DEN5")       ' 測定値14 Den5
            .MS15LDL1 = rs("MS15LDL1")       ' 測定値15 L/DL1
            .MS15LDL2 = rs("MS15LDL2")       ' 測定値15 L/DL2
            .MS15LDL3 = rs("MS15LDL3")       ' 測定値15 L/DL3
            .MS15LDL4 = rs("MS15LDL4")       ' 測定値15 L/DL4
            .MS15LDL5 = rs("MS15LDL5")       ' 測定値15 L/DL5
            .MS15DEN1 = rs("MS15DEN1")       ' 測定値15 Den1
            .MS15DEN2 = rs("MS15DEN2")       ' 測定値15 Den2
            .MS15DEN3 = rs("MS15DEN3")       ' 測定値15 Den3
            .MS15DEN4 = rs("MS15DEN4")       ' 測定値15 Den4
            .MS15DEN5 = rs("MS15DEN5")       ' 測定値15 Den5
            'NULL チェック
            If IsNull(rs("MS01DVD2")) = False Then
                .MS01DVD2 = rs("MS01DVD2")       ' 測定値01 DVD2   2002/7/02 tuku
            Else
                .MS01DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS02DVD2")) = False Then
                .MS02DVD2 = rs("MS02DVD2")       ' 測定値01 DVD2   2002/7/02 tuku
            Else
                .MS02DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS03DVD2")) = False Then
                .MS03DVD2 = rs("MS03DVD2")       ' 測定値01 DVD2   2002/7/02 tuku
            Else
                .MS03DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS04DVD2")) = False Then
                .MS04DVD2 = rs("MS04DVD2")       ' 測定値01 DVD2   2002/7/02 tuku
            Else
                .MS04DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS05DVD2")) = False Then
                .MS05DVD2 = rs("MS05DVD2")       ' 測定値01 DVD2   2002/7/02 tuku
            Else
                .MS05DVD2 = DEF_PARAM_VALUE
            End If
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

    DBDRV_GetTBCMJ006 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMJ007」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ007 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCMJ007_SQL.basより移動)
Public Function DBDRV_GetTBCMJ007(records() As typ_TBCMJ007, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
              " SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ007"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ007 = FUNCTION_RETURN_FAILURE
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
            .MEAS1 = rs("MEAS1")             ' 測定値１
            .MEAS2 = rs("MEAS2")             ' 測定値２
            .MEAS3 = rs("MEAS3")             ' 測定値３
            .MEAS4 = rs("MEAS4")             ' 測定値４
            .MEAS5 = rs("MEAS5")             ' 測定値５
            .MEASPEAK = rs("MEASPEAK")       ' 測定値 ピーク値
            .CALCMEAS = rs("CALCMEAS")       ' 計算結果
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

    DBDRV_GetTBCMJ007 = FUNCTION_RETURN_SUCCESS
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
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCMJ008_SQL.basより移動)
Public Function DBDRV_GetTBCMJ008(records() As typ_TBCMJ008, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN," & _
              " MEASMAX, MEASAVE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
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
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMJ008 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME041」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME041 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCME041_SQL.basより移動)
Public Function DBDRV_GetTBCME041(records() As typ_TBCME041, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACTORY, OPECOND, LENGTH, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME041"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME041 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .Length = rs("LENGTH")           ' 長さ
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME041 = FUNCTION_RETURN_SUCCESS
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
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMJ003_SQL.basより移動)
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
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMJ004_SQL.basより移動)
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
            .CSMEAS = rs("CSMEAS")           ' Cs実測値
            .PRE70P = rs("PRE70P")           ' ７０％推定値
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

'概要      :テーブル「TBCMJ002」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ002 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCMJ002_SQL.basより移動)
Public Function DBDRV_GetTBCMJ002(records() As typ_TBCMJ002, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
              " OPECOND, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID," & _
              " UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ002"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ002 = FUNCTION_RETURN_FAILURE
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
            .TYPE = rs("TYPE")               ' タイプ
            .MEAS1 = rs("MEAS1")             ' 測定値１
            .MEAS2 = rs("MEAS2")             ' 測定値２
            .MEAS3 = rs("MEAS3")             ' 測定値３
            .MEAS4 = rs("MEAS4")             ' 測定値４
            .MEAS5 = rs("MEAS5")             ' 測定値５
            .EFEHS = rs("EFEHS")             ' 実効偏析
            .RRG = rs("RRG")                 ' ＲＲＧ
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

    DBDRV_GetTBCMJ002 = FUNCTION_RETURN_SUCCESS
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
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCMJ001_SQL.basより移動)
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


'///////////////////////////////////////////////////
' @(f)
' 機能    : σ判定基準取得
'
' 返り値  : True  - 正常
' 　　　    False - 失敗
'
' 引き数  : sSigCode  - σ判定基準
'
' 機能説明:
'2006/05/23追加
'///////////////////////////////////////////////////
Public Function GetSigChkCode(Optional ByRef sSigCode As String) As Boolean
    Dim dbIsMine    As Boolean
    Dim sSQL        As String
    Dim objRs       As Object
    
    GetSigChkCode = False
    sSigCode = ""
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc022_SQL.bas -- Function GetSigChkCode"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''ＳＱＬ文作成
    sSQL = ""
    sSQL = sSQL & "SELECT NVL(kcode01a9, ' ')"   '0:σ判定基準
    sSQL = sSQL & "  FROM koda9"
    sSQL = sSQL & " WHERE sysca9 = 'X'"
    sSQL = sSQL & "   AND shuca9 = '19'"
    sSQL = sSQL & "   AND codea9 = 'GFA'"
    
    Set objRs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)
    
    If objRs.EOF Then
        Call MsgOut(0, "σ判定基準のコードが登録されていません", ERR_DISP)
        Exit Function
    End If

    sSigCode = objRs(0)     ''σ判定基準
    
    objRs.Close
    
    ''σ判定基準
    If IsNumeric(sSigCode) = False Then
        Call MsgOut(0, "σ判定基準のコードが正しくありません", ERR_DISP)
        Exit Function
    End If
    
    If dbIsMine Then
        OraDBClose
    End If

    GetSigChkCode = True        ''処理成功を返す

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
    
End Function


