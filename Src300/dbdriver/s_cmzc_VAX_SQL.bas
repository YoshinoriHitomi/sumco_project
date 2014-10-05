Attribute VB_Name = "s_cmzcF_VAX_SQL"
Option Explicit


'                                     2001/10/03
'================================================
' DBアクセス関数
' VAX_DBアクセス用
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------

Public Type typ_VAX_DR_CNDS
    PG_ID   As String * 6      ' プログラムID
    DR_CHRG As Long            ' チャージ量
    DR_CPOS As Integer         ' ルツボ位置
    DR_CSIZ As Integer         ' ルツボサイズ
    DR_DIA  As Integer         ' 直径
    DR_LEN0 As Integer         ' 引上長(1本引き/R0)
    DR_LEN1 As Integer         ' 引上長(R1)
    DR_SR   As String * 9      ' 上軸回転数
    DR_CR   As String * 9      ' 下軸回転数
    DR_GAP  As Integer         ' ギャップ
    DR_PRES7 As String * 8       ' 炉内圧            '2003/05/16 osawa
    'DR_PRES7 As Integer         ' 炉内圧
    DR_AR7   As String * 7       ' アルゴン流量      '2003/05/16 osawa
    'DR_AR7   As Integer         ' アルゴン流量
    UPD_DATE  As Date          ' 更新完了日付
    EXT_DATE  As Date          ' 抽出完了日付
    DR_AR3   As Integer        ' アルゴン№３流量
    DR_DOP   As Integer        ' ドープ
End Type

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :VAX側テーブル「DR_CNDS」から条件にあったレコードを抽出する（複数レコード）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型              ,説明
'          :records()     ,O  ,typ_VAX_DR_CNDS ,抽出レコード
'          :sqlWhere      ,I  ,String          ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String          ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/10/03 作成　蔵本
Public Function DBDRV_VAX_DR_CNDS(records() As typ_VAX_DR_CNDS _
                                   , Optional sqlWhere$ = vbNullString _
                                   , Optional sqlOrder$ = vbNullString _
                                   ) As FUNCTION_RETURN
                                   
    Dim sql As String       'SQL全体
    Dim recCnt As Long      'レコード数
    Dim i As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset


    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_VAX_SQL.bas -- Function DBDRV_VAX_DR_CNDS"
    
    DBDRV_VAX_DR_CNDS = FUNCTION_RETURN_FAILURE
    
    sql = "select "
    sql = sql & "PG_ID, "       ' プログラムID
    sql = sql & "DR_CHRG, "     ' チャージ量
    sql = sql & "DR_CPOS, "     ' ルツボ位置
    sql = sql & "DR_CSIZ, "     ' ルツボサイズ
    sql = sql & "DR_DIA, "      ' 直径
    sql = sql & "DR_LEN0, "     ' 引上長(1本引き/R0)
    sql = sql & "DR_LEN1, "     ' 引上長(R1)
    sql = sql & "DR_SR, "       ' 上軸回転数
    sql = sql & "DR_CR, "       ' 下軸回転数
    sql = sql & "DR_GAP, "      ' ギャップ
    sql = sql & "DR_PRES7, "     ' 炉内圧
    sql = sql & "DR_AR7,  "      ' アルゴン流量
    sql = sql & "DR_AR3,  "     ' アルゴン№３流量
    sql = sql & "DR_DOP,  "     ' ドープ
    sql = sql & "UPD_DATE, "    ' 更新完了日付
    sql = sql & "EXT_DATE "     ' 抽出完了日付
    sql = sql & "from DR_CNDS "
    
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & sqlWhere & sqlOrder
    End If
Debug.Print sql

    Set db = DBEngine.Workspaces(0).OpenDatabase("VAX", dbDriverComplete, True, "ODBC;DATABASE=attach 'filename disk$xtal:[usr.rdb]xtal';UID=xtal;PWD=crystal;DSN=VAX")
    Set rs = db.OpenRecordset(sql)
        
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_VAX_DR_CNDS = FUNCTION_RETURN_FAILURE
        rs.Close
        db.Close
        Set rs = Nothing
        Set db = Nothing
        GoTo proc_exit
    End If
    
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .PG_ID = rs("PG_ID")         ' プログラムID
            .DR_CHRG = IIf(rs("DR_CHRG") <> "", rs("DR_CHRG"), 0) ' チャージ量
            .DR_CPOS = IIf(rs("DR_CPOS") <> "", rs("DR_CPOS"), 0) ' ルツボ位置
            .DR_CSIZ = IIf(rs("DR_CSIZ") <> "", rs("DR_CSIZ"), 0)  ' ルツボサイズ
            .DR_DIA = IIf(rs("DR_DIA") <> "", rs("DR_DIA"), 0)     ' 直径
            .DR_LEN0 = IIf(rs("DR_LEN0") <> "", rs("DR_LEN0"), 0)  ' 引上長(1本引き/R0)
            .DR_LEN1 = IIf(rs("DR_LEN1") <> "", rs("DR_LEN1"), 0)  ' 引上長(R1)
            .DR_SR = IIf(rs("DR_SR") <> "", Trim(rs("DR_SR")), "0") ' 上軸回転数
            .DR_CR = IIf(rs("DR_CR") <> "", Trim(rs("DR_CR")), "0") ' 下軸回転数
            .DR_GAP = IIf(rs("DR_GAP") <> "", rs("DR_GAP"), 0)             ' ギャップ
            .DR_PRES7 = IIf(rs("DR_PRES7") <> "", Trim(rs("DR_PRES7")), "0")      ' 炉内圧
            '.DR_PRES7 = IIf(rs("DR_PRES7") <> "", rs("DR_PRES7"), 0)      ' 炉内圧
            .DR_AR7 = IIf(rs("DR_AR7") <> "", Trim(rs("DR_AR7")), "0")             ' アルゴン流量
            .DR_AR3 = IIf(rs("DR_AR3") <> "", rs("DR_AR3"), " ")       ' アルゴン№３流量　4/30
            .DR_DOP = IIf(rs("DR_DOP") <> "", rs("DR_DOP"), " ")       ' ドープ
            .UPD_DATE = IIf(rs("UPD_DATE") <> "", rs("UPD_DATE"), Now) ' 更新完了日付
            .EXT_DATE = IIf(rs("EXT_DATE") <> "", rs("EXT_DATE"), Now) ' 抽出完了日付
            
            '桁あふれしないようにチェック
            If .DR_CHRG < -9999 Or .DR_CHRG > 9999 Then
                .DR_CHRG = 9999
            End If
            If .DR_CPOS < -999 Or .DR_CPOS > 999 Then
                .DR_CPOS = 999
            End If
            If .DR_CSIZ < -99 Or .DR_CSIZ > 99 Then
                .DR_CSIZ = 99
            End If
            If .DR_DIA < -999 Or .DR_DIA > 999 Then
                .DR_DIA = 999
            End If
            If .DR_LEN0 < -9999 Or .DR_LEN0 > 9999 Then
                .DR_LEN0 = 9999
            End If
            If .DR_LEN1 < -9999 Or .DR_LEN1 > 9999 Then
                .DR_LEN1 = 9999
            End If
            If .DR_GAP < -999 Or .DR_GAP > 999 Then
                .DR_GAP = 999
            End If
        End With
        rs.MoveNext
    Next

'    Do While Not rs.EOF
'        Debug.Print rs.Fields("PG_ID")
'        Debug.Print rs.Fields(3)
'    Loop
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
    
    DBDRV_VAX_DR_CNDS = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_VAX_DR_CNDS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :VAX側テーブル「DR_CNDS」から条件にあったレコードを抽出する(1レコード)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型              ,説明
'          :record        ,O  ,typ_VAX_DR_CNDS ,抽出レコード
'          :sqlWhere      ,I  ,String          ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String          ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/10/03 作成　蔵本
Public Function DBDRV_VAX_DR_CNDS1(record As typ_VAX_DR_CNDS _
                                   , Optional sqlWhere$ = vbNullString _
                                   , Optional sqlOrder$ = vbNullString _
                                   ) As FUNCTION_RETURN
                                   
    Dim sql As String       'SQL全体
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_VAX_SQL.bas -- Function DBDRV_VAX_DR_CNDS1"
    
    DBDRV_VAX_DR_CNDS1 = FUNCTION_RETURN_FAILURE
    
    sql = "select "
    sql = sql & "PG_ID, "       ' プログラムID
    sql = sql & "DR_CHRG, "     ' チャージ量
    sql = sql & "DR_CPOS, "     ' ルツボ位置
    sql = sql & "DR_CSIZ, "     ' ルツボサイズ
    sql = sql & "DR_DIA, "      ' 直径
    sql = sql & "DR_LEN0, "     ' 引上長(1本引き/R0)
    sql = sql & "DR_LEN1, "     ' 引上長(R1)
    sql = sql & "DR_SR, "       ' 上軸回転数
    sql = sql & "DR_CR, "       ' 下軸回転数
    sql = sql & "DR_GAP, "      ' ギャップ
    sql = sql & "DR_PRES7, "     ' 炉内圧
    sql = sql & "DR_AR7,  "      ' アルゴン流量
    sql = sql & "DR_AR3,  "      ' アルゴン№３流量   4/30
    sql = sql & "DR_DOP,  "      ' ドープ
    sql = sql & "UPD_DATE, "    ' 更新完了日付
    sql = sql & "EXT_DATE "     ' 抽出完了日付
    sql = sql & "from DR_CNDS "
    
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & sqlWhere & sqlOrder
    End If

    Set db = DBEngine.Workspaces(0).OpenDatabase("VAX", dbDriverComplete, True, "ODBC;DATABASE=attach 'filename disk$xtal:[usr.rdb]xtal';UID=xtal;PWD=crystal;DSN=VAX")
    Set rs = db.OpenRecordset(sql)
        
    If rs Is Nothing Then
        rs.Close
        db.Close
        Set rs = Nothing
        Set db = Nothing
        DBDRV_VAX_DR_CNDS1 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    'レコード０件時はPGIDをNullにして返す。
    If rs.RecordCount = 0 Then
        record.PG_ID = vbNullString
    Else
        With record
            .PG_ID = rs("PG_ID")         ' プログラムID
            .DR_CHRG = IIf(rs("DR_CHRG") <> "", rs("DR_CHRG"), 0) ' チャージ量
            .DR_CPOS = IIf(rs("DR_CPOS") <> "", rs("DR_CPOS"), 0) ' ルツボ位置
            .DR_CSIZ = IIf(rs("DR_CSIZ") <> "", rs("DR_CSIZ"), 0)  ' ルツボサイズ
            .DR_DIA = IIf(rs("DR_DIA") <> "", rs("DR_DIA"), 0)     ' 直径
            .DR_LEN0 = IIf(rs("DR_LEN0") <> "", rs("DR_LEN0"), 0)  ' 引上長(1本引き/R0)
            .DR_LEN1 = IIf(rs("DR_LEN1") <> "", rs("DR_LEN1"), 0)  ' 引上長(R1)
            .DR_SR = IIf(rs("DR_SR") <> "", Trim(rs("DR_SR")), "") ' 上軸回転数
            .DR_CR = IIf(rs("DR_CR") <> "", Trim(rs("DR_CR")), "") ' 下軸回転数
            .DR_GAP = IIf(rs("DR_GAP") <> "", rs("DR_GAP"), 0)             ' ギャップ
            .DR_PRES7 = IIf(rs("DR_PRES7") <> "", rs("DR_PRES7"), 0)      ' 炉内圧
            .DR_AR7 = IIf(rs("DR_AR7") <> "", rs("DR_AR7"), 0)            ' アルゴン流量
            .DR_AR3 = IIf(rs("DR_AR3") <> "", rs("DR_AR3"), " ")         ' アルゴン№３流量　　4/30
            .DR_DOP = IIf(rs("DR_DOP") <> "", rs("DR_DOP"), " ")         ' ドープ
            .UPD_DATE = IIf(rs("UPD_DATE") <> "", rs("UPD_DATE"), Now) ' 更新完了日付
            .EXT_DATE = IIf(rs("EXT_DATE") <> "", rs("EXT_DATE"), Now) ' 抽出完了日付
            
            '桁あふれしないようにチェック
            If .DR_CHRG < -9999 Or .DR_CHRG > 9999 Then
                .DR_CHRG = 9999
            End If
            If .DR_CPOS < -999 Or .DR_CPOS > 999 Then
                .DR_CPOS = 999
            End If
            If .DR_CSIZ < -99 Or .DR_CSIZ > 99 Then
                .DR_CSIZ = 99
            End If
            If .DR_DIA < -999 Or .DR_DIA > 999 Then
                .DR_DIA = 999
            End If
            If .DR_LEN0 < -9999 Or .DR_LEN0 > 9999 Then
                .DR_LEN0 = 9999
            End If
            If .DR_LEN1 < -9999 Or .DR_LEN1 > 9999 Then
                .DR_LEN1 = 9999
            End If
            If .DR_GAP < -999 Or .DR_GAP > 999 Then
                .DR_GAP = 999
            End If

        End With
    End If

    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing

    DBDRV_VAX_DR_CNDS1 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_VAX_DR_CNDS1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


