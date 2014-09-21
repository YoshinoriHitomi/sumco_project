Attribute VB_Name = "s_cmbc007_SQL"
Option Explicit

'原料在庫修正


'原料在庫取得用
Public Type type_DBDRV_scmzc_fcmgc001e_Disp
    '原料在庫管理
    MTRLNUM As String * 10      ' 原料番号
    WEIGHT As Long              ' 重量
End Type


'原料在庫更新用
Public Type type_DBDRV_scmzc_fcmgc001e_Exec
    '原料在庫管理
    MTRLNUM As String * 10      ' 原料番号
    USABLCLS As String * 1      ' 使用可能区分
    KRPROCCD As String * 5      ' 管理工程コード
    PROCCODE As String * 5      ' 工程コード
    KSTAFFID As String * 8      ' 更新社員ID
    WEIGHT As Long              ' 新重量
    SYORIW As Long              ' 処理量

End Type



'初期表示
'概要    :原料在庫修正 表示用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                    ,説明
'        :record       ,I   ,type_DBDRV_scmzc_fcmgc001e_Disp       ,原料在庫取得用
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                       ,読み込み成否
'説明    :
'履歴    :2001/06/18 蔵本 作成
Public Function DBDRV_scmzc_fcmgc001e_Disp(records() As type_DBDRV_scmzc_fcmgc001e_Disp) As FUNCTION_RETURN
    
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Integer
    Dim i As Long
    
    '原料管理テーブルで使用可能区分１をselect（原料番号、重量）
    

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmgc001e_SQL.bas -- Function DBDRV_scmzc_fcmgc001e_Disp"

    sql = "Select MTRLNUM, USABLCLS, WEIGHT, TSTAFFID, REGDATE, KSTAFFID, UPDDATE "
    sql = sql & "From TBCMG005"
    
        
        sql = "select MTRLNUM, WEIGHT"
        sql = sql & " from ( "
        sql = sql & " select MTRLNUM, WEIGHT"
        sql = sql & " from TBCMG005"
        sql = sql & " where USABLCLS='1'"
        sql = sql & " and WEIGHT > 0 "
        sql = sql & " and substr(MTRLNUM,1,1) not in ('P','N')"
        sql = sql & " order by MTRLNUM ) "
        sql = sql & " union all "
        sql = sql & " select MTRLNUM, WEIGHT"
        sql = sql & " from ( "
        sql = sql & " select MTRLNUM, WEIGHT"
        sql = sql & " from TBCMG005"
        sql = sql & " where USABLCLS='1'"
        sql = sql & " and WEIGHT > 0 "
        sql = sql & " and substr(MTRLNUM,1,1) in ('P','N')"
        sql = sql & " order by MTRLNUM )"


    '   order by 原料番号
    '   substr(原料番号,1,1) not in ('P','N')
    '   union all
    
    'select ...
    '   order by 原料番号
    '   substr(原料番号,1,1) in ('P','N')
    
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    recCnt = rs.RecordCount
    ReDim records(recCnt)
    If recCnt = 0 Then ''2001/07/17 Sano
'2001/07/17 Sano    If rs.RecordCount = 0 Then
        DBDRV_scmzc_fcmgc001e_Disp = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If
    
'2001/07/17 Sano    recCnt = rs.RecordCount
'2001/07/17 Sano    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .MTRLNUM = rs("MTRLNUM")          ' 原料番号
            .WEIGHT = rs("WEIGHT")            ' 重量
        End With
        rs.MoveNext
    Next i
    rs.Close

    DBDRV_scmzc_fcmgc001e_Disp = FUNCTION_RETURN_SUCCESS
   

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmgc001e_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'実行時
'概要    :原料在庫修正 更新、挿入用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                    ,説明
'        :record       ,I   ,type_DBDRV_scmzc_fcmgc001e_Exec       ,原料在庫挿入用
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                       ,読み込み成否
'説明    :
'履歴    :2001/06/18 蔵本 作成
Public Function DBDRV_scmzc_fcmgc001e_Exec(record As type_DBDRV_scmzc_fcmgc001e_Exec) As FUNCTION_RETURN
    
    Dim sql As String

    
    '原料管理テーブルを新重量に更新
        

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmgc001e_SQL.bas -- Function DBDRV_scmzc_fcmgc001e_Exec"

    DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_SUCCESS
    
    sql = "update TBCMG005 set "
    With record
        sql = sql & "WEIGHT=" & .WEIGHT & ", "               ' 重量
        sql = sql & "KSTAFFID='" & .KSTAFFID & "', "         ' 更新社員ID
        sql = sql & "UPDDATE=sysdate "                       ' 更新日付
        sql = sql & "where MTRLNUM='" & .MTRLNUM & "' "
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    
    '原料在庫実績に挿入
    sql = " insert into TBCMG006 ( "
    sql = sql & "MTRLNUM, "          ' 原料番号
    sql = sql & "TRANCNT, "          ' 処理回数
    sql = sql & "KRPROCCD, "         ' 管理工程コード
    sql = sql & "PROCCODE, "         ' 工程コード
    sql = sql & "CLASS, "            ' 区分
    sql = sql & "INWEIGHT, "         ' 入力重量
    sql = sql & "TSTAFFID, "         ' 登録社員ID
    sql = sql & "REGDATE, "          ' 登録日付
    sql = sql & "KSTAFFID, "         ' 更新社員ID
    sql = sql & "UPDDATE, "          ' 更新日付
    sql = sql & "SENDFLAG, "         ' 送信フラグ
    sql = sql & "SENDDATE ) "        ' 送信日付
    With record
        sql = sql & " select "
        sql = sql & " '" & .MTRLNUM & "', "          ' 原料番号
        sql = sql & " nvl(max(TRANCNT),0)+1, "       ' 処理回数
        sql = sql & " '" & .KRPROCCD & "', "         ' 管理工程コード
        sql = sql & " '" & .PROCCODE & "', "         ' 工程コード
        sql = sql & " '" & .USABLCLS & "', "         ' 区分
        sql = sql & " '" & .SYORIW & "', "           ' 入力重量
        sql = sql & " '" & .KSTAFFID & "', "         ' 登録社員ID
        sql = sql & " sysdate, "                     ' 登録日付
        sql = sql & " '" & .KSTAFFID & "', "         ' 更新社員ID
        sql = sql & " sysdate, "                     ' 更新日付
        sql = sql & " '0', "                         ' 送信フラグ
        sql = sql & " sysdate "                      ' 送信日付
        sql = sql & " from TBCMG006 "
        sql = sql & " where MTRLNUM='" & .MTRLNUM & "' "
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
        

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

