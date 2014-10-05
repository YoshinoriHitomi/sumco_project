Attribute VB_Name = "s_cmbc004_SQL"
Option Explicit
'
'' 多結晶受入棚入処理



Public Type type_DBDRV_scmzc_fcmgc001b_Exec
    ' 多結晶受入実績挿入用
    KRPROCCD As String * 5      ' 管理工程コード
    PROCCODE As String * 5      ' 工程コード
    TSTAFFID As String * 8      ' 登録社員ID
    MTRLTYPE As String * 3      ' 原料種類
    MAKERNO As String * 6       ' メーカ管理No
    RVWEIGHT As Double          ' 受入購入重量
    CRYCOMMENT As String        ' コメント
    WEIGHT    As Double         ' 本当の受入量
    
End Type

Public Type type_DBDRV_scmzc_fcmgc001b_Weight
    ' 多結晶受入仕掛り重量抽出用
    MTRL As String '* 10     ' 原料種類
    WEIGHT As Double            ' 受入購入重量
End Type


'概要    :多結晶受入棚入処理 更新、挿入用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                    ,説明
'        :record       ,I   ,type_DBDRV_scmzc_fcmgc001b_Exec       ,多結晶受入実績挿入用
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                       ,読み込み成否
'説明    :
'履歴    :2001/06/18 蔵本 作成
Public Function DBDRV_scmzc_fcmgc001b_Exec(record As type_DBDRV_scmzc_fcmgc001b_Exec) As FUNCTION_RETURN

    Dim sql As String
    Dim MTRLNUM As String
    Dim rs As OraDynaset


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmgc001b_SQL.bas -- Function DBDRV_scmzc_fcmgc001b_Exec"

    DBDRV_scmzc_fcmgc001b_Exec = FUNCTION_RETURN_SUCCESS


    '原料番号
    MTRLNUM = record.MTRLTYPE & record.MAKERNO & "0"

    '多結晶受入実績テーブルへ挿入
    sql = " insert into TBCMG001 ( "
    sql = sql & "MTRLNUM, "          ' 原料番号
    sql = sql & "JDATE, "            ' 日付
'    sql = sql & "TRANCNT, "          ' 処理回数
    sql = sql & "KRPROCCD, "         ' 管理工程コード
    sql = sql & "PROCCODE, "         ' 工程コード
    sql = sql & "MTRLTYPE, "         ' 原料種類
    sql = sql & "MAKERNO, "          ' メーカ管理No
    sql = sql & "RVWEIGHT, "         ' 受入購入重量
    sql = sql & "CRYCOMMENT, "       ' コメント
    sql = sql & "TSTAFFID, "         ' 登録社員ID
    sql = sql & "REGDATE, "          ' 登録日付
    sql = sql & "KSTAFFID, "         ' 更新社員ＩＤ
    sql = sql & "UPDDATE, "          ' 更新日付
    sql = sql & "SENDFLAG, "         ' 送信フラグ
    sql = sql & "SENDDATE) "         ' 送信日付
    With record
        sql = sql & " values ( "
        sql = sql & " '" & MTRLNUM & "', "           ' 原料番号
        sql = sql & " sysdate, "                     ' 日付       sysdateに変更予定#kk#
'        sql = sql & " nvl(max(TRANCNT),0)+1, "       ' 処理回数      はなくなる#kk#
        sql = sql & " '" & .KRPROCCD & "', "         ' 管理工程コード
        sql = sql & " '" & .PROCCODE & "', "         ' 工程コード
        sql = sql & " '" & .MTRLTYPE & "', "         ' 原料種類
        sql = sql & " '" & .MAKERNO & "', "          ' メーカ管理No
        sql = sql & " " & .WEIGHT & ", "             ' 受入購入重量
        sql = sql & " '" & .CRYCOMMENT & "', "       ' コメント
        sql = sql & " '" & .TSTAFFID & "', "         ' 登録社員ID
        sql = sql & " sysdate, "                      ' 登録日付
        sql = sql & " '" & .TSTAFFID & "', "         ' 更新社員ＩＤ
        sql = sql & " sysdate, "                      ' 更新日付
        sql = sql & " '0', "                         ' 送信フラグ
        sql = sql & " sysdate ) "                      ' 送信日付
'        sql = sql & " from TBCMG001 "
'        sql = sql & " where MTRLNUM='" & MTRLNUM & "' "
    End With
    

    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmgc001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    '原料在庫管理の更新 or 挿入
    sql = " select "
    sql = sql & "count(MTRLNUM) as C "
    sql = sql & "from TBCMG005 "
    sql = sql & "where MTRLNUM='" & MTRLNUM & "' "
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    'レコードが無かったら挿入
    If rs("C") = 0 Then
        '原料在庫管理テーブルへの挿入
        sql = "insert into TBCMG005 ( "
        sql = sql & "MTRLNUM, "          ' 原料番号
        sql = sql & "USABLCLS, "         ' 使用可能区分
        sql = sql & "WEIGHT, "           ' 重量
        sql = sql & "TSTAFFID, "         ' 登録社員ID
        sql = sql & "REGDATE, "          ' 登録日付
        sql = sql & "KSTAFFID, "         ' 更新社員ID
        sql = sql & "UPDDATE ) "           ' 更新日付
        
        sql = sql & " values ( "
        sql = sql & " '" & MTRLNUM & "', "
        sql = sql & " '1', "
        sql = sql & " " & record.RVWEIGHT & ", "
        sql = sql & " '" & record.TSTAFFID & "', "   ' 登録社員ID
        sql = sql & " sysdate, "                      ' 登録日付
        sql = sql & " '" & record.TSTAFFID & "', "   ' 更新社員ＩＤ
        sql = sql & " sysdate )"                      ' 更新日付
    
    Else
    
        '原料在庫管理テーブルの更新
        sql = "update TBCMG005 set "
        sql = sql & "WEIGHT=" & record.RVWEIGHT & ", "           ' 重量
        sql = sql & "KSTAFFID='" & record.TSTAFFID & "', "         ' 更新社員ID
        sql = sql & "UPDDATE=sysdate "                                     ' 更新日付
        sql = sql & "where MTRLNUM='" & MTRLNUM & "' "
    
    End If
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
         DBDRV_scmzc_fcmgc001b_Exec = FUNCTION_RETURN_FAILURE
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
    DBDRV_scmzc_fcmgc001b_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'概要    :多結晶受入棚入処理 仕掛り重量抽出ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                    ,説明
'        :record       ,I   ,type_DBDRV_scmzc_fcmgc001b_Weight     ,多結晶受入仕掛り重量抽出用
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                       ,読み込み成否
'説明    :
'履歴    :2001/07/17 Sano 作成
Public Function DBDRV_scmzc_fcmgc001b_Weight(record As type_DBDRV_scmzc_fcmgc001b_Weight) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmgc001b_SQL.bas -- Function DBDRV_scmzc_fcmgc001b_Weight"

    DBDRV_scmzc_fcmgc001b_Weight = FUNCTION_RETURN_SUCCESS

    '原料在庫管理の更新 or 挿入
    sql = " select "
    sql = sql & "nvl(sum(nvl(WEIGHT,0)),0) as W "
    sql = sql & "from TBCMG005 "
    sql = sql & "where MTRLNUM like'" & record.MTRL & "%' "
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    record.WEIGHT = rs("W")
    rs.Close

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmgc001b_Weight = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

