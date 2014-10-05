Attribute VB_Name = "s_cmbc001b_SQL"
Option Explicit

' TBCME017 (製品仕様管理)より
Public Type s_cmzcF_cmfc001b_Disp
    '製品仕様管理
    Hinban12 As String * 12         ' 品番
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    REGDATE As Date                 ' 登録日付
End Type


Public Function DBDRV_s_cmzcF_cmfc001b_Disp(records() As s_cmzcF_cmfc001b_Disp) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001b_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001b_Disp"
    
    ''製品仕様管理があってSXL製作条件がないレコード取得（品番、仕様登録依頼番号、登録日付）
    ''ただし、製作条件付与取消にあるレコードは除く
    sql = "select hinban||ltrim(to_char(mnorevno,'00'))||factory||opecond as hinban12, HMGSTRRNO, REGDATE " & _
          "From tbcme018 " & _
          "where (opecond='1') and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme030) and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme031)"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_s_cmzcF_cmfc001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Hinban12 = rs("HINBAN12")        ' 品番
            .HMGSTRRNO = rs("HMGSTRRNO")    ' 品管理仕様登録依頼番号
            .REGDATE = rs("REGDATE")        ' 登録日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_s_cmzcF_cmfc001b_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'この画面ではExec()はいらない
'Public Function DBDRV_s_cmzcF_cmgc001d_Exec(s_cmzcF_cmfc001a_Disp As type_DBDRV_s_cmzcF_cmgc001d_Exec) As FUNCTION_RETURN
'    s_cmzcF_cmgc001c_Exec = FUNCTION_RETURN_SUCCESS
'
'    'リメルト洗浄払出実績テーブルに原料番号()、管理工程コード()、工程コード()、乾燥後重量()、ロス重量、社員ＩＤをインサート
'
'End Function

