Attribute VB_Name = "SB_cmzcXSDCE"
Option Explicit
'                                     2003/09/05
'======================================================
' ユーザ定義型の宣言
'======================================================

' 振替履歴
Public Type typ_XSDCE
    CRYNUMCE As String * 12         ' ブロックID・結晶番号
    INPOSCE As Integer              ' 結晶内開始位置
    KCNTCE As Integer               ' 工程連番
    HINBCE As String * 8            ' 振替先品番
    REVNUMCE As Integer             ' 製品番号改訂番号(振替先)
    FACTORYCE As String * 1         ' 工場(振替先)
    OPECE As String * 1             ' 操業条件(振替先)
    MOTHINCE As String * 8          ' 振替元品番
    MREVNUMCE As Integer            ' 製品番号改訂番号(振替元)
    MFACTORYCE As String * 1        ' 工場(振替元)
    MOPECE As String * 1            ' 操業条件(振替元)
    SXLIDCE As String * 13          ' SXLID
    WKKTCE As String * 5            ' 工程
    KNKTCE As String * 5            ' 管理工程
    REPSMPLIDTCE As String * 16     ' 代表サンプルID(TOP)
    REPSMPLIDBCE As String * 16     ' 代表サンプルID(BOT)
    TOKNUMCE As String * 10         ' 特採番号
    TOKCAUSECE As String * 200      ' 特採理由
    TOKCODECE As String * 2         ' 特採理由コード
    ERRCAUSECE As String * 50       ' エラー理由
    HULCE As Integer                ' 振替長さ
    HUWCE As Long                   ' 振替重量
    HUMCE As Integer                ' 振替枚数
    TSTAFFCE As String * 8          ' 登録社員ID
    TDAYCE As Date                  ' 登録日付
    KSTAFFCE As String * 8          ' 更新社員ID
    KDAYCE As Date                  ' 更新日付
    SNDKCE As String * 1            ' 送信フラグ
    SNDDAYCE As Date                ' 送信日付
End Type

'更新用
Public Type typ_XSDCE_Update
    CRYNUMCE As String              ' ブロックID・結晶番号
    INPOSCE As String               ' 結晶内開始位置
    KCNTCE As String                ' 工程連番
    HINBCE As String                ' 振替先品番
    REVNUMCE As String              ' 製品番号改訂番号(振替先)
    FACTORYCE As String             ' 工場(振替先)
    OPECE As String                 ' 操業条件(振替先)
    MOTHINCE As String              ' 振替元品番
    MREVNUMCE As String             ' 製品番号改訂番号(振替元)
    MFACTORYCE As String            ' 工場(振替元)
    MOPECE As String                ' 操業条件(振替元)
    SXLIDCE As String               ' SXLID
    WKKTCE As String                ' 工程
    KNKTCE As String                ' 管理工程
    REPSMPLIDTCE As String          ' 代表サンプルID(TOP)
    REPSMPLIDBCE As String          ' 代表サンプルID(BOT)
    TOKNUMCE As String              ' 特採番号
    TOKCAUSECE As String            ' 特採理由
    TOKCODECE As String             ' 特採理由コード
    ERRCAUSECE As String            ' エラー理由
    HULCE As String                 ' 振替長さ
    HUWCE As String                 ' 振替重量
    HUMCE As String                 ' 振替枚数
    TSTAFFCE As String              ' 登録社員ID
    TDAYCE As String                ' 登録日付
    KSTAFFCE As String              ' 更新社員ID
    KDAYCE As String                ' 更新日付
    SNDKCE As String                ' 送信フラグ
    SNDDAYCE As String              ' 送信日付
End Type

'概要      :工程連番を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_sCrynum    ,I  ,String           ,ブロックID・結晶番号
'      　　:p_sGenKotei  ,I  ,String           ,現在工程
'      　　:戻り値       ,O  ,Integer        　,工程連番
'説明      :結晶番号と工程から工程連番を取得する
Public Function GetKCNTC3(p_sCrynum As String, p_sGenKotei As String) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT MAX(KCNTC3) AS MAXKCNTC3 FROM XSDC3 "
    sql = sql & "WHERE CRYNUMC3 = '" & p_sCrynum & "' "
    sql = sql & "AND   WKKTC3   = '" & p_sGenKotei & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("MAXKCNTC3")) Then
        GetKCNTC3 = 0
    Else
        GetKCNTC3 = CInt(rs.Fields("MAXKCNTC3"))
    End If
End Function

'概要      :工程連番を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_sRpCrynum  ,I  ,String           ,親ブロックID
'      　　:p_sGenKotei  ,I  ,String           ,現在工程
'      　　:戻り値       ,O  ,Integer        　,工程連番
'説明      :親ブロックIDと工程から工程連番を取得する
'履歴      :05/09/16 ooba
Public Function GetKCNTC3_New(p_sRpCrynum As String, p_sGenKotei As String) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT MAX(KCNTC3) AS MAXKCNTC3 FROM XSDC3 "
    sql = sql & "WHERE RPCRYNUMC3 = '" & p_sRpCrynum & "' "
    sql = sql & "AND   WKKTC3   = '" & p_sGenKotei & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("MAXKCNTC3")) Then
        GetKCNTC3_New = 0
    Else
        GetKCNTC3_New = CInt(rs.Fields("MAXKCNTC3"))
    End If
End Function

'概要      :結晶内開始位置を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_sCrynum    ,I  ,String           ,ブロックID・結晶番号
'      　　:p_iInpos     ,I  ,Integer          ,結晶内開始位置
'      　　:p_sKcnt      ,I  ,Integer          ,工程連番
'      　　:戻り値       ,O  ,Integer        　,結晶内開始位置
'説明      :結晶番号と工程連番から結晶内開始位置を取得する
Public Function GetINPOSC3(p_sCrynum As String, p_iInpos As Integer, p_iKcnt As Integer) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT INPOSC3 FROM XSDC3 "
    sql = sql & "WHERE CRYNUMC3 = '" & p_sCrynum & "' "
    sql = sql & "AND   INPOSC3  < " & p_iInpos & " "
    sql = sql & "AND   KCNTC3   = " & p_iKcnt & " "
    sql = sql & "ORDER BY INPOSC3 desc "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("INPOSC3")) Then
        GetINPOSC3 = -1
    Else
        GetINPOSC3 = CInt(rs.Fields("INPOSC3"))
    End If
End Function

'●INSERT●  NULLの場合、charならスペース、NumberならNULLを入れる

'概要      :テーブル「XSDCE」にレコードを挿入する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型                ,説明
'      　　:pXSDCE 　　  ,I  ,typ_XSDCE_Update  ,XSDCE更新用ﾃﾞｰﾀ
'      　　:sErrMsg　　　,O  ,String         　 ,エラーメッセージ
'      　　:戻り値       ,O  ,Boolean        　 ,True:OK False:NG
'説明      :履歴管理ＤＢの登録を行う
Public Function CreateXSDCE(pXSDCE As typ_XSDCE_Update, sErrMsg As String) As Boolean
    
    Dim sql As String
    Dim sDbName As String
'    Dim rs As OraDynaset    'RecordSet
    Dim nowtime As Date
    Dim nowtime_sql As String       ''サーバ時間(SQL文)
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDCE.bas -- Function CreateXSDCE"
    sErrMsg = ""
    sDbName = "XSDCE"
    nowtime = getSvrTime()    'サーバーの時間を取得するように変更 2003/6/4 tuku
    
'>>>>> .AddNewをSQL(INSERT)文に変更　2009/06/16 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    With pXSDCE
        sql = "INSERT INTO XSDCE ( "
        sql = sql & " CRYNUMCE"         ' 1:ﾌﾞﾛｯｸID・結晶番号
        sql = sql & ",INPOSCE"          ' 2:結晶内開始位置
        sql = sql & ",KCNTCE"           ' 3:工程連番
        sql = sql & ",HINBCE"           ' 4:振替先品番
        sql = sql & ",REVNUMCE"         ' 5:製品番号改訂番号(振替先)
        sql = sql & ",FACTORYCE"        ' 6:工場(振替先)
        sql = sql & ",OPECE"            ' 7:操業条件(振替先)
        sql = sql & ",MOTHINCE"         ' 8:振替元品番
        sql = sql & ",MREVNUMCE"        ' 9:製品番号改訂番号(振替元)
        sql = sql & ",MFACTORYCE"       '10:工場(振替元)
        sql = sql & ",MOPECE"           '11:操業条件(振替先)
        sql = sql & ",SXLIDCE"          '12:SXLID
        sql = sql & ",WKKTCE"           '13:工程
        sql = sql & ",KNKTCE"           '14:管理工程
        sql = sql & ",REPSMPLIDTCE"     '15:代表サンプルID(TOP)
        sql = sql & ",REPSMPLIDBCE"     '16:代表サンプルID(BOT)
        sql = sql & ",TOKNUMCE"         '17:特採番号
        sql = sql & ",TOKCAUSECE"       '18:特採理由
        sql = sql & ",TOKCODECE"        '19:特採理由コード
        sql = sql & ",ERRCAUSECE"       '20:エラー理由
        sql = sql & ",HULCE"            '21:振替長さ
        sql = sql & ",HUWCE"            '22:振替重量
        sql = sql & ",HUMCE"            '23:振替枚数
        sql = sql & ",TSTAFFCE"         '24:登録社員ID
        sql = sql & ",TDAYCE"           '25:登録日付
        sql = sql & ",KSTAFFCE"         '26:更新社員ID
        sql = sql & ",KDAYCE"           '27:更新日付
        sql = sql & ",SNDKCE"           '28:送信フラグ
        sql = sql & ",SNDDAYCE"         '29:送信日付
        sql = sql & ") "
        sql = sql & "VALUES ( "
        
        ' 1:ﾌﾞﾛｯｸID・結晶番号
        If .CRYNUMCE <> "" Then
            sql = sql & " '" & .CRYNUMCE & "'"
        Else
            sql = sql & " '" & Space(12) & "'"
        End If
        
        ' 2:結晶内開始位置
        If .INPOSCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.INPOSCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        ' 3:工程連番
        If .KCNTCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.KCNTCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        ' 4:振替先品番
        If .HINBCE <> "" Then
            sql = sql & ",'" & .HINBCE & "'"
        Else
            sql = sql & ",'" & Space(8) & "'"
        End If
        
        ' 5:製品番号改訂番号(振替先)
        If .REVNUMCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.REVNUMCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        ' 6:工場(振替先)
        If .FACTORYCE <> "" Then
            sql = sql & ",'" & .FACTORYCE & "'"
        Else
            sql = sql & ",'" & Space(1) & "'"
        End If
        
        ' 7:操業条件(振替先)
        If .OPECE <> "" Then
            sql = sql & ",'" & .OPECE & "'"
        Else
            sql = sql & ",'" & Space(1) & "'"
        End If
        
        ' 8:振替元品番
        If .MOTHINCE <> "" Then
            sql = sql & ",'" & .MOTHINCE & "'"
        Else
            sql = sql & ",'" & Space(8) & "'"
        End If
        
        ' 9:製品番号改訂番号(振替元)
        If .MREVNUMCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.MREVNUMCE)) & "'"
        Else
            sql = sql & ",'0'"
        End If
        
        '10:工場(振替元)
        If .MFACTORYCE <> "" Then
            sql = sql & ",'" & .MFACTORYCE & "'"
        Else
            sql = sql & ",'" & Space(1) & "'"
        End If
        
        '11:操業条件(振替先)
        If .MOPECE <> "" Then
            sql = sql & ",'" & .MOPECE & "'"
        Else
            sql = sql & ",'" & Space(1) & "'"
        End If
        
        '12:SXLID
        If .SXLIDCE <> "" And Left(.SXLIDCE, 1) <> vbNullChar Then
            sql = sql & ",'" & .SXLIDCE & "'"
        Else
            sql = sql & ",'" & Space(13) & "'"
        End If
        
        '13:工程
        If .WKKTCE <> "" Then
            sql = sql & ",'" & .WKKTCE & "'"
        Else
            sql = sql & ",'" & Space(5) & "'"
        End If
        
        '14:管理工程
        If .KNKTCE <> "" Then
            sql = sql & ",'" & .KNKTCE & "'"
        Else
            sql = sql & ",'" & Space(5) & "'"
        End If
        
        '15:代表サンプルID(TOP)
        If .REPSMPLIDTCE <> "" Then
            sql = sql & ",'" & .REPSMPLIDTCE & "'"
        Else
            sql = sql & ",'" & Space(16) & "'"
        End If
        
        '16:代表サンプルID(BOT)
        If .REPSMPLIDBCE <> "" Then
            sql = sql & ",'" & .REPSMPLIDBCE & "'"
        Else
            sql = sql & ",'" & Space(16) & "'"
        End If
        
        '17:特採番号
        If .TOKNUMCE <> "" Then
            sql = sql & ",'" & .TOKNUMCE & "'"
        Else
            sql = sql & ",'" & Space(10) & "'"
        End If
        
        '18:特採理由
        If .TOKCAUSECE <> "" Then
            sql = sql & ",'" & .TOKCAUSECE & "'"
        Else
            sql = sql & ",NULL"
        End If
        
        '19:特採理由コード
        If .TOKCODECE <> "" Then
            sql = sql & ",'" & .TOKCODECE & "'"
        Else
            sql = sql & ",'" & Space(2) & "'"
        End If
        
        '20:エラー理由
        If .ERRCAUSECE <> "" Then
            sql = sql & ",'" & .ERRCAUSECE & "'"
        Else
            sql = sql & ",NULL"
        End If
        
        '21:振替長さ
        If .HULCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.HULCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        '22:振替重量
        If .HUWCE <> "" Then
            sql = sql & ",'" & CStr(CLng(.HUWCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        '23:振替枚数
        If .HUMCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.HUMCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        '24:登録社員ID
        If .TSTAFFCE <> "" Then
            sql = sql & ",'" & .TSTAFFCE & "'"
        Else
            sql = sql & ",'" & Space(8) & "'"
        End If
        
        '25:登録日付
        sql = sql & "," & nowtime_sql
        
        '26:更新社員ID
        If .KSTAFFCE <> "" Then
            sql = sql & ",'" & .KSTAFFCE & "'"
        Else
            sql = sql & ",'" & Space(8) & "'"
        End If
        
        '27:更新日付
        sql = sql & "," & nowtime_sql
        
        '28:送信フラグ
        If .SNDKCE <> "" Then
            sql = sql & ",'" & .SNDKCE & "'"
        Else
            sql = sql & ",'" & Space(1) & "'"
        End If
        
        '29:送信日付
        If .SNDDAYCE <> "" Then
            sql = sql & ",TO_DATE('" & Format$(CDate(.SNDDAYCE), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
        Else
            sql = sql & ",NULL"
        End If
        
        sql = sql & ")"
    
        'SQLを実行
        If OraDB.ExecuteSQL(sql) < 1 Then
            GoTo proc_err
        End If
    
    End With
'<<<<< .AddNewをSQL(INSERT)文に変更　2009/06/16 SETsw kubota ------------------

    CreateXSDCE = True

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
    CreateXSDCE = False
    Resume proc_exit

End Function

