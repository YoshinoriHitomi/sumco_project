Attribute VB_Name = "s_cmbc001e_SQL"
Option Explicit

' 2007/08/30 SPK Tsutsumi Add Start
Public Type typ_Mukesaki
    sMukeCode As String     '' 向先コード
    sMukeName As String     '' 向先名
End Type

Public s_MukesakiBase() As typ_Mukesaki
' 2007/08/30 SPK Tsutsumi Add End

' 製作条件メンテナンス

'概要      :製作条件メンテナンス 製作条件更新／挿入用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:sMkCondNo　　　,I  ,String         　,製作条件№
'      　　:pMkOld   　　　,I  ,typ_TBCMB012   　,製作条件オリジナル
'      　　:pMkNew   　　　,I  ,typ_TBCMB012   　,製作条件
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込み成否
'説明      :
'履歴      :2001/07/30 蔵本 作成
Public Function DBDRV_scmzc_fcmbc001e_UpdInsMkCond(sMkCondNo As String, pMkOld() As typ_TBCMB012, pMkNew As typ_TBCMB012) As FUNCTION_RETURN

    Dim sql As String
    Dim bFlag As Boolean
    Dim i As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_UpdInsMkCond"

    With pMkNew
        bFlag = False
        For i = 1 To UBound(pMkOld)
            If RTrim$(pMkOld(i).MKCONDNO) = RTrim$(sMkCondNo) Then
                bFlag = True
                Exit For
            End If
        Next i

        If bFlag = True Then
            '' 製作条件の更新
            sql = "update TBCMB012 set "
            sql = sql & "MKCONDNO='" & .MKCONDNO & "', "    ' 製作条件No.
            sql = sql & "MODEL='" & .MODEL & "', "          ' 機種
            sql = sql & "RTBSIZE='" & .RTBSIZE & "', "      ' ルツボサイズ
            sql = sql & "CHARGE='" & .CHARGE & "', "        ' チャージ量
            sql = sql & "HZTYPE='" & .HZTYPE & "', "        ' HZタイプ
            sql = sql & "UPSPDTYP='" & .UPSPDTYP & "', "    ' 引上げ速度タイプ
            sql = sql & "MAGTYPE='" & .MAGTYPE & "', "      ' 磁場タイプ
            sql = sql & "USECLS='0', "                      ' 使用区分
            sql = sql & "TSTAFFID='" & .TSTAFFID & "', "    ' 登録社員ID
            sql = sql & "REGDATE=sysdate, "                 ' 登録日付
            sql = sql & "KSTAFFID='" & .KSTAFFID & "', "    ' 更新社員ID
            sql = sql & "UPDDATE=sysdate, "                 ' 更新日付
            sql = sql & "SENDFLAG='0', "                    ' 送信フラグ
            sql = sql & "SENDDATE=sysdate"                  ' 送信日時
            sql = sql & " where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
        Else
            '' 製作条件の挿入
            sql = "insert into TBCMB012 ("
            sql = sql & "MKCONDNO, "        ' 製作条件No.
            sql = sql & "MODEL, "           ' 機種
            sql = sql & "RTBSIZE, "         ' ルツボサイズ
            sql = sql & "CHARGE, "          ' チャージ量
            sql = sql & "HZTYPE, "          ' HZタイプ
            sql = sql & "UPSPDTYP, "        ' 引上げ速度タイプ
            sql = sql & "MAGTYPE, "         ' 磁場タイプ
            sql = sql & "USECLS, "          ' 使用区分
            sql = sql & "TSTAFFID, "        ' 登録社員ID
            sql = sql & "REGDATE, "         ' 登録日付
            sql = sql & "KSTAFFID, "        ' 更新社員ID
            sql = sql & "UPDDATE, "         ' 更新日付
            sql = sql & "SENDFLAG, "        ' 送信フラグ
            sql = sql & "SENDDATE)"         ' 送信日時
            sql = sql & " values ('"
            sql = sql & .MKCONDNO & "', '"  ' 製作条件No.
            sql = sql & .MODEL & "', '"     ' 機種
            sql = sql & .RTBSIZE & "', '"   ' ルツボサイズ
            sql = sql & .CHARGE & "', '"    ' チャージ量
            sql = sql & .HZTYPE & "', '"    ' HZタイプ
            sql = sql & .UPSPDTYP & "', '"  ' 引上げ速度タイプ
            sql = sql & .MAGTYPE & "', "    ' 磁場タイプ
            sql = sql & "'0', '"            ' 使用区分
            sql = sql & .TSTAFFID & "', "   ' 登録社員ID
            sql = sql & "sysdate, '"        ' 登録日付
            sql = sql & .KSTAFFID & "', "   ' 更新社員ID
            sql = sql & "sysdate, "         ' 更新日付
            sql = sql & "'0', "             ' 送信フラグ
            sql = sql & "sysdate)"          ' 送信日時
        End If
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :製作条件メンテナンス 製作条件削除用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:sMkCondNo　　　,I  ,String         　,製作条件№
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込み成否
'説明      :
'履歴      :2001/07/30 蔵本 作成
Public Function DBDRV_scmzc_fcmbc001e_DelMkCond(sMkCondNo As String) As FUNCTION_RETURN

    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_DelMkCond"

    '' 製作条件の削除
    sql = "delete TBCMB012 where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :製作条件メンテナンス 製作条件PG-ID対応更新／挿入用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:sMkCondNo　　　,I  ,String         　,製作条件№
'      　　:pPGIDOld 　　　,I  ,typ_TBCMB013   　,製作条件PG-ID対応オリジナル
'      　　:pPGIDNew 　　　,I  ,typ_TBCMB013   　,製作条件PG-ID対応
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込み成否
'説明      :
'履歴      :2001/07/30 蔵本 作成
Public Function DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng(sMkCondNo As String, pPGIDOld() As typ_TBCMB013, pPGIDNew() As typ_TBCMB013) As FUNCTION_RETURN

    Dim sql As String
    Dim bFlag As Boolean
    Dim i As Long
    Dim j As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_InsPGIDMng"

    For i = 1 To UBound(pPGIDNew)
        With pPGIDNew(i)
            bFlag = False
            For j = 1 To UBound(pPGIDOld)
                If RTrim$(pPGIDOld(j).MKCONDNO) = RTrim$(sMkCondNo) And _
                   RTrim$(pPGIDOld(j).PGIDNO) = RTrim$(.PGIDNO) Then
                    bFlag = True
                    Exit For
                End If
            Next j

            If bFlag = True Then
                '' 製作条件PG-ID対応の更新
                sql = "update TBCMB013 set "
                sql = sql & "MKCONDNO='" & .MKCONDNO & "', "    ' 製作条件No.
                sql = sql & "PGIDNO='" & .PGIDNO & "', "        ' PG-IDNo
                sql = sql & "TSTAFFID='" & .TSTAFFID & "', "    ' 登録社員ID
                sql = sql & "REGDATE=sysdate, "                 ' 登録日付
                sql = sql & "KSTAFFID='" & .KSTAFFID & "', "    ' 更新社員ID
                sql = sql & "UPDDATE=sysdate, "                 ' 更新日付
                sql = sql & "SENDFLAG='0', "                    ' 送信フラグ
                sql = sql & "SENDDATE=sysdate"                  ' 送信日付
                sql = sql & " where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
                sql = sql & " and rtrim(PGIDNO)='" & RTrim$(.PGIDNO) & "'"
            Else
                '' 製作条件PG-ID対応の挿入
                sql = "insert into TBCMB013 ("
                sql = sql & "MKCONDNO, "        ' 製作条件No.
                sql = sql & "PGIDNO, "          ' PG-IDNo
                sql = sql & "TSTAFFID, "        ' 登録社員ID
                sql = sql & "REGDATE, "         ' 登録日付
                sql = sql & "KSTAFFID, "        ' 更新社員ID
                sql = sql & "UPDDATE, "         ' 更新日付
                sql = sql & "SENDFLAG, "        ' 送信フラグ
                sql = sql & "SENDDATE)"         ' 送信日付
                sql = sql & " values ('"
                sql = sql & .MKCONDNO & "', '"  ' 製作条件No.
                sql = sql & .PGIDNO & "', '"    ' PG-IDNo
                sql = sql & .TSTAFFID & "', "   ' 登録社員ID
                sql = sql & "sysdate, '"        ' 登録日付
                sql = sql & .KSTAFFID & "', "   ' 更新社員ID
                sql = sql & "sysdate, "         ' 更新日付
                sql = sql & "'0', "             ' 送信フラグ
                sql = sql & "sysdate)"          ' 送信日付
            End If
        End With
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :製作条件メンテナンス 製作条件PG-ID対応削除用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:sMkCondNo　　　,I  ,String         　,製作条件№
'      　　:sPGIDNo  　　　,I  ,String         　,PG-ID№
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込み成否
'説明      :
'履歴      :2001/07/30 蔵本 作成
Public Function DBDRV_scmzc_fcmbc001e_DelPGIDMng(sMkCondNo As String, sPGIDNo As String) As FUNCTION_RETURN

    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_DelPGIDMng"

    '' 製作条件PG-ID対応の削除
    sql = "delete TBCMB013 where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
    If RTrim$(sPGIDNo) <> "" Then
        sql = sql & " and rtrim(PGIDNO)='" & RTrim$(sPGIDNo) & "'"
    End If
    If OraDB.ExecuteSQL(sql) < 0 Then
        DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMB011」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB011 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMB011_SQL.basより移動)
'           2010/ 1/ 5 SUMCO Akizuki 参照先FROM TBCMB011をTBCMB011_CCVに変更

Public Function DBDRV_GetTBCMB011(records() As typ_TBCMB011, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select PGID, HZPART, HZPTRN, SPACER, UPRING, CHARGE, RTBPOS, RTBSIZE, GAP, UPDM, UPLENGTH, UPRC, RFRNEED, UPSPIN," & _
              " DOWNSPIN, ROPRESS, ARUGON, AIMOIMIN, AIMOIMAX, HCCLASS, HC, AVEUPSPD, UPCNTL, BTMSHAPE, MAGSTR, MAGPOS, CONDGRT," & _
              " MODEL, UPMETHOD, UPCLASS, UPNUM, OPETIME, WTRCOOL, PGID2, RCPT1, RCPT2, RCPT3, RCPT4, RCPT5, CNTL1, CNTL2," & _
              " CNTL3, CNTL4, CNTL5, CNTL6, CNTL7, CNTL8, CNTL9, CNTL10, CNTL11, CNTL12, CNTL13, CNTL14, CNTL15, RUNCOND1," & _
              " RUNCOND2, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCMB011"
    sqlBase = sqlBase & "From TBCMB011_CCV"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB011 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .PGID = rs("PGID")               ' PG-ID
            .HZPART = rs("HZPART")           ' HZパーツ
            .HZPTRN = rs("HZPTRN")           ' HZパターン
            .SPACER = rs("SPACER")           ' スペーサ
            .UPRING = rs("UPRING")           ' アッパーリング
            .CHARGE = rs("CHARGE")           ' チャージ量
            .RTBPOS = rs("RTBPOS")           ' ルツボ位置
            .RTBSIZE = rs("RTBSIZE")         ' ルツボサイズ
            .GAP = rs("GAP")                 ' ギャップ
            .UPDM = rs("UPDM")               ' 引上直径
            .UPLENGTH = rs("UPLENGTH")       ' 引上長（全長）
            .UPRC = rs("UPRC")               ' 引上（RC）
            .RFRNEED = rs("RFRNEED")         ' リフラクタ要否
            .UPSPIN = rs("UPSPIN")           ' 上軸回転数
            .DOWNSPIN = rs("DOWNSPIN")       ' 下軸回転数
            .ROPRESS = rs("ROPRESS")         ' 炉内圧
            .ARUGON = rs("ARUGON")           ' アルゴン量
            .AIMOIMIN = rs("AIMOIMIN")       ' ねらいOi（MIN)
            .AIMOIMAX = rs("AIMOIMAX")       ' ねらいOi（MAX)
            .HCCLASS = rs("HCCLASS")         ' HC種類
            .HC = rs("HC")                   ' HC
            .AVEUPSPD = rs("AVEUPSPD")       ' 平均引上速度
            .UPCNTL = rs("UPCNTL")           ' 引上制御
            .BTMSHAPE = rs("BTMSHAPE")       ' ボトム形状
            .MAGSTR = rs("MAGSTR")           ' 磁場強度
            .MAGPOS = rs("MAGPOS")           ' 磁場位置
            .CONDGRT = rs("CONDGRT")         ' 条件保証登録
            .MODEL = rs("MODEL")             ' 機種
            .UPMETHOD = rs("UPMETHOD")       ' 引上方法
            .UPCLASS = rs("UPCLASS")         ' 引上区分
            .UPNUM = rs("UPNUM")             ' 引上本数
            .OPETIME = rs("OPETIME")         ' 運転時間
            .WTRCOOL = rs("WTRCOOL")         ' 水冷管要否
            .PGID2 = rs("PGID2")             ' PG-ID（一本引）
            .RCPT1 = rs("RCPT1")             ' 対応レシピNo（T1)
            .RCPT2 = rs("RCPT2")             ' 対応レシピNo（T2)
            .RCPT3 = rs("RCPT3")             ' 対応レシピNo（T3)
            .RCPT4 = rs("RCPT4")             ' 対応レシピNo（T4)
            .RCPT5 = rs("RCPT5")             ' 対応レシピNo（T5)
            .CNTL1 = rs("CNTL1")             ' 制限項目（1）
            .CNTL2 = rs("CNTL2")             ' 制限項目（2）
            .CNTL3 = rs("CNTL3")             ' 制限項目（3）
            .CNTL4 = rs("CNTL4")             ' 制限項目（4）
            .CNTL5 = rs("CNTL5")             ' 制限項目（5）
            .CNTL6 = rs("CNTL6")             ' 制限項目（6）
            .CNTL7 = rs("CNTL7")             ' 制限項目（7）
            .CNTL8 = rs("CNTL8")             ' 制限項目（8）
            .CNTL9 = rs("CNTL9")             ' 制限項目（9）
            .CNTL10 = rs("CNTL10")           ' 制限項目（10）
            .CNTL11 = rs("CNTL11")           ' 制限項目（11）
            .CNTL12 = rs("CNTL12")           ' 制限項目（12）
            .CNTL13 = rs("CNTL13")           ' 制限項目（13）
            .CNTL14 = rs("CNTL14")           ' 制限項目（14）
            .CNTL15 = rs("CNTL15")           ' 制限項目（15）
            .RUNCOND1 = rs("RUNCOND1")       ' 運転条件１
            .RUNCOND2 = rs("RUNCOND2")       ' 運転条件２
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

    DBDRV_GetTBCMB011 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMB012」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB012 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMB012_SQL.basより移動)
Public Function DBDRV_GetTBCMB012(records() As typ_TBCMB012, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select MKCONDNO, MODEL, RTBSIZE, CHARGE, HZTYPE, UPSPDTYP, MAGTYPE, USECLS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
              " SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB012"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB012 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .MKCONDNO = rs("MKCONDNO")       ' 制作条件No.
            .MODEL = rs("MODEL")             ' 機種
            .RTBSIZE = rs("RTBSIZE")         ' ルツボサイズ
            .CHARGE = rs("CHARGE")           ' チャージ量
            .HZTYPE = rs("HZTYPE")           ' HZタイプ
            .UPSPDTYP = rs("UPSPDTYP")       ' 引上げ速度タイプ
            .MAGTYPE = rs("MAGTYPE")         ' 磁場タイプ
            .USECLS = rs("USECLS")           ' 使用区分
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日時
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB012 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMB013」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB013 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMB013_SQL.basより移動)
Public Function DBDRV_GetTBCMB013(records() As typ_TBCMB013, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select MKCONDNO, PGIDNO, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB013"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB013 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .MKCONDNO = rs("MKCONDNO")       ' 制作条件No.
            .PGIDNO = rs("PGIDNO")           ' PG-IDNo
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

    DBDRV_GetTBCMB013 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMB005」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB005 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMB005_SQL.basより移動)
Public Function DBDRV_GetTBCMB005(records() As typ_TBCMB005, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select SYSCLASS, CLASS, CODE, INFO1, INFO2, INFO3, INFO4, INFO5, INFO6, INFO7, INFO8, INFO9, NOTE, TSTAFFID," & _
              " REGDATE, KSTAFFID, UPDDATE "
    sqlBase = sqlBase & "From TBCMB005"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB005 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .SYSCLASS = rs("SYSCLASS")       ' システム区分
            .Class = rs("CLASS")             ' 区分
            .CODE = rs("CODE")               ' コード
            .INFO1 = rs("INFO1")             ' 情報１
            .INFO2 = rs("INFO2")             ' 情報２
            .INFO3 = rs("INFO3")             ' 情報３
            .INFO4 = rs("INFO4")             ' 情報４
            .INFO5 = rs("INFO5")             ' 情報５
            .INFO6 = rs("INFO6")             ' 情報６
            .INFO7 = rs("INFO7")             ' 情報７
            .INFO8 = rs("INFO8")             ' 情報８
            .INFO9 = rs("INFO9")             ' 情報９
            .NOTE = rs("NOTE")               ' 備考
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB005 = FUNCTION_RETURN_SUCCESS
End Function

' 2007/09/04 SPK Tsutsumi Add Start
Public Function GetMukeCode() As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Long      'レコード数
    Dim i  As Long
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    
    GetMukeCode = FUNCTION_RETURN_FAILURE
    
    sql = "Select CODEA9,NAMEJA9 "
    sql = sql & "from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '20' "
    sql = sql & "and (CODEA9 = '14' "
    sql = sql & "or CODEA9 = '15' "
    sql = sql & "or CODEA9 = '16' "
    sql = sql & "or CODEA9 = 'ALL') "

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If
    
    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim s_MukesakiBase(recCnt)
    
    If recCnt = 0 Then
        Exit Function
    End If
    
    For i = 1 To recCnt
        With s_MukesakiBase(i)
            If IsNull(rs.Fields("CODEA9")) = False Then
                .sMukeCode = rs.Fields("CODEA9")    ' 向先コード
            End If
            
            If IsNull(rs.Fields("NAMEJA9")) = False Then
                .sMukeName = rs.Fields("NAMEJA9")  ' 向先名
'                f_cmbc061_0.cmbMukesaki.AddItem .sMukeName
            End If
        End With
        rs.MoveNext
    Next
    rs.Close

    GetMukeCode = FUNCTION_RETURN_SUCCESS
proc_exit:
    '終了
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
' 2007/09/04 SPK Tsutsumi Add End
