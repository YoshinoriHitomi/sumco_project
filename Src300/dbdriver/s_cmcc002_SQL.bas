Attribute VB_Name = "s_cmcc002_SQL"
Option Explicit
'                                     2001/08/24
'================================================
' DBアクセス関数
' 定義内容: TBCMB011 (PG-ID管理)
' 参照　　: 060200_全テーブル
'================================================

#If False Then      'テーブルの型定義は別のs_cmzcTableDefs.basで行う
'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_TBCMB011
    PGID As String * 10             ' PG-ID
    HZPART As String * 4            ' HZパーツ
    HZPTRN As String * 2            ' HZパターン
    SPACER As String * 5            ' スペーサ
    UPRING As String * 5            ' アッパーリング
    CHARGE As Long                  ' チャージ量
    RTBPOS As Integer               ' ルツボ位置
    RTBSIZE As String * 2           ' ルツボサイズ
    GAP As Integer                  ' ギャップ
    UPDM As Integer                 ' 引上直径
    UPLENGTH As Integer             ' 引上長（全長）
    UPRC As Integer                 ' 引上（RC）
    RFRNEED As String * 1           ' リフラクタ要否
    UPSPIN As String * 10           ' 上軸回転数
    DOWNSPIN As String * 10         ' 下軸回転数
    ROPRESS As String * 8           ' 炉内圧
    ARUGON As String * 7            ' アルゴン量
    AIMOIMIN As Double              ' ねらいOi（MIN)
    AIMOIMAX As Double              ' ねらいOi（MAX)
    HCCLASS As String * 7           ' HC種類
    HC As String * 3                ' HC
    AVEUPSPD As Double              ' 平均引上速度
    UPCNTL As String * 1            ' 引上制御
    BTMSHAPE As String * 1          ' ボトム形状
    MAGSTR As Double                ' 磁場強度
    MAGPOS As Long                  ' 磁場位置
    CONDGRT As String * 10          ' 条件保証登録
    MODEL As String * 4             ' 機種
    UPMETHOD As String * 4          ' 引上方法
    UPCLASS As String * 2           ' 引上区分
    UPNUM As String * 1             ' 引上本数
    OPETIME As Long                 ' 運転時間
    WTRCOOL As String * 1           ' 水冷管要否
    PGID2 As String * 10            ' PG-ID（一本引）
    RCPT1 As String * 3             ' 対応レシピNo（T1)
    RCPT2 As String * 3             ' 対応レシピNo（T2)
    RCPT3 As String * 3             ' 対応レシピNo（T3)
    RCPT4 As String * 3             ' 対応レシピNo（T4)
    RCPT5 As String * 3             ' 対応レシピNo（T5)
    CNTL1 As String * 1             ' 制限項目（1）
    CNTL2 As String * 1             ' 制限項目（2）
    CNTL3 As String * 1             ' 制限項目（3）
    CNTL4 As String * 1             ' 制限項目（4）
    CNTL5 As String * 1             ' 制限項目（5）
    CNTL6 As String * 1             ' 制限項目（6）
    CNTL7 As String * 1             ' 制限項目（7）
    CNTL8 As String * 1             ' 制限項目（8）
    CNTL9 As String * 1             ' 制限項目（9）
    CNTL10 As String * 1            ' 制限項目（10）
    CNTL11 As String * 1            ' 制限項目（11）
    CNTL12 As String * 1            ' 制限項目（12）
    CNTL13 As String * 1            ' 制限項目（13）
    CNTL14 As String * 1            ' 制限項目（14）
    CNTL15 As String * 1            ' 制限項目（15）
    RUNCOND1 As String              ' 運転条件１
    RUNCOND2 As String              ' 運転条件２
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type
#End If



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
'履歴      :2001/08/24作成　野村
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
    sqlBase = sqlBase & "From TBCMB011"
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

'8/2補足

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
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCMB012(records() As typ_TBCMB012, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
' 払出規制項目追加対応 yakimura 2002.12.01 start
    sqlBase = "Select MKCONDNO, MODEL, RTBSIZE, CHARGE, HZTYPE, UPSPDTYP, MAGTYPE, USECLS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
              " SENDFLAG, SENDDATE, NVL(TOPREG,0) TOPREG, NVL(TAILREG,0) TAILREG, NVL(BTMSPRT,0) BTMSPRT "
    sqlBase = sqlBase & "From TBCMB012"
' 払出規制項目追加対応 yakimura 2002.12.01 end
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
' 払出規制項目追加対応 yakimura 2002.12.01 start
            .Topreg = rs("TOPREG")           ' TOP規制
            .Tailreg = rs("TAILREG")         ' TAIL規制
            .Btmsprt = rs("BTMSPRT")         ' ボトム析出規制
' 払出規制項目追加対応 yakimura 2002.12.01 end
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB012 = FUNCTION_RETURN_SUCCESS
End Function


'8/2 補足
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
'履歴      :2001/08/24作成　野村
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
' 払出規制項目追加対応 yakimura 2002.12.01 start
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
            sql = sql & "SENDDATE=sysdate, "                ' 送信日時
            sql = sql & "TOPREG='" & .Topreg & "', "        ' TOP規制
            sql = sql & "TAILREG='" & .Tailreg & "', "      ' TAIL規制
            sql = sql & "BTMSPRT='" & .Btmsprt & "'"        ' ボトム析出規制
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
            sql = sql & "SENDDATE,"         ' 送信日時
            sql = sql & "TOPREG, "          ' TOP規制
            sql = sql & "TAILREG, "         ' TAIL規制
            sql = sql & "BTMSPRT) "         ' ボトム析出規制
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
            sql = sql & "sysdate, '"        ' 送信日時
            sql = sql & .Topreg & " ', '"   ' TOP規制
            sql = sql & .Tailreg & " ', '"  ' TAIL規制
            sql = sql & .Btmsprt & "')"     ' ボトム析出規制
        End If
    End With
' 払出規制項目追加対応 yakimura 2002.12.01 end
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
