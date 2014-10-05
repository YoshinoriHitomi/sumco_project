Attribute VB_Name = "s_cmcc001_SQL"
Option Explicit
'                                     2001/06/11
'================================================
' DBアクセス関数
' 定義内容: TBCMB011 (PG-ID管理)
' 参照　　: 060200_全テーブル
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmbc001c_Disp
    PGID As String * 10             ' PG-ID
    HZPART As String * 4            ' HZパーツ
    AIMOIMIN As Double              ' ねらいOi（MIN)
    AIMOIMAX As Double              ' ねらいOi（MAX)
    AVEUPSPD As Double              ' 平均引上速度
    MODEL As String * 4             ' 機種
    UPMETHOD As String * 1          ' 引上方法
    UPCLASS As String * 2           ' 引上区分
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
End Type



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
'履歴      :2001/06/11作成　野村
Public Function DBDRV_cmbc001c_Disp(records() As typ_cmbc001c_Disp, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001c_SQL.bas -- Function DBDRV_cmbc001c_Disp"
    
    DBDRV_cmbc001c_Disp = FUNCTION_RETURN_FAILURE
    
    sql = "Select PGID, HZPART, AIMOIMIN, AIMOIMAX, AVEUPSPD, MODEL, UPMETHOD, UPCLASS, REGDATE, UPDDATE "
    sql = sql & "From TBCMB011 "
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & sqlWhere & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_cmbc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .PGID = rs("PGID")               ' PG-ID
            .HZPART = rs("HZPART")           ' HZパーツ
'            .HZPTRN = rs("HZPTRN")           ' HZパターン
            .AIMOIMIN = rs("AIMOIMIN")       ' ねらいOi（MIN)
            .AIMOIMAX = rs("AIMOIMAX")       ' ねらいOi（MAX)
            .AVEUPSPD = rs("AVEUPSPD")       ' 平均引上速度
            .MODEL = rs("MODEL")             ' 機種
            .UPMETHOD = rs("UPMETHOD")       ' 引上方法
            .UPCLASS = rs("UPCLASS")         ' 引上区分
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_cmbc001c_Disp = FUNCTION_RETURN_SUCCESS

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


'概要      :テーブル「TBCMB011」へ挿入
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB011 ,抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/10/04 作成　蔵本
Public Function DBDRV_cmbc001c_Ins(staff As String, records() As typ_VAX_DR_CNDS) As FUNCTION_RETURN

    Dim i As Long
    Dim sql As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001c_SQL.bas -- Function DBDRV_cmbc001c_Ins"

    ' PG-ID管理
    For i = 1 To UBound(records)
        sql = "insert into TBCMB011 ( "
        sql = sql & "PGID, "             ' PG-ID
        sql = sql & "HZPART, "           ' HZパーツ
        sql = sql & "HZPTRN, "           ' HZパターン
        sql = sql & "SPACER, "           ' スペーサ
        sql = sql & "UPRING, "           ' アッパーリング
        sql = sql & "CHARGE, "           ' チャージ量
        sql = sql & "RTBPOS, "           ' ルツボ位置
        sql = sql & "RTBSIZE, "          ' ルツボサイズ
        sql = sql & "GAP, "              ' ギャップ
        sql = sql & "UPDM, "             ' 引上直径
        sql = sql & "UPLENGTH, "         ' 引上長（全長）
        sql = sql & "UPRC, "             ' 引上（RC）
        sql = sql & "RFRNEED, "          ' リフラクタ要否
        sql = sql & "UPSPIN, "           ' 上軸回転数
        sql = sql & "DOWNSPIN, "         ' 下軸回転数
        sql = sql & "ROPRESS, "          ' 炉内圧
        sql = sql & "ARUGON, "           ' アルゴン量
        sql = sql & "DRDOP,"
        sql = sql & "DRAR3,"
        sql = sql & "AIMOIMIN, "         ' ねらいOi（MIN)
        sql = sql & "AIMOIMAX, "         ' ねらいOi（MAX)
        sql = sql & "HCCLASS, "          ' HC種類
        sql = sql & "HC, "               ' HC
        sql = sql & "AVEUPSPD, "         ' 平均引上速度
        sql = sql & "UPCNTL, "           ' 引上制御
        sql = sql & "BTMSHAPE, "         ' ボトム形状
        sql = sql & "MAGSTR, "           ' 磁場強度
        sql = sql & "MAGPOS, "           ' 磁場位置
        sql = sql & "CONDGRT, "          ' 条件保証登録
        sql = sql & "MODEL, "            ' 機種
        sql = sql & "UPMETHOD, "         ' 引上方法
        sql = sql & "UPCLASS, "          ' 引上区分
        sql = sql & "UPNUM, "            ' 引上本数
        sql = sql & "OPETIME, "          ' 運転時間
        sql = sql & "WTRCOOL, "          ' 水冷管要否
        sql = sql & "PGID2, "            ' PG-ID（一本引）
        sql = sql & "RCPT1, "            ' 対応レシピNo（T1)
        sql = sql & "RCPT2, "            ' 対応レシピNo（T2)
        sql = sql & "RCPT3, "            ' 対応レシピNo（T3)
        sql = sql & "RCPT4, "            ' 対応レシピNo（T4)
        sql = sql & "RCPT5, "            ' 対応レシピNo（T5)
        sql = sql & "RCPT6, "            ' 対応レシピNo（T6) 8/30 Yam
        sql = sql & "CNTL1, "            ' 制限項目（1）
        sql = sql & "CNTL2, "            ' 制限項目（2）
        sql = sql & "CNTL3, "            ' 制限項目（3）
        sql = sql & "CNTL4, "            ' 制限項目（4）
        sql = sql & "CNTL5, "            ' 制限項目（5）
        sql = sql & "CNTL6, "            ' 制限項目（6）
        sql = sql & "CNTL7, "            ' 制限項目（7）
        sql = sql & "CNTL8, "            ' 制限項目（8）
        sql = sql & "CNTL9, "            ' 制限項目（9）
        sql = sql & "CNTL10, "           ' 制限項目（10）
        sql = sql & "CNTL11, "           ' 制限項目（11）
        sql = sql & "CNTL12, "           ' 制限項目（12）
        sql = sql & "CNTL13, "           ' 制限項目（13）
        sql = sql & "CNTL14, "           ' 制限項目（14）
        sql = sql & "CNTL15, "           ' 制限項目（15）
        sql = sql & "RUNCOND1, "         ' 運転条件１
        sql = sql & "RUNCOND2, "         ' 運転条件２
        sql = sql & "TSTAFFID, "         ' 登録社員ID
        sql = sql & "REGDATE, "          ' 登録日付
        sql = sql & "KSTAFFID, "         ' 更新社員ID
        sql = sql & "UPDDATE, "          ' 更新日付
        sql = sql & "SENDFLAG, "         ' 送信フラグ
        sql = sql & "SENDDATE )"         ' 送信日付
        With records(i)
            sql = sql & " values ("
            sql = sql & " '" & .PG_ID & "', "            ' PG-ID
            sql = sql & " ' ', "                      ' HZパーツ
            sql = sql & " ' ', "                        ' HZパターン
            sql = sql & " ' ', "                     ' スペーサ
            sql = sql & " ' ', "                     ' アッパーリング
            sql = sql & " " & .DR_CHRG * 100 & ", "      ' チャージ量
            sql = sql & " " & .DR_CPOS & ", "            ' ルツボ位置
            sql = sql & " '" & .DR_CSIZ & "', "          ' ルツボサイズ
            sql = sql & " " & .DR_GAP & ", "             ' ギャップ
            sql = sql & " " & .DR_DIA & ", "             ' 引上直径
            sql = sql & " " & .DR_LEN0 & ", "            ' 引上長（全長）
            sql = sql & " " & .DR_LEN1 & ", "            ' 引上（RC）
            sql = sql & " ' ', "                         ' リフラクタ要否
            sql = sql & " '" & Trim(.DR_SR) & "', "      ' 上軸回転数
            sql = sql & " '" & Trim(.DR_CR) & "', "      ' 下軸回転数
            sql = sql & " '" & Trim(.DR_PRES7) & "', "   ' 炉内圧             '2003/05/16 osawa
            sql = sql & " '" & Trim(.DR_AR7) & "', "     ' アルゴン量         '2003/05/16 osawa
            'sql = sql & " '" & .DR_PRES7 & "', "          ' 炉内圧
            'sql = sql & " '" & .DR_AR7 & "', "            ' アルゴン量
            sql = sql & " '" & .DR_DOP & "', "
            sql = sql & " '" & .DR_AR3 & "', "
            sql = sql & " 0, "                           ' ねらいOi（MIN)
            sql = sql & " 0, "                           ' ねらいOi（MAX)
            sql = sql & " ' ', "                         ' HC種類
            sql = sql & " ' ', "                         ' HC
            sql = sql & " 0, "                           ' 平均引上速度
            sql = sql & " ' ', "                         ' 引上制御
            sql = sql & " ' ', "                         ' ボトム形状
            sql = sql & " 0, "                           ' 磁場強度
            sql = sql & " 0, "                           ' 磁場位置
            sql = sql & " ' ', "                         ' 条件保証登録
            sql = sql & " ' ', "                         ' 機種
            sql = sql & " ' ', "                         ' 引上方法
            sql = sql & " ' ', "                         ' 引上区分
            sql = sql & " ' ', "                         ' 引上本数
            sql = sql & " 0, "                           ' 運転時間
            sql = sql & " ' ', "                         ' 水冷管要否
            sql = sql & " ' ', "                         ' PG-ID（一本引）
            sql = sql & " ' ', "                         ' 対応レシピNo（T1)
            sql = sql & " ' ', "                         ' 対応レシピNo（T2)
            sql = sql & " ' ', "                         ' 対応レシピNo（T3)
            sql = sql & " ' ', "                         ' 対応レシピNo（T4)
            sql = sql & " ' ', "                         ' 対応レシピNo（T5)
            sql = sql & " ' ', "                         ' 対応レシピNo（T6)
            sql = sql & " ' ', "                         ' 制限項目（1）
            sql = sql & " ' ', "                         ' 制限項目（2）
            sql = sql & " ' ', "                         ' 制限項目（3）
            sql = sql & " ' ', "                         ' 制限項目（4）
            sql = sql & " ' ', "                         ' 制限項目（5）
            sql = sql & " ' ', "                         ' 制限項目（6）
            sql = sql & " ' ', "                         ' 制限項目（7）
            sql = sql & " ' ', "                         ' 制限項目（8）
            sql = sql & " ' ', "                         ' 制限項目（9）
            sql = sql & " ' ', "                         ' 制限項目（10）
            sql = sql & " ' ', "                         ' 制限項目（11）
            sql = sql & " ' ', "                         ' 制限項目（12）
            sql = sql & " ' ', "                         ' 制限項目（13）
            sql = sql & " ' ', "                         ' 制限項目（14）
            sql = sql & " ' ', "                         ' 制限項目（15）
            sql = sql & " ' ', "                         ' 運転条件１
            sql = sql & " ' ', "                         ' 運転条件２
            sql = sql & " '" & staff & "', "             ' 登録社員ID
            sql = sql & " sysdate, "                     ' 登録日付
            sql = sql & " '" & staff & "', "             ' 更新社員ID
            sql = sql & " sysdate, "                     ' 更新日付
            sql = sql & " '0', "                         ' 送信フラグ
            sql = sql & " sysdate ) "                    ' 送信日付
        End With

        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_cmbc001c_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    
        DBDRV_cmbc001c_Ins = FUNCTION_RETURN_SUCCESS
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
    DBDRV_cmbc001c_Ins = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'7/30　補足
'------------------------------------------------
' DBアクセス関数（抽出編）
'------------------------------------------------
'概要      :テーブル「TBCMB011」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB011 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/11作成　長野
Public Function DBDRV_cmbc001d_Disp(records() As typ_TBCMB011, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String                   'SQL全体
Dim sqlBase As String                   'SQL基本部(WHERE節の前まで)
Dim rs      As OraDynaset               'RecordSet
Dim recCnt  As Long                     'レコード数
Dim i       As Long                     'ﾙｰﾌﾟｶｳﾝﾄ

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Disp"

    sqlBase = "Select PGID, HZPART, HZPTRN, SPACER, UPRING, CHARGE, RTBPOS, RTBSIZE, GAP, UPDM, UPLENGTH, UPRC, RFRNEED, UPSPIN," & _
              " DOWNSPIN, ROPRESS, ARUGON, AIMOIMIN, AIMOIMAX, HCCLASS, HC, AVEUPSPD, UPCNTL, BTMSHAPE, MAGSTR, MAGPOS, CONDGRT," & _
              " MODEL, UPMETHOD, UPCLASS, UPNUM, OPETIME, WTRCOOL, PGID2, RCPT1, RCPT2, RCPT3, RCPT4, RCPT5, RCPT6, CNTL1, CNTL2," & _
              " CNTL3, CNTL4, CNTL5, CNTL6, CNTL7, CNTL8, CNTL9, CNTL10, CNTL11, CNTL12, CNTL13, CNTL14, CNTL15, RUNCOND1," & _
              " RUNCOND2, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, DRDOP, DRAR3 "
    sqlBase = sqlBase & "From TBCMB011 "
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_cmbc001d_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
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
            .RCPT6 = rs("RCPT6")             ' 対応レシピNo（T6)
            .CNTL1 = rs("CNTL1")             ' 制御項目（1）
            .CNTL2 = rs("CNTL2")             ' 制御項目（2）
            .CNTL3 = rs("CNTL3")             ' 制御項目（3）
            .CNTL4 = rs("CNTL4")             ' 制御項目（4）
            .CNTL5 = rs("CNTL5")             ' 制御項目（5）
            .CNTL6 = rs("CNTL6")             ' 制御項目（6）
            .CNTL7 = rs("CNTL7")             ' 制御項目（7）
            .CNTL8 = rs("CNTL8")             ' 制御項目（8）
            .CNTL9 = rs("CNTL9")             ' 制御項目（9）
            .CNTL10 = rs("CNTL10")           ' 制御項目（10）
            .CNTL11 = rs("CNTL11")           ' 制御項目（11）
            .CNTL12 = rs("CNTL12")           ' 制御項目（12）
            .CNTL13 = rs("CNTL13")           ' 制御項目（13）
            .CNTL14 = rs("CNTL14")           ' 制御項目（14）
            .CNTL15 = rs("CNTL15")           ' 制御項目（15）
            .RUNCOND1 = rs("RUNCOND1")       ' 運転条件１
            .RUNCOND2 = rs("RUNCOND2")       ' 運転条件２
'            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
'            .REGDATE = rs("REGDATE")         ' 登録日付
'            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
'            .UPDDATE = rs("UPDDATE")         ' 更新日付
'            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
'            .SENDDATE = rs("SENDDATE")       ' 送信日付
            .DRDOP = IIf(rs("DRDOP") <> "", rs("DRDOP"), " ") ' ドープ      4/30
            .DRAR3 = IIf(rs("DRAR3") <> "", rs("DRAR3"), " ") ' アルゴン№３流量
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_cmbc001d_Disp = FUNCTION_RETURN_SUCCESS

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

'7/30 補足
Public Function DBDRV_cmbc001d_Exec(records As typ_TBCMB011) As FUNCTION_RETURN
'------------------------------------------------
' DBアクセス関数（更新編）
'------------------------------------------------
'概要      :テーブル「TBCMB011」の条件にあったレコードに更新をかける
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records　     ,O  ,typ_TBCMB011 ,抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :抽出ﾃﾞｰﾀの桁数・書式ﾁｪｯｸは"済み"とする
'履歴      :2001/06/19(TUE)作成　長野
Dim sql     As String                   'SQL全体
Dim rs      As OraDynaset               'RecordSet
Dim UpdID   As String                   '更新対象PGID


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Exec"

    UpdID = records.PGID

'2001/09/05 S.Sano Start 更新日時がセットされていない。
'2001/09/05 S.Sano Start このモードでsysdateのセット方法が不明。
'    sql = "SELECT * FROM TBCMB011 WHERE(PGID = '" & UpdID & "')"
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'
'    rs.Edit
'    With records
'        rs("HZPART") = StrNoNull(.HZPART)          ' HZﾊﾟｰﾂ
'        rs("HZPTRN") = StrNoNull(.HZPTRN)          ' HZﾊﾟﾀｰﾝ
'        rs("SPACER") = StrNoNull(.SPACER)          ' ｽﾍﾟｰｻ
'        rs("UPRING") = StrNoNull(.UPRING)          ' ｱｯﾊﾟｰﾘﾝｸﾞ
'        rs("CHARGE") = .CHARGE          ' ﾁｬｰｼﾞ量
'        rs("RTBPOS") = .RTBPOS          ' ﾙﾂﾎﾞ位置
'        rs("RTBSIZE") = StrNoNull(.RTBSIZE)        ' ﾙﾂﾎﾞｻｲｽﾞ
'        rs("GAP") = .GAP                ' ｷﾞｬｯﾌﾟ
'        rs("UPDM") = .UPDM              ' 引上直径
'        rs("UPLENGTH") = .UPLENGTH      ' 引上長
'        rs("UPRC") = .UPRC              ' 引上RC
'        rs("RFRNEED") = StrNoNull(.RFRNEED)        ' ﾘﾌﾗｸﾀ要否
'        rs("UPSPIN") = StrNoNull(.UPSPIN)          ' 上軸回転数
'        rs("DOWNSPIN") = StrNoNull(.DOWNSPIN)      ' 下軸回転数
'        rs("ROPRESS") = StrNoNull(.ROPRESS)        ' 炉内圧
'        rs("ARUGON") = StrNoNull(.ARUGON)          ' ｱﾙｺﾞﾝ量
'        rs("AIMOIMIN") = .AIMOIMIN      ' ねらいiO(MIN)
'        rs("AIMOIMAX") = .AIMOIMAX      ' ねらいiO(MAX)
'        rs("HCCLASS") = StrNoNull(.HCCLASS)        ' HC種類
'        rs("HC") = StrNoNull(.HC)                  ' HC
'        rs("AVEUPSPD") = .AVEUPSPD      ' 平均引上速度
'        rs("UPCNTL") = StrNoNull(.UPCNTL)          ' 引上制御
'        rs("BTMSHAPE") = StrNoNull(.BTMSHAPE)      ' ﾎﾞﾄﾑ形状
'        rs("MAGSTR") = .MAGSTR          ' 磁場強度
'        rs("MAGPOS") = .MAGPOS          ' 磁場位置
'        rs("CONDGRT") = StrNoNull(.CONDGRT)        ' 条件保証登録
'        rs("MODEL") = StrNoNull(.MODEL)            ' 機種
'        rs("UPMETHOD") = StrNoNull(.UPMETHOD)      ' 引上方法
'        rs("UPCLASS") = StrNoNull(.UPCLASS)        ' 引上区分
'        rs("UPNUM") = StrNoNull(.UPNUM)            ' 引上本数
'        rs("OPETIME") = .OPETIME        ' 運転時間
'        rs("WTRCOOL") = StrNoNull(.WTRCOOL)        ' 水冷管要否
'        rs("PGID2") = StrNoNull(.PGID2)            ' PG-ID（一本引）
'        rs("RCPT1") = StrNoNull(.RCPT1)            ' 対応ﾚｼﾋﾟNo（T1）
'        rs("RCPT2") = StrNoNull(.RCPT2)            ' 対応ﾚｼﾋﾟNo（T2）
'        rs("RCPT3") = StrNoNull(.RCPT3)            ' 対応ﾚｼﾋﾟNo（T3）
'        rs("RCPT4") = StrNoNull(.RCPT4)            ' 対応ﾚｼﾋﾟNo（T4）
'        rs("RCPT5") = StrNoNull(.RCPT5)            ' 対応ﾚｼﾋﾟNo（T5）
'        rs("CNTL1") = StrNoNull(.CNTL1)            ' 制御項目(1)
'        rs("CNTL2") = StrNoNull(.CNTL2)            ' 制御項目(2)
'        rs("CNTL3") = StrNoNull(.CNTL3)            ' 制御項目(3)
'        rs("CNTL4") = StrNoNull(.CNTL4)            ' 制御項目(4)
'        rs("CNTL5") = StrNoNull(.CNTL5)            ' 制御項目(5)
'        rs("CNTL6") = StrNoNull(.CNTL6)            ' 制御項目(6)
'        rs("CNTL7") = StrNoNull(.CNTL7)            ' 制御項目(7)
'        rs("CNTL8") = StrNoNull(.CNTL8)            ' 制御項目(8)
'        rs("CNTL9") = StrNoNull(.CNTL9)            ' 制御項目(9)
'        rs("CNTL10") = StrNoNull(.CNTL10)          ' 制御項目(10)
'        rs("CNTL11") = StrNoNull(.CNTL11)          ' 制御項目(11)
'        rs("CNTL12") = StrNoNull(.CNTL12)          ' 制御項目(12)
'        rs("CNTL13") = StrNoNull(.CNTL13)          ' 制御項目(13)
'        rs("CNTL14") = StrNoNull(.CNTL14)          ' 制御項目(14)
'        rs("CNTL15") = StrNoNull(.CNTL15)          ' 制御項目(15)
'        rs("RUNCOND1") = StrNoNull(.RUNCOND1)      ' 運転条件1
'        rs("RUNCOND2") = StrNoNull(.RUNCOND2)      ' 運転条件2
'    End With
'    rs.Update
'
'    rs.Close
    
'2001/09/05 S.Sano Start
    With records
    sql = "update TBCMB011 set "
    sql = sql & "HZPART = '" & StrNoNull(.HZPART) & "', "       ' HZﾊﾟｰﾂ
    sql = sql & "HZPTRN = '" & StrNoNull(.HZPTRN) & "', "       ' HZﾊﾟﾀｰﾝ
    sql = sql & "SPACER = '" & StrNoNull(.SPACER) & "', "       ' ｽﾍﾟｰｻ
    sql = sql & "UPRING = '" & StrNoNull(.UPRING) & "', "       ' ｱｯﾊﾟｰﾘﾝｸﾞ
    sql = sql & "CHARGE = " & .CHARGE & ", "                    ' ﾁｬｰｼﾞ量
    sql = sql & "RTBPOS = " & .RTBPOS & ", "                    ' ﾙﾂﾎﾞ位置
    sql = sql & "RTBSIZE = '" & StrNoNull(.RTBSIZE) & "', "     ' ﾙﾂﾎﾞｻｲｽﾞ
    sql = sql & "GAP = " & .GAP & ", "                          ' ｷﾞｬｯﾌﾟ
    sql = sql & "UPDM = " & .UPDM & ", "                        ' 引上直径
    sql = sql & "UPLENGTH = " & .UPLENGTH & ", "                ' 引上長
    sql = sql & "UPRC = " & .UPRC & ", "                        ' 引上RC
    sql = sql & "RFRNEED = '" & StrNoNull(.RFRNEED) & "', "     ' ﾘﾌﾗｸﾀ要否
    sql = sql & "UPSPIN = '" & StrNoNull(.UPSPIN) & "', "       ' 上軸回転数
    sql = sql & "DOWNSPIN = '" & StrNoNull(.DOWNSPIN) & "', "   ' 下軸回転数
    sql = sql & "ROPRESS = '" & StrNoNull(.ROPRESS) & "', "     ' 炉内圧
    sql = sql & "ARUGON = '" & StrNoNull(.ARUGON) & "', "       ' ｱﾙｺﾞﾝ量
    sql = sql & "DRAR3 ='" & StrNoNull(.DRAR3) & "', "          ' アルゴン№３流量 ' 4/30 YAM
    sql = sql & "DRDOP ='" & StrNoNull(.DRDOP) & "', "          ' ドープ
    sql = sql & "AIMOIMAX = " & .AIMOIMAX & ", "                ' ねらいiO(MAX)
    sql = sql & "HCCLASS = '" & StrNoNull(.HCCLASS) & "', "     ' HC種類
    sql = sql & "HC = '" & StrNoNull(.HC) & "', "               ' HC
    sql = sql & "AVEUPSPD = " & .AVEUPSPD & ", "                ' 平均引上速度
    sql = sql & "UPCNTL = '" & StrNoNull(.UPCNTL) & "', "       ' 引上制御
    sql = sql & "BTMSHAPE = '" & StrNoNull(.BTMSHAPE) & "', "   ' ﾎﾞﾄﾑ形状
    sql = sql & "MAGSTR = " & .MAGSTR & ", "                    ' 磁場強度
    sql = sql & "MAGPOS = " & .MAGPOS & ", "                    ' 磁場位置
    sql = sql & "CONDGRT = '" & StrNoNull(.CONDGRT) & "', "     ' 条件保証登録
    sql = sql & "MODEL = '" & StrNoNull(.MODEL) & "', "         ' 機種
    sql = sql & "UPMETHOD = '" & StrNoNull(.UPMETHOD) & "', "   ' 引上方法
    sql = sql & "UPCLASS = '" & StrNoNull(.UPCLASS) & "', "     ' 引上区分
    sql = sql & "UPNUM = '" & StrNoNull(.UPNUM) & "', "         ' 引上本数
    sql = sql & "OPETIME = " & .OPETIME & ", "                  ' 運転時間
    sql = sql & "WTRCOOL = '" & StrNoNull(.WTRCOOL) & "', "     ' 水冷管要否
    sql = sql & "PGID2 = '" & StrNoNull(.PGID2) & "', "         ' PG-ID（一本引）
    sql = sql & "RCPT1 = '" & StrNoNull(.RCPT1) & "', "         ' 対応ﾚｼﾋﾟNo（T1）
    sql = sql & "RCPT2 = '" & StrNoNull(.RCPT2) & "', "         ' 対応ﾚｼﾋﾟNo（T2）
    sql = sql & "RCPT3 = '" & StrNoNull(.RCPT3) & "', "         ' 対応ﾚｼﾋﾟNo（T3）
    sql = sql & "RCPT4 = '" & StrNoNull(.RCPT4) & "', "         ' 対応ﾚｼﾋﾟNo（T4）
    sql = sql & "RCPT5 = '" & StrNoNull(.RCPT5) & "', "         ' 対応ﾚｼﾋﾟNo（T5）
    sql = sql & "RCPT6 = '" & StrNoNull(.RCPT6) & "', "         ' 対応ﾚｼﾋﾟNo（T6）
    sql = sql & "CNTL1 = '" & StrNoNull(.CNTL1) & "', "         ' 制御項目(1)
    sql = sql & "CNTL2 = '" & StrNoNull(.CNTL2) & "', "         ' 制御項目(2)
    sql = sql & "CNTL3 = '" & StrNoNull(.CNTL3) & "', "         ' 制御項目(3)
    sql = sql & "CNTL4 = '" & StrNoNull(.CNTL4) & "', "         ' 制御項目(4)
    sql = sql & "CNTL5 = '" & StrNoNull(.CNTL5) & "', "         ' 制御項目(5)
    sql = sql & "CNTL6 = '" & StrNoNull(.CNTL6) & "', "         ' 制御項目(6)
    sql = sql & "CNTL7 = '" & StrNoNull(.CNTL7) & "', "         ' 制御項目(7)
    sql = sql & "CNTL8 = '" & StrNoNull(.CNTL8) & "', "         ' 制御項目(8)
    sql = sql & "CNTL9 = '" & StrNoNull(.CNTL9) & "', "         ' 制御項目(9)
    sql = sql & "CNTL10 = '" & StrNoNull(.CNTL10) & "', "       ' 制御項目(10)
    sql = sql & "CNTL11 = '" & StrNoNull(.CNTL11) & "', "       ' 制御項目(11)
    sql = sql & "CNTL12 = '" & StrNoNull(.CNTL12) & "', "       ' 制御項目(12)
    sql = sql & "CNTL13 = '" & StrNoNull(.CNTL13) & "', "       ' 制御項目(13)
    sql = sql & "CNTL14 = '" & StrNoNull(.CNTL14) & "', "       ' 制御項目(14)
    sql = sql & "CNTL15 = '" & StrNoNull(.CNTL15) & "', "       ' 制御項目(15)
    sql = sql & "RUNCOND1 = '" & StrNoNull(.RUNCOND1) & "', "   ' 運転条件1
    sql = sql & "RUNCOND2 = '" & StrNoNull(.RUNCOND2) & "', "   ' 運転条件2
    sql = sql & "KSTAFFID = '" & .KSTAFFID & "', "              ' 更新社員ID
    sql = sql & "UPDDATE = sysdate "                            ' 更新日付
    sql = sql & "where PGID = '" & UpdID & "'"
    End With
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_cmbc001d_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_cmbc001d_Exec = FUNCTION_RETURN_SUCCESS
'2001/09/05 S.Sano End

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
'8/6 補足
'------------------------------------------------
' DBアクセス関数（削除編）
'------------------------------------------------
'概要      :テーブル「TBCMB011」の条件にあったレコードを削除
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :PGID　        ,O  ,String       ,削除PG-ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/10/05 作成　蔵本
Public Function DBDRV_cmbc001d_Del(PGID As String) As FUNCTION_RETURN

    Dim sql     As String                   'SQL全体

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Del"
    
    sql = "delete "
    sql = sql & "from "
    sql = sql & "TBCMB011 "
    sql = sql & "where "
    sql = sql & "trim(PGID)='" & Trim(PGID) & "'"
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_cmbc001d_Del = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_cmbc001d_Del = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_cmbc001d_Del = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'8/6 補足
Private Function StrNoNull(s$) As String
    If Trim$(s) = vbNullString Then
        StrNoNull = " "
    Else
        StrNoNull = Trim$(s)
    End If
End Function
