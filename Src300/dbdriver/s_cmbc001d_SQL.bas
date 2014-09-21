Attribute VB_Name = "s_cmbc001d_SQL"
Option Explicit
'                                     2001/06/11
'================================================
' DBアクセス関数
' 定義内容: TBCMB011 (PG-ID管理)
' 参照　　: 060200_全テーブル
'================================================
#If False Then
'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmbc001d_Disp
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
    UPSPIN As Double                ' 上軸回転数
    DOWNSPIN As Double              ' 下軸回転数
    ROPRESS As String * 8           ' 炉内圧
    ARUGON As String * 7            ' アルゴン量
    AIMOIMIN As Double              ' ねらいOi（MIN)
    AIMOIMAX As Double              ' ねらいOi（MAX)
    HCCLASS As String * 7           ' HC種類
    HC As String * 3                ' HC
    AVEUPSPD As Double              ' 平均引上速度
    UPCNTL As String * 1            ' 引上制御
    BTMSHAPE As String * 1          ' ボトム形状
    MAGSTR As Long                  ' 磁場強度
    MAGPOS As Long                  ' 磁場位置
    CONDGRT As String * 10          ' 条件保証登録
    MODEL As String * 4             ' 機種
    UPMETHOD As String * 1          ' 引上方法
    UPCLASS As String * 2           ' 引上区分
    UPNUM As String * 1             ' 引上本数
    OPETIME As Long                 ' 運転時間
    WTRCOOL As String * 1           ' 水冷管要否
    PGID2 As String * 8             ' PG-ID（一本引）
    RCPT1 As String * 3             ' 対応レシピNo（T1)
    RCPT2 As String * 3             ' 対応レシピNo（T2)
    RCPT3 As String * 3             ' 対応レシピNo（T3)
    RCPT4 As String * 3             ' 対応レシピNo（T4)
    RCPT5 As String * 3             ' 対応レシピNo（T5)
    CNTL1 As String * 1             ' 制御項目（1）
    CNTL2 As String * 1             ' 制御項目（2）
    CNTL3 As String * 1             ' 制御項目（3）
    CNTL4 As String * 1             ' 制御項目（4）
    CNTL5 As String * 1             ' 制御項目（5）
    CNTL6 As String * 1             ' 制御項目（6）
    CNTL7 As String * 1             ' 制御項目（7）
    CNTL8 As String * 1             ' 制御項目（8）
    CNTL9 As String * 1             ' 制御項目（9）
    CNTL10 As String * 1            ' 制御項目（10）
    CNTL11 As String * 1            ' 制御項目（11）
    CNTL12 As String * 1            ' 制御項目（12）
    CNTL13 As String * 1            ' 制御項目（13）
    CNTL14 As String * 1            ' 制御項目（14）
    CNTL15 As String * 1            ' 制御項目（15）
    RUNCOND1 As String              ' 運転条件１
    RUNCOND2 As String              ' 運転条件２
'    TSTAFFID As String * 5          ' 登録社員ID
'    REGDATE As Date                 ' 登録日付
'    KSTAFFID As String * 8          ' 更新社員ID
'    UPDDATE As Date                 ' 更新日付
'    SENDFLAG As String * 1          ' 送信フラグ
'    SENDDATE As Date                ' 送信日付
End Type
#End If
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
              " MODEL, UPMETHOD, UPCLASS, UPNUM, OPETIME, WTRCOOL, PGID2, RCPT1, RCPT2, RCPT3, RCPT4, RCPT5, CNTL1, CNTL2," & _
              " CNTL3, CNTL4, CNTL5, CNTL6, CNTL7, CNTL8, CNTL9, CNTL10, CNTL11, CNTL12, CNTL13, CNTL14, CNTL15, RUNCOND1," & _
              " RUNCOND2, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
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
    sql = sql & "AIMOIMIN = " & .AIMOIMIN & ", "                ' ねらいiO(MIN)
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

Private Function StrNoNull(s$) As String
    If Trim$(s) = vbNullString Then
        StrNoNull = " "
    Else
        StrNoNull = Trim$(s)
    End If
End Function

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
