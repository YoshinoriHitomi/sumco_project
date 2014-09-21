Attribute VB_Name = "s_cmbc055_SQL"
Option Explicit
'
'================================================
' DBアクセス関数
' 定義内容: TBCMB019 (FRS校正情報)
' 参照　　: 060200_全テーブル
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmjc001j_Disp
    GOUKI       As String * 3       ' 号機
    INPDATE     As Date             ' 日付
    FTIROIL     As Double           ' FTIR（Oi低)
    FTIROIM     As Double           ' FTIR（Oi中）
    FTIROIH     As Double           ' FTIR（Oi高）
    MS1OIL      As Double           ' 測定サンプル1（Oi低)
    MS1OIM      As Double           ' 測定サンプル1（Oi中）
    MS1OIH      As Double           ' 測定サンプル1（Oi高）
    MS2OIL      As Double           ' 測定サンプル2（Oi低)
    MS2OIM      As Double           ' 測定サンプル2（Oi中）
    MS2OIH      As Double           ' 測定サンプル2（Oi高）
    MS3OIL      As Double           ' 測定サンプル3（Oi低)
    MS3OIM      As Double           ' 測定サンプル3（Oi中）
    MS3OIH      As Double           ' 測定サンプル3（Oi高）
    MS4OIL      As Double           ' 測定サンプル4（Oi低)
    MS4OIM      As Double           ' 測定サンプル4（Oi中）
    MS4OIH      As Double           ' 測定サンプル4（Oi高）
    MS5OIL      As Double           ' 測定サンプル5（Oi低)
    MS5OIM      As Double           ' 測定サンプル5（Oi中）
    MS5OIH      As Double           ' 測定サンプル5（Oi高）
    MSAVEOIL    As Double           ' 測定平均（Oi低)
    MSAVEOIM    As Double           ' 測定平均（Oi中）
    MSAVEOIH    As Double           ' 測定平均（Oi高）
    MSSGOIL     As Double           ' 測定σ（Oi低)
    MSSGOIM     As Double           ' 測定σ（Oi中）
    MSSGOIH     As Double           ' 測定σ（Oi高）
    MSPSGOIL    As Double           ' 測定AVE+σ（Oi低)
    MSPSGOIM    As Double           ' 測定AVE+σ（Oi中）
    MSPSGOIH    As Double           ' 測定AVE+σ（Oi高）
    MSNSGOIL    As Double           ' 測定AVE-σ（Oi低)
    MSNSGOIM    As Double           ' 測定AVE-σ（Oi中）
    MSNSGOIH    As Double           ' 測定AVE-σ（Oi高）
    MINOIL      As Double           ' MIN（Oi低)
    MINOIM      As Double           ' MIN（Oi中）
    MINOIH      As Double           ' MIN（Oi高）
    MAXOIL      As Double           ' MAX（Oi低)
    MAXOIM      As Double           ' MAX（Oi中）
    MAXOIH      As Double           ' MAX（Oi高）
    SGCK1OIL    As Double           ' σckサンプル1（Oi低)
    SGCK1OIM    As Double           ' σckサンプル1（Oi中）
    SGCK1OIH    As Double           ' σckサンプル1（Oi高）
    SGCK2OIL    As Double           ' σckサンプル2（Oi低)
    SGCK2OIM    As Double           ' σckサンプル2（Oi中）
    SGCK2OIH    As Double           ' σckサンプル2（Oi高）
    SGCK3OIL    As Double           ' σckサンプル3（Oi低)
    SGCK3OIM    As Double           ' σckサンプル3（Oi中）
    SGCK3OIH    As Double           ' σckサンプル3（Oi高）
    SGCK4OIL    As Double           ' σckサンプル4（Oi低)
    SGCK4OIM    As Double           ' σckサンプル4（Oi中）
    SGCK4OIH    As Double           ' σckサンプル4（Oi高）
    SGCK5OIL    As Double           ' σckサンプル5（Oi低)
    SGCK5OIM    As Double           ' σckサンプル5（Oi中）
    SGCK5OIH    As Double           ' σckサンプル5（Oi高）
    SGCKDOIL    As Double           ' σckデータ数（Oi低)
    SGCKDOIM    As Double           ' σckデータ数（Oi中）
    SGCKDOIH    As Double           ' σckデータ数（Oi高）
    SGCKAOIL    As Double           ' σck平均（Oi低)
    SGCKAAOIM   As Double           ' σck平均（Oi中）
    SGCKAOIH    As Double           ' σck平均（Oi高）
    SGNOIL      As Double           ' σckσ（Oi低)
    SGNOIM      As Double           ' σckσ（Oi中）
    SGNOIH      As Double           ' σckσ（Oi高）
    FTIRKOIL    As Double           ' FTIR換算（Oi低)
    FTIRKOIM    As Double           ' FTIR換算（Oi中）
    FTIRKOIH    As Double           ' FTIR換算（Oi高）
    EFFECTTM    As Integer          ' 有効時間
    YCOEF       As Double           ' ＦＴＩＲ換算式（Ｙ切片）
    XCOEF       As Double           ' ＦＴＩＲ換算式（Ｘ係数）
    RSQUARE     As Double           ' Ｒ２乗
    SGCKST      As Double           ' σ判定基準
    SGCKOIL     As String * 1       ' σ判定（Oi低)
    SGCKOIM     As String * 1       ' σ判定（Oi中）
    SGCKOIH     As String * 1       ' σ判定（Oi高）
    FTIRCKST    As Double           ' FTIR換算判定基準
    FTIRCKOIL   As String * 1       ' FTIR換算判定（Oi低)
    FTIRCKOIM   As String * 1       ' FTIR換算判定（Oi中）
    FTIRCKOIH   As String * 1       ' FTIR換算判定（Oi高）
    MS6OIL      As Double           ' 測定サンプル6（Oi低)
    MS6OIM      As Double           ' 測定サンプル6（Oi中）
    MS6OIH      As Double           ' 測定サンプル6（Oi高）
    SGCK6OIL    As Double           ' σckサンプル6（Oi低)
    SGCK6OIM    As Double           ' σckサンプル6（Oi中）
    SGCK6OIH    As Double           ' σckサンプル6（Oi高）
    CVOIL       As Double           ' CV(%)（Oi低)
    CVOIM       As Double           ' CV(%)（Oi中）
    CVOIH       As Double           ' CV(%)（Oi高）
  '  TSTAFFID As String * 8          ' 登録社員ID
  '  REGDATE As Date                 ' 登録日付
  '  KSTAFFID As String * 8          ' 更新社員ID
  '  UPDDATE As Date                 ' 更新日付
  '  SENDFLAG As String * 1          ' 送信フラグ
  '  SENDDATE As Date                ' 送信日付
End Type

'''''------------------------------------------------
''''' DBアクセス関数
'''''------------------------------------------------
''''
'''''概要      :テーブル「TBCMB019」から条件にあったレコードを抽出する
'''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'''''          :record        ,O  ,typ_cmjc001j_Disp ,抽出レコード
'''''          :GOUK          ,I  ,String       ,「号機」(SQLの抽出条件)
'''''          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'''''説明      :「号機」=引数で、かつ「日付」が最新のデータを抽出する
'''''履歴      :2001/06/20作成　長野
''''Public Function DBDRV_Getcmjc001j_Disp(record As typ_cmjc001j_Disp, GOUK$) As FUNCTION_RETURN
''''Dim sql As String       'SQL全体
''''Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
''''Dim sqlWhere As String  'SQLのWHERE部分
''''Dim sqlGroup As String  'SQLのGROUP部分
''''Dim rs As OraDynaset    'RecordSet
''''Dim recCnt As Long      'レコード数
''''Dim i As Long
''''
''''    DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_FAILURE
''''
''''    ''SQLを組み立てる
''''
''''    'エラーハンドラの設定
''''    On Error GoTo proc_err
''''    gErr.Push "s_cmzcF_cmjc001j_SQL.bas -- Function DBDRV_Getcmjc001j_Disp"
''''
''''    sqlBase = "Select GOUKI, MAX(INPDATE) ""INPDATE"", FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
''''              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
''''              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
''''              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
''''              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
''''              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE "
''''    sqlBase = sqlBase & "From TBCMB019"
''''    ''抽出条件(ｻﾝﾌﾟﾙNO)の取り出し
''''    sqlWhere = "WHERE(GOUKI=" & GOUK & ") "
''''    sqlGroup = "GROUP BY GOUKI"
''''    sql = sqlBase & sqlWhere & sqlGroup
''''
''''    ''データを抽出する
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''    If rs Is Nothing Then
''''        ReDim records(0)
''''        DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_FAILURE
''''        GoTo proc_exit
''''    End If
''''
''''    ''抽出結果を格納する
''''    With record
''''        .GOUKI = rs("GOUKI")             ' 号機
''''        .INPDATE = rs("INPDATE")         ' 日付
''''        .FTIRFZI = rs("FTIRFZI")         ' FTIR（FZ)
''''        .FTIRCZH = rs("FTIRCZH")         ' FTIR（CZ高）
''''        .FTIRCZC = rs("FTIRCZC")         ' FTIR（CZ中）
''''        .MS1FZ = rs("MS1FZ")             ' 測定サンプル1（FZ)
''''        .MS1CZ1 = rs("MS1CZ1")           ' 測定サンプル1（CZ-1)
''''        .MS1CZ2 = rs("MS1CZ2")           ' 測定サンプル1（CZ-2)
''''        .MS2FZ = rs("MS2FZ")             ' 測定サンプル2（FZ)
''''        .MS2CZ1 = rs("MS2CZ1")           ' 測定サンプル2（CZ-1)
''''        .MS2CZ2 = rs("MS2CZ2")           ' 測定サンプル2（CZ-2)
''''        .MS3FZ = rs("MS3FZ")             ' 測定サンプル3（FZ)
''''        .MS3CZ1 = rs("MS3CZ1")           ' 測定サンプル3（CZ-1)
''''        .MS3CZ2 = rs("MS3CZ2")           ' 測定サンプル3（CZ-2)
''''        .MS4FZ = rs("MS4FZ")             ' 測定サンプル4（FZ)
''''        .MS4CZ1 = rs("MS4CZ1")           ' 測定サンプル4（CZ-1)
''''        .MS4CZ2 = rs("MS4CZ2")           ' 測定サンプル4（CZ-2)
''''        .MS5FZ = rs("MS5FZ")             ' 測定サンプル5（FZ)
''''        .MS5CZ1 = rs("MS5CZ1")           ' 測定サンプル5（CZ-1)
''''        .MS5CZ2 = rs("MS5CZ2")           ' 測定サンプル5（CZ-2)
''''        .MSAVEFZ = rs("MSAVEFZ")         ' 測定平均（FZ）
''''        .MSAVECZ1 = rs("MSAVECZ1")       ' 測定平均（CZ-1）
''''        .MSAVECZ2 = rs("MSAVECZ2")       ' 測定平均（CZ-2）
''''        .MSSGFZ = rs("MSSGFZ")           ' 測定σ（FZ）
''''        .MSSGCZ1 = rs("MSSGCZ1")         ' 測定σ（CZ-1）
''''        .MSSGCZ2 = rs("MSSGCZ2")         ' 測定σ（CZ-2）
''''        .MSPSGFZ = rs("MSPSGFZ")         ' 測定AVE+σ（FZ）
''''        .MSPSGCZ1 = rs("MSPSGCZ1")       ' 測定AVE+σ（CZ-1）
''''        .MSPSGCZ2 = rs("MSPSGCZ2")       ' 測定AVE+σ（CZ-2）
''''        .MSNSGFZ = rs("MSNSGFZ")         ' 測定AVE-σ（FZ）
''''        .MSNSGCZ1 = rs("MSNSGCZ1")       ' 測定AVE-σ（CZ-1）
''''        .MSNSGCZ2 = rs("MSNSGCZ2")       ' 測定AVE-σ（CZ-2）
''''        .MINFZ = rs("MINFZ")             ' MIN（FZ）
''''        .MINCZ1 = rs("MINCZ1")           ' MIN（CZ-1）
''''        .MINCZ2 = rs("MINCZ2")           ' MIN（CZ-2）
''''        .MAXFZ = rs("MAXFZ")             ' MAX（FZ）
''''        .MAXCZ1 = rs("MAXCZ1")           ' MAX（CZ-1）
''''        .MAXCZ2 = rs("MAXCZ2")           ' MAX（CZ-2）
''''        .SGCK1FZ = rs("SGCK1FZ")         ' σckサンプル1（FZ)
''''        .SGCK1CZ1 = rs("SGCK1CZ1")       ' σckサンプル1（CZ-1)
''''        .SGCK1CZ2 = rs("SGCK1CZ2")       ' σckサンプル1（CZ-2)
''''        .SGCK2FZ = rs("SGCK2FZ")         ' σckサンプル2（FZ)
''''        .SGCK2CZ1 = rs("SGCK2CZ1")       ' σckサンプル2（CZ-1)
''''        .SGCK2CZ2 = rs("SGCK2CZ2")       ' σckサンプル2（CZ-2)
''''        .SGCK3FZ = rs("SGCK3FZ")         ' σckサンプル3（FZ)
''''        .SGCK3CZ1 = rs("SGCK3CZ1")       ' σckサンプル3（CZ-1)
''''        .SGCK3CZ2 = rs("SGCK3CZ2")       ' σckサンプル3（CZ-2)
''''        .SGCK4FZ = rs("SGCK4FZ")         ' σckサンプル4（FZ)
''''        .SGCK4CZ1 = rs("SGCK4CZ1")       ' σckサンプル4（CZ-1)
''''        .SGCK4CZ2 = rs("SGCK4CZ2")       ' σckサンプル4（CZ-2)
''''        .SGCK5FZ = rs("SGCK5FZ")         ' σckサンプル5（FZ)
''''        .SGCK5CZ1 = rs("SGCK5CZ1")       ' σckサンプル5（CZ-1)
''''        .SGCK5CZ2 = rs("SGCK5CZ2")       ' σckサンプル5（CZ-2)
''''        .SGCKDFZ = rs("SGCKDFZ")         ' σckデータ数（FZ）
''''        .SGCKDCZ1 = rs("SGCKDCZ1")       ' σckデータ数（CZ-1）
''''        .SGCKDCZ2 = rs("SGCKDCZ2")       ' σckデータ数（CZ-2）
''''        .SGCKAFZ = rs("SGCKAFZ")         ' σck平均（FZ）
''''        .SGCKAACZ1 = rs("SGCKAACZ1")     ' σck平均（CZ-1）
''''        .SGCKACZ2 = rs("SGCKACZ2")       ' σck平均（CZ-2）
''''        .SGNFZ = rs("SGNFZ")             ' σckσ（FZ）
''''        .SGNCZ1 = rs("SGNCZ1")           ' σckσ CZ-1）
''''        .SGNCZ2 = rs("SGNCZ2")           ' σckσ（CZ-2）
''''        .FTIRFZ = rs("FTIRFZ")           ' FTIR換算（FZ）
''''        .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR換算（CZ-1）
''''        .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR換算（CZ-2）
''''        .EFFECTTM = rs("EFFECTTM")       ' 有効時間
''''        .YCOEF = rs("YCOEF")             ' ＦＴＩＲ換算式（Ｙ切片）
''''        .XCOEF = rs("XCOEF")             ' ＦＴＩＲ換算式（Ｘ係数）
''''        .RSQUARE = rs("RSQUARE")         ' Ｒ２乗
''''    End With
''''    rs.Close
''''
''''    DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_SUCCESS
''''
''''proc_exit:
''''    '終了
''''    gErr.Pop
''''    Exit Function
''''
''''proc_err:
''''    'エラーハンドラ
''''    Debug.Print "====== Error SQL ======"
''''    Debug.Print sql
''''    gErr.HandleError
''''    Resume proc_exit
''''End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :引数で渡されたレコードをTBCMB019に追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :record        ,I  ,typ_cmjc001j_Disp ,抽出レコード
'          :TSTAFFID      ,I  ,String       ,登録社員ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :
Public Function DBDRV_Getcmjc001j_Exec(record As typ_cmjc001j_Disp, TSTAFFID$) As FUNCTION_RETURN
    Dim sql As String           'SQL全体
    Dim SetDate  As Variant     '入力日付

    DBDRV_Getcmjc001j_Exec = FUNCTION_RETURN_FAILURE
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001j_SQL.bas -- Function DBDRV_Getcmjc001j_Exec"

    SetDate = Format$(record.INPDATE, "yyyy-mm-dd hh:mm:ss")
  
    ''SQLを組み立てる
    sql = "Insert into TBCMB019 ("
    sql = sql & "  GOUKI"                   '' 号機
    sql = sql & ", INPDATE"                 '' 日付
    sql = sql & ", FTIROIL"                 '' FTIR（Oi低)
    sql = sql & ", FTIROIM"                 '' FTIR（Oi中)
    sql = sql & ", FTIROIH"                 '' FTIR（Oi高)
    sql = sql & ", MS1OIL"                  '' 測定サンプル1（Oi低)
    sql = sql & ", MS1OIM"                  '' 測定サンプル1（Oi中)
    sql = sql & ", MS1OIH"                  '' 測定サンプル1（Oi高)
    sql = sql & ", MS2OIL"                  '' 測定サンプル2（Oi低)
    sql = sql & ", MS2OIM"                  '' 測定サンプル2（Oi中)
    sql = sql & ", MS2OIH"                  '' 測定サンプル2（Oi高)
    sql = sql & ", MS3OIL"                  '' 測定サンプル3（Oi低)
    sql = sql & ", MS3OIM"                  '' 測定サンプル3（Oi中)
    sql = sql & ", MS3OIH"                  '' 測定サンプル3（Oi高)
    sql = sql & ", MS4OIL"                  '' 測定サンプル4（Oi低)
    sql = sql & ", MS4OIM"                  '' 測定サンプル4（Oi中)
    sql = sql & ", MS4OIH"                  '' 測定サンプル4（Oi高)
    sql = sql & ", MS5OIL"                  '' 測定サンプル5（Oi低)
    sql = sql & ", MS5OIM"                  '' 測定サンプル5（Oi中)
    sql = sql & ", MS5OIH"                  '' 測定サンプル5（Oi高)
    sql = sql & ", MSAVEOIL"                '' 測定平均（Oi低)
    sql = sql & ", MSAVEOIM"                '' 測定平均（Oi中)
    sql = sql & ", MSAVEOIH"                '' 測定平均（Oi高)
    sql = sql & ", MSSGOIL"                 '' 測定σ（Oi低)
    sql = sql & ", MSSGOIM"                 '' 測定σ（Oi中)
    sql = sql & ", MSSGOIH"                 '' 測定σ（Oi高)
    sql = sql & ", MSPSGOIL"                '' 測定AVE+σ（Oi低)
    sql = sql & ", MSPSGOIM"                '' 測定AVE+σ（Oi中)
    sql = sql & ", MSPSGOIH"                '' 測定AVE+σ（Oi高)
    sql = sql & ", MSNSGOIL"                '' 測定AVE-σ（Oi低)
    sql = sql & ", MSNSGOIM"                '' 測定AVE-σ（Oi中)
    sql = sql & ", MSNSGOIH"                '' 測定AVE-σ（Oi高)
    sql = sql & ", MINOIL"                  '' MIN（Oi低)
    sql = sql & ", MINOIM"                  '' MIN（Oi中)
    sql = sql & ", MINOIH"                  '' MIN（Oi高)
    sql = sql & ", MAXOIL"                  '' MAX（Oi低)
    sql = sql & ", MAXOIM"                  '' MAX（Oi中)
    sql = sql & ", MAXOIH"                  '' MAX（Oi高)
    sql = sql & ", SGCK1OIL"                '' σckサンプル1（Oi低)
    sql = sql & ", SGCK1OIM"                '' σckサンプル1（Oi中)
    sql = sql & ", SGCK1OIH"                '' σckサンプル1（Oi高)
    sql = sql & ", SGCK2OIL"                '' σckサンプル2（Oi低)
    sql = sql & ", SGCK2OIM"                '' σckサンプル2（Oi中)
    sql = sql & ", SGCK2OIH"                '' σckサンプル2（Oi高)
    sql = sql & ", SGCK3OIL"                '' σckサンプル3（Oi低)
    sql = sql & ", SGCK3OIM"                '' σckサンプル3（Oi中)
    sql = sql & ", SGCK3OIH"                '' σckサンプル3（Oi高)
    sql = sql & ", SGCK4OIL"                '' σckサンプル4（Oi低)
    sql = sql & ", SGCK4OIM"                '' σckサンプル4（Oi中)
    sql = sql & ", SGCK4OIH"                '' σckサンプル4（Oi高)
    sql = sql & ", SGCK5OIL"                '' σckサンプル5（Oi低)
    sql = sql & ", SGCK5OIM"                '' σckサンプル5（Oi中)
    sql = sql & ", SGCK5OIH"                '' σckサンプル5（Oi高)
    sql = sql & ", SGCKDOIL"                '' σckデータ数（Oi低)
    sql = sql & ", SGCKDOIM"                '' σckデータ数（Oi中)
    sql = sql & ", SGCKDOIH"                '' σckデータ数（Oi高)
    sql = sql & ", SGCKAOIL"                '' σck平均（Oi低)
    sql = sql & ", SGCKAAOIM"               '' σck平均（Oi中)
    sql = sql & ", SGCKAOIH"                '' σck平均（Oi高)
    sql = sql & ", SGNOIL"                  '' σckσ（Oi低)
    sql = sql & ", SGNOIM"                  '' σckσ（Oi中)
    sql = sql & ", SGNOIH"                  '' σckσ（Oi高)
    sql = sql & ", FTIRKOIL"                '' FTIR換算（Oi低)
    sql = sql & ", FTIRKOIM"                '' FTIR換算（Oi中)
    sql = sql & ", FTIRKOIH"                '' FTIR換算（Oi高)
    sql = sql & ", EFFECTTM"                '' 有効時間
    sql = sql & ", YCOEF"                   '' ＦＴＩＲ換算式（Ｙ切片）
    sql = sql & ", XCOEF"                   '' ＦＴＩＲ換算式（Ｘ係数）
    sql = sql & ", RSQUARE"                 '' Ｒ２乗
    sql = sql & ", SGCKST"                  '' σ判定基準
    sql = sql & ", SGCKOIL"                 '' σ判定（Oi低)
    sql = sql & ", SGCKOIM"                 '' σ判定（Oi中)
    sql = sql & ", SGCKOIH"                 '' σ判定（Oi高)
    sql = sql & ", FTIRCKST"                '' FTIR換算判定基準
    sql = sql & ", FTIRCKOIL"               '' FTIR換算判定（Oi低)
    sql = sql & ", FTIRCKOIM"               '' FTIR換算判定（Oi中)
    sql = sql & ", FTIRCKOIH"               '' FTIR換算判定（Oi高)
    sql = sql & ", MS6OIL"                  '' 測定サンプル6（Oi低)
    sql = sql & ", MS6OIM"                  '' 測定サンプル6（Oi中)
    sql = sql & ", MS6OIH"                  '' 測定サンプル6（Oi高)
    sql = sql & ", SGCK6OIL"                '' σckサンプル6（Oi低)
    sql = sql & ", SGCK6OIM"                '' σckサンプル6（Oi中)
    sql = sql & ", SGCK6OIH"                '' σckサンプル6（Oi高)
    sql = sql & ", CVOIL"                   '' CV（Oi低)
    sql = sql & ", CVOIM"                   '' CV（Oi中)
    sql = sql & ", CVOIH"                   '' CV（Oi高)
    sql = sql & ", TSTAFFID"                '' 登録社員ID
    sql = sql & ", REGDATE"                 '' 登録日付
    sql = sql & ", KSTAFFID"                '' 更新社員ID
    sql = sql & ", UPDDATE"                 '' 更新日付
    sql = sql & ", SENDFLAG"                '' 送信フラグ
    sql = sql & ", SENDDATE"                '' 送信日付
    sql = sql & ")"
    
    sql = sql & "Values("
    sql = sql & "'" & record.GOUKI & "'"                                        '' 号機
    sql = sql & ", " & "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss')"     '' 日付
    sql = sql & ", " & record.FTIROIL                                           '' FTIR（Oi低)
    sql = sql & ", " & record.FTIROIM                                           '' FTIR（Oi中)
    sql = sql & ", " & record.FTIROIH                                           '' FTIR（Oi高)
    sql = sql & ", " & record.MS1OIL                                            '' 測定サンプル1（Oi低)
    sql = sql & ", " & record.MS1OIM                                            '' 測定サンプル1（Oi中)
    sql = sql & ", " & record.MS1OIH                                            '' 測定サンプル1（Oi高)
    sql = sql & ", " & record.MS2OIL                                            '' 測定サンプル2（Oi低)
    sql = sql & ", " & record.MS2OIM                                            '' 測定サンプル2（Oi中)
    sql = sql & ", " & record.MS2OIH                                            '' 測定サンプル2（Oi高)
    sql = sql & ", " & record.MS3OIL                                            '' 測定サンプル3（Oi低)
    sql = sql & ", " & record.MS3OIM                                            '' 測定サンプル3（Oi中)
    sql = sql & ", " & record.MS3OIH                                            '' 測定サンプル3（Oi高)
    sql = sql & ", " & record.MS4OIL                                            '' 測定サンプル4（Oi低)
    sql = sql & ", " & record.MS4OIM                                            '' 測定サンプル4（Oi中)
    sql = sql & ", " & record.MS4OIH                                            '' 測定サンプル4（Oi高)
    sql = sql & ", " & record.MS5OIL                                            '' 測定サンプル5（Oi低)
    sql = sql & ", " & record.MS5OIM                                            '' 測定サンプル5（Oi中)
    sql = sql & ", " & record.MS5OIH                                            '' 測定サンプル5（Oi高)
    sql = sql & ", " & record.MSAVEOIL                                          '' 測定平均（Oi低)
    sql = sql & ", " & record.MSAVEOIM                                          '' 測定平均（Oi中)
    sql = sql & ", " & record.MSAVEOIH                                          '' 測定平均（Oi高)
    sql = sql & ", " & record.MSSGOIL                                           '' 測定σ（Oi低)
    sql = sql & ", " & record.MSSGOIM                                           '' 測定σ（Oi中)
    sql = sql & ", " & record.MSSGOIH                                           '' 測定σ（Oi高)
    sql = sql & ", " & record.MSPSGOIL                                          '' 測定AVE+σ（Oi低)
    sql = sql & ", " & record.MSPSGOIM                                          '' 測定AVE+σ（Oi中)
    sql = sql & ", " & record.MSPSGOIH                                          '' 測定AVE+σ（Oi高)
    sql = sql & ", " & record.MSNSGOIL                                          '' 測定AVE-σ（Oi低)
    sql = sql & ", " & record.MSNSGOIM                                          '' 測定AVE-σ（Oi中)
    sql = sql & ", " & record.MSNSGOIH                                          '' 測定AVE-σ（Oi高)
    sql = sql & ", " & record.MINOIL                                            '' MIN（Oi低)
    sql = sql & ", " & record.MINOIM                                            '' MIN（Oi中)
    sql = sql & ", " & record.MINOIH                                            '' MIN（Oi高)
    sql = sql & ", " & record.MAXOIL                                            '' MAX（Oi低)
    sql = sql & ", " & record.MAXOIM                                            '' MAX（Oi中)
    sql = sql & ", " & record.MAXOIH                                            '' MAX（Oi高)
    sql = sql & ", " & record.SGCK1OIL                                          '' σckサンプル1（Oi低)
    sql = sql & ", " & record.SGCK1OIM                                          '' σckサンプル1（Oi中)
    sql = sql & ", " & record.SGCK1OIH                                          '' σckサンプル1（Oi高)
    sql = sql & ", " & record.SGCK2OIL                                          '' σckサンプル2（Oi低)
    sql = sql & ", " & record.SGCK2OIM                                          '' σckサンプル2（Oi中)
    sql = sql & ", " & record.SGCK2OIH                                          '' σckサンプル2（Oi高)
    sql = sql & ", " & record.SGCK3OIL                                          '' σckサンプル3（Oi低)
    sql = sql & ", " & record.SGCK3OIM                                          '' σckサンプル3（Oi中)
    sql = sql & ", " & record.SGCK3OIH                                          '' σckサンプル3（Oi高)
    sql = sql & ", " & record.SGCK4OIL                                          '' σckサンプル4（Oi低)
    sql = sql & ", " & record.SGCK4OIM                                          '' σckサンプル4（Oi中)
    sql = sql & ", " & record.SGCK4OIH                                          '' σckサンプル4（Oi高)
    sql = sql & ", " & record.SGCK5OIL                                          '' σckサンプル5（Oi低)
    sql = sql & ", " & record.SGCK5OIM                                          '' σckサンプル5（Oi中)
    sql = sql & ", " & record.SGCK5OIH                                          '' σckサンプル5（Oi高)
    sql = sql & ", " & record.SGCKDOIL                                          '' σckデータ数（Oi低)
    sql = sql & ", " & record.SGCKDOIM                                          '' σckデータ数（Oi中)
    sql = sql & ", " & record.SGCKDOIH                                          '' σckデータ数（Oi高)
    sql = sql & ", " & record.SGCKAOIL                                          '' σck平均（Oi低)
    sql = sql & ", " & record.SGCKAAOIM                                         '' σck平均（Oi中)
    sql = sql & ", " & record.SGCKAOIH                                          '' σck平均（Oi高)
    sql = sql & ", " & record.SGNOIL                                            '' σckσ（Oi低)
    sql = sql & ", " & record.SGNOIM                                            '' σckσ（Oi中)
    sql = sql & ", " & record.SGNOIH                                            '' σckσ（Oi高)
    sql = sql & ", " & record.FTIRKOIL                                          '' FTIR換算（Oi低)
    sql = sql & ", " & record.FTIRKOIM                                          '' FTIR換算（Oi中)
    sql = sql & ", " & record.FTIRKOIH                                          '' FTIR換算（Oi高)
    sql = sql & ", " & record.EFFECTTM                                          '' 有効時間
    sql = sql & ", " & record.YCOEF                                             '' ＦＴＩＲ換算式（Ｙ切片）
    sql = sql & ", " & record.XCOEF                                             '' ＦＴＩＲ換算式（Ｘ係数）
    sql = sql & ", " & record.RSQUARE                                           '' Ｒ２乗
    sql = sql & ", " & record.SGCKST                                            '' σ判定基準
    sql = sql & ", '" & record.SGCKOIL & "'"                                    '' σ判定（Oi低)
    sql = sql & ", '" & record.SGCKOIM & "'"                                    '' σ判定（Oi中)
    sql = sql & ", '" & record.SGCKOIH & "'"                                    '' σ判定（Oi高)
    sql = sql & ", " & record.FTIRCKST                                          '' FTIR換算判定基準
    sql = sql & ", '" & record.FTIRCKOIL & "'"                                  '' FTIR換算判定（Oi低)
    sql = sql & ", '" & record.FTIRCKOIM & "'"                                  '' FTIR換算判定（Oi中)
    sql = sql & ", '" & record.FTIRCKOIH & "'"                                  '' FTIR換算判定（Oi高)
    sql = sql & ", " & record.MS6OIL                                            '' 測定サンプル6（Oi低)
    sql = sql & ", " & record.MS6OIM                                            '' 測定サンプル6（Oi中)
    sql = sql & ", " & record.MS6OIH                                            '' 測定サンプル6（Oi高)
    sql = sql & ", " & record.SGCK6OIL                                          '' σckサンプル6（Oi低)
    sql = sql & ", " & record.SGCK6OIM                                          '' σckサンプル6（Oi中)
    sql = sql & ", " & record.SGCK6OIH                                          '' σckサンプル6（Oi高)
    sql = sql & ", " & record.CVOIL                                             '' CV（Oi低)
    sql = sql & ", " & record.CVOIM                                             '' CV（Oi中)
    sql = sql & ", " & record.CVOIH                                             '' CV（Oi高)
    sql = sql & ", '" & TSTAFFID & "'"                                          '' 登録社員ID
    sql = sql & ", SYSDATE"                                                     '' 登録日付
    sql = sql & ", ' '"                                                         '' 更新社員ID
    sql = sql & ", SYSDATE"                                                     '' 更新日付
    sql = sql & ", '0'"                                                         '' 送信フラグ
    sql = sql & ", SYSDATE"                                                     '' 送信日付
    sql = sql & ")"
  
    '' ■SQLの実行
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001j_Exec = FUNCTION_RETURN_SUCCESS

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

'概要      :データ変換を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                ,説明
'          :tblLeft       ,IO   ,typ_TBCMB019      ,テーブルデータ１
'          :tblRight      ,IO   ,typ_cmjc001j_Disp ,テーブルデータ２
'          :bFlg          ,I   ,Boolean           ,TRUE:引数１データ→引数２データへの変換  FALSE:引数１データ←引数２データへの変換
'説明      :
Public Sub ConvDate_F_cmjc001j_a(tblLeft As typ_TBCMB019, tblRight As typ_cmjc001j_Disp, bFlg As Boolean)
    
    If bFlg = True Then
        With tblRight
            .GOUKI = tblLeft.GOUKI
            .INPDATE = tblLeft.INPDATE
            .FTIROIL = tblLeft.FTIROIL
            .FTIROIM = tblLeft.FTIROIM
            .FTIROIH = tblLeft.FTIROIH
            .MS1OIL = tblLeft.MS1OIL
            .MS1OIM = tblLeft.MS1OIM
            .MS1OIH = tblLeft.MS1OIH
            .MS2OIL = tblLeft.MS2OIL
            .MS2OIM = tblLeft.MS2OIM
            .MS2OIH = tblLeft.MS2OIH
            .MS3OIL = tblLeft.MS3OIL
            .MS3OIM = tblLeft.MS3OIM
            .MS3OIH = tblLeft.MS3OIH
            .MS4OIL = tblLeft.MS4OIL
            .MS4OIM = tblLeft.MS4OIM
            .MS4OIH = tblLeft.MS4OIH
            .MS5OIL = tblLeft.MS5OIL
            .MS5OIM = tblLeft.MS5OIM
            .MS5OIH = tblLeft.MS5OIH
            .MSAVEOIL = tblLeft.MSAVEOIL
            .MSAVEOIM = tblLeft.MSAVEOIM
            .MSAVEOIH = tblLeft.MSAVEOIH
            .MSSGOIL = tblLeft.MSSGOIL
            .MSSGOIM = tblLeft.MSSGOIM
            .MSSGOIH = tblLeft.MSSGOIH
            .MSPSGOIL = tblLeft.MSPSGOIL
            .MSPSGOIM = tblLeft.MSPSGOIM
            .MSPSGOIH = tblLeft.MSPSGOIH
            .MSNSGOIL = tblLeft.MSNSGOIL
            .MSNSGOIM = tblLeft.MSNSGOIM
            .MSNSGOIH = tblLeft.MSNSGOIH
            .MINOIL = tblLeft.MINOIL
            .MINOIM = tblLeft.MINOIM
            .MINOIH = tblLeft.MINOIH
            .MAXOIL = tblLeft.MAXOIL
            .MAXOIM = tblLeft.MAXOIM
            .MAXOIH = tblLeft.MAXOIH
            .SGCK1OIL = tblLeft.SGCK1OIL
            .SGCK1OIM = tblLeft.SGCK1OIM
            .SGCK1OIH = tblLeft.SGCK1OIH
            .SGCK2OIL = tblLeft.SGCK2OIL
            .SGCK2OIM = tblLeft.SGCK2OIM
            .SGCK2OIH = tblLeft.SGCK2OIH
            .SGCK3OIL = tblLeft.SGCK3OIL
            .SGCK3OIM = tblLeft.SGCK3OIM
            .SGCK3OIH = tblLeft.SGCK3OIH
            .SGCK4OIL = tblLeft.SGCK4OIL
            .SGCK4OIM = tblLeft.SGCK4OIM
            .SGCK4OIH = tblLeft.SGCK4OIH
            .SGCK5OIL = tblLeft.SGCK5OIL
            .SGCK5OIM = tblLeft.SGCK5OIM
            .SGCK5OIH = tblLeft.SGCK5OIH
            .SGCKDOIL = tblLeft.SGCKDOIL
            .SGCKDOIM = tblLeft.SGCKDOIM
            .SGCKDOIH = tblLeft.SGCKDOIH
            .SGCKAOIL = tblLeft.SGCKAOIL
            .SGCKAAOIM = tblLeft.SGCKAAOIM
            .SGCKAOIH = tblLeft.SGCKAOIH
            .SGNOIL = tblLeft.SGNOIL
            .SGNOIM = tblLeft.SGNOIM
            .SGNOIH = tblLeft.SGNOIH
            .FTIRKOIL = tblLeft.FTIRKOIL
            .FTIRKOIM = tblLeft.FTIRKOIM
            .FTIRKOIH = tblLeft.FTIRKOIH
            .EFFECTTM = tblLeft.EFFECTTM
            .YCOEF = tblLeft.YCOEF
            .XCOEF = tblLeft.XCOEF
            .RSQUARE = tblLeft.RSQUARE
            .SGCKST = tblLeft.SGCKST
            .SGCKOIL = tblLeft.SGCKOIL
            .SGCKOIM = tblLeft.SGCKOIM
            .SGCKOIH = tblLeft.SGCKOIH
            .FTIRCKST = tblLeft.FTIRCKST
            .FTIRCKOIL = tblLeft.FTIRCKOIL
            .FTIRCKOIM = tblLeft.FTIRCKOIM
            .FTIRCKOIH = tblLeft.FTIRCKOIH
            .MS6OIL = tblLeft.MS6OIL
            .MS6OIM = tblLeft.MS6OIM
            .MS6OIH = tblLeft.MS6OIH
            .SGCK6OIL = tblLeft.SGCK6OIL
            .SGCK6OIM = tblLeft.SGCK6OIM
            .SGCK6OIH = tblLeft.SGCK6OIH
            .CVOIL = tblLeft.CVOIL
            .CVOIM = tblLeft.CVOIM
            .CVOIH = tblLeft.CVOIH
        
        End With
    Else
        With tblLeft
            .GOUKI = tblRight.GOUKI
            .INPDATE = tblRight.INPDATE
            .FTIROIL = tblRight.FTIROIL
            .FTIROIM = tblRight.FTIROIM
            .FTIROIH = tblRight.FTIROIH
            .MS1OIL = tblRight.MS1OIL
            .MS1OIM = tblRight.MS1OIM
            .MS1OIH = tblRight.MS1OIH
            .MS2OIL = tblRight.MS2OIL
            .MS2OIM = tblRight.MS2OIM
            .MS2OIH = tblRight.MS2OIH
            .MS3OIL = tblRight.MS3OIL
            .MS3OIM = tblRight.MS3OIM
            .MS3OIH = tblRight.MS3OIH
            .MS4OIL = tblRight.MS4OIL
            .MS4OIM = tblRight.MS4OIM
            .MS4OIH = tblRight.MS4OIH
            .MS5OIL = tblRight.MS5OIL
            .MS5OIM = tblRight.MS5OIM
            .MS5OIH = tblRight.MS5OIH
            .MSAVEOIL = tblRight.MSAVEOIL
            .MSAVEOIM = tblRight.MSAVEOIM
            .MSAVEOIH = tblRight.MSAVEOIH
            .MSSGOIL = tblRight.MSSGOIL
            .MSSGOIM = tblRight.MSSGOIM
            .MSSGOIH = tblRight.MSSGOIH
            .MSPSGOIL = tblRight.MSPSGOIL
            .MSPSGOIM = tblRight.MSPSGOIM
            .MSPSGOIH = tblRight.MSPSGOIH
            .MSNSGOIL = tblRight.MSNSGOIL
            .MSNSGOIM = tblRight.MSNSGOIM
            .MSNSGOIH = tblRight.MSNSGOIH
            .MINOIL = tblRight.MINOIL
            .MINOIM = tblRight.MINOIM
            .MINOIH = tblRight.MINOIH
            .MAXOIL = tblRight.MAXOIL
            .MAXOIM = tblRight.MAXOIM
            .MAXOIH = tblRight.MAXOIH
            .SGCK1OIL = tblRight.SGCK1OIL
            .SGCK1OIM = tblRight.SGCK1OIM
            .SGCK1OIH = tblRight.SGCK1OIH
            .SGCK2OIL = tblRight.SGCK2OIL
            .SGCK2OIM = tblRight.SGCK2OIM
            .SGCK2OIH = tblRight.SGCK2OIH
            .SGCK3OIL = tblRight.SGCK3OIL
            .SGCK3OIM = tblRight.SGCK3OIM
            .SGCK3OIH = tblRight.SGCK3OIH
            .SGCK4OIL = tblRight.SGCK4OIL
            .SGCK4OIM = tblRight.SGCK4OIM
            .SGCK4OIH = tblRight.SGCK4OIH
            .SGCK5OIL = tblRight.SGCK5OIL
            .SGCK5OIM = tblRight.SGCK5OIM
            .SGCK5OIH = tblRight.SGCK5OIH
            .SGCKDOIL = tblRight.SGCKDOIL
            .SGCKDOIM = tblRight.SGCKDOIM
            .SGCKDOIH = tblRight.SGCKDOIH
            .SGCKAOIL = tblRight.SGCKAOIL
            .SGCKAAOIM = tblRight.SGCKAAOIM
            .SGCKAOIH = tblRight.SGCKAOIH
            .SGNOIL = tblRight.SGNOIL
            .SGNOIM = tblRight.SGNOIM
            .SGNOIH = tblRight.SGNOIH
            .FTIRKOIL = tblRight.FTIRKOIL
            .FTIRKOIM = tblRight.FTIRKOIM
            .FTIRKOIH = tblRight.FTIRKOIH
            .EFFECTTM = tblRight.EFFECTTM
            .YCOEF = tblRight.YCOEF
            .XCOEF = tblRight.XCOEF
            .RSQUARE = tblRight.RSQUARE
            .SGCKST = tblRight.SGCKST
            .SGCKOIL = tblRight.SGCKOIL
            .SGCKOIM = tblRight.SGCKOIM
            .SGCKOIH = tblRight.SGCKOIH
            .FTIRCKST = tblRight.FTIRCKST
            .FTIRCKOIL = tblRight.FTIRCKOIL
            .FTIRCKOIM = tblRight.FTIRCKOIM
            .FTIRCKOIH = tblRight.FTIRCKOIH
            .MS6OIL = tblRight.MS6OIL
            .MS6OIM = tblRight.MS6OIM
            .MS6OIH = tblRight.MS6OIH
            .SGCK6OIL = tblRight.SGCK6OIL
            .SGCK6OIM = tblRight.SGCK6OIM
            .SGCK6OIH = tblRight.SGCK6OIH
            .CVOIL = tblRight.CVOIL
            .CVOIM = tblRight.CVOIM
            .CVOIH = tblRight.CVOIH
        
        End With
    End If

End Sub

'''''------------------------------------------------
''''' DBアクセス関数
'''''------------------------------------------------
''''
'''''概要      :テーブル「TBCMB019」から条件にあったレコードを抽出する
'''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'''''          :records()     ,O  ,typ_TBCMB019 ,抽出レコード
'''''          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'''''          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'''''          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'''''説明      :
'''''履歴      :2001/08/24作成　野村
''''Public Function DBDRV_GetTBCMB019(records() As typ_TBCMB019, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
''''Dim sql As String       'SQL全体
''''Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
''''Dim rs As OraDynaset    'RecordSet
''''Dim recCnt As Long      'レコード数
''''Dim i As Long
''''
''''    ''SQLを組み立てる
''''    sqlBase = "Select GOUKI, INPDATE, FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
''''              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
''''              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
''''              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
''''              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
''''              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG," & _
''''              " SENDDATE "
''''    sqlBase = sqlBase & "From TBCMB019"
''''    sql = sqlBase
''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
''''        sql = sql & " " & sqlWhere & " " & sqlOrder
''''    End If
''''
''''    ''データを抽出する
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''    If rs Is Nothing Then
''''        ReDim records(0)
''''        DBDRV_GetTBCMB019 = FUNCTION_RETURN_FAILURE
''''        Exit Function
''''    End If
''''
''''    ''抽出結果を格納する
''''    recCnt = rs.RecordCount
''''    ReDim records(recCnt)
''''    For i = 1 To recCnt
''''        With records(i)
''''            .GOUKI = rs("GOUKI")             ' 号機
''''            .INPDATE = rs("INPDATE")         ' 日付
''''            .FTIRFZI = rs("FTIRFZI")         ' FTIR（FZ)
''''            .FTIRCZH = rs("FTIRCZH")         ' FTIR（CZ高）
''''            .FTIRCZC = rs("FTIRCZC")         ' FTIR（CZ中）
''''            .MS1FZ = rs("MS1FZ")             ' 測定サンプル1（FZ)
''''            .MS1CZ1 = rs("MS1CZ1")           ' 測定サンプル1（CZ-1)
''''            .MS1CZ2 = rs("MS1CZ2")           ' 測定サンプル1（CZ-2)
''''            .MS2FZ = rs("MS2FZ")             ' 測定サンプル2（FZ)
''''            .MS2CZ1 = rs("MS2CZ1")           ' 測定サンプル2（CZ-1)
''''            .MS2CZ2 = rs("MS2CZ2")           ' 測定サンプル2（CZ-2)
''''            .MS3FZ = rs("MS3FZ")             ' 測定サンプル3（FZ)
''''            .MS3CZ1 = rs("MS3CZ1")           ' 測定サンプル3（CZ-1)
''''            .MS3CZ2 = rs("MS3CZ2")           ' 測定サンプル3（CZ-2)
''''            .MS4FZ = rs("MS4FZ")             ' 測定サンプル4（FZ)
''''            .MS4CZ1 = rs("MS4CZ1")           ' 測定サンプル4（CZ-1)
''''            .MS4CZ2 = rs("MS4CZ2")           ' 測定サンプル4（CZ-2)
''''            .MS5FZ = rs("MS5FZ")             ' 測定サンプル5（FZ)
''''            .MS5CZ1 = rs("MS5CZ1")           ' 測定サンプル5（CZ-1)
''''            .MS5CZ2 = rs("MS5CZ2")           ' 測定サンプル5（CZ-2)
''''            .MSAVEFZ = rs("MSAVEFZ")         ' 測定平均（FZ）
''''            .MSAVECZ1 = rs("MSAVECZ1")       ' 測定平均（CZ-1）
''''            .MSAVECZ2 = rs("MSAVECZ2")       ' 測定平均（CZ-2）
''''            .MSSGFZ = rs("MSSGFZ")           ' 測定σ（FZ）
''''            .MSSGCZ1 = rs("MSSGCZ1")         ' 測定σ（CZ-1）
''''            .MSSGCZ2 = rs("MSSGCZ2")         ' 測定σ（CZ-2）
''''            .MSPSGFZ = rs("MSPSGFZ")         ' 測定AVE+σ（FZ）
''''            .MSPSGCZ1 = rs("MSPSGCZ1")       ' 測定AVE+σ（CZ-1）
''''            .MSPSGCZ2 = rs("MSPSGCZ2")       ' 測定AVE+σ（CZ-2）
''''            .MSNSGFZ = rs("MSNSGFZ")         ' 測定AVE-σ（FZ）
''''            .MSNSGCZ1 = rs("MSNSGCZ1")       ' 測定AVE-σ（CZ-1）
''''            .MSNSGCZ2 = rs("MSNSGCZ2")       ' 測定AVE-σ（CZ-2）
''''            .MINFZ = rs("MINFZ")             ' MIN（FZ）
''''            .MINCZ1 = rs("MINCZ1")           ' MIN（CZ-1）
''''            .MINCZ2 = rs("MINCZ2")           ' MIN（CZ-2）
''''            .MAXFZ = rs("MAXFZ")             ' MAX（FZ）
''''            .MAXCZ1 = rs("MAXCZ1")           ' MAX（CZ-1）
''''            .MAXCZ2 = rs("MAXCZ2")           ' MAX（CZ-2）
''''            .SGCK1FZ = rs("SGCK1FZ")         ' σckサンプル1（FZ)
''''            .SGCK1CZ1 = rs("SGCK1CZ1")       ' σckサンプル1（CZ-1)
''''            .SGCK1CZ2 = rs("SGCK1CZ2")       ' σckサンプル1（CZ-2)
''''            .SGCK2FZ = rs("SGCK2FZ")         ' σckサンプル2（FZ)
''''            .SGCK2CZ1 = rs("SGCK2CZ1")       ' σckサンプル2（CZ-1)
''''            .SGCK2CZ2 = rs("SGCK2CZ2")       ' σckサンプル2（CZ-2)
''''            .SGCK3FZ = rs("SGCK3FZ")         ' σckサンプル3（FZ)
''''            .SGCK3CZ1 = rs("SGCK3CZ1")       ' σckサンプル3（CZ-1)
''''            .SGCK3CZ2 = rs("SGCK3CZ2")       ' σckサンプル3（CZ-2)
''''            .SGCK4FZ = rs("SGCK4FZ")         ' σckサンプル4（FZ)
''''            .SGCK4CZ1 = rs("SGCK4CZ1")       ' σckサンプル4（CZ-1)
''''            .SGCK4CZ2 = rs("SGCK4CZ2")       ' σckサンプル4（CZ-2)
''''            .SGCK5FZ = rs("SGCK5FZ")         ' σckサンプル5（FZ)
''''            .SGCK5CZ1 = rs("SGCK5CZ1")       ' σckサンプル5（CZ-1)
''''            .SGCK5CZ2 = rs("SGCK5CZ2")       ' σckサンプル5（CZ-2)
''''            .SGCKDFZ = rs("SGCKDFZ")         ' σckデータ数（FZ）
''''            .SGCKDCZ1 = rs("SGCKDCZ1")       ' σckデータ数（CZ-1）
''''            .SGCKDCZ2 = rs("SGCKDCZ2")       ' σckデータ数（CZ-2）
''''            .SGCKAFZ = rs("SGCKAFZ")         ' σck平均（FZ）
''''            .SGCKAACZ1 = rs("SGCKAACZ1")     ' σck平均（CZ-1）
''''            .SGCKACZ2 = rs("SGCKACZ2")       ' σck平均（CZ-2）
''''            .SGNFZ = rs("SGNFZ")             ' σckσ（FZ）
''''            .SGNCZ1 = rs("SGNCZ1")           ' σckσ CZ-1）
''''            .SGNCZ2 = rs("SGNCZ2")           ' σckσ（CZ-2）
''''            .FTIRFZ = rs("FTIRFZ")           ' FTIR換算（FZ）
''''            .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR換算（CZ-1）
''''            .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR換算（CZ-2）
''''            .EFFECTTM = rs("EFFECTTM")       ' 有効時間
''''            .YCOEF = rs("YCOEF")             ' ＦＴＩＲ換算式（Ｙ切片）
''''            .XCOEF = rs("XCOEF")             ' ＦＴＩＲ換算式（Ｘ係数）
''''            .RSQUARE = rs("RSQUARE")         ' Ｒ２乗
''''            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
''''            .REGDATE = rs("REGDATE")         ' 登録日付
''''            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
''''            .UPDDATE = rs("UPDDATE")         ' 更新日付
''''            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
''''            .SENDDATE = rs("SENDDATE")       ' 送信日付
''''        End With
''''        rs.MoveNext
''''    Next
''''    rs.Close
''''
''''    DBDRV_GetTBCMB019 = FUNCTION_RETURN_SUCCESS
''''End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    : σ判定基準取得
'
' 返り値  : True  - 正常
' 　　　    False - 失敗
'
' 引き数  : sSigCode  - σ判定基準
' 　　　  : sFtirCode - FTIR換算判定基準
' 　　　  : sR2Code   - R2乗判定基準
'
' 機能説明:
'///////////////////////////////////////////////////
Public Function GetSigChkCode(Optional ByRef sSigCode As String _
                            , Optional ByRef sFtirCode As String _
                            , Optional ByRef sR2Code As String _
                            ) As Boolean
    Dim dbIsMine    As Boolean
    Dim sSql        As String
    Dim objRs       As Object
    
    GetSigChkCode = False
    sSigCode = ""
    sFtirCode = ""
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc055_SQL.bas -- Function GetSigChkCode"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''ＳＱＬ文作成
    sSql = ""
    sSql = sSql & "SELECT NVL(kcode01a9, ' ')"   '0:σ判定基準
    sSql = sSql & "      ,NVL(kcode02a9, ' ')"   '1:FTIR換算判定基準
    sSql = sSql & "      ,NVL(kcode03a9, ' ')"   '2:R2乗判定基準
    sSql = sSql & "  FROM koda9"
    sSql = sSql & " WHERE sysca9 = 'X'"
    sSql = sSql & "   AND shuca9 = '19'"
    sSql = sSql & "   AND codea9 = 'FRS'"
    
    Set objRs = OraDB.CreateDynaset(sSql, ORADYN_DEFAULT)
    
    If objRs.EOF Then
        Call MsgOut(0, "σ判定基準のコードが登録されていません", ERR_DISP)
        Exit Function
    End If

    sSigCode = objRs(0)     ''σ判定基準
    sFtirCode = objRs(1)    ''FTIR換算判定基準
    sR2Code = objRs(2)      ''R2乗判定基準
    
    objRs.Close
    
    ''σ判定基準
    If IsNumeric(sSigCode) = False Then
        Call MsgOut(0, "σ判定基準のコードが正しくありません", ERR_DISP)
        Exit Function
    End If
    ' -10~100でない，または小数点第三位以降の入力がある場合はエラー
    If Not (-10# < CDbl(sSigCode) And CDbl(sSigCode) < 100#) Then
        Call MsgOut(0, "σ判定基準のコードが正しくありません", ERR_DISP)
        Exit Function
    End If
    If InStr(1, sSigCode, ".", vbTextCompare) >= 1 Then
        If Len(sSigCode) - InStr(1, sSigCode, ".", vbTextCompare) >= 3 Then
            Call MsgOut(0, "σ判定基準のコードが正しくありません", ERR_DISP)
            Exit Function
        End If
    End If
    
    ''FTIR換算判定基準
    If IsNumeric(sFtirCode) = False Then
        Call MsgOut(0, "FTIR換算判定基準のコードが正しくありません", ERR_DISP)
        Exit Function
    End If
    ' -10~100でない，または小数点第三位以降の入力がある場合はエラー
    If Not (-10# < CDbl(sFtirCode) And CDbl(sFtirCode) < 100#) Then
        Call MsgOut(0, "FTIR換算判定基準のコードが正しくありません", ERR_DISP)
        Exit Function
    End If
    If InStr(1, sFtirCode, ".", vbTextCompare) >= 1 Then
        If Len(sFtirCode) - InStr(1, sFtirCode, ".", vbTextCompare) >= 3 Then
            Call MsgOut(0, "FTIR換算判定基準のコードが正しくありません", ERR_DISP)
            Exit Function
        End If
    End If
    
    ''R2乗判定基準
    If IsNumeric(sR2Code) = False Then
        Call MsgOut(0, "Ｒ2乗判定基準のコードが正しくありません", ERR_DISP)
        Exit Function
    End If
    
    GetSigChkCode = True        ''処理成功を返す

proc_exit:
    If dbIsMine Then
        OraDBClose
    End If
    
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
    
End Function
