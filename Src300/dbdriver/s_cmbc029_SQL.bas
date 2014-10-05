Attribute VB_Name = "s_cmbc029_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DBアクセス関数
' 定義内容: TBCMB014 (GFA校正情報)
' 参照　　: 060200_全テーブル
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmjc001j_Disp
    GOUKI As String * 3             ' 号機
    INPDATE As Date                 ' 日付
    FTIRFZI As Double               ' FTIR（FZ)
    FTIRCZH As Double               ' FTIR（CZ高）
    FTIRCZC As Double               ' FTIR（CZ中）
    MS1FZ As Double                 ' 測定サンプル1（FZ)
    MS1CZ1 As Double                ' 測定サンプル1（CZ-1)
    MS1CZ2 As Double                ' 測定サンプル1（CZ-2)
    MS2FZ As Double                 ' 測定サンプル2（FZ)
    MS2CZ1 As Double                ' 測定サンプル2（CZ-1)
    MS2CZ2 As Double                ' 測定サンプル2（CZ-2)
    MS3FZ As Double                 ' 測定サンプル3（FZ)
    MS3CZ1 As Double                ' 測定サンプル3（CZ-1)
    MS3CZ2 As Double                ' 測定サンプル3（CZ-2)
    MS4FZ As Double                 ' 測定サンプル4（FZ)
    MS4CZ1 As Double                ' 測定サンプル4（CZ-1)
    MS4CZ2 As Double                ' 測定サンプル4（CZ-2)
    MS5FZ As Double                 ' 測定サンプル5（FZ)
    MS5CZ1 As Double                ' 測定サンプル5（CZ-1)
    MS5CZ2 As Double                ' 測定サンプル5（CZ-2)
    MSAVEFZ As Double               ' 測定平均（FZ）
    MSAVECZ1 As Double              ' 測定平均（CZ-1）
    MSAVECZ2 As Double              ' 測定平均（CZ-2）
    MSSGFZ As Double                ' 測定σ（FZ）
    MSSGCZ1 As Double               ' 測定σ（CZ-1）
    MSSGCZ2 As Double               ' 測定σ（CZ-2）
    MSPSGFZ As Double               ' 測定AVE+σ（FZ）
    MSPSGCZ1 As Double              ' 測定AVE+σ（CZ-1）
    MSPSGCZ2 As Double              ' 測定AVE+σ（CZ-2）
    MSNSGFZ As Double               ' 測定AVE-σ（FZ）
    MSNSGCZ1 As Double              ' 測定AVE-σ（CZ-1）
    MSNSGCZ2 As Double              ' 測定AVE-σ（CZ-2）
    MINFZ As Double                 ' MIN（FZ）
    MINCZ1 As Double                ' MIN（CZ-1）
    MINCZ2 As Double                ' MIN（CZ-2）
    MAXFZ As Double                 ' MAX（FZ）
    MAXCZ1 As Double                ' MAX（CZ-1）
    MAXCZ2 As Double                ' MAX（CZ-2）
    SGCK1FZ As Double               ' σckサンプル1（FZ)
    SGCK1CZ1 As Double              ' σckサンプル1（CZ-1)
    SGCK1CZ2 As Double              ' σckサンプル1（CZ-2)
    SGCK2FZ As Double               ' σckサンプル2（FZ)
    SGCK2CZ1 As Double              ' σckサンプル2（CZ-1)
    SGCK2CZ2 As Double              ' σckサンプル2（CZ-2)
    SGCK3FZ As Double               ' σckサンプル3（FZ)
    SGCK3CZ1 As Double              ' σckサンプル3（CZ-1)
    SGCK3CZ2 As Double              ' σckサンプル3（CZ-2)
    SGCK4FZ As Double               ' σckサンプル4（FZ)
    SGCK4CZ1 As Double              ' σckサンプル4（CZ-1)
    SGCK4CZ2 As Double              ' σckサンプル4（CZ-2)
    SGCK5FZ As Double               ' σckサンプル5（FZ)
    SGCK5CZ1 As Double              ' σckサンプル5（CZ-1)
    SGCK5CZ2 As Double              ' σckサンプル5（CZ-2)
    SGCKDFZ As Double               ' σckデータ数（FZ）
    SGCKDCZ1 As Double              ' σckデータ数（CZ-1）
    SGCKDCZ2 As Double              ' σckデータ数（CZ-2）
    SGCKAFZ As Double               ' σck平均（FZ）
    SGCKAACZ1 As Double             ' σck平均（CZ-1）
    SGCKACZ2 As Double              ' σck平均（CZ-2）
    SGNFZ As Double                 ' σckσ（FZ）
    SGNCZ1 As Double                ' σckσ CZ-1）
    SGNCZ2 As Double                ' σckσ（CZ-2）
    FTIRFZ As Double                ' FTIR換算（FZ）
    FTIRCZ1 As Double               ' FTIR換算（CZ-1）
    FTIRCZ2 As Double               ' FTIR換算（CZ-2）
    EFFECTTM As Integer             ' 有効時間
    YCOEF As Double                 ' ＦＴＩＲ換算式（Ｙ切片）
    XCOEF As Double                 ' ＦＴＩＲ換算式（Ｘ係数）
    RSQUARE As Double               ' Ｒ２乗
  '  TSTAFFID As String * 8          ' 登録社員ID
  '  REGDATE As Date                 ' 登録日付
  '  KSTAFFID As String * 8          ' 更新社員ID
  '  UPDDATE As Date                 ' 更新日付
  '  SENDFLAG As String * 1          ' 送信フラグ
  '  SENDDATE As Date                ' 送信日付

'2006/05/22追加
    SGCKST      As Double           ' σ判定基準
    SGCKFZ      As String * 1       ' σ判定(FZ)
    SGCKCZ1     As String * 1       ' σ判定(CZ-1)
    SGCKCZ2     As String * 1       ' σ判定(CZ-2)
    FTIRCKST    As Double           ' FTIR換算判定基準
    FTIRCKFZ    As String * 1       ' FTIR換算判定(FZ)
    FTIRCKCZ1   As String * 1       ' FTIR換算判定(CZ-1)
    FTIRCKCZ2   As String * 1       ' FTIR換算判定(CZ-2)

'2010/03/26追加 SETsw kubota
    MS6FZ       As Double           ' 測定サンプル6（FZ)
    MS6CZ1      As Double           ' 測定サンプル6（CZ-1)
    MS6CZ2      As Double           ' 測定サンプル6（CZ-2)
    SGCK6FZ     As Double           ' σckサンプル6（FZ)
    SGCK6CZ1    As Double           ' σckサンプル6（CZ-1)
    SGCK6CZ2    As Double           ' σckサンプル6（CZ-2)
    CVFZ        As Double           ' CV(%)（FZ）
    CVCZ1       As Double           ' CV(%)（CZ-1）
    CVCZ2       As Double           ' CV(%)（CZ-2）

End Type

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMB014」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :record        ,O  ,typ_cmjc001j_Disp ,抽出レコード
'          :GOUK          ,I  ,String       ,「号機」(SQLの抽出条件)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :「号機」=引数で、かつ「日付」が最新のデータを抽出する
'履歴      :2001/06/20作成　長野
Public Function DBDRV_Getcmjc001j_Disp(record As typ_cmjc001j_Disp, GOUK$) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim sqlWhere As String  'SQLのWHERE部分
Dim sqlGroup As String  'SQLのGROUP部分
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_FAILURE
    
    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001j_SQL.bas -- Function DBDRV_Getcmjc001j_Disp"

    sqlBase = "Select GOUKI, MAX(INPDATE) ""INPDATE"", FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE "
    sqlBase = sqlBase & "From TBCMB014"
    ''抽出条件(ｻﾝﾌﾟﾙNO)の取り出し
    sqlWhere = "WHERE(GOUKI=" & GOUK & ") "
    sqlGroup = "GROUP BY GOUKI"
    sql = sqlBase & sqlWhere & sqlGroup
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    With record
        .GOUKI = rs("GOUKI")             ' 号機
        .INPDATE = rs("INPDATE")         ' 日付
        .FTIRFZI = rs("FTIRFZI")         ' FTIR（FZ)
        .FTIRCZH = rs("FTIRCZH")         ' FTIR（CZ高）
        .FTIRCZC = rs("FTIRCZC")         ' FTIR（CZ中）
        .MS1FZ = rs("MS1FZ")             ' 測定サンプル1（FZ)
        .MS1CZ1 = rs("MS1CZ1")           ' 測定サンプル1（CZ-1)
        .MS1CZ2 = rs("MS1CZ2")           ' 測定サンプル1（CZ-2)
        .MS2FZ = rs("MS2FZ")             ' 測定サンプル2（FZ)
        .MS2CZ1 = rs("MS2CZ1")           ' 測定サンプル2（CZ-1)
        .MS2CZ2 = rs("MS2CZ2")           ' 測定サンプル2（CZ-2)
        .MS3FZ = rs("MS3FZ")             ' 測定サンプル3（FZ)
        .MS3CZ1 = rs("MS3CZ1")           ' 測定サンプル3（CZ-1)
        .MS3CZ2 = rs("MS3CZ2")           ' 測定サンプル3（CZ-2)
        .MS4FZ = rs("MS4FZ")             ' 測定サンプル4（FZ)
        .MS4CZ1 = rs("MS4CZ1")           ' 測定サンプル4（CZ-1)
        .MS4CZ2 = rs("MS4CZ2")           ' 測定サンプル4（CZ-2)
        .MS5FZ = rs("MS5FZ")             ' 測定サンプル5（FZ)
        .MS5CZ1 = rs("MS5CZ1")           ' 測定サンプル5（CZ-1)
        .MS5CZ2 = rs("MS5CZ2")           ' 測定サンプル5（CZ-2)
        .MSAVEFZ = rs("MSAVEFZ")         ' 測定平均（FZ）
        .MSAVECZ1 = rs("MSAVECZ1")       ' 測定平均（CZ-1）
        .MSAVECZ2 = rs("MSAVECZ2")       ' 測定平均（CZ-2）
        .MSSGFZ = rs("MSSGFZ")           ' 測定σ（FZ）
        .MSSGCZ1 = rs("MSSGCZ1")         ' 測定σ（CZ-1）
        .MSSGCZ2 = rs("MSSGCZ2")         ' 測定σ（CZ-2）
        .MSPSGFZ = rs("MSPSGFZ")         ' 測定AVE+σ（FZ）
        .MSPSGCZ1 = rs("MSPSGCZ1")       ' 測定AVE+σ（CZ-1）
        .MSPSGCZ2 = rs("MSPSGCZ2")       ' 測定AVE+σ（CZ-2）
        .MSNSGFZ = rs("MSNSGFZ")         ' 測定AVE-σ（FZ）
        .MSNSGCZ1 = rs("MSNSGCZ1")       ' 測定AVE-σ（CZ-1）
        .MSNSGCZ2 = rs("MSNSGCZ2")       ' 測定AVE-σ（CZ-2）
        .MINFZ = rs("MINFZ")             ' MIN（FZ）
        .MINCZ1 = rs("MINCZ1")           ' MIN（CZ-1）
        .MINCZ2 = rs("MINCZ2")           ' MIN（CZ-2）
        .MAXFZ = rs("MAXFZ")             ' MAX（FZ）
        .MAXCZ1 = rs("MAXCZ1")           ' MAX（CZ-1）
        .MAXCZ2 = rs("MAXCZ2")           ' MAX（CZ-2）
        .SGCK1FZ = rs("SGCK1FZ")         ' σckサンプル1（FZ)
        .SGCK1CZ1 = rs("SGCK1CZ1")       ' σckサンプル1（CZ-1)
        .SGCK1CZ2 = rs("SGCK1CZ2")       ' σckサンプル1（CZ-2)
        .SGCK2FZ = rs("SGCK2FZ")         ' σckサンプル2（FZ)
        .SGCK2CZ1 = rs("SGCK2CZ1")       ' σckサンプル2（CZ-1)
        .SGCK2CZ2 = rs("SGCK2CZ2")       ' σckサンプル2（CZ-2)
        .SGCK3FZ = rs("SGCK3FZ")         ' σckサンプル3（FZ)
        .SGCK3CZ1 = rs("SGCK3CZ1")       ' σckサンプル3（CZ-1)
        .SGCK3CZ2 = rs("SGCK3CZ2")       ' σckサンプル3（CZ-2)
        .SGCK4FZ = rs("SGCK4FZ")         ' σckサンプル4（FZ)
        .SGCK4CZ1 = rs("SGCK4CZ1")       ' σckサンプル4（CZ-1)
        .SGCK4CZ2 = rs("SGCK4CZ2")       ' σckサンプル4（CZ-2)
        .SGCK5FZ = rs("SGCK5FZ")         ' σckサンプル5（FZ)
        .SGCK5CZ1 = rs("SGCK5CZ1")       ' σckサンプル5（CZ-1)
        .SGCK5CZ2 = rs("SGCK5CZ2")       ' σckサンプル5（CZ-2)
        .SGCKDFZ = rs("SGCKDFZ")         ' σckデータ数（FZ）
        .SGCKDCZ1 = rs("SGCKDCZ1")       ' σckデータ数（CZ-1）
        .SGCKDCZ2 = rs("SGCKDCZ2")       ' σckデータ数（CZ-2）
        .SGCKAFZ = rs("SGCKAFZ")         ' σck平均（FZ）
        .SGCKAACZ1 = rs("SGCKAACZ1")     ' σck平均（CZ-1）
        .SGCKACZ2 = rs("SGCKACZ2")       ' σck平均（CZ-2）
        .SGNFZ = rs("SGNFZ")             ' σckσ（FZ）
        .SGNCZ1 = rs("SGNCZ1")           ' σckσ CZ-1）
        .SGNCZ2 = rs("SGNCZ2")           ' σckσ（CZ-2）
        .FTIRFZ = rs("FTIRFZ")           ' FTIR換算（FZ）
        .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR換算（CZ-1）
        .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR換算（CZ-2）
        .EFFECTTM = rs("EFFECTTM")       ' 有効時間
        .YCOEF = rs("YCOEF")             ' ＦＴＩＲ換算式（Ｙ切片）
        .XCOEF = rs("XCOEF")             ' ＦＴＩＲ換算式（Ｘ係数）
        .RSQUARE = rs("RSQUARE")         ' Ｒ２乗
    End With
    rs.Close

    DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_SUCCESS

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


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :引数で渡されたレコードをTBCMB014に追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :record        ,I  ,typ_cmjc001j_Disp ,抽出レコード
'          :TSTAFFID      ,I  ,String       ,登録社員ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/22(Fri)作成　長野

Public Function DBDRV_Getcmjc001j_Exec(record As typ_cmjc001j_Disp, TSTAFFID$) As FUNCTION_RETURN

Dim sql As String           'SQL全体
Dim SetDate  As Variant     '入力日付

    DBDRV_Getcmjc001j_Exec = FUNCTION_RETURN_FAILURE
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001j_SQL.bas -- Function DBDRV_Getcmjc001j_Exec"

    SetDate = Format$(record.INPDATE, "yyyy-mm-dd hh:mm:ss")
  
  ''SQLを組み立てる
    sql = "Insert into TBCMB014 (GOUKI, INPDATE, FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, " & _
          "MS3FZ, MS3CZ1, MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, " & _
          "MSSGFZ, MSSGCZ1, MSSGCZ2, MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, " & _
          "MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ, SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, " & _
          "SGCK4FZ, SGCK4CZ1, SGCK4CZ2, SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, " & _
          "SGNFZ, SGNCZ1, SGNCZ2, FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE"
    sql = sql & ",SGCKST,SGCKFZ,SGCKCZ1,SGCKCZ2,FTIRCKST,FTIRCKFZ,FTIRCKCZ1,FTIRCKCZ2"  '2006/05/23追加 kubota
    sql = sql & ",MS6FZ,MS6CZ1,MS6CZ2,SGCK6FZ,SGCK6CZ1,SGCK6CZ2,CVFZ,CVCZ1,CVCZ2"       '2010/03/26追加 kubota
    sql = sql & ")"
    
    sql = sql & "Values('" & record.GOUKI & "', " & "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss'), " & record.FTIRFZI & ", " & _
          record.FTIRCZH & ", " & record.FTIRCZC & ", " & record.MS1FZ & ", " & record.MS1CZ1 & ", " & record.MS1CZ2 & ", " & _
          record.MS2FZ & ", " & record.MS2CZ1 & ", " & record.MS2CZ2 & ", " & record.MS3FZ & ", " & record.MS3CZ1 & ", " & _
          record.MS3CZ2 & ", " & record.MS4FZ & ", " & record.MS4CZ1 & ", " & record.MS4CZ2 & ", " & record.MS5FZ & ", " & _
          record.MS5CZ1 & ", " & record.MS5CZ2 & ", " & record.MSAVEFZ & ", " & record.MSAVECZ1 & ", " & record.MSAVECZ2 & ", " & _
          record.MSSGFZ & ", " & record.MSSGCZ1 & ", " & record.MSSGCZ2 & ", " & record.MSPSGFZ & ", " & record.MSPSGCZ1 & ", " & _
          record.MSPSGCZ2 & ", " & record.MSNSGFZ & ", " & record.MSNSGCZ1 & ", " & record.MSNSGCZ2 & ", " & record.MINFZ & ", " & _
          record.MINCZ1 & ", " & record.MINCZ2 & ", " & record.MAXFZ & ", " & record.MAXCZ1 & ", " & record.MAXCZ2 & ", " & record.SGCK1FZ & ", " & _
          record.SGCK1CZ1 & ", " & record.SGCK1CZ2 & ", " & record.SGCK2FZ & ", " & record.SGCK2CZ1 & ", " & record.SGCK2CZ2 & ", " & _
          record.SGCK3FZ & ", " & record.SGCK3CZ1 & ", " & record.SGCK3CZ2 & ", " & record.SGCK4FZ & ", " & record.SGCK4CZ1 & ", " & _
          record.SGCK4CZ2 & ", " & record.SGCK5FZ & ", " & record.SGCK5CZ1 & ", " & record.SGCK5CZ2 & ", " & record.SGCKDFZ & ", " & _
          record.SGCKDCZ1 & ", " & record.SGCKDCZ2 & ", " & record.SGCKAFZ & ", " & record.SGCKAACZ1 & ", " & record.SGCKACZ2 & ", " & _
          record.SGNFZ & ", " & record.SGNCZ1 & ", " & record.SGNCZ2 & ", " & record.FTIRFZ & ", " & record.FTIRCZ1 & ", " & _
          record.FTIRCZ2 & ", " & record.EFFECTTM & ", " & record.YCOEF & ", " & record.XCOEF & ", " & record.RSQUARE & ", '" & _
          TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE"
    sql = sql & "," & record.SGCKST & ",'" & record.SGCKFZ & "','" & record.SGCKCZ1 & "','" & record.SGCKCZ2 & "'"          '2006/05/23追加 kubota
    sql = sql & "," & record.FTIRCKST & ",'" & record.FTIRCKFZ & "','" & record.FTIRCKCZ1 & "','" & record.FTIRCKCZ2 & "'"  '2006/05/23追加 kubota
    sql = sql & "," & record.MS6FZ & ",'" & record.MS6CZ1 & "','" & record.MS6CZ2 & "'"                                     '2010/03/26追加 kubota
    sql = sql & "," & record.SGCK6FZ & ",'" & record.SGCK6CZ1 & "','" & record.SGCK6CZ2 & "'"                               '2010/03/26追加 kubota
    sql = sql & "," & record.CVFZ & ",'" & record.CVCZ1 & "','" & record.CVCZ2 & "'"                                        '2010/03/26追加 kubota
    sql = sql & ")"
  
  ''SQLの実行
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
'          :tblLeft       ,IO   ,typ_TBCMB014      ,テーブルデータ１
'          :tblRight      ,IO   ,typ_cmjc001j_Disp ,テーブルデータ２
'          :bFlg          ,I   ,Boolean           ,TRUE:引数１データ→引数２データへの変換  FALSE:引数１データ←引数２データへの変換
'説明      :
Public Sub ConvDate_F_cmjc001j_a(tblLeft As typ_TBCMB014, tblRight As typ_cmjc001j_Disp, bFlg As Boolean)
    If bFlg = True Then
        With tblRight
            .GOUKI = tblLeft.GOUKI
            .INPDATE = tblLeft.INPDATE
            .FTIRFZI = tblLeft.FTIRFZI
            .FTIRCZH = tblLeft.FTIRCZH
            .FTIRCZC = tblLeft.FTIRCZC
            .MS1FZ = tblLeft.MS1FZ
            .MS1CZ1 = tblLeft.MS1CZ1
            .MS1CZ2 = tblLeft.MS1CZ2
            .MS2FZ = tblLeft.MS2FZ
            .MS2CZ1 = tblLeft.MS2CZ1
            .MS2CZ2 = tblLeft.MS2CZ2
            .MS3FZ = tblLeft.MS3FZ
            .MS3CZ1 = tblLeft.MS3CZ1
            .MS3CZ2 = tblLeft.MS3CZ2
            .MS4FZ = tblLeft.MS4FZ
            .MS4CZ1 = tblLeft.MS4CZ1
            .MS4CZ2 = tblLeft.MS4CZ2
            .MS5FZ = tblLeft.MS5FZ
            .MS5CZ1 = tblLeft.MS5CZ1
            .MS5CZ2 = tblLeft.MS5CZ2
            .MSAVEFZ = tblLeft.MSAVEFZ
            .MSAVECZ1 = tblLeft.MSAVECZ1
            .MSAVECZ2 = tblLeft.MSAVECZ2
            .MSSGFZ = tblLeft.MSSGFZ
            .MSSGCZ1 = tblLeft.MSSGCZ1
            .MSSGCZ2 = tblLeft.MSSGCZ2
            .MSPSGFZ = tblLeft.MSPSGFZ
            .MSPSGCZ1 = tblLeft.MSPSGCZ1
            .MSPSGCZ2 = tblLeft.MSPSGCZ2
            .MSNSGFZ = tblLeft.MSNSGFZ
            .MSNSGCZ1 = tblLeft.MSNSGCZ1
            .MSNSGCZ2 = tblLeft.MSNSGCZ2
            .MINFZ = tblLeft.MINFZ
            .MINCZ1 = tblLeft.MINCZ1
            .MINCZ2 = tblLeft.MINCZ2
            .MAXFZ = tblLeft.MAXFZ
            .MAXCZ1 = tblLeft.MAXCZ1
            .MAXCZ2 = tblLeft.MAXCZ2
            .SGCK1FZ = tblLeft.SGCK1FZ
            .SGCK1CZ1 = tblLeft.SGCK1CZ1
            .SGCK1CZ2 = tblLeft.SGCK1CZ2
            .SGCK2FZ = tblLeft.SGCK2FZ
            .SGCK2CZ1 = tblLeft.SGCK2CZ1
            .SGCK2CZ2 = tblLeft.SGCK2CZ2
            .SGCK3FZ = tblLeft.SGCK3FZ
            .SGCK3CZ1 = tblLeft.SGCK3CZ1
            .SGCK3CZ2 = tblLeft.SGCK3CZ2
            .SGCK4FZ = tblLeft.SGCK4FZ
            .SGCK4CZ1 = tblLeft.SGCK4CZ1
            .SGCK4CZ2 = tblLeft.SGCK4CZ2
            .SGCK5FZ = tblLeft.SGCK5FZ
            .SGCK5CZ1 = tblLeft.SGCK5CZ1
            .SGCK5CZ2 = tblLeft.SGCK5CZ2
            .SGCKDFZ = tblLeft.SGCKDFZ
            .SGCKDCZ1 = tblLeft.SGCKDCZ1
            .SGCKDCZ2 = tblLeft.SGCKDCZ2
            .SGCKAFZ = tblLeft.SGCKAFZ
            .SGCKAACZ1 = tblLeft.SGCKAACZ1
            .SGCKACZ2 = tblLeft.SGCKACZ2
            .SGNFZ = tblLeft.SGNFZ
            .SGNCZ1 = tblLeft.SGNCZ1
            .SGNCZ2 = tblLeft.SGNCZ2
            .FTIRFZ = tblLeft.FTIRFZ
            .FTIRCZ1 = tblLeft.FTIRCZ1
            .FTIRCZ2 = tblLeft.FTIRCZ2
            .EFFECTTM = tblLeft.EFFECTTM
            .YCOEF = tblLeft.YCOEF
            .XCOEF = tblLeft.XCOEF
            .RSQUARE = tblLeft.RSQUARE
            
'2006/05/22追加 kubota
            .SGCKST = tblLeft.SGCKST
            .SGCKFZ = tblLeft.SGCKFZ
            .SGCKCZ1 = tblLeft.SGCKCZ1
            .SGCKCZ2 = tblLeft.SGCKCZ2
            .FTIRCKST = tblLeft.FTIRCKST
            .FTIRCKFZ = tblLeft.FTIRCKFZ
            .FTIRCKCZ1 = tblLeft.FTIRCKCZ1
            .FTIRCKCZ2 = tblLeft.FTIRCKCZ2
        
'2010/03/26追加 SETsw kubota
            .MS6FZ = tblLeft.MS6FZ
            .MS6CZ1 = tblLeft.MS6CZ1
            .MS6CZ2 = tblLeft.MS6CZ2
            .SGCK6FZ = tblLeft.SGCK6FZ
            .SGCK6CZ1 = tblLeft.SGCK6CZ1
            .SGCK6CZ2 = tblLeft.SGCK6CZ2
            .CVFZ = tblLeft.CVFZ
            .CVCZ1 = tblLeft.CVCZ1
            .CVCZ2 = tblLeft.CVCZ2
        
        End With
    Else
        With tblLeft
            .GOUKI = tblRight.GOUKI
            .INPDATE = tblRight.INPDATE
            .FTIRFZI = tblRight.FTIRFZI
            .FTIRCZH = tblRight.FTIRCZH
            .FTIRCZC = tblRight.FTIRCZC
            .MS1FZ = tblRight.MS1FZ
            .MS1CZ1 = tblRight.MS1CZ1
            .MS1CZ2 = tblRight.MS1CZ2
            .MS2FZ = tblRight.MS2FZ
            .MS2CZ1 = tblRight.MS2CZ1
            .MS2CZ2 = tblRight.MS2CZ2
            .MS3FZ = tblRight.MS3FZ
            .MS3CZ1 = tblRight.MS3CZ1
            .MS3CZ2 = tblRight.MS3CZ2
            .MS4FZ = tblRight.MS4FZ
            .MS4CZ1 = tblRight.MS4CZ1
            .MS4CZ2 = tblRight.MS4CZ2
            .MS5FZ = tblRight.MS5FZ
            .MS5CZ1 = tblRight.MS5CZ1
            .MS5CZ2 = tblRight.MS5CZ2
            .MSAVEFZ = tblRight.MSAVEFZ
            .MSAVECZ1 = tblRight.MSAVECZ1
            .MSAVECZ2 = tblRight.MSAVECZ2
            .MSSGFZ = tblRight.MSSGFZ
            .MSSGCZ1 = tblRight.MSSGCZ1
            .MSSGCZ2 = tblRight.MSSGCZ2
            .MSPSGFZ = tblRight.MSPSGFZ
            .MSPSGCZ1 = tblRight.MSPSGCZ1
            .MSPSGCZ2 = tblRight.MSPSGCZ2
            .MSNSGFZ = tblRight.MSNSGFZ
            .MSNSGCZ1 = tblRight.MSNSGCZ1
            .MSNSGCZ2 = tblRight.MSNSGCZ2
            .MINFZ = tblRight.MINFZ
            .MINCZ1 = tblRight.MINCZ1
            .MINCZ2 = tblRight.MINCZ2
            .MAXFZ = tblRight.MAXFZ
            .MAXCZ1 = tblRight.MAXCZ1
            .MAXCZ2 = tblRight.MAXCZ2
            .SGCK1FZ = tblRight.SGCK1FZ
            .SGCK1CZ1 = tblRight.SGCK1CZ1
            .SGCK1CZ2 = tblRight.SGCK1CZ2
            .SGCK2FZ = tblRight.SGCK2FZ
            .SGCK2CZ1 = tblRight.SGCK2CZ1
            .SGCK2CZ2 = tblRight.SGCK2CZ2
            .SGCK3FZ = tblRight.SGCK3FZ
            .SGCK3CZ1 = tblRight.SGCK3CZ1
            .SGCK3CZ2 = tblRight.SGCK3CZ2
            .SGCK4FZ = tblRight.SGCK4FZ
            .SGCK4CZ1 = tblRight.SGCK4CZ1
            .SGCK4CZ2 = tblRight.SGCK4CZ2
            .SGCK5FZ = tblRight.SGCK5FZ
            .SGCK5CZ1 = tblRight.SGCK5CZ1
            .SGCK5CZ2 = tblRight.SGCK5CZ2
            .SGCKDFZ = tblRight.SGCKDFZ
            .SGCKDCZ1 = tblRight.SGCKDCZ1
            .SGCKDCZ2 = tblRight.SGCKDCZ2
            .SGCKAFZ = tblRight.SGCKAFZ
            .SGCKAACZ1 = tblRight.SGCKAACZ1
            .SGCKACZ2 = tblRight.SGCKACZ2
            .SGNFZ = tblRight.SGNFZ
            .SGNCZ1 = tblRight.SGNCZ1
            .SGNCZ2 = tblRight.SGNCZ2
            .FTIRFZ = tblRight.FTIRFZ
            .FTIRCZ1 = tblRight.FTIRCZ1
            .FTIRCZ2 = tblRight.FTIRCZ2
            .EFFECTTM = tblRight.EFFECTTM
            .YCOEF = tblRight.YCOEF
            .XCOEF = tblRight.XCOEF
            .RSQUARE = tblRight.RSQUARE
        
'2006/05/22追加 kubota
            .SGCKST = tblRight.SGCKST
            .SGCKFZ = tblRight.SGCKFZ
            .SGCKCZ1 = tblRight.SGCKCZ1
            .SGCKCZ2 = tblRight.SGCKCZ2
            .FTIRCKST = tblRight.FTIRCKST
            .FTIRCKFZ = tblRight.FTIRCKFZ
            .FTIRCKCZ1 = tblRight.FTIRCKCZ1
            .FTIRCKCZ2 = tblRight.FTIRCKCZ2

'2010/03/26追加 SETsw kubota
            .MS6FZ = tblRight.MS6FZ
            .MS6CZ1 = tblRight.MS6CZ1
            .MS6CZ2 = tblRight.MS6CZ2
            .SGCK6FZ = tblRight.SGCK6FZ
            .SGCK6CZ1 = tblRight.SGCK6CZ1
            .SGCK6CZ2 = tblRight.SGCK6CZ2
            .CVFZ = tblRight.CVFZ
            .CVCZ1 = tblRight.CVCZ1
            .CVCZ2 = tblRight.CVCZ2
        
        End With
    End If

End Sub

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMB014」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB014 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCMB014(records() As typ_TBCMB014, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select GOUKI, INPDATE, FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG," & _
              " SENDDATE "
    sqlBase = sqlBase & "From TBCMB014"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB014 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .GOUKI = rs("GOUKI")             ' 号機
            .INPDATE = rs("INPDATE")         ' 日付
            .FTIRFZI = rs("FTIRFZI")         ' FTIR（FZ)
            .FTIRCZH = rs("FTIRCZH")         ' FTIR（CZ高）
            .FTIRCZC = rs("FTIRCZC")         ' FTIR（CZ中）
            .MS1FZ = rs("MS1FZ")             ' 測定サンプル1（FZ)
            .MS1CZ1 = rs("MS1CZ1")           ' 測定サンプル1（CZ-1)
            .MS1CZ2 = rs("MS1CZ2")           ' 測定サンプル1（CZ-2)
            .MS2FZ = rs("MS2FZ")             ' 測定サンプル2（FZ)
            .MS2CZ1 = rs("MS2CZ1")           ' 測定サンプル2（CZ-1)
            .MS2CZ2 = rs("MS2CZ2")           ' 測定サンプル2（CZ-2)
            .MS3FZ = rs("MS3FZ")             ' 測定サンプル3（FZ)
            .MS3CZ1 = rs("MS3CZ1")           ' 測定サンプル3（CZ-1)
            .MS3CZ2 = rs("MS3CZ2")           ' 測定サンプル3（CZ-2)
            .MS4FZ = rs("MS4FZ")             ' 測定サンプル4（FZ)
            .MS4CZ1 = rs("MS4CZ1")           ' 測定サンプル4（CZ-1)
            .MS4CZ2 = rs("MS4CZ2")           ' 測定サンプル4（CZ-2)
            .MS5FZ = rs("MS5FZ")             ' 測定サンプル5（FZ)
            .MS5CZ1 = rs("MS5CZ1")           ' 測定サンプル5（CZ-1)
            .MS5CZ2 = rs("MS5CZ2")           ' 測定サンプル5（CZ-2)
            .MSAVEFZ = rs("MSAVEFZ")         ' 測定平均（FZ）
            .MSAVECZ1 = rs("MSAVECZ1")       ' 測定平均（CZ-1）
            .MSAVECZ2 = rs("MSAVECZ2")       ' 測定平均（CZ-2）
            .MSSGFZ = rs("MSSGFZ")           ' 測定σ（FZ）
            .MSSGCZ1 = rs("MSSGCZ1")         ' 測定σ（CZ-1）
            .MSSGCZ2 = rs("MSSGCZ2")         ' 測定σ（CZ-2）
            .MSPSGFZ = rs("MSPSGFZ")         ' 測定AVE+σ（FZ）
            .MSPSGCZ1 = rs("MSPSGCZ1")       ' 測定AVE+σ（CZ-1）
            .MSPSGCZ2 = rs("MSPSGCZ2")       ' 測定AVE+σ（CZ-2）
            .MSNSGFZ = rs("MSNSGFZ")         ' 測定AVE-σ（FZ）
            .MSNSGCZ1 = rs("MSNSGCZ1")       ' 測定AVE-σ（CZ-1）
            .MSNSGCZ2 = rs("MSNSGCZ2")       ' 測定AVE-σ（CZ-2）
            .MINFZ = rs("MINFZ")             ' MIN（FZ）
            .MINCZ1 = rs("MINCZ1")           ' MIN（CZ-1）
            .MINCZ2 = rs("MINCZ2")           ' MIN（CZ-2）
            .MAXFZ = rs("MAXFZ")             ' MAX（FZ）
            .MAXCZ1 = rs("MAXCZ1")           ' MAX（CZ-1）
            .MAXCZ2 = rs("MAXCZ2")           ' MAX（CZ-2）
            .SGCK1FZ = rs("SGCK1FZ")         ' σckサンプル1（FZ)
            .SGCK1CZ1 = rs("SGCK1CZ1")       ' σckサンプル1（CZ-1)
            .SGCK1CZ2 = rs("SGCK1CZ2")       ' σckサンプル1（CZ-2)
            .SGCK2FZ = rs("SGCK2FZ")         ' σckサンプル2（FZ)
            .SGCK2CZ1 = rs("SGCK2CZ1")       ' σckサンプル2（CZ-1)
            .SGCK2CZ2 = rs("SGCK2CZ2")       ' σckサンプル2（CZ-2)
            .SGCK3FZ = rs("SGCK3FZ")         ' σckサンプル3（FZ)
            .SGCK3CZ1 = rs("SGCK3CZ1")       ' σckサンプル3（CZ-1)
            .SGCK3CZ2 = rs("SGCK3CZ2")       ' σckサンプル3（CZ-2)
            .SGCK4FZ = rs("SGCK4FZ")         ' σckサンプル4（FZ)
            .SGCK4CZ1 = rs("SGCK4CZ1")       ' σckサンプル4（CZ-1)
            .SGCK4CZ2 = rs("SGCK4CZ2")       ' σckサンプル4（CZ-2)
            .SGCK5FZ = rs("SGCK5FZ")         ' σckサンプル5（FZ)
            .SGCK5CZ1 = rs("SGCK5CZ1")       ' σckサンプル5（CZ-1)
            .SGCK5CZ2 = rs("SGCK5CZ2")       ' σckサンプル5（CZ-2)
            .SGCKDFZ = rs("SGCKDFZ")         ' σckデータ数（FZ）
            .SGCKDCZ1 = rs("SGCKDCZ1")       ' σckデータ数（CZ-1）
            .SGCKDCZ2 = rs("SGCKDCZ2")       ' σckデータ数（CZ-2）
            .SGCKAFZ = rs("SGCKAFZ")         ' σck平均（FZ）
            .SGCKAACZ1 = rs("SGCKAACZ1")     ' σck平均（CZ-1）
            .SGCKACZ2 = rs("SGCKACZ2")       ' σck平均（CZ-2）
            .SGNFZ = rs("SGNFZ")             ' σckσ（FZ）
            .SGNCZ1 = rs("SGNCZ1")           ' σckσ CZ-1）
            .SGNCZ2 = rs("SGNCZ2")           ' σckσ（CZ-2）
            .FTIRFZ = rs("FTIRFZ")           ' FTIR換算（FZ）
            .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR換算（CZ-1）
            .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR換算（CZ-2）
            .EFFECTTM = rs("EFFECTTM")       ' 有効時間
            .YCOEF = rs("YCOEF")             ' ＦＴＩＲ換算式（Ｙ切片）
            .XCOEF = rs("XCOEF")             ' ＦＴＩＲ換算式（Ｘ係数）
            .RSQUARE = rs("RSQUARE")         ' Ｒ２乗
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

    DBDRV_GetTBCMB014 = FUNCTION_RETURN_SUCCESS
End Function

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
'2006/05/22追加
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
    gErr.Push "s_cmzc029_SQL.bas -- Function GetSigChkCode"

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
    sSql = sSql & "   AND codea9 = 'GFA'"
    
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

