Attribute VB_Name = "s_cmbc027_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DBアクセス関数
' 定義内容: TBCMJ007 (ライフタイム)
' 参照　　: 060211_結晶検査
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmjc001h_Disp
   ' CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
   ' TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    factory As String * 1           ' 工場
    opecond As String * 1           ' 操業条件
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    GOUKI As String * 3             ' 号機
    MEAS1 As Integer                ' 測定値１
    MEAS2 As Integer                ' 測定値２
    MEAS3 As Integer                ' 測定値３
    MEAS4 As Integer                ' 測定値４
    MEAS5 As Integer                ' 測定値５
    MEASPEAK As Integer             ' 測定値 ピーク値
    CALCMEAS As Integer             ' 計算結果
   ' TSTAFFID As String * 8          ' 登録社員ID
   ' REGDATE As Date                 ' 登録日付
   ' KSTAFFID As String * 8          ' 更新社員ID
   ' UPDDATE As Date                 ' 更新日付
   ' SENDFLAG As String * 1          ' 送信フラグ
   ' SENDDATE As Date                ' 送信日付
' Add Start 2005/11/14 M.Makino
    MEAS6 As Integer                ' 測定値６
    MEAS7 As Integer                ' 測定値７
    MEAS8 As Integer                ' 測定値８
    MEAS9 As Integer                ' 測定値９
    MEAS10 As Integer               ' 測定値１０
    MEASFILE As String              ' 測定データファイル名
    RESVAL As String                ' 実測抵抗
    INCVAL As String                ' 傾き
    CUTVAL As String                ' 切片
    SETVAL As String                ' 設定値
    CONVAL As String                ' 10Ω換算値
    MEAS1DAT1 As String            ' 測定値１　生データ１
    MEAS1DAT2 As String            ' 測定値１　生データ２
    MEAS1DAT3 As String            ' 測定値１　生データ３
    MEAS1DAT4 As String            ' 測定値１　生データ４
    MEAS1DAT5 As String            ' 測定値１　生データ５
    MEAS2DAT1 As String            ' 測定値２　生データ１
    MEAS2DAT2 As String            ' 測定値２　生データ２
    MEAS2DAT3 As String            ' 測定値２　生データ３
    MEAS2DAT4 As String            ' 測定値２　生データ４
    MEAS2DAT5 As String            ' 測定値２　生データ５
    MEAS3DAT1 As String            ' 測定値３　生データ１
    MEAS3DAT2 As String            ' 測定値３　生データ２
    MEAS3DAT3 As String            ' 測定値３　生データ３
    MEAS3DAT4 As String            ' 測定値３　生データ４
    MEAS3DAT5 As String            ' 測定値３　生データ５
    MEAS4DAT1 As String            ' 測定値４　生データ１
    MEAS4DAT2 As String            ' 測定値４　生データ２
    MEAS4DAT3 As String            ' 測定値４　生データ３
    MEAS4DAT4 As String            ' 測定値４　生データ４
    MEAS4DAT5 As String            ' 測定値４　生データ５
    MEAS5DAT1 As String            ' 測定値５　生データ１
    MEAS5DAT2 As String            ' 測定値５　生データ２
    MEAS5DAT3 As String            ' 測定値５　生データ３
    MEAS5DAT4 As String            ' 測定値５　生データ４
    MEAS5DAT5 As String            ' 測定値５　生データ５
    MEAS6DAT1 As String            ' 測定値６　生データ１
    MEAS6DAT2 As String            ' 測定値６　生データ２
    MEAS6DAT3 As String            ' 測定値６　生データ３
    MEAS6DAT4 As String            ' 測定値６　生データ４
    MEAS6DAT5 As String            ' 測定値６　生データ５
    MEAS7DAT1 As String            ' 測定値７　生データ１
    MEAS7DAT2 As String            ' 測定値７　生データ２
    MEAS7DAT3 As String            ' 測定値７　生データ３
    MEAS7DAT4 As String            ' 測定値７　生データ４
    MEAS7DAT5 As String            ' 測定値７　生データ５
    MEAS8DAT1 As String            ' 測定値８　生データ１
    MEAS8DAT2 As String            ' 測定値８　生データ２
    MEAS8DAT3 As String            ' 測定値８　生データ３
    MEAS8DAT4 As String            ' 測定値８　生データ４
    MEAS8DAT5 As String            ' 測定値８　生データ５
    MEAS9DAT1 As String            ' 測定値９　生データ１
    MEAS9DAT2 As String            ' 測定値９　生データ２
    MEAS9DAT3 As String            ' 測定値９　生データ３
    MEAS9DAT4 As String            ' 測定値９　生データ４
    MEAS9DAT5 As String            ' 測定値９　生データ５
    MEAS10DAT1 As String           ' 測定値１０　生データ１
    MEAS10DAT2 As String           ' 測定値１０　生データ２
    MEAS10DAT3 As String           ' 測定値１０　生データ３
    MEAS10DAT4 As String           ' 測定値１０　生データ４
    MEAS10DAT5 As String           ' 測定値１０　生データ５
    LTSPIFLG As String             ' 測定位置判定フラグ
' Add End   2005/11/14 M.Makino
End Type

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMJ007」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_cmjc001h_Disp ,抽出レコード
'          :SPLNUMs()     ,I  ,Integer      ,抽出条件配列(サンプルNo)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/20作成　長野
Public Function DBDRV_Getcmjc001h_Disp(records() As typ_cmjc001h_Disp, SMPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim sqlWhere As String  'SQLのWHERE部分
Dim sqlGroup As String  'SQLのGROUP部分
Dim sqlOrder As String  'SQLのOrder部分
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    DBDRV_Getcmjc001h_Disp = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001h_SQL.bas -- Function DBDRV_Getcmjc001h_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT) , SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS "
    sqlBase = sqlBase & "From TBCMJ007"
    ''抽出条件(ｻﾝﾌﾟﾙNO)の取り出し
    sqlWhere = "Where SMPLNO in ("
    For i = 1 To UBound(SMPLNUMs)
        sqlWhere = sqlWhere & "'" & SMPLNUMs(i) & "'"
        If i < UBound(SMPLNUMs) Then
            sqlWhere = sqlWhere & ", "
        End If
    Next
    sqlWhere = sqlWhere & ") "
    sqlGroup = "GROUP BY CRYNUM, POSITION, SMPKBN, TRANCOND "
    sqlOrder = "ORDER BY POSITION"
    sql = sqlBase & sqlWhere & sqlGroup & sqlOrder
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_Getcmjc001h_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .POSITION = rs("POSITION")       ' 位置
            .SMPKBN = rs("SMPKBN")           ' サンプル区分
            .TRANCOND = rs("TRANCOND")       ' 処理条件
            .SMPLNO = rs("SMPLNO")           ' サンプルＮｏ
            .SMPLUMU = rs("SMPLUMU")         ' サンプル有無
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .GOUKI = rs("GOUKI")             ' 号機
            .MEAS1 = rs("MEAS1")             ' 測定値１
            .MEAS2 = rs("MEAS2")             ' 測定値２
            .MEAS3 = rs("MEAS3")             ' 測定値３
            .MEAS4 = rs("MEAS4")             ' 測定値４
            .MEAS5 = rs("MEAS5")             ' 測定値５
            .MEASPEAK = rs("MEASPEAK")       ' 測定値 ピーク値
            .CALCMEAS = rs("CALCMEAS")       ' 計算結果
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001h_Disp = FUNCTION_RETURN_SUCCESS

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

'概要      :引数で渡されたレコードをTBCMJ007に追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :record        ,I  ,typ_cmjc001h_Disp ,抽出レコード
'          :CRYNUM        ,I  ,String       ,結晶番号
'          :TSTAFFID      ,I  ,String       ,登録社員ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :処理回数はテーブル上の最大値+1とする。
'履歴      :2001/06/25(mon)作成　長野

Public Function DBDRV_Getcmjc001h_Exec(record As typ_cmjc001h_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQLベース部分
Dim sqlWhere As String  'SQLWhere部分
Dim sqlGroup As String  'SQLGroup部分

'    CRYNUM             結晶番号　⇒引数
'    TRANCNT         　 処理回数　⇒最大
'   TSTAFFID            登録社員ID　⇒引数
 '   REGDATE 　　　     登録日付　⇒SYSDATE
 '   KSTAFFID           更新社員ID　⇒" "
 '   UPDDATE            更新日付　⇒SYSDATE
 '   SENDFLAG           送信フラグ　⇒"0"
 '   SENDDATE           送信日付　⇒SYSDATE

    DBDRV_Getcmjc001h_Exec = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001h_SQL.bas -- Function DBDRV_Getcmjc001h_Exec"

' Mod Start 2005/11/14 M.Makino
'    sqlBase = "Insert into TBCMJ007 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, " & _
'              "KRPROCCD, PROCCODE, GOUKI, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) " & vbCrLf
'    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
'               record.SMPLNO & ", '" & record.SMPLUMU & "',  '" & record.HINBAN & "', " & record.REVNUM & ",'" & record.FACTORY & "', '" & record.OPECOND & "', '" & _
'               record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', " & record.MEAS1 & ", " & record.MEAS2 & ", " & record.MEAS3 & ", " & _
'               record.MEAS4 & ", " & record.MEAS5 & ", " & record.MEASPEAK & ", " & record.CALCMEAS & ", '" & TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ007 "
'    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') " & vbCrLf

    sqlBase = "Insert into TBCMJ007"
    sqlBase = sqlBase & " (CRYNUM"          ' [結晶番号]
    sqlBase = sqlBase & ", POSITION"        ' [位置]
    sqlBase = sqlBase & ", SMPKBN"          ' [サンプル区分]
    sqlBase = sqlBase & ", TRANCOND"        ' [処理条件]
    sqlBase = sqlBase & ", TRANCNT"         ' [処理回数]
    sqlBase = sqlBase & ", SMPLNO"          ' [サンプルNo]
    sqlBase = sqlBase & ", SMPLUMU"         ' [サンプル有無]
    sqlBase = sqlBase & ", HINBAN"          ' [品番]
    sqlBase = sqlBase & ", REVNUM"          ' [製品番号改訂番号]
    sqlBase = sqlBase & ", FACTORY"         ' [工場]
    sqlBase = sqlBase & ", OPECOND"         ' [操業条件]
    sqlBase = sqlBase & ", KRPROCCD"        ' [管理工程コード]
    sqlBase = sqlBase & ", PROCCODE"        ' [工程コード]
    sqlBase = sqlBase & ", GOUKI"           ' [号機]
    sqlBase = sqlBase & ", MEAS1"           ' [測定値1]
    sqlBase = sqlBase & ", MEAS2"           ' [測定値2]
    sqlBase = sqlBase & ", MEAS3"           ' [測定値3]
    sqlBase = sqlBase & ", MEAS4"           ' [測定値4]
    sqlBase = sqlBase & ", MEAS5"           ' [測定値5]
    sqlBase = sqlBase & ", MEASPEAK"        ' [測定値 ピーク値]
    sqlBase = sqlBase & ", CALCMEAS"        ' [計算結果]
    sqlBase = sqlBase & ", TSTAFFID"        ' [登録社員ID]
    sqlBase = sqlBase & ", REGDATE"         ' [登録日付]
    sqlBase = sqlBase & ", KSTAFFID"        ' [更新社員ID]
    sqlBase = sqlBase & ", UPDDATE"         ' [更新日付]
    sqlBase = sqlBase & ", SENDFLAG"        ' [送信フラグ]
    sqlBase = sqlBase & ", SENDDATE"        ' [送信日付]
    sqlBase = sqlBase & ", MEAS6"           ' [測定値６]
    sqlBase = sqlBase & ", MEAS7"           ' [測定値７]
    sqlBase = sqlBase & ", MEAS8"           ' [測定値８]
    sqlBase = sqlBase & ", MEAS9"           ' [測定値９]
    sqlBase = sqlBase & ", MEAS10"          ' [測定値１０]
    sqlBase = sqlBase & ", MEASFILE"        ' [測定データファイル名]
    sqlBase = sqlBase & ", RESVAL"          ' [実測抵抗]
    sqlBase = sqlBase & ", INCVAL"          ' [傾き]
    sqlBase = sqlBase & ", CUTVAL"          ' [切片]
    sqlBase = sqlBase & ", SETVAL"          ' [設定値]
    sqlBase = sqlBase & ", CONVAL"          ' [１０Ω換算値]
    sqlBase = sqlBase & ", MEAS1DAT1"       ' [測定値１　生データ１]
    sqlBase = sqlBase & ", MEAS1DAT2"       ' [測定値１　生データ２]
    sqlBase = sqlBase & ", MEAS1DAT3"       ' [測定値１　生データ３]
    sqlBase = sqlBase & ", MEAS1DAT4"       ' [測定値１　生データ４]
    sqlBase = sqlBase & ", MEAS1DAT5"       ' [測定値１　生データ５]
    sqlBase = sqlBase & ", MEAS2DAT1"       ' [測定値２　生データ１]
    sqlBase = sqlBase & ", MEAS2DAT2"       ' [測定値２　生データ２]
    sqlBase = sqlBase & ", MEAS2DAT3"       ' [測定値２　生データ３]
    sqlBase = sqlBase & ", MEAS2DAT4"       ' [測定値２　生データ４]
    sqlBase = sqlBase & ", MEAS2DAT5"       ' [測定値２　生データ５]
    sqlBase = sqlBase & ", MEAS3DAT1"       ' [測定値３　生データ１]
    sqlBase = sqlBase & ", MEAS3DAT2"       ' [測定値３　生データ２]
    sqlBase = sqlBase & ", MEAS3DAT3"       ' [測定値３　生データ３]
    sqlBase = sqlBase & ", MEAS3DAT4"       ' [測定値３　生データ４]
    sqlBase = sqlBase & ", MEAS3DAT5"       ' [測定値３　生データ５]
    sqlBase = sqlBase & ", MEAS4DAT1"       ' [測定値４　生データ１]
    sqlBase = sqlBase & ", MEAS4DAT2"       ' [測定値４　生データ２]
    sqlBase = sqlBase & ", MEAS4DAT3"       ' [測定値４　生データ３]
    sqlBase = sqlBase & ", MEAS4DAT4"       ' [測定値４　生データ４]
    sqlBase = sqlBase & ", MEAS4DAT5"       ' [測定値４　生データ５]
    sqlBase = sqlBase & ", MEAS5DAT1"       ' [測定値５　生データ１]
    sqlBase = sqlBase & ", MEAS5DAT2"       ' [測定値５　生データ２]
    sqlBase = sqlBase & ", MEAS5DAT3"       ' [測定値５　生データ３]
    sqlBase = sqlBase & ", MEAS5DAT4"       ' [測定値５　生データ４]
    sqlBase = sqlBase & ", MEAS5DAT5"       ' [測定値５　生データ５]
    sqlBase = sqlBase & ", MEAS6DAT1"       ' [測定値６　生データ１]
    sqlBase = sqlBase & ", MEAS6DAT2"       ' [測定値６　生データ２]
    sqlBase = sqlBase & ", MEAS6DAT3"       ' [測定値６　生データ３]
    sqlBase = sqlBase & ", MEAS6DAT4"       ' [測定値６　生データ４]
    sqlBase = sqlBase & ", MEAS6DAT5"       ' [測定値６　生データ５]
    sqlBase = sqlBase & ", MEAS7DAT1"       ' [測定値７　生データ１]
    sqlBase = sqlBase & ", MEAS7DAT2"       ' [測定値７　生データ２]
    sqlBase = sqlBase & ", MEAS7DAT3"       ' [測定値７　生データ３]
    sqlBase = sqlBase & ", MEAS7DAT4"       ' [測定値７　生データ４]
    sqlBase = sqlBase & ", MEAS7DAT5"       ' [測定値７　生データ５]
    sqlBase = sqlBase & ", MEAS8DAT1"       ' [測定値８　生データ１]
    sqlBase = sqlBase & ", MEAS8DAT2"       ' [測定値８　生データ２]
    sqlBase = sqlBase & ", MEAS8DAT3"       ' [測定値８　生データ３]
    sqlBase = sqlBase & ", MEAS8DAT4"       ' [測定値８　生データ４]
    sqlBase = sqlBase & ", MEAS8DAT5"       ' [測定値８　生データ５]
    sqlBase = sqlBase & ", MEAS9DAT1"       ' [測定値９　生データ１]
    sqlBase = sqlBase & ", MEAS9DAT2"       ' [測定値９　生データ２]
    sqlBase = sqlBase & ", MEAS9DAT3"       ' [測定値９　生データ３]
    sqlBase = sqlBase & ", MEAS9DAT4"       ' [測定値９　生データ４]
    sqlBase = sqlBase & ", MEAS9DAT5"       ' [測定値９　生データ５]
    sqlBase = sqlBase & ", MEAS10DAT1"      ' [測定値１０　生データ１]
    sqlBase = sqlBase & ", MEAS10DAT2"      ' [測定値１０　生データ２]
    sqlBase = sqlBase & ", MEAS10DAT3"      ' [測定値１０　生データ３]
    sqlBase = sqlBase & ", MEAS10DAT4"      ' [測定値１０　生データ４]
    sqlBase = sqlBase & ", MEAS10DAT5"      ' [測定値１０　生データ５]
    sqlBase = sqlBase & ", LTSPIFLG"        ' [測定位置判定フラグ]
    sqlBase = sqlBase & ") select"
    sqlBase = sqlBase & "  '" & CRYNUM & "'"                    ' [結晶番号]
    sqlBase = sqlBase & ", " & record.POSITION                  ' [位置]
    sqlBase = sqlBase & ", '" & record.SMPKBN & "'"             ' [サンプル区分]
    sqlBase = sqlBase & ", '" & record.TRANCOND & "'"           ' [処理条件]
    sqlBase = sqlBase & ", nvl(MAX(TRANCNT),0) + 1"             ' [処理回数]
    sqlBase = sqlBase & ", " & record.SMPLNO                    ' [サンプルNo]
    sqlBase = sqlBase & ", '" & record.SMPLUMU & "'"            ' [サンプル有無]
    sqlBase = sqlBase & ", '" & record.hinban & "'"             ' [品番]
    sqlBase = sqlBase & ", " & record.REVNUM                    ' [製品番号改訂番号]
    sqlBase = sqlBase & ", '" & record.factory & "'"            ' [工場]
    sqlBase = sqlBase & ", '" & record.opecond & "'"            ' [操業条件]
    sqlBase = sqlBase & ", '" & record.KRPROCCD & "'"           ' [管理工程コード]
    sqlBase = sqlBase & ", '" & record.PROCCODE & "'"           ' [工程コード]
    sqlBase = sqlBase & ", '" & record.GOUKI & "'"              ' [号機]
    sqlBase = sqlBase & ", " & record.MEAS1                     ' [測定値1]
    sqlBase = sqlBase & ", " & record.MEAS2                     ' [測定値2]
    sqlBase = sqlBase & ", " & record.MEAS3                     ' [測定値3]
    sqlBase = sqlBase & ", " & record.MEAS4                     ' [測定値4]
    sqlBase = sqlBase & ", " & record.MEAS5                     ' [測定値5]
    sqlBase = sqlBase & ", " & record.MEASPEAK                  ' [測定値 ピーク値]
    sqlBase = sqlBase & ", " & record.CALCMEAS                  ' [計算結果]
    sqlBase = sqlBase & ", '" & TSTAFFID & "'"                  ' [登録社員ID]
    sqlBase = sqlBase & ", SYSDATE"                             ' [登録日付]
    sqlBase = sqlBase & ", ' '"                                 ' [更新社員ID]
    sqlBase = sqlBase & ", SYSDATE"                             ' [更新日付]
    sqlBase = sqlBase & ", '0'"                                 ' [送信フラグ]
    sqlBase = sqlBase & ", SYSDATE"                             ' [送信日付]
    sqlBase = sqlBase & ", " & record.MEAS6                     ' [測定値６]
    sqlBase = sqlBase & ", " & record.MEAS7                     ' [測定値７]
    sqlBase = sqlBase & ", " & record.MEAS8                     ' [測定値８]
    sqlBase = sqlBase & ", " & record.MEAS9                     ' [測定値９]
    sqlBase = sqlBase & ", " & record.MEAS10                    ' [測定値１０]
    sqlBase = sqlBase & ", '" & record.MEASFILE & "'"           ' [測定データファイル名]
    sqlBase = sqlBase & ", " & LZeroToNull(record.RESVAL)       ' [実測抵抗]
    sqlBase = sqlBase & ", " & LZeroToNull(record.INCVAL)       ' [傾き]
    sqlBase = sqlBase & ", " & LZeroToNull(record.CUTVAL)       ' [切片]
    sqlBase = sqlBase & ", " & LZeroToNull(record.SETVAL)       ' [設定値]
    sqlBase = sqlBase & ", " & LZeroToNull(record.CONVAL)       ' [10Ω換算値]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS1DAT1)    ' [測定値１　生データ１]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS1DAT2)    ' [測定値１　生データ２]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS1DAT3)    ' [測定値１　生データ３]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS1DAT4)    ' [測定値１　生データ４]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS1DAT5)    ' [測定値１　生データ５]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS2DAT1)    ' [測定値２　生データ１]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS2DAT2)    ' [測定値２　生データ２]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS2DAT3)    ' [測定値２　生データ３]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS2DAT4)    ' [測定値２　生データ４]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS2DAT5)    ' [測定値２　生データ５]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS3DAT1)    ' [測定値３　生データ１]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS3DAT2)    ' [測定値３　生データ２]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS3DAT3)    ' [測定値３　生データ３]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS3DAT4)    ' [測定値３　生データ４]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS3DAT5)    ' [測定値３　生データ５]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS4DAT1)    ' [測定値４　生データ１]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS4DAT2)    ' [測定値４　生データ２]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS4DAT3)    ' [測定値４　生データ３]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS4DAT4)    ' [測定値４　生データ４]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS4DAT5)    ' [測定値４　生データ５]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS5DAT1)    ' [測定値５　生データ１]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS5DAT2)    ' [測定値５　生データ２]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS5DAT3)    ' [測定値５　生データ３]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS5DAT4)    ' [測定値５　生データ４]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS5DAT5)    ' [測定値５　生データ５]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS6DAT1)    ' [測定値６　生データ１]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS6DAT2)    ' [測定値６　生データ２]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS6DAT3)    ' [測定値６　生データ３]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS6DAT4)    ' [測定値６　生データ４]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS6DAT5)    ' [測定値６　生データ５]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS7DAT1)    ' [測定値７　生データ１]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS7DAT2)    ' [測定値７　生データ２]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS7DAT3)    ' [測定値７　生データ３]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS7DAT4)    ' [測定値７　生データ４]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS7DAT5)    ' [測定値７　生データ５]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS8DAT1)    ' [測定値８　生データ１]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS8DAT2)    ' [測定値８　生データ２]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS8DAT3)    ' [測定値８　生データ３]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS8DAT4)    ' [測定値８　生データ４]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS8DAT5)    ' [測定値８　生データ５]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS9DAT1)    ' [測定値９　生データ１]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS9DAT2)    ' [測定値９　生データ２]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS9DAT3)    ' [測定値９　生データ３]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS9DAT4)    ' [測定値９　生データ４]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS9DAT5)    ' [測定値９　生データ５]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS10DAT1)   ' [測定値１０　生データ１]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS10DAT2)   ' [測定値１０　生データ２]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS10DAT3)   ' [測定値１０　生データ３]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS10DAT4)   ' [測定値１０　生データ４]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS10DAT5)   ' [測定値１０　生データ５]
    sqlBase = sqlBase & ", '" & record.LTSPIFLG & "'"           ' [測定位置判定フラグ]
    sqlBase = sqlBase & " from TBCMJ007"

    sqlWhere = sqlWhere & " where"
    sqlWhere = sqlWhere & " (CRYNUM='" & CRYNUM & "')"
    sqlWhere = sqlWhere & " and (POSITION=" & record.POSITION & ")"
    sqlWhere = sqlWhere & " and (SMPKBN='" & record.SMPKBN & "')"
    sqlWhere = sqlWhere & " and (TRANCOND='" & record.TRANCOND & "')"
' Mod End   2005/11/14 M.Makino

'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup

    ''SQLの実行
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001h_Exec = FUNCTION_RETURN_SUCCESS

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

'概要      :テーブル「TBCMJ007」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ007 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCMJ007(records() As typ_TBCMJ007, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
' Mod Start 2005/11/14 M.Makino
'    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
'              " PROCCODE, GOUKI, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
'              " SENDFLAG, SENDDATE "
    sqlBase = ""
    sqlBase = sqlBase & "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT"
    sqlBase = sqlBase & ", SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD"
    sqlBase = sqlBase & ", PROCCODE, GOUKI"
    sqlBase = sqlBase & ", nvl(MEAS1, -1) MEAS1"
    sqlBase = sqlBase & ", nvl(MEAS2, -1) MEAS2"
    sqlBase = sqlBase & ", nvl(MEAS3, -1) MEAS3"
    sqlBase = sqlBase & ", nvl(MEAS4, -1) MEAS4"
    sqlBase = sqlBase & ", nvl(MEAS5, -1) MEAS5"
    sqlBase = sqlBase & ", MEASPEAK, CALCMEAS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE"
    sqlBase = sqlBase & ", SENDFLAG, SENDDATE"
    sqlBase = sqlBase & ", nvl(MEAS6, -1) MEAS6"
    sqlBase = sqlBase & ", nvl(MEAS7, -1) MEAS7"
    sqlBase = sqlBase & ", nvl(MEAS8, -1) MEAS8"
    sqlBase = sqlBase & ", nvl(MEAS9, -1) MEAS9"
    sqlBase = sqlBase & ", nvl(MEAS10, -1) MEAS10"
'    sqlBase = sqlBase & ", nvl(MEASFILE, ' ') MEASFILE"
'    sqlBase = sqlBase & ", nvl(RESVAL, -1) RESVAL"
'    sqlBase = sqlBase & ", nvl(INCVAL, -1) INCVAL"
'    sqlBase = sqlBase & ", nvl(CUTVAL, -1) CUTVAL"
'    sqlBase = sqlBase & ", nvl(SETVAL, -1) SETVAL"
'    sqlBase = sqlBase & ", nvl(CONVAL, -1) CONVAL"
'    sqlBase = sqlBase & ", nvl(MEAS1DAT1, -1) MEAS1DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS1DAT2, -1) MEAS1DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS1DAT3, -1) MEAS1DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS1DAT4, -1) MEAS1DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS1DAT5, -1) MEAS1DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS2DAT1, -1) MEAS2DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS2DAT2, -1) MEAS2DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS2DAT3, -1) MEAS2DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS2DAT4, -1) MEAS2DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS2DAT5, -1) MEAS2DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS3DAT1, -1) MEAS3DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS3DAT2, -1) MEAS3DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS3DAT3, -1) MEAS3DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS3DAT4, -1) MEAS3DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS3DAT5, -1) MEAS3DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS4DAT1, -1) MEAS4DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS4DAT2, -1) MEAS4DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS4DAT3, -1) MEAS4DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS4DAT4, -1) MEAS4DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS4DAT5, -1) MEAS4DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS5DAT1, -1) MEAS5DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS5DAT2, -1) MEAS5DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS5DAT3, -1) MEAS5DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS5DAT4, -1) MEAS5DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS5DAT5, -1) MEAS5DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS6DAT1, -1) MEAS6DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS6DAT2, -1) MEAS6DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS6DAT3, -1) MEAS6DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS6DAT4, -1) MEAS6DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS6DAT5, -1) MEAS6DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS7DAT1, -1) MEAS7DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS7DAT2, -1) MEAS7DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS7DAT3, -1) MEAS7DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS7DAT4, -1) MEAS7DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS7DAT5, -1) MEAS7DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS8DAT1, -1) MEAS8DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS8DAT2, -1) MEAS8DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS8DAT3, -1) MEAS8DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS8DAT4, -1) MEAS8DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS8DAT5, -1) MEAS8DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS9DAT1, -1) MEAS9DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS9DAT2, -1) MEAS9DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS9DAT3, -1) MEAS9DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS9DAT4, -1) MEAS9DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS9DAT5, -1) MEAS9DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS10DAT1, -1) MEAS10DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS10DAT2, -1) MEAS10DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS10DAT3, -1) MEAS10DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS10DAT4, -1) MEAS10DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS10DAT5, -1) MEAS10DAT5"
'    sqlBase = sqlBase & ", nvl(LTSPIFLG, -1) LTSPIFLG"
    sqlBase = sqlBase & ", MEASFILE"
    sqlBase = sqlBase & ", RESVAL"
    sqlBase = sqlBase & ", INCVAL"
    sqlBase = sqlBase & ", CUTVAL"
    sqlBase = sqlBase & ", SETVAL"
    sqlBase = sqlBase & ", CONVAL"
    sqlBase = sqlBase & ", MEAS1DAT1"
    sqlBase = sqlBase & ", MEAS1DAT2"
    sqlBase = sqlBase & ", MEAS1DAT3"
    sqlBase = sqlBase & ", MEAS1DAT4"
    sqlBase = sqlBase & ", MEAS1DAT5"
    sqlBase = sqlBase & ", MEAS2DAT1"
    sqlBase = sqlBase & ", MEAS2DAT2"
    sqlBase = sqlBase & ", MEAS2DAT3"
    sqlBase = sqlBase & ", MEAS2DAT4"
    sqlBase = sqlBase & ", MEAS2DAT5"
    sqlBase = sqlBase & ", MEAS3DAT1"
    sqlBase = sqlBase & ", MEAS3DAT2"
    sqlBase = sqlBase & ", MEAS3DAT3"
    sqlBase = sqlBase & ", MEAS3DAT4"
    sqlBase = sqlBase & ", MEAS3DAT5"
    sqlBase = sqlBase & ", MEAS4DAT1"
    sqlBase = sqlBase & ", MEAS4DAT2"
    sqlBase = sqlBase & ", MEAS4DAT3"
    sqlBase = sqlBase & ", MEAS4DAT4"
    sqlBase = sqlBase & ", MEAS4DAT5"
    sqlBase = sqlBase & ", MEAS5DAT1"
    sqlBase = sqlBase & ", MEAS5DAT2"
    sqlBase = sqlBase & ", MEAS5DAT3"
    sqlBase = sqlBase & ", MEAS5DAT4"
    sqlBase = sqlBase & ", MEAS5DAT5"
    sqlBase = sqlBase & ", MEAS6DAT1"
    sqlBase = sqlBase & ", MEAS6DAT2"
    sqlBase = sqlBase & ", MEAS6DAT3"
    sqlBase = sqlBase & ", MEAS6DAT4"
    sqlBase = sqlBase & ", MEAS6DAT5"
    sqlBase = sqlBase & ", MEAS7DAT1"
    sqlBase = sqlBase & ", MEAS7DAT2"
    sqlBase = sqlBase & ", MEAS7DAT3"
    sqlBase = sqlBase & ", MEAS7DAT4"
    sqlBase = sqlBase & ", MEAS7DAT5"
    sqlBase = sqlBase & ", MEAS8DAT1"
    sqlBase = sqlBase & ", MEAS8DAT2"
    sqlBase = sqlBase & ", MEAS8DAT3"
    sqlBase = sqlBase & ", MEAS8DAT4"
    sqlBase = sqlBase & ", MEAS8DAT5"
    sqlBase = sqlBase & ", MEAS9DAT1"
    sqlBase = sqlBase & ", MEAS9DAT2"
    sqlBase = sqlBase & ", MEAS9DAT3"
    sqlBase = sqlBase & ", MEAS9DAT4"
    sqlBase = sqlBase & ", MEAS9DAT5"
    sqlBase = sqlBase & ", MEAS10DAT1"
    sqlBase = sqlBase & ", MEAS10DAT2"
    sqlBase = sqlBase & ", MEAS10DAT3"
    sqlBase = sqlBase & ", MEAS10DAT4"
    sqlBase = sqlBase & ", MEAS10DAT5"
    sqlBase = sqlBase & ", LTSPIFLG"
' Mod End   2005/11/14 M.Makino
    sqlBase = sqlBase & " From TBCMJ007"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ007 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .POSITION = rs("POSITION")       ' 位置
            .SMPKBN = rs("SMPKBN")           ' サンプル区分
            .TRANCOND = rs("TRANCOND")       ' 処理条件
            .TRANCNT = rs("TRANCNT")         ' 処理回数
            .SMPLNO = rs("SMPLNO")           ' サンプルＮｏ
            .SMPLUMU = rs("SMPLUMU")         ' サンプル有無
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .GOUKI = rs("GOUKI")             ' 号機
            .MEAS1 = rs("MEAS1")             ' 測定値１
            .MEAS2 = rs("MEAS2")             ' 測定値２
            .MEAS3 = rs("MEAS3")             ' 測定値３
            .MEAS4 = rs("MEAS4")             ' 測定値４
            .MEAS5 = rs("MEAS5")             ' 測定値５
            .MEASPEAK = rs("MEASPEAK")       ' 測定値 ピーク値
            .CALCMEAS = rs("CALCMEAS")       ' 計算結果
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
' Add Start 2005/11/14 M.Makino
            .MEAS6 = rs("MEAS6")             ' 測定値６
            .MEAS7 = rs("MEAS7")             ' 測定値７
            .MEAS8 = rs("MEAS8")             ' 測定値８
            .MEAS9 = rs("MEAS9")             ' 測定値９
            .MEAS10 = rs("MEAS10")           ' 測定値１０
'            .MEASFILE = rs("MEASFILE")       ' 測定データファイル名
'            .RESVAL = rs("RESVAL")           ' 実測抵抗
'            .INCVAL = rs("INCVAL")           ' 傾き
'            .CUTVAL = rs("CUTVAL")           ' 切片
'            .SETVAL = rs("SETVAL")           ' 設定値
'            .CONVAL = rs("RESVAL")           ' 10Ω換算値
'            .MEAS1DAT1 = rs("MEAS1DAT1")     ' 測定値１　生データ１
'            .MEAS1DAT2 = rs("MEAS1DAT2")     ' 測定値１　生データ２
'            .MEAS1DAT3 = rs("MEAS1DAT3")     ' 測定値１　生データ３
'            .MEAS1DAT4 = rs("MEAS1DAT4")     ' 測定値１　生データ４
'            .MEAS1DAT5 = rs("MEAS1DAT5")     ' 測定値１　生データ５
'            .MEAS2DAT1 = rs("MEAS2DAT1")     ' 測定値２　生データ１
'            .MEAS2DAT2 = rs("MEAS2DAT2")     ' 測定値２　生データ２
'            .MEAS2DAT3 = rs("MEAS2DAT3")     ' 測定値２　生データ３
'            .MEAS2DAT4 = rs("MEAS2DAT4")     ' 測定値２　生データ４
'            .MEAS2DAT5 = rs("MEAS2DAT5")     ' 測定値２　生データ５
'            .MEAS3DAT1 = rs("MEAS3DAT1")     ' 測定値３　生データ１
'            .MEAS3DAT2 = rs("MEAS3DAT2")     ' 測定値３　生データ２
'            .MEAS3DAT3 = rs("MEAS3DAT3")     ' 測定値３　生データ３
'            .MEAS3DAT4 = rs("MEAS3DAT4")     ' 測定値３　生データ４
'            .MEAS3DAT5 = rs("MEAS3DAT5")     ' 測定値３　生データ５
'            .MEAS4DAT1 = rs("MEAS4DAT1")     ' 測定値４　生データ１
'            .MEAS4DAT2 = rs("MEAS4DAT2")     ' 測定値４　生データ２
'            .MEAS4DAT3 = rs("MEAS4DAT3")     ' 測定値４　生データ３
'            .MEAS4DAT4 = rs("MEAS4DAT4")     ' 測定値４　生データ４
'            .MEAS4DAT5 = rs("MEAS4DAT5")     ' 測定値４　生データ５
'            .MEAS5DAT1 = rs("MEAS5DAT1")     ' 測定値５　生データ１
'            .MEAS5DAT2 = rs("MEAS5DAT2")     ' 測定値５　生データ２
'            .MEAS5DAT3 = rs("MEAS5DAT3")     ' 測定値５　生データ３
'            .MEAS5DAT4 = rs("MEAS5DAT4")     ' 測定値５　生データ４
'            .MEAS5DAT5 = rs("MEAS5DAT5")     ' 測定値５　生データ５
'            .MEAS6DAT1 = rs("MEAS6DAT1")     ' 測定値６　生データ１
'            .MEAS6DAT2 = rs("MEAS6DAT2")     ' 測定値６　生データ２
'            .MEAS6DAT3 = rs("MEAS6DAT3")     ' 測定値６　生データ３
'            .MEAS6DAT4 = rs("MEAS6DAT4")     ' 測定値６　生データ４
'            .MEAS6DAT5 = rs("MEAS6DAT5")     ' 測定値６　生データ５
'            .MEAS7DAT1 = rs("MEAS7DAT1")     ' 測定値７　生データ１
'            .MEAS7DAT2 = rs("MEAS7DAT2")     ' 測定値７　生データ２
'            .MEAS7DAT3 = rs("MEAS7DAT3")     ' 測定値７　生データ３
'            .MEAS7DAT4 = rs("MEAS7DAT4")     ' 測定値７　生データ４
'            .MEAS7DAT5 = rs("MEAS7DAT5")     ' 測定値７　生データ５
'            .MEAS8DAT1 = rs("MEAS8DAT1")     ' 測定値８　生データ１
'            .MEAS8DAT2 = rs("MEAS8DAT2")     ' 測定値８　生データ２
'            .MEAS8DAT3 = rs("MEAS8DAT3")     ' 測定値８　生データ３
'            .MEAS8DAT4 = rs("MEAS8DAT4")     ' 測定値８　生データ４
'            .MEAS8DAT5 = rs("MEAS8DAT5")     ' 測定値８　生データ５
'            .MEAS9DAT1 = rs("MEAS9DAT1")     ' 測定値９　生データ１
'            .MEAS9DAT2 = rs("MEAS9DAT2")     ' 測定値９　生データ２
'            .MEAS9DAT3 = rs("MEAS9DAT3")     ' 測定値９　生データ３
'            .MEAS9DAT4 = rs("MEAS9DAT4")     ' 測定値９　生データ４
'            .MEAS9DAT5 = rs("MEAS9DAT5")     ' 測定値９　生データ５
'            .MEAS10DAT1 = rs("MEAS10DAT1")   ' 測定値１０　生データ１
'            .MEAS10DAT2 = rs("MEAS10DAT2")   ' 測定値１０　生データ２
'            .MEAS10DAT3 = rs("MEAS10DAT3")   ' 測定値１０　生データ３
'            .MEAS10DAT4 = rs("MEAS10DAT4")   ' 測定値１０　生データ４
'            .MEAS10DAT5 = rs("MEAS10DAT5")   ' 測定値１０　生データ５
'            .LTSPIFLG = rs("LTSPIFLG")       ' 測定位置判定フラグ
            .MEASFILE = NulltoStr(rs("MEASFILE"))       ' 測定データファイル名
            .RESVAL = NulltoStr(rs("RESVAL"))           ' 実測抵抗
            .INCVAL = NulltoStr(rs("INCVAL"))           ' 傾き
            .CUTVAL = NulltoStr(rs("CUTVAL"))           ' 切片
            .SETVAL = NulltoStr(rs("SETVAL"))           ' 設定値
            .CONVAL = NulltoStr(rs("CONVAL"))           ' 10Ω換算値
            .MEAS1DAT1 = NulltoStr(rs("MEAS1DAT1"))     ' 測定値１　生データ１
            .MEAS1DAT2 = NulltoStr(rs("MEAS1DAT2"))     ' 測定値１　生データ２
            .MEAS1DAT3 = NulltoStr(rs("MEAS1DAT3"))     ' 測定値１　生データ３
            .MEAS1DAT4 = NulltoStr(rs("MEAS1DAT4"))     ' 測定値１　生データ４
            .MEAS1DAT5 = NulltoStr(rs("MEAS1DAT5"))     ' 測定値１　生データ５
            .MEAS2DAT1 = NulltoStr(rs("MEAS2DAT1"))     ' 測定値２　生データ１
            .MEAS2DAT2 = NulltoStr(rs("MEAS2DAT2"))     ' 測定値２　生データ２
            .MEAS2DAT3 = NulltoStr(rs("MEAS2DAT3"))     ' 測定値２　生データ３
            .MEAS2DAT4 = NulltoStr(rs("MEAS2DAT4"))     ' 測定値２　生データ４
            .MEAS2DAT5 = NulltoStr(rs("MEAS2DAT5"))     ' 測定値２　生データ５
            .MEAS3DAT1 = NulltoStr(rs("MEAS3DAT1"))     ' 測定値３　生データ１
            .MEAS3DAT2 = NulltoStr(rs("MEAS3DAT2"))     ' 測定値３　生データ２
            .MEAS3DAT3 = NulltoStr(rs("MEAS3DAT3"))     ' 測定値３　生データ３
            .MEAS3DAT4 = NulltoStr(rs("MEAS3DAT4"))     ' 測定値３　生データ４
            .MEAS3DAT5 = NulltoStr(rs("MEAS3DAT5"))     ' 測定値３　生データ５
            .MEAS4DAT1 = NulltoStr(rs("MEAS4DAT1"))     ' 測定値４　生データ１
            .MEAS4DAT2 = NulltoStr(rs("MEAS4DAT2"))     ' 測定値４　生データ２
            .MEAS4DAT3 = NulltoStr(rs("MEAS4DAT3"))     ' 測定値４　生データ３
            .MEAS4DAT4 = NulltoStr(rs("MEAS4DAT4"))     ' 測定値４　生データ４
            .MEAS4DAT5 = NulltoStr(rs("MEAS4DAT5"))     ' 測定値４　生データ５
            .MEAS5DAT1 = NulltoStr(rs("MEAS5DAT1"))     ' 測定値５　生データ１
            .MEAS5DAT2 = NulltoStr(rs("MEAS5DAT2"))     ' 測定値５　生データ２
            .MEAS5DAT3 = NulltoStr(rs("MEAS5DAT3"))     ' 測定値５　生データ３
            .MEAS5DAT4 = NulltoStr(rs("MEAS5DAT4"))     ' 測定値５　生データ４
            .MEAS5DAT5 = NulltoStr(rs("MEAS5DAT5"))     ' 測定値５　生データ５
            .MEAS6DAT1 = NulltoStr(rs("MEAS6DAT1"))     ' 測定値６　生データ１
            .MEAS6DAT2 = NulltoStr(rs("MEAS6DAT2"))     ' 測定値６　生データ２
            .MEAS6DAT3 = NulltoStr(rs("MEAS6DAT3"))     ' 測定値６　生データ３
            .MEAS6DAT4 = NulltoStr(rs("MEAS6DAT4"))     ' 測定値６　生データ４
            .MEAS6DAT5 = NulltoStr(rs("MEAS6DAT5"))     ' 測定値６　生データ５
            .MEAS7DAT1 = NulltoStr(rs("MEAS7DAT1"))     ' 測定値７　生データ１
            .MEAS7DAT2 = NulltoStr(rs("MEAS7DAT2"))     ' 測定値７　生データ２
            .MEAS7DAT3 = NulltoStr(rs("MEAS7DAT3"))     ' 測定値７　生データ３
            .MEAS7DAT4 = NulltoStr(rs("MEAS7DAT4"))     ' 測定値７　生データ４
            .MEAS7DAT5 = NulltoStr(rs("MEAS7DAT5"))     ' 測定値７　生データ５
            .MEAS8DAT1 = NulltoStr(rs("MEAS8DAT1"))     ' 測定値８　生データ１
            .MEAS8DAT2 = NulltoStr(rs("MEAS8DAT2"))     ' 測定値８　生データ２
            .MEAS8DAT3 = NulltoStr(rs("MEAS8DAT3"))     ' 測定値８　生データ３
            .MEAS8DAT4 = NulltoStr(rs("MEAS8DAT4"))     ' 測定値８　生データ４
            .MEAS8DAT5 = NulltoStr(rs("MEAS8DAT5"))     ' 測定値８　生データ５
            .MEAS9DAT1 = NulltoStr(rs("MEAS9DAT1"))     ' 測定値９　生データ１
            .MEAS9DAT2 = NulltoStr(rs("MEAS9DAT2"))     ' 測定値９　生データ２
            .MEAS9DAT3 = NulltoStr(rs("MEAS9DAT3"))     ' 測定値９　生データ３
            .MEAS9DAT4 = NulltoStr(rs("MEAS9DAT4"))     ' 測定値９　生データ４
            .MEAS9DAT5 = NulltoStr(rs("MEAS9DAT5"))     ' 測定値９　生データ５
            .MEAS10DAT1 = NulltoStr(rs("MEAS10DAT1"))   ' 測定値１０　生データ１
            .MEAS10DAT2 = NulltoStr(rs("MEAS10DAT2"))   ' 測定値１０　生データ２
            .MEAS10DAT3 = NulltoStr(rs("MEAS10DAT3"))   ' 測定値１０　生データ３
            .MEAS10DAT4 = NulltoStr(rs("MEAS10DAT4"))   ' 測定値１０　生データ４
            .MEAS10DAT5 = NulltoStr(rs("MEAS10DAT5"))   ' 測定値１０　生データ５
            .LTSPIFLG = Trim(NulltoStr(rs("LTSPIFLG"))) ' 測定位置判定フラグ
' Mod End   2005/11/14 M.Makino
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMJ007 = FUNCTION_RETURN_SUCCESS
End Function

'概要      :テーブル「KODA9」から条件にあった１０Ω換算式設定レコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tbl_OumConv   ,O  ,typ_OumConvSet   ,10Ω換算値取得構造体
'          :sType         ,I  ,String           ,タイプ
'          :戻り値        ,O  ,FUNCTION_RETURN  ,抽出の成否
'説明      :
'履歴      :2005/11/11作成　牧野
Public Function DBDRV_OumConvGet(record As typ_OumConvSet, sType As String) As Integer

    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet

    DBDRV_OumConvGet = FUNCTION_RETURN_FAILURE

    ' SQL文作成
    sql = ""
    sql = sql & "SELECT CTR01A9, CTR02A9, CTR03A9"
    sql = sql & " FROM  KODA9"
    sql = sql & " WHERE SYSCA9 = 'X'"
    sql = sql & " AND   SHUCA9 = '19'"
    sql = sql & " AND   CODEA9 = '" & sType & "'"

    ' １０Ω換算式設定を取得する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    If rs Is Nothing Then
        Exit Function
    End If

    ' 該当するデータが無い場合判定はエラー
    If rs.EOF Then
        Exit Function
    End If

    ' 抽出結果を格納する
    With record
        ' [傾き]
        .CTR01A9 = Trim(CStr(NulltoStr(rs.Fields("CTR01A9").Value)))
        ' [切片]
        .CTR02A9 = Trim(CStr(NulltoStr(rs.Fields("CTR02A9").Value)))
        ' [設定値]
        .CTR03A9 = Trim(CStr(NulltoStr(rs.Fields("CTR03A9").Value)))
    End With
    
    DBDRV_OumConvGet = FUNCTION_RETURN_SUCCESS

End Function

'概要      :実測抵抗の取得と１０Ω換算値の算出を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型           ,説明
'          :tblCrySmpMan  ,I   ,typ_XSDCS    ,サンプルID
'          :sKekka        ,I   ,String       ,測定結果
'          :sIncval       ,I   ,String       ,傾き
'          :sCutval       ,I   ,String       ,切片
'          :sSetval       ,I   ,String       ,設定値
'          :sJiteiko      ,I   ,String       ,実測抵抗
'          :sKansanchi    ,I   ,String       ,１０Ω換算値
'          :戻り値        ,O  ,FUNCTION_RETURN  ,抽出の成否
'説明      :
'備考      : １０Ω換算式の算出方法
'               Ａ＝ライフタイム測定結果
'               Ｂ＝実測抵抗
'               Ｃ＝切片 [桁数=XXX.XX]
'               Ｄ＝傾き [桁数=XXX.XX]
'               Ｇ＝設定値 [桁数=XXX.XX]
'               Ｅ＝理論値LT＝Ｄ×Ｂ＋Ｃ
'               Ｆ＝汚染量推定値＝１／((1／Ａ)―(1／Ｅ))
'               １０Ω換算値＝１／((１／Ｇ)＋(１／Ｆ)) [桁数=XXXX]
'履歴      :新規 2005/11/14 M.Makino
''Public Function GetKansanchi(tblCrySmpMan As typ_XSDCS, sKekka As String, sIncVal As String, _
''        sCutVal As String, sSetVal As String, sJiteiko As String, sKansanchi As String) As Integer
''    Dim sql As String       'SQL全体
''    Dim rs As OraDynaset    'RecordSet
''    Dim RironchiLT As Double    ' 理論値LT
''    Dim Osenryo As Double       ' 汚染量推定値

''    GetKansanchi = FUNCTION_RETURN_FAILURE

    ' SQL文作成
''    sql = ""
''    sql = sql & "SELECT MEAS1"
''    sql = sql & " FROM  TBCMJ002"
''    sql = sql & " WHERE CRYNUM='" & tblCrySmpMan.XTALCS & "'"
''    sql = sql & " AND   POSITION=" & tblCrySmpMan.INPOSCS
''    sql = sql & " AND   SMPKBN='" & tblCrySmpMan.SMPKBNCS & "'"
''    sql = sql & " AND   TRANCOND='0'"
''    sql = sql & " ORDER BY TRANCNT DESC"

''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
''    If rs Is Nothing Then
''        Exit Function
''    End If

''    If rs.EOF Then
        ' 該当するデータが無い場合判定は空文字
''        sJiteiko = ""
''    Else
''        sJiteiko = Trim(CStr(NulltoStr(rs.Fields("MEAS1").Value)))
''    End If

    ' １０Ω換算値の計算
''    If sKekka <> "" And sIncVal <> "" And sCutVal <> "" And _
''       sSetVal <> "" And sJiteiko <> "" Then

        '0の除算対策
''        On Error GoTo ERROR_CALC

        '１０Ω換算値を算出
''        RironchiLT = CDbl(sIncVal) * CDbl(sJiteiko) + CDbl(sCutVal)
''        Osenryo = 1 / ((1 / CInt(sKekka)) - (1 / RironchiLT))
''        sKansanchi = CStr(Round(1 / ((1 / CDbl(sSetVal)) + (1 / Osenryo)), 0))
''    Else
''        sKansanchi = ""
''    End If
    
''    GetKansanchi = FUNCTION_RETURN_SUCCESS
''    Exit Function

''ERROR_CALC:
''    sKansanchi = ""
''    GetKansanchi = FUNCTION_RETURN_SUCCESS
''End Function

'
' 空文字列（""）に対して『null』を返し，その他の文字列は何もせずに返す
'
'履歴      :2005/11/14追加　牧野
''Private Function LZeroToNull(ByVal sTmp As String) As String
''    If "" = sTmp Then
''        LZeroToNull = "null"
''    Else
''        LZeroToNull = sTmp
''    End If
''End Function

