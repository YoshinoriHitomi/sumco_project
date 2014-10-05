Attribute VB_Name = "s_cmbc025_SQL"
Option Explicit
'
'                                     2001/06/20
'================================================
' DBアクセス関数
' 定義内容: TBCMJ005 (ＯＳＦ実績)
' 参照　　: 060211_結晶検査
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmjc001f_Disp
    
   ' CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
   ' TRANCNT As Integer              ' 処理回数
    SMPLNO As Integer               ' サンプルＮｏ
    SMPLUMU As String * 1           ' サンプル有無
    HINBAN As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    GOUKI As String * 3             ' 号機
    MEASMETH As String * 1          ' 測定方法
    MEASSPOT As Integer             ' 測定点
    MAG As String * 4               ' 倍率
    HTPRC As String * 2             ' 熱処理方法
    KKSP As String * 3              ' 結晶欠陥測定位置
    KKSET As String * 3             ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    CALCMAX As Double               ' 計算結果 Max
    CALCAVE As Double               ' 計算結果 Ave
    MEAS1 As Double                 ' 測定値１
    MEAS2 As Double                 ' 測定値２
    MEAS3 As Double                 ' 測定値３
    MEAS4 As Double                 ' 測定値４
    MEAS5 As Double                 ' 測定値５
    MEAS6 As Double                 ' 測定値６
    MEAS7 As Double                 ' 測定値７
    MEAS8 As Double                 ' 測定値８
    MEAS9 As Double                 ' 測定値９
    MEAS10 As Double                ' 測定値１０
    MEAS11 As Double                ' 測定値１１
    MEAS12 As Double                ' 測定値１２
    MEAS13 As Double                ' 測定値１３
    MEAS14 As Double                ' 測定値１４
    MEAS15 As Double                ' 測定値１５
    MEAS16 As Double                ' 測定値１６
    MEAS17 As Double                ' 測定値１７
    MEAS18 As Double                ' 測定値１８
    MEAS19 As Double                ' 測定値１９
    MEAS20 As Double                ' 測定値２０
   ' TSTAFFID As String * 8          ' 登録社員ID
   ' REGDATE As Date                 ' 登録日付
   ' KSTAFFID As String * 8          ' 更新社員ID
   ' UPDDATE As Date                 ' 更新日付
   ' SENDFLAG As String * 1          ' 送信フラグ
   ' SENDDATE As Date                ' 送信日付
'OSF，BMD項目追加対応  2002.04.02 yakimura
    OSFPOS1 As Double               ' ﾊﾟﾀｰﾝ区分１位置
    OSFWID1 As Double               ' ﾊﾟﾀｰﾝ区分１幅
    OSFRD1  As String               ' ﾊﾟﾀｰﾝ区分１R/D
    OSFPOS2 As Double               ' ﾊﾟﾀｰﾝ区分２位置
    OSFWID2 As Double               ' ﾊﾟﾀｰﾝ区分２幅
    OSFRD2  As String               ' ﾊﾟﾀｰﾝ区分２R/D
    OSFPOS3 As Double               ' ﾊﾟﾀｰﾝ区分３位置
    OSFWID3 As Double               ' ﾊﾟﾀｰﾝ区分３幅
    OSFRD3  As String               ' ﾊﾟﾀｰﾝ区分３R/D
'OSF，BMD項目追加対応  2002.04.02 yakimura
  
End Type

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMJ005」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_cmjc001f_Disp ,抽出レコード
'          :SPLNUMs()     ,I  ,Integer      ,抽出条件配列(サンプルNo)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/20作成　長野
Public Function DBDRV_Getcmjc001f_Disp(records() As typ_cmjc001f_Disp, SMPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim sqlWhere As String  'SQLのWHERE部分
Dim sqlGroup As String  'SQLのGROUP部分
Dim sqlOrder As String  'SQLのOrder部分
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
 
  
    DBDRV_Getcmjc001f_Disp = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmjc001f_SQL.bas -- Function DBDRV_Getcmjc001f_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, CALCMAX, CALCAVE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEAS6," & _
              " MEAS7, MEAS8, MEAS9, MEAS10, MEAS11, MEAS12, MEAS13, MEAS14, MEAS15, MEAS16, MEAS17, MEAS18, MEAS19, MEAS20, " & _
              " OSFPOS1, OSFWID1, OSFRD1, OSFPOS2, OSFWID2, OSFRD2, OSFPOS3, OSFWID3, OSFRD3 "
    sqlBase = sqlBase & "From TBCMJ005"
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
        DBDRV_Getcmjc001f_Disp = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
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
            .HINBAN = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .FACTORY = rs("FACTORY")         ' 工場
            .OPECOND = rs("OPECOND")         ' 操業条件
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .GOUKI = rs("GOUKI")             ' 号機
            .MEASMETH = rs("MEASMETH")       ' 測定方法
            .MEASSPOT = rs("MEASSPOT")       ' 測定点
            .MAG = rs("MAG")                 ' 倍率
            .HTPRC = rs("HTPRC")             ' 熱処理方法
            .KKSP = rs("KKSP")               ' 結晶欠陥測定位置
            .KKSET = rs("KKSET")             ' 結晶欠陥測定条件＋選択ET代
            .CALCMAX = rs("CALCMAX")         ' 計算結果 Max
            .CALCAVE = rs("CALCAVE")         ' 計算結果 Ave
            .MEAS1 = rs("MEAS1")             ' 測定値１
            .MEAS2 = rs("MEAS2")             ' 測定値２
            .MEAS3 = rs("MEAS3")             ' 測定値３
            .MEAS4 = rs("MEAS4")             ' 測定値４
            .MEAS5 = rs("MEAS5")             ' 測定値５
            .MEAS6 = rs("MEAS6")             ' 測定値６
            .MEAS7 = rs("MEAS7")             ' 測定値７
            .MEAS8 = rs("MEAS8")             ' 測定値８
            .MEAS9 = rs("MEAS9")             ' 測定値９
            .MEAS10 = rs("MEAS10")           ' 測定値１０
            .MEAS11 = rs("MEAS11")           ' 測定値１１
            .MEAS12 = rs("MEAS12")           ' 測定値１２
            .MEAS13 = rs("MEAS13")           ' 測定値１３
            .MEAS14 = rs("MEAS14")           ' 測定値１４
            .MEAS15 = rs("MEAS15")           ' 測定値１５
            .MEAS16 = rs("MEAS16")           ' 測定値１６
            .MEAS17 = rs("MEAS17")           ' 測定値１７
            .MEAS18 = rs("MEAS18")           ' 測定値１８
            .MEAS19 = rs("MEAS19")           ' 測定値１９
            .MEAS20 = rs("MEAS20")           ' 測定値２０
'OSF，BMD項目追加対応  2002.04.02 yakimura
            .OSFPOS1 = rs("OSFPOS1")         ' ﾊﾟﾀｰﾝ区分１位置
            .OSFWID1 = rs("OSFWID1")         ' ﾊﾟﾀｰﾝ区分１幅
            .OSFRD1 = rs("OSFRD1")           ' ﾊﾟﾀｰﾝ区分１R/D
            .OSFPOS2 = rs("OSFPOS2")         ' ﾊﾟﾀｰﾝ区分２位置
            .OSFWID2 = rs("OSFWID2")         ' ﾊﾟﾀｰﾝ区分２幅
            .OSFRD2 = rs("OSFRD2")           ' ﾊﾟﾀｰﾝ区分２R/D
            .OSFPOS3 = rs("OSFPOS3")         ' ﾊﾟﾀｰﾝ区分３位置
            .OSFWID3 = rs("OSFWID3")         ' ﾊﾟﾀｰﾝ区分３幅
            .OSFRD3 = rs("OSFRD3")           ' ﾊﾟﾀｰﾝ区分３R/D
'OSF，BMD項目追加対応  2002.04.02 yakimura
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001f_Disp = FUNCTION_RETURN_SUCCESS

PROC_EXIT:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume PROC_EXIT
End Function

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :引数で渡されたレコードをTBCMJ005に追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :record        ,I  ,typ_cmjc001f_Disp ,抽出レコード
'          :CRYNUM        ,I  ,String       ,結晶番号
'          :TSTAFFID      ,I  ,String       ,登録社員ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :処理回数はテーブル上の最大値+1とする。
'履歴      :2001/06/22(Fri)作成　長野

Public Function DBDRV_Getcmjc001f_Exec(record As typ_cmjc001f_Disp, Crynum$, TSTAFFID$) As FUNCTION_RETURN

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
     
     
    DBDRV_Getcmjc001f_Exec = FUNCTION_RETURN_FAILURE
    
    If Left(record.OSFRD1, 1) = "R" Then
       record.OSFRD1 = "R"
    ElseIf Left(record.OSFRD1, 1) = "D" Then
       record.OSFRD1 = "D"
    Else
       record.OSFRD1 = "-"
    End If
    
    If Left(record.OSFRD2, 1) = "R" Then
       record.OSFRD2 = "R"
    ElseIf Left(record.OSFRD2, 1) = "D" Then
       record.OSFRD2 = "D"
    Else
       record.OSFRD2 = "-"
    End If
    
    If Left(record.OSFRD3, 1) = "R" Then
       record.OSFRD3 = "R"
    ElseIf Left(record.OSFRD3, 1) = "D" Then
       record.OSFRD3 = "D"
    Else
       record.OSFRD3 = "-"
    End If
    
    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmjc001f_SQL.bas -- Function DBDRV_Getcmjc001f_Exec"

    sqlBase = "Insert into TBCMJ005 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, CALCMAX, CALCAVE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEAS6, " & _
              "MEAS7, MEAS8, MEAS9, MEAS10, MEAS11, MEAS12, MEAS13, MEAS14, MEAS15, MEAS16, MEAS17, MEAS18, MEAS19, MEAS20, " & _
              "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, OSFPOS1, OSFWID1, OSFRD1, OSFPOS2, OSFWID2, OSFRD2, OSFPOS3, OSFWID3, OSFRD3 ) "
    sqlBase = sqlBase & "select '" & Crynum & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.HINBAN & "', " & record.REVNUM & ", '" & record.FACTORY & "', '" & _
               record.OPECOND & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', '" & record.MEASMETH & "', " & _
               record.MEASSPOT & ", '" & record.MAG & "', '" & record.HTPRC & "', '" & record.KKSP & "', '" & record.KKSET & "', " & _
               record.CALCMAX & ", " & record.CALCAVE & ", " & record.MEAS1 & ", " & record.MEAS2 & ", " & record.MEAS3 & ", " & record.MEAS4 & ", " & _
               record.MEAS5 & ", " & record.MEAS6 & ", " & record.MEAS7 & ", " & record.MEAS8 & ", " & record.MEAS9 & ", " & record.MEAS10 & ", " & _
               record.MEAS11 & ", " & record.MEAS12 & ", " & record.MEAS13 & ", " & record.MEAS14 & ", " & record.MEAS15 & ", " & record.MEAS16 & ", " & _
               record.MEAS17 & ", " & record.MEAS18 & ", " & record.MEAS19 & ", " & record.MEAS20 & ", '" & _
               TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE , " & _
               record.OSFPOS1 & ", " & record.OSFWID1 & ", '" & record.OSFRD1 & "', " & _
               record.OSFPOS2 & ", " & record.OSFWID2 & ", '" & record.OSFRD2 & "', " & _
               record.OSFPOS3 & ", " & record.OSFWID3 & ", '" & record.OSFRD3 & "'" & " From TBCMJ005 "
    sqlWhere = "where (CRYNUM='" & Crynum & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
'yaz
Debug.Print sql
    
    ''SQLの実行
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001f_Exec = FUNCTION_RETURN_SUCCESS

PROC_EXIT:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume PROC_EXIT
End Function




'概要      :データ変換を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                ,説明
'          :tblLeft       ,IO   ,typ_TBCMJ005      ,テーブルデータ１
'          :tblRight      ,IO   ,typ_cmjc001f_Disp ,テーブルデータ２
'          :bFlg          ,I   ,Boolean           ,TRUE:引数１データ→引数２データへの変換  FALSE:引数１データ←引数２データへの変換
'説明      :
Public Sub ConvDate_F_cmjc001f_a(tblLeft As typ_TBCMJ005, tblRight As typ_cmjc001f_Disp, bFlg As Boolean)
    If bFlg = True Then
        With tblRight
            .POSITION = tblLeft.POSITION
            .SMPKBN = tblLeft.SMPKBN
            .TRANCOND = tblLeft.TRANCOND
            .SMPLNO = tblLeft.SMPLNO
            .SMPLUMU = tblLeft.SMPLUMU
            .HINBAN = tblLeft.HINBAN
            .REVNUM = tblLeft.REVNUM
            .FACTORY = tblLeft.FACTORY
            .OPECOND = tblLeft.OPECOND
            .KRPROCCD = tblLeft.KRPROCCD
            .PROCCODE = tblLeft.PROCCODE
            .GOUKI = tblLeft.GOUKI
            .MEASMETH = tblLeft.MEASMETH
            .MEASSPOT = tblLeft.MEASSPOT
            .MAG = tblLeft.MAG
            .HTPRC = tblLeft.HTPRC
            .KKSP = tblLeft.KKSP
            .KKSET = tblLeft.KKSET
            .CALCMAX = tblLeft.CALCMAX
            .CALCAVE = tblLeft.CALCAVE
            .MEAS1 = tblLeft.MEAS1
            .MEAS2 = tblLeft.MEAS2
            .MEAS3 = tblLeft.MEAS3
            .MEAS4 = tblLeft.MEAS4
            .MEAS5 = tblLeft.MEAS5
            .MEAS6 = tblLeft.MEAS6
            .MEAS7 = tblLeft.MEAS7
            .MEAS8 = tblLeft.MEAS8
            .MEAS9 = tblLeft.MEAS9
            .MEAS10 = tblLeft.MEAS10
            .MEAS11 = tblLeft.MEAS11
            .MEAS12 = tblLeft.MEAS12
            .MEAS13 = tblLeft.MEAS13
            .MEAS14 = tblLeft.MEAS14
            .MEAS15 = tblLeft.MEAS15
            .MEAS16 = tblLeft.MEAS16
            .MEAS17 = tblLeft.MEAS17
            .MEAS18 = tblLeft.MEAS18
            .MEAS19 = tblLeft.MEAS19
            .MEAS20 = tblLeft.MEAS20
'OSF，BMD項目追加対応  2002.04.02 yakimura
            .OSFPOS1 = tblLeft.OSFPOS1
            .OSFWID1 = tblLeft.OSFWID1
            .OSFRD1 = tblLeft.OSFRD1
            .OSFPOS2 = tblLeft.OSFPOS2
            .OSFWID2 = tblLeft.OSFWID2
            .OSFRD2 = tblLeft.OSFRD2
            .OSFPOS3 = tblLeft.OSFPOS3
            .OSFWID3 = tblLeft.OSFWID3
            .OSFRD3 = tblLeft.OSFRD3
'OSF，BMD項目追加対応  2002.04.02 yakimura
        End With
    Else
        With tblLeft
            .POSITION = tblRight.POSITION
            .SMPKBN = tblRight.SMPKBN
            .TRANCOND = tblRight.TRANCOND
            .SMPLNO = tblRight.SMPLNO
            .SMPLUMU = tblRight.SMPLUMU
            .HINBAN = tblRight.HINBAN
            .REVNUM = tblRight.REVNUM
            .FACTORY = tblRight.FACTORY
            .OPECOND = tblRight.OPECOND
            .KRPROCCD = tblRight.KRPROCCD
            .PROCCODE = tblRight.PROCCODE
            .GOUKI = tblRight.GOUKI
            .MEASMETH = tblRight.MEASMETH
            .MEASSPOT = tblRight.MEASSPOT
            .MAG = tblRight.MAG
            .HTPRC = tblRight.HTPRC
            .KKSP = tblRight.KKSP
            .KKSET = tblRight.KKSET
            .CALCMAX = tblRight.CALCMAX
            .CALCAVE = tblRight.CALCAVE
            .MEAS1 = tblRight.MEAS1
            .MEAS2 = tblRight.MEAS2
            .MEAS3 = tblRight.MEAS3
            .MEAS4 = tblRight.MEAS4
            .MEAS5 = tblRight.MEAS5
            .MEAS6 = tblRight.MEAS6
            .MEAS7 = tblRight.MEAS7
            .MEAS8 = tblRight.MEAS8
            .MEAS9 = tblRight.MEAS9
            .MEAS10 = tblRight.MEAS10
            .MEAS11 = tblRight.MEAS11
            .MEAS12 = tblRight.MEAS12
            .MEAS13 = tblRight.MEAS13
            .MEAS14 = tblRight.MEAS14
            .MEAS15 = tblRight.MEAS15
            .MEAS16 = tblRight.MEAS16
            .MEAS17 = tblRight.MEAS17
            .MEAS18 = tblRight.MEAS18
            .MEAS19 = tblRight.MEAS19
            .MEAS20 = tblRight.MEAS20
'OSF，BMD項目追加対応  2002.04.02 yakimura
            .OSFPOS1 = tblRight.OSFPOS1
            .OSFWID1 = tblRight.OSFWID1
            .OSFRD1 = tblRight.OSFRD1
            .OSFPOS2 = tblRight.OSFPOS2
            .OSFWID2 = tblRight.OSFWID2
            .OSFRD2 = tblRight.OSFRD2
            .OSFPOS3 = tblRight.OSFPOS3
            .OSFWID3 = tblRight.OSFWID3
            .OSFRD3 = tblRight.OSFRD3
'OSF，BMD項目追加対応  2002.04.02 yakimura
        End With
    End If

End Sub

