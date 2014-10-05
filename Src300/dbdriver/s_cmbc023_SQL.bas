Attribute VB_Name = "s_cmbc023_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DBアクセス関数
' 定義内容: TBCMJ002 (結晶抵抗実績)
' 参照　　: 060211_結晶検査
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmjc001d_Disp
  '  CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
  '  TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long  サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    GOUKI As String * 3             ' 号機
    TYPE As String * 1              ' タイプ
    MEAS1 As Double                 ' 測定値１
    MEAS2 As Double                 ' 測定値２
    MEAS3 As Double                 ' 測定値３
    MEAS4 As Double                 ' 測定値４
    MEAS5 As Double                 ' 測定値５
    EFEHS As Double                 ' 実効偏析
    RRG As Double                   ' ＲＲＧ
    JudgData As Double              ' 検索対象値
    KANSANCHI As String             '10Ω換算値　林
  '  TSTAFFID As String * 8          ' 登録社員ID
  '  REGDATE As Date                 ' 登録日付
  '  KSTAFFID As String * 8          ' 更新社員ID
  '  UPDDATE As Date                 ' 更新日付
  '  SENDFLAG As String * 1          ' 送信フラグ
  '  SENDDATE As Date                ' 送信日付



End Type

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMJ002」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_cmjc001d_Disp ,抽出レコード
'          :SPLNUMs()     ,I  ,Integer      ,抽出条件配列(サンプルNo)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/20作成　長野
Public Function DBDRV_Getcmjc001d_Disp(records() As typ_cmjc001d_Disp, SPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim sqlWhere As String  'SQL文(WHERE節)
Dim sqlGroup As String  'SQL文(GROUP節)
Dim sqlOrder As String  'SQL文(ORDER節)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    DBDRV_Getcmjc001d_Disp = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmjc001d_SQL.bas -- Function DBDRV_Getcmjc001d_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
              " OPECOND, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, JUDGDATA "
    sqlBase = sqlBase & "From TBCMJ002"
    ''抽出条件(ｻﾝﾌﾟﾙNO)の取り出し
    sqlWhere = "Where SMPLNO in ("
    For i = 1 To UBound(SPLNUMs)
        sqlWhere = sqlWhere & "'" & SPLNUMs(i) & "'"
        If i < UBound(SPLNUMs) Then
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
        DBDRV_Getcmjc001d_Disp = FUNCTION_RETURN_FAILURE
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
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .FACTORY = rs("FACTORY")         ' 工場
            .OPECOND = rs("OPECOND")         ' 操業条件
            .GOUKI = rs("GOUKI")             ' 号機
            .TYPE = rs("TYPE")               ' タイプ
            .MEAS1 = rs("MEAS1")             ' 測定値１
            .MEAS2 = rs("MEAS2")             ' 測定値２
            .MEAS3 = rs("MEAS3")             ' 測定値３
            .MEAS4 = rs("MEAS4")             ' 測定値４
            .MEAS5 = rs("MEAS5")             ' 測定値５
            .EFEHS = rs("EFEHS")             ' 実効偏析
            .RRG = rs("RRG")                 ' ＲＲＧ
            .JudgData = rs("JUDGDATA")       ' 検索対象値
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001d_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :引数で渡されたレコードをTBCMJ002に追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :record        ,I  ,typ_cmjc001d_Disp ,抽出レコード
'          :CRYNUM        ,I  ,String       ,結晶番号
'          :TSTAFFID      ,I  ,String       ,登録社員ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :処理回数はテーブル上の最大値+1とする。
'履歴      :2001/06/22(Fri)作成　長野

Public Function DBDRV_Getcmjc001d_Exec(record As typ_cmjc001d_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQLベース部分
Dim sqlWhere As String  'SQLWhere部分
Dim sqlGroup As String  'SQLGroup部分

'    CRYNUM             結晶番号 　⇒引数
'    TRANCNT         　 処理回数　 ⇒最大
'    TSTAFFID           登録社員ID ⇒引数
'    REGDATE 　　　     登録日付　 ⇒SYSDATE
'    KSTAFFID           更新社員ID ⇒" "
'    UPDDATE            更新日付　 ⇒SYSDATE
'    SENDFLAG           送信フラグ ⇒"0"
'    SENDDATE           送信日付　 ⇒SYSDATE

    DBDRV_Getcmjc001d_Exec = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmjc001d_SQL.bas -- Function DBDRV_Getcmjc001d_Exec"

    sqlBase = "Insert into TBCMJ002 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY, " & _
              "OPECOND, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.hinban & "', " & record.REVNUM & ", '" & _
               record.FACTORY & "', '" & record.OPECOND & "', '" & record.GOUKI & "', '" & record.TYPE & "', " & _
               record.MEAS1 & ", " & record.MEAS2 & ", " & record.MEAS3 & ", " & record.MEAS4 & ", " & record.MEAS5 & ", " & record.EFEHS & ", " & _
               record.RRG & ", " & record.JudgData & ", '" & TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ002 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
            
    ''SQLの実行
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001d_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :テーブル「XSDCS」の条件にあったレコードを更新する(推定ｻﾝﾌﾟﾙの実績ﾌﾗｸﾞ)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                       ,説明
'          :tblCrySmpMan  ,I   ,typ_XSDCS               ,新サンプル管理（ブロック）テーブル更新パラメータ
'          :strCryNum     ,I   ,String                  ,結晶番号
'          :iSmpNo        ,I   ,Long                    ,サンプルNo.    Integer→Long 6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O   ,Integer                 ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'履歴      :

Public Function UpdateTbl_CrySmpSuitei(tblCrySmpMan As typ_XSDCS, strCryNum As String, iSmpNo As Long) As Integer
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    
    UpdateTbl_CrySmpSuitei = FUNCTION_RETURN_FAILURE

    ' 推定ｻﾝﾌﾟﾙID1の更新
    With tblCrySmpMan
        sql = "SELECT CRYNUMCS FROM XSDCS "
        sql = sql & "WHERE XTALCS = '" & strCryNum & "' and "
        sql = sql & "      CRYSMPLIDRS1CS = " & iSmpNo
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount <> 0 Then
            
            sql = "update XSDCS set "
            sql = sql & "CRYRESRS1CS='" & .CRYRESRS1CS & "', "          ' 結晶検査実績（Rs-1)
            sql = sql & "KDAYCS=sysdate, "                              ' 更新日付
            sql = sql & "SNDKCS='0' "                                   ' 送信フラグ
            sql = sql & "WHERE XTALCS = '" & strCryNum & "' and "
            sql = sql & "      CRYSMPLIDRS1CS = " & iSmpNo
Debug.Print sql
            If OraDB.ExecuteSQL(sql) <= 0 Then
                rs.Close
                Exit Function
            End If
        End If
        rs.Close
    End With

    ' 推定ｻﾝﾌﾟﾙID2の更新 (注：更新内容は「CRYRESRS1CS」にのみ設定されている。)
    With tblCrySmpMan
        sql = "SELECT CRYNUMCS FROM XSDCS "
        sql = sql & "WHERE XTALCS = '" & strCryNum & "' and "
        sql = sql & "      CRYSMPLIDRS2CS = " & iSmpNo
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount <> 0 Then
        
            sql = "update XSDCS set "
            sql = sql & "CRYRESRS2CS='" & .CRYRESRS1CS & "', "          ' 結晶検査実績（Rs-2)
            sql = sql & "KDAYCS=sysdate, "                              ' 更新日付
            sql = sql & "SNDKCS='0' "                                   ' 送信フラグ
            sql = sql & "WHERE XTALCS = '" & strCryNum & "' and "
            sql = sql & "      CRYSMPLIDRS2CS = " & iSmpNo
Debug.Print sql
            If OraDB.ExecuteSQL(sql) <= 0 Then
                rs.Close
                Exit Function
            End If
        End If
        rs.Close
    End With


    UpdateTbl_CrySmpSuitei = FUNCTION_RETURN_SUCCESS

End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------
'概要      :TBCMJ002に登録データ(TRANCNT=0)となるデータが存在するか確認する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,I/O ,型             ,説明
'          :getCryNum     ,I   ,String         ,結晶番号
'
'          :戻り値        ,O   ,Integer        ,処理成功：FUNCTION_RETURN_SUCCESS
'                                              ,処理失敗：FUNCTION_RETURN_FAILURE
'説明      :TBCMJ002に登録データ(TRANCNT=0)となるデータが存在するか確認する
'履歴      :2011/08/09 Akizuki


Public Function CheckTRANCNT0_UMU(record As typ_cmjc001d_Disp, CRYNUM$, TSTAFFID$) As Boolean

    Dim sql         As String   '実行SQL
    Dim sqlBase     As String   'SQLベース部分
    Dim sqlWhere    As String   'SQLWhere部分
    
    Dim rs As OraDynaset    'RecordSet
    Dim cnt As Integer      '取得件数 保存用


    CheckTRANCNT0_UMU = False
    
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmec066_SQL.bas -- Function getSIRDInfo"

    sqlBase = "select CRYNUM from TBCMJ002" & vbCrLf
   
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and "
    sqlWhere = sqlWhere & "(POSITION=" & record.POSITION & ") and "
    sqlWhere = sqlWhere & "(SMPKBN='" & record.SMPKBN & "') and "
    sqlWhere = sqlWhere & "(TRANCOND='" & record.TRANCOND & "') and "
    sqlWhere = sqlWhere & "(TRANCNT = '0') "
    
    sql = sqlBase & sqlWhere

    '抵抗実績(TRANCNT=0)を取得する。
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
 
    '取得に失敗した場合
    If rs Is Nothing Then
        CheckTRANCNT0_UMU = False
        Exit Function
    End If

    'SXLID情報の件数を取得する。
    cnt = rs.RecordCount
    
    If cnt >= 1 Then
        CheckTRANCNT0_UMU = True
    Else
        CheckTRANCNT0_UMU = False
    End If
    
    Exit Function
    
PROC_ERR:
    'エラーハンドラ
    gErr.HandleError

End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------
'概要      :抵抗実績テーブルTBCMJ002(TRANCNT=0)を作成する

'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型                   ,説明
'          :record        ,I    ,typ_CMJC022i_Disp    ,抽出レコード
'
'          :戻り値        ,O   ,Integer                 ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :TARNCNT=MAX と同じデータをTRANCNT=0で作成する
'履歴      :2011/08/09 SUMCO Akizuki TRANCNT=0対応


 Public Function DBDRV_InsTBCMJ002_TRANCNT0_Exec(record As typ_cmjc001d_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
    Dim sql         As String   '実行SQL
    Dim sqlBase     As String   'SQLベース部分
    Dim sqlWhere    As String   'SQLWhere部分
        
    
    DBDRV_InsTBCMJ002_TRANCNT0_Exec = FUNCTION_RETURN_FAILURE


    ''SQLを組み立てる
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmbc023_SQL.bas - Function DBDRV_InsTBCMJ002_TRANCNT0_Exec"
    
    ''SQLを組み立てる
    sqlBase = "Insert into TBCMJ002 ("
    sqlBase = sqlBase & "CRYNUM,"           '結晶番号
    sqlBase = sqlBase & "POSITION,"         '位置
    sqlBase = sqlBase & "SMPKBN,"           'サンプル区分
    sqlBase = sqlBase & "TRANCOND,"         '処理条件
    sqlBase = sqlBase & "TRANCNT,"          '処理回数
    sqlBase = sqlBase & "SMPLNO,"           'サンプルNo
    sqlBase = sqlBase & "SMPLUMU,"          'サンプル有無
    sqlBase = sqlBase & "KRPROCCD,"         '管理工程コード
    sqlBase = sqlBase & "PROCCODE,"         '工程コード
    sqlBase = sqlBase & "HINBAN,"           '品番
    sqlBase = sqlBase & "REVNUM,"           '製品番号改訂番号
    sqlBase = sqlBase & "FACTORY,"          '工場
    sqlBase = sqlBase & "OPECOND," & vbLf   '操業条件
    sqlBase = sqlBase & "GOUKI,"            '号機
    sqlBase = sqlBase & "TYPE,"             'タイプ
    sqlBase = sqlBase & "MEAS1,"            '測定値1
    sqlBase = sqlBase & "MEAS2,"            '測定値2
    sqlBase = sqlBase & "MEAS3,"            '測定値3
    sqlBase = sqlBase & "MEAS4,"            '測定値4
    sqlBase = sqlBase & "MEAS5,"            '測定値5
    sqlBase = sqlBase & "EFEHS,"            '実効偏析
    sqlBase = sqlBase & "RRG,"              'RRG
    sqlBase = sqlBase & "JUDGDATA,"         '検索対象値
    sqlBase = sqlBase & "TSTAFFID,"         '登録社員ID
    sqlBase = sqlBase & "REGDATE,"          '登録日付
    sqlBase = sqlBase & "KSTAFFID,"         '更新社員ID
    sqlBase = sqlBase & "UPDDATE,"          '更新日付
    sqlBase = sqlBase & "SENDFLAG,"         '送信フラグ
    sqlBase = sqlBase & "SENDDATE)" & vbLf  '送信日付
    
    
    'Select SQLで対象データを取得してセット
    sqlBase = sqlBase & "VALUES(" & vbLf
    sqlBase = sqlBase & "'" & CRYNUM & "',"                     '結晶番号
    sqlBase = sqlBase & "'" & record.POSITION & "',"            '位置
    sqlBase = sqlBase & "'" & record.SMPKBN & "',"              'サンプル区分
    sqlBase = sqlBase & "'" & record.TRANCOND & "',"            '処理条件
    sqlBase = sqlBase & "'0'," & vbLf                           '処理回数　☆TRANCNT=0で作成
    sqlBase = sqlBase & "'" & record.SMPLNO & "',"              'サンプルNo
    sqlBase = sqlBase & "'" & record.SMPLUMU & "',"             'サンプル有無
    sqlBase = sqlBase & "'" & record.KRPROCCD & "',"            '管理工程コード
    sqlBase = sqlBase & "'" & record.PROCCODE & "',"            '工程コード
    sqlBase = sqlBase & "'" & record.hinban & "',"              '品番
    sqlBase = sqlBase & "'" & record.REVNUM & "',"              '製品番号改訂番号
    sqlBase = sqlBase & "'" & record.FACTORY & "',"             '工場
    sqlBase = sqlBase & "'" & record.OPECOND & "',"             '操業条件
    sqlBase = sqlBase & "'" & record.GOUKI & "'," & vbLf        '号機
    sqlBase = sqlBase & "'" & record.TYPE & "',"                'タイプ
    sqlBase = sqlBase & "'" & record.MEAS1 & "',"               '測定値1
    sqlBase = sqlBase & "'" & record.MEAS2 & "',"               '測定値2
    sqlBase = sqlBase & "'" & record.MEAS3 & "',"               '測定値3
    sqlBase = sqlBase & "'" & record.MEAS4 & "',"               '測定値4
    sqlBase = sqlBase & "'" & record.MEAS5 & "',"               '測定値5
    sqlBase = sqlBase & "'" & record.EFEHS & "',"               '実効偏析
    sqlBase = sqlBase & "'" & record.RRG & "',"                 'RRG
    sqlBase = sqlBase & "'" & record.JudgData & "'," & vbLf     '検索対象値
    sqlBase = sqlBase & "'" & TSTAFFID & "',"                   '登録社員ID
    sqlBase = sqlBase & "SYSDATE,"                              '登録日付
    sqlBase = sqlBase & "'" & TSTAFFID & "',"                   '更新社員ID
    sqlBase = sqlBase & "SYSDATE,"                              '更新日付
    sqlBase = sqlBase & "'0',"                                  '送信フラグ
    sqlBase = sqlBase & "SYSDATE)" & vbLf                       '送信日付
    
    sql = sqlBase & sqlWhere
    
    ''SQLの実行
    OraDB.ExecuteSQL (sql)

    DBDRV_InsTBCMJ002_TRANCNT0_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
    

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------
'概要      :抵抗実績テーブルTBCMJ002(TRANCNT=0)を更新する

'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型                   ,説明
'          :record        ,I    ,typ_CMJC022i_Disp    ,抽出レコード
'
'          :戻り値        ,O    ,Integer              ,処理成功：FUNCTION_RETURN_SUCCESS
'                                                     ,処理失敗：FUNCTION_RETURN_FAILURE
'説明      :TARNCNT=MAX と同じデータをTRANCNT=0に更新する
'履歴      :2011/08/09 SUMCO Akizuki TRANCNT=0対応


 Public Function DBDRV_UpdTBCMJ002_TRANCNT0_Exec(record As typ_cmjc001d_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
 
    Dim sql         As String   'SQL
    Dim sqlBase     As String   'SQLベース部分
    Dim sqlWhere    As String   'SQLWhere部分
    
    
    DBDRV_UpdTBCMJ002_TRANCNT0_Exec = FUNCTION_RETURN_FAILURE


    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmbc023_SQL.bas - Function DBDRV_UpdTBCMJ002_TRANCNT0_Exec"
  
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmjc001d_SQL.bas -- Function DBDRV_Getcmjc001d_Exec"
    
    
    ''SQLを組み立てる
    sqlBase = "Update TBCMJ002 set "
    sqlBase = sqlBase & "CRYNUM = '" & CRYNUM & "',"                    '結晶番号
    sqlBase = sqlBase & "POSITION = '" & record.POSITION & "',"         '位置
    sqlBase = sqlBase & "SMPKBN = '" & record.SMPKBN & "',"             'サンプル区分
    sqlBase = sqlBase & "TRANCOND = '" & record.TRANCOND & "',"         '処理条件
    sqlBase = sqlBase & "TRANCNT = 0,"                                  '処理回数 ☆TRANCNT=0で作成
    sqlBase = sqlBase & "SMPLNO = '" & record.SMPLNO & "',"             'サンプルNo
    sqlBase = sqlBase & "SMPLUMU = '" & record.SMPLUMU & "',"           'サンプル有無
    sqlBase = sqlBase & "KRPROCCD = '" & record.KRPROCCD & "',"         '管理工程コード
    sqlBase = sqlBase & "PROCCODE = '" & record.PROCCODE & "',"         '工程コード
    sqlBase = sqlBase & "HINBAN = '" & record.hinban & "',"             '品番
    sqlBase = sqlBase & "REVNUM = '" & record.REVNUM & "',"             '製品番号改訂番号
    sqlBase = sqlBase & "FACTORY = '" & record.FACTORY & "',"           '工場
    sqlBase = sqlBase & "OPECOND = '" & record.OPECOND & "',"           '操業条件
    sqlBase = sqlBase & "GOUKI = '" & record.GOUKI & "'," & vbLf        '号機
    sqlBase = sqlBase & "TYPE = '" & record.TYPE & "',"                 'タイプ
    sqlBase = sqlBase & "MEAS1 = '" & record.MEAS1 & "',"               '測定値1
    sqlBase = sqlBase & "MEAS2 = '" & record.MEAS2 & "',"               '測定値2
    sqlBase = sqlBase & "MEAS3 = '" & record.MEAS3 & "',"               '測定値3
    sqlBase = sqlBase & "MEAS4 = '" & record.MEAS4 & "',"               '測定値4
    sqlBase = sqlBase & "MEAS5 = '" & record.MEAS5 & "',"               '測定値5
    sqlBase = sqlBase & "EFEHS = '" & record.EFEHS & "',"               '実効偏析
    sqlBase = sqlBase & "RRG = '" & record.RRG & "',"                   'RRG
    sqlBase = sqlBase & "JUDGDATA = '" & record.JudgData & "'," & vbLf  '検索対象値
    sqlBase = sqlBase & "KSTAFFID = '" & TSTAFFID & "',"                '更新社員ID
    sqlBase = sqlBase & "UPDDATE = SYSDATE, "                           '更新日付
    sqlBase = sqlBase & "SENDFLAG = '0' " & vbLf                        '送信フラグ
    
    '更新対象となるデータの指定 (TRANCNT=0の同データ)
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and "
    sqlWhere = sqlWhere & "(POSITION=" & record.POSITION & ") and "
    sqlWhere = sqlWhere & "(SMPKBN='" & record.SMPKBN & "') and "
    sqlWhere = sqlWhere & "(TRANCOND=" & record.TRANCOND & ") and"
    sqlWhere = sqlWhere & "(TRANCNT = 0) "

    sql = sqlBase & sqlWhere
    
    ''SQLの実行
    OraDB.ExecuteSQL (sql)
    
    DBDRV_UpdTBCMJ002_TRANCNT0_Exec = FUNCTION_RETURN_SUCCESS


proc_exit:
    '終了

    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
