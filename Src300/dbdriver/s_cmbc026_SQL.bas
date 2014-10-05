Attribute VB_Name = "s_cmbc026_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DBアクセス関数
' 定義内容: TBCMJ006 (ＧＤ実績)
' 参照　　: 060211_結晶検査
'================================================

'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmjc001g_Disp
'    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
'    TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    Factory As String * 1           ' 工場
    OpeCond As String * 1           ' 操業条件
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    GOUKI As String * 3             ' 号機
    MSRSDEN As Integer              ' 測定結果 Den
    MSRSLDL As Integer              ' 測定結果 L/DL
    MSRSDVD2 As Integer             ' 測定結果 DVD2
    MSLDL(1 To 15, 1 To 5) As Integer ' 測定値 LDL (測定位置, n番目)
    MSDEN(1 To 15, 1 To 5) As Integer ' 測定値 DEN (測定位置, n番目)
    MSDVD(1 To 5, 1) As Integer        ' 測定値 DVD (測定位置, n番目) 2002/7/4 tuku
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    MSZEROMN As Integer             ' L/DL0連続数最小値
    MSZEROMX As Integer             ' L/DL0連続数最大値
    PTNJUDGRES As String * 1        ' パターン判定結果
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
 '   TSTAFFID As String * 8          ' 登録社員ID
 '   REGDATE As Date                 ' 登録日付
 '   KSTAFFID As String * 8          ' 更新社員ID
 '   UPDDATE As Date                 ' 更新日付
 '   SENDFLAG As String * 1          ' 送信フラグ
 '   SENDDATE As Date                ' 送信日付
End Type



'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMJ006」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_cmjc001g_Disp ,抽出レコード
'          :SPLNUMs()     ,I  ,Integer      ,抽出条件配列(サンプルNo)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/20作成　長野
Public Function DBDRV_Getcmjc001g_Disp(records() As typ_cmjc001g_Disp, SMPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim sqlWhere As String  'SQLのWHERE部分
Dim sqlGroup As String  'SQLのGROUP部分
Dim sqlOrder As String  'SQLのOrder部分
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
Dim POS As Integer
Dim n As Integer

    DBDRV_Getcmjc001g_Disp = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001g_SQL.bas -- Function DBDRV_Getcmjc001g_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MSRSDEN, MSRSLDL, MSRSDVD2, "
    For POS = 1 To 15
        For n = 1 To 5
            sqlBase = sqlBase & "MS" & POS & "LDL" & n & ", "
        Next
        For n = 1 To 5
            sqlBase = sqlBase & "MS" & POS & "DEN" & n
            If POS = 15 And n = 5 Then
                Exit For
            Else
                sqlBase = sqlBase & ", "
            End If
        Next
        If POS = 15 Then
            sqlBase = sqlBase & " """
        End If
    Next
    sqlBase = sqlBase & "From TBCMJ006"
        
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
        DBDRV_Getcmjc001g_Disp = FUNCTION_RETURN_FAILURE
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
            .Factory = rs("FACTORY")         ' 工場
            .OpeCond = rs("OPECOND")         ' 操業条件
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .GOUKI = rs("GOUKI")             ' 号機
            .MSRSDEN = rs("MSRSDEN")         ' 測定結果 Den
            .MSRSLDL = rs("MSRSLDL")         ' 測定結果 L/DL
            .MSRSDVD2 = rs("MSRSDVD2")       ' 測定結果 DVD2
            
            For POS = 1 To 15
                For n = 1 To 5
                    .MSLDL(POS, n) = rs("MS" & Format$(POS, "00") & "LDL" & n) ' 測定値(pos) L/DL(n)
                    .MSDEN(POS, n) = rs("MS" & Format$(POS, "00") & "DEN" & n) ' 測定値(pos) Den(n)
                Next
            Next

        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001g_Disp = FUNCTION_RETURN_SUCCESS

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

'概要      :引数で渡されたレコードをTBCMJ006に追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :record        ,I  ,typ_cmjc001g_Disp ,抽出レコード
'          :CRYNUM        ,I  ,String       ,結晶番号
'          :TSTAFFID      ,I  ,String       ,登録社員ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :処理回数はテーブル上の最大値+1とする。
'履歴      :2001/06/22(Fri)作成　長野

Public Function DBDRV_Getcmjc001g_Exec(record As typ_cmjc001g_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN

Dim sql As String       'SQL全体
Dim sqlBase As String   'SQLベース部分
Dim sqlWhere As String  'SQLWhere部分
Dim sqlGroup As String  'SQLGroup部分
Dim POS As Integer
Dim n As Integer

'    CRYNUM             結晶番号　⇒引数
'    TRANCNT         　 処理回数　⇒最大
'   TSTAFFID            登録社員ID　⇒引数
 '   REGDATE 　　　     登録日付　⇒SYSDATE
 '   KSTAFFID           更新社員ID　⇒" "
 '   UPDDATE            更新日付　⇒SYSDATE
 '   SENDFLAG           送信フラグ　⇒"0"
 '   SENDDATE           送信日付　⇒SYSDATE
     
    DBDRV_Getcmjc001g_Exec = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001g_SQL.bas -- Function DBDRV_Getcmjc001g_Exec"

    sqlBase = "Insert into TBCMJ006 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MSRSDEN, MSRSLDL, MSRSDVD2, "
            For POS = 1 To 15
                For n = 1 To 5
                    sqlBase = sqlBase & "MS" & Format(POS, "00") & "LDL" & n & ", "
                Next
                For n = 1 To 5
                    sqlBase = sqlBase & "MS" & Format(POS, "00") & "DEN" & n & ", "
                Next
            Next
            For POS = 1 To 5
                sqlBase = sqlBase & "MS" & Format(POS, "00") & "DVD2" & ", "
            Next
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    sqlBase = sqlBase & "MSZEROMN, MSZEROMX, PTNJUDGRES, "
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    sqlBase = sqlBase & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE)"
    sqlBase = sqlBase & " select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.Factory & "', '" & _
               record.OpeCond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', " & record.MSRSDEN & ", " & _
               record.MSRSLDL & ", " & record.MSRSDVD2 & ", "
            For POS = 1 To 15
                For n = 1 To 5
                    sqlBase = sqlBase & record.MSLDL(POS, n) & ", "
                Next
                For n = 1 To 5
                    sqlBase = sqlBase & record.MSDEN(POS, n) & ", "
                Next
            Next
            
            For POS = 1 To 5 'DVD2直接入力カラム追加　2002/7/5 tuku
                    sqlBase = sqlBase & record.MSDVD(POS, 1) & ", "
            Next

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    sqlBase = sqlBase & record.MSZEROMN & "," & record.MSZEROMX & ",'" & record.PTNJUDGRES & "',"
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    sqlBase = sqlBase & "'" & TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ006 "
              
  
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
            
  ''SQLの実行
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001g_Exec = FUNCTION_RETURN_SUCCESS
    

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





'概要      :テーブル「TBCMJ006」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ006 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCMJ006(records() As typ_TBCMJ006, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MSRSDEN, MSRSLDL, MSRSDVD2, MS01LDL1, MS01LDL2, MS01LDL3, MS01LDL4, MS01LDL5, MS01DEN1, MS01DEN2," & _
              " MS01DEN3, MS01DEN4, MS01DEN5, MS02LDL1, MS02LDL2, MS02LDL3, MS02LDL4, MS02LDL5, MS02DEN1, MS02DEN2, MS02DEN3," & _
              " MS02DEN4, MS02DEN5, MS03LDL1, MS03LDL2, MS03LDL3, MS03LDL4, MS03LDL5, MS03DEN1, MS03DEN2, MS03DEN3, MS03DEN4," & _
              " MS03DEN5, MS04LDL1, MS04LDL2, MS04LDL3, MS04LDL4, MS04LDL5, MS04DEN1, MS04DEN2, MS04DEN3, MS04DEN4, MS04DEN5," & _
              " MS05LDL1, MS05LDL2, MS05LDL3, MS05LDL4, MS05LDL5, MS05DEN1, MS05DEN2, MS05DEN3, MS05DEN4, MS05DEN5, MS06LDL1," & _
              " MS06LDL2, MS06LDL3, MS06LDL4, MS06LDL5, MS06DEN1, MS06DEN2, MS06DEN3, MS06DEN4, MS06DEN5, MS07LDL1, MS07LDL2," & _
              " MS07LDL3, MS07LDL4, MS07LDL5, MS07DEN1, MS07DEN2, MS07DEN3, MS07DEN4, MS07DEN5, MS08LDL1, MS08LDL2, MS08LDL3," & _
              " MS08LDL4, MS08LDL5, MS08DEN1, MS08DEN2, MS08DEN3, MS08DEN4, MS08DEN5, MS09LDL1, MS09LDL2, MS09LDL3, MS09LDL4," & _
              " MS09LDL5, MS09DEN1, MS09DEN2, MS09DEN3, MS09DEN4, MS09DEN5, MS10LDL1, MS10LDL2, MS10LDL3, MS10LDL4, MS10LDL5," & _
              " MS10DEN1, MS10DEN2, MS10DEN3, MS10DEN4, MS10DEN5, MS11LDL1, MS11LDL2, MS11LDL3, MS11LDL4, MS11LDL5, MS11DEN1," & _
              " MS11DEN2, MS11DEN3, MS11DEN4, MS11DEN5, MS12LDL1, MS12LDL2, MS12LDL3, MS12LDL4, MS12LDL5, MS12DEN1, MS12DEN2," & _
              " MS12DEN3, MS12DEN4, MS12DEN5, MS13LDL1, MS13LDL2, MS13LDL3, MS13LDL4, MS13LDL5, MS13DEN1, MS13DEN2, MS13DEN3," & _
              " MS13DEN4, MS13DEN5, MS14LDL1, MS14LDL2, MS14LDL3, MS14LDL4, MS14LDL5, MS14DEN1, MS14DEN2, MS14DEN3, MS14DEN4," & _
              " MS14DEN5, MS15LDL1, MS15LDL2, MS15LDL3, MS15LDL4, MS15LDL5, MS15DEN1, MS15DEN2, MS15DEN3, MS15DEN4, MS15DEN5," & _
              " MS01DVD2,  MS02DVD2 , MS03DVD2 , MS04DVD2 , MS05DVD2 , TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    sqlBase = sqlBase & ",MSZEROMN, MSZEROMX, PTNJUDGRES "
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    sqlBase = sqlBase & "From TBCMJ006"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ006 = FUNCTION_RETURN_FAILURE
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
            .Factory = rs("FACTORY")         ' 工場
            .OpeCond = rs("OPECOND")         ' 操業条件
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .GOUKI = rs("GOUKI")             ' 号機
            .MSRSDEN = rs("MSRSDEN")         ' 測定結果 Den
            .MSRSLDL = rs("MSRSLDL")         ' 測定結果 L/DL
            .MSRSDVD2 = rs("MSRSDVD2")       ' 測定結果 DVD2
            .MS01LDL1 = rs("MS01LDL1")       ' 測定値01 L/DL1
            .MS01LDL2 = rs("MS01LDL2")       ' 測定値01 L/DL2
            .MS01LDL3 = rs("MS01LDL3")       ' 測定値01 L/DL3
            .MS01LDL4 = rs("MS01LDL4")       ' 測定値01 L/DL4
            .MS01LDL5 = rs("MS01LDL5")       ' 測定値01 L/DL5
            .MS01DEN1 = rs("MS01DEN1")       ' 測定値01 Den1
            .MS01DEN2 = rs("MS01DEN2")       ' 測定値01 Den2
            .MS01DEN3 = rs("MS01DEN3")       ' 測定値01 Den3
            .MS01DEN4 = rs("MS01DEN4")       ' 測定値01 Den4
            .MS01DEN5 = rs("MS01DEN5")       ' 測定値01 Den5
            .MS02LDL1 = rs("MS02LDL1")       ' 測定値02 L/DL1
            .MS02LDL2 = rs("MS02LDL2")       ' 測定値02 L/DL2
            .MS02LDL3 = rs("MS02LDL3")       ' 測定値02 L/DL3
            .MS02LDL4 = rs("MS02LDL4")       ' 測定値02 L/DL4
            .MS02LDL5 = rs("MS02LDL5")       ' 測定値02 L/DL5
            .MS02DEN1 = rs("MS02DEN1")       ' 測定値02 Den1
            .MS02DEN2 = rs("MS02DEN2")       ' 測定値02 Den2
            .MS02DEN3 = rs("MS02DEN3")       ' 測定値02 Den3
            .MS02DEN4 = rs("MS02DEN4")       ' 測定値02 Den4
            .MS02DEN5 = rs("MS02DEN5")       ' 測定値02 Den5
            .MS03LDL1 = rs("MS03LDL1")       ' 測定値03 L/DL1
            .MS03LDL2 = rs("MS03LDL2")       ' 測定値03 L/DL2
            .MS03LDL3 = rs("MS03LDL3")       ' 測定値03 L/DL3
            .MS03LDL4 = rs("MS03LDL4")       ' 測定値03 L/DL4
            .MS03LDL5 = rs("MS03LDL5")       ' 測定値03 L/DL5
            .MS03DEN1 = rs("MS03DEN1")       ' 測定値03 Den1
            .MS03DEN2 = rs("MS03DEN2")       ' 測定値03 Den2
            .MS03DEN3 = rs("MS03DEN3")       ' 測定値03 Den3
            .MS03DEN4 = rs("MS03DEN4")       ' 測定値03 Den4
            .MS03DEN5 = rs("MS03DEN5")       ' 測定値03 Den5
            .MS04LDL1 = rs("MS04LDL1")       ' 測定値04 L/DL1
            .MS04LDL2 = rs("MS04LDL2")       ' 測定値04 L/DL2
            .MS04LDL3 = rs("MS04LDL3")       ' 測定値04 L/DL3
            .MS04LDL4 = rs("MS04LDL4")       ' 測定値04 L/DL4
            .MS04LDL5 = rs("MS04LDL5")       ' 測定値04 L/DL5
            .MS04DEN1 = rs("MS04DEN1")       ' 測定値04 Den1
            .MS04DEN2 = rs("MS04DEN2")       ' 測定値04 Den2
            .MS04DEN3 = rs("MS04DEN3")       ' 測定値04 Den3
            .MS04DEN4 = rs("MS04DEN4")       ' 測定値04 Den4
            .MS04DEN5 = rs("MS04DEN5")       ' 測定値04 Den5
            .MS05LDL1 = rs("MS05LDL1")       ' 測定値05 L/DL1
            .MS05LDL2 = rs("MS05LDL2")       ' 測定値05 L/DL2
            .MS05LDL3 = rs("MS05LDL3")       ' 測定値05 L/DL3
            .MS05LDL4 = rs("MS05LDL4")       ' 測定値05 L/DL4
            .MS05LDL5 = rs("MS05LDL5")       ' 測定値05 L/DL5
            .MS05DEN1 = rs("MS05DEN1")       ' 測定値05 Den1
            .MS05DEN2 = rs("MS05DEN2")       ' 測定値05 Den2
            .MS05DEN3 = rs("MS05DEN3")       ' 測定値05 Den3
            .MS05DEN4 = rs("MS05DEN4")       ' 測定値05 Den4
            .MS05DEN5 = rs("MS05DEN5")       ' 測定値05 Den5
            .MS06LDL1 = rs("MS06LDL1")       ' 測定値06 L/DL1
            .MS06LDL2 = rs("MS06LDL2")       ' 測定値06 L/DL2
            .MS06LDL3 = rs("MS06LDL3")       ' 測定値06 L/DL3
            .MS06LDL4 = rs("MS06LDL4")       ' 測定値06 L/DL4
            .MS06LDL5 = rs("MS06LDL5")       ' 測定値06 L/DL5
            .MS06DEN1 = rs("MS06DEN1")       ' 測定値06 Den1
            .MS06DEN2 = rs("MS06DEN2")       ' 測定値06 Den2
            .MS06DEN3 = rs("MS06DEN3")       ' 測定値06 Den3
            .MS06DEN4 = rs("MS06DEN4")       ' 測定値06 Den4
            .MS06DEN5 = rs("MS06DEN5")       ' 測定値06 Den5
            .MS07LDL1 = rs("MS07LDL1")       ' 測定値07 L/DL1
            .MS07LDL2 = rs("MS07LDL2")       ' 測定値07 L/DL2
            .MS07LDL3 = rs("MS07LDL3")       ' 測定値07 L/DL3
            .MS07LDL4 = rs("MS07LDL4")       ' 測定値07 L/DL4
            .MS07LDL5 = rs("MS07LDL5")       ' 測定値07 L/DL5
            .MS07DEN1 = rs("MS07DEN1")       ' 測定値07 Den1
            .MS07DEN2 = rs("MS07DEN2")       ' 測定値07 Den2
            .MS07DEN3 = rs("MS07DEN3")       ' 測定値07 Den3
            .MS07DEN4 = rs("MS07DEN4")       ' 測定値07 Den4
            .MS07DEN5 = rs("MS07DEN5")       ' 測定値07 Den5
            .MS08LDL1 = rs("MS08LDL1")       ' 測定値08 L/DL1
            .MS08LDL2 = rs("MS08LDL2")       ' 測定値08 L/DL2
            .MS08LDL3 = rs("MS08LDL3")       ' 測定値08 L/DL3
            .MS08LDL4 = rs("MS08LDL4")       ' 測定値08 L/DL4
            .MS08LDL5 = rs("MS08LDL5")       ' 測定値08 L/DL5
            .MS08DEN1 = rs("MS08DEN1")       ' 測定値08 Den1
            .MS08DEN2 = rs("MS08DEN2")       ' 測定値08 Den2
            .MS08DEN3 = rs("MS08DEN3")       ' 測定値08 Den3
            .MS08DEN4 = rs("MS08DEN4")       ' 測定値08 Den4
            .MS08DEN5 = rs("MS08DEN5")       ' 測定値08 Den5
            .MS09LDL1 = rs("MS09LDL1")       ' 測定値09 L/DL1
            .MS09LDL2 = rs("MS09LDL2")       ' 測定値09 L/DL2
            .MS09LDL3 = rs("MS09LDL3")       ' 測定値09 L/DL3
            .MS09LDL4 = rs("MS09LDL4")       ' 測定値09 L/DL4
            .MS09LDL5 = rs("MS09LDL5")       ' 測定値09 L/DL5
            .MS09DEN1 = rs("MS09DEN1")       ' 測定値09 Den1
            .MS09DEN2 = rs("MS09DEN2")       ' 測定値09 Den2
            .MS09DEN3 = rs("MS09DEN3")       ' 測定値09 Den3
            .MS09DEN4 = rs("MS09DEN4")       ' 測定値09 Den4
            .MS09DEN5 = rs("MS09DEN5")       ' 測定値09 Den5
            .MS10LDL1 = rs("MS10LDL1")       ' 測定値10 L/DL1
            .MS10LDL2 = rs("MS10LDL2")       ' 測定値10 L/DL2
            .MS10LDL3 = rs("MS10LDL3")       ' 測定値10 L/DL3
            .MS10LDL4 = rs("MS10LDL4")       ' 測定値10 L/DL4
            .MS10LDL5 = rs("MS10LDL5")       ' 測定値10 L/DL5
            .MS10DEN1 = rs("MS10DEN1")       ' 測定値10 Den1
            .MS10DEN2 = rs("MS10DEN2")       ' 測定値10 Den2
            .MS10DEN3 = rs("MS10DEN3")       ' 測定値10 Den3
            .MS10DEN4 = rs("MS10DEN4")       ' 測定値10 Den4
            .MS10DEN5 = rs("MS10DEN5")       ' 測定値10 Den5
            .MS11LDL1 = rs("MS11LDL1")       ' 測定値11 L/DL1
            .MS11LDL2 = rs("MS11LDL2")       ' 測定値11 L/DL2
            .MS11LDL3 = rs("MS11LDL3")       ' 測定値11 L/DL3
            .MS11LDL4 = rs("MS11LDL4")       ' 測定値11 L/DL4
            .MS11LDL5 = rs("MS11LDL5")       ' 測定値11 L/DL5
            .MS11DEN1 = rs("MS11DEN1")       ' 測定値11 Den1
            .MS11DEN2 = rs("MS11DEN2")       ' 測定値11 Den2
            .MS11DEN3 = rs("MS11DEN3")       ' 測定値11 Den3
            .MS11DEN4 = rs("MS11DEN4")       ' 測定値11 Den4
            .MS11DEN5 = rs("MS11DEN5")       ' 測定値11 Den5
            .MS12LDL1 = rs("MS12LDL1")       ' 測定値12 L/DL1
            .MS12LDL2 = rs("MS12LDL2")       ' 測定値12 L/DL2
            .MS12LDL3 = rs("MS12LDL3")       ' 測定値12 L/DL3
            .MS12LDL4 = rs("MS12LDL4")       ' 測定値12 L/DL4
            .MS12LDL5 = rs("MS12LDL5")       ' 測定値12 L/DL5
            .MS12DEN1 = rs("MS12DEN1")       ' 測定値12 Den1
            .MS12DEN2 = rs("MS12DEN2")       ' 測定値12 Den2
            .MS12DEN3 = rs("MS12DEN3")       ' 測定値12 Den3
            .MS12DEN4 = rs("MS12DEN4")       ' 測定値12 Den4
            .MS12DEN5 = rs("MS12DEN5")       ' 測定値12 Den5
            .MS13LDL1 = rs("MS13LDL1")       ' 測定値13 L/DL1
            .MS13LDL2 = rs("MS13LDL2")       ' 測定値13 L/DL2
            .MS13LDL3 = rs("MS13LDL3")       ' 測定値13 L/DL3
            .MS13LDL4 = rs("MS13LDL4")       ' 測定値13 L/DL4
            .MS13LDL5 = rs("MS13LDL5")       ' 測定値13 L/DL5
            .MS13DEN1 = rs("MS13DEN1")       ' 測定値13 Den1
            .MS13DEN2 = rs("MS13DEN2")       ' 測定値13 Den2
            .MS13DEN3 = rs("MS13DEN3")       ' 測定値13 Den3
            .MS13DEN4 = rs("MS13DEN4")       ' 測定値13 Den4
            .MS13DEN5 = rs("MS13DEN5")       ' 測定値13 Den5
            .MS14LDL1 = rs("MS14LDL1")       ' 測定値14 L/DL1
            .MS14LDL2 = rs("MS14LDL2")       ' 測定値14 L/DL2
            .MS14LDL3 = rs("MS14LDL3")       ' 測定値14 L/DL3
            .MS14LDL4 = rs("MS14LDL4")       ' 測定値14 L/DL4
            .MS14LDL5 = rs("MS14LDL5")       ' 測定値14 L/DL5
            .MS14DEN1 = rs("MS14DEN1")       ' 測定値14 Den1
            .MS14DEN2 = rs("MS14DEN2")       ' 測定値14 Den2
            .MS14DEN3 = rs("MS14DEN3")       ' 測定値14 Den3
            .MS14DEN4 = rs("MS14DEN4")       ' 測定値14 Den4
            .MS14DEN5 = rs("MS14DEN5")       ' 測定値14 Den5
            .MS15LDL1 = rs("MS15LDL1")       ' 測定値15 L/DL1
            .MS15LDL2 = rs("MS15LDL2")       ' 測定値15 L/DL2
            .MS15LDL3 = rs("MS15LDL3")       ' 測定値15 L/DL3
            .MS15LDL4 = rs("MS15LDL4")       ' 測定値15 L/DL4
            .MS15LDL5 = rs("MS15LDL5")       ' 測定値15 L/DL5
            .MS15DEN1 = rs("MS15DEN1")       ' 測定値15 Den1
            .MS15DEN2 = rs("MS15DEN2")       ' 測定値15 Den2
            .MS15DEN3 = rs("MS15DEN3")       ' 測定値15 Den3
            .MS15DEN4 = rs("MS15DEN4")       ' 測定値15 Den4
            .MS15DEN5 = rs("MS15DEN5")       ' 測定値15 Den5
            'NULL チェック
            If IsNull(rs("MS01DVD2")) = False Then
                .MS01DVD2 = rs("MS01DVD2")       ' 測定値01 DVD2   2002/7/02 tuku
            Else
                .MS01DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS02DVD2")) = False Then
                .MS02DVD2 = rs("MS02DVD2")       ' 測定値01 DVD2   2002/7/02 tuku
            Else
                .MS02DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS03DVD2")) = False Then
                .MS03DVD2 = rs("MS03DVD2")       ' 測定値01 DVD2   2002/7/02 tuku
            Else
                .MS03DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS04DVD2")) = False Then
                .MS04DVD2 = rs("MS04DVD2")       ' 測定値01 DVD2   2002/7/02 tuku
            Else
                .MS04DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS05DVD2")) = False Then
                .MS05DVD2 = rs("MS05DVD2")       ' 測定値01 DVD2   2002/7/02 tuku
            Else
                .MS05DVD2 = DEF_PARAM_VALUE
            End If
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
            If IsNull(rs("MSZEROMN")) = False Then
                .MSZEROMN = rs("MSZEROMN")
            Else
                .MSZEROMN = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MSZEROMX")) = False Then
                .MSZEROMX = rs("MSZEROMX")
            Else
                .MSZEROMX = DEF_PARAM_VALUE
            End If
            If IsNull(rs("PTNJUDGRES")) = False Then
                .PTNJUDGRES = rs("PTNJUDGRES")
            Else
                .PTNJUDGRES = " "
            End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMJ006 = FUNCTION_RETURN_SUCCESS
End Function

