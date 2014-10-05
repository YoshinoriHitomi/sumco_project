Attribute VB_Name = "s_cmbc038_SQL"
Option Explicit

'ブロックラベル払出し  4/16 Yam作成

' ブロック一覧
Public Type typ_BlkLbl
    BLOCKID As String * 12      ' ブロックID
    HIN(5) As tFullHinban       ' 品番
    WFINDDATE As String * 10    ' 最終抜試日付
    CRYNUM As String * 12       ' 結晶番号
    INGOTPOS As Integer         ' インゴット内位置
    LENGTH As Integer           ' ブロック長さ
    REALLEN As Integer          ' ブロック実長さ
    HINLEN(5) As Integer        ' 品番長さ
    DIAMETER As Integer         ' 直径
    SBLOCKID As String * 12     ' 先頭ブロックID
    BLOCKORDER As Integer       ' ブロック順序
    HOLDCLS As String * 1       ' ホールド状態  --- 2001/09/19 kuramoto 追加 ---
    PASSFLAG As String * 1      ' 通過フラグ　　--- 200/04/16 Yam
End Type


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME040」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME040 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCME040(records() As typ_TBCME040, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, REALLEN, BLOCKID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS," & _
              " RSTATCLS, HOLDCLS, BDCAUS, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE," & _
              " PASSFLAG "   '02/07/05 hama
    
    sqlBase = sqlBase & "From TBCME040"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME040 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .LENGTH = rs("LENGTH")           ' 長さ
            .REALLEN = rs("REALLEN")         ' 実長さ
            .BLOCKID = rs("BLOCKID")         ' ブロックID
            .KRPROCCD = rs("KRPROCCD")       ' 現在管理工程
            .NOWPROC = rs("NOWPROC")         ' 現在工程
            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
            .DELCLS = rs("DELCLS")           ' 削除区分
            .LSTATCLS = rs("LSTATCLS")       ' 最終状態区分
            .RSTATCLS = rs("RSTATCLS")       ' 流動状態区分
            .HOLDCLS = rs("HOLDCLS")         ' ホールド区分
            .BDCAUS = rs("BDCAUS")           ' 不良理由
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
            .PASSFLAG = " "   ' 通過フラグのスペースクリア '02/07/05 hama
             If rs("PASSFLAG") = "1" Then
                .PASSFLAG = rs("PASSFLAG")   ' 通過フラグ '02/07/05 hama
            End If

        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME040 = FUNCTION_RETURN_SUCCESS
End Function

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME042」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME042 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCME042(records() As typ_TBCME042, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, SXLID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS, HOLDCLS," & _
              " HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, COUNT, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE, " & _
              " PASSFLAG "   '02/04/16 Yam
    sqlBase = sqlBase & "From TBCME042"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME042 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .LENGTH = rs("LENGTH")           ' 長さ
            .SXLID = rs("SXLID")             ' SXLID
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程
            .NOWPROC = rs("NOWPROC")         ' 現在工程
            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
            .DELCLS = rs("DELCLS")           ' 削除区分
            .LSTATCLS = rs("LSTATCLS")       ' 最終状態区分
            .HOLDCLS = rs("HOLDCLS")         ' ホールド区分
            .HINBAN = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .FACTORY = rs("FACTORY")         ' 工場
            .OPECOND = rs("OPECOND")         ' 操業条件
            .BDCAUS = rs("BDCAUS")           ' 不良理由
            .COUNT = rs("COUNT")             ' 枚数
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
            .PASSFLAG = " "   ' 通過フラグのスペースクリア '02/04/16 Yam
            If rs("PASSFLAG") = "1" Then
                .PASSFLAG = rs("PASSFLAG")   ' 通過フラグ '02/04/05 Yam
            End If
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME042 = FUNCTION_RETURN_SUCCESS
End Function


'概要      :HをNに改造
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sStaffID　　　,I  ,String         　,社員ID
'      　　:pBlkMap 　　　,I  ,typ_BlkLbl     　,ブロック一覧
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :2002/4/4 Yam  SXTLの通過FLGに１をたてる
Public Function DBDRV_s_cmbc038_Exec(ByVal sStaffID As String, pBlkMap() As typ_BlkLbl, sErrMsg As String) As FUNCTION_RETURN

    Dim sql As String
    Dim sDBName As String
    Dim recCnt As Long
    Dim iPos As Integer
    Dim i As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc038_SQL.bas -- Function DBDRV_s_cmbc038_Exec"
    sErrMsg = ""

    recCnt = UBound(pBlkMap)
    For i = 1 To recCnt
        With pBlkMap(i)
            '' SXL管理の更新
            sDBName = "E042"
            iPos = .INGOTPOS + .LENGTH
            sql = "update TBCME040 set "
            sql = sql & "PASSFLAG='1' "
'            sql = sql & "PASSFLAG='1', "
'            sql = sql & "UPDDATE=sysdate "
            sql = sql & " where CRYNUM='" & .CRYNUM & "'"
            sql = sql & " and INGOTPOS>=" & .INGOTPOS
            sql = sql & " and INGOTPOS<" & iPos
            If OraDB.ExecuteSQL(sql) < 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
                DBDRV_s_cmbc038_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End With
    Next i

    DBDRV_s_cmbc038_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:    '' 終了
    gErr.Pop
    Exit Function

proc_err: '' エラーハンドラ
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
    DBDRV_s_cmbc038_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


