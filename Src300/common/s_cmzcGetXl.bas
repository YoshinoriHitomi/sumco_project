Attribute VB_Name = "s_cmzcGetXl"
Option Explicit


Public Function GetXl(CRYNUM$, FormName$) As c_cmzcXl
Dim sqlWhere$
Dim sqlWherePlan$
Dim sqlWhereBlk$
Dim Xl As c_cmzcXl
Dim RET As FUNCTION_RETURN

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetXl"

    ''共通のWHERE句を作る
    sqlWhere = " Where (CRYNUM='" & CRYNUM & "')"
    'リチャージ連番対応（指示No９桁変更） Y.K 2004/09.03 ７桁＋９桁目
    sqlWherePlan = " Where (SUBSTR(CRYNUM,1,7)='" & Mid(CRYNUM, 1, 7) & "') and (SUBSTR(CRYNUM,9,1)='" & Mid(CRYNUM, 9, 1) & "')"
'    sqlWherePlan = " Where (SUBSTR(CRYNUM,1,7)='" & Mid(CRYNUM, 1, 7) & "') "
    
    sqlWhereBlk = " Where (CRYNUM='" & CRYNUM & "') and (INGOTPOS>=0)"
    
    If FormName = "f_cmbc009_3" Then
        Set Xl = New c_cmzcXl
    Else
        ''結晶情報を取得する
        RET = GetTBCME037(Xl, sqlWhere)
        If RET = FUNCTION_RETURN_FAILURE Then
            ''結晶情報の読込に失敗した
            Set GetXl = Nothing
            GoTo proc_exit
        End If
    End If
        
    ''画面名により、各データを読み込む
    Select Case FormName
      Case "f_cmbc009_3"     ' ブロック組合せ
            RET = GetTBCME038(Xl.BlkPlans, sqlWherePlan)
            RET = GetTBCME039(Xl.HinPlans, sqlWherePlan)
      Case "f_cmbc016_1"     ' 加工払出し
            RET = GetTBCME039(Xl.HinPlans, sqlWherePlan)
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            RET = GetTBCME041(Xl.Hins, sqlWhere)
            RET = GetTBCME045(Xl.Cuts, CRYNUM)
      Case "f_cmbc018_2"     ' 切断
            RET = GetTBCME039(Xl.HinPlans, sqlWherePlan)
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            RET = GetTBCME041(Xl.Hins, sqlWhere)
            RET = GetTBCME045(Xl.Cuts, CRYNUM)
'      Case "f_cmbc030_1"     ' 待ち一覧
'            RET = GetTBCME040(xl.Blks, sqlWhereBlk)
'            RET = GetTBCME041(xl.Hins, sqlWhere)
'            ret = GetTBCME044(xl.WfSmps, sqlWhere)
      Case "f_cmbc032_1"     ' 待ち一覧
''↓更新 START SPT用実績作成方法変更 2006/06/30 SMP-OKAMOTO
            RET = GetBlockData_2(Xl.Blks, CRYNUM)
'            RET = GetTBCME040(xl.Blks, sqlWhereBlk)
''↑更新 END   SPT用実績作成方法変更 2006/06/30 SMP-OKAMOTO
            RET = GetTBCME041(Xl.Hins, sqlWhere)
      Case "f_cmbc033_1"     ' 待ち一覧
''↓更新 START SPT用実績作成方法変更 2006/06/30 SMP-OKAMOTO
            RET = GetBlockData_2(Xl.Blks, CRYNUM)
'            RET = GetTBCME040(xl.Blks, sqlWhereBlk)
''↑更新 END   SPT用実績作成方法変更 2006/06/30 SMP-OKAMOTO
            RET = GetTBCME041(Xl.Hins, sqlWhere)

      Case "f_cmbc030_1"     ' 待ち一覧
'            RET = GetTBCME040(xl.Blks, sqlWhereBlk)
''↓更新 START SPT用実績作成方法変更 2006/05/12 SMP-OKAMOTO
            RET = GetBlockData_2(Xl.Blks, CRYNUM)
'            RET = GetBlockData(xl.Blks, CRYNUM)     'ﾌﾞﾛｯｸ管理(TBCME040)参照停止　05/10/05 ooba
''↑更新 END   SPT用実績作成方法変更 2006/05/12 SMP-OKAMOTO
            RET = GetTBCME041(Xl.Hins, sqlWhere)
'            ret = GetTBCME044(xl.WfSmps, sqlWhere)
      Case "f_cmbc031_1"   ' （2002/07　未使用　→), "f_cmkc001e"    ' 再切断指示
'            RET = GetTBCME040(xl.Blks, sqlWhereBlk)
''↓更新 START SPT用実績作成方法変更 2006/05/12 SMP-OKAMOTO
            RET = GetBlockData_2(Xl.Blks, CRYNUM)
'            RET = GetBlockData(xl.Blks, CRYNUM)     'ﾌﾞﾛｯｸ管理(TBCME040)参照停止　05/10/05 ooba
''↑更新 END   SPT用実績作成方法変更 2006/05/12 SMP-OKAMOTO
            RET = GetTBCME041(Xl.Hins, sqlWhere)
'            ret = GetTBCME045(xl.cuts, crynum)
      Case "f_cmbc033_2"     ' 抜試指示
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            RET = GetTBCME041(Xl.Hins, sqlWhere)
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
            sqlWhere = " Where (a.XTALCB='" & CRYNUM & "')"
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
            RET = GetTBCME042(Xl.Sxls, sqlWhere)
            sqlWhere = " Where (XTALCW='" & CRYNUM & "')"
            RET = GetTBCME044(Xl.WfSmps, sqlWhere)
      Case "f_cmbc035_1"     ' 結晶情報変更
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
'            RET = GetTBCME042(xl.Sxls, sqlWhere)
'            RET = GetTBCME044(xl.WfSmps, sqlWhere)
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
            RET = GetTBCME044(Xl.WfSmps, sqlWhere)
            sqlWhere = " Where (a.XTALCB='" & CRYNUM & "')"
            RET = GetTBCME042(Xl.Sxls, sqlWhere)
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本

      Case "f_cmbc036_2"     ' 抜試指示変更
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
'            RET = GetTBCME042(xl.Sxls, sqlWhere)
'            RET = GetTBCME044(xl.WfSmps, sqlWhere)
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
            RET = GetTBCME044(Xl.WfSmps, sqlWhere)
            sqlWhere = " Where (a.XTALCB='" & CRYNUM & "')"
            RET = GetTBCME042(Xl.Sxls, sqlWhere)
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本

            RET = GetReject(Xl.Rejs, CRYNUM)
      Case "f_cmbc039_3"     ' 再抜試
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            RET = GetTBCME041(Xl.Hins, sqlWhere)
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
            sqlWhere = " Where (a.XTALCB='" & CRYNUM & "')"
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
            RET = GetTBCME042(Xl.Sxls, sqlWhere)
            sqlWhere = " where XTALCW = '" & CRYNUM & "' "
            RET = GetTBCME044(Xl.WfSmps, sqlWhere)
            RET = GetReject(Xl.Rejs, CRYNUM)
      Case "block"          ' ブロック情報のみ
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
      Case "hinban"          ' 品番情報のみ
            RET = GetTBCME041(Xl.Hins, sqlWhere)
      Case "All"            ' デバッグ用 全情報
            RET = GetTBCME038(Xl.BlkPlans, sqlWherePlan)
            RET = GetTBCME039(Xl.HinPlans, sqlWherePlan)
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            RET = GetTBCME041(Xl.Hins, sqlWhere)
            
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
'            RET = GetTBCME042(xl.Sxls, sqlWhere)
'            RET = GetTBCME043(xl.XlSmps, sqlWhere)
'            RET = GetTBCME044(xl.WfSmps, sqlWhere)
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
            RET = GetTBCME043(Xl.XlSmps, sqlWhere)
            RET = GetTBCME044(Xl.WfSmps, sqlWhere)
            sqlWhere = " Where (a.XTALCB='" & CRYNUM & "')"
            RET = GetTBCME042(Xl.Sxls, sqlWhere)
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本

            RET = GetTBCME045(Xl.Cuts, CRYNUM)
            RET = GetReject(Xl.Rejs, CRYNUM)
      Case Else
            Debug.Print "GetXl() : FormName が想定外"
            Set Xl = Nothing
    End Select
    
    Set GetXl = Xl

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :テーブル「TBCME037」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME037 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/29作成　野村
Public Function GetTBCME037(Xl As c_cmzcXl, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME037"

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME037"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME037 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    If rs.RecordCount > 0 Then
        Set Xl = New c_cmzcXl
        With Xl
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .DELCLS = rs("DELCLS")           ' 削除区分
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCD = rs("PROCCD")           ' 工程コード
            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
            .RPHINBAN = rs("RPHINBAN")       ' ねらい品番
            .RPREVNUM = rs("RPREVNUM")       ' ねらい品番製品番号改訂番号
            .RPFACT = rs("RPFACT")           ' ねらい品番工場
            .RPOPCOND = rs("RPOPCOND")       ' ねらい品番操業条件
            .PRODCOND = rs("PRODCOND")       ' 製作条件
            .PGID = rs("PGID")               ' ＰＧ－ＩＤ
            .UPLENGTH = rs("UPLENGTH")       ' 引上げ長さ
            .TOPLENG = rs("TOPLENG")         ' ＴＯＰ長さ
            .BODYLENG = rs("BODYLENG")       ' 直胴長さ
            .BOTLENG = rs("BOTLENG")         ' ＢＯＴ長さ
            .FREELENG = rs("FREELENG")       ' フリー長
            .DIAMETER = rs("DIAMETER")       ' 直径
            .CHARGE = rs("CHARGE")           ' チャージ量
            .SEED = rs("SEED")               ' シード
            .ADDDPCLS = rs("ADDDPCLS")       ' 追加ドープ種類
            .ADDDPPOS = rs("ADDDPPOS")       ' 追加ドープ位置
            .ADDDPVAL = rs("ADDDPVAL")       ' 追加ドープ量
        End With
    Else
        GetTBCME037 = FUNCTION_RETURN_FAILURE
        Set Xl = Nothing
        GoTo proc_exit
    End If
    rs.Close

    GetTBCME037 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :テーブル「TBCME038」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :col           ,O  ,c_cmczBlkPlans,ブロック設計コレクション
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/07/01作成　野村
Private Function GetTBCME038(col As c_cmzcBlkPlans, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
Dim target As c_cmzcBlkPlan

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME038"

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, USECLASS, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME038"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME038 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcBlkPlan
        With target
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .Length = rs("LENGTH")           ' 長さ
            .USECLASS = rs("USECLASS")       ' 使用区分
        End With
        col.Add target
        rs.MoveNext
    Next
    rs.Close

    GetTBCME038 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :テーブル「TBCME039」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME039 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/29作成　野村
Private Function GetTBCME039(col As c_cmzcHinPlans, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
Dim target As c_cmzcHinPlan

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME039"

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACT, OPCOND, LENGTH, USECLASS, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME039"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME039 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcHinPlan
        With target
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 改訂番号
            .FACT = rs("FACT")               ' 工場
            .OPCOND = rs("OPCOND")           ' 操業条件
            .Length = rs("LENGTH")           ' 長さ
            .USECLASS = rs("USECLASS")       ' 使用区分
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME039 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :テーブル「TBCME040」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME040 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/29作成　野村
Private Function GetTBCME040(col As c_cmzcBlks, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
Dim target As c_cmzcBlk

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME040"

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, REALLEN, BLOCKID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS," & _
              " RSTATCLS, HOLDCLS, BDCAUS, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME040"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere & " and (LENGTH>0)"
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME040 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcBlk
        With target
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .Length = rs("LENGTH")           ' 長さ
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
        End With
        col.Add target
        rs.MoveNext
    Next
    rs.Close

    GetTBCME040 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :テーブル「XSDC2」「XSDCS」「XSDC4」からブロック情報を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型                ,説明
'          :col           ,O   ,c_cmzcBlks        ,抽出レコード
'          :CRYNUM        ,I   ,String            ,結晶番号
'          :戻り値        ,O   ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :2005/10/05 ooba
Private Function GetBlockData(col As c_cmzcBlks, Optional CRYNUM$) As FUNCTION_RETURN

    Dim sql As String       'SQL全体
    Dim sql2 As String      'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim rs2 As OraDynaset   'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long
    Dim target As c_cmzcBlk


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetBlockData"

    ''SQLを組み立てる
    sql = "select "
    sql = sql & "CSTOP.XTALCS, "                                            '結晶番号
    sql = sql & "CSTOP.INPOSCS, "                                           '結晶内開始位置
    sql = sql & "CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, "                 '長さ
    sql = sql & "nvl(GNLC2,CSBOT.INPOSCS - CSTOP.INPOSCS) as REALLEN, "     '実長さ
    sql = sql & "CSTOP.CRYNUMCS, "                                          'ブロックID
    sql = sql & "nvl(GNKKNTC2,' ') as KRPROCCD, "                           '現在管理工程
    sql = sql & "nvl(GNWKNTC2,' ') as NOWPROC, "                            '現在工程
    sql = sql & "nvl(NEKKNTC2,' ') as LPKRPROCCD, "                         '最終通過管理工程
    sql = sql & "nvl(NEWKNTC2,' ') as LASTPASS, "                           '最終通過工程
    sql = sql & "nvl(SAKJC2,'0') as DELCLS, "                               '削除区分
    sql = sql & "nvl(LSTATBC2,'T') as LSTATCLS, "                           '最終状態区分
    sql = sql & "nvl(RSTATBC2,'T') as RSTATCLS, "                           '流動状態区分
    sql = sql & "nvl(HOLDBC2,'0') as HOLDCLS, "                             'ホールド区分
    sql = sql & "BDCAUSC2 as BDCAUS, "                                      '不良理由
    sql = sql & "C4.KNKTC4, "                                               '最終通過管理工程(XSDC4)
    sql = sql & "C4.WKKTC4, "                                               '最終通過工程(XSDC4)
    sql = sql & "C4.FCODEC4 "                                               '不良理由(XSDC4)
    sql = sql & "from XSDC2, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'T' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSTOP, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'B' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSBOT, "
    sql = sql & "     (select "
    sql = sql & "      XTALC4, "
    sql = sql & "      INPOSC4, "
    sql = sql & "      KNKTC4, "
    sql = sql & "      WKKTC4, "
    sql = sql & "      FCODEC4 "
    sql = sql & "      from XSDC4 TMP4 "
    sql = sql & "      where "
    sql = sql & "      XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "      and (KCKNTC4, KDAYC4) = ("
    sql = sql & "                     select MAX(KCKNTC4), MAX(KDAYC4) "
    sql = sql & "                     from XSDC4 "
    sql = sql & "                     where XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "                     and INPOSC4 = TMP4.INPOSC4) "
    sql = sql & "     ) C4 "
    sql = sql & "where "
    sql = sql & "CSTOP.CRYNUMCS = CRYNUMC2(+) "
    sql = sql & "and CSTOP.CRYNUMCS = CSBOT.CRYNUMCS "
    sql = sql & "and CSTOP.INPOSCS = C4.INPOSC4(+) "
    sql = sql & "and (LIVKC2 is null or LIVKC2 = '0' "
    sql = sql & "     or LSTATBC2 in ('R', 'H', 'B') or KANKC2 = '2') "
    sql = sql & "order by CSTOP.INPOSCS "
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetBlockData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcBlk
        With target
            .CRYNUM = rs("XTALCS")              ' 結晶番号
            .INGOTPOS = rs("INPOSCS")           ' 結晶内開始位置
            .Length = rs("LENGTH")              ' 長さ
            .REALLEN = rs("REALLEN")            ' 実長さ
            .BLOCKID = rs("CRYNUMCS")           ' ブロックID
            .KRPROCCD = rs("KRPROCCD")          ' 現在管理工程
            .NOWPROC = rs("NOWPROC")            ' 現在工程
            .LPKRPROCCD = rs("LPKRPROCCD")      ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")          ' 最終通過工程
            .DELCLS = rs("DELCLS")              ' 削除区分
            .LSTATCLS = rs("LSTATCLS")          ' 最終状態区分
            .RSTATCLS = rs("RSTATCLS")          ' 流動状態区分
            .HOLDCLS = rs("HOLDCLS")            ' ホールド区分
            If InStr(.BLOCKID, "$") <> 0 Then
                .KRPROCCD = MGPRCD_RIMERUTO_UKEIRE          ' 現在管理工程
                .NOWPROC = PROCD_RIMERUTO_UKEIRE            ' 現在工程
                .RSTATCLS = "M"                             ' 流動状態区分
                ' 最終通過管理工程
                If IsNull(rs("KNKTC4")) Then .LPKRPROCCD = "" Else .LPKRPROCCD = rs("KNKTC4")
                ' 最終通過工程
                If IsNull(rs("WKKTC4")) Then .LASTPASS = "" Else .LASTPASS = rs("WKKTC4")
                ' 不良理由
                If IsNull(rs("FCODEC4")) Then .BDCAUS = "0" Else .BDCAUS = rs("FCODEC4")
            Else
                ' 不良理由
                If IsNull(rs("BDCAUS")) Then .BDCAUS = "0" Else .BDCAUS = rs("BDCAUS")
            End If
            If Trim(.NOWPROC) = "" Then .DELCLS = "1"
            
'''            '不良理由をXSDC4から取得
'''            If InStr(.BLOCKID, "$") <> 0 Then
'''                .KRPROCCD = MGPRCD_RIMERUTO_UKEIRE      ' 現在管理工程
'''                .NOWPROC = PROCD_RIMERUTO_UKEIRE        ' 現在工程
'''                .RSTATCLS = "M"                         ' 流動状態区分
'''
'''                sql2 = "select KNKTC4, WKKTC4, FCODEC4 from XSDC4 "
''''                sql2 = sql2 & "where substr(XTALC4, 1 ,10) = '" & Mid(.BLOCKID, 1, 10) & "' "
'''                sql2 = sql2 & "where XTALC4 like '" & Mid(.BLOCKID, 1, 10) & "%' "  '05/12/26
'''                sql2 = sql2 & "and INPOSC4 = " & .INGOTPOS & " "
'''                sql2 = sql2 & "order by KCKNTC4 desc "
'''
'''                Set rs2 = OraDB.DBCreateDynaset(sql2, ORADYN_NO_BLANKSTRIP)
'''
'''                If rs2 Is Nothing Then
'''                    GetBlockData = FUNCTION_RETURN_FAILURE
'''                    GoTo proc_exit
'''                End If
'''
'''                If rs2.RecordCount > 0 Then
'''                    ' 最終通過管理工程
'''                    If IsNull(rs2("KNKTC4")) Then .LPKRPROCCD = "" Else .LPKRPROCCD = rs2("KNKTC4")
'''                    ' 最終通過工程
'''                    If IsNull(rs2("WKKTC4")) Then .LASTPASS = "" Else .LASTPASS = rs2("WKKTC4")
'''                    ' 不良理由
'''                    If IsNull(rs2("FCODEC4")) Then .BDCAUS = "0" Else .BDCAUS = rs2("FCODEC4")
'''                Else
'''                    .BDCAUS = "0"                       ' 不良理由
'''                End If
'''                rs2.Close
'''            Else
'''                ' 不良理由
'''                If IsNull(rs("BDCAUS")) Then .BDCAUS = "0" Else .BDCAUS = rs("BDCAUS")
'''            End If
        End With
        col.Add target
        rs.MoveNext
    Next
    rs.Close

    GetBlockData = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :テーブル「TBCME041」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME041 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/29作成　野村
Private Function GetTBCME041(col As c_cmzcHins, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
Dim target As c_cmzcHin

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME041"

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACTORY, OPECOND, LENGTH, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME041"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME041 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcHin
        With target
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .Factory = rs("FACTORY")         ' 工場
            .OpeCond = rs("OPECOND")         ' 操業条件
            .Length = rs("LENGTH")           ' 長さ
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME041 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :テーブル「TBCME042」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME042 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/29作成　野村
Private Function GetTBCME042(col As c_cmzcSxls, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
Dim target As c_cmzcSxl

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME042"

    ''SQLを組み立てる
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
    sqlBase = ""
    sqlBase = sqlBase & " SELECT"
    sqlBase = sqlBase & "  a.xtalcb as CRYNUM"        ''結晶番号
    sqlBase = sqlBase & " ,a.inposcb as INGOTPOS"     ''結晶内開始位置
    sqlBase = sqlBase & " ,a.rlencb as LENGTH"        ''理論長さ
    sqlBase = sqlBase & " ,a.sxlidcb as SXLID"        ''SXLID
    sqlBase = sqlBase & " ,' ' as KRPROCCD"           ''管理工程(ﾌﾞﾗﾝｸ)
    sqlBase = sqlBase & " ,a.gnwkntcb as NOWPROC"     ''現在工程
    sqlBase = sqlBase & " ,' ' as LPKRPROCCD"         ''最終通過管理工程(ﾌﾞﾗﾝｸ)
    sqlBase = sqlBase & " ,a.newkntcb as LASTPASS"    ''最終通過工程
    sqlBase = sqlBase & " ,a.livkcb as DELCLS"        ''生死区分
    sqlBase = sqlBase & " ,a.lstccb as LSTATCLS"      ''最終状態区分
    sqlBase = sqlBase & " ,a.sholdclscb as HOLDCLS"   ''ホールド区分
    sqlBase = sqlBase & " ,a.hinbcb as HINBAN"        ''品番
    sqlBase = sqlBase & " ,a.revnumcb as REVNUM"      ''製品番号改訂番号
    sqlBase = sqlBase & " ,a.factorycb as FACTORY"    ''工場
    sqlBase = sqlBase & " ,a.opecb as OPECOND"        ''操業条件
    sqlBase = sqlBase & " ,a.furyccb as BDCAUS"       ''不良理由
    sqlBase = sqlBase & " ,a.maicb as COUNT"          ''実枚数
    sqlBase = sqlBase & " ,a.tdaycb as REGDATE"       ''登録日付
    sqlBase = sqlBase & " ,a.kdaycb as UPDDATE"       ''更新日付
    sqlBase = sqlBase & " ,' ' as SUMMITSENDFLAG"     ''SUMMIT送信フラグ(ﾌﾞﾗﾝｸ)
    sqlBase = sqlBase & " ,a.sndkcb as SENDFLAG"      ''送信FLG
    sqlBase = sqlBase & " ,a.sndaycb as SENDDATE"     ''送信日付
    sqlBase = sqlBase & " FROM"
    sqlBase = sqlBase & "  xsdcb a"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
'        sql = sql & " AND a.livkcb = '0'"
'    Else
'        sql = sql & " WHERE a.livkcb = '0'"
    End If
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
'    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, SXLID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS, HOLDCLS," & _
'              " HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, COUNT, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCME042"
'    sql = sqlBase
'    If (sqlWhere <> vbNullString) Then
'        sql = sql & sqlWhere
'    End If
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        GetTBCME042 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcSxl
        With target
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            If IsNull(rs("LENGTH")) = False Then .Length = rs("LENGTH")         ' 長さ
            .SXLID = rs("SXLID")             ' SXLID
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程
            .NOWPROC = rs("NOWPROC")         ' 現在工程
            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
            .DELCLS = rs("DELCLS")           ' 削除区分
            .LSTATCLS = rs("LSTATCLS")       ' 最終状態区分
            If IsNull(rs("HOLDCLS")) = False Then .HOLDCLS = rs("HOLDCLS")      ' ホールド区分
            .hinban = rs("HINBAN")           ' 品番
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
            If .LSTATCLS = "H" Then
                .hinban = "Z"
            End If
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .Factory = rs("FACTORY")         ' 工場
            .OpeCond = rs("OPECOND")         ' 操業条件
            .BDCAUS = rs("BDCAUS")           ' 不良理由
            .COUNT = rs("COUNT")             ' 枚数
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME042 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :テーブル「XSDCS」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_XSDCS    ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/29作成　野村
Private Function GetTBCME043(col As c_cmzcXlSmps, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
Dim target As c_cmzcXlSmp
Dim j As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME043"

    ''SQLを組み立てる
'    sqlBase = "Select CRYNUM, INGOTPOS, SMPKBN, SMPLNO, HINBAN, REVNUM, FACTORY, OPECOND, KTKBN, CRYINDRS, CRYINDOI, CRYINDB1," & _
'              " CRYINDB2, CRYINDB3, CRYINDL1, CRYINDL2, CRYINDL3, CRYINDL4, CRYINDCS, CRYINDGD, CRYINDT, CRYINDEP, CRYRESRS," & _
'              " CRYRESOI, CRYRESB1, CRYRESB2, CRYRESB3, CRYRESL1, CRYRESL2, CRYRESL3, CRYRESL4, CRYRESCS, CRYRESGD, CRYREST," & _
'              " CRYRESEP, SMPLNUM, SMPLPAT, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCME043"
    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, " & _
              " BLKKTFLAGCS, CRYSMPLIDRSCS, CRYSMPLIDRS1CS, CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, CRYRESRS2CS,CRYSMPLIDOICS, " & _
              " CRYINDOICS, CRYRESOICS, CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, " & _
              " CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS,  CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, " & _
              " CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, " & _
              " CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS,  CRYRESTCS, CRYSMPLIDEPCS, CRYINDEPCS,CRYRESEPCS, SMPLNUMCS, " & _
              " SMPLPATCS, TSTAFFCS, TDAYCS, KSTAFFCS, KDAYCS, SNDKCS, SNDDAYCS "
    sqlBase = sqlBase & "From XSDCS"

    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME043 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcXlSmp
        With target
            .CRYNUM = rs("XTALCS")           ' 結晶番号
            .INGOTPOS = rs("INPOSCS")       ' 結晶内位置
            .SMPKBN = rs("SMPKBNCS")           ' サンプル区分
            .SMPLNO = rs("REPSMPLIDCS")           ' サンプルNo
            .hinban = rs("HINBCS")           ' 品番
            .REVNUM = rs("REVNUMCS")           ' 製品番号改訂番号
            .Factory = rs("FACTORYCS")         ' 工場
            .OpeCond = rs("OPECS")         ' 操業条件
            .KTKBN = rs("KTKBNCS")             ' 確定区分
            .CRYINDRS = rs("CRYINDRSCS")       ' 結晶検査指示（Rs)
            .CRYINDOI = rs("CRYINDOICS")       ' 結晶検査指示（Oi)
            .CRYINDB1 = rs("CRYINDB1CS")       ' 結晶検査指示（B1)
            .CRYINDB2 = rs("CRYINDB2CS")       ' 結晶検査指示（B2）
            .CRYINDB3 = rs("CRYINDB3CS")       ' 結晶検査指示（B3)
            .CRYINDL1 = rs("CRYINDL1CS")       ' 結晶検査指示（L1)
            .CRYINDL2 = rs("CRYINDL2CS")       ' 結晶検査指示（L2)
            .CRYINDL3 = rs("CRYINDL3CS")       ' 結晶検査指示（L3)
            .CRYINDL4 = rs("CRYINDL4CS")       ' 結晶検査指示（L4)
            .CRYINDCS = rs("CRYINDCSCS")       ' 結晶検査指示（Cs)
            .CRYINDGD = rs("CRYINDGDCS")       ' 結晶検査指示（GD)
            .CRYINDT = rs("CRYINDTCS")         ' 結晶検査指示（T)
            .CRYINDEP = rs("CRYINDEPCS")       ' 結晶検査指示（EPD)
            .CRYRESRS = rs("CRYRESRSCS")       ' 結晶検査実績（Rs)
            .CRYRESOI = rs("CRYRESOICS")       ' 結晶検査実績（Oi)
            .CRYRESB1 = rs("CRYRESB1CS")       ' 結晶検査実績（B1)
            .CRYRESB2 = rs("CRYRESB2CS")       ' 結晶検査実績（B2）
            .CRYRESB3 = rs("CRYRESB3CS")       ' 結晶検査実績（B3)
            .CRYRESL1 = rs("CRYRESL1CS")       ' 結晶検査実績（L1)
            .CRYRESL2 = rs("CRYRESL2CS")       ' 結晶検査実績（L2)
            .CRYRESL3 = rs("CRYRESL3CS")       ' 結晶検査実績（L3)
            .CRYRESL4 = rs("CRYRESL4CS")       ' 結晶検査実績（L4)
            .CRYRESCS = rs("CRYRESCSCS")       ' 結晶検査実績（Cs)
            .CRYRESGD = rs("CRYRESGDCS")       ' 結晶検査実績（GD)
            .CRYREST = rs("CRYRESTCS")         ' 結晶検査実績（T)
            .CRYRESEP = rs("CRYRESEPCS")       ' 結晶検査実績（EPD)
            .SMPLNUM = rs("SMPLNUMCS")         ' サンプル枚数
            .SMPLPAT = rs("SMPLPATCS")         ' サンプルパターン
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME043 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :テーブル「XSDCW」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_XSDCW    ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/29作成　野村
Private Function GetTBCME044(col As c_cmzcWfSmps, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
Dim target As c_cmzcWfSmp

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME044"

    ''SQLを組み立てる
'    sqlBase = "Select CRYNUM, INGOTPOS, SMPKBN, SMPLID, HINBAN, REVNUM, FACTORY, OPECOND, KTKBN, WFINDRS, WFINDOI, WFINDB1," & _
'              " WFINDB2, WFINDB3, WFINDL1, WFINDL2, WFINDL3, WFINDL4, WFINDDS, WFINDDZ, WFINDSP, WFINDDO1, WFINDDO2, WFINDDO3," & _
'              " WFRESRS, WFRESOI, WFRESB1, WFRESB2, WFRESB3, WFRESL1, WFRESL2, WFRESL3, WFRESL4, WFRESDS, WFRESDZ, WFRESSP," & _
'              " WFRESDO1, WFRESDO2, WFRESDO3, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCME044"

    'GD項目追加　05/01/17 ooba
    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)kondoh
    sqlBase = "Select SXLIDCW, SMPKBNCW, TBKBNCW, REVNUMCW, XTALCW, INPOSCW, REPSMPLIDCW, HINBCW, FACTORYCW, OPECW, KTKBNCW, " & _
              " SMCRYNUMCW, WFSMPLIDRSCW, WFSMPLIDRS1CW, WFSMPLIDRS2CW, WFINDRSCW, WFRESRS1CW, WFRESRS2CW, WFSMPLIDOICW, WFINDOICW, " & _
              " WFRESOICW, WFSMPLIDB1CW, WFINDB1CW, WFRESB1CW, WFSMPLIDB2CW, WFINDB2CW, WFRESB2CW, WFSMPLIDB3CW, WFINDB3CW, " & _
              " WFRESB3CW, WFSMPLIDL1CW, WFINDL1CW, WFRESL1CW, WFSMPLIDL2CW, WFINDL2CW, WFRESL2CW, WFSMPLIDL3CW, WFINDL3CW, WFRESL3CW," & _
              " WFSMPLIDL4CW, WFINDL4CW, WFRESL4CW, WFSMPLIDDSCW, WFINDDSCW, WFRESDSCW, WFSMPLIDDZCW, WFINDDZCW, WFRESDZCW, " & _
              " WFSMPLIDSPCW, WFINDSPCW, WFRESSPCW, WFSMPLIDDO1CW, WFINDDO1CW, WFRESDO1CW, WFSMPLIDDO2CW, WFINDDO2CW, WFRESDO2CW, " & _
              " WFSMPLIDDO3CW, WFINDDO3CW, WFRESDO3CW, WFSMPLIDAOICW, WFINDAOICW, WFRESAOICW, SMPLNUMCW, SMPLPATCW, TSTAFFCW, TDAYCW, " & _
              " KSTAFFCW, KDAYCW, SNDKCW, SNDDAYCW, WFSMPLIDGDCW, WFINDGDCW, WFRESGDCW, WFHSGDCW " & _
              " ,EPSMPLIDB1CW, EPINDB1CW, EPRESB1CW, EPSMPLIDB2CW, EPINDB2CW, EPRESB2CW, EPSMPLIDB3CW, EPINDB3CW, EPRESB3CW, " & _
              " EPSMPLIDL1CW, EPINDL1CW, EPRESL1CW, EPSMPLIDL2CW, EPINDL2CW, EPRESL2CW, EPSMPLIDL3CW, EPINDL3CW, EPRESL3CW "

    sqlBase = sqlBase & "From XSDCW"

    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
'確定区分が’９’だったら除く　濱　平成１５年５月３０日
'       sql = sql & sqlWhere
''      sql = sql & sqlWhere & " and SMPKBN != '9'"
        sql = sql & sqlWhere & " and SMPKBNCW != '9'"
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME044 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcWfSmp
        With target
            If IsNull(rs("XTALCW")) = False Then .CRYNUM = rs("XTALCW")             ' 結晶番号
            If IsNull(rs("INPOSCW")) = False Then .INGOTPOS = rs("INPOSCW")         ' 結晶内位置
            If IsNull(rs("SMPKBNCW")) = False Then .SMPKBN = rs("SMPKBNCW")         ' サンプル区分
            If IsNull(rs("REPSMPLIDCW")) = False Then .SMPLID = rs("REPSMPLIDCW")   ' サンプルID
            If IsNull(rs("HINBCW")) = False Then .hinban = rs("HINBCW")             ' 品番
            If IsNull(rs("REVNUMCW")) = False Then .REVNUM = rs("REVNUMCW")         ' 製品番号改訂番号
            If IsNull(rs("FACTORYCW")) = False Then .Factory = rs("FACTORYCW")      ' 工場
            If IsNull(rs("OPECW")) = False Then .OpeCond = rs("OPECW")              ' 操業条件
            If IsNull(rs("KTKBNCW")) = False Then .KTKBN = rs("KTKBNCW")            ' 確定区分
            If IsNull(rs("WFINDRSCW")) = False Then .WFINDRS = rs("WFINDRSCW")      ' WF検査指示（Rs)
            If IsNull(rs("WFINDOICW")) = False Then .WFINDOI = rs("WFINDOICW")      ' WF検査指示（Oi)
            If IsNull(rs("WFINDB1CW")) = False Then .WFINDB1 = rs("WFINDB1CW")      ' WF検査指示（B1)
            If IsNull(rs("WFINDB2CW")) = False Then .WFINDB2 = rs("WFINDB2CW")      ' WF検査指示（B2）
            If IsNull(rs("WFINDB3CW")) = False Then .WFINDB3 = rs("WFINDB3CW")      ' WF検査指示（B3)
            If IsNull(rs("WFINDL1CW")) = False Then .WFINDL1 = rs("WFINDL1CW")      ' WF検査指示（L1)
            If IsNull(rs("WFINDL2CW")) = False Then .WFINDL2 = rs("WFINDL2CW")      ' WF検査指示（L2)
            If IsNull(rs("WFINDL3CW")) = False Then .WFINDL3 = rs("WFINDL3CW")      ' WF検査指示（L3)
            If IsNull(rs("WFINDL4CW")) = False Then .WFINDL4 = rs("WFINDL4CW")      ' WF検査指示（L4)
            If IsNull(rs("WFINDDSCW")) = False Then .WFINDDS = rs("WFINDDSCW")      ' WF検査指示（DS)
            If IsNull(rs("WFINDDZCW")) = False Then .WFINDDZ = rs("WFINDDZCW")      ' WF検査指示（DZ)
            If IsNull(rs("WFINDSPCW")) = False Then .WFINDSP = rs("WFINDSPCW")      ' WF検査指示（SP)
            If IsNull(rs("WFINDDO1CW")) = False Then .WFINDDO1 = rs("WFINDDO1CW")   ' WF検査指示（DO1)
            If IsNull(rs("WFINDDO2CW")) = False Then .WFINDDO2 = rs("WFINDDO2CW")   ' WF検査指示（DO2)
            If IsNull(rs("WFINDDO3CW")) = False Then .WFINDDO3 = rs("WFINDDO3CW")   ' WF検査指示（DO3)
            If IsNull(rs("WFINDAOICW")) = False Then .WFINDAOI = rs("WFINDAOICW")   ' WF検査指示 (AO)　追加　03/12/05 ooba
            If IsNull(rs("WFINDGDCW")) = False Then .WFINDGD = rs("WFINDGDCW")      ' WF検査指示 (GD)　追加　05/01/17 ooba
            If IsNull(rs("WFRESRS1CW")) = False Then .WFRESRS = rs("WFRESRS1CW")    ' WF検査実績（Rs)
            If IsNull(rs("WFRESOICW")) = False Then .WFRESOI = rs("WFRESOICW")      ' WF検査実績（Oi)
            If IsNull(rs("WFRESB1CW")) = False Then .WFRESB1 = rs("WFRESB1CW")      ' WF検査実績（B1)
            If IsNull(rs("WFRESB2CW")) = False Then .WFRESB2 = rs("WFRESB2CW")      ' WF検査実績（B2）
            If IsNull(rs("WFRESB3CW")) = False Then .WFRESB3 = rs("WFRESB3CW")      ' WF検査実績（B3)
            If IsNull(rs("WFRESL1CW")) = False Then .WFRESL1 = rs("WFRESL1CW")      ' WF検査実績（L1)
            If IsNull(rs("WFRESL2CW")) = False Then .WFRESL2 = rs("WFRESL2CW")      ' WF検査実績（L2)
            If IsNull(rs("WFRESL3CW")) = False Then .WFRESL3 = rs("WFRESL3CW")      ' WF検査実績（L3)
            If IsNull(rs("WFRESL4CW")) = False Then .WFRESL4 = rs("WFRESL4CW")      ' WF検査実績（L4)
            If IsNull(rs("WFRESDSCW")) = False Then .WFRESDS = rs("WFRESDSCW")      ' WF検査実績（DS)
            If IsNull(rs("WFRESDZCW")) = False Then .WFRESDZ = rs("WFRESDZCW")      ' WF検査実績（DZ)
            If IsNull(rs("WFRESSPCW")) = False Then .WFRESSP = rs("WFRESSPCW")      ' WF検査実績（SP)
            If IsNull(rs("WFRESDO1CW")) = False Then .WFRESDO1 = rs("WFRESDO1CW")   ' WF検査実績（DO1)
            If IsNull(rs("WFRESDO2CW")) = False Then .WFRESDO2 = rs("WFRESDO2CW")   ' WF検査実績（DO2)
            If IsNull(rs("WFRESDO3CW")) = False Then .WFRESDO3 = rs("WFRESDO3CW")   ' WF検査実績（DO3)
            If IsNull(rs("WFRESAOICW")) = False Then .WFRESAOI = rs("WFRESAOICW")   ' WF検査実績 (AO)　追加　03/12/05 ooba
            If IsNull(rs("WFRESGDCW")) = False Then .WFRESGD = rs("WFRESGDCW")      ' WF検査実績 (GD)　追加　05/01/17 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
            If IsNull(rs("EPINDB1CW")) = False Then .EPINDB1 = rs("EPINDB1CW")      ' WF検査指示 (BMD1E)
            If IsNull(rs("EPRESB1CW")) = False Then .EPRESB1 = rs("EPRESB1CW")      ' WF検査実績 (BMD1E)
            If IsNull(rs("EPINDB2CW")) = False Then .EPINDB2 = rs("EPINDB2CW")      ' WF検査指示 (BMD2E)
            If IsNull(rs("EPRESB2CW")) = False Then .EPRESB2 = rs("EPRESB2CW")      ' WF検査実績 (BMD2E)
            If IsNull(rs("EPINDB3CW")) = False Then .EPINDB3 = rs("EPINDB3CW")      ' WF検査指示 (BMD3E)
            If IsNull(rs("EPRESB3CW")) = False Then .EPRESB3 = rs("EPRESB3CW")      ' WF検査実績 (BMD3E)
            If IsNull(rs("EPINDL1CW")) = False Then .EPINDL1 = rs("EPINDL1CW")      ' WF検査指示 (OSF1E)
            If IsNull(rs("EPRESL1CW")) = False Then .EPRESL1 = rs("EPRESL1CW")      ' WF検査実績 (OSF1E)
            If IsNull(rs("EPINDL2CW")) = False Then .EPINDL2 = rs("EPINDL2CW")      ' WF検査指示 (OSF2E)
            If IsNull(rs("EPRESL2CW")) = False Then .EPRESL2 = rs("EPRESL2CW")      ' WF検査実績 (OSF2E)
            If IsNull(rs("EPINDL3CW")) = False Then .EPINDL3 = rs("EPINDL3CW")      ' WF検査指示 (OSF3E)
            If IsNull(rs("EPRESL3CW")) = False Then .EPRESL3 = rs("EPRESL3CW")      ' WF検査実績 (OSF3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME044 = FUNCTION_RETURN_SUCCESS

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


'概要      :テーブル「TBCME045」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME045 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/29作成　野村
Private Function GetTBCME045(col As c_cmzcCuts, CRYNUM$) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
Dim target As c_cmzcCut

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME045"

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, TRANCNT, LENGTH, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS," & _
              " STATCLS, BLOCKID, CRYINDRS, CRYINDOI, CRYINDB1, CRYINDB2, CRYINDB3, CRYINDL1, CRYINDL2, CRYINDL3, CRYINDL4," & _
              " CRYINDCS, CRYINDGD, CRYINDT, CRYINDEP, PRIORITY "
    sqlBase = sqlBase & "From TBCME045 IT " & _
              "Where (CRYNUM='" & CRYNUM & "') And (STATCLS<>'1') " & _
              "  and (TRANCNT=(select MAX(TRANCNT) from TBCME045 where (CRYNUM='" & CRYNUM & "') and (STATCLS<>'1')))"
    sql = sqlBase

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME045 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcCut
        With target
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .TRANCNT = rs("TRANCNT")         ' 処理回数
            .Length = rs("LENGTH")           ' 長さ
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 品番製品番号改訂番号
            .Factory = rs("FACTORY")         ' 品番工場
            .OpeCond = rs("OPECOND")         ' 品番操業条件
            .BDCAUS = rs("BDCAUS")           ' 区分コード
            .STATCLS = rs("STATCLS")         ' 状態区分
            .BLOCKID = rs("BLOCKID")         ' ブロックID
            .CRYINDRS = rs("CRYINDRS")       ' 結晶検査指示（Rs)
            .CRYINDOI = rs("CRYINDOI")       ' 結晶検査指示（Oi)
            .CRYINDB1 = rs("CRYINDB1")       ' 結晶検査指示（B1)
            .CRYINDB2 = rs("CRYINDB2")       ' 結晶検査指示（B2）
            .CRYINDB3 = rs("CRYINDB3")       ' 結晶検査指示（B3)
            .CRYINDL1 = rs("CRYINDL1")       ' 結晶検査指示（L1)
            .CRYINDL2 = rs("CRYINDL2")       ' 結晶検査指示（L2)
            .CRYINDL3 = rs("CRYINDL3")       ' 結晶検査指示（L3)
            .CRYINDL4 = rs("CRYINDL4")       ' 結晶検査指示（L4)
            .CRYINDCS = rs("CRYINDCS")       ' 結晶検査指示（Cs)
            .CRYINDGD = rs("CRYINDGD")       ' 結晶検査指示（GD)
            .CRYINDT = rs("CRYINDT")         ' 結晶検査指示（T)
            .CRYINDEP = rs("CRYINDEP")       ' 結晶検査指示（EPD)
            .PRIORITY = rs("PRIORITY")       ' 優先度
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME045 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

Private Function GetReject(col As c_cmzcRejs, CRYNUM$) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
Dim target As c_cmzcRej

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetReject"
    
    GetReject = FUNCTION_RETURN_FAILURE
    
    ''SQLを組み立てる
''    sql = "select LOTID, ALLSCRAP, REJFROM, REJTO " & _
          "from VECMW004 " & _
          "where (LOTID like '" & Left$(CRYNUM, 9) & "%') and (REJCAT<>'C') " & _
          "order by LOTID, REJFROM"

    'ﾋﾞｭｰ参照停止　06/02/06 ooba START ====================================================>
    sql = "select LOTID, ALLSCRAP, REJFROM, REJTO from "
    sql = sql & "("
    sql = sql & "select "
    sql = sql & "  C.CRYNUM,"
    sql = sql & "  XXX.LOTID,"
    sql = sql & "  REJCAT,"
    sql = sql & "  ALLSCRAP,"
    sql = sql & "  case when (XXX.REJFROM<=B.WFFROM) then 0 else XXX.REJFROM end as REJFROM,"
    sql = sql & "  case when (XXX.REJTO>=B.WFTO) then C.LENGTH else XXX.REJTO end as REJTO,"
    sql = sql & "  REJWFFROM,"
    sql = sql & "  REJWFTO"
    sql = sql & " from "
    sql = sql & "("
    sql = sql & "select "
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    0 as REJFROM,"
    sql = sql & "    LENGTH as REJTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCME040 B"
    sql = sql & "  where (A.LOTID=B.BLOCKID)"
    sql = sql & "    and (A.ALLSCRAP='Y')"
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    LENFROM,"
    sql = sql & "    LENTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012"
'    sql = sql & "  where (REJCAT='A') and (ALLSCRAP='N')"
    sql = sql & "  where (REJCAT in ('A','E')) and (ALLSCRAP='N')"      'ﾌﾞﾛｯｸ状態での一部欠量対応 09/02/27 ooba
    sql = sql & " and lotid like '" & left$(CRYNUM, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    A.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    A.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ as REJWFTO"
    sql = sql & "  from TBCMY012 A"
    sql = sql & "  where (A.REJCAT='B') and (ALLSCRAP='N')"
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    B.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    C.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ+A.REJPCS-1 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCMY011 B,"
    sql = sql & "    TBCMY011 C"
    sql = sql & "  where (A.REJCAT='C')"
    sql = sql & "    and (A.LOTID=B.LOTID) and (A.BLOCKSEQ=B.BLOCKSEQ)"
    sql = sql & "    and (A.LOTID=C.LOTID) and (A.BLOCKSEQ+A.REJPCS-1=C.BLOCKSEQ)"
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
    sql = sql & " order by LOTID,REJFROM"
    sql = sql & ") XXX,"
    sql = sql & "  (select LOTID, min(TOP_POS)/10.0 as WFFROM, max(TOP_POS)/10.0 as WFTO from TBCMY011 "
    sql = sql & " where lotid like '" & left$(CRYNUM, 9) & "%'"
    sql = sql & " group by LOTID) B,"
    sql = sql & "  TBCME040 C"
    sql = sql & " where (XXX.LOTID=B.LOTID)"
    sql = sql & "  and (XXX.LOTID=C.BLOCKID)"
    sql = sql & "  and (XXX.ALLSCRAP='N')"
    sql = sql & " union all"
    sql = sql & " select distinct"
    sql = sql & "  C.CRYNUM,"
    sql = sql & "  XXX.LOTID,"
    sql = sql & "  REJCAT,"
    sql = sql & "  ALLSCRAP,"
    sql = sql & "  0 as REJFROM,"
    sql = sql & "  C.LENGTH as REJTO,"
    sql = sql & "  REJWFFROM,"
    sql = sql & "  REJWFTO"
    sql = sql & " from "
    sql = sql & "("
    sql = sql & "select "
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    0 as REJFROM,"
    sql = sql & "    LENGTH as REJTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCME040 B"
    sql = sql & "  where (A.LOTID=B.BLOCKID)"
    sql = sql & "    and (A.ALLSCRAP='Y')"
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    LENFROM,"
    sql = sql & "    LENTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012"
'    sql = sql & "  where (REJCAT='A') and (ALLSCRAP='N')"
    sql = sql & "  where (REJCAT in ('A','E')) and (ALLSCRAP='N')"      'ﾌﾞﾛｯｸ状態での一部欠量対応 09/02/27 ooba
    sql = sql & " and lotid like '" & left$(CRYNUM, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    A.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    A.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ as REJWFTO"
    sql = sql & "  from TBCMY012 A"
    sql = sql & "  where (A.REJCAT='B') and (ALLSCRAP='N')"
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    B.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    C.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ+A.REJPCS-1 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCMY011 B,"
    sql = sql & "    TBCMY011 C"
    sql = sql & "  where (A.REJCAT='C')"
    sql = sql & "    and (A.LOTID=B.LOTID) and (A.BLOCKSEQ=B.BLOCKSEQ)"
    sql = sql & "    and (A.LOTID=C.LOTID) and (A.BLOCKSEQ+A.REJPCS-1=C.BLOCKSEQ)"
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
    sql = sql & " order by LOTID,REJFROM"
    sql = sql & ") XXX,"
    sql = sql & "  TBCME040 C"
    sql = sql & " where (XXX.LOTID=C.BLOCKID)"
    sql = sql & "  and (XXX.ALLSCRAP='Y')"
    sql = sql & ")"
    sql = sql & " where (REJCAT<>'C')"
    sql = sql & " order by LOTID, REJFROM "
    'ﾋﾞｭｰ参照停止　06/02/06 ooba END ======================================================>
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcRej
        With target
            .LOTID = rs("LOTID")                  ' ブロックID
            .ALLSCRAP = rs("ALLSCRAP")            ' 全数スクラップ
            .LENFROM = rs("REJFROM")              ' 長さ　FROM
            .LENTO = rs("REJTO")                  ' 長さ　TO
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close
    
    GetReject = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

''↓追加 START SPT用実績作成方法変更 2006/05/12 SMP-OKAMOTO
'概要      :テーブル「XSDC2」「XSDCS」「XSDC4」からブロック情報を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型                ,説明
'          :col           ,O   ,c_cmzcBlks        ,抽出レコード
'          :CRYNUM        ,I   ,String            ,結晶番号
'          :戻り値        ,O   ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :2005/10/05 ooba
Private Function GetBlockData_2(col As c_cmzcBlks, Optional CRYNUM$) As FUNCTION_RETURN

    Dim sql As String       'SQL全体
    Dim sql2 As String      'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim rs2 As OraDynaset   'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long
    Dim target As c_cmzcBlk

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetBlockData_2"
    
    ''SQLを組み立てる
    ''まずは切断指示なしのブロックを取得
    sql = "select "
    sql = sql & "CSTOP.XTALCS, "                                            '結晶番号
    sql = sql & "CSTOP.INPOSCS　INPOSCS, "                                  '結晶内開始位置
    sql = sql & "CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, "                 '長さ
    sql = sql & "nvl(GNLC2,CSBOT.INPOSCS - CSTOP.INPOSCS) as REALLEN, "     '実長さ
    sql = sql & "CSTOP.CRYNUMCS, "                                          'ブロックID
    sql = sql & "nvl(GNKKNTC2,' ') as KRPROCCD, "                           '現在管理工程
    sql = sql & "nvl(GNWKNTC2,' ') as NOWPROC, "                            '現在工程
    sql = sql & "nvl(NEKKNTC2,' ') as LPKRPROCCD, "                         '最終通過管理工程
    sql = sql & "nvl(NEWKNTC2,' ') as LASTPASS, "                           '最終通過工程
    sql = sql & "nvl(SAKJC2,'0') as DELCLS, "                               '削除区分
    sql = sql & "nvl(LSTATBC2,'T') as LSTATCLS, "                           '最終状態区分
    sql = sql & "nvl(RSTATBC2,'T') as RSTATCLS, "                           '流動状態区分
    sql = sql & "nvl(HOLDBC2,'0') as HOLDCLS, "                             'ホールド区分
    sql = sql & "BDCAUSC2 as BDCAUS, "                                      '不良理由
    sql = sql & "C4.KNKTC4, "                                               '最終通過管理工程(XSDC4)
    sql = sql & "C4.WKKTC4, "                                               '最終通過工程(XSDC4)
    sql = sql & "C4.FCODEC4 "                                               '不良理由(XSDC4)
    sql = sql & "from XSDC2, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS, "
    sql = sql & "      CUTFLGCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'T' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSTOP, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS, "
    sql = sql & "      CUTFLGCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'B' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSBOT, "
    sql = sql & "     (select "
    sql = sql & "      XTALC4, "
    sql = sql & "      INPOSC4, "
    sql = sql & "      KNKTC4, "
    sql = sql & "      WKKTC4, "
    sql = sql & "      FCODEC4 "
    sql = sql & "      from XSDC4 TMP4 "
    sql = sql & "      where "
    sql = sql & "      XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "      and (KCKNTC4, KDAYC4) = ("
    sql = sql & "                     select MAX(KCKNTC4), MAX(KDAYC4) "
    sql = sql & "                     from XSDC4 "
    sql = sql & "                     where XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "                     and INPOSC4 = TMP4.INPOSC4) "
    sql = sql & "     ) C4 "
    sql = sql & "where "
    sql = sql & "CSTOP.CRYNUMCS = CRYNUMC2(+) "
    sql = sql & "and CSTOP.CUTFLGCS is null "
    sql = sql & "and CSTOP.CRYNUMCS = CSBOT.CRYNUMCS "
    sql = sql & "and CSTOP.INPOSCS = C4.INPOSC4(+) "
    sql = sql & "and (LIVKC2 is null or LIVKC2 = '0' "
    sql = sql & "     or LSTATBC2 in ('R', 'H', 'B') or KANKC2 = '2') "
    sql = sql & " UNION ("
    ''次に切断指示ありのブロックを取得
    sql = sql & "select "
    sql = sql & "CSTOP.XTALCS, "                                            '結晶番号
    sql = sql & "CSTOP.INPOSCS　INPOSCS, "                                  '結晶内開始位置
    sql = sql & "CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, "                 '長さ
    sql = sql & "nvl(GNLC2,CSBOT.INPOSCS - CSTOP.INPOSCS) as REALLEN, "     '実長さ
    sql = sql & "CSTOP.CRYNUMCS, "                                          'ブロックID
    sql = sql & "nvl(GNKKNTC2,' ') as KRPROCCD, "                           '現在管理工程
    sql = sql & "nvl(GNWKNTC2,' ') as NOWPROC, "                            '現在工程
    sql = sql & "nvl(NEKKNTC2,' ') as LPKRPROCCD, "                         '最終通過管理工程
    sql = sql & "nvl(NEWKNTC2,' ') as LASTPASS, "                           '最終通過工程
    sql = sql & "nvl(SAKJC2,'0') as DELCLS, "                               '削除区分
    sql = sql & "nvl(LSTATBC2,'T') as LSTATCLS, "                           '最終状態区分
    sql = sql & "nvl(RSTATBC2,'T') as RSTATCLS, "                           '流動状態区分
    sql = sql & "nvl(HOLDBC2,'0') as HOLDCLS, "                             'ホールド区分
    sql = sql & "BDCAUSC2 as BDCAUS, "                                      '不良理由
    sql = sql & "C4.KNKTC4, "                                               '最終通過管理工程(XSDC4)
    sql = sql & "C4.WKKTC4, "                                               '最終通過工程(XSDC4)
    sql = sql & "C4.FCODEC4 "                                               '不良理由(XSDC4)
    sql = sql & "from XSDC2, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS, "
    sql = sql & "      RPCRYNUMCS, "
    sql = sql & "      CUTFLGCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'T' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSTOP, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS, "
    sql = sql & "      RPCRYNUMCS, "
    sql = sql & "      CUTFLGCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'B' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSBOT, "
    sql = sql & "     (select "
    sql = sql & "      XTALC4, "
    sql = sql & "      INPOSC4, "
    sql = sql & "      KNKTC4, "
    sql = sql & "      WKKTC4, "
    sql = sql & "      FCODEC4 "
    sql = sql & "      from XSDC4 TMP4 "
    sql = sql & "      where "
    sql = sql & "      XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "      and (KCKNTC4, KDAYC4) = ("
    sql = sql & "                     select MAX(KCKNTC4), MAX(KDAYC4) "
    sql = sql & "                     from XSDC4 "
    sql = sql & "                     where XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "                     and INPOSC4 = TMP4.INPOSC4) "
    sql = sql & "     ) C4 "
    sql = sql & "where "
    sql = sql & "CSTOP.RPCRYNUMCS = CRYNUMC2(+) "
    sql = sql & "and CSTOP.CUTFLGCS = '1' "
    sql = sql & "and CSTOP.CRYNUMCS = CSBOT.CRYNUMCS "
    sql = sql & "and CSTOP.INPOSCS = C4.INPOSC4(+) "
    sql = sql & "and (LIVKC2 is null or LIVKC2 = '0' "
    sql = sql & "     or LSTATBC2 in ('R', 'H', 'B') or KANKC2 = '2') "
    sql = sql & " ) "
    sql = sql & "order by INPOSCS "
    
    
    
''↓削除 START SPT用実績作成方法変更 IT障害 2006/06/14 SMP-OKAMOTO
'    sql = "select "
'    sql = sql & "DISTINCT "
'    sql = sql & "CSTOP.XTALCS, "                                            '結晶番号
'    sql = sql & "CSTOP.INPOSCS, "                                           '結晶内開始位置
'    sql = sql & "CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, "                 '長さ
'    sql = sql & "nvl(GNLC2,CSBOT.INPOSCS - CSTOP.INPOSCS) as REALLEN, "     '実長さ
'    sql = sql & "CSTOP.CRYNUMCS, "                                          'ブロックID
'    sql = sql & "nvl(GNKKNTC2,' ') as KRPROCCD, "                           '現在管理工程
'    sql = sql & "nvl(GNWKNTC2,' ') as NOWPROC, "                            '現在工程
'    sql = sql & "nvl(NEKKNTC2,' ') as LPKRPROCCD, "                         '最終通過管理工程
'    sql = sql & "nvl(NEWKNTC2,' ') as LASTPASS, "                           '最終通過工程
'    sql = sql & "nvl(SAKJC2,'0') as DELCLS, "                               '削除区分
'    sql = sql & "nvl(LSTATBC2,'T') as LSTATCLS, "                           '最終状態区分
'    sql = sql & "nvl(RSTATBC2,'T') as RSTATCLS, "                           '流動状態区分
'    sql = sql & "nvl(HOLDBC2,'0') as HOLDCLS, "                             'ホールド区分
'    sql = sql & "BDCAUSC2 as BDCAUS, "                                      '不良理由
'    sql = sql & "C4.KNKTC4, "                                               '最終通過管理工程(XSDC4)
'    sql = sql & "C4.WKKTC4, "                                               '最終通過工程(XSDC4)
'    sql = sql & "C4.FCODEC4 "                                               '不良理由(XSDC4)
'    sql = sql & "from XSDC2, "
'    sql = sql & "     (select "
'    sql = sql & "      CRYNUMCS, "
'    sql = sql & "      RPCRYNUMCS, " 'add
'    sql = sql & "      XTALCS, "
'    sql = sql & "      INPOSCS "
'    sql = sql & "      from XSDCS "
'    sql = sql & "      where "
'    sql = sql & "      TBKBNCS = 'T' "
'    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
'    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
'    sql = sql & "     ) CSTOP, "
'    sql = sql & "     (select "
'    sql = sql & "      CRYNUMCS, "
'    sql = sql & "      RPCRYNUMCS, " 'add
'    sql = sql & "      XTALCS, "
'    sql = sql & "      INPOSCS "
'    sql = sql & "      from XSDCS "
'    sql = sql & "      where "
'    sql = sql & "      TBKBNCS = 'B' "
'    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
'    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
'    sql = sql & "     ) CSBOT, "
'    sql = sql & "     (select "
'    sql = sql & "      XTALC4, "
'    sql = sql & "      INPOSC4, "
'    sql = sql & "      KNKTC4, "
'    sql = sql & "      WKKTC4, "
'    sql = sql & "      FCODEC4 "
'    sql = sql & "      from XSDC4 TMP4 "
'    sql = sql & "      where "
'    sql = sql & "      XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
'    sql = sql & "      and (KCKNTC4, KDAYC4) = ("
'    sql = sql & "                     select MAX(KCKNTC4), MAX(KDAYC4) "
'    sql = sql & "                     from XSDC4 "
'    sql = sql & "                     where XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
'    sql = sql & "                     and INPOSC4 = TMP4.INPOSC4) "
'    sql = sql & "     ) C4 "
'    sql = sql & " ,XSDCZ "
'    sql = sql & "where "
'    sql = sql & "CSTOP.CRYNUMCS = CRYNUMCZ(+) "
'    sql = sql & "and RPCRYNUMCZ = CRYNUMC2 "
''    sql = sql & "and CSTOP.CRYNUMCS = CRYNUMC2(+) "
'    sql = sql & "and CSTOP.CRYNUMCS = CSBOT.CRYNUMCS "
'    sql = sql & "and CSTOP.INPOSCS = C4.INPOSC4(+) "
'    sql = sql & "and (LIVKC2 is null or LIVKC2 = '0' "
'    sql = sql & "     or LSTATBC2 in ('R', 'H', 'B') or KANKC2 = '2') "
'    sql = sql & "order by CSTOP.INPOSCS "
''↑削除 END   SPT用実績作成方法変更 IT障害 2006/06/14 SMP-OKAMOTO
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetBlockData_2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcBlk
        With target
            .CRYNUM = rs("XTALCS")              ' 結晶番号
            .INGOTPOS = rs("INPOSCS")           ' 結晶内開始位置
            .Length = rs("LENGTH")              ' 長さ
            .REALLEN = rs("REALLEN")            ' 実長さ
            .BLOCKID = rs("CRYNUMCS")           ' ブロックID
            .KRPROCCD = rs("KRPROCCD")          ' 現在管理工程
            .NOWPROC = rs("NOWPROC")            ' 現在工程
            .LPKRPROCCD = rs("LPKRPROCCD")      ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")          ' 最終通過工程
            .DELCLS = rs("DELCLS")              ' 削除区分
            .LSTATCLS = rs("LSTATCLS")          ' 最終状態区分
            .RSTATCLS = rs("RSTATCLS")          ' 流動状態区分
            .HOLDCLS = rs("HOLDCLS")            ' ホールド区分
            If InStr(.BLOCKID, "$") <> 0 Then
                .KRPROCCD = MGPRCD_RIMERUTO_UKEIRE          ' 現在管理工程
                .NOWPROC = PROCD_RIMERUTO_UKEIRE            ' 現在工程
                .RSTATCLS = "M"                             ' 流動状態区分
                ' 最終通過管理工程
                If IsNull(rs("KNKTC4")) Then .LPKRPROCCD = "" Else .LPKRPROCCD = rs("KNKTC4")
                ' 最終通過工程
                If IsNull(rs("WKKTC4")) Then .LASTPASS = "" Else .LASTPASS = rs("WKKTC4")
                ' 不良理由
                If IsNull(rs("FCODEC4")) Then .BDCAUS = "0" Else .BDCAUS = rs("FCODEC4")
            Else
                ' 不良理由
                If IsNull(rs("BDCAUS")) Then .BDCAUS = "0" Else .BDCAUS = rs("BDCAUS")
            End If
            If Trim(.NOWPROC) = "" Then .DELCLS = "1"
            
        End With
        col.Add target
        rs.MoveNext
    Next
    rs.Close

    GetBlockData_2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

''↑追加 END   SPT用実績作成方法変更 2006/05/12 SMP-OKAMOTO
