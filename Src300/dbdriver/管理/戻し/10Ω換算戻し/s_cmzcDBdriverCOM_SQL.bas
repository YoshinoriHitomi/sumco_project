Attribute VB_Name = "s_cmzcDBdriverCOM_SQL"
Option Explicit

' DBドライバ共通関数

'概要      :引上げ終了実績、コードマスターからシード傾きを取得
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型               ,説明
'      　　:CRYNUM　　　,I  ,String         　,結晶番号
'      　　:SEED  　　　,I  ,Integer        　,シード傾き
'      　　:戻り値      ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_getSEED(ByVal CRYNUM As String, SEED As Integer) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_getSEED"

    sql = "select INFO3"
    sql = sql & " from TBCME037 H, TBCMB005 CM"
    sql = sql & " where H.CRYNUM='" & CRYNUM & "'"
    sql = sql & " and rtrim(CM.CODE,' ')=substr(H.SEED,1,1)"
    sql = sql & " and SYSCLASS='SC'"
    sql = sql & " and CLASS='28'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then
        DBDRV_getSEED = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    SEED = val(rs("INFO3"))
    rs.Close

    DBDRV_getSEED = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_getSEED = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :結晶情報の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型               ,説明
'      　　:CryInf　　　,I  ,typ_TBCME037   　,結晶情報
'      　　:戻り値      ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_CryInf_Ins(CryInf As typ_TBCME037) As FUNCTION_RETURN

    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CryInf_Ins"

    '' 結晶情報の挿入
    With CryInf
        sql = "insert into TBCME037 ("
        sql = sql & "CRYNUM, "              ' 結晶番号
        sql = sql & "DELCLS, "              ' 削除区分
        sql = sql & "KRPROCCD, "            ' 管理工程コード
        sql = sql & "PROCCD, "              ' 工程コード
        sql = sql & "LPKRPROCCD, "          ' 最終通過管理工程
        sql = sql & "LASTPASS, "            ' 最終通過工程
        sql = sql & "RPHINBAN, "            ' ねらい品番
        sql = sql & "RPREVNUM, "            ' ねらい品番製品番号改訂番号
        sql = sql & "RPFACT, "              ' ねらい品番工場
        sql = sql & "RPOPCOND, "            ' ねらい品番操業条件
        sql = sql & "PRODCOND, "            ' 製作条件
        sql = sql & "PGID, "                ' ＰＧ－ＩＤ
        sql = sql & "UPLENGTH, "            ' 引上げ長さ
        sql = sql & "TOPLENG, "             ' ＴＯＰ長さ
        sql = sql & "BODYLENG, "            ' 直胴長さ
        sql = sql & "BOTLENG, "             ' ＢＯＴ長さ
        sql = sql & "FREELENG, "            ' フリー長
        sql = sql & "DIAMETER, "            ' 直径
        sql = sql & "CHARGE, "              ' チャージ量
        sql = sql & "SEED, "                ' シード
        sql = sql & "ADDDPCLS, "            ' 追加ドープ種類
        sql = sql & "ADDDPPOS, "            ' 追加ドープ位置
        sql = sql & "ADDDPVAL, "            ' 追加ドープ量
        sql = sql & "REGDATE, "             ' 登録日付
        sql = sql & "UPDDATE, "             ' 更新日付
        sql = sql & "SENDFLAG, "            ' 送信フラグ
        sql = sql & "SENDDATE)"             ' 送信日付
        sql = sql & " values ('"
        sql = sql & .CRYNUM & "', '"        ' 結晶番号
        sql = sql & .DELCLS & "', '"        ' 削除区分
        sql = sql & .KRPROCCD & "', '"      ' 管理工程コード
        sql = sql & .PROCCD & "', '"        ' 工程コード
        sql = sql & .LPKRPROCCD & "', '"    ' 最終通過管理工程
        sql = sql & .LASTPASS & "', '"      ' 最終通過工程
        sql = sql & .RPHINBAN & "', "       ' ねらい品番
        sql = sql & .RPREVNUM & ", '"       ' ねらい品番製品番号改訂番号
        sql = sql & .RPFACT & "', '"        ' ねらい品番工場
        sql = sql & .RPOPCOND & "', '"      ' ねらい品番操業条件
        sql = sql & .PRODCOND & "', '"      ' 製作条件
        sql = sql & .PGID & "', "           ' ＰＧ－ＩＤ
        sql = sql & .UPLENGTH & ", "        ' 引上げ長さ
        sql = sql & .TOPLENG & ", "         ' ＴＯＰ長さ
        sql = sql & .BODYLENG & ", "        ' 直胴長さ
        sql = sql & .BOTLENG & ", "         ' ＢＯＴ長さ
        sql = sql & .FREELENG & ", "        ' フリー長
        sql = sql & .DIAMETER & ", "        ' 直径
        sql = sql & .CHARGE & ", '"         ' チャージ量
        sql = sql & .SEED & "', '"          ' シード
        sql = sql & .ADDDPCLS & "', "       ' 追加ドープ種類
        sql = sql & .ADDDPPOS & ", "        ' 追加ドープ位置
        sql = sql & .ADDDPVAL & ", "        ' 追加ドープ量
        sql = sql & "sysdate, "             ' 登録日付
        sql = sql & "sysdate, "             ' 更新日付
        sql = sql & "'0', "                 ' 送信フラグ
        sql = sql & "sysdate)"              ' 送信日付
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_CryInf_Ins = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_CryInf_Ins = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_CryInf_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :結晶情報の更新
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型               ,説明
'      　　:CryInf　　　,I  ,typ_TBCME037   　,結晶情報
'      　　:戻り値      ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'-------使用しないほうがよい（蔵本）---------
Public Function DBDRV_CryInf_Upd(CryInf As typ_TBCME037) As FUNCTION_RETURN

    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CryInf_Upd"

    '' 結晶情報の更新
    With CryInf
        sql = "update TBCME037 set "
        sql = sql & "CRYNUM='" & .CRYNUM & "', "            ' 結晶番号
        sql = sql & "DELCLS='" & .DELCLS & "', "            ' 削除区分
        sql = sql & "KRPROCCD='" & .KRPROCCD & "', "        ' 管理工程コード
        sql = sql & "PROCCD='" & .PROCCD & "', "            ' 工程コード
        sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "    ' 最終通過管理工程
        sql = sql & "LASTPASS='" & .LASTPASS & "', "        ' 最終通過工程
        sql = sql & "RPHINBAN='" & .RPHINBAN & "', "        ' ねらい品番
        sql = sql & "RPREVNUM=" & .RPREVNUM & ", "          ' ねらい品番製品番号改訂番号
        sql = sql & "RPFACT='" & .RPFACT & "', "            ' ねらい品番工場
        sql = sql & "RPOPCOND='" & .RPOPCOND & "', "        ' ねらい品番操業条件
        sql = sql & "PRODCOND='" & .PRODCOND & "', "        ' 製作条件
        sql = sql & "PGID='" & .PGID & "', "                ' ＰＧ－ＩＤ
        sql = sql & "UPLENGTH=" & .UPLENGTH & ", "          ' 引上げ長さ
        sql = sql & "TOPLENG=" & .TOPLENG & ", "            ' ＴＯＰ長さ
        sql = sql & "BODYLENG=" & .BODYLENG & ", "          ' 直胴長さ
        sql = sql & "BOTLENG=" & .BOTLENG & ", "            ' ＢＯＴ長さ
        sql = sql & "FREELENG=" & .FREELENG & ", "          ' フリー長
        sql = sql & "DIAMETER=" & .DIAMETER & ", "          ' 直径
        sql = sql & "CHARGE=" & .CHARGE & ", "              ' チャージ量
        sql = sql & "SEED='" & .SEED & "', "                ' シード
        sql = sql & "ADDDPCLS='" & .ADDDPCLS & "', "        ' 追加ドープ種類
        sql = sql & "ADDDPPOS=" & .ADDDPPOS & ", "          ' 追加ドープ位置
        sql = sql & "ADDDPVAL=" & .ADDDPVAL & ", "          ' 追加ドープ量
        sql = sql & "UPDDATE=sysdate, "                     ' 更新日付
        sql = sql & "SENDFLAG='0', "                        ' 送信フラグ
        sql = sql & "SENDDATE=sysdate"                      ' 送信日付
        sql = sql & " where CRYNUM='" & .CRYNUM & "'"
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_CryInf_Upd = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_CryInf_Upd = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_CryInf_Upd = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :ブロック管理の挿入／更新
'ﾊﾟﾗﾒｰﾀ　　:変数名           ,IO ,型               ,説明
'      　　:BlockMngOld　　　,I  ,typ_TBCME040   　,ブロック管理（旧）
'      　　:BlockMngNew　　　,I  ,typ_TBCME040   　,ブロック管理（新）
'      　　:戻り値           ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :古いレコードをみて更新か挿入かを判別する
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_BlockMng_UpdIns(BlockMngOld() As typ_TBCME040, BlockMngNew() As typ_TBCME040) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_BlockMng_UpdIns"

    DBDRV_BlockMng_UpdIns = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(BlockMngNew)
        With BlockMngNew(i)
            lFlg = False
            For j = 1 To UBound(BlockMngOld)
                If BlockMngOld(j).CRYNUM = .CRYNUM And _
                   BlockMngOld(j).INGOTPOS = .INGOTPOS Then
                    '' ブロック管理テーブルの更新
                    sql = "update TBCME040 set "
                    sql = sql & "CRYNUM='" & .CRYNUM & "', "                    ' 結晶番号
                    sql = sql & "INGOTPOS=" & .INGOTPOS & ", "                  ' 結晶内開始位置
                    sql = sql & "LENGTH=" & .Length & ", "                      ' 長さ
                    sql = sql & "REALLEN=" & .REALLEN & ", "                    ' 実長さ
                    sql = sql & "BLOCKID='" & .BLOCKID & "', "                  ' ブロックID
                    sql = sql & "KRPROCCD='" & .KRPROCCD & "', "                ' 現在管理工程
                    sql = sql & "NOWPROC='" & .NOWPROC & "', "                  ' 現在工程
                    sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "            ' 最終通過管理工程
                    sql = sql & "LASTPASS='" & .LASTPASS & "', "                ' 最終通過工程
                    sql = sql & "DELCLS='" & .DELCLS & "', "                    ' 削除区分
                    sql = sql & "LSTATCLS='" & .LSTATCLS & "', "                ' 最終状態区分
                    sql = sql & "RSTATCLS='" & .RSTATCLS & "', "                ' 流動状態区分
                    sql = sql & "HOLDCLS='" & .HOLDCLS & "', "                  ' ホールド区分
                    sql = sql & "BDCAUS='" & .BDCAUS & "', "                    ' 不良理由
                    sql = sql & "UPDDATE=sysdate, "                             ' 更新日付
                    sql = sql & "SUMMITSENDFLAG='" & .SUMMITSENDFLAG & "', "    ' SUMMIT送信フラグ
                    sql = sql & "SENDFLAG='0', "                                ' 送信フラグ
                    sql = sql & "SENDDATE=sysdate "                             ' 送信日付
                    sql = sql & "where CRYNUM='" & .CRYNUM & "' "
                    sql = sql & "and INGOTPOS=" & .INGOTPOS
                    '' WriteDBLog sql
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        DBDRV_BlockMng_UpdIns = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                    lFlg = True
                    Exit For
                End If
            Next j

            If lFlg <> True Then
                '' ブロック管理テーブルの挿入
                sql = "insert into TBCME040 ("
                sql = sql & "CRYNUM, "              ' 結晶番号
                sql = sql & "INGOTPOS, "            ' 結晶内開始位置
                sql = sql & "LENGTH, "              ' 長さ
                sql = sql & "REALLEN, "             ' 実長さ
                sql = sql & "BLOCKID, "             ' ブロックID
                sql = sql & "KRPROCCD, "            ' 現在管理工程
                sql = sql & "NOWPROC, "             ' 現在工程
                sql = sql & "LPKRPROCCD, "          ' 最終通過管理工程
                sql = sql & "LASTPASS, "            ' 最終通過工程
                sql = sql & "DELCLS, "              ' 削除区分
                sql = sql & "LSTATCLS, "            ' 最終状態区分
                sql = sql & "RSTATCLS, "            ' 流動状態区分
                sql = sql & "HOLDCLS, "             ' ホールド区分
                sql = sql & "BDCAUS, "              ' 不良理由
                sql = sql & "REGDATE, "             ' 登録日付
                sql = sql & "UPDDATE, "             ' 更新日付
                sql = sql & "SUMMITSENDFLAG, "      ' SUMMIT送信フラグ
                sql = sql & "SENDFLAG, "            ' 送信フラグ
                sql = sql & "SENDDATE)"             ' 送信日付
                sql = sql & " values ('"
                sql = sql & .CRYNUM & "', "         ' 結晶番号
                sql = sql & .INGOTPOS & ", "        ' 結晶内開始位置
                sql = sql & .Length & ", "          ' 長さ
                sql = sql & .REALLEN & ", '"        ' 実長さ
                sql = sql & .BLOCKID & "', '"       ' ブロックID
                sql = sql & .KRPROCCD & "', '"      ' 現在管理工程
                sql = sql & .NOWPROC & "', '"       ' 現在工程
                sql = sql & .LPKRPROCCD & "', '"    ' 最終通過管理工程
                sql = sql & .LASTPASS & "', '"      ' 最終通過工程
                sql = sql & .DELCLS & "', '"        ' 削除区分
                sql = sql & .LSTATCLS & "', '"      ' 最終状態区分
                sql = sql & .RSTATCLS & "', '"      ' 流動状態区分
                sql = sql & .HOLDCLS & "', '"       ' ホールド区分
                sql = sql & .BDCAUS & "', "         ' 不良理由
                sql = sql & "sysdate, "             ' 登録日付
                sql = sql & "sysdate, '"            ' 更新日付
                sql = sql & .SUMMITSENDFLAG & "', " ' SUMMIT送信フラグ
                sql = sql & "'0', "                 ' 送信フラグ
                sql = sql & "sysdate)"              ' 送信日付
                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_BlockMng_UpdIns = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_BlockMng_UpdIns = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :ブロック管理の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:BlockMng　　　,I  ,typ_TBCME040   　,ブロック管理
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_BlockMng_Ins(BlockMng As typ_TBCME040) As FUNCTION_RETURN

    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_BlockMng_Ins"

    With BlockMng
        sql = "insert into TBCME040 ("
        sql = sql & "CRYNUM, "              ' 結晶番号
        sql = sql & "INGOTPOS, "            ' 結晶内開始位置
        sql = sql & "LENGTH, "              ' 長さ
        sql = sql & "REALLEN, "             ' 実長さ
        sql = sql & "BLOCKID, "             ' ブロックID
        sql = sql & "KRPROCCD, "            ' 現在管理工程
        sql = sql & "NOWPROC, "             ' 現在工程
        sql = sql & "LPKRPROCCD, "          ' 最終通過管理工程
        sql = sql & "LASTPASS, "            ' 最終通過工程
        sql = sql & "DELCLS, "              ' 削除区分
        sql = sql & "LSTATCLS, "            ' 最終状態区分
        sql = sql & "RSTATCLS, "            ' 流動状態区分
        sql = sql & "HOLDCLS, "             ' ホールド区分
        sql = sql & "BDCAUS, "              ' 不良理由
        sql = sql & "REGDATE, "             ' 登録日付
        sql = sql & "UPDDATE, "             ' 更新日付
        sql = sql & "SUMMITSENDFLAG, "      ' SUMMIT送信フラグ
        sql = sql & "SENDFLAG, "            ' 送信フラグ
        sql = sql & "SENDDATE)"             ' 送信日付
        sql = sql & " values ('"
        sql = sql & .CRYNUM & "', "         ' 結晶番号
        sql = sql & .INGOTPOS & ", "        ' 結晶内開始位置
        sql = sql & .Length & ", "          ' 長さ
        sql = sql & .REALLEN & ", '"        ' 実長さ
        sql = sql & .BLOCKID & "', '"       ' ブロックID
        sql = sql & .KRPROCCD & "', '"      ' 現在管理工程
        sql = sql & .NOWPROC & "', '"       ' 現在工程
        sql = sql & .LPKRPROCCD & "', '"    ' 最終通過管理工程
        sql = sql & .LASTPASS & "', '"      ' 最終通過工程
        sql = sql & .DELCLS & "', '"        ' 削除区分
        sql = sql & .LSTATCLS & "', '"      ' 最終状態区分
        sql = sql & .RSTATCLS & "', '"      ' 流動状態区分
        sql = sql & .HOLDCLS & "', '"       ' ホールド区分
        sql = sql & .BDCAUS & "', "         ' 不良理由
        sql = sql & "sysdate, "             ' 登録日付
        sql = sql & "sysdate, "             ' 更新日付
        sql = sql & "'0', "                 ' SUMMIT送信フラグ
        sql = sql & "'0', "                 ' 送信フラグ
        sql = sql & "sysdate)"              ' 送信日付
    End With
    '' WriteDBLog sql
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_BlockMng_Ins = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_BlockMng_Ins = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_BlockMng_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :ブロック管理の更新
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:BlockMng　　　,I  ,typ_TBCME040   　,ブロック管理
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'-------使用しないほうがよい（蔵本）---------
Public Function DBDRV_BlockMng_Upd(BlockMng As typ_TBCME040) As FUNCTION_RETURN

    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_BlockMng_Upd"

    '' ブロック管理テーブルの更新
    With BlockMng
        sql = "update TBCME040 set "
        sql = sql & "LENGTH=" & .Length & ", "              ' 長さ
        sql = sql & "REALLEN=" & .REALLEN & ", "            ' 実長さ
        sql = sql & "BLOCKID='" & .BLOCKID & "', "          ' ブロックID
        sql = sql & "KRPROCCD='" & .KRPROCCD & "', "        ' 現在管理工程
        sql = sql & "NOWPROC='" & .NOWPROC & "', "          ' 現在工程
        sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "    ' 最終通過管理工程
        sql = sql & "LASTPASS='" & .LASTPASS & "', "        ' 最終通過工程
        sql = sql & "DELCLS='" & .DELCLS & "',"             ' 削除区分
        sql = sql & "LSTATCLS='" & .LSTATCLS & "', "        ' 最終状態区分
        sql = sql & "RSTATCLS='" & .RSTATCLS & "', "        ' 流動状態区分
        sql = sql & "HOLDCLS='" & .HOLDCLS & "', "          ' ホールド区分
        sql = sql & "BDCAUS='" & .BDCAUS & "', "            ' 不良理由
        sql = sql & "UPDDATE=sysdate, "                     ' 更新日付
        sql = sql & "SUMMITSENDFLAG='0', "                  ' SUMMIT送信フラグ
        sql = sql & "SENDFLAG='0' "                        ' 送信フラグ
        sql = sql & "where CRYNUM='" & .CRYNUM & "' "
        sql = sql & "and INGOTPOS=" & .INGOTPOS
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_BlockMng_Upd = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_BlockMng_Upd = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_BlockMng_Upd = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :品番管理の挿入／更新
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:HinbanOld　　　,I  ,typ_TBCME041   　,品番管理（旧）
'      　　:HinbanNew　　　,I  ,typ_TBCME041   　,品番管理（新）
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :古いレコードをみて更新か挿入かを判別する
'履歴      :2001/07/12  作成 蔵本
'      　　:2001/11/06  修正 野村
Public Function DBDRV_Hinban_UpdIns(HinbanOld() As typ_TBCME041, HinbanNew() As typ_TBCME041) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long
    Dim nOld As Long
    Dim HIN As tFullHinban

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_Hinban_UpdIns"

    DBDRV_Hinban_UpdIns = FUNCTION_RETURN_SUCCESS

    nOld = UBound(HinbanOld)
    For i = 1 To UBound(HinbanNew)
        With HinbanNew(i)
            For j = 1 To nOld
                If (.INGOTPOS = HinbanOld(j).INGOTPOS) _
                  And (.Length = HinbanOld(j).Length) _
                  And (.hinban = HinbanOld(j).hinban) _
                  And (.REVNUM = HinbanOld(j).REGDATE) _
                  And (.factory = HinbanOld(j).factory) _
                  And (.opecond = HinbanOld(j).opecond) Then
                    '全く同内容のレコードがすでに有る
                    Exit For
                End If
           Next
           If j > nOld Then '同内容のレコードはなかった
                '何か変更があれば、その範囲の品番を置き換える
                HIN.hinban = .hinban
                HIN.mnorevno = .REVNUM
                HIN.factory = .factory
                HIN.opecond = .opecond
                If ChangeAreaHinban(.CRYNUM, .INGOTPOS, .Length, HIN) = FUNCTION_RETURN_FAILURE Then
                    DBDRV_Hinban_UpdIns = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_Hinban_UpdIns = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :品番管理の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:HinbanNew　　　,I  ,typ_TBCME041   　,品番管理
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_Hinban_Ins(HinbanNew() As typ_TBCME041) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_Hinban_Ins"

    DBDRV_Hinban_Ins = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(HinbanNew)
        With HinbanNew(i)
            sql = "insert into TBCME041 ("
            sql = sql & "CRYNUM, "          ' 結晶番号
            sql = sql & "INGOTPOS, "        ' 結晶内開始位置
            sql = sql & "HINBAN, "          ' 品番
            sql = sql & "REVNUM, "          ' 製品番号改訂番号
            sql = sql & "FACTORY, "         ' 工場
            sql = sql & "OPECOND, "         ' 操業条件
            sql = sql & "LENGTH, "          ' 長さ
            sql = sql & "REGDATE, "         ' 登録日付
            sql = sql & "UPDDATE, "         ' 更新日付
            sql = sql & "SENDFLAG, "        ' 送信フラグ
            sql = sql & "SENDDATE)"         ' 送信日付
            sql = sql & " values ('"
            sql = sql & .CRYNUM & "', "     ' 結晶番号
            sql = sql & .INGOTPOS & ", '"   ' 結晶内開始位置
            sql = sql & .hinban & "', "     ' 品番
            sql = sql & .REVNUM & ", '"     ' 製品番号改訂番号
            sql = sql & .factory & "', '"   ' 工場
            sql = sql & .opecond & "', "    ' 操業条件
            sql = sql & .Length & ", "      ' 長さ
            sql = sql & "sysdate, "         ' 登録日付
            sql = sql & "sysdate, "         ' 更新日付
            sql = sql & "'0', "             ' 送信フラグ
            sql = sql & "sysdate)"          ' 送信日付
        End With
        '' WriteDBLog sql
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_Hinban_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_Hinban_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :結晶サンプル管理の挿入／更新
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型                  ,説明
'      　　:CrySmpOld　　　,I  ,typ_XSDCS   　      ,新サンプル管理（ブロック）（旧）
'      　　:CrySmpNew　　　,I  ,typ_XSDCS   　      ,新サンプル管理（ブロック）（新）
'      　　:戻り値         ,O  ,FUNCTION_RETURN　   ,書き込みの成否
'説明      :古いレコードをみて更新か挿入かを判別する
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_CrySmp_UpdIns(CrySmpOld() As typ_XSDCS, CrySmpNew() As typ_XSDCS) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CrySmp_UpdIns"

    DBDRV_CrySmp_UpdIns = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(CrySmpNew)
        With CrySmpNew(i)
            lFlg = False
            For j = 1 To UBound(CrySmpOld)
                If CrySmpOld(j).XTALCS = .XTALCS And _
                   CrySmpOld(j).INPOSCS = .INPOSCS And _
                   CrySmpOld(j).SMPKBNCS = .SMPKBNCS Then
'                    sql = "update TBCME043 set "
'                    sql = sql & "HINBAN='" & .HINBAN & "', "        ' 品番
'                    sql = sql & "REVNUM=" & .REVNUM & ", "          ' 製品番号改訂番号
'                    sql = sql & "FACTORY='" & .factory & "', "      ' 工場
'                    sql = sql & "OPECOND='" & .opecond & "', "      ' 操業条件
'                    sql = sql & "KTKBN='" & .KTKBN & "', "          ' 確定区分
'                    sql = sql & "SMPLNO='" & Abs(.SMPLNO) & "', "   ' サンプルＮｏ
'                    sql = sql & "CRYINDRS='" & .CRYINDRS & "', "    ' 結晶検査指示（Rs)
'                    sql = sql & "CRYINDOI='" & .CRYINDOI & "', "    ' 結晶検査指示（Oi)
'                    sql = sql & "CRYINDB1='" & .CRYINDB1 & "', "    ' 結晶検査指示（B1)
'                    sql = sql & "CRYINDB2='" & .CRYINDB2 & "', "    ' 結晶検査指示（B2)
'                    sql = sql & "CRYINDB3='" & .CRYINDB3 & "', "    ' 結晶検査指示（B3)
'                    sql = sql & "CRYINDL1='" & .CRYINDL1 & "', "    ' 結晶検査指示（L1)
'                    sql = sql & "CRYINDL2='" & .CRYINDL2 & "', "    ' 結晶検査指示（L2)
'                    sql = sql & "CRYINDL3='" & .CRYINDL3 & "', "    ' 結晶検査指示（L3)
'                    sql = sql & "CRYINDL4='" & .CRYINDL4 & "', "    ' 結晶検査指示（L4)
'                    sql = sql & "CRYINDCS='" & .CRYINDCS & "', "    ' 結晶検査指示（Cs)
'                    sql = sql & "CRYINDGD='" & .CRYINDGD & "', "    ' 結晶検査指示（GD)
'                    sql = sql & "CRYINDT='" & .CRYINDT & "', "      ' 結晶検査指示（T)
'                    sql = sql & "CRYINDEP='" & .CRYINDEP & "', "    ' 結晶検査指示（EPD)
'                    sql = sql & "CRYRESRS='" & .CRYRESRS & "', "    ' 結晶検査実績（Rs)
'                    sql = sql & "CRYRESOI='" & .CRYRESOI & "', "    ' 結晶検査実績（Oi)
'                    sql = sql & "CRYRESB1='" & .CRYRESB1 & "', "    ' 結晶検査実績（B1)
'                    sql = sql & "CRYRESB2='" & .CRYRESB2 & "', "    ' 結晶検査実績（B2)
'                    sql = sql & "CRYRESB3='" & .CRYRESB3 & "', "    ' 結晶検査実績（B3)
'                    sql = sql & "CRYRESL1='" & .CRYRESL1 & "', "    ' 結晶検査実績（L1)
'                    sql = sql & "CRYRESL2='" & .CRYRESL2 & "', "    ' 結晶検査実績（L2)
'                    sql = sql & "CRYRESL3='" & .CRYRESL3 & "', "    ' 結晶検査実績（L3)
'                    sql = sql & "CRYRESL4='" & .CRYRESL4 & "', "    ' 結晶検査実績（L4)
'                    sql = sql & "CRYRESCS='" & .CRYRESCS & "', "    ' 結晶検査実績（Cs)
'                    sql = sql & "CRYRESGD='" & .CRYRESGD & "', "    ' 結晶検査実績（GD)
'                    sql = sql & "CRYREST='" & .CRYREST & "', "      ' 結晶検査実績（T)
'                    sql = sql & "CRYRESEP='" & .CRYRESEP & "', "    ' 結晶検査実績（EPD)
'                    sql = sql & "SMPLNUM=" & .SMPLNUM & ", "        ' サンプル枚数
'                    sql = sql & "SMPLPAT='" & .SMPLPAT & "', "      ' サンプルパターン
'                    sql = sql & "UPDDATE=sysdate, "                 ' 更新日付
'                    sql = sql & "SENDFLAG='0' "                     ' 送信フラグ
'                    sql = sql & " where CRYNUM='" & .CRYNUM & "'"
'                    sql = sql & " and INGOTPOS=" & .INGOTPOS
'                    sql = sql & " and SMPKBN='" & .SMPKBN & "'"
                    sql = "update XSDCS set "
                    sql = sql & "CRYNUMCS='" & .CRYNUMCS & "', "            ' ブロックID
                    sql = sql & "TBKBNCS='" & .TBKBNCS & "', "              ' T/B区分
                    sql = sql & "REPSMPLIDCS='" & Abs(.REPSMPLIDCS) & "', " ' 代表サンプルID
                    sql = sql & "HINBCS='" & .HINBCS & "', "                ' 品番
                    sql = sql & "REVNUMCS=" & .REVNUMCS & ", "              ' 製品番号改訂番号
                    sql = sql & "FACTORYCS='" & .FACTORYCS & "', "          ' 工場
                    sql = sql & "OPECS='" & .OPECS & "', "                  ' 操業条件
                    sql = sql & "KTKBNCS='" & .KTKBNCS & "', "              ' 確定区分
                    sql = sql & "BLKKTFLAGCS='" & .BLKKTFLAGCS & "', "      ' ブロック確定フラグ
                    sql = sql & "CRYSMPLIDRSCS=" & .CRYSMPLIDRSCS & ", "    ' サンプルID(Rs)
                    sql = sql & "CRYSMPLIDRS1CS=" & .CRYSMPLIDRS1CS & ", "  ' 推定サンプルID1（Rs）
                    sql = sql & "CRYSMPLIDRS2CS=" & .CRYSMPLIDRS2CS & ", "  ' 推定サンプルID2（Rs）
                    sql = sql & "CRYSMPLIDOICS=" & .CRYSMPLIDOICS & ", "    ' サンプルID(Oi)
                    sql = sql & "CRYSMPLIDB1CS=" & .CRYSMPLIDB1CS & ", "    ' サンプルID(B1)
                    sql = sql & "CRYSMPLIDB2CS=" & .CRYSMPLIDB2CS & ", "    ' サンプルID(B2)
                    sql = sql & "CRYSMPLIDB3CS=" & .CRYSMPLIDB3CS & ", "    ' サンプルID(B3)
                    sql = sql & "CRYSMPLIDL1CS=" & .CRYSMPLIDL1CS & ", "    ' サンプルID(L1)
                    sql = sql & "CRYSMPLIDL2CS=" & .CRYSMPLIDL2CS & ", "    ' サンプルID(L2)
                    sql = sql & "CRYSMPLIDL3CS=" & .CRYSMPLIDL3CS & ", "    ' サンプルID(L3)
                    sql = sql & "CRYSMPLIDL4CS=" & .CRYSMPLIDL4CS & ", "    ' サンプルID(L4)
                    sql = sql & "CRYSMPLIDCSCS=" & .CRYSMPLIDCSCS & ", "    ' サンプルID(Cs)
                    sql = sql & "CRYSMPLIDGDCS=" & .CRYSMPLIDGDCS & ", "    ' サンプルID(GD)
                    sql = sql & "CRYSMPLIDTCS=" & .CRYSMPLIDTCS & ", "      ' サンプルID(T)
                    sql = sql & "CRYSMPLIDEPCS=" & .CRYSMPLIDEPCS & ", "    ' サンプルID(EPD)
                    'Cng Start 2011/03/31 SMPK Y.Hitomi
'                    'Add Start 2010/12/13 SMPK Miyata
'                    sql = sql & "CRYSMPLIDCCS=" & .CRYSMPLIDCCS & ", "      ' サンプルID(C)
'                    sql = sql & "CRYSMPLIDCJCS=" & .CRYSMPLIDCJCS & ", "    ' サンプルID(CJ)
'                    sql = sql & "CRYSMPLIDCJLTCS=" & .CRYSMPLIDCJLTCS & ", " ' サンプルID(CJLT)
'                    sql = sql & "CRYSMPLIDCJ2CS=" & .CRYSMPLIDCJ2CS & ", "  ' サンプルID(CJ2)
'                    'Add End   2010/12/13 SMPK Miyata
                    If .CRYSMPLIDCCS <> 0 Then
                        sql = sql & "CRYSMPLIDCCS=" & .CRYSMPLIDCCS & ", "      ' サンプルID(C)
                    End If
                    If .CRYSMPLIDCJCS <> 0 Then
                        sql = sql & "CRYSMPLIDCJCS=" & .CRYSMPLIDCJCS & ", "    ' サンプルID(CJ)
                    End If
                    If .CRYSMPLIDCJLTCS <> 0 Then
                        sql = sql & "CRYSMPLIDCJLTCS=" & .CRYSMPLIDCJLTCS & ", " ' サンプルID(CJLT)
                    End If
                    If .CRYSMPLIDCJ2CS <> 0 Then
                        sql = sql & "CRYSMPLIDCJ2CS=" & .CRYSMPLIDCJ2CS & ", "  ' サンプルID(CJ2)
                    End If
                    'Add End   2011/03/31 SMPK Y.Hitomi
                    sql = sql & "CRYINDRSCS='" & .CRYINDRSCS & "', "        ' 状態FLG（Rs)
                    sql = sql & "CRYINDOICS='" & .CRYINDOICS & "', "        ' 状態FLG（Oi)
                    sql = sql & "CRYINDB1CS='" & .CRYINDB1CS & "', "        ' 状態FLG（B1)
                    sql = sql & "CRYINDB2CS='" & .CRYINDB2CS & "', "        ' 状態FLG（B2)
                    sql = sql & "CRYINDB3CS='" & .CRYINDB3CS & "', "        ' 状態FLG（B3)
                    sql = sql & "CRYINDL1CS='" & .CRYINDL1CS & "', "        ' 状態FLG（L1)
                    sql = sql & "CRYINDL2CS='" & .CRYINDL2CS & "', "        ' 状態FLG（L2)
                    sql = sql & "CRYINDL3CS='" & .CRYINDL3CS & "', "        ' 状態FLG（L3)
                    sql = sql & "CRYINDL4CS='" & .CRYINDL4CS & "', "        ' 状態FLG（L4)
                    sql = sql & "CRYINDCSCS='" & .CRYINDCSCS & "', "        ' 状態FLG（Cs)
                    sql = sql & "CRYINDGDCS='" & .CRYINDGDCS & "', "        ' 状態FLG（GD)
                    sql = sql & "CRYINDTCS='" & .CRYINDTCS & "', "          ' 状態FLG（T)
                    sql = sql & "CRYINDEPCS='" & .CRYINDEPCS & "', "        ' 状態FLG（EPD)
                    'Cng Start 2011/03/31 SMPK Y.Hitomi
'                    'Add Start 2010/12/13 SMPK Miyata
'                    sql = sql & "CRYINDCCS='" & .CRYINDCCS & "', "          ' 状態FLG（C)
'                    sql = sql & "CRYINDCJCS='" & .CRYINDCJCS & "', "        ' 状態FLG（CJ)
'                    sql = sql & "CRYINDCJLTCS='" & .CRYINDCJLTCS & "', "    ' 状態FLG（CJLT)
'                    sql = sql & "CRYINDCJ2CS='" & .CRYINDCJ2CS & "', "      ' 状態FLG（CJ2)
'                    'Add End   2010/12/13 SMPK Miyata
                    If .CRYINDCCS <> "" And left(.CRYINDCCS, 1) <> vbNullChar Then
                        sql = sql & "CRYINDCCS='" & .CRYINDCCS & "', "          ' 状態FLG（C)
                    End If
                    If .CRYINDCJCS <> "" And left(.CRYINDCJCS, 1) <> vbNullChar Then
                        sql = sql & "CRYINDCJCS='" & .CRYINDCJCS & "', "        ' 状態FLG（CJ)
                    End If
                    If .CRYINDCJLTCS <> "" And left(.CRYINDCJLTCS, 1) <> vbNullChar Then
                        sql = sql & "CRYINDCJLTCS='" & .CRYINDCJLTCS & "', "    ' 状態FLG（CJLT)
                    End If
                    If .CRYINDCJ2CS <> "" And left(.CRYINDCJ2CS, 1) <> vbNullChar Then
                        sql = sql & "CRYINDCJ2CS='" & .CRYINDCJ2CS & "', "      ' 状態FLG（CJ2)
                    End If
                    'Cng End 2011/03/31 SMPK Y.Hitomi
                    sql = sql & "CRYRESRS1CS='" & .CRYRESRS1CS & "', "      ' 実績FLG1（Rs)
                    sql = sql & "CRYRESRS2CS='" & .CRYRESRS2CS & "', "      ' 実績FLG2（Rs)
                    sql = sql & "CRYRESOICS='" & .CRYRESOICS & "', "        ' 実績FLG（Oi)
                    sql = sql & "CRYRESB1CS='" & .CRYRESB1CS & "', "        ' 実績FLG（B1)
                    sql = sql & "CRYRESB2CS='" & .CRYRESB2CS & "', "        ' 実績FLG（B2)
                    sql = sql & "CRYRESB3CS='" & .CRYRESB3CS & "', "        ' 実績FLG（B3)
                    sql = sql & "CRYRESL1CS='" & .CRYRESL1CS & "', "        ' 実績FLG（L1)
                    sql = sql & "CRYRESL2CS='" & .CRYRESL2CS & "', "        ' 実績FLG（L2)
                    sql = sql & "CRYRESL3CS='" & .CRYRESL3CS & "', "        ' 実績FLG（L3)
                    sql = sql & "CRYRESL4CS='" & .CRYRESL4CS & "', "        ' 実績FLG（L4)
                    sql = sql & "CRYRESCSCS='" & .CRYRESCSCS & "', "        ' 実績FLG（Cs)
                    sql = sql & "CRYRESGDCS='" & .CRYRESGDCS & "', "        ' 実績FLG（GD)
                    sql = sql & "CRYRESTCS='" & .CRYRESTCS & "', "          ' 実績FLG（T)
                    sql = sql & "CRYRESEPCS='" & .CRYRESEPCS & "', "        ' 実績FLG（EPD)
'                    'Add Start 2010/12/13 SMPK Miyata
'                    sql = sql & "CRYRESCCS='" & .CRYRESCCS & "', "          ' 実績FLG（C)
'                    sql = sql & "CRYRESCJCS='" & .CRYRESCJCS & "', "        ' 実績FLG（CJ)
'                    sql = sql & "CRYRESCJLTCS='" & .CRYRESCJLTCS & "', "    ' 実績FLG（CJLT)
'                    sql = sql & "CRYRESCJ2CS='" & .CRYRESCJ2CS & "', "      ' 実績FLG（CJ2)
'                    'Add End   2010/12/13 SMPK Miyata
                    'Cng Start 2011/03/31 SMPK Y.Hitomi
                    If .CRYRESCCS <> "" And left(.CRYRESCCS, 1) <> vbNullChar Then
                        sql = sql & "CRYRESCCS='" & .CRYRESCCS & "', "          ' 実績FLG（C)
                    End If
                    If .CRYRESCJCS <> "" And left(.CRYRESCJCS, 1) <> vbNullChar Then
                        sql = sql & "CRYRESCJCS='" & .CRYRESCJCS & "', "        ' 実績FLG（CJ)
                    End If
                    If .CRYRESCJLTCS <> "" And left(.CRYRESCJLTCS, 1) <> vbNullChar Then
                        sql = sql & "CRYRESCJLTCS='" & .CRYRESCJLTCS & "', "    ' 実績FLG（CJLT)
                    End If
                    If .CRYRESCJ2CS <> "" And left(.CRYRESCJ2CS, 1) <> vbNullChar Then
                        sql = sql & "CRYRESCJ2CS='" & .CRYRESCJ2CS & "', "      ' 実績FLG（CJ2)
                    End If
                    'Add End   Cng Start 2011/03/31 SMPK Y.Hitomi
                    
                    sql = sql & "SMPLNUMCS=" & .SMPLNUMCS & ", "            ' サンプル枚数
                    sql = sql & "SMPLPATCS='" & .SMPLPATCS & "', "          ' サンプルパターン
                    sql = sql & "LIVKCS='" & .LIVKCS & "', "                ' 生死区分
                    sql = sql & "KSTAFFCS='" & .KSTAFFCS & "', "            ' 更新社員ID
                    sql = sql & "KDAYCS=sysdate, "                          ' 更新日付
                    sql = sql & "SNDKCS='0' "                               ' 送信フラグ
                    '>>>>> X線測定追加対応 2009/07/28 SETsw kubota ---------------
                    If .CRYSMPLIDXCS <> 0 Then
                        sql = sql & ",CRYSMPLIDXCS=" & .CRYSMPLIDXCS        ' サンプルID(X線)
                    End If
                    If .CRYINDXCS <> "" And left(.CRYINDXCS, 1) <> vbNullChar Then
                        sql = sql & ",CRYINDXCS='" & .CRYINDXCS & "'"       ' 状態FLG（X線)
                    End If
                    If .CRYRESXCS <> "" And left(.CRYRESXCS, 1) <> vbNullChar Then
                        sql = sql & ",CRYRESXCS='" & .CRYRESXCS & "'"       ' 実績FLG（X線)
                    End If
                    '<<<<< X線測定追加対応 2009/07/28 SETsw kubota ---------------
                    '>>>>> 抵抗狙い位置対応 2009/11/06 SETsw kubota ---------------
                    If .QCKBNCS <> "" And left(.QCKBNCS, 1) <> vbNullChar Then
                        sql = sql & ",QCKBNCS='" & .QCKBNCS & "'"           ' (抵抗狙い位置)管理区分
                    End If
                    '<<<<< 抵抗狙い位置対応 2009/11/06 SETsw kubota ---------------
                    
                    '05/10/17 ooba START =====================================================>
                    '親ブロックID
                    If left(.RPCRYNUMCS, 1) <> vbNullChar And Trim(.RPCRYNUMCS) <> "" Then
                        sql = sql & ",RPCRYNUMCS='" & .RPCRYNUMCS & "' "
                    End If
                    '切断フラグ
                    If left(.CUTFLGCS, 1) <> vbNullChar And Trim(.CUTFLGCS) <> "" Then
                        sql = sql & ",CUTFLGCS='" & .CUTFLGCS & "' "
                    Else
                        sql = sql & ",CUTFLGCS=NULL "
                    End If
                    '05/10/17 ooba END =======================================================>
'' 09/03/02 FAE)akiyama start
'                    sql = sql & " where XTALCS='" & .XTALCS & "'"
                    sql = sql & " where CRYNUMCS LIKE '" & left(.XTALCS, 9) & "%'"
'' 09/03/02 FAE)akiyama end
                    sql = sql & " and INPOSCS=" & .INPOSCS
                    sql = sql & " and SMPKBNCS='" & .SMPKBNCS & "'"

                    '' WriteDBLog sql
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        DBDRV_CrySmp_UpdIns = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                    lFlg = True
                    Exit For
                End If
            Next j

            If lFlg <> True Then
'                sql = "insert into TBCME043 ("
'                sql = sql & "CRYNUM, "          ' 結晶番号
'                sql = sql & "INGOTPOS, "        ' 結晶内位置
'                sql = sql & "SMPKBN, "          ' サンプル区分
'                sql = sql & "SMPLNO, "          ' サンプルNo
'                sql = sql & "HINBAN, "          ' 品番
'                sql = sql & "REVNUM, "          ' 製品番号改訂番号
'                sql = sql & "FACTORY, "         ' 工場
'                sql = sql & "OPECOND, "         ' 操業条件
'                sql = sql & "KTKBN, "           ' 確定区分
'                sql = sql & "CRYINDRS, "        ' 結晶検査指示（Rs)
'                sql = sql & "CRYINDOI, "        ' 結晶検査指示（Oi)
'                sql = sql & "CRYINDB1, "        ' 結晶検査指示（B1)
'                sql = sql & "CRYINDB2, "        ' 結晶検査指示（B2)
'                sql = sql & "CRYINDB3, "        ' 結晶検査指示（B3)
'                sql = sql & "CRYINDL1, "        ' 結晶検査指示（L1)
'                sql = sql & "CRYINDL2, "        ' 結晶検査指示（L2)
'                sql = sql & "CRYINDL3, "        ' 結晶検査指示（L3)
'                sql = sql & "CRYINDL4, "        ' 結晶検査指示（L4)
'                sql = sql & "CRYINDCS, "        ' 結晶検査指示（Cs)
'                sql = sql & "CRYINDGD, "        ' 結晶検査指示（GD)
'                sql = sql & "CRYINDT, "         ' 結晶検査指示（T)
'                sql = sql & "CRYINDEP, "        ' 結晶検査指示（EPD)
'                sql = sql & "CRYRESRS, "        ' 結晶検査実績（Rs)
'                sql = sql & "CRYRESOI, "        ' 結晶検査実績（Oi)
'                sql = sql & "CRYRESB1, "        ' 結晶検査実績（B1)
'                sql = sql & "CRYRESB2, "        ' 結晶検査実績（B2)
'                sql = sql & "CRYRESB3, "        ' 結晶検査実績（B3)
'                sql = sql & "CRYRESL1, "        ' 結晶検査実績（L1)
'                sql = sql & "CRYRESL2, "        ' 結晶検査実績（L2)
'                sql = sql & "CRYRESL3, "        ' 結晶検査実績（L3)
'                sql = sql & "CRYRESL4, "        ' 結晶検査実績（L4)
'                sql = sql & "CRYRESCS, "        ' 結晶検査実績（Cs)
'                sql = sql & "CRYRESGD, "        ' 結晶検査実績（GD)
'                sql = sql & "CRYREST, "         ' 結晶検査実績（T)
'                sql = sql & "CRYRESEP, "        ' 結晶検査実績（EPD)
'                sql = sql & "SMPLNUM, "         ' サンプル枚数
'                sql = sql & "SMPLPAT, "         ' サンプルパターン
'                sql = sql & "REGDATE, "         ' 登録日付
'                sql = sql & "UPDDATE, "         ' 更新日付
'                sql = sql & "SENDFLAG, "        ' 送信フラグ
'                sql = sql & "SENDDATE)"         ' 送信日付
'                sql = sql & " values ('"
'                sql = sql & .CRYNUM & "', "
'                sql = sql & .INGOTPOS & ", '"   ' 結晶内位置
'                sql = sql & .SMPKBN & "', "     ' サンプル区分
'                sql = sql & Abs(.SMPLNO) & ", '"     ' サンプルNo
'                sql = sql & .HINBAN & "', "     ' 品番
'                sql = sql & .REVNUM & ", '"     ' 製品番号改訂番号
'                sql = sql & .factory & "', '"   ' 工場
'                sql = sql & .opecond & "', '"   ' 操業条件
'                sql = sql & .KTKBN & "', '"     ' 確定区分
'                sql = sql & .CRYINDRS & "', '"  ' 結晶検査指示（Rs)
'                sql = sql & .CRYINDOI & "', '"  ' 結晶検査指示（Oi)
'                sql = sql & .CRYINDB1 & "', '"  ' 結晶検査指示（B1)
'                sql = sql & .CRYINDB2 & "', '"  ' 結晶検査指示（B2)
'                sql = sql & .CRYINDB3 & "', '"  ' 結晶検査指示（B3)
'                sql = sql & .CRYINDL1 & "', '"  ' 結晶検査指示（L1)
'                sql = sql & .CRYINDL2 & "', '"  ' 結晶検査指示（L2)
'                sql = sql & .CRYINDL3 & "', '"  ' 結晶検査指示（L3)
'                sql = sql & .CRYINDL4 & "', '"  ' 結晶検査指示（L4)
'                sql = sql & .CRYINDCS & "', '"  ' 結晶検査指示（Cs)
'                sql = sql & .CRYINDGD & "', '"  ' 結晶検査指示（GD)
'                sql = sql & .CRYINDT & "', '"   ' 結晶検査指示（T)
'                sql = sql & .CRYINDEP & "', '"  ' 結晶検査指示（EPD)
'                sql = sql & .CRYRESRS & "', '"  ' 結晶検査実績（Rs)
'                sql = sql & .CRYRESOI & "', '"  ' 結晶検査実績（Oi)
'                sql = sql & .CRYRESB1 & "', '"  ' 結晶検査実績（B1)
'                sql = sql & .CRYRESB2 & "', '"  ' 結晶検査実績（B2)
'                sql = sql & .CRYRESB3 & "', '"  ' 結晶検査実績（B3)
'                sql = sql & .CRYRESL1 & "', '"  ' 結晶検査実績（L1)
'                sql = sql & .CRYRESL2 & "', '"  ' 結晶検査実績（L2)
'                sql = sql & .CRYRESL3 & "', '"  ' 結晶検査実績（L3)
'                sql = sql & .CRYRESL4 & "', '"  ' 結晶検査実績（L4)
'                sql = sql & .CRYRESCS & "', '"  ' 結晶検査実績（Cs)
'                sql = sql & .CRYRESGD & "', '"  ' 結晶検査実績（GD)
'                sql = sql & .CRYREST & "', '"   ' 結晶検査実績（T)
'                sql = sql & .CRYRESEP & "', "   ' 結晶検査実績（EPD)
'                sql = sql & .SMPLNUM & ", "     ' サンプル枚数
'                sql = sql & "' ', "             ' サンプルパターン
'                sql = sql & "sysdate, "
'                sql = sql & "sysdate, "
'                sql = sql & "'0', "
'                sql = sql & "sysdate)"
                sql = "insert into XSDCS ("
                sql = sql & "CRYNUMCS,"         'ブロックID
                sql = sql & "SMPKBNCS,"         'サンプル区分
                sql = sql & "TBKBNCS,"          'T/B区分
                sql = sql & "REPSMPLIDCS,"      '代表サンプルID
                sql = sql & "XTALCS,"           '結晶番号
                sql = sql & "INPOSCS,"          '結晶内位置
                sql = sql & "HINBCS,"           '品番
                sql = sql & "REVNUMCS,"         '製品番号改訂番号
                sql = sql & "FACTORYCS,"        '工場
                sql = sql & "OPECS,"            '操業番号
                sql = sql & "KTKBNCS,"          '確定区分
                sql = sql & "BLKKTFLAGCS,"      'ブロック確定フラグ
                sql = sql & "CRYSMPLIDRSCS,"    'サンプルID(Rs)
                sql = sql & "CRYSMPLIDRS1CS,"   '推定サンプルID1（Rs）
                sql = sql & "CRYSMPLIDRS2CS,"   '推定サンプルID2（Rs）
                sql = sql & "CRYINDRSCS,"       '状態FLG(Rs)
                sql = sql & "CRYRESRS1CS,"      '実績FLG1(Rs)
                sql = sql & "CRYRESRS2CS,"      '実績FLG2(Rs)
                sql = sql & "CRYSMPLIDOICS,"    'サンプルID（Oi）
                sql = sql & "CRYINDOICS,"       '状態FLG（Oi）
                sql = sql & "CRYRESOICS,"       '実績FLG（Oi）
                sql = sql & "CRYSMPLIDB1CS,"    'サンプルID（B1）
                sql = sql & "CRYINDB1CS,"       '状態FLG（B1）
                sql = sql & "CRYRESB1CS,"       '実績FLG（B1）
                sql = sql & "CRYSMPLIDB2CS,"    'サンプルID（B2）
                sql = sql & "CRYINDB2CS,"       '状態FLG（B2）
                sql = sql & "CRYRESB2CS,"       '実績FLG（B2）
                sql = sql & "CRYSMPLIDB3CS,"    'サンプルID（B3）
                sql = sql & "CRYINDB3CS,"       '状態FLG（B3）
                sql = sql & "CRYRESB3CS,"       '実績FLG（B3）
                sql = sql & "CRYSMPLIDL1CS,"    'サンプルID（L1）
                sql = sql & "CRYINDL1CS,"       '状態FLG（L1）
                sql = sql & "CRYRESL1CS,"       '実績FLG（L1）
                sql = sql & "CRYSMPLIDL2CS,"    'サンプルID（L2）
                sql = sql & "CRYINDL2CS,"       '状態FLG（L2）
                sql = sql & "CRYRESL2CS,"       '実績FLG（L2）
                sql = sql & "CRYSMPLIDL3CS,"    'サンプルID（L3）
                sql = sql & "CRYINDL3CS,"       '状態FLG（L3）
                sql = sql & "CRYRESL3CS,"       '実績FLG（L3）
                sql = sql & "CRYSMPLIDL4CS,"    'サンプルID（L4）
                sql = sql & "CRYINDL4CS,"       '状態FLG（L4）
                sql = sql & "CRYRESL4CS,"       '実績FLG（L4）
                sql = sql & "CRYSMPLIDCSCS,"    'サンプルID（CS）
                sql = sql & "CRYINDCSCS,"       '状態FLG（CS）
                sql = sql & "CRYRESCSCS,"       '実績FLG（CS）
                sql = sql & "CRYSMPLIDGDCS,"    'サンプルID（GD）
                sql = sql & "CRYINDGDCS,"       '状態FLG（GD）
                sql = sql & "CRYRESGDCS,"       '実績FLG（GD）
                sql = sql & "CRYSMPLIDTCS,"     'サンプルID（T）
                sql = sql & "CRYINDTCS,"        '状態FLG（T）
                sql = sql & "CRYRESTCS,"        '実績FLG（T）
                sql = sql & "CRYSMPLIDEPCS,"    'サンプルID（EPD）
                sql = sql & "CRYINDEPCS,"       '状態FLG（EPD）
                sql = sql & "CRYRESEPCS,"       '実績FLG（EPD）
                sql = sql & "CRYSMPLIDXCS,"     'サンプルID（X線）  'X線測定 2009/07/27追加 SETsw kubota
                sql = sql & "CRYINDXCS,"        '状態FLG（X線）
                sql = sql & "CRYRESXCS,"        '実績FLG（X線）
                'Add Start 2010/12/13 SMPK Miyata
                sql = sql & "CRYSMPLIDCCS,"     'サンプルID（C）
                sql = sql & "CRYINDCCS,"        '状態FLG（C）
                sql = sql & "CRYRESCCS,"        '実績FLG（C）
                sql = sql & "CRYSMPLIDCJCS,"    'サンプルID（CJ）
                sql = sql & "CRYINDCJCS,"       '状態FLG（CJ）
                sql = sql & "CRYRESCJCS,"       '実績FLG（CJ）
                sql = sql & "CRYSMPLIDCJLTCS,"  'サンプルID（CJLT）
                sql = sql & "CRYINDCJLTCS,"     '状態FLG（CJLT）
                sql = sql & "CRYRESCJLTCS,"     '実績FLG（CJLT）
                sql = sql & "CRYSMPLIDCJ2CS,"   'サンプルID（CJ2）
                sql = sql & "CRYINDCJ2CS,"      '状態FLG（CJ2）
                sql = sql & "CRYRESCJ2CS,"      '実績FLG（CJ2）
                'Add End   2010/12/13 SMPK Miyata
                sql = sql & "SMPLNUMCS,"        'サンプル枚数
                sql = sql & "SMPLPATCS,"        'サンプルパターン
                sql = sql & "LIVKCS,"           '生死区分
                sql = sql & "TSTAFFCS,"         '登録社員ID
                sql = sql & "TDAYCS,"           '登録日付
                sql = sql & "KSTAFFCS,"         '更新社員ID
                sql = sql & "KDAYCS,"           '更新日付
                sql = sql & "SNDKCS,"           '送信フラグ
'                sql = sql & "SNDDAYCS)"         '送信日付
                '05/10/17 ooba START =====================================================>
                sql = sql & "SNDDAYCS"          '送信日付
                '親ブロックID
                If left(.RPCRYNUMCS, 1) <> vbNullChar And Trim(.RPCRYNUMCS) <> "" Then
                    sql = sql & ",RPCRYNUMCS"
                End If
                '切断フラグ
                sql = sql & ",CUTFLGCS"
                sql = sql & ",QCKBNCS"          '管理区分       2009/11/06追加 SETsw kubota
                sql = sql & ")"
                '05/10/17 ooba END =======================================================>
                sql = sql & " values ('"
                sql = sql & .CRYNUMCS & "', '"          'ブロックID
                sql = sql & .SMPKBNCS & "', '"          'サンプル区分
                sql = sql & .TBKBNCS & "', "            'T/B区分
                sql = sql & .REPSMPLIDCS & ", '"        '代表サンプルID
                sql = sql & .XTALCS & "', "             '結晶番号
                sql = sql & .INPOSCS & ", '"            '結晶内位置
                sql = sql & .HINBCS & "', "             '品番
                sql = sql & .REVNUMCS & ", '"           '製品番号改訂番号
                sql = sql & .FACTORYCS & "', '"         '工場
                sql = sql & .OPECS & "', '"             '操業条件
                sql = sql & .KTKBNCS & "', '"           '確定区分
                sql = sql & .BLKKTFLAGCS & "', "        'ブロック確定フラグ
                sql = sql & .CRYSMPLIDRSCS & ", "       'サンプルID（Rs）
                sql = sql & .CRYSMPLIDRS1CS & ", "      '推定サンプルID1（Rs）
                sql = sql & .CRYSMPLIDRS2CS & ", '"     '推定サンプルID2（Rs）
                sql = sql & .CRYINDRSCS & "', '"        '状態FLG（Rs）
                sql = sql & .CRYRESRS1CS & "', '"       '実績FLG1（Rs）
                sql = sql & .CRYRESRS2CS & "', "        '実績FLG2（Rs）
                sql = sql & .CRYSMPLIDOICS & ", '"      'サンプルID（Oi）
                sql = sql & .CRYINDOICS & "', '"        '状態FLG（Oi）
                sql = sql & .CRYRESOICS & "', "         '実績FLG（Oi）
                sql = sql & .CRYSMPLIDB1CS & ", '"      'サンプルID（B1）
                sql = sql & .CRYINDB1CS & "', '"        '状態FLG（B1）
                sql = sql & .CRYRESB1CS & "', "         '実績FLG（B1）
                sql = sql & .CRYSMPLIDB2CS & ", '"      'サンプルID（B2）
                sql = sql & .CRYINDB2CS & "', '"        '状態FLG（B2）
                sql = sql & .CRYRESB2CS & "', "         '実績FLG（B2）
                sql = sql & .CRYSMPLIDB3CS & ", '"      'サンプルID（B3）
                sql = sql & .CRYINDB3CS & "', '"        '状態FLG（B3）
                sql = sql & .CRYRESB3CS & "', "         '実績FLG（B3）
                sql = sql & .CRYSMPLIDL1CS & ", '"      'サンプルID（L1）
                sql = sql & .CRYINDL1CS & "', '"        '状態FLG（L1）
                sql = sql & .CRYRESL1CS & "', "         '実績FLG（L1）
                sql = sql & .CRYSMPLIDL2CS & ", '"      'サンプルID（L2）
                sql = sql & .CRYINDL2CS & "', '"        '状態FLG（L2）
                sql = sql & .CRYRESL2CS & "', "         '実績FLG（L2）
                sql = sql & .CRYSMPLIDL3CS & ", '"      'サンプルID（L3）
                sql = sql & .CRYINDL3CS & "', '"        '状態FLG（L3）
                sql = sql & .CRYRESL3CS & "', "         '実績FLG（L3）
                sql = sql & .CRYSMPLIDL4CS & ", '"      'サンプルID（L4）
                sql = sql & .CRYINDL4CS & "', '"        '状態FLG（L4）
                sql = sql & .CRYRESL4CS & "', "         '実績FLG（L4）
                sql = sql & .CRYSMPLIDCSCS & ", '"      'サンプルID（CS）
                sql = sql & .CRYINDCSCS & "', '"        '状態FLG（CS）
                sql = sql & .CRYRESCSCS & "', "         '実績FLG（CS）
                sql = sql & .CRYSMPLIDGDCS & ", '"      'サンプルID（GD）
                sql = sql & .CRYINDGDCS & "', '"        '状態FLG（GD）
                sql = sql & .CRYRESGDCS & "', "         '実績FLG（GD）
                sql = sql & .CRYSMPLIDTCS & ", '"       'サンプルID（T）
                sql = sql & .CRYINDTCS & "', '"         '状態FLG（T）
                sql = sql & .CRYRESTCS & "', "          '実績FLG（T）
                sql = sql & .CRYSMPLIDEPCS & ", '"      'サンプルID（EPD）
                sql = sql & .CRYINDEPCS & "', '"        '状態FLG（EPD）
                sql = sql & .CRYRESEPCS & "', "         '実績FLG（EPD）
                
                '>>>>> X線測定追加対応 2009/07/28 SETsw kubota ---------------
                sql = sql & .CRYSMPLIDXCS               'サンプルID（X線）
                '状態FLG（X線）
                If .CRYINDXCS <> "" And left(.CRYINDXCS, 1) <> vbNullChar Then
                    sql = sql & ",'" & .CRYINDXCS & "'"
                Else
                    sql = sql & ",'0'"
                End If
                '実績FLG（X線）
                If .CRYRESXCS <> "" And left(.CRYRESXCS, 1) <> vbNullChar Then
                    sql = sql & ",'" & .CRYRESXCS & "'"
                Else
                    sql = sql & ",'0'"
                End If
                sql = sql & ", "
                '<<<<< X線測定追加対応 2009/07/28 SETsw kubota ---------------

                'Add Start 2010/12/13 SMPK Miyata
                sql = sql & .CRYSMPLIDCCS & ", '"       'サンプルID（C）
                sql = sql & .CRYINDCCS & "', '"         '状態FLG（C）
                sql = sql & .CRYRESCCS & "', "          '実績FLG（C）
                sql = sql & .CRYSMPLIDCJCS & ", '"      'サンプルID（CJ）
                sql = sql & .CRYINDCJCS & "', '"        '状態FLG（CJ）
                sql = sql & .CRYRESCJCS & "', "         '実績FLG（CJ）
                sql = sql & .CRYSMPLIDCJLTCS & ", '"    'サンプルID（CJLT）
                sql = sql & .CRYINDCJLTCS & "', '"      '状態FLG（CJLT）
                sql = sql & .CRYRESCJLTCS & "', "       '実績FLG（CJLT）
                sql = sql & .CRYSMPLIDCJ2CS & ", '"     'サンプルID（CJ2）
                sql = sql & .CRYINDCJ2CS & "', '"       '状態FLG（CJ2）
                sql = sql & .CRYRESCJ2CS & "', "        '実績FLG（CJ2）
                'Add End   2010/12/13 SMPK Miyata

                sql = sql & .SMPLNUMCS & ", "           'サンプル枚数
                sql = sql & "' ', '"                    'サンプルパターン
                sql = sql & .LIVKCS & "', '"            '生死区分
                sql = sql & .TSTAFFCS & "', "           '登録社員ID
                sql = sql & "sysdate, '"                '登録日付
                sql = sql & .KSTAFFCS & "', "           '更新社員ID
                sql = sql & "sysdate, "                 '更新日付
                sql = sql & "'0', "                     '送信フラグ
'                sql = sql & "sysdate)"                  '送信日付
                '05/10/17 ooba START =====================================================>
                sql = sql & "sysdate"                   '送信日付
                '親ブロックID
                If left(.RPCRYNUMCS, 1) <> vbNullChar And Trim(.RPCRYNUMCS) <> "" Then
                    sql = sql & ", '" & .RPCRYNUMCS & "'"
                End If
                '切断フラグ
                If left(.CUTFLGCS, 1) <> vbNullChar And Trim(.CUTFLGCS) <> "" Then
                    sql = sql & ", '" & .CUTFLGCS & "'"
                Else
                    sql = sql & ", NULL"
                End If
                '05/10/17 ooba END =======================================================>
                
                '>>>>> 抵抗狙い位置対応 2009/11/06 SETsw kubota ---------------
                If .QCKBNCS <> "" And left(.QCKBNCS, 1) <> vbNullChar Then
                    sql = sql & ",'" & .QCKBNCS & "'"
                Else
                    sql = sql & ",NULL"
                End If
                '<<<<< 抵抗狙い位置対応 2009/11/06 SETsw kubota ---------------
                
                sql = sql & ")"
                
                
                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_CrySmp_UpdIns = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_CrySmp_UpdIns = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :結晶サンプル管理の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型                  ,説明
'      　　:CrySmpNew　　　,I  ,typ_XSDCS   　      ,新サンプル管理（ブロック）
'      　　:戻り値         ,O  ,FUNCTION_RETURN　   ,書き込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_CrySmp_Ins(CrySmpNew() As typ_XSDCS) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CrySmp_Ins"

    DBDRV_CrySmp_Ins = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(CrySmpNew)
        With CrySmpNew(i)
'            sql = "insert into TBCME043 ("
'            sql = sql & "CRYNUM, "          ' 結晶番号
'            sql = sql & "INGOTPOS, "        ' 結晶内位置
'            sql = sql & "SMPKBN, "          ' サンプル区分
'            sql = sql & "SMPLNO, "          ' サンプルNo
'            sql = sql & "HINBAN, "          ' 品番
'            sql = sql & "REVNUM, "          ' 製品番号改訂番号
'            sql = sql & "FACTORY, "         ' 工場
'            sql = sql & "OPECOND, "         ' 操業条件
'            sql = sql & "KTKBN, "           ' 確定区分
'            sql = sql & "CRYINDRS, "        ' 結晶検査指示（Rs)
'            sql = sql & "CRYINDOI, "        ' 結晶検査指示（Oi)
'            sql = sql & "CRYINDB1, "        ' 結晶検査指示（B1)
'            sql = sql & "CRYINDB2, "        ' 結晶検査指示（B2)
'            sql = sql & "CRYINDB3, "        ' 結晶検査指示（B3)
'            sql = sql & "CRYINDL1, "        ' 結晶検査指示（L1)
'            sql = sql & "CRYINDL2, "        ' 結晶検査指示（L2)
'            sql = sql & "CRYINDL3, "        ' 結晶検査指示（L3)
'            sql = sql & "CRYINDL4, "        ' 結晶検査指示（L4)
'            sql = sql & "CRYINDCS, "        ' 結晶検査指示（Cs)
'            sql = sql & "CRYINDGD, "        ' 結晶検査指示（GD)
'            sql = sql & "CRYINDT, "         ' 結晶検査指示（T)
'            sql = sql & "CRYINDEP, "        ' 結晶検査指示（EPD)
'            sql = sql & "CRYRESRS, "        ' 結晶検査実績（Rs)
'            sql = sql & "CRYRESOI, "        ' 結晶検査実績（Oi)
'            sql = sql & "CRYRESB1, "        ' 結晶検査実績（B1)
'            sql = sql & "CRYRESB2, "        ' 結晶検査実績（B2)
'            sql = sql & "CRYRESB3, "        ' 結晶検査実績（B3)
'            sql = sql & "CRYRESL1, "        ' 結晶検査実績（L1)
'            sql = sql & "CRYRESL2, "        ' 結晶検査実績（L2)
'            sql = sql & "CRYRESL3, "        ' 結晶検査実績（L3)
'            sql = sql & "CRYRESL4, "        ' 結晶検査実績（L4)
'            sql = sql & "CRYRESCS, "        ' 結晶検査実績（Cs)
'            sql = sql & "CRYRESGD, "        ' 結晶検査実績（GD)
'            sql = sql & "CRYREST, "         ' 結晶検査実績（T)
'            sql = sql & "CRYRESEP, "        ' 結晶検査実績（EPD)
'            sql = sql & "SMPLNUM, "         ' サンプル枚数
'            sql = sql & "SMPLPAT, "         ' サンプルパターン
'            sql = sql & "REGDATE, "         ' 登録日付
'            sql = sql & "UPDDATE, "         ' 更新日付
'            sql = sql & "SENDFLAG, "        ' 送信フラグ
'            sql = sql & "SENDDATE)"         ' 送信日付
'            sql = sql & " values ('"
'            sql = sql & .CRYNUM & "', "
'            sql = sql & .INGOTPOS & ", '"   ' 結晶内位置
'            sql = sql & .SMPKBN & "', "     ' サンプル区分
'            sql = sql & .SMPLNO & ", '"     ' サンプルNo
'            sql = sql & .HINBAN & "', "     ' 品番
'            sql = sql & .REVNUM & ", '"     ' 製品番号改訂番号
'            sql = sql & .factory & "', '"   ' 工場
'            sql = sql & .opecond & "', '"   ' 操業条件
'            sql = sql & .KTKBN & "', '"     ' 確定区分
'            sql = sql & .CRYINDRS & "', '"  ' 結晶検査指示（Rs)
'            sql = sql & .CRYINDOI & "', '"  ' 結晶検査指示（Oi)
'            sql = sql & .CRYINDB1 & "', '"  ' 結晶検査指示（B1)
'            sql = sql & .CRYINDB2 & "', '"  ' 結晶検査指示（B2)
'            sql = sql & .CRYINDB3 & "', '"  ' 結晶検査指示（B3)
'            sql = sql & .CRYINDL1 & "', '"  ' 結晶検査指示（L1)
'            sql = sql & .CRYINDL2 & "', '"  ' 結晶検査指示（L2)
'            sql = sql & .CRYINDL3 & "', '"  ' 結晶検査指示（L3)
'            sql = sql & .CRYINDL4 & "', '"  ' 結晶検査指示（L4)
'            sql = sql & .CRYINDCS & "', '"  ' 結晶検査指示（Cs)
'            sql = sql & .CRYINDGD & "', '"  ' 結晶検査指示（GD)
'            sql = sql & .CRYINDT & "', '"   ' 結晶検査指示（T)
'            sql = sql & .CRYINDEP & "', '"  ' 結晶検査指示（EPD)
'            sql = sql & .CRYRESRS & "', '"  ' 結晶検査実績（Rs)
'            sql = sql & .CRYRESOI & "', '"  ' 結晶検査実績（Oi)
'            sql = sql & .CRYRESB1 & "', '"  ' 結晶検査実績（B1)
'            sql = sql & .CRYRESB2 & "', '"  ' 結晶検査実績（B2)
'            sql = sql & .CRYRESB3 & "', '"  ' 結晶検査実績（B3)
'            sql = sql & .CRYRESL1 & "', '"  ' 結晶検査実績（L1)
'            sql = sql & .CRYRESL2 & "', '"  ' 結晶検査実績（L2)
'            sql = sql & .CRYRESL3 & "', '"  ' 結晶検査実績（L3)
'            sql = sql & .CRYRESL4 & "', '"  ' 結晶検査実績（L4)
'            sql = sql & .CRYRESCS & "', '"  ' 結晶検査実績（Cs)
'            sql = sql & .CRYRESGD & "', '"  ' 結晶検査実績（GD)
'            sql = sql & .CRYREST & "', '"   ' 結晶検査実績（T)
'            sql = sql & .CRYRESEP & "', "   ' 結晶検査実績（EPD)
'            sql = sql & .SMPLNUM & ", "     ' サンプル枚数
'            sql = sql & "' ', "             ' サンプルパターン
'            sql = sql & "sysdate, "
'            sql = sql & "sysdate, "
'            sql = sql & "'0', "
'            sql = sql & "sysdate)"
            sql = "insert into XSDCS ("
            sql = sql & "CRYNUMCS,"         'ブロックID
            sql = sql & "SMPKBNCS,"         'サンプル区分
            sql = sql & "TBKBNCS,"          'T/B区分
            sql = sql & "REPSMPLIDCS,"      '代表サンプルID
            sql = sql & "XTALCS,"           '結晶番号
            sql = sql & "INPOSCS,"          '結晶内位置
            sql = sql & "HINBCS,"           '品番
            sql = sql & "REVNUMCS,"         '製品番号改訂番号
            sql = sql & "FACTORYCS,"        '工場
            sql = sql & "OPECS,"            '操業番号
            sql = sql & "KTKBNCS,"          '確定区分
            sql = sql & "BLKKTFLAGCS,"      'ブロック確定フラグ
            sql = sql & "CRYSMPLIDRSCS,"    'サンプルID(Rs)
            sql = sql & "CRYSMPLIDRS1CS,"   '推定サンプルID1（Rs）
            sql = sql & "CRYSMPLIDRS2CS,"   '推定サンプルID2（Rs）
            sql = sql & "CRYINDRSCS,"       '状態FLG(Rs)
            sql = sql & "CRYRESRS1CS,"      '実績FLG1(Rs)
            sql = sql & "CRYRESRS2CS,"      '実績FLG2(Rs)
            sql = sql & "CRYSMPLIDOICS,"    'サンプルID（Oi）
            sql = sql & "CRYINDOICS,"       '状態FLG（Oi）
            sql = sql & "CRYRESOICS,"       '実績FLG（Oi）
            sql = sql & "CRYSMPLIDB1CS,"    'サンプルID（B1）
            sql = sql & "CRYINDB1CS,"       '状態FLG（B1）
            sql = sql & "CRYRESB1CS,"       '実績FLG（B1）
            sql = sql & "CRYSMPLIDB2CS,"    'サンプルID（B2）
            sql = sql & "CRYINDB2CS,"       '状態FLG（B2）
            sql = sql & "CRYRESB2CS,"       '実績FLG（B2）
            sql = sql & "CRYSMPLIDB3CS,"    'サンプルID（B3）
            sql = sql & "CRYINDB3CS,"       '状態FLG（B3）
            sql = sql & "CRYRESB3CS,"       '実績FLG（B3）
            sql = sql & "CRYSMPLIDL1CS,"    'サンプルID（L1）
            sql = sql & "CRYINDL1CS,"       '状態FLG（L1）
            sql = sql & "CRYRESL1CS,"       '実績FLG（L1）
            sql = sql & "CRYSMPLIDL2CS,"    'サンプルID（L2）
            sql = sql & "CRYINDL2CS,"       '状態FLG（L2）
            sql = sql & "CRYRESL2CS,"       '実績FLG（L2）
            sql = sql & "CRYSMPLIDL3CS,"    'サンプルID（L3）
            sql = sql & "CRYINDL3CS,"       '状態FLG（L3）
            sql = sql & "CRYRESL3CS,"       '実績FLG（L3）
            sql = sql & "CRYSMPLIDL4CS,"    'サンプルID（L4）
            sql = sql & "CRYINDL4CS,"       '状態FLG（L4）
            sql = sql & "CRYRESL4CS,"       '実績FLG（L4）
            sql = sql & "CRYSMPLIDCSCS,"    'サンプルID（CS）
            sql = sql & "CRYINDCSCS,"       '状態FLG（CS）
            sql = sql & "CRYRESCSCS,"       '実績FLG（CS）
            sql = sql & "CRYSMPLIDGDCS,"    'サンプルID（GD）
            sql = sql & "CRYINDGDCS,"       '状態FLG（GD）
            sql = sql & "CRYRESGDCS,"       '実績FLG（GD）
            sql = sql & "CRYSMPLIDTCS,"     'サンプルID（T）
            sql = sql & "CRYINDTCS,"        '状態FLG（T）
            sql = sql & "CRYRESTCS,"        '実績FLG（T）
            sql = sql & "CRYSMPLIDEPCS,"    'サンプルID（EPD）
            sql = sql & "CRYINDEPCS,"       '状態FLG（EPD）
            sql = sql & "CRYRESEPCS,"       '実績FLG（EPD）
            sql = sql & "CRYSMPLIDXCS,"     'サンプルID（X線）  'X線測定 2009/07/27追加 SETsw kubota
            sql = sql & "CRYINDXCS,"        '状態FLG（X線）
            sql = sql & "CRYRESXCS,"        '実績FLG（X線）
            sql = sql & "SMPLNUMCS,"        'サンプル枚数
            sql = sql & "SMPLPATCS,"        'サンプルパターン
            sql = sql & "TSTAFFCS,"         '登録社員ID
            sql = sql & "TDAYCS,"           '登録日付
            sql = sql & "KSTAFFCS,"         '更新社員ID
            sql = sql & "KDAYCS,"           '更新日付
            sql = sql & "SNDKCS,"           '送信フラグ
            sql = sql & "SNDDAYCS)"         '送信日付
            sql = sql & " values ('"
            sql = sql & .CRYNUMCS & "', '"          'ブロックID
            sql = sql & .SMPKBNCS & "', '"          'サンプル区分
            sql = sql & .TBKBNCS & "', "            'T/B区分
            sql = sql & .REPSMPLIDCS & ", '"        '代表サンプルID
            sql = sql & .XTALCS & "', "             '結晶番号
            sql = sql & .INPOSCS & ", '"            '結晶内位置
            sql = sql & .HINBCS & "', "             '品番
            sql = sql & .REVNUMCS & ", '"           '製品番号改訂番号
            sql = sql & .FACTORYCS & "', '"         '工場
            sql = sql & .OPECS & "', '"             '操業条件
            sql = sql & .KTKBNCS & "', '"           '確定区分
            sql = sql & .BLKKTFLAGCS & "', "        'ブロック確定フラグ
            sql = sql & .CRYSMPLIDRSCS & ", "       'サンプルID（Rs）
            sql = sql & .CRYSMPLIDRS1CS & ", "      '推定サンプルID1（Rs）
            sql = sql & .CRYSMPLIDRS2CS & ", '"     '推定サンプルID2（Rs）
            sql = sql & .CRYINDRSCS & "', '"        '状態FLG（Rs）
            sql = sql & .CRYRESRS1CS & "', '"       '実績FLG1（Rs）
            sql = sql & .CRYRESRS2CS & "', "        '実績FLG2（Rs）
            sql = sql & .CRYSMPLIDOICS & ", '"      'サンプルID（Oi）
            sql = sql & .CRYINDOICS & "', '"        '状態FLG（Oi）
            sql = sql & .CRYRESOICS & "', "         '実績FLG（Oi）
            sql = sql & .CRYSMPLIDB1CS & ", '"      'サンプルID（B1）
            sql = sql & .CRYINDB1CS & "', '"        '状態FLG（B1）
            sql = sql & .CRYRESB1CS & "', "         '実績FLG（B1）
            sql = sql & .CRYSMPLIDB2CS & ", '"      'サンプルID（B2）
            sql = sql & .CRYINDB2CS & "', '"        '状態FLG（B2）
            sql = sql & .CRYRESB2CS & "', "         '実績FLG（B2）
            sql = sql & .CRYSMPLIDB3CS & ", '"      'サンプルID（B3）
            sql = sql & .CRYINDB3CS & "', '"        '状態FLG（B3）
            sql = sql & .CRYRESB3CS & "', "         '実績FLG（B3）
            sql = sql & .CRYSMPLIDL1CS & ", '"      'サンプルID（L1）
            sql = sql & .CRYINDL1CS & "', '"        '状態FLG（L1）
            sql = sql & .CRYRESL1CS & "', "         '実績FLG（L1）
            sql = sql & .CRYSMPLIDL2CS & ", '"      'サンプルID（L2）
            sql = sql & .CRYINDL2CS & "', '"        '状態FLG（L2）
            sql = sql & .CRYRESL2CS & "', "         '実績FLG（L2）
            sql = sql & .CRYSMPLIDL3CS & ", '"      'サンプルID（L3）
            sql = sql & .CRYINDL3CS & "', '"        '状態FLG（L3）
            sql = sql & .CRYRESL3CS & "', "         '実績FLG（L3）
            sql = sql & .CRYSMPLIDL4CS & ", '"      'サンプルID（L4）
            sql = sql & .CRYINDL4CS & "', '"        '状態FLG（L4）
            sql = sql & .CRYRESL4CS & "', "         '実績FLG（L4）
            sql = sql & .CRYSMPLIDCSCS & ", '"      'サンプルID（CS）
            sql = sql & .CRYINDCSCS & "', '"        '状態FLG（CS）
            sql = sql & .CRYRESCSCS & "', "         '実績FLG（CS）
            sql = sql & .CRYSMPLIDGDCS & ", '"      'サンプルID（GD）
            sql = sql & .CRYINDGDCS & "', '"        '状態FLG（GD）
            sql = sql & .CRYRESGDCS & "', "         '実績FLG（GD）
            sql = sql & .CRYSMPLIDTCS & ", '"       'サンプルID（T）
            sql = sql & .CRYINDTCS & "', '"         '状態FLG（T）
            sql = sql & .CRYRESTCS & "', "          '実績FLG（T）
            sql = sql & .CRYSMPLIDEPCS & ", '"      'サンプルID（EPD）
            sql = sql & .CRYINDEPCS & "', '"        '状態FLG（EPD）
            sql = sql & .CRYRESEPCS & "', "         '実績FLG（EPD）
            
            '>>>>> X線測定追加対応 2009/07/28 SETsw kubota ---------------
            sql = sql & .CRYSMPLIDXCS               'サンプルID（X線）
            '状態FLG（X線）
            If .CRYINDXCS <> "" And left(.CRYINDXCS, 1) <> vbNullChar Then
                sql = sql & ",'" & .CRYINDXCS & "'"
            Else
                sql = sql & ",'0'"
            End If
            '実績FLG（X線）
            If .CRYRESXCS <> "" And left(.CRYRESXCS, 1) <> vbNullChar Then
                sql = sql & ",'" & .CRYRESXCS & "'"
            Else
                sql = sql & ",'0'"
            End If
            sql = sql & ", "
            '<<<<< X線測定追加対応 2009/07/28 SETsw kubota ---------------
            
            sql = sql & .SMPLNUMCS & ", "           'サンプル枚数
            sql = sql & "' ', '"                    'サンプルパターン
            sql = sql & .TSTAFFCS & "', "           '登録社員ID
            sql = sql & "sysdate, '"                '登録日付
            sql = sql & .KSTAFFCS & "', "           '更新社員ID
            sql = sql & "sysdate, "                 '更新日付
            sql = sql & "'0', "                     '送信フラグ
            sql = sql & "sysdate)"                  '送信日付
        End With
        '' WriteDBLog sql
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_CrySmp_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_CrySmp_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :WFサンプル管理の挿入／更新
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:WfSmpOld　　　,I  ,typ_XSDCW   　   ,新サンプル管理（SXL）（旧）
'      　　:WfSmpNew　　　,I  ,typ_XSDCW   　   ,新サンプル管理（SXL）（新）
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :古いレコードをみて更新か挿入かを判別する
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_WfSmp_UpdIns(WfSmpOld() As typ_XSDCW, WfSmpNew() As typ_XSDCW) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_WfSmp_UpdIns"

    DBDRV_WfSmp_UpdIns = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(WfSmpNew)
        With WfSmpNew(i)
            lFlg = False
            For j = 1 To UBound(WfSmpOld)
                If WfSmpOld(j).XTALCW = .XTALCW And _
                   WfSmpOld(j).INPOSCW = .INPOSCW And _
                   WfSmpOld(j).SMPKBNCW = .SMPKBNCW Then
'                    sql = "update TBCME044 set "
'                    sql = sql & "SMPLID='" & .SMPLID & "', "        ' サンプルID
'                    sql = sql & "HINBAN='" & .HINBAN & "', "        ' 品番
'                    sql = sql & "REVNUM=" & .REVNUM & ", "          ' 製品番号改訂番号
'                    sql = sql & "FACTORY='" & .factory & "', "      ' 工場
'                    sql = sql & "OPECOND='" & .opecond & "', "      ' 操業条件
'                    sql = sql & "KTKBN='" & .KTKBN & "', "          ' 確定区分
'                    sql = sql & "WFINDRS='" & .WFINDRS & "', "      ' WF検査指示（Rs)
'                    sql = sql & "WFINDOI='" & .WFINDOI & "', "      ' WF検査指示（Oi)
'                    sql = sql & "WFINDB1='" & .WFINDB1 & "', "      ' WF検査指示（B1)
'                    sql = sql & "WFINDB2='" & .WFINDB2 & "', "      ' WF検査指示（B2)
'                    sql = sql & "WFINDB3='" & .WFINDB3 & "', "      ' WF検査指示（B3)
'                    sql = sql & "WFINDL1='" & .WFINDL1 & "', "      ' WF検査指示（L1)
'                    sql = sql & "WFINDL2='" & .WFINDL2 & "', "      ' WF検査指示（L2)
'                    sql = sql & "WFINDL3='" & .WFINDL3 & "', "      ' WF検査指示（L3)
'                    sql = sql & "WFINDL4='" & .WFINDL4 & "', "      ' WF検査指示（L4)
'                    sql = sql & "WFINDDS='" & .WFINDDS & "', "      ' WF検査指示（DS)
'                    sql = sql & "WFINDDZ='" & .WFINDDZ & "', "      ' WF検査指示（DZ)
'                    sql = sql & "WFINDSP='" & .WFINDSP & "', "      ' WF検査指示（SP)
'                    sql = sql & "WFINDDO1='" & .WFINDDO1 & "', "    ' WF検査指示（DO1)
'                    sql = sql & "WFINDDO2='" & .WFINDDO2 & "', "    ' WF検査指示（DO2)
'                    sql = sql & "WFINDDO3='" & .WFINDDO3 & "', "    ' WF検査指示（DO3)
'                    'add start 2003/05/21 hitec)matsumoto -------------------------
'                    sql = sql & "WFINDOT1='" & .WFINDOT1 & "', "    ' WF検査指示（OT1)
'                    sql = sql & "WFINDOT2='" & .WFINDOT2 & "', "    ' WF検査指示（OT2)
'                    'add end   2003/05/21 hitec)matsumoto -------------------------
'                    sql = sql & "WFRESRS='" & .WFRESRS & "', "      ' WF検査実績（Rs)
'                    sql = sql & "WFRESOI='" & .WFRESOI & "', "      ' WF検査実績（Oi)
'                    sql = sql & "WFRESB1='" & .WFRESB1 & "', "      ' WF検査実績（B1)
'                    sql = sql & "WFRESB2='" & .WFRESB2 & "', "      ' WF検査実績（B2)
'                    sql = sql & "WFRESB3='" & .WFRESB3 & "', "      ' WF検査実績（B3)
'                    sql = sql & "WFRESL1='" & .WFRESL1 & "', "      ' WF検査実績（L1)
'                    sql = sql & "WFRESL2='" & .WFRESL2 & "', "      ' WF検査実績（L2)
'                    sql = sql & "WFRESL3='" & .WFRESL3 & "', "      ' WF検査実績（L3)
'                    sql = sql & "WFRESL4='" & .WFRESL4 & "', "      ' WF検査実績（L4)
'                    sql = sql & "WFRESDS='" & .WFRESDS & "', "      ' WF検査実績（DS)
'                    sql = sql & "WFRESDZ='" & .WFRESDZ & "', "      ' WF検査実績（DZ)
'                    sql = sql & "WFRESSP='" & .WFRESSP & "', "      ' WF検査実績（SP)
'                    sql = sql & "WFRESDO1='" & .WFRESDO1 & "', "    ' WF検査実績（DO1)
'                    sql = sql & "WFRESDO2='" & .WFRESDO2 & "', "    ' WF検査実績（DO2)
'                    sql = sql & "WFRESDO3='" & .WFRESDO3 & "', "    ' WF検査実績（DO3)
'                    'add start 2003/05/21 hitec)matsumoto -------------------------
'                    sql = sql & "WFRESOT1='" & .WFRESOT1 & "', "    ' WF検査指示（OT1)
'                    sql = sql & "WFRESOT2='" & .WFRESOT2 & "', "    ' WF検査指示（OT2)
'                    'add end   2003/05/21 hitec)matsumoto -------------------------
'                    sql = sql & "UPDDATE=sysdate, "
'                    sql = sql & "SENDFLAG='0'"
'                    sql = sql & " where CRYNUM='" & .CRYNUM & "'"
'                    sql = sql & " and INGOTPOS=" & .INGOTPOS
'                    sql = sql & " and SMPKBN='" & .SMPKBN & "'"

                    sql = "update XSDCW set "
                    sql = sql & "SXLIDCW='" & .SXLIDCW & "', "          ' SXLID
                    sql = sql & "REPSMPLIDCW='" & .REPSMPLIDCW & "', "  ' サンプルID
                    sql = sql & "HINBCW='" & .HINBCW & "', "            ' 品番
                    sql = sql & "REVNUMCW=" & .REVNUMCW & ", "          ' 製品番号改訂番号
                    sql = sql & "FACTORYCW='" & .FACTORYCW & "', "      ' 工場
                    sql = sql & "OPECW='" & .OPECW & "', "              ' 操業条件
                    sql = sql & "KTKBNCW='" & .KTKBNCW & "', "          ' 確定区分
                    sql = sql & "WFINDRSCW='" & .WFINDRSCW & "', "      ' 状態FLG（Rs)
                    sql = sql & "WFRESRS1CW='" & .WFRESRS1CW & "', "    ' 実績FLG1（Rs)
                    sql = sql & "WFINDOICW='" & .WFINDOICW & "', "      ' 状態FLG（Oi)
                    sql = sql & "WFRESOICW='" & .WFRESOICW & "', "      ' 実績FLG（Oi)
                    sql = sql & "WFINDB1CW='" & .WFINDB1CW & "', "      ' 状態FLG（B1)
                    sql = sql & "WFRESB1CW='" & .WFRESB1CW & "', "      ' 実績FLG（B1)
                    sql = sql & "WFINDB2CW='" & .WFINDB2CW & "', "      ' 状態FLG（B2)
                    sql = sql & "WFRESB2CW='" & .WFRESB2CW & "', "      ' 実績FLG（B2)
                    sql = sql & "WFINDB3CW='" & .WFINDB3CW & "', "      ' 状態FLG（B3)
                    sql = sql & "WFRESB3CW='" & .WFRESB3CW & "', "      ' 実績FLG（B3)
                    sql = sql & "WFINDL1CW='" & .WFINDL1CW & "', "      ' 状態FLG（L1)
                    sql = sql & "WFRESL1CW='" & .WFRESL1CW & "', "      ' 実績FLG（L1)
                    sql = sql & "WFINDL2CW='" & .WFINDL2CW & "', "      ' 状態FLG（L2)
                    sql = sql & "WFRESL2CW='" & .WFRESL2CW & "', "      ' 実績FLG（L2)
                    sql = sql & "WFINDL3CW='" & .WFINDL3CW & "', "      ' 状態FLG（L3)
                    sql = sql & "WFRESL3CW='" & .WFRESL3CW & "', "      ' 実績FLG（L3)
                    sql = sql & "WFINDL4CW='" & .WFINDL4CW & "', "      ' 状態FLG（L4)
                    sql = sql & "WFRESL4CW='" & .WFRESL4CW & "', "      ' 実績FLG（L4)
                    sql = sql & "WFINDDSCW='" & .WFINDDSCW & "', "      ' 状態FLG（DS)
                    sql = sql & "WFRESDSCW='" & .WFRESDSCW & "', "      ' 実績FLG（DS)
                    sql = sql & "WFINDDZCW='" & .WFINDDZCW & "', "      ' 状態FLG（DZ)
                    sql = sql & "WFRESDZCW='" & .WFRESDZCW & "', "      ' 実績FLG（DZ)
                    sql = sql & "WFINDSPCW='" & .WFINDSPCW & "', "      ' 状態FLG（SP)
                    sql = sql & "WFRESSPCW='" & .WFRESSPCW & "', "      ' 実績FLG（SP)
                    sql = sql & "WFINDDO1CW='" & .WFINDDO1CW & "', "    ' 状態FLG（DO1)
                    sql = sql & "WFRESDO1CW='" & .WFRESDO1CW & "', "    ' 実績FLG（DO1)
                    sql = sql & "WFINDDO2CW='" & .WFINDDO2CW & "', "    ' 状態FLG（DO2)
                    sql = sql & "WFRESDO2CW='" & .WFRESDO2CW & "', "    ' 実績FLG（DO2)
                    sql = sql & "WFINDDO3CW='" & .WFINDDO3CW & "', "    ' 状態FLG（DO3)
                    sql = sql & "WFRESDO3CW='" & .WFRESDO3CW & "', "    ' 実績FLG（DO3)
                    'add start 2003/05/21 hitec)matsumoto -------------------------
                    sql = sql & "WFINDOT1CW='" & .WFINDOT1CW & "', "    ' 状態FLG（OT1)
                    sql = sql & "WFRESOT1CW='" & .WFRESOT1CW & "', "    ' 実績FLG（OT1)
                    sql = sql & "WFINDOT2CW='" & .WFINDOT2CW & "', "    ' 状態FLG（OT2)
                    sql = sql & "WFRESOT2CW='" & .WFRESOT2CW & "', "    ' 実績FLG（OT2)
                    'add end   2003/05/21 hitec)matsumoto -------------------------
                    '' 残存酸素追加　03/12/05 ooba START ===============================>
                    sql = sql & "WFINDAOICW='" & .WFINDAOICW & "', "    ' 状態FLG (AOI)
                    sql = sql & "WFRESAOICW='" & .WFRESAOICW & "', "    ' 実績FLG (AOI)
                    '' 残存酸素追加　03/12/05 ooba END =================================>
                    '' GD追加　05/01/17 ooba START =====================================>
                    sql = sql & "WFINDGDCW='" & .WFINDGDCW & "', "    ' 状態FLG (GD)
                    sql = sql & "WFRESGDCW='" & .WFRESGDCW & "', "    ' 実績FLG (GD)
                    sql = sql & "WFHSGDCW='" & .WFHSGDCW & "', "      ' 保証FLG (GD)
                    '' GD追加　05/01/17 ooba END =======================================>
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                    sql = sql & "EPINDL1CW = " & .EPINDL1CW & "', "     ' 状態FLG (OSF1E)
                    sql = sql & "EPRESL1CW = " & .EPRESL1CW & "', "     ' 実績FLG (OSF1E)
                    sql = sql & "EPINDL2CW = " & .EPINDL2CW & "', "     ' 状態FLG (OSF2E)
                    sql = sql & "EPRESL2CW = " & .EPRESL2CW & "', "     ' 実績FLG (OSF2E)
                    sql = sql & "EPINDL3CW = " & .EPINDL3CW & "', "     ' 状態FLG (OSF3E)
                    sql = sql & "EPRESL3CW = " & .EPRESL3CW & "', "     ' 実績FLG (OSF3E)
                    sql = sql & "EPINDB1CW = " & .EPINDB1CW & "', "     ' 状態FLG (BMD1E)
                    sql = sql & "EPRESB1CW = " & .EPRESB1CW & "', "     ' 実績FLG (BMD1E)
                    sql = sql & "EPINDB2CW = " & .EPINDB2CW & "', "     ' 状態FLG (BMD2E)
                    sql = sql & "EPRESB2CW = " & .EPRESB2CW & "', "     ' 実績FLG (BMD2E)
                    sql = sql & "EPINDB3CW = " & .EPINDB3CW & "', "     ' 状態FLG (BMD3E)
                    sql = sql & "EPRESB3CW = " & .EPRESB3CW & "', "     ' 実績FLG (BMD3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                    sql = sql & "KDAYCW=sysdate, "
                    sql = sql & "SNDKCW='0'"
                    sql = sql & " where XTALCW='" & .XTALCW & "'"
                    sql = sql & " and INPOSCW=" & .INPOSCW
                    sql = sql & " and SMPKBNCW='" & .SMPKBNCW & "'"

                    '' WriteDBLog sql
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        DBDRV_WfSmp_UpdIns = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                    lFlg = True
                    Exit For
                End If
            Next j

            If lFlg <> True Then
'                sql = "insert into TBCME044 ("
'                sql = sql & "CRYNUM, "          ' 結晶番号
'                sql = sql & "INGOTPOS, "        ' 結晶内位置
'                sql = sql & "SMPKBN, "          ' サンプル区分
'                sql = sql & "SMPLID, "          ' サンプルID
'                sql = sql & "HINBAN, "          ' 品番
'                sql = sql & "REVNUM, "          ' 製品番号改訂番号
'                sql = sql & "FACTORY, "         ' 工場
'                sql = sql & "OPECOND, "         ' 操業条件
'                sql = sql & "KTKBN, "           ' 確定区分
'                sql = sql & "WFINDRS, "         ' WF検査指示（Rs)
'                sql = sql & "WFINDOI, "         ' WF検査指示（Oi)
'                sql = sql & "WFINDB1, "         ' WF検査指示（B1)
'                sql = sql & "WFINDB2, "         ' WF検査指示（B2)
'                sql = sql & "WFINDB3, "         ' WF検査指示（B3)
'                sql = sql & "WFINDL1, "         ' WF検査指示（L1)
'                sql = sql & "WFINDL2, "         ' WF検査指示（L2)
'                sql = sql & "WFINDL3, "         ' WF検査指示（L3)
'                sql = sql & "WFINDL4, "         ' WF検査指示（L4)
'                sql = sql & "WFINDDS, "         ' WF検査指示（DS)
'                sql = sql & "WFINDDZ, "         ' WF検査指示（DZ)
'                sql = sql & "WFINDSP, "         ' WF検査指示（SP)
'                sql = sql & "WFINDDO1, "        ' WF検査指示（DO1)
'                sql = sql & "WFINDDO2, "        ' WF検査指示（DO2)
'                sql = sql & "WFINDDO3, "        ' WF検査指示（DO3)
'               'add start 2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "WFINDOT1, "        ' WF検査指示（OT1)
'                sql = sql & "WFINDOT2, "        ' WF検査指示（OT2)
'               'add end   2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "WFRESRS, "         ' WF検査実績（Rs)
'                sql = sql & "WFRESOI, "         ' WF検査実績（Oi)
'                sql = sql & "WFRESB1, "         ' WF検査実績（B1)
'                sql = sql & "WFRESB2, "         ' WF検査実績（B2)
'                sql = sql & "WFRESB3, "         ' WF検査実績（B3)
'                sql = sql & "WFRESL1, "         ' WF検査実績（L1)
'                sql = sql & "WFRESL2, "         ' WF検査実績（L2)
'                sql = sql & "WFRESL3, "         ' WF検査実績（L3)
'                sql = sql & "WFRESL4, "         ' WF検査実績（L4)
'                sql = sql & "WFRESDS, "         ' WF検査実績（DS)
'                sql = sql & "WFRESDZ, "         ' WF検査実績（DZ)
'                sql = sql & "WFRESSP, "         ' WF検査実績（SP)
'                sql = sql & "WFRESDO1, "        ' WF検査実績（DO1)
'                sql = sql & "WFRESDO2, "        ' WF検査実績（DO2)
'                sql = sql & "WFRESDO3, "        ' WF検査実績（DO3)
'               'add start 2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "WFRESOT1, "        ' WF検査実績（OT1)
'                sql = sql & "WFRESOT2, "        ' WF検査実績（OT2)
'               'add end   2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "REGDATE, "         ' 更新日付
'                sql = sql & "UPDDATE, "         ' 更新日付
'                sql = sql & "SENDFLAG, "        ' 送信フラグ
'                sql = sql & "SENDDATE)"         ' 送信日付
'                sql = sql & " values ('"
'                sql = sql & .XTALCW & "', "     ' 結晶番号
'                sql = sql & .INPOSCW & ", '"   ' 結晶内位置
'                sql = sql & .SMPKBNCW & "', '"    ' サンプル区分
'                sql = sql & .REPSMPLIDCW & "', '"    ' サンプルID
'                sql = sql & .HINBCW & "', "     ' 品番
'                sql = sql & .REVNUMCW & ", '"     ' 製品番号改訂番号
'                sql = sql & .FACTORYCW & "', '"   ' 工場
'                sql = sql & .OPECW & "', '"   ' 操業条件
'                sql = sql & .KTKBNCW & "', '"     ' 確定区分
'                sql = sql & .WFINDRSCW & "', '"   ' WF検査指示（Rs)
'                sql = sql & .WFINDOICW & "', '"   ' WF検査指示（Oi)
'                sql = sql & .WFINDB1CW & "', '"   ' WF検査指示（B1)
'                sql = sql & .WFINDB2CW & "', '"   ' WF検査指示（B2)
'                sql = sql & .WFINDB3CW & "', '"   ' WF検査指示（B3)
'                sql = sql & .WFINDL1CW & "', '"   ' WF検査指示（L1)
'                sql = sql & .WFINDL2CW & "', '"   ' WF検査指示（L2)
'                sql = sql & .WFINDL3CW & "', '"   ' WF検査指示（L3)
'                sql = sql & .WFINDL4CW & "', '"   ' WF検査指示（L4)
'                sql = sql & .WFINDDSCW & "', '"   ' WF検査指示（DS)
'                sql = sql & .WFINDDZCW & "', '"   ' WF検査指示（DZ)
'                sql = sql & .WFINDSPCW & "', '"   ' WF検査指示（SP)
'                sql = sql & .WFINDDO1CW & "', '"  ' WF検査指示（DO1)
'                sql = sql & .WFINDDO2CW & "', '"  ' WF検査指示（DO2)
'                sql = sql & .WFINDDO3CW & "', '"  ' WF検査指示（DO3)
'                'add start 2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & .WFINDOT1CW & "', '"  ' WF検査指示（OT1)
'                sql = sql & .WFINDOT2CW & "', '"  ' WF検査指示（OT2)
'                'add end   2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & .WFRESRS1CW & "', '"   ' WF検査実績（Rs)
'                sql = sql & .WFRESOICW & "', '"   ' WF検査実績（Oi)
'                sql = sql & .WFRESB1CW & "', '"   ' WF検査実績（B1)
'                sql = sql & .WFRESB2CW & "', '"   ' WF検査実績（B2)
'                sql = sql & .WFRESB3CW & "', '"   ' WF検査実績（B3)
'                sql = sql & .WFRESL1CW & "', '"   ' WF検査実績（L1)
'                sql = sql & .WFRESL2CW & "', '"   ' WF検査実績（L2)
'                sql = sql & .WFRESL3CW & "', '"   ' WF検査実績（L3)
'                sql = sql & .WFRESL4CW & "', '"   ' WF検査実績（L4)
'                sql = sql & .WFRESDSCW & "', '"   ' WF検査実績（DS)
'                sql = sql & .WFRESDZCW & "', '"   ' WF検査実績（DZ)
'                sql = sql & .WFRESSPCW & "', '"   ' WF検査実績（SP)
'                sql = sql & .WFRESDO1CW & "', '"  ' WF検査実績（DO1)
'                sql = sql & .WFRESDO2CW & "', '"  ' WF検査実績（DO2)
'                sql = sql & .WFRESDO3CW & "', '"   ' WF検査実績（DO3)
'                'add start 2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & .WFRESOT1CW & "', '"  ' WF検査実績（OT1)
'                sql = sql & .WFRESOT2CW & "',"  ' WF検査実績（OT2)
'                'add end   2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "sysdate, "         ' 登録日付
'                sql = sql & "sysdate, "         ' 更新日付
'                sql = sql & "'0', "             ' 送信フラグ
'                sql = sql & "sysdate)"          ' 送信日付


                sql = "insert into XSDCW ("
                sql = sql & "SXLIDCW, "         ' SXLID
                sql = sql & "SMPKBNCW, "        ' サンプル区分
                sql = sql & "TBKBNCW, "         ' T/B区分
                sql = sql & "REPSMPLIDCW, "     ' サンプルID
                sql = sql & "XTALCW, "          ' 結晶番号
                sql = sql & "INPOSCW, "         ' 結晶内位置
                sql = sql & "HINBCW, "          ' 品番
                sql = sql & "REVNUMCW, "        ' 製品番号改訂番号
                sql = sql & "FACTORYCW, "       ' 工場
                sql = sql & "OPECW, "           ' 操業条件
                sql = sql & "KTKBNCW, "         ' 確定区分
                sql = sql & "SMCRYNUMCW, "      ' サンプルブロックID
                sql = sql & "WFSMPLIDRSCW, "    ' サンプルID(Rs)
                sql = sql & "WFSMPLIDRS1CW, "   ' 推定サンプルID1（Rs）
                sql = sql & "WFSMPLIDRS2CW, "   ' 推定サンプルID2（Rs）
                sql = sql & "WFINDRSCW, "       ' 状態FLG（Rs)
                sql = sql & "WFRESRS1CW, "      ' 実績FLG1（Rs)
                sql = sql & "WFRESRS2CW, "      ' 実績FLG2（Rs)
                sql = sql & "WFSMPLIDOICW, "    ' サンプルID（Oi）
                sql = sql & "WFINDOICW, "       ' 状態FLG（Oi)
                sql = sql & "WFRESOICW, "       ' 実績FLG（Oi)
                sql = sql & "WFSMPLIDB1CW, "    ' サンプルID（B1）
                sql = sql & "WFINDB1CW, "       ' 状態FLG（B1)
                sql = sql & "WFRESB1CW, "       ' 実績FLG（B1)
                sql = sql & "WFSMPLIDB2CW, "    ' サンプルID（B2）
                sql = sql & "WFINDB2CW, "       ' 状態FLG（B2)
                sql = sql & "WFRESB2CW, "       ' 実績FLG（B2)
                sql = sql & "WFSMPLIDB3CW, "    ' サンプルID（B3）
                sql = sql & "WFINDB3CW, "       ' 状態FLG（B3)
                sql = sql & "WFRESB3CW, "       ' 実績FLG（B3)
                sql = sql & "WFSMPLIDL1CW, "    ' サンプルID（L1）
                sql = sql & "WFINDL1CW, "       ' 状態FLG（L1)
                sql = sql & "WFRESL1CW, "       ' 実績FLG（L1)
                sql = sql & "WFSMPLIDL2CW, "    ' サンプルID（L2）
                sql = sql & "WFINDL2CW, "       ' 状態FLG（L2)
                sql = sql & "WFRESL2CW, "       ' 実績FLG（L2)
                sql = sql & "WFSMPLIDL3CW, "    ' サンプルID（L3）
                sql = sql & "WFINDL3CW, "       ' 状態FLG（L3)
                sql = sql & "WFRESL3CW, "       ' 実績FLG（L3)
                sql = sql & "WFSMPLIDL4CW, "    ' サンプルID（L4）
                sql = sql & "WFINDL4CW, "       ' 状態FLG（L4)
                sql = sql & "WFRESL4CW, "       ' 実績FLG（L4)
                sql = sql & "WFSMPLIDDSCW, "    ' サンプルID（DS）
                sql = sql & "WFINDDSCW, "       ' 状態FLG（DS)
                sql = sql & "WFRESDSCW, "       ' 実績FLG（DS)
                sql = sql & "WFSMPLIDDZCW, "    ' サンプルID（DZ）
                sql = sql & "WFINDDZCW, "       ' 状態FLG（DZ)
                sql = sql & "WFRESDZCW, "       ' 実績FLG（DZ)
                sql = sql & "WFSMPLIDSPCW, "    ' サンプルID（SP）
                sql = sql & "WFINDSPCW, "       ' 状態FLG（SP)
                sql = sql & "WFRESSPCW, "       ' 実績FLG（SP)
                sql = sql & "WFSMPLIDDO1CW,"    ' サンプルID（DO1）
                sql = sql & "WFINDDO1CW, "      ' 状態FLG（DO1)
                sql = sql & "WFRESDO1CW, "      ' 実績FLG（DO1)
                sql = sql & "WFSMPLIDDO2CW, "   ' サンプルID（DO2）
                sql = sql & "WFINDDO2CW, "      ' 状態FLG（DO2)
                sql = sql & "WFRESDO2CW, "      ' 実績FLG（DO2)
                sql = sql & "WFSMPLIDDO3CW, "   ' サンプルID（DO3）
                sql = sql & "WFINDDO3CW, "      ' 状態FLG（DO3)
                sql = sql & "WFRESDO3CW, "      ' 実績FLG（DO3)
                sql = sql & "WFSMPLIDOT1CW, "   ' サンプルID（OT1）
                sql = sql & "WFSMPLIDOT2CW, "   ' サンプルID（OT2）
               'add start 2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFINDOT1CW, "      ' 状態FLG（OT1)
                sql = sql & "WFRESOT1CW, "      ' 実績FLG（OT1)
                sql = sql & "WFINDOT2CW, "      ' 状態FLG（OT2)
                sql = sql & "WFRESOT2CW, "      ' 実績FLG（OT2)
               'add end   2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFSMPLIDAOICW, "   ' サンプルID（AOi）
                sql = sql & "WFINDAOICW, "      ' 状態FLG（AOi）
                sql = sql & "WFRESAOICW, "      ' 実績FLG（AOi）
                sql = sql & "SMPLNUMCW, "       ' サンプル枚数
                sql = sql & "SMPLPATCW, "       ' サンプルパターン
                sql = sql & "LIVKCW,"           ' 生死区分
                sql = sql & "TSTAFFCW,"         ' 登録社員ID
                sql = sql & "TDAYCW, "          ' 登録日付
                sql = sql & "KSTAFFCW, "        ' 更新社員ID
                sql = sql & "KDAYCW, "          ' 更新日付
                sql = sql & "SNDKCW, "          ' 送信フラグ
'                sql = sql & "SNDDAYCW)"         ' 送信日付
                '' GD追加　05/01/17 ooba START =====================================>
                sql = sql & "SNDDAYCW, "        ' 送信日付
                sql = sql & "WFSMPLIDGDCW, "    ' サンプルID (GD)
                sql = sql & "WFINDGDCW, "       ' 状態FLG (GD)
                sql = sql & "WFRESGDCW, "       ' 実績FLG (GD)
                sql = sql & "WFHSGDCW"         ' 保証FLG (GD)
                '' GD追加　05/01/17 ooba END =======================================>
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                sql = sql & ", EPSMPLIDB1CW, "  ' サンプルID (BMD1E)
                sql = sql & "EPINDB1CW, "       ' 状態FLG (BMD1E)
                sql = sql & "EPRESB1CW, "       ' 実績FLG (BMD1E)
                sql = sql & "EPSMPLIDB2CW, "    ' サンプルID (BMD2E)
                sql = sql & "EPINDB2CW, "       ' 状態FLG (BMD2E)
                sql = sql & "EPRESB2CW, "       ' 実績FLG (BMD2E)
                sql = sql & "EPSMPLIDB3CW, "    ' サンプルID (BMD3E)
                sql = sql & "EPINDB3CW, "       ' 状態FLG (BMD3E)
                sql = sql & "EPRESB3CW, "       ' 実績FLG (BMD3E)
                sql = sql & "EPSMPLIDL1CW, "    ' サンプルID (OSF1E)
                sql = sql & "EPINDL1CW, "       ' 状態FLG (OSF1E)
                sql = sql & "EPRESL1CW, "       ' 実績FLG (OSF1E)
                sql = sql & "EPSMPLIDL2CW, "    ' サンプルID (OSF2E)
                sql = sql & "EPINDL2CW, "       ' 状態FLG (OSF2E)
                sql = sql & "EPRESL2CW, "       ' 実績FLG (OSF2E)
                sql = sql & "EPSMPLIDL3CW, "    ' サンプルID (OSF3E)
                sql = sql & "EPINDL3CW, "       ' 状態FLG (OSF3E)
                sql = sql & "EPRESL3CW"         ' 実績FLG (OSF3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                sql = sql & " )"
                sql = sql & " values ('"
                sql = sql & .SXLIDCW & "', '"       ' SXLID
                sql = sql & .SMPKBNCW & "', '"      ' サンプル区分
                sql = sql & .TBKBNCW & "', '"       ' T/B区分
                sql = sql & .REPSMPLIDCW & "', '"   ' サンプルID
                sql = sql & .XTALCW & "', "         ' 結晶番号
                sql = sql & .INPOSCW & ", '"        ' 結晶内位置
                sql = sql & .HINBCW & "', "         ' 品番
                sql = sql & .REVNUMCW & ", '"       ' 製品番号改訂番号
                sql = sql & .FACTORYCW & "', '"     ' 工場
                sql = sql & .OPECW & "', '"         ' 操業条件
                sql = sql & .KTKBNCW & "', '"       ' 確定区分
                sql = sql & .SMCRYNUMCW & "', '"    ' サンプルブロックID
                sql = sql & .WFSMPLIDRSCW & "', '"  ' サンプルID（Rs）
                sql = sql & .WFSMPLIDRS1CW & "', '" ' 推定サンプルID1（Rs）
                sql = sql & .WFSMPLIDRS2CW & "', '" ' 推定サンプルID2（Rs）
                sql = sql & .WFINDRSCW & "', '"     ' 状態FLG（Rs)
                sql = sql & .WFRESRS1CW & "', '"    ' 実績FLG1（Rs)
                sql = sql & .WFRESRS2CW & "', '"    ' 実績FLG2（Rs)
                sql = sql & .WFSMPLIDOICW & "', '"  ' サンプルID（Oi）
                sql = sql & .WFINDOICW & "', '"     ' 状態FLG（Oi)
                sql = sql & .WFRESOICW & "', '"     ' 実績FLG（Oi)
                sql = sql & .WFSMPLIDB1CW & "', '"  ' サンプルID（B1）
                sql = sql & .WFINDB1CW & "', '"     ' 状態FLG（B1)
                sql = sql & .WFRESB1CW & "', '"     ' 実績FLG（B1)
                sql = sql & .WFSMPLIDB2CW & "', '"  ' サンプルID（B2）
                sql = sql & .WFINDB2CW & "', '"     ' 状態FLG（B2)
                sql = sql & .WFRESB2CW & "', '"     ' 実績FLG（B2)
                sql = sql & .WFSMPLIDB3CW & "', '"  ' サンプルID（B3）
                sql = sql & .WFINDB3CW & "', '"     ' 状態FLG（B3)
                sql = sql & .WFRESB3CW & "', '"     ' 実績FLG（B3)
                sql = sql & .WFSMPLIDL1CW & "', '"  ' サンプルID（L1）
                sql = sql & .WFINDL1CW & "', '"     ' 状態FLG（L1)
                sql = sql & .WFRESL1CW & "', '"     ' 実績FLG（L1)
                sql = sql & .WFSMPLIDL2CW & "', '"  ' サンプルID（L2）
                sql = sql & .WFINDL2CW & "', '"     ' 状態FLG（L2)
                sql = sql & .WFRESL2CW & "', '"     ' 実績FLG（L2)
                sql = sql & .WFSMPLIDL3CW & "', '"  ' サンプルID（L3）
                sql = sql & .WFINDL3CW & "', '"     ' 状態FLG（L3)
                sql = sql & .WFRESL3CW & "', '"     ' 実績FLG（L3)
                sql = sql & .WFSMPLIDL4CW & "', '"  ' サンプルID（L4）
                sql = sql & .WFINDL4CW & "', '"     ' 状態FLG（L4)
                sql = sql & .WFRESL4CW & "', '"     ' 実績FLG（L4)
                sql = sql & .WFSMPLIDDSCW & "', '"  ' サンプルID（DS）
                sql = sql & .WFINDDSCW & "', '"     ' 状態FLG（DS)
                sql = sql & .WFRESDSCW & "', '"     ' 実績FLG（DS)
                sql = sql & .WFSMPLIDDZCW & "', '"  ' サンプルID（DZ）
                sql = sql & .WFINDDZCW & "', '"     ' 状態FLG（DZ)
                sql = sql & .WFRESDZCW & "', '"     ' 実績FLG（DZ)
                sql = sql & .WFSMPLIDSPCW & "', '"  ' サンプルID（SP）
                sql = sql & .WFINDSPCW & "', '"     ' 状態FLG（SP)
                sql = sql & .WFRESSPCW & "', '"     ' 実績FLG（SP)
                sql = sql & .WFSMPLIDDO1CW & "', '" ' サンプルID（DO1）
                sql = sql & .WFINDDO1CW & "', '"    ' 状態FLG（DO1)
                sql = sql & .WFRESDO1CW & "', '"    ' 実績FLG（DO1)
                sql = sql & .WFSMPLIDDO2CW & "', '" ' サンプルID（DO2）
                sql = sql & .WFINDDO2CW & "', '"    ' 状態FLG（DO2)
                sql = sql & .WFRESDO2CW & "', '"    ' 実績FLG（DO2)
                sql = sql & .WFSMPLIDDO3CW & "', '" ' サンプルID（DO3）
                sql = sql & .WFINDDO3CW & "', '"    ' 状態FLG（DO3)
                sql = sql & .WFRESDO3CW & "', '"    ' 実績FLG（DO3)
                sql = sql & .WFSMPLIDOT1CW & "', '" ' サンプルID（OT1）
                sql = sql & .WFSMPLIDOT2CW & "', '" ' サンプルID（OT2）
                'add start 2003/05/21 hitec)matsumoto -------------------------
                sql = sql & .WFINDOT1CW & "', '"    ' 状態FLG（OT1)
                sql = sql & .WFRESOT1CW & "', '"    ' 実績FLG（OT1)
                sql = sql & .WFINDOT2CW & "', '"    ' 状態FLG（OT2)
                sql = sql & .WFRESOT2CW & "', '"    ' 実績FLG（OT2)
                'add end   2003/05/21 hitec)matsumoto -------------------------
                sql = sql & .WFSMPLIDAOICW & "', '" ' サンプルID（AOi）
                sql = sql & .WFINDAOICW & "', '"    ' 状態FLG（AOi）
                sql = sql & .WFRESAOICW & "', "     ' 実績FLG（AOi）
                sql = sql & .SMPLNUMCW & ", '"      ' サンプル枚数
                sql = sql & .SMPLPATCW & "', '"     ' サンプルパターン
                sql = sql & .LIVKCW & "', '"        ' 生死区分
                sql = sql & .TSTAFFCW & "', "       ' 登録社員ID
                sql = sql & "sysdate, '"            ' 登録日付
                sql = sql & .KSTAFFCW & "', "       ' 更新社員ID
                sql = sql & "sysdate, "             ' 更新日付
                sql = sql & "'0', "                 ' 送信フラグ
'                sql = sql & "sysdate)"              ' 送信日付
                '' GD追加　05/01/17 ooba START =====================================>
                sql = sql & "sysdate, '"            ' 送信日付
                sql = sql & .WFSMPLIDGDCW & "', '"  ' サンプルID (GD)
                sql = sql & .WFINDGDCW & "', '"     ' 状態FLG (GD)
                sql = sql & .WFRESGDCW & "', '"     ' 実績FLG (GD)
                sql = sql & .WFHSGDCW & "', '"      ' 保証FLG (GD)
                '' GD追加　05/01/17 ooba END =======================================>
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                sql = sql & .EPSMPLIDB1CW & "', '"  ' サンプルID (BMD1E)
                sql = sql & .EPINDB1CW & "', '"     ' 状態FLG (BMD1E)
                sql = sql & .EPRESB1CW & "', '"     ' 実績FLG (BMD1E)
                sql = sql & .EPSMPLIDB2CW & "', '"  ' サンプルID (BMD2E)
                sql = sql & .EPINDB2CW & "', '"     ' 状態FLG (BMD2E)
                sql = sql & .EPRESB2CW & "', '"     ' 実績FLG (BMD2E)
                sql = sql & .EPSMPLIDB3CW & "', '"  ' サンプルID (BMD3E)
                sql = sql & .EPINDB3CW & "', '"     ' 状態FLG (BMD3E)
                sql = sql & .EPRESB3CW & "', '"       ' 実績FLG (BMD3E)
                sql = sql & .EPSMPLIDL1CW & "', '"  ' サンプルID (OSF1E)
                sql = sql & .EPINDL1CW & "', '"     ' 状態FLG (OSF1E)
                sql = sql & .EPRESL1CW & "', '"     ' 実績FLG (OSF1E)
                sql = sql & .EPSMPLIDL2CW & "', '"  ' サンプルID (OSF2E)
                sql = sql & .EPINDL2CW & "', '"     ' 状態FLG (OSF2E)
                sql = sql & .EPRESL2CW & "', '"     ' 実績FLG (OSF2E)
                sql = sql & .EPSMPLIDL3CW & "', '"  ' サンプルID (OSF3E)
                sql = sql & .EPINDL3CW & "', '"     ' 状態FLG (OSF3E)
                sql = sql & .EPRESL3CW & "')"       ' 実績FLG (OSF3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_WfSmp_UpdIns = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_WfSmp_UpdIns = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function


'概要      :WFサンプル管理の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型                 ,説明
'      　　:WFSMP 　　　,I  ,typ_XSDCW   　     ,新サンプル管理（SXL）
'      　　:戻り値      ,O  ,FUNCTION_RETURN　  ,書き込みの成否
'説明      :DBDRV_WfSmp_UpdInsに移行する予定
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_WfSmp_INS(WFSMP() As typ_XSDCW) As FUNCTION_RETURN

    Dim sql     As String
    Dim i       As Long
    Dim sDbName As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_WfSmp_INS"

    DBDRV_WfSmp_INS = FUNCTION_RETURN_SUCCESS

    sDbName = "XSDCW"
    For i = 1 To UBound(WFSMP)
        With WFSMP(i)
'            sql = "insert into TBCME044 ("
'            sql = sql & "CRYNUM, "          ' 結晶番号
'            sql = sql & "INGOTPOS, "        ' 結晶内位置
'            sql = sql & "SMPKBN, "          ' サンプル区分
'            sql = sql & "SMPLID, "          ' サンプルID
'            sql = sql & "HINBAN, "          ' 品番
'            sql = sql & "REVNUM, "          ' 製品番号改訂番号
'            sql = sql & "FACTORY, "         ' 工場
'            sql = sql & "OPECOND, "         ' 操業条件
'            sql = sql & "KTKBN, "           ' 確定区分
'            sql = sql & "WFINDRS, "         ' WF検査指示（Rs)
'            sql = sql & "WFINDOI, "         ' WF検査指示（Oi)
'            sql = sql & "WFINDB1, "         ' WF検査指示（B1)
'            sql = sql & "WFINDB2, "         ' WF検査指示（B2)
'            sql = sql & "WFINDB3, "         ' WF検査指示（B3)
'            sql = sql & "WFINDL1, "         ' WF検査指示（L1)
'            sql = sql & "WFINDL2, "         ' WF検査指示（L2)
'            sql = sql & "WFINDL3, "         ' WF検査指示（L3)
'            sql = sql & "WFINDL4, "         ' WF検査指示（L4)
'            sql = sql & "WFINDDS, "         ' WF検査指示（DS)
'            sql = sql & "WFINDDZ, "         ' WF検査指示（DZ)
'            sql = sql & "WFINDSP, "         ' WF検査指示（SP)
'            sql = sql & "WFINDDO1, "        ' WF検査指示（DO1)
'            sql = sql & "WFINDDO2, "        ' WF検査指示（DO2)
'            sql = sql & "WFINDDO3, "        ' WF検査指示（DO3)
'            'add start 2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & "WFINDOT1, "        ' WF検査指示（OT1)
'            sql = sql & "WFINDOT2, "        ' WF検査指示（OT2)
'            'add end   2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & "WFRESRS, "         ' WF検査実績（Rs)
'            sql = sql & "WFRESOI, "         ' WF検査実績（Oi)
'            sql = sql & "WFRESB1, "         ' WF検査実績（B1)
'            sql = sql & "WFRESB2, "         ' WF検査実績（B2)
'            sql = sql & "WFRESB3, "         ' WF検査実績（B3)
'            sql = sql & "WFRESL1, "         ' WF検査実績（L1)
'            sql = sql & "WFRESL2, "         ' WF検査実績（L2)
'            sql = sql & "WFRESL3, "         ' WF検査実績（L3)
'            sql = sql & "WFRESL4, "         ' WF検査実績（L4)
'            sql = sql & "WFRESDS, "         ' WF検査実績（DS)
'            sql = sql & "WFRESDZ, "         ' WF検査実績（DZ)
'            sql = sql & "WFRESSP, "         ' WF検査実績（SP)
'            sql = sql & "WFRESDO1, "        ' WF検査実績（DO1)
'            sql = sql & "WFRESDO2, "        ' WF検査実績（DO2)
'            sql = sql & "WFRESDO3, "        ' WF検査実績（DO3)
'            'add start 2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & "WFRESOT1, "        ' WF検査実績（OT1)
'            sql = sql & "WFRESOT2, "        ' WF検査実績（OT2)
'            'add end   2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & "REGDATE, "         ' 更新日付
'            sql = sql & "UPDDATE, "         ' 更新日付
'            sql = sql & "SENDFLAG, "        ' 送信フラグ
'            sql = sql & "SENDDATE)"         ' 送信日付
'            sql = sql & " values ('"
'            sql = sql & .CRYNUM & "', "     ' 結晶番号
'            sql = sql & .INGOTPOS & ", '"   ' 結晶内位置
'            sql = sql & .SMPKBN & "', '"    ' サンプル区分
'            sql = sql & .SMPLID & "', '"    ' サンプルID
'            sql = sql & .HINBAN & "', "     ' 品番
'            sql = sql & .REVNUM & ", '"     ' 製品番号改訂番号
'            sql = sql & .factory & "', '"   ' 工場
'            sql = sql & .opecond & "', '"   ' 操業条件
'            sql = sql & .KTKBN & "', '"     ' 確定区分
'            sql = sql & .WFINDRS & "', '"   ' WF検査指示（Rs)
'            sql = sql & .WFINDOI & "', '"   ' WF検査指示（Oi)
'            sql = sql & .WFINDB1 & "', '"   ' WF検査指示（B1)
'            sql = sql & .WFINDB2 & "', '"   ' WF検査指示（B2)
'            sql = sql & .WFINDB3 & "', '"   ' WF検査指示（B3)
'            sql = sql & .WFINDL1 & "', '"   ' WF検査指示（L1)
'            sql = sql & .WFINDL2 & "', '"   ' WF検査指示（L2)
'            sql = sql & .WFINDL3 & "', '"   ' WF検査指示（L3)
'            sql = sql & .WFINDL4 & "', '"   ' WF検査指示（L4)
'            sql = sql & .WFINDDS & "', '"   ' WF検査指示（DS)
'            sql = sql & .WFINDDZ & "', '"   ' WF検査指示（DZ)
'            sql = sql & .WFINDSP & "', '"   ' WF検査指示（SP)
'            sql = sql & .WFINDDO1 & "', '"  ' WF検査指示（DO1)
'            sql = sql & .WFINDDO2 & "', '"  ' WF検査指示（DO2)
'            sql = sql & .WFINDDO3 & "', '"  ' WF検査指示（DO3)
'            'add start 2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & .WFINDOT1 & "', '"  ' WF検査指示（OT1)
'            sql = sql & .WFINDOT2 & "', '"  ' WF検査指示（OT2)
'            'add end   2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & .WFRESRS & "', '"   ' WF検査実績（Rs)
'            sql = sql & .WFRESOI & "', '"   ' WF検査実績（Oi)
'            sql = sql & .WFRESB1 & "', '"   ' WF検査実績（B1)
'            sql = sql & .WFRESB2 & "', '"   ' WF検査実績（B2)
'            sql = sql & .WFRESB3 & "', '"   ' WF検査実績（B3)
'            sql = sql & .WFRESL1 & "', '"   ' WF検査実績（L1)
'            sql = sql & .WFRESL2 & "', '"   ' WF検査実績（L2)
'            sql = sql & .WFRESL3 & "', '"   ' WF検査実績（L3)
'            sql = sql & .WFRESL4 & "', '"   ' WF検査実績（L4)
'            sql = sql & .WFRESDS & "', '"   ' WF検査実績（DS)
'            sql = sql & .WFRESDZ & "', '"   ' WF検査実績（DZ)
'            sql = sql & .WFRESSP & "', '"   ' WF検査実績（SP)
'            sql = sql & .WFRESDO1 & "', '"  ' WF検査実績（DO1)
'            sql = sql & .WFRESDO2 & "', '"  ' WF検査実績（DO2)
'            sql = sql & .WFRESDO3 & "', '"   ' WF検査実績（DO3)
'            'add start 2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & .WFRESOT1 & "', '"  ' WF検査実績（OT1)
'            sql = sql & .WFRESOT2 & "',"  ' WF検査実績（OT2)
'            'add end   2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & "sysdate, "         ' 登録日付
'            sql = sql & "sysdate, "         ' 更新日付
'            sql = sql & "'0', "             ' 送信フラグ
'            sql = sql & "sysdate)"          ' 送信日付


                sql = "insert into XSDCW ("
                sql = sql & "SXLIDCW, "             ' SXLID
                sql = sql & "SMPKBNCW, "            ' サンプル区分
                sql = sql & "TBKBNCW, "             ' T/B区分
                sql = sql & "REPSMPLIDCW, "         ' サンプルID
                sql = sql & "XTALCW, "              ' 結晶番号
                sql = sql & "INPOSCW, "             ' 結晶内位置
                sql = sql & "HINBCW, "              ' 品番
                sql = sql & "REVNUMCW, "            ' 製品番号改訂番号
                sql = sql & "FACTORYCW, "           ' 工場
                sql = sql & "OPECW, "               ' 操業条件
                sql = sql & "KTKBNCW, "             ' 確定区分
                sql = sql & "SMCRYNUMCW, "          ' サンプルブロックID
                sql = sql & "WFSMPLIDRSCW, "        ' サンプルID(Rs)
                sql = sql & "WFSMPLIDRS1CW, "       ' 推定サンプルID1（Rs）
                sql = sql & "WFSMPLIDRS2CW, "       ' 推定サンプルID2（Rs）
                sql = sql & "WFINDRSCW, "           ' 状態FLG（Rs)
                sql = sql & "WFRESRS1CW, "          ' 実績FLG1（Rs)
                sql = sql & "WFRESRS2CW, "          ' 実績FLG2（Rs)
                sql = sql & "WFSMPLIDOICW, "        ' サンプルID（Oi）
                sql = sql & "WFINDOICW, "           ' 状態FLG（Oi)
                sql = sql & "WFRESOICW, "           ' 実績FLG（Oi)
                sql = sql & "WFSMPLIDB1CW, "        ' サンプルID（B1）
                sql = sql & "WFINDB1CW, "           ' 状態FLG（B1)
                sql = sql & "WFRESB1CW, "           ' 実績FLG（B1)
                sql = sql & "WFSMPLIDB2CW, "        ' サンプルID（B2）
                sql = sql & "WFINDB2CW, "           ' 状態FLG（B2)
                sql = sql & "WFRESB2CW, "           ' 実績FLG（B2)
                sql = sql & "WFSMPLIDB3CW, "        ' サンプルID（B3）
                sql = sql & "WFINDB3CW, "           ' 状態FLG（B3)
                sql = sql & "WFRESB3CW, "           ' 実績FLG（B3)
                sql = sql & "WFSMPLIDL1CW, "        ' サンプルID（L1）
                sql = sql & "WFINDL1CW, "           ' 状態FLG（L1)
                sql = sql & "WFRESL1CW, "           ' 実績FLG（L1)
                sql = sql & "WFSMPLIDL2CW, "        ' サンプルID（L2）
                sql = sql & "WFINDL2CW, "           ' 状態FLG（L2)
                sql = sql & "WFRESL2CW, "           ' 実績FLG（L2)
                sql = sql & "WFSMPLIDL3CW, "        ' サンプルID（L3）
                sql = sql & "WFINDL3CW, "           ' 状態FLG（L3)
                sql = sql & "WFRESL3CW, "           ' 実績FLG（L3)
                sql = sql & "WFSMPLIDL4CW, "        ' サンプルID（L4）
                sql = sql & "WFINDL4CW, "           ' 状態FLG（L4)
                sql = sql & "WFRESL4CW, "           ' 実績FLG（L4)
                sql = sql & "WFSMPLIDDSCW, "        ' サンプルID（DS）
                sql = sql & "WFINDDSCW, "           ' 状態FLG（DS)
                sql = sql & "WFRESDSCW, "           ' 実績FLG（DS)
                sql = sql & "WFSMPLIDDZCW, "        ' サンプルID（DZ）
                sql = sql & "WFINDDZCW, "           ' 状態FLG（DZ)
                sql = sql & "WFRESDZCW, "           ' 実績FLG（DZ)
                sql = sql & "WFSMPLIDSPCW, "        ' サンプルID（SP）
                sql = sql & "WFINDSPCW, "           ' 状態FLG（SP)
                sql = sql & "WFRESSPCW, "           ' 実績FLG（SP)
                sql = sql & "WFSMPLIDDO1CW,"        ' サンプルID（DO1）
                sql = sql & "WFINDDO1CW, "          ' 状態FLG（DO1)
                sql = sql & "WFRESDO1CW, "          ' 実績FLG（DO1)
                sql = sql & "WFSMPLIDDO2CW, "       ' サンプルID（DO2）
                sql = sql & "WFINDDO2CW, "          ' 状態FLG（DO2)
                sql = sql & "WFRESDO2CW, "          ' 実績FLG（DO2)
                sql = sql & "WFSMPLIDDO3CW, "       ' サンプルID（DO3）
                sql = sql & "WFINDDO3CW, "          ' 状態FLG（DO3)
                sql = sql & "WFRESDO3CW, "          ' 実績FLG（DO3)
                sql = sql & "WFSMPLIDOT1CW, "       ' サンプルID（OT1）
               'add start 2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFINDOT1CW, "          ' 状態FLG（OT1)
                sql = sql & "WFRESOT1CW, "          ' 実績FLG（OT1)
                sql = sql & "WFSMPLIDOT2CW, "       ' サンプルID（OT2）
                sql = sql & "WFINDOT2CW, "          ' 状態FLG（OT2)
                sql = sql & "WFRESOT2CW, "          ' 実績FLG（OT2)
               'add end   2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFSMPLIDAOICW, "       ' サンプルID（AOi）
                sql = sql & "WFINDAOICW, "          ' 状態FLG（AOi）
                sql = sql & "WFRESAOICW, "          ' 実績FLG（AOi）
                sql = sql & "SMPLNUMCW, "           ' サンプル枚数
                sql = sql & "SMPLPATCW, "           ' サンプルパターン
                sql = sql & "TSTAFFCW,"             ' 登録社員ID
                sql = sql & "TDAYCW, "              ' 登録日付
                sql = sql & "KSTAFFCW, "            ' 更新社員ID
                sql = sql & "KDAYCW, "              ' 更新日付
                sql = sql & "SNDKCW, "              ' 送信フラグ
                sql = sql & "SNDDAYCW,"             ' 送信日付
'                sql = sql & "LIVKCW)"               ' 生死区分
                '' GD追加　05/01/17 ooba START =====================================>
                sql = sql & "LIVKCW, "              ' 生死区分
                sql = sql & "WFSMPLIDGDCW, "        ' サンプルID (GD)
                sql = sql & "WFINDGDCW, "           ' 状態FLG (GD)
                sql = sql & "WFRESGDCW, "           ' 実績FLG (GD)
'                sql = sql & "WFHSGDCW)"             ' 保証FLG (GD) 2006/08/15 Del エピ先行評価追加対応 SMP)kondoh
                sql = sql & "WFHSGDCW, "            ' 保証FLG (GD)  2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
                '' GD追加　05/01/17 ooba END =======================================>
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                sql = sql & "EPSMPLIDB1CW, "        ' サンプルID (BMD1E)
                sql = sql & "EPINDB1CW, "           ' 状態FLG (BMD1E)
                sql = sql & "EPRESB1CW, "           ' 実績FLG (BMD1E)
                sql = sql & "EPSMPLIDB2CW, "        ' サンプルID (BMD2E)
                sql = sql & "EPINDB2CW, "           ' 状態FLG (BMD2E)
                sql = sql & "EPRESB2CW, "           ' 実績FLG (BMD2E)
                sql = sql & "EPSMPLIDB3CW, "        ' サンプルID (BMD3E)
                sql = sql & "EPINDB3CW, "           ' 状態FLG (BMD3E)
                sql = sql & "EPRESB3CW, "           ' 実績FLG (BMD3E)
                sql = sql & "EPSMPLIDL1CW, "        ' サンプルID (OSF1E)
                sql = sql & "EPINDL1CW, "           ' 状態FLG (OSF1E)
                sql = sql & "EPRESL1CW, "           ' 実績FLG (OSF1E)
                sql = sql & "EPSMPLIDL2CW, "        ' サンプルID (OSF2E)
                sql = sql & "EPINDL2CW, "           ' 状態FLG (OSF2E)
                sql = sql & "EPRESL2CW, "           ' 実績FLG (OSF2E)
                sql = sql & "EPSMPLIDL3CW, "        ' サンプルID (OSF3E)
                sql = sql & "EPINDL3CW, "           ' 状態FLG (OSF3E)
                sql = sql & "EPRESL3CW"             ' 実績FLG (OSF3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                sql = sql & ")"                     ' 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
                sql = sql & " values ('"
                sql = sql & .SXLIDCW & "', '"       ' SXLID
                sql = sql & .SMPKBNCW & "', '"      ' サンプル区分
                sql = sql & .TBKBNCW & "', '"       ' T/B区分
                sql = sql & .REPSMPLIDCW & "', '"   ' サンプルID
                sql = sql & .XTALCW & "', "         ' 結晶番号
                sql = sql & .INPOSCW & ", '"        ' 結晶内位置
                sql = sql & .HINBCW & "', "         ' 品番
                sql = sql & .REVNUMCW & ", '"       ' 製品番号改訂番号
                sql = sql & .FACTORYCW & "', '"     ' 工場
                sql = sql & .OPECW & "', '"         ' 操業条件
                sql = sql & .KTKBNCW & "', '"       ' 確定区分
                sql = sql & .SMCRYNUMCW & "', '"    ' サンプルブロックID
                sql = sql & .WFSMPLIDRSCW & "', '"  ' サンプルID（Rs）
                sql = sql & .WFSMPLIDRS1CW & "', '" ' 推定サンプルID1（Rs）
'               sql = sql & "Null, "                ' 推定サンプルID1（Rs）
                sql = sql & .WFSMPLIDRS2CW & "', '" ' 推定サンプルID2（Rs）
'               sql = sql & "Null, '"               ' 推定サンプルID2（Rs）
                sql = sql & .WFINDRSCW & "', '"     ' 状態FLG（Rs)
                sql = sql & .WFRESRS1CW & "', "     ' 実績FLG1（Rs)
                sql = sql & "Null, '"               ' 実績FLG2（Rs)
                sql = sql & .WFSMPLIDOICW & "', '"  ' サンプルID（Oi）
                sql = sql & .WFINDOICW & "', '"     ' 状態FLG（Oi)
                sql = sql & .WFRESOICW & "', '"     ' 実績FLG（Oi)
                sql = sql & .WFSMPLIDB1CW & "', '"  ' サンプルID（B1）
                sql = sql & .WFINDB1CW & "', '"     ' 状態FLG（B1)
                sql = sql & .WFRESB1CW & "', '"     ' 実績FLG（B1)
                sql = sql & .WFSMPLIDB2CW & "', '"  ' サンプルID（B2）
                sql = sql & .WFINDB2CW & "', '"     ' 状態FLG（B2)
                sql = sql & .WFRESB2CW & "', '"     ' 実績FLG（B2)
                sql = sql & .WFSMPLIDB3CW & "', '"  ' サンプルID（B3）
                sql = sql & .WFINDB3CW & "', '"     ' 状態FLG（B3)
                sql = sql & .WFRESB3CW & "', '"     ' 実績FLG（B3)
                sql = sql & .WFSMPLIDL1CW & "', '"  ' サンプルID（L1）
                sql = sql & .WFINDL1CW & "', '"     ' 状態FLG（L1)
                sql = sql & .WFRESL1CW & "', '"     ' 実績FLG（L1)
                sql = sql & .WFSMPLIDL2CW & "', '"  ' サンプルID（L2）
                sql = sql & .WFINDL2CW & "', '"     ' 状態FLG（L2)
                sql = sql & .WFRESL2CW & "', '"     ' 実績FLG（L2)
                sql = sql & .WFSMPLIDL3CW & "', '"  ' サンプルID（L3）
                sql = sql & .WFINDL3CW & "', '"     ' 状態FLG（L3)
                sql = sql & .WFRESL3CW & "', '"     ' 実績FLG（L3)
                sql = sql & .WFSMPLIDL4CW & "', '"  ' サンプルID（L4）
                sql = sql & .WFINDL4CW & "', '"     ' 状態FLG（L4)
                sql = sql & .WFRESL4CW & "', '"     ' 実績FLG（L4)
                sql = sql & .WFSMPLIDDSCW & "', '"  ' サンプルID（DS）
                sql = sql & .WFINDDSCW & "', '"     ' 状態FLG（DS)
                sql = sql & .WFRESDSCW & "', '"     ' 実績FLG（DS)
                sql = sql & .WFSMPLIDDZCW & "', '"  ' サンプルID（DZ）
                sql = sql & .WFINDDZCW & "', '"     ' 状態FLG（DZ)
                sql = sql & .WFRESDZCW & "', '"     ' 実績FLG（DZ)
                sql = sql & .WFSMPLIDSPCW & "', '"  ' サンプルID（SP）
                sql = sql & .WFINDSPCW & "', '"     ' 状態FLG（SP)
                sql = sql & .WFRESSPCW & "', '"     ' 実績FLG（SP)
                sql = sql & .WFSMPLIDDO1CW & "', '" ' サンプルID（DO1）
                sql = sql & .WFINDDO1CW & "', '"    ' 状態FLG（DO1)
                sql = sql & .WFRESDO1CW & "', '"    ' 実績FLG（DO1)
                sql = sql & .WFSMPLIDDO2CW & "', '" ' サンプルID（DO2）
                sql = sql & .WFINDDO2CW & "', '"    ' 状態FLG（DO2)
                sql = sql & .WFRESDO2CW & "', '"    ' 実績FLG（DO2)
                sql = sql & .WFSMPLIDDO3CW & "', '" ' サンプルID（DO3）
                sql = sql & .WFINDDO3CW & "', '"    ' 状態FLG（DO3)
                sql = sql & .WFRESDO3CW & "', '"    ' 実績FLG（DO3)
                sql = sql & .WFSMPLIDOT1CW & "', '" ' サンプルID（OT1）
                sql = sql & .WFINDOT1CW & "', '"    ' 状態FLG（OT1)
                sql = sql & .WFRESOT1CW & "', '"    ' 実績FLG（OT1)
                sql = sql & .WFSMPLIDOT2CW & "', '" ' サンプルID（OT2）
                sql = sql & .WFINDOT2CW & "', '"    ' 状態FLG（OT2)
                sql = sql & .WFRESOT2CW & "', '"    ' 実績FLG（OT2)
                sql = sql & .WFSMPLIDAOICW & "', '" ' サンプルID（AOi）
''              sql = sql & "NULL, "                ' サンプルID（AOi）
''              sql = sql & "NULL, "                ' 状態FLG（AOi）
                sql = sql & .WFINDAOICW & "', '"    ' 状態FLG（AOi）
                sql = sql & .WFRESAOICW & "', "     ' 実績FLG（AOi）
''              sql = sql & "NULL, "                ' 実績FLG（AOi）
                sql = sql & "NULL, '"               ' サンプル枚数
                sql = sql & .SMPLPATCW & "', '"     ' サンプルパターン
''              sql = sql & "NULL, "                ' サンプル枚数
''              sql = sql & "NULL, '"               ' サンプルパターン
                sql = sql & .TSTAFFCW & "', "       ' 登録社員ID
                sql = sql & "sysdate, '"            ' 登録日付
                sql = sql & .KSTAFFCW & "', "       ' 更新社員ID
                sql = sql & "sysdate, "             ' 更新日付
                sql = sql & "'0', "                 ' 送信フラグ
                sql = sql & "sysdate,"              ' 送信日付
'                sql = sql & "'0')"                  ' 生死区分
                '' GD追加　05/01/17 ooba START =====================================>
                sql = sql & "'0', '"                ' 生死区分
                sql = sql & .WFSMPLIDGDCW & "', '"  ' サンプルID (GD)
                sql = sql & .WFINDGDCW & "', '"     ' 状態FLG (GD)
                sql = sql & .WFRESGDCW & "', '"     ' 実績FLG (GD)
'                sql = sql & .WFHSGDCW & "')"        ' 保証FLG (GD)  2006/08/15 Del エピ先行評価追加対応 SMP)kondoh
                sql = sql & .WFHSGDCW & "', '"      ' 保証FLG (GD)  2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
                '' GD追加　05/01/17 ooba END =======================================>
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                sql = sql & .EPSMPLIDB1CW & "', '"  ' サンプルID (BMD1E)
                sql = sql & .EPINDB1CW & "', '"     ' 状態FLG (BMD1E)
                sql = sql & .EPRESB1CW & "', '"     ' 実績FLG (BMD1E)
                sql = sql & .EPSMPLIDB2CW & "', '"  ' サンプルID (BMD2E)
                sql = sql & .EPINDB2CW & "', '"     ' 状態FLG (BMD2E)
                sql = sql & .EPRESB2CW & "', '"     ' 実績FLG (BMD2E)
                sql = sql & .EPSMPLIDB3CW & "', '"  ' サンプルID (BMD3E)
                sql = sql & .EPINDB3CW & "', '"     ' 状態FLG (BMD3E)
                sql = sql & .EPRESB3CW & "', '"     ' 実績FLG (BMD3E)
                sql = sql & .EPSMPLIDL1CW & "', '"  ' サンプルID (OSF1E)
                sql = sql & .EPINDL1CW & "', '"     ' 状態FLG (OSF1E)
                sql = sql & .EPRESL1CW & "', '"     ' 実績FLG (OSF1E)
                sql = sql & .EPSMPLIDL2CW & "', '"  ' サンプルID (OSF2E)
                sql = sql & .EPINDL2CW & "', '"     ' 状態FLG (OSF2E)
                sql = sql & .EPRESL2CW & "', '"     ' 実績FLG (OSF2E)
                sql = sql & .EPSMPLIDL3CW & "', '"  ' サンプルID (OSF3E)
                sql = sql & .EPINDL3CW & "', '"     ' 状態FLG (OSF3E)
                sql = sql & .EPRESL3CW & "')"       ' 実績FLG (OSF3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
        
                '' WriteDBLog sql, sDBName
        End With
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_WfSmp_INS = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_WfSmp_INS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function




'概要      :SXL管理の挿入／更新
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型               ,説明
'      　　:SXLOld　　　,I  ,typ_TBCME042   　,SXL管理（旧）
'      　　:SXLNew　　　,I  ,typ_TBCME042   　,SXL管理（新）
'      　　:戻り値      ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :古いレコードをみて更新か挿入かを判別する
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_SXL_UpdIns(SXLOld() As typ_TBCME042, SXLNew() As typ_TBCME042) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_SXL_UpdIns"

    DBDRV_SXL_UpdIns = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(SXLNew)
        With SXLNew(i)
            lFlg = False
            For j = 1 To UBound(SXLOld)
                If SXLOld(j).CRYNUM = .CRYNUM And _
                   SXLOld(j).INGOTPOS = .INGOTPOS Then
                    sql = "update TBCME042 set "
                    sql = sql & "CRYNUM='" & .CRYNUM & "', "            ' 結晶番号
                    sql = sql & "INGOTPOS=" & .INGOTPOS & ", "          ' 結晶内開始位置
                    sql = sql & "LENGTH=" & .Length & ", "              ' 長さ
                    sql = sql & "SXLID='" & .SXLID & "', "              ' SXLID
                    sql = sql & "KRPROCCD='" & .KRPROCCD & "', "        ' 管理工程
                    sql = sql & "NOWPROC='" & .NOWPROC & "', "          ' 現在工程
                    sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "    ' 最終通過管理工程
                    sql = sql & "LASTPASS='" & .LASTPASS & "', "        ' 最終通過工程
                    sql = sql & "DELCLS='" & .DELCLS & "', "            ' 削除区分
                    sql = sql & "LSTATCLS='" & .LSTATCLS & "', "        ' 最終状態区分
                    sql = sql & "HOLDCLS='" & .HOLDCLS & "', "          ' ホールド区分
                    sql = sql & "HINBAN='" & .hinban & "', "            ' 品番
                    sql = sql & "REVNUM=" & .REVNUM & ", "              ' 製品番号改訂番号
                    sql = sql & "FACTORY='" & .factory & "', "          ' 工場
                    sql = sql & "OPECOND='" & .opecond & "', "          ' 操業条件
                    sql = sql & "BDCAUS='" & .BDCAUS & "', "            ' 不良理由
                    sql = sql & "COUNT=" & .COUNT & ", "                ' 枚数
                    sql = sql & "UPDDATE=sysdate, "                     ' 更新日付
                    sql = sql & "SUMMITSENDFLAG='0', "                  ' SUMMIT送信フラグ
                    sql = sql & "SENDFLAG='0'"                          ' 送信フラグ
                    sql = sql & " where CRYNUM='" & .CRYNUM & "'"
                    sql = sql & " and INGOTPOS=" & .INGOTPOS
                    '' WriteDBLog sql
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        GoTo proc_err
                    End If
                    lFlg = True
                    Exit For
                End If
            Next j

            If lFlg <> True Then
                sql = "insert into TBCME042 ("
                sql = sql & "CRYNUM, "              ' 結晶番号
                sql = sql & "INGOTPOS, "            ' 結晶内開始位置
                sql = sql & "LENGTH, "              ' 長さ
                sql = sql & "SXLID, "               ' SXLID
                sql = sql & "KRPROCCD, "            ' 管理工程
                sql = sql & "NOWPROC, "             ' 現在工程
                sql = sql & "LPKRPROCCD, "          ' 最終通過管理工程
                sql = sql & "LASTPASS, "            ' 最終通過工程
                sql = sql & "DELCLS, "              ' 削除区分
                sql = sql & "LSTATCLS, "            ' 最終状態区分
                sql = sql & "HOLDCLS, "             ' ホールド区分
                sql = sql & "HINBAN, "              ' 品番
                sql = sql & "REVNUM, "              ' 製品番号改訂番号
                sql = sql & "FACTORY, "             ' 工場
                sql = sql & "OPECOND, "             ' 操業条件
                sql = sql & "BDCAUS, "              ' 不良理由
                sql = sql & "COUNT, "               ' 枚数
                sql = sql & "REGDATE, "             ' 登録日付
                sql = sql & "UPDDATE, "             ' 更新日付
                sql = sql & "SUMMITSENDFLAG, "      ' SUMMIT送信フラグ
                sql = sql & "SENDFLAG, "            ' 送信フラグ
                sql = sql & "SENDDATE)"             ' 送信日付
                sql = sql & " values ('"
                sql = sql & .CRYNUM & "', "         ' 結晶番号
                sql = sql & .INGOTPOS & ", "        ' 結晶内開始位置
                sql = sql & .Length & ", '"         ' 長さ
                sql = sql & .SXLID & "', '"         ' SXLID
                sql = sql & .KRPROCCD & "', '"      ' 管理工程
                sql = sql & .NOWPROC & "', '"       ' 現在工程
                sql = sql & .LPKRPROCCD & "', '"    ' 最終通過管理工程
                sql = sql & .LASTPASS & "', '"      ' 最終通過工程
                sql = sql & .DELCLS & "', '"        ' 削除区分
                sql = sql & .LSTATCLS & "', '"      ' 最終状態区分
                sql = sql & .HOLDCLS & "', '"       ' ホールド区分
                sql = sql & .hinban & "', "         ' 品番
                sql = sql & .REVNUM & ", '"         ' 製品番号改訂番号
                sql = sql & .factory & "', '"       ' 工場
                sql = sql & .opecond & "', '"       ' 操業条件
                sql = sql & .BDCAUS & "', "         ' 不良理由
                sql = sql & .COUNT & ", "           ' 枚数
                sql = sql & "sysdate, "             ' 登録日付
                sql = sql & "sysdate, "             ' 更新日付
                sql = sql & "'0', "                 ' SUMMIT送信フラグ
                sql = sql & "'0', "                 ' 送信フラグ
                sql = sql & "sysdate)"              ' 送信日付
                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_SXL_UpdIns = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_SXL_UpdIns = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :SXL管理の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型               ,説明
'      　　:SXL   　　　,I  ,typ_TBCME042   　,SXL管理
'      　　:戻り値      ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :DBDRV_SXL_UpdInsに移行する予定
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_SXL_INS(SXL() As typ_TBCME042) As FUNCTION_RETURN

    Dim sql As String
    Dim i As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_SXL_INS"

    DBDRV_SXL_INS = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(SXL)
        If SXL(i).Length > 0 Then
            sql = "delete from TBCME042 where CRYNUM='" & SXL(i).CRYNUM & "' and INGOTPOS=" & SXL(i).INGOTPOS
            OraDB.ExecuteSQL sql
            With SXL(i)
                sql = "insert into TBCME042 ("
                sql = sql & "CRYNUM, "              ' 結晶番号
                sql = sql & "INGOTPOS, "            ' 結晶内開始位置
                sql = sql & "LENGTH, "              ' 長さ
                sql = sql & "SXLID, "               ' SXLID
                sql = sql & "KRPROCCD, "            ' 管理工程
                sql = sql & "NOWPROC, "             ' 現在工程
                sql = sql & "LPKRPROCCD, "          ' 最終通過管理工程
                sql = sql & "LASTPASS, "            ' 最終通過工程
                sql = sql & "DELCLS, "              ' 削除区分
                sql = sql & "LSTATCLS, "            ' 最終状態区分
                sql = sql & "HOLDCLS, "             ' ホールド区分
                sql = sql & "HINBAN, "              ' 品番
                sql = sql & "REVNUM, "              ' 製品番号改訂番号
                sql = sql & "FACTORY, "             ' 工場
                sql = sql & "OPECOND, "             ' 操業条件
                sql = sql & "BDCAUS, "              ' 不良理由
                sql = sql & "COUNT, "               ' 枚数
                sql = sql & "REGDATE, "             ' 登録日付
                sql = sql & "UPDDATE, "             ' 更新日付
                sql = sql & "SUMMITSENDFLAG, "      ' SUMMIT送信フラグ
                sql = sql & "SENDFLAG, "            ' 送信フラグ
                sql = sql & "SENDDATE)"             ' 送信日付
                sql = sql & " values ('"
                sql = sql & .CRYNUM & "', "         ' 結晶番号
                sql = sql & .INGOTPOS & ", "        ' 結晶内開始位置
                sql = sql & .Length & ", '"         ' 長さ
                sql = sql & .SXLID & "', '"         ' SXLID
                sql = sql & .KRPROCCD & "', '"      ' 管理工程
                sql = sql & .NOWPROC & "', '"       ' 現在工程
                sql = sql & .LPKRPROCCD & "', '"    ' 最終通過管理工程
                sql = sql & .LASTPASS & "', '"      ' 最終通過工程
                sql = sql & .DELCLS & "', '"        ' 削除区分
                sql = sql & .LSTATCLS & "', '"      ' 最終状態区分
                sql = sql & .HOLDCLS & "', '"       ' ホールド区分
                sql = sql & .hinban & "', "         ' 品番
                sql = sql & .REVNUM & ", '"         ' 製品番号改訂番号
                sql = sql & .factory & "', '"       ' 工場
                sql = sql & .opecond & "', '"       ' 操業条件
                sql = sql & .BDCAUS & "', "         ' 不良理由
                sql = sql & .COUNT & ", "           ' 枚数
                sql = sql & "sysdate, "             ' 登録日付
                sql = sql & "sysdate, "             ' 更新日付
                sql = sql & "'0', "                 ' SUMMIT送信フラグ
                sql = sql & "'0', "                 ' 送信フラグ
                sql = sql & "sysdate)"              ' 送信日付
            End With
            '' WriteDBLog sql
            If OraDB.ExecuteSQL(sql) <= 0 Then
                DBDRV_SXL_INS = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End If
    Next i

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_SXL_INS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :測定評価方法指示の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:SokuSizi　　　,I  ,typ_TBCMY003   　,測定評価方法指示
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_SokuSizi_Ins(SokuSizi() As typ_TBCMY003) As FUNCTION_RETURN

    Dim sql As String
    Dim i As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_SokuSizi_Ins"

    '' 測定評価方法指示の挿入
    For i = 1 To UBound(SokuSizi)
        With SokuSizi(i)
            sql = " insert into TBCMY003 ("
            sql = sql & "SAMPLEID, "                ' サンプルID
            sql = sql & "OSITEM, "                  ' 評価項目
            sql = sql & "TRANCNT, "                 ' 処理回数
            sql = sql & "SAMPLEKB, "                ' サンプル区分
            sql = sql & "MAISU, "                   ' 評価枚数
            sql = sql & "SPEC, "                    ' 規格値
            sql = sql & "NETSU, "                   ' 熱処理条件
            sql = sql & "ET, "                      ' エッチング条件
            sql = sql & "MES, "                     ' 計測方法
            sql = sql & "DKAN, "                    ' ＤＫアニール条件
            sql = sql & "TXID, "                    ' トランザクションID
            sql = sql & "REGDATE, "                 ' 登録日付
            sql = sql & "SENDFLAG, "                ' 送信フラグ
            sql = sql & "SENDDATE, "                ' 送信日付
            sql = sql & "PLANTCAT, "                ' 向先 2007/08/31 SPK Tsutsumi Add
            '06/06/08 ooba START =======================================================>
            sql = sql & "FEPUA, "                   ' SPV_Fe_PUA値
            sql = sql & "FEPUAPCT, "                ' SPV_Fe_PUA％値
            sql = sql & "FESTD, "                   ' SPV_Fe_STD
            sql = sql & "DIFFPUA, "                 ' SPV_拡散長_PUA値
            sql = sql & "DIFFPUAPCT, "              ' SPV_拡散長_PUA％値
            sql = sql & "NRPUA, "                   ' SPV_NR_PUA値
            sql = sql & "NRPUAPCT, "                ' SPV_NR_PUA%値
            sql = sql & "NRSTD) "                   ' SPV_NR_STD
            '06/06/08 ooba END =========================================================>
            sql = sql & " select '"
            sql = sql & .SAMPLEID & "', '"          ' サンプルID
            sql = sql & .OSITEM & "', "             ' 評価項目
            sql = sql & "nvl(max(TRANCNT),0)+1, '"  ' 処理回数
            sql = sql & .SAMPLEKB & "', '"           ' サンプル区分
            'sql = sql & "'1', '"                    ' 評価枚数   2004/06/23
            sql = sql & .MAISU & "', '"        ' 評価枚数
            sql = sql & .Spec & "', '"              ' 規格値
            sql = sql & .NETSU & "', '"             ' 熱処理条件
            sql = sql & .ET & "', '"                ' エッチング条件
            sql = sql & .MES & "', '"               ' 計測方法
            sql = sql & .DKAN & "', "               ' ＤＫアニール条件
            sql = sql & "'TX851I', "                ' トランザクションID
            sql = sql & "sysdate, "                 ' 登録日付
            sql = sql & "'" & .SENDFLAG & "', "     ' 送信フラグ
            sql = sql & "sysdate, "                 ' 送信日付
            sql = sql & "'" & .MUKESAKI & "', "     ' 向先 2007/08/31 SPK Tsutsumi Add
            '06/06/08 ooba START =======================================================>
            If IsNumeric(.FEPUA) Then sql = sql & CDbl(.FEPUA) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.FEPUAPCT) Then sql = sql & CDbl(.FEPUAPCT) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.FESTD) Then sql = sql & CDbl(.FESTD) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.DIFFPUA) Then sql = sql & CDbl(.DIFFPUA) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.DIFFPUAPCT) Then sql = sql & CDbl(.DIFFPUAPCT) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.NRPUA) Then sql = sql & CDbl(.NRPUA) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.NRPUAPCT) Then sql = sql & CDbl(.NRPUAPCT) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.NRSTD) Then sql = sql & CDbl(.NRSTD) Else sql = sql & "NULL "
            '06/06/08 ooba END =========================================================>
            sql = sql & " from TBCMY003"
            sql = sql & " where SAMPLEID='" & .SAMPLEID & "'"
            sql = sql & " and OSITEM='" & .OSITEM & "'"
            sql = sql & " and SPEC='" & .Spec & "'"
        End With
        '' WriteDBLog sql
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_SokuSizi_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    DBDRV_SokuSizi_Ins = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_SokuSizi_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :転用実績の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:Tenyou　 　　,I  ,typ_TBCMJ013   　,転用実績
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2002/01/15  作成 S.Sano
Public Function DBDRV_Tenyou_Ins(Tenyou As typ_TBCMJ013) As FUNCTION_RETURN

    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_Tenyou_Ins"

    '' 振替廃棄実績の挿入
    With Tenyou
        sql = "insert into TBCMJ013 ("
        sql = sql & "CRYNUM, "                  ' 結晶番号
        sql = sql & "INGOTPOS, "                ' インゴット位置
        sql = sql & "TRANCNT, "                 ' 処理回数
        sql = sql & "LENGTH, "                  ' 長さ
        sql = sql & "KRPROCCD, "                ' 管理工程コード
        sql = sql & "PROCCODE, "                ' 工程コード
        sql = sql & "DUNWNUM, "                 ' 転用先品番
        sql = sql & "DUNWREV, "                 ' 転用先品番 製品番号改訂番号
        sql = sql & "DUNWFACT, "                ' 転用先品番 工場
        sql = sql & "DUNWOPCD, "                ' 転用先品番 操業条件
        sql = sql & "DUOGNUM, "                 ' 転用元品番
        sql = sql & "DUOGREV, "                 ' 転用元品番 製品番号改訂番号
        sql = sql & "DUOGFACT, "                ' 転用元品番 工場
        sql = sql & "DUOGOPCD, "                ' 転用元品番 操業条件
        sql = sql & "TSTAFFID, "                ' 登録社員ID
        sql = sql & "REGDATE, "                 ' 登録日付
        sql = sql & "KSTAFFID, "                ' 更新社員ID
        sql = sql & "UPDDATE, "                 ' 更新日付
        sql = sql & "SENDFLAG)"                 ' 送信フラグ
        sql = sql & " select '"
        sql = sql & .CRYNUM & "', "             ' 結晶番号
        sql = sql & .INGOTPOS & ", "            ' インゴット位置
        sql = sql & "nvl(max(TRANCNT),0)+1, "   ' 処理回数
        sql = sql & .Length & ", '"             ' 長さ
        sql = sql & .KRPROCCD & "', '"          ' 管理工程コード
        sql = sql & .PROCCODE & "', '"          ' 工程コード
        sql = sql & .DUNWNUM & "', "            ' 転用先品番
        sql = sql & .DUNWREV & ", '"            ' 転用先品番 製品番号改訂番号
        sql = sql & .DUNWFACT & "', '"          ' 転用先品番 工場
        sql = sql & .DUNWOPCD & "', '"          ' 転用先品番 操業条件
        sql = sql & .DUOGNUM & "', "            ' 転用元品番
        sql = sql & .DUOGREV & ", '"            ' 転用元品番 製品番号改訂番号
        sql = sql & .DUOGFACT & "', '"          ' 転用元品番 工場
        sql = sql & .DUOGOPCD & "', '"          ' 転用元品番 操業条件
        sql = sql & .TSTAFFID & "', "           ' 登録社員ID
        sql = sql & "sysdate, '"                ' 登録日付
        sql = sql & .KSTAFFID & "', "           ' 更新社員ID
        sql = sql & "sysdate, "                 ' 更新日付
        sql = sql & "'0'"                       ' 送信フラグ
        sql = sql & " from TBCMJ013"
        sql = sql & " where CRYNUM='" & .CRYNUM & "' and INGOTPOS=" & .INGOTPOS
    End With
    '' WriteDBLog sql
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_Tenyou_Ins = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_Tenyou_Ins = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_Tenyou_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :振替廃棄実績の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:Hurikae　　　,I  ,typ_TBCMW006   　,振替廃棄実績
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_Furikae_Ins(Hurikae As typ_TBCMW006) As FUNCTION_RETURN

    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_Furikae_Ins"

    '' 振替廃棄実績の挿入
    With Hurikae
        sql = "insert into TBCMW006 ("
        sql = sql & "CRYNUM, "                  ' 結晶番号
        sql = sql & "INGOTPOS, "                ' インゴット位置
        sql = sql & "TRANCNT, "                 ' 処理回数
        sql = sql & "CRYLEN, "                  ' 長さ
        sql = sql & "KRPROCCD, "                ' 管理工程コード
        sql = sql & "PROCCODE, "                ' 工程コード
        sql = sql & "TRANCLS, "                 ' 処理区分
        sql = sql & "DUNWNUM, "                 ' 転用先品番
        sql = sql & "DUNWREV, "                 ' 転用先品番 製品番号改訂番号
        sql = sql & "DUNWFACT, "                ' 転用先品番 工場
        sql = sql & "DUNWOPCD, "                ' 転用先品番 操業条件
        sql = sql & "DUOGNUM, "                 ' 転用元品番
        sql = sql & "DUOGREV, "                 ' 転用元品番 製品番号改訂番号
        sql = sql & "DUOGFACT, "                ' 転用元品番 工場
        sql = sql & "DUOGOPCD, "                ' 転用元品番 操業条件
        sql = sql & "TSTAFFID, "                ' 登録社員ID
        sql = sql & "REGDATE, "                 ' 登録日付
        sql = sql & "KSTAFFID, "                ' 更新社員ID
        sql = sql & "UPDDATE, "                 ' 更新日付
        ' 2007/09/03 SPK Tsutsumi Add Start
        sql = sql & "SENDFLAG, "                ' 送信フラグ
'        sql = sql & "SENDFLAG) "                ' 送信フラグ
        sql = sql & "PLANTCAT) "                ' 向先
        ' 2007/09/03 SPK Tsutsumi Add End
        sql = sql & " select '"
        sql = sql & .CRYNUM & "', "             ' 結晶番号
        sql = sql & .INGOTPOS & ", "            ' インゴット位置
        sql = sql & "nvl(max(TRANCNT),0)+1, "   ' 処理回数
        sql = sql & .CRYLEN & ", '"             ' 長さ
        sql = sql & .KRPROCCD & "', '"          ' 管理工程コード
        sql = sql & .PROCCODE & "', '"          ' 工程コード
        sql = sql & .TRANCLS & "', '"           ' 処理区分
        sql = sql & .DUNWNUM & "', "            ' 転用先品番
        sql = sql & .DUNWREV & ", '"            ' 転用先品番 製品番号改訂番号
        sql = sql & .DUNWFACT & "', '"          ' 転用先品番 工場
        sql = sql & .DUNWOPCD & "', '"          ' 転用先品番 操業条件
        sql = sql & .DUOGNUM & "', "            ' 転用元品番
        sql = sql & .DUOGREV & ", '"            ' 転用元品番 製品番号改訂番号
        sql = sql & .DUOGFACT & "', '"          ' 転用元品番 工場
        sql = sql & .DUOGOPCD & "', '"          ' 転用元品番 操業条件
        sql = sql & .TSTAFFID & "', "           ' 登録社員ID
        sql = sql & "sysdate, '"                ' 登録日付
        sql = sql & .KSTAFFID & "', "           ' 更新社員ID
        sql = sql & "sysdate, "                 ' 更新日付
        ' 2007/09/03 SPK Tsutsumi Add Start
        sql = sql & "'0',"                      ' 送信フラグ
        sql = sql & "'" & .MUKESAKI & "'"    ' 向先
'        sql = sql & "'0' "                      ' 送信フラグ
        ' 2007/09/03 SPK Tsutsumi Add End
        sql = sql & " from TBCMW006 "
        sql = sql & " where CRYNUM='" & .CRYNUM & "' and INGOTPOS=" & .INGOTPOS
    End With
    '' WriteDBLog sql
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_Furikae_Ins = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_Furikae_Ins = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_Furikae_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'クリスタルカタログ受入実績への挿入
Public Function DBDRV_Catalog_Ins(CryCatalog As typ_TBCMG007) As FUNCTION_RETURN

    Dim sql As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_Catalog_Ins"

    DBDRV_Catalog_Ins = FUNCTION_RETURN_SUCCESS

    ' クリスタルカタログ受入実績への挿入
    sql = "insert into TBCMG007 ( "
    sql = sql & "CRYNUM, "            ' 結晶番号（ブロックID）
    sql = sql & "TRANCNT, "           ' 処理回数
    sql = sql & "KRPROCCD, "          ' 管理工程コード
    sql = sql & "PROCCODE, "          ' 工程コード
    sql = sql & "BDCODE, "            ' 不良理由コード
    sql = sql & "PALTNUM, "           ' パレット番号
    sql = sql & "TSTAFFID, "          ' 登録社員ID
    sql = sql & "REGDATE, "           ' 登録日付
    sql = sql & "KSTAFFID, "          ' 更新社員ID
    sql = sql & "UPDDATE, "           ' 更新日付
    sql = sql & "SENDFLAG, "          ' 送信フラグ
    sql = sql & "SENDDATE) "          ' 送信日付

    With CryCatalog
        sql = sql & "Select "
        sql = sql & " '" & .CRYNUM & "', "          ' 結晶番号
        sql = sql & "nvl(max(TRANCNT),0)+1, "       ' 処理回数
        sql = sql & " '" & .KRPROCCD & "', "        ' 管理工程コード
        sql = sql & " '" & .PROCCODE & "', "        ' 工程コード
        sql = sql & " '" & .BDCODE & "', "          ' 不良理由コード
        sql = sql & " '" & .PALTNUM & "', "         ' パレット番号
        sql = sql & " '" & .TSTAFFID & "', "        ' 登録社員ID
        sql = sql & "sysdate, "                     ' 登録日付
        sql = sql & " '" & .TSTAFFID & "', "        ' 更新社員ID
        sql = sql & "sysdate, "                     ' 更新日付
        sql = sql & "'0', "                         ' 送信フラグ
        sql = sql & "sysdate "                      ' 送信日付
        sql = sql & "From TBCMG007 " & _
              "Where (CRYNUM='" & .CRYNUM & "')"
    End With

    '' WriteDBLog sql
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_Catalog_Ins = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_Catalog_Ins = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶番号からAGR or MGRを取得
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型               ,説明
'      　　:CRYNUM　　  ,I  ,String         　,結晶番号
'      　　:ans  　　　 ,I  ,String         　,A or M or ""
'      　　:戻り値      ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_get_xGR(CRYNUM As String, Ans As String) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_get_xGR"
    DBDRV_get_xGR = FUNCTION_RETURN_FAILURE

    sql = "select PRCMCN from TBCMI001 "
    sql = sql & "where CRYNUM = '" & left(CRYNUM, 9) & "000" & "' and "
    sql = sql & "TRANCNT = any(select max(TRANCNT) from TBCMI001 where CRYNUM = '" & left(CRYNUM, 9) & "000" & "')"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        DBDRV_get_xGR = FUNCTION_RETURN_SUCCESS
        Ans = ""
        rs.Close
        GoTo proc_exit
    End If
    Ans = rs("PRCMCN")
    rs.Close
    
    
    DBDRV_get_xGR = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_get_xGR = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :指定のブロックが存在するかどうかのチェックする
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型               ,説明
'      　　:BLOCKID　　 ,I  ,String         　,結晶番号
'      　　:ans  　　　 ,I  ,Boolean        　,有り(True)無し(False)
'      　　:戻り値      ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_BlockIDCheck(BLOCKID As String, Ans As Boolean) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_BlockIDCheck"
    DBDRV_BlockIDCheck = FUNCTION_RETURN_FAILURE

    sql = "select BLOCKID from TBCME040 "
    sql = sql & "where CRYNUM = '" & left(BLOCKID, 9) & "000" & "' and "
    sql = sql & "BLOCKID = '" & BLOCKID & "'"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Ans = False
    Else
        Ans = True
    End If
    rs.Close
    DBDRV_BlockIDCheck = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_BlockIDCheck = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
        
'概要      :購入単結晶のシード傾きを求める
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型               ,説明
'      　　:BLOCKID　　 ,I  ,String         　,結晶番号
'      　　:ans  　　　 ,I  ,Boolean        　,有り(True)無し(False)
'      　　:戻り値      ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_getSEEDDEG(BLOCKID As String, Ans As Integer) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_getSEEDDEG"
    DBDRV_getSEEDDEG = FUNCTION_RETURN_FAILURE
        
    'ブロック新規情報、シード傾きの求め方を変更
    sql = "select SEEDDEG from TBCMG002 "
    sql = sql & "where TRANCNT=ANY(select MAX(TRANCNT) from TBCMG002 Where CRYNUM='" & BLOCKID & "') and "
    sql = sql & "CRYNUM='" & BLOCKID & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then
        rs.Close
        GoTo proc_exit
    End If
    If rs("SEEDDEG") = 4 Then
        Ans = 4
    Else
        Ans = 0
    End If
    rs.Close
    
    DBDRV_getSEEDDEG = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_getSEEDDEG = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :ブロック内で最もLT仕様が厳しい品番を求める
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型          ,説明
'          :Crynum        ,I  ,String      ,結晶番号
'          :Ingotpos      ,I  ,Integer     ,ブロックの終了位置
'          :hin           ,O  ,tFullHinban ,ブロック内で最もLT仕様の厳しい品番
'          :LTSPI         ,O  ,String      ,ブロック内で最もLT仕様の厳しい品番のLT測定位置コード
'          :戻り値        ,O  ,FUNCTION_RETURN,読込の成否
'説明      :「最もLT仕様が厳しい品番」がなければ、hin.HINBAN='        ', LTSPI=VbNullString
'履歴      :2002/4/23 野村 作成
Public Function DBDRV_getLtHinbanInBlock(CRYNUM As String, INGOTPOS As Integer, HIN As tFullHinban, LTSPI As String) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset
Dim recCnt As Integer
Dim BlkFrom As Integer
Dim BlkTo As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_getLtHinbanInBlock"
    DBDRV_getLtHinbanInBlock = FUNCTION_RETURN_FAILURE
    
    '初期化
    HIN.hinban = vbNullString
    HIN.mnorevno = 0
    HIN.factory = vbNullString
    HIN.factory = vbNullString
    LTSPI = vbNullString
    
    'ブロックの範囲を求める
    BlkTo = INGOTPOS
    sql = "select INGOTPOS from TBCME040 where CRYNUM='" & CRYNUM & "' and INGOTPOS+LENGTH=" & BlkTo
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then     'ブロック終端でないため、対象品番なし
        rs.Close
        DBDRV_getLtHinbanInBlock = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    BlkFrom = rs("INGOTPOS")
    rs.Close
    Set rs = Nothing
    
    'ブロック内の品番で、最もLT仕様が厳しいものを求める
    sql = "select SIYO.HINBAN, SIYO.MNOREVNO, SIYO.FACTORY, SIYO.OPECOND, SIYO.HSXLTHWS, SIYO.HSXLTSPI "
    sql = sql & "from TBCME041 HIN, TBCME019 SIYO "
    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
    sql = sql & "  and HIN.INGOTPOS<" & BlkTo & " and HIN.INGOTPOS+HIN.LENGTH>" & BlkFrom
    sql = sql & "  and SIYO.HINBAN=HIN.HINBAN and SIYO.MNOREVNO=HIN.REVNUM and SIYO.FACTORY=HIN.FACTORY and SIYO.OPECOND=HIN.OPECOND"
    sql = sql & "  and SIYO.HSXLTHWS in ('H','S') "
    sql = sql & "order by SIYO.HSXLTSPI, HIN.INGOTPOS desc"
    sql = "select * from (" & sql & ") where rownum=1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then     '対象品番なし(LTが保証/参考の品番がないと思われる)
        rs.Close
        DBDRV_getLtHinbanInBlock = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    HIN.hinban = rs("HINBAN")
    HIN.mnorevno = rs("MNOREVNO")
    HIN.factory = rs("FACTORY")
    HIN.opecond = rs("OPECOND")
    LTSPI = rs("HSXLTSPI")
    rs.Close
    Set rs = Nothing
    
    DBDRV_getLtHinbanInBlock = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_getLtHinbanInBlock = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :ｻﾝﾌﾟﾙ位置の上品番と下品番が異なる場合、最もGD仕様が厳しい品番を求める
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO ,型          ,説明
'          :tblCrySmp   ,I  ,typ_XSDCS   ,ｻﾝﾌﾟﾙ
'          :HIN         ,I  ,tFullHinban ,最もGD仕様が厳しい品番
'          :戻り値        ,O  ,FUNCTION_RETURN,読込の成否
'説明      :「最もGD仕様が厳しい品番」がなければ、hin.HINBAN='        '
'履歴      :2005/10/05 Y.SIMIZU
Public Function DBDRV_getGDHinbanInBlock(tblCrySmp As typ_XSDCS, HIN As tFullHinban) As FUNCTION_RETURN
    Dim sql     As String
    Dim rs      As OraDynaset
    Dim sGDLine As Single

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_getGDHinbanInBlock"
    DBDRV_getGDHinbanInBlock = FUNCTION_RETURN_FAILURE
    
    '初期化
    HIN.hinban = vbNullString
    HIN.mnorevno = 0
    HIN.factory = vbNullString
    HIN.factory = vbNullString
    
    '同位置の品番のGDﾗｲﾝ数を取得する
    sql = "SELECT  DISTINCT HINBCS,REVNUMCS,FACTORYCS,OPECS,HSXGDLINE,HWFGDLINE "
    sql = sql & "FROM   XSDCS T1,TBCME036 T2 "
    sql = sql & "WHERE  T1.XTALCS = '" & tblCrySmp.XTALCS & "' "
    sql = sql & "AND    T1.INPOSCS = " & tblCrySmp.INPOSCS & " "
    sql = sql & "AND    T1.LIVKCS <> '1' "
    sql = sql & "AND    T1.HINBCS = T2.HINBAN "
    sql = sql & "AND    T1. REVNUMCS = T2.MNOREVNO "
    sql = sql & "AND    T1. FACTORYCS = T2.FACTORY "
    sql = sql & "AND    T1. OPECS = T2.OPECOND "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount <= 0 Then     '対象品番なし
        rs.Close
        DBDRV_getGDHinbanInBlock = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    '自品番のGDﾗｲﾝ数を取得
    Do Until rs.EOF
        With tblCrySmp
            '自品番の場合
            If rs("HINBCS") = .HINBCS And rs("REVNUMCS") = .REVNUMCS And rs("FACTORYCS") = .FACTORYCS And rs("OPECS") = .OPECS Then
                HIN.hinban = rs("HINBCS")
                HIN.mnorevno = rs("REVNUMCS")
                HIN.factory = rs("FACTORYCS")
                HIN.opecond = rs("OPECS")
                sGDLine = fncNullCheck(rs("HSXGDLINE"))
            End If
        End With
        rs.MoveNext
    Loop
    
    rs.MoveFirst
    
    '自品番のﾗｲﾝ数よりもﾗｲﾝ数が多い品番を取得
    Do Until rs.EOF
        '自品番のﾗｲﾝ数よりもﾗｲﾝ数が多い場合
        If fncNullCheck(rs("HSXGDLINE")) > sGDLine Then
            HIN.hinban = rs("HINBCS")
            HIN.mnorevno = rs("REVNUMCS")
            HIN.factory = rs("FACTORYCS")
            HIN.opecond = rs("OPECS")
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    
    Set rs = Nothing
    
    DBDRV_getGDHinbanInBlock = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_getGDHinbanInBlock = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :エピ測定評価方法指示の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:SokuSizi　　　,I  ,typ_TBCMY020   　,エピ測定評価方法指示
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2006/08/15  作成 SMP)kondoh
Public Function DBDRV_SokuSizi_EP_Ins(SokuSizi() As typ_TBCMY020) As FUNCTION_RETURN

    Dim sql As String
    Dim i As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_SokuSizi_EP_Ins"

    '' 測定評価方法指示の挿入
    For i = 1 To UBound(SokuSizi)
        With SokuSizi(i)
            sql = " insert into TBCMY020 ("
            sql = sql & "SAMPLEID, "                ' サンプルID
            sql = sql & "OSITEM, "                  ' 評価項目
            sql = sql & "TRANCNT, "                 ' 処理回数
            sql = sql & "SAMPLEKB, "                ' サンプル区分
            sql = sql & "MAISU, "                   ' 評価枚数
            sql = sql & "SPEC, "                    ' 規格値
            sql = sql & "NETSU, "                   ' 熱処理条件
            sql = sql & "ET, "                      ' エッチング条件
            sql = sql & "MES, "                     ' 計測方法
            sql = sql & "DKAN, "                    ' ＤＫアニール条件
            sql = sql & "TXID, "                    ' トランザクションID
            sql = sql & "REGDATE, "                 ' 登録日付
            sql = sql & "SENDFLAG, "                ' 送信フラグ
            
            ' 2007/08/31 SPK Tsutsumi Add Start
            sql = sql & "SENDDATE, "                ' 送信日付
            sql = sql & "PLANTCAT) "                ' 向先 2007/08/31 SPK Tsutsumi Add
            'sql = sql & "SENDDATE) "                ' 送信日付
            ' 2007/08/31 SPK Tsutsumi Add End

            sql = sql & " select '"
            sql = sql & .SAMPLEID & "', '"          ' サンプルID
            sql = sql & .OSITEM & "', "             ' 評価項目
            sql = sql & "nvl(max(TRANCNT),0)+1, '"  ' 処理回数
            sql = sql & .SAMPLEKB & "', '"           ' サンプル区分
            sql = sql & .MAISU & "', '"             ' 評価枚数
            sql = sql & .Spec & "', '"              ' 規格値
            sql = sql & .NETSU & "', '"             ' 熱処理条件
            sql = sql & .ET & "', '"                 ' エッチング条件
            sql = sql & .MES & "', '"               ' 計測方法
            sql = sql & .DKAN & "', "               ' ＤＫアニール条件
            sql = sql & "'TX871I', "                ' トランザクションID
            sql = sql & "sysdate, "                 ' 登録日付
            sql = sql & "'" & .SENDFLAG & "', "     ' 送信フラグ
            
            ' 2007/08/31 SPK Tsutsumi Add Start
'            sql = sql & "sysdate "                 ' 送信日付
            sql = sql & "sysdate, "                 ' 送信日付
            sql = sql & "'" & .MUKESAKI & "' "   ' 向先
            ' 2007/08/31 SPK Tsutsumi Add End
            
            sql = sql & " from TBCMY020"
            sql = sql & " where SAMPLEID='" & .SAMPLEID & "'"
            sql = sql & " and OSITEM='" & .OSITEM & "'"
            sql = sql & " and SPEC='" & .Spec & "'"
        End With
        '' WriteDBLog sql
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_SokuSizi_EP_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    DBDRV_SokuSizi_EP_Ins = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_SokuSizi_EP_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function
