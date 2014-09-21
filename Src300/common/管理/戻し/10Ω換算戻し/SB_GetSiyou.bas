Attribute VB_Name = "SB_GetSiyou"
Option Explicit

'------------------------------------------------
' TBCME018データ取得(判定用)
'------------------------------------------------

'概要      :テーブル「TBCME018」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2003/09/10 新規作成　システムブレイン

Public Function funGet_TBCME018(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME018"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HSXTYPE, HSXD1CEN, HSXCDIR, HSXRMIN, HSXRMAX, HSXRAMIN, HSXRAMAX, "
    sql = sql & "HSXRMCAL, HSXRMBNP, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS "
    sql = sql & "from TBCME018 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME018 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''抽出結果を格納する
     With tGetRec
        .hin.hinban = rs("HINBAN")          ' 品番
        .hin.mnorevno = rs("MNOREVNO")      ' 製品番号改訂番号
        .hin.factory = rs("FACTORY")        ' 工場
        .hin.opecond = rs("OPECOND")        ' 操業条件
        
        .HSXTYPE = rs("HSXTYPE")                    ' 品ＳＸタイプ
        .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))    ' 品ＳＸ直径１中心          2003/12/12 SystemBrain Null対応
        .HSXCDIR = rs("HSXCDIR")                    ' 品ＳＸ結晶面方位
        .HSXRMIN = fncNullCheck(rs("HSXRMIN"))      ' 品ＳＸ比抵抗下限          2003/12/12 SystemBrain Null対応
        .HSXRMAX = fncNullCheck(rs("HSXRMAX"))      ' 品ＳＸ比抵抗上限          2003/12/12 SystemBrain Null対応
        .HSXRAMIN = fncNullCheck(rs("HSXRAMIN"))    ' 品ＳＸ比抵抗平均下限      2003/12/12 SystemBrain Null対応
        .HSXRAMAX = fncNullCheck(rs("HSXRAMAX"))    ' 品ＳＸ比抵抗平均上限      2003/12/12 SystemBrain Null対応
        .HSXRMCAL = rs("HSXRMCAL")                  ' 品ＳＸ比抵抗面内計算
        .HSXRMBNP = fncNullCheck(rs("HSXRMBNP"))    ' 品ＳＸ比抵抗面内分布      2003/12/12 SystemBrain Null対応
        .HSXRSPOH = rs("HSXRSPOH")                  ' 品ＳＸ比抵抗測定位置＿方
        .HSXRSPOT = rs("HSXRSPOT")                  ' 品ＳＸ比抵抗測定位置＿点
        .HSXRSPOI = rs("HSXRSPOI")                  ' 品ＳＸ比抵抗測定位置＿位
        .HSXRHWYT = rs("HSXRHWYT")                  ' 品ＳＸ比抵抗保証方法＿対
        .HSXRHWYS = rs("HSXRHWYS")                  ' 品ＳＸ比抵抗保証方法＿処
    End With
    Set rs = Nothing

    funGet_TBCME018 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME019データ取得(判定用)
'------------------------------------------------

'概要      :テーブル「TBCME019」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2003/09/10 新規作成　システムブレイン

Public Function funGet_TBCME019(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME019"

    'HSXCNKHI追加 09/01/08 ooba
    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HSXONMIN, HSXONMAX, HSXONAMN, HSXONAMX, HSXONMCL, HSXONMBP, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, "
    sql = sql & "HSXCNMIN, HSXCNMAX, HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKHI, "
    sql = sql & "HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT, HSXLTHWS "
    sql = sql & "from TBCME019 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME019 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''抽出結果を格納する
     With tGetRec
        .hin.hinban = rs("HINBAN")         ' 品番
        .hin.mnorevno = rs("MNOREVNO")     ' 製品番号改訂番号
        .hin.factory = rs("FACTORY")       ' 工場
        .hin.opecond = rs("OPECOND")       ' 操業条件
        
        .HSXONMIN = fncNullCheck(rs("HSXONMIN"))        ' 品ＳＸ酸素濃度下限            2003/12/12 SystemBrain Null対応
        .HSXONMAX = fncNullCheck(rs("HSXONMAX"))        ' 品ＳＸ酸素濃度上限            2003/12/12 SystemBrain Null対応
        .HSXONAMN = fncNullCheck(rs("HSXONAMN"))        ' 品ＳＸ酸素濃度平均下限        2003/12/12 SystemBrain Null対応
        .HSXONAMX = fncNullCheck(rs("HSXONAMX"))        ' 品ＳＸ酸素濃度平均上限        2003/12/12 SystemBrain Null対応
        .HSXONMCL = rs("HSXONMCL")                      ' 品ＳＸ酸素濃度面内計算
        .HSXONMBP = fncNullCheck(rs("HSXONMBP"))        ' 品ＳＸ酸素濃度面内分布        2003/12/12 SystemBrain Null対応
        .HSXONSPH = rs("HSXONSPH")                      ' 品ＳＸ酸素濃度測定位置＿方
        .HSXONSPT = rs("HSXONSPT")                      ' 品ＳＸ酸素濃度測定位置＿点
        .HSXONSPI = rs("HSXONSPI")                      ' 品ＳＸ酸素濃度測定位置＿位
        .HSXONHWT = rs("HSXONHWT")                      ' 品ＳＸ酸素濃度保証方法＿対
        .HSXONHWS = rs("HSXONHWS")                      ' 品ＳＸ酸素濃度保証方法＿処
        
        .HSXCNMIN = fncNullCheck(rs("HSXCNMIN"))        ' 品ＳＸ炭素濃度下限            2003/12/12 SystemBrain Null対応
        .HSXCNMAX = fncNullCheck(rs("HSXCNMAX"))        ' 品ＳＸ炭素濃度上限            2003/12/12 SystemBrain Null対応
        .HSXCNSPH = rs("HSXCNSPH")                      ' 品ＳＸ炭素濃度測定位置＿方
        .HSXCNSPT = rs("HSXCNSPT")                      ' 品ＳＸ炭素濃度測定位置＿点
        .HSXCNSPI = rs("HSXCNSPI")                      ' 品ＳＸ炭素濃度測定位置＿位
        .HSXCNHWT = rs("HSXCNHWT")                      ' 品ＳＸ炭素濃度保証方法＿対
        .HSXCNHWS = rs("HSXCNHWS")                      ' 品ＳＸ炭素濃度保証方法＿処
        .HSXCNKHI = rs("HSXCNKHI")                      ' 品ＳＸ炭素濃度検査頻度＿位 09/01/08 ooba
        
        .HSXLTMIN = fncNullCheck(rs("HSXLTMIN"))        ' 品ＳＸＬタイム下限            2003/12/12 SystemBrain Null対応
        .HSXLTMAX = fncNullCheck(rs("HSXLTMAX"))        ' 品ＳＸＬタイム上限            2003/12/12 SystemBrain Null対応
        .HSXLTSPH = rs("HSXLTSPH")                      ' 品ＳＸＬタイム測定位置＿方
        .HSXLTSPT = rs("HSXLTSPT")                      ' 品ＳＸＬタイム測定位置＿点
        .HSXLTSPI = rs("HSXLTSPI")                      ' 品ＳＸＬタイム測定位置＿位
        .HSXLTHWT = rs("HSXLTHWT")                      ' 品ＳＸＬタイム保証方法＿対
        .HSXLTHWS = rs("HSXLTHWS")                      ' 品ＳＸＬタイム保証方法＿処
    End With
    Set rs = Nothing

    funGet_TBCME019 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME020データ取得(判定用)
'------------------------------------------------

'概要      :テーブル「TBCME020」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2003/09/10 新規作成　システムブレイン

Public Function funGet_TBCME020(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME020"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1NS, "
    sql = sql & "HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS, HSXBM2NS, "
    sql = sql & "HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST, HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3NS, "
    sql = sql & "HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1NS, "
    sql = sql & "HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2NS, "
    sql = sql & "HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR, HSXOF3HT, HSXOF3HS, HSXOF3NS, "
    sql = sql & "HSXOF4AX, HSXOF4MX, HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4NS, "
    sql = sql & "HSXDENMX, HSXDENMN, HSXDENHT, HSXDENHS, HSXDENKU, "
    sql = sql & "HSXLDLMX, HSXLDLMN, HSXLDLHT, HSXLDLHS, HSXLDLKU, "
    sql = sql & "HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXDVDKU, "
    sql = sql & "HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK, HSXBMD1MBP, HSXBMD2MBP, HSXBMD3MBP "
    sql = sql & ", HSXGDPTK "   '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    'Add Start 2011/02/01 SMPK Miyata
    sql = sql & ", HSXCPK, HSXCSZ, HSXCHT, HSXCHS "
    sql = sql & ", HSXCJPK, HSXCJNS, HSXCJHT, HSXCJHS "
    sql = sql & ", HSXCJLTPK, HSXCJLTNS, HSXCJLTHT, HSXCJLTHS "
    sql = sql & ", HSXCJ2PK, HSXCJ2NS, HSXCJ2HT, HSXCJ2HS "
    'Add End   2011/02/01 SMPK Miyata
    sql = sql & "from TBCME020 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME020 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
     
    ''抽出結果を格納する
    With tGetRec
        .hin.hinban = rs("HINBAN")       ' 品番
        .hin.mnorevno = rs("MNOREVNO")   ' 製品番号改訂番号
        .hin.factory = rs("FACTORY")     ' 工場
        .hin.opecond = rs("OPECOND")     ' 操業条件
        
        .HSXBM1AN = fncNullCheck(rs("HSXBM1AN"))        ' 品ＳＸＢＭＤ1平均下限     2003/12/12 SystemBrain Null対応
        .HSXBM1AX = fncNullCheck(rs("HSXBM1AX"))        ' 品ＳＸＢＭＤ1平均上限     2003/12/12 SystemBrain Null対応
        .HSXBM1SH = rs("HSXBM1SH")                      ' 品ＳＸＢＭＤ1測定位置＿方
        .HSXBM1ST = rs("HSXBM1ST")                      ' 品ＳＸＢＭＤ1測定位置＿点
        .HSXBM1SR = rs("HSXBM1SR")                      ' 品ＳＸＢＭＤ1測定位置＿領
        .HSXBM1HT = rs("HSXBM1HT")                      ' 品ＳＸＢＭＤ1保証方法＿対
        .HSXBM1HS = rs("HSXBM1HS")                      ' 品ＳＸＢＭＤ1保証方法＿処
        .HSXBM1NS = rs("HSXBM1NS")                      ' 品ＳＸＢＭＤ1熱処理法
        .HSXBM2AN = fncNullCheck(rs("HSXBM2AN"))        ' 品ＳＸＢＭＤ2平均下限     2003/12/12 SystemBrain Null対応
        .HSXBM2AX = fncNullCheck(rs("HSXBM2AX"))        ' 品ＳＸＢＭＤ2平均上限     2003/12/12 SystemBrain Null対応
        .HSXBM2SH = rs("HSXBM2SH")                      ' 品ＳＸＢＭＤ2測定位置＿方
        .HSXBM2ST = rs("HSXBM2ST")                      ' 品ＳＸＢＭＤ2測定位置＿点
        .HSXBM2SR = rs("HSXBM2SR")                      ' 品ＳＸＢＭＤ2測定位置＿領
        .HSXBM2HT = rs("HSXBM2HT")                      ' 品ＳＸＢＭＤ2保証方法＿対
        .HSXBM2HS = rs("HSXBM2HS")                      ' 品ＳＸＢＭＤ2保証方法＿処
        .HSXBM2NS = rs("HSXBM2NS")                      ' 品ＳＸＢＭＤ2熱処理法
        .HSXBM3AN = fncNullCheck(rs("HSXBM3AN"))        ' 品ＳＸＢＭＤ3平均下限     2003/12/12 SystemBrain Null対応
        .HSXBM3AX = fncNullCheck(rs("HSXBM3AX"))        ' 品ＳＸＢＭＤ3平均上限     2003/12/12 SystemBrain Null対応
        .HSXBM3SH = rs("HSXBM3SH")                      ' 品ＳＸＢＭＤ3測定位置＿方
        .HSXBM3ST = rs("HSXBM3ST")                      ' 品ＳＸＢＭＤ3測定位置＿点
        .HSXBM3SR = rs("HSXBM3SR")                      ' 品ＳＸＢＭＤ3測定位置＿領
        .HSXBM3HT = rs("HSXBM3HT")                      ' 品ＳＸＢＭＤ3保証方法＿対
        .HSXBM3HS = rs("HSXBM3HS")                      ' 品ＳＸＢＭＤ3保証方法＿処
        .HSXBM3NS = rs("HSXBM3NS")                      ' 品ＳＸＢＭＤ3熱処理法
        
        .HSXOF1AX = fncNullCheck(rs("HSXOF1AX"))        ' 品ＳＸＯＳＦ1平均上限     2003/12/12 SystemBrain Null対応
        .HSXOF1MX = fncNullCheck(rs("HSXOF1MX"))        ' 品ＳＸＯＳＦ1上限         2003/12/12 SystemBrain Null対応
        .HSXOF1SH = rs("HSXOF1SH")                      ' 品ＳＸＯＳＦ1測定位置＿方
        .HSXOF1ST = rs("HSXOF1ST")                      ' 品ＳＸＯＳＦ1測定位置＿点
        .HSXOF1SR = rs("HSXOF1SR")                      ' 品ＳＸＯＳＦ1測定位置＿領
        .HSXOF1HT = rs("HSXOF1HT")                      ' 品ＳＸＯＳＦ1保証方法＿対
        .HSXOF1HS = rs("HSXOF1HS")                      ' 品ＳＸＯＳＦ1保証方法＿処
        .HSXOF1NS = rs("HSXOF1NS")                      ' 品ＳＸＯＳＦ1熱処理法
        .HSXOF2AX = fncNullCheck(rs("HSXOF2AX"))        ' 品ＳＸＯＳＦ2平均上限     2003/12/12 SystemBrain Null対応
        .HSXOF2MX = fncNullCheck(rs("HSXOF2MX"))        ' 品ＳＸＯＳＦ2上限         2003/12/12 SystemBrain Null対応
        .HSXOF2SH = rs("HSXOF2SH")                      ' 品ＳＸＯＳＦ2測定位置＿方
        .HSXOF2ST = rs("HSXOF2ST")                      ' 品ＳＸＯＳＦ2測定位置＿点
        .HSXOF2SR = rs("HSXOF2SR")                      ' 品ＳＸＯＳＦ2測定位置＿領
        .HSXOF2HT = rs("HSXOF2HT")                      ' 品ＳＸＯＳＦ2保証方法＿対
        .HSXOF2HS = rs("HSXOF2HS")                      ' 品ＳＸＯＳＦ2保証方法＿処
        .HSXOF2NS = rs("HSXOF2NS")                      ' 品ＳＸＯＳＦ2熱処理法
        .HSXOF3AX = fncNullCheck(rs("HSXOF3AX"))        ' 品ＳＸＯＳＦ3平均上限     2003/12/12 SystemBrain Null対応
        .HSXOF3MX = fncNullCheck(rs("HSXOF3MX"))        ' 品ＳＸＯＳＦ3上限         2003/12/12 SystemBrain Null対応
        .HSXOF3SH = rs("HSXOF3SH")                      ' 品ＳＸＯＳＦ3測定位置＿方
        .HSXOF3ST = rs("HSXOF3ST")                      ' 品ＳＸＯＳＦ3測定位置＿点
        .HSXOF3SR = rs("HSXOF3SR")                      ' 品ＳＸＯＳＦ3測定位置＿領
        .HSXOF3HT = rs("HSXOF3HT")                      ' 品ＳＸＯＳＦ3保証方法＿対
        .HSXOF3HS = rs("HSXOF3HS")                      ' 品ＳＸＯＳＦ3保証方法＿処
        .HSXOF3NS = rs("HSXOF3NS")                      ' 品ＳＸＯＳＦ3熱処理法
        .HSXOF4AX = fncNullCheck(rs("HSXOF4AX"))        ' 品ＳＸＯＳＦ4平均上限     2003/12/12 SystemBrain Null対応
        .HSXOF4MX = fncNullCheck(rs("HSXOF4MX"))        ' 品ＳＸＯＳＦ4上限         2003/12/12 SystemBrain Null対応
        .HSXOF4SH = rs("HSXOF4SH")                      ' 品ＳＸＯＳＦ4測定位置＿方
        .HSXOF4ST = rs("HSXOF4ST")                      ' 品ＳＸＯＳＦ4測定位置＿点
        .HSXOF4SR = rs("HSXOF4SR")                      ' 品ＳＸＯＳＦ4測定位置＿領
        .HSXOF4HT = rs("HSXOF4HT")                      ' 品ＳＸＯＳＦ4保証方法＿対
        .HSXOF4HS = rs("HSXOF4HS")                      ' 品ＳＸＯＳＦ4保証方法＿処
        .HSXOF4NS = rs("HSXOF4NS")                      ' 品ＳＸＯＳＦ4熱処理法
        
        .HSXDENKU = rs("HSXDENKU")                      ' 品ＳＸＤｅｎ検査有無
        .HSXDENMX = fncNullCheck(rs("HSXDENMX"))        ' 品ＳＸＤｅｎ上限          2003/12/12 SystemBrain Null対応
        .HSXDENMN = fncNullCheck(rs("HSXDENMN"))        ' 品ＳＸＤｅｎ下限          2003/12/12 SystemBrain Null対応
        .HSXDENHT = rs("HSXDENHT")                      ' 品ＳＸＤｅｎ保証方法＿対
        .HSXDENHS = rs("HSXDENHS")                      ' 品ＳＸＤｅｎ保証方法＿処
        .HSXDVDKU = rs("HSXDVDKU")                      ' 品ＳＸＤＶＤ２検査有無
        .HSXDVDMX = fncNullCheck(rs("HSXDVDMXN"))       ' 品ＳＸＤＶＤ２上限        2003/12/12 SystemBrain Null対応
        .HSXDVDMN = fncNullCheck(rs("HSXDVDMNN"))       ' 品ＳＸＤＶＤ２下限        2003/12/12 SystemBrain Null対応
        .HSXDVDHT = rs("HSXDVDHT")                      ' 品ＳＸＤＶＤ２保証方法＿対
        .HSXDVDHS = rs("HSXDVDHS")                      ' 品ＳＸＤＶＤ２保証方法＿処
        .HSXLDLKU = rs("HSXLDLKU")                      ' 品ＳＸＬ／ＤＬ検査有無
        .HSXLDLMX = fncNullCheck(rs("HSXLDLMX"))        ' 品ＳＸＬ／ＤＬ上限        2003/12/12 SystemBrain Null対応
        .HSXLDLMN = fncNullCheck(rs("HSXLDLMN"))        ' 品ＳＸＬ／ＤＬ下限        2003/12/12 SystemBrain Null対応
        .HSXLDLHT = rs("HSXLDLHT")                      ' 品ＳＸＬ／ＤＬ保証方法＿対
        .HSXLDLHS = rs("HSXLDLHS")                      ' 品ＳＸＬ／ＤＬ保証方法＿処
        
        If Not IsNull(rs("HSXOSF1PTK")) Then .HSXOSF1PTK = rs("HSXOSF1PTK")     ' 品ＳＸＯＳＦ１パタン区分
        If Not IsNull(rs("HSXOSF2PTK")) Then .HSXOSF2PTK = rs("HSXOSF2PTK")     ' 品ＳＸＯＳＦ２パタン区分
        If Not IsNull(rs("HSXOSF3PTK")) Then .HSXOSF3PTK = rs("HSXOSF3PTK")     ' 品ＳＸＯＳＦ３パタン区分
        If Not IsNull(rs("HSXOSF4PTK")) Then .HSXOSF4PTK = rs("HSXOSF4PTK")     ' 品ＳＸＯＳＦ４パタン区分
'        If Not IsNull(rs("HSXBMD1MBP")) Then .HSXBMD1MBP = rs("HSXBMD1MBP")     ' 品ＳＸＢＭＤ１面内分布
'        If Not IsNull(rs("HSXBMD2MBP")) Then .HSXBMD2MBP = rs("HSXBMD2MBP")     ' 品ＳＸＢＭＤ２面内分布
'        If Not IsNull(rs("HSXBMD3MBP")) Then .HSXBMD3MBP = rs("HSXBMD3MBP")     ' 品ＳＸＢＭＤ３面内分布
        .HSXBMD1MBP = fncNullCheck(rs("HSXBMD1MBP"))                            ' 品ＳＸＢＭＤ１面内分布    2003/12/12 SystemBrain Null対応
        .HSXBMD2MBP = fncNullCheck(rs("HSXBMD2MBP"))                            ' 品ＳＸＢＭＤ２面内分布    2003/12/12 SystemBrain Null対応
        .HSXBMD3MBP = fncNullCheck(rs("HSXBMD3MBP"))                            ' 品ＳＸＢＭＤ３面内分布    2003/12/12 SystemBrain Null対応
        
        If Not IsNull(rs("HSXGDPTK")) Then .HSXGDPTK = rs("HSXGDPTK") Else .HSXGDPTK = " "  ' 品ＳＸＧＤパタン区分  '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    
        'Add Start 2011/02/01 SMPK Miyata
        If Not IsNull(rs("HSXCPK")) Then .HSXCPK = rs("HSXCPK")         ' 品ＳＸＣパターン区分
        If Not IsNull(rs("HSXCSZ")) Then .HSXCSZ = rs("HSXCSZ")         ' 品ＳＸＣ測定条件
        If Not IsNull(rs("HSXCHT")) Then .HSXCHT = rs("HSXCHT")         ' 品ＳＸＣ保証方法＿対
        If Not IsNull(rs("HSXCHS")) Then .HSXCHS = rs("HSXCHS")         ' 品ＳＸＣ保証方法＿処
        If Not IsNull(rs("HSXCJPK")) Then .HSXCJPK = rs("HSXCJPK")      ' 品ＳＸＣＪパターン区分
        If Not IsNull(rs("HSXCJNS")) Then .HSXCJNS = rs("HSXCJNS")      ' 品ＳＸＣＪ熱処理法
        If Not IsNull(rs("HSXCJHT")) Then .HSXCJHT = rs("HSXCJHT")      ' 品ＳＸＣＪ保証方法＿対
        If Not IsNull(rs("HSXCJHS")) Then .HSXCJHS = rs("HSXCJHS")      ' 品ＳＸＣＪ保証方法＿処
        If Not IsNull(rs("HSXCJLTPK")) Then .HSXCJLTPK = rs("HSXCJLTPK")  ' 品ＳＸＣＪＬＴパターン区分
        If Not IsNull(rs("HSXCJLTNS")) Then .HSXCJLTNS = rs("HSXCJLTNS")  ' 品ＳＸＣＪＬＴ熱処理法
        If Not IsNull(rs("HSXCJLTHT")) Then .HSXCJLTHT = rs("HSXCJLTHT")  ' 品ＳＸＣＪＬＴ保証方法＿対
        If Not IsNull(rs("HSXCJLTHS")) Then .HSXCJLTHS = rs("HSXCJLTHS")  ' 品ＳＸＣＪＬＴ保証方法＿処
        If Not IsNull(rs("HSXCJ2PK")) Then .HSXCJ2PK = rs("HSXCJ2PK")   ' 品ＳＸＣＪ２パターン区分
        If Not IsNull(rs("HSXCJ2NS")) Then .HSXCJ2NS = rs("HSXCJ2NS")   ' 品ＳＸＣＪ２熱処理法
        If Not IsNull(rs("HSXCJ2HT")) Then .HSXCJ2HT = rs("HSXCJ2HT")   ' 品ＳＸＣＪ２保証方法＿対
        If Not IsNull(rs("HSXCJ2HS")) Then .HSXCJ2HS = rs("HSXCJ2HS")   ' 品ＳＸＣＪ２保証方法＿処
        'Add End   2011/02/01 SMPK Miyata
    
    End With
    Set rs = Nothing

    funGet_TBCME020 = FUNCTION_RETURN_SUCCESS

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
' TBCME036データ取得(判定用)
'------------------------------------------------

'概要      :テーブル「TBCME036」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2003/09/10 新規作成　システムブレイン

Public Function funGet_TBCME036(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME036"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
'*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ追加
'    sql = sql & "EPDUP, TOPREG, TAILREG, BTMSPRT, BLOCKHFLAG "
'    sql = sql & "EPDUP, TOPREG, TAILREG, BTMSPRT, BLOCKHFLAG,HSXGDLINE "
'*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ追加
    sql = sql & "EPDUP, TOPREG, TAILREG, BTMSPRT, BLOCKHFLAG, HSXGDLINE, COSF3FLAG "
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sql = sql & ",NVL(HSXDKTMP,' ') HSXDKTMP "
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    sql = sql & ",HSXLDLRMN, HSXLDLRMX, HWFLDLRMN, HWFLDLRMX, HSXOF1ARPTK, HSXOFARMIN, HSXOFARMAX, HSXOFARMHMX "
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    'Add Start 2011/02/01 SMPK Miyata
    sql = sql & ",HSXCJLTBND "
    'Add End   2011/02/01 SMPK Miyata

    sql = sql & "from TBCME036 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME036 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''抽出結果を格納する
     With tGetRec
        .hin.hinban = rs("HINBAN")          ' 品番
        .hin.mnorevno = rs("MNOREVNO")      ' 製品番号改訂番号
        .hin.factory = rs("FACTORY")        ' 工場
        .hin.opecond = rs("OPECOND")        ' 操業条件
        
'        If Not IsNull(rs("EPDUP")) Then .EPDUP = rs("EPDUP")                    ' EPD上限
'        If Not IsNull(rs("TOPREG")) Then .TOPREG = rs("TOPREG")                 ' TOP規制
'        If Not IsNull(rs("TAILREG")) Then .TAILREG = rs("TAILREG")              ' TAIL規制
'        If Not IsNull(rs("BTMSPRT")) Then .BTMSPRT = rs("BTMSPRT")              ' ボトム析出規制
        .EPDUP = fncNullCheck(rs("EPDUP"))                                      ' EPD上限                   2003/12/12 SystemBrain Null対応
        .TOPREG = fncNullCheck(rs("TOPREG"))                                    ' TOP規制                   2003/12/12 SystemBrain Null対応
        .TAILREG = fncNullCheck(rs("TAILREG"))                                  ' TAIL規制                  2003/12/12 SystemBrain Null対応
        .BTMSPRT = fncNullCheck(rs("BTMSPRT"))                                  ' ボトム析出規制            2003/12/12 SystemBrain Null対応
        If Not IsNull(rs("BLOCKHFLAG")) Then .BLOCKHFLAG = rs("BLOCKHFLAG")     ' ブロック単位保証品番フラグ
    '*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ追加
        .HSXGDLINE = fncNullCheck(rs("HSXGDLINE"))
    '*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ追加
    
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
        If IsNull(rs("COSF3FLAG")) = False Then .COSF3FLAG = rs("COSF3FLAG") Else .COSF3FLAG = " "            'C-OSF3ﾌﾗｸﾞ
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---

'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        .HSXLDLRMN = fncNullCheck(rs("HSXLDLRMN"))      ' 品SXL/DL連続0下限
        .HSXLDLRMX = fncNullCheck(rs("HSXLDLRMX"))      ' 品SXL/DL連続0上限
        .HWFLDLRMN = fncNullCheck(rs("HWFLDLRMN"))      ' 品WFL/DL連続0下限
        .HWFLDLRMX = fncNullCheck(rs("HWFLDLRMX"))      ' 品WFL/DL連続0上限
        If IsNull(rs("HSXOF1ARPTK")) = False Then .HSXOF1ARPTK = rs("HSXOF1ARPTK") Else .HSXOF1ARPTK = " "  ' 品SXOSF1(ArAN)パタン区分
        .HSXOFARMIN = fncNullCheck(rs("HSXOFARMIN"))    ' 品SXOSF(ArAN)下限
        .HSXOFARMAX = fncNullCheck(rs("HSXOFARMAX"))    ' 品SXOSF(ArAN)上限
        .HSXOFARMHMX = fncNullCheck(rs("HSXOFARMHMX"))  ' 品SXOSF(ArAN)面内比上限
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
        'Add Start 2011/02/01 SMPK Miyata
        .HSXCJLTBND = fncNullCheck(rs("HSXCJLTBND"))    ' 品SXL/CJLTバンド幅 Number(3,0)
        'Add End   2011/02/01 SMPK Miyata

    End With
    Set rs = Nothing

    funGet_TBCME036 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME036データ取得(判定用)
'------------------------------------------------

'概要      :テーブル「TBCME036」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2005/10/12 Y.SIMIZU

Public Function funGet_TBCME036_2(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME036_2"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HWFGDLINE "
    sql = sql & "from TBCME036 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME036_2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''抽出結果を格納する
     With tGetRec
        .HWFGDLINE = fncNullCheck(rs("HWFGDLINE"))                                      ' EPD上限                   2003/12/12 SystemBrain Null対応
    End With
    Set rs = Nothing

    funGet_TBCME036_2 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME021データ取得(判定用)
'------------------------------------------------

'概要      :テーブル「TBCME021」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2003/09/10 新規作成　システムブレイン

Public Function funGet_TBCME021(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME021"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HWFTYPE, HWFRMIN, HWFRMAX, HWFRSPOH, HWFRSPOT, HWFRSPOI, "
    sql = sql & "HWFRHWYT, HWFRHWYS, HWFRMCAL, HWFRAMIN, HWFRAMAX, HWFRMBNP "
    sql = sql & "from TBCME021 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME021 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''抽出結果を格納する
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' 品番
'        .HIN.mnorevno = rs("MNOREVNO")      ' 製品番号改訂番号
'        .HIN.factory = rs("FACTORY")        ' 工場
'        .HIN.opecond = rs("OPECOND")        ' 操業条件
        
        .HWFTYPE = rs("HWFTYPE")                        ' 品ＷＦタイプ
        .HWFRMIN = fncNullCheck(rs("HWFRMIN"))          ' 品ＷＦ比抵抗下限          2003/12/12 SystemBrain Null対応
        .HWFRMAX = fncNullCheck(rs("HWFRMAX"))          ' 品ＷＦ比抵抗上限          2003/12/12 SystemBrain Null対応
        .HWFRSPOH = rs("HWFRSPOH")                      ' 品ＷＦ比抵抗測定位置＿方
        .HWFRSPOT = rs("HWFRSPOT")                      ' 品ＷＦ比抵抗測定位置＿点
        .HWFRSPOI = rs("HWFRSPOI")                      ' 品ＷＦ比抵抗測定位置＿位
        .HWFRHWYT = rs("HWFRHWYT")                      ' 品ＷＦ比抵抗保証方法＿対
        .HWFRHWYS = rs("HWFRHWYS")                      ' 品ＷＦ比抵抗保証方法＿処
        .HWFRMCAL = rs("HWFRMCAL")                      ' 品ＷＦ比抵抗面内計算
        .HWFRAMIN = fncNullCheck(rs("HWFRAMIN"))        ' 品ＷＦ比抵抗平均下限      2003/12/12 SystemBrain Null対応
        .HWFRAMAX = fncNullCheck(rs("HWFRAMAX"))        ' 品ＷＦ比抵抗平均上限      2003/12/12 SystemBrain Null対応
        .HWFRMBNP = fncNullCheck(rs("HWFRMBNP"))        ' 品ＷＦ比抵抗面内分布      2003/12/12 SystemBrain Null対応
    End With
    Set rs = Nothing

    funGet_TBCME021 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME024データ取得(判定用)
'------------------------------------------------

'概要      :テーブル「TBCME024」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2003/09/10 新規作成　システムブレイン

Public Function funGet_TBCME024(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME024"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HWFMKMIN, HWFMKMAX, HWFMKSPH, HWFMKSPT, HWFMKSPR, HWFMKHWT, HWFMKHWS "
    sql = sql & "from TBCME024 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME024 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''抽出結果を格納する
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' 品番
'        .HIN.mnorevno = rs("MNOREVNO")      ' 製品番号改訂番号
'        .HIN.factory = rs("FACTORY")        ' 工場
'        .HIN.opecond = rs("OPECOND")        ' 操業条件
        
        .HWFMKMIN = fncNullCheck(rs("HWFMKMIN"))        ' 品ＷＦ無欠陥層下限            2003/12/12 SystemBrain Null対応
        .HWFMKMAX = fncNullCheck(rs("HWFMKMAX"))        ' 品ＷＦ無欠陥層上限            2003/12/12 SystemBrain Null対応
        .HWFMKSPH = rs("HWFMKSPH")                      ' 品ＷＦ無欠陥層測定位置＿方
        .HWFMKSPT = rs("HWFMKSPT")                      ' 品ＷＦ無欠陥層測定位置＿点
        .HWFMKSPR = rs("HWFMKSPR")                      ' 品ＷＦ無欠陥層測定位置＿領
        .HWFMKHWT = rs("HWFMKHWT")                      ' 品ＷＦ無欠陥層保証方法＿対
        .HWFMKHWS = rs("HWFMKHWS")                      ' 品ＷＦ無欠陥層保証方法＿処
    End With
    Set rs = Nothing

    funGet_TBCME024 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME025データ取得
'------------------------------------------------

'概要      :テーブル「TBCME025」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2003/09/10 新規作成　システムブレイン

Public Function funGet_TBCME025(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME025"

    sql = "select E025.HINBAN, E025.MNOREVNO, E025.FACTORY, E025.OPECOND, "
    sql = sql & "E025.HWFONMIN, E025.HWFONMAX, E025.HWFONSPH, E025.HWFONSPT, E025.HWFONSPI, E025.HWFONHWT, E025.HWFONHWS, "
    sql = sql & "HSXONSPT, HSXONSPI, "
    sql = sql & "E025.HWFONMCL, E025.HWFONMBP, E025.HWFONAMN, E025.HWFONAMX, "
    sql = sql & "E025.HWFOS1MN, E025.HWFOS1MX, E025.HWFOS1SH, E025.HWFOS1ST, E025.HWFOS1SI, E025.HWFOS1HT, E025.HWFOS1HS, E025.HWFOS1NS, "
    sql = sql & "E025.HWFOS2MN, E025.HWFOS2MX, E025.HWFOS2SH, E025.HWFOS2ST, E025.HWFOS2SI, E025.HWFOS2HT, E025.HWFOS2HS, E025.HWFOS2NS, "
    sql = sql & "E025.HWFOS3MN, E025.HWFOS3MX, E025.HWFOS3SH, E025.HWFOS3ST, E025.HWFOS3SI, E025.HWFOS3HT, E025.HWFOS3HS, E025.HWFOS3NS, "
    ''残存酸素仕様取得追加　03/12/09 ooba
    sql = sql & "E025.HWFZOMIN, E025.HWFZOMAX, E025.HWFZOSPH, E025.HWFZOSPT, E025.HWFZOSPI, E025.HWFZOHWT, E025.HWFZOHWS, E025.HWFZONSW, "
    sql = sql & "E025.HWFANTNP, E025.HWFANTIM "
    sql = sql & "from TBCME025 E025, TBCME019 E019 "
    sql = sql & "Where E025.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E025.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E025.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E025.OPECOND = '" & tHIN.opecond & "' and "
    sql = sql & "      E019.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E019.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E019.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E019.OPECOND = '" & tHIN.opecond & "'"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME025 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''抽出結果を格納する
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' 品番
'        .HIN.mnorevno = rs("MNOREVNO")      ' 製品番号改訂番号
'        .HIN.factory = rs("FACTORY")        ' 工場
'        .HIN.opecond = rs("OPECOND")        ' 操業条件
        
        .HWFONMIN = fncNullCheck(rs("HWFONMIN"))        ' 品ＷＦ酸素濃度下限            2003/12/12 SystemBrain Null対応
        .HWFONMAX = fncNullCheck(rs("HWFONMAX"))        ' 品ＷＦ酸素濃度上限            2003/12/12 SystemBrain Null対応
        .HWFONSPH = rs("HWFONSPH")                      ' 品ＷＦ酸素濃度測定位置＿方
'        .HWFONSPT = rs("HWFONSPT")                      ' 品ＷＦ酸素濃度測定位置＿点
'        .HWFONSPI = rs("HWFONSPI")                      ' 品ＷＦ酸素濃度測定位置＿位
        .HWFONSPT = rs("HSXONSPT")                      ' 品ＳＸ酸素濃度測定位置＿点
        .HWFONSPI = rs("HSXONSPI")                      ' 品ＳＸ酸素濃度測定位置＿位
        .HWFONHWT = rs("HWFONHWT")                      ' 品ＷＦ酸素濃度保証方法＿対
        .HWFONHWS = rs("HWFONHWS")                      ' 品ＷＦ酸素濃度保証方法＿処
        .HWFONMCL = rs("HWFONMCL")                      ' 品ＷＦ酸素濃度面内計算
        .HWFONMBP = fncNullCheck(rs("HWFONMBP"))        ' 品ＷＦ酸素濃度面内分布        2003/12/12 SystemBrain Null対応
        .HWFONAMN = fncNullCheck(rs("HWFONAMN"))        ' 品ＷＦ酸素濃度平均下限        2003/12/12 SystemBrain Null対応
        .HWFONAMX = fncNullCheck(rs("HWFONAMX"))        ' 品ＷＦ酸素濃度平均上限        2003/12/12 SystemBrain Null対応
        
        .HWFOS1MN = fncNullCheck(rs("HWFOS1MN"))        ' 品ＷＦ酸素析出１下限          2003/12/12 SystemBrain Null対応
        .HWFOS1MX = fncNullCheck(rs("HWFOS1MX"))        ' 品ＷＦ酸素析出１上限          2003/12/12 SystemBrain Null対応
        .HWFOS1SH = rs("HWFOS1SH")                      ' 品ＷＦ酸素析出１測定位置＿方
        .HWFOS1ST = rs("HWFOS1ST")                      ' 品ＷＦ酸素析出１測定位置＿点
        .HWFOS1SI = rs("HWFOS1SI")                      ' 品ＷＦ酸素析出１測定位置＿位
        .HWFOS1HT = rs("HWFOS1HT")                      ' 品ＷＦ酸素析出１保証方法＿対
        .HWFOS1HS = rs("HWFOS1HS")                      ' 品ＷＦ酸素析出１保証方法＿処
        .HWFOS1NS = rs("HWFOS1NS")                      ' 品ＷＦ酸素析出１熱処理法
        
        .HWFOS2MN = fncNullCheck(rs("HWFOS2MN"))        ' 品ＷＦ酸素析出２下限          2003/12/12 SystemBrain Null対応
        .HWFOS2MX = fncNullCheck(rs("HWFOS2MX"))        ' 品ＷＦ酸素析出２上限          2003/12/12 SystemBrain Null対応
        .HWFOS2SH = rs("HWFOS2SH")                      ' 品ＷＦ酸素析出２測定位置＿方
        .HWFOS2ST = rs("HWFOS2ST")                      ' 品ＷＦ酸素析出２測定位置＿点
        .HWFOS2SI = rs("HWFOS2SI")                      ' 品ＷＦ酸素析出２測定位置＿位
        .HWFOS2HT = rs("HWFOS2HT")                      ' 品ＷＦ酸素析出２保証方法＿対
        .HWFOS2HS = rs("HWFOS2HS")                      ' 品ＷＦ酸素析出２保証方法＿処
        .HWFOS2NS = rs("HWFOS2NS")                      ' 品ＷＦ酸素析出２熱処理法
        
        .HWFOS3MN = fncNullCheck(rs("HWFOS3MN"))        ' 品ＷＦ酸素析出３下限          2003/12/12 SystemBrain Null対応
        .HWFOS3MX = fncNullCheck(rs("HWFOS3MX"))        ' 品ＷＦ酸素析出３上限          2003/12/12 SystemBrain Null対応
        .HWFOS3SH = rs("HWFOS3SH")                      ' 品ＷＦ酸素析出３測定位置＿方
        .HWFOS3ST = rs("HWFOS3ST")                      ' 品ＷＦ酸素析出３測定位置＿点
        .HWFOS3SI = rs("HWFOS3SI")                      ' 品ＷＦ酸素析出３測定位置＿位
        .HWFOS3HT = rs("HWFOS3HT")                      ' 品ＷＦ酸素析出３保証方法＿対
        .HWFOS3HS = rs("HWFOS3HS")                      ' 品ＷＦ酸素析出３保証方法＿処
        .HWFOS3NS = rs("HWFOS3NS")                      ' 品ＷＦ酸素析出３熱処理法
        
        ''残存酸素仕様取得追加　03/12/09 ooba START ==============================>
'''        If IsNull(rs("HWFZOMIN")) = False Then .HWFZOMIN = rs("HWFZOMIN") ' 品ＷＦ残存酸素下限
'''        If IsNull(rs("HWFZOMAX")) = False Then .HWFZOMAX = rs("HWFZOMAX") ' 品ＷＦ残存酸素上限
'''        .HWFZOSPH = rs("HWFZOSPH")                  ' 品ＷＦ残存酸素測定位置＿方
'''        .HWFZOSPT = rs("HWFZOSPT")                  ' 品ＷＦ残存酸素測定位置＿点
'''        .HWFZOSPI = rs("HWFZOSPI")                  ' 品ＷＦ残存酸素測定位置＿位
'''        .HWFZOHWT = rs("HWFZOHWT")                  ' 品ＷＦ残存酸素保証方法＿対
'''        .HWFZOHWS = rs("HWFZOHWS")                  ' 品ＷＦ残存酸素保証方法＿処
'''        .HWFZONSW = rs("HWFZONSW")                  ' 品ＷＦ残存酸素熱処理法

        .HWFZOMIN = fncNullCheck(rs("HWFZOMIN"))    ' 品ＷＦ残存酸素下限
        .HWFZOMAX = fncNullCheck(rs("HWFZOMAX"))    ' 品ＷＦ残存酸素上限
        If IsNull(rs("HWFZOSPH")) = False Then .HWFZOSPH = rs("HWFZOSPH") ' 品ＷＦ残存酸素測定位置＿方
        If IsNull(rs("HWFZOSPT")) = False Then .HWFZOSPT = rs("HWFZOSPT") ' 品ＷＦ残存酸素測定位置＿点
        If IsNull(rs("HWFZOSPI")) = False Then .HWFZOSPI = rs("HWFZOSPI") ' 品ＷＦ残存酸素測定位置＿位
        If IsNull(rs("HWFZOHWT")) = False Then .HWFZOHWT = rs("HWFZOHWT") ' 品ＷＦ残存酸素保証方法＿対
        If IsNull(rs("HWFZOHWS")) = False Then .HWFZOHWS = rs("HWFZOHWS") ' 品ＷＦ残存酸素保証方法＿処
        If IsNull(rs("HWFZONSW")) = False Then .HWFZONSW = rs("HWFZONSW") ' 品ＷＦ残存酸素熱処理法
        ''残存酸素仕様取得追加　03/12/09 ooba END ================================>
        
        .HWFANTIM = fncNullCheck(rs("HWFANTIM"))        ' 品ＷＦＡＮ時間                2003/12/12 SystemBrain Null対応
        .HWFANTNP = fncNullCheck(rs("HWFANTNP"))        ' 品ＷＦＡＮ温度                2003/12/12 SystemBrain Null対応
    End With
    Set rs = Nothing

    funGet_TBCME025 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME026データ取得
'------------------------------------------------

'概要      :テーブル「TBCME026」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2003/09/10 新規作成　システムブレイン

Public Function funGet_TBCME026(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME026"

    'DSODﾊﾟﾀｰﾝ区分取得追加　04/08/09
    'GD仕様取得追加　05/01/26
''    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
''    sql = sql & "HWFDSOPTK, "
''    sql = sql & "HWFDENKU, HWFDENMX, HWFDENMN, HWFDENHT, HWFDENHS, "
''    sql = sql & "HWFDVDKU, HWFDVDMXN, HWFDVDMNN, HWFDVDHT, HWFDVDHS, "
''    sql = sql & "HWFLDLKU, HWFLDLMX, HWFLDLMN, HWFLDLHT, HWFLDLHS, "
''    sql = sql & "HWFGDSPH, HWFGDSPT, HWFGDSPR, "
''    sql = sql & "HWFDSOMX, HWFDSOMN, HWFDSOAX, HWFDSOAN, HWFDSOHT, HWFDSOHS "
''    sql = sql & "from TBCME026 "
''    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
''    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
''    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
''    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    'DKｱﾆｰﾙ温度追加　06/12/22 ooba
    sql = "select E025.HINBAN, E025.MNOREVNO, E025.FACTORY, E025.OPECOND, "
    sql = sql & "E026.HWFDENKU, E026.HWFDENMX, E026.HWFDENMN, E026.HWFDENHT, E026.HWFDENHS, "
    sql = sql & "E026.HWFDVDKU, E026.HWFDVDMXN, E026.HWFDVDMNN, E026.HWFDVDHT, E026.HWFDVDHS, "
    sql = sql & "E026.HWFLDLKU, E026.HWFLDLMX, E026.HWFLDLMN, E026.HWFLDLHT, E026.HWFLDLHS, "
    sql = sql & "E026.HWFGDSPH, E026.HWFGDSPT, E026.HWFGDSPR, "
    sql = sql & "E026.HWFDSOMX, E026.HWFDSOMN, E026.HWFDSOAX, E026.HWFDSOAN, E026.HWFDSOHT, "
    sql = sql & "E026.HWFDSOHS, E026.HWFDSOPTK, E025.HWFANTNP "
    sql = sql & ",E026.HWFGDPTK "    '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    sql = sql & "from TBCME025 E025, TBCME026 E026 "
    sql = sql & "Where E025.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E025.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E025.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E025.OPECOND = '" & tHIN.opecond & "' and "
    sql = sql & "      E026.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E026.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E026.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E026.OPECOND = '" & tHIN.opecond & "'"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME026 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''抽出結果を格納する
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' 品番
'        .HIN.mnorevno = rs("MNOREVNO")      ' 製品番号改訂番号
'        .HIN.factory = rs("FACTORY")        ' 工場
'        .HIN.opecond = rs("OPECOND")        ' 操業条件
        
        .HWFDSOMX = fncNullCheck(rs("HWFDSOMX"))        ' 品ＷＦＤＳＯＤ上限            2003/12/12 SystemBrain Null対応
        .HWFDSOMN = fncNullCheck(rs("HWFDSOMN"))        ' 品ＷＦＤＳＯＤ下限            2003/12/12 SystemBrain Null対応
        .HWFDSOAX = fncNullCheck(rs("HWFDSOAX"))        ' 品ＷＦＤＳＯＤ領域上限        2003/12/12 SystemBrain Null対応
        .HWFDSOAN = fncNullCheck(rs("HWFDSOAN"))        ' 品ＷＦＤＳＯＤ領域下限        2003/12/12 SystemBrain Null対応
        .HWFDSOHT = rs("HWFDSOHT")                      ' 品ＷＦＤＳＯＤ保証方法＿対
        .HWFDSOHS = rs("HWFDSOHS")                      ' 品ＷＦＤＳＯＤ保証方法＿処
        If IsNull(rs("HWFDSOPTK")) = False Then .HWFDSOPTK = rs("HWFDSOPTK") Else .HWFDSOPTK = " "          'パターン区分　04/08/09 ooba
        
        ''GD仕様取得追加　05/01/26 ooba START ========================================>
        .HWFDENKU = rs("HWFDENKU")                      ' 品ＷＦＤｅｎ検査有無
        .HWFDENMX = fncNullCheck(rs("HWFDENMX"))        ' 品ＷＦＤｅｎ上限
        .HWFDENMN = fncNullCheck(rs("HWFDENMN"))        ' 品ＷＦＤｅｎ下限
        .HWFDENHT = rs("HWFDENHT")                      ' 品ＷＦＤｅｎ保証方法＿対
        .HWFDENHS = rs("HWFDENHS")                      ' 品ＷＦＤｅｎ保証方法＿処
        .HWFDVDKU = rs("HWFDVDKU")                      ' 品ＷＦＤＶＤ２検査有無
        .HWFDVDMXN = fncNullCheck(rs("HWFDVDMXN"))      ' 品ＷＦＤＶＤ２上限
        .HWFDVDMNN = fncNullCheck(rs("HWFDVDMNN"))      ' 品ＷＦＤＶＤ２下限
        .HWFDVDHT = rs("HWFDVDHT")                      ' 品ＷＦＤＶＤ２保証方法＿対
        .HWFDVDHS = rs("HWFDVDHS")                      ' 品ＷＦＤＶＤ２保証方法＿処
        .HWFLDLKU = rs("HWFLDLKU")                      ' 品ＷＦＬ／ＤＬ検査有無
        .HWFLDLMX = fncNullCheck(rs("HWFLDLMX"))        ' 品ＷＦＬ／ＤＬ上限
        .HWFLDLMN = fncNullCheck(rs("HWFLDLMN"))        ' 品ＷＦＬ／ＤＬ下限
        .HWFLDLHT = rs("HWFLDLHT")                      ' 品ＷＦＬ／ＤＬ保証方法＿対
        .HWFLDLHS = rs("HWFLDLHS")                      ' 品ＷＦＬ／ＤＬ保証方法＿処
        .HWFGDSPH = rs("HWFGDSPH")                      ' 品ＷＦＧＤ測定位置＿方
        .HWFGDSPT = rs("HWFGDSPT")                      ' 品ＷＦＧＤ測定位置＿点
        .HWFGDSPR = rs("HWFGDSPR")                      ' 品ＷＦＧＤ測定位置＿領
        ''GD仕様取得追加　05/01/26 ooba END ==========================================>
        
        If Not IsNull(rs("HWFANTNP")) Then .HWFANTNP = rs("HWFANTNP")   ' 品ＷＦＡＮ温度　06/12/22 ooba
        
        If Not IsNull(rs("HWFGDPTK")) Then .HWFGDPTK = rs("HWFGDPTK") Else .HWFGDPTK = " "  ' 品ＷＦＧＤパタン区分  '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    End With
    Set rs = Nothing

    funGet_TBCME026 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME028データ取得
'------------------------------------------------

'概要      :テーブル「TBCME028」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2003/09/10 新規作成　システムブレイン

Public Function funGet_TBCME028(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME028"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "

''Upd start 2005/06/28 (TCS)T.terauchi  SPV9点対応
''    sql = sql & "HWFSPVMX, HWFSPVSH, HWFSPVST, HWFSPVSI, HWFSPVHT, HWFSPVHS, "
    sql = sql & "HWFSPVMX, HWFSPVMXN, HWFSPVSH, HWFSPVST, HWFSPVSI, HWFSPVHT, HWFSPVHS, "
    sql = sql & "HWFSPVKN, HWFDLKHN, "
''Upd end   2005/06/28 (TCS)T.Terauchi  SPV9点対応

'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
    sql = sql & "HWFSPVAMN, "
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
    
    sql = sql & "HWFDLMIN, HWFDLMAX, HWFDLSPH, HWFDLSPT, HWFDLSPI, HWFDLHWT, HWFDLHWS "
    sql = sql & "from TBCME028 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME028 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''抽出結果を格納する
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' 品番
'        .HIN.mnorevno = rs("MNOREVNO")      ' 製品番号改訂番号
'        .HIN.factory = rs("FACTORY")        ' 工場
'        .HIN.opecond = rs("OPECOND")        ' 操業条件
        
    ''Upd start 2005/06/28 (TCS)T.Terauchi  SPV9点対応
    ''    .HWFSPVMX = fncNullCheck(rs("HWFSPVMX"))        ' 品ＷＦＳＰＶＦＥ上限          2003/12/12 SystemBrain Null対応
        .HWFSPVMX = fncNullCheck(rs("HWFSPVMXN"))       ' 品ＷＦＳＰＶＦＥ上限
        .HWFSPVKN = rs("HWFSPVKN")                      ' 品ＷＦＳＰＶＦＥ検査頻度＿抜
        .HWFDLKHN = rs("HWFDLKHN")                      ' 品ＷＦ拡散長検査頻度＿抜
    ''Upd end   2005/06/28 (TCS)T.Terauchi  SPV9点対応
    
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        .HWFSPVAM = fncNullCheck(rs("HWFSPVAMN"))       ' 品ＷＦＳＰＶＦＥ平均
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        
        
        .HWFSPVSH = rs("HWFSPVSH")                      ' 品ＷＦＳＰＶＦＥ測定位置＿方
        .HWFSPVST = rs("HWFSPVST")                      ' 品ＷＦＳＰＶＦＥ測定位置＿点
        .HWFSPVSI = rs("HWFSPVSI")                      ' 品ＷＦＳＰＶＦＥ測定位置＿位
        .HWFSPVHT = rs("HWFSPVHT")                      ' 品ＷＦＳＰＶＦＥ保証方法＿対
        .HWFSPVHS = rs("HWFSPVHS")                      ' 品ＷＦＳＰＶＦＥ保証方法＿処
        
        .HWFDLMIN = fncNullCheck(rs("HWFDLMIN"))        ' 品ＷＦ拡散長下限              2003/12/12 SystemBrain Null対応
        .HWFDLMAX = fncNullCheck(rs("HWFDLMAX"))        ' 品ＷＦ拡散長上限              2003/12/12 SystemBrain Null対応
        .HWFDLSPH = rs("HWFDLSPH")                      ' 品ＷＦ拡散長測定位置＿方
        .HWFDLSPT = rs("HWFDLSPT")                      ' 品ＷＦ拡散長測定位置＿点
        .HWFDLSPI = rs("HWFDLSPI")                      ' 品ＷＦ拡散長測定位置＿位
        .HWFDLHWT = rs("HWFDLHWT")                      ' 品ＷＦ拡散長保証方法＿対
        .HWFDLHWS = rs("HWFDLHWS")                      ' 品ＷＦ拡散長保証方法＿処
    End With
    Set rs = Nothing

    funGet_TBCME028 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME029データ取得
'------------------------------------------------

'概要      :テーブル「TBCME029」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2003/09/10 新規作成　システムブレイン

Public Function funGet_TBCME029(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME029"

'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    'AN温度チェックの為にTBCME025から品ＷＦＡＮ温度を取得する
'    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
'    sql = sql & "HWFOF1AX, HWFOF1MX, HWFOF1SH, HWFOF1ST, HWFOF1SR, HWFOF1HT, HWFOF1HS, HWFOF1NS, HWFOF1ET, HWFOF1SZ, "
'    sql = sql & "HWFOF2AX, HWFOF2MX, HWFOF2SH, HWFOF2ST, HWFOF2SR, HWFOF2HT, HWFOF2HS, HWFOF2NS, HWFOF2ET, HWFOF2SZ, "
'    sql = sql & "HWFOF3AX, HWFOF3MX, HWFOF3SH, HWFOF3ST, HWFOF3SR, HWFOF3HT, HWFOF3HS, HWFOF3NS, HWFOF3ET, HWFOF3SZ, "
'    sql = sql & "HWFOF4AX, HWFOF4MX, HWFOF4SH, HWFOF4ST, HWFOF4SR, HWFOF4HT, HWFOF4HS, HWFOF4NS, HWFOF4ET, HWFOF4SZ, "
'    sql = sql & "HWFBM1AN, HWFBM1AX, HWFBM1SH, HWFBM1ST, HWFBM1SR, HWFBM1HT, HWFBM1HS, HWFBM1NS, HWFBM1ET, HWFBM1SZ, "
'    sql = sql & "HWFBM2AN, HWFBM2AX, HWFBM2SH, HWFBM2ST, HWFBM2SR, HWFBM2HT, HWFBM2HS, HWFBM2NS, HWFBM2ET, HWFBM2SZ, "
'    sql = sql & "HWFBM3AN, HWFBM3AX, HWFBM3SH, HWFBM3ST, HWFBM3SR, HWFBM3HT, HWFBM3HS, HWFBM3NS, HWFBM3ET, HWFBM3SZ, "
'    sql = sql & "HWFOSF1PTK, HWFOSF2PTK, HWFOSF3PTK, HWFOSF4PTK, "
'    sql = sql & "HWFBM1MBP, HWFBM2MBP, HWFBM3MBP, HWFBM1MCL, HWFBM2MCL, HWFBM3MCL "
'    sql = sql & "from TBCME029 "
'    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
'    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
'    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
'    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    sql = "select E029.HINBAN, E029.MNOREVNO, E029.FACTORY, E029.OPECOND, "
    sql = sql & "E029.HWFOF1AX, E029.HWFOF1MX, E029.HWFOF1SH, E029.HWFOF1ST, E029.HWFOF1SR, E029.HWFOF1HT, E029.HWFOF1HS, E029.HWFOF1NS, E029.HWFOF1ET, HWFOF1SZ, "
    sql = sql & "E029.HWFOF2AX, E029.HWFOF2MX, E029.HWFOF2SH, E029.HWFOF2ST, E029.HWFOF2SR, E029.HWFOF2HT, E029.HWFOF2HS, E029.HWFOF2NS, E029.HWFOF2ET, HWFOF2SZ, "
    sql = sql & "E029.HWFOF3AX, E029.HWFOF3MX, E029.HWFOF3SH, E029.HWFOF3ST, E029.HWFOF3SR, E029.HWFOF3HT, E029.HWFOF3HS, E029.HWFOF3NS, E029.HWFOF3ET, HWFOF3SZ, "
    sql = sql & "E029.HWFOF4AX, E029.HWFOF4MX, E029.HWFOF4SH, E029.HWFOF4ST, E029.HWFOF4SR, E029.HWFOF4HT, E029.HWFOF4HS, E029.HWFOF4NS, E029.HWFOF4ET, HWFOF4SZ, "
    sql = sql & "E029.HWFBM1AN, E029.HWFBM1AX, E029.HWFBM1SH, E029.HWFBM1ST, E029.HWFBM1SR, E029.HWFBM1HT, E029.HWFBM1HS, E029.HWFBM1NS, E029.HWFBM1ET, HWFBM1SZ, "
    sql = sql & "E029.HWFBM2AN, E029.HWFBM2AX, E029.HWFBM2SH, E029.HWFBM2ST, E029.HWFBM2SR, E029.HWFBM2HT, E029.HWFBM2HS, E029.HWFBM2NS, E029.HWFBM2ET, HWFBM2SZ, "
    sql = sql & "E029.HWFBM3AN, E029.HWFBM3AX, E029.HWFBM3SH, E029.HWFBM3ST, E029.HWFBM3SR, E029.HWFBM3HT, E029.HWFBM3HS, E029.HWFBM3NS, E029.HWFBM3ET, HWFBM3SZ, "
    sql = sql & "E029.HWFOSF1PTK, E029.HWFOSF2PTK, E029.HWFOSF3PTK, E029.HWFOSF4PTK, "
    sql = sql & "E029.HWFBM1MBP, E029.HWFBM2MBP, E029.HWFBM3MBP, E029.HWFBM1MCL, E029.HWFBM2MCL, E029.HWFBM3MCL, "
    sql = sql & "E025.HWFANTNP "
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & ",E048.HWFSIRDMX "          '軸状転位上限
    sql = sql & ",E048.HWFSIRDHT "          '軸状転位保証方法＿対
    sql = sql & ",E048.HWFSIRDHS "          '軸状転位保証方法＿処
    sql = sql & ",E048.HWFSIRDSZ "          '軸状転位測定条件
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
    sql = sql & "from TBCME029 E029 "
    sql = sql & "    ,TBCME025 E025 "
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "    ,TBCME048 E048 "
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
    sql = sql & "Where E029.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E029.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E029.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E029.OPECOND = '" & tHIN.opecond & "' and "
    sql = sql & "      E025.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E025.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E025.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E025.OPECOND = '" & tHIN.opecond & "'"
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "  and E048.HINBAN = '" & tHIN.HINBAN & "' and "
    sql = sql & "      E048.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E048.FACTORY = '" & tHIN.FACTORY & "' and "
    sql = sql & "      E048.OPECOND = '" & tHIN.OPECOND & "'"
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME029 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''抽出結果を格納する
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' 品番
'        .HIN.mnorevno = rs("MNOREVNO")      ' 製品番号改訂番号
'        .HIN.factory = rs("FACTORY")        ' 工場
'        .HIN.opecond = rs("OPECOND")        ' 操業条件
        
        .HWFOF1AX = fncNullCheck(rs("HWFOF1AX"))        ' 品ＷＦＯＳＦ１平均上限        2003/12/12 SystemBrain Null対応
        .HWFOF1MX = fncNullCheck(rs("HWFOF1MX"))        ' 品ＷＦＯＳＦ１上限            2003/12/12 SystemBrain Null対応
        .HWFOF1SH = rs("HWFOF1SH")                      ' 品ＷＦＯＳＦ１測定位置＿方
        .HWFOF1ST = rs("HWFOF1ST")                      ' 品ＷＦＯＳＦ１測定位置＿点
        .HWFOF1SR = rs("HWFOF1SR")                      ' 品ＷＦＯＳＦ１測定位置＿領
        .HWFOF1HT = rs("HWFOF1HT")                      ' 品ＷＦＯＳＦ１保証方法＿対
        .HWFOF1HS = rs("HWFOF1HS")                      ' 品ＷＦＯＳＦ１保証方法＿処
        .HWFOF1NS = rs("HWFOF1NS")                      ' 品ＷＦＯＳＦ１熱処理法
        .HWFOF1ET = fncNullCheck(rs("HWFOF1ET"))        ' 品ＷＦＯＳＦ１選択ＥＴ代      2003/12/12 SystemBrain Null対応
        .HWFOF1SZ = rs("HWFOF1SZ")                      ' 品ＷＦＯＳＦ１測定条件
        .HWFOF2AX = fncNullCheck(rs("HWFOF2AX"))        ' 品ＷＦＯＳＦ２平均上限        2003/12/12 SystemBrain Null対応
        .HWFOF2MX = fncNullCheck(rs("HWFOF2MX"))        ' 品ＷＦＯＳＦ２上限            2003/12/12 SystemBrain Null対応
        .HWFOF2SH = rs("HWFOF2SH")                      ' 品ＷＦＯＳＦ２測定位置＿方
        .HWFOF2ST = rs("HWFOF2ST")                      ' 品ＷＦＯＳＦ２測定位置＿点
        .HWFOF2SR = rs("HWFOF2SR")                      ' 品ＷＦＯＳＦ２測定位置＿領
        .HWFOF2HT = rs("HWFOF2HT")                      ' 品ＷＦＯＳＦ２保証方法＿対
        .HWFOF2HS = rs("HWFOF2HS")                      ' 品ＷＦＯＳＦ２保証方法＿処
        .HWFOF2NS = rs("HWFOF2NS")                      ' 品ＷＦＯＳＦ２熱処理法
        .HWFOF2ET = fncNullCheck(rs("HWFOF2ET"))        ' 品ＷＦＯＳＦ２選択ＥＴ代      2003/12/12 SystemBrain Null対応
        .HWFOF2SZ = rs("HWFOF2SZ")                      ' 品ＷＦＯＳＦ２測定条件
        .HWFOF3AX = fncNullCheck(rs("HWFOF3AX"))        ' 品ＷＦＯＳＦ３平均上限        2003/12/12 SystemBrain Null対応
        .HWFOF3MX = fncNullCheck(rs("HWFOF3MX"))        ' 品ＷＦＯＳＦ３上限            2003/12/12 SystemBrain Null対応
        .HWFOF3SH = rs("HWFOF3SH")                      ' 品ＷＦＯＳＦ３測定位置＿方
        .HWFOF3ST = rs("HWFOF3ST")                      ' 品ＷＦＯＳＦ３測定位置＿点
        .HWFOF3SR = rs("HWFOF3SR")                      ' 品ＷＦＯＳＦ３測定位置＿領
        .HWFOF3HT = rs("HWFOF3HT")                      ' 品ＷＦＯＳＦ３保証方法＿対
        .HWFOF3HS = rs("HWFOF3HS")                      ' 品ＷＦＯＳＦ３保証方法＿処
        .HWFOF3NS = rs("HWFOF3NS")                      ' 品ＷＦＯＳＦ３熱処理法
        .HWFOF3ET = fncNullCheck(rs("HWFOF3ET"))        ' 品ＷＦＯＳＦ３選択ＥＴ代      2003/12/12 SystemBrain Null対応
        .HWFOF3SZ = rs("HWFOF3SZ")                      ' 品ＷＦＯＳＦ３測定条件
        .HWFOF4AX = fncNullCheck(rs("HWFOF4AX"))        ' 品ＷＦＯＳＦ４平均上限        2003/12/12 SystemBrain Null対応
        .HWFOF4MX = fncNullCheck(rs("HWFOF4MX"))        ' 品ＷＦＯＳＦ４上限            2003/12/12 SystemBrain Null対応
        .HWFOF4SH = rs("HWFOF4SH")                      ' 品ＷＦＯＳＦ４測定位置＿方
        .HWFOF4ST = rs("HWFOF4ST")                      ' 品ＷＦＯＳＦ４測定位置＿点
        .HWFOF4SR = rs("HWFOF4SR")                      ' 品ＷＦＯＳＦ４測定位置＿領
        .HWFOF4HT = rs("HWFOF4HT")                      ' 品ＷＦＯＳＦ４保証方法＿対
        .HWFOF4HS = rs("HWFOF4HS")                      ' 品ＷＦＯＳＦ４保証方法＿処
        .HWFOF4NS = rs("HWFOF4NS")                      ' 品ＷＦＯＳＦ４熱処理法
        .HWFOF4ET = fncNullCheck(rs("HWFOF4ET"))        ' 品ＷＦＯＳＦ４選択ＥＴ代      2003/12/12 SystemBrain Null対応
        .HWFOF4SZ = rs("HWFOF4SZ")                      ' 品ＷＦＯＳＦ４測定条件
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
        If IsNull(rs("HWFSIRDMX")) = False Then .HWFOF4MX = rs("HWFSIRDMX") Else .HWFOF4MX = "0"        ' 軸状転位上限
        If IsNull(rs("HWFSIRDHT")) = False Then .HWFOF4HT = rs("HWFSIRDHT") Else .HWFOF4HT = " "        ' 軸状転位保証方法＿対
        If IsNull(rs("HWFSIRDHS")) = False Then .HWFOF4HS = rs("HWFSIRDHS") Else .HWFOF4HS = " "        ' 軸状転位保証方法＿処
        If IsNull(rs("HWFSIRDSZ")) = False Then .HWFOF4SZ = rs("HWFSIRDSZ") Else .HWFOF4SZ = " "        ' 軸状転位測定条件
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
        
        .HWFBM1AN = fncNullCheck(rs("HWFBM1AN"))        ' 品ＷＦＢＭＤ１平均下限        2003/12/12 SystemBrain Null対応
        .HWFBM1AX = fncNullCheck(rs("HWFBM1AX"))        ' 品ＷＦＢＭＤ１平均上限        2003/12/12 SystemBrain Null対応
        .HWFBM1SH = rs("HWFBM1SH")                      ' 品ＷＦＢＭＤ１測定位置＿方
        .HWFBM1ST = rs("HWFBM1ST")                      ' 品ＷＦＢＭＤ１測定位置＿点
        .HWFBM1SR = rs("HWFBM1SR")                      ' 品ＷＦＢＭＤ１測定位置＿領
        .HWFBM1HT = rs("HWFBM1HT")                      ' 品ＷＦＢＭＤ１保証方法＿対
        .HWFBM1HS = rs("HWFBM1HS")                      ' 品ＷＦＢＭＤ１保証方法＿処
        .HWFBM1NS = rs("HWFBM1NS")                      ' 品ＷＦＢＭＤ１熱処理法
        .HWFBM1ET = fncNullCheck(rs("HWFBM1ET"))        ' 品ＷＦＢＭＤ１選択ＥＴ代      2003/12/12 SystemBrain Null対応
        .HWFBM1SZ = rs("HWFBM1SZ")                      ' 品ＷＦＢＭＤ１測定条件
        .HWFBM2AN = fncNullCheck(rs("HWFBM2AN"))        ' 品ＷＦＢＭＤ２平均下限        2003/12/12 SystemBrain Null対応
        .HWFBM2AX = fncNullCheck(rs("HWFBM2AX"))        ' 品ＷＦＢＭＤ２平均上限        2003/12/12 SystemBrain Null対応
        .HWFBM2SH = rs("HWFBM2SH")                      ' 品ＷＦＢＭＤ２測定位置＿方
        .HWFBM2ST = rs("HWFBM2ST")                      ' 品ＷＦＢＭＤ２測定位置＿点
        .HWFBM2SR = rs("HWFBM2SR")                      ' 品ＷＦＢＭＤ２測定位置＿領
        .HWFBM2HT = rs("HWFBM2HT")                      ' 品ＷＦＢＭＤ２保証方法＿対
        .HWFBM2HS = rs("HWFBM2HS")                      ' 品ＷＦＢＭＤ２保証方法＿処
        .HWFBM2NS = rs("HWFBM2NS")                      ' 品ＷＦＢＭＤ２熱処理法
        .HWFBM2ET = fncNullCheck(rs("HWFBM2ET"))        ' 品ＷＦＢＭＤ２選択ＥＴ代      2003/12/12 SystemBrain Null対応
        .HWFBM2SZ = rs("HWFBM2SZ")                      ' 品ＷＦＢＭＤ２測定条件
        .HWFBM3AN = fncNullCheck(rs("HWFBM3AN"))        ' 品ＷＦＢＭＤ３平均下限        2003/12/12 SystemBrain Null対応
        .HWFBM3AX = fncNullCheck(rs("HWFBM3AX"))        ' 品ＷＦＢＭＤ３平均上限        2003/12/12 SystemBrain Null対応
        .HWFBM3SH = rs("HWFBM3SH")                      ' 品ＷＦＢＭＤ３測定位置＿方
        .HWFBM3ST = rs("HWFBM3ST")                      ' 品ＷＦＢＭＤ３測定位置＿点
        .HWFBM3SR = rs("HWFBM3SR")                      ' 品ＷＦＢＭＤ３測定位置＿領
        .HWFBM3HT = rs("HWFBM3HT")                      ' 品ＷＦＢＭＤ３保証方法＿対
        .HWFBM3HS = rs("HWFBM3HS")                      ' 品ＷＦＢＭＤ３保証方法＿処
        .HWFBM3NS = rs("HWFBM3NS")                      ' 品ＷＦＢＭＤ３熱処理法
        .HWFBM3ET = fncNullCheck(rs("HWFBM3ET"))        ' 品ＷＦＢＭＤ３選択ＥＴ代      2003/12/12 SystemBrain Null対応
        .HWFBM3SZ = rs("HWFBM3SZ")                      ' 品ＷＦＢＭＤ３測定条件
        
        If Not IsNull(rs("HWFOSF1PTK")) Then .HWFOSF1PTK = rs("HWFOSF1PTK")   ' 品ＷＦＯＳＦ１パタン区分
        If Not IsNull(rs("HWFOSF2PTK")) Then .HWFOSF2PTK = rs("HWFOSF2PTK")   ' 品ＷＦＯＳＦ２パタン区分
        If Not IsNull(rs("HWFOSF3PTK")) Then .HWFOSF3PTK = rs("HWFOSF3PTK")   ' 品ＷＦＯＳＦ３パタン区分
        If Not IsNull(rs("HWFOSF4PTK")) Then .HWFOSF4PTK = rs("HWFOSF4PTK")   ' 品ＷＦＯＳＦ４パタン区分
        
'        If Not IsNull(rs("HWFBM1MBP")) Then .HWFBM1MBP = rs("HWFBM1MBP")      ' 品ＷＦＢＭＤ１面内分布
'        If Not IsNull(rs("HWFBM2MBP")) Then .HWFBM2MBP = rs("HWFBM2MBP")      ' 品ＷＦＢＭＤ２面内分布
'        If Not IsNull(rs("HWFBM3MBP")) Then .HWFBM3MBP = rs("HWFBM3MBP")      ' 品ＷＦＢＭＤ３面内分布
        .HWFBM1MBP = fncNullCheck(rs("HWFBM1MBP"))      ' 品ＷＦＢＭＤ１面内分布        2003/12/12 SystemBrain Null対応
        .HWFBM2MBP = fncNullCheck(rs("HWFBM2MBP"))      ' 品ＷＦＢＭＤ２面内分布        2003/12/12 SystemBrain Null対応
        .HWFBM3MBP = fncNullCheck(rs("HWFBM3MBP"))      ' 品ＷＦＢＭＤ３面内分布        2003/12/12 SystemBrain Null対応
        If Not IsNull(rs("HWFBM1MCL")) Then .HWFBM1MCL = rs("HWFBM1MCL")      ' 品ＷＦＢＭＤ１面内計算
        If Not IsNull(rs("HWFBM2MCL")) Then .HWFBM2MCL = rs("HWFBM2MCL")      ' 品ＷＦＢＭＤ２面内計算
        If Not IsNull(rs("HWFBM3MCL")) Then .HWFBM3MCL = rs("HWFBM3MCL")      ' 品ＷＦＢＭＤ３面内計算
    
    '↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        'AN温度チェックの為にTBCME025から品ＷＦＡＮ温度を取得する
        If Not IsNull(rs("HWFANTNP")) Then .HWFANTNP = rs("HWFANTNP")       ' 品ＷＦＡＮ温度
    '↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    End With
    Set rs = Nothing

    funGet_TBCME029 = FUNCTION_RETURN_SUCCESS
  

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

'><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><<><><><
'概要      :抵抗推定値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名         ,IO ,型        ,説明
'          :CRYNUM        ,I  ,String    ,結晶番号
'          :TopRs         ,I  ,Double    ,TOP側推定元抵抗実績
'          :TopPos        ,I  ,Double    ,TOP側推定元位置
'          :BotRs         ,I  ,Double    ,TOP側推定元抵抗実績
'          :BotPos        ,I  ,Double    ,TOP側推定元位置
'          :SuiPos        ,I  ,Double    ,推定位置
'          :Suitei  　    ,O  ,Double    ,推定値
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :結晶番号、TOP/BOTの抵抗実績値、位置より抵抗推定を行う。
'履歴      :2003/9/4 作成  筑
Public Function new_ResSuitei(CRYNUM, TopRs, TOPPOS, BotRs, BOTPOS, SuiPos, Suitei As Double) As FUNCTION_RETURN
Dim cc As type_Coefficient  '実行偏析計算用構造体
Dim rp As type_ResPosCal    '推定計算用構造体
Dim Jikouhen As Double  '実行偏析
Dim wgtCharge As Long   'チャージ量
Dim wgtTop As Double    'トップ重量実績値
Dim wgtTopCut As Double 'トップカット重量実績値
Dim DM As Double        '直径１～３の平均
    
    new_ResSuitei = FUNCTION_RETURN_FAILURE
    
    ''実行偏析用パラメータ取得 マルチ引上対応 参照関数変更 2008/04/23 SETsw Nakada
    If GetCoeffParams_new(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
'    If GetCoeffParams(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
        Debug.Print "偏析計算用パラメータの取得に失敗した"
    End If
    
    ''ブロックの実行偏析を求める
    cc.DUNMENSEKI = AreaOfCircle(DM)    '断面積
    cc.CHARGEWEIGHT = wgtCharge         'チャージ量
    cc.TOPWEIGHT = wgtTop + wgtTopCut   'トップ重量
    cc.TOPSMPLPOS = TOPPOS
    cc.BOTSMPLPOS = BOTPOS
    cc.TOPRES = TopRs
    cc.BOTRES = BotRs
    
    Jikouhen = CoefficientCalculation(cc) '実行偏析計算
    
    
    ''推定抵抗値を求める
    If Jikouhen <> -9999 Then
        rp.COEFFICIENT = Jikouhen           '実行偏析
        rp.DUNMENSEKI = cc.DUNMENSEKI       '断面積
        rp.CHARGEWEIGHT = cc.CHARGEWEIGHT   'チャージ量
        rp.TOPWEIGHT = cc.TOPWEIGHT         'トップ重量
        rp.TOPSMPLPOS = TOPPOS
        rp.TOPRES = TopRs
        rp.target = SuiPos
        
        Suitei = ResCalculation(rp)         '推定計算
    Else
        new_ResSuitei = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    new_ResSuitei = FUNCTION_RETURN_SUCCESS

End Function
'
''><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><<><><><
''概要      :偏析計算に必要な各合計重量実績を取得する
''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
''          :CRYNUM        ,I  ,String    ,結晶番号
''          :wgtCharge     ,O  ,Long      ,炉内量（初回チャージ量－前回までの引上げ重量－前回までのﾄｯﾌﾟｶｯﾄ重量）
''          :wgtTop        ,O  ,Double    ,トップ重量実績値
''          :wgtTopCut     ,O  ,Double    ,トップカット重量実績値
''          :DM            ,O  ,Double    ,直径１～３の平均
''          :戻り値        ,O  ,FUNCTION_RETURN,
''説明      :１本引き、残量引きにあわせて実績データを取得する
''履歴      :2001/8/29 作成  野村
'Public Function GetCoeffParams(ByVal CRYNUM$, wgtCharge As Long, wgtTop As Double, wgtTopCut As Double, DM As Double) As FUNCTION_RETURN
'Dim sql As String
'Dim rs As OraDynaset
'
'    On Error GoTo Err
'    GetCoeffParams = FUNCTION_RETURN_FAILURE
'    wgtCharge = 0
'    wgtTop = 0#
'    wgtTopCut = 0#
'    DM = 0#
'
'    sql = "select decode(RONAI,null,CHARGE,RONAI) as RONAI, WGHTTOP, WGTOPCUT, (DM1+DM2+DM3)/3.0 as DM " & _
'          "from TBCMH004 H004, " & _
'          "  (select sum(CHARGE) - sum(UPWEIGHT) - sum(WGTOPCUT) as RONAI" & _
'          "   From TBCMH004" & _
'          "   where (CRYNUM<'" & CRYNUM & "')" & _
'          "    and  (substr(CRYNUM,1,7)='" & Left$(CRYNUM, 7) & "')" & _
'          "  ) SUMDATA " & _
'          "where (CRYNUM='" & CRYNUM & "')"
'
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'    If rs.RecordCount > 0 Then
'        wgtCharge = rs("RONAI")
'        wgtTop = rs("WGHTTOP")
'        wgtTopCut = rs("WGTOPCUT")
'        DM = rs("DM")
'    End If
'    rs.Close
'
'    GetCoeffParams = FUNCTION_RETURN_SUCCESS
'
'proc_exit:
'    On Error GoTo 0
'    Exit Function
'
'Err:
'    Resume proc_exit
'End Function
'
''><><><><><><><><><><><><><><><><><><><><><>><><><><><><><><><><><><><><><><><><><><><><
''概要      :位置に対する抵抗値を推定する。
''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型             ,説明
''          :d             ,IO ,type_ResPosCal ,推定計算構造体
''          :戻り値        ,O  ,Double         ,推定抵抗値
''説明      :
''履歴      :2001/06/23　佐野 信哉　作成
'Public Function ResCalculation(d As type_ResPosCal) As Double
'    Dim GS As Double
'    Dim Ro As Double
'    Dim Gx As Double
'
'    On Error GoTo Err
'    GS = (d.DUNMENSEKI * HIJU_SILICONE * d.TOPSMPLPOS) / (d.CHARGEWEIGHT - d.TOPWEIGHT)
'    Ro = d.TOPRES * (1 - GS) ^ (d.COEFFICIENT - 1)
'    Gx = d.DUNMENSEKI * d.target * HIJU_SILICONE / (d.CHARGEWEIGHT - d.TOPWEIGHT)
'
'    ResCalculation = Ro / (1 - Gx) ^ (d.COEFFICIENT - 1)
'    On Error GoTo 0
'    Exit Function
'Err:
'    On Error GoTo 0
'    ResCalculation = -9999
'End Function

'------------------------------------------------
' TBCME050データ取得
'------------------------------------------------

'概要      :テーブル「TBCME050」から指定品番のレコードを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tHin          ,I  ,tFullHinban                          :品番
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :抽出レコード
'    　　  :sErrMsg 　　  ,O  ,String     　　　　　　　　　　　    :エラーメッセージ
'          :戻り値        ,O  ,FUNCTION_RETURN                      :抽出の成否
'説明      :
'履歴      :2006/08/15 新規作成 エピ先行評価追加対応 SMP)kondoh

Public Function funGet_TBCME050(tHIN As tFullHinban, _
                                tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                                Optional sErrMsg As String = vbNullString) As FUNCTION_RETURN

    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet
    Dim sDBName     As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_GetSiyou.bas -- Function funGet_TBCME050"

    sDBName = "E050"
    '品EPBMD3平均下限(外周),平均上限(外周)追加　09/05/07 ooba
    sql = "SELECT hinban, mnorevno, factory, opecond, hepantnp"
    sql = sql & " ,hepof1ax ,hepof1mx ,hepof1et ,hepof1ns ,hepof1sz ,hepof1sh ,hepof1st ,hepof1sr ,hepof1ht ,hepof1hs "
    sql = sql & " ,hepof1km ,hepof1kn ,hepof1kh ,hepof1ku ,heposf1ptk"
    sql = sql & " ,hepof2ax ,hepof2mx ,hepof2et ,hepof2ns ,hepof2sz ,hepof2sh ,hepof2st ,hepof2sr ,hepof2ht ,hepof2hs"
    sql = sql & " ,hepof2km ,hepof2kn ,hepof2kh ,hepof2ku ,heposf2ptk"
    sql = sql & " ,hepof3ax ,hepof3mx ,hepof3et ,hepof3ns ,hepof3sz ,hepof3sh ,hepof3st ,hepof3sr ,hepof3ht ,hepof3hs"
    sql = sql & " ,hepof3km ,hepof3kn ,hepof3kh ,hepof3ku ,heposf3ptk"
    sql = sql & " ,hepbm1an ,hepbm1ax ,hepbm1et ,hepbm1ns ,hepbm1sz ,hepbm1sh ,hepbm1st ,hepbm1sr ,hepbm1ht ,hepbm1hs"
    sql = sql & " ,hepbm1km ,hepbm1kn ,hepbm1kh ,hepbm1ku ,hepbm1mbp ,hepbm1mcl"
    sql = sql & " ,hepbm2an ,hepbm2ax ,hepbm2et ,hepbm2ns ,hepbm2sz ,hepbm2sh ,hepbm2st ,hepbm2sr ,hepbm2ht ,hepbm2hs"
    sql = sql & " ,hepbm2km ,hepbm2kn ,hepbm2kh ,hepbm2ku ,hepbm2mbp ,hepbm2mcl"
    sql = sql & " ,hepbm3an ,hepbm3ax ,hepbm3gsan ,hepbm3gsax ,hepbm3et ,hepbm3ns ,hepbm3sz ,hepbm3sh ,hepbm3st ,hepbm3sr ,hepbm3ht ,hepbm3hs"
    sql = sql & " ,hepbm3km ,hepbm3kn ,hepbm3kh ,hepbm3ku ,hepbm3mbp ,hepbm3mcl"
    sql = sql & " FROM tbcme050 "
    sql = sql & " WHERE hinban = '" & tHIN.hinban & "' and "
    sql = sql & "      mnorevno = " & tHIN.mnorevno & " and "
    sql = sql & "      factory = '" & tHIN.factory & "' and "
    sql = sql & "      opecond = '" & tHIN.opecond & "'"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funGet_TBCME050 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''抽出結果を格納する
     With tGetRec
        .HEPANTNP = fncNullCheck(rs("HEPANTNP"))                            ' 品EPAN温度
        .HEPOF1AX = fncNullCheck(rs("HEPOF1AX"))                            ' 品EPOSF1平均上限
        .HEPOF1MX = fncNullCheck(rs("HEPOF1MX"))                            ' 品EPOSF1上限
        .HEPOF1ET = fncNullCheck(rs("HEPOF1ET"))                            ' 品EPOSF1選択ET代
        .HEPOF1NS = IIf(IsNull(rs("HEPOF1NS")), "", rs("HEPOF1NS"))         ' 品EPOSF1熱処理法
        .HEPOF1SZ = IIf(IsNull(rs("HEPOF1SZ")), "", rs("HEPOF1SZ"))         ' 品EPOSF1測定条件
        .HEPOF1SH = IIf(IsNull(rs("HEPOF1SH")), "", rs("HEPOF1SH"))         ' 品EPOSF1測定位置_方
        .HEPOF1ST = IIf(IsNull(rs("HEPOF1ST")), "", rs("HEPOF1ST"))         ' 品EPOSF1測定位置_点
        .HEPOF1SR = IIf(IsNull(rs("HEPOF1SR")), "", rs("HEPOF1SR"))         ' 品EPOSF1測定位置_領
        .HEPOF1HT = IIf(IsNull(rs("HEPOF1HT")), "", rs("HEPOF1HT"))         ' 品EPOSF1保証方法_対
        .HEPOF1HS = IIf(IsNull(rs("HEPOF1HS")), "", rs("HEPOF1HS"))         ' 品EPOSF1保証方法_処
        .HEPOF1KM = IIf(IsNull(rs("HEPOF1KM")), "", rs("HEPOF1KM"))         ' 品EPOSF1検査頻度_枚
        .HEPOF1KN = IIf(IsNull(rs("HEPOF1KN")), "", rs("HEPOF1KN"))         ' 品EPOSF1検査頻度_抜
        .HEPOF1KH = IIf(IsNull(rs("HEPOF1KH")), "", rs("HEPOF1KH"))         ' 品EPOSF1検査頻度_保
        .HEPOF1KU = IIf(IsNull(rs("HEPOF1KU")), "", rs("HEPOF1KU"))         ' 品EPOSF1検査頻度_ｳ
        .HEPOSF1PTK = IIf(IsNull(rs("HEPOSF1PTK")), "", rs("HEPOSF1PTK"))   ' 品EPOSF1ﾊﾟﾀﾝ区分
        .HEPOF2AX = fncNullCheck(rs("HEPOF2AX"))                            ' 品EPOSF2平均上限
        .HEPOF2MX = fncNullCheck(rs("HEPOF2MX"))                            ' 品EPOSF2上限
        .HEPOF2ET = fncNullCheck(rs("HEPOF2ET"))                            ' 品EPOSF2選択ET代
        .HEPOF2NS = IIf(IsNull(rs("HEPOF2NS")), "", rs("HEPOF2NS"))         ' 品EPOSF2熱処理法
        .HEPOF2SZ = IIf(IsNull(rs("HEPOF2SZ")), "", rs("HEPOF2SZ"))         ' 品EPOSF2測定条件
        .HEPOF2SH = IIf(IsNull(rs("HEPOF2SH")), "", rs("HEPOF2SH"))         ' 品EPOSF2測定位置_方
        .HEPOF2ST = IIf(IsNull(rs("HEPOF2ST")), "", rs("HEPOF2ST"))         ' 品EPOSF2測定位置_点
        .HEPOF2SR = IIf(IsNull(rs("HEPOF2SR")), "", rs("HEPOF2SR"))         ' 品EPOSF2測定位置_領
        .HEPOF2HT = IIf(IsNull(rs("HEPOF2HT")), "", rs("HEPOF2HT"))         ' 品EPOSF2保証方法_対
        .HEPOF2HS = IIf(IsNull(rs("HEPOF2HS")), "", rs("HEPOF2HS"))         ' 品EPOSF2保証方法_処
        .HEPOF2KM = IIf(IsNull(rs("HEPOF2KM")), "", rs("HEPOF2KM"))         ' 品EPOSF2検査頻度_枚
        .HEPOF2KN = IIf(IsNull(rs("HEPOF2KN")), "", rs("HEPOF2KN"))         ' 品EPOSF2検査頻度_抜
        .HEPOF2KH = IIf(IsNull(rs("HEPOF2KH")), "", rs("HEPOF2KH"))         ' 品EPOSF2検査頻度_保
        .HEPOF2KU = IIf(IsNull(rs("HEPOF2KU")), "", rs("HEPOF2KU"))         ' 品EPOSF2検査頻度_ｳ
        .HEPOSF2PTK = IIf(IsNull(rs("HEPOSF2PTK")), "", rs("HEPOSF2PTK"))   ' 品EPOSF2ﾊﾟﾀﾝ区分
        .HEPOF3AX = fncNullCheck(rs("HEPOF3AX"))                            ' 品EPOSF3平均上限
        .HEPOF3MX = fncNullCheck(rs("HEPOF3MX"))                            ' 品EPOSF3上限
        .HEPOF3ET = fncNullCheck(rs("HEPOF3ET"))                            ' 品EPOSF3選択ET代
        .HEPOF3NS = IIf(IsNull(rs("HEPOF3NS")), "", rs("HEPOF3NS"))         ' 品EPOSF3熱処理法
        .HEPOF3SZ = IIf(IsNull(rs("HEPOF3SZ")), "", rs("HEPOF3SZ"))         ' 品EPOSF3測定条件
        .HEPOF3SH = IIf(IsNull(rs("HEPOF3SH")), "", rs("HEPOF3SH"))         ' 品EPOSF3測定位置_方
        .HEPOF3ST = IIf(IsNull(rs("HEPOF3ST")), "", rs("HEPOF3ST"))         ' 品EPOSF3測定位置_点
        .HEPOF3SR = IIf(IsNull(rs("HEPOF3SR")), "", rs("HEPOF3SR"))         ' 品EPOSF3測定位置_領
        .HEPOF3HT = IIf(IsNull(rs("HEPOF3HT")), "", rs("HEPOF3HT"))         ' 品EPOSF3保証方法_対
        .HEPOF3HS = IIf(IsNull(rs("HEPOF3HS")), "", rs("HEPOF3HS"))         ' 品EPOSF3保証方法_処
        .HEPOF3KM = IIf(IsNull(rs("HEPOF3KM")), "", rs("HEPOF3KM"))         ' 品EPOSF3検査頻度_枚
        .HEPOF3KN = IIf(IsNull(rs("HEPOF3KN")), "", rs("HEPOF3KN"))         ' 品EPOSF3検査頻度_抜
        .HEPOF3KH = IIf(IsNull(rs("HEPOF3KH")), "", rs("HEPOF3KH"))         ' 品EPOSF3検査頻度_保
        .HEPOF3KU = IIf(IsNull(rs("HEPOF3KU")), "", rs("HEPOF3KU"))         ' 品EPOSF3検査頻度_ｳ
        .HEPOSF3PTK = IIf(IsNull(rs("HEPOSF3PTK")), "", rs("HEPOSF3PTK"))   ' 品EPOSF3ﾊﾟﾀﾝ区分
        .HEPBM1AN = fncNullCheck(rs("HEPBM1AN"))                            ' 品EPBMD1平均下限
        .HEPBM1AX = fncNullCheck(rs("HEPBM1AX"))                            ' 品EPBMD1平均上限
        .HEPBM1ET = fncNullCheck(rs("HEPBM1ET"))                            ' 品EPBMD1選択ET代
        .HEPBM1NS = IIf(IsNull(rs("HEPBM1NS")), "", rs("HEPBM1NS"))         ' 品EPBMD1熱処理法
        .HEPBM1SZ = IIf(IsNull(rs("HEPBM1SZ")), "", rs("HEPBM1SZ"))         ' 品EPBMD1測定条件
        .HEPBM1SH = IIf(IsNull(rs("HEPBM1SH")), "", rs("HEPBM1SH"))         ' 品EPBMD1測定位置_方
        .HEPBM1ST = IIf(IsNull(rs("HEPBM1ST")), "", rs("HEPBM1ST"))         ' 品EPBMD1測定位置_点
        .HEPBM1SR = IIf(IsNull(rs("HEPBM1SR")), "", rs("HEPBM1SR"))         ' 品EPBMD1測定位置_領
        .HEPBM1HT = IIf(IsNull(rs("HEPBM1HT")), "", rs("HEPBM1HT"))         ' 品EPBMD1保証方法_対
        .HEPBM1HS = IIf(IsNull(rs("HEPBM1HS")), "", rs("HEPBM1HS"))         ' 品EPBMD1保証方法_処
        .HEPBM1KM = IIf(IsNull(rs("HEPBM1KM")), "", rs("HEPBM1KM"))         ' 品EPBMD1検査頻度_枚
        .HEPBM1KN = IIf(IsNull(rs("HEPBM1KN")), "", rs("HEPBM1KN"))         ' 品EPBMD1検査頻度_抜
        .HEPBM1KH = IIf(IsNull(rs("HEPBM1KH")), "", rs("HEPBM1KH"))         ' 品EPBMD1検査頻度_保
        .HEPBM1KU = IIf(IsNull(rs("HEPBM1KU")), "", rs("HEPBM1KU"))         ' 品EPBMD1検査頻度_ｳ
        .HEPBM1MBP = fncNullCheck(rs("HEPBM1MBP"))                          ' 品EPBMD1面内分布
        .HEPBM1MCL = IIf(IsNull(rs("HEPBM1MCL")), "", rs("HEPBM1MCL"))      ' 品EPBMD1面内計算
        .HEPBM2AN = fncNullCheck(rs("HEPBM2AN"))                            ' 品EPBMD2平均下限
        .HEPBM2AX = fncNullCheck(rs("HEPBM2AX"))                            ' 品EPBMD2平均上限
        .HEPBM2ET = fncNullCheck(rs("HEPBM2ET"))                            ' 品EPBMD2選択ET代
        .HEPBM2NS = IIf(IsNull(rs("HEPBM2NS")), "", rs("HEPBM2NS"))         ' 品EPBMD2熱処理法
        .HEPBM2SZ = IIf(IsNull(rs("HEPBM2SZ")), "", rs("HEPBM2SZ"))         ' 品EPBMD2測定条件
        .HEPBM2SH = IIf(IsNull(rs("HEPBM2SH")), "", rs("HEPBM2SH"))         ' 品EPBMD2測定位置_方
        .HEPBM2ST = IIf(IsNull(rs("HEPBM2ST")), "", rs("HEPBM2ST"))         ' 品EPBMD2測定位置_点
        .HEPBM2SR = IIf(IsNull(rs("HEPBM2SR")), "", rs("HEPBM2SR"))         ' 品EPBMD2測定位置_領
        .HEPBM2HT = IIf(IsNull(rs("HEPBM2HT")), "", rs("HEPBM2HT"))         ' 品EPBMD2保証方法_対
        .HEPBM2HS = IIf(IsNull(rs("HEPBM2HS")), "", rs("HEPBM2HS"))         ' 品EPBMD2保証方法_処
        .HEPBM2KM = IIf(IsNull(rs("HEPBM2KM")), "", rs("HEPBM2KM"))         ' 品EPBMD2検査頻度_枚
        .HEPBM2KN = IIf(IsNull(rs("HEPBM2KN")), "", rs("HEPBM2KN"))         ' 品EPBMD2検査頻度_抜
        .HEPBM2KH = IIf(IsNull(rs("HEPBM2KH")), "", rs("HEPBM2KH"))         ' 品EPBMD2検査頻度_保
        .HEPBM2KU = IIf(IsNull(rs("HEPBM2KU")), "", rs("HEPBM2KU"))         ' 品EPBMD2検査頻度_ｳ
        .HEPBM2MBP = fncNullCheck(rs("HEPBM2MBP"))                          ' 品EPBMD2面内分布
        .HEPBM2MCL = IIf(IsNull(rs("HEPBM2MCL")), "", rs("HEPBM2MCL"))      ' 品EPBMD2面内計算
        .HEPBM3AN = fncNullCheck(rs("HEPBM3AN"))                            ' 品EPBMD3平均下限
        .HEPBM3AX = fncNullCheck(rs("HEPBM3AX"))                            ' 品EPBMD3平均上限
        .HEPBM3GSAN = fncNullCheck(rs("HEPBM3GSAN"))                        ' 品EPBMD3平均下限(外周)　09/05/07 ooba
        .HEPBM3GSAX = fncNullCheck(rs("HEPBM3GSAX"))                        ' 品EPBMD3平均上限(外周)　09/05/07 ooba
        .HEPBM3ET = fncNullCheck(rs("HEPBM3ET"))                            ' 品EPBMD3選択ET代
        .HEPBM3NS = IIf(IsNull(rs("HEPBM3NS")), "", rs("HEPBM3NS"))         ' 品EPBMD3熱処理法
        .HEPBM3SZ = IIf(IsNull(rs("HEPBM3SZ")), "", rs("HEPBM3SZ"))         ' 品EPBMD3測定条件
        .HEPBM3SH = IIf(IsNull(rs("HEPBM3SH")), "", rs("HEPBM3SH"))         ' 品EPBMD3測定位置_方
        .HEPBM3ST = IIf(IsNull(rs("HEPBM3ST")), "", rs("HEPBM3ST"))         ' 品EPBMD3測定位置_点
        .HEPBM3SR = IIf(IsNull(rs("HEPBM3SR")), "", rs("HEPBM3SR"))         ' 品EPBMD3測定位置_領
        .HEPBM3HT = IIf(IsNull(rs("HEPBM3HT")), "", rs("HEPBM3HT"))         ' 品EPBMD3保証方法_対
        .HEPBM3HS = IIf(IsNull(rs("HEPBM3HS")), "", rs("HEPBM3HS"))         ' 品EPBMD3保証方法_処
        .HEPBM3KM = IIf(IsNull(rs("HEPBM3KM")), "", rs("HEPBM3KM"))         ' 品EPBMD3検査頻度_枚
        .HEPBM3KN = IIf(IsNull(rs("HEPBM3KN")), "", rs("HEPBM3KN"))         ' 品EPBMD3検査頻度_抜
        .HEPBM3KH = IIf(IsNull(rs("HEPBM3KH")), "", rs("HEPBM3KH"))         ' 品EPBMD3検査頻度_保
        .HEPBM3KU = IIf(IsNull(rs("HEPBM3KU")), "", rs("HEPBM3KU"))         ' 品EPBMD3検査頻度_ｳ
        .HEPBM3MBP = fncNullCheck(rs("HEPBM3MBP"))                          ' 品EPBMD3面内分布
        .HEPBM3MCL = IIf(IsNull(rs("HEPBM3MCL")), "", rs("HEPBM3MCL"))      ' 品EPBMD3面内計算
    End With
    Set rs = Nothing

    funGet_TBCME050 = FUNCTION_RETURN_SUCCESS
  
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
