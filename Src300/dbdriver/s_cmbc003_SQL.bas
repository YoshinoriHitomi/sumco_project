Attribute VB_Name = "s_cmbc003_SQL"
Option Explicit

#If False Then '---------------- 参考
' TBCME017 (製品仕様管理)より
Public Type s_cmzcF_cmfc001a_Disp
    '製品仕様管理
    hinban As String * 8            ' 品番
    MNOREVNO As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    REGDATE As Date                 ' 登録日付
End Type
#End If '----------------------- ここまで

' 払出規制項目追加対応 yakimura 2002.12.01 start
Public Type s_cmzcF_cmfc002d_Disp
    Hinban12 As String * 12          ' 品番+Revision+工場識別+操業条件
    EPDSETCH As String * 1           ' EPD　選択エッチ
    EPDUP As Integer                 ' EPD　上限
    CUTUNIT As Integer               ' カット単位
    TOPREG As Integer                ' TOP規制
    TAILREG As Double                ' TAIL規制
    BTMSPRT As Integer               ' ボトム析出規制
    REGDATE As Date                  ' 登録日付
    UPDDATE As Date                  ' 更新日付
End Type
' 払出規制項目追加対応 yakimura 2002.12.01 start
Public Type DispData
    KDataName As String
    HDataName As String
    JDataName As String
    KDataNameWf As String
    HDataNameWf As String
End Type
Public DataName() As DispData
' 汎用ｺｰﾄﾞﾏｽﾀ
Public Type typ_GPCodeMaster_003
    codeNo As String            ' コードＮＯ
    CODE As String              ' コード
    codeCont As String          ' コード内容
    INDORDER As Long            ' 表示順
    codename As String          ' コード名称
    KUBUN As String             ' 区分
    READTIME As Double          ' リードタイム
    INDKBN As String            ' 表示区分 0=表示,1=非表示
End Type

'概要      :SQLを元に、データを取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :TbName        ,I  ,String    ,テーブル名
'          :sql           ,I  ,String    ,SQL
'          :rec           ,O  ,c_cmzcrec ,取得データ格納先
'          :戻り値        ,O  ,FUNCTION_RETURN,検索の成否
'説明      :
'履歴      :2001/06/08 作成  野村
Private Function DispSXL_GetData(TbName$, sql$, rec As c_cmzcrec) As FUNCTION_RETURN
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001d_SQL.bas -- Function DispSXL_GetData"
    
    '' データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If (rs Is Nothing) Or (rs.RecordCount = 0) Then
        DispSXL_GetData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    '' 抽出結果を格納する
    Set rec = New c_cmzcrec
    rec.CopyFromRs TbName, rs
    rs.Close
    DispSXL_GetData = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :製品仕様入力画面用 SXLタブに表示する内容を得る
'ﾊﾟﾗﾒｰﾀ    :変数名         ,IO ,型          ,説明
'          :targetHinban   ,   ,tFullHinban ,品番情報
'          :SxlKokyaku_1() ,   ,c_cmzcrec   ,顧客仕様SXL 1 の内容
'          :SxlKokyaku_2() ,   ,c_cmzcrec   ,顧客仕様SXL 2 の内容
'          :SxlKokyaku_3() ,   ,c_cmzcrec   ,顧客仕様SXL 3 の内容
'          :Sxl_1()        ,   ,c_cmzcrec   ,製品仕様SXL_1 の内容
'          :Sxl_2()        ,   ,c_cmzcrec   ,製品仕様SXL_2 の内容
'          :Sxl_3()        ,   ,c_cmzcrec   ,製品仕様SXL_3 の内容
'          :WfKokyaku_2()  ,   ,c_cmzcrec   ,顧客仕様WF 2 の内容
'          :WfKokyaku_8()  ,   ,c_cmzcrec   ,顧客仕様WF 8 の内容
'          :SxlUchigawa()  ,   ,c_cmzcrec   ,内側 の内容
'          :errTbl         ,O  ,String      ,取得できなかったテーブル名
'          :戻り値         ,O  ,FUNCTION_RETURN,検索の成否
'説明      :各出力パラメータの配列は、(1)に仕様データ (2)に指定操業条件のデータ が入る
'          :呼出側で品番を12桁入力可能なため、該当品番が存在しない場合がありうる
'履歴      :2001/06/08 作成  野村
Public Function DBDRV_s_cmzcF_cmgc001d_DispSXL(targetHinban As tFullHinban, SxlKokyaku_1() As c_cmzcrec, _
                                               SxlKokyaku_2() As c_cmzcrec, SxlKokyaku_3() As c_cmzcrec, _
                                               Sxl_1() As c_cmzcrec, Sxl_2() As c_cmzcrec, Sxl_3() As c_cmzcrec, _
                                               WfKokyaku_2() As c_cmzcrec, WfKokyaku_8() As c_cmzcrec, _
                                               Sxluchigawa() As c_cmzcrec, TBCME003 As c_cmzcrec, errTbl$) As FUNCTION_RETURN
Dim sql As String
Dim sqlBase As String       'SQLの基本部
Dim sqlWhere(2) As String   'Where句以降
Dim TbName As String        'テーブル名
Dim i As Integer
Dim HWFMKnSI(4) As Double   'LPDサイズ(0:結果 1～4:候補)
Dim HWFMKnMX(4) As Integer  'LPD上限(0:結果 1～4:候補)
Dim rs As OraDynaset
        
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001d_SQL.bas -- Function DBDRV_s_cmzcF_cmgc001d_DispSXL"
    
    DBDRV_s_cmzcF_cmgc001d_DispSXL = FUNCTION_RETURN_FAILURE
    errTbl = vbNullString
    
    ''SQLの共通部分を準備する
    With targetHinban
        '仕様のWHERE句
        sqlWhere(1) = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='1')" & _
              " Order by OPECOND DESC"
        '内側管理のWHERE句
        sqlWhere(2) = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='" & .OPECOND & "')" & _
              " Order by OPECOND DESC"
    End With
    
    ''出力データを初期化する
    For i = 1 To 2
        Set SxlKokyaku_1(i) = New c_cmzcrec
        Set SxlKokyaku_2(i) = New c_cmzcrec
        Set SxlKokyaku_3(i) = New c_cmzcrec
        Set Sxl_1(i) = New c_cmzcrec
        Set Sxl_2(i) = New c_cmzcrec
        Set Sxl_3(i) = New c_cmzcrec
        Set WfKokyaku_2(i) = New c_cmzcrec
        Set WfKokyaku_8(i) = New c_cmzcrec
        Set Sxluchigawa(i) = New c_cmzcrec
    Next
    Sxluchigawa(1).TABLENAME = "TBCME036"
    Sxluchigawa(1).SetRecDefault
    
    ''無効品番のチェック（製品仕様SXL1に登録されていること、製作条件付与取消に登録されていないこと）
    With targetHinban
        sql = "select A.HINBAN from TBCME018 A, TBCME031 B " & _
              "where (A.HINBAN='" & .hinban & "') and (A.MNOREVNO=" & .MNOREVNO & ") and (A.FACTORY='" & .FACTORY & "')" & _
              " and (A.HINBAN=B.HINBAN(+)) and (A.MNOREVNO=B.MNOREVNO(+)) and (A.FACTORY=B.FACTORY(+))" & _
              " and (B.HINBAN is null)"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If (rs Is Nothing) Or (rs.RecordCount = 0) Then
            errTbl = "TBCME031"
            GoTo proc_exit
        End If
    End With
    
    ''1. 顧客SXL仕様_1 (TBCME005) の内容を取得する
    ''1-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME005"
    sqlBase = "Select KSXTYPKB, KSXRUNIT, KSXRKKBN, KSXDUNIT, KSXD1KBN, KSXD2KBN," & _
                " KSXDFKBN, KSXDPKBN, KSXDWKBN, KSXDDKBN, KSXDAKBN, " & _
                " KSXRMIN, KSXRMAX, KSXRSPOH, KSXRSPOT,KSXRSPOI,KSXRHWYT,KSXRHWYS,KSXRKWAY, " & _
                " KSXRKHNM,KSXRKHNI,KSXRKHNH,KSXRKHNS,KSXRSDEV,KSXRMBNP,KSXRMCAL,KSXTYPE,KSXTYPKW, KSXTYPKB,KSXDOP "
    sqlBase = sqlBase & "From " & TbName
    ''1-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), SxlKokyaku_1(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    '内側管理レコード
    '顧客仕様は内側管理しないため、仕様レコードを取得する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), SxlKokyaku_1(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''2. 顧客SXL仕様_2 (TBCME006) の内容を取得する
    ''2-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME006"
    sqlBase = "Select KSXTMKBN, KSXLTUNT, KSXLTKBN, KSXCNIND, KSXCNUNT," & _
                " KSXCNKBN, KSXONIND, KSXONUNT, KSXONKBN "
    sqlBase = sqlBase & ", KSXLTMIN, KSXLTMAX, KSXLTSPH, KSXLTSPT, KSXLTSPI, KSXLTHWT, KSXLTHWS, KSXLTKWY "
    sqlBase = sqlBase & ", KSXONMIN, KSXONMAX, KSXONSPH, KSXONSPT, KSXONSPI, KSXONHWT, KSXONHWS, KSXONKWY "
    sqlBase = sqlBase & ", KSXONKHM, KSXONKHI, KSXONKHH, KSXONKHS, KSXONSDV, KSXONMBP, KSXONMCL "
    sqlBase = sqlBase & ", KSXCNMIN, KSXCNMAX, KSXCNSPH, KSXCNSPT, KSXCNSPI, KSXCNHWT, KSXCNHWS, KSXCNKWY "
    sqlBase = sqlBase & ", KSXTMMAX, KSXTMSPH, KSXTMSPT, KSXTMSPR, KSXTMKHM, KSXTMKHI, KSXTMKHH, KSXTMKHS "
    sqlBase = sqlBase & "From " & TbName
    ''2-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), SxlKokyaku_2(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    '内側管理レコード
    '顧客仕様は内側管理しないため、仕様レコードを取得する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), SxlKokyaku_2(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''3. 顧客SXL仕様_3 (TBCME007) の内容を取得する
    ''3-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME007"
    sqlBase = "Select KSXOF1KBN, KSXOF1FGS, KSXOF1SO1, KSXOF1ST1, KSXOF2KB, KSXOF2GS, " & _
                "KSXOF2O1, KSXOF2T1, KSXBMKBN, KSXBMFGS, KSXBM2KB, KSXBM2GS "
    sqlBase = sqlBase & ", KSXDENMX, KSXDENMN, KSXDENHT, KSXDENHS, KSXDENKU,KSXLDLMN, KSXLDLMX, " & _
                "KSXLDLHT, KSXLDLHS, KSXLDLKU,KSXDVDMNN, KSXDVDMXN, KSXDVDHT, KSXDVDHS, KSXDVDKU "
    sqlBase = sqlBase & ", KSXOF1MAX, KSXOF1AMX, KSXOF1NSW, KSXOF1CET, KSXOF1SPH, KSXOF1SPT, " & _
                "KSXOF1SPR, KSXOF1HWT, KSXOF1HWS, KSXOF1KHM, KSXOF1KHH, KSXOFPTK, KSXOF1SZY "
    sqlBase = sqlBase & ", KSXOF2MX, KSXOF2AX, KSXOF2NS, KSXOF2ET, KSXOF2SH, KSXOF2ST, " & _
                "KSXOF2SR, KSXOF2HT, KSXOF2HS, KSXOF2KM, KSXOF2KH, KSXOF2PTK, KSXOF2SZ  "
    sqlBase = sqlBase & ", KSXBMMIN, KSXBMMAX, KSXBMCET, KSXBMNS, KSXBMSPH, KSXBMSPT, KSXBMSZY, " & _
                "KSXBMSPR, KSXBMHWT, KSXBMHWS, KSXBMKHM, KSXBMKHI, KSXBMKHH, KSXBMKHS, KSXBM1MBNP, KSXBM1MCAL "
    sqlBase = sqlBase & ", KSXBM2AN, KSXBM2AX, KSXBM2ET, KSXBM2NS, KSXBM2SH, KSXBM2ST, KSXBM2SZ, " & _
                "KSXBM2SR, KSXBM2HT, KSXBM2HS, KSXBM2KM, KSXBM2KI, KSXBM2KH, KSXBM2KS, KSXBM2MBNP, KSXBM2MCAL "
    sqlBase = sqlBase & ", KSXGDSPH, KSXGDSPT, KSXGDSPR, KSXGDSZY, KSXGDZAR, KSXGDKHM, " & _
                "KSXGDKHI, KSXGDKHH, KSXGDKHS,KSXDSOMX,KSXDSOMN,KSXDSOAX,KSXDSOAN,KSXDSOHT, " & _
                "KSXDSOHS,KSXDSOKM,KSXDSOKI,KSXDSOKH,KSXDSOKS "
    sqlBase = sqlBase & "From " & TbName
    ''3-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), SxlKokyaku_3(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    '内側管理レコード
    '顧客仕様は内側管理しないため、仕様レコードを取得する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), SxlKokyaku_3(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''4. 製品SXL仕様_1 (TBCME018) の内容を取得する
    ''4-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME018"
    sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND," & _
                " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH," & _
                " HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
                " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2," & _
                " HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM, HSXD1CEN," & _
                " HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR," & _
                " HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY, HSXCKHNM, HSXCKHNI," & _
                " HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN," & _
                " HSXCTMIN, HSXCTMAX, HSXCYDIR, HSXCYCEN, HSXCYMIN, HSXCYMAX," & _
                " HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY," & _
                " HSXDPDIR, HSXDPMIN, HSXDPMAX, HSXDWCEN, HSXDWMIN, HSXDWMAX," & _
                " HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX," & _
                " SPECRRNO, SXLMCNO, WFMCNO, MCNO, REGDATE, UPDDATE "
    sqlBase = sqlBase & "From " & TbName
    ''4-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), Sxl_1(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    '内側管理レコード
    If DispSXL_GetData(TbName, "select * from TBCME018 " & sqlWhere(2), Sxl_1(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''5. 製品SXL仕様_2 (TBCME019) の内容を取得する
    ''5-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME019"
    sqlBase = "Select HSXTMMAX, HSXTMSPH, HSXTMSPT, HSXTMSPR, HSXTMKHM," & _
                " HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH," & _
                " HSXLTSPT, HSXLTSPI, HSXLTHWT, HSXLTHWS, HSXLTNSW, HSXLTKHM, HSXLTKWY," & _
                " HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN," & _
                " HSXCNMAX, HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS," & _
                " HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
                " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS," & _
                " HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS, HSXONMBP," & _
                " HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX," & _
                " HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH, HSXOS1ST, HSXOS1SI," & _
                " HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS," & _
                " HSXOS2MN, HSXOS2MX, HSXOS2NS, HSXOS2SH, HSXOS2ST, HSXOS2SI," & _
                " HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, HSXTMMAXN "
    sqlBase = sqlBase & "From " & TbName
    ''5-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), Sxl_2(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    '内側管理レコード
    If DispSXL_GetData(TbName, "select * from TBCME019 " & sqlWhere(2), Sxl_2(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''6. 製品SXL仕様_3 (TBCME020) の内容を取得する
    ''6-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME020"
    sqlBase = "Select HSXDENKU, HSXDENMX, HSXDENMN, HSXDENHT, HSXDENHS, HSXDENKU,HSXDVDKU, HSXDVDMX, HSXDVDMN, HSXDVDHT, HSXDVDHS, HSXDVDKU,HSXLDLKU," & _
                " HSXLDLMX, HSXLDLMN, HSXLDLHT, HSXLDLHS, HSXLDLKU,HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH," & _
                " HSXGDKHS, HSXDSOKE, HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS," & _
                " HSXLIFTW, HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
                " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH, HSXOF1KS," & _
                " HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ, HSXOF2KM, HSXOF2KI," & _
                " HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR, HSXOF3HT, HSXOF3HS, HSXOF3SZ," & _
                " HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX, HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT," & _
                " HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS, HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST," & _
                " HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI, HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX," & _
                " HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS, HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET," & _
                " HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST, HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS," & _
                " HSXBM3NS, HSXBM3ET, HSXNOTE, HSXRS1N, HSXRS1Y, HSXRS2N, HSXRS2Y, HSXRS3N, HSXRS3Y, HSXRS4N, HSXRS4Y, HSXRS5N, HSXRS5Y," & _
                " HSXRS6N, HSXRS6Y, HSXRS7N, HSXRS7Y, HSXRS8N, HSXRS8Y, HSXRS9N, HSXRS9Y, HSXRS10N, HSXRS10Y, " & _
                " HSXDVDMNN, HSXDVDMXN, HSXDSONS, HSXCDOPMN, HSXCDOPMX, HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK, " & _
                " HSXBMD1MBP, HSXBMD1MCL, HSXBMD2MBP, HSXBMD2MCL, HSXBMD3MBP, HSXBMD3MCL, HSXDSOPTK,HSXGDPTK,HSXEPNOTE, " & _
                " HSXCOSF3PK,HSXCOSF3SH,HSXCOSF3ST,HSXCOSF3SR,HSXCOSF3HT,HSXCOSF3HS,HSXCOSF3SZ,HSXCOSF3NS,HSXCPK,HSXCSZ, " & _
                " HSXCHT,HSXCHS,HSXCJPK,HSXCJNS,HSXCJHT,HSXCJHS,HSXCJLTPK,HSXCJLTNS,HSXCJLTHT,HSXCJLTHS,HSXCJ2PK,HSXCJ2NS," & _
                " HSXCJ2HT,HSXCJ2HS,HSXDSOSZ "

    sqlBase = sqlBase & "From " & TbName
    ''6-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), Sxl_3(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    '内側管理レコード
    If DispSXL_GetData(TbName, "select * from TBCME020 " & sqlWhere(2), Sxl_3(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''7. 顧客仕様WFﾃﾞｰﾀ2 (TBCME009) の内容を取得する
    ''7-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME009"
    sqlBase = "Select KPRDFORM "
    sqlBase = sqlBase & "From " & TbName
    ''7-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), WfKokyaku_2(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    '内側管理レコード
    '顧客仕様は内側管理しないため、仕様レコードを取得する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), WfKokyaku_2(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''8. 顧客仕様WFﾃﾞｰﾀ8 (TBCME028) の内容を取得する
    ''8-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME028"
    sqlBase = "Select HWFMK1SI, HWFMK2SI, HWFMK3SI, HWFMK4SI, HWFMK1MX, HWFMK2MX, HWFMK3MX, HWFMK4MX "
    sqlBase = sqlBase & "From " & TbName
    ''8-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), WfKokyaku_8(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    '内側管理レコード
    '顧客仕様は内側管理しないため、仕様レコードを取得する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), WfKokyaku_8(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    ''9. 結晶内側管理 (TBCME036) の内容を取得する
    ''9-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
' 払出規制項目追加対応 yakimura 2002.12.01 start
    TbName = "TBCME036"
    'sqlBase = "Select EPDSETCH, EPDUP, CUTUNIT, NVL(TOPREG,0) TOPREG, NVL(TAILREG,0) TAILREG, NVL(BTMSPRT,0) BTMSPRT "
    'sqlBase = sqlBase & ", OTHER1 , OTHER2, OTHERTIME, DCHYUUBU, SNOTE, JNOTE, BLOCKHFLAG, MCNO, SSXLIFTW "
    'sqlBase = sqlBase & ", SPECRRNO, OTHER1MAI, OTHER2MAI, WFCUTT, HSXGDLINE, HWFGDLINE, GLASS, SLICEATU, KUMIDOP, OTHERTIME2 "
    'sqlBase = sqlBase & ", SKPLACE, COSF3FLAG, "
    sqlBase = "Select * "
    sqlBase = sqlBase & "From " & TbName
' 払出規制項目追加対応 yakimura 2002.12.01 end
    ''9-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, "select * from TBCME036 " & sqlWhere(1), Sxluchigawa(1)) = FUNCTION_RETURN_FAILURE Then
        '結晶内側管理Tblは該当レコードなしでもOK。既定値を設定する
        Sxluchigawa(1).TABLENAME = TbName
        Sxluchigawa(1).SetRecDefault
    End If
    '内側管理レコード
    If DispSXL_GetData(TbName, "select * from TBCME036 " & sqlWhere(2), Sxluchigawa(2)) = FUNCTION_RETURN_FAILURE Then
        Sxluchigawa(2).TABLENAME = TbName
        Sxluchigawa(2).SetRecDefault
    End If
    'If DispSXL_GetData(TbName, sqlBase & sqlWhere(2), Sxluchigawa(2)) = FUNCTION_RETURN_FAILURE Then
       '結晶内側管理Tblは該当レコードなしでもOK。既定値を設定する
    '    Sxluchigawa(2).TABLENAME = TbName
    '    Sxluchigawa(2).SetRecDefault
    'End If
    
    '10. LPDサイズ・LPD上限の表示値を設定する
    With WfKokyaku_8(2)
        'HWFMK1SI～HWFMK4SI, HWFMK1MX～HWFMK4MX を取得 '03/12/10 Null対応
        'HWFMKnSI(1) = .Fields.GetValueOrDefault("HWFMK1SI", 0#)
        'HWFMKnMX(1) = .Fields.GetValueOrDefault("HWFMK1MX", 0)
        'HWFMKnSI(2) = .Fields.GetValueOrDefault("HWFMK2SI", 0#)
        'HWFMKnMX(2) = .Fields.GetValueOrDefault("HWFMK2MX", 0)
        'HWFMKnSI(3) = .Fields.GetValueOrDefault("HWFMK3SI", 0#)
        'HWFMKnMX(3) = .Fields.GetValueOrDefault("HWFMK3MX", 0)
        'HWFMKnSI(4) = .Fields.GetValueOrDefault("HWFMK4SI", 0#)
        'HWFMKnMX(4) = .Fields.GetValueOrDefault("HWFMK4MX", 0)
        If IsNull(.Fields("HWFMK1SI")) = False Then
            HWFMKnSI(1) = .Fields.GetValueOrDefault("HWFMK1SI", 0#)
        Else
            HWFMKnSI(1) = 0
        End If
        If IsNull(.Fields("HWFMK1MX")) = False Then
            HWFMKnMX(1) = .Fields.GetValueOrDefault("HWFMK1MX", 0)
        Else
            HWFMKnMX(1) = 0
        End If
        If IsNull(.Fields("HWFMK2SI")) = False Then
            HWFMKnSI(2) = .Fields.GetValueOrDefault("HWFMK2SI", 0)
        Else
            HWFMKnSI(2) = 0
        End If
        If IsNull(.Fields("HWFMK2MX")) = False Then
            HWFMKnMX(2) = .Fields.GetValueOrDefault("HWFMK2MX", 0)
        Else
            HWFMKnMX(2) = 0
        End If
        If IsNull(.Fields("HWFMK3SI")) = False Then
            HWFMKnSI(3) = .Fields.GetValueOrDefault("HWFMK3SI", 0)
        Else
            HWFMKnSI(3) = 0
        End If
        If IsNull(.Fields("HWFMK3MX")) = False Then
            HWFMKnMX(3) = .Fields.GetValueOrDefault("HWFMK3MX", 0)
        Else
            HWFMKnMX(3) = 0
        End If
        If IsNull(.Fields("HWFMK4SI")) = False Then
            HWFMKnSI(4) = .Fields.GetValueOrDefault("HWFMK4SI", 0)
        Else
            HWFMKnSI(4) = 0
        End If
        If IsNull(.Fields("HWFMK4MX")) = False Then
            HWFMKnMX(4) = .Fields.GetValueOrDefault("HWFMK4MX", 0)
        Else
            HWFMKnMX(4) = 0
        End If
    End With
    '候補から絞り込む
    HWFMKnSI(0) = 9999#   '充分大きな値
    For i = 1 To 4
        If HWFMKnSI(0) > HWFMKnSI(i) Then
            HWFMKnSI(0) = HWFMKnSI(i)
            HWFMKnMX(0) = HWFMKnMX(i)
        End If
    Next
    '得られた結果をWfKokyaku_8(1)に登録する
    WfKokyaku_8(1).Fields.Add "HWFMKnSI", HWFMKnSI(0), ORADB_DOUBLE, -1, 5, 3
    WfKokyaku_8(1).Fields.Add "HWFMKnMX", HWFMKnMX(0), ORADB_INTEGER, -1, -1, 4
    
    ''顧客仕様 (TBCME003) の内容を取得する
    ''SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME003"
    sqlBase = "select KTWFBK,KTWFPK,KTWFCK,KTWFASK,KTWFFEK,KTWFALK," & _
                    " KTWFNIK,KTWFCUK,KTWFCRK,KTWFNAK,KTWFZNK,KTWFCAK "
    sqlBase = sqlBase & "From " & TbName
    ''6-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere(1), TBCME003) = FUNCTION_RETURN_FAILURE Then
        'errTbl = TbName
        'GoTo proc_exit
    End If
    
    
    DBDRV_s_cmzcF_cmgc001d_DispSXL = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function
'概要      :製品仕様入力画面用 WFタブに表示する内容を得る
'ﾊﾟﾗﾒｰﾀ    :変数名         ,IO ,型          ,説明
'          :targetHinban   ,   ,tFullHinban ,品番情報
'          :WF()           ,   ,c_cmzcrec   ,製品仕様WF 1-9 の内容
'          :errTbl         ,O  ,String      ,取得できなかったテーブル名
'          :戻り値         ,O  ,FUNCTION_RETURN,検索の成否
'説明      :出力パラメータの配列は、(1 to 9)に製品仕様WFデータ1～9 が入る
'          :呼出側で品番を12桁入力可能なため、該当品番が存在しない場合がありうる
'履歴      :2001/09/28 作成  野村
'          :2010/01/21 顧客仕様取得追加(特記表示用)
Public Function DBDRV_s_cmzcF_cmgc001d_DispWF(targetHinban As tFullHinban, WF() As c_cmzcrec, TBCME001 As c_cmzcrec, errTbl$) As FUNCTION_RETURN
Dim sql As String
Dim sqlBase As String       'SQLの基本部
Dim sqlWhere As String      'Where句以降
Dim sqlWhereE1 As String     'Where句以降
Dim TbName As String        'テーブル名
Dim i As Integer
Dim rs As OraDynaset
        
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001d_SQL.bas -- Function DBDRV_s_cmzcF_cmgc001d_DispWF"
    
    DBDRV_s_cmzcF_cmgc001d_DispWF = FUNCTION_RETURN_FAILURE
    errTbl = vbNullString
    
    ''SQLの共通部分を準備する
    With targetHinban
        '仕様のWHERE句(WF仕様は内側管理しない)
        sqlWhere = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='" & .OPECOND & "')"
        sqlWhereE1 = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='1')"
    End With
    
    ''出力データを初期化する
    ReDim WF(1 To 12)
    For i = 1 To 12
        Set WF(i) = New c_cmzcrec
    Next
    Set TBCME001 = New c_cmzcrec
    
    ''無効品番のチェック（製品仕様SXL1に登録されていること、製作条件付与取消に登録されていないこと）
    With targetHinban
        sql = "select A.HINBAN from TBCME018 A, TBCME031 B " & _
              "where (A.HINBAN='" & .hinban & "') and (A.MNOREVNO=" & .MNOREVNO & ") and (A.FACTORY='" & .FACTORY & "')" & _
              " and (A.HINBAN=B.HINBAN(+)) and (A.MNOREVNO=B.MNOREVNO(+)) and (A.FACTORY=B.FACTORY(+))" & _
              " and (B.HINBAN is null)"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If (rs Is Nothing) Or (rs.RecordCount = 0) Then
            errTbl = "TBCME031"
            GoTo proc_exit
        End If
    End With
    
    
    ''1. 製品WF仕様_1 (TBCME021) の内容を取得する
    ''1-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME021"
    sqlBase = "select HWFFACES, HWFBACKS, HWFBDSWY, HWFTYPE, HWFTYPKW, HWFDOP, HWFFKBWK, HWFFKBWS, HWFRMIN, HWFRMAX, " & _
              "HWFRSPOH||HWFRSPOT||HWFRSPOI as HWFRSPO, " & _
              "HWFRHWYT||HWFRHWYS as HWFRHWY, " & _
              "HWFRKWAY, " & _
              "HWFRKHNM||HWFRKHNN||HWFRKHNH||HWFRKHNU as HWFRKHN, " & _
              "HWFRSDEV, HWFRAMIN, HWFRAMAX, HWFRMBNP, HWFRMCAL, HWFRMBP2, HWFRMCL2, " & _
              "HWFRKBSH||HWFRKBST||HWFRKBSI as HWFRKBS, " & _
              "HWFRKBHT||HWFRKBHS as HWFRKBH, " & _
              "HWFSTMAX, " & _
              "HWFSTSPH||HWFSTSPT||HWFSTSPI as HWFSTSP, " & _
              "HWFSTHWT||HWFSTHWS as HWFSTHW, " & _
              "HWFSTKWY, " & _
              "HWFSTKHM||HWFSTKHN||HWFSTKHH||HWFSTKHU as HWFSTKH, " & _
              "HWFACEN, HWFAMIN, HWFAMAX, " & _
              "HWFASPOH||HWFASPOT||HWFASPOI as HWFASPO, " & _
              "HWFAHWYT||HWFAHWYS as HWFAHWY, "
    sqlBase = sqlBase & "HWFAKWAY, " & _
              "HWFAKHNM||HWFAKHNN||HWFAKHNH||HWFAKHNU as HWFAKHN, " & _
              "HWFASDEV, HWFAAMIN, HWFAAMAX, HWFAMBNP, HWFAMCAL, HWFALTBP, HWFALTCL, HWFALTRA, HWFAMRAN, " & _
              "HWFAKBSH||HWFAKBST||HWFAKBSI as HWFAKBS, " & _
              "HWFAKBHT||HWFAKBHS as HWFAKBH, " & _
              "HWFDIVS, HWFWFORM, HWFD1CEN, HWFD1MIN, HWFD1MAX, HWFD2CEN, HWFD2MIN, HWFD2MAX, " & _
              "HWFDKHNM||HWFDKHNN||HWFDKHNH||HWFDKHNU as HWFDKHN, " & _
              "HWFLPMNP, HWFSGMNP, HWFETMNP, HWFMPMNP, HWFLPKS1, HWFLPKS2, HWFLPKZ1, HWFLPKZ2 "
    sqlBase = sqlBase & "From " & TbName
    ''1-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''2. 製品WF仕様_2 (TBCME022) の内容を取得する
    ''2-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME022"
    sqlBase = "select HWFCDIR, HWFCSCEN, HWFCSMIN, HWFCSMAX, HWFCSDIS, HWFCSDIR, HWFCKWAY, " & _
            "HWFCKHNM||HWFCKHNN||HWFCKHNH||HWFCKHNU as HWFCKHN, " & _
            "HWFCTDIR, HWFCTCEN, HWFCTMIN, HWFCTMAX, HWFCYDIR, HWFCYCEN, HWFCYMIN, HWFCYMAX, HWFKPTNN, " & _
            "HWFOFPKM||HWFOFPKN||HWFOFPKH||HWFOFPKU as HWFOFPK, " & _
            "HWFOF1PD, HWFOF1PN, HWFOF1PX, HWFOF1PD, HWFOF1PN, HWFOF1PX, " & _
            "HWFOFLKM||HWFOFLKN||HWFOFLKH||HWFOFLKU as HWFOFLK, " & _
            "HWFOF1PW, HWFOF1LC, HWFOF1LN, HWFOF1LX, HWFOF1RF, HWFOFRRC, HWFOFRRN, HWFOFRRX, HWFOFRLC, HWFOFRLN, HWFOFRLX, " & _
            "HWFOF1DC, HWFOF1DN, HWFOF1DX, HWFDFKJ, " & _
            "HWFDFKHM||HWFDFKHN||HWFDFKHH||HWFDFKHU as HWFDFKH, " & _
            "HWFDWCEN, HWFDWMIN, HWFDWMAX, HWFDDCEN, HWFDDMIN, HWFDDMAX, HWFDPDRC, HWFDPACN, HWFDPAMN, HWFDPAMX, " & _
            "HWFDACEN, HWFDAMIN, HWFDAMAX, HWFDBRCN, HWFDBRMN, HWFDBRMX, HWFDPDIR, HWFDPMIN, HWFDPMAX, " & _
            "HWFDRRCN, HWFDRRMN, HWFDRRMX, HWFDLRCN, HWFDLRMN, HWFDLRMX, HWFDPKWY, " & _
            "HWFDPKHM||HWFDPKHB||HWFDPKHH||HWFDPKHU as HWFDPKH,HWFDKK "
    sqlBase = sqlBase & "From " & TbName
    ''2-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    
    ''3. 製品WF仕様_3 (TBCME023) の内容を取得する
    ''3-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME023"
    sqlBase = "select HWFMFORM, KWFMM, HWFMWFCN, HWFMWFMN, HWFMWFMX, HWFMWBCN, HWFMWBMN, HWFMWBMX, HWFMACEN, HWFMAMIN, HWFMAMAX, " & _
              "HWFMFKHM||HWFMFKHN||HWFMFKHH||HWFMFKHU as HWFMFKH, " & _
              "HWFMHCEN, HWFMHMIN, HWFMHMAX, HWFMPWCN, HWFMPWMN, HWFMPWMX, HWFMPRCN, HWFMPRMN, HWFMPRMX, " & _
              "HWFDMFRM, HWFDMM, HWFDMACN, HWFDMPRC, HWFIDWAY, HWFIDPRI, HWFIDKND, HWFIDDIR, HWFIDFAC, HWFCSIZE, " & _
              "HWFIDKBU, HWFIDFIG, HWFIDCON, HWFIDPBS, HWFIDZAR, HWFIDPAP, HWFIDDCN, HWFIDDMX, HWFIDDMN, " & _
              "HWFIDSCN, HWFIDSMX, HWFIDSMN, HWFBDPRS, HWFBDTIM, HWFETWAY, HWFMPFIN, HWFLWASW, HWFCMPUL, HWFTPROC, " & _
              "HWFCDOPMI, HWFCDOPMX, HWFIDD2CN, HWFIDD2MN, HWFIDD2MX, HWFIDDCNN, HWFIDDMNN, HWFIDDMXN, " & _
              "HWFIDSCNN, HWFIDSMNN, HWFIDSMXN, HWFIDS2CN, HWFIDS2MN, HWFIDS2MX,HWFMBACEN,HWFMBAMIN,HWFMBAMAX "
    sqlBase = sqlBase & "From " & TbName
    ''3-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF(3)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''4. 製品WF仕様_4 (TBCME024) の内容を取得する
    ''4-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME024"
    sqlBase = "select HWFM1S, HWFM1H, HWFM2S, HWFM2H, HWFNJSUM, HWFOXCEN, HWFOXMIN, HWFOXMAX, " & _
              "HWFOXSPH||HWFOXSPT||HWFOXSPI as HWFOXSP, " & _
              "HWFOXHWT||HWFOXHWS as HWFOXHW, " & _
              "HWFOXHWY, HWFOXNPO, " & _
              "HWFOXKHM||HWFOXKHN||HWFOXKHH||HWFOXKHU as HWFOXKH, " & _
              "HWFOXZAR, HWFOXMBP, HWFOXMCL, HWFOXMRA, HWFOXLTB, HWFOXLTC, HWFOXLTR, HWFPSCEN, HWFPSMIN, HWFPSMAX, " & _
              "HWFPSSPH||HWFPSSPT||HWFPSSPI as HWFPSSP, " & _
              "HWFPSHWT||HWFPSHWS as HWFPSHW, " & _
              "HWFPSKWY, HWFPSNPS, " & _
              "HWFPSKHM||HWFPSKHN||HWFPSKHH||HWFPSKHU as HWFPSKH, " & _
              "HWFPSMBP, HWFPSMCL, HWFPSMRA, HWFNOXCN, HWFNOXMN, HWFNOXMX, " & _
              "HWFNOXSH||HWFNOXST||HWFNOXSI as HWFNOXS, " & _
              "HWFNOXHT||HWFNOXHS as HWFNOXH, " & _
              "HWFNOXHW, HWFNOXNP, " & _
              "HWFNOXKM||HWFNOXKN||HWFNOXKH||HWFNOXKU as HWFNOXK, " & _
              "HWFNOXMB, HWFNOXMC, HWFNOXMR, HWFMKMIN, HWFMKMAX, " & _
              "HWFMKSPH||HWFMKSPT||HWFMKSPR as HWFMKSP, " & _
              "HWFMKHWT||HWFMKHWS as HWFMKHW, " & _
              "HWFMKSZY, " & _
              "HWFMKKHM||HWFMKKHN||HWFMKKHH||HWFMKKHU as HWFMKKH, " & _
              "HWFMKNSW, HWFMKCET, HWFDZSWY, HWFD1STO, HWFD1STT, HWFD1STG, HWFD2NDO, HWFD2NDC, HWFD2NDT, " & _
              "HWFD3RDO, HWFD3RDT, HWFDZMPS, HWFH2ANO, HWFH2ANT, HWFANGZY "
    sqlBase = sqlBase & "From " & TbName
    ''4-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF(4)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''5. 製品WF仕様_5 (TBCME025) の内容を取得する
    ''5-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME025"
    sqlBase = "select HWFTMMAX, " & _
              "HWFTMSPH||HWFTMSPT||HWFTMSPR as HWFTMSP, " & _
              "HWFTMKHM||HWFTMKHN||HWFTMKHH||HWFTMKHU as HWFTMKH, " & _
              "HWFLTMIN, HWFLTMAX, " & _
              "HWFLTSPH||HWFLTSPT||HWFLTSPI as HWFLTSP, " & _
              "HWFLTHWT||HWFLTHWS as HWFLTHW, " & _
              "HWFLTNSW, HWFLTKWY, " & _
              "HWFLTKHM||HWFLTKHN||HWFLTKHH||HWFLTKHU as HWFLTKH, " & _
              "HWFLTMBP, HWFLTMCL, HWFCNMIN, HWFCNMAX, " & _
              "HWFCNSPH||HWFCNSPT||HWFCNSPI as HWFCNSP, " & _
              "HWFCNHWT||HWFCNHWS as HWFCNHW, " & _
              "HWFCNKWY, " & _
              "HWFCNKHM||HWFCNKHN||HWFCNKHH||HWFCNKHU as HWFCNKH, " & _
              "HWFONMIN, HWFONMAX, " & _
              "HWFONSPH||HWFONSPT||HWFONSPI as HWFONSP, " & _
              "HWFONHWT||HWFONHWS as HWFONHW, " & _
              "HWFONKWY, " & _
              "HWFONKHM||HWFONKHN||HWFONKHH||HWFONKHU as HWFONKH, " & _
              "HWFONMBP, HWFONMCL, HWFONLTB, HWFONLTC, HWFONSDV, HWFONAMN, HWFONAMX, " & _
              "HWFOKBSH||HWFOKBST||HWFOKBSI as HWFOKBS, " & _
              "HWFOKBHT||HWFOKBHS as HWFOKBH, "
    sqlBase = sqlBase & "HWFOS1MN, HWFOS1MX, HWFOS1NS, " & _
              "HWFOS1SH||HWFOS1ST||HWFOS1SI as HWFOS1S, " & _
              "HWFOS1HT||HWFOS1HS as HWFOS1H, " & _
              "HWFOS1HM||HWFOS1KN||HWFOS1KH||HWFOS1KU as HWFOS1K, " & _
              "HWFOS2MN, HWFOS2MX, HWFOS2NS, " & _
              "HWFOS2SH||HWFOS2ST||HWFOS2SI as HWFOS2S, " & _
              "HWFOS2HT||HWFOS2HS as HWFOS2H, " & _
              "HWFOS2KM||HWFOS2KN||HWFOS2KH||HWFOS2KU as HWFOS2K, " & _
              "HWFOS3MN, HWFOS3MX, HWFOS3NS, " & _
              "HWFOS3SH||HWFOS3ST||HWFOS3SI as HWFOS3S, " & _
              "HWFOS3HT||HWFOS3HS as HWFOS3H, " & _
              "HWFOS3KM||HWFOS3KN||HWFOS3KH||HWFOS3KU as HWFOS3K, " & _
              "HWFANTNP, HWFANTIM, HWFANTMN, HWFANTMX, HWFZOMIN, HWFZOMAX, " & _
              "HWFZOSPH||HWFZOSPT||HWFZOSPI as HWFZOSP, " & _
              "HWFZOHWT||HWFZOHWS as HWFZOHW, " & _
              "HWFZONSW, HWFZOKWY, " & _
              "HWFZOKHM||HWFZOKHN||HWFZOKHH||HWFZOKHU as HWFZOKH, " & _
              "HWFTMMAXN, HWFANTTAN "
    sqlBase = sqlBase & ", HWFONZMIN,HWFONZMAX,HWFONZMBP,HWFONZMCL,HWFONZSPH, HWFONZSPT, HWFONZSPI "
    sqlBase = sqlBase & ",HWFONZHWT,HWFONZHWS,HWFONZKWY,HWFONZKHM,HWFONZKHN,HWFONZKHH,HWFONZKHU "
    sqlBase = sqlBase & "From " & TbName
    ''5-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF(5)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''6. 製品WF仕様_6 (TBCME026) の内容を取得する
    ''6-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME026"
    sqlBase = "select HWFBDOMN, HWFBDOMX, " & _
            "HWFBDOSH||HWFBDOST||HWFBDOSR as HWFBDOS, " & _
            "HWFBDOHT||HWFBDOHS as HWFBDOH, " & _
            "HWFBDOSZ, HWFBDONS, HWFBDOET, " & _
            "HWFBDOKM||HWFBDOKN||HWFBDOKH||HWFBDOKU as HWFBDOK, " & _
            "HWFBDSMN, HWFBDSMX, " & _
            "HWFBDSSH||HWFBDSST||HWFBDSSR as HWFBDSS, " & _
            "HWFBDSHT||HWFBDSHS as HWFBDSH, " & _
            "HWFBDSSZ, HWFBDSNS, " & _
            "HWFBDSKM||HWFBDSKN||HWFBDSKH||HWFBDSKU as HWFBDSK, " & _
            "HWFBDSET, HWFRNFMX, " & _
            "HWFRNFSH||HWFRNFST||HWFRNFSI as HWFRNFS, " & _
            "HWFRNFKW, HWFRNFZA, HWFRNBMX, " & _
            "HWFRNBSH||HWFRNBST||HWFRNBSI as HWFRNBS, " & _
            "HWFRNBKW, HWFRNBZA, " & _
            "HWFGDSPH||HWFGDSPT||HWFGDSPR as HWFGDSP, " & _
            "HWFGDSZY, HWFGDZAR, " & _
            "HWFGDKHM||HWFGDKHN||HWFGDKHH||HWFGDKHU as HWFGDKH, " & _
            "HWFNTPUM, HWFDVDMNN, HWFDVDMXN, HWFDSOKE, HWFDSOMN, HWFDSOMX, HWFDSOAN, " & _
            "HWFDSOAX, HWFDSOHT, HWFDSOHS, HWFDSONWY, HWFDSOKM, HWFDSOKN, HWFDSOKH, HWFDSOKU, " & _
            "HWFNP1AR, HWFNP1MAX, HWFNP2AR, HWFNP2MAX,HWFDSOPTK, " & _
            "HWFNTPZA,HWFNTPHT||HWFNTPHS as HWFNTPH, " & _
            "HWFNTPKM||HWFNTPKN||HWFNTPKH||HWFNTPKU as HWFNTPK,HWFGDPTK "
    sqlBase = sqlBase & "From " & TbName
    ''6-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF(6)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''7. 製品WF仕様_7 (TBCME027) の内容を取得する
    ''7-1.SQLを組み立てる(仕様レコード:操業条件='1')
    'HWFSZARA,HWFWARZA,HWFGBZAR,HWFGFDZA,HWFGFRZA,HWFSBZAR,HWFSFZAR   Yam 6/22
    TbName = "TBCME027"
    sqlBase = "select HWFSMIN, HWFSMAX, " & _
              "HWFSHWYT||HWFSHWYS as HWFSHWY, " & _
              "HWFSKWAY, " & _
              "HWFSKHM||HWFSKHN||HWFSKHH||HWFSKHU as HWFSKH, " & _
              "HWFSSZYO, HWFSSREC, HWFSZARAN, HWFSSDEV, HWFSAMIN, HWFSAMAX, HWFWARMX, HWFWARSZ, " & _
              "HWFWARHT||HWFWARHS as HWFWARH, " & _
              "HWFWARKW, HWFWARSR, HWFWARZAN, " & _
              "HWFWARKM||HWFWARKN||HWFWARKH||HWFWARKU as HWFWARK, " & _
              "HWFFSZYO, HWFFSREC, HWFGBMAX, HWFGBPUG, HWFGBPUR, " & _
              "HWFGBHWT||HWFGBHWS as HWFGBHW, " & _
              "HWFGBKW, HWFGBZARN, " & _
              "HWFGBKHM||HWFGBKHN||HWFGBKHH||HWFGBKHU as HWFGBKH, " & _
              "HWFGFDMX, HWFGFDPG, HWFGFDPR, " & _
              "HWFGFDHT||HWFGFDHS as HWFGFDH, " & _
              "HWFGFDKW, HWFGFDZAN, " & _
              "HWFGFDKM||HWFGFDKN||HWFGFDKH||HWFGFDKU as HWFGFDK, " & _
              "HWFGFRMX, HWFGFRPG, HWFGFRPR, " & _
              "HWFGFRHT||HWFGFRHS as HWFGFRH, " & _
              "HWFGFRKW, HWFGFRZAN, " & _
              "HWFGFRKM||HWFGFRKN||HWFGFRKH||HWFGFRKU as HWFGFRK, "
    sqlBase = sqlBase & "HWFSBMAX, HWFSBPUG, HWFSBPUR, HWFSBSZX, HWFSBSZY, " & _
              "HWFSBHWT||HWFSBHWS as HWFSBHW, " & _
              "HWFSBKW, HWFSBZARN, " & _
              "HWFSBKHM||HWFSBKHN||HWFSBKHH||HWFSBKHU as HWFSBKH, " & _
              "HWFSFMAX, HWFSFPUG, HWFSFPUR, HWFSFSZX, HWFSFSZY, " & _
              "HWFSFHWT||HWFSFHWS as HWFSFHW, " & _
              "HWFSFKW, HWFSFZARN, " & _
              "HWFSFKHM||HWFSFKHN||HWFSFKHH||HWFSFKHU as HWFSFKH, " & _
              "HWFFSXOF, HWFFSYOF, HWFFPSUM, HWFSBMAXN, HWFSBPUAGN, HWFSFMAXN, HWFSFPUAGN, " & _
              "HWFSBMAXNN, HWFSBPUGNN, HWFSFMAXNN, HWFSFPUGNN,HWFNTPSZ,HWFMKZA2, " & _
              "HWFMBPMX,HWFMBPHT,HWFMBPHS,HWFMBPSZ,HWFMBPSJ,HWFMWPMX,HWFMWPHT,HWFMWPHS,HWFMLOMX, " & _
              "HWFMLOHT,HWFMLOHS,HWFMLTMX,HWFMLTHT,HWFMLTHS,HWFMLGMX,HWFMLGHT,HWFMLGHS,HWFMSFMX, " & _
              "HWFMSFHT,HWFMSFHS,HWFMHLMX,HWFMHLHT,HWFMHLHS,HWFMPAMX,HWFMPAHT,HWFMPAHS,HWFMTPMX, " & _
              "HWFMTPHT,HWFMTPHS,HWFCSGCEN,HWFCSGMIN,HWFCSGMAX,HWFCSXCEN,HWFCSXMIN,HWFCSXMAX, " & _
              "HWFCSYCEN,HWFCSYMIN,HWFCSYMAX,HWFDSOSZ,HWFMBPHM,HWFMBPHN,HWFMBPHH,HWFMBPHU "
    sqlBase = sqlBase & "From " & TbName
    ''7-2.データを抽出・格納する
    'If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF(7)) = FUNCTION_RETURN_FAILURE Then
    If DispSXL_GetData(TbName, sqlBase & sqlWhereE1, WF(7)) = FUNCTION_RETURN_FAILURE Then  'OPECOND='1'に変更（製作条件入力項目作成）
    
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''7_2. 製品WF仕様_12 (TBCME027) の内容を取得する
    ''7-2.SQLを組み立てる(仕様レコード:操業条件='最新')
    TbName = "TBCME027"
    sqlBase = "select  " & _
              "HWFCSGCEN,HWFCSGMIN,HWFCSGMAX,HWFCSXCEN,HWFCSXMIN,HWFCSXMAX, " & _
              "HWFCSYCEN,HWFCSYMIN,HWFCSYMAX "
    sqlBase = sqlBase & "From " & TbName
    ''7-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF(12)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''8. 製品WF仕様_8 (TBCME028) の内容を取得する
    ''8-1.SQLを組み立てる(仕様レコード:操業条件='1')
    'HWFSPVMX,HWFSPVAM 6/22 Yam
    TbName = "TBCME028"
    sqlBase = "select HWFMK1SI, HWFMK1MX, HWFMK1SZ, HWFMK1ZA, "
    sqlBase = sqlBase & "HWFMK1HT||HWFMK1HS as HWFMK1H, "
    sqlBase = sqlBase & "HWFMK1KM||HWFMK1KN||HWFMK1KH||HWFMK1KU as HWFMK1K, "
    sqlBase = sqlBase & "HWFMKSRE, HWFMKKW, HWFM1B1, HWFM1B1B, HWFM1B2, HWFM1B2B, HWFM1B3, HWFM1B3B, HWFMK2SI, HWFMK2MX, "
    sqlBase = sqlBase & "HWFMK2HT||HWFMK2HS as HWFMK2H, "
    sqlBase = sqlBase & "HWFMK2KM||HWFMK2KN||HWFMK2KH||HWFMK2KU as HWFMK2K, "
    sqlBase = sqlBase & "HWFM2B1, HWFM2B1B, HWFM2B2, HWFM2B2B, HWFM2B3, HWFM2B3B, HWFMK3SI, HWFMK3MX, "
    sqlBase = sqlBase & "HWFMK3HT||HWFMK3HS as HWFMK3H, "
    sqlBase = sqlBase & "HWFMK3KM||HWFMK3KN||HWFMK3KH||HWFMK3KU as HWFMK3K, "
    sqlBase = sqlBase & "HWFM3B1, HWFM3B1B, HWFM3B2, HWFM3B2B, HWFM3B3, HWFM3B3B, HWFMK4SI, HWFMK4MX, "
    sqlBase = sqlBase & "HWFMK4HT||HWFMK4HS as HWFMK4H, "
    sqlBase = sqlBase & "HWFMK4KM||HWFMK4KN||HWFMK4KH||HWFMK4KU as HWFMK4K, "
    sqlBase = sqlBase & "HWFM4B1, HWFM4B1B, HWFM4B2, HWFM4B2B, HWFM4B3, HWFM4B3B, HWFMB1SI, HWFMB1MX, HWFMB1SZ, HWFMB1ZA, "
    sqlBase = sqlBase & "HWFMB1HT||HWFMB1HS as HWFMB1H, "
    sqlBase = sqlBase & "HWFMB1KM||HWFMB1KN||HWFMB1KH||HWFMB1KU as HWFMB1K, "
    sqlBase = sqlBase & "HWFGKNO1, HWFMB2SI, HWFMB2MX, HWFMB2SZ, HWFMB2ZA, "
    sqlBase = sqlBase & "HWFMB2HT||HWFMB2HS as HWFMB2H, "
    sqlBase = sqlBase & "HWFMB2KM||HWFMB2KN||HWFMB2KH||HWFMB2KU as HWFMB2K, "
    sqlBase = sqlBase & "HWFTSPHM||HWFTSPHN||HWFTSPHH||HWFTSPHU as HWFTSPH, "
    sqlBase = sqlBase & "HWFMNMAX, "
    sqlBase = sqlBase & "HWFMK7SI, HWFMK7MX, HWFMK7SZ, HWFMK7ZA, "
    sqlBase = sqlBase & "HWFMK7HT||HWFMK7HS as HWFMK7H,HWFMK7MC, "
    'sqlBase = sqlBase & "HWFMK7KM||HWFMK7KN||HWFMK7KH||HWFMK7KU as HWFMK7K, "
    'sqlBase = sqlBase & "HWFM7B1, HWFM7B1B, HWFM7B2, HWFM7B2B, HWFM7B3, HWFM7B3B, "
    sqlBase = sqlBase & "HWFMK8SI, HWFMK8MX, HWFMK8SZ, HWFMK8ZA, "
    sqlBase = sqlBase & "HWFMK8HT||HWFMK8HS as HWFMK8H,HWFMK8MC, "
    'sqlBase = sqlBase & "HWFMK8KM||HWFMK8KN||HWFMK8KH||HWFMK8KU as HWFMK8K, "
    'sqlBase = sqlBase & "HWFM8B1, HWFM8B1B, HWFM8B2, HWFM8B2B, HWFM8B3, HWFM8B3B, "
    sqlBase = sqlBase & "HWFMK9SI, HWFMK9MX, HWFMK9SZ, HWFMK9ZA, "
    sqlBase = sqlBase & "HWFMK9HT||HWFMK9HS as HWFMK9H,HWFMK9MC, "
    'sqlBase = sqlBase & "HWFMK9KM||HWFMK9KN||HWFMK9KH||HWFMK9KU as HWFMK9K, "
    'sqlBase = sqlBase & "HWFM9B1, HWFM9B1B, HWFM9B2, HWFM9B2B, HWFM9B3, HWFM9B3B, "
    sqlBase = sqlBase & "HWFMK10SI, HWFMK10MX, HWFMK10SZ, HWFMK10ZA, "
    sqlBase = sqlBase & "HWFMK10HT||HWFMK10HS as HWFMK10H,HWFMK10MC, "
    'sqlBase = sqlBase & "HWFMK10KM||HWFMK10KN||HWFMK10KH||HWFMK10KU as HWFMK10K, "
    'sqlBase = sqlBase & "HWFM10B1, HWFM10B1B, HWFM10B2, HWFM10B2B, HWFM10B3, HWFM10B3B, "
    sqlBase = sqlBase & "HWFMK11SI, HWFMK11MX, HWFMK11SZ, HWFMK11ZA, "
    sqlBase = sqlBase & "HWFMK11HT||HWFMK11HS as HWFMK11H,HWFMK11MC,  "
    'sqlBase = sqlBase & "HWFMK11KM||HWFMK11KN||HWFMK11KH||HWFMK11KU as HWFMK11K, "
    'sqlBase = sqlBase & "HWFM11B1, HWFM11B1B, HWFM11B2, HWFM11B2B, HWFM11B3, HWFM11B3B, "
    sqlBase = sqlBase & "HWFMK12SI, HWFMK12MX, HWFMK12SZ, HWFMK12ZA, "
    sqlBase = sqlBase & "HWFMK12HT||HWFMK12HS as HWFMK12H,HWFMK12MC,  "
    'sqlBase = sqlBase & "HWFMK12KM||HWFMK12KN||HWFMK12KH||HWFMK12KU as HWFMK12K, "
    'sqlBase = sqlBase & "HWFM12B1, HWFM12B1B, HWFM12B2, HWFM12B2B, HWFM12B3, HWFM12B3B, "
    sqlBase = sqlBase & "HWFMK13SI, HWFMK13MX, HWFMK13SZ, HWFMK13ZA, "
    sqlBase = sqlBase & "HWFMK13HT||HWFMK13HS as HWFMK13H,HWFMK13MC,  "
    'sqlBase = sqlBase & "HWFMK13KM||HWFMK13KN||HWFMK13KH||HWFMK13KU as HWFMK13K, "
    'sqlBase = sqlBase & "HWFM13B1, HWFM13B1B, HWFM13B2, HWFM13B2B, HWFM13B3, HWFM13B3B, "
    sqlBase = sqlBase & "HWFMK14SI, HWFMK14MX, HWFMK14SZ, HWFMK14ZA, "
    sqlBase = sqlBase & "HWFMK14HT||HWFMK14HS as HWFMK14H,HWFMK14MC, "
    'sqlBase = sqlBase & "HWFMK14KM||HWFMK14KN||HWFMK14KH||HWFMK14KU as HWFMK14K, "
    'sqlBase = sqlBase & "HWFM14B1, HWFM14B1B, HWFM14B2, HWFM14B2B, HWFM14B3, HWFM14B3B, "
    sqlBase = sqlBase & "HWFMK15SI, HWFMK15MX, HWFMK15SZ, HWFMK15ZA, "
    sqlBase = sqlBase & "HWFMK15HT||HWFMK15HS as HWFMK15H,HWFMK15MC, "
    'sqlBase = sqlBase & "HWFMK15KM||HWFMK15KN||HWFMK15KH||HWFMK15KU as HWFMK15K, "
    'sqlBase = sqlBase & "HWFM15B1, HWFM15B1B, HWFM15B2, HWFM15B2B, HWFM15B3, HWFM15B3B, "
    sqlBase = sqlBase & "HWFMNSPH||HWFMNSPT||HWFMNSPI as HWFMNSP, " & _
        "HWFMNALX, HWFMNCAX, HWFMNCRX, HWFMNCUX, HWFMNFEX, " & _
        "HWFMNHWT||HWFMNHWS as HWFMNHW, " & _
        "HWFMNKWY, " & _
        "HWFMNKHM||HWFMNKHN||HWFMNKHH||HWFMNKHU as HWFMNKH, " & _
        "HWFMNKMX, HWFMNMGX, HWFMNNAX, HWFMNNIX, HWFMNZNX, HWFSPVMXN, " & _
        "HWFSPVSH||HWFSPVST||HWFSPVSI as HWFSPVS,HWFSPVSH, HWFSPVST, HWFSPVSI, " & _
        "HWFSPVHT||HWFSPVHS as HWFSPVH, " & _
        "HWFSPVKM||HWFSPVKN||HWFSPVKH||HWFSPVKU as HWFSPVK, " & _
        "HWFSPVKM,HWFSPVKN,HWFSPVKH,HWFSPVKU, " & _
        "HWFDLMIN, HWFDLMAX,HWFDLSPH, HWFDLSPT, HWFDLSPI,  " & _
        "HWFDLSPH||HWFDLSPT||HWFDLSPI as HWFDLSP, " & _
        "HWFDLHWT||HWFDLHWS as HWFDLHW, " & _
        "HWFDLKHM||HWFDLKHN||HWFDLKHH||HWFDLKHU as HWFDLKH, " & _
        "HWFOTMIN, HWFOTMX1, HWFOTMX2, " & _
        "HWFOTSPH||HWFOTSPT||HWFOTSPI as HWFOTSP, " & _
        "HWFOTKWY, " & _
        "HWFOTKW1||HWFOTKW2 as HWFOTKW, " & _
        "HWFOTHWT||HWFOTHWS as HWFOTHW, " & _
        "HWFOTKHM||HWFOTKHN||HWFOTKHH||HWFOTKHU as HWFOTKH, " & _
        "HWFMK1MC, HWFMK2SZ, HWFMK2ZAR, HWFMK2MC, HWFMK3SZ, HWFMK3ZAR, HWFMK3MC, " & _
        "HWFMK4SZ, HWFMK4ZAR, HWFMK4MC, " & _
        "HWFMK5MC, HWFMK5B1, HWFMK5B1B, HWFMK5B2, HWFMK5B2B, HWFMK5B3, HWFMK5B3B, " & _
        "HWFMK6MC, HWFMK6B1, HWFMK6B1B, HWFMK6B2, HWFMK6B2B, HWFMK6B3, HWFMK6B3B, HWFSPVAMN,HWFMKMAP,HWFMK1ZAN "
    sqlBase = sqlBase & "From " & TbName
    ''8-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF(8)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''9. 製品WF仕様_9 (TBCME029) の内容を取得する
    ''9-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME029"
    sqlBase = "select HWFOF1AX, HWFOF1MX, HWFOF1ET, HWFOF1NS, HWFOF1SZ, " & _
            "HWFOF1SH||HWFOF1ST||HWFOF1SR as HWFOF1S, " & _
            "HWFOF1HT||HWFOF1HS as HWFOF1H, " & _
            "HWFOF1KM||HWFOF1KN||HWFOF1KH||HWFOF1KU as HWFOF1K, " & _
            "HWFOF2AX, HWFOF2MX, HWFOF2ET, HWFOF2NS, HWFOF2SZ, " & _
            "HWFOF2SH||HWFOF2ST||HWFOF2SR as HWFOF2S, " & _
            "HWFOF2HT||HWFOF2HS as HWFOF2H, " & _
            "HWFOF2KM||HWFOF2KN||HWFOF2KH||HWFOF2KU as HWFOF2K, " & _
            "HWFOF3AX, HWFOF3MX, HWFOF3ET, HWFOF3NS, HWFOF3SZ, " & _
            "HWFOF3SH||HWFOF3ST||HWFOF3SR as HWFOF3S, " & _
            "HWFOF3HT||HWFOF3HS as HWFOF3H, " & _
            "HWFOF3KM||HWFOF3KN||HWFOF3KH||HWFOF3KU as HWFOF3K, " & _
            "HWFOF4AX, HWFOF4MX, HWFOF4ET, HWFOF4NS, HWFOF4SZ, " & _
            "HWFOF4SH||HWFOF4ST||HWFOF4SR as HWFOF4S, " & _
            "HWFOF4HT||HWFOF4HS as HWFOF4H, " & _
            "HWFOF4KM||HWFOF4KN||HWFOF4KH||HWFOF4KU as HWFOF4K, "
    sqlBase = sqlBase & "HWFBM1AN, HWFBM1AX, HWFBM1ET, HWFBM1NS, HWFBM1SZ, " & _
            "HWFBM1SH||HWFBM1ST||HWFBM1SR as HWFBM1S, " & _
            "HWFBM1HT||HWFBM1HS as HWFBM1H, " & _
            "HWFBM1KM||HWFBM1KN||HWFBM1KH||HWFBM1KU as HWFBM1K, " & _
            "HWFBM2AN, HWFBM2AX, HWFBM2ET, HWFBM2NS, HWFBM2SZ, " & _
            "HWFBM2SH||HWFBM2ST||HWFBM2SR as HWFBM2S, " & _
            "HWFBM2HT||HWFBM2HS as HWFBM2H, " & _
            "HWFBM2KM||HWFBM2KN||HWFBM2KH||HWFBM2KU as HWFBM2K, " & _
            "HWFBM3AN, HWFBM3AX, HWFBM3ET, HWFBM3NS, HWFBM3SZ, " & _
            "HWFBM3SH||HWFBM3ST||HWFBM3SR as HWFBM3S, " & _
            "HWFBM3HT||HWFBM3HS as HWFBM3H, " & _
            "HWFBM3KM||HWFBM3KN||HWFBM3KH||HWFBM3KU as HWFBM3K, " & _
            "HWFOSPAX, HWFOSPMX, " & _
            "HWFOSPSH||HWFOSPST||HWFOSPSR as HWFOSPS, " & _
            "HWFOSPHT||HWFOSPHS as HWFOSPH, " & _
            "HWFOSPNS, HWFOSPET, HWFOSPSZ, " & _
            "HWFOSPKM||HWFOSPKN||HWFOSPKH||HWFOSPKU as HWFOSPK, " & _
            "HWFRS1Y, HWFRS1N, HWFRS2Y, HWFRS2N, HWFRS3Y, HWFRS3N, HWFRS4Y, HWFRS4N, HWFRS5Y, HWFRS5N, " & _
            "HWFRS6Y, HWFRS6N, HWFRS7Y, HWFRS7N, HWFRS8Y, HWFRS8N, HWFRS9Y, HWFRS9N, HWFRS10Y, HWFRS10N, " & _
            "HWFNOTE, HWFOSF1PTK, HWFOSF2PTK, HWFOSF3PTK, HWFOSF4PTK, " & _
            "HWFBM1MBP, HWFBM1MCL, HWFBM2MBP, HWFBM2MCL, HWFBM3MBP, HWFBM3MCL, " & _
            "HWFCOSF3PK,HWFCOSF3SH,HWFCOSF3ST,HWFCOSF3SR,HWFCOSF3HT,HWFCOSF3HS,HWFCOSF3SZ,HWFCOSF3NS, " & _
            "HWFCPK,HWFCSZ,HWFCHT,HWFCHS,HWFCJPK,HWFCJNS,HWFCJHT,HWFCJHS,HWFCJLTPK,HWFCJLTNS," & _
            "HWFCJLTHT,HWFCJLTHS,HWFCJ2PK,HWFCJ2NS,HWFCJ2HT,HWFCJ2HS "
    sqlBase = sqlBase & "From " & TbName
    ''9-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF(9)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If

    
    ''11. 製品WF仕様_11 (TBCME048) の内容を取得する
    ''11-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME048"
    
    sqlBase = "select " & _
              "HWFMSR,HWFMHT||HWFMHS as HWFMH,HWFMFBCN,HWFMFBMX,HWFMFBMN," & _
              "HWFMFKHM||HWFMSKHN||HWFMSKHH||HWFMSKHU as HWFMSKH,HWFMSWFCN,HWFMSWFMN,HWFMSWFMX," & _
              "HWFMSWBCN,HWFMSFCN,HWFMSFMN,HWFMSFMX,HWFMSSR,HWFMSHT||HWFMSHS as HWFMSH," & _
              "HWFMSSHCN,HWFMSSHMN,HWFMSSHMX, " & _
              "HWFMSPWCN,HWFMSPWMN,HWFMSPWMX,HWFMSPRCN,HWFMSPRMN,HWFMSPRMX,HWFMSBMN,HWFMSBMX, " & _
              "HWFMSWBMN,HWFMSWBMX,HWFMSFBCN,HWFMSFBMN,HWFMSFBMX,HWFNTPSR,HWFMSBCN,HWFLTPUG,HWFLTPUR, " & _
              "HWFF01ST,HWFF02ST, " & _
              "HWFF03MAX,HWFF03PUR,HWFF03PUG,HWFF03SZX,HWFF03SZY,HWFF03HT||HWFF03HS as HWFF03H,HWFF03KW,HWFF03ST, " & _
              "HWFF04MAX,HWFF04PUR,HWFF04PUG,HWFF04SZX,HWFF04SZY,HWFF04HT||HWFF04HS as HWFF04H,HWFF04KW,HWFF04ST, " & _
              "HWFF05MAX,HWFF05PUR,HWFF05PUG,HWFF05SZX,HWFF05SZY,HWFF05HT||HWFF05HS as HWFF05H,HWFF05KW,HWFF05ST, " & _
              "HWFF06MAX,HWFF06PUR,HWFF06PUG,HWFF06SZX,HWFF06SZY,HWFF06HT||HWFF06HS as HWFF06H,HWFF06KW,HWFF06ST, " & _
              "HWFF07MAX,HWFF07PUR,HWFF07PUG,HWFF07SZX,HWFF07SZY,HWFF07HT||HWFF07HS as HWFF07H,HWFF07KW,HWFF07ST, " & _
              "HWFF08MAX,HWFF08PUR,HWFF08PUG,HWFF08SZX,HWFF08SZY,HWFF08HT||HWFF08HS as HWFF08H,HWFF08KW,HWFF08ST, " & _
              "HWFF09MAX,HWFF09PUR,HWFF09PUG,HWFF09SZX,HWFF09SZY,HWFF09HT||HWFF09HS as HWFF09H,HWFF09KW,HWFF09ST, " & _
              "HWFF10MAX,HWFF10PUR,HWFF10PUG,HWFF10SZX,HWFF10SZY,HWFF10HT||HWFF10HS as HWFF10H,HWFF10KW,HWFF10ST, " & _
              "HWFMNAGX,HWFMNCOX,HWFMNLIX,HWFMNMNX,HWFMNMOX,HWFMNTIX,HWFMNVX,HWFBMNNI," & _
              "HWFBMNSPH||HWFBMNSPT||HWFBMNSPI as HWFBMNSP, " & _
              "HWFBMNALX,HWFBMNCAX,HWFBMNCRX,HWFBMNCUX,HWFBMNFEX,HWFBMNHWT||HWFBMNHWS as HWFBMNHW, " & _
              "HWFBMNKWY,HWFBMNKHM||HWFBMNKHN||HWFBMNKHH||HWFBMNKHU as HWFBMNKH, " & _
              "HWFBMNKX,HWFBMNMGX,HWFBMNNAX,HWFBMNNIX,HWFBMNZNX,HWFBMNAGX,HWFBMNCOX,HWFBMNLIX, " & _
              "HWFBMNMNX,HWFBMNMOX,HWFBMNTIX,HWFBMNVX,HWFSPVPUG,HWFSPVPUR,HWFSPVSTD,HWFDLPUG, " & _
              "HWFDLPUR,HWFNRMX,HWFNRSH||HWFNRST||HWFNRSI as HWFNRS,HWFNRHT||HWFNRHS as HWFNRH, " & _
              "HWFNRKM||HWFNRKN||HWFNRKH||HWFNRKU as HWFNRK, HWFNRAM,HWFNRPUG,HWFNRPUR,HWFNRSTD "
    '2007/04/24 追加
    sqlBase = sqlBase & ",HWFLTCEL,HWFLTZAR,HWFNTPHV,HWFF03MAXN,HWFF03PUGN,HWFF04MAXN,HWFF04PUGN,HWFF05MAXN,HWFF05PUGN, "
    sqlBase = sqlBase & "HWFF06MAXN,HWFF06PUGN,HWFF07MAXN,HWFF07PUGN,HWFF08MAXN,HWFF08PUGN,HWFF09MAXN,HWFF09PUGN,HWFF10MAXN,HWFF10PUGN, "
    sqlBase = sqlBase & "HWFMS,HWFMPACN,HWFMPAMN,HWFMPAMX,HWFMSS,HWFMSPACN,HWFMSPAMN,HWFMSPAMX,HWFMNWMX,HWFBMNWMX "
    sqlBase = sqlBase & ",HWFMSR1CN,HWFMSR1MX,HWFMSR1MN,HWFMSR2CN,HWFMSR2MX,HWFMSR2MN,HWFMSB1CN,HWFMSB1MX,HWFMSB1MN "
    sqlBase = sqlBase & ",HWFMSB2CN,HWFMSB2MX,HWFMSB2MN,HWFMSBXCN,HWFMSBXMX,HWFMSBXMN,HWFMSFECN,HWFMSFEMX,HWFMSFEMN "
    sqlBase = sqlBase & ",HWFMSBECN,HWFMSBEMX,HWFMSBEMN,HWFSIRDMX,HWFSIRDSZ,HWFSIRDHT,HWFSIRDHS,HWFSIRDKM,HWFSIRDKN "
    sqlBase = sqlBase & ",HWFSIRDKH,HWFSIRDKU,HSXNOTE2,HWFNOTE2,HEPNOTE2 "
    sqlBase = sqlBase & " From " & TbName
    ''11-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF(11)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If

    ''顧客仕様 (TBCME001) の内容を取得する（tabは特記欄）
    ''SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME001"
    sqlBase = "select KMGNOTE, KMGEPNOTE "
    sqlBase = sqlBase & "From " & TbName
    ''6-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhereE1, TBCME001) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    DBDRV_s_cmzcF_cmgc001d_DispWF = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :製品仕様入力画面用 タブ3に表示する内容を得る
'ﾊﾟﾗﾒｰﾀ    :変数名         ,IO ,型          ,説明
'          :targetHinban   ,   ,tFullHinban ,品番情報
'          :WF()           ,   ,c_cmzcrec   ,製品仕様WF 1-9 の内容
'          :errTbl         ,O  ,String      ,取得できなかったテーブル名
'          :戻り値         ,O  ,FUNCTION_RETURN,検索の成否
'説明      :出力パラメータの配列は、(1 to 9)に製品仕様WFデータ1～9 が入る
'          :呼出側で品番を12桁入力可能なため、該当品番が存在しない場合がありうる
'履歴      :2001/09/28 作成  野村
Public Function DBDRV_s_cmzcF_cmgc001d_DispTAB3(targetHinban As tFullHinban, Wf_1() As c_cmzcrec, Wf_4() As c_cmzcrec, Wf_5() As c_cmzcrec, _
                                               Wf_6() As c_cmzcrec, Wf_8() As c_cmzcrec, WF_9() As c_cmzcrec, WF_10() As c_cmzcrec, WF_11() As c_cmzcrec, _
                                              WfKokyaku13() As c_cmzcrec, WfKokyaku12() As c_cmzcrec, _
                                              WfKokyaku15() As c_cmzcrec, WfKokyaku16() As c_cmzcrec, _
                                              WfKokyaku11() As c_cmzcrec, WfKokyaku08() As c_cmzcrec, _
                                              EpiSiyou() As c_cmzcrec, EpiKokyaku() As c_cmzcrec, SlgSiyou() As c_cmzcrec, errTbl$) As FUNCTION_RETURN

Dim sql As String
Dim sqlBase As String       'SQLの基本部
Dim sqlWhere As String      'Where句以降
Dim sqlWhere2 As String      'Where句以降
Dim TbName As String        'テーブル名
Dim i As Integer
Dim rs As OraDynaset
        
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001d_SQL.bas -- Function DBDRV_s_cmzcF_cmgc001d_DispTAB3"
    
    DBDRV_s_cmzcF_cmgc001d_DispTAB3 = FUNCTION_RETURN_FAILURE
    errTbl = vbNullString
    
    ''SQLの共通部分を準備する
    With targetHinban
        '仕様のWHERE句(WF仕様は内側管理しない)
        sqlWhere2 = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='" & .OPECOND & "')"
        
        sqlWhere = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='1')" & _
              " Order by OPECOND DESC"
    End With
    
    ''出力データを初期化する
    ''出力データを初期化する
    For i = 1 To 2
        Set Wf_1(i) = New c_cmzcrec
        Set Wf_4(i) = New c_cmzcrec
        Set Wf_5(i) = New c_cmzcrec
        Set Wf_6(i) = New c_cmzcrec
        Set Wf_8(i) = New c_cmzcrec
        Set WF_9(i) = New c_cmzcrec
        Set WF_10(i) = New c_cmzcrec
        Set WF_11(i) = New c_cmzcrec
        Set WfKokyaku13(i) = New c_cmzcrec
        Set WfKokyaku12(i) = New c_cmzcrec
        Set WfKokyaku15(i) = New c_cmzcrec
        Set WfKokyaku16(i) = New c_cmzcrec
        Set WfKokyaku11(i) = New c_cmzcrec
        Set WfKokyaku08(i) = New c_cmzcrec
        Set EpiSiyou(i) = New c_cmzcrec
        Set SlgSiyou(i) = New c_cmzcrec
    Next
    
    ''無効品番のチェック（製品仕様SXL1に登録されていること、製作条件付与取消に登録されていないこと）
    With targetHinban
        sql = "select A.HINBAN from TBCME018 A, TBCME031 B " & _
              "where (A.HINBAN='" & .hinban & "') and (A.MNOREVNO=" & .MNOREVNO & ") and (A.FACTORY='" & .FACTORY & "')" & _
              " and (A.HINBAN=B.HINBAN(+)) and (A.MNOREVNO=B.MNOREVNO(+)) and (A.FACTORY=B.FACTORY(+))" & _
              " and (B.HINBAN is null)"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If (rs Is Nothing) Or (rs.RecordCount = 0) Then
            errTbl = "TBCME031"
            GoTo proc_exit
        End If
    End With
    
    
    ''0. 顧客仕様_1 (TBCME013) の内容を取得する
    ''0-1.SQLを組み立てる
    TbName = "TBCME008"
    sqlBase = "Select KPRRMIN, KPRRMAX, KPRRSPOH, KPRRSPOT,KPRRSPOI,KPRRHWYT,KPRRHWYS,KPRRKWAY, " & _
                " KPRRKHNM,KPRRKHNN,KPRRKHNH,KPRRKHNU,KPRRSDEV,KPRRMBNP,KPRRMCAL "
    sqlBase = sqlBase & "From " & TbName
    ''1-2.データを抽出・格納する
    ''1-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WfKokyaku08(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    Set WfKokyaku08(2) = WfKokyaku08(1)
    
    
    ''0. 顧客仕様_1 (TBCME013) の内容を取得する
    ''0-1.SQLを組み立てる
    TbName = "TBCME013"
    sqlBase = "SELECT KPRDENMX, KPRDENMN, KPRDENHT, KPRDENHS, KPRDENKU,KPRLDLMN, KPRLDLMX, " & _
                "KPRLDLHT, KPRLDLHS, KPRLDLKU,KPRDVDMNN, KPRDVDMXN, KPRDVDHT, KPRDVDHS, KPRDVDKU "
    sqlBase = sqlBase & ", KPRDSOMX, KPRDSOMN, KPRDSOAX, KPRDSOAN, KPRDSOHT, KPRDSOHS, KPRDSONWY "
    sqlBase = sqlBase & ", KPRDSOKM, KPRDSOKH, KPRDSOKN, KPRDSOKU "
    sqlBase = sqlBase & ", KPRGDSPH, KPRGDSPT, KPRGDSPR, KPRGDSZY, KPRGDZAR, KPRGDKHM, KPRGDKHH, KPRGDKHN, KPRGDKHU "
    sqlBase = sqlBase & "From " & TbName
    ''1-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WfKokyaku13(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    Set WfKokyaku13(2) = WfKokyaku13(1)
    ''0. 顧客仕様_1 (TBCME016) の内容を取得する
    ''0-1.SQLを組み立てる
    TbName = "TBCME016"
    sqlBase = "SELECT KPROS1AX, KPROS1MX, KPROS1NS, KPROS1ET, KPROS1SH, KPROS1ST, " & _
                "KPROS1SR, KPROS1HT, KPROS1HS, KPROS1KM, KPROS1KH, KPROSF1PTK, KPROS1SZ "
    sqlBase = sqlBase & ", KPROS2AX, KPROS2MX, KPROS2NS, KPROS2ET, KPROS2SH, KPROS2ST, " & _
                "KPROS2SR, KPROS2HT, KPROS2HS, KPROS2KM, KPROS2KH, KPROSF2PTK, KPROS2SZ "
    sqlBase = sqlBase & ", KPROS3AX, KPROS3MX, KPROS3NS, KPROS3ET, KPROS3SH, KPROS3ST, " & _
                "KPROS3SR, KPROS3HT, KPROS3HS, KPROS3KM, KPROS3KH, KPROSF3PTK, KPROS3SZ "
    sqlBase = sqlBase & ", KPROS4AX, KPROS4MX, KPROS4NS, KPROS4ET, KPROS4SH, KPROS4ST, " & _
                "KPROS4SR, KPROS4HT, KPROS4HS, KPROS4KM, KPROS4KH, KPROSF4PTK, KPROS4SZ "
    sqlBase = sqlBase & ", KPRBM1AN, KPRBM1AX, KPRBM1ET, KPRBM1NS, KPRBM1SH, KPRBM1ST, " & _
                "KPRBM1SR, KPRBM1HT, KPRBM1HS, KPRBM1KM, KPRBM1KN, KPRBM1KH, KPRBM1KU, KPRBM1MBP, KPRBM1MCL, KPRBM1SZ "
    sqlBase = sqlBase & ", KPRBM2AN, KPRBM2AX, KPRBM2ET, KPRBM2NS, KPRBM2SH, KPRBM2ST, " & _
                "KPRBM2SR, KPRBM2HT, KPRBM2HS, KPRBM2KM, KPRBM2KN, KPRBM2KH, KPRBM2KU, KPRBM2MBP, KPRBM2MCL, KPRBM2SZ "
    sqlBase = sqlBase & ", KPRBM3AN, KPRBM3AX, KPRBM3ET, KPRBM3NS, KPRBM3SH, KPRBM3ST, " & _
                "KPRBM3SR, KPRBM3HT, KPRBM3HS, KPRBM3KM, KPRBM3KN, KPRBM3KH, KPRBM3KU, KPRBM3MBP, KPRBM3MCL, KPRBM3SZ "
    sqlBase = sqlBase & ", KPROSPAX, KPROSPMX, KPROSPET, KPROSPNS, KPROSPSH,KPROSPST,KPROSPSR, KPROSPHT,KPROSPHS "
    '2007/04/24
    sqlBase = sqlBase & ", KPRBM1ANN, KPRBM1AXN, KPRBM2ANN, KPRBM2AXN,KPRBM3ANN, KPRBM3AXN "
    sqlBase = sqlBase & "From " & TbName
    ''1-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WfKokyaku16(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    Set WfKokyaku16(2) = WfKokyaku16(1)
    ''0. 顧客仕様_1 (TBCME012) の内容を取得する
    ''0-1.SQLを組み立てる
    TbName = "TBCME012"
    sqlBase = "SELECT KPRLTMIN, KPRLTMAX, KPRLTSPH, KPRLTSPT, KPRLTSPI, KPRLTHWT, KPRLTHWS, KPRLTKWY "
    sqlBase = sqlBase & ", KPRONMIN, KPRONMAX, KPRONSPH, KPRONSPT, KPRONSPI, KPRONHWT, KPRONHWS, KPRONKWY "
    sqlBase = sqlBase & ", KPRONKHM, KPRONKHN, KPRONKHH, KPRONKHU, KPRONSDV, KPRONMBP, KPRONMCL "
    sqlBase = sqlBase & ", KPRCNMIN, KPRCNMAX, KPRCNSPH, KPRCNSPT, KPRCNSPI, KPRCNHWT, KPRCNHWS, KPRCNKWY "
    sqlBase = sqlBase & ", KPRZOMIN, KPRZOMAX, KPRZOSPH, KPRZOSPT, KPRZOSPI, KPRZOHWT, KPRZOHWS, KPRZOKWY, KPRZONSW "
    sqlBase = sqlBase & ", KPROS1MN, KPROS1MX, KPROS1NS, KPROS1SH, KPROS1ST, KPROS1SI, KPROS1HT "
    sqlBase = sqlBase & ", KPROS1HS, KPROS1HM, KPROS1KN, KPROS1KH, KPROS1KU "
    sqlBase = sqlBase & ", KPROS2MN, KPROS2MX, KPROS2NS, KPROS2SH, KPROS2ST, KPROS2SI, KPROS2HT "
    sqlBase = sqlBase & ", KPROS2HS, KPROS2KM, KPROS2KN, KPROS2KH, KPROS2KU "
    sqlBase = sqlBase & ", KPRZOKHM, KPRZOKHN, KPRZOKHH, KPRZOKHU "
    sqlBase = sqlBase & "From " & TbName
    ''1-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WfKokyaku12(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    Set WfKokyaku12(2) = WfKokyaku12(1)
    
    TbName = "TBCME015"
    sqlBase = "select KPRSPVMXN, KPRSPVKM,KPRSPVKN, KPRSPVAMN, KPRSPVHWS, KPRDLMIN, KPRDLMAX, KPRDLHWT,KPRDLHWS "
    sqlBase = sqlBase & ", KPRDLSPH, KPRDLSPT, KPRDLSPI,KPRDLKHN  "
        sqlBase = sqlBase & "From " & TbName
    ''8-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WfKokyaku15(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    Set WfKokyaku15(2) = WfKokyaku15(1)
    TbName = "TBCME011"
    sqlBase = "select  KPRDZSWY, KPRD1STO, KPRD1STT, KPRD1STG, KPRD2NDO, KPRD2NDC, KPRD2NDT, " & _
              "KPRD3RDO, KPRH2ANO, KPRDZMPS, KPRD3RDT, " & _
              "KPRMKMIN, KPRMKMAX, KPRMKSPH, KPRMKSPT, KPRMKSPR, KPRMKHWT, KPRMKHWS, KPRMKSZY, " & _
              "KPRMKKHM, KPRMKKHN, KPRMKKHH, KPRMKKHU, KPRMKNSW, KPRMKCET "
    sqlBase = sqlBase & "From " & TbName
    ''4-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WfKokyaku11(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    Set WfKokyaku11(2) = WfKokyaku11(1)
    ''1. 製品WF仕様_1 (TBCME021) の内容を取得する
    ''1-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME021"
    sqlBase = "select  HWFRMIN, HWFRMAX, " & _
              "HWFRSPOH,HWFRSPOT,HWFRSPOI, " & _
              "HWFRHWYT,HWFRHWYS, " & _
              "HWFRKWAY, " & _
              "HWFRKHNM,HWFRKHNN,HWFRKHNH,HWFRKHNU, " & _
              "HWFRSDEV,HWFRMBNP, HWFRMCAL "
    sqlBase = sqlBase & "From " & TbName
    ''1-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Wf_1(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    If DispSXL_GetData(TbName, "select * from TBCME021 " & sqlWhere2, Wf_1(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''2. 製品WF仕様_2 (TBCME022) の内容を取得する
    ''2-1.SQLを組み立てる(仕様レコード:操業条件='1')
    
    ''3. 製品WF仕様_3 (TBCME023) の内容を取得する
    ''3-1.SQLを組み立てる(仕様レコード:操業条件='1')
    
    ''4. 製品WF仕様_4 (TBCME024) の内容を取得する
    ''4-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME024"
    sqlBase = "select  HWFDZSWY, HWFD1STO, HWFD1STT, HWFD1STG, HWFD2NDO, HWFD2NDC, HWFD2NDT, " & _
              "HWFD3RDO, HWFH2ANO, HWFDZMPS, HWFD3RDT," & _
              "HWFMKMIN, HWFMKMAX, HWFMKSPH, HWFMKSPT, HWFMKSPR, HWFMKHWT, HWFMKHWS, HWFMKSZY, " & _
              "HWFMKKHM, HWFMKKHN, HWFMKKHH, HWFMKKHU, HWFMKNSW, HWFMKCET "
    sqlBase = sqlBase & "From " & TbName
    ''4-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Wf_4(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    If DispSXL_GetData(TbName, "select * from TBCME024 " & sqlWhere2, Wf_4(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''5. 製品WF仕様_5 (TBCME025) の内容を取得する
    ''5-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME025"
    sqlBase = "select  HWFCNMIN, HWFCNMAX, " & _
              "HWFCNSPH,HWFCNSPT,HWFCNSPI, " & _
              "HWFCNHWT,HWFCNHWS, " & _
              "HWFCNKWY, " & _
              "HWFCNKHM,HWFCNKHN,HWFCNKHH,HWFCNKHU, " & _
              "HWFONMIN, HWFONMAX, " & _
              "HWFONSPH,HWFONSPT,HWFONSPI, " & _
              "HWFONHWT,HWFONHWS, " & _
              "HWFONKWY, " & _
              "HWFONKHM,HWFONKHN,HWFONKHH,HWFONKHU, " & _
              "HWFONMBP, HWFONMCL, HWFONSDV "
    sqlBase = sqlBase & ", HWFLTMIN, HWFLTMAX, HWFLTSPH, HWFLTSPT, HWFLTSPI, HWFLTHWT, HWFLTHWS, HWFLTKWY "
    sqlBase = sqlBase & ", HWFCNMIN, HWFCNMAX, HWFCNSPH, HWFCNSPT, HWFCNSPI, HWFCNHWT, HWFCNHWS, HWFCNKWY "
    sqlBase = sqlBase & ", HWFCNKHM, HWFCNKHN, HWFCNKHH, HWFCNKHU "
    sqlBase = sqlBase & ", HWFZOMIN, HWFZOMAX, HWFZOSPH, HWFZOSPT, HWFZOSPI, HWFZOHWT, HWFZOHWS, HWFZOKWY, HWFZONSW "
    sqlBase = sqlBase & ", HWFOS1MN,HWFOS1MX,HWFOS1NS,HWFOS1SH,HWFOS1ST,HWFOS1SI,HWFOS1HT,HWFOS1HS,HWFOS1HM,HWFOS1KN,HWFOS1KH,HWFOS1KU "
    sqlBase = sqlBase & ", HWFOS2MN,HWFOS2MX,HWFOS2NS,HWFOS2SH,HWFOS2ST,HWFOS2SI,HWFOS2HT,HWFOS2HS,HWFOS2KM,HWFOS2KN,HWFOS2KH,HWFOS2KU "
    sqlBase = sqlBase & ", HWFZOKHM,HWFZOKHN,HWFZOKHH,HWFZOKHU "
    sqlBase = sqlBase & "From " & TbName
    ''5-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Wf_5(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    If DispSXL_GetData(TbName, "select * from TBCME025 " & sqlWhere2, Wf_5(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    ''6. 製品WF仕様_6 (TBCME026) の内容を取得する
    ''6-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME026"
    sqlBase = "SELECT HWFDENKU, HWFDENMX, HWFDENMN, HWFDENHT, HWFDENHS, HWFDENKU, HWFLDLKU, HWFLDLMN, HWFLDLMX, " & _
                "HWFLDLHT, HWFLDLHS, HWFLDLKU, HWFDVDKU, HWFDVDMNN, HWFDVDMXN, HWFDVDHT, HWFDVDHS, HWFDVDKU, " & _
                "HWFDSOMX, HWFDSOMN, HWFDSOAX, HWFDSOAN, HWFDSOHT, HWFDSOHS, HWFDSONWY, HWFDSOPTK, " & _
                "HWFDSOKM, HWFDSOKH, HWFDSOKN, HWFDSOKU, HWFGDSPH, HWFGDSPT, HWFGDSPR, HWFGDSZY, " & _
                "HWFGDZAR, HWFGDKHM, HWFGDKHH, HWFGDKHN, HWFGDKHU,HWFGDPTK "
    sqlBase = sqlBase & "From " & TbName
    ''6-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Wf_6(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    If DispSXL_GetData(TbName, "select * from TBCME026 " & sqlWhere2, Wf_6(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    ''7. 製品WF仕様_7 (TBCME027) の内容を取得する
    ''7-1.SQLを組み立てる(仕様レコード:操業条件='1')
    
    ''8. 製品WF仕様_8 (TBCME028) の内容を取得する
    ''8-1.SQLを組み立てる(仕様レコード:操業条件='1')
    ' HWFSPVMX,HWFSPVAM 6/22 Yam
    TbName = "TBCME028"
    sqlBase = "select HWFSPVMXN, " & _
        "HWFSPVHT,HWFSPVHS, HWFSPVKN,HWFSPVAMN,HWFSPVSH, HWFSPVST, HWFSPVSI, " & _
        "HWFDLMIN, HWFDLMAX,HWFDLSPH, HWFDLSPT, HWFDLSPI, " & _
        "HWFDLHWT,HWFDLHWS, HWFDLKHN,HWFSPVPS,HWFDLPS "
        sqlBase = sqlBase & "From " & TbName
    ''8-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Wf_8(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    If DispSXL_GetData(TbName, "select * from TBCME028 " & sqlWhere2, Wf_8(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''9. 製品WF仕様_9 (TBCME029) の内容を取得する
    ''9-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME029"
    sqlBase = "select HWFOF1AX,HWFOF1MX,HWFOF1ET,HWFOF1NS,HWFOF1SH,HWFOF1ST,HWFOF1SR,HWFOF1HT,HWFOF1HS,HWFOF1KM,HWFOF1KH, HWFOSF1PTK,HWFOF1SZ, " & _
            "HWFOF2AX,HWFOF2MX,HWFOF2ET,HWFOF2NS,HWFOF2SH,HWFOF2ST,HWFOF2SR,HWFOF2HT,HWFOF2HS,HWFOF2KM,HWFOF2KH,HWFOSF2PTK, HWFOF2SZ, " & _
            "HWFOF3AX,HWFOF3MX,HWFOF3ET,HWFOF3NS,HWFOF3SH,HWFOF3ST,HWFOF3SR,HWFOF3HT,HWFOF3HS,HWFOF3KM,HWFOF3KH,HWFOSF3PTK, HWFOF3SZ, " & _
            "HWFOF4AX,HWFOF4MX,HWFOF4ET,HWFOF4NS,HWFOF4SH,HWFOF4ST,HWFOF4SR,HWFOF4HT,HWFOF4HS,HWFOF4KM,HWFOF4KH,HWFOSF4PTK, HWFOF4SZ, " & _
            "HWFBM1AN,HWFBM1AX,HWFBM1ET,HWFBM1NS,HWFBM1SH,HWFBM1ST,HWFBM1SR,HWFBM1HT,HWFBM1HS,HWFBM1KM,HWFBM1KN,HWFBM1KH,HWFBM1KU, HWFBM1MBP, HWFBM1MCL, HWFBM1SZ, " & _
            "HWFBM2AN,HWFBM2AX,HWFBM2ET,HWFBM2NS,HWFBM2SH,HWFBM2ST,HWFBM2SR,HWFBM2HT,HWFBM2HS,HWFBM2KM,HWFBM2KN,HWFBM2KH,HWFBM2KU, HWFBM2MBP, HWFBM2MCL, HWFBM2SZ, " & _
            "HWFBM3AN,HWFBM3AX,HWFBM3ET,HWFBM3NS,HWFBM3SH,HWFBM3ST,HWFBM3SR,HWFBM3HT,HWFBM3HS,HWFBM3KM,HWFBM3KN,HWFBM3KH,HWFBM3KU, HWFBM3MBP, HWFBM3MCL, HWFBM3SZ, " & _
            "HWFOSPAX, HWFOSPMX, HWFOSPET, " & _
            "HWFOSPNS, " & _
            "HWFOSPSH,HWFOSPST,HWFOSPSR, " & _
            "HWFOSPHT,HWFOSPHS "
    sqlBase = sqlBase & "From " & TbName
    ''9-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF_9(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    If DispSXL_GetData(TbName, "select * from TBCME029 " & sqlWhere2, WF_9(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''10. 製品WF仕様_10 (TBCME047) の内容を取得する
    ''10-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME047"
    sqlBase = "select KPRSPVPUG,KPRSPVPUR,KPRDLPUG,KPRDLPUR "
    sqlBase = sqlBase & "From " & TbName
    ''9-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF_10(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    Set WF_10(2) = WF_10(1)
    
    ''11. 製品WF仕様_11 (TBCME048) の内容を取得する
    ''11-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME048"
    sqlBase = "select HWFSPVPUG,HWFSPVPUR,HWFSPVSTD,HWFDLPUG,HWFDLPUR,HWFNRMX,HWFNRAM,HWFNRPUG,HWFNRPUR, " & _
            "HWFNRSTD,HWFNRSH,HWFNRST,HWFNRSI,HWFNRHT,HWFNRHS,HWFNRKN,HWFNRKM,HWFNRKH, HWFNRKU "
    sqlBase = sqlBase & ",HWFSIRDMX,HWFSIRDSZ,HWFSIRDHT,HWFSIRDHS,HWFSIRDKM,HWFSIRDKN,HWFSIRDKH,HWFSIRDKU,HWFSIRDPS,HWFNRPS "
    sqlBase = sqlBase & "From " & TbName
    ''11-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF_11(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    If DispSXL_GetData(TbName, "select * from TBCME048 " & sqlWhere2, WF_11(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''12. エピ仕様_11 (TBCME049) の内容を取得する
    ''12-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME049"
    sqlBase = "select * "
    sqlBase = sqlBase & "From " & TbName
    ''9-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, EpiKokyaku(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    Set EpiKokyaku(2) = EpiKokyaku(1)
    
    ''12. エピ仕様_12 (TBCME050) の内容を取得する
    ''12-2.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME050"
    sqlBase = "select * "
    sqlBase = sqlBase & "From " & TbName
    ''11-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, EpiSiyou(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    If DispSXL_GetData(TbName, "select * from TBCME050 " & sqlWhere2, EpiSiyou(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    DBDRV_s_cmzcF_cmgc001d_DispTAB3 = FUNCTION_RETURN_SUCCESS

    
    ''13. スラグ仕様_13 (TBCME051) の内容を取得する
    ''13-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME051"
    sqlBase = "select * "
    sqlBase = sqlBase & "From " & TbName
    ''11-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, SlgSiyou(1)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    If DispSXL_GetData(TbName, "select * from TBCME051 " & sqlWhere2, SlgSiyou(2)) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    DBDRV_s_cmzcF_cmgc001d_DispTAB3 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

Public Function AddRecord(rec As c_cmzcrec) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001d_SQL.bas -- Function AddRecord"

    sql = "select * from " & rec.TABLENAME
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    rs.AddNew
    rec.CopyToRs rs
    rs.Update
    rs.Close
    
    AddRecord = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


Public Function GetMaxOpecond(fullHinban As tFullHinban) As String
Dim sql As String
Dim rs As OraDynaset    'RecordSet
    

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001d_SQL.bas -- Function GetMaxOpecond"

    With fullHinban
        sql = "select MAX(opecond) as max_opecond " & _
              "from TBCME018 " & _
              "where (hinban='" & .hinban & "') and (MNOREVNO=" & .MNOREVNO & ") and (FACTORY='" & .FACTORY & "') " & _
              "group by hinban"
    End With
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        GetMaxOpecond = vbNullString
    Else
        GetMaxOpecond = rs("max_opecond")
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

Public Function DBDRV_s_cmzcF_cmfc001d_Exec(insMode As Boolean, targetHinban As tFullHinban, SxlKokyaku_1() As c_cmzcrec, SxlKokyaku_2() As c_cmzcrec, SxlKokyaku_3() As c_cmzcrec, Sxl_1() As c_cmzcrec, Sxl_2() As c_cmzcrec, Sxl_3() As c_cmzcrec, Sxluchigawa() As c_cmzcrec, _
                                            Wf_1() As c_cmzcrec, Wf_2() As c_cmzcrec, Wf_3() As c_cmzcrec, Wf_4() As c_cmzcrec, Wf_5() As c_cmzcrec, _
                                            Wf_6() As c_cmzcrec, Wf_7() As c_cmzcrec, Wf_8() As c_cmzcrec, WF_9() As c_cmzcrec, WF_11() As c_cmzcrec, EpiSiyou() As c_cmzcrec, SlgSiyou() As c_cmzcrec) As FUNCTION_RETURN
Dim sql_top As String
Dim sql_sel As String
Dim sql As String
Dim fld As OraField
Dim rs As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001d_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001d_Exec"

    ''トランザクション開始
    Debug.Print "BeginTrans ======="
    OraDB.BeginTrans

    If insMode Then
        With Sxl_1(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
        With Sxl_2(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
        With Sxl_3(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
                '製WF仕様1～9
        'UpOpecond "TBCME021", targetHinban
        UpOpecond "TBCME022", targetHinban
        UpOpecond "TBCME023", targetHinban
        'UpOpecond "TBCME024", targetHinban
        'UpOpecond "TBCME025", targetHinban
        'UpOpecond "TBCME026", targetHinban
        UpOpecond "TBCME027", targetHinban
        'UpOpecond "TBCME028", targetHinban
        'UpOpecond "TBCME029", targetHinban
        With Wf_1(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
        'With Wf_2(0)
        '    .Fields("HINBAN") = targetHinban.HINBAN
        '    .Fields("MNOREVNO") = targetHinban.MNOREVNO
        '    .Fields("FACTORY") = targetHinban.FACTORY
        '    .Fields("OPECOND") = targetHinban.OPECOND
        '    sql = .SqlInsert
        '    Debug.Print "ExecuteSQL ==========="
        '    Debug.Print sql
        '    OraDB.ExecuteSQL sql
        'End With
        'With Wf_3(0)
        '    .Fields("HINBAN") = targetHinban.HINBAN
        '    .Fields("MNOREVNO") = targetHinban.MNOREVNO
        '    .Fields("FACTORY") = targetHinban.FACTORY
        '    .Fields("OPECOND") = targetHinban.OPECOND
        '    sql = .SqlInsert
        '    Debug.Print "ExecuteSQL ==========="
        '    Debug.Print sql
        '    OraDB.ExecuteSQL sql
        'End With
        With Wf_4(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
        With Wf_5(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
        With Wf_6(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
        'With Wf_7(0)
        '    .Fields("HINBAN") = targetHinban.HINBAN
        '    .Fields("MNOREVNO") = targetHinban.MNOREVNO
        '    .Fields("FACTORY") = targetHinban.FACTORY
        '    .Fields("OPECOND") = targetHinban.OPECOND
        '    sql = .SqlInsert
        '    Debug.Print "ExecuteSQL ==========="
        '    Debug.Print sql
        '    OraDB.ExecuteSQL sql
        'End With
        With Wf_8(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
        With WF_9(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
        With WF_11(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
        With EpiSiyou(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
        With SlgSiyou(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
        
        '結晶内側管理
        With Sxluchigawa(0)
            .Fields("HINBAN") = targetHinban.hinban
            .Fields("MNOREVNO") = targetHinban.MNOREVNO
            .Fields("FACTORY") = targetHinban.FACTORY
            .Fields("OPECOND") = targetHinban.OPECOND
            .Fields("SPECRRNO") = Sxl_1(2).Fields("SPECRRNO")
            .Fields("SXLMCNO") = Sxl_1(2).Fields("SXLMCNO")
            .Fields("WFMCNO") = Sxl_1(2).Fields("WFMCNO")
            sql = .SqlInsert
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            OraDB.ExecuteSQL sql
        End With
    Else
        With Sxl_1(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(Sxl_1(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
        With Sxl_2(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(Sxl_2(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
        With Sxl_3(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(Sxl_3(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
        With Sxluchigawa(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(Sxluchigawa(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                If 0 >= OraDB.ExecuteSQL(sql) Then
                    .SetRecDefault
                    .Fields("SPECRRNO") = Sxl_1(2).Fields("SPECRRNO")
                    .Fields("SXLMCNO") = Sxl_1(2).Fields("SXLMCNO")
                    .Fields("WFMCNO") = Sxl_1(2).Fields("WFMCNO")
                    sql = .SqlInsert
                    Debug.Print "ExecuteSQL ===========(UpdateできなかったのでINSERT)"
                    Debug.Print sql
                    OraDB.ExecuteSQL sql
                End If
            End If
        End With
    
        With Wf_1(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(Wf_1(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
        'With Wf_2(0)
        '    .Fields.Add "HINBAN", targetHinban.HINBAN, ORADB_TEXT, -1
        '    .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
        '    .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
        '    .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
        '    sql = .SqlUpdate(Wf_2(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
        '    Debug.Print "ExecuteSQL ==========="
        '    Debug.Print sql
        '    If sql <> vbNullString Then
        '        OraDB.ExecuteSQL sql
        '    End If
        'End With
        'With Wf_3(0)
        '    .Fields.Add "HINBAN", targetHinban.HINBAN, ORADB_TEXT, -1
        '    .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
        '    .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
        '    .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
        '    sql = .SqlUpdate(Wf_3(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
        '    Debug.Print "ExecuteSQL ==========="
        '    Debug.Print sql
        '    If sql <> vbNullString Then
        '        OraDB.ExecuteSQL sql
        '    End If
        'End With
        With Wf_4(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(Wf_4(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
        With Wf_5(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(Wf_5(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
        With Wf_6(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(Wf_6(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
        'With Wf_7(0)
        '    .Fields.Add "HINBAN", targetHinban.HINBAN, ORADB_TEXT, -1
        '    .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
        '    .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
        '    .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
        '    sql = .SqlUpdate(Wf_7(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
        '    Debug.Print "ExecuteSQL ==========="
        '    Debug.Print sql
        '    If sql <> vbNullString Then
        '        OraDB.ExecuteSQL sql
        '    End If
        'End With
        With Wf_8(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(Wf_8(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
        With WF_9(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(WF_9(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
        With WF_11(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(WF_11(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
        With EpiSiyou(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(EpiSiyou(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
        With SlgSiyou(0)
            .Fields.Add "HINBAN", targetHinban.hinban, ORADB_TEXT, -1
            .Fields.Add "MNOREVNO", targetHinban.MNOREVNO, ORADB_BYTE, -1
            .Fields.Add "FACTORY", targetHinban.FACTORY, ORADB_TEXT, -1
            .Fields.Add "OPECOND", targetHinban.OPECOND, ORADB_TEXT, -1
            sql = .SqlUpdate(SlgSiyou(2), "HINBAN", "MNOREVNO", "FACTORY", "OPECOND")
            Debug.Print "ExecuteSQL ==========="
            Debug.Print sql
            If sql <> vbNullString Then
                OraDB.ExecuteSQL sql
            End If
        End With
    End If
    
    '結晶面傾き更新(TBCME027)・コピー(TBCME022)
    If UPDATE_TBCME022(targetHinban) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    
    DBDRV_s_cmzcF_cmfc001d_Exec = FUNCTION_RETURN_SUCCESS

    ''正常終了ならコミット
    Debug.Print "CommitTrans ======="
    OraDB.CommitTrans

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    ''エラー時はロールバック
    Debug.Print "RollBack ======="
    OraDB.Rollback
    Resume proc_exit
End Function

'概要      :仕様レコードをリビジョンのみあげ、内側管理データとして追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型          ,説明
'          :TbName        ,I  ,String      ,対象テーブル名
'          :targetHinban  ,I  ,tFullHinban ,対象品番
'説明      :製WF仕様1～9のデータコピー用に作成
'履歴      :2001/06/28 作成  野村
Private Sub UpOpecond(ByVal TbName$, targetHinban As tFullHinban)
Dim sql As String
Dim rs As OraDynaset
Dim rec As New c_cmzcrec
 

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001d_SQL.bas -- Sub UpOpecond"

    sql = "Select * From " & TbName & _
          " Where (HINBAN='" & targetHinban.hinban & "')" & _
          " And (MNOREVNO=" & targetHinban.MNOREVNO & ")" & _
          " And (FACTORY='" & targetHinban.FACTORY & "')" & _
          " And (OPECOND='1')"
    On Error Resume Next
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    On Error GoTo proc_err
    If rs.RecordCount = 0 Then
        rec.TABLENAME = TbName
        rec.SetRecDefault
    Else
        rec.CopyFromRs TbName, rs
    End If
    rs.Close
    Set rs = Nothing
    rec("OPECOND") = targetHinban.OPECOND
    sql = rec.SqlInsert
    Debug.Print "ExecuteSQL ==========="
    Debug.Print sql
    OraDB.ExecuteSQL sql

proc_exit:
    '終了
    Set rec = Nothing
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :内側管理データ（規制項目）の抽出
'説明      :Spread_sheet への出力用
'履歴      :2002/12/01  作成  yakimura

Public Function DBDRV_s_cmzcF_cmfc002d_Disp(records() As s_cmzcF_cmfc002d_Disp) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc002d_SQL.bas -- Function DBDRV_s_cmzcF_cmfc002d_Disp"

    
    ''製品仕様管理があってSXL製作条件がないレコード取得
    ''ただし、製作条件付与取消にあるレコードは除く

    sql = "select hinban||ltrim(to_char(mnorevno,'00'))||factory||opecond as hinban12, " & _
          "EPDSETCH, NVL(EPDUP,0) EPDUP, NVL(CUTUNIT,0) CUTUNIT , " & _
          "NVL(TOPREG,0) TOPREG, NVL(TAILREG,0) TAILREG, NVL(BTMSPRT,0) BTMSPRT, REGDATE, UPDDATE From tbcme036 " & _
          "where " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme031)  " & _
          "order by hinban12"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_s_cmzcF_cmfc002d_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Hinban12 = rs("HINBAN12")        ' 品番 + Revision + 工場識別 + 操業条件
            .EPDSETCH = rs("EPDSETCH")        ' EPD 選択エッチ
            .EPDUP = rs("EPDUP")              ' EPD 上限
            .CUTUNIT = rs("CUTUNIT")          ' カット単位
            .TOPREG = rs("TOPREG")            ' TOP規制
            .TAILREG = rs("TAILREG")          ' TAIL規制
            .BTMSPRT = rs("BTMSPRT")          ' ボトム析出規制
            .REGDATE = rs("REGDATE")          ' 登録日付
            .UPDDATE = rs("UPDDATE")          ' 更新日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_s_cmzcF_cmfc002d_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function
Public Function DBDRV_s_cmzcF_cmfc003c_Syounin(targetHinban As tFullHinban, StaffID, Snote, Jnote, sFlg, Optional Tflag As Boolean = False) As FUNCTION_RETURN
Dim sqlWhere As String
Dim sqlSet As String
Dim sql As String
Dim fld As OraField
Dim rs As OraDynaset



    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc003c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc003c_Syounin"
    ''SQLの共通部分を準備する
    With targetHinban
        '仕様のWHERE句
        If sFlg = "9" Then  '廃棄は全てのリビジョン対象
            sqlWhere = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') "
        Else
            sqlWhere = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='" & .OPECOND & "')"
        End If
    End With
    sqlSet = "SSTAFFID = '" & StaffID & "', " & _
            "SYNFLAG = '" & sFlg & "', " & _
            "SYNDATE = SYSDATE  "
    
    ''トランザクション開始
    Debug.Print "BeginTrans ======="
    OraDB.BeginTrans
    
    ''TBCME018
    sql = "update TBCME018 set "
    sql = sql & sqlSet & sqlWhere
    
    Debug.Print sql
    If 0 >= OraDB.ExecuteSQL(sql) Then
        GoTo proc_err
    End If
    ''TBCME019
    sql = "update TBCME019 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME020
    sql = "update TBCME020 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME021
    sql = "update TBCME021 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME022
    sql = "update TBCME022 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME023
    sql = "update TBCME023 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME024
    sql = "update TBCME024 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME025
    sql = "update TBCME025 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME026
    sql = "update TBCME026 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME027
    sql = "update TBCME027 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME028
    sql = "update TBCME028 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME029
    sql = "update TBCME029 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME048
    sql = "update TBCME048 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME050
    sql = "update TBCME050 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME051
    sql = "update TBCME051 set "
    sql = sql & sqlSet & sqlWhere
    
    OraDB.ExecuteSQL sql
    Debug.Print sql
    ''TBCME036
    sql = "update TBCME036 set "
    sql = sql & sqlSet
    If Tflag = False Then
        sql = sql & ", SNOTE = '" & Snote & "', JNOTE = '" & Jnote & "'"
    End If
    sql = sql & sqlWhere
    OraDB.ExecuteSQL sql
    Debug.Print sql
    
    DBDRV_s_cmzcF_cmfc003c_Syounin = FUNCTION_RETURN_SUCCESS

    ''正常終了ならコミット
    Debug.Print "CommitTrans ======="
    OraDB.CommitTrans

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    ''エラー時はロールバック
    Debug.Print "RollBack ======="
    OraDB.Rollback
    Resume proc_exit
End Function

Public Function DBDRV_s_cmzcF_cmgc001d_DispCopy(CopyHinban As tFullHinban, Sxl_1C As c_cmzcrec, Sxl_2C As c_cmzcrec, Sxl_3C As c_cmzcrec, SxluchigawaC As c_cmzcrec, _
                                            Wf_1C As c_cmzcrec, Wf_4C As c_cmzcrec, Wf_5C As c_cmzcrec, _
                                            Wf_6C As c_cmzcrec, Wf_8C As c_cmzcrec, WF_9C As c_cmzcrec, WF_11C As c_cmzcrec, _
                                            EpiC As c_cmzcrec, SlgC As c_cmzcrec, errTbl$) As FUNCTION_RETURN
Dim sql As String
Dim sqlBase As String       'SQLの基本部
Dim sqlWhere As String   'Where句以降
Dim TbName As String        'テーブル名
Dim i As Integer
Dim rs As OraDynaset
        
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001d_SQL.bas -- Function DBDRV_s_cmzcF_cmgc001d_DispCopy"
    
    DBDRV_s_cmzcF_cmgc001d_DispCopy = FUNCTION_RETURN_FAILURE
    errTbl = vbNullString
    
    ''SQLの共通部分を準備する
    With CopyHinban
        '内側管理のWHERE句
        sqlWhere = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='" & .OPECOND & "')" & _
              " Order by OPECOND DESC"
    End With
    
    ''出力データを初期化する
    Set Sxl_1C = New c_cmzcrec
    Set Sxl_2C = New c_cmzcrec
    Set Sxl_3C = New c_cmzcrec
    Set SxluchigawaC = New c_cmzcrec
    Set EpiC = New c_cmzcrec
    
    ''無効品番のチェック（製品仕様SXL1に登録されていること、製作条件付与取消に登録されていないこと）
    With CopyHinban
        sql = "select A.HINBAN from TBCME018 A, TBCME031 B " & _
              "where (A.HINBAN='" & .hinban & "') and (A.MNOREVNO=" & .MNOREVNO & ") and (A.FACTORY='" & .FACTORY & "')" & _
              " and (A.HINBAN=B.HINBAN(+)) and (A.MNOREVNO=B.MNOREVNO(+)) and (A.FACTORY=B.FACTORY(+))" & _
              " and (B.HINBAN is null)"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If (rs Is Nothing) Or (rs.RecordCount = 0) Then
            errTbl = "TBCME031"
            GoTo proc_exit
        End If
    End With
    
    
    ''4. 製品SXL仕様_1 (TBCME018) の内容を取得する
    ''4-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME018"
    sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND," & _
                " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH," & _
                " HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
                " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2," & _
                " HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM, HSXD1CEN," & _
                " HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR," & _
                " HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY, HSXCKHNM, HSXCKHNI," & _
                " HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN," & _
                " HSXCTMIN, HSXCTMAX, HSXCYDIR, HSXCYCEN, HSXCYMIN, HSXCYMAX," & _
                " HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY," & _
                " HSXDPDIR, HSXDPMIN, HSXDPMAX, HSXDWCEN, HSXDWMIN, HSXDWMAX," & _
                " HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX," & _
                " SPECRRNO, SXLMCNO, WFMCNO, MCNO, REGDATE, UPDDATE "
    sqlBase = sqlBase & "From " & TbName
    ''4-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxl_1C) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    ''5. 製品SXL仕様_2 (TBCME019) の内容を取得する
    ''5-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME019"
    sqlBase = "Select HSXTMMAX, HSXTMSPH, HSXTMSPT, HSXTMSPR, HSXTMKHM," & _
                " HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH," & _
                " HSXLTSPT, HSXLTSPI, HSXLTHWT, HSXLTHWS, HSXLTNSW, HSXLTKHM, HSXLTKWY," & _
                " HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN," & _
                " HSXCNMAX, HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS," & _
                " HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
                " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS," & _
                " HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS, HSXONMBP," & _
                " HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX," & _
                " HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH, HSXOS1ST, HSXOS1SI," & _
                " HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS," & _
                " HSXOS2MN, HSXOS2MX, HSXOS2NS, HSXOS2SH, HSXOS2ST, HSXOS2SI," & _
                " HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, HSXTMMAXN "
    sqlBase = sqlBase & "From " & TbName
    ''5-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxl_2C) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''6. 製品SXL仕様_3 (TBCME020) の内容を取得する
    ''6-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    '内側管理レコード
    TbName = "TBCME020"
    sqlBase = "Select HSXDENKU, HSXDENMX, HSXDENMN, HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMX, HSXDVDMN, HSXDVDHT, HSXDVDHS, HSXLDLKU," & _
                " HSXLDLMX, HSXLDLMN, HSXLDLHT, HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH," & _
                " HSXGDKHS, HSXDSOKE, HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS," & _
                " HSXLIFTW, HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
                " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH, HSXOF1KS," & _
                " HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ, HSXOF2KM, HSXOF2KI," & _
                " HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR, HSXOF3HT, HSXOF3HS, HSXOF3SZ," & _
                " HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX, HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT," & _
                " HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS, HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST," & _
                " HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI, HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX," & _
                " HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS, HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET," & _
                " HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST, HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS," & _
                " HSXBM3NS, HSXBM3ET, HSXNOTE, HSXRS1N, HSXRS1Y, HSXRS2N, HSXRS2Y, HSXRS3N, HSXRS3Y, HSXRS4N, HSXRS4Y, HSXRS5N, HSXRS5Y," & _
                " HSXRS6N, HSXRS6Y, HSXRS7N, HSXRS7Y, HSXRS8N, HSXRS8Y, HSXRS9N, HSXRS9Y, HSXRS10N, HSXRS10Y, " & _
                " HSXDVDMNN, HSXDVDMXN, HSXDSONS, HSXCDOPMN, HSXCDOPMX, HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK, " & _
                " HSXBMD1MBP, HSXBMD1MCL, HSXBMD2MBP, HSXBMD2MCL, HSXBMD3MBP, HSXBMD3MCL, HSXDSOPTK,HSXGDPTK, " & _
                " HSXCOSF3PK,HSXCOSF3SH,HSXCOSF3ST,HSXCOSF3SR,HSXCOSF3HT,HSXCOSF3HS,HSXCOSF3SZ,HSXCOSF3NS,HSXCPK,HSXCSZ, " & _
                " HSXCHT,HSXCHS,HSXCJPK,HSXCJNS,HSXCJHT,HSXCJHS,HSXCJLTPK,HSXCJLTNS,HSXCJLTHT,HSXCJLTHS,HSXCJ2PK,HSXCJ2NS," & _
                " HSXCJ2HT,HSXCJ2HS,HSXDSOSZ "
    
    sqlBase = sqlBase & "From " & TbName
    ''6-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxl_3C) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''9. 結晶内側管理 (TBCME036) の内容を取得する
    ''9-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME036"
    'sqlBase = "Select EPDSETCH, EPDUP, CUTUNIT, NVL(TOPREG,0) TOPREG, NVL(TAILREG,0) TAILREG, NVL(BTMSPRT,0) BTMSPRT "
    'sqlBase = sqlBase & ", OTHER1 , OTHER2, OTHERTIME, DCHYUUBU, SNOTE, JNOTE, BLOCKHFLAG "
    'sqlBase = sqlBase & ", OTHER1MAI,WFCUTT,GLASS,SLICEATU,KUMIDOP "
    'sqlBase = sqlBase & ", CUTUNIT,  OTHER1MAI, SKPLACE,COSF3FLAG,HSXDKTMP "
    sqlBase = "Select * "
    sqlBase = sqlBase & " From " & TbName
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, SxluchigawaC) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    TbName = "TBCME021"
    sqlBase = "select  HWFRMIN, HWFRMAX, " & _
              "HWFRSPOH,HWFRSPOT,HWFRSPOI, " & _
              "HWFRHWYT,HWFRHWYS, " & _
              "HWFRKWAY, " & _
              "HWFRKHNM,HWFRKHNN,HWFRKHNH,HWFRKHNU, " & _
              "HWFRSDEV,HWFRMBNP, HWFRMCAL "
    sqlBase = sqlBase & "From " & TbName
    ''1-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Wf_1C) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    ''4. 製品WF仕様_4 (TBCME024) の内容を取得する
    ''4-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME024"
    sqlBase = "select  HWFDZSWY, HWFD1STO, HWFD1STT, HWFD1STG, HWFD2NDO, HWFD2NDC, HWFD2NDT, " & _
              "HWFD3RDO, HWFD3RDT, HWFH2ANO, HWFDZMPS, " & _
              "HWFMKMIN, HWFMKMAX, HWFMKSPH, HWFMKSPT, HWFMKSPR, HWFMKHWT, HWFMKHWS, HWFMKSZY, " & _
              "HWFMKKHM, HWFMKKHN, HWFMKKHH, HWFMKKHU, HWFMKNSW, HWFMKCET "
    sqlBase = sqlBase & "From " & TbName
    ''4-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Wf_4C) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    ''5. 製品WF仕様_5 (TBCME025) の内容を取得する
    ''5-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME025"
    sqlBase = "select  HWFCNMIN, HWFCNMAX, " & _
              "HWFCNSPH,HWFCNSPT,HWFCNSPI, " & _
              "HWFCNHWT,HWFCNHWS, " & _
              "HWFCNKWY, " & _
              "HWFCNKHM,HWFCNKHN,HWFCNKHH,HWFCNKHU, " & _
              "HWFONMIN, HWFONMAX, " & _
              "HWFONSPH,HWFONSPT,HWFONSPI, " & _
              "HWFONHWT,HWFONHWS, " & _
              "HWFONKWY, " & _
              "HWFONKHM,HWFONKHN,HWFONKHH,HWFONKHU, " & _
              "HWFONMBP, HWFONMCL, HWFONSDV "
    sqlBase = sqlBase & ", HWFLTMIN, HWFLTMAX, HWFLTSPH, HWFLTSPT, HWFLTSPI, HWFLTHWT, HWFLTHWS, HWFLTKWY "
    sqlBase = sqlBase & ", HWFCNMIN, HWFCNMAX, HWFCNSPH, HWFCNSPT, HWFCNSPI, HWFCNHWT, HWFCNHWS, HWFCNKWY "
    sqlBase = sqlBase & ", HWFCNKHM, HWFCNKHN, HWFCNKHH, HWFCNKHU "
    sqlBase = sqlBase & ", HWFZOMIN, HWFZOMAX, HWFZOSPH, HWFZOSPT, HWFZOSPI, HWFZOHWT, HWFZOHWS, HWFZOKWY, HWFZONSW  "
    sqlBase = sqlBase & ", HWFOS1MN, HWFOS1MX, HWFOS1NS, HWFOS1SH, HWFOS1ST, HWFOS1SI, HWFOS1HT, HWFOS1HS "
    sqlBase = sqlBase & ", HWFOS1HM, HWFOS1KN, HWFOS1KH, HWFOS1KU "
    sqlBase = sqlBase & ", HWFOS2MN, HWFOS2MX, HWFOS2NS, HWFOS2SH, HWFOS2ST, HWFOS2SI, HWFOS2HT, HWFOS2HS "
    sqlBase = sqlBase & ", HWFOS2KM, HWFOS2KN, HWFOS2KH, HWFOS2KU "
    sqlBase = sqlBase & ", HWFZOKHM, HWFZOKHN, HWFZOKHH, HWFZOKHU "
    sqlBase = sqlBase & "From " & TbName
    ''5-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Wf_5C) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    ''6. 製品WF仕様_6 (TBCME026) の内容を取得する
    ''6-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME026"
    sqlBase = "SELECT HWFDENKU, HWFDENMX, HWFDENMN, HWFDENHT, HWFDENHS, HWFLDLKU, HWFLDLMN, HWFLDLMX, " & _
                "HWFLDLHT, HWFLDLHS, HWFDVDKU, HWFDVDMNN, HWFDVDMXN, HWFDVDHT, HWFDVDHS, " & _
                "HWFDSOMX, HWFDSOMN, HWFDSOAX, HWFDSOAN, HWFDSOHT, HWFDSOHS, HWFDSONWY, HWFDSOPTK, " & _
                "HWFDSOKM, HWFDSOKH, HWFDSOKN, HWFDSOKU, HWFGDSPH, HWFGDSPT, HWFGDSPR, HWFGDSZY, " & _
                "HWFGDZAR, HWFGDKHM, HWFGDKHH, HWFGDKHN, HWFGDKHU,HWFGDPTK "
    sqlBase = sqlBase & "From " & TbName
    ''6-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Wf_6C) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    ''8. 製品WF仕様_8 (TBCME028) の内容を取得する
    ''8-1.SQLを組み立てる(仕様レコード:操業条件='1')
    'HWFSPVMX,HWFSPVAM 6/22 Yam
    TbName = "TBCME028"
    sqlBase = "select HWFSPVMXN, " & _
        "HWFSPVHT,HWFSPVHS, HWFSPVKN, HWFSPVAMN,HWFSPVSH, HWFSPVST, HWFSPVSI,  " & _
        "HWFDLMIN, HWFDLMAX,HWFDLSPH, HWFDLSPT, HWFDLSPI,  " & _
        "HWFDLHWT,HWFDLHWS,HWFDLKHN,HWFSPVPS,HWFDLPS "
        sqlBase = sqlBase & "From " & TbName
    ''8-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Wf_8C) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    ''9. 製品WF仕様_9 (TBCME029) の内容を取得する
    ''9-1.SQLを組み立てる(仕様レコード:操業条件='1')
    TbName = "TBCME029"
    sqlBase = "select HWFOF1AX,HWFOF1MX,HWFOF1ET,HWFOF1NS,HWFOF1SH,HWFOF1ST,HWFOF1SR,HWFOF1HT,HWFOF1HS,HWFOF1KM, HWFOF1KH, HWFOSF1PTK, " & _
            "HWFOF2AX,HWFOF2MX,HWFOF2ET,HWFOF2NS,HWFOF2SH,HWFOF2ST,HWFOF2SR,HWFOF2HT,HWFOF2HS,HWFOF2KM,HWFOF2KH,HWFOSF2PTK, " & _
            "HWFOF3AX,HWFOF3MX,HWFOF3ET,HWFOF3NS,HWFOF3SH,HWFOF3ST,HWFOF3SR,HWFOF3HT,HWFOF3HS,HWFOF3KM,HWFOF3KH,HWFOSF3PTK, " & _
            "HWFOF4AX,HWFOF4MX,HWFOF4ET,HWFOF4NS,HWFOF4SH,HWFOF4ST,HWFOF4SR,HWFOF4HT,HWFOF4HS,HWFOF4KM,HWFOF4KH,HWFOSF4PTK, " & _
            "HWFBM1AN,HWFBM1AX,HWFBM1ET,HWFBM1NS,HWFBM1SH,HWFBM1ST,HWFBM1SR,HWFBM1HT,HWFBM1HS,HWFBM1KM,HWFBM1KN,HWFBM1KH,HWFBM1KU, HWFBM1MBP, HWFBM1MCL, " & _
            "HWFBM2AN,HWFBM2AX,HWFBM2ET,HWFBM2NS,HWFBM2SH,HWFBM2ST,HWFBM2SR,HWFBM2HT,HWFBM2HS,HWFBM2KM,HWFBM2KN,HWFBM2KH,HWFBM2KU, HWFBM2MBP, HWFBM2MCL, " & _
            "HWFBM3AN,HWFBM3AX,HWFBM3ET,HWFBM3NS,HWFBM3SH,HWFBM3ST,HWFBM3SR,HWFBM3HT,HWFBM3HS,HWFBM3KM,HWFBM3KN,HWFBM3KH,HWFBM3KU, HWFBM3MBP, HWFBM3MCL, " & _
            "HWFOSPAX, HWFOSPMX, HWFOSPET, " & _
            "HWFOSPNS, " & _
            "HWFOSPSH,HWFOSPST,HWFOSPSR, " & _
            "HWFOSPHT,HWFOSPHS,HWFOF1SZ,HWFOF2SZ,HWFOF3SZ,HWFOF4SZ,HWFBM1SZ,HWFBM2SZ,HWFBM3SZ "
    sqlBase = sqlBase & "From " & TbName
    ''9-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF_9C) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    ''11. 製品WF仕様_11 (TBCME048) の内容を取得する
    TbName = "TBCME048"
    sqlBase = "select HWFSPVPUG,HWFSPVPUR,HWFSPVSTD,HWFDLPUG,HWFDLPUR,HWFNRMX,HWFNRAM,HWFNRPUG,HWFNRPUR, " & _
            "HWFNRSTD,HWFNRSH,HWFNRST,HWFNRSI,HWFNRHT,HWFNRHS,HWFNRKN,HWFNRKM,HWFNRKH, HWFNRKU "
    sqlBase = sqlBase & ",HWFMSBECN,HWFMSBEMX,HWFMSBEMN,HWFSIRDMX,HWFSIRDSZ,HWFSIRDHT,HWFSIRDHS,HWFSIRDKM,HWFSIRDKN "
    sqlBase = sqlBase & ",HWFSIRDKH,HWFSIRDKU,HWFSIRDPS,HWFNRPS "
    sqlBase = sqlBase & "From " & TbName
    ''11-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WF_11C) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''12. エピ仕様 (TBCME050) の内容を取得する
    TbName = "TBCME050"
    sqlBase = "select * "
    sqlBase = sqlBase & "From " & TbName
    ''11-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, EpiC) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    ''13. スラグ仕様 (TBCME051) の内容を取得する
    TbName = "TBCME051"
    sqlBase = "select * "
    sqlBase = sqlBase & "From " & TbName
    ''13.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, SlgC) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    DBDRV_s_cmzcF_cmgc001d_DispCopy = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :製品仕様画面 エピタブに表示する内容を得る
'ﾊﾟﾗﾒｰﾀ    :変数名         ,IO ,型          ,説明
'          :targetHinban   ,   ,tFullHinban ,品番情報
'          :EP           ,   ,c_cmzcrec   ,
'          :errTbl         ,O  ,String      ,取得できなかったテーブル名
'          :戻り値         ,O  ,FUNCTION_RETURN,検索の成否
'説明      :出力パラメータの配列は、(1 to 9)に製品仕様WFデータ1～9 が入る
'          :呼出側で品番を12桁入力可能なため、該当品番が存在しない場合がありうる
'履歴      :
Public Function DBDRV_s_cmzcF_cmgc001d_DispEP(targetHinban As tFullHinban, EP As c_cmzcrec, errTbl$) As FUNCTION_RETURN
Dim sql As String
Dim sqlBase As String       'SQLの基本部
Dim sqlWhere As String      'Where句以降
Dim TbName As String        'テーブル名
Dim i As Integer
Dim rs As OraDynaset
        
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001d_SQL.bas -- Function DBDRV_s_cmzcF_cmgc001d_DispEP"
    
    DBDRV_s_cmzcF_cmgc001d_DispEP = FUNCTION_RETURN_FAILURE
    errTbl = vbNullString
    
    ''SQLの共通部分を準備する
    With targetHinban
        '仕様のWHERE句(WF仕様は内側管理しない)
        sqlWhere = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='" & .OPECOND & "')"
    End With
    
     Set EP = New c_cmzcrec
    
    ''無効品番のチェック（製品仕様SXL1に登録されていること、製作条件付与取消に登録されていないこと）
    With targetHinban
        sql = "select A.HINBAN from TBCME018 A, TBCME031 B " & _
              "where (A.HINBAN='" & .hinban & "') and (A.MNOREVNO=" & .MNOREVNO & ") and (A.FACTORY='" & .FACTORY & "')" & _
              " and (A.HINBAN=B.HINBAN(+)) and (A.MNOREVNO=B.MNOREVNO(+)) and (A.FACTORY=B.FACTORY(+))" & _
              " and (B.HINBAN is null)"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If (rs Is Nothing) Or (rs.RecordCount = 0) Then
            errTbl = "TBCME031"
            GoTo proc_exit
        End If
    End With
    
    
    '' エピ仕様 (TBCME050) の内容を取得する
    TbName = "TBCME050"
    sqlBase = "select HEPIGKBN,HEPANTNP,HEPANTIM,HEPANGZY,HEPACEN,HEPOF1AX, HEPOF1MX, HEPOF1ET, HEPOF1NS, HEPOF1SZ, " & _
            "HEPOF1SH||HEPOF1ST||HEPOF1SR as HEPOF1S, " & _
            "HEPOF1HT||HEPOF1HS as HEPOF1H, " & _
            "HEPOF1KM||HEPOF1KN||HEPOF1KH||HEPOF1KU as HEPOF1K, " & _
            "HEPOF2AX, HEPOF2MX, HEPOF2ET, HEPOF2NS, HEPOF2SZ, " & _
            "HEPOF2SH||HEPOF2ST||HEPOF2SR as HEPOF2S, " & _
            "HEPOF2HT||HEPOF2HS as HEPOF2H, " & _
            "HEPOF2KM||HEPOF2KN||HEPOF2KH||HEPOF2KU as HEPOF2K, " & _
            "HEPOF3AX, HEPOF3MX, HEPOF3ET, HEPOF3NS, HEPOF3SZ, " & _
            "HEPOF3SH||HEPOF3ST||HEPOF3SR as HEPOF3S, " & _
            "HEPOF3HT||HEPOF3HS as HEPOF3H, " & _
            "HEPOF3KM||HEPOF3KN||HEPOF3KH||HEPOF3KU as HEPOF3K, "
    sqlBase = sqlBase & "HEPBM1AN, HEPBM1AX, HEPBM1ET, HEPBM1NS, HEPBM1SZ, " & _
            "HEPBM1SH||HEPBM1ST||HEPBM1SR as HEPBM1S, " & _
            "HEPBM1HT||HEPBM1HS as HEPBM1H, " & _
            "HEPBM1KM||HEPBM1KN||HEPBM1KH||HEPBM1KU as HEPBM1K, " & _
            "HEPBM2AN, HEPBM2AX, HEPBM2ET, HEPBM2NS, HEPBM2SZ, " & _
            "HEPBM2SH||HEPBM2ST||HEPBM2SR as HEPBM2S, " & _
            "HEPBM2HT||HEPBM2HS as HEPBM2H, " & _
            "HEPBM2KM||HEPBM2KN||HEPBM2KH||HEPBM2KU as HEPBM2K, " & _
            "HEPBM3AN, HEPBM3AX, HEPBM3ET, HEPBM3NS, HEPBM3SZ, " & _
            "HEPBM3SH||HEPBM3ST||HEPBM3SR as HEPBM3S, " & _
            "HEPBM3HT||HEPBM3HS as HEPBM3H, " & _
            "HEPBM3KM||HEPBM3KN||HEPBM3KH||HEPBM3KU as HEPBM3K, " & _
            "HEPOSF1PTK, HEPOSF2PTK, HEPOSF3PTK,  " & _
            "HEPBM1MBP, HEPBM1MCL, HEPBM2MBP, HEPBM2MCL, HEPBM3MBP, HEPBM3MCL "
    sqlBase = sqlBase & "From " & TbName
    ''9-2.データを抽出・格納する
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, EP) = FUNCTION_RETURN_FAILURE Then
        errTbl = TbName
        GoTo proc_exit
    End If
    
    DBDRV_s_cmzcF_cmgc001d_DispEP = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function
Public Function GetIGkbn(fullHinban As tFullHinban) As String
Dim sql As String
Dim rs As OraDynaset    'RecordSet
    

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "cmbc003_SQL.bas -- Function GetIGkbn"

    With fullHinban
        sql = "select HWFIGKBN " & _
              "from TBCME017 " & _
              "where HINBAN = '" & .hinban & "'"
    End With
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        GetIGkbn = vbNullString
    Else
        If IsNull(rs("HWFIGKBN")) Then
            GetIGkbn = vbNullString
        Else
            GetIGkbn = rs("HWFIGKBN")
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function
Public Function GetTBCME001(fullHinban As tFullHinban, sClipDataX As Variant, sClipDataW As Variant, _
                            sShKbn As String, sCsCod As String) As Boolean
Dim sql As String
Dim rs As OraDynaset    'RecordSet
'2007/09 変更
'結晶事業所
'CFCTFLAG1 米沢事業所
'CFCTFLAG2 佐賀事業所
'CFCTFLAG3 伊万里事業所
'CFCTFLAG4 ""
'CFCTFLAG5 ""
'CFCTFLAG6 ""
'WF事業所
'KFCTFLAG1 4棟
'KFCTFLAG2 5棟
'KFCTFLAG3 6棟
'KFCTFLAG4 ""
'KFCTFLAG5 ""
'KFCTFLAG6 ""
'※それぞれ1桁の1が立っていれば該当作業

'2008/05/29 製品区分取得追加

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "cmbc003_SQL.bas -- Function GetTBCME001"
    GetTBCME001 = False
    sClipDataX = vbNullString
    sClipDataW = vbNullString
    
    With fullHinban
        sql = "select CFCTFLAG1, CFCTFLAG2, CFCTFLAG3, CFCTFLAG4, CFCTFLAG5 "
        sql = sql & " ,KFCTFLAG1, KFCTFLAG2, KFCTFLAG3, KFCTFLAG4, KFCTFLAG5,KMGSHKBN,KMGCSCOD "
        sql = sql & "from TBCME001 "
        sql = sql & "where HINBAN = '" & .hinban & "'"
        sql = sql & " And MNOREVNO= '" & .MNOREVNO & "' "
        sql = sql & " And FACTORY= '" & .FACTORY & "'"
        sql = sql & " And OPECOND= '1' "
    End With
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        GetTBCME001 = False
        Exit Function
    Else
        If IsNull(rs("CFCTFLAG1")) = False And rs("CFCTFLAG1") = "1" Then
            sClipDataX = "米沢" & Chr$(13)
        End If
        If IsNull(rs("CFCTFLAG2")) = False And rs("CFCTFLAG2") = "1" Then
            sClipDataX = sClipDataX & "佐賀" & Chr$(13)
        End If
        If IsNull(rs("CFCTFLAG3")) = False And rs("CFCTFLAG3") = "1" Then
            sClipDataX = sClipDataX & "伊万里" & Chr$(13)
        End If
        If IsNull(rs("CFCTFLAG4")) = False And rs("CFCTFLAG4") = "1" Then
            sClipDataX = sClipDataX & rs("CFCTFLAG4") & Chr$(13)
        End If
        If IsNull(rs("CFCTFLAG5")) = False And rs("CFCTFLAG5") = "1" Then
            sClipDataX = sClipDataX & rs("CFCTFLAG5") & Chr$(13)
        End If
        
        If IsNull(rs("KFCTFLAG1")) = False And rs("KFCTFLAG1") = "1" Then
            sClipDataW = "４棟" & Chr$(13)
        End If
        If IsNull(rs("KFCTFLAG2")) = False And rs("KFCTFLAG2") = "1" Then
            sClipDataW = sClipDataW & "５棟" & Chr$(13)
        End If
        If IsNull(rs("KFCTFLAG3")) = False And rs("KFCTFLAG3") = "1" Then
            sClipDataW = sClipDataW & "６棟" & Chr$(13)
        End If
        If IsNull(rs("KFCTFLAG4")) = False And rs("KFCTFLAG4") = "1" Then
            sClipDataW = sClipDataW & rs("KFCTFLAG4") & Chr$(13)
        End If
        If IsNull(rs("KFCTFLAG5")) = False And rs("KFCTFLAG5") = "1" Then
            sClipDataW = sClipDataW & rs("KFCTFLAG5") & Chr$(13)
        End If
        
        If IsNull(rs("KMGSHKBN")) = False Then
            sShKbn = rs("KMGSHKBN")
        Else
            sShKbn = ""
        End If
        
        If IsNull(rs("KMGCSCOD")) = False Then
            sCsCod = rs("KMGCSCOD")
        Else
            sCsCod = ""
        End If
        
    End If
    GetTBCME001 = True
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

Public Function GetKokyakuHin(fullHinban As tFullHinban) As String
Dim sql As String
Dim rs As OraDynaset    'RecordSet
    

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "cmbc003_SQL.bas -- Function GetKokyakuHin"

    With fullHinban
        sql = "select KMGCSHN " & _
              "from TBCME001 " & _
              "where (hinban='" & .hinban & "') and (MNOREVNO=" & .MNOREVNO & ") and (FACTORY='" & .FACTORY & "') "
    End With
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        GetKokyakuHin = vbNullString
    Else
        If IsNull(rs("KMGCSHN")) Then
            GetKokyakuHin = vbNullString
        Else
            GetKokyakuHin = rs("KMGCSHN")
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :汎用コードマスタから、コードNOに対応するコードの一覧を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :CODENO        ,I  ,String           ,コードNO
'          :GPCodeList()  ,O  ,typ_GPCodeMaster ,対応するコードデータの一覧
'          :戻り値        ,O  ,Integer          ,成功/失敗
'説明      :
'履歴      :2001/06/04 作成  野村
Public Function GetGPCodeList_003(ByVal codeNo As String, GPCodeList() As typ_GPCodeMaster_003) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim i As Integer
Dim recCnt As Integer

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''汎用コードマスタから、コードNOに対応するコードの一覧を得る
    sql = "select CODE, CODECONT, CODENAME, INDORDER, KUBUN, READTIME,INDKBN from TBCME033 where (rtrim(CODENO)='" & Trim$(codeNo) & "') order by INDORDER"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.EOF Then
        ''見つからなかったら、0件としてFUNCTION_RETURN_FAILUREを返す
        ReDim GPCodeList(0)
        GetGPCodeList_003 = FUNCTION_RETURN_FAILURE
    Else
        ''見つかったら、その件数分のデータをコピーしてFUNCTION_RETURN_SUCCESSを返す
        recCnt = rs.RecordCount
        ReDim GPCodeList(recCnt)
        For i = 1 To recCnt
            With GPCodeList(i)
                .codeNo = codeNo
                .CODE = rs("CODE")
                .codeCont = rs("CODECONT")
                .codename = rs("CODENAME")
                .INDORDER = rs("INDORDER")
                .KUBUN = rs("KUBUN")
                .READTIME = rs("READTIME")
                If IsNull(rs("INDKBN")) Then
                    .INDKBN = "0"
                Else
                    .INDKBN = rs("INDKBN")
                End If
                rs.MoveNext
            End With
        Next
        GetGPCodeList_003 = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
End Function

'概要      :コードDB書込処理
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :sCodeNo       ,I  ,String    ,コード№
'          :sCode         ,I  ,String    ,種別コード
'          :sIndKbn       ,I  ,String    ,表示区分
'          :戻り値        ,O  ,FUNCTION_RETURN,検索の成否
Public Function UPDATE_TBCME033(ByVal sCodeNo$, ByVal sCode$, ByVal sIndKbn$) As FUNCTION_RETURN
Dim sql As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc003c_SQL.bas -- Function UPDATE_TBCME033"

    
    sql = "UPDATE TBCME033 set " & _
          "INDKBN = '" & sIndKbn & "' ," & _
          "UPDDATE = SYSDATE, " & _
          "STAFFID = '" & f_cmbc003_1.txtStaffID.Text & "' " & _
          "where " & _
          "(trim(CODENO) = '" & sCodeNo & "') and " & _
          "(trim(CODE) = '" & sCode & "') "
    OraDB.ExecuteSQL sql
    If 0 >= OraDB.ExecuteSQL(sql) Then
        GoTo proc_err
    End If
    
    UPDATE_TBCME033 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "==== Error SQL ===="
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
'概要      :結晶面データDB書込処理
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :sCodeNo       ,I  ,String    ,コード№
'          :sCode         ,I  ,String    ,種別コード
'          :sIndKbn       ,I  ,String    ,表示区分
'          :戻り値        ,O  ,FUNCTION_RETURN,検索の成否
Public Function UPDATE_TBCME022(fullHinban As tFullHinban) As FUNCTION_RETURN
Dim sql As String
Dim dData()
Dim rs As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc003c_SQL.bas -- Function UPDATE_TBCME022"
    
    UPDATE_TBCME022 = FUNCTION_RETURN_FAILURE
    
    '画面の取得
    If GetSprData(dData) = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If
    
    
    With fullHinban
    '    sql = "select HWFCSGCEN,HWFCSGMIN,HWFCSGMAX "      '面傾き中心
    '    sql = sql & ",HWFCSXCEN,HWFCSXMIN,HWFCSXMAX "      '面傾きＸ軸（横）
    '    sql = sql & ",HWFCSYCEN,HWFCSYMIN,HWFCSYMAX "      '面傾きＹ軸（縦）
    '    sql = sql & "from TBCME027 "
    '    sql = sql & "where HINBAN = '" & .hinban & "'"
    '    sql = sql & " And MNOREVNO= '" & .MNOREVNO & "' "
    '    sql = sql & " And FACTORY= '" & .FACTORY & "'"
    '    sql = sql & " And OPECOND= '" & .OPECOND & "' "
    '    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '    If rs.RecordCount = 0 Then
    '        UPDATE_TBCME022 = False
    '        Exit Function
    '    Else
    '        If IsNull(rs("HWFCSGCEN")) Then
    '            iCenG = "Null"
    '        Else
    '            iCenG = CDbl(rs("HWFCSGCEN"))
    '        End If
    '        If IsNull(rs("HWFCSGMIN")) Then
    '            iMinG = "Null"
    '        Else
    '            iMinG = CDbl(rs("HWFCSGMIN"))
    '        End If
    '        If IsNull(rs("HWFCSGMAX")) Then
    '            iMaxG = "Null"
    '        Else
    '            iMaxG = CDbl(rs("HWFCSGMAX"))
    '        End If
    '
    '        If IsNull(rs("HWFCSXCEN")) Then
    '            iCenX = "Null"
    '        Else
    '            iCenX = CDbl(rs("HWFCSXCEN"))
    '        End If
    '        If IsNull(rs("HWFCSXMIN")) Then
    '            iMinX = "Null"
    '        Else
    '            iMinX = CDbl(rs("HWFCSXMIN"))
    '        End If
    '        If IsNull(rs("HWFCSXMAX")) Then
    '            iMaxX = "Null"
    '        Else
    '            iMaxX = CDbl(rs("HWFCSXMAX"))
    '        End If
    '
    '        If IsNull(rs("HWFCSYCEN")) Then
    '            iCenY = "Null"
    '        Else
    '            iCenY = CDbl(rs("HWFCSYCEN"))
    '        End If
    '        If IsNull(rs("HWFCSYMIN")) Then
    '            iMinY = "Null"
    '        Else
    '            iMinY = CDbl(rs("HWFCSYMIN"))
    '        End If
    '        If IsNull(rs("HWFCSYMAX")) Then
    '            iMaxY = "Null"
    '        Else
    '            iMaxY = CDbl(rs("HWFCSYMAX"))
    '        End If
    '
    '    End If
    '    rs.Close
        
        
        sql = "UPDATE TBCME027 set "
        sql = sql & "HWFCSGCEN = " & dData(0) & " ,"
        sql = sql & "HWFCSGMIN = " & dData(1) & " ,"
        sql = sql & "HWFCSGMAX = " & dData(2) & " ,"
        sql = sql & "HWFCSXCEN = " & dData(3) & " ,"
        sql = sql & "HWFCSXMIN = " & dData(4) & " ,"
        sql = sql & "HWFCSXMAX = " & dData(5) & " ,"
        sql = sql & "HWFCSYCEN = " & dData(6) & " ,"
        sql = sql & "HWFCSYMIN = " & dData(7) & " ,"
        sql = sql & "HWFCSYMAX = " & dData(8) & " ,"
        sql = sql & "UPDDATE = SYSDATE, "
        sql = sql & "KSTAFFID = '" & f_cmbc003_1.txtStaffID.Text & "' "
        sql = sql & "where "
        sql = sql & "     HINBAN = '" & .hinban & "'"
        sql = sql & " And MNOREVNO= " & .MNOREVNO & " "
        sql = sql & " And FACTORY= '" & .FACTORY & "'"
        sql = sql & " And OPECOND= '" & .OPECOND & "' "
    
        OraDB.ExecuteSQL sql
        If 0 >= OraDB.ExecuteSQL(sql) Then
            GoTo proc_err
        End If
    
        sql = ""
        sql = "UPDATE TBCME022 set "
        sql = sql & "HWFCSCEN = " & dData(0) & " ,"
        sql = sql & "HWFCSMIN = " & dData(1) & " ,"
        sql = sql & "HWFCSMAX = " & dData(2) & " ,"
        sql = sql & "HWFCTCEN = " & dData(3) & " ,"
        sql = sql & "HWFCTMIN = " & dData(4) & " ,"
        sql = sql & "HWFCTMAX = " & dData(5) & " ,"
        sql = sql & "HWFCYCEN = " & dData(6) & " ,"
        sql = sql & "HWFCYMIN = " & dData(7) & " ,"
        sql = sql & "HWFCYMAX = " & dData(8) & " ,"
        sql = sql & "UPDDATE = SYSDATE, "
        sql = sql & "KSTAFFID = '" & f_cmbc003_1.txtStaffID.Text & "' "
        sql = sql & "where "
        sql = sql & "     HINBAN = '" & .hinban & "'"
        sql = sql & " And MNOREVNO= " & .MNOREVNO & " "
        sql = sql & " And FACTORY= '" & .FACTORY & "'"
        sql = sql & " And OPECOND= '" & .OPECOND & "' "
    
        OraDB.ExecuteSQL sql
        If 0 >= OraDB.ExecuteSQL(sql) Then
            GoTo proc_err
        End If
    
    End With
    
    
    UPDATE_TBCME022 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "==== Error SQL ===="
    Debug.Print sql
    gErr.HandleError
    OraDB.Rollback
    Resume proc_exit
End Function


