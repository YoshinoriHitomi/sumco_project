Attribute VB_Name = "s_cmbc001c_SQL"
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
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DispSXL_GetData"
    
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
'          :SxlKokyaku_1 ,   ,c_cmzcrec   ,顧客仕様SXL 1 の内容
'          :SxlKokyaku_2 ,   ,c_cmzcrec   ,顧客仕様SXL 2 の内容
'          :SxlKokyaku_3 ,   ,c_cmzcrec   ,顧客仕様SXL 3 の内容
'          :Sxl_1        ,   ,c_cmzcrec   ,製品仕様SXL_1 の内容
'          :Sxl_2        ,   ,c_cmzcrec   ,製品仕様SXL_2 の内容
'          :Sxl_3        ,   ,c_cmzcrec   ,製品仕様SXL_3 の内容
'          :WfKokyaku_2  ,   ,c_cmzcrec   ,顧客仕様WF 2 の内容
'          :WfKokyaku_8  ,   ,c_cmzcrec   ,顧客仕様WF 8 の内容
'          :SxlUchigawa  ,   ,c_cmzcrec   ,内側 の内容
'          :戻り値         ,O  ,FUNCTION_RETURN,検索の成否
'説明      :各出力パラメータの配列は、(1)に仕様データ (2)に指定操業条件のデータ が入る
'          :呼出側で品番を12桁入力可能なため、該当品番が存在しない場合がありうる
'履歴      :2001/06/08 作成  野村
Public Function DBDRV_s_cmzcF_cmfc001c_DispSXL(targetHinban As tFullHinban, SxlKokyaku_1 As c_cmzcrec, SxlKokyaku_2 As c_cmzcrec, SxlKokyaku_3 As c_cmzcrec, Sxl_1 As c_cmzcrec, Sxl_2 As c_cmzcrec, Sxl_3 As c_cmzcrec, WfKokyaku_2 As c_cmzcrec, WfKokyaku_8 As c_cmzcrec, SxlUchigawa As c_cmzcrec) As FUNCTION_RETURN
Dim sql As String
Dim sqlBase As String       'SQLの基本部
Dim sqlWhere As String      'Where句以降
Dim TbName As String        'テーブル名
Dim i As Integer
Dim HWFMKnSI(4) As Double   'LPDサイズ(0:結果 1～4:候補)
Dim HWFMKnMX(4) As Integer  'LPD上限(0:結果 1～4:候補)
Dim rs As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001c_DispSXL"
        
    DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
    
    ''SQLの共通部分を準備する
    With targetHinban
        'WHERE句
        sqlWhere = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='" & .OPECOND & "')"
    End With
    
    ''出力データを初期化する
    Set SxlKokyaku_1 = New c_cmzcrec
    Set SxlKokyaku_2 = New c_cmzcrec
    Set SxlKokyaku_3 = New c_cmzcrec
    Set Sxl_1 = New c_cmzcrec
    Set Sxl_2 = New c_cmzcrec
    Set Sxl_3 = New c_cmzcrec
    Set WfKokyaku_2 = New c_cmzcrec
    Set WfKokyaku_8 = New c_cmzcrec
    Set SxlUchigawa = New c_cmzcrec
    
    ''無効品番のチェック（製品仕様SXL1に登録されていること、製作条件付与取消に登録されていないこと）
    With targetHinban
        sql = "select A.HINBAN from TBCME018 A, TBCME031 B " & _
              "where (A.HINBAN='" & .hinban & "') and (A.MNOREVNO=" & .MNOREVNO & ") and (A.FACTORY='" & .FACTORY & "')" & _
              " and (A.HINBAN=B.HINBAN(+)) and (A.MNOREVNO=B.MNOREVNO(+)) and (A.FACTORY=B.FACTORY(+))" & _
              " and (B.HINBAN is null)"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If (rs Is Nothing) Or (rs.RecordCount = 0) Then
            GoTo proc_exit
        End If
    End With
    
    ''1. 顧客SXL仕様_1 (TBCME005) の内容を取得する
    ''1-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME005"
    sqlBase = "Select KSXTYPKB, KSXRUNIT, KSXRKKBN, KSXD1KBN, KSXD2KBN," & _
                " KSXDFKBN, KSXDPKBN, KSXDWKBN, KSXDDKBN, KSXDAKBN "
    sqlBase = sqlBase & "From " & TbName
    ''1-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, SxlKokyaku_1) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''2. 顧客SXL仕様_2 (TBCME006) の内容を取得する
    ''2-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME006"
    sqlBase = "Select KSXTMKBN, KSXLTUNT, KSXLTKBN, KSXCNIND, KSXCNUNT," & _
                " KSXCNKBN, KSXONIND, KSXONUNT, KSXONKBN "
    sqlBase = sqlBase & "From " & TbName
    ''2-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, SxlKokyaku_2) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''3. 顧客SXL仕様_3 (TBCME007) の内容を取得する
    ''3-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME007"
    sqlBase = "Select KSXOF1KBN, KSXOF1FGS, KSXOF1SO1, KSXOF1ST1, KSXOF2KB, KSXOF2GS, " & _
                "KSXOF2O1, KSXOF2ST, KSXBMKBN, KSXBMFGS, KSXBM2KB, KSXBM2GS "
    sqlBase = sqlBase & "From " & TbName
    ''3-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, SxlKokyaku_3) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''4. 製品SXL仕様_1 (TBCME018) の内容を取得する
    ''4-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME018"
    sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND," & _
                " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX," & _
                " HSXRSPOH||HSXRSPOT||HSXRSPOI as HSXSPO," & _
                " HSXRHWYT||HSXRHWYS as HSXRHWY," & _
                " HSXRKWAY," & _
                " HSXRKHNM||HSXRKHNI||HSXRKHNH||HSXRKHNS as HSXRKHN," & _
                " HSXRMCAL, HSXRMBNP, HSXRMCL2," & _
                " HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM, HSXD1CEN," & _
                " HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR," & _
                " HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
                " HSXCKHNM||HSXCKHNI||HSXCKHNH||HSXCKHNS as HSXCKHN," & _
                " HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN," & _
                " HSXCTMIN, HSXCTMAX, HSXCYDIR, HSXCYCEN, HSXCYMIN, HSXCYMAX," & _
                " HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY," & _
                " HSXDPDIR, HSXDPMIN, HSXDPMAX, HSXDWCEN, HSXDWMIN, HSXDWMAX," & _
                " HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX," & _
                " SPECRRNO, SXLMCNO, WFMCNO "
    sqlBase = sqlBase & "From " & TbName
    ''4-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxl_1) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''5. 製品SXL仕様_2 (TBCME019) の内容を取得する
    ''5-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME019"
    sqlBase = "Select HSXTMMAX," & _
                " HSXTMSPH||HSXTMSPT||HSXTMSPR as HSXTMSP," & _
                " HSXTMKHM||HSXTMKHI||HSXTMKHH||HSXTMKHS as HSXTMKH," & _
                " HSXLTMIN, HSXLTMAX," & _
                " HSXLTSPH||HSXLTSPT||HSXLTSPI as HSXLTSP," & _
                " HSXLTHWT||HSXLTHWS as HSXLTHW," & _
                " HSXLTNSW," & _
                " HSXLTKHM||HSXLTKHI||HSXLTKHH||HSXLTKHS as HSXLTKH," & _
                " HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
                " HSXCNSPH||HSXCNSPT||HSXCNSPI as HSXCNSP," & _
                " HSXCNHWT||HSXCNHWS as HSXCNHW," & _
                " HSXCNKWY," & _
                " HSXCNKHM||HSXCNKHI||HSXCNKHH||HSXCNKHS as HSXCNKH," & _
                " HSXONMIN, HSXONMAX,"
    sqlBase = sqlBase & " HSXONSPH||HSXONSPT||HSXONSPI as HSXONSP," & _
                " HSXONHWT||HSXONHWS as HSXONHW," & _
                " HSXONKWY," & _
                " HSXONKHM||HSXONKHI||HSXONKHH||HSXONKHS as HSXONKH," & _
                " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS," & _
                " HSXOS1SH||HSXOS1ST||HSXOS1SI as HSXOS1S," & _
                " HSXOS1HT||HSXOS1HS as HSXOS1H," & _
                " HSXOS1HM||HSXOS1KI||HSXOS1KH||HSXOS1KS as HSXOS1K," & _
                " HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
                " HSXOS2SH||HSXOS2ST||HSXOS2SI as HSXOS2S," & _
                " HSXOS2HT||HSXOS2HS as HSXOS2H," & _
                " HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU as HSXOS2K "
    sqlBase = sqlBase & "From " & TbName
    ''5-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxl_2) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''6. 製品SXL仕様_3 (TBCME020) の内容を取得する
    ''6-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME020"
    sqlBase = "Select HSXDENKU," & _
              " HSXDENHT||HSXDENHS as HSXDENH," & _
              " HSXDVDKU," & _
              " HSXDVDHT||HSXDVDHS as HSXDVDH," & _
              " HSXLDLKU," & _
              " HSXLDLHT||HSXLDLHS as HSXLDLH," & _
              " HSXGDSPH||HSXGDSPT||HSXGDSPR as HSXGDSP," & _
              " HSXGDSZY," & _
              " HSXGDKHM||HSXGDKHI||HSXGDKHH||HSXGDKHS as HSXGDKH," & _
              " HSXDSOHT||HSXDSOHS as HSXDSOH," & _
              " HSXDSOKM||HSXDSOKI||HSXDSOKH||HSXDSOKS as HSXDSOK," & _
              " HSXLIFTW, HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDPNI, HSXGSFIN, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SZ," & _
              " HSXOF1SH||HSXOF1ST||HSXOF1SR as HSXOF1S," & _
              " HSXOF1HT||HSXOF1HS as HSXOF1H," & _
              " HSXOF1NS, HSXOF1ET,"
    sqlBase = sqlBase & " HSXOF1KM||HSXOF1KI||HSXOF1KH||HSXOF1KS as HSXOF1K," & _
              " HSXOF2AX, HSXOF2MX, HSXOF2SZ," & _
              " HSXOF2SH||HSXOF2ST||HSXOF2SR as HSXOF2S," & _
              " HSXOF2HT||HSXOF2HS as HSXOF2H," & _
              " HSXOF2NS, HSXOF2ET," & _
              " HSXOF2KM||HSXOF2KI||HSXOF2KH||HSXOF2KS as HSXOF2K," & _
              " HSXOF3AX, HSXOF3MX, HSXOF3SZ," & _
              " HSXOF3SH||HSXOF3ST||HSXOF3SR as HSXOF3S," & _
              " HSXOF3HT||HSXOF3HS as HSXOF3H," & _
              " HSXOF3NS, HSXOF3ET," & _
              " HSXOF3KM||HSXOF3KI||HSXOF3KH||HSXOF3KS as HSXOF3K," & _
              " HSXOF4AX, HSXOF4MX, HSXOF4SZ," & _
              " HSXOF4SH||HSXOF4ST||HSXOF4SR as HSXOF4S," & _
              " HSXOF4HT||HSXOF4HS as HSXOF4H," & _
              " HSXOF4NS, HSXOF4ET," & _
              " HSXOF4KM||HSXOF4KI||HSXOF4KH||HSXOF4KS as HSXOF4K,"
    sqlBase = sqlBase & " HSXBM1SZ, HSXBM1AN, HSXBM1AX," & _
              " HSXBM1SH||HSXBM1ST||HSXBM1SR as HSXBM1S," & _
              " HSXBM1HT||HSXBM1HS as HSXBM1H," & _
              " HSXBM1NS,HSXBM1ET," & _
              " HSXBM1KM||HSXBM1KI||HSXBM1KH||HSXBM1KS as HSXBM1K," & _
              " HSXBM2SZ, HSXBM2AN, HSXBM2AX," & _
              " HSXBM2SH||HSXBM2ST||HSXBM2SR as HSXBM2S," & _
              " HSXBM2HT||HSXBM2HS as HSXBM2H," & _
              " HSXBM2NS,HSXBM2ET," & _
              " HSXBM2KM||HSXBM2KI||HSXBM2KH||HSXBM2KS as HSXBM2K," & _
              " HSXBM3SZ, HSXBM3AN, HSXBM3AX," & _
              " HSXBM3SH||HSXBM3ST||HSXBM3SR as HSXBM3S," & _
              " HSXBM3HT||HSXBM3HS as HSXBM3H," & _
              " HSXBM3NS,HSXBM3ET," & _
              " HSXBM3KM||HSXBM3KI||HSXBM3KH||HSXBM3KS as HSXBM3K "
    sqlBase = sqlBase & "From " & TbName
    ''6-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxl_3) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''7. 顧客仕様WFﾃﾞｰﾀ2 (TBCME009) の内容を取得する
    ''7-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME009"
    sqlBase = "Select KPRDFORM "
    sqlBase = sqlBase & "From " & TbName
    ''7-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WfKokyaku_2) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''8. 顧客仕様WFﾃﾞｰﾀ8 (TBCME028) の内容を取得する
    ''8-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME028"
    sqlBase = "Select HWFMK1SI, HWFMK2SI, HWFMK3SI, HWFMK4SI, HWFMK1MX, HWFMK2MX, HWFMK3MX, HWFMK4MX "
    sqlBase = sqlBase & "From " & TbName
    ''8-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WfKokyaku_8) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''9. 結晶内側管理 (TBCME036) の内容を取得する
    ''9-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME036"
    sqlBase = "Select EPDSETCH, EPDUP, CUTUNIT "
    sqlBase = sqlBase & "From " & TbName
    ''9-2.データを抽出・格納する
    '仕様レコード
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, SxlUchigawa) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''10. LPDサイズ・LPD上限の表示値を設定する
    With WfKokyaku_8
        'HWFMK1SI～HWFMK4SI, HWFMK1MX～HWFMK4MX を取得
        HWFMKnSI(1) = .Fields.GetValueOrDefault("HWFMK1SI", 0#)
        HWFMKnMX(1) = .Fields.GetValueOrDefault("HWFMK1MX", 0)
        HWFMKnSI(2) = .Fields.GetValueOrDefault("HWFMK2SI", 0#)
        HWFMKnMX(2) = .Fields.GetValueOrDefault("HWFMK2MX", 0)
        HWFMKnSI(3) = .Fields.GetValueOrDefault("HWFMK3SI", 0#)
        HWFMKnMX(3) = .Fields.GetValueOrDefault("HWFMK3MX", 0)
        HWFMKnSI(4) = .Fields.GetValueOrDefault("HWFMK4SI", 0#)
        HWFMKnMX(4) = .Fields.GetValueOrDefault("HWFMK4MX", 0)
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
    WfKokyaku_8.Fields.Add "HWFMKnSI", HWFMKnSI(0), ORADB_DOUBLE, -1
    WfKokyaku_8.Fields.Add "HWFMKnMX", HWFMKnMX(0), ORADB_INTEGER, -1
    
    DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :製作条件の概要を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :jokenNo       ,I  ,String    ,製作条件番号
'          :rec           ,O  ,c_cmzcrec ,概要
'          :戻り値        ,O  ,FUNCTION_RETURN,検索の成否
'説明      :
'履歴      :2001/06/08 作成  野村
Public Function DBDRV_s_cmzcF_cmfc001c_GetSJoken(ByVal jokenNo$, rec As c_cmzcrec) As FUNCTION_RETURN
Dim sql As String           'SQL
Dim TbName As String        'テーブル名

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001c_GetSJoken"

    ''出力データを初期化する
    Set rec = New c_cmzcrec
    
    ''1. 製作条件 (TBCMB012) の内容を取得する
    ''1-1.SQLを組み立てる
    TbName = "TBCMB012"
    sql = "Select MODEL, RTBSIZE, CHARGE, HZTYPE, UPSPDTYP, MAGTYPE" & _
          " From " & TbName & _
          " Where (rtrim(MKCONDNO)='" & jokenNo & "')"
    ''1-2.データを抽出・格納する
    DBDRV_s_cmzcF_cmfc001c_GetSJoken = DispSXL_GetData(TbName, sql, rec)

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :製作条件番号に対応したPGIDの一覧を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :jokenNo       ,I  ,String    ,製作条件番号
'          :PGIDs()       ,O  ,String    ,対応するPGIDの一覧
'          :戻り値        ,O  ,FUNCTION_RETURN,検索の成否
'説明      :
'履歴      :2001/06/08 作成  野村
Public Function DBDRV_s_cmzcF_cmfc001c_GetPGID(ByVal jokenNo$, PGIDs() As String) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001c_GetPGID"

    sql = "Select PGIDNO " & _
          "From TBCMB013 " & _
          "Where (rtrim(MKCONDNO)='" & jokenNo & "')"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim PGIDs(0)
        DBDRV_s_cmzcF_cmfc001c_GetPGID = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim PGIDs(recCnt)
    For i = 1 To recCnt
        PGIDs(i) = rs("PGIDNO")     ' PG-IDNo
        rs.MoveNext
    Next
    rs.Close

    DBDRV_s_cmzcF_cmfc001c_GetPGID = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :[実行]時データ書込処理
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :IraiNo        ,I  ,String    ,仕様登録依頼番号
'          :SXLMCNO       ,I  ,String    ,SXL製作条件(仕様側)
'          :WFMCNO        ,I  ,String    ,WF製作条件(仕様側)
'          :Hinban12      ,I  ,String    ,12桁品番
'          :SJokenNo      ,I  ,String    ,製作条件番号
'          :Hikiage       ,I  ,String    ,引上方法
'          :StaffID       ,I  ,String    ,担当者コード
'          :戻り値        ,O  ,FUNCTION_RETURN,検索の成否
'説明      :
'履歴      :2001/06/08 作成  野村
Public Function DBDRV_s_cmzcF_cmfc001c_Exec(ByVal IraiNo$, ByVal SXLMCNO$, ByVal WFMCNO$, ByVal Hinban12$, ByVal SJokenNo$, ByVal Hikiage$, ByVal StaffID$) As FUNCTION_RETURN
Dim sql_top As String
Dim sql_sel As String
Dim sql As String
Dim fld As OraField
Dim rs As OraDynaset
Dim fullHinban As tFullHinban


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001c_Exec"

    If Len(Hinban12) <> 12 Then
        DBDRV_s_cmzcF_cmfc001c_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    With fullHinban
        .hinban = Left$(Hinban12, 8)
        .MNOREVNO = Val(Mid$(Hinban12, 9, 2))
        .FACTORY = Mid$(Hinban12, 11)
        .OPECOND = Right$(Hinban12, 1)
    End With
    sql = "insert into TBCME030 " & _
          "(HINBAN, MNOREVNO, FACTORY, OPECOND, SSXLIFTW, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO, " & _
          "STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE) " & _
          "values ("
    sql = sql & "'" & fullHinban.hinban & "', "     ' 品番
    sql = sql & fullHinban.MNOREVNO & ", "          ' 製品番号改訂番号
    sql = sql & "'" & fullHinban.FACTORY & "', "    ' 工場
    sql = sql & "'" & fullHinban.OPECOND & "', "    ' 操業条件
    sql = sql & "'" & Hikiage & "', "               ' 製ＳＸ引上方法
    sql = sql & "' ', "                             ' Ｉ／Ｆ区分
    sql = sql & "' ', "                             ' 処理区分
    sql = sql & "'" & IraiNo & "', "                ' 仕様登録依頼番号
    sql = sql & "'" & SXLMCNO & "', "               ' ＳＸＬ製作条件番号
    sql = sql & "'" & WFMCNO & "', "                ' ＷＦ製作条件番号
    sql = sql & "'" & StaffID & "', "               ' 社員ID
    sql = sql & "SYSDATE, "                         ' 登録日付
    sql = sql & "SYSDATE, "                         ' 更新日付
    sql = sql & "'0', "                             ' 送信フラグ
    sql = sql & "SYSDATE "                          ' 送信日付
    sql = sql & ")"
    If 0 >= OraDB.ExecuteSQL(sql) Then
        GoTo proc_err
    End If
    
    ''品番データに製作条件Noを書き込む(受取仕様であるリビジョン１を除く各リビジョンに）
    sql = "update TBCME018 set " & _
          "MCNO = '" & SJokenNo & "' " & _
          "where " & _
          "(HINBAN = '" & fullHinban.hinban & "') and " & _
          "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
          "(FACTORY = '" & fullHinban.FACTORY & "')"
    OraDB.ExecuteSQL sql
    If 0 >= OraDB.ExecuteSQL(sql) Then
        GoTo proc_err
    End If
    
    DBDRV_s_cmzcF_cmfc001c_Exec = FUNCTION_RETURN_SUCCESS

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

