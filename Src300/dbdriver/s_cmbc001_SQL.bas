Attribute VB_Name = "s_cmbc001_SQL"
' TBCME017 (製品仕様管理)より
Public Type s_cmzcF_cmfc001b_Disp
    '製品仕様管理
    Hinban12 As String * 12         ' 品番
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    REGDATE As Date                 ' 登録日付
    SENDDATE As Date                ' 送信日付
    SENDFLAG As String              ' フラグ
    TOUROKU As Date
End Type

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

'                                     2001/06/11
'================================================
' DBアクセス関数
' 定義内容: TBCMB011 (PG-ID管理)
' 参照　　: 060200_全テーブル
'================================================
#If False Then
'------------------------------------------------
' ユーザ定義型の宣言
'------------------------------------------------
Public Type typ_cmbc001d_Disp
    PGID As String * 10             ' PG-ID
    HZPART As String * 4            ' HZパーツ
    HZPTRN As String * 2            ' HZパターン
    SPACER As String * 5            ' スペーサ
    UPRING As String * 5            ' アッパーリング
    CHARGE As Long                  ' チャージ量
    RTBPOS As Integer               ' ルツボ位置
    RTBSIZE As String * 2           ' ルツボサイズ
    GAP As Integer                  ' ギャップ
    UPDM As Integer                 ' 引上直径
    UPLENGTH As Integer             ' 引上長（全長）
    UPRC As Integer                 ' 引上（RC）
    RFRNEED As String * 1           ' リフラクタ要否
    UPSPIN As Double                ' 上軸回転数
    DOWNSPIN As Double              ' 下軸回転数
    ROPRESS As String * 8           ' 炉内圧
    ARUGON As String * 7            ' アルゴン量
    AIMOIMIN As Double              ' ねらいOi（MIN)
    AIMOIMAX As Double              ' ねらいOi（MAX)
    HCCLASS As String * 7           ' HC種類
    HC As String * 3                ' HC
    AVEUPSPD As Double              ' 平均引上速度
    UPCNTL As String * 1            ' 引上制御
    BTMSHAPE As String * 1          ' ボトム形状
    MAGSTR As Long                  ' 磁場強度
    MAGPOS As Long                  ' 磁場位置
    CONDGRT As String * 10          ' 条件保証登録
    MODEL As String * 4             ' 機種
    UPMETHOD As String * 1          ' 引上方法
    UPCLASS As String * 2           ' 引上区分
    UPNUM As String * 1             ' 引上本数
    OPETIME As Long                 ' 運転時間
    WTRCOOL As String * 1           ' 水冷管要否
    PGID2 As String * 8             ' PG-ID（一本引）
    RCPT1 As String * 3             ' 対応レシピNo（T1)
    RCPT2 As String * 3             ' 対応レシピNo（T2)
    RCPT3 As String * 3             ' 対応レシピNo（T3)
    RCPT4 As String * 3             ' 対応レシピNo（T4)
    RCPT5 As String * 3             ' 対応レシピNo（T5)
    CNTL1 As String * 1             ' 制御項目（1）
    CNTL2 As String * 1             ' 制御項目（2）
    CNTL3 As String * 1             ' 制御項目（3）
    CNTL4 As String * 1             ' 制御項目（4）
    CNTL5 As String * 1             ' 制御項目（5）
    CNTL6 As String * 1             ' 制御項目（6）
    CNTL7 As String * 1             ' 制御項目（7）
    CNTL8 As String * 1             ' 制御項目（8）
    CNTL9 As String * 1             ' 制御項目（9）
    CNTL10 As String * 1            ' 制御項目（10）
    CNTL11 As String * 1            ' 制御項目（11）
    CNTL12 As String * 1            ' 制御項目（12）
    CNTL13 As String * 1            ' 制御項目（13）
    CNTL14 As String * 1            ' 制御項目（14）
    CNTL15 As String * 1            ' 制御項目（15）
    RUNCOND1 As String              ' 運転条件１
    RUNCOND2 As String              ' 運転条件２
'    TSTAFFID As String * 5          ' 登録社員ID
'    REGDATE As Date                 ' 登録日付
'    KSTAFFID As String * 8          ' 更新社員ID
'    UPDDATE As Date                 ' 更新日付
'    SENDFLAG As String * 1          ' 送信フラグ
'    SENDDATE As Date                ' 送信日付
End Type
#End If



'概要      :引上指示番号の連番部に値を加える
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :sijiNo        ,I  ,String    ,元の引上指示番号
'          :addVal        ,I  ,Integer   ,加算値(マイナスも可)
'          :戻り値        ,O  ,String    ,加算後の引上指示番号
'説明      :
'履歴      :2001/07/09 作成  野村 (2002/07 s_cmzcF_cmhc001d_SQL.basより移動)
Public Function SijiNoAdd(sijiNo$, addVal%) As String
Dim seq As Integer
Dim newNo As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmhc001d_SQL.bas -- Function SijiNoAdd"

    seq = val(Mid$(sijiNo, 5, 3))
    SijiNoAdd = Left$(sijiNo, 4) & Format$(seq + addVal, "000") & Mid$(sijiNo, 8)

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'現在の仕様登録では未使用
Public Function DBDRV_s_cmzcF_cmfc001b_Disp(records() As s_cmzcF_cmfc001b_Disp) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001b_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001b_Disp"
    
    ''製品仕様管理があってSXL製作条件がないレコード取得（品番、仕様登録依頼番号、登録日付）
    ''ただし、製作条件付与取消にあるレコードは除く
    'sql = "select hinban||ltrim(to_char(mnorevno,'00'))||factory||opecond as hinban12, HMGSTRRNO, REGDATE " & _
          "From tbcme018 " & _
          "where (opecond='1') and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme030) and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme031)"
    'sql = "select A.hinban||ltrim(to_char(A.mnorevno,'00'))||A.factory||A.opecond as hinban12, A.HMGSTRRNO, A.REGDATE , B.UPDDATE as TOUROKU " & _
          "From tbcme018 A , tbcme036 B " & _
          "where (A.opecond='1') and (B.opecond='1') and " & _
          "(A.hinban||A.mnorevno||A.factory) = (B.hinban||B.mnorevno||B.factory) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select D.hinban||D.mnorevno||D.factory from tbcme030 D) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select C.hinban||C.mnorevno||C.factory from tbcme031 C)"
    sql = "select A.hinban||ltrim(to_char(A.mnorevno,'00'))||A.factory||A.opecond as hinban12, A.HMGSTRRNO, A.REGDATE , B.UPDDATE as TOUROKU " & _
          "From tbcme018 A , tbcme036 B " & _
          "where (A.opecond='1')  and (B.opecond(+) = '1') and " & _
          "(A.hinban||A.mnorevno||A.factory) = (B.hinban(+)||B.mnorevno(+)||B.factory(+)) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select D.hinban||D.mnorevno||D.factory from tbcme030 D) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select C.hinban||C.mnorevno||C.factory from tbcme031 C)"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_s_cmzcF_cmfc001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Hinban12 = rs("HINBAN12")        ' 品番
            .HMGSTRRNO = rs("HMGSTRRNO")    ' 品管理仕様登録依頼番号
            .REGDATE = rs("REGDATE")        ' 登録日付
            If IsNull(rs("TOUROKU")) = False Then
                .TOUROKU = rs("TOUROKU")        ' 登録日付
            Else
                .SENDFLAG = "X"
            End If
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_s_cmzcF_cmfc001b_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'この画面ではExec()はいらない
'Public Function DBDRV_s_cmzcF_cmgc001d_Exec(s_cmzcF_cmfc001a_Disp As type_DBDRV_s_cmzcF_cmgc001d_Exec) As FUNCTION_RETURN
'    s_cmzcF_cmgc001c_Exec = FUNCTION_RETURN_SUCCESS
'
'    'リメルト洗浄払出実績テーブルに原料番号()、管理工程コード()、工程コード()、乾燥後重量()、ロス重量、社員ＩＤをインサート
'
'End Function



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

'現在の仕様登録では未使用
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
Public Function DBDRV_s_cmzcF_cmfc001c_DispSXL(targetHinban As tFullHinban, SxlKokyaku_1 As c_cmzcrec, SxlKokyaku_2 As c_cmzcrec, SxlKokyaku_3 As c_cmzcrec, Sxl_1 As c_cmzcrec, Sxl_2 As c_cmzcrec, Sxl_3 As c_cmzcrec, WfKokyaku_2 As c_cmzcrec, WfKokyaku_8 As c_cmzcrec, Sxluchigawa As c_cmzcrec) As FUNCTION_RETURN
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
    Set Sxluchigawa = New c_cmzcrec
    
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
    'sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND," & _
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
                " SPECRRNO, SXLMCNO, WFMCNO, MCNO, SSTAFFID, SYNDATE "
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
    'sqlBase = "Select HSXTMMAX," & _
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
    'sqlBase = sqlBase & " HSXONSPH||HSXONSPT||HSXONSPI as HSXONSP," & _
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
                " HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU as HSXOS2K, HSXTMMAXN "
    sqlBase = "Select HSXTMMAX, HSXTMSPH, HSXTMSPT, HSXTMSPR, HSXTMKHM," & _
                " HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH," & _
                " HSXLTSPT, HSXLTSPI, HSXLTHWT, HSXLTHWS, HSXLTNSW, HSXLTKHM," & _
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
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxl_2) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''6. 製品SXL仕様_3 (TBCME020) の内容を取得する
    ''6-1.SQLを組み立てる(指定の操業条件までのレコード:操業条件逆順)
    TbName = "TBCME020"
    'sqlBase = "Select HSXDENKU," & _
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
    'sqlBase = sqlBase & " HSXOF1KM||HSXOF1KI||HSXOF1KH||HSXOF1KS as HSXOF1K," & _
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
    'sqlBase = sqlBase & " HSXBM1SZ, HSXBM1AN, HSXBM1AX," & _
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
              " HSXBM3KM||HSXBM3KI||HSXBM3KH||HSXBM3KS as HSXBM3K," & _
              " HSXDVDMNN, HSXDVDMXN, HSXDSONS, HSXCDOPMN, HSXCDOPMX," & _
              " HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK, " & _
              " HSXBMD1MBP, HSXBMD1MCL, HSXBMD2MBP, HSXBMD2MCL, HSXBMD3MBP, HSXBMD3MCL "
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
                " HSXBMD1MBP, HSXBMD1MCL, HSXBMD2MBP, HSXBMD2MCL, HSXBMD3MBP, HSXBMD3MCL, HSXDSOPTK "
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
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxluchigawa) = FUNCTION_RETURN_FAILURE Then
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

Public Function DBDRV_s_cmzcF_cmfc001c_GetHikiage(targetHinban As tFullHinban, Hikiage As String) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001c_GetHikiage"
    With targetHinban
    'sql = "Select SSXLIFTW " & _
          "From TBCME030 " & _
          " Where (HINBAN='" & .HINBAN & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='1')"
    sql = "Select SSXLIFTW " & _
          "From TBCME036 " & _
          " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='1')"
    End With
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        DBDRV_s_cmzcF_cmfc001c_GetHikiage = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    If IsNull(rs("SSXLIFTW")) Then
        Hikiage = ""
    Else
        Hikiage = rs("SSXLIFTW")     '    rs.Close
    End If
    DBDRV_s_cmzcF_cmfc001c_GetHikiage = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'現在の仕様登録では未使用
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
Public Function DBDRV_s_cmzcF_cmfc001c_Exec(ByVal IraiNo$, ByVal SXLMCNO$, ByVal WFMCNO$, ByVal Hinban12$, _
                                            ByVal SJokenNo$, ByVal Hikiage$, ByVal StaffID$) As FUNCTION_RETURN
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
        .MNOREVNO = val(Mid$(Hinban12, 9, 2))
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



'------------------------------------------------
' DBアクセス関数（抽出編）
'------------------------------------------------
'概要      :テーブル「TBCMB011」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB011 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/06/11作成　長野
Public Function DBDRV_cmbc001d_Disp(records() As typ_TBCMB011, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String                   'SQL全体
Dim sqlBase As String                   'SQL基本部(WHERE節の前まで)
Dim rs      As OraDynaset               'RecordSet
Dim recCnt  As Long                     'レコード数
Dim i       As Long                     'ﾙｰﾌﾟｶｳﾝﾄ

    ''SQLを組み立てる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Disp"

    sqlBase = "Select PGID, HZPART, HZPTRN, SPACER, UPRING, CHARGE, RTBPOS, RTBSIZE, GAP, UPDM, UPLENGTH, UPRC, RFRNEED, UPSPIN," & _
              " DOWNSPIN, ROPRESS, ARUGON, AIMOIMIN, AIMOIMAX, HCCLASS, HC, AVEUPSPD, UPCNTL, BTMSHAPE, MAGSTR, MAGPOS, CONDGRT," & _
              " MODEL, UPMETHOD, UPCLASS, UPNUM, OPETIME, WTRCOOL, PGID2, RCPT1, RCPT2, RCPT3, RCPT4, RCPT5, CNTL1, CNTL2," & _
              " CNTL3, CNTL4, CNTL5, CNTL6, CNTL7, CNTL8, CNTL9, CNTL10, CNTL11, CNTL12, CNTL13, CNTL14, CNTL15, RUNCOND1," & _
              " RUNCOND2, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB011 "
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_cmbc001d_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .PGID = rs("PGID")               ' PG-ID
            .HZPART = rs("HZPART")           ' HZパーツ
            .HZPTRN = rs("HZPTRN")           ' HZパターン
            .SPACER = rs("SPACER")           ' スペーサ
            .UPRING = rs("UPRING")           ' アッパーリング
            .CHARGE = rs("CHARGE")           ' チャージ量
            .RTBPOS = rs("RTBPOS")           ' ルツボ位置
            .RTBSIZE = rs("RTBSIZE")         ' ルツボサイズ
            .GAP = rs("GAP")                 ' ギャップ
            .UPDM = rs("UPDM")               ' 引上直径
            .UPLENGTH = rs("UPLENGTH")       ' 引上長（全長）
            .UPRC = rs("UPRC")               ' 引上（RC）
            .RFRNEED = rs("RFRNEED")         ' リフラクタ要否
            .UPSPIN = rs("UPSPIN")           ' 上軸回転数
            .DOWNSPIN = rs("DOWNSPIN")       ' 下軸回転数
            .ROPRESS = rs("ROPRESS")         ' 炉内圧
            .ARUGON = rs("ARUGON")           ' アルゴン量
            .AIMOIMIN = rs("AIMOIMIN")       ' ねらいOi（MIN)
            .AIMOIMAX = rs("AIMOIMAX")       ' ねらいOi（MAX)
            .HCCLASS = rs("HCCLASS")         ' HC種類
            .HC = rs("HC")                   ' HC
            .AVEUPSPD = rs("AVEUPSPD")       ' 平均引上速度
            .UPCNTL = rs("UPCNTL")           ' 引上制御
            .BTMSHAPE = rs("BTMSHAPE")       ' ボトム形状
            .MAGSTR = rs("MAGSTR")           ' 磁場強度
            .MAGPOS = rs("MAGPOS")           ' 磁場位置
            .CONDGRT = rs("CONDGRT")         ' 条件保証登録
            .MODEL = rs("MODEL")             ' 機種
            .UPMETHOD = rs("UPMETHOD")       ' 引上方法
            .UPCLASS = rs("UPCLASS")         ' 引上区分
            .UPNUM = rs("UPNUM")             ' 引上本数
            .OPETIME = rs("OPETIME")         ' 運転時間
            .WTRCOOL = rs("WTRCOOL")         ' 水冷管要否
            .PGID2 = rs("PGID2")             ' PG-ID（一本引）
            .RCPT1 = rs("RCPT1")             ' 対応レシピNo（T1)
            .RCPT2 = rs("RCPT2")             ' 対応レシピNo（T2)
            .RCPT3 = rs("RCPT3")             ' 対応レシピNo（T3)
            .RCPT4 = rs("RCPT4")             ' 対応レシピNo（T4)
            .RCPT5 = rs("RCPT5")             ' 対応レシピNo（T5)
            .CNTL1 = rs("CNTL1")             ' 制御項目（1）
            .CNTL2 = rs("CNTL2")             ' 制御項目（2）
            .CNTL3 = rs("CNTL3")             ' 制御項目（3）
            .CNTL4 = rs("CNTL4")             ' 制御項目（4）
            .CNTL5 = rs("CNTL5")             ' 制御項目（5）
            .CNTL6 = rs("CNTL6")             ' 制御項目（6）
            .CNTL7 = rs("CNTL7")             ' 制御項目（7）
            .CNTL8 = rs("CNTL8")             ' 制御項目（8）
            .CNTL9 = rs("CNTL9")             ' 制御項目（9）
            .CNTL10 = rs("CNTL10")           ' 制御項目（10）
            .CNTL11 = rs("CNTL11")           ' 制御項目（11）
            .CNTL12 = rs("CNTL12")           ' 制御項目（12）
            .CNTL13 = rs("CNTL13")           ' 制御項目（13）
            .CNTL14 = rs("CNTL14")           ' 制御項目（14）
            .CNTL15 = rs("CNTL15")           ' 制御項目（15）
            .RUNCOND1 = rs("RUNCOND1")       ' 運転条件１
            .RUNCOND2 = rs("RUNCOND2")       ' 運転条件２
'            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
'            .REGDATE = rs("REGDATE")         ' 登録日付
'            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
'            .UPDDATE = rs("UPDDATE")         ' 更新日付
'            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
'            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_cmbc001d_Disp = FUNCTION_RETURN_SUCCESS

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

Public Function DBDRV_cmbc001d_Exec(records As typ_TBCMB011) As FUNCTION_RETURN
'------------------------------------------------
' DBアクセス関数（更新編）
'------------------------------------------------
'概要      :テーブル「TBCMB011」の条件にあったレコードに更新をかける
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records　     ,O  ,typ_TBCMB011 ,抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :抽出ﾃﾞｰﾀの桁数・書式ﾁｪｯｸは"済み"とする
'履歴      :2001/06/19(TUE)作成　長野
Dim sql     As String                   'SQL全体
Dim rs      As OraDynaset               'RecordSet
Dim UpdID   As String                   '更新対象PGID


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Exec"

    UpdID = records.PGID

'2001/09/05 S.Sano Start 更新日時がセットされていない。
'2001/09/05 S.Sano Start このモードでsysdateのセット方法が不明。
'    sql = "SELECT * FROM TBCMB011 WHERE(PGID = '" & UpdID & "')"
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'
'    rs.Edit
'    With records
'        rs("HZPART") = StrNoNull(.HZPART)          ' HZﾊﾟｰﾂ
'        rs("HZPTRN") = StrNoNull(.HZPTRN)          ' HZﾊﾟﾀｰﾝ
'        rs("SPACER") = StrNoNull(.SPACER)          ' ｽﾍﾟｰｻ
'        rs("UPRING") = StrNoNull(.UPRING)          ' ｱｯﾊﾟｰﾘﾝｸﾞ
'        rs("CHARGE") = .CHARGE          ' ﾁｬｰｼﾞ量
'        rs("RTBPOS") = .RTBPOS          ' ﾙﾂﾎﾞ位置
'        rs("RTBSIZE") = StrNoNull(.RTBSIZE)        ' ﾙﾂﾎﾞｻｲｽﾞ
'        rs("GAP") = .GAP                ' ｷﾞｬｯﾌﾟ
'        rs("UPDM") = .UPDM              ' 引上直径
'        rs("UPLENGTH") = .UPLENGTH      ' 引上長
'        rs("UPRC") = .UPRC              ' 引上RC
'        rs("RFRNEED") = StrNoNull(.RFRNEED)        ' ﾘﾌﾗｸﾀ要否
'        rs("UPSPIN") = StrNoNull(.UPSPIN)          ' 上軸回転数
'        rs("DOWNSPIN") = StrNoNull(.DOWNSPIN)      ' 下軸回転数
'        rs("ROPRESS") = StrNoNull(.ROPRESS)        ' 炉内圧
'        rs("ARUGON") = StrNoNull(.ARUGON)          ' ｱﾙｺﾞﾝ量
'        rs("AIMOIMIN") = .AIMOIMIN      ' ねらいiO(MIN)
'        rs("AIMOIMAX") = .AIMOIMAX      ' ねらいiO(MAX)
'        rs("HCCLASS") = StrNoNull(.HCCLASS)        ' HC種類
'        rs("HC") = StrNoNull(.HC)                  ' HC
'        rs("AVEUPSPD") = .AVEUPSPD      ' 平均引上速度
'        rs("UPCNTL") = StrNoNull(.UPCNTL)          ' 引上制御
'        rs("BTMSHAPE") = StrNoNull(.BTMSHAPE)      ' ﾎﾞﾄﾑ形状
'        rs("MAGSTR") = .MAGSTR          ' 磁場強度
'        rs("MAGPOS") = .MAGPOS          ' 磁場位置
'        rs("CONDGRT") = StrNoNull(.CONDGRT)        ' 条件保証登録
'        rs("MODEL") = StrNoNull(.MODEL)            ' 機種
'        rs("UPMETHOD") = StrNoNull(.UPMETHOD)      ' 引上方法
'        rs("UPCLASS") = StrNoNull(.UPCLASS)        ' 引上区分
'        rs("UPNUM") = StrNoNull(.UPNUM)            ' 引上本数
'        rs("OPETIME") = .OPETIME        ' 運転時間
'        rs("WTRCOOL") = StrNoNull(.WTRCOOL)        ' 水冷管要否
'        rs("PGID2") = StrNoNull(.PGID2)            ' PG-ID（一本引）
'        rs("RCPT1") = StrNoNull(.RCPT1)            ' 対応ﾚｼﾋﾟNo（T1）
'        rs("RCPT2") = StrNoNull(.RCPT2)            ' 対応ﾚｼﾋﾟNo（T2）
'        rs("RCPT3") = StrNoNull(.RCPT3)            ' 対応ﾚｼﾋﾟNo（T3）
'        rs("RCPT4") = StrNoNull(.RCPT4)            ' 対応ﾚｼﾋﾟNo（T4）
'        rs("RCPT5") = StrNoNull(.RCPT5)            ' 対応ﾚｼﾋﾟNo（T5）
'        rs("CNTL1") = StrNoNull(.CNTL1)            ' 制御項目(1)
'        rs("CNTL2") = StrNoNull(.CNTL2)            ' 制御項目(2)
'        rs("CNTL3") = StrNoNull(.CNTL3)            ' 制御項目(3)
'        rs("CNTL4") = StrNoNull(.CNTL4)            ' 制御項目(4)
'        rs("CNTL5") = StrNoNull(.CNTL5)            ' 制御項目(5)
'        rs("CNTL6") = StrNoNull(.CNTL6)            ' 制御項目(6)
'        rs("CNTL7") = StrNoNull(.CNTL7)            ' 制御項目(7)
'        rs("CNTL8") = StrNoNull(.CNTL8)            ' 制御項目(8)
'        rs("CNTL9") = StrNoNull(.CNTL9)            ' 制御項目(9)
'        rs("CNTL10") = StrNoNull(.CNTL10)          ' 制御項目(10)
'        rs("CNTL11") = StrNoNull(.CNTL11)          ' 制御項目(11)
'        rs("CNTL12") = StrNoNull(.CNTL12)          ' 制御項目(12)
'        rs("CNTL13") = StrNoNull(.CNTL13)          ' 制御項目(13)
'        rs("CNTL14") = StrNoNull(.CNTL14)          ' 制御項目(14)
'        rs("CNTL15") = StrNoNull(.CNTL15)          ' 制御項目(15)
'        rs("RUNCOND1") = StrNoNull(.RUNCOND1)      ' 運転条件1
'        rs("RUNCOND2") = StrNoNull(.RUNCOND2)      ' 運転条件2
'    End With
'    rs.Update
'
'    rs.Close
    
'2001/09/05 S.Sano Start
    With records
    sql = "update TBCMB011 set "
    sql = sql & "HZPART = '" & StrNoNull(.HZPART) & "', "       ' HZﾊﾟｰﾂ
    sql = sql & "HZPTRN = '" & StrNoNull(.HZPTRN) & "', "       ' HZﾊﾟﾀｰﾝ
    sql = sql & "SPACER = '" & StrNoNull(.SPACER) & "', "       ' ｽﾍﾟｰｻ
    sql = sql & "UPRING = '" & StrNoNull(.UPRING) & "', "       ' ｱｯﾊﾟｰﾘﾝｸﾞ
    sql = sql & "CHARGE = " & .CHARGE & ", "                    ' ﾁｬｰｼﾞ量
    sql = sql & "RTBPOS = " & .RTBPOS & ", "                    ' ﾙﾂﾎﾞ位置
    sql = sql & "RTBSIZE = '" & StrNoNull(.RTBSIZE) & "', "     ' ﾙﾂﾎﾞｻｲｽﾞ
    sql = sql & "GAP = " & .GAP & ", "                          ' ｷﾞｬｯﾌﾟ
    sql = sql & "UPDM = " & .UPDM & ", "                        ' 引上直径
    sql = sql & "UPLENGTH = " & .UPLENGTH & ", "                ' 引上長
    sql = sql & "UPRC = " & .UPRC & ", "                        ' 引上RC
    sql = sql & "RFRNEED = '" & StrNoNull(.RFRNEED) & "', "     ' ﾘﾌﾗｸﾀ要否
    sql = sql & "UPSPIN = '" & StrNoNull(.UPSPIN) & "', "       ' 上軸回転数
    sql = sql & "DOWNSPIN = '" & StrNoNull(.DOWNSPIN) & "', "   ' 下軸回転数
    sql = sql & "ROPRESS = '" & StrNoNull(.ROPRESS) & "', "     ' 炉内圧
    sql = sql & "ARUGON = '" & StrNoNull(.ARUGON) & "', "       ' ｱﾙｺﾞﾝ量
    sql = sql & "AIMOIMIN = " & .AIMOIMIN & ", "                ' ねらいiO(MIN)
    sql = sql & "AIMOIMAX = " & .AIMOIMAX & ", "                ' ねらいiO(MAX)
    sql = sql & "HCCLASS = '" & StrNoNull(.HCCLASS) & "', "     ' HC種類
    sql = sql & "HC = '" & StrNoNull(.HC) & "', "               ' HC
    sql = sql & "AVEUPSPD = " & .AVEUPSPD & ", "                ' 平均引上速度
    sql = sql & "UPCNTL = '" & StrNoNull(.UPCNTL) & "', "       ' 引上制御
    sql = sql & "BTMSHAPE = '" & StrNoNull(.BTMSHAPE) & "', "   ' ﾎﾞﾄﾑ形状
    sql = sql & "MAGSTR = " & .MAGSTR & ", "                    ' 磁場強度
    sql = sql & "MAGPOS = " & .MAGPOS & ", "                    ' 磁場位置
    sql = sql & "CONDGRT = '" & StrNoNull(.CONDGRT) & "', "     ' 条件保証登録
    sql = sql & "MODEL = '" & StrNoNull(.MODEL) & "', "         ' 機種
    sql = sql & "UPMETHOD = '" & StrNoNull(.UPMETHOD) & "', "   ' 引上方法
    sql = sql & "UPCLASS = '" & StrNoNull(.UPCLASS) & "', "     ' 引上区分
    sql = sql & "UPNUM = '" & StrNoNull(.UPNUM) & "', "         ' 引上本数
    sql = sql & "OPETIME = " & .OPETIME & ", "                  ' 運転時間
    sql = sql & "WTRCOOL = '" & StrNoNull(.WTRCOOL) & "', "     ' 水冷管要否
    sql = sql & "PGID2 = '" & StrNoNull(.PGID2) & "', "         ' PG-ID（一本引）
    sql = sql & "RCPT1 = '" & StrNoNull(.RCPT1) & "', "         ' 対応ﾚｼﾋﾟNo（T1）
    sql = sql & "RCPT2 = '" & StrNoNull(.RCPT2) & "', "         ' 対応ﾚｼﾋﾟNo（T2）
    sql = sql & "RCPT3 = '" & StrNoNull(.RCPT3) & "', "         ' 対応ﾚｼﾋﾟNo（T3）
    sql = sql & "RCPT4 = '" & StrNoNull(.RCPT4) & "', "         ' 対応ﾚｼﾋﾟNo（T4）
    sql = sql & "RCPT5 = '" & StrNoNull(.RCPT5) & "', "         ' 対応ﾚｼﾋﾟNo（T5）
    sql = sql & "CNTL1 = '" & StrNoNull(.CNTL1) & "', "         ' 制御項目(1)
    sql = sql & "CNTL2 = '" & StrNoNull(.CNTL2) & "', "         ' 制御項目(2)
    sql = sql & "CNTL3 = '" & StrNoNull(.CNTL3) & "', "         ' 制御項目(3)
    sql = sql & "CNTL4 = '" & StrNoNull(.CNTL4) & "', "         ' 制御項目(4)
    sql = sql & "CNTL5 = '" & StrNoNull(.CNTL5) & "', "         ' 制御項目(5)
    sql = sql & "CNTL6 = '" & StrNoNull(.CNTL6) & "', "         ' 制御項目(6)
    sql = sql & "CNTL7 = '" & StrNoNull(.CNTL7) & "', "         ' 制御項目(7)
    sql = sql & "CNTL8 = '" & StrNoNull(.CNTL8) & "', "         ' 制御項目(8)
    sql = sql & "CNTL9 = '" & StrNoNull(.CNTL9) & "', "         ' 制御項目(9)
    sql = sql & "CNTL10 = '" & StrNoNull(.CNTL10) & "', "       ' 制御項目(10)
    sql = sql & "CNTL11 = '" & StrNoNull(.CNTL11) & "', "       ' 制御項目(11)
    sql = sql & "CNTL12 = '" & StrNoNull(.CNTL12) & "', "       ' 制御項目(12)
    sql = sql & "CNTL13 = '" & StrNoNull(.CNTL13) & "', "       ' 制御項目(13)
    sql = sql & "CNTL14 = '" & StrNoNull(.CNTL14) & "', "       ' 制御項目(14)
    sql = sql & "CNTL15 = '" & StrNoNull(.CNTL15) & "', "       ' 制御項目(15)
    sql = sql & "RUNCOND1 = '" & StrNoNull(.RUNCOND1) & "', "   ' 運転条件1
    sql = sql & "RUNCOND2 = '" & StrNoNull(.RUNCOND2) & "', "   ' 運転条件2
    sql = sql & "KSTAFFID = '" & .KSTAFFID & "', "              ' 更新社員ID
    sql = sql & "UPDDATE = sysdate "                            ' 更新日付
    sql = sql & "where PGID = '" & UpdID & "'"
    End With
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_cmbc001d_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_cmbc001d_Exec = FUNCTION_RETURN_SUCCESS
'2001/09/05 S.Sano End

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

Private Function StrNoNull(s$) As String
    If Trim$(s) = vbNullString Then
        StrNoNull = " "
    Else
        StrNoNull = Trim$(s)
    End If
End Function

'------------------------------------------------
' DBアクセス関数（削除編）
'------------------------------------------------
'概要      :テーブル「TBCMB011」の条件にあったレコードを削除
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :PGID　        ,O  ,String       ,削除PG-ID
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/10/05 作成　蔵本
Public Function DBDRV_cmbc001d_Del(PGID As String) As FUNCTION_RETURN

    Dim sql     As String                   'SQL全体

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Del"
    
    sql = "delete "
    sql = sql & "from "
    sql = sql & "TBCMB011 "
    sql = sql & "where "
    sql = sql & "trim(PGID)='" & Trim(PGID) & "'"
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_cmbc001d_Del = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_cmbc001d_Del = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_cmbc001d_Del = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


' 製作条件メンテナンス

'概要      :製作条件メンテナンス 製作条件更新／挿入用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:sMkCondNo　　　,I  ,String         　,製作条件№
'      　　:pMkOld   　　　,I  ,typ_TBCMB012   　,製作条件オリジナル
'      　　:pMkNew   　　　,I  ,typ_TBCMB012   　,製作条件
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込み成否
'説明      :
'履歴      :2001/07/30 蔵本 作成
Public Function DBDRV_scmzc_fcmbc001e_UpdInsMkCond(sMkCondNo As String, pMkOld() As typ_TBCMB012, pMkNew As typ_TBCMB012) As FUNCTION_RETURN

    Dim sql As String
    Dim bFlag As Boolean
    Dim i As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_UpdInsMkCond"

    With pMkNew
        bFlag = False
        For i = 1 To UBound(pMkOld)
            If RTrim$(pMkOld(i).MKCONDNO) = RTrim$(sMkCondNo) Then
                bFlag = True
                Exit For
            End If
        Next i

        If bFlag = True Then
            '' 製作条件の更新
            sql = "update TBCMB012 set "
            sql = sql & "MKCONDNO='" & .MKCONDNO & "', "    ' 製作条件No.
            sql = sql & "MODEL='" & .MODEL & "', "          ' 機種
            sql = sql & "RTBSIZE='" & .RTBSIZE & "', "      ' ルツボサイズ
            sql = sql & "CHARGE='" & .CHARGE & "', "        ' チャージ量
            sql = sql & "HZTYPE='" & .HZTYPE & "', "        ' HZタイプ
            sql = sql & "UPSPDTYP='" & .UPSPDTYP & "', "    ' 引上げ速度タイプ
            sql = sql & "MAGTYPE='" & .MAGTYPE & "', "      ' 磁場タイプ
            sql = sql & "USECLS='0', "                      ' 使用区分
            sql = sql & "TSTAFFID='" & .TSTAFFID & "', "    ' 登録社員ID
            sql = sql & "REGDATE=sysdate, "                 ' 登録日付
            sql = sql & "KSTAFFID='" & .KSTAFFID & "', "    ' 更新社員ID
            sql = sql & "UPDDATE=sysdate, "                 ' 更新日付
            sql = sql & "SENDFLAG='0', "                    ' 送信フラグ
            sql = sql & "SENDDATE=sysdate"                  ' 送信日時
            sql = sql & " where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
        Else
            '' 製作条件の挿入
            sql = "insert into TBCMB012 ("
            sql = sql & "MKCONDNO, "        ' 製作条件No.
            sql = sql & "MODEL, "           ' 機種
            sql = sql & "RTBSIZE, "         ' ルツボサイズ
            sql = sql & "CHARGE, "          ' チャージ量
            sql = sql & "HZTYPE, "          ' HZタイプ
            sql = sql & "UPSPDTYP, "        ' 引上げ速度タイプ
            sql = sql & "MAGTYPE, "         ' 磁場タイプ
            sql = sql & "USECLS, "          ' 使用区分
            sql = sql & "TSTAFFID, "        ' 登録社員ID
            sql = sql & "REGDATE, "         ' 登録日付
            sql = sql & "KSTAFFID, "        ' 更新社員ID
            sql = sql & "UPDDATE, "         ' 更新日付
            sql = sql & "SENDFLAG, "        ' 送信フラグ
            sql = sql & "SENDDATE)"         ' 送信日時
            sql = sql & " values ('"
            sql = sql & .MKCONDNO & "', '"  ' 製作条件No.
            sql = sql & .MODEL & "', '"     ' 機種
            sql = sql & .RTBSIZE & "', '"   ' ルツボサイズ
            sql = sql & .CHARGE & "', '"    ' チャージ量
            sql = sql & .HZTYPE & "', '"    ' HZタイプ
            sql = sql & .UPSPDTYP & "', '"  ' 引上げ速度タイプ
            sql = sql & .MAGTYPE & "', "    ' 磁場タイプ
            sql = sql & "'0', '"            ' 使用区分
            sql = sql & .TSTAFFID & "', "   ' 登録社員ID
            sql = sql & "sysdate, '"        ' 登録日付
            sql = sql & .KSTAFFID & "', "   ' 更新社員ID
            sql = sql & "sysdate, "         ' 更新日付
            sql = sql & "'0', "             ' 送信フラグ
            sql = sql & "sysdate)"          ' 送信日時
        End If
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :製作条件メンテナンス 製作条件削除用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:sMkCondNo　　　,I  ,String         　,製作条件№
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込み成否
'説明      :
'履歴      :2001/07/30 蔵本 作成
Public Function DBDRV_scmzc_fcmbc001e_DelMkCond(sMkCondNo As String) As FUNCTION_RETURN

    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_DelMkCond"

    '' 製作条件の削除
    sql = "delete TBCMB012 where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :製作条件メンテナンス 製作条件PG-ID対応更新／挿入用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:sMkCondNo　　　,I  ,String         　,製作条件№
'      　　:pPGIDOld 　　　,I  ,typ_TBCMB013   　,製作条件PG-ID対応オリジナル
'      　　:pPGIDNew 　　　,I  ,typ_TBCMB013   　,製作条件PG-ID対応
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込み成否
'説明      :
'履歴      :2001/07/30 蔵本 作成
Public Function DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng(sMkCondNo As String, pPGIDOld() As typ_TBCMB013, pPGIDNew() As typ_TBCMB013) As FUNCTION_RETURN

    Dim sql As String
    Dim bFlag As Boolean
    Dim i As Long
    Dim j As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_InsPGIDMng"

    For i = 1 To UBound(pPGIDNew)
        With pPGIDNew(i)
            bFlag = False
            For j = 1 To UBound(pPGIDOld)
                If RTrim$(pPGIDOld(j).MKCONDNO) = RTrim$(sMkCondNo) And _
                   RTrim$(pPGIDOld(j).PGIDNO) = RTrim$(.PGIDNO) Then
                    bFlag = True
                    Exit For
                End If
            Next j

            If bFlag = True Then
                '' 製作条件PG-ID対応の更新
                sql = "update TBCMB013 set "
                sql = sql & "MKCONDNO='" & .MKCONDNO & "', "    ' 製作条件No.
                sql = sql & "PGIDNO='" & .PGIDNO & "', "        ' PG-IDNo
                sql = sql & "TSTAFFID='" & .TSTAFFID & "', "    ' 登録社員ID
                sql = sql & "REGDATE=sysdate, "                 ' 登録日付
                sql = sql & "KSTAFFID='" & .KSTAFFID & "', "    ' 更新社員ID
                sql = sql & "UPDDATE=sysdate, "                 ' 更新日付
                sql = sql & "SENDFLAG='0', "                    ' 送信フラグ
                sql = sql & "SENDDATE=sysdate"                  ' 送信日付
                sql = sql & " where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
                sql = sql & " and rtrim(PGIDNO)='" & RTrim$(.PGIDNO) & "'"
            Else
                '' 製作条件PG-ID対応の挿入
                sql = "insert into TBCMB013 ("
                sql = sql & "MKCONDNO, "        ' 製作条件No.
                sql = sql & "PGIDNO, "          ' PG-IDNo
                sql = sql & "TSTAFFID, "        ' 登録社員ID
                sql = sql & "REGDATE, "         ' 登録日付
                sql = sql & "KSTAFFID, "        ' 更新社員ID
                sql = sql & "UPDDATE, "         ' 更新日付
                sql = sql & "SENDFLAG, "        ' 送信フラグ
                sql = sql & "SENDDATE)"         ' 送信日付
                sql = sql & " values ('"
                sql = sql & .MKCONDNO & "', '"  ' 製作条件No.
                sql = sql & .PGIDNO & "', '"    ' PG-IDNo
                sql = sql & .TSTAFFID & "', "   ' 登録社員ID
                sql = sql & "sysdate, '"        ' 登録日付
                sql = sql & .KSTAFFID & "', "   ' 更新社員ID
                sql = sql & "sysdate, "         ' 更新日付
                sql = sql & "'0', "             ' 送信フラグ
                sql = sql & "sysdate)"          ' 送信日付
            End If
        End With
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :製作条件メンテナンス 製作条件PG-ID対応削除用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:sMkCondNo　　　,I  ,String         　,製作条件№
'      　　:sPGIDNo  　　　,I  ,String         　,PG-ID№
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込み成否
'説明      :
'履歴      :2001/07/30 蔵本 作成
Public Function DBDRV_scmzc_fcmbc001e_DelPGIDMng(sMkCondNo As String, sPGIDNo As String) As FUNCTION_RETURN

    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_DelPGIDMng"

    '' 製作条件PG-ID対応の削除
    sql = "delete TBCMB013 where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
    If RTrim$(sPGIDNo) <> "" Then
        sql = sql & " and rtrim(PGIDNO)='" & RTrim$(sPGIDNo) & "'"
    End If
    If OraDB.ExecuteSQL(sql) < 0 Then
        DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMB011」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB011 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMB011_SQL.basより移動)
Public Function DBDRV_GetTBCMB011(records() As typ_TBCMB011, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select PGID, HZPART, HZPTRN, SPACER, UPRING, CHARGE, RTBPOS, RTBSIZE, GAP, UPDM, UPLENGTH, UPRC, RFRNEED, UPSPIN," & _
              " DOWNSPIN, ROPRESS, ARUGON, AIMOIMIN, AIMOIMAX, HCCLASS, HC, AVEUPSPD, UPCNTL, BTMSHAPE, MAGSTR, MAGPOS, CONDGRT," & _
              " MODEL, UPMETHOD, UPCLASS, UPNUM, OPETIME, WTRCOOL, PGID2, RCPT1, RCPT2, RCPT3, RCPT4, RCPT5, CNTL1, CNTL2," & _
              " CNTL3, CNTL4, CNTL5, CNTL6, CNTL7, CNTL8, CNTL9, CNTL10, CNTL11, CNTL12, CNTL13, CNTL14, CNTL15, RUNCOND1," & _
              " RUNCOND2, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB011"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB011 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .PGID = rs("PGID")               ' PG-ID
            .HZPART = rs("HZPART")           ' HZパーツ
            .HZPTRN = rs("HZPTRN")           ' HZパターン
            .SPACER = rs("SPACER")           ' スペーサ
            .UPRING = rs("UPRING")           ' アッパーリング
            .CHARGE = rs("CHARGE")           ' チャージ量
            .RTBPOS = rs("RTBPOS")           ' ルツボ位置
            .RTBSIZE = rs("RTBSIZE")         ' ルツボサイズ
            .GAP = rs("GAP")                 ' ギャップ
            .UPDM = rs("UPDM")               ' 引上直径
            .UPLENGTH = rs("UPLENGTH")       ' 引上長（全長）
            .UPRC = rs("UPRC")               ' 引上（RC）
            .RFRNEED = rs("RFRNEED")         ' リフラクタ要否
            .UPSPIN = rs("UPSPIN")           ' 上軸回転数
            .DOWNSPIN = rs("DOWNSPIN")       ' 下軸回転数
            .ROPRESS = rs("ROPRESS")         ' 炉内圧
            .ARUGON = rs("ARUGON")           ' アルゴン量
            .AIMOIMIN = rs("AIMOIMIN")       ' ねらいOi（MIN)
            .AIMOIMAX = rs("AIMOIMAX")       ' ねらいOi（MAX)
            .HCCLASS = rs("HCCLASS")         ' HC種類
            .HC = rs("HC")                   ' HC
            .AVEUPSPD = rs("AVEUPSPD")       ' 平均引上速度
            .UPCNTL = rs("UPCNTL")           ' 引上制御
            .BTMSHAPE = rs("BTMSHAPE")       ' ボトム形状
            .MAGSTR = rs("MAGSTR")           ' 磁場強度
            .MAGPOS = rs("MAGPOS")           ' 磁場位置
            .CONDGRT = rs("CONDGRT")         ' 条件保証登録
            .MODEL = rs("MODEL")             ' 機種
            .UPMETHOD = rs("UPMETHOD")       ' 引上方法
            .UPCLASS = rs("UPCLASS")         ' 引上区分
            .UPNUM = rs("UPNUM")             ' 引上本数
            .OPETIME = rs("OPETIME")         ' 運転時間
            .WTRCOOL = rs("WTRCOOL")         ' 水冷管要否
            .PGID2 = rs("PGID2")             ' PG-ID（一本引）
            .RCPT1 = rs("RCPT1")             ' 対応レシピNo（T1)
            .RCPT2 = rs("RCPT2")             ' 対応レシピNo（T2)
            .RCPT3 = rs("RCPT3")             ' 対応レシピNo（T3)
            .RCPT4 = rs("RCPT4")             ' 対応レシピNo（T4)
            .RCPT5 = rs("RCPT5")             ' 対応レシピNo（T5)
            .CNTL1 = rs("CNTL1")             ' 制限項目（1）
            .CNTL2 = rs("CNTL2")             ' 制限項目（2）
            .CNTL3 = rs("CNTL3")             ' 制限項目（3）
            .CNTL4 = rs("CNTL4")             ' 制限項目（4）
            .CNTL5 = rs("CNTL5")             ' 制限項目（5）
            .CNTL6 = rs("CNTL6")             ' 制限項目（6）
            .CNTL7 = rs("CNTL7")             ' 制限項目（7）
            .CNTL8 = rs("CNTL8")             ' 制限項目（8）
            .CNTL9 = rs("CNTL9")             ' 制限項目（9）
            .CNTL10 = rs("CNTL10")           ' 制限項目（10）
            .CNTL11 = rs("CNTL11")           ' 制限項目（11）
            .CNTL12 = rs("CNTL12")           ' 制限項目（12）
            .CNTL13 = rs("CNTL13")           ' 制限項目（13）
            .CNTL14 = rs("CNTL14")           ' 制限項目（14）
            .CNTL15 = rs("CNTL15")           ' 制限項目（15）
            .RUNCOND1 = rs("RUNCOND1")       ' 運転条件１
            .RUNCOND2 = rs("RUNCOND2")       ' 運転条件２
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB011 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMB012」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB012 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMB012_SQL.basより移動)
Public Function DBDRV_GetTBCMB012(records() As typ_TBCMB012, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select MKCONDNO, MODEL, RTBSIZE, CHARGE, HZTYPE, UPSPDTYP, MAGTYPE, USECLS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
              " SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB012"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB012 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .MKCONDNO = rs("MKCONDNO")       ' 制作条件No.
            .MODEL = rs("MODEL")             ' 機種
            .RTBSIZE = rs("RTBSIZE")         ' ルツボサイズ
            .CHARGE = rs("CHARGE")           ' チャージ量
            .HZTYPE = rs("HZTYPE")           ' HZタイプ
            .UPSPDTYP = rs("UPSPDTYP")       ' 引上げ速度タイプ
            .MAGTYPE = rs("MAGTYPE")         ' 磁場タイプ
            .USECLS = rs("USECLS")           ' 使用区分
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日時
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB012 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMB013」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB013 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMB013_SQL.basより移動)
Public Function DBDRV_GetTBCMB013(records() As typ_TBCMB013, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select MKCONDNO, PGIDNO, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB013"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB013 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .MKCONDNO = rs("MKCONDNO")       ' 制作条件No.
            .PGIDNO = rs("PGIDNO")           ' PG-IDNo
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB013 = FUNCTION_RETURN_SUCCESS
End Function

Public Function DBDRV_Syounin_Disp(records() As s_cmzcF_cmfc001b_Disp) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "Syounin_SQL.bas -- Function DBDRV_Syounin_Disp"
    
    ''未承認製品仕様管理コード取得（品番、仕様登録依頼番号、更新日付）
    ''ただし、製作条件付与取消にあるレコードは除く
    'sql = "select hinban||ltrim(to_char(mnorevno,'00'))||factory||opecond as hinban12, HMGSTRRNO, REGDATE " & _
          "From tbcme018 " & _
          "where (opecond > '1') and (nvl(synflag, ' ') = '0') and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme031) and " & _
          "(hinban||mnorevno||factory) in (select hinban||mnorevno||factory from tbcme032) "
    'sql = "select hinban||ltrim(to_char(mnorevno,'00'))||factory||opecond as hinban12, HMGSTRRNO, REGDATE " & _
          "From tbcme018 " & _
          "where (opecond > '1') and (nvl(synflag, ' ') = '0') and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme031)  "
    'sql = "select A.hinban||ltrim(to_char(A.mnorevno,'00'))||A.factory||A.opecond as hinban12, A.HMGSTRRNO, A.REGDATE, B.SENDFLAG, B.SENDDATE " & _
          "From tbcme018 A , tbcme032 B " & _
          "where (A.opecond > '1') and (nvl(A.synflag, ' ') = '0') and " & _
          "(A.hinban||A.mnorevno||A.factory = B.hinban(+)||B.mnorevno(+)||B.factory(+) ) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select hinban||mnorevno||factory from tbcme031) "
    'sql = "select A.hinban||ltrim(to_char(A.mnorevno,'00'))||A.factory||A.opecond as hinban12, A.HMGSTRRNO, A.REGDATE , B.SYNFLAG, B.UPDDATE as TOUROKU " & _
          "From tbcme018 A , tbcme036 B " & _
          "where (A.opecond='1')  and (B.opecond(+) = '1') and " & _
          "(A.hinban||A.mnorevno||A.factory) = (B.hinban(+)||B.mnorevno(+)||B.factory(+)) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select D.hinban||D.mnorevno||D.factory from tbcme030 D) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select C.hinban||C.mnorevno||C.factory from tbcme031 C)"
    sql = "select A.hinban||ltrim(to_char(A.mnorevno,'00'))||A.factory||A.opecond as hinban12, A.HMGSTRRNO, A.REGDATE , B.UPDDATE as TOUROKU , C.SENDFLAG, C.SENDDATE " & _
          "From tbcme018 A , tbcme036 B , tbcme032 C " & _
          "where (nvl(A.synflag, ' ') = '0') and A.opecond = B.opecond(+) and " & _
          "(A.hinban||A.mnorevno||A.factory) = (B.hinban(+)||B.mnorevno(+)||B.factory(+)) and " & _
          "(A.hinban||A.mnorevno||A.factory = C.hinban(+)||C.mnorevno(+)||C.factory(+) ) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select hinban||mnorevno||factory from tbcme031 )"
          Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_Syounin_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Hinban12 = rs("HINBAN12")       ' 品番
            .HMGSTRRNO = rs("HMGSTRRNO")   ' 品管理仕様登録依頼番号
            .REGDATE = rs("REGDATE")       ' 登録日付
            If IsNull(rs("TOUROKU")) = False Then .TOUROKU = rs("TOUROKU")
            If IsNull(rs("SENDDATE")) = False Then .SENDDATE = rs("SENDDATE")
            If IsNull(rs("SENDFLAG")) = False Then .SENDFLAG = rs("SENDFLAG")
        End With
        rs.MoveNext
    Next
    rs.Close


    DBDRV_Syounin_Disp = FUNCTION_RETURN_SUCCESS

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
'Public Function DBDRV_s_cmzcF_cmfc003c_Exec(ByVal IraiNo$, ByVal Sxluchigawa As c_cmzcrec, Hinban12$, ByVal StaffID$, ByVal Snote$, ByVal Jnote$) As FUNCTION_RETURN
Public Function DBDRV_s_cmzcF_cmfc003c_Exec(ByVal IraiNo$, ByVal Sxluchigawa As c_cmzcrec, Hinban12$, ByVal StaffID$) As FUNCTION_RETURN
Dim sql_top As String
Dim sql_sel As String
Dim sql As String
Dim fld As OraField
Dim rs As OraDynaset
Dim fullHinban As tFullHinban


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc003c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc003c_Exec"
    ''トランザクション開始
    Debug.Print "BeginTrans ======="
    OraDB.BeginTrans
    
    If Len(Hinban12) <> 12 Then
        DBDRV_s_cmzcF_cmfc003c_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    With fullHinban
        .hinban = Left$(Hinban12, 8)
        .MNOREVNO = val(Mid$(Hinban12, 9, 2))
        .FACTORY = Mid$(Hinban12, 11)
        .OPECOND = Right$(Hinban12, 1)
    End With
    
    sql = "insert into TBCME030 " & _
          "(HINBAN, MNOREVNO, FACTORY, OPECOND, SSXLIFTW, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO, " & _
          "TOPREG, BTMSPRT, MCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE, QASENDFLAG) " & _
          "values ("
    sql = sql & "'" & fullHinban.hinban & "', "     ' 品番
    sql = sql & fullHinban.MNOREVNO & ", "          ' 製品番号改訂番号
    sql = sql & "'" & fullHinban.FACTORY & "', "    ' 工場
    'sql = sql & "'" & fullHinban.OPECOND & "', "    ' 操業条件
    sql = sql & "'1', "    ' 操業条件
    'sql = sql & "'" & Hikiage & "', "               ' 製ＳＸ引上方法
    sql = sql & "'" & Sxluchigawa("SSXLIFTW") & "', "               ' 製ＳＸ引上方法
    sql = sql & "' ', "                             ' Ｉ／Ｆ区分
    sql = sql & "' ', "                             ' 処理区分
    sql = sql & "'" & IraiNo & "', "                ' 仕様登録依頼番号
    sql = sql & "'" & Sxluchigawa("SXLMCNO") & "', "    ' ＳＸＬ製作条件番号
    sql = sql & "'" & Sxluchigawa("WFMCNO") & "', "     ' ＷＦ製作条件番号
    sql = sql & "'" & Sxluchigawa("TOPREG") & "', "     ' TOP規制          04/07/09
    sql = sql & "'" & Sxluchigawa("BTMSPRT") & "', "    ' ボトム析出規制    04/07/09
    sql = sql & "'" & Sxluchigawa("MCNO") & "', "       ' 製作条件    04/09/01
    sql = sql & "'" & StaffID & "', "               ' 社員ID
    sql = sql & "SYSDATE, "                         ' 登録日付
    sql = sql & "SYSDATE, "                         ' 更新日付
    sql = sql & "'0', "                             ' 送信フラグ
    sql = sql & "SYSDATE, "                         ' 送信日付
    sql = sql & "'0' "                              ' 品質送信フラグ
    sql = sql & ")"
    If 0 >= OraDB.ExecuteSQL(sql) Then
        GoTo proc_err
    End If
    ''品番データに製作条件Noを書き込む(受取仕様であるリビジョン１を除く各リビジョンに）<---登録に移動
    'sql = "update TBCME018 set " & _
          "MCNO = '" & SJokenNo & "' " & _
          "where " & _
          "(HINBAN = '" & fullHinban.HINBAN & "') and " & _
          "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
          "(FACTORY = '" & fullHinban.FACTORY & "')"
    'OraDB.ExecuteSQL sql
    'If 0 >= OraDB.ExecuteSQL(sql) Then
    '    GoTo proc_err
    'End If
    
    ''特記事項の更新は承認処理で行うので削除
    'sql = "update TBCME036 set " & _
    '      "UPDDATE = sysdate ," & _
    '      "SNOTE = '" & Snote & "' ," & _
    '      "JNOTE = '" & Jnote & "' " & _
    '      "where " & _
    '      "(HINBAN = '" & fullHinban.HINBAN & "') and " & _
    '      "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
    '      "(FACTORY = '" & fullHinban.FACTORY & "') and " & _
    '      "(OPECOND = '" & fullHinban.OPECOND & "')"
    'If 0 >= OraDB.ExecuteSQL(sql) Then
    '    GoTo proc_err
    'End If
    
    DBDRV_s_cmzcF_cmfc003c_Exec = FUNCTION_RETURN_SUCCESS
    
    ''正常終了ならコミット
    Debug.Print "CommitTrans ======="
    OraDB.CommitTrans

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "==== Error SQL ===="
    Debug.Print sql
    gErr.HandleError
    ''エラー時はロールバック
    Debug.Print "RollBack ======="
    OraDB.Rollback
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
Public Function DBDRV_TOUROKU_Exec(ByVal IraiNo$, ByVal Sxl_1 As c_cmzcrec, Sxluchigawa As c_cmzcrec, Hinban12$, _
                                            ByVal SJokenNo$, ByVal Hikiage$, ByVal StaffID$, ByVal Snote$, ByVal Jnote$) As FUNCTION_RETURN
Dim sql_top As String
Dim sql_sel As String
Dim sql As String
Dim fld As OraField
Dim rs As OraDynaset
Dim fullHinban As tFullHinban


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "TOUROKU_SQL.bas -- Function DBDRV_TOUROKU_Exec"
    ''トランザクション開始
    Debug.Print "BeginTrans ======="
    OraDB.BeginTrans
    
    If Len(Hinban12) <> 12 Then
        DBDRV_TOUROKU_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    With fullHinban
        .hinban = Left$(Hinban12, 8)
        .MNOREVNO = val(Mid$(Hinban12, 9, 2))
        .FACTORY = Mid$(Hinban12, 11)
        .OPECOND = Right$(Hinban12, 1)
    End With
    ''品番データに製作条件Noを書き込む
    sql = "update TBCME018 set " & _
          "UPDDATE = SYSDATE, " & _
          "MCNO = '" & SJokenNo & "' " & _
          "where " & _
          "(HINBAN = '" & fullHinban.hinban & "') and " & _
          "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
          "(FACTORY = '" & fullHinban.FACTORY & "')"
    Debug.Print "ExecuteSQL ==========="
    Debug.Print sql
    If 0 >= OraDB.ExecuteSQL(sql) Then
        GoTo proc_err
    End If
    With Sxluchigawa
        ''全てのリビジョンの製作条件Noを書きかえる
        sql = "update TBCME036 set " & _
              "UPDDATE = SYSDATE ," & _
              "MCNO = '" & SJokenNo & "' ," & _
              "SSXLIFTW = '" & Hikiage & "' " & _
              "where " & _
              "(HINBAN = '" & fullHinban.hinban & "') and " & _
              "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
              "(FACTORY = '" & fullHinban.FACTORY & "') "
              '"(OPECOND = '" & fullHinban.OPECOND & "')"
        If 0 >= OraDB.ExecuteSQL(sql) Then
            .Fields("HINBAN") = fullHinban.hinban
            .Fields("MNOREVNO") = fullHinban.MNOREVNO
            .Fields("FACTORY") = fullHinban.FACTORY
            '.Fields("OPECOND") = "1"
            .Fields("OPECOND") = fullHinban.OPECOND
            .Fields("SPECRRNO") = IraiNo
            .Fields("SXLMCNO") = Sxl_1("SXLMCNO")
            .Fields("WFMCNO") = Sxl_1("WFMCNO")
            .Fields("SNOTE") = Snote
            .Fields("JNOTE") = Jnote
            .Fields("STAFFID") = StaffID
            .Fields("MCNO") = SJokenNo
            .Fields("SSXLIFTW") = Hikiage
            sql = .SqlInsert
            If 0 >= OraDB.ExecuteSQL(sql) Then
                    GoTo proc_err
            End If
        End If
        Debug.Print "ExecuteSQL ==========="
        Debug.Print sql
        ''特記事項を更新する
        'sql = "update TBCME036 set " & _
        '      "SNOTE = '" & Snote & "' ," & _
        '      "JNOTE = '" & Jnote & "' " & _
        '      "where " & _
        '      "(HINBAN = '" & fullHinban.HINBAN & "') and " & _
        '      "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
        '      "(FACTORY = '" & fullHinban.FACTORY & "') and " & _
        '      "(OPECOND = '" & fullHinban.OPECOND & "')"
        'Debug.Print "ExecuteSQL ==========="
        'Debug.Print sql
        'If 0 >= OraDB.ExecuteSQL(sql) Then
        '    GoTo proc_err
        'End If
        
    End With
    
    DBDRV_TOUROKU_Exec = FUNCTION_RETURN_SUCCESS
    
    ''正常終了ならコミット
    Debug.Print "CommitTrans ======="
    OraDB.CommitTrans

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "==== Error SQL ===="
    Debug.Print sql
    gErr.HandleError
    ''エラー時はロールバック
    Debug.Print "RollBack ======="
    OraDB.Rollback
    Resume proc_exit
End Function
