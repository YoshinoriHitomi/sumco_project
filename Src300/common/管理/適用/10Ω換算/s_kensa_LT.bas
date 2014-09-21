Attribute VB_Name = "s_kensa_LT"
' ライフタイム測定点数（新データは１０点固定）
Public Const SS_SOKUETI_TENSU = 10
' ライフタイム測定点数（旧データは５点固定）
Public Const SS_SOKUETI_TENSU_OLD = 5
' 配列初期化値
Public Const DEF_PARAM_VALUE_LT = -1

''　LT関数戻り値定義
Public Enum FUNC_RET_LT         ''関数の戻り値
    FUNC_RET_LT_NOSAMPLE = FUNCTION_RETURN_FAILURE + 2  '' サンプル無
    FUNC_RET_LT_NODATA = FUNCTION_RETURN_SUCCESS + 1    '' LTデータなし
    FUNC_RET_LT_SUCCESS = FUNCTION_RETURN_SUCCESS       '' 正常
    FUNC_RET_LT_FAILURE = FUNCTION_RETURN_FAILURE       '' 異常
    FUNC_RET_LT_CALCFAIL = FUNCTION_RETURN_FAILURE - 1  '' ライフタイム測定結果の算出エラー
End Enum

'' ライフタイム測定値
Public Type typ_LTMEAS
    CRYNUMCS As String * 12         'ブロックID
    XTALCS As String * 12           '結晶番号
    HINBCS As String * 8            '品番
    REVNUMCS As Integer             '製品番号改訂番号
    FACTORYCS As String * 1         '工場
    OPECS As String * 1             '操業条件
    MEAS1 As Integer                '測定値1
    MEAS2 As Integer                '測定値2
    MEAS3 As Integer                '測定値3
    MEAS4 As Integer                '測定値4
    MEAS5 As Integer                '測定値5
    MEAS6 As Integer                '測定値6
    MEAS7 As Integer                '測定値7
    MEAS8 As Integer                '測定値8
    MEAS9 As Integer                '測定値9
    MEAS10 As Integer               '測定値10
    LTSPIFLG As String * 1          '測定位置判定フラグ
End Type



'概要      :ライフタイム測定結果を算出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型              ,説明
'          :iCalcMeas     ,O   ,Integer         ,ライフタイム測定結果
'          :sCrynum       ,I   ,String          ,分割結晶番号
'          :iSmplIDLt     ,I   ,Long            ,サンプルID(LT)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O   ,FUNC_RET_LT     ,LT関数戻り値定義の通り
'説明      :
'履歴      :05/11/24 (SET)M.makino
Public Function KNS_GetLtCalcMeas(ByRef iCalcMeas As Integer, _
                                  sCryNum As String, iSmplIDLt As Long) As FUNC_RET_LT
    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tLtMeas     As typ_LTMEAS
    Dim tHinInf     As tFullHinban
    Dim tSXLData    As typ_TBCME019
    Dim iLtParam()  As Integer
    Dim iOldFlg     As Integer
    Dim sHsxLtspi   As String
    Dim sXtal        As String

    KNS_GetLtCalcMeas = FUNC_RET_LT_FAILURE

    '' ライフタイムデータの取得
    iRet = DBDRV_KNS_GetLTMeas(tLtMeas, sCryNum, iSmplIDLt)
    If iRet <> FUNC_RET_LT_SUCCESS Then
        KNS_GetLtCalcMeas = iRet
        Exit Function
    End If

    '' 新形式のデータは測定位置を取得する
    If (Trim(tLtMeas.LTSPIFLG) <> "") Then
        iOldFlg = 0
    
        '' Z品番、G品番の場合はねらい品番に置き換える
        If (Trim(tLtMeas.HINBCS) = "Z") Or (Trim(tLtMeas.HINBCS) = "G") Then
            iRet = DBDRV_KNS_GetNeraiZuban(tHinInf, tLtMeas.XTALCS)
            If iRet <> FUNC_RET_LT_SUCCESS Then Exit Function
        Else
            tHinInf.hinban = tLtMeas.HINBCS
            tHinInf.mnorevno = tLtMeas.REVNUMCS
            tHinInf.factory = tLtMeas.FACTORYCS
            tHinInf.opecond = tLtMeas.OPECS
        End If

        '' 製品仕様SXLデータ２取得(TBCME019)
        iRet = DBDRV_KNS_GetSXLData(tSXLData, tHinInf)
        If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
        
        sHsxLtspi = tSXLData.HSXLTSPI
    '' 旧形式は測定位置を取得する必要がない
    Else
        iOldFlg = 1
        sHsxLtspi = ""
    End If
    
    ReDim iLtParam(KNS_GetMeasureNum_LT(iOldFlg) - 1)
    
    ''初期化
    For Index = 0 To UBound(iLtParam)
        iLtParam(Index) = DEF_PARAM_VALUE_LT
    Next Index
    
    iLtParam(0) = tLtMeas.MEAS1
    iLtParam(1) = tLtMeas.MEAS2
    iLtParam(2) = tLtMeas.MEAS3
    iLtParam(3) = tLtMeas.MEAS4
    iLtParam(4) = tLtMeas.MEAS5
    If iOldFlg = 0 Then
        iLtParam(5) = tLtMeas.MEAS6
        iLtParam(6) = tLtMeas.MEAS7
        iLtParam(7) = tLtMeas.MEAS8
        iLtParam(8) = tLtMeas.MEAS9
        iLtParam(9) = tLtMeas.MEAS10
    End If
    
    '' 測定結果の算出
    iRet = KNS_CalculateMeasResult_LT(iCalcMeas, iLtParam(), sHsxLtspi, iOldFlg)
    If iRet <> FUNCTION_RETURN_SUCCESS Then
        KNS_GetLtCalcMeas = FUNC_RET_LT_CALCFAIL
        Exit Function
    End If

    KNS_GetLtCalcMeas = FUNC_RET_LT_SUCCESS
    
End Function

'概要      :取得した製品仕様SXLデータより測定点数を取得する（ライフタイム実績）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型             ,説明
'          :iOldFlg       ,I  ,Integer        ,旧データフラグ   (旧データ[5点測定]は1を設定する)
'          :戻り値        ,O  ,Integer        ,測定点数
'説明      :
' Mod Start 2005/11/14 M.Makino
'Private Function KNS_GetMeasureNum_LT(tHinInf As tFullHinban) As Integer
Public Function KNS_GetMeasureNum_LT(iOldFlg As Integer) As Integer
' Mod End   2005/11/14 M.Makino

    Dim Index   As Integer
    Dim strMN   As String
    Dim iNum    As Integer

' Mod Start 2005/11/14 M.Makino
'    '' ライフタイムは５点固定
'    iNum = 5
    If iOldFlg = 1 Then
        iNum = SS_SOKUETI_TENSU_OLD
    Else
        iNum = SS_SOKUETI_TENSU
    End If
' Mod End   2005/11/14 M.Makino

    KNS_GetMeasureNum_LT = iNum

End Function

'概要      :測定結果を計算する（ライフタイム実績）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :iResult       ,O   ,Integer   ,計算結果
'          :iParam()      ,I   ,Integer   ,測定値配列
'          :sHsxLtspi     ,I   ,String    ,測定位置         (新データ[10点測定]は3,5,Aのどれかを設定する)
'          :iOldFlg       ,I   ,Integer   ,旧データフラグ   (旧データ[5点測定]は1を設定する)
'          :戻り値        ,O   ,Integer   ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :2005/11/07 牧野 変更　10点測定対応
Public Function KNS_CalculateMeasResult_LT(iResult As Integer, iParam() As Integer, _
                    sHsxLtspi As String, iOldFlg As Integer) As Integer
    Dim Index   As Integer
    Dim iAve    As Integer

    On Error GoTo Err
    KNS_CalculateMeasResult_LT = FUNCTION_RETURN_FAILURE

' Mod Start 2005/11/14 M.Makino
'    '' パラメータ入力チェック
'    For Index = 0 To UBound(iParam)
'        If iParam(Index) = DEF_PARAM_VALUE_LT Then
'            Exit Function
'        End If
'    Next Index
'
'    ''３，４，５点の測定点のAVEを求める
'    iAve = RoundDown((iParam(2) + iParam(3) + iParam(4)) / 3#, 0)
'
'    '' 測定点２とAVE値を比較、値の小さい方を測定結果とする
'    If iAve < iParam(1) Then
'        iResult = iAve
'    Else
'        iResult = iParam(1)
'    End If

    
    '' 旧データの場合（５点測定）
    If iOldFlg = 1 Then
        '' パラメータ入力チェック
        For Index = 0 To KNS_GetMeasureNum_LT(iOldFlg) - 1
            If iParam(Index) = DEF_PARAM_VALUE_LT Then
                Exit Function
            End If
        Next Index
        ''３，４，５点の測定点のAVEを求める
        iAve = RoundDown((iParam(2) + iParam(3) + iParam(4)) / 3#, 0)

        '' 測定点２とAVE値を比較、値の小さい方を測定結果とする
        If iAve < iParam(1) Then
            iResult = iAve
        Else
            iResult = iParam(1)
        End If

    '' 新データの場合（１０点測定）
    Else
        '' パラメータ入力チェック
        For Index = 0 To KNS_GetMeasureNum_LT(iOldFlg) - 1
            If iParam(Index) = DEF_PARAM_VALUE_LT Then
                Exit Function
            End If
        Next Index

        ''' [A:Ce,Inside3mm]の場合
        If Trim(sHsxLtspi) = "3" Then
            ''８，９，１０点の測定点のAVEを求める
            iAve = RoundDown((iParam(7) + iParam(8) + iParam(9)) / 3#, 0)

        ''' [A:Ce,Inside5mm]の場合
        ElseIf Trim(sHsxLtspi) = "5" Then
            ''５，６，７点の測定点のAVEを求める
            iAve = RoundDown((iParam(4) + iParam(5) + iParam(6)) / 3#, 0)

        ''' [A:Ce,Inside10mm]の場合
        ElseIf Trim(sHsxLtspi) = "A" Then
            ''２，３，４点の測定点のAVEを求める
            iAve = RoundDown((iParam(1) + iParam(2) + iParam(3)) / 3#, 0)

' Mod Start 2005/12/13 M.Makino
'        ''' その他の場合はエラー
'        Else
'            Exit Function

        ''' その他の場合は[A:Ce,Inside10mm]の仕様とする
        Else
            ''２，３，４点の測定点のAVEを求める
            iAve = RoundDown((iParam(1) + iParam(2) + iParam(3)) / 3#, 0)
' Mod End   2005/12/13 M.Makino

        End If
    
        '' 測定点１とAVE値を比較、値の小さい方を測定結果とする
        If iAve < iParam(0) Then
            iResult = iAve
        Else
            iResult = iParam(0)
        End If
    End If
' Mod End   2005/11/14 M.Makino

    KNS_CalculateMeasResult_LT = FUNCTION_RETURN_SUCCESS
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
End Function

'概要      :新サンプル管理、ライフタイムテーブルからライフタイム測定値を抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records       ,O  ,typ_LTMEAS   ,抽出レコード
'          :sCrynum       ,I  ,String       ,分割結晶番号
'          :iSmplIDLt     ,I  ,Long         ,サンプル番号   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O  ,FUNC_RET_LT  ,抽出の成否
'説明      :
'履歴      :2005/11/24 Create (SET)M.Makino
Public Function DBDRV_KNS_GetLTMeas(records As typ_LTMEAS, sCryNum As String, _
                                    iSmplIDLt As Long) As FUNC_RET_LT
    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet

    DBDRV_KNS_GetLTMeas = FUNC_RET_LT_FAILURE

    ''SQLを組み立てる
    sql = sql & "select nvl(T1.CRYNUMCS, '')  CRYNUMCS" 'ブロックID
    sql = sql & ", nvl(T1.XTALCS, '')   XTALCS"         '結晶番号
    sql = sql & ", nvl(T1.HINBCS, '')    HINBCS"        '品番
    sql = sql & ", nvl(T1.REVNUMCS, 0)   REVNUMCS"      '製品番号改訂番号
    sql = sql & ", nvl(T1.FACTORYCS, '') FACTORYCS"     '工場
    sql = sql & ", nvl(T1.OPECS, '')     OPECS"         '操業条件
    sql = sql & ", nvl(T2.MEAS1, -1) MEAS1"             '測定値１
    sql = sql & ", nvl(T2.MEAS2, -1) MEAS2"             '測定値２
    sql = sql & ", nvl(T2.MEAS3, -1) MEAS3"             '測定値３
    sql = sql & ", nvl(T2.MEAS4, -1) MEAS4"             '測定値４
    sql = sql & ", nvl(T2.MEAS5, -1) MEAS5"             '測定値５
    sql = sql & ", nvl(T2.MEAS6, -1) MEAS6"             '測定値６
    sql = sql & ", nvl(T2.MEAS7, -1) MEAS7"             '測定値７
    sql = sql & ", nvl(T2.MEAS8, -1) MEAS8"             '測定値８
    sql = sql & ", nvl(T2.MEAS9, -1) MEAS9"             '測定値９
    sql = sql & ", nvl(T2.MEAS10, -1) MEAS10"           '測定値１０
    sql = sql & ", LTSPIFLG"                            '測定位置判定フラグ
    sql = sql & ", SMPLUMU"                             'サンプル有無
    sql = sql & " from XSDCS T1, TBCMJ007 T2"
    sql = sql & " where T1.CRYNUMCS = '" & sCryNum & "'"
    sql = sql & " and T1.CRYSMPLIDTCS = " & iSmplIDLt
    sql = sql & " and T1.XTALCS = T2.CRYNUM"
'    sql = sql & " and T1.INPOSCS = T2.POSITION"
    sql = sql & " and T1.TBKBNCS = T2.SMPKBN"
    sql = sql & " and T2.TRANCOND = '0'"
    sql = sql & " and T2.TRANCNT = "
    sql = sql & "("
    sql = sql & " select max(T2.TRANCNT)"
    sql = sql & " from XSDCS T1, TBCMJ007 T2"
    sql = sql & " where T1.CRYNUMCS = '" & sCryNum & "'"
    sql = sql & " and T1.CRYSMPLIDTCS = " & iSmplIDLt
    sql = sql & " and T1.XTALCS = T2.CRYNUM"
'    sql = sql & " and T1.INPOSCS = T2.POSITION"
    sql = sql & " and T1.TBKBNCS = T2.SMPKBN"
    sql = sql & " and T2.TRANCOND = '0'"
    sql = sql & ")"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    If rs Is Nothing Then
        Exit Function
    End If

    ' 該当するデータが無い場合
    If rs.EOF Then
        DBDRV_KNS_GetLTMeas = FUNC_RET_LT_NODATA
        Exit Function
    End If
    
    '' サンプル無の場合はサンプル無コードを返す
    If rs.Fields("SMPLUMU").Value <> "0" Then
        DBDRV_KNS_GetLTMeas = FUNC_RET_LT_NOSAMPLE
        Exit Function
    End If


    With records
        .CRYNUMCS = rs.Fields("CRYNUMCS").Value   'ブロックID
        .XTALCS = rs.Fields("XTALCS").Value       '結晶番号
        .HINBCS = rs.Fields("HINBCS").Value       '品番
        .REVNUMCS = rs.Fields("REVNUMCS").Value   '製品番号改訂番号
        .FACTORYCS = rs.Fields("FACTORYCS").Value '工場
        .OPECS = rs.Fields("OPECS").Value         '操業条件
        .MEAS1 = rs.Fields("MEAS1").Value         '測定値１
        .MEAS2 = rs.Fields("MEAS2").Value         '測定値２
        .MEAS3 = rs.Fields("MEAS3").Value         '測定値３
        .MEAS4 = rs.Fields("MEAS4").Value         '測定値４
        .MEAS5 = rs.Fields("MEAS5").Value         '測定値５
        .MEAS6 = rs.Fields("MEAS6").Value         '測定値６
        .MEAS7 = rs.Fields("MEAS7").Value         '測定値７
        .MEAS8 = rs.Fields("MEAS8").Value         '測定値８
        .MEAS9 = rs.Fields("MEAS9").Value         '測定値９
        .MEAS10 = rs.Fields("MEAS10").Value       '測定値１０
        .LTSPIFLG = Trim(CStr(NulltoStr(rs.Fields("LTSPIFLG").Value)))  '測定位置判定フラグ
    End With

    rs.Close

    DBDRV_KNS_GetLTMeas = FUNC_RET_LT_SUCCESS
End Function

'概要      :ねらい品番を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,tFullHinban  ,抽出レコード
'          :sXtal         ,I  ,String       ,結晶番号
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2005/11/24 Create (SET)M.Makino
Public Function DBDRV_KNS_GetNeraiZuban(records As tFullHinban, sXtal As String) As FUNCTION_RETURN
    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet

    DBDRV_KNS_GetNeraiZuban = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる
    sql = sql & "select RPHINBAN, RPREVNUM, RPFACT, RPOPCOND"
    sql = sql & " from TBCME037"
    sql = sql & " where CRYNUM = '" & sXtal & "'"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    If rs Is Nothing Then
        Exit Function
    End If

    ' 該当するデータが無い場合
    If rs.EOF Then Exit Function

    With records
        .hinban = Trim(CStr(NulltoStr(rs.Fields("RPHINBAN").Value)))
        .mnorevno = Trim(CStr(NulltoStr(rs.Fields("RPREVNUM").Value)))
        .factory = Trim(CStr(NulltoStr(rs.Fields("RPFACT").Value)))
        .opecond = Trim(CStr(NulltoStr(rs.Fields("RPOPCOND").Value)))
    End With

    rs.Close

    DBDRV_KNS_GetNeraiZuban = FUNCTION_RETURN_SUCCESS
End Function


'概要      :品番等から測定点を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :siyou         ,O  ,typ_TBCME019 ,測定位置格納構造体
'          :tHinInf       ,I  ,tFullHinban  ,品番等
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2005/11/24 Create (SET)T.Takasaki
Public Function DBDRV_KNS_GetSXLData(siyou As typ_TBCME019, tHinInf As tFullHinban) As FUNCTION_RETURN
    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet

    DBDRV_KNS_GetSXLData = FUNCTION_RETURN_FAILURE
    
    ''SQLを組み立てる
    sql = sql & "select HSXLTSPI "
    sql = sql & "from TBCME019 "
    sql = sql & "where HINBAN = '" & tHinInf.hinban & "' "
    sql = sql & "and MNOREVNO = '" & tHinInf.mnorevno & "' "
    sql = sql & "and FACTORY = '" & tHinInf.factory & "' "
    sql = sql & "and OPECOND = '" & tHinInf.opecond & "' "
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    If rs Is Nothing Then
        Exit Function
    End If

    ' 該当するデータが無い場合
    If rs.EOF Then Exit Function

    ''データ格納
    With siyou
        .HSXLTSPI = Trim(CStr(NulltoStr(rs.Fields("HSXLTSPI").Value)))
    End With
    
    rs.Close
    
    DBDRV_KNS_GetSXLData = FUNCTION_RETURN_SUCCESS

End Function

'概要      :実測抵抗の取得と１０Ω換算値の算出を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型           ,説明
'          :tblCrySmpMan  ,I   ,typ_XSDCS    ,サンプルID
'          :sKekka        ,I   ,String       ,測定結果
'          :sIncval       ,I   ,String       ,傾き
'          :sCutval       ,I   ,String       ,切片
'          :sSetval       ,I   ,String       ,設定値
'          :sJiteiko      ,I   ,String       ,実測抵抗
'          :sKansanchi    ,I   ,String       ,１０Ω換算値
'          :戻り値        ,O  ,FUNCTION_RETURN  ,抽出の成否
'説明      :
'備考      : １０Ω換算式の算出方法
'               Ａ＝ライフタイム測定結果
'               Ｂ＝実測抵抗
'               Ｃ＝切片 [桁数=XXX.XX]
'               Ｄ＝傾き [桁数=XXX.XX]
'               Ｇ＝設定値 [桁数=XXX.XX]
'               Ｅ＝理論値LT＝Ｄ×Ｂ＋Ｃ
'               Ｆ＝汚染量推定値＝１／((1／Ａ)―(1／Ｅ))
'               １０Ω換算値＝１／((１／Ｇ)＋(１／Ｆ)) [桁数=XXXX]
'履歴      :新規 2005/11/14 M.Makino
Public Function GetKansanchi(tblCrySmpMan As typ_XSDCS, sKekka As String, sIncVal As String, _
        sCutVal As String, sSetVal As String, sJiteiko As String, sKansanchi As String) As Integer
    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim RironchiLT As Double    ' 理論値LT
    Dim Osenryo As Double       ' 汚染量推定値

    GetKansanchi = FUNCTION_RETURN_FAILURE

    ' SQL文作成
    sql = ""
    sql = sql & "SELECT MEAS1"
    sql = sql & " FROM  TBCMJ002"
    sql = sql & " WHERE CRYNUM='" & tblCrySmpMan.XTALCS & "'"
    sql = sql & " AND   POSITION=" & tblCrySmpMan.INPOSCS
    sql = sql & " AND   SMPKBN='" & tblCrySmpMan.SMPKBNCS & "'"
    sql = sql & " AND   TRANCOND='0'"
    sql = sql & " ORDER BY TRANCNT DESC"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    If rs Is Nothing Then
        Exit Function
    End If

    If rs.EOF Then
        ' 該当するデータが無い場合判定は空文字
        sJiteiko = ""
    Else
        sJiteiko = Trim(CStr(NulltoStr(rs.Fields("MEAS1").Value)))
    End If

    ' １０Ω換算値の計算
    If sKekka <> "" And sIncVal <> "" And sCutVal <> "" And _
       sSetVal <> "" And sJiteiko <> "" Then

        '0の除算対策
        On Error GoTo ERROR_CALC

        '１０Ω換算値を算出
        RironchiLT = CDbl(sIncVal) * CDbl(sJiteiko) + CDbl(sCutVal)
        Osenryo = 1 / ((1 / CInt(sKekka)) - (1 / RironchiLT))
        sKansanchi = CStr(Round(1 / ((1 / CDbl(sSetVal)) + (1 / Osenryo)), 0))
    Else
        sKansanchi = ""
    End If
    
    GetKansanchi = FUNCTION_RETURN_SUCCESS
    Exit Function

ERROR_CALC:
    sKansanchi = ""
    GetKansanchi = FUNCTION_RETURN_SUCCESS
End Function

'
' 空文字列（""）に対して『null』を返し，その他の文字列は何もせずに返す
'
'履歴      :2005/11/14追加　牧野
Public Function LZeroToNull(ByVal sTmp As String) As String
    If "" = sTmp Then
        LZeroToNull = "null"
    Else
        LZeroToNull = sTmp
    End If
End Function



