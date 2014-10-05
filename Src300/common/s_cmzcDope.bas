Attribute VB_Name = "s_cmzcDope"
Type DopeData
    DopKind As String * 7       'ドープ名称(CODE)
    IonDensity As Double        'イオン濃度(OutPut)
    CoreCoeff As Integer        '補正係数
End Type
Type CodeList
    RCLSCODE As String * 3      '精製原料名称(ρ区分CODE)
    WEIGHT As Long              '重量
    IonDensity As Double        'イオン濃度(OutPut)
End Type
Type type_DopeCal
    NTYPE As String * 1         ' 狙いタイプ
    res As Double               ' 狙い抵抗
    CHARGE As Long              ' チャージ量
    Dope As DopeData
    CryList() As CodeList
    FixNumA As Double           '定数A
    FixNumB As Double           '定数B
End Type

'*ADD* TCS)K.Kunori 2004.11.29 START >>>
'■ドーパント量データ
Public Type typ_DpData
    DPWEIGHT    As Double           ' 必要ドーパント量
    ZanDp       As Double           ' 残液ドーパント量
    AddDp       As Double           ' 追加ドーパント量
    InpDp       As Double           ' 入力ドーパント量
    '*ADD* TCS)K.Kunori 2004.10.14
    PutDp       As Double           ' 投入ドーパント量
End Type

Public sDpData As typ_DpData

Public strRes1  As String           '精製原料濃度合計(×f×γ)
Public strRes2  As String           'ﾄﾞｰﾊﾟﾝﾄ希釈率(f)
'*ADD* TCS)K.Kunori 2004.11.29 END <<<
'*ADD* 補正係数追加 TCS)K.Kunori 2004.12.16
Public strRes3  As String           '補正係数(γ)

Option Explicit

Public Function Log10(x)
   Log10 = Log(x) / Log(10#)
End Function

Public Function Exp10(x)
   Exp10 = Exp(x) ^ Log(10#)
End Function

Public Function DopeCalculation(CC As type_DopeCal) As Double
    Dim Ion As Double
    Dim temp As Double
    Dim c0 As Integer
    
    DopeCalculation = -9999
    'ドーパント量計算用データ収集。
    If GetDopeCalData(CC) = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If
    
    On Error GoTo Err
    temp = 0
    For c0 = 1 To UBound(CC.CryList())
        temp = temp + CC.CryList(c0).IonDensity * CC.CryList(c0).WEIGHT * 10 ^ 14
    Next
    
    Ion = Exp10((Log10(CC.res) - CC.FixNumB) / CC.FixNumA) / 2.34
    '計算式の変更   2008/09/08 Kameda
    'DopeCalculation = ((Ion * CC.CHARGE - temp) / (CC.Dope.IonDensity * 10 ^ 14)) / CC.Dope.CoreCoeff
    DopeCalculation = ((Ion * CC.CHARGE - temp) / (CC.Dope.IonDensity * 10 ^ 14)) * CC.Dope.CoreCoeff
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    DopeCalculation = -9999
End Function

Public Function GetDopeCalData(CC As type_DopeCal) As FUNCTION_RETURN
    Dim sql As String       'SQL全体
    Dim sql1 As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim c0 As Integer
    Dim c1 As Integer
    Dim MaxRec As Integer
    Dim MaxRec1 As Integer
    Dim temp() As CodeList
    Dim sFactor As Single
    Dim sHenseki As Single
    
    GetDopeCalData = FUNCTION_RETURN_FAILURE
    
    MaxRec = UBound(CC.CryList())
    
    '精製原料のイオン濃度を求める。
    If MaxRec > 0 Then
        sql1 = ""
        For c0 = 1 To MaxRec
            sql1 = sql1 & "'" & CC.CryList(c0).RCLSCODE & "',"
        Next
        sql1 = Left(sql1, Len(sql1) - 1)
        
        sql = "select RCLSCODE, IonDensity from TBCMB007 where "
        sql = sql & "RCLSCODE in (" & sql1 & ")"
    
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        If rs.RecordCount = 0 Then
            Exit Function
        End If
        
        MaxRec1 = rs.RecordCount
        ReDim temp(1 To MaxRec1) As CodeList
        For c0 = 1 To MaxRec1
            temp(c0).RCLSCODE = rs("RCLSCODE")
            temp(c0).IonDensity = rs("IonDensity")
            rs.MoveNext
        Next
        rs.Close
        For c0 = 1 To MaxRec
            For c1 = 1 To MaxRec1
                If CC.CryList(c0).RCLSCODE = temp(c1).RCLSCODE Then
                    CC.CryList(c0).IonDensity = temp(c1).IonDensity
                End If
            Next
        Next
    End If
    
    '指定ドープのイオン濃度を求める。 2011/05/31 kameda
    'sql = "select IonDensity,CoreCoeff from TBCMB009 where "
    'sql = sql & "DopKind = '" & CC.Dope.DopKind & "' "

    'Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    'If rs.RecordCount = 0 Then
    '    Exit Function
    'End If
    
    'CC.Dope.IonDensity = rs("IonDensity")
    CC.Dope.IonDensity = (Getrnoudo(CC.Dope.DopKind, CC.res)) * 100000
    'CC.Dope.CoreCoeff = rs("CoreCoeff")
    If GetFactor(CC.NTYPE, sFactor, sHenseki) = False Then
        Call MsgOut(0, "ファクター・偏析係数取得エラー", ERR_DISP)
        Exit Function
    End If
    CC.Dope.CoreCoeff = sFactor
    'rs.Close
    
    '定数A、Bを求める。
    sql = "select FIXNUMA, FIZNUMB from TBCMB010 where "
    sql = sql & "TYPE = '" & CC.NTYPE & "' "
    sql = sql & "and RESFROM >= " & CC.res & " "
    sql = sql & "and RESTO <= " & CC.res

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        Exit Function
    End If
    CC.FixNumA = rs("FixNumA")
    CC.FixNumB = rs("FizNumB")
    rs.Close
    
    GetDopeCalData = FUNCTION_RETURN_SUCCESS
End Function

'*ADD* 計量ﾄﾞｰﾌﾟ量計算方法変更対応 ※ﾓｼﾞｭｰﾙ統一 TCS)K.Kunori 2004.11.29 START >>>
'=============================================================================
'@(f)       DopeCalculationPart
'
'機能       計量ﾄﾞｰﾌﾟ量を求める
'
'戻り値     計量ﾄﾞｰﾌﾟ量計算結果(dpVal as Double)
'
'引数       狙いﾀｲﾌﾟ                    strprmType As String
'           狙い抵抗                    dblprmRes As Double
'           ﾁｬｰｼﾞ量                     dblprmChrg As Double
'           指示ﾄﾞｰﾊﾟﾝﾄ量               dblprmSijiDp As Double
'           ﾁｬｰｼﾞ№                     strprmChrgNo As String(txtChargeNo.Text)
'
'機能概要   //計量ﾄﾞｰﾌﾟ量計算式//
'　　　　　 計量ﾄﾞｰﾌﾟ量 = ﾄﾞｰﾊﾟﾝﾄ量指示(g) - ﾄﾞｰﾊﾟﾝﾄ希釈率(f) × 投入ﾄﾞｰﾊﾟﾝﾄ量合計(mg)/1000 - ﾄﾞｰﾊﾟﾝﾄ量実績(g)
'
'備考       計算以上発生時は、'-9999'を返す
'=============================================================================
Public Function DopeCalculationPart(strprmType As String, _
                                    dblprmRes As Double, _
                                    dblprmChrg As Double, _
                                    dblprmSijiDp As Double, _
                                    strprmChrgNo As String) As Variant
    
    Dim strType         As String       '狙いﾀｲﾌﾟ
    Dim dblRes          As Double       '狙い抵抗
    Dim dblChrg         As Double       'ﾁｬｰｼﾞ量
    Dim dblTgResBtm     As Double       '狙い抵抗(下限)
    Dim dblTgResTop     As Double       '狙い抵抗(上限)
    Dim dblData(0 To 2) As Double       'EE,EZ,EC
    Dim varmo           As Variant      'mo
    Dim varDilTmp       As Variant      'ﾄﾞｰﾊﾟﾝﾄ希釈率算出用変数(mo*W*α)
    Dim varDilution     As Variant      'ﾄﾞｰﾊﾟﾝﾄ希釈率
    Dim varCoefficient  As Variant      '軸係数
    Dim strErrData      As String       'ｴﾗｰﾃﾞｰﾀ
    '*ADD* TCS)K.Kunori 2004.12.16
    Dim dblSupplCoefficient As Double   '補正係数
    
    DopeCalculationPart = -9999
    
    On Error GoTo ErrHand
    
    '■■■■■■■■■■■■■■■
    '  必要ﾄﾞｰﾊﾟﾝﾄ量(Ｍ)算出処理
    '■■■■■■■■■■■■■■■
    
    '●必要ﾄﾞｰﾊﾟﾝﾄ量計算式●
    '計量ﾄﾞｰﾊﾟﾝﾄ量 = ﾄﾞｰﾊﾟﾝﾄ量指示 - ﾄﾞｰﾊﾟﾝﾄ希釈率(f) × 精製原料含有ﾄﾞｰﾊﾟﾝﾄ量合計 - ﾄﾞｰﾊﾟﾝﾄ量実績
    '>>> Ｍ = dblprmSijiDp - f * sDpData.PutDp - sDpData.InpDp
    'ﾄﾞｰﾊﾟﾝﾄ希釈率 = (mo * ﾁｬｰｼﾞ量 * 軸係数) / ﾄﾞｰﾊﾟﾝﾄ量実績
    '>>> f = (mo * dblChrg * α) / sDpData.InpDp
    '>>> mo = ((EC * log10(dblRes)) ^ 2) - (EE * log10(dblRes)) - EZ)
    
    '///ﾛｰｶﾙ変数に格納
    strType = strprmType    '狙いﾀｲﾌﾟ
    dblRes = dblprmRes      '狙い抵抗(ρ)
    dblChrg = dblprmChrg    'ﾁｬｰｼﾞ量
    
    '+++++++++++++++
    '  mo算出処理
    '+++++++++++++++
    
    '------------------------------
    '  狙い抵抗下限･上限取得処理
    '------------------------------
    '///狙い抵抗(下限)とρを比較して、ｷｰとなる条件(狙い抵抗下限･上限)を決定
    '///引数：狙いﾀｲﾌﾟ
    '   　　　狙い抵抗(ρ)
    '   　　　狙い抵抗(下限)
    '   　　　狙い抵抗(上限)
    If GetResData(strType, dblRes, dblTgResBtm, dblTgResTop) = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If
    
    '--------------
    '  mo算出処理
    '--------------
    '◆狙い抵抗(下限) <> 0 かつ 狙い抵抗(上限) <> 0 の場合
    If dblTgResBtm <> 0 Or dblTgResTop <> 0 Then
        
        '///ｷｰ条件により、mo算出用ﾃﾞｰﾀ取得(EC,EE,EZ)
        '///引数：狙いﾀｲﾌﾟ
        '   　　　狙い抵抗(下限)
        '   　　　狙い抵抗(上限)
        '   　　　EE,EZ,EC格納用変数
        If GetmoCalData(strType, dblTgResBtm, dblTgResTop, dblData()) = FUNCTION_RETURN_FAILURE Then
            Exit Function
        End If
        
        '///mo算出 ※Logに１０をかける
        varmo = 10 ^ ((dblData(2) * (Log10(dblRes)) ^ 2) - (dblData(0) * (Log10(dblRes))) - dblData(1))
        
    '◆ρが10.0以上の場合
    Else
        '///mo算出
        varmo = 0.501 * (dblRes ^ -1.0185)
    End If
        
    '++++++++++++++++++++++++++++
    '  ﾄﾞｰﾊﾟﾝﾄ希釈率(f)算出処理
    '++++++++++++++++++++++++++++
    
    '------------------
    '  軸係数取得処理
    '------------------
    '///ｷｰ条件により、軸係数を取得
    '///引数：狙いﾀｲﾌﾟ
    '   　　　軸
    '   　　　軸係数
    If GetCoefficientData(strType, varCoefficient, strprmChrgNo) = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If
    
    '----------------------------
    '  ﾄﾞｰﾊﾟﾝﾄ希釈率(f)算出処理
    '----------------------------
    '///mo*W*α算出 ※ﾁｬｰｼﾞ量をｷﾛｸﾞﾗﾑに換算する為に1000で割る
    varDilTmp = varmo * (dblChrg / 1000) * varCoefficient
    '◆(mo*W*α)が０以外の場合
    If varDilTmp <> 0 Then
        '///ﾄﾞｰﾊﾟﾝﾄ希釈率(f)算出 ※指示･ﾄﾞｰﾊﾟﾝﾄ量をﾐﾘｸﾞﾗﾑに換算する為に1000をかける
        '*CHG* 指示･ﾄﾞｰﾊﾟﾝﾄ量を使用するよう修正 TCS)K.Kunori 2004.11.24
'''        varDilution = (sDpData.InpDp * 1000) / varDilTmp
        varDilution = (dblprmSijiDp * 1000) / varDilTmp
    '◆(mo*W*α)が０の場合
    Else
        '///ﾄﾞｰﾊﾟﾝﾄ希釈率を０とする
        varDilution = 0
    End If
    
    '*ADD* 補正係数値取得処理追加 TCS)K.Kunori 2004.12.16 START >>>
    '-------------------
    '  補正係数値取得処
    '-------------------
    '///ﾄﾞｰﾊﾟﾝﾄ希釈率(f)により、補正係数値を取得
    '///引数：ﾄﾞｰﾊﾟﾝﾄ希釈率(f)
'    If GetSupplCoefficientData(CDbl(varDilution), dblSupplCoefficient) = FUNCTION_RETURN_FAILURE Then  '狙い抵抗別に補正係数を変更 2007/03/05 SETsw kubota
    If GetSupplCoefficientData(CDbl(varDilution), dblSupplCoefficient, strType, dblprmRes) = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If
    '*ADD* 補正係数値取得処理追加 TCS)K.Kunori 2004.12.16 END <<<
    
    '+++++++++++++++++++++++++
    '  計量ﾄﾞｰﾊﾟﾝﾄ量算出処理
    '+++++++++++++++++++++++++
    '///計量ﾄﾞｰﾊﾟﾝﾄ量(g)算出
    '*CHG* 補正係数追加 TCS)K.Kunori 2004.12.16
'''    DopeCalculationPart = CDec(dblprmSijiDp - CDec(varDilution) * sDpData.PutDp - sDpData.InpDp)
    DopeCalculationPart = CDec(dblprmSijiDp - _
                               CDec(varDilution) * dblSupplCoefficient * sDpData.PutDp - _
                               sDpData.InpDp)
    
    '---------------------------------------------------------
    '  精製原料濃度合計(×f×γ)＆ﾄﾞｰﾊﾟﾝﾄ希釈率(f)退避処理(表示用)
    '---------------------------------------------------------
    '///精製原料濃度合計(×f×γ)
    '*CHG* 補正係数追加の為計算式変更 TCS)K.Kunori 2004.12.16
'''    strRes1 = CStr(CDec(varDilution) * sDpData.PutDp)
    strRes1 = CStr(CDec(varDilution) * dblSupplCoefficient * sDpData.PutDp)
    '///ﾄﾞｰﾊﾟﾝﾄ希釈率(f)
    strRes2 = CStr(CDec(varDilution))
    '*ADD* 補正係数追加 TCS)K.Kunori 2004.12.16
    strRes3 = CStr(dblSupplCoefficient)
    
    '■ﾃｽﾄ用■
    Debug.Print "ﾀｲﾌﾟ：" & strType
    Debug.Print "ﾁｬｰｼﾞ量：" & CStr(CDec(dblChrg))
    Debug.Print "ρ：" & CStr(dblRes)
    Debug.Print "EE：" & dblData(0)
    Debug.Print "EZ：" & dblData(1)
    Debug.Print "EC：" & dblData(2)
    Debug.Print "軸係数：" & CStr(varCoefficient)
    Debug.Print "m0：" & CStr(CDec(varmo))
    Debug.Print "ﾄﾞｰﾊﾟﾝﾄ希釈率(f)：" & CStr(CDec(varDilution))
    Debug.Print "ﾄﾞｰﾊﾟﾝﾄ指示：" & CStr(CDec(dblprmSijiDp))
    Debug.Print "ﾄﾞｰﾊﾟﾝﾄ実績：" & CStr(CDec(sDpData.InpDp))
    Debug.Print "精製原料含有不純物量：" & CStr(CDec(sDpData.PutDp))
    Debug.Print "計量ﾄﾞｰﾌﾟ量：" & CStr(CDec(DopeCalculationPart))
    
    Exit Function

ErrHand:
    '///終了
    gErr.Pop
    Exit Function
End Function

'=============================================================================
'@(f)       GetResData
'
'機能       狙い抵抗下限･上限取得処理
'
'戻り値     True/False
'
'引数       狙いﾀｲﾌﾟ                    strType As String
'　　　     狙い抵抗(ρ)                dblRes As Double
'　　　     狙い抵抗(下限)              dblTgResBtm As Double
'　　　     狙い抵抗(上限)              dblTgResTop As Double
'
'機能概要
'
'備考
'=============================================================================
Public Function GetResData(strType As String, _
                           dblRes As Double, _
                           dblTgResBtm As Double, _
                           dblTgResTop As Double) As FUNCTION_RETURN
    
    GetResData = FUNCTION_RETURN_FAILURE
        
    '◆狙いﾀｲﾌﾟがＰの場合
    If strType = "P" Then
        Select Case dblRes
            '◆ρが10.0以上の場合
            Case Is >= 10
                dblTgResBtm = 10
                dblTgResTop = 99999
            '◆ρが1.0以上の場合
            Case Is >= 1
                dblTgResBtm = 1
                dblTgResTop = 10
            '◆ρが0.1以上の場合
            Case Is >= 0.1
                dblTgResBtm = 0.1
                dblTgResTop = 1
            '◆ρが0.0195以上の場合
            Case Is >= 0.0195
                dblTgResBtm = 0.0195
                dblTgResTop = 0.1
            '◆ρが0.01以上の場合
            Case Is >= 0.01
                dblTgResBtm = 0.01
                dblTgResTop = 0.0195
            '◆ρが0.005以上の場合
            Case Is >= 0.005
                dblTgResBtm = 0.005
                dblTgResTop = 0.01
            '◆ρが0.001以上の場合
            Case Is >= 0.001
                dblTgResBtm = 0.001
                dblTgResTop = 0.005
        End Select
    '◆狙いﾀｲﾌﾟがＮの場合
    ElseIf strType = "N" Then
        Select Case dblRes
            '◆ρが10.0以上の場合
            Case Is >= 10
                dblTgResBtm = 0
                dblTgResTop = 0
            '◆ρが1.0以上の場合
            Case Is >= 1
                dblTgResBtm = 1
                dblTgResTop = 10
            '◆ρが0.1以上の場合
            Case Is >= 0.1
                dblTgResBtm = 0.1
                dblTgResTop = 1
            '◆ρが0.245以上の場合
            Case Is >= 0.0245
                dblTgResBtm = 0.0245
                dblTgResTop = 0.1
            '◆ρが0.01以上の場合
            Case Is >= 0.01
                dblTgResBtm = 0.01
                dblTgResTop = 0.0245
            '◆ρが0.01より小さい場合
            Case Is < 0.01
                dblTgResBtm = 0
                dblTgResTop = 0.01
        End Select
    '◆狙いﾀｲﾌﾟがsbの場合
    ElseIf strType = "sb" Then
        Select Case dblRes
            '◆ρが0.05以上の場合
            Case Is >= 0.05
                dblTgResBtm = 0.05
                dblTgResTop = 0
            '◆ρが0.01以上の場合
            Case Is >= 0.015
                dblTgResBtm = 0.015
                dblTgResTop = 0.05
            '◆ρが0.01より小さい場合
            Case Is < 0.015
                dblTgResBtm = 0
                dblTgResTop = 0.015
        End Select
    End If
    
    GetResData = FUNCTION_RETURN_SUCCESS
    
    Exit Function

End Function

'=============================================================================
'@(f)       GetmoCalData
'
'機能       mo算出用ﾃﾞｰﾀ取得
'
'戻り値     True/False
'
'引数       狙いﾀｲﾌﾟ                    strType As String
'　　　     狙い抵抗(下限)              dblTgResBtm As Double
'　　　     狙い抵抗(上限)              dblTgResTop As Double
'　　　     EE,EZ,EC                    dblData() As Double
'
'機能概要
'
'備考
'=============================================================================
Public Function GetmoCalData(strType As String, _
                             dblTgResBtm As Double, _
                             dblTgResTop As Double, _
                             dblData() As Double) As FUNCTION_RETURN
    
    Dim strSql          As String       'SQL
    Dim rs              As OraDynaset   'ﾚｺｰﾄﾞｾｯﾄ
    Dim strTgResBtm     As String       '狙い抵抗(下限)
    Dim strTgResTop     As String       '狙い抵抗(上限)
    
    GetmoCalData = FUNCTION_RETURN_FAILURE
    
    '///型変換処理
    strTgResBtm = CStr(dblTgResBtm)     '狙い抵抗(下限)
    strTgResTop = CStr(dblTgResTop)     '狙い抵抗(上限)
    
    '-----------
    '  SQL発行
    '-----------
    '///条件：ｼｽﾃﾑ区分 = 'K'
    '   　　　種別ｺｰﾄﾞ = 'A5'
    '   　　　関連ｺｰﾄﾞ = strType(狙いﾀｲﾌﾟ)
    '　　　 　ﾃﾞｰﾀ1 　 = strTgResBtm(狙い抵抗(下限))
    '　　　   ﾃﾞｰﾀ2 　 = strTgResTop(狙い抵抗(上限))
    strSql = ""
    strSql = strSql & "SELECT KCODE03A9, KCODE04A9, KCODE05A9"
    strSql = strSql & "  FROM KODA9"
    strSql = strSql & " WHERE SYSCA9 = 'K'"
    strSql = strSql & "   AND SHUCA9 = 'A5'"
    strSql = strSql & "   AND KCODEA9 = '" & LCase$(strType) & "' "
    strSql = strSql & "   AND KCODE01A9 = '" & strTgResBtm & "' "
    strSql = strSql & "   AND KCODE02A9 = '" & strTgResTop & "' "
    
    Set rs = OraDB.DBCreateDynaset(strSql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        Exit Function
    End If
    
    '///取得ﾃﾞｰﾀ格納
    dblData(0) = rs("KCODE03A9")        'EE
    dblData(1) = rs("KCODE04A9")        'EZ
    dblData(2) = rs("KCODE05A9")        'EC
    
    rs.Close
    
    GetmoCalData = FUNCTION_RETURN_SUCCESS
    
    Exit Function

End Function

'=============================================================================
'@(f)       GetCoefficientData
'
'機能       軸係数取得処理
'
'戻り値     True/False
'
'引数       狙いﾀｲﾌﾟ                    strType As String
'　　　     軸係数                      varCoefficient As Variant
'           ﾁｬｰｼﾞ№                     strChrgNo As String(txtChargeNo.Text)
'
'機能概要
'
'備考
'=============================================================================
Public Function GetCoefficientData(strType As String, _
                                   varCoefficient As Variant, _
                                   strChrgNo As String) As FUNCTION_RETURN
    
    Dim strSql          As String       'SQL
    Dim rs              As OraDynaset   'ﾚｺｰﾄﾞｾｯﾄ
    Dim strCore         As String       '軸
    Dim strEdit         As String       '軸ﾃﾞｰﾀ編集用変数
    Dim intInstr        As Integer      '"<"開始位置

    GetCoefficientData = FUNCTION_RETURN_FAILURE

    '-----------
    '  SQL発行
    '-----------
    '///条件：ｼｽﾃﾑ区分         = 'SC'(TBCMB005)
    '   　　　種別ｺｰﾄﾞ         = '2'(TBCMB005)
    '   　　　ｺｰﾄﾞ             = TBCME018.HSXCDIR(TBCMB005)
    '　　　 　品番  　         = TBCMH001.HINBAN(TBCME018)
    '　　　   製品番号改訂番号 = TBCMH001.NMNOREVNO(TBCME018)
    '   　　　工場             = TBCMH001.NFACTORY(TBCME018)
    '         操業条件         = TBCMH001.NOPECOND(TBCME018)
    '         品SX結晶面方位   = TBCMB005.CODE(TBCME018)
    strSql = ""
    strSql = strSql & "SELECT DA9.KCODEA9 AS JIKU"
    strSql = strSql & "  FROM KODA9 DA9, TBCME018 TE18, TBCMH001 H01"
    strSql = strSql & " WHERE DA9.SYSCA9 = 'K'"
    strSql = strSql & "   AND DA9.SHUCA9 = 'AI'"
    strSql = strSql & "   AND H01.UPINDNO = '" & strChrgNo & "' "
    strSql = strSql & "   AND TE18.HINBAN = H01.HINBAN"
    strSql = strSql & "   AND TE18.MNOREVNO = H01.NMNOREVNO"
    strSql = strSql & "   AND TE18.FACTORY = H01.NFACTORY"
    strSql = strSql & "   AND TE18.OPECOND = H01.NOPECOND"
    '*CHG* 参照ｶﾗﾑ変更 TCS)K.Kunori 2004.11.24
'''    strSql = strSql & "   AND TRIM(DA9.CODEA9) = TRIM(TE18.HSXCDIR)"
    strSql = strSql & "   AND TRIM(DA9.CODEA9) = SUBSTR(TE18.MCNO,2,1)"
    
    Set rs = OraDB.DBCreateDynaset(strSql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        Exit Function
    End If
    
    '///軸ﾃﾞｰﾀ編集＆格納
    strCore = IIf(IsNull(rs("JIKU")), vbNullString, rs("JIKU")) '軸
    
    rs.Close
    
    '-----------
    '  SQL発行
    '-----------
    '///条件：ｼｽﾃﾑ区分 = 'K'
    '   　　　種別ｺｰﾄﾞ = 'A6'
    '   　　　関連ｺｰﾄﾞ = strType(狙いﾀｲﾌﾟ)
    '　　　   ﾃﾞｰﾀ2 　 = strCore(軸)
    strSql = ""
    strSql = strSql & "SELECT CTR01A9 AS COEFF"
    strSql = strSql & "  FROM KODA9"
    strSql = strSql & " WHERE SYSCA9 = 'K'"
    strSql = strSql & "   AND SHUCA9 = 'A6'"
    strSql = strSql & "   AND KCODE01A9 = '" & LCase$(strType) & "' "
    strSql = strSql & "   AND KCODE02A9 = '" & "<" & strCore & ">" & "' "
    
    Set rs = OraDB.DBCreateDynaset(strSql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        Exit Function
    End If
    
    '///取得ﾃﾞｰﾀ格納
    varCoefficient = CDbl(rs("COEFF"))    '軸係数
    
    rs.Close
    
    GetCoefficientData = FUNCTION_RETURN_SUCCESS
    
    Exit Function

End Function
'*ADD* 計量ﾄﾞｰﾌﾟ量計算方法変更対応 ※ﾓｼﾞｭｰﾙ統一 TCS)K.Kunori 2004.11.29 END <<<

'*ADD* 補正係数取得処理追加 TCS)K.Kunori 2004.12.16 START >>>
'=============================================================================
'@(f)       GetSupplCoefficientData
'
'機能       補正係数取得処理
'
'戻り値     True/False
'
'引数       ﾄﾞｰﾊﾟﾝﾄ希釈率(f)            varDilution As Variant
'　　　     補正係数                    varSupplCoefficient As Variant
'　　　     タイプ                      strType As String
'　　　     狙い抵抗                    dblNerai As String
'
'機能概要
'
'備考
'=============================================================================
Public Function GetSupplCoefficientData(dblDilution As Double _
                                      , dblSupplCoefficient As Double _
                                      , ByVal strType As String _
                                      , ByVal dblNerai As Double _
                                      ) As FUNCTION_RETURN
    
    Dim strSql          As String       'SQL
    Dim rs              As OraDynaset   'ﾚｺｰﾄﾞｾｯﾄ
    Dim dblfBtm         As Double       'ｆ値下限
    Dim dblfTop         As Double       'ｆ値上限
    Dim dblGanma        As Double       '補正係数(退避用)
    
    GetSupplCoefficientData = FUNCTION_RETURN_FAILURE
    
    '///ﾃﾞｰﾀ無の場合は'1'固定とする
    dblSupplCoefficient = 1
    
    '-----------
    '  SQL発行
    '-----------
    '///条件：ｼｽﾃﾑ区分 = 'K'
    '   　　　種別ｺｰﾄﾞ = 'AO'
    strSql = ""
    strSql = strSql & "SELECT KCODE01A9, KCODE02A9, KCODE03A9" & vbLf
    strSql = strSql & "  FROM KODA9" & vbLf
    strSql = strSql & " WHERE SYSCA9 = 'K'" & vbLf
    strSql = strSql & "   AND SHUCA9 = 'AO'"
    
    Set rs = OraDB.DBCreateDynaset(strSql, ORADYN_DEFAULT)
    
    '◆ﾃﾞｰﾀ有の場合
    If rs.RecordCount <> 0 Then
        
        '///取得ﾃﾞｰﾀ格納
        dblfBtm = val(NulltoStr(rs("KCODE01A9")))               'ｆ値下限
        dblfTop = val(NulltoStr(rs("KCODE02A9")))               'ｆ値上限
        dblGanma = val(NulltoStr(rs("KCODE03A9")))              '補正係数
        
        '◆ｆ値上限が０以外の場合
        If dblfTop <> 0 Then
            '◆ｆ値が範囲内の場合(下限値以上上限値未満)
            If dblfBtm <= dblDilution And dblDilution < dblfTop Then
                dblSupplCoefficient = dblGanma
                '補正係数算出方法変更 2007/03/05追加 SETsw kubota
                If GetHosei_Nerai(strType, CStr(dblNerai), dblSupplCoefficient) = False Then
                    Exit Function
                End If
            End If
        '◆ｆ値上限が０(NULL)の場合
        Else
            '◆ｆ値が下限値以上の場合
            If dblfBtm <= dblDilution Then
                dblSupplCoefficient = dblGanma
                '補正係数算出方法変更 2007/03/05追加 SETsw kubota
                If GetHosei_Nerai(strType, CStr(dblNerai), dblSupplCoefficient) = False Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    rs.Close
    
    GetSupplCoefficientData = FUNCTION_RETURN_SUCCESS
    
    Exit Function

End Function
'*ADD* 補正係数取得処理追加 TCS)K.Kunori 2004.12.16 END <<<

'*ADD* 狙い抵抗別補正係数取得処理追加(200mm関数のコピー) SETsw kubota 2007.03.05 START >>>
'概要      :補正係数の取得
'ﾊﾟﾗﾒｰﾀ(In):タイプ
'           ねらい抵抗
'          :戻り値：正常／異常
'説明      :
'履歴      :2007.02.19 作成
Public Function GetHosei_Nerai(ByVal sType As String _
                             , ByVal sNerai As String _
                             , ByRef dblSupplCoefficient As Double _
                             ) As Boolean

    Dim sSql        As String
    Dim objDS       As Object
    Dim dblLow      As Double
    Dim dblHigh     As Double
    Dim dblG        As Double
    Dim dblNerai    As Double
    Dim lCnt        As Long
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    
    GetHosei_Nerai = False
    
    dblNerai = val(sNerai)
    
    'コード設定なし等の場合、補正係数を設定せずに終了する
    '''gdblGanma = 1   '存在しない場合、１固定とする。
    
    '補正係数取得
    sSql = ""
    sSql = sSql & "SELECT kcode01a9,kcode02a9,ctr01a9" & vbLf
    sSql = sSql & "  FROM KODA9 " & vbLf
    sSql = sSql & " WHERE SYSCA9 = 'X'" & vbCrLf
    sSql = sSql & "   AND SHUCA9 = 'RS'" & vbCrLf
    sSql = sSql & "   AND CODEA9 LIKE '" & Left$(UCase$(sType), 1) & "%'" & vbCrLf
    sSql = sSql & " ORDER BY CODEA9" & vbCrLf
    
    Set objDS = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    
    For lCnt = 1 To objDS.RecordCount
    
        dblLow = val(NulltoStr(objDS(0)))
        dblHigh = val(NulltoStr(objDS(1)))
        dblG = val(NulltoStr(objDS(2)))
        
        If dblLow = 0 Then
            '下限値がNull(= 0)の場合は、上限値未満の時取得データを設定する。
            If dblHigh > dblNerai Then
                dblSupplCoefficient = dblG
                Exit For
            End If
        ElseIf dblHigh = 0 Then
            '上限値がNull(= 0)の場合は、下限値以上の時取得データを設定する。
            If dblLow <= dblNerai Then
                dblSupplCoefficient = dblG
                Exit For
            End If
        Else
            '上下限値が設定されている場合は、上下限範囲内の時取得データを設定する
            If dblLow <= dblNerai And dblNerai < dblHigh Then
                dblSupplCoefficient = dblG
                Exit For
            End If
        End If
        objDS.MoveNext
    
    Next lCnt
    
    GetHosei_Nerai = True
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit
    
End Function
'*ADD* 狙い抵抗別補正係数取得処理追加(200mm関数のコピー) SETsw kubota 2007.03.05 END <<<

'///////////////////////////////////////////////////
' @(f)
' 機能    : 抵抗濃度取得 Kameda
' 返り値  : True:正常
'           False:失敗
' 引き数  : ﾄﾞｰﾊﾟﾝﾄ種類
' 機能説明: 抵抗濃度から登録偏析係数を読み込む
' 修正履歴: 濃度(CTR01A)少数以下４桁→３桁   2009/12/17
'///////////////////////////////////////////////////
Public Function Getrnoudo(sDopant As String, sPuro As Double) As Double
    Dim sSqlStmt As String
    Dim objOraDyn As Object
    Dim lCount As Long
    Dim syubetu As String
    Dim sSql As String
    
    Getrnoudo = 0
    
    'SQL編集
    sSqlStmt = "SELECT  NVL(shuca9   , ' ')"                  ' 0:[種別コード]
    sSqlStmt = sSqlStmt & ",NVL(kcode01a9, ' ')"                  ' 1:[データ１] … 抵抗下限
    sSqlStmt = sSqlStmt & ",NVL(kcode02a9, ' ')"                  ' 2:[データ２] … 抵抗上限
    sSqlStmt = sSqlStmt & ",NVL(ctr01a9, 0)"                      ' 3:[ｶｳﾝﾀｰ１]　… ドーパント濃度
    sSqlStmt = sSqlStmt & "  FROM koda9 "
    sSqlStmt = sSqlStmt & " WHERE sysca9 = 'X'"
    sSqlStmt = sSqlStmt & "   AND shuca9 >= 'D0'"
    sSqlStmt = sSqlStmt & "   AND shuca9 <= 'D9'"
    sSqlStmt = sSqlStmt & "   AND codea9 = '" & sDopant & "'"
    sSqlStmt = sSqlStmt & " ORDER BY shuca9"
    
    ''ダイナセット作成
    If DynSet2(objOraDyn, sSqlStmt) = False Then
        Exit Function
    End If
    If objOraDyn.EOF Then
        ''該当する種類が無かった
        Exit Function
    End If
    
    For lCount = 1 To objOraDyn.RecordCount
    
       If objOraDyn(1) <= sPuro Or objOraDyn(1) = "" Then
          If sPuro < objOraDyn(2) Or objOraDyn(2) = "" Then
                Getrnoudo = objOraDyn(3) / 10
          End If
       End If
       objOraDyn.MoveNext
   Next lCount
      

End Function
'///////////////////////////////////////////////////
' @(f)
' 機能    : ドーパント種類取得 kameda
' 返り値  : True:正常
'           False:失敗
' 引き数  :
' 機能説明:
' 修正履歴:
'///////////////////////////////////////////////////
Public Function GetDopeKind(sDopeKind() As String) As FUNCTION_RETURN
    Dim sSqlStmt As String
    Dim objOraDyn As Object
    Dim lCount As Long
    Dim syubetu As String
    Dim sSql As String
    
    GetDopeKind = FUNCTION_RETURN_FAILURE
    
    'SQL編集
    sSqlStmt = "SELECT  NVL(codea9   , ' ')"
    sSqlStmt = sSqlStmt & "  FROM koda9 "
    sSqlStmt = sSqlStmt & " WHERE sysca9 = 'X'"
    sSqlStmt = sSqlStmt & "   AND shuca9 = 'D0'"
    sSqlStmt = sSqlStmt & " ORDER BY codea9"
    
    ''ダイナセット作成
    If DynSet2(objOraDyn, sSqlStmt) = False Then
        Exit Function
    End If
    If objOraDyn.EOF Then
        ''該当する種類が無かった
        Exit Function
    End If
    
    ReDim sDopeKind(objOraDyn.RecordCount)
    
    For lCount = 1 To objOraDyn.RecordCount
          sDopeKind(lCount) = objOraDyn(0)
       objOraDyn.MoveNext
   Next lCount
      
      GetDopeKind = FUNCTION_RETURN_SUCCESS

End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    : ドーパント計算用ファクター・偏析係数取得
'
' 返り値  : false:失敗
'           true:取得件数
'
' 引き数  : 処理区分
'           タイプ:"P"/"N"
'           ファクター(OUT)
'           偏析係数(OUT)
'
' 機能説明: ドーパント計算用ファクター・偏析係数取得
'///////////////////////////////////////////////////
Function GetFactor(ByVal sType As String, ByRef sngFactor As Single, ByRef sngHenseki As Single) As Boolean
    GetFactor = False
    
    Dim sSqlStmt As String
    Dim objOraDyn As Object
    
    
    ''ＳＱＬ文作成
    sSqlStmt = "SELECT NVL(kcodea9, ' '),                   "
    sSqlStmt = sSqlStmt & "NVL(kcode01a9, ' ')                "
    sSqlStmt = sSqlStmt & "FROM koda9                       "
    sSqlStmt = sSqlStmt & "WHERE sysca9 = 'X'               "
    sSqlStmt = sSqlStmt & "  AND shuca9 = '36'              "
    sSqlStmt = sSqlStmt & "  AND codea9 = '" & sType & "' "
    
    ''ダイナセット作成
    If DynSet2(objOraDyn, sSqlStmt) = False Then
        ''ダイナセット作成失敗
        Call MsgOut(100, sSqlStmt, ERR_DISP_LOG)
        
        GetFactor = False
        Exit Function
    End If
    If objOraDyn.EOF Then
        GetFactor = False
        Exit Function
    End If

    sngHenseki = objOraDyn(0)   ''偏析係数
    sngFactor = objOraDyn(1)    ''ファクター
   
    GetFactor = True
End Function

