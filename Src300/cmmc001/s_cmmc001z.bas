Attribute VB_Name = "s_cmmc001z"

'概要      :偏析計算に必要な各合計重量実績を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :CRYNUM        ,I  ,String    ,結晶番号
'          :wgtCharge     ,O  ,Long      ,炉内量（初回チャージ量－前回までの引上げ重量－前回までのﾄｯﾌﾟｶｯﾄ重量）
'          :wgtTop        ,O  ,Double    ,トップ重量実績値
'          :wgtTopCut     ,O  ,Double    ,トップカット重量実績値
'          :DM            ,O  ,Double    ,直径１～３の平均
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :１本引き、残量引きにあわせて実績データを取得する
'履歴      :2001/8/29 作成  野村
Public Function GetCoeffParams(ByVal CRYNUM$, wgtCharge As Long, wgtTop As Double, wgtTopCut As Double, DM As Double) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset

    On Error GoTo Err
    GetCoeffParams = FUNCTION_RETURN_FAILURE
    wgtCharge = 0
    wgtTop = 0#
    wgtTopCut = 0#
    DM = 0#
    
    sql = "select decode(RONAI,null,CHARGE,RONAI) as RONAI, WGHTTOP, WGTOPCUT, (DM1+DM2+DM3)/3.0 as DM " & _
          "from TBCMH004 H004, " & _
          "  (select sum(CHARGE) - sum(UPWEIGHT) - sum(WGTOPCUT) as RONAI" & _
          "   From TBCMH004" & _
          "   where (CRYNUM<'" & CRYNUM & "')" & _
          "    and  (substr(CRYNUM,1,7)='" & left$(CRYNUM, 7) & "')" & _
          "  ) SUMDATA " & _
          "where (CRYNUM='" & CRYNUM & "')"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        wgtCharge = rs("RONAI")
        wgtTop = rs("WGHTTOP")
        wgtTopCut = rs("WGTOPCUT")
        DM = rs("DM")
    End If
    rs.Close
    
    GetCoeffParams = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    On Error GoTo 0
    Exit Function

Err:
    Resume proc_exit
End Function


'概要      :抵抗値に対する位置を推定する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型             ,説明
'          :d             ,IO ,type_ResPosCal ,推定計算構造体
'          :戻り値        ,O  ,Double         ,推定位置
'説明      :
'履歴      :2001/06/23　佐野 信哉　作成
Public Function PosCalculation(d As type_ResPosCal) As Double
    Dim GS As Double        'ρTop位置引上げ率
    Dim Ro As Double        '基準抵抗値
    Dim Gx As Double
    
    On Error GoTo Err
    GS = (d.DUNMENSEKI * HIJU_SILICONE * d.TOPSMPLPOS) / (d.CHARGEWEIGHT - d.TOPWEIGHT)
    Ro = d.TOPRES * ((1 - GS) ^ (d.COEFFICIENT - 1))
    Gx = 1 - ((Ro / d.target) ^ (1 / (d.COEFFICIENT - 1)))
    
    PosCalculation = ((d.CHARGEWEIGHT - d.TOPWEIGHT) * Gx) / (d.DUNMENSEKI * HIJU_SILICONE)
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    PosCalculation = -9999
End Function

'概要      :位置に対する抵抗値を推定する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型             ,説明
'          :d             ,IO ,type_ResPosCal ,推定計算構造体
'          :戻り値        ,O  ,Double         ,推定抵抗値
'説明      :
'履歴      :2001/06/23　佐野 信哉　作成
Public Function ResCalculation(d As type_ResPosCal) As Double
    Dim GS As Double        'ρTop位置引上げ率
    Dim Ro As Double        '基準抵抗値
    Dim Gx As Double        '推定対象引上げ率

    On Error GoTo Err
    GS = (d.DUNMENSEKI * HIJU_SILICONE * d.TOPSMPLPOS) / (d.CHARGEWEIGHT - d.TOPWEIGHT)
    Ro = d.TOPRES * (1 - GS) ^ (d.COEFFICIENT - 1)
    Gx = d.DUNMENSEKI * d.target * HIJU_SILICONE / (d.CHARGEWEIGHT - d.TOPWEIGHT)

    ResCalculation = Ro / (1 - Gx) ^ (d.COEFFICIENT - 1)
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    ResCalculation = -9999
End Function

'概要      :偏析係数を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :d             ,IO ,type_Coefficient ,偏析係数計算構造体
'          :戻り値        ,O  ,Double           ,偏析係数
'説明      :
'履歴      :2001/06/23　佐野 信哉　作成
Public Function CoefficientCalculation(d As type_Coefficient) As Double
    Dim GT As Double
    Dim GB As Double
    
    On Error GoTo Err
    GT = (d.DUNMENSEKI * d.TOPSMPLPOS * HIJU_SILICONE) / (d.CHARGEWEIGHT - d.TOPWEIGHT)
    GB = (d.DUNMENSEKI * d.BOTSMPLPOS * HIJU_SILICONE) / (d.CHARGEWEIGHT - d.TOPWEIGHT)
    
    CoefficientCalculation = Log(d.BOTRES / (d.TOPRES * 1)) / Log((1 - GT) / (1 - GB)) + 1
    
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    CoefficientCalculation = -9999
End Function


'概要      :シリコン円柱の重量を求める
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dblDiameter   ,I  ,Double    ,直径(mm)
'          :dblHeight     ,I  ,Double    ,高さ(mm)
'          :戻り値        ,O  ,Double    ,重量(g)
'説明      :
'履歴      :2001/06/29 作成  野村
Public Function WeightOfCylinder(ByVal dblDiameter As Double, ByVal dblHeight As Double) As Double
Dim dblRadius As Double

    dblRadius = dblDiameter / 2#
    WeightOfCylinder = HIJU_SILICONE * cdblPI * (dblRadius ^ 2) * dblHeight
End Function


'概要      :シリコン円錐の重量を求める
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dblDiameter   ,I  ,Double    ,直径(mm)
'          :dblHeight     ,I  ,Double    ,高さ(mm)
'          :戻り値        ,O  ,Double    ,重量(g)
'説明      :TOP・BOT重量の計算用
'履歴      :2001/06/29 作成  野村
Public Function WeightOfCone(ByVal dblDiameter As Double, ByVal dblHeight As Double) As Double
Dim dblRadius As Double

    dblRadius = dblDiameter / 2#
    WeightOfCone = HIJU_SILICONE * (cdblPI * (dblRadius ^ 2) * dblHeight) / 3#
End Function


'概要      :円の面積を求める
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dblDiameter   ,I  ,Double    ,直径(mm)
'          :戻り値        ,O  ,Double    ,面積(mm2)
'説明      :
'履歴      :2001/07/05 作成  野村
Public Function AreaOfCircle(ByVal dblDiameter As Double) As Double
Dim dblRadius As Double

    dblRadius = dblDiameter / 2#
    AreaOfCircle = cdblPI * (dblRadius ^ 2)
End Function


'概要      :偏析計算に必要な各合計重量実績を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :CRYNUM        ,I  ,String    ,結晶番号
'          :wgtCharge     ,O  ,Long      ,炉内量（初回チャージ量－前回までの引上げ重量－前回までのﾄｯﾌﾟｶｯﾄ重量）
'          :wgtTop        ,O  ,Double    ,トップ重量実績値
'          :wgtTopCut     ,O  ,Double    ,トップカット重量実績値
'          :DM            ,O  ,Double    ,直径１～３の平均
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :【マルチ引上対応】 全量引き､残量引き､RC引きにあわせて実績データを取得する
'履歴      :2008/4/21 作成  SETsw Nakada
Public Function GetCoeffParams_new(ByVal CRYNUM$, wgtCharge As Long, wgtTop As Double, wgtTopCut As Double, DM As Double) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset

    On Error GoTo Err
    GetCoeffParams_new = FUNCTION_RETURN_FAILURE
    wgtCharge = 0
    wgtTop = 0#
    wgtTopCut = 0#
    DM = 0#
    
    '' 推定チャージ、重量（TOP）、トップカット重量、直胴直径の平均値 取得
    sql = " SELECT C1.SUICHARGE, C1.WGHTTOC1, C1.PUTCUTWC1, "
    sql = sql & " (C1.DIA1C1 + C1.DIA2C1 + C1.DIA3C1) / 3.0 AS DM "
    sql = sql & " FROM XSDC1 C1 "
    sql = sql & " WHERE C1.XTALC1 = '" & CRYNUM & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        wgtCharge = rs("SUICHARGE")       ''推定チャージ
        wgtTop = rs("WGHTTOC1")           ''重量（TOP）
        wgtTopCut = rs("PUTCUTWC1")       ''トップカット重量
        DM = rs("DM")                     ''直胴直径(平均値)
    End If
    rs.Close
    
    GetCoeffParams_new = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    On Error GoTo 0
    Exit Function

Err:
    Resume proc_exit
End Function

''2011/01/17 tkimura ADD START ==========================================================>
'概要      :偏析計算に必要な各合計重量実績を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :CRYNUM        ,I  ,String    ,結晶番号
'          :wgtCharge     ,O  ,Long      ,炉内量
'          :wgtChargeA    ,O  ,Long      ,A結晶の炉内量
'          :wgtTop        ,O  ,Double    ,トップ重量実績値
'          :wgtTopCut     ,O  ,Double    ,トップカット重量実績値
'          :DM            ,O  ,Double    ,直径１～３の平均
'          :hikiFlg       ,O  ,Integer   ,引上げフラグ(1=通常、2=BC結晶)
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :【マルチ引上対応】 全量引き､残量引き､RC引きにあわせて実績データを取得する
'履歴      :2008/4/21 作成  SETsw Nakada
Public Function GetCoeffParams_new2(ByVal CRYNUM$, _
                                    wgtCharge As Long, _
                                    wgtChargeA As Long, _
                                    wgtTop As Double, _
                                    wgtTopCut As Double, _
                                    DM As Double, _
                                    HIKIFLG As Integer) As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim cryNumA As String       'BC結晶処理でのA結晶を格納する。

    On Error GoTo Err
    GetCoeffParams_new2 = FUNCTION_RETURN_FAILURE
    wgtCharge = 0
    wgtChargeA = 0
    wgtTop = 0#
    wgtTopCut = 0#
    DM = 0#
    
    '' 推定チャージ、重量（TOP）、トップカット重量、直胴直径の平均値 取得
    sql = " SELECT C1.SUICHARGE, C1.WGHTTOC1, C1.PUTCUTWC1, "
    sql = sql & " (C1.DIA1C1 + C1.DIA2C1 + C1.DIA3C1) / 3.0 AS DM "
    sql = sql & " FROM XSDC1 C1 "
    sql = sql & " WHERE C1.XTALC1 = '" & CRYNUM & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        wgtCharge = rs("SUICHARGE")       ''推定チャージ
        wgtTop = rs("WGHTTOC1")           ''重量（TOP）
        wgtTopCut = rs("PUTCUTWC1")       ''トップカット重量
        DM = rs("DM")                     ''直胴直径(平均値)
    End If
    rs.Close
    
    '結晶番号の9桁がBorCならばBC結晶となる。
    If Mid(CRYNUM, 9, 1) = "B" Or Mid(CRYNUM, 9, 1) = "C" Then
        HIKIFLG = "2"       'BC結晶
    Else
        HIKIFLG = "1"       '通常
    End If
    
    'このあとにwgtChargeAを求める必要がある。(HIKIFLG="2"のときのみ)
    If HIKIFLG = "2" Then
        cryNumA = Mid(CRYNUM, 1, 8) & "A" & Mid(CRYNUM, 10, 3)      '結晶番号の9桁目をAにする。
        sql = " SELECT C1.SUICHARGE "
        sql = sql & " FROM XSDC1 C1 "
        sql = sql & " WHERE C1.XTALC1 = '" & cryNumA & "'"

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        If rs.RecordCount > 0 Then
            wgtChargeA = rs("SUICHARGE")       ''推定チャージ
        End If
        rs.Close
    End If
    
    Set rs = Nothing
    GetCoeffParams_new2 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    On Error GoTo 0
    Exit Function

Err:
    Resume proc_exit
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'概要      :偏析係数を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                       ,説明
'          :d             ,I ,type_Coefficient_new2     ,推定抵抗,推定引上率計算構造体
'          :戻り値        ,O  ,Double                   ,偏析係数
'説明      :
'履歴      :2001/06/23　佐野 信哉　作成
Public Function CoefficientCalculation_new2(ByRef d As type_Coefficient_new2) As Double
    
    On Error GoTo Err
    
    CoefficientCalculation_new2 = Log(d.BOTRES / (d.TOPRES * 1)) / Log((1 - d.GT) / (1 - d.GB)) + 1
    
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    CoefficientCalculation_new2 = -9999
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'概要       :引上げ率を計算する。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :d              ,I  ,type_Coefficient_new2      ,推定抵抗,推定引上率計算構造体
'           :戻り値         ,O  ,Double                     ,位置引上率
'説明       :
'履歴       :2011/01/17 tkimura
Public Function HikiageCalculation(ByRef d As type_Coefficient_new2) As Double
    Dim result As Double

    '通常
    If d.HIKIFLG = "1" Then
        result = (d.DUNMENSEKI * d.SMPLPOS * HIJU_SILICONE + d.TOPWEIGHT) / (d.CHARGEWEIGHT)
    'BC結晶
    Else
        result = (d.DUNMENSEKI * d.SMPLPOS * HIJU_SILICONE + d.TOPWEIGHT + d.CHARGEWEIGHTA - d.CHARGEWEIGHT) / (d.CHARGEWEIGHTA)
    End If
    
    HikiageCalculation = result
    
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'概要       :基準抵抗値を計算する。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :d              ,I  ,type_Coefficient_new2      ,推定抵抗,推定引上率計算構造体
'           :戻り値         ,O ,Double                      ,基準抵抗値
'説明       :
'履歴       :2011/01/17 tkimura
Public Function StandardResCalculation(d As type_Coefficient_new2) As Double
    
    On Error GoTo Err
        
    StandardResCalculation = d.TOPRES * (1 - d.GT) ^ (d.Henseki - 1)
        
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    StandardResCalculation = -9999
    
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'概要       :推定位置比抵抗値を計算する。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :d              ,I  ,type_Coefficient_new2      ,推定抵抗,推定引上率計算構造体
'           :戻り値         ,O ,Double                      ,推定位置比抵抗値
'説明       :
'履歴       :2011/01/17 tkimura
Public Function SuiteiResCalculation(ByRef d As type_Coefficient_new2) As Double
    
    On Error GoTo Err
        
    SuiteiResCalculation = d.KIJUNTEIKOU / (1 - d.SUITEIHIKIRITU) ^ (d.Henseki - 1)
        
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    SuiteiResCalculation = -9999
    
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================
