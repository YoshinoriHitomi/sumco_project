Attribute VB_Name = "s_cmzcwjoi"
Option Explicit
''WFセンターOi判定構造体
Type W_OI
    GuaranteeOi     As Guarantee    ''品質保証情報構造体
    GuaranteeCal    As String * 1   ''品WF酸素濃度面内計算
    SpecOiMin       As Double       ''品WF酸素濃度下限
    SpecOiMax       As Double       ''品WF酸素濃度上限
    SpecORG         As Double       ''品WF酸素濃度面内分布
    SpecOiAveMin    As Double       ''品WF酸素濃度平均下限
    SpecOiAveMax    As Double       ''品WF酸素濃度平均上限
    Oi(9)           As Double       ''Oi測定値
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    OiAntnp         As Double       ''AN温度Oi測定値
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    ORG             As Double       ''ORG計算値
    OiMin           As Double       ''OiMin計算値
    OiMax           As Double       ''OiMax計算値
    JudgOi          As Boolean      ''Oi判定結果
    JudgOrg         As Boolean      ''ORG判定結果
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    'チェック用AN温度を追加
    JudgAntnp       As Boolean      ''ＡＮ温度判定結果
    Antnp           As Integer      ''品ＷＦＡＮ温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
End Type

'概要      :WFセンターOi判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Oi            ,I  ,W_OI             ,WFセンターOi判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WfOiJudg(Oi As W_OI, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFOI_JUDG = 2                  ''判定識別フラグ(Oi)
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim center As Double
    Dim r2 As Double
    Dim AllTemp() As Double
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
    Dim liRet As Integer
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    

    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Oi.JudgOi = JUDG_NG
    Oi.JudgOrg = JUDG_NG
    If Oi.GuaranteeOi.cJudg = JudgCodeW01 Then ''Oi判定有り
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "酸素濃度判定 **********"
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_点 = " & Oi.GuaranteeOi.cCount
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_方 = " & Oi.GuaranteeOi.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_位 = " & Oi.GuaranteeOi.cPos
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_対 = " & Oi.GuaranteeOi.cObj
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_処 = " & Oi.GuaranteeOi.cJudg
'''''        WFCJudgDialog.WFCErrorMessage "分布計算 = " & Oi.GuaranteeCal
'''''        WFCJudgDialog.WFCErrorMessage "酸素濃度面内分布 = " & Str(Oi.SpecORG)

        ''ORG判定
        'ORGの小数桁数を6桁(7桁目四捨五入)に変更(ここでは丸めない) 2011/11/25 SETsw kubota
        'Oi.ORG = WFCORGCal(Oi.Oi(), Oi.GuaranteeOi, Oi.GuaranteeCal)
        Oi.ORG = WFCORGCal_NotRound(Oi.Oi(), Oi.GuaranteeOi, Oi.GuaranteeCal)
        Oi.OiMin = WFCMin(WFOI_JUDG, Oi.Oi(), Oi.GuaranteeOi)
        Oi.OiMax = WFCMax(WFOI_JUDG, Oi.Oi(), Oi.GuaranteeOi)
        
'2002/02/27 S.Sano ORGの仕様が0の場合は、判定を行わず必ずOKとする。
'2002/02/27 S.Sano 面内分布計算は行う。
        If Oi.SpecORG = 0 Then                                      '2002/02/27 S.Sano
            Oi.JudgOrg = JUDG_OK                                    '2002/02/27 S.Sano
        Else                                                        '2002/02/27 S.Sano
            If Oi.ORG = -1 Then
                Oi.JudgOrg = JUDG_NG
            Else
                Oi.JudgOrg = RangeDecision_nl(Oi.ORG, 0, Oi.SpecORG)
            End If
        End If                                                      '2002/02/27 S.Sano
        
        ''Oi判定
        If (InStr(ObjCodeGrp01, Oi.GuaranteeOi.cObj) <> 0) Then
            Select Case Oi.GuaranteeOi.cObj
            Case ObjCode01  ''中心1点
                center = WFCCenterP(WFOI_JUDG, Oi.Oi(), Oi.GuaranteeOi)
                If center = -9999 Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
'''''                    WFCJudgDialog.WFCErrorMessage "酸素濃度判定、対象データ無し。"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
                Else
                    Oi.JudgOi = RangeDecision_nl(center, Oi.SpecOiMin, Oi.SpecOiMax)
                End If
            Case ObjCode02  ''中央値
                center = WFCCenterP(WFOI_JUDG, Oi.Oi(), Oi.GuaranteeOi)
                If center = -9999 Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
'''''                    WFCJudgDialog.WFCErrorMessage "酸素濃度判定、対象データ無し。"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
                Else
                    Oi.JudgOi = RangeDecision_nl(center, Oi.SpecOiMin, Oi.SpecOiMax)
                End If
            Case ObjCode03 ''全域
                If WFCJudgDataSelect_All(Oi.Oi(), Oi.GuaranteeOi, AllTemp()) = FUNCTION_RETURN_FAILURE Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
'''''                    WFCJudgDialog.WFCErrorMessage "酸素濃度判定、対象データ無し。"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
                Else
                    Oi.JudgOi = JUDG_OK
                    For c0 = 0 To UBound(AllTemp())
                        If RangeDecision_nl(AllTemp(c0), Oi.SpecOiMin, Oi.SpecOiMax) = JUDG_NG Then
                            Oi.JudgOi = JUDG_NG
                        End If
                    Next
                End If
            Case ObjCode04 ''R/2
                r2 = WFCR2(WFOI_JUDG, Oi.Oi(), Oi.GuaranteeOi)
                If r2 = -9999 Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
'''''                    WFCJudgDialog.WFCErrorMessage "酸素濃度判定、対象データ無し。"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
                Else
                    Oi.JudgOi = RangeDecision_nl(r2, Oi.SpecOiMin, Oi.SpecOiMax)
                End If
            End Select
        Else
            ''狙い、規格無し以外の場合
            If (Oi.GuaranteeOi.cObj <> ObjCode13) And (Oi.GuaranteeOi.cObj <> ObjCode15) Then
                ''対象データ無し
                ''エラー情報構造体に情報を代入。
'''''                WFCJudgDialog.WFCErrorMessage "酸素濃度判定、対象データ無し。"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
            End If
        End If
        
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
        ''AN温度判定
        liRet = funCodeDBGetMatrixReturn("SB", "AO", CStr(Oi.Antnp), CStr(Oi.OiAntnp))
        If liRet = -1 Then
            FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
        ElseIf liRet = 0 Then
            Oi.JudgAntnp = JUDG_NG
        Else
            Oi.JudgAntnp = JUDG_OK
        End If
                
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    Else
        Oi.JudgOi = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -s-
        Oi.JudgAntnp = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -e-
'        If InStr(JudgCodeW02, Oi.GuaranteeOi.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, OI_JUDG, Oi.GuaranteeOi.cJudg)
'        End If
    End If
    
    WfOiJudg = FuncAns
End Function

'概要      :ORGを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :d()           ,I  ,double    ,測定値
'          :iMax          ,I  ,Integer   ,測定点数
'          :戻り値        ,O  ,double    ,org,ORG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
'端数丸めを行う関数と行わない関数の二つに分離 2011/11/25 SETsw kubota
'Public Function WFCORGCal(O() As Double, G As Guarantee, calcode As String) As Double
Public Function WFCORGCal_NotRound(O() As Double, G As Guarantee, calcode As String) As Double
    Dim Min As Double
    Dim max As Double
    Dim AVE As Double
    Dim center As Double
    Dim side As Double
    Dim side_ce As Double
    Dim r2 As Double
    Dim ORG As Double
    Dim errflag As Boolean
    Dim deverrflag As Boolean
    Dim C As Double
    
    errflag = False
    deverrflag = False
    
    ORG = -1
    
    Select Case calcode
    Case "A" '(max-min)/min×100
        Min = WFCMin(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If (Min <> 0) Then
                ORG = (max - Min) * 100 / Min
            Else
                deverrflag = True
            End If
        End If
    Case "B" '(max-min)/max×100
        Min = WFCMin(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If (max <> 0) Then
                ORG = (max - Min) * 100 / max
            Else
                deverrflag = True
            End If
        End If
    Case "C" '(max-min)/center×100
        Min = WFCMin(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        center = WFCCenterP(WFOI_JUDG, O(), G)
        If (Min <> -9999) And (max <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                ORG = (max - Min) * 100 / center
            Else
                deverrflag = True
            End If
        End If
    Case "D" '|center-side|max/center×100
        side = WFCCE_Side_Max(WFOI_JUDG, O(), G)
        center = WFCCenterP(WFOI_JUDG, O(), G)
        If (side <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                ORG = side * 100 / center
            Else
                deverrflag = True
            End If
        End If
    Case "E" '(centerave-sideave)/centerave×100
        center = WFCCEAve(WFOI_JUDG, O(), G)
        side = WFCSideAve(WFOI_JUDG, O(), G)
        If (side <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                ORG = Abs(center - side) * 100 / center  '|center-side| 2002/6/28 osawa
            Else
                deverrflag = True
            End If
        End If
    Case "F" '|center-R/2|max/center×100
        r2 = WFCCE_R2_Max(WFOI_JUDG, O(), G)
        center = WFCCenterP(WFOI_JUDG, O(), G)
        If (r2 <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                ORG = r2 * 100 / center
            Else
                deverrflag = True
            End If
        End If
    Case "G" '2(|center-side|max)/(sideave+center)×100
        side_ce = WFCCE_Side_Max(WFOI_JUDG, O(), G)
        side = WFCSideAve(WFOI_JUDG, O(), G)
        center = WFCCenterP(WFOI_JUDG, O(), G)
        If (side_ce <> -9999) And (side <> -9999) And (center <> -9999) Then
            If ((side + center) <> 0) Then
                ORG = 2 * Abs(side_ce) * 100 / (side + center)  '|side_ce|  2002/6/28 osawa
            Else
                deverrflag = True
            End If
        End If
    Case "H" '(max-ave)/ave×100
        AVE = WFCAve(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        If (max <> -9999) And (AVE <> -9999) Then
            If (AVE <> 0) Then
                ORG = (max - AVE) * 100 / AVE
            Else
                deverrflag = True
            End If
        End If
    Case "K" '(max-min)/(max+min)×100
        Min = WFCMin(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If ((max + Min) <> 0) Then
                ORG = (max - AVE) * 100 / (max + Min)
            Else
                deverrflag = True
            End If
        End If
    Case "L" '(max-min)/2×ave×100
        Min = WFCMin(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        AVE = WFCAve(WFOI_JUDG, O(), G)
        If (Min <> -9999) And (max <> -9999) And (AVE <> -9999) Then
            If (AVE <> 0) Then
                ORG = (max - Min) * 100 / 2 * AVE
            Else
                deverrflag = True
            End If
        End If
    Case "M" '(max-min)/ave×100
        Min = WFCMin(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        AVE = WFCAve(WFOI_JUDG, O(), G)
        If (Min <> -9999) And (max <> -9999) And (AVE <> -9999) Then
            If (AVE <> 0) Then
                ORG = (max - Min) * 100 / AVE
            Else
                deverrflag = True
            End If
        End If

'コード"N"に対応　受付No.20409  <2002.10.11 yakimura> start
    
    Case "N" '|(center-side)/(center+side)|×200
        
        C = WFCCenterP(WFOI_JUDG, O(), G)
        ORG = WF_TypeN_Exc(WFOI_JUDG, O(), G, C)
        
        If ORG = -9999 Then
              errflag = True
        End If

'コード"N"に対応　受付No.20409  <2002.10.11 yakimura> end
    
    Case " "

'''''        WFCJudgDialog.WFCErrorMessage "分布計算未定義 A にて計算"
        Min = WFCMin(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If (Min <> 0) Then
                ORG = (max - Min) * 100 / Min
            Else
                deverrflag = True
            End If
        End If
    Case Else
        errflag = True
    End Select


    If errflag Then
'''''        WFCJudgDialog.WFCErrorMessage "分布計算未定義"
    ElseIf deverrflag Then
'''''        WFCJudgDialog.WFCErrorMessage "分布計算 0 除算エラー"
    ElseIf ORG = -1 Then
'''''        WFCJudgDialog.WFCErrorMessage "測定位置、対象データ、分布計算矛盾"
    End If
    
    '2002/07/24 Update T.Hayashi
    'WFCORGCal = ORG
    'WFCORGCal = RoundUp(ORG, 2)
    WFCORGCal_NotRound = ORG        '丸めない値を返すように変更 2011/11/25 SETsw kubota
    
End Function

'端数丸めを行う関数と行わない関数の二つに分離 2011/11/25 SETsw kubota
Public Function WFCORGCal(O() As Double, G As Guarantee, calcode As String) As Double
    '端数丸めを行わない関数を呼び出し、小数2桁(3桁目切り上げ)にして返す
    WFCORGCal = RoundUp(WFCORGCal_NotRound(O(), G, calcode), 2)
End Function


Private Function WFCJudgDataSelect_All(d() As Double, G As Guarantee, T() As Double) As FUNCTION_RETURN
    Dim Func_Ans As FUNCTION_RETURN
    
    Func_Ans = FUNCTION_RETURN_FAILURE
    
    Select Case G.cPos
    Case "1"                                  '2003/05/15 追加　osawa 依頼No.030130
        ReDim T(3) As Double
        T(0) = d(0)
        T(1) = d(2)
        T(2) = d(6)
        T(3) = d(9)
        Func_Ans = FUNCTION_RETURN_SUCCESS    '依頼No.030130　追加ここまで
    Case "2", "3", "4", "5"
        ReDim T(2) As Double
        T(0) = d(0)
        T(1) = d(6)
        T(2) = d(9)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "6", "7", "8"
        ReDim T(2) As Double
        T(0) = d(0)
        T(1) = d(5)
        T(2) = d(8)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "A", "L"
        ReDim T(2) As Double
        T(0) = d(0)
        T(1) = d(4)
        T(2) = d(7)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "D"
        ReDim T(1) As Double
        T(0) = d(0)
        T(1) = d(3)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "E"
        ReDim T(1) As Double
        T(0) = d(0)
        T(1) = d(2)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "F"
        ReDim T(1) As Double
        T(0) = d(0)
        T(1) = d(1)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "G"
        ReDim T(3) As Double
        T(0) = d(0)
        T(1) = d(2)
        T(2) = d(6)
        T(3) = d(9)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "H"
        ReDim T(3) As Double
        T(0) = d(0)
        T(1) = d(2)
        T(2) = d(5)
        T(3) = d(8)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "J", "M"
        ReDim T(3) As Double
        T(0) = d(0)
        T(1) = d(2)
        T(2) = d(4)
        T(3) = d(7)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "K"
        ReDim T(2) As Double
        T(0) = d(0)
        T(1) = d(1)
        T(2) = d(3)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "Y"
        ReDim T(0) As Double
        T(0) = d(0)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "R"
        ReDim T(9) As Double
        T(0) = d(0)
        T(1) = d(1)
        T(2) = d(2)
        T(3) = d(3)
        T(4) = d(4)
        T(5) = d(5)
        T(6) = d(6)
        T(7) = d(7)
        T(8) = d(8)
        T(9) = d(9)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    End Select

    WFCJudgDataSelect_All = Func_Ans
End Function
