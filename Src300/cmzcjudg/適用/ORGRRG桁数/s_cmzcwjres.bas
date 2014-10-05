Attribute VB_Name = "s_cmzcwjres"
Option Explicit

''WFセンター比抵抗判定構造体
Type W_RES
    GuaranteeRes    As Guarantee    ''品質保証情報構造体
    GuaranteeCal    As String * 1   ''品WF比抵抗面内計算
    SpecResMin      As Double       ''品WF比抵抗下限
    SpecResMax      As Double       ''品WF比抵抗上限
    SpecRrg         As Double       ''品WF比抵抗面内分布
    SpecResAveMin   As Double       ''品WF比抵抗平均下限
    SpecResAveMax   As Double       ''品WF比抵抗平均上限
    Res(4)          As Double       ''比抵抗測定値
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    'AN温度分を追加
    ResAntnp        As Double       ''AN温度比抵抗測定値
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RRG             As Double       ''RRG計算値
    JudgRes         As Boolean      ''比抵抗判定値
    JudgRrg         As Boolean      ''RRG判定結果
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    'チェック用AN温度を追加
    JudgAntnp       As Boolean      ''ＡＮ温度判定結果
    Antnp           As Integer      ''品ＷＦＡＮ温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    DkTmpSiyo       As String       ''DK温度（仕様）
    DkTmpJsk        As String       ''DK温度（実績）
    JudgDkTmp       As Boolean      ''DK温度判定値
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
End Type

'概要      :WFセンター比抵抗判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Res           ,I  ,W_RES            ,WFセンター比抵抗判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WfRESJudg(Res As W_RES, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFRES_JUDG = 1                 ''判定識別フラグ(RES)
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim center As Double
    Dim r2 As Double
    Dim AllTemp() As Double
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
    Dim liRet As Integer
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Dim iRet    As Integer
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    
    Res.JudgRrg = JUDG_NG
    Res.JudgRes = JUDG_NG
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Res.JudgDkTmp = JUDG_NG
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    If Res.GuaranteeRes.cJudg = JudgCodeW01 Then ''RES判定有り
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "比抵抗判定 **********"
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_点 = " & Res.GuaranteeRes.cCount
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_方 = " & Res.GuaranteeRes.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_位 = " & Res.GuaranteeRes.cPos
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_対 = " & Res.GuaranteeRes.cObj
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_処 = " & Res.GuaranteeRes.cJudg
'''''        WFCJudgDialog.WFCErrorMessage "分布計算 = " & Res.GuaranteeCal
'''''        WFCJudgDialog.WFCErrorMessage "比抵抗面内分布 = " & Str(Res.SpecRrg)

        
        ''RRG判定
        'RRGの小数桁数を6桁(7桁目四捨五入)に変更(ここでは丸めない) 2011/11/25 SETsw kubota
        'Res.RRG = WFCRRGCal(Res.Res(), Res.GuaranteeRes, Res.GuaranteeCal)
        Res.RRG = WFCRRGCal_NotRound(Res.Res(), Res.GuaranteeRes, Res.GuaranteeCal)
'2002/02/27 S.Sano RRGの仕様が0の場合は、判定を行わず必ずOKとする。
'2002/02/27 S.Sano 面内分布計算は行う。
        If Res.SpecRrg = 0 Then                                     '2002/02/27 S.Sano
            Res.JudgRrg = JUDG_OK                                   '2002/02/27 S.Sano
        Else                                                        '2002/02/27 S.Sano
            If Res.RRG = -1 Then
                Res.JudgRrg = JUDG_NG
            Else
                Res.JudgRrg = RangeDecision_nl(Res.RRG, 0, Res.SpecRrg)
            End If
        End If                                                      '2002/02/27 S.Sano
        
        ''RES判定
        If (InStr(ObjCodeGrp01, Res.GuaranteeRes.cObj) <> 0) Then
            Select Case Res.GuaranteeRes.cObj
            Case ObjCode01  ''中心1点
                center = WFCCenterP(WFRES_JUDG, Res.Res(), Res.GuaranteeRes)
                If center = -9999 Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
'''''                    WFCJudgDialog.WFCErrorMessage "比抵抗判定、対象データ無し。"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
                Else
                    Res.JudgRes = RangeDecision_nl(center, Res.SpecResMin, Res.SpecResMax)
                End If
            Case ObjCode02  ''中央値
                center = WFCCenterD(WFRES_JUDG, Res.Res(), Res.GuaranteeRes)
                If center = -9999 Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
'''''                    WFCJudgDialog.WFCErrorMessage "比抵抗判定、対象データ無し。"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
                Else
                    Res.JudgRes = RangeDecision_nl(center, Res.SpecResMin, Res.SpecResMax)
                End If
            Case ObjCode03 ''全域
                If WFCJudgDataSelect_All(Res.Res(), Res.GuaranteeRes, AllTemp()) = FUNCTION_RETURN_FAILURE Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
'''''                    WFCJudgDialog.WFCErrorMessage "比抵抗判定、対象データ無し。"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
                Else
                    Res.JudgRes = JUDG_OK
                    For c0 = 0 To UBound(AllTemp())
                        If RangeDecision_nl(AllTemp(c0), Res.SpecResMin, Res.SpecResMax) = JUDG_NG Then
                            Res.JudgRes = JUDG_NG
                        End If
                    Next
                End If
            Case ObjCode04 ''R/2
                r2 = WFCR2(WFRES_JUDG, Res.Res(), Res.GuaranteeRes)
                If r2 = -9999 Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
'''''                    WFCJudgDialog.WFCErrorMessage "比抵抗判定、対象データ無し。"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
                Else
                    Res.JudgRes = RangeDecision_nl(r2, Res.SpecResMin, Res.SpecResMax)
                End If
            End Select
        Else
            If (Res.GuaranteeRes.cObj <> ObjCode13) And (Res.GuaranteeRes.cObj <> ObjCode15) Then
                ''対象データ無し
                ''エラー情報構造体に情報を代入。
'''''                WFCJudgDialog.WFCErrorMessage "比抵抗判定、対象データ無し。"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
            End If
        End If
    
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
        ''AN温度判定
        liRet = funCodeDBGetMatrixReturn("SB", "AR", CStr(Res.Antnp), CStr(Res.ResAntnp))
        If liRet = -1 Then
            FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
        ElseIf liRet = 0 Then
            Res.JudgAntnp = JUDG_NG
        Else
            Res.JudgAntnp = JUDG_OK
        End If
        
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------

'--------------- 2008/08/25 INSERT START  By Systech ---------------
        ''DK温度判定
        If Trim(Res.DkTmpJsk) = "" Or Trim(Res.DkTmpSiyo) = "" Then
            Res.JudgDkTmp = JUDG_OK
        Else
            iRet = funCodeDBGetMatrixReturn(DKTMP_TBCMB005SYS, DKTMP_TBCMB005CLS, Res.DkTmpJsk, Res.DkTmpSiyo)
            If iRet = -1 Then
                FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
            ElseIf iRet = 0 Then
                Res.JudgDkTmp = JUDG_NG
            Else
                Res.JudgDkTmp = JUDG_OK
            End If
        End If
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
    Else
        Res.JudgRes = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -s-
        Res.JudgAntnp = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -e-
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        Res.JudgDkTmp = JUDG_OK
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
'        If InStr(JudgCodeW02, Res.GuaranteeRes.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, RES_JUDG, Res.GuaranteeRes.cJudg)
'        End If
    End If

    WfRESJudg = FuncAns
End Function

'概要      :RRGを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
'端数丸めを行う関数と行わない関数の二つに分離 2011/11/25 SETsw kubota
'Public Function WFCRRGCal(R() As Double, G As Guarantee, calcode As String) As Double
Public Function WFCRRGCal_NotRound(R() As Double, G As Guarantee, calcode As String) As Double
    Dim Min As Double
    Dim max As Double
    Dim AVE As Double
    Dim center As Double
    Dim side As Double
    Dim side_ce As Double
    Dim r2 As Double
    Dim RRG As Double
    Dim errflag As Boolean
    Dim deverrflag As Boolean
    Dim C As Double
    
    errflag = False
    deverrflag = False
    
    RRG = -1
    
    Select Case calcode
    Case "A" '(max-min)/min×100
        Min = WFCMin(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If (Min <> 0) Then
                RRG = (max - Min) * 100 / Min
            Else
                deverrflag = True
            End If
        End If
    Case "B" '(max-min)/max×100
        Min = WFCMin(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If (max <> 0) Then
                RRG = (max - Min) * 100 / max
            Else
                deverrflag = True
            End If
        End If
    Case "C" '(max-min)/center×100
        Min = WFCMin(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        center = WFCCenterP(WFRES_JUDG, R(), G)
        If (Min <> -9999) And (max <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                RRG = (max - Min) * 100 / center
            Else
                deverrflag = True
            End If
        End If
    Case "D" '|center-side|max/center×100
        side = WFCCE_Side_Max(WFRES_JUDG, R(), G)
        center = WFCCenterP(WFRES_JUDG, R(), G)
        If (side <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                RRG = side * 100 / center
            Else
                deverrflag = True
            End If
        End If
    Case "E" '(centerave-sideave)/centerave×100
        center = WFCCEAve(WFRES_JUDG, R(), G)
        side = WFCSideAve(WFRES_JUDG, R(), G)
        If (side <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                RRG = Abs(center - side) * 100 / center  '|center-side| 2002/6/28 osawa
            Else
                deverrflag = True
            End If
        End If
    Case "F" '|center-R/2|max/center×100
        r2 = WFCCE_R2_Max(WFRES_JUDG, R(), G)
        center = WFCCenterP(WFRES_JUDG, R(), G)
        If (r2 <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                RRG = r2 * 100 / center
            Else
                deverrflag = True
            End If
        End If
    Case "G" '2(side-center)/(sideave+center)×100
        side_ce = WFCCE_Side_Max(WFRES_JUDG, R(), G)
        side = WFCSideAve(WFRES_JUDG, R(), G)
        center = WFCCenterP(WFRES_JUDG, R(), G)
        If (side_ce <> -9999) And (side <> -9999) And (center <> -9999) Then
            If ((side + center) <> 0) Then
                RRG = 2 * Abs(side_ce) * 100 / (side + center)  '|side_ce| 2002/6/28 osawa
            Else
                deverrflag = True
            End If
        End If
    Case "H" '(max-ave)/ave×100
        AVE = WFCAve(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        If (max <> -9999) And (AVE <> -9999) Then
            If (AVE <> 0) Then
                RRG = (max - AVE) * 100 / AVE
            Else
                deverrflag = True
            End If
        End If
    Case "K" '(max-min)/(max+min)×100
        Min = WFCMin(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If ((max + Min) <> 0) Then
                RRG = (max - AVE) * 100 / (max + Min)
            Else
                deverrflag = True
            End If
        End If
    Case "L" '(max-min)/2×ave×100
        Min = WFCMin(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        AVE = WFCAve(WFRES_JUDG, R(), G)
        If (Min <> -9999) And (max <> -9999) And (AVE <> -9999) Then
            If (AVE <> 0) Then
                RRG = (max - Min) * 100 / 2 * AVE
            Else
                deverrflag = True
            End If
        End If
    Case "M" '(max-min)/ave×100
        Min = WFCMin(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        AVE = WFCAve(WFRES_JUDG, R(), G)
        If (Min <> -9999) And (max <> -9999) And (AVE <> -9999) Then
            If (AVE <> 0) Then
                RRG = (max - Min) * 100 / AVE
            Else
                deverrflag = True
            End If
        End If

'コード"N"に対応　受付No.20409  <2002.10.11 yakimura> start
    
    Case "N" '|(center-side)/(center+side)|×200
        
        C = WFCCenterP(WFRES_JUDG, R(), G)
        'RRG = WF_TypeN_Exc(WFOI_JUDG, R(), G, C)
        RRG = WF_TypeN_Exc(WFRES_JUDG, R(), G, C)                       '2003/5/15
        
        If RRG = -9999 Then
              errflag = True
        End If

'コード"N"に対応　受付No.20409  <2002.10.11 yakimura> end
    
    Case " "

'''''        WFCJudgDialog.WFCErrorMessage "分布計算未定義 A にて計算"

        Min = WFCMin(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If (Min <> 0) Then
                RRG = (max - Min) * 100 / Min
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
    ElseIf RRG = -1 Then
'''''        WFCJudgDialog.WFCErrorMessage "測定位置、対象データ、分布計算矛盾"
    End If

    '2002/07/24 Update T.Hayashi
    'WFCRRGCal = RRG
    'WFCRRGCal = RoundUp(RRG, 4)
    WFCRRGCal_NotRound = RRG        '丸めない値を返すように変更 2011/11/25 SETsw kubota

End Function

'端数丸めを行う関数と行わない関数の二つに分離 2011/11/25 SETsw kubota
Public Function WFCRRGCal(R() As Double, G As Guarantee, calcode As String) As Double
    '端数丸めを行わない関数を呼び出し、小数2桁(3桁目切り上げ)にして返す
    WFCRRGCal = RoundUp(WFCRRGCal_NotRound(R(), G, calcode), 4)
End Function


Private Function WFCJudgDataSelect_All(d() As Double, G As Guarantee, T() As Double) As FUNCTION_RETURN
    Dim Func_Ans As FUNCTION_RETURN
    
    Func_Ans = FUNCTION_RETURN_FAILURE
    
    Select Case G.cPos
    Case "1"                                 '2003/05/15 追加　osawa 依頼No.030130
        ReDim T(2) As Double
        T(0) = d(0)
        T(1) = d(3)
        T(2) = d(4)
        Func_Ans = FUNCTION_RETURN_SUCCESS   '依頼No.030130　追加ここまで
    Case "2", "3", "4"
        ReDim T(1) As Double
        T(0) = d(0)
        T(1) = d(4)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "5", "6", "7", "8"
        ReDim T(1) As Double
        T(0) = d(1)
        T(1) = d(4)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "A", "L"
        ReDim T(1) As Double
        T(0) = d(2)
        T(1) = d(4)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "E"
        ReDim T(1) As Double
        T(0) = d(3)
        T(1) = d(4)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "G", "H"
        ReDim T(2) As Double
        T(0) = d(1)
        T(1) = d(3)
        T(2) = d(4)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "J", "M"
        ReDim T(2) As Double
        T(0) = d(2)
        T(1) = d(3)
        T(2) = d(4)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "Y"
        ReDim T(0) As Double
        T(0) = d(4)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    Case "P"
        ReDim T(4) As Double
        T(0) = d(0)
        T(1) = d(1)
        T(2) = d(2)
        T(3) = d(3)
        T(4) = d(4)
        Func_Ans = FUNCTION_RETURN_SUCCESS
    End Select

    WFCJudgDataSelect_All = Func_Ans
End Function
