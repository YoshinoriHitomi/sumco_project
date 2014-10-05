Attribute VB_Name = "s_cmzcwj"
Option Explicit

''WFセンターΔOi判定構造体
Type W_DOI
    GuaranteeDoi    As Guarantee    ''品質保証情報構造体
    SpecDoiMin      As Double       ''品WF酸素析出1～3下限
    SpecDoiMax      As Double       ''品WF酸素析出1～3上限
    Doi(5)          As Double       ''ΔOi測定値
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    DoiAntnp        As Double      ''ＡＮ温度ΔOi測定値
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    JudgDoi         As Boolean      ''ΔOi判定結果
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    'チェック用AN温度を追加
    JudgAntnp       As Boolean      ''ＡＮ温度判定結果
    Antnp           As Integer      ''品ＷＦＡＮ温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
End Type

''WFセンターAOi判定構造体　03/12/09 ooba
Type W_AOI
    GuaranteeAoi    As Guarantee    ''品質保証情報構造体
    SpecAoiMin      As Double       ''品WF残存酸素下限
    SpecAoiMax      As Double       ''品WF残存酸素上限
    AOI(2)          As Double       ''AOi測定値
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    AoiAntnp        As Double      ''ＡＮ温度AOi測定値
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    JudgAoi         As Boolean      ''AOi判定結果
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    'チェック用AN温度を追加
    JudgAntnp       As Boolean      ''ＡＮ温度判定結果
    Antnp           As Integer      ''品ＷＦＡＮ温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
End Type

''WFセンターOSF判定構造体
Type W_OSF
    GuaranteeOsf    As Guarantee    ''品質保証情報構造体
    SpecOsfAveMax   As Double       ''品WFOSF平均上限
    SpecOsfMax      As Double       ''品WFOSF上限
    OSF(4)          As Double       ''OSF測定値
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    OsfAntnp        As Double      ''AN温度OSF測定値
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    OSFp(2)         As String * 1   ''OSFパターン実績　2003/05/17 ooba
    Min             As Double       ''最小値
    max             As Double       ''最大値
    AVE             As Double       ''平均値
    JudgOsf         As Boolean      ''OSF判定結果
    JudgDataMin     As Double       ''最少判定値
    JudgDataMax     As Double       ''最大判定値
    JudgDataAve     As Double       ''平均判定値
    JudgDataPTK     As String * 1   ''パターン区分　2003/05/17 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    'チェック用AN温度を追加
    JudgAntnp       As Boolean      ''ＡＮ温度判定結果
    Antnp           As Integer      ''品ＷＦＡＮ温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
End Type

''WFセンターBMD判定構造体
Type W_BMD
    GuaranteeBmd    As Guarantee    ''品質保証情報構造体
    SpecBmdAveMin   As Double       ''品WFBMD平均下限
    SpecBmdAveMax   As Double       ''品WFBMD平均上限
    SpecBmdGsAveMin   As Double     ''BMD平均下限(外周)　09/05/07 ooba
    SpecBmdGsAveMax   As Double     ''BMD平均上限(外周)　09/05/07 ooba
    SpecBmdMBP      As Double       ''品WFBMD面内分布　2003/05/20 ooba
    SpecBmdMCL      As String * 2   ''品WFBMD面内計算　2003/05/20 ooba
'    BMD(3)          As Double       ''BMD測定値
    BMD(4)          As Double       ''BMD測定値　2003/05/20 ooba　5点対応
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    'AN温度分を追加
    BmdAntnp        As Double      ''AN温度BMD測定値
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Min             As Double       ''最小値
    max             As Double       ''最大値
    AVE             As Double       ''平均値
    JudgBmd         As Boolean      ''BMD判定結果
    JudgDataMin     As Double       ''最少判定値
    JudgDataMax     As Double       ''最大判定値
    JudgDataAve     As Double       ''平均判定値
    JudgDataMBP     As Double       ''面内分布判定値　2003/05/20 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    'チェック用AN温度を追加
    JudgAntnp       As Boolean      ''ＡＮ温度判定結果
    Antnp           As Integer      ''品ＷＦＡＮ温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
End Type

''WFセンターDZ判定構造体
Type W_DZ
    GuaranteeDz     As Guarantee    ''品質保証情報構造体
    SpecDzMin       As Double       ''品WF無欠陥層下限
    SpecDzMax       As Double       ''品WF無欠陥層上限
    DZ(3)           As Double       ''DZ測定値
    JudgDz          As Boolean      ''DZ判定結果
    JudgDataMin     As Double       ''最少判定値
    JudgDataMax     As Double       ''最大判定値
    JudgDataAve     As Double       ''平均判定値
End Type

''WFセンターDSOD判定構造体
Type W_DSOD
    GuaranteeDsod   As Guarantee    ''品質保証情報構造体
    SpecDsodMin     As Double       ''品WFDSOD下限
    SpecDsodMax     As Double       ''品WFDSOD上限
    Dsod            As Double       ''DSOD測定値
    Dsodp(1)        As String * 3   ''DSODパターン実績　04/07/23 ooba
    JudgDataPTK     As String * 1   ''DSODパターン区分　04/07/23 ooba
    JudgDsod        As Boolean      ''DSOD判定結果
    DsodAntnp       As Double       ''AN温度DSOD測定値  06/12/22 ooba
    JudgAntnp       As Boolean      ''AN温度判定結果    06/12/22 ooba
    Antnp           As Integer      ''品WFAN温度        06/12/22 ooba
End Type

''WFセンターSPV判定構造体
Type W_SPV
    GuaranteeSpv    As Guarantee    ''品質保証情報構造体
    GuaranteeSpvFe  As Guarantee    ''品質保証情報構造体
    SpecSpvMin      As Double       ''品WF拡散長下限
    SpecSpvMax      As Double       ''品WF拡散長上限
    SpecSpvFeMax    As Double       ''品WFFe濃度上限
    Spv(5)          As Double       ''SPV測定値
    JudgSpv         As Boolean      ''SPV判定結果
    '-----TEST2004/10
    SpecSpvAvMax      As Double       ''品WF平均上限
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
    GuaranteeSpvNr  As Guarantee    ''品質保証情報構造体
    SpecSpvNrMax    As Double       ''品WFNr濃度上限
    SpecSpvNrAvMax  As Double       ''品WFNr平均上限
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
End Type

''WFセンターGD判定構造体
Public Type W_GD
    GuaranteeDen        As Guarantee    ''品質保証情報構造体
    GuaranteeLdl        As Guarantee    ''品質保証情報構造体
    GuaranteeDvd2       As Guarantee    ''品質保証情報構造体
    JudgFlagDen         As String * 1   ''品WFDen検査有無
    JudgFlagLdl         As String * 1   ''品WFL/DL検査有無
    JudgFlagDvd2        As String * 1   ''品WFDVD2検査有無
    SpecDenMin          As Double       ''品WFDen下限
    SpecDenMax          As Double       ''品WFDen上限
    SpecLdlMin          As Double       ''品WFL/DL下限
    SpecLdlMax          As Double       ''品WFL/DL上限
    SpecDvd2Min         As Double       ''品WFDVD2下限
    SpecDvd2Max         As Double       ''品WFDVD2上限
'*** UPDATE ↓ Y.SIMIZU 2005/10/1 品WFGDﾗｲﾝ数
    SpecGdLine          As Single       ''品WFGDﾗｲﾝ数
'*** UPDATE ↑ Y.SIMIZU 2005/10/1 品WFGDﾗｲﾝ数
    Den                 As Double       ''Den計算値
    Ldl                 As Double       ''L/DL計算値
    Dvd2                As Double       ''DVD2計算値
    JudgDen             As Boolean      ''Den判定結果
    JudgLdl             As Boolean      ''L/DL判定結果
    JudgDvd2            As Boolean      ''DVD2判定結果
    GdAntnp             As Double       ''AN温度GD測定値    06/12/22 ooba
    JudgAntnp           As Boolean      ''AN温度判定結果    06/12/22 ooba
    Antnp               As Integer      ''品WFAN温度        06/12/22 ooba
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    GDPTK               As String * 1   ''品ＷＦＧＤパタン区分
    LdlMin              As Integer      ''L/DL連続0MIN
    LdlMax              As Integer      ''L/DL連続0MAX
    ZeroLdlMin          As Integer      ''品SXLdl連続0下限
    ZeroLdlMax          As Integer      ''品SXLdl連続0上限
    JudgLdlPtn          As Boolean      ''L/DLパターン判定結果
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
End Type

''↓Add 2010/01/07 SIRD対応 Y.Hitomi
''WFセンターSIRD判定構造体
Type W_SD
    GuaranteeSd         As Guarantee    ''品質保証情報構造体
    SpecSdMax           As Integer      ''軸状転位(SIRD)上限
    SdMeasData          As Integer      ''SIRD測定結果
    JudgSD              As Boolean      ''SIRD判定結果
End Type
''↑Add 2010/01/07 SIRD対応 Y.Hitomi

'概要      :WFセンターΔOi判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Doi           ,I  ,W_DOI            ,WFセンターΔOi判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WfDOiJudg(Doi As W_DOI, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFDOI_JUDG = 3                 ''判定識別フラグ(ΔOi)
    Dim FuncAns As FUNCTION_RETURN
    Dim TempDOi(2) As Double
    Dim JData(2) As Double
    Dim c0 As Integer
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
    Dim liRet           As Integer
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Doi.JudgDoi = JUDG_NG
    If Doi.GuaranteeDoi.cJudg = JudgCodeW01 Then ''DOi判定有り
        
''''''        WFCJudgDialog.WFCErrorMessage " "
''''''        WFCJudgDialog.WFCErrorMessage "ΔOi判定 **************"
''''''        WFCJudgDialog.WFCErrorMessage "測定位置_点 = " & Doi.GuaranteeDoi.cCount
''''''        WFCJudgDialog.WFCErrorMessage "測定位置_方 = " & Doi.GuaranteeDoi.cMeth
''''''        WFCJudgDialog.WFCErrorMessage "測定位置_位 = " & Doi.GuaranteeDoi.cPos
''''''        WFCJudgDialog.WFCErrorMessage "保証方法_対 = " & Doi.GuaranteeDoi.cObj
''''''        WFCJudgDialog.WFCErrorMessage "保証方法_処 = " & Doi.GuaranteeDoi.cJudg
        
        ''ΔOi = Initial_Oi - After_Oi
        ''中心から外周へ並べ替え
        For c0 = 0 To 2
            TempDOi(c0) = Doi.Doi(2 - c0) - Doi.Doi(5 - c0)
        Next
        
        ''DOi判定
        FuncAns = GetWfJudgData(WFDOI_JUDG, Doi.GuaranteeDoi, TempDOi(), JData())
        If (InStr(ObjCodeGrp01, Doi.GuaranteeDoi.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
            Select Case Doi.GuaranteeDoi.cObj
            Case ObjCode01, ObjCode02  ''中心1点、中央値
                Doi.JudgDoi = RangeDecision_nl(JData(0), Doi.SpecDoiMin, Doi.SpecDoiMax)
            Case ObjCode03 ''全域
                Doi.JudgDoi = JUDG_OK
                For c0 = 0 To 2
                    If JData(c0) <> -1 Then
                        If RangeDecision_nl(JData(c0), Doi.SpecDoiMin, Doi.SpecDoiMax) = JUDG_NG Then
                            Doi.JudgDoi = JUDG_NG
                        End If
                    End If
                Next
            Case ObjCode04 ''R/2''''''''''''''''''''''''''''''''''''''''''''''''''''''''''未定
'''''                WFCJudgDialog.WFCErrorMessage "ΔOi判定、対象データ無し。"
            End Select
        Else
            ''狙い、規格無し以外の場合
            If (Doi.GuaranteeDoi.cObj <> ObjCode13) And (Doi.GuaranteeDoi.cObj <> ObjCode15) Then
                ''対象データ無し
                ''エラー情報構造体に情報を代入。
'''''                WFCJudgDialog.WFCErrorMessage "ΔOi判定、対象データ無し。"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, DOI_JUDG, Doi.GuaranteeDoi.cObj)
            End If
        End If
    
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
        'AN温度チェックを追加
        ''AN温度判定
        'マトリックスからチェックの成否を取得
        liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(Doi.Antnp), CStr(Doi.DoiAntnp))
        If liRet = -1 Then
            FuncAns = SetErrInfo(ErrInfo, EZJ00, DOI_JUDG, Doi.GuaranteeDoi.cObj)
        ElseIf liRet = 0 Then
            Doi.JudgAntnp = JUDG_NG
        Else
            Doi.JudgAntnp = JUDG_OK
        End If
        
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Else
        Doi.JudgDoi = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -s-
        Doi.JudgAntnp = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -e-
'        If InStr(JudgCodeW02, Doi.GuaranteeDoi.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, DOI_JUDG, Doi.GuaranteeDoi.cJudg)
'        End If
    End If
    
    WfDOiJudg = FuncAns
End Function

'概要      :WFセンターAOi判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Aoi           ,I  ,W_AOI            ,WFセンターAOi判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :03/12/09 ooba

Public Function WfAOiJudg(AOI As W_AOI, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN

    Dim FuncAns As FUNCTION_RETURN
    Dim TempAOi(2) As Double
    Dim JData(2) As Double
    Dim c0 As Integer
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
    Dim liRet As Integer
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    AOI.JudgAoi = JUDG_NG
    If AOI.GuaranteeAoi.cJudg = JudgCodeW01 Then ''AOi判定有り
        
        ''ΔOi = Initial_Oi - After_Oi
        ''中心から外周へ並べ替え
        For c0 = 0 To 2
'''            TempDOi(c0) = Doi.Doi(2 - c0) - Doi.Doi(5 - c0)
            TempAOi(c0) = AOI.AOI(2 - c0)
        Next
        
        ''AOi判定
        FuncAns = GetWfJudgData(WFAOI_JUDG, AOI.GuaranteeAoi, TempAOi(), JData())
        If (InStr(ObjCodeGrp01, AOI.GuaranteeAoi.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
            Select Case AOI.GuaranteeAoi.cObj
            Case ObjCode01, ObjCode02  ''中心1点、中央値
                AOI.JudgAoi = RangeDecision_nl(JData(0), AOI.SpecAoiMin, AOI.SpecAoiMax)
            Case ObjCode03 ''全域
                AOI.JudgAoi = JUDG_OK
                For c0 = 0 To 2
                    If JData(c0) <> -1 Then
                        If RangeDecision_nl(JData(c0), AOI.SpecAoiMin, AOI.SpecAoiMax) = JUDG_NG Then
                            AOI.JudgAoi = JUDG_NG
                        End If
                    End If
                Next
            Case ObjCode04 ''R/2''''''''''''''''''''''''''''''''''''''''''''''''''''''''''未定
'''''                WFCJudgDialog.WFCErrorMessage "ΔOi判定、対象データ無し。"
            End Select
        Else
            ''狙い、規格無し以外の場合
            If (AOI.GuaranteeAoi.cObj <> ObjCode13) And (AOI.GuaranteeAoi.cObj <> ObjCode15) Then
                ''対象データ無し
                ''エラー情報構造体に情報を代入。
'''''                WFCJudgDialog.WFCErrorMessage "ΔOi判定、対象データ無し。"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, AOI_JUDG, AOI.GuaranteeAoi.cObj)
            End If
        End If
    
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
        ''AN温度判定
        liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(AOI.Antnp), CStr(AOI.AoiAntnp))
        If liRet = -1 Then
            FuncAns = SetErrInfo(ErrInfo, EZJ00, AOI_JUDG, AOI.GuaranteeAoi.cObj)
        ElseIf liRet = 0 Then
            AOI.JudgAntnp = JUDG_NG
        Else
            AOI.JudgAntnp = JUDG_OK
        End If
                
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Else
        AOI.JudgAoi = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -s-
        AOI.JudgAntnp = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -e-
    End If
    
    WfAOiJudg = FuncAns
End Function

'概要      :WFセンターOSF判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Osf           ,I  ,W_OSF            ,WFセンターOSF判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WfOSFJudg(OSF As W_OSF, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFOSF_JUDG = 4                 ''判定識別フラグ(OSF)
    Dim FuncAns As FUNCTION_RETURN
    Dim JData(4) As Double
    Dim c0 As Integer
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
    Dim liRet As Integer
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '' 2006/09/19 SMP)kondoh Add -s-
    Dim JudgOsfTmp As Boolean
    JudgOsfTmp = JUDG_OK
    '' 2006/09/19 SMP)kondoh Add -e-
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    OSF.JudgOsf = JUDG_NG
    
    OSF.JudgDataAve = JudgAve(OSF.OSF())
    OSF.JudgDataMax = JudgMax(OSF.OSF())
    OSF.JudgDataMin = JudgMin(OSF.OSF())
    
    If OSF.GuaranteeOsf.cJudg = JudgCodeW01 Then ''OSF判定有り
        
''''''        WFCJudgDialog.WFCErrorMessage " "
''''''        WFCJudgDialog.WFCErrorMessage "OSF判定 ***************"
''''''        WFCJudgDialog.WFCErrorMessage "測定位置_点 = " & OSF.GuaranteeOsf.cCount
''''''        WFCJudgDialog.WFCErrorMessage "測定位置_方 = " & OSF.GuaranteeOsf.cMeth
''''''        WFCJudgDialog.WFCErrorMessage "測定位置_位 = " & OSF.GuaranteeOsf.cPos
''''''        WFCJudgDialog.WFCErrorMessage "保証方法_対 = " & OSF.GuaranteeOsf.cObj
''''''        WFCJudgDialog.WFCErrorMessage "保証方法_処 = " & OSF.GuaranteeOsf.cJudg
        
        ''OSF判定
        FuncAns = GetWfJudgData(WFOSF_JUDG, OSF.GuaranteeOsf, OSF.OSF(), JData())
'        If (InStr(ObjCodeGrp03, OSF.GuaranteeOsf.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
        If (InStr(ObjCodeGrp03 & ObjCode08 & ObjCode09, OSF.GuaranteeOsf.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
            Select Case OSF.GuaranteeOsf.cObj
            Case ObjCode05, ObjCode08  ''全点の平均値、全点の最小値
                OSF.JudgOsf = RangeDecision_nl(JData(0), 0, OSF.SpecOsfAveMax)
            Case ObjCode06  ''全点の最大値
                OSF.JudgOsf = RangeDecision_nl(JData(0), 0, OSF.SpecOsfMax)
            Case ObjCode07 ''全点の平均値と最大値
                '' 2006/09/19 SMP)kondoh Cng -s-
''                If RangeDecision_nl(JData(1), 0, OSF.SpecOsfAveMax) Then
                If RangeDecision_nl(JData(0), 0, OSF.SpecOsfAveMax) Then
                '' 2006/09/19 SMP)kondoh Cng -e-
                    OSF.JudgOsf = RangeDecision_nl(JData(1), 0, OSF.SpecOsfMax)
                Else
                    OSF.JudgOsf = JUDG_NG
                End If
            Case ObjCode09 ''内周部2点、外周部2点(5点測定で1,2,4,5)
                For c0 = 0 To 3
                '' 2006/09/19 SMP)kondoh Cng -s-
''                    If RangeDecision_nl(JData(c0), 0, OSF.SpecOsfAveMax) Then
''                        OSF.JudgOsf = JUDG_NG
                    If RangeDecision_nl(JData(c0), 0, OSF.SpecOsfAveMax) = False Then
                        JudgOsfTmp = JUDG_NG
                '' 2006/09/19 SMP)kondoh Cng -e-
                    End If
                Next
                '' 2006/09/19 SMP)kondoh Add -s-
                If JudgOsfTmp = JUDG_OK Then OSF.JudgOsf = JUDG_OK
                '' 2006/09/19 SMP)kondoh Add -e-
            End Select
        Else
            ''狙い、規格無し以外の場合
            If (OSF.GuaranteeOsf.cObj <> ObjCode13) And (OSF.GuaranteeOsf.cObj <> ObjCode15) Then
                ''対象データ無し
                ''エラー情報構造体に情報を代入。
'''''                WFCJudgDialog.WFCErrorMessage "OSF判定、対象データ無し。"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, OSF_JUDG, OSF.GuaranteeOsf.cObj)
            End If
        End If
        With OSF
        If .JudgOsf Then    'パターンの判定　2003/05/17 ooba
            If .JudgDataPTK = "1" Or .JudgDataPTK = "2" Or .JudgDataPTK = "3" _
            Or .JudgDataPTK = "4" Or .JudgDataPTK = " " Then
                If InStr("RD ", .OSFp(0)) > 0 And InStr("RD ", .OSFp(1)) > 0 _
                And InStr("RD ", .OSFp(2)) > 0 Then
                    .JudgOsf = JudgPattern(.JudgDataPTK, .OSFp())
                Else
                    FuncAns = FUNCTION_RETURN_FAILURE
                End If
            Else
                FuncAns = FUNCTION_RETURN_FAILURE
            End If
        End If
        End With
    
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
        ''AN温度判定
        liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(OSF.Antnp), CStr(OSF.OsfAntnp))
        If liRet = -1 Then
            FuncAns = FUNCTION_RETURN_FAILURE
        ElseIf liRet = 0 Then
            OSF.JudgAntnp = JUDG_NG
        Else
            OSF.JudgAntnp = JUDG_OK
        End If
        
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------

    Else
        OSF.JudgOsf = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -s-
        OSF.JudgAntnp = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -e-
'        If InStr(JudgCodeW02, Osf.GuaranteeOsf.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, OSF_JUDG, Osf.GuaranteeOsf.cJudg)
'        End If
    End If
    
    WfOSFJudg = FuncAns
End Function

'概要      :WFセンターBMD判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Bmd           ,I  ,W_BMD            ,WFセンターBMD判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :Bno           ,I  ,Integer          ,BMDno(1:BMD1,2:BMD2,3:BMD3)(ｴﾋﾟ用)　09/05/07 ooba
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WfBMDJudg(BMD As W_BMD, ErrInfo As ERROR_INFOMATION, _
                                                Optional Bno As Integer = 0) As FUNCTION_RETURN
'WFBMD_JUDG = 5                 ''判定識別フラグ(BMD)
    Dim FuncAns As FUNCTION_RETURN
    Dim TempBmd(3) As Double
    Dim JData(3) As Double
    Dim c0 As Integer
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
    Dim liRet As Integer
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    BMD.JudgBmd = JUDG_NG
    
    BMD.JudgDataAve = JudgAve(BMD.BMD())
    BMD.JudgDataMax = JudgMax(BMD.BMD())
    BMD.JudgDataMin = JudgMin(BMD.BMD())
    If BMD.SpecBmdMCL = "P " Then                '面内分布の計算　2003/05/20 ooba
        BMD.JudgDataMBP = JudgBmdMBP(BMD.BMD())
    Else
        BMD.JudgDataMBP = 0                      '面内分布が"P"以外の時は計算結果を0とする　2003/06/06 ooba
    End If
    
    If BMD.GuaranteeBmd.cJudg = JudgCodeW01 Then ''BMD判定有り
        
''''''        WFCJudgDialog.WFCErrorMessage " "
''''''        WFCJudgDialog.WFCErrorMessage "BMD判定 ***************"
''''''        WFCJudgDialog.WFCErrorMessage "測定位置_点 = " & BMD.GuaranteeBmd.cCount
''''''        WFCJudgDialog.WFCErrorMessage "測定位置_方 = " & BMD.GuaranteeBmd.cMeth
''''''        WFCJudgDialog.WFCErrorMessage "測定位置_位 = " & BMD.GuaranteeBmd.cPos
''''''        WFCJudgDialog.WFCErrorMessage "保証方法_対 = " & BMD.GuaranteeBmd.cObj
''''''        WFCJudgDialog.WFCErrorMessage "保証方法_処 = " & BMD.GuaranteeBmd.cJudg
        
        ''BMD判定
'--- 2006/08/15 Del エピ先行評価追加対応 SMP)kondoh -s-
'        FuncAns = GetWfJudgData(WFBMD_JUDG, BMD.GuaranteeBmd, BMD.BMD(), JData())
'--- 2006/08/15 Del エピ先行評価追加対応 SMP)kondoh -e-
        If (InStr(ObjCodeGrp02, BMD.GuaranteeBmd.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
            FuncAns = GetWfJudgData(WFBMD_JUDG, BMD.GuaranteeBmd, BMD.BMD(), JData())
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
            Select Case BMD.GuaranteeBmd.cObj
            ''全点の平均値、全点の最大値、全点の最小値、MAX(2,4点目)、MAX(2,3,4点目)
            Case ObjCode05, ObjCode06, ObjCode08, ObjCode10, ObjCode11
                BMD.JudgBmd = RangeDecision_nl(JData(0), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
            Case ObjCode07 ''全点の平均値と最大値
                If RangeDecision_nl(JData(0), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax) Then
                    BMD.JudgBmd = RangeDecision_nl(JData(1), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
                Else
                    BMD.JudgBmd = JUDG_NG
                End If
            '----TEST2004/10
            Case ObjCode16 ''全点の最大値と最小値
                If RangeDecision_nl(JData(0), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax) Then
                    BMD.JudgBmd = RangeDecision_nl(JData(1), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
                Else
                    BMD.JudgBmd = JUDG_NG
                End If
'' 2008/10/20 BMD評価,外周1点保証機能追加 ADD By Systech Start
            Case ObjCode18  ''AVE+外周1点
'                If RangeDecision_nl(JData(0), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax) Then
                '外周1点判定方法変更　09/05/07 ooba
                If Bno = 3 Then
                    If RangeDecision_nl(JData(0), BMD.SpecBmdGsAveMin, BMD.SpecBmdGsAveMax) Then
                        BMD.JudgBmd = RangeDecision_nl(JData(1), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
                    Else
                        BMD.JudgBmd = JUDG_NG
                    End If
                Else
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, BMD_JUDG, BMD.GuaranteeBmd.cObj)
                End If
'' 2008/10/20 BMD評価,外周1点保証機能追加 ADD By Systech End
            End Select
        Else
            ''狙い、規格無し以外の場合
            If (BMD.GuaranteeBmd.cObj <> ObjCode13) And (BMD.GuaranteeBmd.cObj <> ObjCode15) Then
                ''対象データ無し
                ''エラー情報構造体に情報を代入。
'''''                WFCJudgDialog.WFCErrorMessage "OSF判定、対象データ無し。"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, BMD_JUDG, BMD.GuaranteeBmd.cObj)
            End If
        End If
        '面内分布での判定　2003/05/20 ooba
        If BMD.SpecBmdMCL = "P " Then      '仕様がPの時のみ判定を行い、それ以外は判定を行わずOKとする
'            If BMD.SpecBmdMBP = -1 Then
'                FuncAns = FUNCTION_RETURN_FAILURE
            If BMD.SpecBmdMBP <> 0 Or BMD.SpecBmdMBP = -1 Then      '仕様の面内分布が0,-1(NULL)の時は判定を行わずOKとする
                If BMD.JudgDataMBP = -1 Then
                    FuncAns = FUNCTION_RETURN_FAILURE
                Else
                    If BMD.JudgBmd Then
                        BMD.JudgBmd = RangeDecision_nl(BMD.JudgDataMBP, 0, BMD.SpecBmdMBP)
                    End If
                End If
            End If
        End If
    
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
        ''AN温度判定
        liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(BMD.Antnp), CStr(BMD.BmdAntnp))
        If liRet = -1 Then
            FuncAns = FUNCTION_RETURN_FAILURE
        ElseIf liRet = 0 Then
            BMD.JudgAntnp = JUDG_NG
        Else
            BMD.JudgAntnp = JUDG_OK
        End If
        
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------

    Else
        BMD.JudgBmd = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -s-
        BMD.JudgAntnp = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -e-
'        If InStr(JudgCodeW02, Bmd.GuaranteeBmd.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, BMD_JUDG, Bmd.GuaranteeBmd.cJudg)
'        End If
    End If
    
    WfBMDJudg = FuncAns
End Function

'概要      :WFセンターDZ判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Dz            ,I  ,W_DZ             ,WFセンターDZ判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WfDZJudg(DZ As W_DZ, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFDZ_JUDG = 6                  ''判定識別フラグ(DZ)
    Dim FuncAns As FUNCTION_RETURN
    Dim JData(3) As Double
    Dim c0 As Integer
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    DZ.JudgDz = JUDG_NG
    
    DZ.JudgDataAve = JudgAve(DZ.DZ())
    DZ.JudgDataMax = JudgMax(DZ.DZ())
    DZ.JudgDataMin = JudgMin(DZ.DZ())
    
    If DZ.GuaranteeDz.cJudg = JudgCodeW01 Then ''DZ判定有り
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "DZ判定 ****************"
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_点 = " & Dz.GuaranteeDz.cCount
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_方 = " & Dz.GuaranteeDz.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_位 = " & Dz.GuaranteeDz.cPos
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_対 = " & Dz.GuaranteeDz.cObj
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_処 = " & Dz.GuaranteeDz.cJudg
        
        ''DZ判定
        FuncAns = GetWfJudgData(WFDZ_JUDG, DZ.GuaranteeDz, DZ.DZ(), JData())
        If (InStr(ObjCodeGrp06, DZ.GuaranteeDz.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
            Select Case DZ.GuaranteeDz.cObj
            ''全点の平均値、全点の最大値、全点の最小値、MAX(2,4点目)、MAX(2,3,4点目)
            Case ObjCode05, ObjCode06, ObjCode08, ObjCode10, ObjCode11
                DZ.JudgDz = RangeDecision_nl(JData(0), DZ.SpecDzMin, DZ.SpecDzMax)
            Case ObjCode07 ''全点の平均値と最大値
                If RangeDecision_nl(JData(0), DZ.SpecDzMin, DZ.SpecDzMax) Then
                    DZ.JudgDz = RangeDecision_nl(JData(1), DZ.SpecDzMin, DZ.SpecDzMax)
                Else
                    DZ.JudgDz = JUDG_NG
                End If
            
            '佐賀システムと共通関数合わせで修正　hama 2004/11/30 start
            Case ObjCode03
                DZ.JudgDz = JUDG_OK
                  'For c0 = 0 To 3    '2004/12/21
                  For c0 = 0 To CInt(DZ.GuaranteeDz.cCount) - 1
                    If JData(c0) <> -1 Then
                        If RangeDecision_nl(JData(c0), DZ.SpecDzMin, DZ.SpecDzMax) = JUDG_NG Then
                            DZ.JudgDz = JUDG_NG
                        End If
                    Else
                            DZ.JudgDz = JUDG_NG
                    End If
                Next
            Case ObjCode16
               If RangeDecision_nl(JData(0), DZ.SpecDzMin, DZ.SpecDzMax) Then
                   DZ.JudgDz = RangeDecision_nl(JData(1), DZ.SpecDzMin, DZ.SpecDzMax)
               Else
                    DZ.JudgDz = JUDG_NG
               End If
            End Select
           '佐賀システムと共通関数合わせで修正　hama 2004/11/30 end
        Else
            ''狙い、規格無し以外の場合
            If (DZ.GuaranteeDz.cObj <> ObjCode13) And (DZ.GuaranteeDz.cObj <> ObjCode15) Then
                ''対象データ無し
                ''エラー情報構造体に情報を代入。
'''''                WFCJudgDialog.WFCErrorMessage "DZ判定、対象データ無し。"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, DZ_JUDG, DZ.GuaranteeDz.cObj)
            End If
        End If
    Else
        DZ.JudgDz = JUDG_OK
'        If InStr(JudgCodeW02, Dz.GuaranteeDz.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, DZ_JUDG, Dz.GuaranteeDz.cJudg)
'        End If
    End If
    
    WfDZJudg = FuncAns
End Function

'概要      :WFセンターDSOD判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Dsod          ,I  ,W_DSOD           ,WFセンターDSOD判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WfDSODJudg(Dsod As W_DSOD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFDSOD_JUDG = 7                ''判定識別フラグ(DSOD)
    Dim FuncAns As FUNCTION_RETURN
    Dim liRet As Integer        '06/12/22 ooba
    Dim sResult As String       '06/12/22 ooba
    Dim RET As Integer          '06/12/22 ooba
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Dsod.JudgDsod = JUDG_NG
    Dsod.JudgAntnp = JUDG_OK        '06/12/22 ooba
    If Dsod.GuaranteeDsod.cJudg = JudgCodeW01 Then ''DSOD判定有り
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "DSOD判定 **************"
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_点 = " & Dsod.GuaranteeDsod.cCount
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_方 = " & Dsod.GuaranteeDsod.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_位 = " & Dsod.GuaranteeDsod.cPos
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_対 = " & Dsod.GuaranteeDsod.cObj
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_処 = " & Dsod.GuaranteeDsod.cJudg
        
        If Dsod.GuaranteeDsod.cObj = ObjCodeGrp04 Then
            ''DSOD判定
            Dsod.JudgDsod = RangeDecision_nl(Dsod.Dsod, Dsod.SpecDsodMin, Dsod.SpecDsodMax)
        Else
            ''狙い、規格無し以外の場合
            If (Dsod.GuaranteeDsod.cObj <> ObjCode13) And (Dsod.GuaranteeDsod.cObj <> ObjCode15) Then
                ''対象データ無し
                ''エラー情報構造体に情報を代入。
'''''                WFCJudgDialog.WFCErrorMessage "DSOD判定、対象データ無し。"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, DSOD_JUDG, Dsod.GuaranteeDsod.cObj)
            End If
        End If
        'DSODﾊﾟﾀｰﾝ判定追加　04/07/28 ooba START ================================>
        If Dsod.JudgDsod = JUDG_OK Then
            Dsod.JudgDsod = JudgDsodPattern(Dsod.JudgDataPTK, Dsod.Dsodp())
        End If
        'DSODﾊﾟﾀｰﾝ判定追加　04/07/28 ooba END ==================================>
        
        'GD/DSOD熱処理条件追加　06/12/22 ooba START =========================================>
        RET = 0
        sResult = ""
        RET = funCodeDBGet("SB", "15", "DS", 0, " ", sResult)
        If RET = 0 And Mid(sResult, 16, 1) = "2" Then
            liRet = funCodeDBGetMatrixReturn("SB", "AD", CStr(Dsod.Antnp), CStr(Dsod.DsodAntnp))
            If liRet = -1 Then
                FuncAns = FUNCTION_RETURN_FAILURE
            ElseIf liRet = 0 Then
                Dsod.JudgAntnp = JUDG_NG
            End If
        End If
        'GD/DSOD熱処理条件追加　06/12/22 ooba END ===========================================>
    Else
        Dsod.JudgDsod = JUDG_OK
'        If InStr(JudgCodeW02, Dsod.GuaranteeDsod.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, EZJ00, DSOD_JUDG, Dsod.GuaranteeDsod.cJudg)
'        End If
    End If
    
    WfDSODJudg = FuncAns
End Function

'概要      :WFセンターSPV判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Spv           ,I  ,W_SPV            ,WFセンターSPV判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WfSPVJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFSPV_JUDG = 8                 ''判定識別フラグ(SPV)
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim j1 As Boolean
    Dim j2 As Boolean
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
    j1 = JUDG_NG
    j2 = JUDG_NG
    
    If Spv.GuaranteeSpv.cJudg = JudgCodeW01 Then ''SPV判定有り
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "SPV判定 ***************"
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_点 = " & Spv.GuaranteeSpv.cCount
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_方 = " & Spv.GuaranteeSpv.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_位 = " & Spv.GuaranteeSpv.cPos
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_対 = " & Spv.GuaranteeSpv.cObj
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_処 = " & Spv.GuaranteeSpv.cJudg
        
        ''SPV判定
        '-----TEST2004/10
        'If Spv.GuaranteeSpv.cObj = ObjCode03 Then
        Select Case Spv.GuaranteeSpv.cObj
            Case ObjCode03  '全測定点=3
                ''拡散長AVEが規格値範囲内ならOK
                'j1 = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax)
                ''MAX,AVE,MINが全て規格値範囲内ならOKに変更
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    If RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                        j1 = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
                    Else
                        j1 = JUDG_NG
                    End If
                Else
                    j1 = JUDG_NG
                End If
            Case ObjCode05 'ＡＶＥ=A
                ''AVEが規格値範囲内ならOK
                j1 = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax)
            Case ObjCode06 'ＭＡＸ=B
                ''MAXが規格値範囲内ならOK
                j1 = RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax)
            Case ObjCode07 'ＡＶＥ+ＭＡＸ=C
                ''MAX,AVEが規格値範囲内ならOK
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    j1 = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax)
                Else
                    j1 = JUDG_NG
                End If
            Case ObjCode08 'ＭＩＮ=D
                ''MINが規格値範囲内ならOK
                j1 = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
            Case ObjCode16 'ＭＩＮ+ＭＡＸ=K
                ''MAX,MINが規格値範囲内ならOK
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    j1 = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
                Else
                    j1 = JUDG_NG
                End If
            Case Else
                ''狙い、規格無し以外の場合
                If (Spv.GuaranteeSpv.cObj <> ObjCode13) And (Spv.GuaranteeSpv.cObj <> ObjCode15) Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
    '''''                WFCJudgDialog.WFCErrorMessage "SPV判定、対象データ無し。"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, SPV_JUDG, Spv.GuaranteeSpv.cObj)
                    GoTo EXIT_FUNC
                End If
                j1 = JUDG_OK
        End Select

    Else
        j1 = JUDG_OK
    End If
    
    If Spv.GuaranteeSpvFe.cJudg = JudgCodeW01 Then ''SPVFE判定有り
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "SPV判定 ***************"
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_点 = " & Spv.GuaranteeSpvFe.cCount
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_方 = " & Spv.GuaranteeSpvFe.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "測定位置_位 = " & Spv.GuaranteeSpvFe.cPos
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_対 = " & Spv.GuaranteeSpvFe.cObj
'''''        WFCJudgDialog.WFCErrorMessage "保証方法_処 = " & Spv.GuaranteeSpvFe.cJudg
        
        ''SPVFE判定
        '----TEST2004/10
        'If Spv.GuaranteeSpvFe.cObj = ObjCode03 Then
        Select Case Spv.GuaranteeSpvFe.cObj
        Case ObjCode03, ObjCode06 '全測定点、ＭＡＸ
            j2 = RangeDecision_nl(Spv.Spv(3), 0, Spv.SpecSpvFeMax)
        Case ObjCode05 'ＡＶＥ
            j2 = RangeDecision_nl(Spv.Spv(4), 0, Spv.SpecSpvAvMax)
        Case ObjCode07 'ＡＶＥ+ＭＩＮ
            If RangeDecision_nl(Spv.Spv(4), 0, Spv.SpecSpvAvMax) Then
                j2 = RangeDecision_nl(Spv.Spv(3), 0, Spv.SpecSpvFeMax)
            Else
                j2 = JUDG_NG
            End If
        Case Else  '
            ''狙い、規格無し以外の場合
            If (Spv.GuaranteeSpvFe.cObj <> ObjCode13) And (Spv.GuaranteeSpvFe.cObj <> ObjCode15) Then
                ''対象データ無し
                ''エラー情報構造体に情報を代入。
'''''                WFCJudgDialog.WFCErrorMessage "SPV判定、対象データ無し。"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, SPV_JUDG, Spv.GuaranteeSpvFe.cObj)
                GoTo EXIT_FUNC
            End If
            j2 = JUDG_OK
        End Select

    Else
        j2 = JUDG_OK
    End If
    
    Spv.JudgSpv = (j1 And j2)

EXIT_FUNC:
    
    WfSPVJudg = FuncAns
End Function

'概要      :WFセンターGD判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO    ,型                  ,説明
'          :GD            ,I     ,W_GD                ,WFセンターGD判定構造体
'          :sGDhsflg      ,I     ,String              ,保証ﾌﾗｸﾞ(1：WF保証)　06/12/22 ooba
'          :ErrInfo       ,O     ,ERROR_INFOMATION    ,エラー情報構造体
'          :戻り値        ,O     ,FUNCTION_RETURN     ,
'説明      :
'履歴      :05/01/31 ooba
Public Function WfGdJudg(GD As W_GD, sGDhsflg As String, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN

    Dim FuncAns As FUNCTION_RETURN
    Dim liRet As Integer        '06/12/22 ooba
    Dim sResult As String       '06/12/22 ooba
    Dim RET As Integer          '06/12/22 ooba
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS

    ''Den検査有無判断
    GD.JudgDen = JUDG_OK
'    If GD.JudgFlagDen = "1" Then
        If GD.GuaranteeDen.cJudg = JudgCodeW01 Then ''Den判定あり
            GD.JudgDen = RangeDecision_nl(GD.Den, GD.SpecDenMin, GD.SpecDenMax)
        Else
            GD.JudgDen = JUDG_OK
        End If
'    End If

    ''L/DL検査有無判断
    GD.JudgLdl = JUDG_OK
'    If GD.JudgFlagLdl = "1" Then
        If GD.GuaranteeLdl.cJudg = JudgCodeW01 Then ''L/DL判定あり
            GD.JudgLdl = RangeDecision_nl(GD.Ldl, GD.SpecLdlMin, GD.SpecLdlMax)
        Else
            GD.JudgLdl = JUDG_OK
        End If
'    End If

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    GD.JudgLdlPtn = JUDG_OK
'    If WFGD.JudgFlagLdl = "1" Then
        If GD.GuaranteeLdl.cJudg = JudgCodeW01 Then ''L/DL判定あり
            If GD.GDPTK = "1" Then
                ' "0"連続数(MIN)　≧　品__L/DL連続0下限(SX/WF)
                If GD.ZeroLdlMin = -1 Then
                    GD.JudgLdlPtn = JUDG_OK
                Else
                    If GD.LdlMin >= GD.ZeroLdlMin Then
                        GD.JudgLdlPtn = JUDG_OK
                    Else
                        GD.JudgLdlPtn = JUDG_NG
                    End If
                End If
            ElseIf GD.GDPTK = "2" Then
                ' "0"連続数(MAX)　≦　品__L/DL連続0上限(SX/WF)
                If GD.ZeroLdlMax = -1 Then
                    GD.JudgLdlPtn = JUDG_OK
                Else
                    If GD.LdlMax <= GD.ZeroLdlMax Then
                        GD.JudgLdlPtn = JUDG_OK
                    Else
                        GD.JudgLdlPtn = JUDG_NG
                    End If
                End If
            Else
                ' 判定無し
                
            End If
        End If
'    End If
    GD.JudgLdl = GD.JudgLdl And GD.JudgLdlPtn
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    
    ''DVD2検査有無判断
    GD.JudgDvd2 = JUDG_OK
'    If GD.JudgFlagDvd2 = "1" Then
        If GD.GuaranteeDvd2.cJudg = JudgCodeW01 Then ''Dvd2判定あり
            GD.JudgDvd2 = RangeDecision_nl(GD.Dvd2, GD.SpecDvd2Min, GD.SpecDvd2Max)
        Else
            GD.JudgDvd2 = JUDG_OK
        End If
'    End If
    
    'GD/DSOD熱処理条件追加　06/12/22 ooba START =========================================>
    GD.JudgAntnp = JUDG_OK
    If sGDhsflg = "1" And _
       (GD.GuaranteeDen.cJudg = JudgCodeW01 Or _
        GD.GuaranteeLdl.cJudg = JudgCodeW01 Or _
        GD.GuaranteeDvd2.cJudg = JudgCodeW01) Then
        RET = 0
        sResult = ""
        'DEN-AN温度ﾁｪｯｸ
        RET = funCodeDBGet("SB", "15", "DEN", 0, " ", sResult)
        If RET = 0 And Mid(sResult, 16, 1) = "2" Then
            'LDL-AN温度ﾁｪｯｸ
            RET = funCodeDBGet("SB", "15", "LDL", 0, " ", sResult)
            If RET = 0 And Mid(sResult, 16, 1) = "2" Then
                'DVD-AN温度ﾁｪｯｸ
                RET = funCodeDBGet("SB", "15", "DVD", 0, " ", sResult)
                If RET = 0 And Mid(sResult, 16, 1) = "2" Then
                    liRet = funCodeDBGetMatrixReturn("SB", "AG", CStr(GD.Antnp), CStr(GD.GdAntnp))
                    If liRet = -1 Then
                        FuncAns = FUNCTION_RETURN_FAILURE
                    ElseIf liRet = 0 Then
                        GD.JudgAntnp = JUDG_NG
                    End If
                End If
            End If
        End If
    End If
    'GD/DSOD熱処理条件追加　06/12/22 ooba END ===========================================>
        
    WfGdJudg = FuncAns

End Function
'概要      :WFセンターSIRD(SD)判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Dz            ,I  ,W_DZ             ,WFセンターSIRD判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2010/01/07 Y.Hitomi
Public Function WfSDJudg(SD As W_SD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim JData(3) As Double
    Dim c0 As Integer
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
        
    If SD.GuaranteeSd.cJudg = JudgCodeW01 Then  ''SD判定有り
        ''SD判定
        If SD.SdMeasData <= SD.SpecSdMax Then
            SD.JudgSD = JUDG_OK
        Else
            SD.JudgSD = JUDG_NG
        End If
    Else
        'Cng Start 2010/10/05 Y.Hitomi
        SD.JudgSD = JUDG_OK
        'FuncAns = FUNCTION_RETURN_FAILURE
        'Cng End   2010/10/05 Y.Hitomi
    End If
    
    WfSDJudg = FuncAns
End Function

'概要      :対象コードに従って判定対象データを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Flag          ,I  ,GUARANTEE ,対象コード
'          :d()           ,I  ,double    ,測定値
'          :d1()          ,O  ,double    ,判定対象データ
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :Flag.cObjの値,
'          :1       ,d1(0)=中心測定値
'          :2       ,d1(0)=中央測定値
'          :3       ,d1()=全測定点
'          :4       ,d1(0)=R/2
'          :A       ,d1(0)=平均値
'          :B       ,d1(0)=最大値
'          :C       ,d1(0)=平均値,d1(1)=最大値
'          :D       ,d1(0)=最小値
'          :E       ,d1(0～3)=内周部2点、外周部2点(5点測定で1,2,4,5)
'          :F       ,d1(0)=2,4点目の内大きい値
'          :G       ,d1(0)=2,3,4点目の内大きい値
'          :K       ,d1(0)=最大値,d1(1)=最小値　TEST2004/10
'履歴      :2001/06/06 佐野 信哉 作成
Public Function GetWfJudgData(JudgFlag As Integer, flag As Guarantee, d() As Double, d1() As Double) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim COUNT As Integer
    Dim High As Integer
    
    '' 配列の上限を取得します。
    High = UBound(d)
    
    FuncAns = FUNCTION_RETURN_SUCCESS '' 正常
    Select Case flag.cObj
    Case ObjCode01 ''中心測定値
        Select Case JudgFlag
        Case WFOI_JUDG  ''判定識別フラグ(Oi)
            d1(0) = d(0)
        Case WFRES_JUDG ''判定識別フラグ(RES)
            d1(0) = d(4)
        Case WFDOI_JUDG, WFAOI_JUDG ''判定識別フラグ(ΔOi)、判定識別フラグ(AOi) 追加 03/12/09 ooba
        '' 取得データ変更　2003/10/15 ooba
'            d1(0) = d(5)
            d1(0) = d(2)
        Case WFBMD_JUDG ''判定識別フラグ(BMD)
            d1(0) = d(3)
'        Case WFDZ_JUDG  ''判定識別フラグ(DZ)
'        Case WFDSOD_JUDG ''判定識別フラグ(DSOD)
        Case WFSPV_JUDG ''判定識別フラグ(SPV)
        Case WFOSF_JUDG ''判定識別フラグ(OSF)
        Case Else
'''''            WFCJudgDialog.WFCErrorMessage "対象データ無し。"
            FuncAns = FUNCTION_RETURN_FAILURE '' 異常
        End Select
    Case ObjCode02 ''測定値の中央値
        d1(0) = JudgCenter(d())
    Case ObjCode03 ''全測定点
        DataCopy d(), d1()
    Case ObjCode04 ''R/2
        If InStr(PosCodeGrp01, flag.cPos) <> 0 Then
            Select Case JudgFlag
            Case WFRES_JUDG ''判定識別フラグ(RES)
                d1(0) = d(3)
            Case WFOI_JUDG ''判定識別フラグ(Oi)
                d1(0) = d(2)
            Case WFDOI_JUDG, WFBMD_JUDG, WFDZ_JUDG, WFAOI_JUDG ''判定識別フラグ(ΔOi)、判定識別フラグ(BMD)、判定識別フラグ(DZ)、判定識別フラグ(AOi) 追加 03/12/09 ooba
                d1(0) = d(1)
'            Case WFDSOD_JUDG ''判定識別フラグ(DSOD)
'            Case WFSPV_JUDG ''判定識別フラグ(SPV)
'            Case WFOSF_JUDG ''判定識別フラグ(OSF)
            Case Else
'''''                WFCJudgDialog.WFCErrorMessage "対象データ無し。"
                FuncAns = FUNCTION_RETURN_FAILURE '' 異常
            End Select
        Else
'''''            WFCJudgDialog.WFCErrorMessage "対象データ無し。"
            FuncAns = FUNCTION_RETURN_FAILURE '' 異常
        End If
    Case ObjCode05 ''全点の平均値
        d1(0) = JudgAve(d())
    Case ObjCode06 ''全点の最大値
        d1(0) = JudgMax(d())
    Case ObjCode07 ''全点の平均値と最大値
        d1(0) = JudgAve(d())
        d1(1) = JudgMax(d())
    Case ObjCode08 ''全点の最小値
        d1(0) = JudgMin(d())
    Case ObjCode09 ''内周部2点、外周部2点(5点測定で1,2,4,5)
        DataCopy d(), d1()
        COUNT = 0
        For c0 = High To 0 Step -1
            '' 2006/09/19 SMP)kondoh Cng -s-
            If d(c0) <> -1 Then
                d1(3 - COUNT) = d(c0)
''            If d1(c0) <> -1 Then
''                d1(3 - COUNT) = d1(c0)
            '' 2006/09/19 SMP)kondoh Cng -e-
                COUNT = COUNT + 1
            End If
            If COUNT = 2 Then Exit For
        Next
    Case ObjCode10 ''MAX(2,4点目)
        If (d(1) <> -1) And (d(3) <> -1) Then
            If d(1) >= d(3) Then
                d1(0) = d(1)
            Else
                d1(0) = d(3)
            End If
        Else
'''''            WFCJudgDialog.WFCErrorMessage "対象データ無し。"
            FuncAns = FUNCTION_RETURN_FAILURE '' 異常
        End If
    Case ObjCode11 ''MAX(2,3,4点目)
        If (d(1) <> -1) And (d(2) <> -1) And (d(3) <> -1) Then
            If d(1) >= d(2) Then
                If d(1) >= d(3) Then
                    d1(0) = d(1)
                Else
                    d1(0) = d(3)
                End If
            Else
                If d(2) >= d(3) Then
                    d1(0) = d(2)
                Else
                    d1(0) = d(3)
                End If
            End If
        Else
            ''あり得ないエラー
'''''            WFCJudgDialog.WFCErrorMessage "対象データ無し。"
            FuncAns = FUNCTION_RETURN_FAILURE '' 異常
        End If
''    Case ObjCode12 ''個数保証
''    Case ObjCode13 ''狙い
''    Case ObjCode14 ''形状測定(平坦度、反返り、WARP)
''    Case ObjCode15 ''規格なし
    '----TEST2004/10
    Case ObjCode16 ''全点の最大値と最小値
        'Select Case JudgFlag 04/12/16削除
            'Case WFBMD_JUDG
                d1(0) = JudgMax(d())
                d1(1) = JudgMin(d())
        'End Select
        
'' 2008/10/20 BMD評価,外周1点保証機能追加 ADD By Systech Start
    Case ObjCode18 ''AVE+外周1点
        d1(0) = d(0)            '外周１点
        d1(1) = JudgAve(d())    '全点平均値
'' 2008/10/20 BMD評価,外周1点保証機能追加 ADD By Systech End
        
    Case Else
'''''        WFCJudgDialog.WFCErrorMessage "対象データ無し。"
        FuncAns = FUNCTION_RETURN_FAILURE '' 異常
    End Select
    
    GetWfJudgData = FuncAns
End Function

'概要      :WFC判定対象MINデータを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WFCMin(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim Min As Double
    
    Min = -9999
    
    Select Case G.cPos
    Case "1"                                           '2003/05/15 追加　osawa 依頼No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(0) < d(3), d(0), d(3))
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
            Min = IIf(d(6) < Min, d(6), Min)
            Min = IIf(d(9) < Min, d(9), Min)
        End Select                                     '依頼No.030130　追加ここまで
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(0) < d(4), d(0), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(6), d(0), d(6))
            Min = IIf(d(9) < Min, d(9), Min)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(0) < d(4), d(0), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(6), d(0), d(6))
            Min = IIf(d(9) < Min, d(9), Min)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(0) < d(4), d(0), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(6), d(0), d(6))
            Min = IIf(d(9) < Min, d(9), Min)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(1) < d(4), d(1), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(6), d(0), d(6))
            Min = IIf(d(9) < Min, d(9), Min)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(1) < d(4), d(1), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(5), d(0), d(5))
            Min = IIf(d(8) < Min, d(8), Min)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(1) < d(4), d(1), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(5), d(0), d(5))
            Min = IIf(d(8) < Min, d(8), Min)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(1) < d(4), d(1), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(5), d(0), d(5))
            Min = IIf(d(8) < Min, d(8), Min)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(2) < d(4), d(2), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(4), d(0), d(4))
            Min = IIf(d(7) < Min, d(7), Min)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(3), d(0), d(3))
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(3) < d(4), d(3), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(1), d(0), d(1))
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(1) < d(3), d(1), d(3))
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
            Min = IIf(d(6) < Min, d(6), Min)
            Min = IIf(d(9) < Min, d(9), Min)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(1) < d(3), d(1), d(3))
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
            Min = IIf(d(5) < Min, d(5), Min)
            Min = IIf(d(8) < Min, d(8), Min)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(2) < d(3), d(2), d(3))
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
            Min = IIf(d(4) < Min, d(4), Min)
            Min = IIf(d(7) < Min, d(7), Min)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(2) < d(4), d(2), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(4), d(0), d(4))
            Min = IIf(d(7) < Min, d(7), Min)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(2) < d(3), d(2), d(3))
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
            Min = IIf(d(4) < Min, d(4), Min)
            Min = IIf(d(7) < Min, d(7), Min)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(1), d(0), d(1))
            Min = IIf(d(3) < Min, d(3), Min)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            Min = IIf(d(0) < d(1), d(0), d(1))
            Min = IIf(d(2) < Min, d(2), Min)
            Min = IIf(d(3) < Min, d(3), Min)
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            Min = IIf(d(0) < d(1), d(0), d(1))
            Min = IIf(d(2) < Min, d(2), Min)
            Min = IIf(d(3) < Min, d(3), Min)
            Min = IIf(d(4) < Min, d(4), Min)
            Min = IIf(d(5) < Min, d(5), Min)
            Min = IIf(d(6) < Min, d(6), Min)
            Min = IIf(d(7) < Min, d(7), Min)
            Min = IIf(d(8) < Min, d(8), Min)
            Min = IIf(d(9) < Min, d(9), Min)
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case " "
    End Select
    
    WFCMin = Min
End Function

'概要      :WFC判定対象MAXデータを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WFCMax(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim max As Double
    
    max = -9999
    
    Select Case G.cPos
    Case "1"                                         '2003/05/15 追加　osawa 依頼No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(0) > d(3), d(0), d(3))
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
            max = IIf(d(6) > max, d(6), max)
            max = IIf(d(9) > max, d(9), max)
        End Select                                   '依頼No.030130　追加ここまで
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(0) > d(4), d(0), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(6), d(0), d(6))
            max = IIf(d(9) > max, d(9), max)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(0) > d(4), d(0), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(6), d(0), d(6))
            max = IIf(d(9) > max, d(9), max)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(0) > d(4), d(0), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(6), d(0), d(6))
            max = IIf(d(9) > max, d(9), max)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(1) > d(4), d(1), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(6), d(0), d(6))
            max = IIf(d(9) > max, d(9), max)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(1) > d(4), d(1), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(5), d(0), d(5))
            max = IIf(d(8) > max, d(8), max)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(1) > d(4), d(1), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(5), d(0), d(5))
            max = IIf(d(8) > max, d(8), max)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(1) > d(4), d(1), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(5), d(0), d(5))
            max = IIf(d(8) > max, d(8), max)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(2) > d(4), d(2), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(4), d(0), d(4))
            max = IIf(d(7) > max, d(7), max)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(3), d(0), d(3))
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(3) > d(4), d(3), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(1), d(0), d(1))
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(1) > d(3), d(1), d(3))
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
            max = IIf(d(6) > max, d(6), max)
            max = IIf(d(9) > max, d(9), max)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(1) > d(3), d(1), d(3))
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
            max = IIf(d(5) > max, d(5), max)
            max = IIf(d(8) > max, d(8), max)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(2) > d(3), d(2), d(3))
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
            max = IIf(d(4) > max, d(4), max)
            max = IIf(d(7) > max, d(7), max)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(2) > d(4), d(2), d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(4), d(0), d(4))
            max = IIf(d(7) > max, d(7), max)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(2) > d(3), d(2), d(3))
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
            max = IIf(d(4) > max, d(4), max)
            max = IIf(d(7) > max, d(7), max)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(1), d(0), d(1))
            max = IIf(d(3) > max, d(3), max)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            max = IIf(d(0) > d(1), d(0), d(1))
            max = IIf(d(2) > max, d(2), max)
            max = IIf(d(3) > max, d(3), max)
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            max = IIf(d(0) > d(1), d(0), d(1))
            max = IIf(d(2) > max, d(2), max)
            max = IIf(d(3) > max, d(3), max)
            max = IIf(d(4) > max, d(4), max)
            max = IIf(d(5) > max, d(5), max)
            max = IIf(d(6) > max, d(6), max)
            max = IIf(d(7) > max, d(7), max)
            max = IIf(d(8) > max, d(8), max)
            max = IIf(d(9) > max, d(9), max)
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case " "
    End Select

    WFCMax = max
End Function

'概要      :WFC判定対象aveデータを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WFCAve(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim AVE As Double
    
    AVE = -9999
    
    Select Case G.cPos
    Case "1"                                               '2003/05/15 追加　osawa 依頼No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(0) + d(3) + d(4)) / 3
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(2) + d(6) + d(9)) / 4
        End Select                                         '依頼No.030130　追加ここまで
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(0) + d(4)) / 2
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(6) + d(9)) / 3
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(0) + d(4)) / 2
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(6) + d(9)) / 3
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(0) + d(4)) / 2
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(6) + d(9)) / 3
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(1) + d(4)) / 2
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(6) + d(9)) / 3
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(1) + d(4)) / 2
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(5) + d(8)) / 3
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(1) + d(4)) / 2
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(5) + d(8)) / 3
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(1) + d(4)) / 2
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(5) + d(8)) / 3
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(2) + d(4)) / 2
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(4) + d(7)) / 3
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(3)) / 2
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(3) + d(4)) / 2
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(2)) / 2
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(1)) / 2
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(1) + d(3) + d(4)) / 3
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(2) + d(6) + d(9)) / 4
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(1) + d(3) + d(4)) / 3
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(2) + d(5) + d(8)) / 4
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(2) + d(3) + d(4)) / 3
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(2) + d(4) + d(7)) / 4
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(2) + d(4)) / 2
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(4) + d(7)) / 3
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(2) + d(3) + d(4)) / 3
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(2) + d(4) + d(7)) / 4
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(1) + d(3)) / 3
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            AVE = (d(0) + d(1) + d(2) + d(3) + d(4)) / 5
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            AVE = (d(0) + d(1) + d(2) + d(3) + d(4) + d(5) + d(6) + d(7) + d(8) + d(9)) / 10
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case " "
    End Select

    WFCAve = AVE
End Function

'概要      :WFC判定対象センター位置データを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WFCCenterP(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim CenterP As Double
    
    CenterP = -9999
    
    Select Case G.cPos
    Case "1"                                      '2003/05/15 追加　osawa 依頼No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select                                '依頼No.030130　追加ここまで
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "D"
            CenterP = d(0)
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterP = d(0)
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case " "
    End Select

    WFCCenterP = CenterP
End Function

'概要      :WFC判定対象中央値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WFCCenterD(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim CenterD As Double
    Dim temp() As Double
    
    CenterD = -9999
    
    Select Case G.cPos
    Case "1"                                       '2003/05/15 追加　osawa 依頼No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(3)
            temp(2) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(3) As Double
            temp(0) = d(0)
            temp(1) = d(2)
            temp(2) = d(6)
            temp(3) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        End Select                                '依頼No.030130　追加ここまで
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(1) As Double
            temp(0) = d(0)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(6)
            temp(2) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(1) As Double
            temp(0) = d(0)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(6)
            temp(2) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(1) As Double
            temp(0) = d(1)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(6)
            temp(2) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(1) As Double
            temp(0) = d(1)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(6)
            temp(2) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(1) As Double
            temp(0) = d(1)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(5)
            temp(2) = d(8)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(1) As Double
            temp(0) = d(1)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(5)
            temp(2) = d(8)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(1) As Double
            temp(0) = d(1)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(5)
            temp(2) = d(8)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(1) As Double
            temp(0) = d(2)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(4)
            temp(2) = d(7)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(1) As Double
            temp(0) = d(0)
            temp(1) = d(3)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(1) As Double
            temp(0) = d(3)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(1) As Double
            temp(0) = d(0)
            temp(1) = d(2)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(1) As Double
            temp(0) = d(0)
            temp(1) = d(1)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(2) As Double
            temp(0) = d(1)
            temp(1) = d(3)
            temp(2) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(3) As Double
            temp(0) = d(0)
            temp(1) = d(2)
            temp(2) = d(6)
            temp(3) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(2) As Double
            temp(0) = d(1)
            temp(1) = d(3)
            temp(2) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(3) As Double
            temp(0) = d(0)
            temp(1) = d(2)
            temp(2) = d(5)
            temp(3) = d(8)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(2) As Double
            temp(0) = d(2)
            temp(1) = d(3)
            temp(2) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(3) As Double
            temp(0) = d(0)
            temp(1) = d(2)
            temp(2) = d(4)
            temp(3) = d(7)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(1) As Double
            temp(0) = d(2)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(4)
            temp(2) = d(7)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(2) As Double
            temp(0) = d(2)
            temp(1) = d(3)
            temp(2) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(3) As Double
            temp(0) = d(0)
            temp(1) = d(2)
            temp(2) = d(4)
            temp(3) = d(7)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(2)
            temp(2) = d(3)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            CenterD = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            CenterD = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ReDim temp(4) As Double
            temp(0) = d(0)
            temp(1) = d(1)
            temp(2) = d(2)
            temp(3) = d(3)
            temp(4) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((4 + 1) / 2))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ReDim temp(9) As Double
            temp(0) = d(0)
            temp(1) = d(1)
            temp(2) = d(2)
            temp(3) = d(3)
            temp(4) = d(4)
            temp(5) = d(5)
            temp(6) = d(6)
            temp(7) = d(7)
            temp(8) = d(8)
            temp(9) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((4 + 1) / 2))
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case " "
    End Select
    
    WFCCenterD = CenterD
End Function

'概要      :WFC判定対象R/2値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WFCR2(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim r2 As Double
    
    r2 = -9999
    
    Select Case G.cPos
    Case "1"                                         '2003/05/15 追加　osawa 依頼No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            r2 = d(3)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            r2 = d(2)
        End Select                                   '依頼No.030130　追加ここまで
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            r2 = d(3)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            r2 = d(2)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            r2 = d(3)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            r2 = d(2)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            r2 = d(3)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            r2 = d(2)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            r2 = d(3)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            r2 = d(2)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            r2 = d(3)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            r2 = d(2)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            r2 = d(3)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            r2 = d(2)
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case " "
    End Select
    
    WFCR2 = r2
End Function

'概要      :WFC判定対象（|Center-Side|Max）値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WFCCE_Side_Max(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim ce_side_max As Double
    Dim ce_side0 As Double
    Dim ce_side1 As Double
    Dim ce_side2 As Double
    Dim ce_side3 As Double
    Dim ce_side4 As Double
    Dim ce_side5 As Double
    
    ce_side_max = -9999
    
    Select Case G.cPos
    Case "1"                                              '2003/05/15 追加　osawa 依頼No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(0) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select                                         '依頼No.030130　追加ここまで
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(0) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(0) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(0) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(5))
            ce_side1 = Abs(d(0) - d(8))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(5))
            ce_side1 = Abs(d(0) - d(8))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(5))
            ce_side1 = Abs(d(0) - d(8))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(2) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(4))
            ce_side1 = Abs(d(0) - d(7))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(5))
            ce_side1 = Abs(d(0) - d(8))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(2) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(4))
            ce_side1 = Abs(d(0) - d(7))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(2) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(4))
            ce_side1 = Abs(d(0) - d(7))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side_max = Abs(d(2) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(4))
            ce_side1 = Abs(d(0) - d(7))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ce_side0 = Abs(d(0) - d(4))
            ce_side1 = Abs(d(1) - d(4))
            ce_side2 = Abs(d(2) - d(4))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
            ce_side_max = IIf(ce_side2 > ce_side_max, ce_side2, ce_side_max)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ce_side0 = Abs(d(0) - d(4))
            ce_side1 = Abs(d(0) - d(5))
            ce_side2 = Abs(d(0) - d(6))
            ce_side3 = Abs(d(0) - d(7))
            ce_side4 = Abs(d(0) - d(8))
            ce_side5 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
            ce_side_max = IIf(ce_side2 > ce_side_max, ce_side2, ce_side_max)
            ce_side_max = IIf(ce_side3 > ce_side_max, ce_side3, ce_side_max)
            ce_side_max = IIf(ce_side4 > ce_side_max, ce_side4, ce_side_max)
            ce_side_max = IIf(ce_side5 > ce_side_max, ce_side5, ce_side_max)
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case " "
    End Select
    
    WFCCE_Side_Max = ce_side_max
End Function

'概要      :WFC判定対象中心平均値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WFCCEAve(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim ceave As Double
    
    ceave = -9999
    
    Select Case G.cPos
    Case "1"                                        '2003/05/15 追加　osawa 依頼No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select                                   '依頼No.030130　追加ここまで
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(0)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            ceave = d(4)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            ceave = d(0)
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case " "
    End Select
    
    WFCCEAve = ceave
End Function

'概要      :WFC判定対象Side平均値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WFCSideAve(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim sideave As Double
    
    sideave = -9999
    
    Select Case G.cPos
    Case "1"                                       '2003/05/15 追加　osawa 依頼No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(0)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select                                  '依頼No.030130　追加ここまで
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(0)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(0)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(0)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(1)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(1)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(5) + d(8)) / 2
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(1)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(5) + d(8)) / 2
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(1)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(5) + d(8)) / 2
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(2)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(4) + d(7)) / 2
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(1)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(1)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(5) + d(8)) / 2
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(2)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(4) + d(7)) / 2
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(2)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(4) + d(7)) / 2
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = d(2)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(4) + d(7)) / 2
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            sideave = (d(0) + d(1) + d(2)) / 3
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            sideave = (d(4) + d(5) + d(6) + d(7) + d(8) + d(9)) / 6
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case " "
    End Select
    
    WFCSideAve = sideave
End Function

'概要      :WFC判定対象（|Center-R/2|Max）値を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function WFCCE_R2_Max(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim cd_r2_max As Double
    
    cd_r2_max = -9999
    
    Select Case G.cPos
    Case "1"                                           '2003/05/15 追加　osawa 依頼No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select                                      '依頼No.030130　追加ここまで
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case " "
    End Select
    
    WFCCE_R2_Max = cd_r2_max
End Function



'概要      :面内分布計算式[N]に対応し、ＲＯＧ値を算出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Res           ,I  ,W_RES     ,WFセンター比抵抗判定構造体
'          :戻り値        ,O  ,double    ,RRG
'説明      :
'履歴      :2002/10/18  yakimura  作成
'
Public Function WF_TypeN_Exc(JudgFlag As Integer, d() As Double, G As Guarantee, C As Double) As Double
    Dim auto_cal As Double
    Dim auto_cal1 As Double
    Dim auto_cal2 As Double
    Dim auto_cal3 As Double
    Dim auto_cal4 As Double
    Dim auto_cal5 As Double
    Dim auto_cal6 As Double
    
    auto_cal = -9999
    
    Select Case G.cPos
    Case "1"                                                      '2003/05/15 追加　osawa 依頼No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(0))
             auto_cal = Abs(C - d(0)) / Abs(C + d(0)) * 200
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(6) <> -9999) Then
            auto_cal1 = Abs(C - d(6)) / Abs(C + d(6)) * 200
          End If

          If (C <> -9999) And (d(9) <> -9999) Then
                auto_cal2 = Abs(C - d(9)) / Abs(C + d(9)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select                                                '依頼No.030130　追加ここまで
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(4))
            auto_cal = Abs(C - d(4)) / Abs(C + d(4)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(6) <> -9999) Then
            auto_cal1 = Abs(C - d(6)) / Abs(C + d(6)) * 200
          End If

          If (C <> -9999) And (d(9) <> -9999) Then
                auto_cal2 = Abs(C - d(9)) / Abs(C + d(9)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(4))
            auto_cal = Abs(C - d(4)) / Abs(C + d(4)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(6) <> -9999) Then
            auto_cal1 = Abs(C - d(6)) / Abs(C + d(6)) * 200
          End If

          If (C <> -9999) And (d(9) <> -9999) Then
                auto_cal2 = Abs(C - d(9)) / Abs(C + d(9)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(4))
            auto_cal = Abs(C - d(4)) / Abs(C + d(4)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(6) <> -9999) Then
            auto_cal1 = Abs(C - d(6)) / Abs(C + d(6)) * 200
          End If

          If (C <> -9999) And (d(9) <> -9999) Then
                auto_cal2 = Abs(C - d(9)) / Abs(C + d(9)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(6) <> -9999) Then
            auto_cal1 = Abs(C - d(6)) / Abs(C + d(6)) * 200
          End If

          If (C <> -9999) And (d(9) <> -9999) Then
                auto_cal2 = Abs(C - d(9)) / Abs(C + d(9)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(5) <> -9999) Then
            auto_cal1 = Abs(C - d(5)) / Abs(C + d(5)) * 200
          End If

          If (C <> -9999) And (d(8) <> -9999) Then
                auto_cal2 = Abs(C - d(8)) / Abs(C + d(8)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(5) <> -9999) Then
            auto_cal1 = Abs(C - d(5)) / Abs(C + d(5)) * 200
          End If

          If (C <> -9999) And (d(8) <> -9999) Then
                auto_cal2 = Abs(C - d(8)) / Abs(C + d(8)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(5) <> -9999) Then
            auto_cal1 = Abs(C - d(5)) / Abs(C + d(5)) * 200
          End If

          If (C <> -9999) And (d(8) <> -9999) Then
                auto_cal2 = Abs(C - d(8)) / Abs(C + d(8)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(2))
            auto_cal = Abs(C - d(2)) / Abs(C + d(2)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(4) <> -9999) Then
            auto_cal1 = Abs(C - d(4)) / Abs(C + d(4)) * 200
          End If

          If (C <> -9999) And (d(7) <> -9999) Then
                auto_cal2 = Abs(C - d(7)) / Abs(C + d(7)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(6) <> -9999) Then
            auto_cal1 = Abs(C - d(6)) / Abs(C + d(6)) * 200
          End If

          If (C <> -9999) And (d(9) <> -9999) Then
                auto_cal2 = Abs(C - d(9)) / Abs(C + d(9)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(5) <> -9999) Then
            auto_cal1 = Abs(C - d(5)) / Abs(C + d(5)) * 200
          End If

          If (C <> -9999) And (d(8) <> -9999) Then
                auto_cal2 = Abs(C - d(8)) / Abs(C + d(8)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(2))
            auto_cal = Abs(C - d(2)) / Abs(C + d(2)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(4) <> -9999) Then
            auto_cal1 = Abs(C - d(4)) / Abs(C + d(4)) * 200
          End If

          If (C <> -9999) And (d(7) <> -9999) Then
                auto_cal2 = Abs(C - d(7)) / Abs(C + d(7)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(2))
            auto_cal = Abs(C - d(2)) / Abs(C + d(2)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(4) <> -9999) Then
            auto_cal1 = Abs(C - d(4)) / Abs(C + d(4)) * 200
          End If

          If (C <> -9999) And (d(7) <> -9999) Then
                auto_cal2 = Abs(C - d(7)) / Abs(C + d(7)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
            'auto_cal = Abs(d(2))
            auto_cal = Abs(C - d(2)) / Abs(C + d(2)) * 200           '2003/5/16
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(4) <> -9999) Then
            auto_cal1 = Abs(C - d(4)) / Abs(C + d(4)) * 200
          End If

          If (C <> -9999) And (d(7) <> -9999) Then
                auto_cal2 = Abs(C - d(7)) / Abs(C + d(7)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
          If (C <> -9999) And (d(0) <> -9999) Then
            auto_cal1 = Abs(C - d(0)) / Abs(C + d(0)) * 200
          End If

          If (C <> -9999) And (d(1) <> -9999) Then
            auto_cal2 = Abs(C - d(1)) / Abs(C + d(1)) * 200
          End If

          If (C <> -9999) And (d(2) <> -9999) Then
            auto_cal3 = Abs(C - d(2)) / Abs(C + d(2)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
          auto_cal = IIf(auto_cal > auto_cal3, auto_cal, auto_cal3)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select

    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
          If (C <> -9999) And (d(4) <> -9999) Then
            auto_cal1 = Abs(C - d(4)) / Abs(C + d(4)) * 200
          End If

          If (C <> -9999) And (d(5) <> -9999) Then
            auto_cal2 = Abs(C - d(5)) / Abs(C + d(5)) * 200
          End If

          If (C <> -9999) And (d(6) <> -9999) Then
            auto_cal3 = Abs(C - d(6)) / Abs(C + d(6)) * 200
          End If

          If (C <> -9999) And (d(7) <> -9999) Then
            auto_cal4 = Abs(C - d(7)) / Abs(C + d(7)) * 200
          End If

          If (C <> -9999) And (d(8) <> -9999) Then
            auto_cal5 = Abs(C - d(8)) / Abs(C + d(8)) * 200
          End If

          If (C <> -9999) And (d(9) <> -9999) Then
            auto_cal6 = Abs(C - d(9)) / Abs(C + d(9)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
          auto_cal = IIf(auto_cal > auto_cal3, auto_cal, auto_cal3)
          auto_cal = IIf(auto_cal > auto_cal4, auto_cal, auto_cal4)
          auto_cal = IIf(auto_cal > auto_cal5, auto_cal, auto_cal5)
          auto_cal = IIf(auto_cal > auto_cal6, auto_cal, auto_cal6)
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''判定識別フラグ(RES)
        Case WFOI_JUDG ''判定識別フラグ(Oi)
        End Select
    Case " "
    End Select
    
    WF_TypeN_Exc = auto_cal
End Function

''Upd start 2005/06/22 (TCS)t.terauchi  SPV9点対応
'概要      :WFセンターSPV(Fe濃度 MAP測定)判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Spv           ,I  ,W_SPV            ,WFセンターSPV判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2005/06/22 新規作成 (TCS)t.terauchi
Public Function WfSPV_Fe_AMXJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
        
    If Spv.GuaranteeSpvFe.cJudg = JudgCodeW01 Then ''SPVFE濃度　判定有り
                
        ''SPV(Fe濃度 MAP測定)判定
        Select Case Spv.GuaranteeSpvFe.cObj
            
            '全測定点(3)・MAX(B)
            Case ObjCode03, ObjCode06
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(0), 0, Spv.SpecSpvFeMax)
            
            'AVE(A)
            Case ObjCode05
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(2), 0, Spv.SpecSpvAvMax)
            
            'AVE+MAX(C)
            Case ObjCode07
                If RangeDecision_nl(Spv.Spv(2), 0, Spv.SpecSpvAvMax) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(0), 0, Spv.SpecSpvFeMax)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
            
            'その他
            Case Else
                ''狙い、規格無し以外の場合
                If (Spv.GuaranteeSpvFe.cObj <> ObjCode13) And (Spv.GuaranteeSpvFe.cObj <> ObjCode15) Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, SPV_JUDG, Spv.GuaranteeSpvFe.cObj)
                    GoTo EXIT_FUNC
                End If
                Spv.JudgSpv = JUDG_OK
        
        End Select
    Else
        Spv.JudgSpv = JUDG_OK
    End If
    
EXIT_FUNC:
    
    WfSPV_Fe_AMXJudg = FuncAns
End Function

'概要      :WFセンターSPV(Fe濃度 9点測定)判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Spv           ,I  ,W_SPV            ,WFセンターSPV判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2005/06/22 新規作成 (TCS)t.terauchi
Public Function WfSPV_Fe_V9TJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
        
    If Spv.GuaranteeSpvFe.cJudg = JudgCodeW01 Then ''SPVFE濃度　判定有り
        
        ''SPV(Fe濃度 9点測定)判定
        Select Case Spv.GuaranteeSpvFe.cObj
        
            '中心点(1)
            Case ObjCode01
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(3), 0, Spv.SpecSpvFeMax)
        
            '全測定点(3)・MAX(B)
            Case ObjCode03, ObjCode06
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(0), 0, Spv.SpecSpvFeMax)
            
            'AVE(A)
            Case ObjCode05
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(2), 0, Spv.SpecSpvAvMax)
            
            'AVE+MAX(C)
            Case ObjCode07
                If RangeDecision_nl(Spv.Spv(2), 0, Spv.SpecSpvAvMax) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(0), 0, Spv.SpecSpvFeMax)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
            
            'その他
            Case Else
                ''狙い、規格無し以外の場合
                If (Spv.GuaranteeSpvFe.cObj <> ObjCode13) And (Spv.GuaranteeSpvFe.cObj <> ObjCode15) Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, SPV_JUDG, Spv.GuaranteeSpvFe.cObj)
                    GoTo EXIT_FUNC
                End If
                Spv.JudgSpv = JUDG_OK
        
        End Select
    Else
        Spv.JudgSpv = JUDG_OK
    End If

EXIT_FUNC:
    
    WfSPV_Fe_V9TJudg = FuncAns
    
End Function

'概要      :WFセンターSPV(拡散長 MAP測定)判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Spv           ,I  ,W_SPV            ,WFセンターSPV判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2005/06/22 新規作成 (TCS)t.terauchi
Public Function WfSPV_DIFF_AMXJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
    
    If Spv.GuaranteeSpv.cJudg = JudgCodeW01 Then ''SPV拡散長　判定有り
                
        ''SPV(拡散長 MAP測定)判定
        Select Case Spv.GuaranteeSpv.cObj
        
            '全測定点(3)
            Case ObjCode03
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    If RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                        Spv.JudgSpv = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
                    Else
                        Spv.JudgSpv = JUDG_NG
                    End If
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
            
            'AVE(A)
            Case ObjCode05
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
            
            'MAX(B)
            Case ObjCode06
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax)
            
            'AVE+MAX(C)
            Case ObjCode07
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
            
            'MIN(D)
            Case ObjCode08
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax)
            
            'MIN+MAX(K)
            Case ObjCode16
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
               
            'AVE+MIN(L)　08/03/13 ooba
            Case ObjCode17
                'AVEが上限(AVE下限)以上かつ,MINが下限(MIN下限)以上
                If RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMax, -1) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, -1)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
                
            'その他
            Case Else
                ''狙い、規格無し以外の場合
                If (Spv.GuaranteeSpv.cObj <> ObjCode13) And (Spv.GuaranteeSpv.cObj <> ObjCode15) Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, SPV_JUDG, Spv.GuaranteeSpv.cObj)
                    GoTo EXIT_FUNC
                End If
                Spv.JudgSpv = JUDG_OK
        End Select
    Else
        Spv.JudgSpv = JUDG_OK
    End If
    
EXIT_FUNC:
    
    WfSPV_DIFF_AMXJudg = FuncAns
    
End Function

'概要      :WFセンターSPV(拡散長 9点測定)判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Spv           ,I  ,W_SPV            ,WFセンターSPV判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2005/06/22 新規作成 (TCS)t.terauchi
Public Function WfSPV_DIFF_V9TJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
    
    If Spv.GuaranteeSpv.cJudg = JudgCodeW01 Then ''SPV拡散長　判定有り
        
        ''SPV(拡散長 9点測定)
        Select Case Spv.GuaranteeSpv.cObj
            
            '中心点(1)
            Case ObjCode01
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(3), Spv.SpecSpvMin, Spv.SpecSpvMax)
            
            '全測定点(3)
            Case ObjCode03
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    If RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                        Spv.JudgSpv = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
                    Else
                        Spv.JudgSpv = JUDG_NG
                    End If
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
            
            'AVE(A)
            Case ObjCode05
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
            
            'MAX(B)
            Case ObjCode06
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax)
            
            'AVE+MAX(C)
            Case ObjCode07
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
            
            'MIN(D)
            Case ObjCode08
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax)
            
            'MIN+MAX(K)
            Case ObjCode16
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
            
            'AVE+MIN(L)　08/03/13 ooba
            Case ObjCode17
                'AVEが上限(AVE下限)以上かつ,MINが下限(MIN下限)以上
                If RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMax, -1) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, -1)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
                
            'その他
            Case Else
                ''狙い、規格無し以外の場合
                If (Spv.GuaranteeSpv.cObj <> ObjCode13) And (Spv.GuaranteeSpv.cObj <> ObjCode15) Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, SPV_JUDG, Spv.GuaranteeSpv.cObj)
                    GoTo EXIT_FUNC
                End If
                Spv.JudgSpv = JUDG_OK
        End Select
    Else
        Spv.JudgSpv = JUDG_OK
    End If
    
EXIT_FUNC:
    
    WfSPV_DIFF_V9TJudg = FuncAns
End Function

'概要      :WFセンターSPV(Nr濃度 MAP測定)判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Spv           ,I  ,W_SPV            ,WFセンターSPV判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2006/06/12 新規作成 SMP)kondoh
Public Function WfSPV_Nr_AMXJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
        
    If Spv.GuaranteeSpvNr.cJudg = JudgCodeW01 Then ''SPVNR濃度　判定有り
                
        ''SPV(Nr濃度 MAP測定)判定
        Select Case Spv.GuaranteeSpvNr.cObj
            
            '全測定点(3)・MAX(B)
            Case ObjCode03, ObjCode06
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(0), 0, Spv.SpecSpvNrMax)
            
            'AVE(A)
            Case ObjCode05
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(2), 0, Spv.SpecSpvNrAvMax)
            
            'AVE+MAX(C)
            Case ObjCode07
                If RangeDecision_nl(Spv.Spv(2), 0, Spv.SpecSpvNrAvMax) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(0), 0, Spv.SpecSpvNrMax)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
            
            'その他
            Case Else
                ''狙い、規格無し以外の場合
                If (Spv.GuaranteeSpvNr.cObj <> ObjCode13) And (Spv.GuaranteeSpvNr.cObj <> ObjCode15) Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, SPV_JUDG, Spv.GuaranteeSpvNr.cObj)
                    GoTo EXIT_FUNC
                End If
                Spv.JudgSpv = JUDG_OK
        
        End Select
    Else
        Spv.JudgSpv = JUDG_OK
    End If
    
EXIT_FUNC:
    
    WfSPV_Nr_AMXJudg = FuncAns
End Function

'概要      :WFセンターSPV(Nr濃度 9点測定)判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Spv           ,I  ,W_SPV            ,WFセンターSPV判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2006/06/12 新規作成 SMP)kondoh
Public Function WfSPV_Nr_V9TJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
                
    If Spv.GuaranteeSpvNr.cJudg = JudgCodeW01 Then ''SPVNR濃度　判定有り
        
        ''SPV(Fe濃度 9点測定)判定
        Select Case Spv.GuaranteeSpvNr.cObj

'' DB上にNr濃度の中心の項目が無いため、判定不可能
''            '中心点(1)
''            Case ObjCode01
''                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(3), 0, Spv.SpecSpvNrMax)
        
            '全測定点(3)・MAX(B)
            Case ObjCode03, ObjCode06
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(0), 0, Spv.SpecSpvNrMax)
            
            'AVE(A)
            Case ObjCode05
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(2), 0, Spv.SpecSpvNrMax)
            
            'AVE+MAX(C)
            Case ObjCode07
                If RangeDecision_nl(Spv.Spv(2), 0, Spv.SpecSpvNrAvMax) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(0), 0, Spv.SpecSpvNrMax)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
            
            'その他
            Case Else
                ''狙い、規格無し以外の場合
                If (Spv.GuaranteeSpvNr.cObj <> ObjCode13) And (Spv.GuaranteeSpvNr.cObj <> ObjCode15) Then
                    ''対象データ無し
                    ''エラー情報構造体に情報を代入。
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, SPV_JUDG, Spv.GuaranteeSpvNr.cObj)
                    GoTo EXIT_FUNC
                End If
                Spv.JudgSpv = JUDG_OK
        
        End Select
    Else
        Spv.JudgSpv = JUDG_OK
    End If

EXIT_FUNC:
    
    WfSPV_Nr_V9TJudg = FuncAns
    
End Function

