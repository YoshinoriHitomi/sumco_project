Attribute VB_Name = "s_cmzcwjres"
Option Explicit

''WF�Z���^�[���R����\����
Type W_RES
    GuaranteeRes    As Guarantee    ''�i���ۏ؏��\����
    GuaranteeCal    As String * 1   ''�iWF���R�ʓ��v�Z
    SpecResMin      As Double       ''�iWF���R����
    SpecResMax      As Double       ''�iWF���R���
    SpecRrg         As Double       ''�iWF���R�ʓ����z
    SpecResAveMin   As Double       ''�iWF���R���ω���
    SpecResAveMax   As Double       ''�iWF���R���Ϗ��
    Res(4)          As Double       ''���R����l
'���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    'AN���x����ǉ�
    ResAntnp        As Double       ''AN���x���R����l
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    RRG             As Double       ''RRG�v�Z�l
    JudgRes         As Boolean      ''���R����l
    JudgRrg         As Boolean      ''RRG���茋��
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    '�`�F�b�N�pAN���x��ǉ�
    JudgAntnp       As Boolean      ''�`�m���x���茋��
    Antnp           As Integer      ''�i�v�e�`�m���x
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    DkTmpSiyo       As String       ''DK���x�i�d�l�j
    DkTmpJsk        As String       ''DK���x�i���сj
    JudgDkTmp       As Boolean      ''DK���x����l
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
End Type

'�T�v      :WF�Z���^�[���R������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Res           ,I  ,W_RES            ,WF�Z���^�[���R����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WfRESJudg(Res As W_RES, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFRES_JUDG = 1                 ''���莯�ʃt���O(RES)
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim center As Double
    Dim r2 As Double
    Dim AllTemp() As Double
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    Dim liRet As Integer
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Dim iRet    As Integer
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    
    Res.JudgRrg = JUDG_NG
    Res.JudgRes = JUDG_NG
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Res.JudgDkTmp = JUDG_NG
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    If Res.GuaranteeRes.cJudg = JudgCodeW01 Then ''RES����L��
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "���R���� **********"
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�_ = " & Res.GuaranteeRes.cCount
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Res.GuaranteeRes.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Res.GuaranteeRes.cPos
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Res.GuaranteeRes.cObj
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Res.GuaranteeRes.cJudg
'''''        WFCJudgDialog.WFCErrorMessage "���z�v�Z = " & Res.GuaranteeCal
'''''        WFCJudgDialog.WFCErrorMessage "���R�ʓ����z = " & Str(Res.SpecRrg)

        
        ''RRG����
        'RRG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX(�����ł͊ۂ߂Ȃ�) 2011/11/25 SETsw kubota
        'Res.RRG = WFCRRGCal(Res.Res(), Res.GuaranteeRes, Res.GuaranteeCal)
        Res.RRG = WFCRRGCal_NotRound(Res.Res(), Res.GuaranteeRes, Res.GuaranteeCal)
'2002/02/27 S.Sano RRG�̎d�l��0�̏ꍇ�́A������s�킸�K��OK�Ƃ���B
'2002/02/27 S.Sano �ʓ����z�v�Z�͍s���B
        If Res.SpecRrg = 0 Then                                     '2002/02/27 S.Sano
            Res.JudgRrg = JUDG_OK                                   '2002/02/27 S.Sano
        Else                                                        '2002/02/27 S.Sano
            If Res.RRG = -1 Then
                Res.JudgRrg = JUDG_NG
            Else
                Res.JudgRrg = RangeDecision_nl(Res.RRG, 0, Res.SpecRrg)
            End If
        End If                                                      '2002/02/27 S.Sano
        
        ''RES����
        If (InStr(ObjCodeGrp01, Res.GuaranteeRes.cObj) <> 0) Then
            Select Case Res.GuaranteeRes.cObj
            Case ObjCode01  ''���S1�_
                center = WFCCenterP(WFRES_JUDG, Res.Res(), Res.GuaranteeRes)
                If center = -9999 Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
'''''                    WFCJudgDialog.WFCErrorMessage "���R����A�Ώۃf�[�^�����B"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
                Else
                    Res.JudgRes = RangeDecision_nl(center, Res.SpecResMin, Res.SpecResMax)
                End If
            Case ObjCode02  ''�����l
                center = WFCCenterD(WFRES_JUDG, Res.Res(), Res.GuaranteeRes)
                If center = -9999 Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
'''''                    WFCJudgDialog.WFCErrorMessage "���R����A�Ώۃf�[�^�����B"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
                Else
                    Res.JudgRes = RangeDecision_nl(center, Res.SpecResMin, Res.SpecResMax)
                End If
            Case ObjCode03 ''�S��
                If WFCJudgDataSelect_All(Res.Res(), Res.GuaranteeRes, AllTemp()) = FUNCTION_RETURN_FAILURE Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
'''''                    WFCJudgDialog.WFCErrorMessage "���R����A�Ώۃf�[�^�����B"
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
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
'''''                    WFCJudgDialog.WFCErrorMessage "���R����A�Ώۃf�[�^�����B"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
                Else
                    Res.JudgRes = RangeDecision_nl(r2, Res.SpecResMin, Res.SpecResMax)
                End If
            End Select
        Else
            If (Res.GuaranteeRes.cObj <> ObjCode13) And (Res.GuaranteeRes.cObj <> ObjCode15) Then
                ''�Ώۃf�[�^����
                ''�G���[���\���̂ɏ������B
'''''                WFCJudgDialog.WFCErrorMessage "���R����A�Ώۃf�[�^�����B"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
            End If
        End If
    
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
        ''AN���x����
        liRet = funCodeDBGetMatrixReturn("SB", "AR", CStr(Res.Antnp), CStr(Res.ResAntnp))
        If liRet = -1 Then
            FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
        ElseIf liRet = 0 Then
            Res.JudgAntnp = JUDG_NG
        Else
            Res.JudgAntnp = JUDG_OK
        End If
        
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------

'--------------- 2008/08/25 INSERT START  By Systech ---------------
        ''DK���x����
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
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, RES_JUDG, Res.GuaranteeRes.cJudg)
'        End If
    End If

    WfRESJudg = FuncAns
End Function

'�T�v      :RRG�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
'�[���ۂ߂��s���֐��ƍs��Ȃ��֐��̓�ɕ��� 2011/11/25 SETsw kubota
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
    Case "A" '(max-min)/min�~100
        Min = WFCMin(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If (Min <> 0) Then
                RRG = (max - Min) * 100 / Min
            Else
                deverrflag = True
            End If
        End If
    Case "B" '(max-min)/max�~100
        Min = WFCMin(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If (max <> 0) Then
                RRG = (max - Min) * 100 / max
            Else
                deverrflag = True
            End If
        End If
    Case "C" '(max-min)/center�~100
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
    Case "D" '|center-side|max/center�~100
        side = WFCCE_Side_Max(WFRES_JUDG, R(), G)
        center = WFCCenterP(WFRES_JUDG, R(), G)
        If (side <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                RRG = side * 100 / center
            Else
                deverrflag = True
            End If
        End If
    Case "E" '(centerave-sideave)/centerave�~100
        center = WFCCEAve(WFRES_JUDG, R(), G)
        side = WFCSideAve(WFRES_JUDG, R(), G)
        If (side <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                RRG = Abs(center - side) * 100 / center  '|center-side| 2002/6/28 osawa
            Else
                deverrflag = True
            End If
        End If
    Case "F" '|center-R/2|max/center�~100
        r2 = WFCCE_R2_Max(WFRES_JUDG, R(), G)
        center = WFCCenterP(WFRES_JUDG, R(), G)
        If (r2 <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                RRG = r2 * 100 / center
            Else
                deverrflag = True
            End If
        End If
    Case "G" '2(side-center)/(sideave+center)�~100
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
    Case "H" '(max-ave)/ave�~100
        AVE = WFCAve(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        If (max <> -9999) And (AVE <> -9999) Then
            If (AVE <> 0) Then
                RRG = (max - AVE) * 100 / AVE
            Else
                deverrflag = True
            End If
        End If
    Case "K" '(max-min)/(max+min)�~100
        Min = WFCMin(WFRES_JUDG, R(), G)
        max = WFCMax(WFRES_JUDG, R(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If ((max + Min) <> 0) Then
                RRG = (max - AVE) * 100 / (max + Min)
            Else
                deverrflag = True
            End If
        End If
    Case "L" '(max-min)/2�~ave�~100
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
    Case "M" '(max-min)/ave�~100
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

'�R�[�h"N"�ɑΉ��@��tNo.20409  <2002.10.11 yakimura> start
    
    Case "N" '|(center-side)/(center+side)|�~200
        
        C = WFCCenterP(WFRES_JUDG, R(), G)
        'RRG = WF_TypeN_Exc(WFOI_JUDG, R(), G, C)
        RRG = WF_TypeN_Exc(WFRES_JUDG, R(), G, C)                       '2003/5/15
        
        If RRG = -9999 Then
              errflag = True
        End If

'�R�[�h"N"�ɑΉ��@��tNo.20409  <2002.10.11 yakimura> end
    
    Case " "

'''''        WFCJudgDialog.WFCErrorMessage "���z�v�Z����` A �ɂČv�Z"

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
'''''        WFCJudgDialog.WFCErrorMessage "���z�v�Z����`"
    ElseIf deverrflag Then
'''''        WFCJudgDialog.WFCErrorMessage "���z�v�Z 0 ���Z�G���["
    ElseIf RRG = -1 Then
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu�A�Ώۃf�[�^�A���z�v�Z����"
    End If

    '2002/07/24 Update T.Hayashi
    'WFCRRGCal = RRG
    'WFCRRGCal = RoundUp(RRG, 4)
    WFCRRGCal_NotRound = RRG        '�ۂ߂Ȃ��l��Ԃ��悤�ɕύX 2011/11/25 SETsw kubota

End Function

'�[���ۂ߂��s���֐��ƍs��Ȃ��֐��̓�ɕ��� 2011/11/25 SETsw kubota
Public Function WFCRRGCal(R() As Double, G As Guarantee, calcode As String) As Double
    '�[���ۂ߂��s��Ȃ��֐����Ăяo���A����2��(3���ڐ؂�グ)�ɂ��ĕԂ�
    WFCRRGCal = RoundUp(WFCRRGCal_NotRound(R(), G, calcode), 4)
End Function


Private Function WFCJudgDataSelect_All(d() As Double, G As Guarantee, T() As Double) As FUNCTION_RETURN
    Dim Func_Ans As FUNCTION_RETURN
    
    Func_Ans = FUNCTION_RETURN_FAILURE
    
    Select Case G.cPos
    Case "1"                                 '2003/05/15 �ǉ��@osawa �˗�No.030130
        ReDim T(2) As Double
        T(0) = d(0)
        T(1) = d(3)
        T(2) = d(4)
        Func_Ans = FUNCTION_RETURN_SUCCESS   '�˗�No.030130�@�ǉ������܂�
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
