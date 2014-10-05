Attribute VB_Name = "s_cmzcwjoi"
Option Explicit
''WF�Z���^�[Oi����\����
Type W_OI
    GuaranteeOi     As Guarantee    ''�i���ۏ؏��\����
    GuaranteeCal    As String * 1   ''�iWF�_�f�Z�x�ʓ��v�Z
    SpecOiMin       As Double       ''�iWF�_�f�Z�x����
    SpecOiMax       As Double       ''�iWF�_�f�Z�x���
    SpecORG         As Double       ''�iWF�_�f�Z�x�ʓ����z
    SpecOiAveMin    As Double       ''�iWF�_�f�Z�x���ω���
    SpecOiAveMax    As Double       ''�iWF�_�f�Z�x���Ϗ��
    Oi(9)           As Double       ''Oi����l
'���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    OiAntnp         As Double       ''AN���xOi����l
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    ORG             As Double       ''ORG�v�Z�l
    OiMin           As Double       ''OiMin�v�Z�l
    OiMax           As Double       ''OiMax�v�Z�l
    JudgOi          As Boolean      ''Oi���茋��
    JudgOrg         As Boolean      ''ORG���茋��
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    '�`�F�b�N�pAN���x��ǉ�
    JudgAntnp       As Boolean      ''�`�m���x���茋��
    Antnp           As Integer      ''�i�v�e�`�m���x
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
End Type

'�T�v      :WF�Z���^�[Oi������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Oi            ,I  ,W_OI             ,WF�Z���^�[Oi����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WfOiJudg(Oi As W_OI, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFOI_JUDG = 2                  ''���莯�ʃt���O(Oi)
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim center As Double
    Dim r2 As Double
    Dim AllTemp() As Double
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    Dim liRet As Integer
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    

    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Oi.JudgOi = JUDG_NG
    Oi.JudgOrg = JUDG_NG
    If Oi.GuaranteeOi.cJudg = JudgCodeW01 Then ''Oi����L��
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "�_�f�Z�x���� **********"
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�_ = " & Oi.GuaranteeOi.cCount
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Oi.GuaranteeOi.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Oi.GuaranteeOi.cPos
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Oi.GuaranteeOi.cObj
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Oi.GuaranteeOi.cJudg
'''''        WFCJudgDialog.WFCErrorMessage "���z�v�Z = " & Oi.GuaranteeCal
'''''        WFCJudgDialog.WFCErrorMessage "�_�f�Z�x�ʓ����z = " & Str(Oi.SpecORG)

        ''ORG����
        'ORG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX(�����ł͊ۂ߂Ȃ�) 2011/11/25 SETsw kubota
        'Oi.ORG = WFCORGCal(Oi.Oi(), Oi.GuaranteeOi, Oi.GuaranteeCal)
        Oi.ORG = WFCORGCal_NotRound(Oi.Oi(), Oi.GuaranteeOi, Oi.GuaranteeCal)
        Oi.OiMin = WFCMin(WFOI_JUDG, Oi.Oi(), Oi.GuaranteeOi)
        Oi.OiMax = WFCMax(WFOI_JUDG, Oi.Oi(), Oi.GuaranteeOi)
        
'2002/02/27 S.Sano ORG�̎d�l��0�̏ꍇ�́A������s�킸�K��OK�Ƃ���B
'2002/02/27 S.Sano �ʓ����z�v�Z�͍s���B
        If Oi.SpecORG = 0 Then                                      '2002/02/27 S.Sano
            Oi.JudgOrg = JUDG_OK                                    '2002/02/27 S.Sano
        Else                                                        '2002/02/27 S.Sano
            If Oi.ORG = -1 Then
                Oi.JudgOrg = JUDG_NG
            Else
                Oi.JudgOrg = RangeDecision_nl(Oi.ORG, 0, Oi.SpecORG)
            End If
        End If                                                      '2002/02/27 S.Sano
        
        ''Oi����
        If (InStr(ObjCodeGrp01, Oi.GuaranteeOi.cObj) <> 0) Then
            Select Case Oi.GuaranteeOi.cObj
            Case ObjCode01  ''���S1�_
                center = WFCCenterP(WFOI_JUDG, Oi.Oi(), Oi.GuaranteeOi)
                If center = -9999 Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
'''''                    WFCJudgDialog.WFCErrorMessage "�_�f�Z�x����A�Ώۃf�[�^�����B"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
                Else
                    Oi.JudgOi = RangeDecision_nl(center, Oi.SpecOiMin, Oi.SpecOiMax)
                End If
            Case ObjCode02  ''�����l
                center = WFCCenterP(WFOI_JUDG, Oi.Oi(), Oi.GuaranteeOi)
                If center = -9999 Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
'''''                    WFCJudgDialog.WFCErrorMessage "�_�f�Z�x����A�Ώۃf�[�^�����B"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
                Else
                    Oi.JudgOi = RangeDecision_nl(center, Oi.SpecOiMin, Oi.SpecOiMax)
                End If
            Case ObjCode03 ''�S��
                If WFCJudgDataSelect_All(Oi.Oi(), Oi.GuaranteeOi, AllTemp()) = FUNCTION_RETURN_FAILURE Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
'''''                    WFCJudgDialog.WFCErrorMessage "�_�f�Z�x����A�Ώۃf�[�^�����B"
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
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
'''''                    WFCJudgDialog.WFCErrorMessage "�_�f�Z�x����A�Ώۃf�[�^�����B"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
                Else
                    Oi.JudgOi = RangeDecision_nl(r2, Oi.SpecOiMin, Oi.SpecOiMax)
                End If
            End Select
        Else
            ''�_���A�K�i�����ȊO�̏ꍇ
            If (Oi.GuaranteeOi.cObj <> ObjCode13) And (Oi.GuaranteeOi.cObj <> ObjCode15) Then
                ''�Ώۃf�[�^����
                ''�G���[���\���̂ɏ������B
'''''                WFCJudgDialog.WFCErrorMessage "�_�f�Z�x����A�Ώۃf�[�^�����B"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
            End If
        End If
        
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
        ''AN���x����
        liRet = funCodeDBGetMatrixReturn("SB", "AO", CStr(Oi.Antnp), CStr(Oi.OiAntnp))
        If liRet = -1 Then
            FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
        ElseIf liRet = 0 Then
            Oi.JudgAntnp = JUDG_NG
        Else
            Oi.JudgAntnp = JUDG_OK
        End If
                
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    Else
        Oi.JudgOi = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -s-
        Oi.JudgAntnp = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -e-
'        If InStr(JudgCodeW02, Oi.GuaranteeOi.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, OI_JUDG, Oi.GuaranteeOi.cJudg)
'        End If
    End If
    
    WfOiJudg = FuncAns
End Function

'�T�v      :ORG�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :iMax          ,I  ,Integer   ,����_��
'          :�߂�l        ,O  ,double    ,org,ORG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
'�[���ۂ߂��s���֐��ƍs��Ȃ��֐��̓�ɕ��� 2011/11/25 SETsw kubota
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
    Case "A" '(max-min)/min�~100
        Min = WFCMin(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If (Min <> 0) Then
                ORG = (max - Min) * 100 / Min
            Else
                deverrflag = True
            End If
        End If
    Case "B" '(max-min)/max�~100
        Min = WFCMin(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If (max <> 0) Then
                ORG = (max - Min) * 100 / max
            Else
                deverrflag = True
            End If
        End If
    Case "C" '(max-min)/center�~100
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
    Case "D" '|center-side|max/center�~100
        side = WFCCE_Side_Max(WFOI_JUDG, O(), G)
        center = WFCCenterP(WFOI_JUDG, O(), G)
        If (side <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                ORG = side * 100 / center
            Else
                deverrflag = True
            End If
        End If
    Case "E" '(centerave-sideave)/centerave�~100
        center = WFCCEAve(WFOI_JUDG, O(), G)
        side = WFCSideAve(WFOI_JUDG, O(), G)
        If (side <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                ORG = Abs(center - side) * 100 / center  '|center-side| 2002/6/28 osawa
            Else
                deverrflag = True
            End If
        End If
    Case "F" '|center-R/2|max/center�~100
        r2 = WFCCE_R2_Max(WFOI_JUDG, O(), G)
        center = WFCCenterP(WFOI_JUDG, O(), G)
        If (r2 <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                ORG = r2 * 100 / center
            Else
                deverrflag = True
            End If
        End If
    Case "G" '2(|center-side|max)/(sideave+center)�~100
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
    Case "H" '(max-ave)/ave�~100
        AVE = WFCAve(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        If (max <> -9999) And (AVE <> -9999) Then
            If (AVE <> 0) Then
                ORG = (max - AVE) * 100 / AVE
            Else
                deverrflag = True
            End If
        End If
    Case "K" '(max-min)/(max+min)�~100
        Min = WFCMin(WFOI_JUDG, O(), G)
        max = WFCMax(WFOI_JUDG, O(), G)
        If (Min <> -9999) And (max <> -9999) Then
            If ((max + Min) <> 0) Then
                ORG = (max - AVE) * 100 / (max + Min)
            Else
                deverrflag = True
            End If
        End If
    Case "L" '(max-min)/2�~ave�~100
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
    Case "M" '(max-min)/ave�~100
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

'�R�[�h"N"�ɑΉ��@��tNo.20409  <2002.10.11 yakimura> start
    
    Case "N" '|(center-side)/(center+side)|�~200
        
        C = WFCCenterP(WFOI_JUDG, O(), G)
        ORG = WF_TypeN_Exc(WFOI_JUDG, O(), G, C)
        
        If ORG = -9999 Then
              errflag = True
        End If

'�R�[�h"N"�ɑΉ��@��tNo.20409  <2002.10.11 yakimura> end
    
    Case " "

'''''        WFCJudgDialog.WFCErrorMessage "���z�v�Z����` A �ɂČv�Z"
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
'''''        WFCJudgDialog.WFCErrorMessage "���z�v�Z����`"
    ElseIf deverrflag Then
'''''        WFCJudgDialog.WFCErrorMessage "���z�v�Z 0 ���Z�G���["
    ElseIf ORG = -1 Then
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu�A�Ώۃf�[�^�A���z�v�Z����"
    End If
    
    '2002/07/24 Update T.Hayashi
    'WFCORGCal = ORG
    'WFCORGCal = RoundUp(ORG, 2)
    WFCORGCal_NotRound = ORG        '�ۂ߂Ȃ��l��Ԃ��悤�ɕύX 2011/11/25 SETsw kubota
    
End Function

'�[���ۂ߂��s���֐��ƍs��Ȃ��֐��̓�ɕ��� 2011/11/25 SETsw kubota
Public Function WFCORGCal(O() As Double, G As Guarantee, calcode As String) As Double
    '�[���ۂ߂��s��Ȃ��֐����Ăяo���A����2��(3���ڐ؂�グ)�ɂ��ĕԂ�
    WFCORGCal = RoundUp(WFCORGCal_NotRound(O(), G, calcode), 2)
End Function


Private Function WFCJudgDataSelect_All(d() As Double, G As Guarantee, T() As Double) As FUNCTION_RETURN
    Dim Func_Ans As FUNCTION_RETURN
    
    Func_Ans = FUNCTION_RETURN_FAILURE
    
    Select Case G.cPos
    Case "1"                                  '2003/05/15 �ǉ��@osawa �˗�No.030130
        ReDim T(3) As Double
        T(0) = d(0)
        T(1) = d(2)
        T(2) = d(6)
        T(3) = d(9)
        Func_Ans = FUNCTION_RETURN_SUCCESS    '�˗�No.030130�@�ǉ������܂�
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
