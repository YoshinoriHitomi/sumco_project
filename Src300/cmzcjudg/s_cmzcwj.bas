Attribute VB_Name = "s_cmzcwj"
Option Explicit

''WF�Z���^�[��Oi����\����
Type W_DOI
    GuaranteeDoi    As Guarantee    ''�i���ۏ؏��\����
    SpecDoiMin      As Double       ''�iWF�_�f�͏o1�`3����
    SpecDoiMax      As Double       ''�iWF�_�f�͏o1�`3���
    Doi(5)          As Double       ''��Oi����l
'���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    DoiAntnp        As Double      ''�`�m���x��Oi����l
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    JudgDoi         As Boolean      ''��Oi���茋��
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    '�`�F�b�N�pAN���x��ǉ�
    JudgAntnp       As Boolean      ''�`�m���x���茋��
    Antnp           As Integer      ''�i�v�e�`�m���x
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
End Type

''WF�Z���^�[AOi����\���́@03/12/09 ooba
Type W_AOI
    GuaranteeAoi    As Guarantee    ''�i���ۏ؏��\����
    SpecAoiMin      As Double       ''�iWF�c���_�f����
    SpecAoiMax      As Double       ''�iWF�c���_�f���
    AOI(2)          As Double       ''AOi����l
'���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    AoiAntnp        As Double      ''�`�m���xAOi����l
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    JudgAoi         As Boolean      ''AOi���茋��
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    '�`�F�b�N�pAN���x��ǉ�
    JudgAntnp       As Boolean      ''�`�m���x���茋��
    Antnp           As Integer      ''�i�v�e�`�m���x
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
End Type

''WF�Z���^�[OSF����\����
Type W_OSF
    GuaranteeOsf    As Guarantee    ''�i���ۏ؏��\����
    SpecOsfAveMax   As Double       ''�iWFOSF���Ϗ��
    SpecOsfMax      As Double       ''�iWFOSF���
    OSF(4)          As Double       ''OSF����l
'���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    OsfAntnp        As Double      ''AN���xOSF����l
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    OSFp(2)         As String * 1   ''OSF�p�^�[�����с@2003/05/17 ooba
    Min             As Double       ''�ŏ��l
    max             As Double       ''�ő�l
    AVE             As Double       ''���ϒl
    JudgOsf         As Boolean      ''OSF���茋��
    JudgDataMin     As Double       ''�ŏ�����l
    JudgDataMax     As Double       ''�ő唻��l
    JudgDataAve     As Double       ''���ϔ���l
    JudgDataPTK     As String * 1   ''�p�^�[���敪�@2003/05/17 ooba
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    '�`�F�b�N�pAN���x��ǉ�
    JudgAntnp       As Boolean      ''�`�m���x���茋��
    Antnp           As Integer      ''�i�v�e�`�m���x
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
End Type

''WF�Z���^�[BMD����\����
Type W_BMD
    GuaranteeBmd    As Guarantee    ''�i���ۏ؏��\����
    SpecBmdAveMin   As Double       ''�iWFBMD���ω���
    SpecBmdAveMax   As Double       ''�iWFBMD���Ϗ��
    SpecBmdGsAveMin   As Double     ''BMD���ω���(�O��)�@09/05/07 ooba
    SpecBmdGsAveMax   As Double     ''BMD���Ϗ��(�O��)�@09/05/07 ooba
    SpecBmdMBP      As Double       ''�iWFBMD�ʓ����z�@2003/05/20 ooba
    SpecBmdMCL      As String * 2   ''�iWFBMD�ʓ��v�Z�@2003/05/20 ooba
'    BMD(3)          As Double       ''BMD����l
    BMD(4)          As Double       ''BMD����l�@2003/05/20 ooba�@5�_�Ή�
'���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    'AN���x����ǉ�
    BmdAntnp        As Double      ''AN���xBMD����l
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    Min             As Double       ''�ŏ��l
    max             As Double       ''�ő�l
    AVE             As Double       ''���ϒl
    JudgBmd         As Boolean      ''BMD���茋��
    JudgDataMin     As Double       ''�ŏ�����l
    JudgDataMax     As Double       ''�ő唻��l
    JudgDataAve     As Double       ''���ϔ���l
    JudgDataMBP     As Double       ''�ʓ����z����l�@2003/05/20 ooba
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    '�`�F�b�N�pAN���x��ǉ�
    JudgAntnp       As Boolean      ''�`�m���x���茋��
    Antnp           As Integer      ''�i�v�e�`�m���x
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
End Type

''WF�Z���^�[DZ����\����
Type W_DZ
    GuaranteeDz     As Guarantee    ''�i���ۏ؏��\����
    SpecDzMin       As Double       ''�iWF�����בw����
    SpecDzMax       As Double       ''�iWF�����בw���
    DZ(3)           As Double       ''DZ����l
    JudgDz          As Boolean      ''DZ���茋��
    JudgDataMin     As Double       ''�ŏ�����l
    JudgDataMax     As Double       ''�ő唻��l
    JudgDataAve     As Double       ''���ϔ���l
End Type

''WF�Z���^�[DSOD����\����
Type W_DSOD
    GuaranteeDsod   As Guarantee    ''�i���ۏ؏��\����
    SpecDsodMin     As Double       ''�iWFDSOD����
    SpecDsodMax     As Double       ''�iWFDSOD���
    Dsod            As Double       ''DSOD����l
    Dsodp(1)        As String * 3   ''DSOD�p�^�[�����с@04/07/23 ooba
    JudgDataPTK     As String * 1   ''DSOD�p�^�[���敪�@04/07/23 ooba
    JudgDsod        As Boolean      ''DSOD���茋��
    DsodAntnp       As Double       ''AN���xDSOD����l  06/12/22 ooba
    JudgAntnp       As Boolean      ''AN���x���茋��    06/12/22 ooba
    Antnp           As Integer      ''�iWFAN���x        06/12/22 ooba
End Type

''WF�Z���^�[SPV����\����
Type W_SPV
    GuaranteeSpv    As Guarantee    ''�i���ۏ؏��\����
    GuaranteeSpvFe  As Guarantee    ''�i���ۏ؏��\����
    SpecSpvMin      As Double       ''�iWF�g�U������
    SpecSpvMax      As Double       ''�iWF�g�U�����
    SpecSpvFeMax    As Double       ''�iWFFe�Z�x���
    Spv(5)          As Double       ''SPV����l
    JudgSpv         As Boolean      ''SPV���茋��
    '-----TEST2004/10
    SpecSpvAvMax      As Double       ''�iWF���Ϗ��
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
    GuaranteeSpvNr  As Guarantee    ''�i���ۏ؏��\����
    SpecSpvNrMax    As Double       ''�iWFNr�Z�x���
    SpecSpvNrAvMax  As Double       ''�iWFNr���Ϗ��
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
End Type

''WF�Z���^�[GD����\����
Public Type W_GD
    GuaranteeDen        As Guarantee    ''�i���ۏ؏��\����
    GuaranteeLdl        As Guarantee    ''�i���ۏ؏��\����
    GuaranteeDvd2       As Guarantee    ''�i���ۏ؏��\����
    JudgFlagDen         As String * 1   ''�iWFDen�����L��
    JudgFlagLdl         As String * 1   ''�iWFL/DL�����L��
    JudgFlagDvd2        As String * 1   ''�iWFDVD2�����L��
    SpecDenMin          As Double       ''�iWFDen����
    SpecDenMax          As Double       ''�iWFDen���
    SpecLdlMin          As Double       ''�iWFL/DL����
    SpecLdlMax          As Double       ''�iWFL/DL���
    SpecDvd2Min         As Double       ''�iWFDVD2����
    SpecDvd2Max         As Double       ''�iWFDVD2���
'*** UPDATE �� Y.SIMIZU 2005/10/1 �iWFGDײݐ�
    SpecGdLine          As Single       ''�iWFGDײݐ�
'*** UPDATE �� Y.SIMIZU 2005/10/1 �iWFGDײݐ�
    Den                 As Double       ''Den�v�Z�l
    Ldl                 As Double       ''L/DL�v�Z�l
    Dvd2                As Double       ''DVD2�v�Z�l
    JudgDen             As Boolean      ''Den���茋��
    JudgLdl             As Boolean      ''L/DL���茋��
    JudgDvd2            As Boolean      ''DVD2���茋��
    GdAntnp             As Double       ''AN���xGD����l    06/12/22 ooba
    JudgAntnp           As Boolean      ''AN���x���茋��    06/12/22 ooba
    Antnp               As Integer      ''�iWFAN���x        06/12/22 ooba
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    GDPTK               As String * 1   ''�i�v�e�f�c�p�^���敪
    LdlMin              As Integer      ''L/DL�A��0MIN
    LdlMax              As Integer      ''L/DL�A��0MAX
    ZeroLdlMin          As Integer      ''�iSXLdl�A��0����
    ZeroLdlMax          As Integer      ''�iSXLdl�A��0���
    JudgLdlPtn          As Boolean      ''L/DL�p�^�[�����茋��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
End Type

''��Add 2010/01/07 SIRD�Ή� Y.Hitomi
''WF�Z���^�[SIRD����\����
Type W_SD
    GuaranteeSd         As Guarantee    ''�i���ۏ؏��\����
    SpecSdMax           As Integer      ''����]��(SIRD)���
    SdMeasData          As Integer      ''SIRD���茋��
    JudgSD              As Boolean      ''SIRD���茋��
End Type
''��Add 2010/01/07 SIRD�Ή� Y.Hitomi

'�T�v      :WF�Z���^�[��Oi������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Doi           ,I  ,W_DOI            ,WF�Z���^�[��Oi����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WfDOiJudg(Doi As W_DOI, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFDOI_JUDG = 3                 ''���莯�ʃt���O(��Oi)
    Dim FuncAns As FUNCTION_RETURN
    Dim TempDOi(2) As Double
    Dim JData(2) As Double
    Dim c0 As Integer
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    Dim liRet           As Integer
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Doi.JudgDoi = JUDG_NG
    If Doi.GuaranteeDoi.cJudg = JudgCodeW01 Then ''DOi����L��
        
''''''        WFCJudgDialog.WFCErrorMessage " "
''''''        WFCJudgDialog.WFCErrorMessage "��Oi���� **************"
''''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�_ = " & Doi.GuaranteeDoi.cCount
''''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Doi.GuaranteeDoi.cMeth
''''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Doi.GuaranteeDoi.cPos
''''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Doi.GuaranteeDoi.cObj
''''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Doi.GuaranteeDoi.cJudg
        
        ''��Oi = Initial_Oi - After_Oi
        ''���S����O���֕��בւ�
        For c0 = 0 To 2
            TempDOi(c0) = Doi.Doi(2 - c0) - Doi.Doi(5 - c0)
        Next
        
        ''DOi����
        FuncAns = GetWfJudgData(WFDOI_JUDG, Doi.GuaranteeDoi, TempDOi(), JData())
        If (InStr(ObjCodeGrp01, Doi.GuaranteeDoi.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
            Select Case Doi.GuaranteeDoi.cObj
            Case ObjCode01, ObjCode02  ''���S1�_�A�����l
                Doi.JudgDoi = RangeDecision_nl(JData(0), Doi.SpecDoiMin, Doi.SpecDoiMax)
            Case ObjCode03 ''�S��
                Doi.JudgDoi = JUDG_OK
                For c0 = 0 To 2
                    If JData(c0) <> -1 Then
                        If RangeDecision_nl(JData(c0), Doi.SpecDoiMin, Doi.SpecDoiMax) = JUDG_NG Then
                            Doi.JudgDoi = JUDG_NG
                        End If
                    End If
                Next
            Case ObjCode04 ''R/2''''''''''''''''''''''''''''''''''''''''''''''''''''''''''����
'''''                WFCJudgDialog.WFCErrorMessage "��Oi����A�Ώۃf�[�^�����B"
            End Select
        Else
            ''�_���A�K�i�����ȊO�̏ꍇ
            If (Doi.GuaranteeDoi.cObj <> ObjCode13) And (Doi.GuaranteeDoi.cObj <> ObjCode15) Then
                ''�Ώۃf�[�^����
                ''�G���[���\���̂ɏ������B
'''''                WFCJudgDialog.WFCErrorMessage "��Oi����A�Ώۃf�[�^�����B"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, DOI_JUDG, Doi.GuaranteeDoi.cObj)
            End If
        End If
    
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
        'AN���x�`�F�b�N��ǉ�
        ''AN���x����
        '�}�g���b�N�X����`�F�b�N�̐��ۂ��擾
        liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(Doi.Antnp), CStr(Doi.DoiAntnp))
        If liRet = -1 Then
            FuncAns = SetErrInfo(ErrInfo, EZJ00, DOI_JUDG, Doi.GuaranteeDoi.cObj)
        ElseIf liRet = 0 Then
            Doi.JudgAntnp = JUDG_NG
        Else
            Doi.JudgAntnp = JUDG_OK
        End If
        
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    Else
        Doi.JudgDoi = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -s-
        Doi.JudgAntnp = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -e-
'        If InStr(JudgCodeW02, Doi.GuaranteeDoi.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, DOI_JUDG, Doi.GuaranteeDoi.cJudg)
'        End If
    End If
    
    WfDOiJudg = FuncAns
End Function

'�T�v      :WF�Z���^�[AOi������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Aoi           ,I  ,W_AOI            ,WF�Z���^�[AOi����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :03/12/09 ooba

Public Function WfAOiJudg(AOI As W_AOI, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN

    Dim FuncAns As FUNCTION_RETURN
    Dim TempAOi(2) As Double
    Dim JData(2) As Double
    Dim c0 As Integer
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    Dim liRet As Integer
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    AOI.JudgAoi = JUDG_NG
    If AOI.GuaranteeAoi.cJudg = JudgCodeW01 Then ''AOi����L��
        
        ''��Oi = Initial_Oi - After_Oi
        ''���S����O���֕��בւ�
        For c0 = 0 To 2
'''            TempDOi(c0) = Doi.Doi(2 - c0) - Doi.Doi(5 - c0)
            TempAOi(c0) = AOI.AOI(2 - c0)
        Next
        
        ''AOi����
        FuncAns = GetWfJudgData(WFAOI_JUDG, AOI.GuaranteeAoi, TempAOi(), JData())
        If (InStr(ObjCodeGrp01, AOI.GuaranteeAoi.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
            Select Case AOI.GuaranteeAoi.cObj
            Case ObjCode01, ObjCode02  ''���S1�_�A�����l
                AOI.JudgAoi = RangeDecision_nl(JData(0), AOI.SpecAoiMin, AOI.SpecAoiMax)
            Case ObjCode03 ''�S��
                AOI.JudgAoi = JUDG_OK
                For c0 = 0 To 2
                    If JData(c0) <> -1 Then
                        If RangeDecision_nl(JData(c0), AOI.SpecAoiMin, AOI.SpecAoiMax) = JUDG_NG Then
                            AOI.JudgAoi = JUDG_NG
                        End If
                    End If
                Next
            Case ObjCode04 ''R/2''''''''''''''''''''''''''''''''''''''''''''''''''''''''''����
'''''                WFCJudgDialog.WFCErrorMessage "��Oi����A�Ώۃf�[�^�����B"
            End Select
        Else
            ''�_���A�K�i�����ȊO�̏ꍇ
            If (AOI.GuaranteeAoi.cObj <> ObjCode13) And (AOI.GuaranteeAoi.cObj <> ObjCode15) Then
                ''�Ώۃf�[�^����
                ''�G���[���\���̂ɏ������B
'''''                WFCJudgDialog.WFCErrorMessage "��Oi����A�Ώۃf�[�^�����B"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, AOI_JUDG, AOI.GuaranteeAoi.cObj)
            End If
        End If
    
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
        ''AN���x����
        liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(AOI.Antnp), CStr(AOI.AoiAntnp))
        If liRet = -1 Then
            FuncAns = SetErrInfo(ErrInfo, EZJ00, AOI_JUDG, AOI.GuaranteeAoi.cObj)
        ElseIf liRet = 0 Then
            AOI.JudgAntnp = JUDG_NG
        Else
            AOI.JudgAntnp = JUDG_OK
        End If
                
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    Else
        AOI.JudgAoi = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -s-
        AOI.JudgAntnp = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -e-
    End If
    
    WfAOiJudg = FuncAns
End Function

'�T�v      :WF�Z���^�[OSF������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Osf           ,I  ,W_OSF            ,WF�Z���^�[OSF����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WfOSFJudg(OSF As W_OSF, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFOSF_JUDG = 4                 ''���莯�ʃt���O(OSF)
    Dim FuncAns As FUNCTION_RETURN
    Dim JData(4) As Double
    Dim c0 As Integer
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    Dim liRet As Integer
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '' 2006/09/19 SMP)kondoh Add -s-
    Dim JudgOsfTmp As Boolean
    JudgOsfTmp = JUDG_OK
    '' 2006/09/19 SMP)kondoh Add -e-
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    OSF.JudgOsf = JUDG_NG
    
    OSF.JudgDataAve = JudgAve(OSF.OSF())
    OSF.JudgDataMax = JudgMax(OSF.OSF())
    OSF.JudgDataMin = JudgMin(OSF.OSF())
    
    If OSF.GuaranteeOsf.cJudg = JudgCodeW01 Then ''OSF����L��
        
''''''        WFCJudgDialog.WFCErrorMessage " "
''''''        WFCJudgDialog.WFCErrorMessage "OSF���� ***************"
''''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�_ = " & OSF.GuaranteeOsf.cCount
''''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & OSF.GuaranteeOsf.cMeth
''''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & OSF.GuaranteeOsf.cPos
''''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & OSF.GuaranteeOsf.cObj
''''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & OSF.GuaranteeOsf.cJudg
        
        ''OSF����
        FuncAns = GetWfJudgData(WFOSF_JUDG, OSF.GuaranteeOsf, OSF.OSF(), JData())
'        If (InStr(ObjCodeGrp03, OSF.GuaranteeOsf.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
        If (InStr(ObjCodeGrp03 & ObjCode08 & ObjCode09, OSF.GuaranteeOsf.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
            Select Case OSF.GuaranteeOsf.cObj
            Case ObjCode05, ObjCode08  ''�S�_�̕��ϒl�A�S�_�̍ŏ��l
                OSF.JudgOsf = RangeDecision_nl(JData(0), 0, OSF.SpecOsfAveMax)
            Case ObjCode06  ''�S�_�̍ő�l
                OSF.JudgOsf = RangeDecision_nl(JData(0), 0, OSF.SpecOsfMax)
            Case ObjCode07 ''�S�_�̕��ϒl�ƍő�l
                '' 2006/09/19 SMP)kondoh Cng -s-
''                If RangeDecision_nl(JData(1), 0, OSF.SpecOsfAveMax) Then
                If RangeDecision_nl(JData(0), 0, OSF.SpecOsfAveMax) Then
                '' 2006/09/19 SMP)kondoh Cng -e-
                    OSF.JudgOsf = RangeDecision_nl(JData(1), 0, OSF.SpecOsfMax)
                Else
                    OSF.JudgOsf = JUDG_NG
                End If
            Case ObjCode09 ''������2�_�A�O����2�_(5�_�����1,2,4,5)
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
            ''�_���A�K�i�����ȊO�̏ꍇ
            If (OSF.GuaranteeOsf.cObj <> ObjCode13) And (OSF.GuaranteeOsf.cObj <> ObjCode15) Then
                ''�Ώۃf�[�^����
                ''�G���[���\���̂ɏ������B
'''''                WFCJudgDialog.WFCErrorMessage "OSF����A�Ώۃf�[�^�����B"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, OSF_JUDG, OSF.GuaranteeOsf.cObj)
            End If
        End If
        With OSF
        If .JudgOsf Then    '�p�^�[���̔���@2003/05/17 ooba
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
    
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
        ''AN���x����
        liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(OSF.Antnp), CStr(OSF.OsfAntnp))
        If liRet = -1 Then
            FuncAns = FUNCTION_RETURN_FAILURE
        ElseIf liRet = 0 Then
            OSF.JudgAntnp = JUDG_NG
        Else
            OSF.JudgAntnp = JUDG_OK
        End If
        
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------

    Else
        OSF.JudgOsf = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -s-
        OSF.JudgAntnp = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -e-
'        If InStr(JudgCodeW02, Osf.GuaranteeOsf.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, OSF_JUDG, Osf.GuaranteeOsf.cJudg)
'        End If
    End If
    
    WfOSFJudg = FuncAns
End Function

'�T�v      :WF�Z���^�[BMD������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Bmd           ,I  ,W_BMD            ,WF�Z���^�[BMD����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :Bno           ,I  ,Integer          ,BMDno(1:BMD1,2:BMD2,3:BMD3)(��ߗp)�@09/05/07 ooba
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WfBMDJudg(BMD As W_BMD, ErrInfo As ERROR_INFOMATION, _
                                                Optional Bno As Integer = 0) As FUNCTION_RETURN
'WFBMD_JUDG = 5                 ''���莯�ʃt���O(BMD)
    Dim FuncAns As FUNCTION_RETURN
    Dim TempBmd(3) As Double
    Dim JData(3) As Double
    Dim c0 As Integer
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    Dim liRet As Integer
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    BMD.JudgBmd = JUDG_NG
    
    BMD.JudgDataAve = JudgAve(BMD.BMD())
    BMD.JudgDataMax = JudgMax(BMD.BMD())
    BMD.JudgDataMin = JudgMin(BMD.BMD())
    If BMD.SpecBmdMCL = "P " Then                '�ʓ����z�̌v�Z�@2003/05/20 ooba
        BMD.JudgDataMBP = JudgBmdMBP(BMD.BMD())
    Else
        BMD.JudgDataMBP = 0                      '�ʓ����z��"P"�ȊO�̎��͌v�Z���ʂ�0�Ƃ���@2003/06/06 ooba
    End If
    
    If BMD.GuaranteeBmd.cJudg = JudgCodeW01 Then ''BMD����L��
        
''''''        WFCJudgDialog.WFCErrorMessage " "
''''''        WFCJudgDialog.WFCErrorMessage "BMD���� ***************"
''''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�_ = " & BMD.GuaranteeBmd.cCount
''''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & BMD.GuaranteeBmd.cMeth
''''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & BMD.GuaranteeBmd.cPos
''''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & BMD.GuaranteeBmd.cObj
''''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & BMD.GuaranteeBmd.cJudg
        
        ''BMD����
'--- 2006/08/15 Del �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
'        FuncAns = GetWfJudgData(WFBMD_JUDG, BMD.GuaranteeBmd, BMD.BMD(), JData())
'--- 2006/08/15 Del �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
        If (InStr(ObjCodeGrp02, BMD.GuaranteeBmd.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
            FuncAns = GetWfJudgData(WFBMD_JUDG, BMD.GuaranteeBmd, BMD.BMD(), JData())
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
            Select Case BMD.GuaranteeBmd.cObj
            ''�S�_�̕��ϒl�A�S�_�̍ő�l�A�S�_�̍ŏ��l�AMAX(2,4�_��)�AMAX(2,3,4�_��)
            Case ObjCode05, ObjCode06, ObjCode08, ObjCode10, ObjCode11
                BMD.JudgBmd = RangeDecision_nl(JData(0), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
            Case ObjCode07 ''�S�_�̕��ϒl�ƍő�l
                If RangeDecision_nl(JData(0), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax) Then
                    BMD.JudgBmd = RangeDecision_nl(JData(1), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
                Else
                    BMD.JudgBmd = JUDG_NG
                End If
            '----TEST2004/10
            Case ObjCode16 ''�S�_�̍ő�l�ƍŏ��l
                If RangeDecision_nl(JData(0), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax) Then
                    BMD.JudgBmd = RangeDecision_nl(JData(1), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
                Else
                    BMD.JudgBmd = JUDG_NG
                End If
'' 2008/10/20 BMD�]��,�O��1�_�ۏ؋@�\�ǉ� ADD By Systech Start
            Case ObjCode18  ''AVE+�O��1�_
'                If RangeDecision_nl(JData(0), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax) Then
                '�O��1�_������@�ύX�@09/05/07 ooba
                If Bno = 3 Then
                    If RangeDecision_nl(JData(0), BMD.SpecBmdGsAveMin, BMD.SpecBmdGsAveMax) Then
                        BMD.JudgBmd = RangeDecision_nl(JData(1), BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
                    Else
                        BMD.JudgBmd = JUDG_NG
                    End If
                Else
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, BMD_JUDG, BMD.GuaranteeBmd.cObj)
                End If
'' 2008/10/20 BMD�]��,�O��1�_�ۏ؋@�\�ǉ� ADD By Systech End
            End Select
        Else
            ''�_���A�K�i�����ȊO�̏ꍇ
            If (BMD.GuaranteeBmd.cObj <> ObjCode13) And (BMD.GuaranteeBmd.cObj <> ObjCode15) Then
                ''�Ώۃf�[�^����
                ''�G���[���\���̂ɏ������B
'''''                WFCJudgDialog.WFCErrorMessage "OSF����A�Ώۃf�[�^�����B"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, BMD_JUDG, BMD.GuaranteeBmd.cObj)
            End If
        End If
        '�ʓ����z�ł̔���@2003/05/20 ooba
        If BMD.SpecBmdMCL = "P " Then      '�d�l��P�̎��̂ݔ�����s���A����ȊO�͔�����s�킸OK�Ƃ���
'            If BMD.SpecBmdMBP = -1 Then
'                FuncAns = FUNCTION_RETURN_FAILURE
            If BMD.SpecBmdMBP <> 0 Or BMD.SpecBmdMBP = -1 Then      '�d�l�̖ʓ����z��0,-1(NULL)�̎��͔�����s�킸OK�Ƃ���
                If BMD.JudgDataMBP = -1 Then
                    FuncAns = FUNCTION_RETURN_FAILURE
                Else
                    If BMD.JudgBmd Then
                        BMD.JudgBmd = RangeDecision_nl(BMD.JudgDataMBP, 0, BMD.SpecBmdMBP)
                    End If
                End If
            End If
        End If
    
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
        ''AN���x����
        liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(BMD.Antnp), CStr(BMD.BmdAntnp))
        If liRet = -1 Then
            FuncAns = FUNCTION_RETURN_FAILURE
        ElseIf liRet = 0 Then
            BMD.JudgAntnp = JUDG_NG
        Else
            BMD.JudgAntnp = JUDG_OK
        End If
        
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------

    Else
        BMD.JudgBmd = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -s-
        BMD.JudgAntnp = JUDG_OK
        '' 2006/09/19 SMP)kondoh Add -e-
'        If InStr(JudgCodeW02, Bmd.GuaranteeBmd.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, BMD_JUDG, Bmd.GuaranteeBmd.cJudg)
'        End If
    End If
    
    WfBMDJudg = FuncAns
End Function

'�T�v      :WF�Z���^�[DZ������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Dz            ,I  ,W_DZ             ,WF�Z���^�[DZ����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WfDZJudg(DZ As W_DZ, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFDZ_JUDG = 6                  ''���莯�ʃt���O(DZ)
    Dim FuncAns As FUNCTION_RETURN
    Dim JData(3) As Double
    Dim c0 As Integer
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    DZ.JudgDz = JUDG_NG
    
    DZ.JudgDataAve = JudgAve(DZ.DZ())
    DZ.JudgDataMax = JudgMax(DZ.DZ())
    DZ.JudgDataMin = JudgMin(DZ.DZ())
    
    If DZ.GuaranteeDz.cJudg = JudgCodeW01 Then ''DZ����L��
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "DZ���� ****************"
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�_ = " & Dz.GuaranteeDz.cCount
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Dz.GuaranteeDz.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Dz.GuaranteeDz.cPos
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Dz.GuaranteeDz.cObj
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Dz.GuaranteeDz.cJudg
        
        ''DZ����
        FuncAns = GetWfJudgData(WFDZ_JUDG, DZ.GuaranteeDz, DZ.DZ(), JData())
        If (InStr(ObjCodeGrp06, DZ.GuaranteeDz.cObj) <> 0) And (FuncAns = FUNCTION_RETURN_SUCCESS) Then
            Select Case DZ.GuaranteeDz.cObj
            ''�S�_�̕��ϒl�A�S�_�̍ő�l�A�S�_�̍ŏ��l�AMAX(2,4�_��)�AMAX(2,3,4�_��)
            Case ObjCode05, ObjCode06, ObjCode08, ObjCode10, ObjCode11
                DZ.JudgDz = RangeDecision_nl(JData(0), DZ.SpecDzMin, DZ.SpecDzMax)
            Case ObjCode07 ''�S�_�̕��ϒl�ƍő�l
                If RangeDecision_nl(JData(0), DZ.SpecDzMin, DZ.SpecDzMax) Then
                    DZ.JudgDz = RangeDecision_nl(JData(1), DZ.SpecDzMin, DZ.SpecDzMax)
                Else
                    DZ.JudgDz = JUDG_NG
                End If
            
            '����V�X�e���Ƌ��ʊ֐����킹�ŏC���@hama 2004/11/30 start
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
           '����V�X�e���Ƌ��ʊ֐����킹�ŏC���@hama 2004/11/30 end
        Else
            ''�_���A�K�i�����ȊO�̏ꍇ
            If (DZ.GuaranteeDz.cObj <> ObjCode13) And (DZ.GuaranteeDz.cObj <> ObjCode15) Then
                ''�Ώۃf�[�^����
                ''�G���[���\���̂ɏ������B
'''''                WFCJudgDialog.WFCErrorMessage "DZ����A�Ώۃf�[�^�����B"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, DZ_JUDG, DZ.GuaranteeDz.cObj)
            End If
        End If
    Else
        DZ.JudgDz = JUDG_OK
'        If InStr(JudgCodeW02, Dz.GuaranteeDz.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, DZ_JUDG, Dz.GuaranteeDz.cJudg)
'        End If
    End If
    
    WfDZJudg = FuncAns
End Function

'�T�v      :WF�Z���^�[DSOD������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Dsod          ,I  ,W_DSOD           ,WF�Z���^�[DSOD����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WfDSODJudg(Dsod As W_DSOD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFDSOD_JUDG = 7                ''���莯�ʃt���O(DSOD)
    Dim FuncAns As FUNCTION_RETURN
    Dim liRet As Integer        '06/12/22 ooba
    Dim sResult As String       '06/12/22 ooba
    Dim RET As Integer          '06/12/22 ooba
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Dsod.JudgDsod = JUDG_NG
    Dsod.JudgAntnp = JUDG_OK        '06/12/22 ooba
    If Dsod.GuaranteeDsod.cJudg = JudgCodeW01 Then ''DSOD����L��
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "DSOD���� **************"
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�_ = " & Dsod.GuaranteeDsod.cCount
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Dsod.GuaranteeDsod.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Dsod.GuaranteeDsod.cPos
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Dsod.GuaranteeDsod.cObj
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Dsod.GuaranteeDsod.cJudg
        
        If Dsod.GuaranteeDsod.cObj = ObjCodeGrp04 Then
            ''DSOD����
            Dsod.JudgDsod = RangeDecision_nl(Dsod.Dsod, Dsod.SpecDsodMin, Dsod.SpecDsodMax)
        Else
            ''�_���A�K�i�����ȊO�̏ꍇ
            If (Dsod.GuaranteeDsod.cObj <> ObjCode13) And (Dsod.GuaranteeDsod.cObj <> ObjCode15) Then
                ''�Ώۃf�[�^����
                ''�G���[���\���̂ɏ������B
'''''                WFCJudgDialog.WFCErrorMessage "DSOD����A�Ώۃf�[�^�����B"
                FuncAns = SetErrInfo(ErrInfo, EZJ00, DSOD_JUDG, Dsod.GuaranteeDsod.cObj)
            End If
        End If
        'DSOD����ݔ���ǉ��@04/07/28 ooba START ================================>
        If Dsod.JudgDsod = JUDG_OK Then
            Dsod.JudgDsod = JudgDsodPattern(Dsod.JudgDataPTK, Dsod.Dsodp())
        End If
        'DSOD����ݔ���ǉ��@04/07/28 ooba END ==================================>
        
        'GD/DSOD�M���������ǉ��@06/12/22 ooba START =========================================>
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
        'GD/DSOD�M���������ǉ��@06/12/22 ooba END ===========================================>
    Else
        Dsod.JudgDsod = JUDG_OK
'        If InStr(JudgCodeW02, Dsod.GuaranteeDsod.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, EZJ00, DSOD_JUDG, Dsod.GuaranteeDsod.cJudg)
'        End If
    End If
    
    WfDSODJudg = FuncAns
End Function

'�T�v      :WF�Z���^�[SPV������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Spv           ,I  ,W_SPV            ,WF�Z���^�[SPV����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WfSPVJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
'WFSPV_JUDG = 8                 ''���莯�ʃt���O(SPV)
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim j1 As Boolean
    Dim j2 As Boolean
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
    j1 = JUDG_NG
    j2 = JUDG_NG
    
    If Spv.GuaranteeSpv.cJudg = JudgCodeW01 Then ''SPV����L��
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "SPV���� ***************"
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�_ = " & Spv.GuaranteeSpv.cCount
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Spv.GuaranteeSpv.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Spv.GuaranteeSpv.cPos
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Spv.GuaranteeSpv.cObj
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Spv.GuaranteeSpv.cJudg
        
        ''SPV����
        '-----TEST2004/10
        'If Spv.GuaranteeSpv.cObj = ObjCode03 Then
        Select Case Spv.GuaranteeSpv.cObj
            Case ObjCode03  '�S����_=3
                ''�g�U��AVE���K�i�l�͈͓��Ȃ�OK
                'j1 = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax)
                ''MAX,AVE,MIN���S�ċK�i�l�͈͓��Ȃ�OK�ɕύX
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    If RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                        j1 = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
                    Else
                        j1 = JUDG_NG
                    End If
                Else
                    j1 = JUDG_NG
                End If
            Case ObjCode05 '�`�u�d=A
                ''AVE���K�i�l�͈͓��Ȃ�OK
                j1 = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax)
            Case ObjCode06 '�l�`�w=B
                ''MAX���K�i�l�͈͓��Ȃ�OK
                j1 = RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax)
            Case ObjCode07 '�`�u�d+�l�`�w=C
                ''MAX,AVE���K�i�l�͈͓��Ȃ�OK
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    j1 = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, Spv.SpecSpvMax)
                Else
                    j1 = JUDG_NG
                End If
            Case ObjCode08 '�l�h�m=D
                ''MIN���K�i�l�͈͓��Ȃ�OK
                j1 = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
            Case ObjCode16 '�l�h�m+�l�`�w=K
                ''MAX,MIN���K�i�l�͈͓��Ȃ�OK
                If RangeDecision_nl(Spv.Spv(0), Spv.SpecSpvMin, Spv.SpecSpvMax) Then
                    j1 = RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMin, Spv.SpecSpvMax)
                Else
                    j1 = JUDG_NG
                End If
            Case Else
                ''�_���A�K�i�����ȊO�̏ꍇ
                If (Spv.GuaranteeSpv.cObj <> ObjCode13) And (Spv.GuaranteeSpv.cObj <> ObjCode15) Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
    '''''                WFCJudgDialog.WFCErrorMessage "SPV����A�Ώۃf�[�^�����B"
                    FuncAns = SetErrInfo(ErrInfo, EZJ00, SPV_JUDG, Spv.GuaranteeSpv.cObj)
                    GoTo EXIT_FUNC
                End If
                j1 = JUDG_OK
        End Select

    Else
        j1 = JUDG_OK
    End If
    
    If Spv.GuaranteeSpvFe.cJudg = JudgCodeW01 Then ''SPVFE����L��
        
'''''        WFCJudgDialog.WFCErrorMessage " "
'''''        WFCJudgDialog.WFCErrorMessage "SPV���� ***************"
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�_ = " & Spv.GuaranteeSpvFe.cCount
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Spv.GuaranteeSpvFe.cMeth
'''''        WFCJudgDialog.WFCErrorMessage "����ʒu_�� = " & Spv.GuaranteeSpvFe.cPos
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Spv.GuaranteeSpvFe.cObj
'''''        WFCJudgDialog.WFCErrorMessage "�ۏؕ��@_�� = " & Spv.GuaranteeSpvFe.cJudg
        
        ''SPVFE����
        '----TEST2004/10
        'If Spv.GuaranteeSpvFe.cObj = ObjCode03 Then
        Select Case Spv.GuaranteeSpvFe.cObj
        Case ObjCode03, ObjCode06 '�S����_�A�l�`�w
            j2 = RangeDecision_nl(Spv.Spv(3), 0, Spv.SpecSpvFeMax)
        Case ObjCode05 '�`�u�d
            j2 = RangeDecision_nl(Spv.Spv(4), 0, Spv.SpecSpvAvMax)
        Case ObjCode07 '�`�u�d+�l�h�m
            If RangeDecision_nl(Spv.Spv(4), 0, Spv.SpecSpvAvMax) Then
                j2 = RangeDecision_nl(Spv.Spv(3), 0, Spv.SpecSpvFeMax)
            Else
                j2 = JUDG_NG
            End If
        Case Else  '
            ''�_���A�K�i�����ȊO�̏ꍇ
            If (Spv.GuaranteeSpvFe.cObj <> ObjCode13) And (Spv.GuaranteeSpvFe.cObj <> ObjCode15) Then
                ''�Ώۃf�[�^����
                ''�G���[���\���̂ɏ������B
'''''                WFCJudgDialog.WFCErrorMessage "SPV����A�Ώۃf�[�^�����B"
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

'�T�v      :WF�Z���^�[GD������s���B
'���Ұ�    :�ϐ���        ,IO    ,�^                  ,����
'          :GD            ,I     ,W_GD                ,WF�Z���^�[GD����\����
'          :sGDhsflg      ,I     ,String              ,�ۏ��׸�(1�FWF�ۏ�)�@06/12/22 ooba
'          :ErrInfo       ,O     ,ERROR_INFOMATION    ,�G���[���\����
'          :�߂�l        ,O     ,FUNCTION_RETURN     ,
'����      :
'����      :05/01/31 ooba
Public Function WfGdJudg(GD As W_GD, sGDhsflg As String, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN

    Dim FuncAns As FUNCTION_RETURN
    Dim liRet As Integer        '06/12/22 ooba
    Dim sResult As String       '06/12/22 ooba
    Dim RET As Integer          '06/12/22 ooba
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS

    ''Den�����L�����f
    GD.JudgDen = JUDG_OK
'    If GD.JudgFlagDen = "1" Then
        If GD.GuaranteeDen.cJudg = JudgCodeW01 Then ''Den���肠��
            GD.JudgDen = RangeDecision_nl(GD.Den, GD.SpecDenMin, GD.SpecDenMax)
        Else
            GD.JudgDen = JUDG_OK
        End If
'    End If

    ''L/DL�����L�����f
    GD.JudgLdl = JUDG_OK
'    If GD.JudgFlagLdl = "1" Then
        If GD.GuaranteeLdl.cJudg = JudgCodeW01 Then ''L/DL���肠��
            GD.JudgLdl = RangeDecision_nl(GD.Ldl, GD.SpecLdlMin, GD.SpecLdlMax)
        Else
            GD.JudgLdl = JUDG_OK
        End If
'    End If

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    GD.JudgLdlPtn = JUDG_OK
'    If WFGD.JudgFlagLdl = "1" Then
        If GD.GuaranteeLdl.cJudg = JudgCodeW01 Then ''L/DL���肠��
            If GD.GDPTK = "1" Then
                ' "0"�A����(MIN)�@���@�i__L/DL�A��0����(SX/WF)
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
                ' "0"�A����(MAX)�@���@�i__L/DL�A��0���(SX/WF)
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
                ' ���薳��
                
            End If
        End If
'    End If
    GD.JudgLdl = GD.JudgLdl And GD.JudgLdlPtn
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    
    ''DVD2�����L�����f
    GD.JudgDvd2 = JUDG_OK
'    If GD.JudgFlagDvd2 = "1" Then
        If GD.GuaranteeDvd2.cJudg = JudgCodeW01 Then ''Dvd2���肠��
            GD.JudgDvd2 = RangeDecision_nl(GD.Dvd2, GD.SpecDvd2Min, GD.SpecDvd2Max)
        Else
            GD.JudgDvd2 = JUDG_OK
        End If
'    End If
    
    'GD/DSOD�M���������ǉ��@06/12/22 ooba START =========================================>
    GD.JudgAntnp = JUDG_OK
    If sGDhsflg = "1" And _
       (GD.GuaranteeDen.cJudg = JudgCodeW01 Or _
        GD.GuaranteeLdl.cJudg = JudgCodeW01 Or _
        GD.GuaranteeDvd2.cJudg = JudgCodeW01) Then
        RET = 0
        sResult = ""
        'DEN-AN���x����
        RET = funCodeDBGet("SB", "15", "DEN", 0, " ", sResult)
        If RET = 0 And Mid(sResult, 16, 1) = "2" Then
            'LDL-AN���x����
            RET = funCodeDBGet("SB", "15", "LDL", 0, " ", sResult)
            If RET = 0 And Mid(sResult, 16, 1) = "2" Then
                'DVD-AN���x����
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
    'GD/DSOD�M���������ǉ��@06/12/22 ooba END ===========================================>
        
    WfGdJudg = FuncAns

End Function
'�T�v      :WF�Z���^�[SIRD(SD)������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Dz            ,I  ,W_DZ             ,WF�Z���^�[SIRD����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2010/01/07 Y.Hitomi
Public Function WfSDJudg(SD As W_SD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim JData(3) As Double
    Dim c0 As Integer
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
        
    If SD.GuaranteeSd.cJudg = JudgCodeW01 Then  ''SD����L��
        ''SD����
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

'�T�v      :�ΏۃR�[�h�ɏ]���Ĕ���Ώۃf�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Flag          ,I  ,GUARANTEE ,�ΏۃR�[�h
'          :d()           ,I  ,double    ,����l
'          :d1()          ,O  ,double    ,����Ώۃf�[�^
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :Flag.cObj�̒l,
'          :1       ,d1(0)=���S����l
'          :2       ,d1(0)=��������l
'          :3       ,d1()=�S����_
'          :4       ,d1(0)=R/2
'          :A       ,d1(0)=���ϒl
'          :B       ,d1(0)=�ő�l
'          :C       ,d1(0)=���ϒl,d1(1)=�ő�l
'          :D       ,d1(0)=�ŏ��l
'          :E       ,d1(0�`3)=������2�_�A�O����2�_(5�_�����1,2,4,5)
'          :F       ,d1(0)=2,4�_�ڂ̓��傫���l
'          :G       ,d1(0)=2,3,4�_�ڂ̓��傫���l
'          :K       ,d1(0)=�ő�l,d1(1)=�ŏ��l�@TEST2004/10
'����      :2001/06/06 ���� �M�� �쐬
Public Function GetWfJudgData(JudgFlag As Integer, flag As Guarantee, d() As Double, d1() As Double) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim COUNT As Integer
    Dim High As Integer
    
    '' �z��̏�����擾���܂��B
    High = UBound(d)
    
    FuncAns = FUNCTION_RETURN_SUCCESS '' ����
    Select Case flag.cObj
    Case ObjCode01 ''���S����l
        Select Case JudgFlag
        Case WFOI_JUDG  ''���莯�ʃt���O(Oi)
            d1(0) = d(0)
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            d1(0) = d(4)
        Case WFDOI_JUDG, WFAOI_JUDG ''���莯�ʃt���O(��Oi)�A���莯�ʃt���O(AOi) �ǉ� 03/12/09 ooba
        '' �擾�f�[�^�ύX�@2003/10/15 ooba
'            d1(0) = d(5)
            d1(0) = d(2)
        Case WFBMD_JUDG ''���莯�ʃt���O(BMD)
            d1(0) = d(3)
'        Case WFDZ_JUDG  ''���莯�ʃt���O(DZ)
'        Case WFDSOD_JUDG ''���莯�ʃt���O(DSOD)
        Case WFSPV_JUDG ''���莯�ʃt���O(SPV)
        Case WFOSF_JUDG ''���莯�ʃt���O(OSF)
        Case Else
'''''            WFCJudgDialog.WFCErrorMessage "�Ώۃf�[�^�����B"
            FuncAns = FUNCTION_RETURN_FAILURE '' �ُ�
        End Select
    Case ObjCode02 ''����l�̒����l
        d1(0) = JudgCenter(d())
    Case ObjCode03 ''�S����_
        DataCopy d(), d1()
    Case ObjCode04 ''R/2
        If InStr(PosCodeGrp01, flag.cPos) <> 0 Then
            Select Case JudgFlag
            Case WFRES_JUDG ''���莯�ʃt���O(RES)
                d1(0) = d(3)
            Case WFOI_JUDG ''���莯�ʃt���O(Oi)
                d1(0) = d(2)
            Case WFDOI_JUDG, WFBMD_JUDG, WFDZ_JUDG, WFAOI_JUDG ''���莯�ʃt���O(��Oi)�A���莯�ʃt���O(BMD)�A���莯�ʃt���O(DZ)�A���莯�ʃt���O(AOi) �ǉ� 03/12/09 ooba
                d1(0) = d(1)
'            Case WFDSOD_JUDG ''���莯�ʃt���O(DSOD)
'            Case WFSPV_JUDG ''���莯�ʃt���O(SPV)
'            Case WFOSF_JUDG ''���莯�ʃt���O(OSF)
            Case Else
'''''                WFCJudgDialog.WFCErrorMessage "�Ώۃf�[�^�����B"
                FuncAns = FUNCTION_RETURN_FAILURE '' �ُ�
            End Select
        Else
'''''            WFCJudgDialog.WFCErrorMessage "�Ώۃf�[�^�����B"
            FuncAns = FUNCTION_RETURN_FAILURE '' �ُ�
        End If
    Case ObjCode05 ''�S�_�̕��ϒl
        d1(0) = JudgAve(d())
    Case ObjCode06 ''�S�_�̍ő�l
        d1(0) = JudgMax(d())
    Case ObjCode07 ''�S�_�̕��ϒl�ƍő�l
        d1(0) = JudgAve(d())
        d1(1) = JudgMax(d())
    Case ObjCode08 ''�S�_�̍ŏ��l
        d1(0) = JudgMin(d())
    Case ObjCode09 ''������2�_�A�O����2�_(5�_�����1,2,4,5)
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
    Case ObjCode10 ''MAX(2,4�_��)
        If (d(1) <> -1) And (d(3) <> -1) Then
            If d(1) >= d(3) Then
                d1(0) = d(1)
            Else
                d1(0) = d(3)
            End If
        Else
'''''            WFCJudgDialog.WFCErrorMessage "�Ώۃf�[�^�����B"
            FuncAns = FUNCTION_RETURN_FAILURE '' �ُ�
        End If
    Case ObjCode11 ''MAX(2,3,4�_��)
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
            ''���蓾�Ȃ��G���[
'''''            WFCJudgDialog.WFCErrorMessage "�Ώۃf�[�^�����B"
            FuncAns = FUNCTION_RETURN_FAILURE '' �ُ�
        End If
''    Case ObjCode12 ''���ۏ�
''    Case ObjCode13 ''�_��
''    Case ObjCode14 ''�`�󑪒�(���R�x�A���Ԃ�AWARP)
''    Case ObjCode15 ''�K�i�Ȃ�
    '----TEST2004/10
    Case ObjCode16 ''�S�_�̍ő�l�ƍŏ��l
        'Select Case JudgFlag 04/12/16�폜
            'Case WFBMD_JUDG
                d1(0) = JudgMax(d())
                d1(1) = JudgMin(d())
        'End Select
        
'' 2008/10/20 BMD�]��,�O��1�_�ۏ؋@�\�ǉ� ADD By Systech Start
    Case ObjCode18 ''AVE+�O��1�_
        d1(0) = d(0)            '�O���P�_
        d1(1) = JudgAve(d())    '�S�_���ϒl
'' 2008/10/20 BMD�]��,�O��1�_�ۏ؋@�\�ǉ� ADD By Systech End
        
    Case Else
'''''        WFCJudgDialog.WFCErrorMessage "�Ώۃf�[�^�����B"
        FuncAns = FUNCTION_RETURN_FAILURE '' �ُ�
    End Select
    
    GetWfJudgData = FuncAns
End Function

'�T�v      :WFC����Ώ�MIN�f�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WFCMin(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim Min As Double
    
    Min = -9999
    
    Select Case G.cPos
    Case "1"                                           '2003/05/15 �ǉ��@osawa �˗�No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(0) < d(3), d(0), d(3))
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
            Min = IIf(d(6) < Min, d(6), Min)
            Min = IIf(d(9) < Min, d(9), Min)
        End Select                                     '�˗�No.030130�@�ǉ������܂�
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(0) < d(4), d(0), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(6), d(0), d(6))
            Min = IIf(d(9) < Min, d(9), Min)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(0) < d(4), d(0), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(6), d(0), d(6))
            Min = IIf(d(9) < Min, d(9), Min)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(0) < d(4), d(0), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(6), d(0), d(6))
            Min = IIf(d(9) < Min, d(9), Min)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(1) < d(4), d(1), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(6), d(0), d(6))
            Min = IIf(d(9) < Min, d(9), Min)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(1) < d(4), d(1), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(5), d(0), d(5))
            Min = IIf(d(8) < Min, d(8), Min)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(1) < d(4), d(1), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(5), d(0), d(5))
            Min = IIf(d(8) < Min, d(8), Min)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(1) < d(4), d(1), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(5), d(0), d(5))
            Min = IIf(d(8) < Min, d(8), Min)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(2) < d(4), d(2), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(4), d(0), d(4))
            Min = IIf(d(7) < Min, d(7), Min)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(3), d(0), d(3))
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(3) < d(4), d(3), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(1), d(0), d(1))
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(1) < d(3), d(1), d(3))
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
            Min = IIf(d(6) < Min, d(6), Min)
            Min = IIf(d(9) < Min, d(9), Min)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(1) < d(3), d(1), d(3))
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
            Min = IIf(d(5) < Min, d(5), Min)
            Min = IIf(d(8) < Min, d(8), Min)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(2) < d(3), d(2), d(3))
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
            Min = IIf(d(4) < Min, d(4), Min)
            Min = IIf(d(7) < Min, d(7), Min)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(2) < d(4), d(2), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(4), d(0), d(4))
            Min = IIf(d(7) < Min, d(7), Min)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(2) < d(3), d(2), d(3))
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(2), d(0), d(2))
            Min = IIf(d(4) < Min, d(4), Min)
            Min = IIf(d(7) < Min, d(7), Min)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = IIf(d(0) < d(1), d(0), d(1))
            Min = IIf(d(3) < Min, d(3), Min)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            Min = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            Min = IIf(d(0) < d(1), d(0), d(1))
            Min = IIf(d(2) < Min, d(2), Min)
            Min = IIf(d(3) < Min, d(3), Min)
            Min = IIf(d(4) < Min, d(4), Min)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case " "
    End Select
    
    WFCMin = Min
End Function

'�T�v      :WFC����Ώ�MAX�f�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WFCMax(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim max As Double
    
    max = -9999
    
    Select Case G.cPos
    Case "1"                                         '2003/05/15 �ǉ��@osawa �˗�No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(0) > d(3), d(0), d(3))
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
            max = IIf(d(6) > max, d(6), max)
            max = IIf(d(9) > max, d(9), max)
        End Select                                   '�˗�No.030130�@�ǉ������܂�
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(0) > d(4), d(0), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(6), d(0), d(6))
            max = IIf(d(9) > max, d(9), max)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(0) > d(4), d(0), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(6), d(0), d(6))
            max = IIf(d(9) > max, d(9), max)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(0) > d(4), d(0), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(6), d(0), d(6))
            max = IIf(d(9) > max, d(9), max)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(1) > d(4), d(1), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(6), d(0), d(6))
            max = IIf(d(9) > max, d(9), max)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(1) > d(4), d(1), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(5), d(0), d(5))
            max = IIf(d(8) > max, d(8), max)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(1) > d(4), d(1), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(5), d(0), d(5))
            max = IIf(d(8) > max, d(8), max)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(1) > d(4), d(1), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(5), d(0), d(5))
            max = IIf(d(8) > max, d(8), max)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(2) > d(4), d(2), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(4), d(0), d(4))
            max = IIf(d(7) > max, d(7), max)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(3), d(0), d(3))
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(3) > d(4), d(3), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(1), d(0), d(1))
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(1) > d(3), d(1), d(3))
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
            max = IIf(d(6) > max, d(6), max)
            max = IIf(d(9) > max, d(9), max)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(1) > d(3), d(1), d(3))
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
            max = IIf(d(5) > max, d(5), max)
            max = IIf(d(8) > max, d(8), max)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(2) > d(3), d(2), d(3))
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
            max = IIf(d(4) > max, d(4), max)
            max = IIf(d(7) > max, d(7), max)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(2) > d(4), d(2), d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(4), d(0), d(4))
            max = IIf(d(7) > max, d(7), max)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(2) > d(3), d(2), d(3))
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(2), d(0), d(2))
            max = IIf(d(4) > max, d(4), max)
            max = IIf(d(7) > max, d(7), max)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = IIf(d(0) > d(1), d(0), d(1))
            max = IIf(d(3) > max, d(3), max)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            max = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            max = IIf(d(0) > d(1), d(0), d(1))
            max = IIf(d(2) > max, d(2), max)
            max = IIf(d(3) > max, d(3), max)
            max = IIf(d(4) > max, d(4), max)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case " "
    End Select

    WFCMax = max
End Function

'�T�v      :WFC����Ώ�ave�f�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WFCAve(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim AVE As Double
    
    AVE = -9999
    
    Select Case G.cPos
    Case "1"                                               '2003/05/15 �ǉ��@osawa �˗�No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(0) + d(3) + d(4)) / 3
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(2) + d(6) + d(9)) / 4
        End Select                                         '�˗�No.030130�@�ǉ������܂�
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(0) + d(4)) / 2
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(6) + d(9)) / 3
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(0) + d(4)) / 2
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(6) + d(9)) / 3
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(0) + d(4)) / 2
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(6) + d(9)) / 3
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(1) + d(4)) / 2
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(6) + d(9)) / 3
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(1) + d(4)) / 2
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(5) + d(8)) / 3
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(1) + d(4)) / 2
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(5) + d(8)) / 3
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(1) + d(4)) / 2
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(5) + d(8)) / 3
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(2) + d(4)) / 2
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(4) + d(7)) / 3
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(3)) / 2
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(3) + d(4)) / 2
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(2)) / 2
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(1)) / 2
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(1) + d(3) + d(4)) / 3
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(2) + d(6) + d(9)) / 4
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(1) + d(3) + d(4)) / 3
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(2) + d(5) + d(8)) / 4
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(2) + d(3) + d(4)) / 3
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(2) + d(4) + d(7)) / 4
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(2) + d(4)) / 2
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(4) + d(7)) / 3
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(2) + d(3) + d(4)) / 3
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(2) + d(4) + d(7)) / 4
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(1) + d(3)) / 3
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            AVE = (d(0) + d(1) + d(2) + d(3) + d(4)) / 5
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            AVE = (d(0) + d(1) + d(2) + d(3) + d(4) + d(5) + d(6) + d(7) + d(8) + d(9)) / 10
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case " "
    End Select

    WFCAve = AVE
End Function

'�T�v      :WFC����ΏۃZ���^�[�ʒu�f�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WFCCenterP(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim CenterP As Double
    
    CenterP = -9999
    
    Select Case G.cPos
    Case "1"                                      '2003/05/15 �ǉ��@osawa �˗�No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select                                '�˗�No.030130�@�ǉ������܂�
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "D"
            CenterP = d(0)
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterP = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterP = d(0)
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case " "
    End Select

    WFCCenterP = CenterP
End Function

'�T�v      :WFC����Ώے����l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WFCCenterD(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim CenterD As Double
    Dim temp() As Double
    
    CenterD = -9999
    
    Select Case G.cPos
    Case "1"                                       '2003/05/15 �ǉ��@osawa �˗�No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(3)
            temp(2) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(3) As Double
            temp(0) = d(0)
            temp(1) = d(2)
            temp(2) = d(6)
            temp(3) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        End Select                                '�˗�No.030130�@�ǉ������܂�
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(1) As Double
            temp(0) = d(0)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(6)
            temp(2) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(1) As Double
            temp(0) = d(0)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(6)
            temp(2) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(1) As Double
            temp(0) = d(1)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(6)
            temp(2) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(1) As Double
            temp(0) = d(1)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(6)
            temp(2) = d(9)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(1) As Double
            temp(0) = d(1)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(5)
            temp(2) = d(8)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(1) As Double
            temp(0) = d(1)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(5)
            temp(2) = d(8)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(1) As Double
            temp(0) = d(1)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(5)
            temp(2) = d(8)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(1) As Double
            temp(0) = d(2)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(4)
            temp(2) = d(7)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(1) As Double
            temp(0) = d(0)
            temp(1) = d(3)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(1) As Double
            temp(0) = d(3)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(1) As Double
            temp(0) = d(0)
            temp(1) = d(2)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(1) As Double
            temp(0) = d(0)
            temp(1) = d(1)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(2) As Double
            temp(0) = d(1)
            temp(1) = d(3)
            temp(2) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(2) As Double
            temp(0) = d(1)
            temp(1) = d(3)
            temp(2) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(2) As Double
            temp(0) = d(2)
            temp(1) = d(3)
            temp(2) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(1) As Double
            temp(0) = d(2)
            temp(1) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(4)
            temp(2) = d(7)
            BubbleSort temp()
            CenterD = temp(Int((1 + 1) / 2))
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(2) As Double
            temp(0) = d(2)
            temp(1) = d(3)
            temp(2) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ReDim temp(2) As Double
            temp(0) = d(0)
            temp(1) = d(2)
            temp(2) = d(3)
            BubbleSort temp()
            CenterD = temp(Int((2 + 1) / 2))
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            CenterD = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            CenterD = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ReDim temp(4) As Double
            temp(0) = d(0)
            temp(1) = d(1)
            temp(2) = d(2)
            temp(3) = d(3)
            temp(4) = d(4)
            BubbleSort temp()
            CenterD = temp(Int((4 + 1) / 2))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case " "
    End Select
    
    WFCCenterD = CenterD
End Function

'�T�v      :WFC����Ώ�R/2�l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WFCR2(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim r2 As Double
    
    r2 = -9999
    
    Select Case G.cPos
    Case "1"                                         '2003/05/15 �ǉ��@osawa �˗�No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            r2 = d(3)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            r2 = d(2)
        End Select                                   '�˗�No.030130�@�ǉ������܂�
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            r2 = d(3)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            r2 = d(2)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            r2 = d(3)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            r2 = d(2)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            r2 = d(3)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            r2 = d(2)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            r2 = d(3)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            r2 = d(2)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            r2 = d(3)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            r2 = d(2)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            r2 = d(3)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            r2 = d(2)
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case " "
    End Select
    
    WFCR2 = r2
End Function

'�T�v      :WFC����Ώہi|Center-Side|Max�j�l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
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
    Case "1"                                              '2003/05/15 �ǉ��@osawa �˗�No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(0) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select                                         '�˗�No.030130�@�ǉ������܂�
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(0) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(0) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(0) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(5))
            ce_side1 = Abs(d(0) - d(8))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(5))
            ce_side1 = Abs(d(0) - d(8))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(5))
            ce_side1 = Abs(d(0) - d(8))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(2) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(4))
            ce_side1 = Abs(d(0) - d(7))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(6))
            ce_side1 = Abs(d(0) - d(9))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(1) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(5))
            ce_side1 = Abs(d(0) - d(8))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(2) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(4))
            ce_side1 = Abs(d(0) - d(7))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(2) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(4))
            ce_side1 = Abs(d(0) - d(7))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side_max = Abs(d(2) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ce_side0 = Abs(d(0) - d(4))
            ce_side1 = Abs(d(0) - d(7))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ce_side0 = Abs(d(0) - d(4))
            ce_side1 = Abs(d(1) - d(4))
            ce_side2 = Abs(d(2) - d(4))
            ce_side_max = IIf(ce_side0 > ce_side1, ce_side0, ce_side1)
            ce_side_max = IIf(ce_side2 > ce_side_max, ce_side2, ce_side_max)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case " "
    End Select
    
    WFCCE_Side_Max = ce_side_max
End Function

'�T�v      :WFC����Ώے��S���ϒl�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WFCCEAve(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim ceave As Double
    
    ceave = -9999
    
    Select Case G.cPos
    Case "1"                                        '2003/05/15 �ǉ��@osawa �˗�No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select                                   '�˗�No.030130�@�ǉ������܂�
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(0)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            ceave = d(4)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            ceave = d(0)
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case " "
    End Select
    
    WFCCEAve = ceave
End Function

'�T�v      :WFC����Ώ�Side���ϒl�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WFCSideAve(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim sideave As Double
    
    sideave = -9999
    
    Select Case G.cPos
    Case "1"                                       '2003/05/15 �ǉ��@osawa �˗�No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(0)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select                                  '�˗�No.030130�@�ǉ������܂�
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(0)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(0)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(0)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(1)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(1)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(5) + d(8)) / 2
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(1)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(5) + d(8)) / 2
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(1)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(5) + d(8)) / 2
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(2)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(4) + d(7)) / 2
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(1)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(6) + d(9)) / 2
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(1)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(5) + d(8)) / 2
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(2)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(4) + d(7)) / 2
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(2)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(4) + d(7)) / 2
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = d(2)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(4) + d(7)) / 2
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            sideave = (d(0) + d(1) + d(2)) / 3
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            sideave = (d(4) + d(5) + d(6) + d(7) + d(8) + d(9)) / 6
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case " "
    End Select
    
    WFCSideAve = sideave
End Function

'�T�v      :WFC����Ώہi|Center-R/2|Max�j�l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function WFCCE_R2_Max(JudgFlag As Integer, d() As Double, G As Guarantee) As Double
    Dim cd_r2_max As Double
    
    cd_r2_max = -9999
    
    Select Case G.cPos
    Case "1"                                           '2003/05/15 �ǉ��@osawa �˗�No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select                                      '�˗�No.030130�@�ǉ������܂�
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "3"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "4"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "5"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "6"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "7"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "8"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "A"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "B"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "H"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "J"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "L"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "M"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "K"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            cd_r2_max = Abs(d(3) - d(4))
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
            cd_r2_max = Abs(d(0) - d(2))
        End Select
    Case "S"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case " "
    End Select
    
    WFCCE_R2_Max = cd_r2_max
End Function



'�T�v      :�ʓ����z�v�Z��[N]�ɑΉ����A�q�n�f�l���Z�o����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Res           ,I  ,W_RES     ,WF�Z���^�[���R����\����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2002/10/18  yakimura  �쐬
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
    Case "1"                                                      '2003/05/15 �ǉ��@osawa �˗�No.030130
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(0))
             auto_cal = Abs(C - d(0)) / Abs(C + d(0)) * 200
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
          If (C <> -9999) And (d(6) <> -9999) Then
            auto_cal1 = Abs(C - d(6)) / Abs(C + d(6)) * 200
          End If

          If (C <> -9999) And (d(9) <> -9999) Then
                auto_cal2 = Abs(C - d(9)) / Abs(C + d(9)) * 200
          End If

          auto_cal = IIf(auto_cal1 > auto_cal2, auto_cal1, auto_cal2)
        End Select                                                '�˗�No.030130�@�ǉ������܂�
    Case "2"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(4))
            auto_cal = Abs(C - d(4)) / Abs(C + d(4)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(4))
            auto_cal = Abs(C - d(4)) / Abs(C + d(4)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(4))
            auto_cal = Abs(C - d(4)) / Abs(C + d(4)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(2))
            auto_cal = Abs(C - d(2)) / Abs(C + d(2)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "C"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "D"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "E"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "F"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "G"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(1))
            auto_cal = Abs(C - d(1)) / Abs(C + d(1)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(2))
            auto_cal = Abs(C - d(2)) / Abs(C + d(2)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(2))
            auto_cal = Abs(C - d(2)) / Abs(C + d(2)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
            'auto_cal = Abs(d(2))
            auto_cal = Abs(C - d(2)) / Abs(C + d(2)) * 200           '2003/5/16
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Y"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Q"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "N"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "P"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
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
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select

    Case "R"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
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
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "T"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "U"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "V"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "W"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "X"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case "Z"
        Select Case JudgFlag
        Case WFRES_JUDG ''���莯�ʃt���O(RES)
        Case WFOI_JUDG ''���莯�ʃt���O(Oi)
        End Select
    Case " "
    End Select
    
    WF_TypeN_Exc = auto_cal
End Function

''Upd start 2005/06/22 (TCS)t.terauchi  SPV9�_�Ή�
'�T�v      :WF�Z���^�[SPV(Fe�Z�x MAP����)������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Spv           ,I  ,W_SPV            ,WF�Z���^�[SPV����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2005/06/22 �V�K�쐬 (TCS)t.terauchi
Public Function WfSPV_Fe_AMXJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
        
    If Spv.GuaranteeSpvFe.cJudg = JudgCodeW01 Then ''SPVFE�Z�x�@����L��
                
        ''SPV(Fe�Z�x MAP����)����
        Select Case Spv.GuaranteeSpvFe.cObj
            
            '�S����_(3)�EMAX(B)
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
            
            '���̑�
            Case Else
                ''�_���A�K�i�����ȊO�̏ꍇ
                If (Spv.GuaranteeSpvFe.cObj <> ObjCode13) And (Spv.GuaranteeSpvFe.cObj <> ObjCode15) Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
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

'�T�v      :WF�Z���^�[SPV(Fe�Z�x 9�_����)������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Spv           ,I  ,W_SPV            ,WF�Z���^�[SPV����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2005/06/22 �V�K�쐬 (TCS)t.terauchi
Public Function WfSPV_Fe_V9TJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
        
    If Spv.GuaranteeSpvFe.cJudg = JudgCodeW01 Then ''SPVFE�Z�x�@����L��
        
        ''SPV(Fe�Z�x 9�_����)����
        Select Case Spv.GuaranteeSpvFe.cObj
        
            '���S�_(1)
            Case ObjCode01
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(3), 0, Spv.SpecSpvFeMax)
        
            '�S����_(3)�EMAX(B)
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
            
            '���̑�
            Case Else
                ''�_���A�K�i�����ȊO�̏ꍇ
                If (Spv.GuaranteeSpvFe.cObj <> ObjCode13) And (Spv.GuaranteeSpvFe.cObj <> ObjCode15) Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
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

'�T�v      :WF�Z���^�[SPV(�g�U�� MAP����)������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Spv           ,I  ,W_SPV            ,WF�Z���^�[SPV����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2005/06/22 �V�K�쐬 (TCS)t.terauchi
Public Function WfSPV_DIFF_AMXJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
    
    If Spv.GuaranteeSpv.cJudg = JudgCodeW01 Then ''SPV�g�U���@����L��
                
        ''SPV(�g�U�� MAP����)����
        Select Case Spv.GuaranteeSpv.cObj
        
            '�S����_(3)
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
               
            'AVE+MIN(L)�@08/03/13 ooba
            Case ObjCode17
                'AVE�����(AVE����)�ȏォ��,MIN������(MIN����)�ȏ�
                If RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMax, -1) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, -1)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
                
            '���̑�
            Case Else
                ''�_���A�K�i�����ȊO�̏ꍇ
                If (Spv.GuaranteeSpv.cObj <> ObjCode13) And (Spv.GuaranteeSpv.cObj <> ObjCode15) Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
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

'�T�v      :WF�Z���^�[SPV(�g�U�� 9�_����)������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Spv           ,I  ,W_SPV            ,WF�Z���^�[SPV����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2005/06/22 �V�K�쐬 (TCS)t.terauchi
Public Function WfSPV_DIFF_V9TJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
    
    If Spv.GuaranteeSpv.cJudg = JudgCodeW01 Then ''SPV�g�U���@����L��
        
        ''SPV(�g�U�� 9�_����)
        Select Case Spv.GuaranteeSpv.cObj
            
            '���S�_(1)
            Case ObjCode01
                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(3), Spv.SpecSpvMin, Spv.SpecSpvMax)
            
            '�S����_(3)
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
            
            'AVE+MIN(L)�@08/03/13 ooba
            Case ObjCode17
                'AVE�����(AVE����)�ȏォ��,MIN������(MIN����)�ȏ�
                If RangeDecision_nl(Spv.Spv(2), Spv.SpecSpvMax, -1) Then
                    Spv.JudgSpv = RangeDecision_nl(Spv.Spv(1), Spv.SpecSpvMin, -1)
                Else
                    Spv.JudgSpv = JUDG_NG
                End If
                
            '���̑�
            Case Else
                ''�_���A�K�i�����ȊO�̏ꍇ
                If (Spv.GuaranteeSpv.cObj <> ObjCode13) And (Spv.GuaranteeSpv.cObj <> ObjCode15) Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
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

'�T�v      :WF�Z���^�[SPV(Nr�Z�x MAP����)������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Spv           ,I  ,W_SPV            ,WF�Z���^�[SPV����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2006/06/12 �V�K�쐬 SMP)kondoh
Public Function WfSPV_Nr_AMXJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
        
    If Spv.GuaranteeSpvNr.cJudg = JudgCodeW01 Then ''SPVNR�Z�x�@����L��
                
        ''SPV(Nr�Z�x MAP����)����
        Select Case Spv.GuaranteeSpvNr.cObj
            
            '�S����_(3)�EMAX(B)
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
            
            '���̑�
            Case Else
                ''�_���A�K�i�����ȊO�̏ꍇ
                If (Spv.GuaranteeSpvNr.cObj <> ObjCode13) And (Spv.GuaranteeSpvNr.cObj <> ObjCode15) Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
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

'�T�v      :WF�Z���^�[SPV(Nr�Z�x 9�_����)������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Spv           ,I  ,W_SPV            ,WF�Z���^�[SPV����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2006/06/12 �V�K�쐬 SMP)kondoh
Public Function WfSPV_Nr_V9TJudg(Spv As W_SPV, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Spv.JudgSpv = JUDG_NG
                
    If Spv.GuaranteeSpvNr.cJudg = JudgCodeW01 Then ''SPVNR�Z�x�@����L��
        
        ''SPV(Fe�Z�x 9�_����)����
        Select Case Spv.GuaranteeSpvNr.cObj

'' DB���Nr�Z�x�̒��S�̍��ڂ��������߁A����s�\
''            '���S�_(1)
''            Case ObjCode01
''                Spv.JudgSpv = RangeDecision_nl(Spv.Spv(3), 0, Spv.SpecSpvNrMax)
        
            '�S����_(3)�EMAX(B)
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
            
            '���̑�
            Case Else
                ''�_���A�K�i�����ȊO�̏ꍇ
                If (Spv.GuaranteeSpvNr.cObj <> ObjCode13) And (Spv.GuaranteeSpvNr.cObj <> ObjCode15) Then
                    ''�Ώۃf�[�^����
                    ''�G���[���\���̂ɏ������B
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

