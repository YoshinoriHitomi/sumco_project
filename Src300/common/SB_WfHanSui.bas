Attribute VB_Name = "SB_WfHanSui"
Option Explicit

'-------------------------------------------------------------------------------
' �萔��`
'-------------------------------------------------------------------------------
'XSDCW
Private Const cWFSMPLID     As String = "WFSMPLID"      'XSDCW�̃T���v���h�c
Private Const cWFIND        As String = "WFIND"         'XSDCW�̏��FLG
Private Const cWFRES        As String = "WFRES"         'XSDCW�̎���FLG
Private Const cWFHS         As String = "WFHS"          'XSDCW�̕ۏ�FLG '�ǉ� 05/01/28 ooba
Private Const cCW           As String = "CW"            'XSDCW�̍��ڍŏI����
Private Const cWF_RS        As String = "RS"            'XSDCW��Rs
Private Const cWF_OI        As String = "OI"            'XSDCW��Oi
Private Const cWF_B1        As String = "B1"            'XSDCW��BMD1
Private Const cWF_B2        As String = "B2"            'XSDCW��BMD2
Private Const cWF_B3        As String = "B3"            'XSDCW��BMD3
Private Const cWF_O1        As String = "L1"            'XSDCW��OSF1
Private Const cWF_O2        As String = "L2"            'XSDCW��OSF2
Private Const cWF_O3        As String = "L3"            'XSDCW��OSF3
Private Const cWF_O4        As String = "L4"            'XSDCW��OSF4
Private Const cWF_DS        As String = "DS"            'XSDCW��DS
Private Const cWF_DZ        As String = "DZ"            'XSDCW��DZ
Private Const cWF_SP        As String = "SP"            'XSDCW��SP
Private Const cWF_DO1       As String = "DO1"           'XSDCW��DO1
Private Const cWF_DO2       As String = "DO2"           'XSDCW��DO2
Private Const cWF_DO3       As String = "DO3"           'XSDCW��DO3
Private Const cWF_OT1       As String = "OT1"           'XSDCW��OT1
Private Const cWF_OT2       As String = "OT2"           'XSDCW��OT2
Private Const cWF_AOI       As String = "AOI"           'XSDCW��AOI
Private Const cWF_GD        As String = "GD"            'XSDCW��GD      '�ǉ� 05/01/28 ooba
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
Private Const cEPSMPLID     As String = "EPSMPLID"      'XSDCW�̃T���v��ID(�G�s����)
Private Const cEPIND        As String = "EPIND"         'XSDCW�̏��FLG(�G�s����)
Private Const cEPRES        As String = "EPRES"         'XSDCW�̎���FLG(�G�s����)
Private Const cEP_B1        As String = "B1"            'XSDCW��BMD1E
Private Const cEP_B2        As String = "B2"            'XSDCW��BMD2E
Private Const cEP_B3        As String = "B3"            'XSDCW��BMD3E
Private Const cEP_O1        As String = "L1"            'XSDCW��OSF1E
Private Const cEP_O2        As String = "L2"            'XSDCW��OSF2E
Private Const cEP_O3        As String = "L3"            'XSDCW��OSF3E
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

'------------------------------------------------
' �v�e���f/����`�F�b�N���ʊ֐�
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�]�����ڇ��ɂ��A���f�����肩�𔻒f���A�v�e���f�`�F�b�N�A�܂��́A�v�e����`�F�b�N���Ăяo���B�i���ʊ֐��j
'           ���ʊ֐��̃`�F�b�N���ʂ𓖊֐��̌��ʂƂ��āA�Ăяo�����֕Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sSXLid        ,I  ,String       :SXL-ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :�����ԍ�
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :iSmplPos      ,I  ,Integer      :�V�T���v���ʒu(mm)
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� =  1 RS     �� �ΏۊO
'                                                       =  2 Oi     �� �����1
'                                                       =  3 BMD1   �� �����1
'                                                       =  4 BMD2   �� �����1
'                                                       =  5 BMD3   �� �����1
'                                                       =  6 OSF1   �� �����1
'                                                       =  7 OSF2   �� �����1
'                                                       =  8 OSF3   �� �����1
'                                                       =  9 OSF4   �� �����1
'                                                       = 10 DS     �� �����1
'                                                       = 11 DZ     �� �����1
'                                                       = 12 SP     �� �����2
'                                                       = 13 D1     �� �����1
'                                                       = 14 D2     �� �����1
'                                                       = 15 D3     �� �����1
'                                                       = 18 AO     �� �����1   '�c���_�f�ǉ��@03/12/09 ooba
'                                                       = 19 GD     �� �����1   'GD�ǉ��@05/01/26 ooba
'                                                       = 20 BMD1E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 21 BMD2E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 22 BMD3E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 23 OSF1E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 24 OSF2E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 25 OSF3E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'          :iFromPos      ,I  ,Integer      :�����͈�From
'          :iToPos        ,I  ,Integer      :�����͈�To
'          :iHanSuiKBN    ,O  ,Integer      :���f/����敪(0:���f,1:����)
'          :sGetSmplID1   ,O  ,String       :���T���v��ID1
'          :sGetSmplID2   ,O  ,String       :���T���v��ID2 (���f�����g�p)
'          :sGetHSflg1    ,O  ,String       :���T���v���̕ۏ�FLG    '�ǉ��@05/01/28 ooba
'          :�߂�l        ,O  ,Integer      :�`�F�b�N���� = 0 : ����I��(���f/����OK)
'                                                           1 : ����I��(���f/����NG)
'                                                          -1 : ���͈����l�G���[
'                                                          -2 : ��L�ȊO�̃G���[
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funChkWfHanSui(sSXLID As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                iItemNo As Integer, iFromPos As Integer, iToPos As Integer, iHanSuiKBN As Integer, _
                                sGetSmplID1 As String, sGetSmplID2 As String, Optional sGetHSflg1 As String = "") As Integer
    Dim retCode As Integer
    
    '���T���v��ID������
    sGetSmplID1 = ""
    sGetSmplID2 = ""
    sGetHSflg1 = "0"     '05/02/18 ooba
    
    '�p�����[�^�`�F�b�N
    If (Len(sSXLID) <> 13) Then GoTo ChkWfHanSuiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkWfHanSuiParameterErr
    
    '�w�肳�ꂽ�]�����ڇ��ɂ��A���f�����肩�𔻒f���A�v�e���f�`�F�b�N�A�܂��́A�v�e����`�F�b�N���Ăяo���B
    Select Case iItemNo
    Case 1          'RS(���R)
        retCode = 1
        iHanSuiKBN = 1
'    Case 2 To 15
'    Case 2 To 18
    '�c���_�f�ǉ��@03/12/09 ooba
    'GD�ǉ��@05/01/26 ooba
'    Case 2 To 19
    '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
    Case 2 To 25    'Oi(�_�f�Z�x),BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4,DS,DZ,SP,D1,D2,D3,--,--,AO,GD,BMD1E,BMD2E,BMD3E,OSF1E,OSF2E,OSF3E
'        retCode = funChkWfHanei(sSXLid, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, sGetSmplID1)
        '�ۏ�FLG�ǉ��@05/01/28 ooba
        retCode = funChkWfHanei(sSXLID, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, sGetSmplID1, sGetHSflg1)
        iHanSuiKBN = 0
    Case Else
        GoTo ChkWfHanSuiParameterErr
    End Select
    
    '���ʊ֐��̃`�F�b�N���ʂ𓖊֐��̌��ʂƂ��āA�Ăяo�����֕Ԃ��B
    funChkWfHanSui = retCode
    Exit Function

ChkWfHanSuiParameterErr:
    funChkWfHanSui = -1
    Exit Function

ChkWfHanSuiSonotaErr:
    funChkWfHanSui = -2
End Function

'------------------------------------------------
' �v�e���f�`�F�b�N
'------------------------------------------------

'�T�v      :�w�肳�ꂽ��񂩂�A�v�e���f�`�F�b�N���s�Ȃ����ʂ�Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sSXLid        ,I  ,String       :SXL-ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :�����ԍ�
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :iSmplPos      ,I  ,Integer      :�V�T���v���ʒu(mm)
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� =  1 RS     �� �ΏۊO
'                                                       =  2 Oi     �� �����1
'                                                       =  3 BMD1   �� �����1
'                                                       =  4 BMD2   �� �����1
'                                                       =  5 BMD3   �� �����1
'                                                       =  6 OSF1   �� �����1
'                                                       =  7 OSF2   �� �����1
'                                                       =  8 OSF3   �� �����1
'                                                       =  9 OSF4   �� �����1
'                                                       = 10 DS     �� �����1
'                                                       = 11 DZ     �� �����1
'                                                       = 12 SP     �� �����2
'                                                       = 13 D1     �� �����1
'                                                       = 14 D2     �� �����1
'                                                       = 15 D3     �� �����1
'                                                       = 18 AO     �� �����1   '�c���_�f�ǉ��@03/12/09 ooba
'                                                       = 19 GD     �� �����1   'GD�ǉ��@05/01/26 ooba
'                                                       = 20 BMD1E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 21 BMD2E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 22 BMD3E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 23 OSF1E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 24 OSF2E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 25 OSF3E  �� �����1   '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'          :iFromPos      ,I  ,Integer      :�����͈�From
'          :iToPos        ,I  ,Integer      :�����͈�To
'          :sGetSmplID    ,O  ,String       :���f���T���v��ID
'          :sGetHSflg     ,O  ,String       :���f���T���v���̕ۏ�FLG    '�ǉ��@05/01/28 ooba
'          :�߂�l        ,O  ,Integer      :�`�F�b�N���� = 0 : ����I��(���fOK)
'                                                           1 : ����I��(���fNG)
'                                                          -1 : ���͈����l�G���[
'                                                          -2 : ��L�ȊO�̃G���[
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funChkWfHanei(sSXLID As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, iFromPos As Integer, iToPos As Integer, sGetSmplID As String, sGetHSflg As String) As Integer
    Dim wHPtrn          As Integer
    Dim tSiyou          As type_DBDRV_scmzc_fcmlc001c_Siyou
    Dim wGetSXLid       As String
    Dim wGetSmpKbn      As String
    Dim wGetSmplID      As String
    Dim wGetHSflg       As String       '05/01/28 ooba
    
    Dim tTBCMY013       As typ_TBCMY013
    Dim tGDjisseki      As typ_TBCMJ015 '05/01/31 ooba

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    Dim tTBCMY022       As typ_TBCMY022
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

''Upd start 2005/06/28 (TCS)t.terauchi  SPV9�_�Ή�
    Dim tSPVJisseki     As typ_TBCMJ016
    Dim sPos            As String
''Upd end   2005/06/28 (TCS)t.terauchi  SPV9�_�Ή�
    
    Dim retJudg         As Boolean
    Dim wIdFlg          As Integer
    Dim TmpData(2)      As String
    
    Dim dShiyo()        As Double       '2003/12/11 Null�Ή��ǉ�
    Dim sHosyo          As String       '2003/12/11 Null�Ή��ǉ�
    
    Dim tSiyou_Sxl      As type_DBDRV_scmzc_fcmkc001c_Siyou '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    
    '������
    wGetSmplID = ""
    wGetHSflg = ""      '05/01/28 ooba
    
    '�p�����[�^�`�F�b�N
    If (Len(sSXLID) <> 13) Then GoTo ChkWfHaneiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkWfHaneiParameterErr
    
    '�w�肳�ꂽ�]�����ڇ����ɕK�v�ȕi�Ԏd�l�l���擾���A�v�e���f�l�擾�p�^�[�������肷��B�i�w�肳�ꂽ�]�����ڇ��ɂ��A�������������B�j
    Select Case iItemNo
    Case 1              'RS(���R)
        GoTo ChkWfHaneiNG
    Case 2              'Oi(�_�f�Z�x)
        If funGet_TBCME025(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
        
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        ReDim dShiyo(5)
        dShiyo(1) = tSiyou.HWFONMIN         ' �i�v�e�_�f�Z�x����
        dShiyo(2) = tSiyou.HWFONMAX         ' �i�v�e�_�f�Z�x���
        dShiyo(3) = tSiyou.HWFONMBP         ' �i�v�e�_�f�Z�x�ʓ����z
        dShiyo(4) = tSiyou.HWFONAMN         ' �i�v�e�_�f�Z�x���ω���
        dShiyo(5) = tSiyou.HWFONAMX         ' �i�v�e�_�f�Z�x���Ϗ��
        If fncJissekiHantei_nl(tSiyou.HWFONHWS, dShiyo) = False Then GoTo ChkWfHaneiSonotaErr
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        
    Case 3 To 9         'BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4
        If funGet_TBCME029(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
        
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        'BMD�d�lNULL�`�F�b�N���폜�i�����OK�Ƃ���B�j�@          2003/12/19 tuku
''''        ReDim dShiyo(1)
''''        If iItemNo = 3 Then         'BMD1
''''            sHosyo = tSiyou.HWFBM1HS            ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
''''            dShiyo(1) = tSiyou.HWFBM1MBP        ' �i�v�e�a�l�c�P�ʓ����z
''''        ElseIf iItemNo = 4 Then     'BMD2
''''            sHosyo = tSiyou.HWFBM2HS            ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
''''            dShiyo(1) = tSiyou.HWFBM2MBP        ' �i�v�e�a�l�c�Q�ʓ����z
''''        ElseIf iItemNo = 5 Then     'BMD3
''''            sHosyo = tSiyou.HWFBM3HS            ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
''''            dShiyo(1) = tSiyou.HWFBM3MBP        ' �i�v�e�a�l�c�R�ʓ����z
''''        ElseIf iItemNo = 6 Then     'OSF1
''''        ElseIf iItemNo = 7 Then     'OSF2
''''        ElseIf iItemNo = 8 Then     'OSF3
''''        ElseIf iItemNo = 9 Then     'OSF4
''''        End If
''''        If fncJissekiHantei_nl(sHosyo, dShiyo) = False Then GoTo ChkWfHaneiSonotaErr
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        
    Case 10             'DSOD
        If funGet_TBCME026(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
        
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        ReDim dShiyo(4)
        dShiyo(1) = tSiyou.HWFDSOMX         ' �i�v�e�c�r�n�c���
        dShiyo(2) = tSiyou.HWFDSOMN         ' �i�v�e�c�r�n�c����
        dShiyo(3) = tSiyou.HWFDSOAX         ' �i�v�e�c�r�n�c�̈���
        dShiyo(4) = tSiyou.HWFDSOAN         ' �i�v�e�c�r�n�c�̈扺��
        If fncJissekiHantei_nl(tSiyou.HWFDSOHS, dShiyo) = False Then GoTo ChkWfHaneiSonotaErr
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        
    Case 11             'DZ��
        If funGet_TBCME024(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
    Case 12             'SPVFE
        If funGet_TBCME028(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
'        wHPtrn = 2
        'SPV�̔��f����݂�1�ɕύX�@04/04/27 ooba
        wHPtrn = 1
    Case 13, 14, 15     'DOI1(�_�f�͏o1),DOI2(�_�f�͏o2),DOI3(�_�f�͏o3)
        If funGet_TBCME025(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
    Case 18             'AO     '�c���_�f�ǉ��@03/12/09 ooba
        If funGet_TBCME025(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
    Case 19             'GD     'GD�ǉ��@05/01/26 ooba
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        If funGet_TBCME020(tFullHin, tSiyou_Sxl) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG

        If funGet_TBCME036(tFullHin, tSiyou_Sxl) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        
        tSiyou.HSXGDPTK = tSiyou_Sxl.HSXGDPTK
        tSiyou.HSXLDLRMN = tSiyou_Sxl.HSXLDLRMN
        tSiyou.HSXLDLRMX = tSiyou_Sxl.HSXLDLRMX
        tSiyou.HWFLDLRMN = tSiyou_Sxl.HWFLDLRMN
        tSiyou.HWFLDLRMX = tSiyou_Sxl.HWFLDLRMX
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
        
        If funGet_TBCME026(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        
    '*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݐ��擾
        If funGet_TBCME036_2(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
    '*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݐ��擾
        wHPtrn = 1
            
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    Case 20 To 25
        If funGet_TBCME050(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    Case Else
        GoTo ChkWfHaneiParameterErr
    End Select

    '�v�e���f���T���v���h�c�̎擾
    If wHPtrn = 1 Then              '�������f�l�擾�p�^�[���P
'        If funGetWfHanei1(sSXLid, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos,
'                                                                    wGetSXLid, wGetSmpKbn, wGetSmplID) <> 0 Then GoTo ChkWfHaneiNG
'        '�ۏ�FLG�ǉ��@05/01/28 ooba
'        If funGetWfHanei1(sSXLid, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, _
'                                                                    wGetSXLid, wGetSmpKbn, wGetSmplID, wGetHSflg) <> 0 Then GoTo ChkWfHaneiNG

        If funGetWfHanei1(sSXLID, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, _
                                                                    wGetSXLid, wGetSmpKbn, wGetSmplID, wGetHSflg) <> 0 Then
            '' ����GD���f�Ή��@05/06/13 ooba START ======================================>
            'GD
            If iItemNo = 19 Then
                'TOP�̏ꍇ
                If sTB = "T" And IsNumeric(CrySampleID.TsmplidGD) Then
                    '���������ID�ƕۏ�FLG(1:�����ۏ�)���
                    wGetSmplID = CrySampleID.TsmplidGD
                    wGetHSflg = "1"
                'BOT�̏ꍇ
                ElseIf sTB = "B" And IsNumeric(CrySampleID.BsmplidGD) Then
                    '���������ID�ƕۏ�FLG(1:�����ۏ�)���
                    wGetSmplID = CrySampleID.BsmplidGD
                    wGetHSflg = "1"
                Else
                    GoTo ChkWfHaneiNG
                End If
            Else
                GoTo ChkWfHaneiNG
            End If
            '' ����GD���f�Ή��@05/06/13 ooba END ========================================>
        End If
        
    ElseIf wHPtrn = 2 Then          '�������f�l�擾�p�^�[���Q
        If funGetWfHanei2(sSXLID, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, _
                                                                    wGetSXLid, wGetSmpKbn, wGetSmplID) <> 0 Then GoTo ChkWfHaneiNG
    End If
    
    '�������f�������ID����A�������f�l�i���ђl�j���擾����B�i�w�肳�ꂽ�]�����ڇ��ɂ��A�������������B�j
    Select Case iItemNo
'    Case 1              'RS(���R)
'        GoTo ChkWfHaneiNG
    Case 2              'Oi(�_�f�Z�x)
        'Oi�̎��ђl���擾����
        If funGetTBCMY013(wGetSmplID, "OI", "OI", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
        'Oi����������s�Ȃ�
        If Not WfCrOiJudg(tSiyou, tTBCMY013, retJudg) Then GoTo ChkWfHaneiNG

    Case 3, 4, 5                'BMD1, BMD2, BMD3
        If iItemNo = 3 Then
            'BMD1�̎��ђl���擾����
            If funGetTBCMY013(wGetSmplID, "BMD", "BMD1", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 4 Then
            'BMD2�̎��ђl���擾����
            If funGetTBCMY013(wGetSmplID, "BMD", "BMD2", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 5 Then
            'BMD3�̎��ђl���擾����
            If funGetTBCMY013(wGetSmplID, "BMD", "BMD3", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 3
        End If
        'BMD�̑���������s�Ȃ�
        If Not WfCrBmdJudg(tSiyou, tTBCMY013, retJudg, wIdFlg) Then GoTo ChkWfHaneiNG
        
    Case 6, 7, 8, 9             'OSF1, OSF2, OSF3, OSF4
        If iItemNo = 6 Then
            'OSF1�̎��ђl���擾����
            If funGetTBCMY013(wGetSmplID, "OSF", "OSF1", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 7 Then
            'OSF2�̎��ђl���擾����
            If funGetTBCMY013(wGetSmplID, "OSF", "OSF2", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 8 Then
            'OSF3�̎��ђl���擾����
            If funGetTBCMY013(wGetSmplID, "OSF", "OSF3", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 3
        ElseIf iItemNo = 9 Then
            'OSF4�̎��ђl���擾����
            If funGetTBCMY013(wGetSmplID, "OSF", "OSF4", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 4
        End If
        'OSF�̑���������s�Ȃ�
        If Not WfCrOsfJudg(tSiyou, tTBCMY013, retJudg, wIdFlg, TmpData) Then GoTo ChkWfHaneiNG
    
    Case 10             'DSOD
        'DSOD�̎��ђl���擾����
        If funGetTBCMY013(wGetSmplID, "DSOD", "DSOD", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
    
        'DSOD����������s�Ȃ�
        If Not WfCrDsodjudg(tSiyou, tTBCMY013, retJudg) Then GoTo ChkWfHaneiNG
    
    Case 11             'DZ��
        'DZ���̎��ђl���擾����
        If funGetTBCMY013(wGetSmplID, "DZ", "DZ", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
    
        'DZ������������s�Ȃ�
        If Not WfCrDzjudg(tSiyou, tTBCMY013, retJudg) Then GoTo ChkWfHaneiNG
    
    Case 12             'SPVFE
        
    ''Upd start 2005/06/28 (TCS)t.terauchi  SPV9�_�Ή�
'        'SPVFE�̎��ђl���擾����
'        If funGetTBCMY013(wGetSmplID, "SPV", "SPV", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
'
'        'SPVFE����������s�Ȃ�
'        If Not WfCrSpvjudg(tSiyou, tTBCMY013, retJudg) Then GoTo ChkWfHaneiNG
        
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '' WF�d�l(SPV)�擾
        If funWfcGetDataEtc_SPV(tFullHin, _
                                tSiyou) <> FUNCTION_RETURN_SUCCESS Then GoTo ChkWfHaneiNG
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                
        'SPVFE�̎��ђl���擾����
        If funGetSPVJisseki_J016(sCryNum, wGetSmplID, _
                        tSPVJisseki, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        
        
        '�����ް��Ȃ�
        If Trim(tSPVJisseki.SMPLNO) = "0" Then GoTo ChkWfHaneiNG
        
        If sTB = "T" Then
            sPos = "TOP"
        ElseIf sTB = "B" Then
            sPos = "BOT"
        Else
            GoTo ChkWfHaneiParameterErr
        End If
        
        'SPV(Fe�Z�x)����������s�Ȃ�
        If ((tSiyou.HWFSPVHS = "H") And CheckKHN(tSiyou.HWFSPVKN, 15, sPos)) Then
            If Not WfCrSpvJudg_New(tSiyou, tSPVJisseki, retJudg, 1, sPos) Then GoTo ChkWfHaneiNG
        Else
            retJudg = True
        End If
        
        If retJudg = True Then
            'SPV(�g�U��)����������s�Ȃ�
            If ((tSiyou.HWFDLHWS = "H") And CheckKHN(tSiyou.HWFDLKHN, 16, sPos)) Then
                If Not WfCrSpvJudg_New(tSiyou, tSPVJisseki, retJudg, 2, sPos) Then GoTo ChkWfHaneiNG
            End If
        End If
    ''Upd end   2005/06/28 (TCS)t.terauchi  SPV9�_�Ή�

'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        If retJudg = True Then
            'SPV(Nr�Z�x)����������s�Ȃ�
            If ((tSiyou.HWFNRHS = "H") And CheckKHN(tSiyou.HWFNRKN, 19, sPos)) Then
                If Not WfCrSpvJudg_New(tSiyou, tSPVJisseki, retJudg, 3, sPos) Then GoTo ChkWfHaneiNG
            Else
                retJudg = True
            End If
        End If
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------


    
    Case 13, 14, 15             'DOI1(�_�f�͏o1),DOI2(�_�f�͏o2),DOI3(�_�f�͏o3)
        If iItemNo = 13 Then
            'DOI1�̎��ђl���擾����
            If funGetTBCMY013(wGetSmplID, "DOI", "DOI1", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 14 Then
            'DOI2�̎��ђl���擾����
            If funGetTBCMY013(wGetSmplID, "DOI", "DOI2", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 15 Then
            'DOI3�̎��ђl���擾����
            If funGetTBCMY013(wGetSmplID, "DOI", "DOI2", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 3
        End If
        'DOI�̑���������s�Ȃ�
        If Not WfCrDoiJudg(tSiyou, tTBCMY013, retJudg, wIdFlg) Then GoTo ChkWfHaneiNG
    
    ''�c���_�f�ǉ��@03/12/09 ooba START ==================================================>
    Case 18             'AOi
        'AOi�̎��ђl���擾����
        If funGetTBCMY013(wGetSmplID, "AOI", "AOI", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
        
        'AOi�̑���������s�Ȃ�
        If Not WfCrAoiJudg(tSiyou, tTBCMY013, retJudg) Then GoTo ChkWfHaneiNG
    ''�c���_�f�ǉ��@03/12/09 ooba END ====================================================>
    
    ''GD�ǉ��@05/01/31 ooba START ========================================================>
    Case 19             'GD
        If wGetHSflg = "1" Then
            'GD�̌������ђl���擾����
            If funGetGDJisseki_J006(sCryNum, wGetSmplID, tGDjisseki) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
            
            tSiyou.WFHSGDCW = "1"   '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
        Else
            'GD��WF���ђl���擾����
            If funGetGDJisseki_J015(sCryNum, wGetSmplID, wGetHSflg, tGDjisseki) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
            
            tSiyou.WFHSGDCW = "0"   '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
        End If
        '�����ް��Ȃ�
        If Trim(tGDjisseki.SMPLNO) = "" Then GoTo ChkWfHaneiNG
        
        'GD�̑���������s�Ȃ�
        If Not WfCrGdJudg(tSiyou, tGDjisseki, retJudg) Then GoTo ChkWfHaneiNG
    ''GD�ǉ��@05/01/31 ooba END ==========================================================>

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    Case 20 To 22
        If iItemNo = 20 Then
            'BMD1(EP)�̎��ђl���擾����
            If funGetTBCMY022(wGetSmplID, "BMD", "BMD1", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 21 Then
            'BMD2(EP)�̎��ђl���擾����
            If funGetTBCMY022(wGetSmplID, "BMD", "BMD2", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 22 Then
            'BMD3(EP)�̎��ђl���擾����
            If funGetTBCMY022(wGetSmplID, "BMD", "BMD3", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 3
        End If
        'BMD�̑���������s�Ȃ�
        If Not EpBmdJudg(tSiyou, tTBCMY022, retJudg, wIdFlg) Then GoTo ChkWfHaneiNG
    Case 23 To 25
        If iItemNo = 23 Then
            'OSF1(EP)�̎��ђl���擾����
            If funGetTBCMY022(wGetSmplID, "OSF", "OSF1", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 24 Then
            'OSF2(EP)�̎��ђl���擾����
            If funGetTBCMY022(wGetSmplID, "OSF", "OSF2", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 25 Then
            'OSF3(EP)�̎��ђl���擾����
            If funGetTBCMY022(wGetSmplID, "OSF", "OSF3", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 3
        End If
        'OSF�̑���������s�Ȃ�
        If Not EpOsfJudg(tSiyou, tTBCMY022, retJudg, wIdFlg, TmpData) Then GoTo ChkWfHaneiNG

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

'    Case Else
'        GoTo ChkWfHaneiParameterErr
    End Select
    
    '�w�肳�ꂽ�]�����ڇ��̑������肪OK�̏ꍇ�A���f���T���v��ID��ݒ肵�A�߂�l��'0'(����I��(���fOK))��ݒ肵�A�������I������B
    '�������肪NG�̏ꍇ�A�߂�l��'1'(����I��(���fNG))��ݒ肵�A�������I������B
    If retJudg = False Then GoTo ChkWfHaneiNG
        
    sGetSmplID = wGetSmplID
    sGetHSflg = wGetHSflg       '05/01/28 ooba
    funChkWfHanei = 0
    Exit Function

ChkWfHaneiNG:
    sGetSmplID = wGetSmplID
    sGetHSflg = wGetHSflg       '05/01/28 ooba
    funChkWfHanei = 1
    Exit Function

ChkWfHaneiParameterErr:
    funChkWfHanei = -1
    Exit Function

ChkWfHaneiSonotaErr:
    funChkWfHanei = -2
End Function

'------------------------------------------------
' �v�e���f�l�擾�i�p�^�[���P�j
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�V�T���v���ʒu��񂩂�A�v�e���f���T���v���h�c��V�T���v���Ǘ�(SXL)(XSDCW)��茟�����A���ʂ�Ԃ��B
'           ���f���悤�Ƃ���V�T���v���ʒu���ATOP�̏ꍇ��BOT�̏ꍇ�Ō������@(����)���قȂ�B
'           ���f���T���v���h�c����������ꍇ�A��{�I�ɂ́A�V�T���v���ʒu���猩�āA�㉺�T���v���̒��ŋ߂��ق��̃T���v���h�c�𒊏o����B
'           ��������ۂ̌����͈͂́A�w�肳�ꂽ�͈͓��̂ݗL���Ƃ��A�����͈͓��ɂ݂���Ȃ��ꍇ�A�u�Y������قȂ��v�Ƃ���B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sSXLid        ,I  ,String       :SXL-ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :�����ԍ�
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :iSmplPos      ,I  ,Integer      :�V�T���v���ʒu(mm)
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� =  1 RS
'                                                       =  2 Oi     ���Ώ�
'                                                       =  3 BMD1   ���Ώ�
'                                                       =  4 BMD2   ���Ώ�
'                                                       =  5 BMD3   ���Ώ�
'                                                       =  6 OSF1   ���Ώ�
'                                                       =  7 OSF2   ���Ώ�
'                                                       =  8 OSF3   ���Ώ�
'                                                       =  9 OSF4   ���Ώ�
'                                                       = 10 DS     ���Ώ�
'                                                       = 11 DZ     ���Ώ�
'                                                       = 12 SP
'                                                       = 13 D1     ���Ώ�
'                                                       = 14 D2     ���Ώ�
'                                                       = 15 D3     ���Ώ�
'                                                       = 18 AO     ���Ώ�  '�ǉ��@03/12/09 ooba
'                                                       = 19 GD     ���Ώ�  '�ǉ��@05/01/26 ooba
'                                                       = 20 BMD1E  ���Ώ�  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 21 BMD2E  ���Ώ�  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 22 BMD3E  ���Ώ�  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 23 OSF1E  ���Ώ�  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 24 OSF2E  ���Ώ�  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 25 OSF3E  ���Ώ�  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'          :iFromPos      ,I  ,Integer      :�����͈�From
'          :iToPos        ,I  ,Integer      :�����͈�To
'          :sGetSXLid     ,O  ,String       :���f��SXL-ID
'          :sGetSmpKbn    ,O  ,String       :���f���T���v���敪
'          :sGetSmplID    ,O  ,String       :���f���T���v���h�c
'          :sGetHSflg     ,O  ,String       :���f���T���v���̕ۏ�FLG    '�ǉ��@05/01/28 ooba
'          :�߂�l        ,O  ,Integer      :�擾���� = 0 : ����I��
'                                                       1 : ����I��(�Y���T���v���Ȃ�)
'                                                      -1 : �ُ�I��
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetWfHanei1(sSXLID As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, iFromPos As Integer, iToPos As Integer, _
                               sGetSXLid As String, sGetSmpKbn As String, sGetSmplID As String, sGetHSflg As String) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       '�����ID����
    Dim ediInd      As String       '���FLG����
    Dim ediRes      As String       '����FLG����
    Dim ediHs       As String       '�ۏ�FLG����    '05/01/28 ooba
    
    '�p�����[�^�`�F�b�N
    If (Len(sSXLID) <> 13) Then GoTo GetWfHanei1ParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetWfHanei1ParameterErr
    
    '�w�肳�ꂽ�]�����ڇ�����A�����Ώƕ]�����ږ������肷��B
    kName = funGetWfKensaName(iItemNo)
    If kName = " " Then GoTo GetWfHanei1ParameterErr
        
    Select Case iItemNo
    Case 20 To 25
        'SQL�����Ŏg�p���閼�̂ɕҏW
        ediSmpid = cEPSMPLID & kName & cCW     '�����ID
        ediInd = cEPIND & kName & cCW          '���FLG
        ediRes = cEPRES & kName & cCW          '����FLG
    Case Else
        'SQL�����Ŏg�p���閼�̂ɕҏW
        ediSmpid = cWFSMPLID & kName & cCW     '�����ID
        ediInd = cWFIND & kName & cCW          '���FLG
        ediRes = cWFRES & kName & cCW          '����FLG
    End Select
    
    '�ۏ��׸ސݒ�@05/01/28 ooba
    Select Case iItemNo
    Case 19     'GD
        ediHs = cWFHS & kName & cCW        '�ۏ�FLG
    Case Else
        ediHs = "'0'"
    End Select
    
    
    '�w�肳�ꂽ�������ɁA�V����يǗ�(SXL)(XSDCW)����������B
'    sql = "select SXLIDCW, SMPKBNCW, " & ediSmpid & " as SMPLID from XSDCW "
    '�ۏ��׸ޒǉ��@05/01/28 ooba
    sql = "select "
    sql = sql & "SXLIDCW, "
    sql = sql & "SMPKBNCW, "
    sql = sql & ediSmpid & " as SMPLID, "
    sql = sql & ediHs & " as HSFLG "
    sql = sql & "from XSDCW "

    'TOP�ʒu(T/B�敪='T')�̌���
    If sTB = "T" Then
        sql = sql & "where tbkbncw = '" & sTB & "' and "
        sql = sql & "      xtalcw = '" & sCryNum & "' and "
        sql = sql & "      inposcw <= " & iSmplPos & " and "
        sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
        sql = sql & "  " & ediRes & " <> '0' and "
        sql = sql & "      inposcw >= " & iFromPos & " and "
        sql = sql & "      inposcw <= " & iToPos & " "
        sql = sql & "order by inposcw desc"
    
    'BOT�ʒu(T/B�敪='B')�̌���
    ElseIf sTB = "B" Then
        sql = sql & "where tbkbncw = '" & sTB & "' and "
        sql = sql & "      xtalcw = '" & sCryNum & "' and "
        sql = sql & "      inposcw >= " & iSmplPos & " and "
        sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
        sql = sql & "  " & ediRes & " <> '0' and "
        sql = sql & "      inposcw >= " & iFromPos & " and "
        sql = sql & "      inposcw <= " & iToPos & " "
        sql = sql & "order by inposcw asc"
    Else
        GoTo GetWfHanei1ParameterErr
    End If
    
    'SQL���̎��s
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetWfHanei1 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '�Ăяo�����ւ̌��ʒʒm
    sGetSXLid = rs("SXLIDCW")
    sGetSmpKbn = rs("SMPKBNCW")
    sGetSmplID = rs("SMPLID")
    sGetHSflg = rs("HSFLG")     '05/01/28 ooba
    Set rs = Nothing
    
    funGetWfHanei1 = 0
    Exit Function

GetWfHanei1ParameterErr:
    funGetWfHanei1 = -1
End Function

'------------------------------------------------
' �v�e���f�l�擾�i�p�^�[���Q�j
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�V�T���v���ʒu��񂩂�A�v�e���f���T���v���h�c��V�T���v���Ǘ�(SXL)(XSDCW)��茟�����A���ʂ�Ԃ��B
'           ���f���T���v���h�c����������ꍇ�A��{�I�ɂ́A�V�T���v���ʒu���猩�āA���T���v���̒��ŋ߂��ق��̃T���v���h�c�𒊏o����B
'           ��������ۂ̌����͈͂́A�w�肳�ꂽ�͈͓��̂ݗL���Ƃ��A�����͈͓��ɂ݂���Ȃ��ꍇ�A�u�Y������قȂ��v�Ƃ���B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sSXLid        ,I  ,String       :SXL-ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :�����ԍ�
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :iSmplPos      ,I  ,Integer      :�V�T���v���ʒu(mm)
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� =  1 RS
'                                                       =  2 Oi
'                                                       =  3 BMD1
'                                                       =  4 BMD2
'                                                       =  5 BMD3
'                                                       =  6 OSF1
'                                                       =  7 OSF2
'                                                       =  8 OSF3
'                                                       =  9 OSF4
'                                                       = 10 DS
'                                                       = 11 DZ
'                                                       = 12 SP     ���Ώ�
'                                                       = 13 D1
'                                                       = 14 D2
'                                                       = 15 D3
'                                                       = 18 AO
'                                                       = 19 GD
'                                                       = 20 BMD1E  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 21 BMD2E  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 22 BMD3E  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 23 OSF1E  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 24 OSF2E  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                       = 25 OSF3E  '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
'          :iFromPos      ,I  ,Integer      :�����͈�From
'          :iToPos        ,I  ,Integer      :�����͈�To
'          :sGetSXLid     ,O  ,String       :���f��SXL-ID
'          :sGetSmpKbn    ,O  ,String       :���f���T���v���敪
'          :sGetSmplID    ,O  ,String       :���f���T���v���h�c
'          :�߂�l        ,O  ,Integer      :�擾���� = 0 : ����I��
'                                                       1 : ����I��(�Y���T���v���Ȃ�)
'                                                      -1 : �ُ�I��
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetWfHanei2(sSXLID As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, iFromPos As Integer, iToPos As Integer, _
                               sGetSXLid As String, sGetSmpKbn As String, sGetSmplID As String) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       '�����ID����
    Dim ediInd      As String       '���FLG����
    Dim ediRes      As String       '����FLG����
    
    '�p�����[�^�`�F�b�N
    If (Len(sSXLID) <> 13) Then GoTo GetWfHanei2ParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetWfHanei2ParameterErr
    
    '�w�肳�ꂽ�]�����ڇ�����A�����Ώƕ]�����ږ������肷��B
    kName = funGetWfKensaName(iItemNo)
    If kName = " " Then GoTo GetWfHanei2ParameterErr
    

    Select Case iItemNo
    Case 20 To 25
        'SQL�����Ŏg�p���閼�̂ɕҏW
        ediSmpid = cEPSMPLID & kName & cCW     '�����ID
        ediInd = cEPIND & kName & cCW          '���FLG
        ediRes = cEPRES & kName & cCW          '����FLG
    Case Else
        'SQL�����Ŏg�p���閼�̂ɕҏW
        ediSmpid = cWFSMPLID & kName & cCW     '�����ID
        ediInd = cWFIND & kName & cCW          '���FLG
        ediRes = cWFRES & kName & cCW          '����FLG
    End Select
    
    '�w�肳�ꂽ�������ɁA�V����يǗ�(SXL)(XSDCW)����������B
    sql = "select SXLIDCW, SMPKBNCW, " & ediSmpid & " as SMPLID from XSDCW "
    sql = sql & "where xtalcw = '" & sCryNum & "' and "
    'SPV�͕K��BOT������擾����悤�ɕύX�@04/04/23 ooba
    If kName = "SP" Then sql = sql & "TBKBNCW = 'B' and "
    sql = sql & "      inposcw > " & iSmplPos & " and "
    sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
    sql = sql & "  " & ediRes & " <> '0' and "
    sql = sql & "      inposcw >= " & iFromPos & " and "
    sql = sql & "      inposcw <= " & iToPos & " "
    sql = sql & "order by inposcw asc"
    
    'SQL���̎��s
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetWfHanei2 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '�Ăяo�����ւ̌��ʒʒm
    sGetSXLid = rs("SXLIDCW")
    sGetSmpKbn = rs("SMPKBNCW")
    sGetSmplID = rs("SMPLID")
    Set rs = Nothing
    
    funGetWfHanei2 = 0
    Exit Function
    
GetWfHanei2ParameterErr:
    funGetWfHanei2 = -1
End Function

'------------------------------------------------
' �v�e�����Ώە]�����ږ��擾
'------------------------------------------------

'�T�v      :�]�����ڇ�����A�v�e�����Ώە]�����ږ���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� �� Sxl =  1 RS
'                                                              =  2 Oi
'                                                              =  3 BMD1
'                                                              =  4 BMD2
'                                                              =  5 BMD3
'                                                              =  6 OSF1
'                                                              =  7 OSF2
'                                                              =  8 OSF3
'                                                              =  9 OSF4
'                                                              = 10 DS
'                                                              = 11 DZ
'                                                              = 12 SP
'                                                              = 13 DO1
'                                                              = 14 DO2
'                                                              = 15 DO3
'                                                              = 18 AO     '�ǉ��@03/12/09 ooba
'                                                              = 19 GD     '�ǉ��@05/01/28 ooba
'                                                              = 20 BMD1E  '2006/08/15 �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                              = 21 BMD2E  '2006/08/15 �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                              = 22 BMD3E  '2006/08/15 �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                              = 23 OSF1E  '2006/08/15 �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                              = 24 OSF2E  '2006/08/15 �G�s��s�]���ǉ��Ή� SMP)kondoh
'                                                              = 25 OSF3E  '2006/08/15 �G�s��s�]���ǉ��Ή� SMP)kondoh
'          :�߂�l        ,O  ,Sting        :�����Ώۍ��ږ�(���Ұ��װ���́A�󔒂�Ԃ�)
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetWfKensaName(iItemNo As Integer) As String
    
    '�p�����[�^�`�F�b�N
'    If iItemNo < 1 Or iItemNo > 15 Then GoTo GetWfKensaNameParameterErr
    ''�c���_�f�������ڒǉ��ɂ��ύX�@03/12/15 ooba
'    If iItemNo < 1 Or iItemNo > 18 Then GoTo GetWfKensaNameParameterErr
    'GD�ǉ��ɂ��ύX�@05/01/28 ooba
'    If iItemNo < 1 Or iItemNo > 19 Then GoTo GetWfKensaNameParameterErr
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    If iItemNo < 1 Or iItemNo > 25 Then GoTo GetWfKensaNameParameterErr
    
    'SXL
    Select Case iItemNo
    Case 1:     funGetWfKensaName = cWF_RS        'RS(���R)
    Case 2:     funGetWfKensaName = cWF_OI        'Oi(�_�f�Z�x)
    Case 3:     funGetWfKensaName = cWF_B1        'BMD1
    Case 4:     funGetWfKensaName = cWF_B2        'BMD2
    Case 5:     funGetWfKensaName = cWF_B3        'BMD3
    Case 6:     funGetWfKensaName = cWF_O1        'OSF1
    Case 7:     funGetWfKensaName = cWF_O2        'OSF2
    Case 8:     funGetWfKensaName = cWF_O3        'OSF3
    Case 9:     funGetWfKensaName = cWF_O4        'OSF4
    Case 10:    funGetWfKensaName = cWF_DS        'CS(�Y�f�Z�x)
    Case 11:    funGetWfKensaName = cWF_DZ        'GD
    Case 12:    funGetWfKensaName = cWF_SP        'LT(ײ����)
    Case 13:    funGetWfKensaName = cWF_DO1       'EPD
    Case 14:    funGetWfKensaName = cWF_DO2       'LT(ײ����)
    Case 15:    funGetWfKensaName = cWF_DO3       'EPD
    Case 18:    funGetWfKensaName = cWF_AOI       'AOi      '�c���_�f�ǉ��@03/12/15 ooba
    Case 19:    funGetWfKensaName = cWF_GD        'GD       'GD�ǉ��@05/01/28 ooba
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    Case 20:    funGetWfKensaName = cEP_B1        'BMD1E
    Case 21:    funGetWfKensaName = cEP_B2        'BMD2E
    Case 22:    funGetWfKensaName = cEP_B3        'BMD3E
    Case 23:    funGetWfKensaName = cEP_O1        'OSF1E
    Case 24:    funGetWfKensaName = cEP_O2        'OSF2E
    Case 25:    funGetWfKensaName = cEP_O3        'OSF3E
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    End Select
    
    Exit Function

GetWfKensaNameParameterErr:
    funGetWfKensaName = " "
End Function

'------------------------------------------------
' ����]������(WF�̊e�����)�擾�֐�
'------------------------------------------------

'�T�v      :�T���v���h�c����ATBCMY013���������A����]������(WF�̊e�����)���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                   :����
'          :sSmplID       ,I  ,String               :�T���v���h�c
'          :sItemName     ,I  ,String               :�]�����ږ���(RES,OI,BMD,OSF,DSOD,DZ,SPV,DOI)
'          :sItemDetail   ,I  ,String               :�]�����ڏڍז���(RES,OI,BMD1�`BMD3,OSF1�`OSF4,DSOD,DZ,SPV,DOI1�`DOI3)
'          :tTBCMY013     ,O  ,typ_TBCMY013         :����]������(�\����)
'          :�߂�l        ,O  ,Integer              :�擾���� = 0 : ����
'                                                              -1 : �ُ�
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetTBCMY013(sSmplID As String, sItemName As String, sItemDetail As String, tTBCMY013 As typ_TBCMY013) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "SB_WfHanSui.bas -- Function funGetTBCMY013"
    
    '�T���v���h�c������TBCMY013�̑���]������(WF�̊e�����)����������B
    sql = "select SAMPLEID, OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5, "
    sql = sql & "MESDATA6, MESDATA7, MESDATA8, MESDATA9, MESDATA10, MESDATA11, MESDATA12, MESDATA13, MESDATA14, MESDATA15, "
    sql = sql & "TXID, REGDATE, SENDFLAG, SENDDATE "
    sql = sql & "from TBCMY013 "
    sql = sql & "where SAMPLEID = '" & sSmplID & "' and "
    sql = sql & "      OSITEM = '" & sItemName & "' and "
    sql = sql & "      SPEC = '" & sItemDetail & "'"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetTBCMY013 = -1
        GoTo PROC_EXIT
    End If
    
     ''���o���ʂ��i�[����
    With tTBCMY013
        .SAMPLEID = rs("SAMPLEID")       ' �T���v��ID
        .OSITEM = rs("OSITEM")           ' �]������
        .MAISU = rs("MAISU")             ' �]������
        .Spec = rs("SPEC")               ' �K�i�l
        .NETSU = rs("NETSU")             ' �M��������
        .ET = rs("ET")                   ' �G�b�`���O����
        .MES = rs("MES")                 ' �v�����@
        .DKAN = rs("DKAN")               ' �c�j�A�j�[������
        .MESDATA1 = rs("MESDATA1")       ' ����f�[�^���̂P
        .MESDATA2 = rs("MESDATA2")       ' ����f�[�^���̂Q
        .MESDATA3 = rs("MESDATA3")       ' ����f�[�^���̂R
        .MESDATA4 = rs("MESDATA4")       ' ����f�[�^���̂S
        .MESDATA5 = rs("MESDATA5")       ' ����f�[�^���̂T
        .MESDATA6 = rs("MESDATA6")       ' ����f�[�^���̂U
        .MESDATA7 = rs("MESDATA7")       ' ����f�[�^���̂V
        .MESDATA8 = rs("MESDATA8")       ' ����f�[�^���̂W
        .MESDATA9 = rs("MESDATA9")       ' ����f�[�^���̂X
        .MESDATA10 = rs("MESDATA10")     ' ����f�[�^���̂P�O
        .MESDATA11 = rs("MESDATA11")     ' ����f�[�^����1�P
        .MESDATA12 = rs("MESDATA12")     ' ����f�[�^����1�Q
        .MESDATA13 = rs("MESDATA13")     ' ����f�[�^����1�R
        .MESDATA14 = rs("MESDATA14")     ' ����f�[�^����1�S
        .MESDATA15 = rs("MESDATA15")     ' ����f�[�^����1�T
        .TXID = rs("TXID")               ' �g�����U�N�V����ID
        .REGDATE = rs("REGDATE")         ' �o�^���t
        .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
        .SENDDATE = rs("SENDDATE")       ' ���M���t
    End With
    
    funGetTBCMY013 = 0

PROC_EXIT:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

PROC_ERR:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume PROC_EXIT
End Function

'------------------------------------------------
' EP��s�]�����ʎ擾�֐�
'------------------------------------------------

'�T�v      :�T���v���h�c�A�]�����ڂ���TBCMY022���������AEP��s�]�����ʂ��擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                   :����
'          :sSmplID       ,I  ,String               :�T���v���h�c
'          :sItemName     ,I  ,String               :�G�s�]�����ږ���(BMD,OSF)
'          :sItemDetail   ,I  ,String               :�G�s�]�����ڏڍז���(BMD1�`BMD3,OSF1�`OSF3)
'          :tTBCMY022     ,O  ,typ_TBCMY022         :�G�s����]������(�\����)
'          :�߂�l        ,O  ,Integer              :�擾���� = 0 : ����
'                                                              -1 : �ُ�
'����      :SB_WfHanSui.funGetTBCMY013����ɍ쐬
'����      :�V�K�쐬 2006/08/15 �G�s��s�]���ǉ��Ή� SMP)kondoh

Public Function funGetTBCMY022(sSmplID As String, sItemName As String, sItemDetail As String, tTBCMY022 As typ_TBCMY022) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "SB_WfHanSui.bas -- Function funGetTBCMY022"
    
    '�T���v���h�c������TBCMY022��EP��s����]�����ʂ���������B
    sql = "select SAMPLEID, OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5, "
    sql = sql & "MESDATA6, MESDATA7, MESDATA8, MESDATA9, MESDATA10, MESDATA11, MESDATA12, MESDATA13, MESDATA14, MESDATA15, "
    sql = sql & "TXID, REGDATE, SENDFLAG, SENDDATE "
    sql = sql & "from TBCMY022 "
    sql = sql & "where SAMPLEID = '" & sSmplID & "' and "
    sql = sql & "      OSITEM = '" & sItemName & "' and "
    sql = sql & "      SPEC = '" & sItemDetail & "'"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetTBCMY022 = -1
        GoTo PROC_EXIT
    End If
    
     ''���o���ʂ��i�[����
    With tTBCMY022
        .SAMPLEID = rs("SAMPLEID")       ' �T���v��ID
        .OSITEM = rs("OSITEM")           ' �]������
        .MAISU = rs("MAISU")             ' �]������
        .Spec = rs("SPEC")               ' �K�i�l
        .NETSU = rs("NETSU")             ' �M��������
        .ET = rs("ET")                   ' �G�b�`���O����
        .MES = rs("MES")                 ' �v�����@
        .DKAN = rs("DKAN")               ' �c�j�A�j�[������
        .MESDATA1 = rs("MESDATA1")       ' ����f�[�^���̂P
        .MESDATA2 = rs("MESDATA2")       ' ����f�[�^���̂Q
        .MESDATA3 = rs("MESDATA3")       ' ����f�[�^���̂R
        .MESDATA4 = rs("MESDATA4")       ' ����f�[�^���̂S
        .MESDATA5 = rs("MESDATA5")       ' ����f�[�^���̂T
        .MESDATA6 = rs("MESDATA6")       ' ����f�[�^���̂U
        .MESDATA7 = rs("MESDATA7")       ' ����f�[�^���̂V
        .MESDATA8 = rs("MESDATA8")       ' ����f�[�^���̂W
        .MESDATA9 = rs("MESDATA9")       ' ����f�[�^���̂X
        .MESDATA10 = rs("MESDATA10")     ' ����f�[�^���̂P�O
        .MESDATA11 = rs("MESDATA11")     ' ����f�[�^����1�P
        .MESDATA12 = rs("MESDATA12")     ' ����f�[�^����1�Q
        .MESDATA13 = rs("MESDATA13")     ' ����f�[�^����1�R
        .MESDATA14 = rs("MESDATA14")     ' ����f�[�^����1�S
        .MESDATA15 = rs("MESDATA15")     ' ����f�[�^����1�T
        .TXID = rs("TXID")               ' �g�����U�N�V����ID
        .REGDATE = rs("REGDATE")         ' �o�^���t
        .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
        .SENDDATE = rs("SENDDATE")       ' ���M���t
    End With
    
    funGetTBCMY022 = 0

PROC_EXIT:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function
    
PROC_ERR:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume PROC_EXIT
End Function


