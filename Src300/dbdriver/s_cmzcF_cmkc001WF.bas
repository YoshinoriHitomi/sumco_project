Attribute VB_Name = "s_cmzcF_cmkc001WF"
Option Explicit

' WF�T���v���d�l(*�͖��`�F�b�N�̃p�����[�^)
Public Type typ_SpWFSamp
    hin As tFullHinban      ' �i��

    HWFRHWYS As String * 1  ' �������@(Rs)
    HWFRSPOH As String * 1  ' ������@(Rs)*
    HWFRSPOT As String * 1  ' ����_��(Rs) -> Heavy
    HWFRSPOI As String * 1  ' ����ʒu(Rs)*

    HWFONHWS As String * 1  ' �������@(Oi)
    HWFONKWY As String * 2  ' �������@(Oi)
    HWFONSPH As String * 1  ' ������@(Oi)
    HWFONSPT As String * 1  ' ����_��(Oi) -> Heavy
    HWFONSPI As String * 1  ' ����ʒu(Oi)

    HWFBM1HS As String * 1  ' �������@(B1)
    HWFBM1SH As String * 1  ' ������@(B1)
    HWFBM1ST As String * 1  ' ����_��(B1)
    HWFBM1SR As String * 1  ' ���O�̈�(B1)
    HWFBM1NS As String * 2  ' �M�����@(B1)
    HWFBM1SZ As String * 1  ' �������(B1)
    HWFBM1ET As Integer     ' �I���G�b�`(B1)

    HWFBM2HS As String * 1  ' �������@(B2)
    HWFBM2SH As String * 1  ' ������@(B2)
    HWFBM2ST As String * 1  ' ����_��(B2)
    HWFBM2SR As String * 1  ' ���O�̈�(B2)
    HWFBM2NS As String * 2  ' �M�����@(B2)
    HWFBM2SZ As String * 1  ' �������(B2)
    HWFBM2ET As Integer     ' �I���G�b�`(B2)

    HWFBM3HS As String * 1  ' �������@(B3)
    HWFBM3SH As String * 1  ' ������@(B3)
    HWFBM3ST As String * 1  ' ����_��(B3)
    HWFBM3SR As String * 1  ' ���O�̈�(B3)
    HWFBM3NS As String * 2  ' �M�����@(B3)
    HWFBM3SZ As String * 1  ' �������(B3)
    HWFBM3ET As Integer     ' �I���G�b�`(B3)

    HWFOF1HS As String * 1  ' �������@(L1)
    HWFOF1SH As String * 1  ' ������@(L1)
    HWFOF1ST As String * 1  ' ����_��(L1)
    HWFOF1SR As String * 1  ' ���O�̈�(L1)
    HWFOF1NS As String * 2  ' �M�����@(L1)
    HWFOF1SZ As String * 1  ' �������(L1)
    HWFOF1ET As Integer     ' �I���G�b�`(L1)

    HWFOF2HS As String * 1  ' �������@(L2)
    HWFOF2SH As String * 1  ' ������@(L2)
    HWFOF2ST As String * 1  ' ����_��(L2)
    HWFOF2SR As String * 1  ' ���O�̈�(L2)
    HWFOF2NS As String * 2  ' �M�����@(L2)
    HWFOF2SZ As String * 1  ' �������(L2)
    HWFOF2ET As Integer     ' �I���G�b�`(L2)

    HWFOF3HS As String * 1  ' �������@(L3)
    HWFOF3SH As String * 1  ' ������@(L3)
    HWFOF3ST As String * 1  ' ����_��(L3)
    HWFOF3SR As String * 1  ' ���O�̈�(L3)
    HWFOF3NS As String * 2  ' �M�����@(L3)
    HWFOF3SZ As String * 1  ' �������(L3)
    HWFOF3ET As Integer     ' �I���G�b�`(L3)

'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
'''    HWFOF4HS As String * 1  ' �������@(L4)
'''    HWFOF4SH As String * 1  ' ������@(L4)
'''    HWFOF4ST As String * 1  ' ����_��(L4)
'''    HWFOF4SR As String * 1  ' ���O�̈�(L4)
'''    HWFOF4NS As String * 2  ' �M�����@(L4)
'''    HWFOF4SZ As String * 1  ' �������(L4)
'''    HWFOF4ET As Integer     ' �I���G�b�`(L4)
    
    HWFSIRDMX As Integer       '����]�ʏ��(SIRD)
    HWFSIRDSZ As String * 1    '����]�ʑ������(SIRD)
    HWFSIRDHT As String * 1    '����]�ʕۏؕ��@�Q��(SIRD)
    HWFSIRDHS As String * 1    '����]�ʕۏؕ��@�Q��(SIRD)
    HWFSIRDKM As String * 1    '����]�ʌ����p�x�Q��(SIRD)
    HWFSIRDKH As String * 1    '����]�ʌ����p�x�Q��(SIRD)
    HWFSIRDKU As String * 1    '����]�ʌ����p�x�Q�E(SIRD)
    HWFSIRDPS As String * 2    '����]��TB�ۏ؈ʒu(SIRD)
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)

    HWFDSOHS As String * 1  ' �������@(DS)

    HWFMKHWS As String * 1  ' �������@(DZ)
    HWFMKSPH As String * 1  ' ������@(DZ)
    HWFMKSPT As String * 1  ' ����_��(DZ)
    HWFMKSPR As String * 1  ' ���O�̈�(DZ)
    HWFMKNSW As String * 2  ' �M�����@(DZ)
    HWFMKSZY As String * 1  ' �������(DZ)
    HWFMKCET As Integer     ' �I���G�b�`(DZ)

    HWFSPVHS As String * 1  ' �������@(SP/Fe�Z�x)
    HWFSPVSH As String * 1  ' ������@(SP/Fe�Z�x)*
    HWFSPVST As String * 1  ' ����_��(SP/Fe�Z�x)*
    HWFSPVSI As String * 1  ' ����ʒu(SP/Fe�Z�x)*
    HWFDLHWS As String * 1  ' �������@(SP/�g�U��)
    HWFDLSPH As String * 1  ' ������@(SP/�g�U��)*
    HWFDLSPT As String * 1  ' ����_��(SP/�g�U��)*
    HWFDLSPI As String * 1  ' ����ʒu(SP/�g�U��)*
    HWFNRHS  As String * 1  ' �������@(SP/Nr�Z�x)               06/06/08 ooba START ======>
    HWFNRSH  As String * 1  ' ������@(SP/Nr�Z�x)*
    HWFNRST  As String * 1  ' ����_��(SP/Nr�Z�x)*
    HWFNRSI  As String * 1  ' ����ʒu(SP/Nr�Z�x)*
    HWFSPVPUG   As String * 10      ' PUA��(SP/Fe�Z�x)*
    HWFSPVPUR   As String * 10      ' PUA��(SP/Fe�Z�x)*
    HWFSPVSTD   As String * 10      ' �W���΍�(SP/Fe�Z�x)*
    HWFDLPUG    As String * 10      ' PUA��(SP/�g�U��)*
    HWFDLPUR    As String * 10      ' PUA��(SP/�g�U��)*
    HWFNRPUG    As String * 10      ' PUA��(SP/Nr�Z�x)*
    HWFNRPUR    As String * 10      ' PUA��(SP/Nr�Z�x)*
    HWFNRSTD    As String * 10      ' �W���΍�(SP/Nr�Z�x)*      06/06/08 ooba END ========>

    HWFOS1HS As String * 1  ' �������@(D1)
    HWFOS1SH As String * 1  ' ������@(D1)*
    HWFOS1ST As String * 1  ' ����_��(D1)*
    HWFOS1SI As String * 1  ' ����ʒu(D1)*
    HWFOS1NS As String * 2  ' �M�����@(D1)

    HWFOS2HS As String * 1  ' �������@(D2)
    HWFOS2SH As String * 1  ' ������@(D2)*
    HWFOS2ST As String * 1  ' ����_��(D2)*
    HWFOS2SI As String * 1  ' ����ʒu(D2)*
    HWFOS2NS As String * 2  ' �M�����@(D2)

    HWFOS3HS As String * 1  ' �������@(D3)
    HWFOS3SH As String * 1  ' ������@(D3)*
    HWFOS3ST As String * 1  ' ����_��(D3)*
    HWFOS3SI As String * 1  ' ����ʒu(D3)*
    HWFOS3NS As String * 2  ' �M�����@(D3)
    
    HWOTHER1 As String * 1  ' ��������(OT1) '03/05/21
    HWOTHER2 As String * 1  ' ��������(OT1) '03/05/21
    
    HWFZOHWS As String * 1  ' �������@(AO)  ''�ǉ� 03/12/05 ooba START ======>
    HWFZOSPH As String * 1  ' ������@(AO)*
    HWFZOSPT As String * 1  ' ����_��(AO)*
    HWFZOSPI As String * 1  ' ����ʒu(AO)*
    HWFZONSW As String * 2  ' �M�����@(AO)  ''�ǉ� 03/12/05 ooba END ========>
    
    HWFDENHS As String * 1  ' �������@(GD/DEN)  '�ǉ��@05/01/18 ooba START ====>
    HWFLDLHS As String * 1  ' �������@(GD/LDL)
    HWFDVDHS As String * 1  ' �������@(GD/DVD2) '�ǉ��@05/01/18 ooba END ======>
    HWFGDSPH As String * 1  ' ������@(GD)�@    '05/10/25 ooba
    HWFGDSPT As String * 1  ' ����_��(GD)�@    '05/10/25 ooba
    HWFGDZAR As String * 1  ' ���O�̈�(GD)�@    '05/10/25 ooba
    
    HWFRKHNN As String * 1  ' �����p�x_��(Rs)   '�ǉ��@04/04/08 ooba START ====>
    HWFONKHN As String * 1  ' �����p�x_��(Oi)
    HWFOF1KN As String * 1  ' �����p�x_��(L1)
    HWFOF2KN As String * 1  ' �����p�x_��(L2)
    HWFOF3KN As String * 1  ' �����p�x_��(L3)
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
'''    HWFOF4KN As String * 1  ' �����p�x_��(L4)
    HWFSIRDKN As String * 1  ' �����p�x_��(SIRD)
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
    HWFBM1KN As String * 1  ' �����p�x_��(B1)
    HWFBM2KN As String * 1  ' �����p�x_��(B2)
    HWFBM3KN As String * 1  ' �����p�x_��(B3)
    HWFOS1KN As String * 1  ' �����p�x_��(D1)
    HWFOS2KN As String * 1  ' �����p�x_��(D2)
    HWFOS3KN As String * 1  ' �����p�x_��(D3)
    HWFDSOKN As String * 1  ' �����p�x_��(DS)
    HWFMKKHN As String * 1  ' �����p�x_��(DZ)
    HWFSPVKN As String * 1  ' �����p�x_��(SP/Fe�Z�x)
    HWFDLKHN As String * 1  ' �����p�x_��(SP/�g�U��)
    HWFZOKHN As String * 1  ' �����p�x_��(AO)   '�ǉ��@04/04/08 ooba END ======>
    HWFGDKHN As String * 1  ' �����p�x_��(GD)�@05/01/18 ooba
    HWFNRKN  As String * 1  ' �����p�x_��(SP/Nr�Z�x)  06/06/08 ooba
    
    HWFIGKBN As String * 1  ' IG�敪
    HWFANTNP As Integer     ' DK�A�j�[������(���x)
    HWFANTIM As Integer     ' DK�A�j�[������(����)
    HWFANGZY As String * 1  ' DK�A�j�[������(�K�X)�@04/07/23 ooba
    
    HWOTHER1MAI As String * 1  ' �T���v������(OT1) '04/06/23
    HWOTHER2MAI As String * 1  ' �T���v������(OT2) '04/06/23

''Upd Start (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�
    HWFGDLINE   As String * 3   '�iWFGDײݐ�(TBCME036)
''Upd End   (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    HEPOF1NS As String * 2  ' �i�M�����@(OSF1E)
    HEPOF1SZ As String * 1  ' �i�������(OSF1E)
    HEPOF1ET As Integer     ' �i�I��ET��(OSF1E)
    HEPOF1HS As String * 1  ' �i�ۏؕ��@_��(OSF1E)
    HEPOF1SH As String * 1  ' �i����ʒu_��(OSF1E)
    HEPOF1ST As String * 1  ' �i����ʒu_�_(OSF1E)
    HEPOF1SR As String * 1  ' �i����ʒu_��(OSF1E)
    HEPOF1KN As String * 1  ' �i�����p�x_��(OSF1E)
    HEPOF2NS As String * 2  ' �i�M�����@(OSF2E)
    HEPOF2SZ As String * 1  ' �i�������(OSF2E)
    HEPOF2ET As Integer     ' �i�I��ET��(OSF2E)
    HEPOF2HS As String * 1  ' �i�ۏؕ��@_��(OSF2E)
    HEPOF2SH As String * 1  ' �i����ʒu_��(OSF2E)
    HEPOF2ST As String * 1  ' �i����ʒu_�_(OSF2E)
    HEPOF2SR As String * 1  ' �i����ʒu_��(OSF2E)
    HEPOF2KN As String * 1  ' �i�����p�x_��(OSF2E)
    HEPOF3NS As String * 2  ' �i�M�����@(OSF3E)
    HEPOF3SZ As String * 1  ' �i�������(OSF3E)
    HEPOF3ET As Integer     ' �i�I��ET��(OSF3E)
    HEPOF3HS As String * 1  ' �i�ۏؕ��@_��(OSF3E)
    HEPOF3SH As String * 1  ' �i����ʒu_��(OSF3E)
    HEPOF3ST As String * 1  ' �i����ʒu_�_(OSF3E)
    HEPOF3SR As String * 1  ' �i����ʒu_��(OSF3E)
    HEPOF3KN As String * 1  ' �i�����p�x_��(OSF3E)
    HEPBM1NS As String * 2  ' �i�M�����@(BMD1E)
    HEPBM1SZ As String * 1  ' �i�������(BMD1E)
    HEPBM1ET As Integer     ' �i�I��ET��(BMD1E)
    HEPBM1HS As String * 1  ' �i�ۏؕ��@_��(BMD1E)
    HEPBM1SH As String * 1  ' �i����ʒu_��(BMD1E)
    HEPBM1ST As String * 1  ' �i����ʒu_�_(BMD1E)
    HEPBM1SR As String * 1  ' �i����ʒu_��(BMD1E)
    HEPBM1KN As String * 1  ' �i�����p�x_��(BMD1E)
    HEPBM2NS As String * 2  ' �i�M�����@(BMD2E)
    HEPBM2SZ As String * 1  ' �i�������(BMD2E)
    HEPBM2ET As Integer     ' �i�I��ET��(BMD2E)
    HEPBM2HS As String * 1  ' �i�ۏؕ��@_��(BMD2E)
    HEPBM2SH As String * 1  ' �i����ʒu_��(BMD2E)
    HEPBM2ST As String * 1  ' �i����ʒu_�_(BMD2E)
    HEPBM2SR As String * 1  ' �i����ʒu_��(BMD2E)
    HEPBM2KN As String * 1  ' �i�����p�x_��(BMD2E)
    HEPBM3NS As String * 2  ' �i�M�����@(BMD3E)
    HEPBM3SZ As String * 1  ' �i�������(BMD3E)
    HEPBM3ET As Integer     ' �i�I��ET��(BMD3E)
    HEPBM3HS As String * 1  ' �i�ۏؕ��@_��(BMD3E)
    HEPBM3SH As String * 1  ' �i����ʒu_��(BMD3E)
    HEPBM3ST As String * 1  ' �i����ʒu_�_(BMD3E)
    HEPBM3SR As String * 1  ' �i����ʒu_��(BMD3E)
    HEPBM3KN As String * 1  ' �i�����p�x_��(BMD3E)
    HEPACEN  As Double      ' �iE1�����S
    HEPANTNP As Integer     ' �iEPAN���x
    HEPANTIM As Integer     ' �iEPAN����
    HEPIGKBN As String * 1  ' �iEPIG�敪
    HEPANGZY As String * 1  ' �iEP����AN�K�X����
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga START   ---
    HWFGDSZY As String * 1  ' �i�v�e�f�c�������
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga END     ---
End Type

' WF�T���v���e�[�u��
Public Type typ_WFSample
    CRYINDRS As String * 1  ' ��������(Rs)
    CRYINDOI As String * 1  ' ��������(Oi)
    CRYINDB1 As String * 1  ' ��������(B1)
    CRYINDB2 As String * 1  ' ��������(B2�j
    CRYINDB3 As String * 1  ' ��������(B3)
    CRYINDL1 As String * 1  ' ��������(L1)
    CRYINDL2 As String * 1  ' ��������(L2)
    CRYINDL3 As String * 1  ' ��������(L3)
    CRYINDL4 As String * 1  ' ��������(L4)
    CRYINDDS As String * 1  ' ��������(DS)
    CRYINDDZ As String * 1  ' ��������(DZ)
    CRYINDSP As String * 1  ' ��������(SP)
    CRYINDD1 As String * 1  ' ��������(D1)
    CRYINDD2 As String * 1  ' ��������(D2)
    CRYINDD3 As String * 1  ' ��������(D3)
    CRYINDOT1 As String * 1 ' ��������(OT1) 'Add.03/05/20
    CRYINDOT2 As String * 1 ' ��������(OT2) 'Add.03/05/20
    CRYINDAO As String * 1  ' �����L��(AO)  '�ǉ� 03/12/05 ooba
    CRYINDGD As String * 1  ' �����L��(GD)  '�ǉ� 05/01/18 ooba
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga START   ---
    CRYINDGD2 As String * 1  ' �����L��(GD��������p)
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga END     ---
    WFHSGD As String * 1    ' �ۏ�FLG(GD)   '�ǉ� 05/01/18 ooba
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    EPIINDL1 As String * 1 ' �����L��(OSF1E)
    EPIINDL2 As String * 1 ' �����L��(OSF2E)
    EPIINDL3 As String * 1 ' �����L��(OSF3E)
    EPIINDB1 As String * 1 ' �����L��(BMD1E)
    EPIINDB2 As String * 1 ' �����L��(BMD2E)
    EPIINDB3 As String * 1 ' �����L��(BMD3E)
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
End Type

'�T�v      :���i�d�lWF�f�[�^�̎擾�h���C�o
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'�@�@      :pSpWFSamp�@�@�@,IO ,typ_SpWFSamp   �@,WF�T���v���d�l
'�@�@      :�߂�l         ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function scmzc_getWF(pSpWFSamp As typ_SpWFSamp) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim sOT1    As String '03/05/21 �㓡
    Dim sOT2    As String '03/05/21 �㓡
    Dim sMAI1    As String '04/06/23
    Dim sMAI2    As String '04/06/23
    Dim rtn     As FUNCTION_RETURN
    '' �G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function scmzc_getWF"

    '' ���i�d�l�̎擾
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
'''    sql = "select " & _
'''          "E021HWFRSPOH, E021HWFRSPOT, E021HWFRSPOI, E021HWFRHWYS, E024HWFMKSPH, " & _
'''          "E024HWFMKSPT, E024HWFMKSPR, E024HWFMKHWS, E024HWFMKSZY, E024HWFMKNSW, " & _
'''          "E024HWFMKCET, E025HWFONSPH, E025HWFONSPT, E025HWFONSPI, E025HWFONHWS, " & _
'''          "E025HWFONKWY, E025HWFOS1NS, E025HWFOS1SH, E025HWFOS1ST, E025HWFOS1SI, " & _
'''          "E025HWFOS1HS, E025HWFOS2NS, E025HWFOS2SH, E025HWFOS2ST, E025HWFOS2SI, " & _
'''          "E025HWFOS2HS, E025HWFOS3NS, E025HWFOS3SH, E025HWFOS3ST, E025HWFOS3SI, " & _
'''          "E025HWFOS3HS, E025HWFANTNP, E025HWFANTIM, E026HWFDSOHS, E028HWFSPVSH, " & _
'''          "E028HWFSPVST, E028HWFSPVSI, E028HWFSPVHS, E028HWFDLSPH, E028HWFDLSPT, " & _
'''          "E028HWFDLSPI, E028HWFDLHWS, E029HWFOF1ET, E029HWFOF1NS, E029HWFOF1SZ, " & _
'''          "E029HWFOF1SH, E029HWFOF1ST, E029HWFOF1SR, E029HWFOF1HS, E029HWFOF2ET, " & _
'''          "E029HWFOF2NS, E029HWFOF2SZ, E029HWFOF2SH, E029HWFOF2ST, E029HWFOF2SR, " & _
'''          "E029HWFOF2HS, E029HWFOF3ET, E029HWFOF3NS, E029HWFOF3SZ, E029HWFOF3SH, " & _
'''          "E029HWFOF3ST, E029HWFOF3SR, E029HWFOF3HS, E029HWFOF4ET, E029HWFOF4NS, " & _
'''          "E029HWFOF4SZ, E029HWFOF4SH, E029HWFOF4ST, E029HWFOF4SR, E029HWFOF4HS, " & _
'''          "E029HWFBM1ET, E029HWFBM1NS, E029HWFBM1SZ, E029HWFBM1SH, E029HWFBM1ST, " & _
'''          "E029HWFBM1SR, E029HWFBM1HS, E029HWFBM2ET, E029HWFBM2NS, E029HWFBM2SZ, " & _
'''          "E029HWFBM2SH, E029HWFBM2ST, E029HWFBM2SR, E029HWFBM2HS, E029HWFBM3ET, " & _
'''          "E029HWFBM3NS, E029HWFBM3SZ, E029HWFBM3SH, E029HWFBM3ST, E029HWFBM3SR, E029HWFBM3HS" & _
'''          " from VECME001" & _
'''          " where E018HINBAN='" & pSpWFSamp.hin.hinban & "' and E018MNOREVNO=" & pSpWFSamp.hin.mnorevno & _
'''          " and E018FACTORY='" & pSpWFSamp.hin.FACTORY & "' and E018OPECOND='" & pSpWFSamp.hin.OPECOND & "'"
          
    sql = "select " & _
          "E021HWFRSPOH, E021HWFRSPOT, E021HWFRSPOI, E021HWFRHWYS, E024HWFMKSPH, " & _
          "E024HWFMKSPT, E024HWFMKSPR, E024HWFMKHWS, E024HWFMKSZY, E024HWFMKNSW, " & _
          "E024HWFMKCET, E025HWFONSPH, E025HWFONSPT, E025HWFONSPI, E025HWFONHWS, " & _
          "E025HWFONKWY, E025HWFOS1NS, E025HWFOS1SH, E025HWFOS1ST, E025HWFOS1SI, " & _
          "E025HWFOS1HS, E025HWFOS2NS, E025HWFOS2SH, E025HWFOS2ST, E025HWFOS2SI, " & _
          "E025HWFOS2HS, E025HWFOS3NS, E025HWFOS3SH, E025HWFOS3ST, E025HWFOS3SI, " & _
          "E025HWFOS3HS, E025HWFANTNP, E025HWFANTIM, E026HWFDSOHS, E028HWFSPVSH, " & _
          "E028HWFSPVST, E028HWFSPVSI, E028HWFSPVHS, E028HWFDLSPH, E028HWFDLSPT, " & _
          "E028HWFDLSPI, E028HWFDLHWS, E029HWFOF1ET, E029HWFOF1NS, E029HWFOF1SZ, " & _
          "E029HWFOF1SH, E029HWFOF1ST, E029HWFOF1SR, E029HWFOF1HS, E029HWFOF2ET, " & _
          "E029HWFOF2NS, E029HWFOF2SZ, E029HWFOF2SH, E029HWFOF2ST, E029HWFOF2SR, " & _
          "E029HWFOF2HS, E029HWFOF3ET, E029HWFOF3NS, E029HWFOF3SZ, E029HWFOF3SH, " & _
          "E029HWFOF3ST, E029HWFOF3SR, E029HWFOF3HS, " & _
          "E029HWFBM1ET, E029HWFBM1NS, E029HWFBM1SZ, E029HWFBM1SH, E029HWFBM1ST, " & _
          "E029HWFBM1SR, E029HWFBM1HS, E029HWFBM2ET, E029HWFBM2NS, E029HWFBM2SZ, " & _
          "E029HWFBM2SH, E029HWFBM2ST, E029HWFBM2SR, E029HWFBM2HS, E029HWFBM3ET, " & _
          "E029HWFBM3NS, E029HWFBM3SZ, E029HWFBM3SH, E029HWFBM3ST, E029HWFBM3SR, E029HWFBM3HS" & _
          " from VECME001" & _
          " where E018HINBAN='" & pSpWFSamp.hin.hinban & "' and E018MNOREVNO=" & pSpWFSamp.hin.mnorevno & _
          " and E018FACTORY='" & pSpWFSamp.hin.factory & "' and E018OPECOND='" & pSpWFSamp.hin.opecond & "'"
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If

    With pSpWFSamp
        .HWFRSPOH = rs("E021HWFRSPOH")
        .HWFRSPOT = rs("E021HWFRSPOT")
        .HWFRSPOI = rs("E021HWFRSPOI")
        .HWFRHWYS = rs("E021HWFRHWYS")
        .HWFMKSPH = rs("E024HWFMKSPH")
        .HWFMKSPT = rs("E024HWFMKSPT")
        .HWFMKSPR = rs("E024HWFMKSPR")
        .HWFMKHWS = rs("E024HWFMKHWS")
        .HWFMKSZY = rs("E024HWFMKSZY")
        .HWFMKNSW = rs("E024HWFMKNSW")
        .HWFMKCET = fncNullCheck(rs("E024HWFMKCET"))
        .HWFONSPH = rs("E025HWFONSPH")
        .HWFONSPT = rs("E025HWFONSPT")
        .HWFONSPI = rs("E025HWFONSPI")
        .HWFONHWS = rs("E025HWFONHWS")
        .HWFONKWY = rs("E025HWFONKWY")
        .HWFOS1NS = rs("E025HWFOS1NS")
        .HWFOS1HS = rs("E025HWFOS1HS")
        .HWFOS1SH = rs("E025HWFOS1SH")
        .HWFOS1ST = rs("E025HWFOS1ST")
        .HWFOS1SI = rs("E025HWFOS1SI")
        .HWFOS2NS = rs("E025HWFOS2NS")
        .HWFOS2SH = rs("E025HWFOS2SH")
        .HWFOS2ST = rs("E025HWFOS2ST")
        .HWFOS2SI = rs("E025HWFOS2SI")
        .HWFOS2HS = rs("E025HWFOS2HS")
        .HWFOS3NS = rs("E025HWFOS3NS")
        .HWFOS3SH = rs("E025HWFOS3SH")
        .HWFOS3ST = rs("E025HWFOS3ST")
        .HWFOS3SI = rs("E025HWFOS3SI")
        .HWFOS3HS = rs("E025HWFOS3HS")
        .HWFANTNP = fncNullCheck(rs("E025HWFANTNP"))
        .HWFANTIM = fncNullCheck(rs("E025HWFANTIM"))
        .HWFDSOHS = rs("E026HWFDSOHS")
        .HWFSPVSH = rs("E028HWFSPVSH")
        .HWFSPVST = rs("E028HWFSPVST")
        .HWFSPVSI = rs("E028HWFSPVSI")
        .HWFSPVHS = rs("E028HWFSPVHS")
        .HWFDLSPH = rs("E028HWFDLSPH")
        .HWFDLSPT = rs("E028HWFDLSPT")
        .HWFDLSPI = rs("E028HWFDLSPI")
        .HWFDLHWS = rs("E028HWFDLHWS")
        .HWFOF1ET = fncNullCheck(rs("E029HWFOF1ET"))
        .HWFOF1NS = rs("E029HWFOF1NS")
        .HWFOF1SZ = rs("E029HWFOF1SZ")
        .HWFOF1SH = rs("E029HWFOF1SH")
        .HWFOF1ST = rs("E029HWFOF1ST")
        .HWFOF1SR = rs("E029HWFOF1SR")
        .HWFOF1HS = rs("E029HWFOF1HS")
        .HWFOF2ET = fncNullCheck(rs("E029HWFOF2ET"))
        .HWFOF2NS = rs("E029HWFOF2NS")
        .HWFOF2SZ = rs("E029HWFOF2SZ")
        .HWFOF2SH = rs("E029HWFOF2SH")
        .HWFOF2ST = rs("E029HWFOF2ST")
        .HWFOF2SR = rs("E029HWFOF2SR")
        .HWFOF2HS = rs("E029HWFOF2HS")
        .HWFOF3ET = fncNullCheck(rs("E029HWFOF3ET"))
        .HWFOF3NS = rs("E029HWFOF3NS")
        .HWFOF3SZ = rs("E029HWFOF3SZ")
        .HWFOF3SH = rs("E029HWFOF3SH")
        .HWFOF3ST = rs("E029HWFOF3ST")
        .HWFOF3SR = rs("E029HWFOF3SR")
        .HWFOF3HS = rs("E029HWFOF3HS")
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
'''        .HWFOF4ET = fncNullCheck(rs("E029HWFOF4ET"))
'''        .HWFOF4NS = rs("E029HWFOF4NS")
'''        .HWFOF4SZ = rs("E029HWFOF4SZ")
'''        .HWFOF4SH = rs("E029HWFOF4SH")
'''        .HWFOF4ST = rs("E029HWFOF4ST")
'''        .HWFOF4SR = rs("E029HWFOF4SR")
'''        .HWFOF4HS = rs("E029HWFOF4HS")
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
        .HWFBM1ET = fncNullCheck(rs("E029HWFBM1ET"))
        .HWFBM1NS = rs("E029HWFBM1NS")
        .HWFBM1SZ = rs("E029HWFBM1SZ")
        .HWFBM1SH = rs("E029HWFBM1SH")
        .HWFBM1ST = rs("E029HWFBM1ST")
        .HWFBM1SR = rs("E029HWFBM1SR")
        .HWFBM1HS = rs("E029HWFBM1HS")
        .HWFBM2ET = fncNullCheck(rs("E029HWFBM2ET"))
        .HWFBM2NS = rs("E029HWFBM2NS")
        .HWFBM2SZ = rs("E029HWFBM2SZ")
        .HWFBM2SH = rs("E029HWFBM2SH")
        .HWFBM2ST = rs("E029HWFBM2ST")
        .HWFBM2SR = rs("E029HWFBM2SR")
        .HWFBM2HS = rs("E029HWFBM2HS")
        .HWFBM3ET = fncNullCheck(rs("E029HWFBM3ET"))
        .HWFBM3NS = rs("E029HWFBM3NS")
        .HWFBM3SZ = rs("E029HWFBM3SZ")
        .HWFBM3SH = rs("E029HWFBM3SH")
        .HWFBM3ST = rs("E029HWFBM3ST")
        .HWFBM3SR = rs("E029HWFBM3SR")
        .HWFBM3HS = rs("E029HWFBM3HS")
        'rtn = scmzc_getE036(pSpWFSamp.HIN, sOT1, sOT2)   2004/06/23
        'rtn = scmzc_getE036(pSpWFSamp.HIN, sOT1, sOT2)    '2004/07/12 koyama update
        rtn = scmzc_getE036(pSpWFSamp.hin, sOT1, sOT2, sMAI1, sMAI2) '2004/07/12 koyama update
        If rtn = FUNCTION_RETURN_FAILURE Then
            rs.Close
            scmzc_getWF = FUNCTION_RETURN_FAILURE
            GoTo PROC_EXIT
        End If
        .HWOTHER1 = sOT1 '### 03/05/20
        .HWOTHER2 = sOT2
        .HWOTHER1MAI = sMAI1   '04/06/23
        .HWOTHER2MAI = sMAI2   '04/06/23
    End With
    rs.Close
    
    '�����p�x_���ް��擾�@04/04/08 ooba START ==========================================>
    sql = "select "
    sql = sql & "TBCME026.HWFGDKHN, "   '�����p�x_��(GD)�@05/01/18 ooba
    sql = sql & "TBCME024.HWFANGZY, "   '�i�v�e�����`�m�K�X�����@04/07/23 ooba
    sql = sql & "TBCME021.HWFRKHNN, "
    sql = sql & "TBCME025.HWFONKHN, "
    sql = sql & "TBCME029.HWFOF1KN, "
    sql = sql & "TBCME029.HWFOF2KN, "
    sql = sql & "TBCME029.HWFOF3KN, "
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
'''    sql = sql & "TBCME029.HWFOF4KN, "
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
    sql = sql & "TBCME029.HWFBM1KN, "
    sql = sql & "TBCME029.HWFBM2KN, "
    sql = sql & "TBCME029.HWFBM3KN, "
    sql = sql & "TBCME025.HWFOS1KN, "
    sql = sql & "TBCME025.HWFOS2KN, "
    sql = sql & "TBCME025.HWFOS3KN, "
    sql = sql & "TBCME026.HWFDSOKN, "
    sql = sql & "TBCME024.HWFMKKHN, "
    sql = sql & "TBCME028.HWFSPVKN, "
    sql = sql & "TBCME028.HWFDLKHN, "
    sql = sql & "TBCME025.HWFZOKHN "
    sql = sql & "from TBCME021, TBCME024, TBCME025, TBCME026, TBCME028, TBCME029 "
    sql = sql & "where TBCME021.HINBAN = TBCME024.HINBAN "
    sql = sql & "and TBCME021.MNOREVNO = TBCME024.MNOREVNO "
    sql = sql & "and TBCME021.FACTORY = TBCME024.FACTORY "
    sql = sql & "and TBCME021.OPECOND = TBCME024.OPECOND "
    sql = sql & "and TBCME021.HINBAN = TBCME025.HINBAN "
    sql = sql & "and TBCME021.MNOREVNO = TBCME025.MNOREVNO "
    sql = sql & "and TBCME021.FACTORY = TBCME025.FACTORY "
    sql = sql & "and TBCME021.OPECOND = TBCME025.OPECOND "
    sql = sql & "and TBCME021.HINBAN = TBCME026.HINBAN "
    sql = sql & "and TBCME021.MNOREVNO = TBCME026.MNOREVNO "
    sql = sql & "and TBCME021.FACTORY = TBCME026.FACTORY "
    sql = sql & "and TBCME021.OPECOND = TBCME026.OPECOND "
    sql = sql & "and TBCME021.HINBAN = TBCME028.HINBAN "
    sql = sql & "and TBCME021.MNOREVNO = TBCME028.MNOREVNO "
    sql = sql & "and TBCME021.FACTORY = TBCME028.FACTORY "
    sql = sql & "and TBCME021.OPECOND = TBCME028.OPECOND "
    sql = sql & "and TBCME021.HINBAN = TBCME029.HINBAN "
    sql = sql & "and TBCME021.MNOREVNO = TBCME029.MNOREVNO "
    sql = sql & "and TBCME021.FACTORY = TBCME029.FACTORY "
    sql = sql & "and TBCME021.OPECOND = TBCME029.OPECOND "
    sql = sql & "and TBCME021.HINBAN = '" & pSpWFSamp.hin.hinban & "' "
    sql = sql & "and TBCME021.MNOREVNO = " & pSpWFSamp.hin.mnorevno & " "
    sql = sql & "and TBCME021.FACTORY = '" & pSpWFSamp.hin.factory & "' "
    sql = sql & "and TBCME021.OPECOND = '" & pSpWFSamp.hin.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    
    With pSpWFSamp
        If IsNull(rs("HWFGDKHN")) = False Then .HWFGDKHN = rs("HWFGDKHN") Else .HWFGDKHN = " "  '05/01/18 ooba
        If IsNull(rs("HWFANGZY")) = False Then .HWFANGZY = rs("HWFANGZY") Else .HWFANGZY = " "  '04/07/23 ooba
        If IsNull(rs("HWFRKHNN")) = False Then .HWFRKHNN = rs("HWFRKHNN") Else .HWFRKHNN = " "
        If IsNull(rs("HWFONKHN")) = False Then .HWFONKHN = rs("HWFONKHN") Else .HWFONKHN = " "
        If IsNull(rs("HWFOF1KN")) = False Then .HWFOF1KN = rs("HWFOF1KN") Else .HWFOF1KN = " "
        If IsNull(rs("HWFOF2KN")) = False Then .HWFOF2KN = rs("HWFOF2KN") Else .HWFOF2KN = " "
        If IsNull(rs("HWFOF3KN")) = False Then .HWFOF3KN = rs("HWFOF3KN") Else .HWFOF3KN = " "
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
'''        If IsNull(rs("HWFOF4KN")) = False Then .HWFOF4KN = rs("HWFOF4KN") Else .HWFOF4KN = " "
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
        If IsNull(rs("HWFBM1KN")) = False Then .HWFBM1KN = rs("HWFBM1KN") Else .HWFBM1KN = " "
        If IsNull(rs("HWFBM2KN")) = False Then .HWFBM2KN = rs("HWFBM2KN") Else .HWFBM2KN = " "
        If IsNull(rs("HWFBM3KN")) = False Then .HWFBM3KN = rs("HWFBM3KN") Else .HWFBM3KN = " "
        If IsNull(rs("HWFOS1KN")) = False Then .HWFOS1KN = rs("HWFOS1KN") Else .HWFOS1KN = " "
        If IsNull(rs("HWFOS2KN")) = False Then .HWFOS2KN = rs("HWFOS2KN") Else .HWFOS2KN = " "
        If IsNull(rs("HWFOS3KN")) = False Then .HWFOS3KN = rs("HWFOS3KN") Else .HWFOS3KN = " "
        If IsNull(rs("HWFDSOKN")) = False Then .HWFDSOKN = rs("HWFDSOKN") Else .HWFDSOKN = " "
        If IsNull(rs("HWFMKKHN")) = False Then .HWFMKKHN = rs("HWFMKKHN") Else .HWFMKKHN = " "
        If IsNull(rs("HWFSPVKN")) = False Then .HWFSPVKN = rs("HWFSPVKN") Else .HWFSPVKN = " "
        If IsNull(rs("HWFDLKHN")) = False Then .HWFDLKHN = rs("HWFDLKHN") Else .HWFDLKHN = " "
        If IsNull(rs("HWFZOKHN")) = False Then .HWFZOKHN = rs("HWFZOKHN") Else .HWFZOKHN = " "
    End With
    rs.Close
    '�����p�x_���ް��擾�@04/04/08 ooba END ============================================>
    
    '' �c���_�f�d�l�擾�@03/12/05 ooba START ===========================================>
    sql = "select HWFZOHWS, HWFZOSPH, HWFZOSPT, HWFZOSPI, HWFZONSW from TBCME025 "
    sql = sql & "where HINBAN = '" & pSpWFSamp.hin.hinban & "' "
    sql = sql & "and MNOREVNO = " & pSpWFSamp.hin.mnorevno & " "
    sql = sql & "and FACTORY = '" & pSpWFSamp.hin.factory & "' "
    sql = sql & "and OPECOND = '" & pSpWFSamp.hin.opecond & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    
    If IsNull(rs("HWFZOHWS")) = False Then pSpWFSamp.HWFZOHWS = rs("HWFZOHWS") Else pSpWFSamp.HWFZOHWS = " "
    If IsNull(rs("HWFZOSPH")) = False Then pSpWFSamp.HWFZOSPH = rs("HWFZOSPH") Else pSpWFSamp.HWFZOSPH = " "
    If IsNull(rs("HWFZOSPT")) = False Then pSpWFSamp.HWFZOSPT = rs("HWFZOSPT") Else pSpWFSamp.HWFZOSPT = " "
    If IsNull(rs("HWFZOSPI")) = False Then pSpWFSamp.HWFZOSPI = rs("HWFZOSPI") Else pSpWFSamp.HWFZOSPI = " "
    If IsNull(rs("HWFZONSW")) = False Then pSpWFSamp.HWFZONSW = rs("HWFZONSW") Else pSpWFSamp.HWFZONSW = " "
    
    rs.Close
    '' �c���_�f�d�l�擾�@03/12/05 ooba END =============================================>
    
    '' GD�d�l�擾�@05/01/18 ooba START ================================================>
    
''Upd start (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�
'    sql = "select "
'    sql = sql & "HWFDENHS, "        '�������@(GD/DEN)
'    sql = sql & "HWFLDLHS, "        '�������@(GD/LDL)
'    sql = sql & "HWFDVDHS "         '�������@(GD/DVD2)
'    sql = sql & "from TBCME026 "
'    sql = sql & "where HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
'    sql = sql & "and MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
'    sql = sql & "and FACTORY = '" & pSpWFSamp.HIN.factory & "' "
'    sql = sql & "and OPECOND = '" & pSpWFSamp.HIN.opecond & "' "
    sql = "select "
    sql = sql & "T1.HWFGDSPH AS HWFGDSPH, "         '������@(GD)�@05/10/25 ooba
    sql = sql & "T1.HWFGDSPT AS HWFGDSPT, "         '����_��(GD)�@05/10/25 ooba
    sql = sql & "T1.HWFGDZAR AS HWFGDZAR, "         '���O�̈�(GD)�@05/10/25 ooba
    sql = sql & "T1.HWFDENHS AS HWFDENHS, "         '�������@(GD/DEN)
    sql = sql & "T1.HWFLDLHS AS HWFLDLHS, "         '�������@(GD/LDL)
    sql = sql & "T1.HWFDVDHS AS HWFDVDHS"           '�������@(GD/DVD2)
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga START   ---
    sql = sql & ",T1.HWFGDSZY AS HWFGDSZY"          '�������(GD)
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga END     ---
    sql = sql & ",T2.HWFGDLINE AS HWFGDLINE "       'ײݐ�
    sql = sql & "from TBCME026 T1,TBCME036 T2 "
    sql = sql & "where T1.HINBAN = '" & pSpWFSamp.hin.hinban & "' "
    sql = sql & "and T1.MNOREVNO = " & pSpWFSamp.hin.mnorevno & " "
    sql = sql & "and T1.FACTORY = '" & pSpWFSamp.hin.factory & "' "
    sql = sql & "and T1.OPECOND = '" & pSpWFSamp.hin.opecond & "' "
    sql = sql & "and T1.HINBAN = T2.HINBAN "
    sql = sql & "and T1.MNOREVNO = T2.MNOREVNO "
    sql = sql & "and T1.FACTORY = T2.FACTORY "
    sql = sql & "and T1.OPECOND = T2.OPECOND "
''Upd end   (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    
    If IsNull(rs("HWFGDSPH")) = False Then pSpWFSamp.HWFGDSPH = rs("HWFGDSPH") Else pSpWFSamp.HWFGDSPH = " "  '05/10/25 ooba
    If IsNull(rs("HWFGDSPT")) = False Then pSpWFSamp.HWFGDSPT = rs("HWFGDSPT") Else pSpWFSamp.HWFGDSPT = " "  '05/10/25 ooba
    If IsNull(rs("HWFGDZAR")) = False Then pSpWFSamp.HWFGDZAR = rs("HWFGDZAR") Else pSpWFSamp.HWFGDZAR = " "  '05/10/25 ooba
    If IsNull(rs("HWFDENHS")) = False Then pSpWFSamp.HWFDENHS = rs("HWFDENHS") Else pSpWFSamp.HWFDENHS = " "
    If IsNull(rs("HWFLDLHS")) = False Then pSpWFSamp.HWFLDLHS = rs("HWFLDLHS") Else pSpWFSamp.HWFLDLHS = " "
    If IsNull(rs("HWFDVDHS")) = False Then pSpWFSamp.HWFDVDHS = rs("HWFDVDHS") Else pSpWFSamp.HWFDVDHS = " "
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga START   ---
    If IsNull(rs("HWFGDSZY")) = False Then pSpWFSamp.HWFGDSZY = rs("HWFGDSZY") Else pSpWFSamp.HWFGDSZY = " "
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga END     ---
''Upd Start (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�
    If IsNull(rs("HWFGDLINE")) = False Then pSpWFSamp.HWFGDLINE = CStr(rs("HWFGDLINE"))
''Upd End   (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�
    
    rs.Close
    '' GD�d�l�擾�@05/01/18 ooba END ==================================================>
    
    '' SPV�d�l�擾�@06/06/08 ooba START ===============================================>
    sql = "select HWFNRHS, "                    '�iWFSPVNR�ۏؕ��@_��
    sql = sql & "HWFNRSH, "                     '�iWFSPVNR����ʒu_��
    sql = sql & "HWFNRST, "                     '�iWFSPVNR����ʒu_�_
    sql = sql & "HWFNRSI, "                     '�iWFSPVNR����ʒu_��
    sql = sql & "HWFNRKN, "                     '�iWFSPVNR�����p�x_��
    sql = sql & "HWFSPVPUG, "                   '�iWFSPVFEPUA��
    sql = sql & "HWFSPVPUR, "                   '�iWFSPVFEPUA��
    sql = sql & "HWFSPVSTD, "                   '�iWFSPVFE�W���΍�
    sql = sql & "HWFDLPUG, "                    '�iWF�g�U��PUA��
    sql = sql & "HWFDLPUR, "                    '�iWF�g�U��PUA��
    sql = sql & "HWFNRPUG, "                    '�iWFSPVNRPUA��
    sql = sql & "HWFNRPUR, "                    '�iWFSPVNRPUA��
    sql = sql & "HWFNRSTD "                     '�iWFSPVNR�W���΍�
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
    sql = sql & ",HWFSIRDMX, "                  '����]�ʏ��
    sql = sql & "HWFSIRDSZ, "                   '����]�ʑ������
    sql = sql & "HWFSIRDHT, "                   '����]�ʕۏؕ��@�Q��
    sql = sql & "HWFSIRDHS, "                   '����]�ʕۏؕ��@_��
    sql = sql & "HWFSIRDKM, "                   '����]�ʌ����p�x�Q��
    sql = sql & "HWFSIRDKN, "                   '����]�ʌ����p�x_��
    sql = sql & "HWFSIRDKH, "                   '����]�ʌ����p�x�Q��
    sql = sql & "HWFSIRDKU, "                   '����]�ʌ����p�x�Q�E
    sql = sql & "HWFSIRDPS  "                   '����]��TB�ۏ؈ʒu
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
    sql = sql & "from TBCME048 "
    sql = sql & "where HINBAN = '" & pSpWFSamp.hin.hinban & "' "
    sql = sql & "and MNOREVNO = " & pSpWFSamp.hin.mnorevno & " "
    sql = sql & "and FACTORY = '" & pSpWFSamp.hin.factory & "' "
    sql = sql & "and OPECOND = '" & pSpWFSamp.hin.opecond & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    
    If IsNull(rs("HWFNRHS")) = False Then pSpWFSamp.HWFNRHS = rs("HWFNRHS") Else pSpWFSamp.HWFNRHS = " "
    If IsNull(rs("HWFNRSH")) = False Then pSpWFSamp.HWFNRSH = rs("HWFNRSH") Else pSpWFSamp.HWFNRSH = " "
    If IsNull(rs("HWFNRST")) = False Then pSpWFSamp.HWFNRST = rs("HWFNRST") Else pSpWFSamp.HWFNRST = " "
    If IsNull(rs("HWFNRSI")) = False Then pSpWFSamp.HWFNRSI = rs("HWFNRSI") Else pSpWFSamp.HWFNRSI = " "
    If IsNull(rs("HWFNRKN")) = False Then pSpWFSamp.HWFNRKN = rs("HWFNRKN") Else pSpWFSamp.HWFNRKN = " "
    If IsNull(rs("HWFSPVPUG")) = False Then pSpWFSamp.HWFSPVPUG = rs("HWFSPVPUG") Else pSpWFSamp.HWFSPVPUG = " "
    If IsNull(rs("HWFSPVPUR")) = False Then pSpWFSamp.HWFSPVPUR = rs("HWFSPVPUR") Else pSpWFSamp.HWFSPVPUR = " "
    If IsNull(rs("HWFSPVSTD")) = False Then pSpWFSamp.HWFSPVSTD = rs("HWFSPVSTD") Else pSpWFSamp.HWFSPVSTD = " "
    If IsNull(rs("HWFDLPUG")) = False Then pSpWFSamp.HWFDLPUG = rs("HWFDLPUG") Else pSpWFSamp.HWFDLPUG = " "
    If IsNull(rs("HWFDLPUR")) = False Then pSpWFSamp.HWFDLPUR = rs("HWFDLPUR") Else pSpWFSamp.HWFDLPUR = " "
    If IsNull(rs("HWFNRPUG")) = False Then pSpWFSamp.HWFNRPUG = rs("HWFNRPUG") Else pSpWFSamp.HWFNRPUG = " "
    If IsNull(rs("HWFNRPUR")) = False Then pSpWFSamp.HWFNRPUR = rs("HWFNRPUR") Else pSpWFSamp.HWFNRPUR = " "
    If IsNull(rs("HWFNRSTD")) = False Then pSpWFSamp.HWFNRSTD = rs("HWFNRSTD") Else pSpWFSamp.HWFNRSTD = " "
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
    If IsNull(rs("HWFSIRDMX")) = False Then pSpWFSamp.HWFSIRDMX = rs("HWFSIRDMX") Else pSpWFSamp.HWFSIRDMX = "0"    '����]�ʏ��
    If IsNull(rs("HWFSIRDSZ")) = False Then pSpWFSamp.HWFSIRDSZ = rs("HWFSIRDSZ") Else pSpWFSamp.HWFSIRDSZ = " "    '����]�ʑ������
    If IsNull(rs("HWFSIRDHT")) = False Then pSpWFSamp.HWFSIRDHT = rs("HWFSIRDHT") Else pSpWFSamp.HWFSIRDHT = " "    '����]�ʕۏؕ��@�Q��
    If IsNull(rs("HWFSIRDHS")) = False Then pSpWFSamp.HWFSIRDHS = rs("HWFSIRDHS") Else pSpWFSamp.HWFSIRDHS = " "    '����]�ʕۏؕ��@�Q��
    If IsNull(rs("HWFSIRDKM")) = False Then pSpWFSamp.HWFSIRDKM = rs("HWFSIRDKM") Else pSpWFSamp.HWFSIRDKM = " "    '����]�ʌ����p�x�Q��
    If IsNull(rs("HWFSIRDKN")) = False Then pSpWFSamp.HWFSIRDKN = rs("HWFSIRDKN") Else pSpWFSamp.HWFSIRDKN = " "    '����]�ʌ����p�x�Q��
    If IsNull(rs("HWFSIRDKH")) = False Then pSpWFSamp.HWFSIRDKH = rs("HWFSIRDKH") Else pSpWFSamp.HWFSIRDKH = " "    '����]�ʌ����p�x�Q��
    If IsNull(rs("HWFSIRDKU")) = False Then pSpWFSamp.HWFSIRDKU = rs("HWFSIRDKU") Else pSpWFSamp.HWFSIRDKU = " "    '����]�ʌ����p�x�Q�E
    If IsNull(rs("HWFSIRDPS")) = False Then pSpWFSamp.HWFSIRDPS = Trim(rs("HWFSIRDPS")) Else pSpWFSamp.HWFSIRDPS = " "    '����]��TB�ۏ؈ʒu
    
    '�u����]��TB�ۏ؈ʒu�v�𔻒肵�A�u����]�ʌ����p�x�Q���v�ɕҏW�i���Ή��j
    Select Case Trim(pSpWFSamp.HWFSIRDPS)
    Case "T"
        pSpWFSamp.HWFSIRDKN = "3"
    Case "B"
        pSpWFSamp.HWFSIRDKN = "4"
    Case "TB"
        pSpWFSamp.HWFSIRDKN = "6"
    Case Else
        pSpWFSamp.HWFSIRDKN = " "
    End Select
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
    
    rs.Close
    '' SPV�d�l�擾�@06/06/08 ooba END =================================================>
    
    '' ���i�d�l�Ǘ��̎擾
    sql = "select HWFIGKBN from TBCME017" & _
          " where HINBAN='" & pSpWFSamp.hin.hinban & "' and MNOREVNO=" & pSpWFSamp.hin.mnorevno & _
          " and FACTORY='" & pSpWFSamp.hin.factory & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    pSpWFSamp.HWFIGKBN = rs("HWFIGKBN")
    rs.Close

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    '' �G�s�d�l�擾(BMD1E�`BMD3E,OSF1E�`OSF3E)
    sql = "select HEPOF1NS, "                   ' �i�M�����@(OSF1E)
    sql = sql & "HEPOF1SZ, "                    ' �i�������(OSF1E)
    sql = sql & "HEPOF1ET, "                    ' �i�I��ET��(OSF1E)
    sql = sql & "HEPOF1HS, "                    ' �i�ۏؕ��@_��(OSF1E)
    sql = sql & "HEPOF1SH, "                    ' �i����ʒu_��(OSF1E)
    sql = sql & "HEPOF1ST, "                    ' �i����ʒu_�_(OSF1E)
    sql = sql & "HEPOF1SR, "                    ' �i����ʒu_��(OSF1E)
    sql = sql & "HEPOF1KN, "                    ' �i�����p�x_��(OSF1E)
    sql = sql & "HEPOF2NS, "                    ' �i�M�����@(OSF2E)
    sql = sql & "HEPOF2SZ, "                    ' �i�������(OSF2E)
    sql = sql & "HEPOF2ET, "                    ' �i�I��ET��(OSF2E)
    sql = sql & "HEPOF2HS, "                    ' �i�ۏؕ��@_��(OSF2E)
    sql = sql & "HEPOF2SH, "                    ' �i����ʒu_��(OSF2E)
    sql = sql & "HEPOF2ST, "                    ' �i����ʒu_�_(OSF2E)
    sql = sql & "HEPOF2SR, "                    ' �i����ʒu_��(OSF2E)
    sql = sql & "HEPOF2KN, "                    ' �i�����p�x_��(OSF2E)
    sql = sql & "HEPOF3NS, "                    ' �i�M�����@(OSF3E)
    sql = sql & "HEPOF3SZ, "                    ' �i�������(OSF3E)
    sql = sql & "HEPOF3ET, "                    ' �i�I��ET��(OSF3E)
    sql = sql & "HEPOF3HS, "                    ' �i�ۏؕ��@_��(OSF3E)
    sql = sql & "HEPOF3SH, "                    ' �i����ʒu_��(OSF3E)
    sql = sql & "HEPOF3ST, "                    ' �i����ʒu_�_(OSF3E)
    sql = sql & "HEPOF3SR, "                    ' �i����ʒu_��(OSF3E)
    sql = sql & "HEPOF3KN, "                    ' �i�����p�x_��(OSF3E)
    sql = sql & "HEPBM1NS, "                    ' �i�M�����@(BMD1E)
    sql = sql & "HEPBM1SZ, "                    ' �i�������(BMD1E)
    sql = sql & "HEPBM1ET, "                    ' �i�I��ET��(BMD1E)
    sql = sql & "HEPBM1HS, "                    ' �i�ۏؕ��@_��(BMD1E)
    sql = sql & "HEPBM1SH, "                    ' �i����ʒu_��(BMD1E)
    sql = sql & "HEPBM1ST, "                    ' �i����ʒu_�_(BMD1E)
    sql = sql & "HEPBM1SR, "                    ' �i����ʒu_��(BMD1E)
    sql = sql & "HEPBM1KN, "                    ' �i�����p�x_��(BMD1E)
    sql = sql & "HEPBM2NS, "                    ' �i�M�����@(BMD2E)
    sql = sql & "HEPBM2SZ, "                    ' �i�������(BMD2E)
    sql = sql & "HEPBM2ET, "                    ' �i�I��ET��(BMD2E)
    sql = sql & "HEPBM2HS, "                    ' �i�ۏؕ��@_��(BMD2E)
    sql = sql & "HEPBM2SH, "                    ' �i����ʒu_��(BMD2E)
    sql = sql & "HEPBM2ST, "                    ' �i����ʒu_�_(BMD2E)
    sql = sql & "HEPBM2SR, "                    ' �i����ʒu_��(BMD2E)
    sql = sql & "HEPBM2KN, "                    ' �i�����p�x_��(BMD2E)
    sql = sql & "HEPBM3NS, "                    ' �i�M�����@(BMD3E)
    sql = sql & "HEPBM3SZ, "                    ' �i�������(BMD3E)
    sql = sql & "HEPBM3ET, "                    ' �i�I��ET��(BMD3E)
    sql = sql & "HEPBM3HS, "                    ' �i�ۏؕ��@_��(BMD3E)
    sql = sql & "HEPBM3SH, "                    ' �i����ʒu_��(BMD3E)
    sql = sql & "HEPBM3ST, "                    ' �i����ʒu_�_(BMD3E)
    sql = sql & "HEPBM3SR, "                    ' �i����ʒu_��(BMD3E)
    sql = sql & "HEPBM3KN, "                    ' �i�����p�x_��(BMD3E)
    sql = sql & "HEPACEN, "                     ' �iE1�����S
    sql = sql & "HEPANTNP, "                    ' �iEPAN���x
    sql = sql & "HEPANTIM, "                    ' �iEPAN����
    sql = sql & "HEPIGKBN, "                    ' �iEPIG�敪
    sql = sql & "HEPANGZY "                     ' �iEP����AN�K�X����
    sql = sql & "from TBCME050 "
    sql = sql & "where HINBAN = '" & pSpWFSamp.hin.hinban & "' "
    sql = sql & "and MNOREVNO = " & pSpWFSamp.hin.mnorevno & " "
    sql = sql & "and FACTORY = '" & pSpWFSamp.hin.factory & "' "
    sql = sql & "and OPECOND = '" & pSpWFSamp.hin.opecond & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    
    If IsNull(rs("HEPOF1NS")) = False Then pSpWFSamp.HEPOF1NS = rs("HEPOF1NS") Else pSpWFSamp.HEPOF1NS = " "
    If IsNull(rs("HEPOF1SZ")) = False Then pSpWFSamp.HEPOF1SZ = rs("HEPOF1SZ") Else pSpWFSamp.HEPOF1SZ = " "
    pSpWFSamp.HEPOF1ET = fncNullCheck(rs("HEPOF1ET"))
    If IsNull(rs("HEPOF1HS")) = False Then pSpWFSamp.HEPOF1HS = rs("HEPOF1HS") Else pSpWFSamp.HEPOF1HS = " "
    If IsNull(rs("HEPOF1SH")) = False Then pSpWFSamp.HEPOF1SH = rs("HEPOF1SH") Else pSpWFSamp.HEPOF1SH = " "
    If IsNull(rs("HEPOF1ST")) = False Then pSpWFSamp.HEPOF1ST = rs("HEPOF1ST") Else pSpWFSamp.HEPOF1ST = " "
    If IsNull(rs("HEPOF1SR")) = False Then pSpWFSamp.HEPOF1SR = rs("HEPOF1SR") Else pSpWFSamp.HEPOF1SR = " "
    If IsNull(rs("HEPOF1KN")) = False Then pSpWFSamp.HEPOF1KN = rs("HEPOF1KN") Else pSpWFSamp.HEPOF1KN = " "
    If IsNull(rs("HEPOF2NS")) = False Then pSpWFSamp.HEPOF2NS = rs("HEPOF2NS") Else pSpWFSamp.HEPOF2NS = " "
    If IsNull(rs("HEPOF2SZ")) = False Then pSpWFSamp.HEPOF2SZ = rs("HEPOF2SZ") Else pSpWFSamp.HEPOF2SZ = " "
    pSpWFSamp.HEPOF2ET = fncNullCheck(rs("HEPOF2ET"))
    If IsNull(rs("HEPOF2HS")) = False Then pSpWFSamp.HEPOF2HS = rs("HEPOF2HS") Else pSpWFSamp.HEPOF2HS = " "
    If IsNull(rs("HEPOF2SH")) = False Then pSpWFSamp.HEPOF2SH = rs("HEPOF2SH") Else pSpWFSamp.HEPOF2SH = " "
    If IsNull(rs("HEPOF2ST")) = False Then pSpWFSamp.HEPOF2ST = rs("HEPOF2ST") Else pSpWFSamp.HEPOF2ST = " "
    If IsNull(rs("HEPOF2SR")) = False Then pSpWFSamp.HEPOF2SR = rs("HEPOF2SR") Else pSpWFSamp.HEPOF2SR = " "
    If IsNull(rs("HEPOF2KN")) = False Then pSpWFSamp.HEPOF2KN = rs("HEPOF2KN") Else pSpWFSamp.HEPOF2KN = " "
    If IsNull(rs("HEPOF3NS")) = False Then pSpWFSamp.HEPOF3NS = rs("HEPOF3NS") Else pSpWFSamp.HEPOF3NS = " "
    If IsNull(rs("HEPOF3SZ")) = False Then pSpWFSamp.HEPOF3SZ = rs("HEPOF3SZ") Else pSpWFSamp.HEPOF3SZ = " "
    pSpWFSamp.HEPOF3ET = fncNullCheck(rs("HEPOF3ET"))
    If IsNull(rs("HEPOF3HS")) = False Then pSpWFSamp.HEPOF3HS = rs("HEPOF3HS") Else pSpWFSamp.HEPOF3HS = " "
    If IsNull(rs("HEPOF3SH")) = False Then pSpWFSamp.HEPOF3SH = rs("HEPOF3SH") Else pSpWFSamp.HEPOF3SH = " "
    If IsNull(rs("HEPOF3ST")) = False Then pSpWFSamp.HEPOF3ST = rs("HEPOF3ST") Else pSpWFSamp.HEPOF3ST = " "
    If IsNull(rs("HEPOF3SR")) = False Then pSpWFSamp.HEPOF3SR = rs("HEPOF3SR") Else pSpWFSamp.HEPOF3SR = " "
    If IsNull(rs("HEPOF3KN")) = False Then pSpWFSamp.HEPOF3KN = rs("HEPOF3KN") Else pSpWFSamp.HEPOF3KN = " "
    If IsNull(rs("HEPBM1NS")) = False Then pSpWFSamp.HEPBM1NS = rs("HEPBM1NS") Else pSpWFSamp.HEPBM1NS = " "
    If IsNull(rs("HEPBM1SZ")) = False Then pSpWFSamp.HEPBM1SZ = rs("HEPBM1SZ") Else pSpWFSamp.HEPBM1SZ = " "
    pSpWFSamp.HEPBM1ET = fncNullCheck(rs("HEPBM1ET"))
    If IsNull(rs("HEPBM1HS")) = False Then pSpWFSamp.HEPBM1HS = rs("HEPBM1HS") Else pSpWFSamp.HEPBM1HS = " "
    If IsNull(rs("HEPBM1SH")) = False Then pSpWFSamp.HEPBM1SH = rs("HEPBM1SH") Else pSpWFSamp.HEPBM1SH = " "
    If IsNull(rs("HEPBM1ST")) = False Then pSpWFSamp.HEPBM1ST = rs("HEPBM1ST") Else pSpWFSamp.HEPBM1ST = " "
    If IsNull(rs("HEPBM1SR")) = False Then pSpWFSamp.HEPBM1SR = rs("HEPBM1SR") Else pSpWFSamp.HEPBM1SR = " "
    If IsNull(rs("HEPBM1KN")) = False Then pSpWFSamp.HEPBM1KN = rs("HEPBM1KN") Else pSpWFSamp.HEPBM1KN = " "
    If IsNull(rs("HEPBM2NS")) = False Then pSpWFSamp.HEPBM2NS = rs("HEPBM2NS") Else pSpWFSamp.HEPBM2NS = " "
    If IsNull(rs("HEPBM2SZ")) = False Then pSpWFSamp.HEPBM2SZ = rs("HEPBM2SZ") Else pSpWFSamp.HEPBM2SZ = " "
    pSpWFSamp.HEPBM2ET = fncNullCheck(rs("HEPBM2ET"))
    If IsNull(rs("HEPBM2HS")) = False Then pSpWFSamp.HEPBM2HS = rs("HEPBM2HS") Else pSpWFSamp.HEPBM2HS = " "
    If IsNull(rs("HEPBM2SH")) = False Then pSpWFSamp.HEPBM2SH = rs("HEPBM2SH") Else pSpWFSamp.HEPBM2SH = " "
    If IsNull(rs("HEPBM2ST")) = False Then pSpWFSamp.HEPBM2ST = rs("HEPBM2ST") Else pSpWFSamp.HEPBM2ST = " "
    If IsNull(rs("HEPBM2SR")) = False Then pSpWFSamp.HEPBM2SR = rs("HEPBM2SR") Else pSpWFSamp.HEPBM2SR = " "
    If IsNull(rs("HEPBM2KN")) = False Then pSpWFSamp.HEPBM2KN = rs("HEPBM2KN") Else pSpWFSamp.HEPBM2KN = " "
    If IsNull(rs("HEPBM3NS")) = False Then pSpWFSamp.HEPBM3NS = rs("HEPBM3NS") Else pSpWFSamp.HEPBM3NS = " "
    If IsNull(rs("HEPBM3SZ")) = False Then pSpWFSamp.HEPBM3SZ = rs("HEPBM3SZ") Else pSpWFSamp.HEPBM3SZ = " "
    pSpWFSamp.HEPBM3ET = fncNullCheck(rs("HEPBM3ET"))
    If IsNull(rs("HEPBM3HS")) = False Then pSpWFSamp.HEPBM3HS = rs("HEPBM3HS") Else pSpWFSamp.HEPBM3HS = " "
    If IsNull(rs("HEPBM3SH")) = False Then pSpWFSamp.HEPBM3SH = rs("HEPBM3SH") Else pSpWFSamp.HEPBM3SH = " "
    If IsNull(rs("HEPBM3ST")) = False Then pSpWFSamp.HEPBM3ST = rs("HEPBM3ST") Else pSpWFSamp.HEPBM3ST = " "
    If IsNull(rs("HEPBM3SR")) = False Then pSpWFSamp.HEPBM3SR = rs("HEPBM3SR") Else pSpWFSamp.HEPBM3SR = " "
    If IsNull(rs("HEPBM3KN")) = False Then pSpWFSamp.HEPBM3KN = rs("HEPBM3KN") Else pSpWFSamp.HEPBM3KN = " "
    pSpWFSamp.HEPACEN = fncNullCheck(rs("HEPACEN"))
    pSpWFSamp.HEPANTNP = fncNullCheck(rs("HEPANTNP"))
    pSpWFSamp.HEPANTIM = fncNullCheck(rs("HEPANTIM"))
    If IsNull(rs("HEPIGKBN")) = False Then pSpWFSamp.HEPIGKBN = rs("HEPIGKBN") Else pSpWFSamp.HEPIGKBN = " "
    If IsNull(rs("HEPANGZY")) = False Then pSpWFSamp.HEPANGZY = rs("HEPANGZY") Else pSpWFSamp.HEPANGZY = " "
    rs.Close
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    scmzc_getWF = FUNCTION_RETURN_SUCCESS

PROC_EXIT:
    '' �I��
    gErr.Pop
    Exit Function

PROC_ERR:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getWF = FUNCTION_RETURN_FAILURE
    Resume PROC_EXIT

End Function
'----------------------------------------------------------------------------
'�T�v      :���i�d�lWF�f�[�^�iOT�P�AOT2)�̎擾�h���C�o
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'�@�@      :pSpWFSamp�@�@�@,IO ,typ_SpWFSamp   �@,WF�T���v���d�l
'�@�@      :�߂�l         ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :03/05/21 �㓡     2004/06/23 �ύX ���̑��T���v�������擾
'----------------------------------------------------------------------------
Public Function scmzc_getE036(pHin As tFullHinban, strOT1 As String, strOT2 As String, _
                              strMAI1 As String, strMAI2 As String) As FUNCTION_RETURN
    Dim sql     As String
    Dim rs As OraDynaset
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function scmzc_getE036"
    '--- 2004/06/23
    'sql = "select " & _
          "OTHER1, OTHER2, OTHERTIME" & _
          " from TBCME036" & _
          " where HINBAN ='" & pHin.hinban & "' and MNOREVNO=" & pHin.mnorevno & _
          " and FACTORY ='" & pHin.factory & "' and OPECOND ='" & pHin.opecond & "'" & _
          " and OTHERTIME > sysdate"
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
''    sql = "select " & _
''          "OTHER1, OTHER2, OTHERTIME, OTHER1MAI, OTHER2MAI " & _
''          " from TBCME036" & _
''          " where HINBAN ='" & pHin.hinban & "' and MNOREVNO=" & pHin.mnorevno & _
''          " and FACTORY ='" & pHin.factory & "' and OPECOND ='" & pHin.opecond & "'" & _
''          " and OTHERTIME > sysdate"
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   a.ot1 AS other1"
    sql = sql & "  ,a.ot1m AS other1mai"
    sql = sql & "  ,b.ot2 AS other2"
    sql = sql & "  ,b.ot2m AS other2mai"
    sql = sql & " FROM"
    sql = sql & "   ("
    sql = sql & "    SELECT"
    sql = sql & "      COUNT(other1)"
    sql = sql & "     ,MAX(other1) AS ot1"
    sql = sql & "     ,MAX(other1mai) AS ot1m"
    sql = sql & "    FROM"
    sql = sql & "      tbcme036"
    sql = sql & "    WHERE hinban   = '" & pHin.hinban & "'"
    sql = sql & "      AND mnorevno = " & pHin.mnorevno
    sql = sql & "      AND factory  = '" & pHin.factory & "'"
    sql = sql & "      AND opecond  = '" & pHin.opecond & "'"
    sql = sql & "      AND othertime > SYSDATE"
    sql = sql & "   ) a"
    sql = sql & "  ,("
    sql = sql & "    SELECT"
    sql = sql & "      COUNT(other2)"
    sql = sql & "     ,MAX(other2) AS ot2"
    sql = sql & "     ,MAX(other2mai) AS ot2m"
    sql = sql & "    FROM"
    sql = sql & "      tbcme036"
    sql = sql & "    WHERE hinban   = '" & pHin.hinban & "'"
    sql = sql & "      AND mnorevno = " & pHin.mnorevno
    sql = sql & "      AND factory  = '" & pHin.factory & "'"
    sql = sql & "      AND opecond  = '" & pHin.opecond & "'"
    sql = sql & "      AND othertime2 > SYSDATE"
    sql = sql & "   ) b"
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    '---------------
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        rs.Close
        strOT1 = "0"
        strOT2 = "0"
        strMAI1 = "0"    '2004/06/23
        strMAI2 = "0"    '2004/06/23
''        scmzc_getE036 = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    If IsNull(rs("OTHER1")) = True Then
        strOT1 = "0"
    Else
        strOT1 = rs("OTHER1")
    End If
    If IsNull(rs("OTHER2")) = True Then
        strOT2 = "0"
    Else
        strOT2 = rs("OTHER2")
    End If
    '----- 2004/06/23
    If IsNull(rs("OTHER1MAI")) = True Then
        strMAI1 = "0"
    Else
        strMAI1 = rs("OTHER1MAI")
    End If
    If IsNull(rs("OTHER2MAI")) = True Then
        strMAI2 = "0"
    Else
        strMAI2 = rs("OTHER2MAI")
    End If
    '-----------------
    
    scmzc_getE036 = FUNCTION_RETURN_SUCCESS
    rs.Close
    
PROC_EXIT:
    '' �I��
    gErr.Pop
    Exit Function
    
PROC_ERR:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getE036 = FUNCTION_RETURN_FAILURE
    Resume PROC_EXIT
    
End Function


'�T�v      :�_�f�͏o�Ǝc���_�f�̎d�l�`�F�b�N
'���Ұ��@�@:�ϐ���        ,IO ,�^              ,����
'      �@�@:pHin�@�@    �@,I  ,tFullHinban   �@,�i��
'      �@�@:�߂�l        ,O  ,Integer       �@,�d�l�`�F�b�N����(-1:�װ�C0:AOi�d�l���C1:AOi�d�l�L)
'����      :�_�f�͏o(��oi)�Ǝc���_�f�̗����Ɏd�l�������Ă����ꍇ�G���[��Ԃ�
'����      :03/12/05 ooba

Public Function ChkAoiSiyou(pHin As tFullHinban) As Integer

    Dim sSql As String
    Dim rs As OraDynaset
    Dim sDoiSiyou(2) As String  '�����L��(DOi1�`3)
    Dim sAoiSiyou As String     '�����L��(AOi)
    Dim iCnt As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function ChkAoiSiyou"

    sSql = "select HWFOS1HS, HWFOS2HS, HWFOS3HS, HWFZOHWS from TBCME025 "
    sSql = sSql & "where HINBAN = '" & pHin.hinban & "' "
    sSql = sSql & "and MNOREVNO = " & pHin.mnorevno & " "
    sSql = sSql & "and FACTORY = '" & pHin.factory & "' "
    sSql = sSql & "and OPECOND = '" & pHin.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        rs.Close
        ChkAoiSiyou = -1
        GoTo PROC_EXIT
    End If
    
    If IsNull(rs("HWFOS1HS")) = False Then sDoiSiyou(0) = rs("HWFOS1HS") '�iWF�_�f�͏o1�ۏؕ��@_��
    If IsNull(rs("HWFOS2HS")) = False Then sDoiSiyou(1) = rs("HWFOS2HS") '�iWF�_�f�͏o2�ۏؕ��@_��
    If IsNull(rs("HWFOS3HS")) = False Then sDoiSiyou(2) = rs("HWFOS3HS") '�iWF�_�f�͏o3�ۏؕ��@_��
    If IsNull(rs("HWFZOHWS")) = False Then sAoiSiyou = rs("HWFZOHWS")    '�iWF�c���_�f�ۏؕ��@_��
    
'--------------- 2008/07/25 INSERT START  By Systech ---------------
    rs.Close
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
    
    '�_�f�͏o�Ǝc���_�f�̎d�l�`�F�b�N
    ChkAoiSiyou = 0
    For iCnt = 0 To 2
        If sDoiSiyou(iCnt) = "H" Or sDoiSiyou(iCnt) = "S" Then
            '�_�f�͏o(��oi)�Ǝc���_�f�̗����Ɏd�l�������Ă����ꍇ�̓G���[
            If sAoiSiyou = "H" Or sAoiSiyou = "S" Then
                ChkAoiSiyou = -1
                Exit For
            End If
        Else
            If sAoiSiyou = "H" Or sAoiSiyou = "S" Then
                ChkAoiSiyou = 1
            End If
        End If
    Next
    
PROC_EXIT:
    '' �I��
    gErr.Pop
    Exit Function
    
PROC_ERR:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    ChkAoiSiyou = -1
    Resume PROC_EXIT
    
End Function

