Attribute VB_Name = "SB_CryJudg"
Option Explicit

' ������t���[��
' �d�l�ۏؕ��@�Q�� --+--�Ȃ� --���сi�Y���ʒu�j--�����Ă��Ȃ��Ă�����OK
'�@�@�@�@�@�@�@�@�@�@|
'                   +--���� --���сi�Y���ʒu) --+--���� -- ����`�F�b�N --+-- OK
'                                              |                        |
'                                              |                        +-- MG
'                                              |
'                                              +--�Ȃ� --+-- �����w���T�E�U�ȊO�̏ꍇ --+--EPD�ACs�ALT�̏ꍇ����T�� --+-- �Ȃ� -- NG
'                                                        |                            |                          �@ |
'                                                        |                            +--EPD,Cs�ALT�ȊO -- NG       +-- ���� -- ����`�F�b�N --+-- OK
'                                                        |                                                                                    |
'                                                        |                                                                                    +-- NG
'                                                        |
'                                                        +-- �����w���T�̏ꍇ (Rs, Cs) �Ȃ琄��Ȃ̂őS�̂�����т�T�� --+
'                                                        |  �@�@�@�@�@�@�@�@�@                                          |
'                                                        |                                                             |
'                                                        +-- �����w���U�̏ꍇ TOP�Ȃ��ցATAIL�Ȃ牺�֎��т�T��       --+-- ����`�F�b�N --+-- OK
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@                 �@|
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�i�����w���́A�w���𗧂Ă鑤������ɗ��ĂĂ���ƍl���Ă���j�@�@�@�@�@                 +-- NG

'' �f�o�b�N��`
''
'' �萔��`
''
Private Const MAXCNT    As Integer = 16         ' �ő匏��
Public Const BlkTop     As Integer = 1          ' TOP��
Public Const BlkTail    As Integer = 2          ' TAIL��
Public Const MSYSCLASS  As String = "NM"        ' �V�X�e���敪
Public Const KCLASS     As String = "01"        ' �N���X
Public Const KCODE      As String = "1"         ' �R�[�h

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : ���уp�^�[���敪�Ɛ��i�d�l�p�^�[���敪�̒�`�ǉ�
Private Const CNST_JSK_PTN_None As String = "0"
Private Const CNST_JSK_PTN_Ring As String = "1"
Private Const CNST_JSK_PTN_Disk As String = "2"
Private Const CNST_JSK_PTN_DiskRing As String = "3"
Private Const CNST_JSK_PTN_PBband As String = "5"
Private Const CNST_JSK_PTN_Pband As String = "6"
Private Const CNST_JSK_PTN_Bband As String = "7"

Private Const CNST_SIYO_NO_Ring As String = "1"
Private Const CNST_SIYO_NO_Disk As String = "2"
Private Const CNST_SIYO_NO_Pattern As String = "3"
Private Const CNST_SIYO_Fumon As String = "4"
Private Const CNST_SIYO_NO_PBband As String = "5"
Private Const CNST_SIYO_NO_Pband As String = "6"
Private Const CNST_SIYO_NO_Bband As String = "7"
'Add End   2011/01/17 SMPK A.Nagamine

'�e���茋�ʏ��
Public Type typ_ALLRSLT
    pos     As Integer                  ' �������J�n�ʒu
    NAIYO   As String                   ' ���e
    INFO1   As String                   ' ���P
    INFO2   As String                   ' ���Q
    INFO3   As String                   ' ���R
    INFO4   As String                   ' ���S
    OKNG    As String                   ' ���茋��
    SMPLNO  As Long                     ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLID  As String                   ' �T���v��ID�iWF_Judg�Ŏg�p�j
    BLOCKNG As Boolean                  ' GD�G���[�ƂȂ�i�Ԃ��܂ނ�����
    hinban  As String                   ' �i��(12��)
End Type
    
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
'�����̍\���̂ɍ��ڒǉ������VB�̐����Ɉ���������̂ŁA�ʂŊǗ�����B
'�e���茋�ʏ��
Public Type typ_ALLRSLT_EX
    pos     As Integer                  ' �������J�n�ʒu
    NAIYO   As String                   ' ���e
    INFO1   As String                   ' ���P
    INFO2   As String                   ' ���Q
    INFO3   As String                   ' ���R
    INFO4   As String                   ' ���S
    INFO5   As String                   ' ���T�iAN���x�j
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
'����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
    INFO6   As String                   ' ���U�iPUA�l�j
    INFO7   As String                   ' ���V�iPUA%�l�j
    INFO8   As String                   ' ���W�iSTD�l�j
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
    OKNG    As String                   ' ���茋��
    SMPLNO  As Long                     ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLID  As String                   ' �T���v��ID�iWF_Judg�Ŏg�p�j
    BLOCKNG As Boolean                  ' GD�G���[�ƂȂ�i�Ԃ��܂ނ�����
    hinban  As String                   ' �i��(12��)
End Type
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------

'�S���\����
Type typ_AllTypesB
    intPFlg             As Integer                              ' �\���t���O
    StrStaffId          As String                               ' �X�^�b�tID
    strStaffName        As String                               ' �X�^�b�t��
    BLOCKID             As String * 12                          ' �u���b�NID
    Cut(2)              As Double                               ' �ăJ�b�g�ʒu
    COEF(2)             As Double                               ' �ΐ͌W��
    CRCOEF              As Double                               ' �����ΐ͌W��
    OKNG(2)             As Boolean                              ' ���R����
    Henseki             As Boolean                              ' ���R���їL��(�����S��TOP/TAIL)
    JudgRes(2)          As Boolean                              ' ���R����    2001/10/02 S.Sano
    JudgRrg(2)          As Boolean                              ' RRG����       2001/10/02 S.Sano
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    JudgDkTmp(2)        As Boolean                              ' DK���x����
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    typ_rsz()           As typ_TBCMJ002                         ' ������R����(�����S��TOP/TAIL)
    typ_hage(2)         As typ_TBCMH004                         ' ���グ�I������
    typ_rslt(2, MAXCNT) As typ_ALLRSLT                          ' �e���я��
    typ_zi              As type_DBDRV_scmzc_fcmkc001c_Zisseki   ' ���т��܂Ƃ߂��\����
    typ_si()            As type_DBDRV_scmzc_fcmkc001c_Siyou     ' �d�l
    typ_cr()            As type_DBDRV_scmzc_fcmkc001c_CrySmp    ' �����T���v���Ǘ��擾�p (TOP,TAIL���łQ���R�[�h�擾)
    blYONE              As Boolean                              ' �đ�t���O�i��w���P�����@yaz�j
    COEFflg             As Boolean                              ' �u���b�N�ΐ͔���t���O   2005/1/11�ǉ�
    Hinsyu              As String                               ' �u���b�N�ΐ͔���(�i��j  2005/1/11�ǉ�
    DOPEflg             As Boolean                              ' ���ް�߈ʒu����         2005/1/11�ǉ�
End Type

Public typ_b        As typ_AllTypesB        '�S���\����
Public JudgSC_B(2)  As Judg_Spec_Cry        '�d�l�����x���\����
Public ciSmpGetFlg  As Integer              '����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
Public ciKcnt       As Integer              '�H���A��
'----2005/1/11
Type typ_Suitei
    COEF                As Double                               ' �ΐ͌W��
    Henseki             As Boolean                              ' ���R���їL��(�����S��TOP/TAIL)
    SuiSpec             As type_DBDRV_scmzc_fcmkc001c_Siyou     ' �d�l
    SuiData(2)          As type_DBDRV_scmzc_fcmkc001c_CryR
    COEFflg             As Boolean                              ' �u���b�N�ΐ͔���t���O
    Hinsyu              As String                               ' �u���b�N�ΐ͔���(�i��j
    DOPEflg             As Boolean                              ' ���ް�߈ʒu����
    RsJudg(2)           As Boolean
End Type
'---TEST2004/10
Public SuiteiData() As typ_Suitei
''==�����i�Ԕ���Ή��@20060501 SMP����
'' 0:�S�������ڂō��۔���,1:Cs,LT,EPD�ō��۔���,2:Skip
Public giTpMultiFlg As Integer ''Top�ł̍��۔���U�蕪��
Public giBtMultiFlg As Integer ''bottom�ł̍��۔���U�蕪��
Private pJMEAS_Top() As Double
Private psKSTAFFID  As String
Private psHSXRSPOT  As String
Private psHSXRSPOI  As String
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
Public gsCOSF3Flg As String
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---

'--------------- 2008/07/25 INSERT START  By Systech ---------------
Private pbGDJudgeTbl(3) As Boolean          ' GD���茋�ʑޔ�
'--------------- 2008/07/25 INSERT  END   By Systech ---------------

'------------------------------------------------
' ��������
'------------------------------------------------

'�T�v      :���ђl�̑���������s���B
'���Ұ�    :�ϐ���          ,IO ,�^             :����
'          :sKeyID          ,I  ,String         :��ۯ�ID�A���́A�����ԍ�
'          :tNew_Hinban     ,I  ,tFullHinban    :�U�֌��i��
'          :bTotalJudg      ,O  ,Boolean        :�g�[�^������
'          :iErr_Code       ,O  ,Integer        :�װ����(�߂�l�Ɠ���)
'          :sErr_Msg        ,O  ,String         :�װү���޺���
'          :typ_B           ,O  ,typ_AllTypesB  :�S���\����(�\����)
'          :iSmpGetFlg      ,I  ,Integer        :����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
'          :iSamplID1       ,I  ,Long           :TOP�����ID(�ȗ���)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iSamplID2       ,I  ,Long           :BOT�����ID(�ȗ���)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iKcnt           ,I  ,Integer        :�H���A��(�ȗ���)
'          :�߂�l          ,O  ,Integer        :�擾�̐���(0:����I��, -1:�ُ�I��)
'����      :
'����      :2003/09/19 �V�K�쐬�@SB

Public Function funCrySogoHantei(sKeyID As String, tNew_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_b As typ_AllTypesB, _
                iSmpGetFlg As Integer, Optional iSamplID1 As Long = 0, Optional iSamplID2 As Long = 0, _
                Optional iKcnt As Integer = 0) As Integer
    
    On Error GoTo Apl_down
    
    '�߂�l������
    funCrySogoHantei = FUNCTION_RETURN_FAILURE
    TotalJudg = True
    
    '�O���[�o���ϐ��ɐݒ�
    ciSmpGetFlg = iSmpGetFlg
    ciKcnt = iKcnt
    
    '�u���b�NID��ݒ�
    sErr_Msg = "������������(��ۯ�ID�ݒ�)"
    typ_b.BLOCKID = sKeyID
    
    '��ʏ��ݒ�
    sErr_Msg = "������������(SetAllData)"
    If SetAllData(typ_b, tNew_Hinban, iErr_Code, sErr_Msg, iSmpGetFlg, iSamplID1, iSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '�d�l�����w���擾
    sErr_Msg = "������������(SpecJudgCheck)"
    Call SpecJudgCheck
    
    '2003/12/13 SystemBrain Null�Ή��ǉ���
    '�d�lNull�`�F�b�N
    sErr_Msg = "�d�lNull����"
    If funCryChkNull(typ_b.typ_si(BlkTop), sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    '2003/12/13 SystemBrain Null�Ή��ǉ���
    
    '���уf�[�^����(TOP)
    sErr_Msg = "������������(����(TOP))"
    
    '----TEST2004/10
    '��ʏo�͗p�Ɏ�����R�l��ޔ����Ă���
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS5
        
    If Trim(typ_b.typ_zi.CRYRZ(BlkTop).KSTAFFID) <> KSTAFF_J002 Then
        '��R�l�𑪒�ʒu�R�[�h�ɂ����בւ���
        
        If Set_Rs_Ichi(typ_b.typ_si(BlkTop).HSXRSPOT, typ_b.typ_si(BlkTop).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTop).MEAS1, _
                        typ_b.typ_zi.CRYRZ(BlkTop).MEAS2, typ_b.typ_zi.CRYRZ(BlkTop).MEAS3, typ_b.typ_zi.CRYRZ(BlkTop).MEAS4, typ_b.typ_zi.CRYRZ(BlkTop).MEAS5) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If
    If CrAllJudg(typ_b, tNew_Hinban, BlkTop) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    '���уf�[�^����(TAIL)
    sErr_Msg = "������������(����(TAIL))"
    '��ʏo�͗p�Ɏ�����R�l��ޔ����Ă���
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS5
        
    If Trim(typ_b.typ_zi.CRYRZ(BlkTail).KSTAFFID) <> KSTAFF_J002 Then
        '��R�l�𑪒�ʒu�R�[�h�ɂ����בւ���
        If Set_Rs_Ichi(typ_b.typ_si(BlkTail).HSXRSPOT, typ_b.typ_si(BlkTail).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTail).MEAS1, _
                        typ_b.typ_zi.CRYRZ(BlkTail).MEAS2, typ_b.typ_zi.CRYRZ(BlkTail).MEAS3, typ_b.typ_zi.CRYRZ(BlkTail).MEAS4, typ_b.typ_zi.CRYRZ(BlkTail).MEAS5) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If
    If CrAllJudg(typ_b, tNew_Hinban, BlkTail) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    bTotalJudg = TotalJudg
    
    funCrySogoHantei = FUNCTION_RETURN_SUCCESS
'------------------------------------------ �I������  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funCrySogoHantei = -4
    iErr_Code = funCrySogoHantei
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' ��������(���f�f�[�^�̍��۔�����s��Ȃ�)
'------------------------------------------------

'�T�v      :���ђl�̑���������s���B
'���Ұ�    :�ϐ���          ,IO ,�^             :����
'          :sKeyID          ,I  ,String         :��ۯ�ID�A���́A�����ԍ�
'          :Top_Hinban      ,I  ,tFullHinban    :TOP�i��
'          :Tail_Hinban     ,I  ,tFullHinban    :TAIL�i��
'          :bTotalJudg      ,O  ,Boolean        :�g�[�^������
'          :iErr_Code       ,O  ,Integer        :�װ����(�߂�l�Ɠ���)
'          :sErr_Msg        ,O  ,String         :�װү���޺���
'          :typ_B           ,O  ,typ_AllTypesB  :�S���\����(�\����)
'          :iSmpGetFlg      ,I  ,Integer        :����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
'          :iSamplID1       ,I  ,Long           :TOP�����ID(�ȗ���)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iSamplID2       ,I  ,Long           :BOT�����ID(�ȗ���)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iKcnt           ,I  ,Integer        :�H���A��(�ȗ���)
'          :�߂�l          ,O  ,Integer        :�擾�̐���(0:����I��, -1:�ُ�I��)
'����      :
'����      :2005/02/07 �V�K�쐬�@�ǉ� ffc)tanabe

Public Function funCrySogoHantei2(sKeyID As String, Top_Hinban As tFullHinban, Tail_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_b As typ_AllTypesB, _
                iSmpGetFlg As Integer, Optional iSamplID1 As Long = 0, Optional iSamplID2 As Long = 0, _
                Optional iKcnt As Integer = 0) As Integer
    
    On Error GoTo Apl_down
    
    '�߂�l������
    funCrySogoHantei2 = FUNCTION_RETURN_FAILURE
    TotalJudg = True
    
    '�O���[�o���ϐ��ɐݒ�
    ciSmpGetFlg = iSmpGetFlg
    ciKcnt = iKcnt
    
    '�u���b�NID��ݒ�
    sErr_Msg = "������������(��ۯ�ID�ݒ�)"
    typ_b.BLOCKID = sKeyID
    
    '��ʏ��ݒ�(TOP��)
    sErr_Msg = "������������(SetAllData2)"
    If SetAllData2(typ_b, Top_Hinban, Tail_Hinban, iErr_Code, sErr_Msg, iSmpGetFlg, iSamplID1, iSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '�d�l�����w���擾
    sErr_Msg = "������������(SpecJudgCheck)"
    Call SpecJudgCheck
    
    '�d�lNull�`�F�b�N
    sErr_Msg = "�d�lNull����"
    If funCryChkNull(typ_b.typ_si(BlkTop), sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '���уf�[�^����(TOP)
    sErr_Msg = "������������(����(TOP))"
    
    '��ʏo�͗p�Ɏ�����R�l��ޔ����Ă���
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS5
        
    If Trim(typ_b.typ_zi.CRYRZ(BlkTop).KSTAFFID) <> KSTAFF_J002 Then
        '��R�l�𑪒�ʒu�R�[�h�ɂ����בւ���
        
        If Set_Rs_Ichi(typ_b.typ_si(BlkTop).HSXRSPOT, typ_b.typ_si(BlkTop).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTop).MEAS1, _
                        typ_b.typ_zi.CRYRZ(BlkTop).MEAS2, typ_b.typ_zi.CRYRZ(BlkTop).MEAS3, typ_b.typ_zi.CRYRZ(BlkTop).MEAS4, typ_b.typ_zi.CRYRZ(BlkTop).MEAS5) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If
    If CrAllJudg(typ_b, Top_Hinban, BlkTop) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    '���уf�[�^����(TAIL)
    sErr_Msg = "������������(����(TAIL))"
    '��ʏo�͗p�Ɏ�����R�l��ޔ����Ă���
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS5
        
    If Trim(typ_b.typ_zi.CRYRZ(BlkTail).KSTAFFID) <> KSTAFF_J002 Then
        '��R�l�𑪒�ʒu�R�[�h�ɂ����בւ���
        If Set_Rs_Ichi(typ_b.typ_si(BlkTail).HSXRSPOT, typ_b.typ_si(BlkTail).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTail).MEAS1, _
                        typ_b.typ_zi.CRYRZ(BlkTail).MEAS2, typ_b.typ_zi.CRYRZ(BlkTail).MEAS3, typ_b.typ_zi.CRYRZ(BlkTail).MEAS4, typ_b.typ_zi.CRYRZ(BlkTail).MEAS5) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If

    If CrAllJudg(typ_b, Tail_Hinban, BlkTail) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    bTotalJudg = TotalJudg
    
    funCrySogoHantei2 = FUNCTION_RETURN_SUCCESS

Apl_Exit:
    
    Exit Function
    
Apl_down:
    funCrySogoHantei2 = -4
    iErr_Code = funCrySogoHantei2
    GoTo Apl_Exit
    
End Function

'�T�v      :��ʏ��f�[�^�ݒ�
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_A         ,IO ,typ_AllTypes ,�e���\����
'����      :��ʏ������\���̂ɐݒ肷��
'����      :

Public Function SetAllData(typ_b As typ_AllTypesB, tNew_Hinban As tFullHinban, iErr_Code As Integer, _
                           sErr_Msg As String, iSmpGetFlg As Integer, iSamplID1 As Long, iSamplID2 As Long) As FUNCTION_RETURN
    
    Dim typ_hi()    As typ_TBCMH004
    Dim typ_tan     As typ_TBCMG002
    Dim sErrMsg     As String
    Dim i           As Integer

    '�u���b�NID��3���Ŕ��f����
    '�܂��A�ۗ�
    
    SetAllData = FUNCTION_RETURN_FAILURE ''2001/07/25 Sano�C��

    '�������� �e��f�[�^�擾
    sErr_Msg = "������������(funCryGetDataEtc)"
    If funCryGetDataEtc(typ_b.BLOCKID, tNew_Hinban, _
                        typ_b.typ_si, _
                        typ_b.typ_cr, _
                        typ_b.typ_zi, _
                        sErrMsg, _
                        iSmpGetFlg, iSamplID1, iSamplID2) <> FUNCTION_RETURN_SUCCESS Then
        If sErrMsg = "0" Then sErr_Msg = "�����敪�G���["
        Exit Function
    End If
    
    typ_b.blYONE = True
    With typ_b
        ' ���������w���iRs)
        sErr_Msg = "������������(RS-Top)"
        If InStr("123", .typ_cr(BlkTop).CRYINDRSCS) <> 0 And _
            .typ_zi.CRYRZ(BlkTop).SMPLUMU = "0" Then
            
            '���グ�I�����ю擾
            ReDim typ_hi(0)
            sErr_Msg = "������������(RS-Top���グ�I�����ю擾)"
            If s_cmmc001db_Sql(.typ_cr(BlkTop).XTALCS, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                If UBound(typ_hi) = 0 Then
                   '���グ�I�����ю擾���s
                    Exit Function
                Else
                    .typ_hage(BlkTop) = typ_hi(1)
                End If
            End If
        End If
        
        ' ���������w���iRs)
        sErr_Msg = "������������(RS-Bot)"
        If InStr("123", .typ_cr(BlkTail).CRYINDRSCS) <> 0 And _
            .typ_zi.CRYRZ(BlkTail).SMPLUMU = "0" Then
            
            '���グ�I�����ю擾
            ReDim typ_hi(0)
            sErr_Msg = "������������(RS-Bot���グ�I�����ю擾)"
            If s_cmmc001db_Sql(.typ_cr(BlkTail).XTALCS, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                If UBound(typ_hi) = 0 Then
                   '���グ�I�����ю擾���s
                    Exit Function
                Else
                    .typ_hage(BlkTail) = typ_hi(1)
                End If
            End If
        End If
    End With
    
    '�����S��TOP/TAIL��R���ђl�擾
    sErr_Msg = "������������(�����S��TOP/TAIL��R���ђl�擾)"
    If s_cmmc001db2_sql(typ_b.typ_si(1).CRYNUM, _
                        typ_b.typ_si(1).ADDDPPOS, _
                        typ_b.typ_si(1).FREELENG, _
                        typ_b.typ_cr(2).INPOSCS, _
                        typ_b.typ_rsz()) <> FUNCTION_RETURN_SUCCESS Then
       '��R���ђl���s
        typ_b.Henseki = False
    Else
        typ_b.Henseki = True
    End If
        
    SetAllData = FUNCTION_RETURN_SUCCESS
End Function

'�T�v      :��ʏ��f�[�^�ݒ�(������������F���f�f�[�^�̍��۔�����s��Ȃ��p�֐�)
'���Ұ�    :�ϐ���          ,IO ,�^             ,����
'          :typ_B           ,IO ,typ_AllTypes   ,�e���\����
'          :Top_Hinban      ,I  ,tFullHinban    ,TOP�i��
'          :Tail_Hinban     ,I  ,tFullHinban    ,TAIL�i��
'          :iErr_Code       ,O  ,Integer        ,�װ����(�߂�l�Ɠ���)
'          :sErr_Msg        ,O  ,String         ,�װү���޺���
'          :iSmpGetFlg      ,I  ,Integer        ,����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
'          :iSamplID1       ,I  ,Long           ,TOP�����ID(�ȗ���)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iSamplID2       ,I  ,Long           ,BOT�����ID(�ȗ���)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'����      :��ʏ������\���̂ɐݒ肷��
'����      :2005/02/08 ffc)tanabe

Public Function SetAllData2(typ_b As typ_AllTypesB, Top_Hinban As tFullHinban, Tail_Hinban As tFullHinban, iErr_Code As Integer, _
                           sErr_Msg As String, iSmpGetFlg As Integer, iSamplID1 As Long, iSamplID2 As Long) As FUNCTION_RETURN
    
    Dim typ_hi()    As typ_TBCMH004
    Dim typ_tan     As typ_TBCMG002
    Dim sErrMsg     As String
    Dim i           As Integer
    
    SetAllData2 = FUNCTION_RETURN_FAILURE

    '�������� �e��f�[�^�擾
    sErr_Msg = "������������(funCryGetDataEtc2)"
    If funCryGetDataEtc2(typ_b.BLOCKID, Top_Hinban, Tail_Hinban, _
                        typ_b.typ_si, _
                        typ_b.typ_cr, _
                        typ_b.typ_zi, _
                        sErrMsg, _
                        iSmpGetFlg, iSamplID1, iSamplID2) <> FUNCTION_RETURN_SUCCESS Then
    If sErrMsg = "0" Then sErr_Msg = "�����敪�G���["
        Exit Function
    End If
    
    typ_b.blYONE = True
    With typ_b
        ' ���������w���iRs)
        sErr_Msg = "������������(RS-Top)"
        If InStr("123", .typ_cr(BlkTop).CRYINDRSCS) <> 0 And _
            .typ_zi.CRYRZ(BlkTop).SMPLUMU = "0" Then
            
            '���グ�I�����ю擾
            ReDim typ_hi(0)
            sErr_Msg = "������������(RS-Top���グ�I�����ю擾)"
            If s_cmmc001db_Sql(.typ_cr(BlkTop).XTALCS, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                If UBound(typ_hi) = 0 Then
                   '���グ�I�����ю擾���s
                    Exit Function
                Else
                    .typ_hage(BlkTop) = typ_hi(1)
                End If
            End If
        End If
        
        ' ���������w���iRs)
        sErr_Msg = "������������(RS-Bot)"
        If InStr("123", .typ_cr(BlkTail).CRYINDRSCS) <> 0 And _
            .typ_zi.CRYRZ(BlkTail).SMPLUMU = "0" Then
            
            '���グ�I�����ю擾
            ReDim typ_hi(0)
            sErr_Msg = "������������(RS-Bot���グ�I�����ю擾)"
            If s_cmmc001db_Sql(.typ_cr(BlkTail).XTALCS, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                If UBound(typ_hi) = 0 Then
                   '���グ�I�����ю擾���s
                    Exit Function
                Else
                    .typ_hage(BlkTail) = typ_hi(1)
                End If
            End If
        End If
    End With
    
    '�����S��TOP/TAIL��R���ђl�擾
    sErr_Msg = "������������(�����S��TOP/TAIL��R���ђl�擾)"
    If s_cmmc001db2_sql(typ_b.typ_si(1).CRYNUM, _
                        typ_b.typ_si(1).ADDDPPOS, _
                        typ_b.typ_si(1).FREELENG, _
                        typ_b.typ_cr(2).INPOSCS, _
                        typ_b.typ_rsz()) <> FUNCTION_RETURN_SUCCESS Then
       '��R���ђl���s
        typ_b.Henseki = False
    Else
        typ_b.Henseki = True
    End If
    
    SetAllData2 = FUNCTION_RETURN_SUCCESS

End Function

Public Sub SpecJudgCheck()
    Dim c0              As Integer
    Dim UDHinSpec(2)    As Judg_Spec_Cry
    Dim smpShared       As Boolean
    Dim KouteiKbn       As Integer              '�H���敪�@08/04/15 ooba
    Dim sSxlPos         As String               'SXL�ʒu(TOP/BOT)�@08/04/15 ooba
    
    '08/04/15 ooba START ======================================================>
    '�H���ɂ�茋������L���𔻒f����B
    '�@�Ĕ����w��(CW760)�ȊO�̏ꍇ�͌����ۏ؂ɂ�蔻�f�B
    '              (�����ۏ�)=(X) ���Ȃ�
    '                        =(H) ������
    '�A�Ĕ����w��(CW760)�̏ꍇ�͌����ۏ؂�WF�ۏ؂̑g�����ɂ�蔻�f�B
    '       (�����ۏ�,WF�ۏ�)=(X,X) ���Ȃ�
    '                        =(X,H) ���Ȃ�
    '                        =(H,X) ������
    '                        =(H,H) ���Ȃ�
    '�BWF�H��(CC720)�ȍ~��COSF3�̔�����s�Ȃ�Ȃ��B
    
    '�H�����f
    Select Case left(JudgKoutei, 4)
    '--�����H��
    Case "CC10", "CC20", "CC30", "CC31", "CC40", "CC45", "CC46", "CC60", "CC61", "CC70", "CC72"
        KouteiKbn = 0
    '--WF�H��(WF����O)
    Case "CC73", "CW74", "CW75"
        KouteiKbn = 1
    '--WF�H��(WF�����)
    Case "CW76", "CW80"
        KouteiKbn = 2
    '--���̑�
    Case Else
        KouteiKbn = 0
    End Select
    
    For c0 = 1 To 2
        With typ_b.typ_si(c0)
            sSxlPos = IIf(c0 = SxlTop, "TOP", "BOT")
            '�����H��
            If KouteiKbn = 0 Then
                JudgSC_B(c0).rs = (.HSXRHWYS = "H")
                JudgSC_B(c0).Oi = (.HSXONHWS = "H")
                JudgSC_B(c0).B1 = (.HSXBM1HS = "H")
                JudgSC_B(c0).B2 = (.HSXBM2HS = "H")
                JudgSC_B(c0).B3 = (.HSXBM3HS = "H")
                JudgSC_B(c0).L1 = (.HSXOF1HS = "H")
                JudgSC_B(c0).L2 = (.HSXOF2HS = "H")
                JudgSC_B(c0).L3 = (.HSXOF3HS = "H")
                JudgSC_B(c0).L4 = (.HSXOF4HS = "H")
                JudgSC_B(c0).COSF3 = (.COSF3FLAG = "H")
                JudgSC_B(c0).GD = (.HSXDENHS = "H") Or (.HSXLDLHS = "H") Or (.HSXDVDHS = "H")
                JudgSC_B(c0).Cs = (.HSXCNHWS = "H")
                JudgSC_B(c0).Lt = (.HSXLTHWS = "H")
                JudgSC_B(c0).EPD = True
                
              'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ�
                JudgSC_B(c0).CuC = (.HSXCHS = "H")
                JudgSC_B(c0).CuCJ = (.HSXCJHS = "H")
                JudgSC_B(c0).CuCJLT = (.HSXCJLTHS = "H")
                JudgSC_B(c0).CuCJ2 = (.HSXCJ2HS = "H")
              'Add End   2011/01/17 SMPK A.Nagamine
            'WF�H��(WF����O)
            ElseIf KouteiKbn = 1 Then
                JudgSC_B(c0).rs = (.HSXRHWYS = "H")
                JudgSC_B(c0).Oi = (.HSXONHWS = "H")
                JudgSC_B(c0).B1 = (.HSXBM1HS = "H")
                JudgSC_B(c0).B2 = (.HSXBM2HS = "H")
                JudgSC_B(c0).B3 = (.HSXBM3HS = "H")
                JudgSC_B(c0).L1 = (.HSXOF1HS = "H")
                JudgSC_B(c0).L2 = (.HSXOF2HS = "H")
                JudgSC_B(c0).L3 = (.HSXOF3HS = "H")
                JudgSC_B(c0).L4 = (.HSXOF4HS = "H")
              'Add Start 2011/02/01 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ�
                'JudgSC_B(c0).COSF3 = False
                JudgSC_B(c0).COSF3 = (.COSF3FLAG = "H")
              'Add End   2011/02/01 SMPK A.Nagamine
                JudgSC_B(c0).GD = (.HSXDENHS = "H") Or (.HSXLDLHS = "H") Or (.HSXDVDHS = "H")
                JudgSC_B(c0).Cs = (.HSXCNHWS = "H")
                JudgSC_B(c0).Lt = (.HSXLTHWS = "H")
                JudgSC_B(c0).EPD = True
                
              'Add Start 2011/02/01 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ�
                JudgSC_B(c0).CuC = (.HSXCHS = "H")
                JudgSC_B(c0).CuCJ = (.HSXCJHS = "H")
                JudgSC_B(c0).CuCJLT = (.HSXCJLTHS = "H")
                JudgSC_B(c0).CuCJ2 = (.HSXCJ2HS = "H")
              'Add End   2011/02/01 SMPK A.Nagamine
            'WF�H��(WF�����)
            ElseIf KouteiKbn = 2 Then
                JudgSC_B(c0).rs = (.HSXRHWYS = "H") And _
                                    ((.HWFRHWYS <> "H") Or Not CheckKHN(.HWFRKHNN, 1, sSxlPos))
                JudgSC_B(c0).Oi = (.HSXONHWS = "H") And _
                                    ((.HWFONHWS <> "H") Or Not CheckKHN(.HWFONKHN, 2, sSxlPos))
                JudgSC_B(c0).B1 = (.HSXBM1HS = "H") And _
                                    ((.HWFBM1HS <> "H") Or Not CheckKHN(.HWFBM1KN, 7, sSxlPos))
                JudgSC_B(c0).B2 = (.HSXBM2HS = "H") And _
                                    ((.HWFBM2HS <> "H") Or Not CheckKHN(.HWFBM2KN, 8, sSxlPos))
                JudgSC_B(c0).B3 = (.HSXBM3HS = "H") And _
                                    ((.HWFBM3HS <> "H") Or Not CheckKHN(.HWFBM3KN, 9, sSxlPos))
                JudgSC_B(c0).L1 = (.HSXOF1HS = "H") And _
                                    ((.HWFOF1HS <> "H") Or Not CheckKHN(.HWFOF1KN, 3, sSxlPos))
                JudgSC_B(c0).L2 = (.HSXOF2HS = "H") And _
                                    ((.HWFOF2HS <> "H") Or Not CheckKHN(.HWFOF2KN, 4, sSxlPos))
                JudgSC_B(c0).L3 = (.HSXOF3HS = "H") And _
                                    ((.HWFOF3HS <> "H") Or Not CheckKHN(.HWFOF3KN, 5, sSxlPos))
                JudgSC_B(c0).L4 = (.HSXOF4HS = "H") And _
                                    ((.HWFOF4HS <> "H") Or Not CheckKHN(.HWFOF4KN, 6, sSxlPos))
              'Add Start 2011/02/01 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ�
                'JudgSC_B(c0).COSF3 = False
                JudgSC_B(c0).COSF3 = (.COSF3FLAG = "H")
              'Add End   2011/02/01 SMPK A.Nagamine
                JudgSC_B(c0).GD = ((.HSXDENHS = "H") Or (.HSXLDLHS = "H") Or (.HSXDVDHS = "H")) And _
                                    (((.HWFDENHS <> "H") And (.HWFLDLHS <> "H") And (.HWFDVDHS <> "H")) Or Not CheckKHN(.HWFGDKHN, 18, sSxlPos))
                JudgSC_B(c0).Cs = (.HSXCNHWS = "H")
                JudgSC_B(c0).Lt = (.HSXLTHWS = "H")
                JudgSC_B(c0).EPD = True
                
              'Add Start 2011/02/01 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ�
                JudgSC_B(c0).CuC = (.HSXCHS = "H")
                JudgSC_B(c0).CuCJ = (.HSXCJHS = "H")
                JudgSC_B(c0).CuCJLT = (.HSXCJLTHS = "H")
                JudgSC_B(c0).CuCJ2 = (.HSXCJ2HS = "H")
              'Add End   2011/02/01 SMPK A.Nagamine
            End If
        End With
    Next
    '08/04/15 ooba END ========================================================>
    
End Sub

'�T�v      :���㌋������(�S)
'���Ұ�    :�ϐ���        ,IO ,�^               :����
'          :typ_B         ,I  ,typ_AllTypesB    :�e���\����
'          :tNew_Hinban   ,I  ,tFullHinban      :�U�֌��i��
'          :tt            ,I  ,Integer          :TopTail����p
'����      :�����w���ɏ]���A���є�����s��
'����      :
Public Function CrAllJudg(typ_b As typ_AllTypesB, tNew_Hinban As tFullHinban, tt As Integer) As FUNCTION_RETURN
    Dim IND         As String                   '�����w��
    Dim bJudg       As Boolean
    Dim i           As Integer
    Dim cnt         As Integer
    Dim typTmList() As typ_TBCMB005
    Dim minwk       As String, maxwk As String
    Dim vTemp       As Variant
    Dim RET         As FUNCTION_RETURN
    Dim Gd_si()     As type_DBDRV_scmzc_fcmkc001c_Siyou
    Dim jCs         As String                               '�u���b�N���i�Ԃ�Cs�ۏ�
    Dim jCsFromTo   As String                               '�u���b�N���i�Ԃ�Cs�ۏ�(FromTo)
    Dim hasSiji     As Boolean                              '�����w������
    Dim sHinban12   As String                               '�i��(12��)
    Dim bJudgXY     As Boolean                              'X������p�t���O�ǉ� 2009/10/22
    Dim bJudgX      As Boolean                              'X������p�t���O�ǉ� 2009/10/22
    Dim bJudgY      As Boolean                              'X������p�t���O�ǉ� 2009/10/22
    Dim Oi          As C_Oi       '2010/03/12
    
    CrAllJudg = FUNCTION_RETURN_FAILURE
    
    sHinban12 = tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond
    
    i = 0
       
    '�����R�[�h���X�g�擾
    If GetCodeList(MSYSCLASS, KCLASS, typTmList) <> FUNCTION_RETURN_SUCCESS Then
        '�����R�[�h���X�g�擾���s
        Exit Function
    End If
    With typ_b
        '' ���������w��(Rs)*****************************************************************
        '�����w���ݒ�
        IND = IIf(tt = BlkTop, "123", "123")
        If JudgSC_B(tt).rs Then
            ' �w���������ꍇ�́ANG�Ƃ��ĕ\��
            .OKNG(tt) = False
            If (InStr(IND, .typ_cr(tt).CRYINDRSCS) <> 0) Then
                If left(.typ_zi.CRYRZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    ' �T���v���������ꍇ�́ANG�Ƃ��ĕ\��
                    If .typ_zi.CRYRZ(tt).SMPLUMU = "0" Then
                        '���R����
                        If Not CrResJudg(1, .typ_si(tt), .typ_zi.CRYRZ(tt), .OKNG(tt), tt) Then
                            '���R���莸�s
                        End If
                    End If
                End If
            End If
            If .OKNG(tt) = False Then
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00100"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDRSCS) <> 0) Then
                .OKNG(tt) = True
                If left(.typ_zi.CRYRZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    ' �T���v���������ꍇ�́AOK�Ƃ��ĕ\��
                    If .typ_zi.CRYRZ(tt).SMPLUMU = "0" Then
                        '���R����
                        If Not CrResJudg(1, .typ_si(tt), .typ_zi.CRYRZ(tt), .OKNG(tt), tt) Then
                            '���R���莸�s
                        End If
                    End If
                End If
            End If
        End If
        
        
        '�����w���ݒ�
        IND = IIf(tt = BlkTop, "123", "123")
        '' ���������w��(Oi)*****************************************************************
        If JudgSC_B(tt).Oi Then
            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).BLOCKNG = False
            .typ_rslt(tt, i).pos = -1                                       ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())       ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                               ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                               ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                     ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
            .typ_rslt(tt, i).SMPLNO = -1                                    ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                    ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
            bJudg = False
            If (InStr(IND, .typ_cr(tt).CRYINDOICS) <> 0) Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.OIZ(tt).POSITION                 ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())       ' ���e
                .typ_rslt(tt, i).SMPLNO = .typ_zi.OIZ(tt).SMPLNO                ' �T���v���m��
                .typ_rslt(tt, i).INFO2 = "�����L"                               ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                               ' ���Q
                If left(.typ_zi.OIZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(tt, i).INFO3 = "����ٖ�"                          ' ���R
                    If .typ_zi.OIZ(tt).SMPLUMU = "0" Then
                        'OI���莸�s
                        .typ_rslt(tt, i).INFO3 = "����Err"                      ' ���Q
                        'OI����
                        If CrOiJudg(.typ_si(tt), .typ_zi.OIZ(tt), bJudg) Then
                            Call GetOiMaxMin(.typ_zi.OIZ(tt), minwk, maxwk)
                            '��ʕ\�����e�ݒ�
                            vTemp = CStr(.typ_zi.OIZ(tt).OIMEAS1)                       ' ���P
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' ���P
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(maxwk, "0.00")     ' ���Q
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(minwk, "0.00")     ' ���R
                            vTemp = CStr(.typ_zi.OIZ(tt).ORGRES)                        ' ���S
                            'ORG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                            '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' ���S
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' ���S
                        Else
                            If .typ_zi.OIZ(tt).ORGRES = -999 Then               ' 2010/03/12 Kameda
                                ReDim Oi.Oi(4)
                                Oi.Oi(0) = .typ_zi.OIZ(tt).OIMEAS1
                                Oi.Oi(1) = .typ_zi.OIZ(tt).OIMEAS2
                                Oi.Oi(2) = .typ_zi.OIZ(tt).OIMEAS3
                                Oi.Oi(3) = .typ_zi.OIZ(tt).OIMEAS4
                                Oi.Oi(4) = .typ_zi.OIZ(tt).OIMEAS5
                                .typ_rslt(tt, i).INFO1 = "�d�l" & .typ_si(tt).HSXONSPT & "�_"   ' ���P
                                .typ_rslt(tt, i).INFO2 = "����" & GetTensu(Oi) & "�_"                                ' ���Q
                                .typ_rslt(tt, i).INFO4 = "�_���s��"     ' ���S
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                               ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                               ' ���茋��
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00101"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDOICS) <> 0) Then
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).OKNG = "OK"                                ' ���茋��
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.OIZ(tt).POSITION             ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())   ' ���e
                .typ_rslt(tt, i).SMPLNO = .typ_zi.OIZ(tt).SMPLNO            ' �T���v���m��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N�Q"                                ' ���茋��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).INFO1 = "�d�l��"                           ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                           ' ���Q
                .typ_rslt(tt, i).INFO3 = "����ٖ�"                          ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
                .typ_rslt(tt, i).hinban = sHinban12                         ' �i��(12��)
                If .typ_zi.OIZ(tt).SMPLUMU = "0" Then
                    'OI���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                          ' ���Q
                    'OI����
                    If CrOiJudg(.typ_si(tt), .typ_zi.OIZ(tt), bJudg) Then
                        Call GetOiMaxMin(.typ_zi.OIZ(tt), minwk, maxwk)
                        '��ʕ\�����e�ݒ�
                        vTemp = CStr(.typ_zi.OIZ(tt).OIMEAS1)                       ' ���P
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' ���P
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(maxwk, "0.00")     ' ���Q
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(minwk, "0.00")     ' ���R
                        vTemp = CStr(.typ_zi.OIZ(tt).ORGRES)                        ' ���S
                        'ORG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                        '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' ���S
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' ���S
                    End If
                End If
                i = i + 1
            End If
        End If
        '' ���������w��(B1)*****************************************************************
        BMDDataSet 1, tt, i, typTmList(), sHinban12
        '' ���������w��(B2)*****************************************************************
        BMDDataSet 2, tt, i, typTmList(), sHinban12
        '' ���������w��(B3)*****************************************************************
        BMDDataSet 3, tt, i, typTmList(), sHinban12
        '' ���������w��(L1)*****************************************************************
        OSFDataSet 1, tt, i, typTmList(), sHinban12, .typ_si(tt).HSXOF1ARPTK    '' ������, .typ_si(tt).HSXOF1ARPTK��ǉ� 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
        '' ���������w��(L2)*****************************************************************
        OSFDataSet 2, tt, i, typTmList(), sHinban12, " "    '' ������, .typ_si(tt).HSXOF1ARPTK��ǉ� 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
        ' ���������w��(L3)*****************************************************************
        OSFDataSet 3, tt, i, typTmList(), sHinban12, " "    '' ������, .typ_si(tt).HSXOF1ARPTK��ǉ� 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
        '' ���������w��(L4)*****************************************************************
        OSFDataSet 4, tt, i, typTmList(), sHinban12, " "    '' ������, .typ_si(tt).HSXOF1ARPTK��ǉ� 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
        '' ���������w��(Cs)*****************************************************************
        If JudgSC_B(tt).Cs And (tt = BlkTail Or .typ_si(tt).HSXCNKHI = "6" Or .typ_si(tt).HSXCNKHI = "9") Then  'TOP/BOT�ۏؑΉ� 09/01/08 ooba
            '��ʕ\�����e������
            .typ_rslt(tt, i).BLOCKNG = False
            .typ_rslt(tt, i).pos = -1                                   ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())   ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                           ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                           ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                 ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
            .typ_rslt(tt, i).SMPLNO = -1                                ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                         ' �i��(12��)
            bJudg = False
            If (InStr(IND, .typ_cr(tt).CRYINDCSCS) <> 0) Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.CSZ(tt).POSITION                 ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())       ' ���e
                .typ_rslt(tt, i).SMPLNO = .typ_zi.CSZ(tt).SMPLNO                ' �T���v���m��
                .typ_rslt(tt, i).INFO2 = "�����L"                               ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                               ' ���Q
                If left(.typ_zi.CSZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(tt, i).INFO3 = "����ٖ�"                          ' ���R
                    If .typ_zi.CSZ(tt).SMPLUMU = "0" Then
                        'Cs���莸�s
                        .typ_rslt(tt, i).INFO3 = "����Err"                      ' ���Q
                        'CS����擾
                        If CrCsjudg(.typ_si(tt), .typ_zi.CSZ(tt), bJudg) Then
                            '��ʕ\�����e�ݒ�
                            vTemp = CStr(.typ_zi.CSZ(tt).CSMEAS)
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00") ' ���P
                            .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                            .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                        End If
                    End If
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                               ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                               ' ���茋��
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00111"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDCSCS) <> 0) Then
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).OKNG = "OK"                                    ' ���茋��
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.CSZ(tt).POSITION                 ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())       ' ���e
                .typ_rslt(tt, i).SMPLNO = .typ_zi.CSZ(tt).SMPLNO                ' �T���v���m��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N�Q"                                   ' ���茋��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).INFO1 = "�d�l��"                               ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                               ' ���Q
                .typ_rslt(tt, i).INFO3 = "����ٖ�"                              ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
                If .typ_zi.CSZ(tt).SMPLUMU = "0" Then
                    'Cs���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                      ' ���Q
                    'CS����擾
                    If CrCsjudg(.typ_si(tt), .typ_zi.CSZ(tt), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CStr(.typ_zi.CSZ(tt).CSMEAS)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00") ' ���P
                        .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                        .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                    End If
                End If
                i = i + 1
            End If
        End If
        '' ���������w��(GD)*****************************************************************
        '�u���b�N���̑S�i�Ԃ̎d�l���擾
        .typ_rslt(tt, i).BLOCKNG = False
        If JudgSC_B(tt).GD Then
            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).pos = -1           ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())       ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                               ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                               ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                     ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
            .typ_rslt(tt, i).SMPLNO = -1                                    ' �T���v���m��
            .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
            bJudg = False
            If (InStr(IND, .typ_cr(tt).CRYINDGDCS) <> 0) Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.GDZ(tt).POSITION                 ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())       ' ���e
                .typ_rslt(tt, i).SMPLNO = .typ_zi.GDZ(tt).SMPLNO                ' �T���v���m��
                .typ_rslt(tt, i).INFO2 = "�����L"                               ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                               ' ���R
                .typ_rslt(tt, i).INFO4 = "���і�"                               ' ���S    '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
                If left(.typ_zi.GDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(tt, i).INFO3 = "����ٖ�"                              ' ���R
                    If .typ_zi.GDZ(tt).SMPLUMU = "0" Then
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(tt, i).INFO3 = "����Err"                          ' ���R
                        'GD����擾
                        If CrGdjudg(.typ_si(tt), .typ_zi.GDZ(tt), bJudg) Then
                            '��ʕ\�����e�ݒ�
                            vTemp = CStr(.typ_zi.GDZ(tt).MSRSDEN)                       ' ���P
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' ���P
                            vTemp = CStr(.typ_zi.GDZ(tt).MSRSLDL)                       ' ���Q
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")        ' ���Q
                            vTemp = CStr(.typ_zi.GDZ(tt).MSRSDVD2)                      ' ���R
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' ���R
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
                            vTemp = CStr(.typ_zi.GDZ(tt).MSZEROMN)                      ' ���S
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' ���S
                            .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & " , "     ' ���S
                            vTemp = CStr(.typ_zi.GDZ(tt).MSZEROMX)                      ' ���S
                            .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & _
                                                     DBData2DispData(vTemp, "0")        ' ���S
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
                        End If
                    End If
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                               ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                               ' ���茋��
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                If pbGDJudgeTbl(3) = False Then
                    gsTbcmy028ErrCode = "00114"
                ElseIf pbGDJudgeTbl(3) = False Then
                    gsTbcmy028ErrCode = "00113"
                Else
                    gsTbcmy028ErrCode = "00112"
                End If
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDGDCS) <> 0) Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).pos = .typ_zi.GDZ(tt).POSITION                 ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())       ' ���e
                .typ_rslt(tt, i).INFO1 = "�d�l��"                               ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                               ' ���Q
                .typ_rslt(tt, i).INFO3 = "����ٖ�"                              ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                .typ_rslt(tt, i).INFO4 = "���і���"                              ' ���S
                .typ_rslt(tt, i).SMPLNO = .typ_zi.GDZ(tt).SMPLNO                ' �T���v���m��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N�Q"                                    ' ���茋��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
                If .typ_zi.GDZ(tt).SMPLUMU = "0" Then
                    'GD���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                          ' ���R
                    'GD����擾
                    If CrGdjudg(.typ_si(tt), .typ_zi.GDZ(tt), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CStr(.typ_zi.GDZ(tt).MSRSDEN)                       ' ���P
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' ���P
                        vTemp = CStr(.typ_zi.GDZ(tt).MSRSLDL)                       ' ���Q
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")        ' ���Q
                        vTemp = CStr(.typ_zi.GDZ(tt).MSRSDVD2)                      ' ���R
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' ���R
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
                        vTemp = CStr(.typ_zi.GDZ(tt).MSZEROMN)                      ' ���S
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' ���S
                        .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & " , "     ' ���S
                        vTemp = CStr(.typ_zi.GDZ(tt).MSZEROMX)                      ' ���S
                        .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & _
                                                 DBData2DispData(vTemp, "0")        ' ���S
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
                    End If
                End If
                i = i + 1
            End If
        End If
        '' ���������w��(T)*****************************************************************
Dim HIN As tFullHinban
Dim LTSPI As String

        If (InStr(IND, .typ_cr(tt).CRYINDTCS) <> 0) Then
            hasSiji = True
        Else
            hasSiji = False
        End If
        bJudg = True                                        '2004/01/15 SystemBrain
        If (JudgSC_B(tt).Lt) And (tt = BlkTail) Then        '2004/01/15 SystemBrain
            bJudg = False                                   '2004/01/15 SystemBrain
        Else                                                '2004/01/15 SystemBrain
            JudgSC_B(tt).Lt = False                         '2004/01/15 SystemBrain
        End If                                              '2004/01/15 SystemBrain
        
        'LT��Bot�[�Ńu���b�N�S��𔻒肷�邱�ƂɂȂ������߁A�uTop�[�i�Ԃ�LT�w���������Bot�ŕ\���v�͕s�v�ƂȂ���
        If (JudgSC_B(tt).Lt) Or (hasSiji And (tt = BlkTail)) Then '�d�l���� or Bot�[�Ō�������
            .typ_rslt(tt, i).BLOCKNG = False
            
            '��ʕ\�����e������
            .typ_rslt(tt, i).pos = .typ_zi.LTZ(tt).POSITION             ' �������J�n�ʒu
            .typ_rslt(tt, i).SMPLNO = -1                                ' �T���v���m��
            .typ_rslt(tt, i).NAIYO = Search_CrCode("T", typTmList())    ' ���e
            If JudgSC_B(tt).Lt Then
                .typ_rslt(tt, i).INFO1 = "�d�l�L"                       ' ���P
            Else
                .typ_rslt(tt, i).INFO1 = "�d�l��"
                bJudg = True
            End If
            If hasSiji Then
                .typ_rslt(tt, i).INFO2 = "�����L"                       ' ���Q
            Else
                .typ_rslt(tt, i).INFO2 = "������"
            End If
            .typ_rslt(tt, i).INFO3 = "���і�"                           ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
            .typ_rslt(tt, i).hinban = sHinban12                         ' �i��(12��)
            
            '���C�t�^�C��
            bJudgX = True   '10������
            '����ƌ��ʓo�^
            If .typ_zi.LTZ(tt).CRYNUM = .typ_si(1).CRYNUM Then
                .typ_rslt(tt, i).pos = .typ_zi.LTZ(tt).POSITION                 ' �������J�n�ʒu
                .typ_rslt(tt, i).SMPLNO = .typ_zi.LTZ(tt).SMPLNO                ' �T���v���m��
                If (.typ_zi.LTZ(tt).SMPLUMU <> "0") Then
                    .typ_rslt(tt, i).INFO3 = "����ٖ�"                          ' ���R
                Else
                    '2005/12/02 add SET���� LT�v�Z�֐�call ->
                    '���C�t�^�C���l���v�Z���Ȃ���
                    Call Sub_LTReCalc(.typ_si(tt), .typ_zi.LTZ(tt))
                    '2005/12/02 add SET���� LT�v�Z�֐�call <-

                    'LT����擾
                    If CrLtjudg(.typ_si(tt), .typ_zi.LTZ(tt), bJudg) Then
''Add Start 2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
                        If CrLt10judg(.typ_si(tt), .typ_zi.LTZ(tt), .typ_cr(tt), bJudgX) Then
''Add End   2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
                            '��ʕ\�����e�ݒ�
                            vTemp = CStr(.typ_zi.LTZ(tt).CALCMEAS)                  ' ���P
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")
                            vTemp = CStr(.typ_zi.LTZ(tt).MEASPEAK)                  ' ���Q
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")
                            .typ_rslt(tt, i).INFO3 = ""                             ' ���R
''Add Start 2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
                            ' ���S
                            If .typ_zi.LTZ(tt).CONVAL = (-1) Then
                                .typ_rslt(tt, i).INFO4 = "NULL"
                            Else
                                .typ_rslt(tt, i).INFO4 = CStr(.typ_zi.LTZ(tt).CONVAL)
                            End If
                        Else
                            .typ_rslt(tt, i).INFO3 = "LT10����Err"                  ' ���R
                        End If
''Add End   2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
                    Else
                        .typ_rslt(tt, i).INFO3 = "����Err"                      ' ���R
                    End If
                End If
            Else    '���тȂ�
                If JudgSC_B(tt).Lt Then bJudg = False
            End If
            
''Add Start 2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
            If bJudg = True Then
                If bJudgX = True Then
                    bJudg = True
                Else
                    bJudg = False
                End If
            End If
''Add End   2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
            
            If (bJudg = False) Then
                .typ_rslt(tt, i).OKNG = "NG"
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00110"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
'====================== Debug Debug =====================================
            ElseIf .typ_si(tt).HSXLTHWS = "S" Then
                .typ_rslt(tt, i).OKNG = "N�Q"                            ' ���茋��
'====================== Debug Debug =====================================
            Else
                .typ_rslt(tt, i).OKNG = "OK"                            ' ���茋��
            End If
            i = i + 1
        End If
        '' ���������w��(EPD)*****************************************************************
        If JudgSC_B(tt).EPD Then
            If tt = BlkTop Then
                .typ_rslt(tt, i).BLOCKNG = False
                If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                ' �������J�n�ʒu
                    .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO               ' �T���v���m��
                    .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())      ' ���e
                    .typ_rslt(tt, i).INFO1 = "�d�l�L"                               ' ���P
                    .typ_rslt(tt, i).INFO2 = "�����L"                               ' ���Q
                    .typ_rslt(tt, i).INFO3 = "���і�"                               ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                    .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
                    bJudg = False
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).INFO3 = "����ٖ�"                          ' ���R
                        If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                            'EPD���莸�s
                            .typ_rslt(tt, i).INFO3 = "����Err"                      ' ���R
                            'EPD����擾
                            If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                                '��ʕ\�����e�ݒ�
                                vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                  ' ���P
                                .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' ���P
                                .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                                .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                            End If
                        End If
                    End If
                    If bJudg = True Then
                        .typ_rslt(tt, i).OKNG = "OK"                               ' ���茋��
                    Else
                        .typ_rslt(tt, i).OKNG = "NG"                               ' ���茋��
                        TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                        gsTbcmy028ErrCode = "00102"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
                    End If
                    i = i + 1
                End If
            Else
                '��ʕ\�����e�ݒ�

'>>>>>  �T���v�������Ή� 2006/05/09�ύX kubota
                .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION    ' �������J�n�ʒu
'<<<<<  �T���v�������Ή� 2006/05/09�ύX kubota
                
                .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())          ' ���e
                .typ_rslt(tt, i).INFO1 = "�d�l�L"                                   ' ���P
                
'>>>>>  �T���v�������Ή� 2006/05/09�ύX kubota
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
'<<<<<  �T���v�������Ή� 2006/05/09�ύX kubota
                
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
                
'>>>>>  �T���v�������Ή� 2006/05/09�ύX kubota
                .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO   ' �T���v���m��
'<<<<<  �T���v�������Ή� 2006/05/09�ύX kubota
                
                .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                bJudg = False
                If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).INFO3 = "����ٖ�"                              ' ���R
                        If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                            .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION            ' �������J�n�ʒu
                            .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO           ' �T���v���m��
                            'EPD���莸�s
                            .typ_rslt(tt, i).INFO3 = "����Err"                          ' ���R
                            'EPD����擾
                            If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                                '��ʕ\�����e�ݒ�
                                vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                  ' ���P
                                .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' ���P
                                .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                                .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                            End If
                        End If
                    End If
                Else
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                ' �������J�n�ʒu
                        .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO               ' �T���v���m��
                        'EPD���莸�s
                        .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���R
                        'EPD����擾
                        If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                            '��ʕ\�����e�ݒ�
                            vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)          ' ���P
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' ���P
                            .typ_rslt(tt, i).INFO2 = ""                                 ' ���Q
                            .typ_rslt(tt, i).INFO3 = ""                                 ' ���R
                        End If
                    End If
                End If
                If bJudg = True Then
                    .typ_rslt(tt, i).OKNG = "OK"                               ' ���茋��
                Else
                    .typ_rslt(tt, i).OKNG = "NG"                               ' ���茋��
                    TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                    gsTbcmy028ErrCode = "00111"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
                End If
                i = i + 1
            End If
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                    ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())          ' ���e
                .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "����ٖ�"                                  ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
                .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO                   ' �T���v���m��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N�Q"                                        ' ���茋��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                    'EPD���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���R
                    'EPD����擾
                    If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                      ' ���P
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' ���P
                        .typ_rslt(tt, i).INFO2 = ""                                 ' ���Q
                        .typ_rslt(tt, i).INFO3 = ""                                 ' ���R
                    End If
                End If
                i = i + 1
            End If
        End If
        'SIRD����f�[�^�ݒ�   2010/02/04 add Kameda
        If tt = BlkTop Then
            If .typ_cr(tt).SIRDKBNY3 = "1" Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.SIRD.POSITION                ' �������J�n�ʒu
                '.typ_rslt(tt, i).SMPLNO = .typ_zi.SIRD.SMPLNO               ' �T���v���m��
                .typ_rslt(tt, i).NAIYO = Search_CrCode("SD", typTmList())       ' ���e
                .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
                bJudg = False
                'SIRD����擾
                If CrSIRDjudg(.typ_si(tt), .typ_zi.SIRD, bJudg) Then
                    '��ʕ\�����e�ݒ�
                    vTemp = CStr(.typ_zi.SIRD.SIRDCNT)                  ' ���P
                    .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' ���P
                    .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                    .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                End If
                If bJudg = True Then
                    .typ_rslt(tt, i).OKNG = "OK"                               ' ���茋��
                Else
                    .typ_rslt(tt, i).OKNG = "NG"                               ' ���茋��
                    TotalJudg = False
                    ''gsTbcmy028ErrCode = ""
                End If
                '�]���҂��Q�Ǝ�  2010/02/18 Kameda
                If .typ_zi.SIRD.NothingFlg = "1" Then
                    .typ_rslt(tt, i).INFO1 = ""                                ' ���P
                    .typ_rslt(tt, i).OKNG = "�]���҂�"                         ' ���茋��
                    'Add Start 2012/01/31 Y.Hitomi
                    TotalJudg = False
                    'Add End 2012/01/31 Y.Hitomi
                End If
                i = i + 1
            ElseIf .typ_cr(tt).SIRDKBNY3 = "2" Then       '2010/02/16 add Kameda
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.SIRD.POSITION                 ' �������J�n�ʒu
                '.typ_rslt(tt, i).SMPLNO = .typ_zi.SIRD.SMPLNO               ' �T���v���m��
                .typ_rslt(tt, i).NAIYO = Search_CrCode("SD", typTmList())       ' ���e
                .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
                'bJudg = False    �\���̂�
                'SIRD�\��
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).INFO1 = "��s�]��"                     ' ���P
                .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                .typ_rslt(tt, i).OKNG = "OK"                            ' ���茋��
                i = i + 1
            End If
        End If
        
        'X������f�[�^�ݒ�   2009/08/12 add Kameda
        '�����p�݂̂Ŕ��� X,Y�͌x�����o��(�w�i�ԁj  2009/10/22 add Kameda
        If tt = BlkTail Then
            If .typ_cr(tt).CRYINDXC1 <> 0 Then
                'If CrXjudg(.typ_si(tt), .typ_zi.XZ, bJudg) Then     2009/10/22 Kameda
                If CrXjudg(.typ_si(tt), .typ_zi.XZ, bJudgXY, bJudgX, bJudgY) Then
                    If bJudgXY Then
                        '.typ_zi.XZ.JUDG = "OK"    2009/10/22
                        .typ_zi.XZ.JUDGXY = "OK"
                    Else
                        '.typ_zi.XZ.JUDG = "NG"    2009/10/22
                        .typ_zi.XZ.JUDGXY = "NG"
                        TotalJudg = False
                    End If
                    '�x�����o�����߂ɍ��ڒǉ�     2009/10/22 Kameda
                    If bJudgX Then
                        .typ_zi.XZ.JUDGX = "OK"
                    Else
                        .typ_zi.XZ.JUDGX = "NG"
                    End If
                    If bJudgY Then
                        .typ_zi.XZ.JUDGY = "OK"
                    Else
                        .typ_zi.XZ.JUDGY = "NG"
                    End If
                End If
            Else
                '.typ_zi.XZ.JUDG = ""     2009/10/22
                .typ_zi.XZ.JUDGXY = ""
                .typ_zi.XZ.JUDGX = ""
                .typ_zi.XZ.JUDGY = ""
            End If
        End If
        
      'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎��є��菈��
        Call CuDecoDataSet_C(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJ(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJLT(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJ2(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
      ''Add End   2011/01/17 SMPK A.Nagamine
        
    End With
    
    CrAllJudg = FUNCTION_RETURN_SUCCESS
End Function

'�T�v      :�R�[�h���擾
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :strCode       ,I  ,String       ,�����R�[�h
'          :CodeData      ,I  ,typ_TBCMB005 ,�R�[�h���X�g�\����
'          :�߂�l        ,O  ,String       ,�Y���R�[�h������
'����      :�R�[�h��񃊃X�g����Y���R�[�h�̏����擾����
'����      :
Private Function Search_CrCode(strCode As String, typ_CodeData() As typ_TBCMB005) As String
    Dim i As Integer
    
    '���X�g����Y���R�[�h�̏��P������
    i = 1
    Do While typ_CodeData(i).INFO1 <> ""
        If strCode = Trim(typ_CodeData(i).CODE) Then
            Search_CrCode = typ_CodeData(i).INFO1
            Exit Function
        End If
        i = i + 1
    Loop
    Search_CrCode = ""
End Function

'�T�v      :OI���ё���lMIN/MAX�l�擾
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_oiz       ,I  ,type_DBDRV_scmzc_fcmkc001c_Oi ,�e���\����
'          :min           ,O  ,String       ,MIN�l
'          :max           ,O  ,String       ,MAX�l
'����      :OI���ё���l����MIN�EMAX�l���擾����
'����      :
Private Sub GetOiMaxMin(typ_oiz As type_DBDRV_scmzc_fcmkc001c_Oi, _
                            OiMin As String, OiMax As String)
    Dim wk(4) As Double

    With typ_oiz
        wk(0) = .OIMEAS1                ' �n������l�P
        wk(1) = .OIMEAS2                ' �n������l�Q
        wk(2) = .OIMEAS3                ' �n������l�R
        wk(3) = .OIMEAS4                ' �n������l�S
        wk(4) = .OIMEAS5                ' �n������l�T
    End With
    OiMin = JudgMin(wk())
    OiMax = JudgMax(wk())
End Sub

'�T�v      :��R����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :iCompFlg      ,I   ,Integer                             :�ΐ͌v�Z�׸�(0:�ΐ͌v�Z�Ȃ�, 1:�ΐ͌v�Z����)
'          :typ_si        ,I   ,type_DBDRV_scmzc_fcmkc001c_Siyou    :�d�l���\����
'          :typ_cryrz     ,I   ,type_DBDRV_scmzc_fcmkc001c_CryR     :RS���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :��R������s��
'����      :
Public Function CrResJudg(iCompFlg As Integer, _
                          typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_cryrz As type_DBDRV_scmzc_fcmkc001c_CryR, _
                          bJudg As Boolean, _
                          tt As Integer) As Boolean
    Dim ErrInfo     As ERROR_INFOMATION     '�G���[���\����
    Dim rs          As C_RES                'RS����\����
    Dim cc          As type_Coefficient
    Dim rp          As type_ResPosCal
    Dim COEF        As Double
    Dim wgtCharge   As Long                 '�ΐ͌v�Z�p�p�����[�^
    Dim wgtTop      As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim wgtTopCut   As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim DM          As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim test As String
    Dim cf As C_COEF
    Dim sMcno2 As String
    Dim sMcno1 As String
    
    bJudg = True
    
    '��R��������ݒ�
    rs.GuaranteeRes.cMeth = typ_si.HSXRSPOH     '����ʒu_��
    rs.GuaranteeRes.cCount = typ_si.HSXRSPOT    '����ʒu_�_
    rs.GuaranteeRes.cPos = typ_si.HSXRSPOI      '����ʒu_��(OSF�̏ꍇ ��)
    rs.GuaranteeRes.cObj = typ_si.HSXRHWYT      '�ۏؕ��@_��
    rs.GuaranteeRes.cJudg = typ_si.HSXRHWYS     '�ۏؕ��@_��
    rs.GuaranteeRes.cBunp = typ_si.HSXRMCAL     '���z�v�Z
    rs.SpecResMin = typ_si.HSXRMIN              ' �i�r�w���R����
    rs.SpecResMax = typ_si.HSXRMAX              ' �i�r�w���R���
    rs.SpecResAveMin = typ_si.HSXRAMIN          ' �i�r�w���R���ω���
    rs.SpecResAveMax = typ_si.HSXRAMAX          ' �i�r�w���R���Ϗ��
    rs.SpecRrg = typ_si.HSXRMBNP                ' �i�r�w���R�ʓ����z
    rs.Res(0) = typ_cryrz.MEAS1                 ' ����l�P
    rs.Res(1) = typ_cryrz.MEAS2                 ' ����l�Q
    rs.Res(2) = typ_cryrz.MEAS3                 ' ����l�R
    rs.Res(3) = typ_cryrz.MEAS4                 ' ����l�S
    rs.Res(4) = typ_cryrz.MEAS5                 ' ����l�T
    rs.RRG = typ_cryrz.RRG                      ' �q�q�f
'--------------- 2008/08/25 INSERT START  By Systech --------------
    rs.DkTmpSiyo = typ_si.HSXDKTMP
    rs.DkTmpJsk = typ_cryrz.HSXDKTMP
'--------------- 2008/08/25 INSERT  END   By Systech --------------
    '��R����
    If CrystalRESJudg(rs, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrResJudg = False
        typ_cryrz.RRG = rs.RRG '�Čv�Z���ʂ��Ƃ肠�����\�����Ȃ�
        If iCompFlg = 1 Then
            typ_b.JudgRes(tt) = rs.JudgRes1 '2001/10/02 S.Sano
            typ_b.JudgRrg(tt) = rs.JudgRrg '2001/10/02 S.Sano
'--------------- 2008/08/25 INSERT START  By Systech --------------
            typ_b.JudgDkTmp(tt) = rs.JudgDkTmp
'--------------- 2008/08/25 INSERT  END   By Systech --------------
        End If
        Exit Function
    End If
    
    typ_cryrz.RRG = rs.RRG '2001/10/02 S.Sano �Čv�Z���ʂ��Ƃ肠�����\�����Ȃ�
    If (iCompFlg = 1) And (ciSmpGetFlg = 0) Then
        typ_b.JudgRes(tt) = rs.JudgRes1 '2001/10/02 S.Sano
        typ_b.JudgRrg(tt) = rs.JudgRrg '2001/10/02 S.Sano
'--------------- 2008/08/25 INSERT START  By Systech --------------
        typ_b.JudgDkTmp(tt) = rs.JudgDkTmp
'--------------- 2008/08/25 INSERT  END   By Systech --------------
    
        '�ΐ͌W���v�Z �}���`����Ή� �Q�Ɗ֐��ύX 2008/04/23 SETsw Nakada
        If GetCoeffParams_new(typ_cryrz.CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
            Debug.Print "�ΐ͌v�Z�p�p�����[�^�̎擾�Ɏ��s����"
        End If
        With typ_b
            cc.DUNMENSEKI = AreaOfCircle(DM)
            cc.TOPSMPLPOS = .typ_zi.CRYRZ(1).POSITION
            cc.BOTSMPLPOS = .typ_zi.CRYRZ(2).POSITION
            cc.CHARGEWEIGHT = wgtCharge
            cc.TOPWEIGHT = wgtTop + wgtTopCut
            cc.TOPRES = .typ_zi.CRYRZ(1).MEAS1
            cc.BOTRES = .typ_zi.CRYRZ(2).MEAS1
            .COEF(tt) = CoefficientCalculation(cc)
            If .Henseki = True Then
                '�����ΐ͌W���v�Z
                cc.DUNMENSEKI = AreaOfCircle(DM)
                cc.TOPSMPLPOS = .typ_rsz(1).POSITION
                cc.BOTSMPLPOS = .typ_rsz(2).POSITION
                cc.CHARGEWEIGHT = wgtCharge
                cc.TOPWEIGHT = wgtTop + wgtTopCut
                cc.TOPRES = .typ_rsz(1).MEAS1
                cc.BOTRES = .typ_rsz(2).MEAS1
                .CRCOEF = CoefficientCalculation(cc)
            End If
            '2005/01/11 �u���b�N�ΐ͔��菈���ǉ� -------
            sMcno1 = Mid(Trim(.typ_si(tt).PRODCOND), 2, 1)
            sMcno2 = Mid(Trim(.typ_si(tt).PRODCOND), 1, 1)
            cf.JudgCOEF = True
            Select Case sMcno1
                Case "H", "I", "J", "K"
                    cf.NP = "n"
                Case "A", "B", "C"
                    Select Case sMcno2
                        Case "A", "B"
                            cf.NP = "p+"
                        Case "1", "2", "3", "4", "5", "6", "7", "C", "E"
                            cf.NP = "p-"
                        Case Else
                            cf.JudgCOEF = False
                    End Select
                Case Else
                    cf.JudgCOEF = False
            End Select
            If cf.JudgCOEF Then
                cf.COEF = .COEF(tt)
                If CrystalCOEFJudg(cf, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                    cf.JudgCOEF = False
                End If
            End If
            '�G���[�\���p�Ƀt���O���Z�b�g����
            If cf.JudgCOEF Then
                .COEFflg = True
            Else
                .COEFflg = False
            End If
            .Hinsyu = cf.NP
            '�ǉ��ް�߈ʒu�̃`�F�b�N
            .DOPEflg = True
            If .typ_si(tt).ADDDPPOS <> 0 Then
                If .typ_si(tt).INGOTPOS <= .typ_si(tt).ADDDPPOS And _
                   .typ_si(tt).INGOTPOS + .typ_si(tt).Length >= .typ_si(tt).ADDDPPOS Then
                   .DOPEflg = False
                End If
            End If
            '2005/01/11 --------------------------------
        
        End With
    End If
    
    If Not rs.JudgRes Then '2001/10/02 S.Sano
        If (iCompFlg = 1) And (ciSmpGetFlg = 0) Then
            With typ_b
                '�ΐ͌v�Z����ăJ�b�g�ʒu���v�Z
                rp.COEFFICIENT = .COEF(tt)
                rp.DUNMENSEKI = AreaOfCircle(DM)
                rp.CHARGEWEIGHT = wgtCharge
                rp.TOPWEIGHT = wgtTop + wgtTopCut
                rp.TOPSMPLPOS = .typ_zi.CRYRZ(1).POSITION
                rp.TOPRES = .typ_zi.CRYRZ(1).MEAS1
                rp.target = IIf(tt = BlkTop, .typ_si(tt).HSXRMAX, .typ_si(tt).HSXRMIN)
                .Cut(tt) = PosCalculation(rp)
            End With
        End If
        bJudg = False
    End If
    CrResJudg = True
    
End Function

'�T�v      :OI����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_oiz       ,I  ,type_DBDRV_scmzc_fcmkc001c_Oi        :OI���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :OI������s��
'����      :
Public Function CrOiJudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                         typ_oiz As type_DBDRV_scmzc_fcmkc001c_Oi, _
                         bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim Oi      As C_Oi                     'Oi����\����
    
    ReDim Oi.Oi(4) As Double
    
    bJudg = True
        
    'OI��������ݒ�
    Oi.GuaranteeOi.cMeth = typ_si.HSXONSPH      '����ʒu_��
    Oi.GuaranteeOi.cCount = typ_si.HSXONSPT     '����ʒu_�_
    Oi.GuaranteeOi.cPos = typ_si.HSXONSPI       '����ʒu_��(OSF�̏ꍇ ��)
    Oi.GuaranteeOi.cObj = typ_si.HSXONHWT       '�ۏؕ��@_��
    Oi.GuaranteeOi.cJudg = typ_si.HSXONHWS      '�ۏؕ��@_��
    Oi.GuaranteeOi.cBunp = typ_si.HSXONMCL      '���z�v�Z
    Oi.SpecOiMin = typ_si.HSXONMIN              '�iSX�_�f�Z�x����
    Oi.SpecOiMax = typ_si.HSXONMAX              '�iSX�_�f�Z�x���
    Oi.SpecORG = typ_si.HSXONMBP                '�iSX�_�f�Z�x�ʓ����z
    Oi.SpecOiAveMin = typ_si.HSXONAMN           '�iSX�_�f�Z�x���ω���
    Oi.SpecOiAveMax = typ_si.HSXONAMX           '�iSX�_�f�Z�x���Ϗ��
    
    Oi.Oi(0) = typ_oiz.OIMEAS1             'Oi����l
    Oi.Oi(1) = typ_oiz.OIMEAS2             'Oi����l
    Oi.Oi(2) = typ_oiz.OIMEAS3             'Oi����l
    Oi.Oi(3) = typ_oiz.OIMEAS4             'Oi����l
    Oi.Oi(4) = typ_oiz.OIMEAS5             'Oi����l
    Oi.ORG = typ_oiz.ORGRES                'ORG�v�Z�l
    '2010/05/10 �Q�l�d�l�Ή� Y.Hitomi
    If Oi.GuaranteeOi.cCount >= "1" Then
        '����_���̃`�F�b�N   2010/03/12 Kameda
        If Oi.Oi(CInt(Oi.GuaranteeOi.cCount) - 1) = -1 Then
            typ_oiz.ORGRES = -999   '����_���s��
            CrOiJudg = False
            bJudg = False
            Exit Function
        End If
    End If
    
    'OI����
    If CrystalOiJudg(Oi, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        CrOiJudg = False
        bJudg = False
        Exit Function
    End If
    
    'ORG�̍Čv�Z�̒l��\������
    typ_oiz.ORGRES = Oi.ORG                   'ORG�v�Z�l
    
    If Oi.JudgOi <> True Or Oi.JudgOrg <> True Then
        bJudg = False
    End If
    
    CrOiJudg = True
End Function

'�T�v      :BMD����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_bmdz      ,I  ,type_DBDRV_scmzc_fcmkc001c_BMD       :BMD���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :bmflg         ,I  ,Integer                              :BMD�׸�(1:BMD1, 2:BMD2, 3:BMD3)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :BMD������s��
'����      :
Public Function CrBmdJudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_bmdz As type_DBDRV_scmzc_fcmkc001c_BMD, _
                          bJudg As Boolean, _
                          bmflg As Integer) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim bm      As C_BMD                    'BMD�\����
    Dim w_Bunpu As Double

    bJudg = True

    'BMD��������ݒ�
    Select Case bmflg
    Case 1
        bm.GuaranteeBmd.cMeth = typ_si.HSXBM1SH   '����ʒu_��
        bm.GuaranteeBmd.cCount = typ_si.HSXBM1ST  '����ʒu_�_
        bm.GuaranteeBmd.cPos = typ_si.HSXBM1SR    '����ʒu_��(OSF�̏ꍇ ��)
        bm.GuaranteeBmd.cObj = typ_si.HSXBM1HT    '�ۏؕ��@_��
        bm.GuaranteeBmd.cJudg = typ_si.HSXBM1HS   '�ۏؕ��@_��
        bm.SpecBmdAveMin = typ_si.HSXBM1AN        '�iSXBMD���ω���
        bm.SpecBmdAveMax = typ_si.HSXBM1AX        '�iSXBMD���Ϗ��
    Case 2
        bm.GuaranteeBmd.cMeth = typ_si.HSXBM2SH   '����ʒu_��
        bm.GuaranteeBmd.cCount = typ_si.HSXBM2ST  '����ʒu_�_
        bm.GuaranteeBmd.cPos = typ_si.HSXBM2SR    '����ʒu_��(OSF�̏ꍇ ��)
        bm.GuaranteeBmd.cObj = typ_si.HSXBM2HT    '�ۏؕ��@_��
        bm.GuaranteeBmd.cJudg = typ_si.HSXBM2HS   '�ۏؕ��@_��
        bm.SpecBmdAveMin = typ_si.HSXBM2AN        '�iSXBMD���ω���
        bm.SpecBmdAveMax = typ_si.HSXBM2AX        '�iSXBMD���Ϗ��
    Case 3
        bm.GuaranteeBmd.cMeth = typ_si.HSXBM3SH   '����ʒu_��
        bm.GuaranteeBmd.cCount = typ_si.HSXBM3ST  '����ʒu_�_
        bm.GuaranteeBmd.cPos = typ_si.HSXBM3SR    '����ʒu_��(OSF�̏ꍇ ��)
        bm.GuaranteeBmd.cObj = typ_si.HSXBM3HT    '�ۏؕ��@_��
        bm.GuaranteeBmd.cJudg = typ_si.HSXBM3HS   '�ۏؕ��@_��
        bm.SpecBmdAveMin = typ_si.HSXBM3AN        '�iSXBMD���ω���
        bm.SpecBmdAveMax = typ_si.HSXBM3AX        '�iSXBMD���Ϗ��
    End Select
    
    bm.BMD(0) = typ_bmdz.MEAS1                      'BMD����l
    bm.BMD(1) = typ_bmdz.MEAS2                      'BMD����l
    bm.BMD(2) = typ_bmdz.MEAS3                      'BMD����l
    bm.BMD(3) = typ_bmdz.MEAS4                      'BMD����l
    bm.BMD(4) = typ_bmdz.MEAS5                      'BMD����l
    bm.Min = typ_bmdz.MEASMIN                       '�ŏ��l
    bm.max = typ_bmdz.MEASMAX                       '�ő�l
    bm.AVE = typ_bmdz.MEASAVE                       '���ϒl
    
    w_Bunpu = typ_bmdz.BMDMNBUNP

    If typ_si.HSXBM1HS = "H" And typ_si.HSXBM1HT <> "" Then
       If bmflg = "1" Then
          If typ_si.HSXBMD1MBP < w_Bunpu And typ_si.HSXBMD1MBP <> -1 Then
             bJudg = False
          End If
       ElseIf bmflg = "2" Then
          If typ_si.HSXBMD2MBP < w_Bunpu And typ_si.HSXBMD2MBP <> -1 Then
             bJudg = False
          End If
       ElseIf bmflg = "3" Then
          If typ_si.HSXBMD3MBP < w_Bunpu And typ_si.HSXBMD3MBP <> -1 Then
             bJudg = False
          End If
       End If
    End If
    
    'BMD����
    If CrystalBMDJudg(bm, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrBmdJudg = False
        Exit Function
    End If
    If bm.JudgBmd <> True Then
        bJudg = False
    End If
    
    CrBmdJudg = True

End Function

'�T�v      :OSF����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_osfz      ,I  ,type_DBDRV_scmzc_fcmkc001c_OSF       :OSF���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :osfflg        ,I  ,Integer                              :OSF�׸�(1:OSF1, 2:OSF2, 3:OSF3, 4:OSF4)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :OSF������s��
'����      :
Public Function CrOsfJudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_Osfz As type_DBDRV_scmzc_fcmkc001c_OSF, _
                          bJudg As Boolean, _
                          osfflg As Integer) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim os      As C_OSF                    'OSF�\����
    Dim w_RD    As String
    Dim j       As Integer

    bJudg = True

    'OSF��������ݒ�
    Select Case osfflg
    Case 1
        os.GuaranteeOsf.cMeth = typ_si.HSXOF1SH   '����ʒu_��
        os.GuaranteeOsf.cCount = typ_si.HSXOF1ST  '����ʒu_�_
        os.GuaranteeOsf.cPos = typ_si.HSXOF1SR    '����ʒu_��(OSF�̏ꍇ ��)
        os.GuaranteeOsf.cObj = typ_si.HSXOF1HT    '�ۏؕ��@_��
        os.GuaranteeOsf.cJudg = typ_si.HSXOF1HS   '�ۏؕ��@_��
        os.SpecOsfAveMax = typ_si.HSXOF1AX        '�iSXOSF���Ϗ��
        os.SpecOsfMax = typ_si.HSXOF1MX           '�iSX���
    Case 2
        os.GuaranteeOsf.cMeth = typ_si.HSXOF2SH   '����ʒu_��
        os.GuaranteeOsf.cCount = typ_si.HSXOF2ST  '����ʒu_�_
        os.GuaranteeOsf.cPos = typ_si.HSXOF2SR    '����ʒu_��(OSF�̏ꍇ ��)
        os.GuaranteeOsf.cObj = typ_si.HSXOF2HT    '�ۏؕ��@_��
        os.GuaranteeOsf.cJudg = typ_si.HSXOF2HS   '�ۏؕ��@_��
        os.SpecOsfAveMax = typ_si.HSXOF2AX        '�iSXOSF���Ϗ��
        os.SpecOsfMax = typ_si.HSXOF2MX           '�iSX���
    Case 3
        os.GuaranteeOsf.cMeth = typ_si.HSXOF3SH   '����ʒu_��
        os.GuaranteeOsf.cCount = typ_si.HSXOF3ST  '����ʒu_�_
        os.GuaranteeOsf.cPos = typ_si.HSXOF3SR    '����ʒu_��(OSF�̏ꍇ ��)
        os.GuaranteeOsf.cObj = typ_si.HSXOF3HT    '�ۏؕ��@_��
        os.GuaranteeOsf.cJudg = typ_si.HSXOF3HS   '�ۏؕ��@_��
        os.SpecOsfAveMax = typ_si.HSXOF3AX        '�iSXOSF���Ϗ��
        os.SpecOsfMax = typ_si.HSXOF3MX           '�iSX���
    Case 4
        os.GuaranteeOsf.cMeth = typ_si.HSXOF4SH   '����ʒu_��
        os.GuaranteeOsf.cCount = typ_si.HSXOF4ST  '����ʒu_�_
        os.GuaranteeOsf.cPos = typ_si.HSXOF4SR    '����ʒu_��(OSF�̏ꍇ ��)
        os.GuaranteeOsf.cObj = typ_si.HSXOF4HT    '�ۏؕ��@_��
        os.GuaranteeOsf.cJudg = typ_si.HSXOF4HS   '�ۏؕ��@_��
        os.SpecOsfAveMax = typ_si.HSXOF4AX        '�iSXOSF���Ϗ��
        os.SpecOsfMax = typ_si.HSXOF4MX           '�iSX���
    End Select

    os.OSF(0) = typ_Osfz.MEAS1        'OSF����l
    os.OSF(1) = typ_Osfz.MEAS2        'OSF����l
    os.OSF(2) = typ_Osfz.MEAS3        'OSF����l
    os.OSF(3) = typ_Osfz.MEAS4        'OSF����l
    os.OSF(4) = typ_Osfz.MEAS5        'OSF����l
    os.OSF(5) = typ_Osfz.MEAS6        'OSF����l
    os.OSF(6) = typ_Osfz.MEAS7        'OSF����l
    os.OSF(7) = typ_Osfz.MEAS8        'OSF����l
    os.OSF(8) = typ_Osfz.MEAS9        'OSF����l
    os.OSF(9) = typ_Osfz.MEAS10       'OSF����l
    os.OSF(10) = typ_Osfz.MEAS11      'OSF����l
    os.OSF(11) = typ_Osfz.MEAS12      'OSF����l
    os.OSF(12) = typ_Osfz.MEAS13      'OSF����l
    os.OSF(13) = typ_Osfz.MEAS14      'OSF����l
    os.OSF(14) = typ_Osfz.MEAS15      'OSF����l
    os.OSF(15) = typ_Osfz.MEAS16      'OSF����l
    os.OSF(16) = typ_Osfz.MEAS17      'OSF����l
    os.OSF(17) = typ_Osfz.MEAS18      'OSF����l
    os.OSF(18) = typ_Osfz.MEAS19      'OSF����l
    os.OSF(19) = typ_Osfz.MEAS20      'OSF����l
    os.max = typ_Osfz.CALCMAX         '�ő�l
    os.AVE = typ_Osfz.CALCAVE         '���ϒl
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    os.ARPTK = typ_si.HSXOF1ARPTK       '�iSXOSF1(ArAN)�p�^���敪
    os.ARMIN = typ_si.HSXOFARMIN        '�iSXOSF(ArAN)����
    os.ARMAX = typ_si.HSXOFARMAX        '�iSXOSF(ArAN)���
    os.ARMHMX = typ_si.HSXOFARMHMX      '�iSXOSF(ArAN)�ʓ�����
    os.CALCMH = typ_Osfz.CALCMH         '�ʓ���(MAX/MIN)
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

    w_RD = typ_Osfz.OSFRD1 + typ_Osfz.OSFRD2 + typ_Osfz.OSFRD3

    If os.GuaranteeOsf.cJudg = "H" And os.GuaranteeOsf.cObj <> "" Then
       If osfflg = 1 Then
           Select Case typ_si.HSXOSF1PTK
               Case "1"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Then bJudg = False
                   Next
               Case "2"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
               Case "3"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Or Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
            End Select
       ElseIf osfflg = 2 Then
           Select Case typ_si.HSXOSF2PTK
               Case "1"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Then bJudg = False
                   Next
               Case "2"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
               Case "3"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Or Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
            End Select
       ElseIf osfflg = 3 Then
           Select Case typ_si.HSXOSF3PTK
               Case "1"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Then bJudg = False
                   Next
               Case "2"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
               Case "3"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Or Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
            End Select
       ElseIf osfflg = 4 Then
           Select Case typ_si.HSXOSF4PTK
               Case "1"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Then bJudg = False
                   Next
               Case "2"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
               Case "3"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Or Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
            End Select
       End If
    End If
    
    'OSF����
    If CrystalOSFJudg(os, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrOsfJudg = False
        Exit Function
    End If
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    If osfflg = 1 Then
        'OSF����
        If CrystalOSFJudg_02(os, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
            bJudg = False
            CrOsfJudg = False
            Exit Function
        End If
        
        os.JudgOsf = os.JudgOsf And os.JudgOsfPtn
        
        If os.ARPTK = "1" Or os.ARPTK = "2" Then
            If os.JudgOsfPtn = True Then
                typ_Osfz.PTNJUDGRES = "1"
            Else
                typ_Osfz.PTNJUDGRES = "9"
            End If
        Else
            typ_Osfz.PTNJUDGRES = " "
        End If
    End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    
    If os.JudgOsf <> True Then
        bJudg = False
    End If
    
    CrOsfJudg = True

End Function

'�T�v      :CS����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_csz       ,I  ,type_DBDRV_scmzc_fcmkc001c_CS        :CS���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :CS������s��
'����      :
Public Function CrCsjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                         typ_csz As type_DBDRV_scmzc_fcmkc001c_CS, _
                         bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim Cs      As C_Cs                     'CS�\����
    
    bJudg = True
        
    'CS��������ݒ�
    Cs.GuaranteeCs.cMeth = typ_si.HSXCNSPH   '����ʒu_��
    Cs.GuaranteeCs.cCount = typ_si.HSXCNSPT  '����ʒu_�_
    Cs.GuaranteeCs.cPos = typ_si.HSXCNSPI    '����ʒu_��(OSF�̏ꍇ ��)
    Cs.GuaranteeCs.cObj = typ_si.HSXCNHWT    '�ۏؕ��@_��
    Cs.GuaranteeCs.cJudg = typ_si.HSXCNHWS   '�ۏؕ��@_��
    Cs.SpecCsMin = typ_si.HSXCNMIN           '�iSX�Y�f�Z�x����
    Cs.SpecCsMax = typ_si.HSXCNMAX           '�iSX�Y�f�Z�x���
    Cs.SpecCsKHI = typ_si.HSXCNKHI           '�����p�x_�� 09/01/08 ooba
    Cs.Cs = typ_csz.CSMEAS                   'Cs����l
    
    'CS����
    If CrystalCsJudg(Cs, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrCsjudg = False
        Exit Function
    End If
    
    If Cs.JudgCs <> True Then
        bJudg = False
    End If
    
    CrCsjudg = True

End Function

'�T�v      :GD����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_gdz       ,I  ,type_DBDRV_scmzc_fcmkc001c_GD        :GD���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :GD������s��
'����      :
Public Function CrGdjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                         typ_gdz As type_DBDRV_scmzc_fcmkc001c_GD, _
                         bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim GD      As C_GD                     'GD�\����
    
    bJudg = True
        
    'GD��������ݒ�
    GD.GuaranteeDen.cMeth = ""                  '����ʒu_��
    GD.GuaranteeDen.cCount = ""                 '����ʒu_�_
    GD.GuaranteeDen.cPos = ""                   '����ʒu_��(OSF�̏ꍇ ��)
    GD.GuaranteeDen.cObj = typ_si.HSXDENHT      '�ۏؕ��@_��
    GD.GuaranteeDen.cJudg = typ_si.HSXDENHS     '�ۏؕ��@_��
    
    GD.GuaranteeLdl.cMeth = ""                  '����ʒu_��
    GD.GuaranteeLdl.cCount = ""                 '����ʒu_�_
    GD.GuaranteeLdl.cPos = ""                   '����ʒu_��(OSF�̏ꍇ ��)
    GD.GuaranteeLdl.cObj = typ_si.HSXLDLHT      '�ۏؕ��@_��
    GD.GuaranteeLdl.cJudg = typ_si.HSXLDLHS     '�ۏؕ��@_��
    
    GD.GuaranteeDvd2.cMeth = ""                 '����ʒu_��
    GD.GuaranteeDvd2.cCount = ""                '����ʒu_�_
    GD.GuaranteeDvd2.cPos = ""                  '����ʒu_��(OSF�̏ꍇ ��)
    GD.GuaranteeDvd2.cObj = typ_si.HSXDVDHT     '�ۏؕ��@_��
    GD.GuaranteeDvd2.cJudg = typ_si.HSXDVDHS    '�ۏؕ��@_��
    
    GD.JudgFlagDen = typ_si.HSXDENKU            '�iSXDen�����L��
    GD.JudgFlagLdl = typ_si.HSXLDLKU            '�iSXL/DL�����L��
    GD.JudgFlagDvd2 = typ_si.HSXDVDKU           '�iSXDVD2�����L��
    
    GD.SpecDenMin = typ_si.HSXDENMN             '�iSXDen����
    GD.SpecDenMax = typ_si.HSXDENMX             '�iSXDen���
    GD.SpecLdlMin = typ_si.HSXLDLMN             '�iSXLdl����
    GD.SpecLdlMax = typ_si.HSXLDLMX             '�iSXLdl���
    GD.SpecDvd2Min = typ_si.HSXDVDMN            '�iSXDvd2����
    GD.SpecDvd2Max = typ_si.HSXDVDMX            '�iSXDvd2���
'*** UPDATE �� Y.SIMIZU 2005/10/13 �iSXGDײݐ��ǉ�
    GD.SpecGdLine = typ_si.HSXGDLINE            '�iSXGDײݐ�
'*** UPDATE �� Y.SIMIZU 2005/10/13 �iSXGDײݐ��ǉ�
    
    GD.Den = typ_gdz.MSRSDEN                    'Den�v�Z�l
    GD.Ldl = typ_gdz.MSRSLDL                    'L/DL�v�Z�l
    GD.Dvd2 = typ_gdz.MSRSDVD2                  'Dvd2�v�Z�l
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    GD.ZeroLdlMin = typ_si.HSXLDLRMN            '�iSXL/DL�A��0����
    GD.ZeroLdlMax = typ_si.HSXLDLRMX            '�iSXL/DL�A��0���
    GD.LdlMin = typ_gdz.MSZEROMN                'L/DL0�A�����ŏ��l
    GD.LdlMax = typ_gdz.MSZEROMX                'L/DL0�A�����ő�l
    GD.GDPTK = typ_si.HSXGDPTK                  '�i�r�w�f�c�p�^���敪
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    
'*** UPDATE �� Y.SIMIZU 2005/10/13 ײݐ��Ή�
    'GDײݐ���3����4.5����5�łȂ��ꍇ�͔���װ
    If GD.SpecGdLine <> 3 And GD.SpecGdLine <> 4.5 And GD.SpecGdLine <> 5 Then
        bJudg = False
        CrGdjudg = False
        Exit Function
    End If
    
    'GDײݐ����̎��т����邩����������
    If ChkGD_Data(typ_gdz, GD) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrGdjudg = False
        Exit Function
    End If
'*** UPDATE �� Y.SIMIZU 2005/10/13 ײݐ��Ή�

    'GD����
    If CrystalGDJudg(GD, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrGdjudg = False
        Exit Function
    End If
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    GD.JudgLdl = GD.JudgLdl And GD.JudgLdlPtn

    If GD.GDPTK = "1" Or GD.GDPTK = "2" Then
        If GD.JudgLdlPtn = True Then
            typ_gdz.PTNJUDGRES = "1"
        Else
            typ_gdz.PTNJUDGRES = "9"
        End If
    Else
        typ_gdz.PTNJUDGRES = " "
    End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    
    If GD.JudgDen <> True Or GD.JudgLdl <> True Or GD.JudgDvd2 <> True Then
        bJudg = False
    End If
    
'--------------- 2008/07/25 INSERT START  By Systech ---------------
    pbGDJudgeTbl(1) = GD.JudgDen
    pbGDJudgeTbl(2) = GD.JudgDvd2
    pbGDJudgeTbl(3) = GD.JudgLdl
'--------------- 2008/07/25 INSERT  END   By Systech ---------------

    CrGdjudg = True

End Function

'�T�v      :�d�l��GDײݐ�������l�����݂��邩����������
'���Ұ�    :�ϐ���      ,IO ,�^                             :����
'          :tGDdata    ,I   ,type_DBDRV_scmzc_fcmkc001c_GD  :GD���э\����
'          :GD         ,O   ,C_GD                           :GD�d�l�\����
'          :�߂�l      ,O  ,FUNCTION_RETURN                :���� = FUNCTION_RETURN_SUCCESS : OK
'                                                           FUNCTION_RETURN_FAILURE : NG
'����      :
'����      :05/10/13 Y.SIMIZU
Private Function ChkGD_Data(tGDdata As type_DBDRV_scmzc_fcmkc001c_GD, GD As C_GD) As FUNCTION_RETURN
    Dim iCnt            As Integer
    Dim iPoint          As Integer
    Dim iLine           As Integer
    Dim iTden(5, 15)    As Integer
    Dim iTldl(5, 15)    As Integer
    Dim iTdvd2(5)       As Integer

    'GD����l���
    With tGDdata
        iTden(1, 1) = .MS01DEN1         '����l01 Den1
        iTden(2, 1) = .MS01DEN2         '����l01 Den2
        iTden(3, 1) = .MS01DEN3         '����l01 Den3
        iTden(4, 1) = .MS01DEN4         '����l01 Den4
        iTden(5, 1) = .MS01DEN5         '����l01 Den5
        iTden(1, 2) = .MS02DEN1         '����l02 Den1
        iTden(2, 2) = .MS02DEN2         '����l02 Den2
        iTden(3, 2) = .MS02DEN3         '����l02 Den3
        iTden(4, 2) = .MS02DEN4         '����l02 Den4
        iTden(5, 2) = .MS02DEN5         '����l02 Den5
        iTden(1, 3) = .MS03DEN1         '����l03 Den1
        iTden(2, 3) = .MS03DEN2         '����l03 Den2
        iTden(3, 3) = .MS03DEN3         '����l03 Den3
        iTden(4, 3) = .MS03DEN4         '����l03 Den4
        iTden(5, 3) = .MS03DEN5         '����l03 Den5
        iTden(1, 4) = .MS04DEN1         '����l04 Den1
        iTden(2, 4) = .MS04DEN2         '����l04 Den2
        iTden(3, 4) = .MS04DEN3         '����l04 Den3
        iTden(4, 4) = .MS04DEN4         '����l04 Den4
        iTden(5, 4) = .MS04DEN5         '����l04 Den5
        iTden(1, 5) = .MS05DEN1         '����l05 Den1
        iTden(2, 5) = .MS05DEN2         '����l05 Den2
        iTden(3, 5) = .MS05DEN3         '����l05 Den3
        iTden(4, 5) = .MS05DEN4         '����l05 Den4
        iTden(5, 5) = .MS05DEN5         '����l05 Den5
        iTden(1, 6) = .MS06DEN1         '����l06 Den1
        iTden(2, 6) = .MS06DEN2         '����l06 Den2
        iTden(3, 6) = .MS06DEN3         '����l06 Den3
        iTden(4, 6) = .MS06DEN4         '����l06 Den4
        iTden(5, 6) = .MS06DEN5         '����l06 Den5
        iTden(1, 7) = .MS07DEN1         '����l07 Den1
        iTden(2, 7) = .MS07DEN2         '����l07 Den2
        iTden(3, 7) = .MS07DEN3         '����l07 Den3
        iTden(4, 7) = .MS07DEN4         '����l07 Den4
        iTden(5, 7) = .MS07DEN5         '����l07 Den5
        iTden(1, 8) = .MS08DEN1         '����l08 Den1
        iTden(2, 8) = .MS08DEN2         '����l08 Den2
        iTden(3, 8) = .MS08DEN3         '����l08 Den3
        iTden(4, 8) = .MS08DEN4         '����l08 Den4
        iTden(5, 8) = .MS08DEN5         '����l08 Den5
        iTden(1, 9) = .MS09DEN1         '����l09 Den1
        iTden(2, 9) = .MS09DEN2         '����l09 Den2
        iTden(3, 9) = .MS09DEN3         '����l09 Den3
        iTden(4, 9) = .MS09DEN4         '����l09 Den4
        iTden(5, 9) = .MS09DEN5         '����l09 Den5
        iTden(1, 10) = .MS10DEN1        '����l10 Den1
        iTden(2, 10) = .MS10DEN2        '����l10 Den2
        iTden(3, 10) = .MS10DEN3        '����l10 Den3
        iTden(4, 10) = .MS10DEN4        '����l10 Den4
        iTden(5, 10) = .MS10DEN5        '����l10 Den5
        iTden(1, 11) = .MS11DEN1        '����l11 Den1
        iTden(2, 11) = .MS11DEN2        '����l11 Den2
        iTden(3, 11) = .MS11DEN3        '����l11 Den3
        iTden(4, 11) = .MS11DEN4        '����l11 Den4
        iTden(5, 11) = .MS11DEN5        '����l11 Den5
        iTden(1, 12) = .MS12DEN1        '����l12 Den1
        iTden(2, 12) = .MS12DEN2        '����l12 Den2
        iTden(3, 12) = .MS12DEN3        '����l12 Den3
        iTden(4, 12) = .MS12DEN4        '����l12 Den4
        iTden(5, 12) = .MS12DEN5        '����l12 Den5
        iTden(1, 13) = .MS13DEN1        '����l13 Den1
        iTden(2, 13) = .MS13DEN2        '����l13 Den2
        iTden(3, 13) = .MS13DEN3        '����l13 Den3
        iTden(4, 13) = .MS13DEN4        '����l13 Den4
        iTden(5, 13) = .MS13DEN5        '����l13 Den5
        iTden(1, 14) = .MS14DEN1        '����l14 Den1
        iTden(2, 14) = .MS14DEN2        '����l14 Den2
        iTden(3, 14) = .MS14DEN3        '����l14 Den3
        iTden(4, 14) = .MS14DEN4        '����l14 Den4
        iTden(5, 14) = .MS14DEN5        '����l14 Den5
        iTden(1, 15) = .MS15DEN1        '����l15 Den1
        iTden(2, 15) = .MS15DEN2        '����l15 Den2
        iTden(3, 15) = .MS15DEN3        '����l15 Den3
        iTden(4, 15) = .MS15DEN4        '����l15 Den4
        iTden(5, 15) = .MS15DEN5        '����l15 Den5
        
        iTldl(1, 1) = .MS01LDL1         '����l01 L/DL1
        iTldl(2, 1) = .MS01LDL2         '����l01 L/DL2
        iTldl(3, 1) = .MS01LDL3         '����l01 L/DL3
        iTldl(4, 1) = .MS01LDL4         '����l01 L/DL4
        iTldl(5, 1) = .MS01LDL5         '����l01 L/DL5
        iTldl(1, 2) = .MS02LDL1         '����l02 L/DL1
        iTldl(2, 2) = .MS02LDL2         '����l02 L/DL2
        iTldl(3, 2) = .MS02LDL3         '����l02 L/DL3
        iTldl(4, 2) = .MS02LDL4         '����l02 L/DL4
        iTldl(5, 2) = .MS02LDL5         '����l02 L/DL5
        iTldl(1, 3) = .MS03LDL1         '����l03 L/DL1
        iTldl(2, 3) = .MS03LDL2         '����l03 L/DL2
        iTldl(3, 3) = .MS03LDL3         '����l03 L/DL3
        iTldl(4, 3) = .MS03LDL4         '����l03 L/DL4
        iTldl(5, 3) = .MS03LDL5         '����l03 L/DL5
        iTldl(1, 4) = .MS04LDL1         '����l04 L/DL1
        iTldl(2, 4) = .MS04LDL2         '����l04 L/DL2
        iTldl(3, 4) = .MS04LDL3         '����l04 L/DL3
        iTldl(4, 4) = .MS04LDL4         '����l04 L/DL4
        iTldl(5, 4) = .MS04LDL5         '����l04 L/DL5
        iTldl(1, 5) = .MS05LDL1         '����l05 L/DL1
        iTldl(2, 5) = .MS05LDL2         '����l05 L/DL2
        iTldl(3, 5) = .MS05LDL3         '����l05 L/DL3
        iTldl(4, 5) = .MS05LDL4         '����l05 L/DL4
        iTldl(5, 5) = .MS05LDL5         '����l05 L/DL5
        iTldl(1, 6) = .MS06LDL1         '����l06 L/DL1
        iTldl(2, 6) = .MS06LDL2         '����l06 L/DL2
        iTldl(3, 6) = .MS06LDL3         '����l06 L/DL3
        iTldl(4, 6) = .MS06LDL4         '����l06 L/DL4
        iTldl(5, 6) = .MS06LDL5         '����l06 L/DL5
        iTldl(1, 7) = .MS07LDL1         '����l07 L/DL1
        iTldl(2, 7) = .MS07LDL2         '����l07 L/DL2
        iTldl(3, 7) = .MS07LDL3         '����l07 L/DL3
        iTldl(4, 7) = .MS07LDL4         '����l07 L/DL4
        iTldl(5, 7) = .MS07LDL5         '����l07 L/DL5
        iTldl(1, 8) = .MS08LDL1         '����l08 L/DL1
        iTldl(2, 8) = .MS08LDL2         '����l08 L/DL2
        iTldl(3, 8) = .MS08LDL3         '����l08 L/DL3
        iTldl(4, 8) = .MS08LDL4         '����l08 L/DL4
        iTldl(5, 8) = .MS08LDL5         '����l08 L/DL5
        iTldl(1, 9) = .MS09LDL1         '����l09 L/DL1
        iTldl(2, 9) = .MS09LDL2         '����l09 L/DL2
        iTldl(3, 9) = .MS09LDL3         '����l09 L/DL3
        iTldl(4, 9) = .MS09LDL4         '����l09 L/DL4
        iTldl(5, 9) = .MS09LDL5         '����l09 L/DL5
        iTldl(1, 10) = .MS10LDL1        '����l10 L/DL1
        iTldl(2, 10) = .MS10LDL2        '����l10 L/DL2
        iTldl(3, 10) = .MS10LDL3        '����l10 L/DL3
        iTldl(4, 10) = .MS10LDL4        '����l10 L/DL4
        iTldl(5, 10) = .MS10LDL5        '����l10 L/DL5
        iTldl(1, 11) = .MS11LDL1        '����l11 L/DL1
        iTldl(2, 11) = .MS11LDL2        '����l11 L/DL2
        iTldl(3, 11) = .MS11LDL3        '����l11 L/DL3
        iTldl(4, 11) = .MS11LDL4        '����l11 L/DL4
        iTldl(5, 11) = .MS11LDL5        '����l11 L/DL5
        iTldl(1, 12) = .MS12LDL1        '����l12 L/DL1
        iTldl(2, 12) = .MS12LDL2        '����l12 L/DL2
        iTldl(3, 12) = .MS12LDL3        '����l12 L/DL3
        iTldl(4, 12) = .MS12LDL4        '����l12 L/DL4
        iTldl(5, 12) = .MS12LDL5        '����l12 L/DL5
        iTldl(1, 13) = .MS13LDL1        '����l13 L/DL1
        iTldl(2, 13) = .MS13LDL2        '����l13 L/DL2
        iTldl(3, 13) = .MS13LDL3        '����l13 L/DL3
        iTldl(4, 13) = .MS13LDL4        '����l13 L/DL4
        iTldl(5, 13) = .MS13LDL5        '����l13 L/DL5
        iTldl(1, 14) = .MS14LDL1        '����l14 L/DL1
        iTldl(2, 14) = .MS14LDL2        '����l14 L/DL2
        iTldl(3, 14) = .MS14LDL3        '����l14 L/DL3
        iTldl(4, 14) = .MS14LDL4        '����l14 L/DL4
        iTldl(5, 14) = .MS14LDL5        '����l14 L/DL5
        iTldl(1, 15) = .MS15LDL1        '����l15 L/DL1
        iTldl(2, 15) = .MS15LDL2        '����l15 L/DL2
        iTldl(3, 15) = .MS15LDL3        '����l15 L/DL3
        iTldl(4, 15) = .MS15LDL4        '����l15 L/DL4
        iTldl(5, 15) = .MS15LDL5        '����l15 L/DL5
        
        iTdvd2(1) = .MS01DVD2           '����l01 DVD2
        iTdvd2(2) = .MS02DVD2           '����l02 DVD2
        iTdvd2(3) = .MS03DVD2           '����l03 DVD2
        iTdvd2(4) = .MS04DVD2           '����l04 DVD2
        iTdvd2(5) = .MS05DVD2           '����l05 DVD2
    End With
    
    'Den�̎d�l�������L��,�ۏؗL��̏ꍇ
    If (GD.JudgFlagDen = "1" And GD.GuaranteeDen.cJudg = JudgCodeC01) Or _
       (GD.JudgFlagDvd2 = "1" And GD.GuaranteeDvd2.cJudg = JudgCodeC01 And GD.SpecDvd2Min = 0 And GD.SpecDvd2Max = 0) Then
    
        'Den�̑���l��ײݐ������邩������
        For iPoint = 1 To 15
            '����_7�܂�
            If iPoint <= 7 Then
                '�d�l��3ײ݂̏ꍇ
                If GD.SpecGdLine = 3 Then
                    iLine = 3
                '�d�l��4.5ײݖ���5ײ݂̏ꍇ
                ElseIf GD.SpecGdLine = 4.5 Or GD.SpecGdLine = 5 Then
                    iLine = 5
                End If
            '����_8����
            Else
                '�d�l��3ײ݂̏ꍇ
                If GD.SpecGdLine = 3 Then
                    iLine = 3
                '�d�l��4.5ײ݂̏ꍇ
                ElseIf GD.SpecGdLine = 4.5 Then
                    iLine = 4
                '�d�l��5ײ݂̏ꍇ
                Else
                    iLine = 5
                End If
            End If
            
            For iCnt = 1 To iLine
                'DEN�̑���l���Ȃ��ꍇ
                If iTden(iCnt, iPoint) = -1 Then
                    ChkGD_Data = FUNCTION_RETURN_FAILURE
                    '����װ(�����𔲂���)
                    Exit Function
                End If
            Next iCnt
        Next iPoint
    End If
    
    'LDL�̎d�l�������L��,�ۏؗL��̏ꍇ
    If GD.JudgFlagLdl = "1" And GD.GuaranteeLdl.cJudg = JudgCodeC01 Then
    
        'L/DL�̑���l��ײݐ������邩������
        For iPoint = 1 To 15
            '����_7�܂�
            If iPoint <= 7 Then
                '�d�l��3ײ݂̏ꍇ
                If GD.SpecGdLine = 3 Then
                    iLine = 3
                '�d�l��4.5ײ݂̏ꍇ
                ElseIf GD.SpecGdLine = 4.5 Or GD.SpecGdLine = 5 Then
                    iLine = 5
                End If
            '����_8����
            Else
                '�d�l��3ײ݂̏ꍇ
                If GD.SpecGdLine = 3 Then
                    iLine = 3
                '�d�l��4.5ײ݂̏ꍇ
                ElseIf GD.SpecGdLine = 4.5 Then
                    iLine = 4
                '�d�l��5ײ݂̏ꍇ
                Else
                    iLine = 5
                End If
            End If
            
            For iCnt = 1 To iLine
                'L/DL�̑���l���Ȃ��ꍇ
                If iTldl(iCnt, iPoint) = -1 Then
                    ChkGD_Data = FUNCTION_RETURN_FAILURE
                    '����װ(�����𔲂���)
                    Exit Function
                End If
            Next iCnt
        Next iPoint
    End If
    
    ChkGD_Data = FUNCTION_RETURN_SUCCESS
    
End Function

'�T�v      :LifeTime����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_ltz       ,I  ,type_DBDRV_scmzc_fcmkc001c_LT        :LT���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :LifeTime������s��
'����      :
Public Function CrLtjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                         typ_ltz As type_DBDRV_scmzc_fcmkc001c_LT, _
                         bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim Lt      As C_LT                     'LifeTime�\����
    
    bJudg = True
        
    'LT��������ݒ�
    Lt.GuaranteeLt.cMeth = typ_si.HSXLTSPH         '����ʒu_��
    Lt.GuaranteeLt.cCount = typ_si.HSXLTSPT        '����ʒu_�_
    Lt.GuaranteeLt.cPos = typ_si.HSXLTSPI          '����ʒu_��(OSF�̏ꍇ ��)
    Lt.GuaranteeLt.cObj = typ_si.HSXLTHWT          '�ۏؕ��@_��
    Lt.GuaranteeLt.cJudg = typ_si.HSXLTHWS         '�ۏؕ��@_��
    
    Lt.SpecLtMin = typ_si.HSXLTMIN                 '�iSXL�^�C������
    Lt.SpecLtMax = typ_si.HSXLTMAX                 '�iSXL�^�C�����

    Lt.Lt = typ_ltz.CALCMEAS                       '���C�t�^�C���v�Z�l
    
    'LT����
    If CrystalLTJudg(Lt, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrLtjudg = False
        Exit Function
    End If
    
    If Lt.JudgLt <> True Then
        bJudg = False
    End If
    
    CrLtjudg = True

End Function

'�T�v      :LT10����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_ltz       ,I  ,type_DBDRV_scmzc_fcmkc001c_LT        :LT���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :LT10������s��
'����      :2011/07/22 T.Koi(SETsw)
Public Function CrLt10judg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                         typ_ltz As type_DBDRV_scmzc_fcmkc001c_LT, _
                         typ_cr As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                         bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim Lt      As C_LT                     'LifeTime�\����
    Dim CRYREST10CS As String
    
    bJudg = True
    
    'LT��������ݒ�
    Lt.GuaranteeLt.cMeth = typ_si.HSXLTSPH         '����ʒu_��
    Lt.GuaranteeLt.cCount = typ_si.HSXLTSPT        '����ʒu_�_
    Lt.GuaranteeLt.cPos = typ_si.HSXLTSPI          '����ʒu_��(OSF�̏ꍇ ��)
    Lt.GuaranteeLt.cObj = typ_si.HSXLTHWT          '�ۏؕ��@_��
    Lt.GuaranteeLt.cJudg = typ_si.HSXLTHWS         '�ۏؕ��@_��
    
    'LT10���уt���O
    CRYREST10CS = typ_cr.CRYREST10CS               'LT10���уt���O
    If CRYREST10CS = "9" Then                      '�ΏۊO
        CrLt10judg = True
        Exit Function
    End If
    
    Lt.SpecLt10Min = typ_si.HSXLT10MIN               '�iSXL�^�C������

    Lt.Lt10 = typ_ltz.CONVAL                         'LT10���v�Z�l
    
    'LT10����
    If CrystalLT10Judg(Lt, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrLt10judg = False
        Exit Function
    End If
    
    If Lt.JudgLt10 <> True Then
        bJudg = False
    End If
    
    CrLt10judg = True

End Function

'�T�v      :EPD����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_epdz      ,I  ,type_DBDRV_scmzc_fcmkc001c_EPD       :EPD���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :EPD������s��
'����      :
Public Function CrEpdjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_epdz As type_DBDRV_scmzc_fcmkc001c_EPD, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim ep      As C_EPD                    'EPD�\����
    
    bJudg = True
        
    'EPD��������ݒ�
    ep.SpecEpdMax = typ_si.EPDUP            '���������Ǘ��EPD���
    ep.EPD = typ_epdz.MEASURE               'EPD����l
        
    'EPD����
    If CrystalEPDJudg(ep, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrEpdjudg = False
        Exit Function
    End If
        
    If ep.JudgEpd <> True Then
        bJudg = False
    End If
    
    CrEpdjudg = True

End Function
'�T�v      :X������ 2009/08/12 Kameda
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_xz        ,I  ,type_DBDRV_scmzc_fcmkc001c_X         :X�����э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :X��������s��
'����      :2009/10/22 ����͍����p�݂̂ōs��(X,Y���O��Ă��鎞�͐ԕ\��)
Public Function CrXjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_xz As type_DBDRV_scmzc_fcmkc001c_X, _
                          bJudgXY As Boolean, bJudgX As Boolean, bJudgY As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim x      As C_XY                    'X����\����
    
    'bJudg = True       2009/10/22 Kameda
    bJudgXY = True
    bJudgX = True
    bJudgY = True
        
    'X����������ݒ�
    '����
    x.SpecXY_Min = typ_si.HSXCSMIN
    x.SpecXY_Max = typ_si.HSXCSMAX
    x.Spec_XY = typ_xz.XXY
    
    
    '�c
    x.SpecY_Min = typ_si.HSXCTMIN
    x.SpecY_Max = typ_si.HSXCTMAX
    x.Spec_Y = typ_xz.XY
    
    '
    x.SpecX_Min = typ_si.HSXCYMIN
    x.SpecX_Max = typ_si.HSXCYMAX
    x.Spec_X = typ_xz.XX
    
    If CrystalXYJudg(x, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        'bJudg = False      2009/10/22
        bJudgXY = False
        CrXjudg = False
        Exit Function
    End If
        
    If x.JudgResult_XY <> True Then
        'bJudg = False      2009/10/22
        bJudgXY = False
    End If
    If x.JudgResult_Y <> True Then
        'bJudg = False      2009/10/22
        bJudgY = False
    End If
    If x.JudgResult_X <> True Then
        'bJudg = False      2009/10/22
        bJudgX = False
    End If
    
    CrXjudg = True

End Function

'�T�v      :SIRD����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_epdz      ,I  ,type_DBDRV_scmzc_fcmkc001c_SIRD      :SIRD���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :SIRD������s��
'����      :2010/02/04 Kameda
Public Function CrSIRDjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_sird As type_DBDRV_scmzc_fcmkc001c_SIRD, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim SIRD      As C_SIRD                    'SIRD�\����
    
    bJudg = True
        
    'SIRD��������ݒ�
    SIRD.SpecSirdMax = typ_si.HWFSIRDMX     '�d�l�ʓ������
    SIRD.SIRDCNT = typ_sird.SIRDCNT         'SIRD����l
        
    'SIRD����
    If CrystalSIRDJudg(SIRD, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrSIRDjudg = False
        Exit Function
    End If
        
    If SIRD.JudgSird <> True Then
        bJudg = False
    End If
    
    CrSIRDjudg = True

End Function

'Add Start 2011/01/31 SMPK Miyata
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_cz        ,I  ,type_DBDRV_scmzc_fcmkc001c_C         :C���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :C������s��
'����      :
Public Function CrCjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_cz As type_DBDRV_scmzc_fcmkc001c_C, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim C       As C_C                      'C����\����
    
    bJudg = True
        
    'C��������ݒ�
    C.GuaranteeC.cObj = typ_si.HSXCHT       '�i�r�w�b�ۏؕ��@�Q��
    C.GuaranteeC.cJudg = typ_si.HSXCHS      '�i�r�w�b�ۏؕ��@�Q��

    C.HSXCPK = typ_si.HSXCPK                ''�i�r�w�b�p�^�[���敪
    C.HSXCSZ = typ_si.HSXCSZ                ''�i�r�w�b�������
    C.CPTNJSK = typ_cz.CPTNJSK              ''C �p�^�[������
    C.CDISKJSK = typ_cz.CDISKJSK            ''C Disk���a����
    C.CRINGNKJSK = typ_cz.CRINGNKJSK        ''C Ring���a����
    C.CRINGGKJSK = typ_cz.CRINGGKJSK        ''C Ring�O�a����

    If CrystalCJudg(C, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrCjudg = False
        Exit Function
    End If
        
    If C.JudgC <> True Then
        bJudg = False
    End If
    
    CrCjudg = True

End Function

'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_cjz       ,I  ,type_DBDRV_scmzc_fcmkc001c_CJ        :CJ���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :CJ������s��
'����      :
Public Function CrCJjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_cjz As type_DBDRV_scmzc_fcmkc001c_CJ, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim CJ      As C_CJ                     'CJ����\����
    
    bJudg = True
        
    'CJ��������ݒ�
    CJ.GuaranteeCJ.cObj = typ_si.HSXCJHT        '�i�r�w�b�i�ۏؕ��@�Q��
    CJ.GuaranteeCJ.cJudg = typ_si.HSXCJHS       '�i�r�w�b�i�ۏؕ��@�Q��

    CJ.HSXCJPK = typ_si.HSXCJPK                 ''�i�r�w�b�i�p�^�[���敪
    CJ.HSXCJNS = typ_si.HSXCJNS                 ''�i�r�w�b�i�M�����@

    CJ.CJPTNJSK = typ_cjz.CJPTNJSK              ''CJ �p�^�[������
    CJ.CJDISKJSK = typ_cjz.CJDISKJSK            ''CJ Disk���a����
    CJ.CJRINGNKJSK = typ_cjz.CJRINGNKJSK        ''CJ Ring���a����
    CJ.CJRINGGKJSK = typ_cjz.CJRINGGKJSK        ''CJ Ring�O�a����
    CJ.CJBANDNKJSK = typ_cjz.CJBANDNKJSK        ''CJ Band���a����
    CJ.CJBANDGKJSK = typ_cjz.CJBANDGKJSK        ''CJ Band�O�a����
    CJ.CJRINGCALC = typ_cjz.CJRINGCALC          ''CJ Ring���v�Z
    CJ.CJPICALC = typ_cjz.CJPICALC              ''CJ Pi���v�Z
    CJ.CJHANTEI = typ_cjz.CJHANTEI              ''CJ ���茋��
    CJ.CJDMAXPIC5 = typ_cjz.CJDMAXPIC5          ''CJ Disk�̂݃p�^�[�� Pi������l
    CJ.CJRMAXPIC5 = typ_cjz.CJRMAXPIC5          ''CJ Ring�̂݃p�^�[�� Pi������l
    CJ.CJDRMAXPIC5 = typ_cjz.CJDRMAXPIC5        ''CJ DiskRing�p�^�[�� Pi������l
    CJ.CJALLMAXDIC5 = typ_cjz.CJALLMAXDIC5      ''CJ ����Disk���a����l
    CJ.CJALLMINRINC5 = typ_cjz.CJALLMINRINC5    ''CJ ����Ring���a�����l
    CJ.CJALLMAXRIGC5 = typ_cjz.CJALLMAXRIGC5    ''CJ ����Ring�O�a����l

    If CrystalCJJudg(CJ, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrCJjudg = False
        Exit Function
    End If
        
    If CJ.JudgCJ <> True Then
        bJudg = False
    End If
    
    CrCJjudg = True

End Function

'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_cjltz     ,I  ,type_DBDRV_scmzc_fcmkc001c_CJLT      :CJLT���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :CJLT������s��
'����      :
Public Function CrCJLTjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_cjltz As type_DBDRV_scmzc_fcmkc001c_CJLT, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim CJLT    As C_CJLT                   'CJ����\����
    
    bJudg = True
        
    'CJLT��������ݒ�
    CJLT.GuaranteeCJLT.cObj = typ_si.HSXCJLTHT  '�i�r�w�b�i�k�s�ۏؕ��@�Q��
    CJLT.GuaranteeCJLT.cJudg = typ_si.HSXCJLTHS '�i�r�w�b�i�k�s�ۏؕ��@�Q��

    CJLT.HSXCJLTPK = typ_si.HSXCJLTPK           '�i�r�w�b�i�k�s�p�^�[���敪
    CJLT.HSXCJLTNS = typ_si.HSXCJLTNS           '�i�r�w�b�i�k�s�M�����@
    CJLT.CJLTPTNJSK = typ_cjltz.CJLTPTNJSK      ''CJ(LT) �p�^�[������
    CJLT.CJLTDISKJSK = typ_cjltz.CJLTDISKJSK    ''CJ(LT) Disk���a����
    CJLT.CJLTRINGNKJSK = typ_cjltz.CJLTRINGNKJSK ''CJ(LT) Ring���a����
    CJLT.CJLTRINGGKJSK = typ_cjltz.CJLTRINGGKJSK ''CJ(LT) Ring�O�a����
    CJLT.CJLTBANDNKJSK = typ_cjltz.CJLTBANDNKJSK ''CJ(LT) Band���a����
    CJLT.CJLTBANDGKJSK = typ_cjltz.CJLTBANDGKJSK ''CJ(LT) Band�O�a����
    CJLT.CJLTRINGCALC = typ_cjltz.CJLTRINGCALC  ''CJ(LT) Ring���v�Z
    CJLT.CJLTPICALC = typ_cjltz.CJLTPICALC      ''CJ(LT) Pi���v�Z
    CJLT.CJLTBANDCALC = typ_cjltz.CJLTBANDCALC  ''CJ(LT) Band���v�Z
    CJLT.HSXCJLTBND = typ_cjltz.HSXCJLTBND      ''CJ(LT) Band������l

    If CrystalCJLTJudg(CJLT, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrCJLTjudg = False
        Exit Function
    End If
        
    If CJLT.JudgCJLT <> True Then
        bJudg = False
    End If
    
    CrCJLTjudg = True

End Function

'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :typ_cj2z      ,I  ,type_DBDRV_scmzc_fcmkc001c_CJ2       :CJ2���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(True:����OK, False:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :CJLT������s��
'����      :
Public Function CrCJ2judg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_cj2z As type_DBDRV_scmzc_fcmkc001c_CJ2, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim CJ2     As C_CJ2                    'CJ����\����
    
    bJudg = True
        
    'CJ2��������ݒ�
    CJ2.GuaranteeCJ2.cObj = typ_si.HSXCJ2HT     '�i�r�w�b�i�Q�ۏؕ��@�Q��
    CJ2.GuaranteeCJ2.cJudg = typ_si.HSXCJ2HS    '�i�r�w�b�i�Q�ۏؕ��@�Q��

    CJ2.HSXCJ2PK = typ_si.HSXCJ2PK              '�i�r�w�b�i�k�s�p�^�[���敪
    CJ2.HSXCJ2NS = typ_si.HSXCJ2NS              '�i�r�w�b�i�k�s�M�����@

    CJ2.CJ2PTNJSK = typ_cj2z.CJ2PTNJSK          ''CJ2 �p�^�[������
    CJ2.CJ2DISKJSK = typ_cj2z.CJ2DISKJSK        ''CJ2 Disk���a����
    CJ2.CJ2RINGNKJSK = typ_cj2z.CJ2RINGNKJSK    ''CJ2 Ring���a����
    CJ2.CJ2RINGGKJSK = typ_cj2z.CJ2RINGGKJSK    ''CJ2 Ring�O�a����
    CJ2.CJ2PICALC = typ_cj2z.CJ2PICALC          ''CJ2 Pi���v�Z
    CJ2.CJ2HANTEI = typ_cj2z.CJ2HANTEI          ''CJ2 ���茋��
    CJ2.CJ2DMAXPIC5 = typ_cj2z.CJ2DMAXPIC5      ''CJ2 Disk�̂݃p�^�[�� Pi������l
    CJ2.CJ2RMAXPIC5 = typ_cj2z.CJ2RMAXPIC5      ''CJ2 Ring�̂݃p�^�[�� Pi������l
    CJ2.CJ2RMINRINC5 = typ_cj2z.CJ2RMINRINC5    ''CJ2 Ring�̂݃p�^�[�� Ring���a�����l
    CJ2.CJ2RMAXRIGC5 = typ_cj2z.CJ2RMAXRIGC5    ''CJ2 Ring�̂݃p�^�[�� Ring�O�a����l
    CJ2.CJ2DRMAXPIC5 = typ_cj2z.CJ2DRMAXPIC5    ''CJ2 DiskRing�p�^�[�� Pi������l
    CJ2.CJ2DRMINRINC5 = typ_cj2z.CJ2DRMINRINC5  ''CJ2 DiskRing�p�^�[�� Ring���a�����l
    CJ2.CJ2DRMAXRIGC5 = typ_cj2z.CJ2DRMAXRIGC5  ''CJ2 DiskRing�p�^�[�� Ring�O�a����l

    If CrystalCJ2Judg(CJ2, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrCJ2judg = False
        Exit Function
    End If
        
    If CJ2.JudgCJ2 <> True Then
        bJudg = False
    End If
    
    CrCJ2judg = True

End Function

'Add End   2011/01/31 SMPK Miyata

Private Function NtoS(strWk As String) As String
    If Mid(strWk, 1, 1) = Chr(0) Then
        NtoS = " "
        Exit Function
    End If
    NtoS = strWk
End Function

Private Sub BMDDataSet(BmdNo As Integer, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                       '�����w��
    Dim typ_bmdz        As type_DBDRV_scmzc_fcmkc001c_BMD
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    
    '�����w���ݒ�
    IND = IIf(UpDo = BlkTop, "123", "123")
    
    With typ_b
        Select Case BmdNo
        Case 1
            JudgSpecCode = JudgSC_B(UpDo).B1
            SCC = "B1"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDB1CS) <> 0)
            typ_bmdz = .typ_zi.BMD1Z(UpDo)
        Case 2
            JudgSpecCode = JudgSC_B(UpDo).B2
            SCC = "B2"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDB2CS) <> 0)
            typ_bmdz = .typ_zi.BMD2Z(UpDo)
        Case 3
            JudgSpecCode = JudgSC_B(UpDo).B3
            SCC = "B3"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDB3CS) <> 0)
            typ_bmdz = .typ_zi.BMD3Z(UpDo)
        End Select
        
        If JudgSpecCode Then
            '��ʕ\�����e�ݒ�
            .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' �������J�n�ʒu
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
            .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                             ' ���P
            .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                             ' ���Q
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' ���R
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
            .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' �T���v���m��
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
            bJudg = False
            If shiji Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                         ' ���R
                If left(typ_bmdz.CRYNUM, 9) = left(.BLOCKID, 9) Then
                    .typ_rslt(UpDo, DispLineCount).pos = typ_bmdz.POSITION              ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_bmdz.SMPLNO             ' �T���v���m��
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                    ' ���R
                    If typ_bmdz.SMPLUMU = "0" Then
                        'BMD1���莸�s
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                                ' ���Q
                        'BMD1����
                        If CrBmdJudg(.typ_si(UpDo), typ_bmdz, bJudg, BmdNo) Then
                            vTemp = CStr(typ_bmdz.MEASAVE)                                              ' ���P
                            .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.0")        ' ���P
                            vTemp = CStr(typ_bmdz.MEASMAX)                                              ' ���Q
                            .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")        ' ���Q
                            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' ���R
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                            vTemp = CStr(typ_bmdz.BMDMNBUNP)
                            .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.0")        ' ���4
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                        End If
                    End If
                End If
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' ���茋��
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' ���茋��
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case BmdNo
                Case 1
                    gsTbcmy028ErrCode = "00107"
                Case 2
                    gsTbcmy028ErrCode = "00108"
                Case 3
                    gsTbcmy028ErrCode = "00109"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = typ_bmdz.POSITION                      ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                             ' ���P
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                             ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                            ' ���R
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���4
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_bmdz.SMPLNO                     ' �T���v���m��
'====================== Debug Debug =====================================
                .typ_rslt(UpDo, DispLineCount).OKNG = "N�Q"                                  ' ���茋��
'====================== Debug Debug =====================================
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                If typ_bmdz.SMPLUMU = "0" Then
                    'BMD1���莸�s
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                                ' ���Q
                    'BMD1����
                    If CrBmdJudg(.typ_si(UpDo), typ_bmdz, bJudg, BmdNo) Then
                        vTemp = CStr(typ_bmdz.MEASAVE)                                              ' ���P
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.0")        ' ���P
                        vTemp = CStr(typ_bmdz.MEASMAX)                                              ' ���Q
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")        ' ���Q
                        .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' ���Q
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                        vTemp = CStr(typ_bmdz.BMDMNBUNP)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.0")        ' ���4
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                    End If
                End If
                DispLineCount = DispLineCount + 1
            End If
        End If
    End With
End Sub

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech Start
'' �iSXOSF1(ArAN)�p�^���敪�������ɒǉ�
Private Sub OSFDataSet(OsfNo As Integer, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String, sAranPtn As String)
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech End
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                  '�����w��
    Dim typ_Osfz        As type_DBDRV_scmzc_fcmkc001c_OSF
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim w_1             As String
    Dim w_2             As String
    Dim w_3             As String
    
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
    Dim strXTALC1       As String
    Dim strJDGEIDC      As String
    Dim strSYNFLG       As String
    Dim strYMKFLG       As String
    Dim strOSFRD1       As String
    Dim strOSFRD2       As String
    Dim lngOSFWID1      As Long
    Dim lngOSFWID2      As Long
    Dim lngSMPPOS       As Long
    Dim strRMAXC5       As String
    Dim strDMAXC5       As String
    Dim strDRRMAXC5     As String
    Dim strDRDMAXC5     As String
    Dim lngRMAXC5       As Long
    Dim lngDMAXC5       As Long
    Dim lngDRRMAXC5     As Long
    Dim lngDRDMAXC5     As Long
    Dim ErrFlg          As Boolean
    Dim SYNErrFlg       As Boolean
    Dim DBErrFlg        As Boolean
    Dim YFlg            As Boolean
       
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END ---

    'OSF3����p�׸ޏ�����
    gsCOSF3Flg = ""

    '�����w���ݒ�
    IND = IIf(UpDo = BlkTop, "123", "123")
        
    With typ_b
        Select Case OsfNo
        Case 1
            JudgSpecCode = JudgSC_B(UpDo).L1
            SCC = "L1"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDL1CS) <> 0)
            typ_Osfz = .typ_zi.OSF1Z(UpDo)
        Case 2
            JudgSpecCode = JudgSC_B(UpDo).L2
            SCC = "L2"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDL2CS) <> 0)
            typ_Osfz = .typ_zi.OSF2Z(UpDo)
        Case 3
            JudgSpecCode = JudgSC_B(UpDo).L3
            SCC = "L3"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDL3CS) <> 0)
            typ_Osfz = .typ_zi.OSF3Z(UpDo)
        Case 4
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
            JudgSpecCode = JudgSC_B(UpDo).COSF3
            SCC = "COSF3"
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---
            'SCC = "L4"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDL4CS) <> 0)
            typ_Osfz = .typ_zi.OSF4Z(UpDo)
        End Select
                   
        '�ۏ��׸�="H"�̏ꍇ
        If JudgSpecCode Then
        
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
            '�M�����敪:OSF4�̏ꍇ
            If OsfNo = 4 Then
                '��ۯ�ID
                strXTALC1 = Trim(typ_b.BLOCKID)
                '�����ԍ�
                strXTALC1 = left(strXTALC1, 9) & "000"

                'OSF���ѓ��͔��菈��
                '���������&���ђl�ޔ�
                
                If Trim(typ_Osfz.OSFRD1) = "R" Or Trim(typ_Osfz.OSFRD1) = "D" Then
                    strOSFRD1 = Trim(typ_Osfz.OSFRD1)
                Else
                    strOSFRD1 = "-"
                End If
                
                If Trim(typ_Osfz.OSFRD2) = "D" Then
                    strOSFRD2 = Trim(typ_Osfz.OSFRD2)
                Else
                    strOSFRD2 = "-"
                End If
                
                If IsNull(typ_Osfz.OSFWID1) = True Then
                   lngOSFWID1 = -1
                ElseIf IsNumeric(typ_Osfz.OSFWID1) = False Then
                   lngOSFWID1 = -1
                Else
                   lngOSFWID1 = Trim(typ_Osfz.OSFWID1)
                End If
                
                If IsNull(typ_Osfz.OSFWID2) = True Then
                   lngOSFWID2 = -1
                ElseIf IsNumeric(typ_Osfz.OSFWID2) = False Then
                   lngOSFWID2 = -1
                Else
                   lngOSFWID2 = Trim(typ_Osfz.OSFWID2)
                End If
                
                '-1�ȊO�̐��l�l��
                If lngOSFWID1 < 0 Then
                   lngOSFWID1 = -1
                End If
                If lngOSFWID2 < 0 Then
                   lngOSFWID2 = -1
                End If

               '����وʒu
               lngSMPPOS = Trim(typ_Osfz.POSITION)

               '�����׸ޏ�����
               ErrFlg = True
               YFlg = False
                               
               '����݋敪�A���ђl��NULL�̏ꍇ
               If strOSFRD1 = "-" And strOSFRD2 <> "-" Then
                   ErrFlg = False
               ElseIf strOSFRD1 <> "-" And lngOSFWID1 = -1 Then
                    ErrFlg = False
               ElseIf strOSFRD2 <> "-" And lngOSFWID2 = -1 Then
                    ErrFlg = False
               ElseIf strOSFRD2 = "-" And lngOSFWID2 > 0 Then
                   ErrFlg = False
               End If
               
                '�����ԍ��𷰂Ƃ���XSDC1���C�|OSF3����ID���l������
               If GetCOSF3ID(strJDGEIDC, strXTALC1) <> FUNCTION_RETURN_SUCCESS Then
                    ErrFlg = False
               Else
                   If Trim(strJDGEIDC) = "" Then
                        gsCOSF3Flg = "1"
                        ErrFlg = False
                   '����ID=�u9�v�̏ꍇ�͔���Ȃ�(����OK)�@07/08/01 M.Kaga
                   ElseIf Trim(strJDGEIDC) = "9" Then
                        YFlg = True
                        bJudg = True
                   Else
                       '�l������C-OSF3����ID��XODC5_OSF30��菳�F�׸ނ̊l��
                       If GetSYNFLAGC5(strSYNFLG, strYMKFLG, strJDGEIDC) <> FUNCTION_RETURN_SUCCESS Then
                          ErrFlg = False
                       Else
                          '���F�׸�:0�@�����F�̏ꍇ
                          If Trim(strSYNFLG) = "0" Or Trim(strSYNFLG) = "" Or IsNull(strSYNFLG) Then
                             gsCOSF3Flg = "2"
                             ErrFlg = False
                          End If
                       End If
                   End If
               End If
               
               If ErrFlg = False Then
                   '��ʕ\�����e�ݒ�
                   .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' �������J�n�ʒu
                   .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' ���e
                   .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                   .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                   .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                      ' ���R
                   .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' �T���v���m��
                   .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' �i��(12��)
                   If left(typ_Osfz.CRYNUM, 9) = left(.BLOCKID, 9) Then
                       .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION              ' �������J�n�ʒu
                       .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO             ' �T���v���m��
                   End If
                   bJudg = False
               Else
                   If YFlg = False Then
                       '����݋敪�ɂ�菈������
                       'R�݂̂̏ꍇ
                       If strOSFRD1 = "R" And strOSFRD2 = "-" Then
                           'R�̂ݏ���l�̊l�����s��
                           If GetCOSF3PTN(strJDGEIDC, lngSMPPOS, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                               ErrFlg = False
                           End If
                           'ں��ޖ��FVB�G���[(��ōl����)
                           If Trim(strRMAXC5) = "" Then
                               ErrFlg = False
                           Else
                               lngRMAXC5 = Trim(strRMAXC5)
                               '���ђl�̔���
    
                               If lngOSFWID1 <= lngRMAXC5 Then
                                   '����OK
                                   bJudg = True
                               ElseIf lngOSFWID1 > lngRMAXC5 Then
                                   '����NG
                                   bJudg = False
                               End If
                           End If
                       'D�݂̂̏ꍇ
                       ElseIf strOSFRD1 = "D" Then
                           'D�̂ݏ���l�̊l�����s��
                           If GetCOSF3PTN(strJDGEIDC, lngSMPPOS, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                               ErrFlg = False
                           End If
    
                           'ں��ޖ�����Ͻ��̎��ђl��NULL�FVB�G���[(��ōl����)
                           If Trim(strDMAXC5) = "" Then
                               ErrFlg = False
                           Else
                               lngDMAXC5 = Trim(strDMAXC5)
                               '���ђl�̔���
                               If lngOSFWID1 <= lngDMAXC5 Then
                                   '����OK
                                   bJudg = True
                               ElseIf lngOSFWID1 > lngDMAXC5 Then
                                   '����NG
                                   bJudg = False
                               End If
                           End If
                       'R&D�̏ꍇ
                       ElseIf strOSFRD1 = "R" And strOSFRD2 = "D" Then
                           'D��������l����R��������l�̊l�����s��
                           If GetCOSF3PTN(strJDGEIDC, lngSMPPOS, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                               ErrFlg = False
                           End If
    
                           'ں��ޖ�����Ͻ��̎��ђl��NULL�FVB�G���[(��ōl����)
                           If Trim(strDRRMAXC5) = "" Or Trim(strDRDMAXC5) = "" Then
                               ErrFlg = False
                           Else
                               lngDRRMAXC5 = Trim(strDRRMAXC5)
                               lngDRDMAXC5 = Trim(strDRDMAXC5)
                               '���ђl�̔���
                               If lngOSFWID1 <= lngDRRMAXC5 And lngOSFWID2 <= lngDRDMAXC5 Then
                                   '����OK
                                   bJudg = True
                               ElseIf lngOSFWID1 > lngDRRMAXC5 Or lngOSFWID2 > lngDRDMAXC5 Then
                                   '����NG
                                   bJudg = False
                               End If
                           End If
                       Else
                           '���ђl�����A��������ݖ����̏ꍇ����OK
                           bJudg = True
                       End If
                   End If
                   If ErrFlg = False Then
                       '��ʕ\�����e�ݒ�
                       .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' �������J�n�ʒu
                       .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' ���e
                       .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                       .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                       .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                        ' ���R
                       .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' �T���v���m��
                       .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' �i��(12��)
                       If left(typ_Osfz.CRYNUM, 9) = left(.BLOCKID, 9) Then
                           .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION              ' �������J�n�ʒu
                           .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO             ' �T���v���m��
                       End If
                       bJudg = False

                   Else
                       '��ʕ\�����e�ݒ�
                       .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' �������J�n�ʒu
                       .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                       .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                             ' ���P
                       .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                             ' ���Q
                       .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' ���R
                       .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                       .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' �T���v���m��
                       .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                       If shiji Then
                           '��ʕ\�����e�ݒ�
                           .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                           .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                           .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                         ' ���R
                           If left(typ_Osfz.CRYNUM, 9) = left(.BLOCKID, 9) Then
                               .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION              ' �������J�n�ʒu
                               .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO             ' �T���v���m��
                               .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                    ' ���R
                               If typ_Osfz.SMPLUMU = "0" Then
                                   .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                ' ���R
                                   '��ʕ\�����e�ݒ�
                                    .typ_rslt(UpDo, DispLineCount).INFO1 = strOSFRD1               ' ���P
                                    .typ_rslt(UpDo, DispLineCount).INFO2 = lngOSFWID1              ' ���Q
                                    .typ_rslt(UpDo, DispLineCount).INFO3 = strOSFRD2               ' ���R
                                    .typ_rslt(UpDo, DispLineCount).INFO4 = lngOSFWID2              ' ���S
                               End If
                           End If
                       End If
                   End If
               End If
               If bJudg = True Then
                   .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' ���茋��
               Else
                   .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' ���茋��
                   TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                    Select Case OsfNo
                    Case 1
                        gsTbcmy028ErrCode = "00103"
                    Case 2
                        gsTbcmy028ErrCode = "00104"
                    Case 3
                        gsTbcmy028ErrCode = "00105"
                    Case 4
                        gsTbcmy028ErrCode = "00106"
                    End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
               End If
               DispLineCount = DispLineCount + 1
           Else
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END ---

                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                             ' ���P
                .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                             ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' ���R
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' �T���v���m��
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                bJudg = False
                If shiji Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                         ' ���R
                    If left(typ_Osfz.CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION              ' �������J�n�ʒu
                        .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO             ' �T���v���m��
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                    ' ���R
                        If typ_Osfz.SMPLUMU = "0" Then
                            'OSF���莸�s
                            .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                                        ' ���R
                            'OSF����擾
                            If CrOsfJudg(.typ_si(UpDo), typ_Osfz, bJudg, OsfNo) Then
                                '��ʕ\�����e�ݒ�
                                vTemp = CStr(typ_Osfz.CALCAVE)                                                      ' ���P
                                .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")               ' ���P
                                vTemp = CStr(typ_Osfz.CALCMAX)                                                      ' ���Q
                                .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")                ' ���Q
                                vTemp = CStr(typ_Osfz.MEAS6)                                                        ' ���R
                                .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")               ' ���R
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
                                If sAranPtn = "1" Or sAranPtn = "2" Then
                                    vTemp = CStr(typ_Osfz.CALCMH)                                                            ' ���R
                                    .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.0")                   ' ���R
                                End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

    ' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                                w_1 = IIf(typ_Osfz.OSFRD1 = Null Or typ_Osfz.OSFRD1 = " ", "�|", typ_Osfz.OSFRD1)
                                w_2 = IIf(typ_Osfz.OSFRD2 = Null Or typ_Osfz.OSFRD2 = " ", "�|", typ_Osfz.OSFRD2)
                                w_3 = IIf(typ_Osfz.OSFRD3 = Null Or typ_Osfz.OSFRD3 = " ", "�|", typ_Osfz.OSFRD3)
                                .typ_rslt(UpDo, DispLineCount).INFO4 = w_1 & w_2 & w_3                              ' ���4
                                
    ' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                            End If
                        End If
                    End If
                End If
                If bJudg = True Then
                    .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' ���茋��
                Else
                    .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' ���茋��
                    TotalJudg = False
                End If
                DispLineCount = DispLineCount + 1
            End If
            
        '�ۏ��׸�=S����NULL�̏ꍇ
        Else
            If shiji Then
                If OsfNo = 4 Then
            
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION                      ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                             ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                             ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                            ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO                     ' �T���v���m��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N�Q"                                  ' ���茋��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                    If typ_Osfz.SMPLUMU = "0" Then
                        '��ʕ\�����e�ݒ�
                        If IsNull(typ_Osfz.OSFRD1) = True Then
                            .typ_rslt(UpDo, DispLineCount).INFO1 = ""
                        Else
                            .typ_rslt(UpDo, DispLineCount).INFO1 = typ_Osfz.OSFRD1
                        End If
                        If IsNull(typ_Osfz.OSFWID1) = True Then
                            .typ_rslt(UpDo, DispLineCount).INFO2 = ""
                        Else
                            .typ_rslt(UpDo, DispLineCount).INFO2 = typ_Osfz.OSFWID1
                        End If
                        If IsNull(typ_Osfz.OSFRD2) = True Then
                            .typ_rslt(UpDo, DispLineCount).INFO3 = ""
                        Else
                            .typ_rslt(UpDo, DispLineCount).INFO3 = typ_Osfz.OSFRD2
                        End If
                        If IsNull(typ_Osfz.OSFWID2) = True Then
                            .typ_rslt(UpDo, DispLineCount).INFO4 = ""
                        Else
                            .typ_rslt(UpDo, DispLineCount).INFO4 = typ_Osfz.OSFWID2
                        End If
                       
                    End If
                Else
            
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION                      ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                             ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                             ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                            ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO                     ' �T���v���m��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N�Q"                                  ' ���茋��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                    If typ_Osfz.SMPLUMU = "0" Then
                        'OSF���莸�s
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                                            ' ���Q
                        'OSF����
                        If CrOsfJudg(.typ_si(UpDo), typ_Osfz, bJudg, OsfNo) Then
                            vTemp = CStr(typ_Osfz.CALCAVE)                                                          ' ���P
                            .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")                   ' ���P
                            vTemp = CStr(typ_Osfz.CALCMAX)                                                          ' ���Q
                            .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")                    ' ���Q
                            vTemp = CStr(typ_Osfz.MEAS6)                                                            ' ���R
                            .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")                   ' ���R
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
                            If sAranPtn = "1" Or sAranPtn = "2" Then
                                vTemp = CStr(typ_Osfz.CALCMH)                                                            ' ���R
                                .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.0")                   ' ���R
                            End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

    ' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                            w_1 = IIf(typ_Osfz.OSFRD1 = Null Or typ_Osfz.OSFRD1 = " ", "�|", typ_Osfz.OSFRD1)
                            w_2 = IIf(typ_Osfz.OSFRD2 = Null Or typ_Osfz.OSFRD2 = " ", "�|", typ_Osfz.OSFRD2)
                            w_3 = IIf(typ_Osfz.OSFRD3 = Null Or typ_Osfz.OSFRD3 = " ", "�|", typ_Osfz.OSFRD3)
                            .typ_rslt(UpDo, DispLineCount).INFO4 = w_1 & w_2 & w_3                                  ' ���4
    ' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
                        End If
                    End If
                End If
                DispLineCount = DispLineCount + 1
            End If
        End If
    End With
End Sub

Private Function AllHinGdjudg(Gd_si() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_gdz As type_DBDRV_scmzc_fcmkc001c_GD, _
                          cnt As Integer) As Boolean
Dim RET     As Boolean
Dim Judg    As Boolean
Dim i       As Integer

    AllHinGdjudg = False
    
    For i = 1 To cnt
        RET = CrGdjudg(Gd_si(i), typ_gdz, Judg)
        If Judg = False Then
            AllHinGdjudg = True
        End If
    Next

End Function

Public Function DBData2DispData(data As Variant, Optional Formatstr As String) As Variant
    If data = -1 Then
        DBData2DispData = ""
    Else
        If Formatstr = "" Then
            DBData2DispData = data
        Else
            DBData2DispData = Format(data, Formatstr)
        End If
    End If
End Function

'------------------------------------------------
' �d�lNull�`�F�b�N(����)
'------------------------------------------------

'�T�v      :������������̊e�������ڂ̕ۏؕ��@��'H'�܂���'S'�̏ꍇ�A�d�l�l��Null(-1)���ǂ����𔻒f����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tSiyou        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�i�ԁA�d�l�A���������擾�p
'          :sErrMsg       ,IO ,String                               :�װү����
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���� = FUNCTION_RETURN_SUCCESS : OK
'                                                                           FUNCTION_RETURN_FAILURE : NG
'����      :
'����      :2003/12/13 �V�K�쐬�@�V�X�e���u���C��

Private Function funCryChkNull(tSiyou As type_DBDRV_scmzc_fcmkc001c_Siyou, sErrMsg As String) As FUNCTION_RETURN
    Dim dShiyo()    As Double
    Dim sHosyo      As String
    Dim cnt         As Integer
    
    '������
    funCryChkNull = FUNCTION_RETURN_SUCCESS
    
    '--------------- RS(���R) ---------------
    ReDim dShiyo(5)
    dShiyo(1) = tSiyou.HSXRMIN          ' �i�r�w���R����
    dShiyo(2) = tSiyou.HSXRMAX          ' �i�r�w���R���
    dShiyo(3) = tSiyou.HSXRAMIN         ' �i�r�w���R���ω���
    dShiyo(4) = tSiyou.HSXRAMAX         ' �i�r�w���R���Ϗ��
    dShiyo(5) = tSiyou.HSXRMBNP         ' �i�r�w���R�ʓ����z
    If fncJissekiHantei_nl(tSiyou.HSXRHWYS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(RS)"
        funCryChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00100"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
    '--------------- Oi(�_�f�Z�x) ---------------
    ReDim dShiyo(5)
    dShiyo(1) = tSiyou.HSXONMIN         ' �i�r�w�_�f�Z�x����
    dShiyo(2) = tSiyou.HSXONMAX         ' �i�r�w�_�f�Z�x���
    dShiyo(3) = tSiyou.HSXONAMN         ' �i�r�w�_�f�Z�x���ω���
    dShiyo(4) = tSiyou.HSXONAMX         ' �i�r�w�_�f�Z�x���Ϗ��
    dShiyo(5) = tSiyou.HSXONMBP         ' �i�r�w�_�f�Z�x�ʓ����z
    If fncJissekiHantei_nl(tSiyou.HSXONHWS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(Oi)"
        funCryChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00101"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
        
    '--------------- CS(�Y�f�Z�x) ---------------
    ReDim dShiyo(2)
    dShiyo(1) = tSiyou.HSXCNMIN         ' �i�r�w�Y�f�Z�x����
    dShiyo(2) = tSiyou.HSXCNMAX         ' �i�r�w�Y�f�Z�x���
    If fncJissekiHantei_nl(tSiyou.HSXCNHWS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(CS)"
        funCryChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00111"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
    '--------------- LT(ײ����) ---------------
    ReDim dShiyo(1)
'   ReDim dShiyo(2)

    dShiyo(1) = tSiyou.HSXLTMIN         ' �i�r�w�k�^�C������
'   dShiyo(2) = tSiyou.HSXLTMAX         ' �i�r�w�k�^�C�����
    If fncJissekiHantei_nl(tSiyou.HSXLTHWS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(LT)"
        funCryChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00110"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
'''Add Start 2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
'    dShiyo(1) = tSiyou.HSXLT10MIN         ' �i�r�w�kLT10����
'    If fncJissekiHantei_nl(tSiyou.HSXLTHWS, dShiyo) = False Then
'        sErrMsg = sErrMsg & "(LT10)"
'        funCryChkNull = FUNCTION_RETURN_FAILURE
'        gsTbcmy028ErrCode = "00110"
'        Exit Function
'    End If
'
'''Add End   2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
    
    '--------------- EPD ---------------
    If tSiyou.EPDUP = -1 Then           ' EPD���
        sErrMsg = sErrMsg & "(EPD)"
        funCryChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00102"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If

End Function
'�T�v      :�u���b�N�ΐ́A���ް�߈ʒu����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :TopPos        ,I   ,�g�b�v����وʒu
'          :BotPos        ,I   ,�{�g������وʒu
'          :TopMeas       ,I   ,�g�b�v���S�����l
'          :BotMeas       ,I   ,�{�g�����S�����l
'          :JDCryNum      ,I   ,�����ԍ�
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :�u���b�N�ΐ͔͈͂���ђ��ް�߈ʒu���܂ނ��`�F�b�N���s���i���莞����j
'����      :2005/1/11
Public Function HenDopeJudg(TOPPOS As Integer, BOTPOS As Integer, TopMeas As Double, BotMeas As Double, _
                          JDCryNum As String, JDHinb As tFullHinban) As Boolean
    Dim COEF        As Double
    Dim wgtCharge   As Long                 '�ΐ͌v�Z�p�p�����[�^
    Dim wgtTop      As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim wgtTopCut   As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim DM          As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim cf As C_COEF
    Dim sMcno2 As String
    Dim sMcno1 As String
    Dim sMcno  As String
    Dim cc          As type_Coefficient
    Dim ErrInfo     As ERROR_INFOMATION     '�G���[���\����
    Dim i As Integer
    
    HenDopeJudg = True
    
    '�ΐ͌W���v�Z �}���`����Ή� �Q�Ɗ֐��ύX 2008/04/23 SETsw Nakada
    If GetCoeffParams_new(JDCryNum, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
        Debug.Print "�ΐ͌v�Z�p�p�����[�^�̎擾�Ɏ��s����"
    End If
    
    cc.DUNMENSEKI = AreaOfCircle(DM)
    cc.TOPSMPLPOS = TOPPOS
    cc.BOTSMPLPOS = BOTPOS
    cc.CHARGEWEIGHT = wgtCharge
    cc.TOPWEIGHT = wgtTop + wgtTopCut
    cc.TOPRES = TopMeas
    cc.BOTRES = BotMeas
    COEF = CoefficientCalculation(cc)
    '�u���b�N�ΐ͔��菈�� -------
    '�i�Ԃ�萻������i���o�[�����߂�
    sMcno = Trim(GetMcno(JDHinb))
    i = UBound(SuiteiData)
    SuiteiData(i).SuiSpec.PRODCOND = sMcno
    sMcno1 = Mid(sMcno, 2, 1)
    sMcno2 = Mid(sMcno, 1, 1)
    cf.JudgCOEF = True
    Select Case sMcno1
        Case "H", "I", "J", "K"
            cf.NP = "n"
        Case "A", "B", "C"
            Select Case sMcno2
                Case "A", "B"
                    cf.NP = "p+"
                Case "1", "2", "3", "4", "5", "6", "7", "C", "E"
                    cf.NP = "p-"
                Case Else
                    cf.JudgCOEF = False
            End Select
        Case Else
            cf.JudgCOEF = False
    End Select
    If cf.JudgCOEF Then
        cf.COEF = COEF
        If CrystalCOEFJudg(cf, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
            cf.JudgCOEF = False
        End If
    End If
    With SuiteiData(i)
        '�G���[�\���p�Ƀt���O���Z�b�g����,����s��Ԃ�
        If cf.JudgCOEF Then
            .COEFflg = True
        Else
            .COEFflg = False
            HenDopeJudg = False
        End If
        .Hinsyu = cf.NP
        .COEF = cf.COEF
        '�ǉ��ް�߈ʒu�̃`�F�b�N
        .DOPEflg = True
        If typ_b.typ_si(1).ADDDPPOS <> 0 Then
            If TOPPOS <= typ_b.typ_si(1).ADDDPPOS And BOTPOS >= typ_b.typ_si(1).ADDDPPOS Then
               .DOPEflg = False
                HenDopeJudg = False
            End If
        End If
    End With
        
End Function

'�T�v      :������������擾����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:jdHinban    ,I  ,String           ,
'      �@�@:�߂�l       ,O  ,Long           �@,������
'�쐬      :2005/1/11
Public Function GetMcno(jd_Hinban As tFullHinban) As String

    Dim sSql As String
    Dim rs As OraDynaset
    
    
    sSql = "SELECT MCNO FROM TBCME036"
    sSql = sSql & " WHERE HINBAN = '" & jd_Hinban.hinban & "' "
    sSql = sSql & " and MNOREVNO = '" & jd_Hinban.mnorevno & "'"
    sSql = sSql & " and FACTORY  = '" & jd_Hinban.factory & "'"
    sSql = sSql & " and OPECOND  = '" & jd_Hinban.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        GetMcno = ""
    Else
        GetMcno = rs("MCNO")
    End If
    
End Function

'
'���C�t�^�C�����Čv�Z����
'�T�v      :LT�v�Z�֐����Ăяo���l��Ԃ�
'���Ұ��@�@:�ϐ���   ,IO ,�^                                ,����
'      �@�@:Siyou    ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou  ,����ʒu���擾
'      �@�@:jisseki  ,IO ,type_DBDRV_scmzc_fcmkc001c_LT     ,����l1�`10���擾���A�v�Z���ʂ�Ԃ�
'      �@�@:�߂�l   �Ȃ�
'�쐬      :2005/12/02 SETsw�@����@�L�s
'          :
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Sub_LTReCalc(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                jisseki As type_DBDRV_scmzc_fcmkc001c_LT)

    Dim MEAS(9) As Integer      '����l�i�[�z��
    Dim iOldFlg As Integer      '���f�[�^����t���O
    Dim iRet As Integer         '�߂�l
    Dim iResult As Integer      '�v�Z����
    
    MEAS(0) = jisseki.MEAS1
    MEAS(1) = jisseki.MEAS2
    MEAS(2) = jisseki.MEAS3
    MEAS(3) = jisseki.MEAS4
    MEAS(4) = jisseki.MEAS5
    MEAS(5) = jisseki.MEAS6
    MEAS(6) = jisseki.MEAS7
    MEAS(7) = jisseki.MEAS8
    MEAS(8) = jisseki.MEAS9
    MEAS(9) = jisseki.MEAS10
    
    '���f�[�^����t���O�m�F
    If jisseki.LTSPIFLG <> "" Then
        iOldFlg = 0
    Else
        iOldFlg = 1
    End If
    
    '���C�t�^�C���v�Z�l
    iRet = KNS_CalculateMeasResult_LT(iResult, MEAS, siyou.HSXLTSPI, iOldFlg)
    If iRet <> FUNC_RET_LT_SUCCESS Then
        jisseki.CALCMEAS = -1
    Else
        jisseki.CALCMEAS = iResult
    End If
End Sub

'------------------------------------------------
' �����i�Ԕ���Ή�
'------------------------------------------------

'�T�v      :���ђl�̑���������s���B
'���Ұ�    :�ϐ���          ,IO ,�^             :����
'          :sKeyID          ,I  ,String         :��ۯ�ID�A���́A�����ԍ�
'          :tNew_Hinban     ,I  ,tFullHinban    :�U�֌��i��
'          :bTotalJudg      ,O  ,Boolean        :�g�[�^������
'          :iErr_Code       ,O  ,Integer        :�װ����(�߂�l�Ɠ���)
'          :sErr_Msg        ,O  ,String         :�װү���޺���
'          :typ_B           ,O  ,typ_AllTypesB  :�S���\����(�\����)
'          :iSmpGetFlg      ,I  ,Integer        :����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
'          :iSamplID1       ,I  ,Long           :TOP�����ID(�ȗ���)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iSamplID2       ,I  ,Long           :BOT�����ID(�ȗ���)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iKcnt           ,I  ,Integer        :�H���A��(�ȗ���)
'          :�߂�l          ,O  ,Integer        :�擾�̐���(0:����I��, -1:�ُ�I��)
'����      :
'����      :�����i�Ԕ���Ή� 20060501 SMP���� funCrySogoHantei������
''memo:            tOld_Hinban = TOP tNew_Hinban=Tail
Public Function funCrySogoHantei_CC600Multi(sKeyID As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_b As typ_AllTypesB, _
                iSmpGetFlg As Integer, Optional iSamplID1 As Long = 0, Optional iSamplID2 As Long = 0, _
                Optional iKcnt As Integer = 0) As Integer
    
    On Error GoTo Apl_down
    Dim liCnt As Integer
    
    '�߂�l������
    funCrySogoHantei_CC600Multi = FUNCTION_RETURN_FAILURE
    TotalJudg = True
    
    '�O���[�o���ϐ��ɐݒ�
    ciSmpGetFlg = iSmpGetFlg
    ciKcnt = iKcnt
    
    '�u���b�NID��ݒ�
    sErr_Msg = "������������(��ۯ�ID�ݒ�)"
    typ_b.BLOCKID = sKeyID
    
    '��ʏ��ݒ�
    sErr_Msg = "������������(SetAllData)"
    
    If SetAllData2(typ_b, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg, iSmpGetFlg, iSamplID1, iSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
  
  
    '�d�l�����w���擾
    sErr_Msg = "������������(SpecJudgCheck)"
    Call SpecJudgCheck
    
    '2003/12/13 SystemBrain Null�Ή��ǉ���
    '�d�lNull�`�F�b�N
    sErr_Msg = "�d�lNull����"
    If funCryChkNull(typ_b.typ_si(BlkTop), sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    '2003/12/13 SystemBrain Null�Ή��ǉ���
    
    '���уf�[�^����(TOP)
    sErr_Msg = "������������(����(TOP))"
    
    
    
    '----TEST2004/10
    '��ʏo�͗p�Ɏ�����R�l��ޔ����Ă���
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS5
        
    If giTpMultiFlg = 0 Then ''<<�����i�Ԕ���Ή����ϕ����@���Top�̂Ƃ�������
        '--Top�̂Ƃ���typ_B.typ_zi.CRYRZ(BlkTop).JMEASxx��ێ�
        Erase pJMEAS_Top
        ReDim pJMEAS_Top(5)
        pJMEAS_Top(1) = typ_b.typ_zi.CRYRZ(BlkTop).MEAS1
        pJMEAS_Top(2) = typ_b.typ_zi.CRYRZ(BlkTop).MEAS2
        pJMEAS_Top(3) = typ_b.typ_zi.CRYRZ(BlkTop).MEAS3
        pJMEAS_Top(4) = typ_b.typ_zi.CRYRZ(BlkTop).MEAS4
        pJMEAS_Top(5) = typ_b.typ_zi.CRYRZ(BlkTop).MEAS5
        psKSTAFFID = typ_b.typ_zi.CRYRZ(BlkTop).KSTAFFID
        psHSXRSPOT = typ_b.typ_si(BlkTop).HSXRSPOT
        psHSXRSPOI = typ_b.typ_si(BlkTop).HSXRSPOI

        If Trim(typ_b.typ_zi.CRYRZ(BlkTop).KSTAFFID) <> KSTAFF_J002 Then
            '��R�l�𑪒�ʒu�R�[�h�ɂ����בւ���
            ''--<<<<���Top
            If Set_Rs_Ichi(typ_b.typ_si(BlkTop).HSXRSPOT, typ_b.typ_si(BlkTop).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTop).MEAS1, _
                            typ_b.typ_zi.CRYRZ(BlkTop).MEAS2, typ_b.typ_zi.CRYRZ(BlkTop).MEAS3, typ_b.typ_zi.CRYRZ(BlkTop).MEAS4, typ_b.typ_zi.CRYRZ(BlkTop).MEAS5) = FUNCTION_RETURN_FAILURE Then
                GoTo Apl_down
            End If
        End If
    End If
    

    If giBtMultiFlg = 0 And giTpMultiFlg = 1 Then ''<<<<���ϕ����@��Ԃ�����Bottom�̂Ƃ�������
        typ_b.typ_zi.CRYRZ(BlkTop).MEAS1 = pJMEAS_Top(1)
        typ_b.typ_zi.CRYRZ(BlkTop).MEAS2 = pJMEAS_Top(2)
        typ_b.typ_zi.CRYRZ(BlkTop).MEAS3 = pJMEAS_Top(3)
        typ_b.typ_zi.CRYRZ(BlkTop).MEAS4 = pJMEAS_Top(4)
        typ_b.typ_zi.CRYRZ(BlkTop).MEAS5 = pJMEAS_Top(5)

        typ_b.typ_zi.CRYRZ(BlkTop).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS1
        typ_b.typ_zi.CRYRZ(BlkTop).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS2
        typ_b.typ_zi.CRYRZ(BlkTop).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS3
        typ_b.typ_zi.CRYRZ(BlkTop).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS4
        typ_b.typ_zi.CRYRZ(BlkTop).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS5

        If Trim(psKSTAFFID) <> KSTAFF_J002 Then
            '��R�l�𑪒�ʒu�R�[�h�ɂ����בւ���
            ''--<<<<���Top
            If Set_Rs_Ichi(psHSXRSPOT, psHSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTop).MEAS1, _
                            typ_b.typ_zi.CRYRZ(BlkTop).MEAS2, typ_b.typ_zi.CRYRZ(BlkTop).MEAS3, typ_b.typ_zi.CRYRZ(BlkTop).MEAS4, typ_b.typ_zi.CRYRZ(BlkTop).MEAS5) = FUNCTION_RETURN_FAILURE Then
                GoTo Apl_down
            End If
        End If
    End If
    
    ''-Top�̔���>>>>>>>>>>>>>>>>>���ϕ���
    If giTpMultiFlg = 0 Then '�S��������
        If CrAllJudg(typ_b, tNew_Hinban, BlkTop) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    ElseIf giTpMultiFlg = 1 Then ''Cs,LT,EPD�ō��۔���
        If CrAllJudgCC600Multi(typ_b, tOld_Hinban, BlkTop) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If
    ''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '���уf�[�^����(TAIL)
    sErr_Msg = "������������(����(TAIL))"
    
    '��ʏo�͗p�Ɏ�����R�l��ޔ����Ă���
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS5
        
    If giBtMultiFlg = 0 Then ''<<<<���ϕ����@��Ԃ�����Bottom�̂Ƃ�������
        If Trim(typ_b.typ_zi.CRYRZ(BlkTail).KSTAFFID) <> KSTAFF_J002 Then
            '��R�l�𑪒�ʒu�R�[�h�ɂ����בւ���
            ''--<<<<���Bottom
            If Set_Rs_Ichi(typ_b.typ_si(BlkTail).HSXRSPOT, typ_b.typ_si(BlkTail).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTail).MEAS1, _
                            typ_b.typ_zi.CRYRZ(BlkTail).MEAS2, typ_b.typ_zi.CRYRZ(BlkTail).MEAS3, typ_b.typ_zi.CRYRZ(BlkTail).MEAS4, typ_b.typ_zi.CRYRZ(BlkTail).MEAS5) = FUNCTION_RETURN_FAILURE Then
                GoTo Apl_down
            End If
        End If
    End If
    ''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>���ϕ���
    '--Bottom�̔���
    If giBtMultiFlg = 0 Then ''�S��������
        If CrAllJudg(typ_b, tNew_Hinban, BlkTail) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    ElseIf giBtMultiFlg = 1 Then ''Cs,LT,EPD�ō��۔���
        If CrAllJudgCC600Multi(typ_b, tNew_Hinban, BlkTail) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If
    ''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    bTotalJudg = TotalJudg
    
    funCrySogoHantei_CC600Multi = FUNCTION_RETURN_SUCCESS
'------------------------------------------ �I������  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funCrySogoHantei_CC600Multi = -4
    iErr_Code = funCrySogoHantei_CC600Multi
    GoTo Apl_Exit
    
End Function

'�T�v      :���㌋������  CC600�}���`�u���b�N�Ή�
'���Ұ�    :�ϐ���        ,IO ,�^               :����
'          :typ_B         ,I  ,typ_AllTypesB    :�e���\����
'          :tNew_Hinban   ,I  ,tFullHinban      :�U�֌��i��
'          :tt            ,I  ,Integer          :TopTail����p
'����      :�����w���ɏ]���ACs,EPD,LT�̎��є�����s��
'����      : �����i�Ԕ���Ή��@20060501 SMP����  CrAllJudg�̉���
'''
Public Function CrAllJudgCC600Multi(typ_b As typ_AllTypesB, tNew_Hinban As tFullHinban, tt As Integer) As FUNCTION_RETURN
    Dim IND         As String                   '�����w��
    Dim bJudg       As Boolean
    Dim i           As Integer
    Dim cnt         As Integer
    Dim typTmList() As typ_TBCMB005
    Dim minwk       As String, maxwk As String
    Dim vTemp       As Variant
    Dim RET         As FUNCTION_RETURN
    Dim Gd_si()     As type_DBDRV_scmzc_fcmkc001c_Siyou
    Dim jCs         As String                               '�u���b�N���i�Ԃ�Cs�ۏ�
    Dim jCsFromTo   As String                               '�u���b�N���i�Ԃ�Cs�ۏ�(FromTo)
    Dim hasSiji     As Boolean                              '�����w������
    Dim sHinban12   As String                               '�i��(12��)
    Dim bJudgXY     As Boolean                              'X������p�t���O�ǉ� 2009/10/22
    Dim bJudgX      As Boolean                              'X������p�t���O�ǉ� 2009/10/22
    Dim bJudgY      As Boolean                              'X������p�t���O�ǉ� 2009/10/22
    Dim Oi          As C_Oi       '2010/03/12
    
    CrAllJudgCC600Multi = FUNCTION_RETURN_FAILURE
    
    sHinban12 = tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond
    
    i = 0
       
    '�����R�[�h���X�g�擾
    If GetCodeList(MSYSCLASS, KCLASS, typTmList) <> FUNCTION_RETURN_SUCCESS Then
        '�����R�[�h���X�g�擾���s
        Exit Function
    End If
    With typ_b
'>>>>> Oi�̒ǉ� 2011/02/09 SETsw kubota -------------------------
        '�����w���ݒ�
        IND = IIf(tt = BlkTop, "123", "123")
        '' ���������w��(Oi)*****************************************************************
        If JudgSC_B(tt).Oi Then
            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).BLOCKNG = False
            .typ_rslt(tt, i).pos = -1                                       ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())       ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                               ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                               ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                     ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
            .typ_rslt(tt, i).SMPLNO = -1                                    ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                    ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
            bJudg = False
            If (InStr(IND, .typ_cr(tt).CRYINDOICS) <> 0) Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.OIZ(tt).POSITION                 ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())       ' ���e
                .typ_rslt(tt, i).SMPLNO = .typ_zi.OIZ(tt).SMPLNO                ' �T���v���m��
                .typ_rslt(tt, i).INFO2 = "�����L"                               ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                               ' ���Q
                If left(.typ_zi.OIZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(tt, i).INFO3 = "����ٖ�"                          ' ���R
                    If .typ_zi.OIZ(tt).SMPLUMU = "0" Then
                        'OI���莸�s
                        .typ_rslt(tt, i).INFO3 = "����Err"                      ' ���Q
                        'OI����
                        If CrOiJudg(.typ_si(tt), .typ_zi.OIZ(tt), bJudg) Then
                            Call GetOiMaxMin(.typ_zi.OIZ(tt), minwk, maxwk)
                            '��ʕ\�����e�ݒ�
                            vTemp = CStr(.typ_zi.OIZ(tt).OIMEAS1)                       ' ���P
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' ���P
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(maxwk, "0.00")     ' ���Q
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(minwk, "0.00")     ' ���R
                            vTemp = CStr(.typ_zi.OIZ(tt).ORGRES)                        ' ���S
                            'ORG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                            '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' ���S
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' ���S
                        Else
                            If .typ_zi.OIZ(tt).ORGRES = -999 Then               ' 2010/03/12 Kameda
                                ReDim Oi.Oi(4)
                                Oi.Oi(0) = .typ_zi.OIZ(tt).OIMEAS1
                                Oi.Oi(1) = .typ_zi.OIZ(tt).OIMEAS2
                                Oi.Oi(2) = .typ_zi.OIZ(tt).OIMEAS3
                                Oi.Oi(3) = .typ_zi.OIZ(tt).OIMEAS4
                                Oi.Oi(4) = .typ_zi.OIZ(tt).OIMEAS5
                                .typ_rslt(tt, i).INFO1 = "�d�l" & .typ_si(tt).HSXONSPT & "�_"   ' ���P
                                .typ_rslt(tt, i).INFO2 = "����" & GetTensu(Oi) & "�_"                                ' ���Q
                                .typ_rslt(tt, i).INFO4 = "�_���s��"     ' ���S
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                               ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                               ' ���茋��
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00101"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDOICS) <> 0) Then
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).OKNG = "OK"                                ' ���茋��
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.OIZ(tt).POSITION             ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())   ' ���e
                .typ_rslt(tt, i).SMPLNO = .typ_zi.OIZ(tt).SMPLNO            ' �T���v���m��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N�Q"                                ' ���茋��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).INFO1 = "�d�l��"                           ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                           ' ���Q
                .typ_rslt(tt, i).INFO3 = "����ٖ�"                          ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
                .typ_rslt(tt, i).hinban = sHinban12                         ' �i��(12��)
                If .typ_zi.OIZ(tt).SMPLUMU = "0" Then
                    'OI���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                          ' ���Q
                    'OI����
                    If CrOiJudg(.typ_si(tt), .typ_zi.OIZ(tt), bJudg) Then
                        Call GetOiMaxMin(.typ_zi.OIZ(tt), minwk, maxwk)
                        '��ʕ\�����e�ݒ�
                        vTemp = CStr(.typ_zi.OIZ(tt).OIMEAS1)                       ' ���P
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' ���P
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(maxwk, "0.00")     ' ���Q
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(minwk, "0.00")     ' ���R
                        vTemp = CStr(.typ_zi.OIZ(tt).ORGRES)                        ' ���S
                        'ORG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                        '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' ���S
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' ���S
                    End If
                End If
                i = i + 1
            End If
        End If
'<<<<< Oi�̒ǉ� 2011/02/09 SETsw kubota -------------------------
        '' ���������w��(Cs)*****************************************************************
        '�����w���ݒ�
        IND = IIf(tt = BlkTop, "123", "123")
        If JudgSC_B(tt).Cs And (tt = BlkTail Or .typ_si(tt).HSXCNKHI = "6" Or .typ_si(tt).HSXCNKHI = "9") Then  'TOP/BOT�ۏؑΉ� 09/01/08 ooba
            '��ʕ\�����e������
            .typ_rslt(tt, i).BLOCKNG = False
            .typ_rslt(tt, i).pos = -1                                   ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())   ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                           ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                           ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                 ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
            .typ_rslt(tt, i).SMPLNO = -1                                ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                         ' �i��(12��)
            bJudg = False
            If (InStr(IND, .typ_cr(tt).CRYINDCSCS) <> 0) Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.CSZ(tt).POSITION                 ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())       ' ���e
                .typ_rslt(tt, i).SMPLNO = .typ_zi.CSZ(tt).SMPLNO                ' �T���v���m��
                .typ_rslt(tt, i).INFO2 = "�����L"                               ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                               ' ���Q
                If left(.typ_zi.CSZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(tt, i).INFO3 = "����ٖ�"                          ' ���R
                    If .typ_zi.CSZ(tt).SMPLUMU = "0" Then
                        'Cs���莸�s
                        .typ_rslt(tt, i).INFO3 = "����Err"                      ' ���Q
                        'CS����擾
                        If CrCsjudg(.typ_si(tt), .typ_zi.CSZ(tt), bJudg) Then
                            '��ʕ\�����e�ݒ�
                            vTemp = CStr(.typ_zi.CSZ(tt).CSMEAS)
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00") ' ���P
                            .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                            .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                        End If
                    End If
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                               ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                               ' ���茋��
                TotalJudg = False
            End If
            i = i + 1
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDCSCS) <> 0) Then
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).OKNG = "OK"                                    ' ���茋��
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.CSZ(tt).POSITION                 ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())       ' ���e
                .typ_rslt(tt, i).SMPLNO = .typ_zi.CSZ(tt).SMPLNO                ' �T���v���m��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N�Q"                                   ' ���茋��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).INFO1 = "�d�l��"                               ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                               ' ���Q
                .typ_rslt(tt, i).INFO3 = "����ٖ�"                              ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
                If .typ_zi.CSZ(tt).SMPLUMU = "0" Then
                    'Cs���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                      ' ���Q
                    'CS����擾
                    If CrCsjudg(.typ_si(tt), .typ_zi.CSZ(tt), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CStr(.typ_zi.CSZ(tt).CSMEAS)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00") ' ���P
                        .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                        .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                    End If
                End If
                i = i + 1
            End If
        End If
        
        '' ���������w��(T)*****************************************************************
Dim HIN As tFullHinban
Dim LTSPI As String

        If (InStr(IND, .typ_cr(tt).CRYINDTCS) <> 0) Then
            hasSiji = True
        Else
            hasSiji = False
        End If
        bJudg = True                                        '2004/01/15 SystemBrain
        If (JudgSC_B(tt).Lt) And (tt = BlkTail) Then        '2004/01/15 SystemBrain
            bJudg = False                                   '2004/01/15 SystemBrain
        Else                                                '2004/01/15 SystemBrain
            JudgSC_B(tt).Lt = False                         '2004/01/15 SystemBrain
        End If                                              '2004/01/15 SystemBrain
        
        'LT��Bot�[�Ńu���b�N�S��𔻒肷�邱�ƂɂȂ������߁A�uTop�[�i�Ԃ�LT�w���������Bot�ŕ\���v�͕s�v�ƂȂ���
        If (JudgSC_B(tt).Lt) Or (hasSiji And (tt = BlkTail)) Then '�d�l���� or Bot�[�Ō�������
            .typ_rslt(tt, i).BLOCKNG = False
            
            '��ʕ\�����e������
            .typ_rslt(tt, i).pos = .typ_zi.LTZ(tt).POSITION             ' �������J�n�ʒu
            .typ_rslt(tt, i).SMPLNO = -1                                ' �T���v���m��
            .typ_rslt(tt, i).NAIYO = Search_CrCode("T", typTmList())    ' ���e
            If JudgSC_B(tt).Lt Then
                .typ_rslt(tt, i).INFO1 = "�d�l�L"                       ' ���P
            Else
                .typ_rslt(tt, i).INFO1 = "�d�l��"
                bJudg = True
            End If
            If hasSiji Then
                .typ_rslt(tt, i).INFO2 = "�����L"                       ' ���Q
            Else
                .typ_rslt(tt, i).INFO2 = "������"
            End If
            .typ_rslt(tt, i).INFO3 = "���і�"                           ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
            .typ_rslt(tt, i).hinban = sHinban12                         ' �i��(12��)
            
            '���C�t�^�C��
            bJudgX = True   '10������
            '����ƌ��ʓo�^
            If .typ_zi.LTZ(tt).CRYNUM = .typ_si(1).CRYNUM Then
                .typ_rslt(tt, i).pos = .typ_zi.LTZ(tt).POSITION                 ' �������J�n�ʒu
                .typ_rslt(tt, i).SMPLNO = .typ_zi.LTZ(tt).SMPLNO                ' �T���v���m��
                If (.typ_zi.LTZ(tt).SMPLUMU <> "0") Then
                    .typ_rslt(tt, i).INFO3 = "����ٖ�"                          ' ���R
                Else
                    '2005/12/02 add SET���� LT�v�Z�֐�call ->
                    '���C�t�^�C���l���v�Z���Ȃ���
                    Call Sub_LTReCalc(.typ_si(tt), .typ_zi.LTZ(tt))
                    '2005/12/02 add SET���� LT�v�Z�֐�call <-
                    
                    'LT����擾
                    If CrLtjudg(.typ_si(tt), .typ_zi.LTZ(tt), bJudg) Then
''Add Start 2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
                        If CrLt10judg(.typ_si(tt), .typ_zi.LTZ(tt), .typ_cr(tt), bJudgX) Then
''Add End   2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
                            '��ʕ\�����e�ݒ�
                            vTemp = CStr(.typ_zi.LTZ(tt).CALCMEAS)                  ' ���P
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")
                            vTemp = CStr(.typ_zi.LTZ(tt).MEASPEAK)                  ' ���Q
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")
                            .typ_rslt(tt, i).INFO3 = ""                             ' ���R
''Add Start 2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
                            ' ���S
                            If .typ_zi.LTZ(tt).CONVAL = (-1) Then
                                .typ_rslt(tt, i).INFO4 = "NULL"
                            Else
                                .typ_rslt(tt, i).INFO4 = CStr(.typ_zi.LTZ(tt).CONVAL)
                            End If
                        Else
                            .typ_rslt(tt, i).INFO3 = "LT10����Err"                  ' ���R
                        End If
''Add End   2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
                    Else
                        .typ_rslt(tt, i).INFO3 = "����Err"                      ' ���R
                    End If
                End If
            Else    '���тȂ�
                If JudgSC_B(tt).Lt Then bJudg = False
            End If
            
''Add Start 2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
            If bJudg = True Then
                If bJudgX = True Then
                    bJudg = True
                Else
                    bJudg = False
                End If
            End If
''Add End   2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
            
            If tt = BlkTail Then ''<<Tail�̂Ƃ��̂ݔ��肳����
            If (bJudg = False) Then
                .typ_rslt(tt, i).OKNG = "NG"
                TotalJudg = False
'====================== Debug Debug =====================================
            ElseIf .typ_si(tt).HSXLTHWS = "S" Then
                .typ_rslt(tt, i).OKNG = "N�Q"                            ' ���茋��
'====================== Debug Debug =====================================
            Else
                .typ_rslt(tt, i).OKNG = "OK"                            ' ���茋��
            End If
            End If
            i = i + 1
        End If
        '' ���������w��(EPD)*****************************************************************
        If JudgSC_B(tt).EPD Then
            If tt = BlkTop Then
                .typ_rslt(tt, i).BLOCKNG = False
                If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                ' �������J�n�ʒu
                    .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO               ' �T���v���m��
                    .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())      ' ���e
                    .typ_rslt(tt, i).INFO1 = "�d�l�L"                               ' ���P
                    .typ_rslt(tt, i).INFO2 = "�����L"                               ' ���Q
                    .typ_rslt(tt, i).INFO3 = "���і�"                               ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                    .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
                    bJudg = False
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).INFO3 = "����ٖ�"                          ' ���R
                        If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                            'EPD���莸�s
                            .typ_rslt(tt, i).INFO3 = "����Err"                      ' ���R
                            'EPD����擾
                            If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                                '��ʕ\�����e�ݒ�
                                vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                  ' ���P
                                .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' ���P
                                .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                                .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                            End If
                        End If
                    End If
                    .typ_rslt(tt, i).OKNG = "OK"                               ' ���茋�ʒ��K���킹
                    i = i + 1
                End If
            Else
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = -1           ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())          ' ���e
                .typ_rslt(tt, i).INFO1 = "�d�l�L"                                   ' ���P
                .typ_rslt(tt, i).INFO2 = "������"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
                .typ_rslt(tt, i).SMPLNO = -1                                        ' �T���v���m��
                .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                bJudg = False
                If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).INFO3 = "����ٖ�"                              ' ���R
                        If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                            .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION            ' �������J�n�ʒu
                            .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO           ' �T���v���m��
                            'EPD���莸�s
                            .typ_rslt(tt, i).INFO3 = "����Err"                          ' ���R
                            'EPD����擾
                            If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                                '��ʕ\�����e�ݒ�
                                vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                  ' ���P
                                .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' ���P
                                .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                                .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                            End If
                        End If
                    End If
                Else
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                ' �������J�n�ʒu
                        .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO               ' �T���v���m��
                        'EPD���莸�s
                        .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���R
                        'EPD����擾
                        If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                            '��ʕ\�����e�ݒ�
                            vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)          ' ���P
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' ���P
                            .typ_rslt(tt, i).INFO2 = ""                                 ' ���Q
                            .typ_rslt(tt, i).INFO3 = ""                                 ' ���R
                        End If
                    End If
                End If
                If bJudg = True Then
                    .typ_rslt(tt, i).OKNG = "OK"                               ' ���茋��
                Else
                    .typ_rslt(tt, i).OKNG = "NG"                               ' ���茋��
                    TotalJudg = False
                End If
                i = i + 1
            End If
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                    ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())          ' ���e
                .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "����ٖ�"                                  ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
                .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO                   ' �T���v���m��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N�Q"                                        ' ���茋��
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                    'EPD���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���R
                    'EPD����擾
                    If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                      ' ���P
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' ���P
                        .typ_rslt(tt, i).INFO2 = ""                                 ' ���Q
                        .typ_rslt(tt, i).INFO3 = ""                                 ' ���R
                    End If
                End If
                i = i + 1
            End If
        End If
        'SIRD����f�[�^�ݒ�   2010/02/04 add Kameda
        If tt = BlkTop Then
            .typ_rslt(tt, i).BLOCKNG = False
            If .typ_cr(tt).SIRDKBNY3 = "1" Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.SIRD.POSITION                ' �������J�n�ʒu
                '.typ_rslt(tt, i).SMPLNO = .typ_zi.SIRD.SMPLNO               ' �T���v���m��
                .typ_rslt(tt, i).NAIYO = Search_CrCode("SD", typTmList())       ' ���e
                .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
                bJudg = False
                'SIRD����擾
                If CrSIRDjudg(.typ_si(tt), .typ_zi.SIRD, bJudg) Then
                    '��ʕ\�����e�ݒ�
                    vTemp = CStr(.typ_zi.SIRD.SIRDCNT)                  ' ���P
                    .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' ���P
                    .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                    .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                End If
                If bJudg = True Then
                    .typ_rslt(tt, i).OKNG = "OK"                               ' ���茋��
                Else
                    .typ_rslt(tt, i).OKNG = "NG"                               ' ���茋��
                    TotalJudg = False
                    ''gsTbcmy028ErrCode = ""
                End If
                '�]���҂��Q�Ǝ�  2010/02/18 Kameda
                If .typ_zi.SIRD.NothingFlg = "1" Then
                    .typ_rslt(tt, i).INFO1 = ""                                ' ���P
                    .typ_rslt(tt, i).OKNG = "�]���҂�"                         ' ���茋��
                End If
                i = i + 1
            ElseIf .typ_cr(tt).SIRDKBNY3 = "2" Then       '2010/02/16 add Kameda
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = .typ_zi.SIRD.POSITION                 ' �������J�n�ʒu
                '.typ_rslt(tt, i).SMPLNO = .typ_zi.SIRD.SMPLNO               ' �T���v���m��
                .typ_rslt(tt, i).NAIYO = Search_CrCode("SD", typTmList())       ' ���e
                .typ_rslt(tt, i).hinban = sHinban12                             ' �i��(12��)
                'bJudg = False    �\���̂�
                'SIRD�\��
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).INFO1 = "��s�]��"                     ' ���P
                .typ_rslt(tt, i).INFO2 = ""                             ' ���Q
                .typ_rslt(tt, i).INFO3 = ""                             ' ���R
                .typ_rslt(tt, i).OKNG = "OK"                            ' ���茋��
                i = i + 1
            End If
        End If
        
        'X������f�[�^�ݒ�   2009/08/12 add Kameda
        '�����p�݂̂Ŕ��� X,Y�͌x�����o��(�w�i�ԁj  2009/10/22 add Kameda
        If tt = BlkTail Then
            If .typ_cr(tt).CRYINDXC1 <> 0 Then
                'If CrXjudg(.typ_si(tt), .typ_zi.XZ, bJudg) Then     2009/10/22 Kameda
                If CrXjudg(.typ_si(tt), .typ_zi.XZ, bJudgXY, bJudgX, bJudgY) Then
                    If bJudgXY Then
                        '.typ_zi.XZ.JUDG = "OK"    2009/10/22
                        .typ_zi.XZ.JUDGXY = "OK"
                    Else
                        '.typ_zi.XZ.JUDG = "NG"    2009/10/22
                        .typ_zi.XZ.JUDGXY = "NG"
                        TotalJudg = False
                    End If
                    '�x�����o�����߂ɍ��ڒǉ�     2009/10/22 Kameda
                    If bJudgX Then
                        .typ_zi.XZ.JUDGX = "OK"
                    Else
                        .typ_zi.XZ.JUDGX = "NG"
                    End If
                    If bJudgY Then
                        .typ_zi.XZ.JUDGY = "OK"
                    Else
                        .typ_zi.XZ.JUDGY = "NG"
                    End If
                End If
            Else
                '.typ_zi.XZ.JUDG = ""     2009/10/22
                .typ_zi.XZ.JUDGXY = ""
                .typ_zi.XZ.JUDGX = ""
                .typ_zi.XZ.JUDGY = ""
            End If
        End If
        
      'Add Start 2011/02/01 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎��є��菈��
        Call CuDecoDataSet_C(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJ(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJLT(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJ2(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
      ''Add End   2011/02/01 SMPK A.Nagamine
        
    End With
        
    CrAllJudgCC600Multi = FUNCTION_RETURN_SUCCESS
End Function
'�T�v      :����_�����擾����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:OI           ,I  ,           ,
'      �@�@:�߂�l       ,O  ,Long           �@,����_��
'�쐬      :2010/3/12 Kameda
Public Function GetTensu(Oi As C_Oi) As Integer

    Dim i As Integer
        GetTensu = 5
        For i = 0 To 4
            If Oi.Oi(i) = -1 Then
                GetTensu = i
                Exit Function
            End If
        Next
End Function

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C)�̎��є��菈��
Public Function CuDecoDataSet_C(pJudgSC_B() As Judg_Spec_Cry, ptyp_b As typ_AllTypesB, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String, Optional pblnFlag As Boolean = False) As Boolean
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech End
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                  '�����w��
    Dim bJudg           As Boolean
    Dim strXTALC1       As String
    Dim strJDGEIDC      As String
    Dim strSYNFLG       As String
    Dim strYMKFLG       As String
    Dim lngSMPPOS       As Long
    Dim iErrFlg         As Integer
    Dim YFlg            As Boolean
    Dim intRet          As Integer
    Dim sResult         As String
    Dim strSampUmu      As String
    Dim strInfo(4)      As String
    Dim intSiyou        As Integer
    Dim intJisseki      As Integer
    Dim StrCryNum       As String
    Dim lngSampNo       As Long
    Dim strPtnJsk       As String
    Dim str028ErrCode   As String
    Dim blnRet          As Boolean
    Dim typ_Ret_CuDeco As typ_SB_com_xodb5_osf31_Cudeco
    Dim intNG_Num       As Integer
    
    blnRet = False
    intNG_Num = 0
    
    '�����w���ݒ�
    IND = IIf(UpDo = BlkTop, "123", "123")
    
    With ptyp_b
        JudgSpecCode = pJudgSC_B(UpDo).CuC
        SCC = "C"
        shiji = (InStr(IND, .typ_cr(UpDo).CRYINDCCS) <> 0)
        strSampUmu = .typ_zi.CuC(UpDo).SMPLUMUC                 ' C.�T���v���L�� TBCMJ023.SMPLUMUC "0"=�T���v���L��,"1"=�T���v������
        strInfo(0) = CStr(.typ_zi.CuC(UpDo).CDISKJSK)           ' ���P C.Disk���a
        strInfo(1) = CStr(.typ_zi.CuC(UpDo).CRINGNKJSK)         ' ���Q C.Ring���a
        strInfo(2) = CStr(.typ_zi.CuC(UpDo).CRINGGKJSK)         ' ���R C.Ring�O�a
        strInfo(3) = .typ_zi.CuC(UpDo).CPTNJSK                  ' ���S C.�p�^�[������
        lngSMPPOS = .typ_zi.CuC(UpDo).POSITION                  ' C.����وʒu
        StrCryNum = .typ_zi.CuC(UpDo).CRYNUM                    ' C.�����ԍ�
        lngSampNo = .typ_zi.CuC(UpDo).SMPLNO                    ' C.�T���v����(��\�T���v��ID)
        strPtnJsk = .typ_zi.CuC(UpDo).CPTNJSK                   ' C.�p�^�[������
        str028ErrCode = "00155"
        
        '�ۏ��׸�="H"�̏ꍇ����V�p�����[�^Optional ��True�w�肳�ꂽ�Ƃ�
        If (JudgSpecCode) Or (pblnFlag) Then
            
            bJudg = False
            
            '�����׸ޏ�����
            iErrFlg = 0
            YFlg = False
            
            '��ۯ�ID
            strXTALC1 = Trim(ptyp_b.BLOCKID)
            '�����ԍ�
            strXTALC1 = left(strXTALC1, 9) & "000"
            
            '�����ԍ��𷰂Ƃ���XSDC1���C�|OSF3����ID���l������
            If GetCOSF3ID(strJDGEIDC, strXTALC1) <> FUNCTION_RETURN_SUCCESS Then
                iErrFlg = 11
            Else
                If Trim(strJDGEIDC) = "" Then
                    iErrFlg = 12
'                '����ID=�u9�v�̏ꍇ�͔���Ȃ�(����OK)�@07/08/01 M.Kaga
'                ElseIf Trim(strJDGEIDC) = "9" Then
'                    YFlg = True
'                    bJudg = True
                Else
                    '�l������C-OSF3����ID��XODC5_OSF30��菳�F�׸ނ̊l��
                    If GetSYNFLAGC5(strSYNFLG, strYMKFLG, strJDGEIDC) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 13
                    Else
                        '���F�׸�:0�@�����F�̏ꍇ
                        If Trim(strSYNFLG) = "0" Or Trim(strSYNFLG) = "" Or IsNull(strSYNFLG) Then
                            iErrFlg = 14
                        End If
                    End If
                End If
            End If
            
            If iErrFlg > 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' ���e
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                        ' ���R
                .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' �T���v���m��
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' �i��(12��)
                If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS           ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo          ' �T���v���m��
                End If
                bJudg = False
            Else
                If YFlg = False Then
                    
                    If GetOsf31_CuDeco(strJDGEIDC, lngSMPPOS, typ_Ret_CuDeco) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 21
                    Else
                        intSiyou = -1
                        intJisseki = -1
                        
                        If (IsNumeric(.typ_si(UpDo).HSXCPK)) Then
                            intSiyou = CInt(.typ_si(UpDo).HSXCPK)
                        End If
                        
                        If (IsNumeric(strPtnJsk)) Then
                            intJisseki = CInt(strPtnJsk) + 1
                        End If
                        
                        If (intSiyou >= 1) And (intSiyou <= 4) And (intJisseki >= 1) And (intJisseki <= 4) Then
                        
                           'intRet = funCodeDBGet(SYSCLASS, CLASS, ���i�d�l�p�^�[���敪, 1, ���уp�^�[���敪, �߂�l(tbcmb005.info1))
                            intRet = funCodeDBGet("SB", "S1", CStr(intSiyou), 1, CStr(intJisseki), sResult)
                            If (intRet = 0) And (sResult <> vbNullString) And (Len(sResult) >= 1) Then
                                If sResult = "1" Then
                                    bJudg = True
                                Else
                                    bJudg = False
                                    intNG_Num = 51
                                End If
                            Else
                                'sErr_Msg = sAdd_Msg & sErr_Msg & "���d�l:" & CStr(intSiyou) & ", ����:" & CStr(intJisseki)
                                bJudg = False
                                intNG_Num = 52
'                                GoTo CodeDBGet_Error
                            End If
                            
                        End If
                    End If
                End If
                
                If iErrFlg > 0 Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                        ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = CInt(iErrFlg)                    ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' �T���v���m��
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' �i��(12��)
                    If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' �������J�n�ʒu
                        .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' �T���v���m��
                    End If
                    bJudg = False
                    
                Else
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                             ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                             ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' �T���v���m��
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                    
                    If shiji Then
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                        .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                         ' ���R
                        
                        If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                            .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' �������J�n�ʒu
                            .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' �T���v���m��
                            .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                    ' ���R
                            
                            If strSampUmu = "0" Then
                                .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                ' ���R
                                '��ʕ\�����e�ݒ�
                                .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                                .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                                .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                                .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' ���茋��
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' ���茋��
                TotalJudg = False
                gsTbcmy028ErrCode = str028ErrCode
            End If
            DispLineCount = DispLineCount + 1
            
        Else
        ' Add Start 2011/02/15 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : �u�ۏؕ��@_���v�Q�l���̏����ǉ�
            If shiji Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS                              ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                             ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                             ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                            ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo                           ' �T���v���m��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N�Q"                                 ' ���茋��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                    
                    If strSampUmu = "0" Then
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                    End If
                DispLineCount = DispLineCount + 1
            End If
        ' Add End   2011/02/15 SMPK A.Nagamine
        End If      '/*  End of If (JudgSpecCode) Or (pblnFlag) Then */
    End With
    
    CuDecoDataSet_C = blnRet
    
End Function
''Add End   2011/01/17 SMPK A.Nagamine


'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : Cu-deco(CJ)�̎��є��菈��
Public Function CuDecoDataSet_CJ(pJudgSC_B() As Judg_Spec_Cry, ptyp_b As typ_AllTypesB, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String, Optional pblnFlag As Boolean = False) As Boolean
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech End
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                  '�����w��
    Dim bJudg           As Boolean
    Dim strXTALC1       As String
    Dim strJDGEIDC      As String
    Dim strSYNFLG       As String
    Dim strYMKFLG       As String
    Dim lngSMPPOS       As Long
    Dim iErrFlg         As Integer
    Dim YFlg            As Boolean
    Dim intRet          As Integer
    Dim sResult         As String
    Dim strSampUmu      As String
    Dim strInfo(4)      As String
    Dim intSiyou        As Integer
    Dim intJisseki      As Integer
    Dim StrCryNum       As String
    Dim lngSampNo       As Long
    Dim strPtnJsk       As String
    Dim str028ErrCode   As String
    Dim blnRet          As Boolean
    Dim typ_Ret_CuDeco  As typ_SB_com_xodb5_osf31_Cudeco
    Dim intNG_Num       As Integer
    Dim intMax          As Integer
    
    blnRet = False
    intNG_Num = 0
    
    '�����w���ݒ�
    IND = IIf(UpDo = BlkTop, "123", "123")
    
    With ptyp_b
        ' CJ ����
        JudgSpecCode = pJudgSC_B(UpDo).CuCJ
        SCC = "CJ"
        shiji = (InStr(IND, .typ_cr(UpDo).CRYINDCJCS) <> 0)
        strSampUmu = .typ_zi.CuCJ(UpDo).SMPLUMUCJ               ' CJ.�T���v���L�� TBCMJ023.SMPLUMUCJ "0"=�T���v���L��,"1"=�T���v������
        strInfo(0) = CStr(.typ_zi.CuCJ(UpDo).CJDISKJSK)         ' ���P CJ.Disk���a
        strInfo(1) = CStr(.typ_zi.CuCJ(UpDo).CJRINGNKJSK)       ' ���Q CJ.Ring���a
        strInfo(2) = CStr(.typ_zi.CuCJ(UpDo).CJRINGGKJSK)       ' ���R CJ.Ring�O�a
        strInfo(3) = .typ_zi.CuCJ(UpDo).CJPTNJSK                ' ���S CJ.�p�^�[������
        lngSMPPOS = .typ_zi.CuCJ(UpDo).POSITION                 ' CJ.����وʒu
        StrCryNum = .typ_zi.CuCJ(UpDo).CRYNUM                   ' CJ.�����ԍ�
        lngSampNo = .typ_zi.CuCJ(UpDo).SMPLNO                   ' CJ.�T���v����(��\�T���v��ID)
        strPtnJsk = .typ_zi.CuCJ(UpDo).CJPTNJSK                 ' CJ.�p�^�[������
        str028ErrCode = "00156"
        
        '�ۏ��׸�="H"�̏ꍇ����V�p�����[�^Optional ��True�w�肳�ꂽ�Ƃ�
        If (JudgSpecCode) Or (pblnFlag) Then
            
            bJudg = False
            
            '�����׸ޏ�����
            iErrFlg = 0
            YFlg = False
            
            '��ۯ�ID
            strXTALC1 = Trim(ptyp_b.BLOCKID)
            '�����ԍ�
            strXTALC1 = left(strXTALC1, 9) & "000"
            
            '�����ԍ��𷰂Ƃ���XSDC1���C�|OSF3����ID���l������
            If GetCOSF3ID(strJDGEIDC, strXTALC1) <> FUNCTION_RETURN_SUCCESS Then
                iErrFlg = 11
            Else
                If Trim(strJDGEIDC) = "" Then
                    iErrFlg = 12
'                '����ID=�u9�v�̏ꍇ�͔���Ȃ�(����OK)�@07/08/01 M.Kaga
'                ElseIf Trim(strJDGEIDC) = "9" Then
'                    YFlg = True
'                    bJudg = True
                Else
                    '�l������C-OSF3����ID��XODC5_OSF30��菳�F�׸ނ̊l��
                    If GetSYNFLAGC5(strSYNFLG, strYMKFLG, strJDGEIDC) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 13
                    Else
                        '���F�׸�:0�@�����F�̏ꍇ
                        If Trim(strSYNFLG) = "0" Or Trim(strSYNFLG) = "" Or IsNull(strSYNFLG) Then
                            iErrFlg = 14
                        End If
                    End If
                End If
            End If
            
            If iErrFlg > 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' ���e
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                        ' ���R
                .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' �T���v���m��
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' �i��(12��)
                If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS           ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo          ' �T���v���m��
                End If
                bJudg = False
            Else
                If YFlg = False Then
                    
                    If GetOsf31_CuDeco(strJDGEIDC, lngSMPPOS, typ_Ret_CuDeco) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 21
                    Else
                        intSiyou = -1
                        intJisseki = -1
                        
                        If (IsNumeric(.typ_si(UpDo).HSXCJPK)) Then
                            intSiyou = CInt(.typ_si(UpDo).HSXCJPK)
                        End If
                        
                        If (IsNumeric(strPtnJsk)) Then
                            intJisseki = CInt(strPtnJsk) + 1
                        End If
                        
                        If (intSiyou >= 1) And (intSiyou <= 4) And (intJisseki >= 1) And (intJisseki <= 4) Then
                        
                           'intRet = funCodeDBGet(SYSCLASS, CLASS, ���i�d�l�p�^�[���敪, 1, ���уp�^�[���敪, �߂�l(tbcmb005.info1))
                            intRet = funCodeDBGet("SB", "S1", CStr(intSiyou), 1, CStr(intJisseki), sResult)
                            If (intRet = 0) And (sResult <> vbNullString) And (Len(sResult) >= 1) Then
                                If sResult = "1" Then
                                    bJudg = True
                                Else
                                    bJudg = False
                                    intNG_Num = 51
                                End If
                            Else
                                'sErr_Msg = sAdd_Msg & sErr_Msg & "���d�l:" & CStr(intSiyou) & ", ����:" & CStr(intJisseki)
                                bJudg = False
                                intNG_Num = 52
'                                GoTo CodeDBGet_Error
                            End If
                            
                        End If
                        
                        ' CJ Ring���a�E�O�a�̔���
                        If bJudg Then
                            If (strPtnJsk = CNST_JSK_PTN_Ring) Or (strPtnJsk = CNST_JSK_PTN_DiskRing) Then
                                If (typ_Ret_CuDeco.CJALLMINRINC5 = -1) Or (typ_Ret_CuDeco.CJALLMINRINC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 61
                                ElseIf (.typ_zi.CuCJ(UpDo).CJRINGNKJSK = -1) Or (.typ_zi.CuCJ(UpDo).CJRINGNKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 62
                                ElseIf (typ_Ret_CuDeco.CJALLMAXRIGC5 = -1) Or (typ_Ret_CuDeco.CJALLMAXRIGC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 63
                                ElseIf (.typ_zi.CuCJ(UpDo).CJRINGGKJSK = -1) Or (.typ_zi.CuCJ(UpDo).CJRINGGKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 64
                                ElseIf (typ_Ret_CuDeco.CJALLMINRINC5 > .typ_zi.CuCJ(UpDo).CJRINGNKJSK) Then
                                    bJudg = False
                                    intNG_Num = 65
                                ElseIf (typ_Ret_CuDeco.CJALLMAXRIGC5 < .typ_zi.CuCJ(UpDo).CJRINGGKJSK) Then
                                    bJudg = False
                                    intNG_Num = 66
                                End If
                            End If
                        End If
                        
                        ' CJ Disk���a�̔���
                        If bJudg Then
                            If (strPtnJsk = CNST_JSK_PTN_Disk) Or (strPtnJsk = CNST_JSK_PTN_DiskRing) Then
                                If (typ_Ret_CuDeco.CJALLMAXDIC5 = -1) Or (typ_Ret_CuDeco.CJALLMAXDIC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 61
                                ElseIf (.typ_zi.CuCJ(UpDo).CJDISKJSK = -1) Or (.typ_zi.CuCJ(UpDo).CJDISKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 62
                                ElseIf (typ_Ret_CuDeco.CJALLMAXDIC5 < .typ_zi.CuCJ(UpDo).CJDISKJSK) Then
                                    bJudg = False
                                    intNG_Num = 71
                                End If
                            End If
                        End If
                        
                        'CJ �v�ZPi���̔���(����l�`�F�b�N)
                        If bJudg Then
                            If (strPtnJsk = CNST_JSK_PTN_Disk) Or (strPtnJsk = CNST_JSK_PTN_Ring) Or (strPtnJsk = CNST_JSK_PTN_DiskRing) Then
                                If (strPtnJsk = CNST_JSK_PTN_Disk) Then
                                    intMax = typ_Ret_CuDeco.CJDMAXPIC5
                                ElseIf (strPtnJsk = CNST_JSK_PTN_Ring) Then
                                    intMax = typ_Ret_CuDeco.CJRMAXPIC5
                                Else
                                    intMax = typ_Ret_CuDeco.CJDRMAXPIC5
                                End If
                                
                                If (intMax = -1) Or (intMax > 150) Then
                                    bJudg = False
                                    intNG_Num = 81
                                ElseIf (.typ_zi.CuCJ(UpDo).CJPICALC = -1) Or (.typ_zi.CuCJ(UpDo).CJPICALC > 150) Then
                                    bJudg = False
                                    intNG_Num = 82
                                ElseIf (intMax < .typ_zi.CuCJ(UpDo).CJPICALC) Then
                                    bJudg = False
                                    intNG_Num = 83
                                End If
                            End If
                        End If
                        
                    End If
                End If
                
                
                If iErrFlg > 0 Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                        ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = CInt(iErrFlg)                    ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' �T���v���m��
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' �i��(12��)
                    If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' �������J�n�ʒu
                        .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' �T���v���m��
                    End If
                    bJudg = False
                    
                Else
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                             ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                             ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' �T���v���m��
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                    
                    If shiji Then
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                        .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                         ' ���R
                        
                        If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                            .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' �������J�n�ʒu
                            .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' �T���v���m��
                            .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                    ' ���R
                            
                            If strSampUmu = "0" Then
                                .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                ' ���R
                                '��ʕ\�����e�ݒ�
                                .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                                .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                                .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                                .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' ���茋��
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' ���茋��
                TotalJudg = False
                gsTbcmy028ErrCode = str028ErrCode
            End If
            DispLineCount = DispLineCount + 1
            
        Else
        ' Add Start 2011/02/15 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : �u�ۏؕ��@_���v�Q�l���̏����ǉ�
            If shiji Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS                              ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                             ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                             ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                            ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo                           ' �T���v���m��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N�Q"                                 ' ���茋��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                    
                    If strSampUmu = "0" Then
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                    End If
                DispLineCount = DispLineCount + 1
            End If
        ' Add End   2011/02/15 SMPK A.Nagamine
        End If      '/*  End of If (JudgSpecCode) Or (pblnFlag) Then */
    End With
    
    CuDecoDataSet_CJ = blnRet
    
End Function
''Add End   2011/01/17 SMPK A.Nagamine

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : Cu-deco(CJ(LT))�̎��є��菈��
Public Function CuDecoDataSet_CJLT(pJudgSC_B() As Judg_Spec_Cry, ptyp_b As typ_AllTypesB, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String, Optional pblnFlag As Boolean = False) As Boolean
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech End
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                  '�����w��
    Dim bJudg           As Boolean
    Dim strXTALC1       As String
    Dim strJDGEIDC      As String
    Dim strSYNFLG       As String
    Dim strYMKFLG       As String
    Dim lngSMPPOS       As Long
    Dim iErrFlg         As Integer
    Dim YFlg            As Boolean
    Dim intRet          As Integer
    Dim sResult         As String
    Dim strSampUmu      As String
    Dim strInfo(4)      As String
    Dim intSiyou        As Integer
    Dim intJisseki      As Integer
    Dim StrCryNum       As String
    Dim lngSampNo       As Long
    Dim strPtnJsk       As String
    Dim str028ErrCode   As String
    Dim blnRet          As Boolean
    Dim typ_Ret_CuDeco  As typ_SB_com_xodb5_osf31_Cudeco
    Dim intNG_Num       As Integer
    
    blnRet = False
    intNG_Num = 0
    
    '�����w���ݒ�
    IND = IIf(UpDo = BlkTop, "123", "123")
    
    With ptyp_b
        ' CJ(LT) ����
        JudgSpecCode = pJudgSC_B(UpDo).CuCJLT
        SCC = "CJLT"
        shiji = (InStr(IND, .typ_cr(UpDo).CRYINDCJLTCS) <> 0)
        strSampUmu = .typ_zi.CuCJLT(UpDo).SMPLUMUCJLT           ' CJ(LT).�T���v���L�� TBCMJ023.SMPLUMUCJLT "0"=�T���v���L��,"1"=�T���v������
        strInfo(0) = CStr(.typ_zi.CuCJLT(UpDo).CJLTPICALC)      ' ���P CJ(LT).Pi���v�Z
        strInfo(1) = CStr(.typ_zi.CuCJLT(UpDo).CJLTBANDNKJSK)   ' ���Q CJ(LT).Band���a����
        strInfo(2) = CStr(.typ_zi.CuCJLT(UpDo).CJLTBANDGKJSK)   ' ���R CJ(LT).Band�O�a����
        strInfo(3) = .typ_zi.CuCJLT(UpDo).CJLTPTNJSK            ' ���S CJ(LT).�p�^�[������
        lngSMPPOS = .typ_zi.CuCJLT(UpDo).POSITION               ' CJ(LT).����وʒu
        StrCryNum = .typ_zi.CuCJLT(UpDo).CRYNUM                 ' CJ(LT).�����ԍ�
        lngSampNo = .typ_zi.CuCJLT(UpDo).SMPLNO                 ' CJ(LT).�T���v����(��\�T���v��ID)
        strPtnJsk = .typ_zi.CuCJLT(UpDo).CJLTPTNJSK             ' CJ(LT).�p�^�[������
        str028ErrCode = "00157"
        
        '�ۏ��׸�="H"�̏ꍇ����V�p�����[�^Optional ��True�w�肳�ꂽ�Ƃ�
        If (JudgSpecCode) Or (pblnFlag) Then
            
            bJudg = False
            
            '�����׸ޏ�����
            iErrFlg = 0
            YFlg = False
            
            '��ۯ�ID
            strXTALC1 = Trim(ptyp_b.BLOCKID)
            '�����ԍ�
            strXTALC1 = left(strXTALC1, 9) & "000"
            
            '�����ԍ��𷰂Ƃ���XSDC1���C�|OSF3����ID���l������
            If GetCOSF3ID(strJDGEIDC, strXTALC1) <> FUNCTION_RETURN_SUCCESS Then
                iErrFlg = 11
            Else
                If Trim(strJDGEIDC) = "" Then
                    iErrFlg = 12
'                '����ID=�u9�v�̏ꍇ�͔���Ȃ�(����OK)�@07/08/01 M.Kaga
'                ElseIf Trim(strJDGEIDC) = "9" Then
'                    YFlg = True
'                    bJudg = True
                Else
                    '�l������C-OSF3����ID��XODC5_OSF30��菳�F�׸ނ̊l��
                    If GetSYNFLAGC5(strSYNFLG, strYMKFLG, strJDGEIDC) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 13
                    Else
                        '���F�׸�:0�@�����F�̏ꍇ
                        If Trim(strSYNFLG) = "0" Or Trim(strSYNFLG) = "" Or IsNull(strSYNFLG) Then
                            iErrFlg = 14
                        End If
                    End If
                End If
            End If
            
            If iErrFlg > 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' ���e
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                        ' ���R
                .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' �T���v���m��
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' �i��(12��)
                If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS           ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo          ' �T���v���m��
                End If
                bJudg = False
            Else
                If YFlg = False Then
                    
                    If GetOsf31_CuDeco(strJDGEIDC, lngSMPPOS, typ_Ret_CuDeco) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 21
                    Else
                        intSiyou = -1
                        intJisseki = -1
                        
                        If (IsNumeric(.typ_si(UpDo).HSXCJLTPK)) Then
                            intSiyou = CInt(.typ_si(UpDo).HSXCJLTPK)
                        End If
                        
                        If (IsNumeric(strPtnJsk)) Then
                            intJisseki = CInt(strPtnJsk) + 1
                        End If
                        
                        If (intSiyou >= 1) And (intSiyou <= 7) And (intJisseki >= 1) And (intJisseki <= 8) Then
                        
                           'intRet = funCodeDBGet(SYSCLASS, CLASS, ���i�d�l�p�^�[���敪, 1, ���уp�^�[���敪, �߂�l(tbcmb005.info1))
                            intRet = funCodeDBGet("SB", "S2", CStr(intSiyou), 1, CStr(intJisseki), sResult)
                            If (intRet = 0) And (sResult <> vbNullString) And (Len(sResult) >= 1) Then
                                If sResult = "1" Then
                                    bJudg = True
                                Else
                                    bJudg = False
                                    intNG_Num = 51
                                End If
                            Else
                                'sErr_Msg = sAdd_Msg & sErr_Msg & "���d�l:" & CStr(intSiyou) & ", ����:" & CStr(intJisseki)
                                bJudg = False
                                intNG_Num = 52
'                                GoTo CodeDBGet_Error
                            End If
                            
                        End If
                        
                        ' CJ(LT) �v�ZBand���̔���
                        If bJudg Then
                            'Cng Start 2012/06/05 Y.Hitomi
                                If (.typ_si(UpDo).HSXCJLTBND = -1) Or (.typ_si(UpDo).HSXCJLTBND > 150) Then
'                            If (strPtnJsk = CNST_JSK_PTN_PBband) Or (strPtnJsk = CNST_JSK_PTN_Pband) Or (strPtnJsk = CNST_JSK_PTN_Bband) Then
'                                If (.typ_si(UpDo).HSXCJLTBND = -1) Or (.typ_si(UpDo).HSXCJLTBND > 150) Then
                            'Cng Start 2012/06/05 Y.Hitomi
                                    bJudg = False
                                    intNG_Num = 61
                                ElseIf (.typ_zi.CuCJLT(UpDo).CJLTBANDGKJSK = -1) Or (.typ_zi.CuCJLT(UpDo).CJLTBANDGKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 62
                                ElseIf (.typ_zi.CuCJLT(UpDo).CJLTBANDNKJSK = -1) Or (.typ_zi.CuCJLT(UpDo).CJLTBANDNKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 63
                                ElseIf (.typ_zi.CuCJLT(UpDo).CJLTBANDGKJSK < .typ_zi.CuCJLT(UpDo).CJLTBANDNKJSK) Then
                                    bJudg = False
                                    intNG_Num = 64
                                ElseIf (.typ_si(UpDo).HSXCJLTBND < (.typ_zi.CuCJLT(UpDo).CJLTBANDGKJSK - .typ_zi.CuCJLT(UpDo).CJLTBANDNKJSK)) Then
                                    bJudg = False
                                    intNG_Num = 65
                                End If
'                            End If
                        End If
                    End If
                End If
                
                If iErrFlg > 0 Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                        ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = CInt(iErrFlg)                    ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' �T���v���m��
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' �i��(12��)
                    If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' �������J�n�ʒu
                        .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' �T���v���m��
                    End If
                    bJudg = False
                    
                Else
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                             ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                             ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' �T���v���m��
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                    
                    If shiji Then
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                        .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                         ' ���R
                        
                        If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                            .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' �������J�n�ʒu
                            .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' �T���v���m��
                            .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                    ' ���R
                            
                            If strSampUmu = "0" Then
                                .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                ' ���R
                                '��ʕ\�����e�ݒ�
                                .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                                .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                                .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                                .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' ���茋��
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' ���茋��
                TotalJudg = False
                gsTbcmy028ErrCode = str028ErrCode
            End If
            DispLineCount = DispLineCount + 1
            
        Else
        ' Add Start 2011/02/15 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : �u�ۏؕ��@_���v�Q�l���̏����ǉ�
            If shiji Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS                              ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                             ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                             ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                            ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo                           ' �T���v���m��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N�Q"                                 ' ���茋��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                    
                    If strSampUmu = "0" Then
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                    End If
                DispLineCount = DispLineCount + 1
            End If
        ' Add End   2011/02/15 SMPK A.Nagamine
        End If      '/*  End of If (JudgSpecCode) Or (pblnFlag) Then */
    End With
    
    CuDecoDataSet_CJLT = blnRet
    
End Function
''Add End   2011/01/17 SMPK A.Nagamine

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : Cu-deco(CJ2)�̎��є��菈��
Public Function CuDecoDataSet_CJ2(pJudgSC_B() As Judg_Spec_Cry, ptyp_b As typ_AllTypesB, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String, Optional pblnFlag As Boolean = False) As Boolean
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech End
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                  '�����w��
    Dim bJudg           As Boolean
    Dim strXTALC1       As String
    Dim strJDGEIDC      As String
    Dim strSYNFLG       As String
    Dim strYMKFLG       As String
    Dim lngSMPPOS       As Long
    Dim iErrFlg         As Integer
    Dim YFlg            As Boolean
    Dim intRet          As Integer
    Dim sResult         As String
    Dim strSampUmu      As String
    Dim strInfo(4)      As String
    Dim intSiyou        As Integer
    Dim intJisseki      As Integer
    Dim StrCryNum       As String
    Dim lngSampNo       As Long
    Dim strPtnJsk       As String
    Dim str028ErrCode   As String
    Dim blnRet          As Boolean
    Dim typ_Ret_CuDeco  As typ_SB_com_xodb5_osf31_Cudeco
    Dim intNG_Num       As Integer
    Dim intMin          As Integer
    
    blnRet = False
    intNG_Num = 0
    
    '�����w���ݒ�
    IND = IIf(UpDo = BlkTop, "123", "123")
    
    With ptyp_b
        ' CJ2 ����
        JudgSpecCode = pJudgSC_B(UpDo).CuCJ2
        SCC = "CJ2"
        shiji = (InStr(IND, .typ_cr(UpDo).CRYINDCJ2CS) <> 0)
        strSampUmu = .typ_zi.CuCJ2(UpDo).SMPLUMUCJ2             ' CJ2.�T���v���L�� TBCMJ023.SMPLUMUCJ2 "0"=�T���v���L��,"1"=�T���v������
        strInfo(0) = CStr(.typ_zi.CuCJ2(UpDo).CJ2DISKJSK)         ' ���P CJ2.Disk���a
        strInfo(1) = CStr(.typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK)       ' ���Q CJ2.Ring���a
        strInfo(2) = CStr(.typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK)       ' ���R CJ2.Ring�O�a
        strInfo(3) = .typ_zi.CuCJ2(UpDo).CJ2PTNJSK              ' ���S CJ2.�p�^�[������
        lngSMPPOS = .typ_zi.CuCJ2(UpDo).POSITION                ' CJ2.����وʒu
        StrCryNum = .typ_zi.CuCJ2(UpDo).CRYNUM                  ' CJ2.�����ԍ�
        lngSampNo = .typ_zi.CuCJ2(UpDo).SMPLNO                  ' CJ2.�T���v����(��\�T���v��ID)
        strPtnJsk = .typ_zi.CuCJ2(UpDo).CJ2PTNJSK               ' CJ2.�p�^�[������
        str028ErrCode = "00158"
        
        '�ۏ��׸�="H"�̏ꍇ����V�p�����[�^Optional ��True�w�肳�ꂽ�Ƃ�
        If (JudgSpecCode) Or (pblnFlag) Then
            
            bJudg = False
            
            '�����׸ޏ�����
            iErrFlg = 0
            YFlg = False
            
            '��ۯ�ID
            strXTALC1 = Trim(ptyp_b.BLOCKID)
            '�����ԍ�
            strXTALC1 = left(strXTALC1, 9) & "000"
            
            '�����ԍ��𷰂Ƃ���XSDC1���C�|OSF3����ID���l������
            If GetCOSF3ID(strJDGEIDC, strXTALC1) <> FUNCTION_RETURN_SUCCESS Then
                iErrFlg = 11
            Else
                If Trim(strJDGEIDC) = "" Then
                    iErrFlg = 12
'                '����ID=�u9�v�̏ꍇ�͔���Ȃ�(����OK)�@07/08/01 M.Kaga
'                ElseIf Trim(strJDGEIDC) = "9" Then
'                    YFlg = True
'                    bJudg = True
                Else
                    '�l������C-OSF3����ID��XODC5_OSF30��菳�F�׸ނ̊l��
                    If GetSYNFLAGC5(strSYNFLG, strYMKFLG, strJDGEIDC) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 13
                    Else
                        '���F�׸�:0�@�����F�̏ꍇ
                        If Trim(strSYNFLG) = "0" Or Trim(strSYNFLG) = "" Or IsNull(strSYNFLG) Then
                            iErrFlg = 14
                        End If
                    End If
                End If
            End If
            
            If iErrFlg > 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' ���e
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                        ' ���R
                .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' �T���v���m��
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' �i��(12��)
                If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS           ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo          ' �T���v���m��
                End If
                bJudg = False
            Else
                If YFlg = False Then
                    
                    If GetOsf31_CuDeco(strJDGEIDC, lngSMPPOS, typ_Ret_CuDeco) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 21
                    Else
                        intSiyou = -1
                        intJisseki = -1
                        
                        If (IsNumeric(.typ_si(UpDo).HSXCJ2PK)) Then
                            intSiyou = CInt(.typ_si(UpDo).HSXCJ2PK)
                        End If
                        
                        If (IsNumeric(strPtnJsk)) Then
                            intJisseki = CInt(strPtnJsk) + 1
                        End If
                        
                        If (intSiyou >= 1) And (intSiyou <= 4) And (intJisseki >= 1) And (intJisseki <= 4) Then
                        
                           'intRet = funCodeDBGet(SYSCLASS, CLASS, ���i�d�l�p�^�[���敪, 1, ���уp�^�[���敪, �߂�l(tbcmb005.info1))
                            intRet = funCodeDBGet("SB", "S1", CStr(intSiyou), 1, CStr(intJisseki), sResult)
                            If (intRet = 0) And (sResult <> vbNullString) And (Len(sResult) >= 1) Then
                                If sResult = "1" Then
                                    bJudg = True
                                Else
                                    bJudg = False
                                    intNG_Num = 51
                                End If
                            Else
                                'sErr_Msg = sAdd_Msg & sErr_Msg & "���d�l:" & CStr(intSiyou) & ", ����:" & CStr(intJisseki)
                                bJudg = False
                                intNG_Num = 52
'                                GoTo CodeDBGet_Error
                            End If
                            
                        End If
                        
                        ' CJ2 Ring���a�E�O�a�̔���
                        If bJudg Then
                            If (strPtnJsk = CNST_JSK_PTN_Ring) Then
                                If (typ_Ret_CuDeco.CJ2RMINRINC5 = -1) Or (typ_Ret_CuDeco.CJ2RMINRINC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 61
                                ElseIf (.typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK = -1) Or (.typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 62
                                ElseIf (typ_Ret_CuDeco.CJ2RMAXRIGC5 = -1) Or (typ_Ret_CuDeco.CJ2RMAXRIGC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 63
                                ElseIf (.typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK = -1) Or (.typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 64
                                ElseIf (typ_Ret_CuDeco.CJ2RMINRINC5 > .typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK) Then
                                    bJudg = False
                                    intNG_Num = 65
                                ElseIf (typ_Ret_CuDeco.CJ2RMAXRIGC5 < .typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK) Then
                                    bJudg = False
                                    intNG_Num = 66
                                End If
                            ElseIf (strPtnJsk = CNST_JSK_PTN_DiskRing) Then
                                If (typ_Ret_CuDeco.CJ2DRMINRINC5 = -1) Or (typ_Ret_CuDeco.CJ2DRMINRINC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 71
                                ElseIf (.typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK = -1) Or (.typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 72
                                ElseIf (typ_Ret_CuDeco.CJ2DRMAXRIGC5 = -1) Or (typ_Ret_CuDeco.CJ2DRMAXRIGC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 73
                                ElseIf (.typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK = -1) Or (.typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 74
                                ElseIf (typ_Ret_CuDeco.CJ2DRMINRINC5 > .typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK) Then
                                    bJudg = False
                                    intNG_Num = 75
                                ElseIf (typ_Ret_CuDeco.CJ2DRMAXRIGC5 < .typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK) Then
                                    bJudg = False
                                    intNG_Num = 76
                                End If
                            End If
                        End If
                        
                        'CJ2 �v�ZPi���̔���(�����l�`�F�b�N)
                        If bJudg Then
                            If (strPtnJsk = CNST_JSK_PTN_Disk) Or (strPtnJsk = CNST_JSK_PTN_Ring) Or (strPtnJsk = CNST_JSK_PTN_DiskRing) Then
                                If (strPtnJsk = CNST_JSK_PTN_Disk) Then
                                    intMin = typ_Ret_CuDeco.CJ2DMAXPIC5
                                ElseIf (strPtnJsk = CNST_JSK_PTN_Ring) Then
                                    intMin = typ_Ret_CuDeco.CJ2RMAXPIC5
                                Else
                                    intMin = typ_Ret_CuDeco.CJ2DRMAXPIC5
                                End If
                                
                                If (intMin = -1) Or (intMin > 150) Then
                                    bJudg = False
                                    intNG_Num = 81
                                ElseIf (.typ_zi.CuCJ2(UpDo).CJ2PICALC = -1) Or (.typ_zi.CuCJ2(UpDo).CJ2PICALC > 150) Then
                                    bJudg = False
                                    intNG_Num = 82
                                ElseIf (intMin > .typ_zi.CuCJ2(UpDo).CJ2PICALC) Then
                                    bJudg = False
                                    intNG_Num = 83
                                End If
                            End If
                        End If
                        
                    End If
                End If
                
                If iErrFlg > 0 Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                        ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = CInt(iErrFlg)                    ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' �T���v���m��
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' �i��(12��)
                    If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' �������J�n�ʒu
                        .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' �T���v���m��
                    End If
                    bJudg = False
                    
                Else
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                             ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                             ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' �T���v���m��
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                    
                    If shiji Then
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                         ' ���P
                        .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                         ' ���Q
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                         ' ���R
                        
                        If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                            .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' �������J�n�ʒu
                            .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' �T���v���m��
                            .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                    ' ���R
                            
                            If strSampUmu = "0" Then
                                .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                ' ���R
                                '��ʕ\�����e�ݒ�
                                .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                                .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                                .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                                .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' ���茋��
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' ���茋��
                TotalJudg = False
                gsTbcmy028ErrCode = str028ErrCode
            End If
            DispLineCount = DispLineCount + 1
            
        Else
        ' Add Start 2011/02/15 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : �u�ۏؕ��@_���v�Q�l���̏����ǉ�
            If shiji Then
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS                              ' �������J�n�ʒu
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' ���e
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                             ' ���P
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                             ' ���Q
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����ٖ�"                            ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo                           ' �T���v���m��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N�Q"                                 ' ���茋��
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' �i��(12��)
                    
                    If strSampUmu = "0" Then
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                    End If
                DispLineCount = DispLineCount + 1
            End If
        ' Add End   2011/02/15 SMPK A.Nagamine
        End If      '/*  End of If (JudgSpecCode) Or (pblnFlag) Then */
    End With
    
    CuDecoDataSet_CJ2 = blnRet
    
End Function
''Add End   2011/01/17 SMPK A.Nagamine

