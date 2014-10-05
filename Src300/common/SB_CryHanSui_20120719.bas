Attribute VB_Name = "SB_CryHanSui"
Option Explicit

'-------------------------------------------------------------------------------
' �萔��`
'-------------------------------------------------------------------------------
'XSDCS
Private Const cCRYSMPLID    As String = "CRYSMPLID"     'XSDCS�̃T���v���h�c
Private Const cCRYIND       As String = "CRYIND"        'XSDCS�̏��FLG
Private Const cCRYRES       As String = "CRYRES"        'XSDCS�̎���FLG
Private Const cCS           As String = "CS"            'XSDCS�̍��ڍŏI����
Private Const cCRY_RS       As String = "RS"            'XSDCS��Rs
Private Const cCRY_OI       As String = "OI"            'XSDCS��Oi
Private Const cCRY_B1       As String = "B1"            'XSDCS��BMD1
Private Const cCRY_B2       As String = "B2"            'XSDCS��BMD2
Private Const cCRY_B3       As String = "B3"            'XSDCS��BMD3
Private Const cCRY_O1       As String = "L1"            'XSDCS��OSF1
Private Const cCRY_O2       As String = "L2"            'XSDCS��OSF2
Private Const cCRY_O3       As String = "L3"            'XSDCS��OSF3
Private Const cCRY_O4       As String = "L4"            'XSDCS��OSF4
Private Const cCRY_CS       As String = "CS"            'XSDCS��Cs
Private Const cCRY_GD       As String = "GD"            'XSDCS��GD
Private Const cCRY_LT       As String = "T"             'XSDCS��LT
Private Const cCRY_EP       As String = "EP"            'XSDCS��EPD
'Add Start 2011/01/19 SMPK Miyata
Private Const cCRY_C        As String = "C"             'XSDCS��C
Private Const cCRY_CJ       As String = "CJ"            'XSDCS��CJ
Private Const cCRY_CJLT     As String = "CJLT"          'XSDCS��CJLT
Private Const cCRY_CJ2      As String = "CJ2"           'XSDCS��CJ2
'Add End   2011/01/19 SMPK Miyata

'������R����
Public Type type_DBDRV_scmzc_fcmkc001c_CryR
    CRYNUM      As String * 12          ' �����ԍ�
    POSITION    As Integer              ' �ʒu
    SMPKBN      As String * 1           ' �T���v���敪
    TRANCOND    As String * 1           ' ��������
    TRANCNT     As String * 1           ' ������
    SMPLNO      As Long                 ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' �T���v���L��
    MEAS1       As Double               ' ����l�P
    MEAS2       As Double               ' ����l�Q
    MEAS3       As Double               ' ����l�R
    MEAS4       As Double               ' ����l�S
    MEAS5       As Double               ' ����l�T
    EFEHS       As Double               ' �����ΐ�
    RRG         As Double               ' �q�q�f
    REGDATE     As Date                 ' �o�^���t
    '-----TEST2004/10
    JMEAS1       As Double               ' ����l�P
    JMEAS2       As Double               ' ����l�Q
    JMEAS3       As Double               ' ����l�R
    JMEAS4       As Double               ' ����l�S
    JMEAS5       As Double               ' ����l�T
    KSTAFFID     As String
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP     As String * 1          'DK���x(����)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
End Type

'Oi����
Public Type type_DBDRV_scmzc_fcmkc001c_Oi
    CRYNUM      As String * 12          ' �����ԍ�
    POSITION    As Integer              ' �ʒu
    SMPKBN      As String * 1           ' �T���v���敪
    TRANCOND    As String * 1           ' ��������
    TRANCNT     As String * 1           ' ������
    SMPLNO      As Long                 ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' �T���v���L��
    OIMEAS1     As Double               ' �n������l�P
    OIMEAS2     As Double               ' �n������l�Q
    OIMEAS3     As Double               ' �n������l�R
    OIMEAS4     As Double               ' �n������l�S
    OIMEAS5     As Double               ' �n������l�T
    ORGRES      As Double               ' �n�q�f����
    AVE         As Double               ' �`�u�d
    FTIRCONV    As Double               ' �e�s�h�q���Z
    INSPECTWAY  As String * 2           ' �������@
    REGDATE     As Date                 ' �o�^���t
End Type

'BMD1�`3����
Public Type type_DBDRV_scmzc_fcmkc001c_BMD
    CRYNUM      As String * 12          ' �����ԍ�
    POSITION    As Integer              ' �ʒu
    SMPKBN      As String * 1           ' �T���v���敪
    TRANCOND    As String * 1           ' ��������
    TRANCNT     As String * 1           ' ������
    SMPLNO      As Long                 ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' �T���v���L��
    HTPRC       As String * 2           ' �M�������@
    KKSP        As String * 3           ' �������ב���ʒu
    KKSET       As String * 3           ' �������ב�������{�I��ET��@�@char(1)�{number(2)
    MEAS1       As Double               ' ����l�P
    MEAS2       As Double               ' ����l�Q
    MEAS3       As Double               ' ����l�R
    MEAS4       As Double               ' ����l�S
    MEAS5       As Double               ' ����l�T
    MEASMIN     As Double               ' MIN
    MEASMAX     As Double               ' MAX
    MEASAVE     As Double               ' AVE
    BMDMNBUNP   As Double               ' BMD�ʓ����z
    REGDATE     As Date                 ' �o�^���t
End Type

'OSF1�`4����
Public Type type_DBDRV_scmzc_fcmkc001c_OSF
    CRYNUM      As String * 12          ' �����ԍ�
    POSITION    As Integer              ' �ʒu
    SMPKBN      As String * 1           ' �T���v���敪
    TRANCOND    As String * 1           ' ��������
    TRANCNT     As Integer              ' ������      String * 1 -> Integer 2008/10/28 L/DL,OSF����ۼޯ��ǉ�(IT) UPD By Systech
    SMPLNO      As Long                 ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' �T���v���L��
    HTPRC       As String * 2           ' �M�������@
    KKSP        As String * 3           ' �������ב���ʒu
    KKSET       As String * 3           ' �������ב�������{�I��ET��@�@char(1)�{number(2)
    CALCMAX     As Double               ' �v�Z���� Max
    CALCAVE     As Double               ' �v�Z���� Ave
    MEAS1       As Double               ' ����l�P
    MEAS2       As Double               ' ����l�Q
    MEAS3       As Double               ' ����l�R
    MEAS4       As Double               ' ����l�S
    MEAS5       As Double               ' ����l�T
    MEAS6       As Double               ' ����l�U
    MEAS7       As Double               ' ����l�V
    MEAS8       As Double               ' ����l�W
    MEAS9       As Double               ' ����l�X
    MEAS10      As Double               ' ����l�P�O
    MEAS11      As Double               ' ����l�P�P
    MEAS12      As Double               ' ����l�P�Q
    MEAS13      As Double               ' ����l�P�R
    MEAS14      As Double               ' ����l�P�S
    MEAS15      As Double               ' ����l�P�T
    MEAS16      As Double               ' ����l�P�U
    MEAS17      As Double               ' ����l�P�V
    MEAS18      As Double               ' ����l�P�W
    MEAS19      As Double               ' ����l�P�X
    MEAS20      As Double               ' ����l�Q�O
    OSFPOS1     As Double               ' ����݋敪�P�ʒu
    OSFWID1     As Double               ' ����݋敪�P��
    OSFRD1      As String * 1           ' ����݋敪�PR/D
    OSFPOS2     As Double               ' ����݋敪�Q�ʒu
    OSFWID2     As Double               ' ����݋敪�Q��
    OSFRD2      As String * 1           ' ����݋敪�QR/D
    OSFPOS3     As Double               ' ����݋敪�R�ʒu
    OSFWID3     As Double               ' ����݋敪�R��
    OSFRD3      As String * 1           ' ����݋敪�RR/D
    CALCMH      As Double               ' �ʓ���(MAX/MIN)   '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    PTNJUDGRES  As String               ' �p�^�[�����茋��   '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    REGDATE     As Date                 ' �o�^���t
End Type

'CS����
Public Type type_DBDRV_scmzc_fcmkc001c_CS
    CRYNUM      As String * 12          ' �����ԍ�
    POSITION    As Integer              ' �ʒu
    SMPKBN      As String * 1           ' �T���v���敪
    TRANCOND    As String * 1           ' ��������
    TRANCNT     As String * 1           ' ������
    SMPLNO      As Long                 ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' �T���v���L��
    CSMEAS      As Double               ' Cs�����l
    PRE70P      As Double               ' �V�O������l
    INSPECTWAY  As String * 2           ' �������@
    REGDATE     As Date                 ' �o�^���t
End Type

'GD����
Public Type type_DBDRV_scmzc_fcmkc001c_GD
    CRYNUM      As String * 12          ' �����ԍ�
    POSITION    As Integer              ' �ʒu
    SMPKBN      As String * 1           ' �T���v���敪
    TRANCOND    As String * 1           ' ��������
    TRANCNT     As Integer              ' ������      String * 1 -> Integer 2008/10/28 L/DL,OSF����ۼޯ��ǉ�(IT) UPD By Systech
    SMPLNO      As Long                 ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' �T���v���L��
    MSRSDEN     As Integer              ' ���茋�� Den
    MSRSLDL     As Integer              ' ���茋�� L/DL
    MSRSDVD2    As Integer              ' ���茋�� DVD2
    MS01LDL1    As Integer              ' ����l01 L/DL1
    MS01LDL2    As Integer              ' ����l01 L/DL2
    MS01LDL3    As Integer              ' ����l01 L/DL3
    MS01LDL4    As Integer              ' ����l01 L/DL4
    MS01LDL5    As Integer              ' ����l01 L/DL5
    MS01DEN1    As Integer              ' ����l01 Den1
    MS01DEN2    As Integer              ' ����l01 Den2
    MS01DEN3    As Integer              ' ����l01 Den3
    MS01DEN4    As Integer              ' ����l01 Den4
    MS01DEN5    As Integer              ' ����l01 Den5
    MS02LDL1    As Integer              ' ����l02 L/DL1
    MS02LDL2    As Integer              ' ����l02 L/DL2
    MS02LDL3    As Integer              ' ����l02 L/DL3
    MS02LDL4    As Integer              ' ����l02 L/DL4
    MS02LDL5    As Integer              ' ����l02 L/DL5
    MS02DEN1    As Integer              ' ����l02 Den1
    MS02DEN2    As Integer              ' ����l02 Den2
    MS02DEN3    As Integer              ' ����l02 Den3
    MS02DEN4    As Integer              ' ����l02 Den4
    MS02DEN5    As Integer              ' ����l02 Den5
    MS03LDL1    As Integer              ' ����l03 L/DL1
    MS03LDL2    As Integer              ' ����l03 L/DL2
    MS03LDL3    As Integer              ' ����l03 L/DL3
    MS03LDL4    As Integer              ' ����l03 L/DL4
    MS03LDL5    As Integer              ' ����l03 L/DL5
    MS03DEN1    As Integer              ' ����l03 Den1
    MS03DEN2    As Integer              ' ����l03 Den2
    MS03DEN3    As Integer              ' ����l03 Den3
    MS03DEN4    As Integer              ' ����l03 Den4
    MS03DEN5    As Integer              ' ����l03 Den5
    MS04LDL1    As Integer              ' ����l04 L/DL1
    MS04LDL2    As Integer              ' ����l04 L/DL2
    MS04LDL3    As Integer              ' ����l04 L/DL3
    MS04LDL4    As Integer              ' ����l04 L/DL4
    MS04LDL5    As Integer              ' ����l04 L/DL5
    MS04DEN1    As Integer              ' ����l04 Den1
    MS04DEN2    As Integer              ' ����l04 Den2
    MS04DEN3    As Integer              ' ����l04 Den3
    MS04DEN4    As Integer              ' ����l04 Den4
    MS04DEN5    As Integer              ' ����l04 Den5
    MS05LDL1    As Integer              ' ����l05 L/DL1
    MS05LDL2    As Integer              ' ����l05 L/DL2
    MS05LDL3    As Integer              ' ����l05 L/DL3
    MS05LDL4    As Integer              ' ����l05 L/DL4
    MS05LDL5    As Integer              ' ����l05 L/DL5
    MS05DEN1    As Integer              ' ����l05 Den1
    MS05DEN2    As Integer              ' ����l05 Den2
    MS05DEN3    As Integer              ' ����l05 Den3
    MS05DEN4    As Integer              ' ����l05 Den4
    MS05DEN5    As Integer              ' ����l05 Den5
    MS06LDL1    As Integer              ' ����l06 L/DL1
    MS06LDL2    As Integer              ' ����l06 L/DL2
    MS06LDL3    As Integer              ' ����l06 L/DL3
    MS06LDL4    As Integer              ' ����l06 L/DL4
    MS06LDL5    As Integer              ' ����l06 L/DL5
    MS06DEN1    As Integer              ' ����l06 Den1
    MS06DEN2    As Integer              ' ����l06 Den2
    MS06DEN3    As Integer              ' ����l06 Den3
    MS06DEN4    As Integer              ' ����l06 Den4
    MS06DEN5    As Integer              ' ����l06 Den5
    MS07LDL1    As Integer              ' ����l07 L/DL1
    MS07LDL2    As Integer              ' ����l07 L/DL2
    MS07LDL3    As Integer              ' ����l07 L/DL3
    MS07LDL4    As Integer              ' ����l07 L/DL4
    MS07LDL5    As Integer              ' ����l07 L/DL5
    MS07DEN1    As Integer              ' ����l07 Den1
    MS07DEN2    As Integer              ' ����l07 Den2
    MS07DEN3    As Integer              ' ����l07 Den3
    MS07DEN4    As Integer              ' ����l07 Den4
    MS07DEN5    As Integer              ' ����l07 Den5
    MS08LDL1    As Integer              ' ����l08 L/DL1
    MS08LDL2    As Integer              ' ����l08 L/DL2
    MS08LDL3    As Integer              ' ����l08 L/DL3
    MS08LDL4    As Integer              ' ����l08 L/DL4
    MS08LDL5    As Integer              ' ����l08 L/DL5
    MS08DEN1    As Integer              ' ����l08 Den1
    MS08DEN2    As Integer              ' ����l08 Den2
    MS08DEN3    As Integer              ' ����l08 Den3
    MS08DEN4    As Integer              ' ����l08 Den4
    MS08DEN5    As Integer              ' ����l08 Den5
    MS09LDL1    As Integer              ' ����l09 L/DL1
    MS09LDL2    As Integer              ' ����l09 L/DL2
    MS09LDL3    As Integer              ' ����l09 L/DL3
    MS09LDL4    As Integer              ' ����l09 L/DL4
    MS09LDL5    As Integer              ' ����l09 L/DL5
    MS09DEN1    As Integer              ' ����l09 Den1
    MS09DEN2    As Integer              ' ����l09 Den2
    MS09DEN3    As Integer              ' ����l09 Den3
    MS09DEN4    As Integer              ' ����l09 Den4
    MS09DEN5    As Integer              ' ����l09 Den5
    MS10LDL1    As Integer              ' ����l10 L/DL1
    MS10LDL2    As Integer              ' ����l10 L/DL2
    MS10LDL3    As Integer              ' ����l10 L/DL3
    MS10LDL4    As Integer              ' ����l10 L/DL4
    MS10LDL5    As Integer              ' ����l10 L/DL5
    MS10DEN1    As Integer              ' ����l10 Den1
    MS10DEN2    As Integer              ' ����l10 Den2
    MS10DEN3    As Integer              ' ����l10 Den3
    MS10DEN4    As Integer              ' ����l10 Den4
    MS10DEN5    As Integer              ' ����l10 Den5
    MS11LDL1    As Integer              ' ����l11 L/DL1
    MS11LDL2    As Integer              ' ����l11 L/DL2
    MS11LDL3    As Integer              ' ����l11 L/DL3
    MS11LDL4    As Integer              ' ����l11 L/DL4
    MS11LDL5    As Integer              ' ����l11 L/DL5
    MS11DEN1    As Integer              ' ����l11 Den1
    MS11DEN2    As Integer              ' ����l11 Den2
    MS11DEN3    As Integer              ' ����l11 Den3
    MS11DEN4    As Integer              ' ����l11 Den4
    MS11DEN5    As Integer              ' ����l11 Den5
    MS12LDL1    As Integer              ' ����l12 L/DL1
    MS12LDL2    As Integer              ' ����l12 L/DL2
    MS12LDL3    As Integer              ' ����l12 L/DL3
    MS12LDL4    As Integer              ' ����l12 L/DL4
    MS12LDL5    As Integer              ' ����l12 L/DL5
    MS12DEN1    As Integer              ' ����l12 Den1
    MS12DEN2    As Integer              ' ����l12 Den2
    MS12DEN3    As Integer              ' ����l12 Den3
    MS12DEN4    As Integer              ' ����l12 Den4
    MS12DEN5    As Integer              ' ����l12 Den5
    MS13LDL1    As Integer              ' ����l13 L/DL1
    MS13LDL2    As Integer              ' ����l13 L/DL2
    MS13LDL3    As Integer              ' ����l13 L/DL3
    MS13LDL4    As Integer              ' ����l13 L/DL4
    MS13LDL5    As Integer              ' ����l13 L/DL5
    MS13DEN1    As Integer              ' ����l13 Den1
    MS13DEN2    As Integer              ' ����l13 Den2
    MS13DEN3    As Integer              ' ����l13 Den3
    MS13DEN4    As Integer              ' ����l13 Den4
    MS13DEN5    As Integer              ' ����l13 Den5
    MS14LDL1    As Integer              ' ����l14 L/DL1
    MS14LDL2    As Integer              ' ����l14 L/DL2
    MS14LDL3    As Integer              ' ����l14 L/DL3
    MS14LDL4    As Integer              ' ����l14 L/DL4
    MS14LDL5    As Integer              ' ����l14 L/DL5
    MS14DEN1    As Integer              ' ����l14 Den1
    MS14DEN2    As Integer              ' ����l14 Den2
    MS14DEN3    As Integer              ' ����l14 Den3
    MS14DEN4    As Integer              ' ����l14 Den4
    MS14DEN5    As Integer              ' ����l14 Den5
    MS15LDL1    As Integer              ' ����l15 L/DL1
    MS15LDL2    As Integer              ' ����l15 L/DL2
    MS15LDL3    As Integer              ' ����l15 L/DL3
    MS15LDL4    As Integer              ' ����l15 L/DL4
    MS15LDL5    As Integer              ' ����l15 L/DL5
    MS15DEN1    As Integer              ' ����l15 Den1
    MS15DEN2    As Integer              ' ����l15 Den2
    MS15DEN3    As Integer              ' ����l15 Den3
    MS15DEN4    As Integer              ' ����l15 Den4
    MS15DEN5    As Integer              ' ����l15 Den5
    MS01DVD2    As Integer              ' ����l01 DVD2
    MS02DVD2    As Integer              ' ����l02 DVD2
    MS03DVD2    As Integer              ' ����l03 DVD2
    MS04DVD2    As Integer              ' ����l04 DVD2
    MS05DVD2    As Integer              ' ����l05 DVD2
    MSZEROMN    As Integer              ' L/DL0�A�����ŏ��l '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    MSZEROMX    As Integer              ' L/DL0�A�����ő�l '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    PTNJUDGRES  As String               ' �p�^�[�����茋��   '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    REGDATE     As Date                 ' �o�^���t
End Type

'���C�t�^�C�����ю擾�֐�
Public Type type_DBDRV_scmzc_fcmkc001c_LT
    CRYNUM      As String * 12          ' �����ԍ�
    POSITION    As Integer              ' �ʒu
    SMPKBN      As String * 1           ' �T���v���敪
    TRANCOND    As String * 1           ' ��������
    TRANCNT     As String * 1           ' ������
    SMPLNO      As Long                 ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' �T���v���L��
    MEAS1       As Integer              ' ����l�P
    MEAS2       As Integer              ' ����l�Q
    MEAS3       As Integer              ' ����l�R
    MEAS4       As Integer              ' ����l�S
    MEAS5       As Integer              ' ����l�T
    MEASPEAK    As Integer              ' ����l �s�[�N�l
    CALCMEAS    As Integer              ' �v�Z����
    REGDATE     As Date                 ' �o�^���t
    LTSPI       As String               ' ����ʒu�R�[�h
'2005/12/02 add SET���� ����l6�`10�ǉ��A����t���O->
    MEAS6       As Integer              ' ����l�U
    MEAS7       As Integer              ' ����l�V
    MEAS8       As Integer              ' ����l�W
    MEAS9       As Integer              ' ����l�X
    MEAS10      As Integer              ' ����l�P�O
    LTSPIFLG    As String               ' ����t���O
'2005/12/02 add SET���� ����l6�`10�ǉ��A����t���O<-
''Add Start 2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
    CONVAL      As Integer               ' LT10�����Z
''Add End   2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)

End Type


'EPD���ю擾�֐�
Public Type type_DBDRV_scmzc_fcmkc001c_EPD
    CRYNUM      As String * 12          ' �����ԍ�
    POSITION    As Integer              ' �ʒu
    SMPKBN      As String * 1           ' �T���v���敪
    TRANCOND    As String * 1           ' ��������
    TRANCNT     As String * 1           ' ������
    SMPLNO      As Long                 ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' �T���v���L��
    MEASURE     As Integer              ' ����l
    REGDATE     As Date                 ' �o�^���t
End Type
'X�����ю擾�֐�   2009/08/12 add Kameda
Public Type type_DBDRV_scmzc_fcmkc001c_X
    CRYNUM      As String * 12          ' �����ԍ�
    POSITION    As Integer              ' �ʒu
    SMPKBN      As String * 1           ' �T���v���敪
    TRANCOND    As String * 1           ' ��������
    TRANCNT     As String * 1           ' ������
    SMPLNO      As Long                 ' �T���v���m��
    SMPLUMU     As String * 1           ' �T���v���L��
    XX          As Double               ' ����lX
    XY          As Double               ' ����lY
    XXY         As Double               ' ����l����
    REGDATE     As Date                 ' �o�^���t
    'JUDG        As String              ' ���荀�ڒǉ�   2009/10/22 Kameda
    JUDGXY       As String              ' ���荀�ڒǉ�   2009/10/22 Kameda
    JUDGX        As String              ' ���荀�ڒǉ�   2009/10/22 Kameda
    JUDGY        As String              ' ���荀�ڒǉ�   2009/10/22 Kameda
End Type
'SIRD���ю擾�֐�   2010/02/04 add Kameda
Public Type type_DBDRV_scmzc_fcmkc001c_SIRD
    CRYNUM      As String * 12          ' �����ԍ�
    POSITION    As Integer              ' �ʒu
    SMPKBN      As String * 1           ' �T���v���敪
    TRANCOND    As String * 1           ' ��������
    TRANCNT     As String * 1           ' ������
    'SMPLNO      As Long                 ' �T���v���m��
    SMPLNO      As String                ' �T���v���m��
    SMPLUMU     As String * 1           ' �T���v���L��
    SIRDCNT     As Double               ' �]������
    REGDATE     As Date                 ' �o�^���t
    NothingFlg  As String               ' �f�[�^�Ȃ��t���O   2010/02/18 Kameda
End Type

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ�
Public Type type_DBDRV_scmzc_fcmkc001c_C
    CRYNUM          As String * 12      ' �����ԍ�
    POSITION        As Integer          ' �ʒu
    SMPKBN          As String * 1       ' �T���v���敪
    TRANCNT         As Integer          ' ������
'    TRANCOND        As String * 1       '��������
    SMPLNO          As Long             ' �T���v���m��
    SMPLUMUC        As String * 1       ' �T���v���L���iC�j
    
    CPTNJSK         As String * 1       ' C �p�^�[������
    CDISKJSK        As Integer          ' C Disk���a����
    CRINGNKJSK      As Integer          ' C Ring���a����
    CRINGGKJSK      As Integer          ' C Ring�O�a����
    CHANTEI         As String * 1       ' C ���茋��
    REGDATE         As Date             ' C �o�^���t
End Type

Public Type type_DBDRV_scmzc_fcmkc001c_CJ
    CRYNUM          As String * 12      ' �����ԍ�
    POSITION        As Integer          ' �ʒu
    SMPKBN          As String * 1       ' �T���v���敪
    TRANCNT         As Integer          ' ������
'    TRANCOND        As String * 1       '��������
    SMPLNO          As Long             ' �T���v���m��
    SMPLUMUCJ       As String * 1       ' �T���v���L���iCJ�j
    
    CJPTNJSK        As String * 1       ' CJ �p�^�[������
    CJDISKJSK       As Integer          ' CJ Disk���a����
    CJRINGNKJSK     As Integer          ' CJ Ring���a����
    CJRINGGKJSK     As Integer          ' CJ Ring�O�a����
    CJBANDNKJSK     As Integer          ' CJ Band���a����
    CJBANDGKJSK     As Integer          ' CJ Band�O�a����
    CJRINGCALC      As Integer          ' CJ Ring���v�Z
    CJPICALC        As Integer          ' CJ Pi���v�Z
    CJHANTEI        As String * 1       ' CJ ���茋��
'    CJBUIUMU        As String * 1       ' CJ ���ʕʔ���L��
    CJDMAXPIC5      As Integer          ' CJ Disk�̂݃p�^�[�� Pi������l
    CJRMAXPIC5      As Integer          ' CJ Ring�̂݃p�^�[�� Pi������l
    CJDRMAXPIC5     As Integer          ' CJ DiskRing�p�^�[�� Pi������l
    CJALLMAXDIC5    As Integer          ' CJ ����Disk���a����l
    CJALLMINRINC5   As Integer          ' CJ ����Ring���a�����l
    CJALLMAXRIGC5   As Integer          ' CJ ����Ring�O�a����l
    REGDATE         As Date             ' CJ �o�^���t
End Type

Public Type type_DBDRV_scmzc_fcmkc001c_CJLT
    CRYNUM          As String * 12      ' �����ԍ�
    POSITION        As Integer          ' �ʒu
    SMPKBN          As String * 1       ' �T���v���敪
    TRANCNT         As Integer          ' ������
'    TRANCOND        As String * 1       '��������
    SMPLNO          As Long             ' �T���v���m��
    SMPLUMUCJLT     As String * 1       ' �T���v���L���iCJ(LT)�j
    
    CJLTPTNJSK      As String * 1       ' CJ(LT) �p�^�[������
    CJLTDISKJSK     As Integer          ' CJ(LT) Disk���a����
    CJLTRINGNKJSK   As Integer          ' CJ(LT) Ring���a����
    CJLTRINGGKJSK   As Integer          ' CJ(LT) Ring�O�a����
    CJLTBANDNKJSK   As Integer          ' CJ(LT) Band���a����
    CJLTBANDGKJSK   As Integer          ' CJ(LT) Band�O�a����
    CJLTRINGCALC    As Integer          ' CJ(LT) Ring���v�Z
    CJLTPICALC      As Integer          ' CJ(LT) Pi���v�Z
    CJLTBANDCALC    As Integer          ' CJ(LT) Band���v�Z
    CJLTHANTEI      As String * 1       ' CJ(LT) ���茋��
    HSXCJLTBND      As Integer          ' CJ(LT) Band������l
    REGDATE         As Date             ' CJ(LT) �o�^���t
End Type

Public Type type_DBDRV_scmzc_fcmkc001c_CJ2
    CRYNUM          As String * 12      ' �����ԍ�
    POSITION        As Integer          ' �ʒu
    SMPKBN          As String * 1       ' �T���v���敪
    TRANCNT         As Integer          ' ������
'    TRANCOND        As String * 1       '��������
    SMPLNO          As Long             ' �T���v���m��
    SMPLUMUCJ2      As String * 1       ' �T���v���L���iCJ2�j
    
    CJ2PTNJSK       As String * 1       ' CJ2 �p�^�[������
    CJ2DISKJSK      As Integer          ' CJ2 Disk���a����
    CJ2RINGNKJSK    As Integer          ' CJ2 Ring���a����
    CJ2RINGGKJSK    As Integer          ' CJ2 Ring�O�a����
    CJ2PICALC       As Integer          ' CJ2 Pi���v�Z
    CJ2HANTEI       As String * 1       ' CJ2 ���茋��
'    CJ2BUIUMU       As String * 1       ' CJ2 ���ʕʔ���L��
    CJ2DMAXPIC5     As Integer          ' CJ2 Disk�̂݃p�^�[�� Pi�������l(MAX���������ł�)
    CJ2RMAXPIC5     As Integer          ' CJ2 Ring�̂݃p�^�[�� Pi�������l(MAX���������ł�)
    CJ2RMINRINC5    As Integer          ' CJ2 Ring�̂݃p�^�[�� Ring���a�����l
    CJ2RMAXRIGC5    As Integer          ' CJ2 Ring�̂݃p�^�[�� Ring�O�a����l
    CJ2DRMAXPIC5    As Integer          ' CJ2 DiskRing�p�^�[�� Pi�������l(MAX���������ł�)
    CJ2DRMINRINC5   As Integer          ' CJ2 DiskRing�p�^�[�� Ring���a�����l
    CJ2DRMAXRIGC5   As Integer          ' CJ2 DiskRing�p�^�[�� Ring�O�a����l
    REGDATE         As Date             ' CJ2 �o�^���t
End Type
'Add End   2011/01/17 SMPK A.Nagamine



'------------------------------------------------
' �������f/����`�F�b�N���ʊ֐�
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�]�����ڇ��ɂ��A���f�����肩�𔻒f���A�������f�`�F�b�N�A�܂��́A��������`�F�b�N���Ăяo���B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sBlockid      ,I  ,String       :��ۯ�ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :�����ԍ�
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :iSmplPos      ,I  ,Integer      :�V�T���v���ʒu(mm)
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� =  1 RS     �� ����ΐ͌v�Z
'                                                       =  2 Oi     �� �����1
'                                                       =  3 BMD1   �� �����1
'                                                       =  4 BMD2   �� �����1
'                                                       =  5 BMD3   �� �����1
'                                                       =  6 OSF1   �� �����1
'                                                       =  7 OSF2   �� �����1
'                                                       =  8 OSF3   �� �����1
'                                                       =  9 OSF4   �� �����1
'                                                       = 10 CS     �� �����2(����l,�����l��0����(0<)�̏ꍇ,�����1)
'                                                       = 11 GD     �� �����1
'                                                       = 12 LT     �� �����3
'                                                       = 13 EPD    �� �����2
'          :iFromPos      ,I  ,Integer      :�����͈�From
'          :iToPos        ,I  ,Integer      :�����͈�To
'          :iHanSuiKBN    ,O  ,Integer      :���f/����敪(0:���f,1:����)
'          :iGetSmplID1   ,O  ,long         :���T���v��ID1                  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iGetSmplID2   ,O  ,long         :���T���v��ID2 (���f�����g�p)   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O  ,Integer      :�`�F�b�N���� = 0 : ����I��(���f/����OK)
'                                                           1 : ����I��(���f/����NG)
'                                                          -1 : ���͈����l�G���[
'                                                          -2 : ��L�ȊO�̃G���[
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funChkSxlHanSui(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                iItemNo As Integer, iFromPos As Integer, iToPos As Integer, iHanSuiKBN As Integer, _
                                iGetSmplID1 As Long, iGetSmplID2 As Long, tFullhin2 As tFullHinban) As Integer
    Dim retCode As Integer
    
    '���T���v��ID������
    iGetSmplID1 = -1
    iGetSmplID2 = -1
    
    '�p�����[�^�`�F�b�N
    If (Len(sBlockId) <> 12) Then GoTo ChkSxlHanSuiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkSxlHanSuiParameterErr
    
    '�w�肳�ꂽ�]�����ڇ��ɂ��A���f�����肩�𔻒f���A�������f�`�F�b�N�A�܂��́A��������`�F�b�N���Ăяo���B
    Select Case iItemNo
    Case 1              'RS(���R)
        retCode = funChkSxlSuitei(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iGetSmplID1, iGetSmplID2, tFullhin2)
        iHanSuiKBN = 1
    'Chg Start 2011/01/31 SMPK Miyata
    'Case 2 To 13        'Oi(�_�f�Z�x),BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4,CS(�Y�f�Z�x),GD,LT(ײ����),EPD
    Case 2 To 13, 15 To 18  'Oi(�_�f�Z�x),BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4,CS(�Y�f�Z�x),GD,LT(ײ����),EPD,C,CJ,CJLT,CJ2
    'Chg End   2011/01/31 SMPK Miyata
        retCode = funChkSxlHanei(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, iGetSmplID1)
        iHanSuiKBN = 0
    Case Else
        GoTo ChkSxlHanSuiParameterErr
    End Select
    
    '���ʊ֐��̃`�F�b�N���ʂ𓖊֐��̌��ʂƂ��āA�Ăяo�����֕Ԃ��B
    funChkSxlHanSui = retCode
    Exit Function

ChkSxlHanSuiParameterErr:
    funChkSxlHanSui = -1
    Exit Function

ChkSxlHanSuiSonotaErr:
    funChkSxlHanSui = -2
End Function

'------------------------------------------------
' �������f�`�F�b�N
'------------------------------------------------

'�T�v      :�w�肳�ꂽ��񂩂�A�������f�`�F�b�N���s�Ȃ����ʂ�Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sBlockid      ,I  ,String       :��ۯ�ID
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
'                                                       = 10 CS     ���Ώ�
'                                                       = 11 GD     ���Ώ�
'                                                       = 12 LT     ���Ώ�
'                                                       = 13 EPD    ���Ώ�
'          :iFromPos      ,I  ,Integer      :�����͈�From
'          :iToPos        ,I  ,Integer      :�����͈�To
'          :iGetSmplID    ,O  ,long         :���f���T���v��ID   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O  ,Integer      :�`�F�b�N���� = 0 : ����I��(���fOK)
'                                                           1 : ����I��(���fNG)
'                                                          -1 : ���͈����l�G���[
'                                                          -2 : ��L�ȊO�̃G���[
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funChkSxlHanei(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, iFromPos As Integer, iToPos As Integer, iGetSmplID As Long) As Integer
    Dim wHPtrn          As Integer
    Dim tSiyou          As type_DBDRV_scmzc_fcmkc001c_Siyou
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
    Dim strOSFRD1       As String
    Dim strOSFRD2       As String
    Dim lngOSFWID1      As Long
    Dim lngOSFWID2      As Long
    Dim strJDGEIDC      As String
    Dim strSynFlagc5    As String
    Dim strYmkFlagc5    As String
    Dim lSmpPos         As Long
    Dim strRMAXC5       As String
    Dim strDMAXC5       As String
    Dim strDRRMAXC5     As String
    Dim strDRDMAXC5     As String
    Dim lRMaxc5         As Long
    Dim lDMaxc5         As Long
    Dim lDrrMaxc5       As Long
    Dim lDrdMaxc5       As Long
    
    Dim tSiyou2         As type_DBDRV_scmzc_fcmkc001c_Siyou
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---
    Dim wGetBlockid     As String
    Dim wGetSmpKbn      As String
    Dim wGetSmplID      As Long     'Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    
    Dim tCryOi          As type_DBDRV_scmzc_fcmkc001c_Oi
    Dim tCryBMD         As type_DBDRV_scmzc_fcmkc001c_BMD
    Dim tCryOSF         As type_DBDRV_scmzc_fcmkc001c_OSF
    Dim tCryCS          As type_DBDRV_scmzc_fcmkc001c_CS
    Dim tCryGD          As type_DBDRV_scmzc_fcmkc001c_GD
    Dim tCryLT          As type_DBDRV_scmzc_fcmkc001c_LT
    Dim tCryEPD         As type_DBDRV_scmzc_fcmkc001c_EPD
    'Add Start 2011/01/31 SMPK Miyata
    Dim tCryC           As type_DBDRV_scmzc_fcmkc001c_C
    Dim tCryCJ          As type_DBDRV_scmzc_fcmkc001c_CJ
    Dim tCryCJLT        As type_DBDRV_scmzc_fcmkc001c_CJLT
    Dim tCryCJ2         As type_DBDRV_scmzc_fcmkc001c_CJ2
    'Add End   2011/01/31 SMPK Miyata

    Dim retJudg         As Boolean
    Dim wIdFlg          As Integer
    
    Dim dShiyo()        As Double       '2003/12/11 Null�Ή��ǉ�
    Dim sHosyo          As String       '2003/12/11 Null�Ή��ǉ�
    
    '������
    wGetSmplID = -1
    
    '�p�����[�^�`�F�b�N
    If (Len(sBlockId) <> 12) Then GoTo ChkSxlHaneiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkSxlHaneiParameterErr
    
    '�w�肳�ꂽ�]�����ڇ����ɕK�v�ȕi�Ԏd�l�l���擾���A�������f�l�擾�p�^�[�������肷��B�i�w�肳�ꂽ�]�����ڇ��ɂ��A�������������B�j
    Select Case iItemNo
    Case 1                      'RS(���R)
        GoTo ChkSxlHaneiNG
    Case 2                      'Oi(�_�f�Z�x)
        If funGet_TBCME019(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 1
        
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        ReDim dShiyo(5)
        dShiyo(1) = tSiyou.HSXONMIN         ' �i�r�w�_�f�Z�x����
        dShiyo(2) = tSiyou.HSXONMAX         ' �i�r�w�_�f�Z�x���
        dShiyo(3) = tSiyou.HSXONAMN         ' �i�r�w�_�f�Z�x���ω���
        dShiyo(4) = tSiyou.HSXONAMX         ' �i�r�w�_�f�Z�x���Ϗ��
        dShiyo(5) = tSiyou.HSXONMBP         ' �i�r�w�_�f�Z�x�ʓ����z
        'NULL�͕s��(NULL�����֐����ı��) 09/03/13 ooba
'        If fncJissekiHantei_nl(tSiyou.HSXONHWS, dShiyo) = False Then GoTo ChkSxlHaneiSonotaErr
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        
    Case 3, 4, 5                'BMD1,BMD2,BMD3
        If funGet_TBCME020(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 1
                
    Case 6, 7, 8, 9             'OSF1,OSF2,OSF3,OSF4
        If funGet_TBCME020(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 1
        
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        If funGet_TBCME036(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
            
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        ReDim dShiyo(2)
        If iItemNo = 6 Then         'OSF1
            sHosyo = tSiyou.HSXOF1HS            ' �i�r�w�n�r�e1�ۏؕ��@�Q��
            dShiyo(1) = tSiyou.HSXOF1AX         ' �i�r�w�n�r�e1���Ϗ��
            dShiyo(2) = tSiyou.HSXOF1MX         ' �i�r�w�n�r�e1���
        ElseIf iItemNo = 7 Then     'OSF2
            sHosyo = tSiyou.HSXOF2HS            ' �i�r�w�n�r�e2�ۏؕ��@�Q��
            dShiyo(1) = tSiyou.HSXOF2AX         ' �i�r�w�n�r�e2���Ϗ��
            dShiyo(2) = tSiyou.HSXOF2MX         ' �i�r�w�n�r�e2���
        ElseIf iItemNo = 8 Then     'OSF3
            sHosyo = tSiyou.HSXOF3HS            ' �i�r�w�n�r�e3�ۏؕ��@�Q��
            dShiyo(1) = tSiyou.HSXOF3AX         ' �i�r�w�n�r�e3���Ϗ��
            dShiyo(2) = tSiyou.HSXOF3MX         ' �i�r�w�n�r�e3���
        ElseIf iItemNo = 9 Then     'OSF4
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
            'C-OSF3�׸ފl��
            If funGet_TBCME036(tFullHin, tSiyou2) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
            wHPtrn = 1
            '�׸ނ̌����ύX
            sHosyo = tSiyou2.COSF3FLAG          ' �b�|�n�r�e�R�׸�
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END  ---
            dShiyo(1) = tSiyou.HSXOF4AX         ' �i�r�w�n�r�e4���Ϗ��
            dShiyo(2) = tSiyou.HSXOF4MX         ' �i�r�w�n�r�e4���
        End If
        
'C�|OSF3����@�\�ǉ� 2007/06/14 M.Kaga STRAT ---
        If iItemNo <> 9 Then
            'NULL�͕s��(NULL�����֐����ı��) 09/03/13 ooba
'            If fncJissekiHantei_nl(sHosyo, dShiyo) = False Then GoTo ChkSxlHaneiSonotaErr
            'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        End If
'C�|OSF3����@�\�ǉ� 2007/06/14 M.Kaga END ---

    Case 10                     'CS(�Y�f�Z�x)
        If funGet_TBCME019(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        'TOP/BOT�ۏ؂͔��f�����1,BOT�ۏ؂͔��f�����2 09/01/08 ooba
        If tSiyou.HSXCNKHI = "6" Or tSiyou.HSXCNKHI = "9" Then
            wHPtrn = 1
        Else
            wHPtrn = 2
        End If
        
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        ReDim dShiyo(2)
        dShiyo(1) = tSiyou.HSXCNMIN         ' �i�r�w�Y�f�Z�x����
        dShiyo(2) = tSiyou.HSXCNMAX         ' �i�r�w�Y�f�Z�x���
        'NULL�͕s��(NULL�����֐����ı��) 09/03/13 ooba
'        If fncJissekiHantei_nl(tSiyou.HSXCNHWS, dShiyo) = False Then GoTo ChkSxlHaneiSonotaErr
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        
    Case 11                     'GD
        If funGet_TBCME020(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
    '*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݎ擾�ǉ�
        If funGet_TBCME036(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
    '*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݎ擾�ǉ�
        wHPtrn = 1
                
    Case 12                     'LT(ײ����)
        If funGet_TBCME019(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        If funGet_TBCME036(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 3
    
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        ReDim dShiyo(2)
        dShiyo(1) = tSiyou.HSXLTMIN         ' �i�r�w�k�^�C������
        dShiyo(2) = tSiyou.HSXLTMAX         ' �i�r�w�k�^�C�����
        'NULL�͕s��(NULL�����֐����ı��) 09/03/13 ooba
'        If fncJissekiHantei_nl(tSiyou.HSXLTHWS, dShiyo) = False Then GoTo ChkSxlHaneiSonotaErr
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        
    Case 13                     'EPD
        If funGet_TBCME036(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 2
    
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        If tSiyou.EPDUP = -1 Then GoTo ChkSxlHaneiSonotaErr     ' EPD���
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��

    'Add Start 2011/01/31 SMPK Miyata
    Case 15, 16, 17, 18         'C,CJ,CJLT,CJ2
        If funGet_TBCME020(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 1
        
        If iItemNo = 17 Then
            If funGet_TBCME036(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        End If
    'Add End   2011/01/31 SMPK Miyata

    Case Else
        GoTo ChkSxlHaneiParameterErr
    End Select

    '�������f���T���v���h�c�̎擾
    If wHPtrn = 1 Then              '�������f�l�擾�p�^�[���P
        If funGetSxlHanei1(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, _
                                                                    wGetBlockid, wGetSmpKbn, wGetSmplID) <> 0 Then GoTo ChkSxlHaneiNG
    
    ElseIf wHPtrn = 2 Then          '�������f�l�擾�p�^�[���Q
        If funGetSxlHanei2(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, _
                                                                    wGetBlockid, wGetSmpKbn, wGetSmplID) <> 0 Then GoTo ChkSxlHaneiNG
    
    ElseIf wHPtrn = 3 Then          '�������f�l�擾�p�^�[���R
        If funGetSxlHanei3(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, tSiyou.HSXLTSPI, _
                                                                    wGetBlockid, wGetSmpKbn, wGetSmplID) <> 0 Then GoTo ChkSxlHaneiNG
    
    End If
    
    '�������f�������ID����A�������f�l�i���ђl�j���擾����B�i�w�肳�ꂽ�]�����ڇ��ɂ��A�������������B�j
    Select Case iItemNo
    Case 2                      'Oi(�_�f�Z�x)
        'Oi�̎��ђl���擾����
        If funGetCryOiJisseki(sCryNum, wGetSmplID, tCryOi) <> 0 Then GoTo ChkSxlHaneiNG
        'Oi����������s�Ȃ�
        If Not CrOiJudg(tSiyou, tCryOi, retJudg) Then GoTo ChkSxlHaneiNG
    
    Case 3, 4, 5                'BMD1, BMD2, BMD3
        If iItemNo = 3 Then
            'BMD1�̎��ђl���擾����
            If funGetCryBMDJisseki(sCryNum, wGetSmplID, 1, tCryBMD) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 4 Then
            'BMD2�̎��ђl���擾����
            If funGetCryBMDJisseki(sCryNum, wGetSmplID, 2, tCryBMD) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 5 Then
            'BMD3�̎��ђl���擾����
            If funGetCryBMDJisseki(sCryNum, wGetSmplID, 3, tCryBMD) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 3
        End If
        'BMD�̑���������s�Ȃ�
        If Not CrBmdJudg(tSiyou, tCryBMD, retJudg, wIdFlg) Then GoTo ChkSxlHaneiNG
    
    Case 6, 7, 8, 9             'OSF1, OSF2, OSF3, OSF4
        If iItemNo = 6 Then
            'OSF1�̎��ђl���擾����
            If funGetCryOSFJisseki(sCryNum, wGetSmplID, 1, tCryOSF) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 7 Then
            'OSF2�̎��ђl���擾����
            If funGetCryOSFJisseki(sCryNum, wGetSmplID, 2, tCryOSF) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 8 Then
            'OSF3�̎��ђl���擾����
            If funGetCryOSFJisseki(sCryNum, wGetSmplID, 3, tCryOSF) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 3
        ElseIf iItemNo = 9 Then
            'OSF4�̎��ђl���擾����
            If funGetCryOSFJisseki(sCryNum, wGetSmplID, 4, tCryOSF) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 4
        End If
        
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
        'OSF1,2,3�̏ꍇ
        If iItemNo = 6 Or iItemNo = 7 Or iItemNo = 8 Then
            'OSF�̑���������s�Ȃ�
            If Not CrOsfJudg(tSiyou, tCryOSF, retJudg, wIdFlg) Then GoTo ChkSxlHaneiNG
        Else
            'OSF���ѓ��͔��菈��
            '���������&���ђl�ޔ�
            If Trim(tCryOSF.OSFRD1) = "R" Or Trim(tCryOSF.OSFRD1) = "D" Then
                strOSFRD1 = Trim(tCryOSF.OSFRD1)
            Else
                strOSFRD1 = "-"
            End If
            If Trim(tCryOSF.OSFRD2) = "D" Then
                strOSFRD2 = Trim(tCryOSF.OSFRD2)
            Else
                strOSFRD2 = "-"
            End If
            If IsNull(tCryOSF.OSFWID1) = True Then
               lngOSFWID1 = -1
            ElseIf IsNumeric(tCryOSF.OSFWID1) = False Then
               lngOSFWID1 = -1
            Else
               lngOSFWID1 = Trim(tCryOSF.OSFWID1)
            End If
            If IsNull(tCryOSF.OSFWID2) = True Then
               lngOSFWID2 = -1
            ElseIf IsNumeric(tCryOSF.OSFWID2) = False Then
               lngOSFWID2 = -1
            Else
               lngOSFWID2 = Trim(tCryOSF.OSFWID2)
            End If
            
            '-1�ȊO�̐��l�l��
            If lngOSFWID1 < 0 Then
               lngOSFWID1 = -1
            End If
            If lngOSFWID2 < 0 Then
               lngOSFWID2 = -1
            End If

            '����݋敪�A���ђl��NULL�̏ꍇ
            If strOSFRD1 = "-" And strOSFRD2 <> "-" Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            ElseIf strOSFRD1 <> "-" And lngOSFWID1 = -1 Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            ElseIf strOSFRD2 <> "-" And lngOSFWID2 = -1 Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            ElseIf strOSFRD2 = "-" And lngOSFWID2 > 0 Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
        
             '�����ԍ��𷰂Ƃ���XSDC1���C�|OSF3����ID���l������
            If GetCOSF3ID(strJDGEIDC, Trim(sCryNum)) <> FUNCTION_RETURN_SUCCESS Then
                '�Y��ں��ޖ����̏ꍇ
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
            'C�|OSF3����ID��NULL�̏ꍇ
            If Trim(strJDGEIDC) = "" Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            Else
                '�l������C-OSF3����ID��XODC5_OSF30��菳�F�׸ނ̊l��
                If GetSYNFLAGC5(strSynFlagc5, strYmkFlagc5, Trim(strJDGEIDC)) <> FUNCTION_RETURN_SUCCESS Then
                    '�Y��ں��ޖ����̏ꍇ
                    retJudg = False
                    GoTo ChkSxlHaneiNG
                End If
                '���F�׸�:0�@�����F�̏ꍇ
                If Trim(strSynFlagc5) = "0" Or Trim(strSynFlagc5) = "" Or IsNull(strSynFlagc5) Then
                    retJudg = False
                    GoTo ChkSxlHaneiNG
                '�폜�׸�:1�@�����̏ꍇ
                ElseIf Trim(strYmkFlagc5) = "1" Or Trim(strYmkFlagc5) = "" Or IsNull(strYmkFlagc5) Then
                    retJudg = False
                    GoTo ChkSxlHaneiNG
                Else
                
                    '����وʒu�擾(���f�������NO�ɕR�t���ʒu)
                    lSmpPos = Trim(iSmplPos)
                    
                    '����݋敪�ɂ�菈������
                    'R�݂̂̏ꍇ
                    If strOSFRD1 = "R" And strOSFRD2 = "-" Then
                        'R�̂ݏ���l�̊l�����s��
                        If GetCOSF3PTN(strJDGEIDC, lSmpPos, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                            '�Y��ں��ޖ����̏ꍇ
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        End If
                        'ں��ޖ��FVB�G���[(��ōl����)
                        If Trim(strRMAXC5) = "" Then
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        Else
                            lRMaxc5 = Trim(strRMAXC5)
                            '���ђl�̔���
                            If lngOSFWID1 <= lRMaxc5 Then
                                '���fOK
                                retJudg = True
                            ElseIf lngOSFWID1 > lRMaxc5 Then
                                '���fNG
                                retJudg = False
                            End If
                        End If
                    'D�݂̂̏ꍇ
                    ElseIf strOSFRD1 = "D" Then
                        'D�̂ݏ���l�̊l�����s��
                        If GetCOSF3PTN(strJDGEIDC, lSmpPos, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                            '�Y��ں��ޖ����̏ꍇ
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        End If
                                                
                        'ں��ޖ�����Ͻ��̎��ђl��NULL�FVB�G���[(��ōl����)
                        If Trim(strDMAXC5) = "" Then
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        Else
                            lDMaxc5 = Trim(strDMAXC5)
                            '���ђl�̔���
                            If lngOSFWID1 <= lDMaxc5 Then
                                '���fOK
                                retJudg = True
                            ElseIf lngOSFWID1 > lDMaxc5 Then
                                '���fNG
                                retJudg = False
                            End If
                        End If
                    'R&D�̏ꍇ
                    ElseIf strOSFRD1 = "R" And strOSFRD2 = "D" Then
                        'D��������l����R��������l�̊l�����s��
                        If GetCOSF3PTN(strJDGEIDC, lSmpPos, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                         '�Y��ں��ޖ����̏ꍇ
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        End If
                                                
                        'ں��ޖ�����Ͻ��̎��ђl��NULL�FVB�G���[(��ōl����)
                        If Trim(strDRRMAXC5) = "" Or Trim(strDRDMAXC5) = "" Then
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        Else
                            lDrrMaxc5 = Trim(strDRRMAXC5)
                            lDrdMaxc5 = Trim(strDRDMAXC5)
                            '���ђl�̔���
                            If lngOSFWID1 <= lDrrMaxc5 And lngOSFWID2 <= lDrdMaxc5 Then
                                '���fOK
                                retJudg = True
                            ElseIf lngOSFWID1 > lDrrMaxc5 Or lngOSFWID2 > lDrdMaxc5 Then
                                '���fNG
                                retJudg = False
                            End If
                        End If
                    Else
                        '���ђl���A����݋敪���̏ꍇ���fOK
                         retJudg = True
                    End If
                End If
            End If
        End If
        
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---
    
    Case 10             'CS(�Y�f�Z�x)
        'CS�̎��ђl���擾����
        If funGetCryCSJisseki(sCryNum, wGetSmplID, tCryCS) <> 0 Then GoTo ChkSxlHaneiNG
        'CS����������s�Ȃ�
        If Not CrCsjudg(tSiyou, tCryCS, retJudg) Then GoTo ChkSxlHaneiNG
    
    Case 11             'GD
        'GD�̎��ђl���擾����
        If funGetCryGDJisseki(sCryNum, wGetSmplID, tCryGD) <> 0 Then GoTo ChkSxlHaneiNG
        'GD����������s�Ȃ�
        If Not CrGdjudg(tSiyou, tCryGD, retJudg) Then GoTo ChkSxlHaneiNG
    
    Case 12             'LT(ײ����)
        'LT�̎��ђl���擾����
        If funGetCryLTJisseki(sCryNum, wGetSmplID, tCryLT) <> 0 Then GoTo ChkSxlHaneiNG
        '2005/12/02 add SET���� LT�v�Z�֐���call���� ->
        '���C�t�^�C���l���v�Z���Ȃ���
        Call Sub_LTReCalc(tSiyou, tCryLT)
        '2005/12/02 add SET���� LT�v�Z�֐���call���� <-
        'LT����������s�Ȃ�
        If Not CrLtjudg(tSiyou, tCryLT, retJudg) Then GoTo ChkSxlHaneiNG
    
    Case 13             'EPD
        'EPD�̎��ђl���擾����
        If funGetCryEPDJisseki(sCryNum, wGetSmplID, tCryEPD) <> 0 Then GoTo ChkSxlHaneiNG
        'EPD����������s�Ȃ�
        If Not CrEpdjudg(tSiyou, tCryEPD, retJudg) Then GoTo ChkSxlHaneiNG

    'Add Start 2011/01/31 SMPK Miyata
    Case 15             'C
        If funGetCryCJisseki(sCryNum, wGetSmplID, tCryC) <> 0 Then GoTo ChkSxlHaneiNG

         '�����ԍ��𷰂Ƃ���XSDC1���C�|OSF3����ID���l������
        If GetCOSF3ID(strJDGEIDC, Trim(sCryNum)) <> FUNCTION_RETURN_SUCCESS Then
            '�Y��ں��ޖ����̏ꍇ
            retJudg = False
            GoTo ChkSxlHaneiNG
        End If
        'C�|OSF3����ID��NULL�̏ꍇ
        If Trim(strJDGEIDC) = "" Then
            retJudg = False
            GoTo ChkSxlHaneiNG
        Else
            '�l������C-OSF3����ID��XODC5_OSF30��菳�F�׸ނ̊l��
            If GetSYNFLAGC5(strSynFlagc5, strYmkFlagc5, Trim(strJDGEIDC)) <> FUNCTION_RETURN_SUCCESS Then
                '�Y��ں��ޖ����̏ꍇ
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
            '���F�׸�:0�@�����F�̏ꍇ
            If Trim(strSynFlagc5) = "0" Or Trim(strSynFlagc5) = "" Or IsNull(strSynFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            '�폜�׸�:1�@�����̏ꍇ
            ElseIf Trim(strYmkFlagc5) = "1" Or Trim(strYmkFlagc5) = "" Or IsNull(strYmkFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
        End If

        'C����������s�Ȃ�
        If Not CrCjudg(tSiyou, tCryC, retJudg) Then GoTo ChkSxlHaneiNG

    Case 16             'CJ
        If funGetCryCJJisseki(sCryNum, wGetSmplID, tCryCJ) <> 0 Then GoTo ChkSxlHaneiNG

         '�����ԍ��𷰂Ƃ���XSDC1���C�|OSF3����ID���l������
        If GetCOSF3ID(strJDGEIDC, Trim(sCryNum)) <> FUNCTION_RETURN_SUCCESS Then
            '�Y��ں��ޖ����̏ꍇ
            retJudg = False
            GoTo ChkSxlHaneiNG
        End If
        'C�|OSF3����ID��NULL�̏ꍇ
        If Trim(strJDGEIDC) = "" Then
            retJudg = False
            GoTo ChkSxlHaneiNG
        Else
            '�l������C-OSF3����ID��XODC5_OSF30��菳�F�׸ނ̊l��
            If GetSYNFLAGC5(strSynFlagc5, strYmkFlagc5, Trim(strJDGEIDC)) <> FUNCTION_RETURN_SUCCESS Then
                '�Y��ں��ޖ����̏ꍇ
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
            '���F�׸�:0�@�����F�̏ꍇ
            If Trim(strSynFlagc5) = "0" Or Trim(strSynFlagc5) = "" Or IsNull(strSynFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            '�폜�׸�:1�@�����̏ꍇ
            ElseIf Trim(strYmkFlagc5) = "1" Or Trim(strYmkFlagc5) = "" Or IsNull(strYmkFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
        End If

        'CJ����������s�Ȃ�
        If Not CrCJjudg(tSiyou, tCryCJ, retJudg) Then GoTo ChkSxlHaneiNG
    
    Case 17             'CJLT
        If funGetCryCJLTJisseki(sCryNum, wGetSmplID, tCryCJLT) <> 0 Then GoTo ChkSxlHaneiNG

         '�����ԍ��𷰂Ƃ���XSDC1���C�|OSF3����ID���l������
        If GetCOSF3ID(strJDGEIDC, Trim(sCryNum)) <> FUNCTION_RETURN_SUCCESS Then
            '�Y��ں��ޖ����̏ꍇ
            retJudg = False
            GoTo ChkSxlHaneiNG
        End If
        'C�|OSF3����ID��NULL�̏ꍇ
        If Trim(strJDGEIDC) = "" Then
            retJudg = False
            GoTo ChkSxlHaneiNG
        Else
            '�l������C-OSF3����ID��XODC5_OSF30��菳�F�׸ނ̊l��
            If GetSYNFLAGC5(strSynFlagc5, strYmkFlagc5, Trim(strJDGEIDC)) <> FUNCTION_RETURN_SUCCESS Then
                '�Y��ں��ޖ����̏ꍇ
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
            '���F�׸�:0�@�����F�̏ꍇ
            If Trim(strSynFlagc5) = "0" Or Trim(strSynFlagc5) = "" Or IsNull(strSynFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            '�폜�׸�:1�@�����̏ꍇ
            ElseIf Trim(strYmkFlagc5) = "1" Or Trim(strYmkFlagc5) = "" Or IsNull(strYmkFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
        End If

        'CJLT����������s�Ȃ�
        If Not CrCJLTjudg(tSiyou, tCryCJLT, retJudg) Then GoTo ChkSxlHaneiNG

    Case 18             'CJ2
        If funGetCryCJ2Jisseki(sCryNum, wGetSmplID, tCryCJ2) <> 0 Then GoTo ChkSxlHaneiNG

         '�����ԍ��𷰂Ƃ���XSDC1���C�|OSF3����ID���l������
        If GetCOSF3ID(strJDGEIDC, Trim(sCryNum)) <> FUNCTION_RETURN_SUCCESS Then
            '�Y��ں��ޖ����̏ꍇ
            retJudg = False
            GoTo ChkSxlHaneiNG
        End If
        'C�|OSF3����ID��NULL�̏ꍇ
        If Trim(strJDGEIDC) = "" Then
            retJudg = False
            GoTo ChkSxlHaneiNG
        Else
            '�l������C-OSF3����ID��XODC5_OSF30��菳�F�׸ނ̊l��
            If GetSYNFLAGC5(strSynFlagc5, strYmkFlagc5, Trim(strJDGEIDC)) <> FUNCTION_RETURN_SUCCESS Then
                '�Y��ں��ޖ����̏ꍇ
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
            '���F�׸�:0�@�����F�̏ꍇ
            If Trim(strSynFlagc5) = "0" Or Trim(strSynFlagc5) = "" Or IsNull(strSynFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            '�폜�׸�:1�@�����̏ꍇ
            ElseIf Trim(strYmkFlagc5) = "1" Or Trim(strYmkFlagc5) = "" Or IsNull(strYmkFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
        End If

        'CJ2����������s�Ȃ�
        If Not CrCJ2judg(tSiyou, tCryCJ2, retJudg) Then GoTo ChkSxlHaneiNG
    'Add End   2011/01/31 SMPK Miyata

    End Select

    '�w�肳�ꂽ�]�����ڇ��̑������肪OK�̏ꍇ�A���f���T���v��ID��ݒ肵�A�߂�l��'0'(����I��(���fOK))��ݒ肵�A�������I������B
    '�������肪NG�̏ꍇ�A�߂�l��'1'(����I��(���fNG))��ݒ肵�A�������I������B
    If retJudg = False Then GoTo ChkSxlHaneiNG
        
    iGetSmplID = wGetSmplID
    funChkSxlHanei = 0
    Exit Function

ChkSxlHaneiNG:
    iGetSmplID = wGetSmplID
    funChkSxlHanei = 1
    Exit Function

ChkSxlHaneiParameterErr:
    funChkSxlHanei = -1
    Exit Function

ChkSxlHaneiSonotaErr:
    funChkSxlHanei = -2
    Exit Function
    
End Function

'------------------------------------------------
' ��������`�F�b�N
'------------------------------------------------

'�T�v      :�w�肳�ꂽ��񂩂�A��������`�F�b�N���s�Ȃ����ʂ�Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sBlockid      ,I  ,String       :��ۯ�ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :�����ԍ�
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :iSmplPos      ,I  ,Integer      :�V�T���v���ʒu(mm)
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� =  1 RS     ���Ώ�
'                                                       =  2 Oi
'                                                       =  3 BMD1
'                                                       =  4 BMD2
'                                                       =  5 BMD3
'                                                       =  6 OSF1
'                                                       =  7 OSF2
'                                                       =  8 OSF3
'                                                       =  9 OSF4
'                                                       = 10 CS
'                                                       = 11 GD
'                                                       = 12 LT
'                                                       = 13 EPD
'          :iGetSmplID1   ,O  ,long         :���茳�T���v��ID1      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iGetSmplID2   ,O  ,long         :���茳�T���v��ID2      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O  ,Integer      :�`�F�b�N���� = 0 : ����I��(����OK)
'                                                           1 : ����I��(����NG)
'                                                          -1 : ���͈����l�G���[
'                                                          -2 : ��L�ȊO�̃G���[
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funChkSxlSuitei(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                iItemNo As Integer, iGetSmplID1 As Long, iGetSmplID2 As Long, tfullhin1 As tFullHinban) As Integer
    Dim retCode         As Integer
    Dim tSiyou          As type_DBDRV_scmzc_fcmkc001c_Siyou
    Dim tTBCME037       As c_cmzcXl
    Dim sqlWhere        As String
    
    Dim wGetBlockidTop  As String
    Dim wGetSmpKbnTop   As String
    Dim wGetSmplIDTop   As Long         'Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    Dim wGetSPtrnTop    As String
    Dim wGetBlockidBot  As String
    Dim wGetSmpKbnBot   As String
    Dim wGetSmplIDBot   As Long         'Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    Dim wGetSPtrnBot    As String
    Dim wGetPosTop      As Integer    '2005/1/11
    Dim wGetPosBot      As Integer    '2005/1/11
    Dim tCryRs(2)       As type_DBDRV_scmzc_fcmkc001c_CryR        '(0)�����茳Top, (1)�����茳Bot, (2)�������
    Dim wcnt            As Integer
    Dim wMeasTop(4)     As Double                   'Top����l
    Dim wMeasBot(4)     As Double                   'Bot����l
    Dim wMeasSui()      As Double                   '�Z�o����l
    Dim retJudg         As Boolean
    
    Dim dShiyo(5)       As Double       '2003/12/11 Null�Ή��ǉ�
    
    Dim i      As Integer  'TEST2004/10
    Dim i2     As Integer  'TEST2004/10
    Dim sCnt   As Integer  'TEST2004/10
    '������
    wGetSmplIDTop = -1
    wGetSmplIDBot = -1
    sCnt = UBound(SuiteiData) + 1
    ReDim Preserve SuiteiData(sCnt)
    For i2 = 0 To 2
        SuiteiData(sCnt).SuiData(i2).MEAS1 = 0
        SuiteiData(sCnt).SuiData(i2).MEAS2 = 0
        SuiteiData(sCnt).SuiData(i2).MEAS3 = 0
        SuiteiData(sCnt).SuiData(i2).MEAS4 = 0
        SuiteiData(sCnt).SuiData(i2).MEAS5 = 0
    Next
    '�p�����[�^�`�F�b�N
    If (Len(sBlockId) <> 12) Then GoTo ChkSxlSuiteiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkSxlSuiteiParameterErr
    
    '�w�肳�ꂽ�]�����ڇ����ɕK�v�ȕi�Ԏd�l�l���擾����B�i�w�肳�ꂽ�]�����ڇ��ɂ��A�������������B�j
    Select Case iItemNo
    Case 1              'RS(���R)
        If Trim(tFullHin.hinban) <> "Z" And Trim(tFullHin.hinban) <> "G" Then
            If funGet_TBCME018(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlSuiteiNG
        End If
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        dShiyo(1) = tSiyou.HSXRMIN          ' �i�r�w���R����
        dShiyo(2) = tSiyou.HSXRMAX          ' �i�r�w���R���
        dShiyo(3) = tSiyou.HSXRAMIN         ' �i�r�w���R���ω���
        dShiyo(4) = tSiyou.HSXRAMAX         ' �i�r�w���R���Ϗ��
        dShiyo(5) = tSiyou.HSXRMBNP         ' �i�r�w���R�ʓ����z
        'NULL�͕s��(NULL�����֐����ı��) 09/03/13 ooba
'        If fncJissekiHantei_nl(tSiyou.HSXRHWYS, dShiyo) = False Then GoTo ChkSxlSuiteiSonotaErr
        'Null�Ή������ǉ� 2003/12/11 SystenBrain ��
        
    Case 2 To 13        'Oi(�_�f�Z�x),BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4,CS(�Y�f�Z�x),GD,LT(ײ����),EPD
        GoTo ChkSxlSuiteiNG
    Case Else
        GoTo ChkSxlSuiteiParameterErr
    End Select
    '�������茳�T���v���h�c�̎擾    '2005/1/11 �C���@�ʒu�ǉ�
    If funGetSuitei(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, _
                    wGetBlockidTop, wGetSmpKbnTop, wGetSmplIDTop, wGetSPtrnTop, _
                    wGetBlockidBot, wGetSmpKbnBot, wGetSmplIDBot, wGetSPtrnBot, wGetPosTop, wGetPosBot) <> 0 Then GoTo ChkSxlSuiteiNG
    
    '�������茳��ۯ�ID1�A�������茳����ً敪1�A�������茳�����ID1����A���茳���ђl���擾����B
    'RS�̎��ђl���擾����
    If funGetCryRsJisseki(sCryNum, wGetSmplIDTop, tCryRs(0)) <> 0 Then GoTo ChkSxlSuiteiNG
    '�������茳�u���b�NID2��������茳�T���v���敪2��������茳�T���v��ID2���礐��茳���ђl���擾����
    'RS�̎��ђl���擾����
    If funGetCryRsJisseki(sCryNum, wGetSmplIDBot, tCryRs(1)) <> 0 Then GoTo ChkSxlSuiteiNG
    
    With SuiteiData(sCnt)
        .SuiSpec = tSiyou
        .SuiData(0) = tCryRs(0)
        .SuiData(1) = tCryRs(1)
        Debug.Print .SuiSpec.HIN.hinban
        '�����l�̃`�F�b�N�i�e�i�Ԃ̐���������N���A���Ȃ��Ɛ���ł��Ȃ��j
        .RsJudg(1) = True
        .RsJudg(2) = True
        If Trim(tFullHin.hinban) <> "Z" And Trim(tFullHin.hinban) <> "G" Then
            If funChkJissoku(tFullHin, tCryRs(0)) = False Then .RsJudg(1) = False
            If funChkJissoku(tFullHin, tCryRs(1)) = False Then .RsJudg(2) = False
        End If
        If Trim(tfullhin1.hinban) <> "Z" And Trim(tfullhin1.hinban) <> "G" Then
            If funChkJissoku(tfullhin1, tCryRs(0)) = False Then .RsJudg(1) = False
            If funChkJissoku(tfullhin1, tCryRs(1)) = False Then .RsJudg(2) = False
        End If
        If .RsJudg(1) = False Or .RsJudg(2) = False Then GoTo ChkSxlSuiteiNG
    End With
    '--------------
    '�����̎��уf�[�^�ҏW
    With tCryRs(2)
        .CRYNUM = sCryNum
        .POSITION = iSmplPos
        .SMPKBN = sTB
        .TRANCOND = "0"
        .TRANCNT = 1
        .SMPLNO = -1
        .SMPLUMU = "1"
    End With
    
    '------TEST2004/10
    If tCryRs(0).KSTAFFID = KSTAFF_J002 Then
        GoTo ChkSxlSuiteiNG
    End If
    wcnt = funGetRsCnt(tCryRs(0))
    If wcnt <> 5 Then GoTo ChkSxlSuiteiNG
    
    'Top/Bot����l�𐄒�l�Z�o�p�ɃZ�b�g
    If wGetSPtrnTop = "A" Then                  '����p�^�[��A
        wMeasTop(0) = tCryRs(0).MEAS1
        wMeasTop(1) = tCryRs(0).MEAS2
        wMeasTop(2) = tCryRs(0).MEAS3
        wMeasTop(3) = tCryRs(0).MEAS4
        wMeasTop(4) = tCryRs(0).MEAS5
    ElseIf wGetSPtrnTop = "B" Then              '����p�^�[��B
        wMeasTop(0) = tCryRs(0).MEAS1
        wMeasTop(1) = tCryRs(0).MEAS4
        wMeasTop(2) = tCryRs(0).MEAS5
        wMeasTop(3) = 0
        wMeasTop(4) = 0
    End If
    
    '------TEST2004/10
    If tCryRs(1).KSTAFFID = KSTAFF_J002 Then
        GoTo ChkSxlSuiteiNG
    End If
    wcnt = funGetRsCnt(tCryRs(1))
    If wcnt <> 5 Then GoTo ChkSxlSuiteiNG
    
    If wGetSPtrnBot = "A" Then                  '����p�^�[��A
        wMeasBot(0) = tCryRs(1).MEAS1
        wMeasBot(1) = tCryRs(1).MEAS2
        wMeasBot(2) = tCryRs(1).MEAS3
        wMeasBot(3) = tCryRs(1).MEAS4
        wMeasBot(4) = tCryRs(1).MEAS5
    ElseIf wGetSPtrnBot = "B" Then              '����p�^�[��B
        wMeasBot(0) = tCryRs(1).MEAS1
        wMeasBot(1) = tCryRs(1).MEAS4
        wMeasBot(2) = tCryRs(1).MEAS5
        wMeasBot(3) = 0
        wMeasBot(4) = 0
    End If
    
    ReDim wMeasSui(5 - 1)
    For wcnt = 0 To 5 - 1
        '����l�̎Z�o
        retCode = new_ResSuitei(sCryNum, wMeasTop(wcnt), tCryRs(0).POSITION, wMeasBot(wcnt), tCryRs(1).POSITION, iSmplPos, wMeasSui(wcnt))
        If retCode = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlSuiteiNG
    
    Next wcnt
    '�������(TBCME037)�f�[�^�̎擾
    sqlWhere = " Where (CRYNUM='" & sCryNum & "')"
    If GetTBCME037(tTBCME037, sqlWhere) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlSuiteiSonotaErr
    
    '-----TEST2004/10
    '����l�̐ݒ�
    tCryRs(2).MEAS1 = wMeasSui(0)
    tCryRs(2).MEAS2 = wMeasSui(1)
    tCryRs(2).MEAS3 = wMeasSui(2)
    tCryRs(2).MEAS4 = wMeasSui(3)
    tCryRs(2).MEAS5 = wMeasSui(4)
    '-----TEST2004/10
    SuiteiData(sCnt).SuiData(2) = tCryRs(2)
    '��R�f�[�^�𑪒�ʒu�ɂ����בւ���
    If Trim(tFullHin.hinban) = "Z" Or Trim(tFullHin.hinban) = "G" Then
        retJudg = True
        SuiteiData(sCnt).SuiSpec.HIN.hinban = Trim(tFullHin.hinban)
        SuiteiData(sCnt).COEFflg = True
        SuiteiData(sCnt).DOPEflg = True
    Else
        If Set_Rs_Ichi(tSiyou.HSXRSPOT, tSiyou.HSXRSPOI, tCryRs(2).MEAS1, tCryRs(2).MEAS2, _
                           tCryRs(2).MEAS3, tCryRs(2).MEAS4, tCryRs(2).MEAS5) Then GoTo ChkSxlSuiteiNG
        
        '����v�Z�ŎZ�o��������l��RS����������s�Ȃ�
        If Not CrResJudg(0, tSiyou, tCryRs(2), retJudg, 1) Then GoTo ChkSxlSuiteiNG
        '2005/1/11 �u���b�N�ΐ͒l�͈͊O,���ް�߈ʒu���܂ރu���b�N�͐���s��
        If HenDopeJudg(wGetPosTop, iSmplPos, tCryRs(0).MEAS1, tCryRs(2).MEAS1, sCryNum, tFullHin) = False Then
            GoTo ChkSxlSuiteiNG
        End If
        If HenDopeJudg(iSmplPos, wGetPosBot, tCryRs(2).MEAS1, tCryRs(1).MEAS1, sCryNum, tFullHin) = False Then
            GoTo ChkSxlSuiteiNG
        End If
    End If
    '�w�肳�ꂽ�]�����ڇ��̑������肪OK�̏ꍇ�A���茳�T���v��ID1�Ɛ��茳�T���v��ID2��ݒ肵�A�߂�l��'0'(����I��(����OK))��ݒ肵�A�������I������B
    '�������肪NG�̏ꍇ�A�߂�l��'1'(����I��(����NG))��ݒ肵�A�������I������B
    
    If retJudg = False Then GoTo ChkSxlSuiteiNG
        
    iGetSmplID1 = wGetSmplIDTop
    iGetSmplID2 = wGetSmplIDBot
    funChkSxlSuitei = 0
    Exit Function

ChkSxlSuiteiNG:
    iGetSmplID1 = wGetSmplIDTop
    iGetSmplID2 = wGetSmplIDBot
    funChkSxlSuitei = 1
    Exit Function

ChkSxlSuiteiParameterErr:
    funChkSxlSuitei = -1
    Exit Function

ChkSxlSuiteiSonotaErr:
    funChkSxlSuitei = -2
End Function

'------------------------------------------------
' �������f�l�擾�i�p�^�[���P�j
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�V�T���v���ʒu��񂩂�A�������f���T���v���h�c��V�T���v���Ǘ�(��ۯ�)(XSDCS)��茟�����A���ʂ�Ԃ��B
'           ���f���悤�Ƃ���V�T���v���ʒu���ATOP�̏ꍇ��BOT�̏ꍇ�Ō������@(����)���قȂ�B
'           ���f���T���v���h�c����������ꍇ�A��{�I�ɂ́A�V�T���v���ʒu���猩�āA�㉺�T���v���̒��ŋ߂��ق��̃T���v���h�c�𒊏o����B
'           ��������ۂ̌����͈͂́A�w�肳�ꂽ�͈͓��̂ݗL���Ƃ��A�����͈͓��ɂ݂���Ȃ��ꍇ�A�u�Y������قȂ��v�Ƃ���B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sBlockid      ,I  ,String       :��ۯ�ID
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
'                                                       = 10 CS
'                                                       = 11 GD     ���Ώ�
'                                                       = 12 LT
'                                                       = 13 EPD
'          :iFromPos      ,I  ,Integer      :�����͈�From
'          :iToPos        ,I  ,Integer      :�����͈�To
'          :sGetBlockid   ,O  ,String       :���f���u���b�N�h�c
'          :sGetSmpKbn    ,O  ,String       :���f���T���v���敪
'          :iGetSmplID    ,O  ,long         :���f���T���v���h�c     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O  ,Integer      :�擾���� = 0 : ����I��
'                                                       1 : ����I��(�Y���T���v���Ȃ�)
'                                                      -1 : �ُ�I��
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetSxlHanei1(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                              iItemNo As Integer, iFromPos As Integer, iToPos As Integer, _
                               sGetBlockid As String, sGetSmpKbn As String, iGetSmplID As Long) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       '�����ID����
    Dim ediInd      As String       '���FLG����
    Dim ediRes      As String       '����FLG����
    
    '�p�����[�^�`�F�b�N
    If (Len(sBlockId) <> 12) Then GoTo GetSxlHanei1ParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetSxlHanei1ParameterErr
    
    '�w�肳�ꂽ�]�����ڇ�����A�����Ώƕ]�����ږ������肷��B
    kName = funGetCryKensaName(iItemNo)
    If kName = " " Then GoTo GetSxlHanei1ParameterErr
    
    'SQL�����Ŏg�p���閼�̂ɕҏW
    ediSmpid = cCRYSMPLID & kName & cCS     '�����ID
    ediInd = cCRYIND & kName & cCS          '���FLG
    ediRes = cCRYRES & kName & cCS          '����FLG
    
    '�w�肳�ꂽ�������ɁA�V����يǗ�(��ۯ�)(XSDCS)����������B
    sql = "select CRYNUMCS, SMPKBNCS, " & ediSmpid & " as SMPLID from XSDCS "

    'TOP�ʒu(T/B�敪='T')�̌���
    If sTB = "T" Then
        sql = sql & "where tbkbncs = '" & sTB & "' and "
        sql = sql & "      xtalcs = '" & sCryNum & "' and "
        sql = sql & "      inposcs <= " & iSmplPos & " and "
        sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
        sql = sql & "  " & ediRes & " <> '0' and "
        sql = sql & "      inposcs >= " & iFromPos & " and "
        sql = sql & "      inposcs <= " & iToPos & " "
        sql = sql & "order by inposcs desc"
    
    'BOT�ʒu(T/B�敪='B')�̌���
    ElseIf sTB = "B" Then
        sql = sql & "where tbkbncs = '" & sTB & "' and "
        sql = sql & "      xtalcs = '" & sCryNum & "' and "
        sql = sql & "      inposcs >= " & iSmplPos & " and "
        sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
        sql = sql & "  " & ediRes & " <> '0' and "
        sql = sql & "      inposcs >= " & iFromPos & " and "
        sql = sql & "      inposcs <= " & iToPos & " "
        sql = sql & "order by inposcs asc"
    Else
        GoTo GetSxlHanei1ParameterErr
    End If
    
    'SQL���̎��s
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetSxlHanei1 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '�Ăяo�����ւ̌��ʒʒm
    sGetBlockid = rs("CRYNUMCS")
    sGetSmpKbn = rs("SMPKBNCS")
    iGetSmplID = rs("SMPLID")
    Set rs = Nothing
    
    funGetSxlHanei1 = 0
    Exit Function

GetSxlHanei1ParameterErr:
    funGetSxlHanei1 = -1
End Function

'------------------------------------------------
' �������f�l�擾�i�p�^�[���Q�j
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�V�T���v���ʒu��񂩂�A�������f���T���v���h�c��V�T���v���Ǘ�(��ۯ�)(XSDCS)��茟�����A���ʂ�Ԃ��B
'           ���f���T���v���h�c����������ꍇ�A��{�I�ɂ́A�V�T���v���ʒu���猩�āA���T���v���̒��ŋ߂��ق��̃T���v���h�c�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sBlockid      ,I  ,String       :��ۯ�ID
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
'                                                       = 10 CS     ���Ώ�
'                                                       = 11 GD
'                                                       = 12 LT
'                                                       = 13 EPD    ���Ώ�
'          :sGetBlockid   ,O  ,String       :���f���u���b�N�h�c
'          :sGetSmpKbn    ,O  ,String       :���f���T���v���敪
'          :iGetSmplID    ,O  ,Long         :���f���T���v���h�c     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O  ,Integer      :�擾���� = 0 : ����I��
'                                                       1 : ����I��(�Y���T���v���Ȃ�)
'                                                      -1 : �ُ�I��
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetSxlHanei2(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, sGetBlockid As String, sGetSmpKbn As String, iGetSmplID As Long) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       '�����ID����
    Dim ediInd      As String       '���FLG����
    Dim ediRes      As String       '����FLG����
    
    '�p�����[�^�`�F�b�N
    If (Len(sBlockId) <> 12) Then GoTo GetSxlHanei2ParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetSxlHanei2ParameterErr
    
    '�w�肳�ꂽ�]�����ڇ�����A�����Ώƕ]�����ږ������肷��B
    kName = funGetCryKensaName(iItemNo)
    If kName = " " Then GoTo GetSxlHanei2ParameterErr
    
    'SQL�����Ŏg�p���閼�̂ɕҏW
    ediSmpid = cCRYSMPLID & kName & cCS     '�����ID
    ediInd = cCRYIND & kName & cCS          '���FLG
    ediRes = cCRYRES & kName & cCS          '����FLG
    
    '�w�肳�ꂽ�������ɁA�V����يǗ�(��ۯ�)(XSDCS)����������B
    sql = "select CRYNUMCS, SMPKBNCS, " & ediSmpid & " as SMPLID from XSDCS "
'' 09/03/02 FAE)akiyama start
'    sql = sql & "where xtalcs = '" & sCryNum & "' and "
    sql = sql & "where CRYNUMCS LIKE '" & left(sCryNum, 9) & "%' and "
'' 09/03/02 FAE)akiyama end
    sql = sql & "      inposcs > " & iSmplPos & " and "
    sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
    sql = sql & "  " & ediRes & " <> '0' "
    sql = sql & "order by inposcs asc"
    
    'SQL���̎��s
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetSxlHanei2 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '�Ăяo�����ւ̌��ʒʒm
    sGetBlockid = rs("CRYNUMCS")
    sGetSmpKbn = rs("SMPKBNCS")
    iGetSmplID = rs("SMPLID")
    Set rs = Nothing
    
    funGetSxlHanei2 = 0
    Exit Function

GetSxlHanei2ParameterErr:
    funGetSxlHanei2 = -1
End Function


'------------------------------------------------
' �������f�l�擾 (�p�^�[��3)
'------------------------------------------------
'
'�T�v      :�w�肳�ꂽ�V�T���v���ʒu��񂩂�A�������f���T���v���h�c��V�T���v���Ǘ�(��ۯ�)(XSDCS)��茟�����A���ʂ�Ԃ��B
'           ���f���T���v���h�c����������ꍇ�A�������ň�Ԍ������d�l�����i�Ԃ̃T���v�����A
'            �V�T���v���ʒu���猩�āA���T���v���̒��ŋ߂��ق��̃T���v���h�c�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sBlockid      ,I  ,String       :��ۯ�ID
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
'                                                       = 10 CS
'                                                       = 11 GD
'                                                       = 12 LT     ���Ώ�
'                                                       = 13 EPD
'          :sGetBlockid   ,O  ,String       :���f���u���b�N�h�c
'          :sGetSmpKbn    ,O  ,String       :���f���T���v���敪
'          :iGetSmplID    ,O  ,Long         :���f���T���v���h�c     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O  ,Integer      :�擾���� = 0 : ����I��
'                                                       1 : ����I��(�Y���T���v���Ȃ�)
'                                                      -1 : �ُ�I��
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��
'
Public Function funGetSxlHanei3(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                iItemNo As Integer, sHsxLtspi As String, _
                                sGetBlockid As String, sGetSmpKbn As String, iGetSmplID As Long) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       '�����ID����
    Dim ediInd      As String       '���FLG����
    Dim ediRes      As String       '����FLG����
    Dim GETHINBAN   As tFullHinban  '��Ԍ������d�l�̕i��
    Dim LTsmpid    As String       '���т̌������ʃT���v��NO
    
    '�p�����[�^�`�F�b�N
    If (Len(sBlockId) <> 12) Then GoTo GetSxlHanei3ParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetSxlHanei3ParameterErr
    
    '�w�肳�ꂽ�]�����ڇ�����A�����Ώƕ]�����ږ������肷��B
    kName = funGetCryKensaName(iItemNo)
    If kName = " " Then GoTo GetSxlHanei3ParameterErr
    
    'SQL�����Ŏg�p���閼�̂ɕҏW
    ediSmpid = cCRYSMPLID & kName & cCS     '�����ID
    ediInd = cCRYIND & kName & cCS          '���FLG
    ediRes = cCRYRES & kName & cCS          '����FLG
    
    ''�d�l�������ꍇ�A���߂��ق��̻����ID�ɑ΂���i�Ԃ��擾����@2004/01/26 ooba START ===========>
    sql = sql & "select E019.HINBAN, E019.MNOREVNO, E019.FACTORY, E019.OPECOND, J007.SMPLNO "
    sql = sql & "from TBCME019 E019, TBCMJ007 J007, XSDCS CS "
    sql = sql & "where E019.HINBAN = J007.HINBAN "
    sql = sql & "and E019.MNOREVNO = J007.REVNUM "
    sql = sql & "and E019.FACTORY = J007.FACTORY "
    sql = sql & "and E019.OPECOND = J007.OPECOND "
    sql = sql & "and J007.CRYNUM = CS.XTALCS "
    sql = sql & "and J007.SMPLNO = CS." & ediSmpid & " "
    sql = sql & "and CRYNUM = '" & sCryNum & "' "
    sql = sql & "and J007.TRANCNT = (select max(TRANCNT) from TBCMJ007 "
    sql = sql & "where CRYNUM = J007.CRYNUM "
    sql = sql & "and SMPLNO = J007.SMPLNO) "
    sql = sql & "and CS." & ediInd & " = '1' "
    sql = sql & "and CS." & ediRes & " <> '0' "
    sql = sql & "and E019.HSXLTSPI != ' ' "
    sql = sql & "and E019.HSXLTSPI <= '" & sHsxLtspi & "' "
    sql = sql & "order by E019.HSXLTSPI asc, J007.POSITION asc "
    ''�d�l�������ꍇ�A���߂��ق��̻����ID�ɑ΂���i�Ԃ��擾����@2004/01/26 ooba END =============>
    
    'SQL���̎��s
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetSxlHanei3 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '�擾�i�Ԃ̐ݒ�
    GETHINBAN.hinban = rs("HINBAN")
    GETHINBAN.mnorevno = rs("MNOREVNO")
    GETHINBAN.factory = rs("FACTORY")
    GETHINBAN.opecond = rs("OPECOND")
    LTsmpid = rs("SMPLNO")              '�擾�i�Ԃ̃T���v��ID��ݒ肷��悤�ɕύX�@04/02/06 tuku
    
    Set rs = Nothing
    
    '�w�肳�ꂽ�������ɁA�V����يǗ�(��ۯ�)(XSDCS)����������B
    '�擾����J007�̎��т̃T���v��ID�����Ɍ�������悤�ɕύX 04/02/06 tuku
    sql = "select CRYNUMCS, SMPKBNCS, " & ediSmpid & " as SMPLID from XSDCS "
'' 09/03/02 FAE)akiyama start
'    sql = sql & "where XTALCS = '" & sCryNum & "' and "
    sql = sql & "where CRYNUMCS LIKE '" & left(sCryNum, 9) & "%' and "
'' 09/03/02 FAE)akiyama end
    sql = sql & "      INPOSCS > " & iSmplPos & " and "
    sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
    sql = sql & "  " & ediRes & " <> '0' and "
    sql = sql & "  " & ediSmpid & " = '" & LTsmpid & "'  "
    sql = sql & "order by inposcs asc"
    
    'SQL���̎��s
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetSxlHanei3 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '�Ăяo�����ւ̌��ʒʒm
    sGetBlockid = rs("CRYNUMCS")
    sGetSmpKbn = rs("SMPKBNCS")
    iGetSmplID = rs("SMPLID")
    Set rs = Nothing
    
    funGetSxlHanei3 = 0
    Exit Function

GetSxlHanei3ParameterErr:
    funGetSxlHanei3 = -1
End Function

'------------------------------------------------
' ��������l�擾
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�V����وʒu��񂩂�A�������茳�����ID1�ƌ������茳�����ID2��V����يǗ�(��ۯ�)(XSDCS)��茟�����A���ʂ�Ԃ��B
'           ���茳�����ID1�Ɛ��茳�����ID2����������ꍇ�A�������̍�TOP�ƍ�BOT�ʒu�̎����f�[�^��ΏۂƂ���B
'           �V����وʒu�̕i�Ԏd�l�ƌ������̍�TOP�^��BOT�ʒu�̕i�Ԏd�l���A���ꂼ��u3�_����v�u5�_����v�ł�����݂��l�����邪�A��������݂ɂ���Ă�����ۂ𔻒f����B
'           ��L�̑���_���p�^�[���̑g�ݍ��킹�ɂ��A�擾���ׂ�����_�f�[�^�̈ʒu(�ꏊ)���قȂ�B
'           ����ۂ̔��f�Ƃ��āAXSDC1��SUIFLGC1�̒l(0:���苖��,1:����֎~)���l������B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sBlockid      ,I  ,String       :��ۯ�ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :�����ԍ�
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :iSmplPos      ,I  ,Integer      :�V�T���v���ʒu(mm)
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� =  1 RS     ���Ώ�
'                                                       =  2 Oi
'                                                       =  3 BMD1
'                                                       =  4 BMD2
'                                                       =  5 BMD3
'                                                       =  6 OSF1
'                                                       =  7 OSF2
'                                                       =  8 OSF3
'                                                       =  9 OSF4
'                                                       = 10 CS
'                                                       = 11 GD
'                                                       = 12 LT
'                                                       = 13 EPD
'          :sGetBlockid1  ,O  ,String       :���茳�u���b�N�h�c�P
'          :sGetSmpKbn1   ,O  ,String       :���茳�T���v���敪�P
'          :iGetSmplID1   ,O  ,Long         :���茳�T���v���h�c�P   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iGetPCode1    ,O  ,String       :���茳�p�^�[���P('A' or 'B')
'          :iGetPos1      ,O  ,Integr       :���茳�T���v���ʒu1  2005/1/11
'          :sGetBlockid2  ,O  ,String       :���茳�u���b�N�h�c�Q
'          :sGetSmpKbn2   ,O  ,String       :���茳�T���v���敪�Q
'          :iGetSmplID2   ,O  ,Long         :���茳�T���v���h�c�Q   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iGetPCode2    ,O  ,String       :���茳�p�^�[���Q('A' or 'B')
'          :iGetPos2      ,O  ,Integr       :���茳�T���v���ʒu2  2005/1/11
'          :�߂�l        ,O  ,Integer      :�擾���� = 0 : ����I��
'                                                       1 : ����I��(�Y���T���v���Ȃ�)
'                                                      -1 : �ُ�I��
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��
'          :TEST2004/10 �����ǉ�,2005/1/11 �����ǉ�
Public Function funGetSuitei(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, iItemNo As Integer, _
                             sGetBlockid1 As String, sGetSmpKbn1 As String, iGetSmplID1 As Long, iGetPCode1 As String, _
                             sGetBlockid2 As String, sGetSmpKbn2 As String, iGetSmplID2 As Long, iGetPCode2 As String, _
                             iGetPos1 As Integer, iGetPos2 As Integer) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       '�����ID����
    Dim ediInd      As String       '���FLG����
    Dim ediRes      As String       '����FLG����
    Dim getNewSpec  As String       '�V����وʒu���R�d�l�l
    Dim getTopBlkID As String       'TOP�ʒu��ۯ�ID
    Dim getTopSmpK  As String       'TOP�ʒu����ً敪
    Dim getTopSmpID As String       'TOP�ʒu�����ID
    Dim getTopHin   As tFullHinban  'TOP�ʒu�i��
    Dim getTopSpec  As String       'TOP�ʒu���R�d�l�l
    Dim getTopPtrn  As String       'TOP�ʒu����ݺ���
    Dim getTopPos   As Integer      'TOP�ʒu�������ʒu
    
    Dim getBotBlkID As String       'BOT�ʒu��ۯ�ID
    Dim getBotSmpK  As String       'BOT�ʒu����ً敪
    Dim getBotSmpID As String       'BOT�ʒu�����ID
    Dim getBotHin   As tFullHinban  'BOT�ʒu�i��
    Dim getBotSpec  As String       'BOT�ʒu���R�d�l�l
    Dim getBotPtrn  As String       'BOT�ʒu����ݺ���
    Dim getBotPos   As Integer      'BOT�ʒu�������ʒu
    '�p�����[�^�`�F�b�N
    If (Len(sBlockId) <> 12) Then GoTo GetSuiteiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetSuiteiParameterErr
    
    '�w�肳�ꂽ�]�����ڇ�����A�����Ώƕ]�����ږ������肷��B
    kName = funGetCryKensaName(iItemNo)
    If kName = " " Then GoTo GetSuiteiParameterErr
    
    'SQL�����Ŏg�p���閼�̂ɕҏW
    ediSmpid = cCRYSMPLID & kName & cCS     '�����ID
    ediInd = cCRYIND & kName & cCS          '���FLG
    ediRes = cCRYRES & kName & "1" & cCS    '����FLG
    
    '�w�肳�ꂽ�������ɁA�V����يǗ�(��ۯ�)(XSDCS)����������B
    '�ᐄ�茳�T���v���h�c�P(TOP�ʒu)�̎擾��
    sql = "select CS.CRYNUMCS, CS.SMPKBNCS, CS." & ediSmpid & " as SMPLID, CS.HINBCS, CS.REVNUMCS, CS.FACTORYCS, CS.OPECS, CS.INPOSCS "  '2005/1/11
    sql = sql & "from XSDCS CS, XSDC1 C1 "
    sql = sql & "where CS.XTALCS = '" & sCryNum & "' and "
    sql = sql & "      CS.INPOSCS < " & iSmplPos & " and "
    sql = sql & "      CS." & ediInd & " = '1' and "
    sql = sql & "      CS." & ediInd & " <> '0' and "
    sql = sql & "      CS." & ediRes & " <> '0' and "
    sql = sql & "      C1.XTALC1 = CS.XTALCS and "
    sql = sql & "      C1.SUIFLG = '0' "
    sql = sql & "order by CS.INPOSCS desc"
    
    'SQL���̎��s
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        Set rs = Nothing
        GoTo GetSuiteiEmpty
    End If
    
    'TOP�ʒu�f�[�^�̐ݒ�
    getTopBlkID = rs("CRYNUMCS")            'TOP�ʒu��ۯ�ID
    getTopSmpK = rs("SMPKBNCS")             'TOP�ʒu����ً敪
    getTopSmpID = rs("SMPLID")                'TOP�ʒu�����ID
    getTopHin.hinban = rs("HINBCS")         'TOP�ʒu�i��
    getTopHin.mnorevno = rs("REVNUMCS")     'TOP�ʒu���i�ԍ������ԍ�
    getTopHin.factory = rs("FACTORYCS")     'TOP�ʒu�H��
    getTopHin.opecond = rs("OPECS")         'TOP�ʒu���Ə���
    getTopPos = rs("INPOSCS")               'TOP�ʒu   2005/1/11
    Set rs = Nothing
    
    '�ᐄ�茳�T���v���h�c�Q(BOT�ʒu)�̎擾��
    sql = "select CS.CRYNUMCS, CS.SMPKBNCS, CS." & ediSmpid & " as SMPLID, CS.HINBCS, CS.REVNUMCS, CS.FACTORYCS, CS.OPECS, CS.INPOSCS "   '2005/1/11
    sql = sql & "from XSDCS CS, XSDC1 C1 "
    sql = sql & "where CS.XTALCS = '" & sCryNum & "' and "
    sql = sql & "      CS.INPOSCS > " & iSmplPos & " and "
    sql = sql & "      CS." & ediInd & " = '1' and "
    sql = sql & "      CS." & ediInd & " <> '0' and "
    sql = sql & "      CS." & ediRes & " <> '0' and "
    sql = sql & "      C1.XTALC1 = CS.XTALCS and "
    sql = sql & "      C1.SUIFLG = '0' "
    sql = sql & "order by CS.INPOSCS asc"
    
    'SQL���̎��s
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        Set rs = Nothing
        GoTo GetSuiteiEmpty
    End If
    
    'BOT�ʒu�f�[�^�̐ݒ�
    getBotBlkID = rs("CRYNUMCS")            'BOT�ʒu��ۯ�ID
    getBotSmpK = rs("SMPKBNCS")             'BOT�ʒu����ً敪
    getBotSmpID = rs("SMPLID")                'BOT�ʒu�����ID
    getBotHin.hinban = rs("HINBCS")         'BOT�ʒu�i��
    getBotHin.mnorevno = rs("REVNUMCS")     'BOT�ʒu���i�ԍ������ԍ�
    getBotHin.factory = rs("FACTORYCS")     'BOT�ʒu�H��
    getBotHin.opecond = rs("OPECS")         'BOT�ʒu���Ə���
    getBotPos = rs("INPOSCS")               '2005/1/11
    Set rs = Nothing
    
    '�e�i�Ԃ̔��R�d�l�l�擾
    getTopPtrn = "A"
    getBotPtrn = "A"
    '------------------------------------------------------
    
    '�Ăяo�����ւ̌��ʒʒm
    sGetBlockid1 = getTopBlkID      '���茳�u���b�N�h�c�P
    sGetSmpKbn1 = getTopSmpK        '���茳�T���v���敪�P
    iGetSmplID1 = getTopSmpID       '���茳�T���v���h�c�P
    iGetPCode1 = getTopPtrn         '���茳�p�^�[���P('A' or 'B')
    iGetPos1 = getTopPos            '���茳�ʒu  2005/1/11
    
    sGetBlockid2 = getBotBlkID      '���茳�u���b�N�h�c�Q
    sGetSmpKbn2 = getBotSmpK        '���茳�T���v���敪�Q
    iGetSmplID2 = getBotSmpID       '���茳�T���v���h�c�Q
    iGetPCode2 = getBotPtrn         '���茳�p�^�[���Q('A' or 'B')
    iGetPos2 = getBotPos            '���茳�ʒu  2005/1/11
    
    funGetSuitei = 0
    Exit Function

GetSuiteiEmpty:
    funGetSuitei = 1
    Exit Function

GetSuiteiParameterErr:
    funGetSuitei = -1
End Function

'------------------------------------------------
' �������� ���R�d�l�l�擾�֐�
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�i�Ԃ���ATBCME018���������A���R�d�l�l(�iSX���R����ʒu_��)���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :�߂�l        ,O  ,Sting        :���R�d�l�l(�iSX���R����ʒu_��)
'                                            (�擾�ł��Ȃ��ꍇ�́A�󔒂�Ԃ�)
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetSuiSpecRS(tFullHin As tFullHinban) As String
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    
    '�w�肳�ꂽ�i�Ԃ���TBCME018��HSXRSPOI(�iSX���R����ʒu_��)����������B
    sql = "select HSXRSPOI from TBCME018 "
    sql = sql & "where HINBAN = '" & Trim(tFullHin.hinban) & "' and "
    sql = sql & "      MNOREVNO = " & tFullHin.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tFullHin.factory & "' and "
    sql = sql & "      OPECOND = '" & tFullHin.opecond & "'"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetSuiSpecRS = " "
        Set rs = Nothing
        Exit Function
    End If
    
    'TOP�ʒu�f�[�^�̐ݒ�
    funGetSuiSpecRS = rs("HSXRSPOI")
    Set rs = Nothing

End Function

'------------------------------------------------
' ���������Ώە]�����ږ��擾
'------------------------------------------------

'�T�v      :�]�����ڇ�����A���������Ώە]�����ږ���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� �� ��ۯ� =  1 RS
'                                                                =  2 Oi
'                                                                =  3 BMD1
'                                                                =  4 BMD2
'                                                                =  5 BMD3
'                                                                =  6 OSF1
'                                                                =  7 OSF2
'                                                                =  8 OSF3
'                                                                =  9 OSF4
'                                                                = 10 CS
'                                                                = 11 GD
'                                                                = 12 LT
'                                                                = 13 EPD
'          :�߂�l        ,O  ,Sting        :�����Ώۍ��ږ�(���Ұ��װ���́A�󔒂�Ԃ�)
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetCryKensaName(iItemNo As Integer) As String
    
    '�p�����[�^�`�F�b�N
    'If iItemNo < 1 Or iItemNo > 13 Then GoTo GetCryKensaNameParameterErr    'Chg 2011/01/19 SMPK Miyata
    If iItemNo < 1 Or iItemNo > 18 Then GoTo GetCryKensaNameParameterErr

    '��ۯ�
    Select Case iItemNo
    Case 1:     funGetCryKensaName = cCRY_RS       'RS(���R)
    Case 2:     funGetCryKensaName = cCRY_OI       'Oi(�_�f�Z�x)
    Case 3:     funGetCryKensaName = cCRY_B1       'BMD1
    Case 4:     funGetCryKensaName = cCRY_B2       'BMD2
    Case 5:     funGetCryKensaName = cCRY_B3       'BMD3
    Case 6:     funGetCryKensaName = cCRY_O1       'OSF1
    Case 7:     funGetCryKensaName = cCRY_O2       'OSF2
    Case 8:     funGetCryKensaName = cCRY_O3       'OSF3
    Case 9:     funGetCryKensaName = cCRY_O4       'OSF4
    Case 10:    funGetCryKensaName = cCRY_CS       'CS(�Y�f�Z�x)
    Case 11:    funGetCryKensaName = cCRY_GD       'GD
    Case 12:    funGetCryKensaName = cCRY_LT       'LT(ײ����)
    Case 13:    funGetCryKensaName = cCRY_EP       'EPD
'Add Start 2011/01/19 SMPK Miyata
    Case 15:    funGetCryKensaName = cCRY_C        'C
    Case 16:    funGetCryKensaName = cCRY_CJ       'CJ
    Case 17:    funGetCryKensaName = cCRY_CJLT     'CJLT
    Case 18:    funGetCryKensaName = cCRY_CJ2      'CJ2
'Add End   2011/01/19 SMPK Miyata
    End Select
    
    Exit Function

GetCryKensaNameParameterErr:
    funGetCryKensaName = " "
End Function

'------------------------------------------------
' ������R���ю擾�֐�
'------------------------------------------------

'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ002���������A������R���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :tCryRs        ,O  ,type_DBDRV_scmzc_fcmkc001c_CryR  :������R����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetCryRsJisseki(sCryNum As String, iSmplID As Long, tCryRs As type_DBDRV_scmzc_fcmkc001c_CryR) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Dim wkXsdcs     As typ_XSDCS
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryRsJisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ002�̌�����R���ђl����������B
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, REGDATE, KSTAFFID "    '----TEST2004/10
    sql = sql & "from TBCMJ002 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryRsJisseki = -1
        GoTo proc_exit
    End If
    
     ''���o���ʂ��i�[����
    With tCryRs
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCOND = rs("TRANCOND")          ' ��������
        .TRANCNT = rs("TRANCNT")            ' ������
        .SMPLNO = rs("SMPLNO")              ' �T���v���m��
        .SMPLUMU = rs("SMPLUMU")            ' �T���v���L��
        .MEAS1 = rs("MEAS1")                ' ����l�P
        .MEAS2 = rs("MEAS2")                ' ����l�Q
        .MEAS3 = rs("MEAS3")                ' ����l�R
        .MEAS4 = rs("MEAS4")                ' ����l�S
        .MEAS5 = rs("MEAS5")                ' ����l�T
        .EFEHS = rs("EFEHS")                ' �����ΐ�
        .RRG = rs("RRG")                    ' �q�q�f
        .REGDATE = rs("REGDATE")            ' �o�^���t
        .KSTAFFID = rs("KSTAFFID")          '----TEST2004/10
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        ' DK���x�i���сj
        wkXsdcs.XTALCS = .CRYNUM
        wkXsdcs.CRYSMPLIDRSCS = .SMPLNO
        wkXsdcs.CRYINDRSCS = "0"
        .HSXDKTMP = GetDKTmpCode(False, wkXsdcs)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    End With
    
    funGetCryRsJisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ����Oi���ю擾�֐�
'------------------------------------------------

'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ003���������A����Oi���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :tCryOi        ,O  ,type_DBDRV_scmzc_fcmkc001c_Oi    :����Oi����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetCryOiJisseki(sCryNum As String, iSmplID As Long, tCryOi As type_DBDRV_scmzc_fcmkc001c_Oi) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryOiJisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ003�̌���Oi���ђl����������B
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, AVE, FTIRCONV, INSPECTWAY, REGDATE "
    sql = sql & "from TBCMJ003 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "  and TRANCOND = 0 "       'GFA��FTIR���Z�l�\���ُ�Ή� 2011/01/20�ǉ� SETsw kubota
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryOiJisseki = -1
        GoTo proc_exit
    End If
    
     ''���o���ʂ��i�[����
    With tCryOi
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCOND = rs("TRANCOND")          ' ��������
        .TRANCNT = rs("TRANCNT")            ' ������
        .SMPLNO = rs("SMPLNO")              ' �T���v���m��
        .SMPLUMU = rs("SMPLUMU")            ' �T���v���L��
'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
        If IsNull(rs("OIMEAS1")) = False Then .OIMEAS1 = rs("OIMEAS1") Else .OIMEAS1 = -1  '�n������l1
        If IsNull(rs("OIMEAS2")) = False Then .OIMEAS2 = rs("OIMEAS2") Else .OIMEAS2 = -1  '�n������l2
        If IsNull(rs("OIMEAS3")) = False Then .OIMEAS3 = rs("OIMEAS3") Else .OIMEAS3 = -1  '�n������l3
        If IsNull(rs("OIMEAS4")) = False Then .OIMEAS4 = rs("OIMEAS4") Else .OIMEAS4 = -1  '�n������l4
        If IsNull(rs("OIMEAS5")) = False Then .OIMEAS5 = rs("OIMEAS5") Else .OIMEAS5 = -1  '�n������l5
        If IsNull(rs("ORGRES")) = False Then .ORGRES = rs("ORGRES") Else .ORGRES = -1    ' �n�q�f����
'OI_NULL�Ή��@2005/03/08 TUKU END   --------------------------------------------------
        .AVE = rs("AVE")                    ' AVE
        .FTIRCONV = rs("FTIRCONV")          ' FTIR���Z
        .INSPECTWAY = rs("INSPECTWAY")      ' �������@
        .REGDATE = rs("REGDATE")            ' �o�^���t
    End With
    
    funGetCryOiJisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ����BMD���ю擾�֐�
'------------------------------------------------

'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ008���������A����BMD���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iTranCond     ,I  ,Integer                          :��������(1:BMD1, 2:BMD2, 3:BMD3)
'          :tCryBMD       ,O  ,type_DBDRV_scmzc_fcmkc001c_BMD   :����BMD����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetCryBMDJisseki(sCryNum As String, iSmplID As Long, iTranCond As Integer, tCryBMD As type_DBDRV_scmzc_fcmkc001c_BMD) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryBMDJisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ008�̌���BMD���ђl����������B
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "HTPRC, KKSP, KKSET, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN, MEASMAX, MEASAVE, BMDMNBUNP, REGDATE "
    sql = sql & "from TBCMJ008 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      TRANCOND = '" & iTranCond & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryBMDJisseki = -1
        GoTo proc_exit
    End If
    
     ''���o���ʂ��i�[����
    With tCryBMD
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCOND = rs("TRANCOND")          ' ��������
        .TRANCNT = rs("TRANCNT")            ' ������
        .SMPLNO = rs("SMPLNO")              ' �T���v���m��
        .SMPLUMU = rs("SMPLUMU")            ' �T���v���L��
        .HTPRC = rs("HTPRC")                ' �M�������@
        .KKSP = rs("KKSP")                  ' �������ב���ʒu
        .KKSET = rs("KKSET")                ' �������ב�������{�I��ET��@�@char(1)�{number(2)
        .MEAS1 = rs("MEAS1")                ' ����l1
        .MEAS2 = rs("MEAS2")                ' ����l2
        .MEAS3 = rs("MEAS3")                ' ����l3
        .MEAS4 = rs("MEAS4")                ' ����l4
        .MEAS5 = rs("MEAS5")                ' ����l5
        .MEASMIN = rs("MEASMIN")            ' Min
        .MEASMAX = rs("MEASMAX")            ' max
        .MEASAVE = rs("MEASAVE")            ' AVE
        If Not IsNull(rs("BMDMNBUNP")) Then .BMDMNBUNP = rs("BMDMNBUNP")      ' BMD�ʓ����z
        .REGDATE = rs("REGDATE")            ' �o�^���t
    End With
    
    funGetCryBMDJisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ����OSF���ю擾�֐�
'------------------------------------------------

'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ005���������A����OSF���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iTranCond     ,I  ,Integer                          :��������(1:OSF1, 2:OSF2, 3:OSF3, 4:OSF4)
'          :tCryOSF       ,O  ,type_DBDRV_scmzc_fcmkc001c_OSF   :����OSF����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetCryOSFJisseki(sCryNum As String, iSmplID As Long, iTranCond As Integer, tCryOSF As type_DBDRV_scmzc_fcmkc001c_OSF) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryOSFJisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ005�̌���OSF���ђl����������B
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "HTPRC, KKSP, KKSET, CALCMAX, CALCAVE, "
    sql = sql & "MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEAS6, MEAS7, MEAS8, MEAS9, MEAS10, "
    sql = sql & "MEAS11, MEAS12, MEAS13, MEAS14, MEAS15, MEAS16, MEAS17, MEAS18, MEAS19, MEAS20, "
    sql = sql & "OSFPOS1, OSFWID1, OSFRD1, OSFPOS2, OSFWID2, OSFRD2, OSFPOS3, OSFWID3, OSFRD3, REGDATE "
    sql = sql & ",CALCMH "  '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    sql = sql & "from TBCMJ005 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      TRANCOND = '" & iTranCond & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryOSFJisseki = -1
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
    With tCryOSF
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCOND = rs("TRANCOND")          ' ��������
        .TRANCNT = rs("TRANCNT")            ' ������
        .SMPLNO = rs("SMPLNO")              ' �T���v���m��
        .SMPLUMU = rs("SMPLUMU")            ' �T���v���L��
        .HTPRC = rs("HTPRC")                ' �M�������@
        .KKSP = rs("KKSP")                  ' �������ב���ʒu
        .KKSET = rs("KKSET")                ' �������ב�������{�I��ET��@�@char(1)�{number(2)
        .CALCMAX = rs("CALCMAX")            ' �v�Z���� Max
        .CALCAVE = rs("CALCAVE")            ' �v�Z���� Ave
        .MEAS1 = rs("MEAS1")                ' ����l1
        .MEAS2 = rs("MEAS2")                ' ����l2
        .MEAS3 = rs("MEAS3")                ' ����l3
        .MEAS4 = rs("MEAS4")                ' ����l4
        .MEAS5 = rs("MEAS5")                ' ����l5
        .MEAS6 = rs("MEAS6")                ' ����l6
        .MEAS7 = rs("MEAS7")                ' ����l7
        .MEAS8 = rs("MEAS8")                ' ����l8
        .MEAS9 = rs("MEAS9")                ' ����l9
        .MEAS10 = rs("MEAS10")              ' ����l10
        .MEAS11 = rs("MEAS11")              ' ����l11
        .MEAS12 = rs("MEAS12")              ' ����l12
        .MEAS13 = rs("MEAS13")              ' ����l13
        .MEAS14 = rs("MEAS14")              ' ����l14
        .MEAS15 = rs("MEAS15")              ' ����l15
        .MEAS16 = rs("MEAS16")              ' ����l16
        .MEAS17 = rs("MEAS17")              ' ����l17
        .MEAS18 = rs("MEAS18")              ' ����l18
        .MEAS19 = rs("MEAS19")              ' ����l19
        .MEAS20 = rs("MEAS20")              ' ����l20
        If Not IsNull(rs("OSFPOS1")) Then .OSFPOS1 = rs("OSFPOS1")      ' �p�^�[���敪1�ʒu
        If Not IsNull(rs("OSFWID1")) Then .OSFWID1 = rs("OSFWID1")      ' �p�^�[���敪1��
        If Not IsNull(rs("OSFRD1")) Then .OSFRD1 = rs("OSFRD1")         ' �p�^�[���敪1R / D
        If Not IsNull(rs("OSFPOS2")) Then .OSFPOS2 = rs("OSFPOS2")      ' �p�^�[���敪2�ʒu
        If Not IsNull(rs("OSFWID2")) Then .OSFWID2 = rs("OSFWID2")      ' �p�^�[���敪2��
        If Not IsNull(rs("OSFRD2")) Then .OSFRD2 = rs("OSFRD2")         ' �p�^�[���敪2R / D
        If Not IsNull(rs("OSFPOS3")) Then .OSFPOS3 = rs("OSFPOS3")      ' �p�^�[���敪3�ʒu
        If Not IsNull(rs("OSFWID3")) Then .OSFWID3 = rs("OSFWID3")      ' �p�^�[���敪3��
        If Not IsNull(rs("OSFRD3")) Then .OSFRD3 = rs("OSFRD3")         ' �p�^�[���敪3R / D
        
        .CALCMH = fncNullCheck(rs("CALCMH"))    ' �ʓ���(MAX/MIN)   '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
        
        .REGDATE = rs("REGDATE")            ' �o�^���t
    End With
    
    funGetCryOSFJisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ����CS���ю擾�֐�
'------------------------------------------------

'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ004���������A����CS���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :tCryCS        ,O  ,type_DBDRV_scmzc_fcmkc001c_CS    :����CS����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetCryCSJisseki(sCryNum As String, iSmplID As Long, tCryCS As type_DBDRV_scmzc_fcmkc001c_CS) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryCSJisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ004�̌���CS���ђl����������B
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "CSMEAS, PRE70P, INSPECTWAY, REGDATE "
    sql = sql & "from TBCMJ004 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryCSJisseki = -1
        GoTo proc_exit
    End If
    
     ''���o���ʂ��i�[����
    With tCryCS
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCOND = rs("TRANCOND")          ' ��������
        .TRANCNT = rs("TRANCNT")            ' ������
        .SMPLNO = rs("SMPLNO")              ' �T���v���m��
        .SMPLUMU = rs("SMPLUMU")            ' �T���v���L��
'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
            If IsNull(rs("CSMEAS")) = False Then .CSMEAS = rs("CSMEAS") Else .CSMEAS = -1  ' Cs�����l
            If IsNull(rs("PRE70P")) = False Then .PRE70P = rs("PRE70P") Else .PRE70P = -1  ' �V�O������l
'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
        .INSPECTWAY = rs("INSPECTWAY")      ' �������@
        .REGDATE = rs("REGDATE")            ' �o�^���t
    End With
    
    funGetCryCSJisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ����GD���ю擾�֐�
'------------------------------------------------

'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ006���������A����GD���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :tCryGD        ,O  ,type_DBDRV_scmzc_fcmkc001c_GD    :����GD����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetCryGDJisseki(sCryNum As String, iSmplID As Long, tCryGD As type_DBDRV_scmzc_fcmkc001c_GD) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryGDJisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ006�̌���GD���ђl����������B
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "MSRSDEN, MSRSLDL, MSRSDVD2, "
    sql = sql & "MS01LDL1, MS01LDL2, MS01LDL3, MS01LDL4, MS01LDL5, MS01DEN1, MS01DEN2, MS01DEN3, MS01DEN4, MS01DEN5, "
    sql = sql & "MS02LDL1, MS02LDL2, MS02LDL3, MS02LDL4, MS02LDL5, MS02DEN1, MS02DEN2, MS02DEN3, MS02DEN4, MS02DEN5, "
    sql = sql & "MS03LDL1, MS03LDL2, MS03LDL3, MS03LDL4, MS03LDL5, MS03DEN1, MS03DEN2, MS03DEN3, MS03DEN4, MS03DEN5, "
    sql = sql & "MS04LDL1, MS04LDL2, MS04LDL3, MS04LDL4, MS04LDL5, MS04DEN1, MS04DEN2, MS04DEN3, MS04DEN4, MS04DEN5, "
    sql = sql & "MS05LDL1, MS05LDL2, MS05LDL3, MS05LDL4, MS05LDL5, MS05DEN1, MS05DEN2, MS05DEN3, MS05DEN4, MS05DEN5, "
    sql = sql & "MS06LDL1, MS06LDL2, MS06LDL3, MS06LDL4, MS06LDL5, MS06DEN1, MS06DEN2, MS06DEN3, MS06DEN4, MS06DEN5, "
    sql = sql & "MS07LDL1, MS07LDL2, MS07LDL3, MS07LDL4, MS07LDL5, MS07DEN1, MS07DEN2, MS07DEN3, MS07DEN4, MS07DEN5, "
    sql = sql & "MS08LDL1, MS08LDL2, MS08LDL3, MS08LDL4, MS08LDL5, MS08DEN1, MS08DEN2, MS08DEN3, MS08DEN4, MS08DEN5, "
    sql = sql & "MS09LDL1, MS09LDL2, MS09LDL3, MS09LDL4, MS09LDL5, MS09DEN1, MS09DEN2, MS09DEN3, MS09DEN4, MS09DEN5, "
    sql = sql & "MS10LDL1, MS10LDL2, MS10LDL3, MS10LDL4, MS10LDL5, MS10DEN1, MS10DEN2, MS10DEN3, MS10DEN4, MS10DEN5, "
    sql = sql & "MS11LDL1, MS11LDL2, MS11LDL3, MS11LDL4, MS11LDL5, MS11DEN1, MS11DEN2, MS11DEN3, MS11DEN4, MS11DEN5, "
    sql = sql & "MS12LDL1, MS12LDL2, MS12LDL3, MS12LDL4, MS12LDL5, MS12DEN1, MS12DEN2, MS12DEN3, MS12DEN4, MS12DEN5, "
    sql = sql & "MS13LDL1, MS13LDL2, MS13LDL3, MS13LDL4, MS13LDL5, MS13DEN1, MS13DEN2, MS13DEN3, MS13DEN4, MS13DEN5, "
    sql = sql & "MS14LDL1, MS14LDL2, MS14LDL3, MS14LDL4, MS14LDL5, MS14DEN1, MS14DEN2, MS14DEN3, MS14DEN4, MS14DEN5, "
    sql = sql & "MS15LDL1, MS15LDL2, MS15LDL3, MS15LDL4, MS15LDL5, MS15DEN1, MS15DEN2, MS15DEN3, MS15DEN4, MS15DEN5, "
    sql = sql & "MS01DVD2, MS02DVD2, MS03DVD2, MS04DVD2, MS05DVD2, REGDATE "
    
    sql = sql & ", MSZEROMN, MSZEROMX " '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    
    sql = sql & "from TBCMJ006 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryGDJisseki = -1
        GoTo proc_exit
    End If
    
     ''���o���ʂ��i�[����
    With tCryGD
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCOND = rs("TRANCOND")          ' ��������
        .TRANCNT = rs("TRANCNT")            ' ������
        .SMPLNO = rs("SMPLNO")              ' �T���v���m��
        .SMPLUMU = rs("SMPLUMU")            ' �T���v���L��
        .MSRSDEN = rs("MSRSDEN")            ' ���茋�� Den
        .MSRSLDL = rs("MSRSLDL")            ' ���茋�� L/DL
        .MSRSDVD2 = rs("MSRSDVD2")          ' ���茋�� DVD2
        .MS01LDL1 = rs("MS01LDL1")          ' ����l01 L/DL1
        .MS01LDL2 = rs("MS01LDL2")          ' ����l01 L/DL2
        .MS01LDL3 = rs("MS01LDL3")          ' ����l01 L/DL3
        .MS01LDL4 = rs("MS01LDL4")          ' ����l01 L/DL4
        .MS01LDL5 = rs("MS01LDL5")          ' ����l01 L/DL5
        .MS01DEN1 = rs("MS01DEN1")          ' ����l01 Den1
        .MS01DEN2 = rs("MS01DEN2")          ' ����l01 Den2
        .MS01DEN3 = rs("MS01DEN3")          ' ����l01 Den3
        .MS01DEN4 = rs("MS01DEN4")          ' ����l01 Den4
        .MS01DEN5 = rs("MS01DEN5")          ' ����l01 Den5
        .MS02LDL1 = rs("MS02LDL1")          ' ����l02 L/DL1
        .MS02LDL2 = rs("MS02LDL2")          ' ����l02 L/DL2
        .MS02LDL3 = rs("MS02LDL3")          ' ����l02 L/DL3
        .MS02LDL4 = rs("MS02LDL4")          ' ����l02 L/DL4
        .MS02LDL5 = rs("MS02LDL5")          ' ����l02 L/DL5
        .MS02DEN1 = rs("MS02DEN1")          ' ����l02 Den1
        .MS02DEN2 = rs("MS02DEN2")          ' ����l02 Den2
        .MS02DEN3 = rs("MS02DEN3")          ' ����l02 Den3
        .MS02DEN4 = rs("MS02DEN4")          ' ����l02 Den4
        .MS02DEN5 = rs("MS02DEN5")          ' ����l02 Den5
        .MS03LDL1 = rs("MS03LDL1")          ' ����l03 L/DL1
        .MS03LDL2 = rs("MS03LDL2")          ' ����l03 L/DL2
        .MS03LDL3 = rs("MS03LDL3")          ' ����l03 L/DL3
        .MS03LDL4 = rs("MS03LDL4")          ' ����l03 L/DL4
        .MS03LDL5 = rs("MS03LDL5")          ' ����l03 L/DL5
        .MS03DEN1 = rs("MS03DEN1")          ' ����l03 Den1
        .MS03DEN2 = rs("MS03DEN2")          ' ����l03 Den2
        .MS03DEN3 = rs("MS03DEN3")          ' ����l03 Den3
        .MS03DEN4 = rs("MS03DEN4")          ' ����l03 Den4
        .MS03DEN5 = rs("MS03DEN5")          ' ����l03 Den5
        .MS04LDL1 = rs("MS04LDL1")          ' ����l04 L/DL1
        .MS04LDL2 = rs("MS04LDL2")          ' ����l04 L/DL2
        .MS04LDL3 = rs("MS04LDL3")          ' ����l04 L/DL3
        .MS04LDL4 = rs("MS04LDL4")          ' ����l04 L/DL4
        .MS04LDL5 = rs("MS04LDL5")          ' ����l04 L/DL5
        .MS04DEN1 = rs("MS04DEN1")          ' ����l04 Den1
        .MS04DEN2 = rs("MS04DEN2")          ' ����l04 Den2
        .MS04DEN3 = rs("MS04DEN3")          ' ����l04 Den3
        .MS04DEN4 = rs("MS04DEN4")          ' ����l04 Den4
        .MS04DEN5 = rs("MS04DEN5")          ' ����l04 Den5
        .MS05LDL1 = rs("MS05LDL1")          ' ����l05 L/DL1
        .MS05LDL2 = rs("MS05LDL2")          ' ����l05 L/DL2
        .MS05LDL3 = rs("MS05LDL3")          ' ����l05 L/DL3
        .MS05LDL4 = rs("MS05LDL4")          ' ����l05 L/DL4
        .MS05LDL5 = rs("MS05LDL5")          ' ����l05 L/DL5
        .MS05DEN1 = rs("MS05DEN1")          ' ����l05 Den1
        .MS05DEN2 = rs("MS05DEN2")          ' ����l05 Den2
        .MS05DEN3 = rs("MS05DEN3")          ' ����l05 Den3
        .MS05DEN4 = rs("MS05DEN4")          ' ����l05 Den4
        .MS05DEN5 = rs("MS05DEN5")          ' ����l05 Den5
        .MS06LDL1 = rs("MS06LDL1")          ' ����l06 L/DL1
        .MS06LDL2 = rs("MS06LDL2")          ' ����l06 L/DL2
        .MS06LDL3 = rs("MS06LDL3")          ' ����l06 L/DL3
        .MS06LDL4 = rs("MS06LDL4")          ' ����l06 L/DL4
        .MS06LDL5 = rs("MS06LDL5")          ' ����l06 L/DL5
        .MS06DEN1 = rs("MS06DEN1")          ' ����l06 Den1
        .MS06DEN2 = rs("MS06DEN2")          ' ����l06 Den2
        .MS06DEN3 = rs("MS06DEN3")          ' ����l06 Den3
        .MS06DEN4 = rs("MS06DEN4")          ' ����l06 Den4
        .MS06DEN5 = rs("MS06DEN5")          ' ����l06 Den5
        .MS07LDL1 = rs("MS07LDL1")          ' ����l07 L/DL1
        .MS07LDL2 = rs("MS07LDL2")          ' ����l07 L/DL2
        .MS07LDL3 = rs("MS07LDL3")          ' ����l07 L/DL3
        .MS07LDL4 = rs("MS07LDL4")          ' ����l07 L/DL4
        .MS07LDL5 = rs("MS07LDL5")          ' ����l07 L/DL5
        .MS07DEN1 = rs("MS07DEN1")          ' ����l07 Den1
        .MS07DEN2 = rs("MS07DEN2")          ' ����l07 Den2
        .MS07DEN3 = rs("MS07DEN3")          ' ����l07 Den3
        .MS07DEN4 = rs("MS07DEN4")          ' ����l07 Den4
        .MS07DEN5 = rs("MS07DEN5")          ' ����l07 Den5
        .MS08LDL1 = rs("MS08LDL1")          ' ����l08 L/DL1
        .MS08LDL2 = rs("MS08LDL2")          ' ����l08 L/DL2
        .MS08LDL3 = rs("MS08LDL3")          ' ����l08 L/DL3
        .MS08LDL4 = rs("MS08LDL4")          ' ����l08 L/DL4
        .MS08LDL5 = rs("MS08LDL5")          ' ����l08 L/DL5
        .MS08DEN1 = rs("MS08DEN1")          ' ����l08 Den1
        .MS08DEN2 = rs("MS08DEN2")          ' ����l08 Den2
        .MS08DEN3 = rs("MS08DEN3")          ' ����l08 Den3
        .MS08DEN4 = rs("MS08DEN4")          ' ����l08 Den4
        .MS08DEN5 = rs("MS08DEN5")          ' ����l08 Den5
        .MS09LDL1 = rs("MS09LDL1")          ' ����l09 L/DL1
        .MS09LDL2 = rs("MS09LDL2")          ' ����l09 L/DL2
        .MS09LDL3 = rs("MS09LDL3")          ' ����l09 L/DL3
        .MS09LDL4 = rs("MS09LDL4")          ' ����l09 L/DL4
        .MS09LDL5 = rs("MS09LDL5")          ' ����l09 L/DL5
        .MS09DEN1 = rs("MS09DEN1")          ' ����l09 Den1
        .MS09DEN2 = rs("MS09DEN2")          ' ����l09 Den2
        .MS09DEN3 = rs("MS09DEN3")          ' ����l09 Den3
        .MS09DEN4 = rs("MS09DEN4")          ' ����l09 Den4
        .MS09DEN5 = rs("MS09DEN5")          ' ����l09 Den5
        .MS10LDL1 = rs("MS10LDL1")          ' ����l10 L/DL1
        .MS10LDL2 = rs("MS10LDL2")          ' ����l10 L/DL2
        .MS10LDL3 = rs("MS10LDL3")          ' ����l10 L/DL3
        .MS10LDL4 = rs("MS10LDL4")          ' ����l10 L/DL4
        .MS10LDL5 = rs("MS10LDL5")          ' ����l10 L/DL5
        .MS10DEN1 = rs("MS10DEN1")          ' ����l10 Den1
        .MS10DEN2 = rs("MS10DEN2")          ' ����l10 Den2
        .MS10DEN3 = rs("MS10DEN3")          ' ����l10 Den3
        .MS10DEN4 = rs("MS10DEN4")          ' ����l10 Den4
        .MS10DEN5 = rs("MS10DEN5")          ' ����l10 Den5
        .MS11LDL1 = rs("MS11LDL1")          ' ����l11 L/DL1
        .MS11LDL2 = rs("MS11LDL2")          ' ����l11 L/DL2
        .MS11LDL3 = rs("MS11LDL3")          ' ����l11 L/DL3
        .MS11LDL4 = rs("MS11LDL4")          ' ����l11 L/DL4
        .MS11LDL5 = rs("MS11LDL5")          ' ����l11 L/DL5
        .MS11DEN1 = rs("MS11DEN1")          ' ����l11 Den1
        .MS11DEN2 = rs("MS11DEN2")          ' ����l11 Den2
        .MS11DEN3 = rs("MS11DEN3")          ' ����l11 Den3
        .MS11DEN4 = rs("MS11DEN4")          ' ����l11 Den4
        .MS11DEN5 = rs("MS11DEN5")          ' ����l11 Den5
        .MS12LDL1 = rs("MS12LDL1")          ' ����l12 L/DL1
        .MS12LDL2 = rs("MS12LDL2")          ' ����l12 L/DL2
        .MS12LDL3 = rs("MS12LDL3")          ' ����l12 L/DL3
        .MS12LDL4 = rs("MS12LDL4")          ' ����l12 L/DL4
        .MS12LDL5 = rs("MS12LDL5")          ' ����l12 L/DL5
        .MS12DEN1 = rs("MS12DEN1")          ' ����l12 Den1
        .MS12DEN2 = rs("MS12DEN2")          ' ����l12 Den2
        .MS12DEN3 = rs("MS12DEN3")          ' ����l12 Den3
        .MS12DEN4 = rs("MS12DEN4")          ' ����l12 Den4
        .MS12DEN5 = rs("MS12DEN5")          ' ����l12 Den5
        .MS13LDL1 = rs("MS13LDL1")          ' ����l13 L/DL1
        .MS13LDL2 = rs("MS13LDL2")          ' ����l13 L/DL2
        .MS13LDL3 = rs("MS13LDL3")          ' ����l13 L/DL3
        .MS13LDL4 = rs("MS13LDL4")          ' ����l13 L/DL4
        .MS13LDL5 = rs("MS13LDL5")          ' ����l13 L/DL5
        .MS13DEN1 = rs("MS13DEN1")          ' ����l13 Den1
        .MS13DEN2 = rs("MS13DEN2")          ' ����l13 Den2
        .MS13DEN3 = rs("MS13DEN3")          ' ����l13 Den3
        .MS13DEN4 = rs("MS13DEN4")          ' ����l13 Den4
        .MS13DEN5 = rs("MS13DEN5")          ' ����l13 Den5
        .MS14LDL1 = rs("MS14LDL1")          ' ����l14 L/DL1
        .MS14LDL2 = rs("MS14LDL2")          ' ����l14 L/DL2
        .MS14LDL3 = rs("MS14LDL3")          ' ����l14 L/DL3
        .MS14LDL4 = rs("MS14LDL4")          ' ����l14 L/DL4
        .MS14LDL5 = rs("MS14LDL5")          ' ����l14 L/DL5
        .MS14DEN1 = rs("MS14DEN1")          ' ����l14 Den1
        .MS14DEN2 = rs("MS14DEN2")          ' ����l14 Den2
        .MS14DEN3 = rs("MS14DEN3")          ' ����l14 Den3
        .MS14DEN4 = rs("MS14DEN4")          ' ����l14 Den4
        .MS14DEN5 = rs("MS14DEN5")          ' ����l14 Den5
        .MS15LDL1 = rs("MS15LDL1")          ' ����l15 L/DL1
        .MS15LDL2 = rs("MS15LDL2")          ' ����l15 L/DL2
        .MS15LDL3 = rs("MS15LDL3")          ' ����l15 L/DL3
        .MS15LDL4 = rs("MS15LDL4")          ' ����l15 L/DL4
        .MS15LDL5 = rs("MS15LDL5")          ' ����l15 L/DL5
        .MS15DEN1 = rs("MS15DEN1")          ' ����l15 Den1
        .MS15DEN2 = rs("MS15DEN2")          ' ����l15 Den2
        .MS15DEN3 = rs("MS15DEN3")          ' ����l15 Den3
        .MS15DEN4 = rs("MS15DEN4")          ' ����l15 Den4
        .MS15DEN5 = rs("MS15DEN5")          ' ����l15 Den5
        If Not IsNull(rs("MS01DVD2")) Then .MS01DVD2 = rs("MS01DVD2")      ' ����l01 DVD2
        If Not IsNull(rs("MS02DVD2")) Then .MS02DVD2 = rs("MS02DVD2")      ' ����l02 DVD2
        If Not IsNull(rs("MS03DVD2")) Then .MS03DVD2 = rs("MS03DVD2")      ' ����l03 DVD2
        If Not IsNull(rs("MS04DVD2")) Then .MS04DVD2 = rs("MS04DVD2")      ' ����l04 DVD2
        If Not IsNull(rs("MS05DVD2")) Then .MS05DVD2 = rs("MS05DVD2")      ' ����l05 DVD2
        
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        .MSZEROMN = fncNullCheck(rs("MSZEROMN"))    ' L/DL0�A�����ŏ��l
        .MSZEROMX = fncNullCheck(rs("MSZEROMX"))    ' L/DL0�A�����ő�l
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
        
        .REGDATE = rs("REGDATE")            ' �o�^���t
    End With
    
    funGetCryGDJisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ����LT���ю擾�֐�
'------------------------------------------------

'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ007���������A����LT���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :tCryLT        ,O  ,type_DBDRV_scmzc_fcmkc001c_LT    :����LT����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetCryLTJisseki(sCryNum As String, iSmplID As Long, tCryLT As type_DBDRV_scmzc_fcmkc001c_LT) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryLTJisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ007�̌���LT���ђl����������B

    '2005/12/02 mod SET���� LT����l10�_�ANULL���Ή� ->
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MEASPEAK, CALCMEAS, REGDATE, "
    sql = sql & "NVL(MEAS1, -1) MEAS1,"
    sql = sql & "NVL(MEAS2, -1) MEAS2,"
    sql = sql & "NVL(MEAS3, -1) MEAS3,"
    sql = sql & "NVL(MEAS4, -1) MEAS4,"
    sql = sql & "NVL(MEAS5, -1) MEAS5,"
    sql = sql & "NVL(MEAS6, -1) MEAS6,"
    sql = sql & "NVL(MEAS7, -1) MEAS7,"
    sql = sql & "NVL(MEAS8, -1) MEAS8,"
    sql = sql & "NVL(MEAS9, -1) MEAS9,"
    sql = sql & "NVL(MEAS10, -1) MEAS10,"
    sql = sql & "LTSPIFLG "
    sql = sql & ",NVL(CONVAL, -1) CONVAL "
    sql = sql & "from TBCMJ007 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    '2005/12/02 mod SET���� LT����l10�_�ANULL���Ή� <-
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryLTJisseki = -1
        GoTo proc_exit
    End If
    
     ''���o���ʂ��i�[����
    With tCryLT
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCOND = rs("TRANCOND")          ' ��������
        .TRANCNT = rs("TRANCNT")            ' ������
        .SMPLNO = rs("SMPLNO")              ' �T���v���m��
        .SMPLUMU = rs("SMPLUMU")            ' �T���v���L��
        .MEAS1 = rs("MEAS1")                ' ����l�P
        .MEAS2 = rs("MEAS2")                ' ����l�Q
        .MEAS3 = rs("MEAS3")                ' ����l�R
        .MEAS4 = rs("MEAS4")                ' ����l�S
        .MEAS5 = rs("MEAS5")                ' ����l�T
        .MEASPEAK = rs("MEASPEAK")          ' ����l �s�[�N�l
        .CALCMEAS = rs("CALCMEAS")          ' �v�Z����
        .REGDATE = rs("REGDATE")            ' �o�^���t
''Add Start 2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
        .CONVAL = rs.Fields("CONVAL")       '10�����Z�l
''Add End   2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
        '2005/12/02 add SET���� ����l�U�`�P�O�J�����ǉ��̂��ߒǉ� ->
        '                       ����t���O�J�����ǉ�
        .MEAS6 = rs("MEAS6")            ' ����l�U
        .MEAS7 = rs("MEAS7")            ' ����l�V
        .MEAS8 = rs("MEAS8")            ' ����l�W
        .MEAS9 = rs("MEAS9")            ' ����l�X
        .MEAS10 = rs("MEAS10")          ' ����l�P�O
        .LTSPIFLG = Trim(CStr(NulltoStr(rs.Fields("LTSPIFLG").Value)))  '����ʒu����t���O
        
        '2005/12/02 add SET���� ����l�U�`�P�O�J�����ǉ��̂��ߒǉ� <-
        '                       ����t���O�J�����ǉ�
    End With
    
    funGetCryLTJisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ����EPD���ю擾�֐�
'------------------------------------------------

'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ001���������A����EPD���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :tCryEPD       ,O  ,type_DBDRV_scmzc_fcmkc001c_EPD   :����EPD����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetCryEPDJisseki(sCryNum As String, iSmplID As Long, tCryEPD As type_DBDRV_scmzc_fcmkc001c_EPD) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryEPDJisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ001�̌���EPD���ђl����������B
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MEASURE, REGDATE "
    sql = sql & "from TBCMJ001 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryEPDJisseki = -1
        GoTo proc_exit
    End If
    
     ''���o���ʂ��i�[����
    With tCryEPD
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCOND = rs("TRANCOND")          ' ��������
        .TRANCNT = rs("TRANCNT")            ' ������
        .SMPLNO = rs("SMPLNO")              ' �T���v���m��
        .SMPLUMU = rs("SMPLUMU")            ' �T���v���L��
        .MEASURE = rs("MEASURE")            ' ����l
        .REGDATE = rs("REGDATE")            ' �o�^���t
    End With
    
    funGetCryEPDJisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ������R���т̃f�[�^�����擾
'------------------------------------------------
'�T�v      :������R����(�\����)�ɑ��݂���f�[�^�������擾����B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :tCryRs        ,I  ,type_DBDRV_scmzc_fcmkc001c_CryR      :������R���э\����
'          :�߂�l        ,O  ,Integer                              :�f�[�^����
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Private Function funGetRsCnt(tCryRs As type_DBDRV_scmzc_fcmkc001c_CryR) As Integer
    
    Dim sql         As String
    Dim rs          As OraDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function funGetRsCnt"

    funGetRsCnt = 0
    
    If tCryRs.MEAS1 = -1 Then GoTo proc_exit
    funGetRsCnt = funGetRsCnt + 1
    If tCryRs.MEAS2 = -1 Then GoTo proc_exit
    funGetRsCnt = funGetRsCnt + 1
    If tCryRs.MEAS3 = -1 Then GoTo proc_exit
    funGetRsCnt = funGetRsCnt + 1
    If tCryRs.MEAS4 = -1 Then GoTo proc_exit
    funGetRsCnt = funGetRsCnt + 1
    If tCryRs.MEAS5 = -1 Then GoTo proc_exit
    funGetRsCnt = funGetRsCnt + 1

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    funGetRsCnt = -1
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ���������l�`�F�b�N
'------------------------------------------------

'�T�v      :���������`�F�b�N���s�Ȃ����ʂ�Ԃ��B
'����      :
'����      :TEST2004/10

Public Function funChkJissoku(tFullHin As tFullHinban, tCryRs As type_DBDRV_scmzc_fcmkc001c_CryR) As Boolean
    Dim tSiyou          As type_DBDRV_scmzc_fcmkc001c_Siyou
    Dim retJudg         As Boolean
    Dim sCryRs As type_DBDRV_scmzc_fcmkc001c_CryR
    
    If funGet_TBCME018(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then
        funChkJissoku = False
        Exit Function
    End If
    sCryRs = tCryRs
    '����������s�Ȃ�
    '����_���A�ۏ؂͑S�_�ɌŒ�(�S�Ă̒l���`�F�b�N���邽�߁j
    tSiyou.HSXRSPOT = "5"
    tSiyou.HSXRHWYT = "3"
    If Not CrResJudg(0, tSiyou, sCryRs, retJudg, 1) Then
        funChkJissoku = False
        Exit Function
    End If
    If retJudg = False Then
        funChkJissoku = False
        Exit Function
    End If
    funChkJissoku = True
    
End Function

'Add Start 2011/01/31 SMPK Miyata
'------------------------------------------------
' ����C���ю擾�֐�
'------------------------------------------------
'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ023���������A����C���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c
'          :tCryC         ,O  ,type_DBDRV_scmzc_fcmkc001c_C     :����C����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :
Public Function funGetCryCJisseki(sCryNum As String, iSmplID As Long, tCryC As type_DBDRV_scmzc_fcmkc001c_C) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryCJisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ023�̌���C���ђl����������B
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO, "
    sql = sql & "SMPLUMUC, CPTNJSK, CDISKJSK, CRINGNKJSK, CRINGGKJSK, CHANTEI, "
    sql = sql & "REGDATE "
    sql = sql & "from TBCMJ023 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryCJisseki = -1
        GoTo proc_exit
    End If
    
     ''���o���ʂ��i�[����
    With tCryC
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCNT = rs("TRANCNT")            ' ������

        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")             ' �T���v���m��
        If IsNull(rs("SMPLUMUC")) = False Then .SMPLUMUC = rs("SMPLUMUC")       ' �T���v���L���iC�j
        If IsNull(rs("CPTNJSK")) = False Then .CPTNJSK = rs("CPTNJSK")          ' C �p�^�[������
        If IsNull(rs("CDISKJSK")) = False Then .CDISKJSK = rs("CDISKJSK")       ' C Disk���a����
        If IsNull(rs("CRINGNKJSK")) = False Then .CRINGNKJSK = rs("CRINGNKJSK") ' C Ring���a����
        If IsNull(rs("CRINGGKJSK")) = False Then .CRINGGKJSK = rs("CRINGGKJSK") ' C Ring�O�a����
        If IsNull(rs("CHANTEI")) = False Then .CHANTEI = rs("CHANTEI")          ' C ���茋��
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")          ' �o�^���t
    End With

    funGetCryCJisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ����CJ���ю擾�֐�
'------------------------------------------------
'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ023���������A����CJ���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c
'          :tCryCJ        ,O  ,type_DBDRV_scmzc_fcmkc001c_CJ    :����CJ����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :
Public Function funGetCryCJJisseki(sCryNum As String, iSmplID As Long, tCryCJ As type_DBDRV_scmzc_fcmkc001c_CJ) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryCJJisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ023�̌���C���ђl����������B
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO, "
    sql = sql & "SMPLUMUCJ, CJPTNJSK, CJDISKJSK, CJRINGNKJSK, CJRINGGKJSK, CJBANDNKJSK, CJBANDGKJSK, CJRINGCALC, "
    sql = sql & "CJPICALC , CJHANTEI, CJDMAXPIC5, CJRMAXPIC5, CJDRMAXPIC5, CJALLMAXDIC5, CJALLMINRINC5, CJALLMAXRIGC5, "
    sql = sql & "REGDATE "
    sql = sql & "from TBCMJ023 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryCJJisseki = -1
        GoTo proc_exit
    End If
    
     ''���o���ʂ��i�[����
    With tCryCJ
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCNT = rs("TRANCNT")            ' ������

        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                 ' �T���v���m��
        If IsNull(rs("SMPLUMUCJ")) = False Then .SMPLUMUCJ = rs("SMPLUMUCJ")        ' �T���v���L���iCJ�j
        If IsNull(rs("CJPTNJSK")) = False Then .CJPTNJSK = rs("CJPTNJSK")           ' CJ �p�^�[������
        If IsNull(rs("CJDISKJSK")) = False Then .CJDISKJSK = rs("CJDISKJSK")        ' CJ Disk���a����
        If IsNull(rs("CJRINGNKJSK")) = False Then .CJRINGNKJSK = rs("CJRINGNKJSK")  ' CJ Ring���a����
        If IsNull(rs("CJRINGGKJSK")) = False Then .CJRINGGKJSK = rs("CJRINGGKJSK")  ' CJ Ring�O�a����
        If IsNull(rs("CJBANDNKJSK")) = False Then .CJBANDNKJSK = rs("CJBANDNKJSK")  ' CJ Band���a����
        If IsNull(rs("CJBANDGKJSK")) = False Then .CJBANDGKJSK = rs("CJBANDGKJSK")  ' CJ Band�O�a����
        If IsNull(rs("CJRINGCALC")) = False Then .CJRINGCALC = rs("CJRINGCALC")     ' CJ Ring���v�Z
        If IsNull(rs("CJPICALC")) = False Then .CJPICALC = rs("CJPICALC")           ' CJ Pi���v�Z
        If IsNull(rs("CJHANTEI")) = False Then .CJHANTEI = rs("CJHANTEI")           ' CJ ���茋��
        If IsNull(rs("CJDMAXPIC5")) = False Then .CJDMAXPIC5 = rs("CJDMAXPIC5")     ' CJ Disk�̂݃p�^�[�� Pi������l
        If IsNull(rs("CJRMAXPIC5")) = False Then .CJRMAXPIC5 = rs("CJRMAXPIC5")     ' CJ Ring�̂݃p�^�[�� Pi������l
        If IsNull(rs("CJDRMAXPIC5")) = False Then .CJDRMAXPIC5 = rs("CJDRMAXPIC5")          ' CJ DiskRing�p�^�[�� Pi������l
        If IsNull(rs("CJALLMAXDIC5")) = False Then .CJALLMAXDIC5 = rs("CJALLMAXDIC5")       ' CJ ����Disk���a����l
        If IsNull(rs("CJALLMINRINC5")) = False Then .CJALLMINRINC5 = rs("CJALLMINRINC5")    ' CJ ����Ring���a�����l
        If IsNull(rs("CJALLMAXRIGC5")) = False Then .CJALLMAXRIGC5 = rs("CJALLMAXRIGC5")    ' CJ ����Ring�O�a����l
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")                      ' �o�^���t
    End With

    funGetCryCJJisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ����CJLT���ю擾�֐�
'------------------------------------------------
'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ023���������A����CJLT���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c
'          :tCryCJLT      ,O  ,type_DBDRV_scmzc_fcmkc001c_CJLT  :����CJLT����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :
Public Function funGetCryCJLTJisseki(sCryNum As String, iSmplID As Long, tCryCJLT As type_DBDRV_scmzc_fcmkc001c_CJLT) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryCJLTJisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ023�̌���C���ђl����������B
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO, "
    sql = sql & "SMPLUMUCJLT, CJLTPTNJSK, CJLTDISKJSK, CJLTRINGNKJSK, CJLTRINGGKJSK, "
    sql = sql & "CJLTBANDNKJSK , CJLTBANDGKJSK, CJLTRINGCALC, CJLTPICALC, CJLTHANTEI, "
    sql = sql & "REGDATE "
    sql = sql & "from TBCMJ023 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"


    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryCJLTJisseki = -1
        GoTo proc_exit
    End If
    
     ''���o���ʂ��i�[����
    With tCryCJLT
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCNT = rs("TRANCNT")            ' ������

        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                 ' �T���v���m��
        If IsNull(rs("SMPLUMUCJLT")) = False Then .SMPLUMUCJLT = rs("SMPLUMUCJLT")    ' �T���v���L���iCJ(LT)�j
        
        If IsNull(rs("CJLTPTNJSK")) = False Then .CJLTPTNJSK = rs("CJLTPTNJSK")     ' CJ(LT) �p�^�[������
        If IsNull(rs("CJLTDISKJSK")) = False Then .CJLTDISKJSK = rs("CJLTDISKJSK")  ' CJ(LT) Disk���a����
        If IsNull(rs("CJLTRINGNKJSK")) = False Then .CJLTRINGNKJSK = rs("CJLTRINGNKJSK")  ' CJ(LT) Ring���a����
        If IsNull(rs("CJLTRINGGKJSK")) = False Then .CJLTRINGGKJSK = rs("CJLTRINGGKJSK")  ' CJ(LT) Ring�O�a����
        If IsNull(rs("CJLTBANDNKJSK")) = False Then .CJLTBANDNKJSK = rs("CJLTBANDNKJSK")  ' CJ(LT) Band���a����
        If IsNull(rs("CJLTBANDGKJSK")) = False Then .CJLTBANDGKJSK = rs("CJLTBANDGKJSK")  ' CJ(LT) Band�O�a����
        If IsNull(rs("CJLTRINGCALC")) = False Then .CJLTRINGCALC = rs("CJLTRINGCALC")     ' CJ(LT) Ring���v�Z
        If IsNull(rs("CJLTPICALC")) = False Then .CJLTPICALC = rs("CJLTPICALC")           ' CJ(LT) Pi���v�Z
        If IsNull(rs("CJLTHANTEI")) = False Then .CJLTHANTEI = rs("CJLTHANTEI")           ' CJ(LT) ���茋��
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")                    ' �o�^���t
    End With

    funGetCryCJLTJisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ����CJ2���ю擾�֐�
'------------------------------------------------
'�T�v      :�����ԍ��A�T���v���h�c����ATBCMJ023���������A����CJLT���ђl���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^                               :����
'          :sCryNum       ,I  ,String                           :�����ԍ�
'          :iSmplID       ,I  ,Long                             :�T���v���h�c
'          :tCryCJ2       ,O  ,type_DBDRV_scmzc_fcmkc001c_CJ2   :����CJ2����(�\����)
'          :�߂�l        ,O  ,Integer                          :�擾���� = 0 : ����
'                                                                          -1 : �ُ�
'����      :
'����      :
Public Function funGetCryCJ2Jisseki(sCryNum As String, iSmplID As Long, tCryCJ2 As type_DBDRV_scmzc_fcmkc001c_CJ2) As Integer

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryCJ2Jisseki"
    
    '�����ԍ��A�T���v���h�c����TBCMJ023�̌���C���ђl����������B
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO, "
    sql = sql & "SMPLUMUCJ2, CJ2PTNJSK, CJ2DISKJSK, CJ2RINGNKJSK, CJ2RINGGKJSK,CJ2PICALC, "
    sql = sql & "CJ2HANTEI , CJ2DMAXPIC5, CJ2RMAXPIC5, CJ2RMINRINC5, CJ2RMAXRIGC5, CJ2DRMAXPIC5, "
    sql = sql & "CJ2DRMINRINC5, CJ2DRMAXRIGC5, "
    sql = sql & "REGDATE "
    sql = sql & "from TBCMJ023 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"

''
''
''

    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetCryCJ2Jisseki = -1
        GoTo proc_exit
    End If
    
     ''���o���ʂ��i�[����
    With tCryCJ2
        .CRYNUM = rs("CRYNUM")              ' �����ԍ�
        .POSITION = rs("POSITION")          ' �ʒu
        .SMPKBN = rs("SMPKBN")              ' �T���v���敪
        .TRANCNT = rs("TRANCNT")            ' ������

        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                 ' �T���v���m��
        If IsNull(rs("SMPLUMUCJ2")) = False Then .SMPLUMUCJ2 = rs("SMPLUMUCJ2")     ' �T���v���L���iCJ2�j
        
        If IsNull(rs("CJ2PTNJSK")) = False Then .CJ2PTNJSK = rs("CJ2PTNJSK")        ' CJ2 �p�^�[������
        If IsNull(rs("CJ2DISKJSK")) = False Then .CJ2DISKJSK = rs("CJ2DISKJSK")     ' CJ2 Disk���a����
        If IsNull(rs("CJ2RINGNKJSK")) = False Then .CJ2RINGNKJSK = rs("CJ2RINGNKJSK")   ' CJ2 Ring���a����
        If IsNull(rs("CJ2RINGGKJSK")) = False Then .CJ2RINGGKJSK = rs("CJ2RINGGKJSK")   ' CJ2 Ring�O�a����
        If IsNull(rs("CJ2PICALC")) = False Then .CJ2PICALC = rs("CJ2PICALC")            ' CJ2 Pi���v�Z
        If IsNull(rs("CJ2HANTEI")) = False Then .CJ2HANTEI = rs("CJ2HANTEI")            ' CJ2 ���茋��
        If IsNull(rs("CJ2DMAXPIC5")) = False Then .CJ2DMAXPIC5 = rs("CJ2DMAXPIC5")      ' CJ2 Disk�̂݃p�^�[�� Pi�������l(MAX���������ł�)
        If IsNull(rs("CJ2RMAXPIC5")) = False Then .CJ2RMAXPIC5 = rs("CJ2RMAXPIC5")      ' CJ2 Ring�̂݃p�^�[�� Pi�������l(MAX���������ł�)
        If IsNull(rs("CJ2RMINRINC5")) = False Then .CJ2RMINRINC5 = rs("CJ2RMINRINC5")   ' CJ2 Ring�̂݃p�^�[�� Ring���a�����l
        If IsNull(rs("CJ2RMAXRIGC5")) = False Then .CJ2RMAXRIGC5 = rs("CJ2RMAXRIGC5")   ' CJ2 Ring�̂݃p�^�[�� Ring�O�a����l
        
        If IsNull(rs("CJ2DRMAXPIC5")) = False Then .CJ2DRMAXPIC5 = rs("CJ2DRMAXPIC5")   ' CJ2 DiskRing�p�^�[�� Pi�������l(MAX���������ł�)
        If IsNull(rs("CJ2DRMINRINC5")) = False Then .CJ2DRMINRINC5 = rs("CJ2DRMINRINC5") ' CJ2 DiskRing�p�^�[�� Ring���a�����l
        If IsNull(rs("CJ2DRMAXRIGC5")) = False Then .CJ2DRMAXRIGC5 = rs("CJ2DRMAXRIGC5") ' CJ2 DiskRing�p�^�[�� Ring�O�a����l
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")                  ' �o�^���t
    End With

    funGetCryCJ2Jisseki = 0

proc_exit:
    '�I��
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
'Add End   2011/01/31 SMPK Miyata
