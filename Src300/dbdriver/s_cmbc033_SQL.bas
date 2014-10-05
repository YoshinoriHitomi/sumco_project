Attribute VB_Name = "s_cmbc033_SQL"
Option Explicit

' �����w��

Public lStfMst As Long
Public intEnCmd As Integer
Public Const MAXCNT As Integer = 16                             ' �ő匏��
Public Const BlkTop As Integer = 1                                 ' TOP��
Public Const BlkTail As Integer = 2                                ' TAIL��
Public Const KSYSCLASS As String = "GP"                         ' �V�X�e���敪
Public Const MSYSCLASS As String = "NM"                         ' �V�X�e���敪
Public Const KCLASS As String = "01"                            ' �N���X
Public Const KCODE As String = "1"                              ' �R�[�h

' �u���b�N���
Public Type typ_BlkInf1
    BLOCKID As String * 12      ' �u���b�NID
    LENGTH As Integer           ' ����
    REALLEN As Integer          ' ������
    KRPROCCD As String * 5      ' ���݊Ǘ��H��
    NOWPROC As String * 5       ' ���ݍH��
    LPKRPROCCD As String * 5    ' �ŏI�ʉߊǗ��H��
    LASTPASS As String * 5      ' �ŏI�ʉߍH��
    RSTATCLS As String * 1      ' ������ԋ敪
    BDCODE As String * 3        ' �s�Ǘ��R�R�[�h
    PALTNUM As String * 4       ' �p���b�g�ԍ�
    SEED As String * 4          ' �V�[�h
    COF As type_Coefficient     ' �ΐ͌W���v�Z
    SAMPFLAG As Boolean         ' �T���v���擾�t���O
End Type

Type cmkc001b_LockWait
    flag As Boolean
    Grp As Integer
End Type
Type cmkc001b_Wait3_HINBAN
    HINBAN As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    factory As String * 1           ' �H��
    opecond As String * 1           ' ���Ə���
End Type
Type cmkc001b_Wait3_BLK
    BLOCKID As String * 12          ' �u���b�NID
    IngotPos As Integer             ' �������J�n�ʒu
    LENGTH As Integer               ' ����
    NOWPROC As String * 5           ' ���ݍH��
    HOLDCLS As String * 1           ' �z�[���h�敪 ---kuramoto �ǉ� 2001/09/19----
    GRPFLG1 As Integer           ' �O���[�v���
    GRPFLG2 As Integer           ' �O���[�v���
    COLORFLG As Boolean
    topHin As cmkc001b_Wait3_HINBAN
    botHin As cmkc001b_Wait3_HINBAN
End Type
Type cmkc001b_Wait3
    CRYNUM As String * 12           ' �����ԍ�
    blkInfo() As cmkc001b_Wait3_BLK
End Type

Type type_cmkc001b_SmpMng
    CRYNUM As String * 12
    IngotPos As Integer
    SMPKBN As String * 1
    
    HINBAN As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    factory As String * 1           ' �H��
    opecond As String * 1           ' ���Ə���
    
    
    CRYINDRS As String * 1
    CRYRESRS As String * 1
    CRYINDOI As String * 1
    CRYRESOI As String * 1
    CRYINDB1 As String * 1
    CRYRESB1 As String * 1
    CRYINDB2 As String * 1
    CRYRESB2 As String * 1
    CRYINDB3 As String * 1
    CRYRESB3 As String * 1
    CRYINDL1 As String * 1
    CRYRESL1 As String * 1
    CRYINDL2 As String * 1
    CRYRESL2 As String * 1
    CRYINDL3 As String * 1
    CRYRESL3 As String * 1
    CRYINDL4 As String * 1
    CRYRESL4 As String * 1
    CRYINDCS As String * 1
    CRYRESCS As String * 1
    CRYINDGD As String * 1
    CRYRESGD As String * 1
    CRYINDT As String * 1
    CRYREST As String * 1
    CRYINDEP As String * 1
    CRYRESEP As String * 1
    
    HSXCNHWS As String * 1          ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    HSXLTHWS As String * 1          ' �i�r�w�k�^�C���ۏؕ��@�Q��
    EPD As String * 1               ' EPD
End Type

#If SPEEDUP Then   '���������� 02.1.28-2.15 �쑺
Private Type tSmpMng
    BLOCKID As String * 12
    TOPPOS As Integer
    BOTPOS As Integer
    
    CRYNUM As String * 12
    IngotPos As Integer
    SMPKBN As String * 1
    
    HINBAN As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    factory As String * 1           ' �H��
    opecond As String * 1           ' ���Ə���
    
    CRYINDRS As String * 1
    CRYRESRS As String * 1
    CRYINDOI As String * 1
    CRYRESOI As String * 1
    CRYINDB1 As String * 1
    CRYRESB1 As String * 1
    CRYINDB2 As String * 1
    CRYRESB2 As String * 1
    CRYINDB3 As String * 1
    CRYRESB3 As String * 1
    CRYINDL1 As String * 1
    CRYRESL1 As String * 1
    CRYINDL2 As String * 1
    CRYRESL2 As String * 1
    CRYINDL3 As String * 1
    CRYRESL3 As String * 1
    CRYINDL4 As String * 1
    CRYRESL4 As String * 1
    CRYINDCS As String * 1
    CRYRESCS As String * 1
    CRYINDGD As String * 1
    CRYRESGD As String * 1
    CRYINDT As String * 1
    CRYREST As String * 1
    CRYINDEP As String * 1
    CRYRESEP As String * 1
End Type
#End If


'�҂��ꗗ

'�����\���p
Public Type type_DBDRV_scmzc_fcmkc001b_Disp
    CRYNUM As String * 12           ' �����ԍ�
    IngotPos As Integer             ' �������J�n�ʒu
'    LENGTH As Integer               ' ����              '2001/11/8
    BLOCKID As String * 12          ' �u���b�NID
    HSXTYPE As String * 1           ' �i�r�w�^�C�v
    HSXCDIR As String * 1           ' �i�r�w�����ʕ���
    UPDDATE As Date                 ' �X�V���t
    Judg As String                  ' ����
    hin() As tFullHinban            ' �i��(full)
    HOLDCLS As String * 1           ' �z�[���h�敪 ---kuramoto �ǉ� 2001/09/25----
    SMP() As type_cmkc001b_SmpMng   ' �T���v���Ǘ�
End Type

'�i�ԁA�d�l�A���������擾�p (TOP,TAIL���łQ���R�[�h�擾)
Public Type type_DBDRV_scmzc_fcmkc001c_Siyou
    '�u���b�N�Ǘ�
    CRYNUM As String * 12             ' �����ԍ�
    IngotPos As Integer               ' �������J�n�ʒu
    LENGTH As Integer                 ' ����
    '�i�ԊǗ�
    hin As tFullHinban                ' �i��(full)
        
        '�������
    PRODCOND As String * 4            ' �������
    PGID As String * 8                ' �o�f�|�h�c
    UPLENGTH As Integer               ' ���グ����
    FREELENG As Integer               ' �t���[��
    DIAMETER As Integer               ' ���a 2002/05/01 S.Sano
    CHARGE As Double                  ' �`���[�W��
    SEED As String * 4                ' �V�[�h
    ADDDPPOS As Integer                 ' �ǉ��h�[�v�ʒu

    '���i�d�l
    HSXTYPE As String * 1             ' �i�r�w�^�C�v
    HSXD1CEN As Double                ' �i�r�w���a�P���S
    HSXCDIR As String * 1             ' �i�r�w�����ʕ���
    HSXRMIN As Double                 ' �i�r�w���R����
    HSXRMAX As Double                 ' �i�r�w���R���
    HSXRAMIN As Double                ' �i�r�w���R���ω���
    HSXRAMAX As Double                ' �i�r�w���R���Ϗ��
    HSXRMBNP As Double                ' �i�r�w���R�ʓ����z
    HSXRSPOH As String * 1            ' �i�r�w���R����ʒu�Q��
    HSXRSPOT As String * 1            ' �i�r�w���R����ʒu�Q�_
    HSXRSPOI As String * 1            ' �i�r�w���R����ʒu�Q��
    HSXRHWYT As String * 1            ' �i�r�w���R�ۏؕ��@�Q��
    HSXRHWYS As String * 1            ' �i�r�w���R�ۏؕ��@�Q��

    HSXONMIN As Double                ' �i�r�w�_�f�Z�x����
    HSXONMAX As Double                ' �i�r�w�_�f�Z�x���
    HSXONAMN As Double                ' �i�r�w�_�f�Z�x���ω���
    HSXONAMX As Double                ' �i�r�w�_�f�Z�x���Ϗ��
    HSXONMBP As Double                ' �i�r�w�_�f�Z�x�ʓ����z
    HSXONSPH As String * 1            ' �i�r�w�_�f�Z�x����ʒu�Q��
    HSXONSPT As String * 1            ' �i�r�w�_�f�Z�x����ʒu�Q�_
    HSXONSPI As String * 1            ' �i�r�w�_�f�Z�x����ʒu�Q��
    HSXONHWT As String * 1            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    HSXONHWS As String * 1            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��

    HSXBM1AN As Double                ' �i�r�w�a�l�c�P���ω���
    HSXBM1AX As Double                ' �i�r�w�a�l�c�P���Ϗ��
    HSXBM2AN As Double                ' �i�r�w�a�l�c�Q���ω���
    HSXBM2AX As Double                ' �i�r�w�a�l�c�Q���Ϗ��
    HSXBM3AN As Double                ' �i�r�w�a�l�c�R���ω���
    HSXBM3AX As Double                ' �i�r�w�a�l�c�R���Ϗ��
    HSXBM1SH As String * 1            ' �i�r�w�a�l�c�P����ʒu�Q��
    HSXBM1ST As String * 1            ' �i�r�w�a�l�c�P����ʒu�Q�_
    HSXBM1SR As String * 1            ' �i�r�w�a�l�c�P����ʒu�Q��
    HSXBM1HT As String * 1            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    HSXBM1HS As String * 1            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    HSXBM2SH As String * 1            ' �i�r�w�a�l�c�Q����ʒu�Q��
    HSXBM2ST As String * 1            ' �i�r�w�a�l�c�Q����ʒu�Q�_
    HSXBM2SR As String * 1            ' �i�r�w�a�l�c�Q����ʒu�Q��
    HSXBM2HT As String * 1            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    HSXBM2HS As String * 1            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    HSXBM3SH As String * 1            ' �i�r�w�a�l�c�R����ʒu�Q��
    HSXBM3ST As String * 1            ' �i�r�w�a�l�c�R����ʒu�Q�_
    HSXBM3SR As String * 1            ' �i�r�w�a�l�c�R����ʒu�Q��
    HSXBM3HT As String * 1            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
    HSXBM3HS As String * 1            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��

    HSXOS1AX As Double                ' �i�r�w�n�r�e�P���Ϗ��
    HSXOS1MX As Double                ' �i�r�w�n�r�e�P���
    HSXOS2AX As Double                ' �i�r�w�n�r�e�Q���Ϗ��
    HSXOS2MX As Double                ' �i�r�w�n�r�e�Q���
    HSXOS3AX As Double                ' �i�r�w�n�r�e�R���Ϗ��
    HSXOS3MX As Double                ' �i�r�w�n�r�e�R���
    HSXOS4AX As Double                ' �i�r�w�n�r�e�S���Ϗ��
    HSXOS4MX As Double                ' �i�r�w�n�r�e�S���
    HSXOS1SH As String * 1            ' �i�r�w�n�r�e�P����ʒu�Q��
    HSXOS1ST As String * 1            ' �i�r�w�n�r�e�P����ʒu�Q�_
    HSXOS1SR As String * 1            ' �i�r�w�n�r�e�P����ʒu�Q��
    HSXOS1HT As String * 1            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    HSXOS1HS As String * 1            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    HSXOS2SH As String * 1            ' �i�r�w�n�r�e�Q����ʒu�Q��
    HSXOS2ST As String * 1            ' �i�r�w�n�r�e�Q����ʒu�Q�_
    HSXOS2SR As String * 1            ' �i�r�w�n�r�e�Q����ʒu�Q��
    HSXOS2HT As String * 1            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    HSXOS2HS As String * 1            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    HSXOS3SH As String * 1            ' �i�r�w�n�r�e�R����ʒu�Q��
    HSXOS3ST As String * 1            ' �i�r�w�n�r�e�R����ʒu�Q�_
    HSXOS3SR As String * 1            ' �i�r�w�n�r�e�R����ʒu�Q��
    HSXOS3HT As String * 1            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    HSXOS3HS As String * 1            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    HSXOS4SH As String * 1            ' �i�r�w�n�r�e�S����ʒu�Q��
    HSXOS4ST As String * 1            ' �i�r�w�n�r�e�S����ʒu�Q�_
    HSXOS4SR As String * 1            ' �i�r�w�n�r�e�S����ʒu�Q��
    HSXOS4HT As String * 1            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    HSXOS4HS As String * 1            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    HSXOS1NS As String * 2            ' �i�r�w�n�r�e�P�M�����@
    HSXOS2NS As String * 2            ' �i�r�w�n�r�e�Q�M�����@
    HSXOS3NS As String * 2            ' �i�r�w�n�r�e�R�M�����@
    HSXOS4NS As String * 2            ' �i�r�w�n�r�e�S�M�����@
    HSXBM1NS As String * 2            ' �i�r�w�a�l�c�P�M�����@
    HSXBM2NS As String * 2            ' �i�r�w�a�l�c�Q�M�����@
    HSXBM3NS As String * 2            ' �i�r�w�a�l�c�R�M�����@

    HSXCNMIN As Double                ' �i�r�w�Y�f�Z�x����
    HSXCNMAX As Double                ' �i�r�w�Y�f�Z�x���
    HSXCNSPH As String * 1            ' �i�r�w�Y�f�Z�x����ʒu�Q��
    HSXCNSPT As String * 1            ' �i�r�w�Y�f�Z�x����ʒu�Q�_
    HSXCNSPI As String * 1            ' �i�r�w�Y�f�Z�x����ʒu�Q��
    HSXCNHWT As String * 1            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    HSXCNHWS As String * 1            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��

    HSXDENMX As Integer               ' �i�r�w�c�������
    HSXDENMN As Integer               ' �i�r�w�c��������
    HSXLDLMX As Integer               ' �i�r�w�k�^�c�k���
    HSXLDLMN As Integer               ' �i�r�w�k�^�c�k����
    HSXDVDMX As Integer               ' �i�r�w�c�u�c�Q���
    HSXDVDMN As Integer               ' �i�r�w�c�u�c�Q����
    HSXDENHT As String * 1            ' �i�r�w�c�����ۏؕ��@�Q��
    HSXDENHS As String * 1            ' �i�r�w�c�����ۏؕ��@�Q��
    HSXLDLHT As String * 1            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    HSXLDLHS As String * 1            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    HSXDVDHT As String * 1            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    HSXDVDHS As String * 1            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    HSXDENKU As String * 1            ' �i�r�w�c���������L��
    HSXDVDKU As String * 1            ' �i�r�w�c�u�c�Q�����L��
    HSXLDLKU As String * 1            ' �i�r�w�k�^�c�k�����L��

    HSXLTMIN As Integer               ' �i�r�w�k�^�C������
    HSXLTMAX As Integer               ' �i�r�w�k�^�C�����
    HSXLTSPH As String * 1            ' �i�r�w�k�^�C������ʒu�Q��
    HSXLTSPT As String * 1            ' �i�r�w�k�^�C������ʒu�Q�_
    HSXLTSPI As String * 1            ' �i�r�w�k�^�C������ʒu�Q��
    HSXLTHWT As String * 1            ' �i�r�w�k�^�C���ۏؕ��@�Q��
    HSXLTHWS As String * 1            ' �i�r�w�k�^�C���ۏؕ��@�Q��
    '���������Ǘ�
    EPDUP As Integer                  ' EPD�@���
End Type


' �����T���v���Ǘ��擾�p (TOP,TAIL���łQ���R�[�h�擾)
Public Type type_DBDRV_scmzc_fcmkc001c_CrySmp
    CRYNUM As String * 12             ' �����ԍ�
    IngotPos As Integer               ' �������ʒu
    LENGTH As Integer                 ' ����
    BLOCKID As String * 12            ' �u���b�NID
    SMPKBN As String * 1              ' �T���v���敪
    SMPLNO As Integer                 ' �T���v��No
    HINBAN As String * 12             ' �i��
    REVNUM As Integer                 ' ���i�ԍ������ԍ�
    factory As String * 1             ' �H��
    opecond As String * 1             ' ���Ə���
    KTKBN  As String * 1              ' �m��敪
    CRYINDRS As String * 1            ' ���������w���iRs)
    CRYINDOI As String * 1            ' ���������w���iOi)
    CRYINDB1 As String * 1            ' ���������w���iB1)
    CRYINDB2 As String * 1            ' ���������w���iB2�j
    CRYINDB3 As String * 1            ' ���������w���iB3)
    CRYINDL1 As String * 1            ' ���������w���iL1)
    CRYINDL2 As String * 1            ' ���������w���iL2)
    CRYINDL3 As String * 1            ' ���������w���iL3)
    CRYINDL4 As String * 1            ' ���������w���iL4)
    CRYINDCS As String * 1            ' ���������w���iCs)
    CRYINDGD As String * 1            ' ���������w���iGD)
    CRYINDT As String * 1             ' ���������w���iT)
    CRYINDEP As String * 1            ' ���������w���iEPD)
End Type


'������R����
Public Type type_DBDRV_scmzc_fcmkc001c_CryR
    CRYNUM As String * 12             ' �����ԍ�
    POSITION As Integer               ' �ʒu
    SMPKBN As String * 1              ' �T���v���敪
    SMPLNO As Integer                 ' �T���v���m��
    SMPLUMU As String * 1             ' �T���v���L��
    TRANCOND As String * 1            ' ��������
    MEAS1 As Double                   ' ����l�P
    MEAS2 As Double                   ' ����l�Q
    MEAS3 As Double                   ' ����l�R
    MEAS4 As Double                   ' ����l�S
    MEAS5 As Double                   ' ����l�T
    RRG As Double                     ' �q�q�f
    REGDATE As Date                   ' �o�^���t
End Type


'Oi����
Public Type type_DBDRV_scmzc_fcmkc001c_Oi
    CRYNUM As String * 12             ' �����ԍ�
    POSITION As Integer               ' �ʒu
    SMPKBN As String * 1              ' �T���v���敪
    SMPLNO As Integer                 ' �T���v���m��
    SMPLUMU As String * 1             ' �T���v���L��
    TRANCOND As String * 1            ' ��������
    OIMEAS1 As Double                 ' �n������l�P
    OIMEAS2 As Double                 ' �n������l�Q
    OIMEAS3 As Double                 ' �n������l�R
    OIMEAS4 As Double                 ' �n������l�S
    OIMEAS5 As Double                 ' �n������l�T
    ORGRES As Double                  ' �n�q�f����
    AVE As Double                     ' �`�u�d
    FTIRCONV As Double                ' �e�s�h�q���Z
    INSPECTWAY As String * 2          ' �������@
    REGDATE As Date                   ' �o�^���t
End Type


'BMD1�`3����
Public Type type_DBDRV_scmzc_fcmkc001c_BMD
    CRYNUM As String * 12             ' �����ԍ�
    POSITION As Integer               ' �ʒu
    SMPKBN As String * 1              ' �T���v���敪
    SMPLNO As Integer                 ' �T���v���m��
    SMPLUMU As String * 1             ' �T���v���L��
    HTPRC As String * 2               ' �M�������@
    KKSP As String * 3                ' �������ב���ʒu
    KKSET As String * 3               ' �������ב�������{�I��ET��@�@char(1)�{number(2)
    TRANCOND As String * 1            ' ��������
    MEAS1 As Double                   ' ����l�P
    MEAS2 As Double                   ' ����l�Q
    MEAS3 As Double                   ' ����l�R
    MEAS4 As Double                   ' ����l�S
    MEAS5 As Double                   ' ����l�T
    Min As Double                     ' MIN
    max As Double                     ' MAX
    AVE As Double                     ' AVE
    REGDATE As Date                   ' �o�^���t
End Type


'OSF1�`4����
Public Type type_DBDRV_scmzc_fcmkc001c_OSF
    CRYNUM As String * 12             ' �����ԍ�
    POSITION As Integer               ' �ʒu
    SMPKBN As String * 1              ' �T���v���敪
    SMPLNO As Integer                 ' �T���v���m��
    SMPLUMU As String * 1             ' �T���v���L��
    HTPRC As String * 2               ' �M�������@
    KKSP As String * 3                ' �������ב���ʒu
    KKSET As String * 3               ' �������ב�������{�I��ET��@�@char(1)�{number(2)
    TRANCOND As String * 1            ' ��������
    CALCMAX As Double                 ' �v�Z���� Max
    CALCAVE As Double                 ' �v�Z���� Ave
    MEAS1 As Double                   ' ����l�P
    MEAS2 As Double                   ' ����l�Q
    MEAS3 As Double                   ' ����l�R
    MEAS4 As Double                   ' ����l�S
    MEAS5 As Double                   ' ����l�T
    MEAS6 As Double                   ' ����l�U
    MEAS7 As Double                   ' ����l�V
    MEAS8 As Double                   ' ����l�W
    MEAS9 As Double                   ' ����l�X
    MEAS10 As Double                  ' ����l�P�O
    MEAS11 As Double                  ' ����l�P�P
    MEAS12 As Double                  ' ����l�P�Q
    MEAS13 As Double                  ' ����l�P�R
    MEAS14 As Double                  ' ����l�P�S
    MEAS15 As Double                  ' ����l�P�T
    MEAS16 As Double                  ' ����l�P�U
    MEAS17 As Double                  ' ����l�P�V
    MEAS18 As Double                  ' ����l�P�W
    MEAS19 As Double                  ' ����l�P�X
    MEAS20 As Double                  ' ����l�Q�O
    REGDATE As Date                   ' �o�^���t
End Type


'CS����
Public Type type_DBDRV_scmzc_fcmkc001c_CS
    CRYNUM As String * 12             ' �����ԍ�
    POSITION As Integer               ' �ʒu
    SMPKBN As String * 1              ' �T���v���敪
    SMPLNO As Integer                 ' �T���v���m��
    SMPLUMU As String * 1             ' �T���v���L��
    TRANCOND As String * 1            ' ��������
    CSMEAS As Double                  ' Cs�����l
    PRE70P As Double                  ' �V�O������l
    REGDATE As Date                   ' �o�^���t
End Type


'GD����
Public Type type_DBDRV_scmzc_fcmkc001c_GD
    CRYNUM As String * 12             ' �����ԍ�
    POSITION As Integer               ' �ʒu
    SMPKBN As String * 1              ' �T���v���敪
    SMPLNO As Integer                 ' �T���v���m��
    SMPLUMU As String * 1             ' �T���v���L��
    TRANCOND As String * 1            ' ��������
    MSRSDEN As Integer                ' ���茋�� Den
    MSRSLDL As Integer                ' ���茋�� L/DL
    MSRSDVD2 As Integer               ' ���茋�� DVD2
    MS01LDL1 As Integer             ' ����l01 L/DL1
    MS01LDL2 As Integer             ' ����l01 L/DL2
    MS01LDL3 As Integer             ' ����l01 L/DL3
    MS01LDL4 As Integer             ' ����l01 L/DL4
    MS01LDL5 As Integer             ' ����l01 L/DL5
    MS01DEN1 As Integer             ' ����l01 Den1
    MS01DEN2 As Integer             ' ����l01 Den2
    MS01DEN3 As Integer             ' ����l01 Den3
    MS01DEN4 As Integer             ' ����l01 Den4
    MS01DEN5 As Integer             ' ����l01 Den5
    MS02LDL1 As Integer             ' ����l02 L/DL1
    MS02LDL2 As Integer             ' ����l02 L/DL2
    MS02LDL3 As Integer             ' ����l02 L/DL3
    MS02LDL4 As Integer             ' ����l02 L/DL4
    MS02LDL5 As Integer             ' ����l02 L/DL5
    MS02DEN1 As Integer             ' ����l02 Den1
    MS02DEN2 As Integer             ' ����l02 Den2
    MS02DEN3 As Integer             ' ����l02 Den3
    MS02DEN4 As Integer             ' ����l02 Den4
    MS02DEN5 As Integer             ' ����l02 Den5
    MS03LDL1 As Integer             ' ����l03 L/DL1
    MS03LDL2 As Integer             ' ����l03 L/DL2
    MS03LDL3 As Integer             ' ����l03 L/DL3
    MS03LDL4 As Integer             ' ����l03 L/DL4
    MS03LDL5 As Integer             ' ����l03 L/DL5
    MS03DEN1 As Integer             ' ����l03 Den1
    MS03DEN2 As Integer             ' ����l03 Den2
    MS03DEN3 As Integer             ' ����l03 Den3
    MS03DEN4 As Integer             ' ����l03 Den4
    MS03DEN5 As Integer             ' ����l03 Den5
    MS04LDL1 As Integer             ' ����l04 L/DL1
    MS04LDL2 As Integer             ' ����l04 L/DL2
    MS04LDL3 As Integer             ' ����l04 L/DL3
    MS04LDL4 As Integer             ' ����l04 L/DL4
    MS04LDL5 As Integer             ' ����l04 L/DL5
    MS04DEN1 As Integer             ' ����l04 Den1
    MS04DEN2 As Integer             ' ����l04 Den2
    MS04DEN3 As Integer             ' ����l04 Den3
    MS04DEN4 As Integer             ' ����l04 Den4
    MS04DEN5 As Integer             ' ����l04 Den5
    MS05LDL1 As Integer             ' ����l05 L/DL1
    MS05LDL2 As Integer             ' ����l05 L/DL2
    MS05LDL3 As Integer             ' ����l05 L/DL3
    MS05LDL4 As Integer             ' ����l05 L/DL4
    MS05LDL5 As Integer             ' ����l05 L/DL5
    MS05DEN1 As Integer             ' ����l05 Den1
    MS05DEN2 As Integer             ' ����l05 Den2
    MS05DEN3 As Integer             ' ����l05 Den3
    MS05DEN4 As Integer             ' ����l05 Den4
    MS05DEN5 As Integer             ' ����l05 Den5
    MS06LDL1 As Integer             ' ����l06 L/DL1
    MS06LDL2 As Integer             ' ����l06 L/DL2
    MS06LDL3 As Integer             ' ����l06 L/DL3
    MS06LDL4 As Integer             ' ����l06 L/DL4
    MS06LDL5 As Integer             ' ����l06 L/DL5
    MS06DEN1 As Integer             ' ����l06 Den1
    MS06DEN2 As Integer             ' ����l06 Den2
    MS06DEN3 As Integer             ' ����l06 Den3
    MS06DEN4 As Integer             ' ����l06 Den4
    MS06DEN5 As Integer             ' ����l06 Den5
    MS07LDL1 As Integer             ' ����l07 L/DL1
    MS07LDL2 As Integer             ' ����l07 L/DL2
    MS07LDL3 As Integer             ' ����l07 L/DL3
    MS07LDL4 As Integer             ' ����l07 L/DL4
    MS07LDL5 As Integer             ' ����l07 L/DL5
    MS07DEN1 As Integer             ' ����l07 Den1
    MS07DEN2 As Integer             ' ����l07 Den2
    MS07DEN3 As Integer             ' ����l07 Den3
    MS07DEN4 As Integer             ' ����l07 Den4
    MS07DEN5 As Integer             ' ����l07 Den5
    MS08LDL1 As Integer             ' ����l08 L/DL1
    MS08LDL2 As Integer             ' ����l08 L/DL2
    MS08LDL3 As Integer             ' ����l08 L/DL3
    MS08LDL4 As Integer             ' ����l08 L/DL4
    MS08LDL5 As Integer             ' ����l08 L/DL5
    MS08DEN1 As Integer             ' ����l08 Den1
    MS08DEN2 As Integer             ' ����l08 Den2
    MS08DEN3 As Integer             ' ����l08 Den3
    MS08DEN4 As Integer             ' ����l08 Den4
    MS08DEN5 As Integer             ' ����l08 Den5
    MS09LDL1 As Integer             ' ����l09 L/DL1
    MS09LDL2 As Integer             ' ����l09 L/DL2
    MS09LDL3 As Integer             ' ����l09 L/DL3
    MS09LDL4 As Integer             ' ����l09 L/DL4
    MS09LDL5 As Integer             ' ����l09 L/DL5
    MS09DEN1 As Integer             ' ����l09 Den1
    MS09DEN2 As Integer             ' ����l09 Den2
    MS09DEN3 As Integer             ' ����l09 Den3
    MS09DEN4 As Integer             ' ����l09 Den4
    MS09DEN5 As Integer             ' ����l09 Den5
    MS10LDL1 As Integer             ' ����l10 L/DL1
    MS10LDL2 As Integer             ' ����l10 L/DL2
    MS10LDL3 As Integer             ' ����l10 L/DL3
    MS10LDL4 As Integer             ' ����l10 L/DL4
    MS10LDL5 As Integer             ' ����l10 L/DL5
    MS10DEN1 As Integer             ' ����l10 Den1
    MS10DEN2 As Integer             ' ����l10 Den2
    MS10DEN3 As Integer             ' ����l10 Den3
    MS10DEN4 As Integer             ' ����l10 Den4
    MS10DEN5 As Integer             ' ����l10 Den5
    MS11LDL1 As Integer             ' ����l11 L/DL1
    MS11LDL2 As Integer             ' ����l11 L/DL2
    MS11LDL3 As Integer             ' ����l11 L/DL3
    MS11LDL4 As Integer             ' ����l11 L/DL4
    MS11LDL5 As Integer             ' ����l11 L/DL5
    MS11DEN1 As Integer             ' ����l11 Den1
    MS11DEN2 As Integer             ' ����l11 Den2
    MS11DEN3 As Integer             ' ����l11 Den3
    MS11DEN4 As Integer             ' ����l11 Den4
    MS11DEN5 As Integer             ' ����l11 Den5
    MS12LDL1 As Integer             ' ����l12 L/DL1
    MS12LDL2 As Integer             ' ����l12 L/DL2
    MS12LDL3 As Integer             ' ����l12 L/DL3
    MS12LDL4 As Integer             ' ����l12 L/DL4
    MS12LDL5 As Integer             ' ����l12 L/DL5
    MS12DEN1 As Integer             ' ����l12 Den1
    MS12DEN2 As Integer             ' ����l12 Den2
    MS12DEN3 As Integer             ' ����l12 Den3
    MS12DEN4 As Integer             ' ����l12 Den4
    MS12DEN5 As Integer             ' ����l12 Den5
    MS13LDL1 As Integer             ' ����l13 L/DL1
    MS13LDL2 As Integer             ' ����l13 L/DL2
    MS13LDL3 As Integer             ' ����l13 L/DL3
    MS13LDL4 As Integer             ' ����l13 L/DL4
    MS13LDL5 As Integer             ' ����l13 L/DL5
    MS13DEN1 As Integer             ' ����l13 Den1
    MS13DEN2 As Integer             ' ����l13 Den2
    MS13DEN3 As Integer             ' ����l13 Den3
    MS13DEN4 As Integer             ' ����l13 Den4
    MS13DEN5 As Integer             ' ����l13 Den5
    MS14LDL1 As Integer             ' ����l14 L/DL1
    MS14LDL2 As Integer             ' ����l14 L/DL2
    MS14LDL3 As Integer             ' ����l14 L/DL3
    MS14LDL4 As Integer             ' ����l14 L/DL4
    MS14LDL5 As Integer             ' ����l14 L/DL5
    MS14DEN1 As Integer             ' ����l14 Den1
    MS14DEN2 As Integer             ' ����l14 Den2
    MS14DEN3 As Integer             ' ����l14 Den3
    MS14DEN4 As Integer             ' ����l14 Den4
    MS14DEN5 As Integer             ' ����l14 Den5
    MS15LDL1 As Integer             ' ����l15 L/DL1
    MS15LDL2 As Integer             ' ����l15 L/DL2
    MS15LDL3 As Integer             ' ����l15 L/DL3
    MS15LDL4 As Integer             ' ����l15 L/DL4
    MS15LDL5 As Integer             ' ����l15 L/DL5
    MS15DEN1 As Integer             ' ����l15 Den1
    MS15DEN2 As Integer             ' ����l15 Den2
    MS15DEN3 As Integer             ' ����l15 Den3
    MS15DEN4 As Integer             ' ����l15 Den4
    MS15DEN5 As Integer             ' ����l15 Den5
    REGDATE As Date                   ' �o�^���t
End Type


'���C�t�^�C�����ю擾�֐�
Public Type type_DBDRV_scmzc_fcmkc001c_LT
    CRYNUM As String * 12             ' �����ԍ�
    POSITION As Integer               ' �ʒu
    SMPKBN As String * 1              ' �T���v���敪
    SMPLNO As Integer                 ' �T���v���m��
    SMPLUMU As String * 1             ' �T���v���L��
    MEAS1 As Integer                  ' ����l�P
    MEAS2 As Integer                  ' ����l�Q
    MEAS3 As Integer                  ' ����l�R
    MEAS4 As Integer                  ' ����l�S
    MEAS5 As Integer                  ' ����l�T
    TRANCOND As String * 1            ' ��������
    MEASPEAK As Integer               ' ����l �s�[�N�l
    CALCMEAS As Integer               ' �v�Z����
    REGDATE As Date                   ' �o�^���t
    LTSPI As String                 '����ʒu�R�[�h
End Type


'EPD���ю擾�֐�
Public Type type_DBDRV_scmzc_fcmkc001c_EPD
    CRYNUM As String * 12             ' �����ԍ�
    POSITION As Integer               ' �ʒu
    SMPKBN As String * 1              ' �T���v���敪
    SMPLNO As Integer                 ' �T���v���m��
    SMPLUMU As String * 1             ' �T���v���L��
    TRANCOND As String * 1            ' ��������
    MEASURE As Integer                ' ����l
    REGDATE As Date                   ' �o�^���t
End Type


'���т��܂Ƃ߂��\����
Public Type type_DBDRV_scmzc_fcmkc001c_Zisseki
    CRYRZ() As type_DBDRV_scmzc_fcmkc001c_CryR
    OIZ() As type_DBDRV_scmzc_fcmkc001c_Oi
    BMD1Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD2Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD3Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    OSF1Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF2Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF3Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF4Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    csz() As type_DBDRV_scmzc_fcmkc001c_CS
    GDZ() As type_DBDRV_scmzc_fcmkc001c_GD
    LTZ() As type_DBDRV_scmzc_fcmkc001c_LT
    EPDZ() As type_DBDRV_scmzc_fcmkc001c_EPD
    SURSZ() As type_DBDRV_scmzc_fcmkc001c_CryR
End Type


'�u���b�N�Ǘ��X�V�p�i���ݍH���A�ŏI�ʉߍH���j
Public Type type_DBDRV_scmzc_fcmkc001c_UpdBlock1
    CRYNUM As String * 12           ' �����ԍ�
    IngotPos As Integer             ' �������J�n�ʒu
    NOWPROC As String * 5           ' ���ݍH��
    LASTPASS As String * 5          ' �ŏI�ʉߍH��
End Type


'�u���b�N�Ǘ��X�V�p�i�폜�敪�A�ŏI��ԋ敪�A������ԋ敪�j
Public Type type_DBDRV_scmzc_fcmkc001c_UpdBlock2
    CRYNUM As String * 12           ' �����ԍ�
    IngotPos As Integer             ' �������J�n�ʒu
    DELCLS As String * 1            ' �폜�敪
    LSTATCLS As String * 1          ' �ŏI��ԋ敪
    RSTATCLS As String * 1          ' ������ԋ敪
End Type

'�u���b�N�Ǘ��X�V�p�i�N���X�^���J�^���O�A�������g�p�j
Public Type typ_DBDRV_fcmkc001c_UpdBlkCR
    CRYNUM As String * 12           ' �����ԍ�
    IngotPos As Integer             ' �������J�n�ʒu
    NOWPROC As String * 5           ' ���ݍH��
'    LASTPASS As String * 5          ' �ŏI�ʉߍH��
    DELCLS As String * 1            ' �폜�敪
    BDCAUS As String * 3            ' �s�Ǘ��R
    LSTATCLS As String * 1          ' �ŏI��ԋ敪
    RSTATCLS As String * 1          ' ������ԋ敪
End Type



'�����T���v���Ǘ��X�V�p
Public Type type_DBDRV_scmzc_fcmkc001c_UpdCrySmp
    CRYNUM As String * 12           ' �����ԍ�
    IngotPos As Integer             ' �������ʒu
    SMPKBN As String * 1            ' �T���v���敪
End Type

'���茋�ʂ�J014�����v�ۍ\����
Public Type Judg_Spec_Cry
    Enable As Boolean           '�L���ȕi�Ԃł���
    rs As Boolean               'Rs�͗v����
    Oi As Boolean               'Oi�͗v����
    B1 As Boolean               'BMD1�͗v����
    B2 As Boolean               'BMD2�͗v����
    B3 As Boolean               'BMD3�͗v����
    L1 As Boolean               'OSF1�͗v����
    L2 As Boolean               'OSF2�͗v����
    L3 As Boolean               'OSF3�͗v����
    L4 As Boolean               'OSF4�͗v����
    Cs As Boolean               'Cs�͗v����
    GD As Boolean               'GD�͗v����
    Lt As Boolean               'LT�͗v����
    EPD As Boolean              'EPD�͗v����
End Type

' �d�l�̎w���������Ă��锻�f�p
Public Const SIJI = "H"
Public Const SANKOU = "S"

'2002/08/01 M.Tomita------------------------------------------------------

'===========================================
' �v�e���H�p���ʃe�[�u��
'===========================================

' �����w��
Public Type typ_WafInd
    BLOCKID As String * 12      ' �u���b�NID
    BlockPos As Integer         ' �u���b�N�o
    IngotPos As Integer         ' �����o
    LENGTH As Integer           ' ����
    HINUP As tFullHinban        ' ��i��
    HINDN As tFullHinban        ' ���i��
    SMP As typ_WFSample         ' ��������
    HINFLG As Boolean           ' �i�ԋ�؂�t���O
    SMPFLG As Boolean           ' WF�T���v����؂�t���O
    ERRDNFLG As Boolean         ' ���i�ԃG���[�t���O
    SMPLKBN1 As String * 1      ' �T���v���敪�P
    SMPLKBN2 As String * 1      ' �T���v���敪�Q
End Type

' ���i�d�l
Public Type typ_HinSpec
    hin As tFullHinban          ' �i��
    IngotPos As Integer         ' �������J�n�ʒu
    LENGTH As Integer           ' ����
    HWFRMIN As Double           ' ���R����
    HWFRMAX As Double           ' ���R���
    HWFRHWYS As String * 1      ' �����L��(Rs)
    HWFONHWS As String * 1      ' �����L��(Oi)
    HWFBM1HS As String * 1      ' �����L��(B1)
    HWFBM2HS As String * 1      ' �����L��(B2)
    HWFBM3HS As String * 1      ' �����L��(B3)
    HWFOF1HS As String * 1      ' �����L��(L1)
    HWFOF2HS As String * 1      ' �����L��(L2)
    HWFOF3HS As String * 1      ' �����L��(L3)
    HWFOF4HS As String * 1      ' �����L��(L4)
    HWFDSOHS As String * 1      ' �����L��(DS)
    HWFMKHWS As String * 1      ' �����L��(DZ)
    HWFSPVHS As String * 1      ' �����L��(SP/Fe�Z�x)
    HWFDLHWS As String * 1      ' �����L��(SP/�g�U��)
    HWFOS1HS As String * 1      ' �����L��(D1)
    HWFOS2HS As String * 1      ' �����L��(D2)
    HWFOS3HS As String * 1      ' �����L��(D3)
    HWFOTHER1 As String * 1     ' �����L��(OT2) ''Add.03/05/20 �㓡
    HWFOTHER2 As String * 1     ' �����L��(OT1) ''Add.03/05/20
End Type

' �����E�F�n�[
Public Type typ_LackMap
    BLOCKID As String * 12      ' �u���b�NID
    LACKPOSS As Double          ' �����ʒu(From)
    LACKPOSE As Double          ' �����ʒu(To)
    REJCAT As String * 1        ' �������R
    LACKCNTS As Integer         ' ��������(From)
    LACKCNTE As Integer         ' ��������(To)
End Type

'�e���я��
Public Type typ_ALLRSLT
    pos As Integer                    ' �������J�n�ʒu
    NAIYO As String                   ' ���e
    INFO1 As String                   ' ���P
    INFO2 As String                   ' ���Q
    INFO3 As String                   ' ���R
    INFO4 As String                   ' ���S
    OKNG  As String                   ' ���茋��
    SMPLNO As Integer                 ' �T���v���m��
    BLOCKNG As Boolean                'GD�G���[�ƂȂ�i�Ԃ��܂ނ�����
End Type

'�S���\����
Type typ_AllTypes
    intPFlg As Integer                              ' �\���t���O
    strStaffID As String                            ' �X�^�b�tID
    strStaffName As String                          ' �X�^�b�t��
    BLOCKID  As String * 12                         ' �u���b�NID
    Cut(2) As Double                                ' �ăJ�b�g�ʒu
    COEF(2) As Double                               ' �ΐ͌W��
    CRCOEF As Double                                ' �����ΐ͌W��
    OKNG(2) As Boolean                              ' ���R����
    Henseki As Boolean                              ' ���R���їL��(�����S��TOP/TAIL)
    JudgRes(2) As Boolean                              ' ���R����    2001/10/02 S.Sano
    JudgRrg(2) As Boolean                              ' RRG����       2001/10/02 S.Sano
    typ_rsz() As typ_TBCMJ002                       ' ������R����(�����S��TOP/TAIL)
    typ_hage(2) As typ_TBCMH004                     ' ���グ�I������
    typ_rslt(2, MAXCNT) As typ_ALLRSLT              ' �e���я��
    typ_zi As type_DBDRV_scmzc_fcmkc001c_Zisseki    ' ���т��܂Ƃ߂��\����
    typ_si() As type_DBDRV_scmzc_fcmkc001c_Siyou    ' �d�l
    typ_cr() As type_DBDRV_scmzc_fcmkc001c_CrySmp   ' �����T���v���Ǘ��擾�p (TOP,TAIL���łQ���R�[�h�擾)
    blYONE As Boolean                               ' �đ�t���O
End Type

Public typ_A As typ_AllTypes        '�S���\����
Public JudgSC(2) As Judg_Spec_Cry        '�d�l�����x���\����
Public TotalJudg As Boolean         '�g�[�^������
Public MeasFlag(2) As Judg_Spec_Cry        '�d�l�����x���\����
Public Kakou As type_KakouJudg      '���H���є���\����


'�u���b�N���x�����o��  4/16 Yam�쐬

' �u���b�N�ꗗ
Public Type typ_BlkLbl
    BLOCKID As String * 12      ' �u���b�NID
    hin(5) As tFullHinban       ' �i��
    WFINDDATE As String * 10    ' �ŏI�������t
    CRYNUM As String * 12       ' �����ԍ�
    IngotPos As Integer         ' �C���S�b�g���ʒu
    LENGTH As Integer           ' �u���b�N����
    REALLEN As Integer          ' �u���b�N������
    HINLEN(5) As Integer        ' �i�Ԓ���
    DIAMETER As Integer         ' ���a
    SBLOCKID As String * 12     ' �擪�u���b�NID
    BLOCKORDER As Integer       ' �u���b�N����
    HOLDCLS As String * 1       ' �z�[���h���  --- 2001/09/19 kuramoto �ǉ� ---
    PASSFLAG As String * 1      ' �ʉ߃t���O�@�@--- 200/04/16 Yam
End Type



'�T�v      :�����w���p ��ʕ\�����c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sBlockID�@�@�@,I  ,String         �@,�u���b�NID
'      �@�@:pCryInf �@�@�@,O  ,typ_TBCME037   �@,�������
'      �@�@:pHinDsn �@�@�@,O  ,typ_TBCME039   �@,�i�Ԑ݌v
'      �@�@:pHinMng �@�@�@,O  ,typ_TBCME041   �@,�i�ԊǗ�
'      �@�@:pBlkInf �@�@�@,O  ,typ_BlkInf1    �@,�u���b�N���
'      �@�@:pHinSpec�@�@�@,O  ,typ_HinSpec    �@,���i�d�l
'      �@�@:dNeraiRes �@�@,O  ,Double         �@,�˂炢�i�Ԃ̔��R����l�iP+�̔��f�p�j
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmkc001g_Disp(ByVal SBLOCKID As String, pCryInf As typ_TBCME037, _
                                           pHinDsn() As typ_TBCME039, pHinMng() As typ_TBCME041, _
                                           pBlkInf() As typ_BlkInf1, pHinSpec() As typ_HinSpec, _
                                           dNeraiRes As Double, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim rs As OraDynaset
    Dim sql As String
    Dim sDbName As String
    Dim sCryNum As String
    Dim sHin As String
    Dim sSeed As String
    Dim dMenseki As Double
    Dim dTopWght As Double
    Dim dCharge As Double
    Dim dMeas(4) As Double
    Dim bFlag As Boolean
    Dim recCnt As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001g_SQL.bas -- Function DBDRV_scmzc_fcmkc001g_Disp"
    sErrMsg = ""

    '' �u���b�N�Ǘ��̎擾
    sDbName = "E040"
    sCryNum = Left(SBLOCKID, 9) & "000"
    sql = "select "
    sql = sql & "INGOTPOS, LENGTH, REALLEN, BLOCKID, "
    sql = sql & "KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, RSTATCLS"
    sql = sql & " from TBCME040 where CRYNUM='" & sCryNum & "'"
    sql = sql & " and INGOTPOS>=0 and LENGTH>0 order by INGOTPOS"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    bFlag = False
    ReDim pBlkInf(recCnt)
    For i = 1 To recCnt
        With pBlkInf(i)
            .COF.TOPSMPLPOS = rs("INGOTPOS")
            .LENGTH = rs("LENGTH")
            .REALLEN = rs("REALLEN")
            .BLOCKID = rs("BLOCKID")
            .KRPROCCD = rs("KRPROCCD")
            .NOWPROC = rs("NOWPROC")
            .LPKRPROCCD = rs("LPKRPROCCD")
            .LASTPASS = rs("LASTPASS")
            .RSTATCLS = rs("RSTATCLS")
            .COF.BOTSMPLPOS = .COF.TOPSMPLPOS + .LENGTH
            .SAMPFLAG = False
            If .BLOCKID = SBLOCKID Then
                bFlag = True
            End If
        End With
        rs.MoveNext
    Next i
    rs.Close

    '' �u���b�NID���݃`�F�b�N
    If bFlag = False Then
        sErrMsg = GetMsgStr("EBLK0")
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �������̎擾(s_cmzcTBCME037_SQL.bas ���K�v)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)

    '' �i�Ԑ݌v�̎擾(s_cmzcTBCME039_SQL.bas ���K�v)
    sDbName = "E039"
    sql = " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' and LENGTH>0 order by INGOTPOS"
    If DBDRV_GetTBCME039(pHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �i�ԊǗ��̎擾(s_cmzcTBCME041_SQL.bas ���K�v)
    sDbName = "E041"
    sql = " where CRYNUM='" & sCryNum & "'�@and LENGTH>0 order by INGOTPOS"
    If DBDRV_GetTBCME041(pHinMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' ���グ�I�����т̎擾
    sDbName = "H004"
    sql = "select (DM1+DM2+DM3)/3.0 as DM, WGHTTOP, CHARGE, SEED from TBCMH004 where CRYNUM='" & sCryNum & "'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        dMenseki = AreaOfCircle(rs("DM"))
        dTopWght = rs("WGHTTOP")
        dCharge = rs("CHARGE")
        sSeed = rs("SEED")
    Else
        dMenseki = 0
        dTopWght = 0
        dCharge = 0
        sSeed = ""
    End If
    rs.Close

    '' ������R���т̎擾
    sDbName = "J002"
    recCnt = UBound(pBlkInf)
    For i = 1 To recCnt
        With pBlkInf(i)
            .SEED = sSeed                   ' �V�[�h
            .COF.DUNMENSEKI = dMenseki      ' �f�ʐ�
            .COF.CHARGEWEIGHT = dCharge     ' �`���[�W��
            .COF.TOPWEIGHT = dTopWght       ' �g�b�v�d��

            '' �g�b�v�����R�����l�̎擾
            sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.TOPSMPLPOS & " and SMPKBN='T'"
            sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.TOPSMPLPOS & " and SMPKBN='T')"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount > 0 Then
                dMeas(0) = rs("MEAS1")
                dMeas(1) = rs("MEAS2")
                dMeas(2) = rs("MEAS3")
                dMeas(3) = rs("MEAS4")
                dMeas(4) = rs("MEAS5")
                .COF.TOPRES = JudgCenter(dMeas())
            Else
                .COF.TOPRES = -9999
            End If
            rs.Close

            '' �{�g�������R�����l�̎擾
            sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='B'"
            sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='B')"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
                sql = sql & " where CRYNUM='" & sCryNum & "'"
                sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='T'"
                sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
                sql = sql & " where CRYNUM='" & sCryNum & "'"
                sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='T')"
                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            End If
            If rs.RecordCount > 0 Then
                dMeas(0) = rs("MEAS1")
                dMeas(1) = rs("MEAS2")
                dMeas(2) = rs("MEAS3")
                dMeas(3) = rs("MEAS4")
                dMeas(4) = rs("MEAS5")
                .COF.BOTRES = JudgCenter(dMeas())
            Else
                .COF.BOTRES = -9999
            End If
            rs.Close
        End With
    Next i

    '' ���i�d�l�̎擾
    sDbName = "VE004"
    recCnt = UBound(pHinMng)
    ReDim pHinSpec(recCnt)
    k = 0
    For i = 1 To recCnt
        With pHinMng(i)
            sHin = RTrim$(.HINBAN)
            If sHin <> "" And sHin <> "G" And sHin <> "Z" Then
                For j = 1 To k
                    If pHinSpec(j).hin.HINBAN = .HINBAN Then
                        pHinSpec(j).LENGTH = pHinSpec(j).LENGTH + .LENGTH
                        Exit For
                    End If
                Next j
                If j > k Then
                    k = k + 1
                    pHinSpec(k).IngotPos = .IngotPos
                    pHinSpec(k).hin.HINBAN = .HINBAN
                    pHinSpec(k).hin.mnorevno = .REVNUM
                    pHinSpec(k).hin.factory = .factory
                    pHinSpec(k).hin.opecond = .opecond
                    pHinSpec(k).LENGTH = .LENGTH
                    If DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec(k)) = FUNCTION_RETURN_FAILURE Then
                        sErrMsg = GetMsgStr("EGET2", sDbName)
                        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
            End If
        End With
    Next i
    ReDim Preserve pHinSpec(k)

    '' �˂炢�i�Ԃ̔��R����l���擾
    sql = "select HSXRMAX"
    sql = sql & " from TBCME037 E37, TBCME018 E18"
    sql = sql & " where (E37.CRYNUM='" & Left$(SBLOCKID, 9) & "000')"
    sql = sql & " and (E37.RPHINBAN=E18.HINBAN) and (E37.RPREVNUM=E18.MNOREVNO)"
    sql = sql & " and (E37.RPFACT=E18.FACTORY) and (E37.RPOPCOND=E18.OPECOND)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        dNeraiRes = rs("HSXRMAX")
    Else
        dNeraiRes = 0#      '�����܂ł͂��Ȃ��͂�
    End If
    rs.Close

    DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("EGET2", sDbName)
    DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�����w���p ���i�d�l��p�c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:pHinSpec�@�@�@,IO ,typ_HinSpec    �@,���i�d�l
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec As typ_HinSpec) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim sOT1    As String   '03/05/21
    Dim sOT2    As String
    Dim rtn     As FUNCTION_RETURN

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001g_GetSpec"

    '' ���i�d�l�̎擾
    With pHinSpec
        sql = "select "
        sql = sql & "E021HWFRMIN, E021HWFRMAX, E021HWFRHWYS, "
        sql = sql & "E024HWFMKHWS, E025HWFONHWS, E025HWFOS1HS, E025HWFOS2HS, E025HWFOS3HS, "
        sql = sql & "E026HWFDSOHS, E028HWFSPVHS, E028HWFDLHWS, E029HWFOF1HS, E029HWFOF2HS, "
        sql = sql & "E029HWFOF3HS, E029HWFOF4HS, E029HWFBM1HS, E029HWFBM2HS, E029HWFBM3HS"
        sql = sql & " from VECME004"
        sql = sql & " where E018HINBAN='" & .hin.HINBAN & "'"
        sql = sql & " and E018MNOREVNO=" & .hin.mnorevno
        sql = sql & " and E018FACTORY='" & .hin.factory & "'"
        sql = sql & " and E018OPECOND='" & .hin.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        .HWFRMIN = fncNullCheck(rs("E021HWFRMIN"))
        .HWFRMAX = fncNullCheck(rs("E021HWFRMAX"))
        .HWFRHWYS = rs("E021HWFRHWYS")
        .HWFMKHWS = rs("E024HWFMKHWS")
        .HWFONHWS = rs("E025HWFONHWS")
        .HWFOS1HS = rs("E025HWFOS1HS")
        .HWFOS2HS = rs("E025HWFOS2HS")
        .HWFOS3HS = rs("E025HWFOS3HS")
        .HWFDSOHS = rs("E026HWFDSOHS")
        .HWFSPVHS = rs("E028HWFSPVHS")
        .HWFDLHWS = rs("E028HWFDLHWS")
        .HWFOF1HS = rs("E029HWFOF1HS")
        .HWFOF2HS = rs("E029HWFOF2HS")
        .HWFOF3HS = rs("E029HWFOF3HS")
        .HWFOF4HS = rs("E029HWFOF4HS")
        .HWFBM1HS = rs("E029HWFBM1HS")
        .HWFBM2HS = rs("E029HWFBM2HS")
        .HWFBM3HS = rs("E029HWFBM3HS")
        rtn = scmzc_getE036(pHinSpec.hin, sOT1, sOT2)   '03/05/21
        If rtn = FUNCTION_RETURN_FAILURE Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        .HWFOTHER1 = sOT1 '### 03/05/21
        .HWFOTHER2 = sOT2
 
        rs.Close
    End With

    DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�����w���p ���s���c�a�h���C�o
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:sCryNum�@�@�@,I  ,String         �@,�����ԍ�
'      �@�@:pBlkInf�@�@�@,I  ,typ_BlkInf1    �@,�u���b�N���
'      �@�@:pSXLMng�@�@�@,I  ,typ_TBCME042   �@,SXL�Ǘ�
'      �@�@:pWafSmp�@�@�@,I  ,typ_XSDCW   �@   ,�V�T���v���Ǘ��iSXL�j
'      �@�@:pCryCat�@�@�@,I  ,typ_TBCMG007   �@,�N���X�^���J�^���O�������
'      �@�@:pBsInd �@�@�@,I  ,typ_TBCMW001   �@,�����w������
'      �@�@:pMesInd�@�@�@,I  ,typ_TBCMY003   �@,����]�����@�w��
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function DBDRV_scmzc_fcmkc001g_Exec(ByVal sCryNum As String, pBlkInf() As typ_BlkInf1, _
                                           pSXLMng() As typ_TBCME042, pWafSmp() As typ_XSDCW, pCryCat() As typ_TBCMG007, _
                                           pBsInd() As typ_TBCMW001, pMesInd() As typ_TBCMY003, sErrMsg As String) As FUNCTION_RETURN

Dim sql As String
Dim sDbName As String
Dim recCnt As Long
Dim i As Long
Dim hin As tFullHinban

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001g_Exec"
    sErrMsg = ""

    '' WriteDBLog " ", "Start"

    '' �������̍X�V
    sDbName = "E037"
    sql = "update TBCME037 set "
    sql = sql & "KRPROCCD='" & MGPRCD_WFC_HARAIDASI & "', "
    sql = sql & "PROCCD='" & PROCD_WFC_HARAIDASI & "', "
    sql = sql & "LPKRPROCCD='" & MGPRCD_NUKISI_SIJI & "', "
    sql = sql & "LASTPASS='" & PROCD_NUKISI_SIJI & "', "
    sql = sql & "UPDDATE=sysdate, "
    sql = sql & "SENDFLAG='0'"
    sql = sql & " where CRYNUM='" & sCryNum & "'"
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �u���b�N�Ǘ��̍X�V
    sDbName = "E040"
    recCnt = UBound(pBlkInf)
    For i = 1 To recCnt
        With pBlkInf(i)
            sql = "update TBCME040 set "
            sql = sql & "KRPROCCD='" & .KRPROCCD & "', "
            sql = sql & "NOWPROC='" & .NOWPROC & "', "
            sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "
            sql = sql & "LASTPASS='" & .LASTPASS & "', "
            sql = sql & "RSTATCLS='" & .RSTATCLS & "', "
            sql = sql & "UPDDATE=sysdate, "
            sql = sql & "SENDFLAG='0' "
            sql = sql & " where CRYNUM='" & sCryNum & "' and INGOTPOS=" & .COF.TOPSMPLPOS
        End With
        '' WriteDBLog sql, sDbName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    
    '�i�ԊǗ��e�[�u���̍X�V
    sDbName = "E041"
    recCnt = UBound(pBlkInf)
    With hin
        .mnorevno = 0
        .factory = " "
        .opecond = " "
    End With
    For i = 1 To recCnt
        With pBlkInf(i)
            If .RSTATCLS = "G" Then
                'G�i�ԂɕύX
                hin.HINBAN = "G"
                If ChangeAreaHinban(sCryNum, CInt(.COF.TOPSMPLPOS), .LENGTH, hin) = FUNCTION_RETURN_FAILURE Then
                    DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
                End If
            ElseIf .RSTATCLS = "M" Then
                'Z�i�ԂɕύX
                hin.HINBAN = "Z"
                If ChangeAreaHinban(sCryNum, CInt(.COF.TOPSMPLPOS), .LENGTH, hin) = FUNCTION_RETURN_FAILURE Then
                    DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
                End If
            End If
        End With
    Next

    '' SXL�Ǘ��̑}��(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    sDbName = "E042"
    If DBDRV_SXL_INS(pSXLMng()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' WF�T���v���Ǘ��̑}��(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    sDbName = "E044"
'''' --TEST--
''''If DBDRV_WfSmp_INS(pWafSmp()) = FUNCTION_RETURN_FAILURE Then
    If DBDRV_WfSmp_INS(pWafSmp(), i) = FUNCTION_RETURN_FAILURE Then
        
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �N���X�^���J�^���O������т̑}��
    sDbName = "G007"
    recCnt = UBound(pCryCat)
    For i = 1 To recCnt
        With pCryCat(i)
            sql = "insert into TBCMG007 "
            sql = sql & "(CRYNUM, TRANCNT, KRPROCCD, PROCCODE, BDCODE, PALTNUM, "
            sql = sql & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE)"
            sql = sql & " select '"
            sql = sql & .CRYNUM & "', "             ' �����ԍ�
            sql = sql & "nvl(max(TRANCNT),0)+1, '"  ' ������
            sql = sql & MGPRCD_NUKISI_SIJI & "', '" ' �Ǘ��H���R�[�h
            sql = sql & PROCD_NUKISI_SIJI & "', '"  ' �H���R�[�h
            sql = sql & .BDCODE & "', '"            ' �s�Ǘ��R�R�[�h
            sql = sql & .PALTNUM & "', '"           ' �p���b�g�ԍ�
            sql = sql & .TSTAFFID & "', "           ' �o�^�Ј�ID
            sql = sql & "sysdate, '"                ' �o�^���t
            sql = sql & .KSTAFFID & "', "           ' �X�V�Ј�ID
            sql = sql & "sysdate, "                 ' �X�V���t
            sql = sql & "'0', "                     ' ���M�t���O
            sql = sql & "sysdate"                   ' ���M���t
            sql = sql & " from TBCMG007"
            sql = sql & " where CRYNUM='" & .CRYNUM & "'"
        End With
        '' WriteDBLog sql, sDbName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' �����w�����т̑}��
    sDbName = "W001"
    recCnt = UBound(pBsInd)
    For i = 1 To recCnt
        With pBsInd(i)
            sql = "insert into TBCMW001 "
            sql = sql & "(CRYNUM, INGOTPOS, TRANCNT, "
            sql = sql & "CRYLEN, KRPROCCD, PROCCODE, BLOCKID, "
            sql = sql & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE)"
            sql = sql & " select '"
            sql = sql & .CRYNUM & "', "             ' �����ԍ�
            sql = sql & .IngotPos & ", "            ' �C���S�b�g�ʒu
            sql = sql & "nvl(max(TRANCNT),0)+1, "   ' ������
            sql = sql & .CRYLEN & ", '"             ' ����
            sql = sql & MGPRCD_NUKISI_SIJI & "', '" ' �Ǘ��H���R�[�h
            sql = sql & PROCD_NUKISI_SIJI & "', '"  ' �H���R�[�h
            sql = sql & .BLOCKID & "', '"           ' �u���b�NID
            sql = sql & .TSTAFFID & "', "           ' �o�^�Ј�ID
            sql = sql & "sysdate, '"                ' �o�^���t
            sql = sql & .TSTAFFID & "', "           ' �X�V�Ј�ID
            sql = sql & "sysdate, "                 ' �X�V���t
            sql = sql & "'0', "                     ' ���M�t���O
            sql = sql & "sysdate"                   ' ���M���t
            sql = sql & " from TBCMW001"
            sql = sql & " where CRYNUM='" & sCryNum & "' and INGOTPOS=" & .IngotPos
        End With
        '' WriteDBLog sql, sDbName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' ����]�����@�w���̑}��(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    sDbName = "Y003"
    If DBDRV_SokuSizi_Ins(pMesInd()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    '' WriteDBLog " ", "End"
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
    DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME037�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME037 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcF_TBCME037_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME037(records() As typ_TBCME037, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME037"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME037 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .DELCLS = rs("DELCLS")           ' �폜�敪
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCD = rs("PROCCD")           ' �H���R�[�h
            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
            .RPHINBAN = rs("RPHINBAN")       ' �˂炢�i��
            .RPREVNUM = rs("RPREVNUM")       ' �˂炢�i�Ԑ��i�ԍ������ԍ�
            .RPFACT = rs("RPFACT")           ' �˂炢�i�ԍH��
            .RPOPCOND = rs("RPOPCOND")       ' �˂炢�i�ԑ��Ə���
            .PRODCOND = rs("PRODCOND")       ' �������
            .PGID = rs("PGID")               ' �o�f�|�h�c
            .UPLENGTH = rs("UPLENGTH")       ' ���グ����
            .TOPLENG = rs("TOPLENG")         ' �s�n�o����
            .BODYLENG = rs("BODYLENG")       ' ��������
            .BOTLENG = rs("BOTLENG")         ' �a�n�s����
            .FREELENG = rs("FREELENG")       ' �t���[��
            .DIAMETER = rs("DIAMETER")       ' ���a
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .SEED = rs("SEED")               ' �V�[�h
            .ADDDPCLS = rs("ADDDPCLS")       ' �ǉ��h�[�v���
            .ADDDPPOS = rs("ADDDPPOS")       ' �ǉ��h�[�v�ʒu
            .ADDDPVAL = rs("ADDDPVAL")       ' �ǉ��h�[�v��
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME037 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME039�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME039 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcF_TBCME039_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME039(records() As typ_TBCME039, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACT, OPCOND, LENGTH, USECLASS, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME039"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME039 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .IngotPos = rs("INGOTPOS")       ' �������J�n�ʒu
            .HINBAN = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' �����ԍ�
            .FACT = rs("FACT")               ' �H��
            .OPCOND = rs("OPCOND")           ' ���Ə���
            .LENGTH = rs("LENGTH")           ' ����
            .USECLASS = rs("USECLASS")       ' �g�p�敪
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME039 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME041�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME041 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCME041_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME041(records() As typ_TBCME041, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACTORY, OPECOND, LENGTH, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME041"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME041 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .IngotPos = rs("INGOTPOS")       ' �������J�n�ʒu
            .HINBAN = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .LENGTH = rs("LENGTH")           ' ����
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME041 = FUNCTION_RETURN_SUCCESS
End Function


'�T�v    :�҂��ꗗ �����\���p�c�a�h���C�o�i�����҂��j
'���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
'        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'����    :
'����    :2001/07/06 ���{ �쐬
Public Function DBDRV_scmzc_fcmkc001b_Disp1(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN

    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '�u���b�N�Ǘ��̃��R�[�h��
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim BlockIdBuf As String
    
    '<�����҂���
    '�u���b�N�Ǘ��e�[�u������u���b�NID�A�X�V���t�擾�i�������т��������̂��́j
    

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp1"

    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_SUCCESS

    '�u���b�NID�A�X�V���t�̎擾
    sql = "select distinct "
    sql = sql & " V.E040CRYNUM, "
    sql = sql & " V.E040INGOTPOS, "
    sql = sql & " V.E040BLOCKID, "
    sql = sql & " V.E040UPDDATE, "
    sql = sql & " V.E040HOLDCLS, "
    sql = sql & " H.HINBAN, "            ' �i��
    sql = sql & " H.REVNUM, "            ' ���i�ԍ������ԍ�
    sql = sql & " H.FACTORY, "           ' �H��
    sql = sql & " H.OPECOND, "           ' ���Ə���
    sql = sql & " S.HSXTYPE, "           ' �i�r�w�^�C�v
    sql = sql & " S.HSXCDIR, "            ' �i�r�w�����ʕ���
    sql = sql & " H.INGOTPOS "
    sql = sql & " from "
    sql = sql & " VECME010 V, TBCME041 H, TBCME018 S "
    sql = sql & " where "
    sql = sql & " V.E040CRYNUM = H.CRYNUM "
    sql = sql & " and H.HINBAN = S.HINBAN "
    sql = sql & " and H.REVNUM = S.MNOREVNO "
    sql = sql & " and H.FACTORY = S.FACTORY "
    sql = sql & " and H.OPECOND = S.OPECOND "
                '�u���b�N���̕i�Ԍ���
    sql = sql & " and (( V.E040INGOTPOS >= H.INGOTPOS "
    sql = sql & " and V.E040INGOTPOS < H.INGOTPOS + H.LENGTH ) "
    sql = sql & " or ( V.E040INGOTPOS + V.E040LENGTH > H.INGOTPOS "
    sql = sql & " and V.E040INGOTPOS + V.E040LENGTH < H.INGOTPOS + H.LENGTH  ) "
    sql = sql & " or ( H.INGOTPOS >= V.E040INGOTPOS "
    sql = sql & " and H.INGOTPOS < V.E040INGOTPOS + V.E040LENGTH ) "
    sql = sql & " or ( H.INGOTPOS + H.LENGTH > V.E040INGOTPOS "
    sql = sql & " and H.INGOTPOS + H.LENGTH < V.E040INGOTPOS + V.E040LENGTH )) "
                '�H���R�[�h�A��ԁA�敪�̏����w��
    sql = sql & " and V.E040NOWPROC='CC600' "
    sql = sql & " and V.E040LSTATCLS='T' "
    sql = sql & " and V.E040RSTATCLS='T' "
    sql = sql & " and V.E040DELCLS='0' "
    'sql = sql & " and V.E040HOLDCLS='0' " ' �z�[���h�u���b�N���擾
                '�w����0�łȂ����т�0
    sql = sql & " and ((V.E043CRYINDRS<>'0' and V.E043CRYRESRS='0') "         ' �����������сiRs)
    sql = sql & " or (V.E043CRYINDOI<>'0' and V.E043CRYRESOI='0') "         ' �����������сiOi)
    sql = sql & " or (V.E043CRYINDB1<>'0' and V.E043CRYRESB1='0')"          ' �����������сiB1)
    sql = sql & " or (V.E043CRYINDB2<>'0' and V.E043CRYRESB2='0') "         ' �����������сiB2�j
    sql = sql & " or (V.E043CRYINDB3<>'0' and V.E043CRYRESB3='0') "         ' �����������сiB3)
    sql = sql & " or (V.E043CRYINDL1<>'0' and V.E043CRYRESL1='0') "         ' �����������сiL1)
    sql = sql & " or (V.E043CRYINDL2<>'0' and V.E043CRYRESL2='0') "         ' �����������сiL2)
    sql = sql & " or (V.E043CRYINDL3<>'0' and V.E043CRYRESL3='0') "         ' �����������сiL3)
    sql = sql & " or (V.E043CRYINDL4<>'0' and V.E043CRYRESL4='0') "         ' �����������сiL4)
    sql = sql & " or (V.E043CRYINDCS<>'0' and V.E043CRYRESCS='0') "         ' �����������сiCs)
    sql = sql & " or (V.E043CRYINDGD<>'0' and V.E043CRYRESGD='0') "         ' �����������сiGD)
    sql = sql & " or (V.E043CRYINDT<>'0' and V.E043CRYREST='0') "           ' �����������сiT)
    sql = sql & " or (V.E043CRYINDEP<>'0' and V.E043CRYRESEP='0')) "         ' �����������сiEPD)
    sql = sql & " order by V.E040BLOCKID, H.INGOTPOS "

    '�f�[�^�𒊏o����
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '���R�[�h0�����͐���
    If rs.RecordCount = 0 Then
        rs.Close
        ReDim records(0)
    Else
        BlockIdBuf = vbNullString
        recCnt = rs.RecordCount
        j = 0
        For i = 1 To recCnt
            DoEvents
        '�u���b�NID���̊i�[
            If rs("E040BLOCKID") <> BlockIdBuf Then
            
                j = j + 1
                ReDim Preserve records(j)
                
                With records(j)
                    .CRYNUM = rs("E040CRYNUM")
                    .IngotPos = rs("E040INGOTPOS")
                    .BLOCKID = rs("E040BLOCKID")   ' �u���b�NID
                    .UPDDATE = rs("E040UPDDATE")   ' �X�V���t
                    .HOLDCLS = rs("E040HOLDCLS")   ' �z�[���h�敪
                    BlockIdBuf = records(j).BLOCKID
                    .HSXTYPE = rs("HSXTYPE")
                    .HSXCDIR = rs("HSXCDIR")
                    .Judg = " "
                End With
                
                k = 1
            End If
            
            '�i�Ԃ̊i�[
            ReDim Preserve records(j).hin(k)
            records(j).hin(k).HINBAN = rs("HINBAN")
            records(j).hin(k).mnorevno = rs("REVNUM")
            records(j).hin(k).factory = rs("FACTORY")
            records(j).hin(k).opecond = rs("OPECOND")
            k = k + 1
            rs.MoveNext
        Next i
        rs.Close
            
    End If
        
    

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v    :�҂��ꗗ �����\���p�c�a�h���C�o�i����҂��j
'���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
'        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'����    :
'����    :2001/07/06 ���{ �쐬
Public Function DBDRV_scmzc_fcmkc001b_Disp2(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN

    '������҂���
    '�����҂���������Ă���ꍇ�Ƌt�łO������Ȃ�����
    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '�u���b�N�Ǘ��̃��R�[�h��
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim BlockIdBuf As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp2"

    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_SUCCESS
    
    sql = "select distinct "
    sql = sql & " B.CRYNUM, "
    sql = sql & " B.INGOTPOS as ss, "
'    sql = sql & " B.LENGTH, "             ' �����ǉ� 2001/11/8
    sql = sql & " B.BLOCKID, "
    sql = sql & " B.UPDDATE, "
    sql = sql & " B.HOLDCLS, "
    sql = sql & " H.HINBAN, "            ' �i��
    sql = sql & " H.REVNUM, "            ' ���i�ԍ������ԍ�
    sql = sql & " H.FACTORY, "           ' �H��
    sql = sql & " H.OPECOND, "           ' ���Ə���
    sql = sql & " S.HSXTYPE, "           ' �i�r�w�^�C�v
    sql = sql & " S.HSXCDIR, "            ' �i�r�w�����ʕ���
    sql = sql & " H.INGOTPOS, "
                '����NG�����邩�ǂ���
    sql = sql & " (select count(*) from VECME010 V1 "
    sql = sql & "  where V1.E040BLOCKID=B.BLOCKID "
    sql = sql & "  and ((V1.E043CRYINDRS<>'0' and V1.E043CRYRESRS='2') "         ' �����������сiRs)
    sql = sql & "  or (V1.E043CRYINDOI<>'0' and V1.E043CRYRESOI='2') "         ' �����������сiOi)
    sql = sql & "  or (V1.E043CRYINDB1<>'0' and V1.E043CRYRESB1='2')"          ' �����������сiB1)
    sql = sql & "  or (V1.E043CRYINDB2<>'0' and V1.E043CRYRESB2='2') "         ' �����������сiB2�j
    sql = sql & "  or (V1.E043CRYINDB3<>'0' and V1.E043CRYRESB3='2') "         ' �����������сiB3)
    sql = sql & "  or (V1.E043CRYINDL1<>'0' and V1.E043CRYRESL1='2') "         ' �����������сiL1)
    sql = sql & "  or (V1.E043CRYINDL2<>'0' and V1.E043CRYRESL2='2') "         ' �����������сiL2)
    sql = sql & "  or (V1.E043CRYINDL3<>'0' and V1.E043CRYRESL3='2') "         ' �����������сiL3)
    sql = sql & "  or (V1.E043CRYINDL4<>'0' and V1.E043CRYRESL4='2') "         ' �����������сiL4)
    sql = sql & "  or (V1.E043CRYINDCS<>'0' and V1.E043CRYRESCS='2') "         ' �����������сiCs)
    sql = sql & "  or (V1.E043CRYINDGD<>'0' and V1.E043CRYRESGD='2') "         ' �����������сiGD)
    sql = sql & "  or (V1.E043CRYINDT<>'0' and V1.E043CRYREST='2') "           ' �����������сiT)
    sql = sql & "  or (V1.E043CRYINDEP<>'0' and V1.E043CRYRESEP='2')) ) as J "         ' �����������сiEPD)
    sql = sql & " from "
    sql = sql & " TBCME040 B, TBCME041 H, TBCME018 S"
    sql = sql & " where "
    sql = sql & " B.CRYNUM = H.CRYNUM "
    sql = sql & " and H.HINBAN = S.HINBAN "
    sql = sql & " and H.REVNUM = S.MNOREVNO "
    sql = sql & " and H.FACTORY = S.FACTORY "
    sql = sql & " and H.OPECOND = S.OPECOND "
    
                '�H���R�[�h�A��ԁA�敪�̏����w��
    sql = sql & " and B.NOWPROC='CC600' "
    sql = sql & " and B.LSTATCLS='T' "
    sql = sql & " and B.RSTATCLS='T' "
    sql = sql & " and B.DELCLS='0' "
    'sql = sql & " and B.HOLDCLS='0' " ' �z�[���h�u���b�N���擾
                '�u���b�N���Ɋ܂܂��i�Ԃ�����
    sql = sql & " and (( B.INGOTPOS >= H.INGOTPOS "
    sql = sql & " and B.INGOTPOS < H.INGOTPOS + H.LENGTH ) "
    sql = sql & " or ( B.INGOTPOS + B.LENGTH > H.INGOTPOS "
    sql = sql & " and B.INGOTPOS + B.LENGTH < H.INGOTPOS + H.LENGTH  ) "
    sql = sql & " or ( H.INGOTPOS >= B.INGOTPOS "
    sql = sql & " and H.INGOTPOS < B.INGOTPOS + B.LENGTH ) "
    sql = sql & " or ( H.INGOTPOS + H.LENGTH > B.INGOTPOS "
    sql = sql & " and H.INGOTPOS + H.LENGTH < B.INGOTPOS + B.LENGTH )) "
                '�w����0�łȂ����т�0�łȂ��T���v�����㉺�Q�����邩
    sql = sql & " and 2=( select count(*) "
    sql = sql & "  from VECME010 V2 "
    sql = sql & "  where "
    sql = sql & "  B.BLOCKID=V2.E040BLOCKID"
    sql = sql & "  and (V2.E043CRYINDRS='0' or V2.E043CRYRESRS<>'0') "         ' �����������сiRs)
    sql = sql & "  and (V2.E043CRYINDOI='0' or V2.E043CRYRESOI<>'0') "         ' �����������сiOi)
    sql = sql & "  and (V2.E043CRYINDB1='0' or V2.E043CRYRESB1<>'0')"          ' �����������сiB1)
    sql = sql & "  and (V2.E043CRYINDB2='0' or V2.E043CRYRESB2<>'0') "         ' �����������сiB2�j
    sql = sql & "  and (V2.E043CRYINDB3='0' or V2.E043CRYRESB3<>'0') "         ' �����������сiB3)
    sql = sql & "  and (V2.E043CRYINDL1='0' or V2.E043CRYRESL1<>'0') "         ' �����������сiL1)
    sql = sql & "  and (V2.E043CRYINDL2='0' or V2.E043CRYRESL2<>'0') "         ' �����������сiL2)
    sql = sql & "  and (V2.E043CRYINDL3='0' or V2.E043CRYRESL3<>'0') "         ' �����������сiL3)
    sql = sql & "  and (V2.E043CRYINDL4='0' or V2.E043CRYRESL4<>'0') "         ' �����������сiL4)
    sql = sql & "  and (V2.E043CRYINDCS='0' or V2.E043CRYRESCS<>'0') "         ' �����������сiCs)
    sql = sql & "  and (V2.E043CRYINDGD='0' or V2.E043CRYRESGD<>'0') "         ' �����������сiGD)
    sql = sql & "  and (V2.E043CRYINDT='0' or V2.E043CRYREST<>'0') "           ' �����������сiT)
    sql = sql & "  and (V2.E043CRYINDEP='0' or V2.E043CRYRESEP<>'0') )"         ' �����������сiEPD)
    sql = sql & " order by B.BLOCKID, H.INGOTPOS "
    
    '�f�[�^�𒊏o����
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '���R�[�h0�����͐���
    If rs.RecordCount = 0 Then
        rs.Close
        ReDim records(0)
    Else
        BlockIdBuf = vbNullString
        recCnt = rs.RecordCount
        j = 0
        For i = 1 To recCnt
            DoEvents
        '�u���b�NID���̊i�[
            If rs("BLOCKID") <> BlockIdBuf Then
            
                j = j + 1
                ReDim Preserve records(j)
                
                With records(j)
                    .CRYNUM = rs("CRYNUM")
                    .IngotPos = rs("ss")
'                    .LENGTH = rs("LENGTH")      ' ����
                    .BLOCKID = rs("BLOCKID")   ' �u���b�NID
                    .UPDDATE = rs("UPDDATE")   ' �X�V���t
                    .HOLDCLS = rs("HOLDCLS")   ' �z�[���h�敪
                    BlockIdBuf = records(j).BLOCKID
                    .HSXTYPE = rs("HSXTYPE")
                    .HSXCDIR = rs("HSXCDIR")
                    If rs("J") > 0 Then
                        
                        .Judg = "2"
                    Else
                        .Judg = "1"
                    End If
                
                End With
                k = 1
            End If
            
            '�i�Ԃ̊i�[
            ReDim Preserve records(j).hin(k)
            records(j).hin(k).HINBAN = rs("HINBAN")
            records(j).hin(k).mnorevno = rs("REVNUM")
            records(j).hin(k).factory = rs("FACTORY")
            records(j).hin(k).opecond = rs("OPECOND")
            k = k + 1
            rs.MoveNext
        Next i
        rs.Close
            
    End If

    
    '�w���P�������ю擾
    If getKouBlock(records(), "CC600") = FUNCTION_RETURN_FAILURE Then
       DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
       GoTo proc_exit
    End If
    
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



'�T�v    :�҂��ꗗ �����\���p�c�a�h���C�o�i���o�҂��j
'���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
'        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'����    :
'����    :2001/07/06 ���{ �쐬
Public Function DBDRV_scmzc_fcmkc001b_Disp3(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN

    '�����o�҂���
    'CC700�̂���
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp3"


    DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_SUCCESS
    
    '�u���b�NID��X�V���t�A�i�ԓ��擾
    If getBlockID(records(), "CC700") = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


'    '�w���P�������ю擾
'    If getKouBlock(records(), "CC700") = FUNCTION_RETURN_FAILURE Then
'       DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
'       GoTo proc_exit
'    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function



'�T�v    :�҂��ꗗ �����\���p�c�a�h���C�o�i�����w���҂��j
'���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
'        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'����    :
'����    :2001/07/06 ���{ �쐬
Public Function DBDRV_scmzc_fcmkc001b_Disp4(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN

    '�������w���҂���
    'CC710�̂���
    
    '�u���b�NID��X�V���t�擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp4"

    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_SUCCESS


    '�u���b�NID��X�V���t�A�i�ԓ��擾
    If getBlockID(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
'2000/08/24 S.Sano Start
'    '�w���P�������ю擾
'    If getKouBlock(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
'       DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
'       GoTo proc_exit
'    End If
'2000/08/24 S.Sano End


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function
Public Function cmkc001b_DBDataCheck1(LWD() As cmkc001b_LockWait, Wd1() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'    Dim typ_A As typ_AllTypes        '�S���\����
'    Dim c0 As Integer
'    Dim sErrMsg As String
'    Dim NothingFlag As Boolean
'    Dim FuncAns As FUNCTION_RETURN
'    For c0 = 1 To UBound(Wd1())
'        NothingFlag = False
'        FuncAns = DBDRV_scmzc_fcmkc001b_Disp(Wd1(c0).BLOCKID, typ_A.typ_si, typ_A.typ_cr, typ_A.typ_zi, sErrMsg, NothingFlag)
'        LWD(c0).flag = NothingFlag
'    Next
    
   
    Dim l As Long, m As Long
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function cmkc001b_DBDataCheck1"

    
    cmkc001b_DBDataCheck1 = FUNCTION_RETURN_SUCCESS
    
    Set rs = Nothing
    
#If SPEEDUP Then   '���������� 02.1.28-2.15 �쑺
'����������
'���ƂȂ�u���b�N�Ƃ��̗��[�T���v���ɂ��āA������Ԃ��܂Ƃ߂Ď擾
'SQL�̔��s�񐔂�}�����ă��������ł̏����ɐ؂芷����
Dim SMP() As tSmpMng
Dim idx As Integer
Dim topIdx As Integer
Dim botIdx As Integer

Debug.Print " 1:" & Time
    sql = vbNullString
'    sql = sql & "select"
'    sql = sql & "  B.BLOCKID, B.INGOTPOS as TOPPOS, B.INGOTPOS+LENGTH as BOTPOS"
'    sql = sql & ", S.CRYNUM, S.INGOTPOS, SMPKBN, HINBAN, REVNUM, FACTORY, OPECOND"
'    sql = sql & ", CRYINDRS, CRYRESRS, CRYINDOI, CRYRESOI"
'    sql = sql & ", CRYINDB1, CRYRESB1, CRYINDB2, CRYRESB2, CRYINDB3, CRYRESB3"
'    sql = sql & ", CRYINDL1, CRYRESL1, CRYINDL2, CRYRESL2, CRYINDL3, CRYRESL3, CRYINDL4, CRYRESL4"
'    sql = sql & ", CRYINDCS, CRYRESCS, CRYINDGD, CRYRESGD, CRYINDT, CRYREST, CRYINDEP, CRYRESEP "
'    sql = sql & "from TBCME043 S, TBCME040 B "
'    sql = sql & "where S.CRYNUM=B.CRYNUM"
'    sql = sql & "  and B.INGOTPOS>=0"
'    sql = sql & "  and B.DELCLS='0'"
'    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
'    sql = sql & "  and B.RSTATCLS='T'"
'    sql = sql & "  and B.HOLDCLS='0'"
'    sql = sql & "  and ((S.INGOTPOS=B.INGOTPOS) or (S.INGOTPOS=B.INGOTPOS+B.LENGTH)) "
'    sql = sql & "order by B.BLOCKID, S.INGOTPOS, S.SMPKBN"
    sql = sql & "select"
    sql = sql & "  B.BLOCKID, B.INGOTPOS as TOPPOS, B.INGOTPOS+LENGTH as BOTPOS"
    sql = sql & ", S.XTALCS, S.INPOSCS, SMPKBNCS, HINBCS, REVNUMCS, FACTORYCS, OPECS"
    sql = sql & ", CRYINDRSCS, CRYRESRS1CS, CRYINDOICS, CRYRESOICS"
    sql = sql & ", CRYINDB1CS, CRYRESB1CS, CRYINDB2CS, CRYRESB2CS, CRYINDB3CS, CRYRESB3CS"
    sql = sql & ", CRYINDL1CS, CRYRESL1CS, CRYINDL2CS, CRYRESL2CS, CRYINDL3CS, CRYRESL3CS, CRYINDL4CS, CRYRESL4CS"
    sql = sql & ", CRYINDCSCS, CRYRESCSCS, CRYINDGDCS, CRYRESGDCS, CRYINDTCS, CRYRESTCS, CRYINDEPCS, CRYRESEPCS "
    sql = sql & "from XSDCS S, TBCME040 B "
    sql = sql & "where S.XTALCS=B.CRYNUM"
    sql = sql & "  and B.INGOTPOS>=0"
    sql = sql & "  and B.DELCLS='0'"
    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
    sql = sql & "  and B.RSTATCLS='T'"
    sql = sql & "  and B.HOLDCLS='0'"
    sql = sql & "  and ((S.INPOSCS=B.INGOTPOS) or (S.INPOSCS=B.INGOTPOS+B.LENGTH)) "
    sql = sql & "order by B.BLOCKID, S.INPOSCS, S.SMPKBNCS"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    ReDim SMP(rs.RecordCount)
    With SMP(0)
        .BLOCKID = " "
        .CRYNUM = " "
        .SMPKBN = " "
        .HINBAN = " "
        .factory = " "
        .opecond = " "
        .CRYINDRS = " "
        .CRYRESRS = " "
        .CRYINDOI = " "
        .CRYRESOI = " "
        .CRYINDB1 = " "
        .CRYRESB1 = " "
        .CRYINDB2 = " "
        .CRYRESB2 = " "
        .CRYINDB3 = " "
        .CRYRESB3 = " "
        .CRYINDL1 = " "
        .CRYRESL1 = " "
        .CRYINDL2 = " "
        .CRYRESL2 = " "
        .CRYINDL3 = " "
        .CRYRESL3 = " "
        .CRYINDL4 = " "
        .CRYRESL4 = " "
        .CRYINDCS = " "
        .CRYRESCS = " "
        .CRYINDGD = " "
        .CRYRESGD = " "
        .CRYINDT = " "
        .CRYREST = " "
        .CRYINDEP = " "
        .CRYRESEP = " "
    End With

    For l = 1 To rs.RecordCount
        With SMP(l)
            .BLOCKID = rs("BLOCKID")
            .TOPPOS = rs("TOPPOS")
            .BOTPOS = rs("BOTPOS")
            .CRYNUM = rs("CRYNUM")
            .IngotPos = rs("INGOTPOS")
            .SMPKBN = rs("SMPKBN")
            .HINBAN = rs("HINBAN")
            .REVNUM = rs("REVNUM")
            .factory = rs("FACTORY")
            .opecond = rs("OPECOND")
            .CRYINDRS = rs("CRYINDRS")
            .CRYRESRS = rs("CRYRESRS")
            .CRYINDOI = rs("CRYINDOI")
            .CRYRESOI = rs("CRYRESOI")
            .CRYINDB1 = rs("CRYINDB1")
            .CRYRESB1 = rs("CRYRESB1")
            .CRYINDB2 = rs("CRYINDB2")
            .CRYRESB2 = rs("CRYRESB2")
            .CRYINDB3 = rs("CRYINDB3")
            .CRYRESB3 = rs("CRYRESB3")
            .CRYINDL1 = rs("CRYINDL1")
            .CRYRESL1 = rs("CRYRESL1")
            .CRYINDL2 = rs("CRYINDL2")
            .CRYRESL2 = rs("CRYRESL2")
            .CRYINDL3 = rs("CRYINDL3")
            .CRYRESL3 = rs("CRYRESL3")
            .CRYINDL4 = rs("CRYINDL4")
            .CRYRESL4 = rs("CRYRESL4")
            .CRYINDCS = rs("CRYINDCS")
            .CRYRESCS = rs("CRYRESCS")
            .CRYINDGD = rs("CRYINDGD")
            .CRYRESGD = rs("CRYRESGD")
            .CRYINDT = rs("CRYINDT")
            .CRYREST = rs("CRYREST")
            .CRYINDEP = rs("CRYINDEP")
            .CRYRESEP = rs("CRYRESEP")
        End With
        rs.MoveNext
    Next
    rs.Close
    Set rs = Nothing
Debug.Print " 2:" & Time
#End If
    
    For l = 1 To UBound(Wd1())
        DoEvents
        LWD(l).flag = False
'Debug.Print " " & l & ":" & Time
        
        With Wd1(l)
        
        ' �w���P�����̃u���b�N�͖������łn�j
        If Mid$(.BLOCKID, 1, 1) <> "8" Then
        
            ReDim .SMP(2)
                        
            ' �㉺�̃T���v�����擾
#If SPEEDUP Then   '���������� 02.1.28-2.15 �쑺
'����������
'�ꊇ�擾����������Ԕz�񂩂�A�f�[�^���擾����悤�ɉ���
            For m = 1 To 2
                DoEvents
                
                topIdx = 0
                botIdx = 0
                For idx = 1 To UBound(SMP)
                    If (SMP(idx).BLOCKID = .BLOCKID) Then
                        If (SMP(idx).SMPKBN = "T") Then
                            topIdx = idx
                        Else
                            botIdx = idx
                        End If
                    ElseIf SMP(idx).BLOCKID > .BLOCKID Then
                        Exit For
                    End If
                Next
                If m = 1 Then
                    If topIdx > 0 Then
                        idx = topIdx
                    Else
                        idx = botIdx
                    End If
                Else
                    If botIdx > 0 Then
                        idx = botIdx
                    Else
                        idx = topIdx
                    End If
                End If
                
                With .SMP(m)
                    .CRYNUM = SMP(idx).CRYNUM
                    .IngotPos = SMP(idx).IngotPos
                    .SMPKBN = SMP(idx).SMPKBN
                    .HINBAN = SMP(idx).HINBAN
                    .REVNUM = SMP(idx).REVNUM
                    .factory = SMP(idx).factory
                    .opecond = SMP(idx).opecond
                    .CRYINDRS = SMP(idx).CRYINDRS
                    .CRYRESRS = SMP(idx).CRYRESRS
                    .CRYINDOI = SMP(idx).CRYINDOI
                    .CRYRESOI = SMP(idx).CRYRESOI
                    .CRYINDB1 = SMP(idx).CRYINDB1
                    .CRYRESB1 = SMP(idx).CRYRESB1
                    .CRYINDB2 = SMP(idx).CRYINDB2
                    .CRYRESB2 = SMP(idx).CRYRESB2
                    .CRYINDB3 = SMP(idx).CRYINDB3
                    .CRYRESB3 = SMP(idx).CRYRESB3
                    .CRYINDL1 = SMP(idx).CRYINDL1
                    .CRYRESL1 = SMP(idx).CRYRESL1
                    .CRYINDL2 = SMP(idx).CRYINDL2
                    .CRYRESL2 = SMP(idx).CRYRESL2
                    .CRYINDL3 = SMP(idx).CRYINDL3
                    .CRYRESL3 = SMP(idx).CRYRESL3
                    .CRYINDL4 = SMP(idx).CRYINDL4
                    .CRYRESL4 = SMP(idx).CRYRESL4
                    .CRYINDCS = SMP(idx).CRYINDCS
                    .CRYRESCS = SMP(idx).CRYRESCS
                    .CRYINDGD = SMP(idx).CRYINDGD
                    .CRYRESGD = SMP(idx).CRYRESGD
                    .CRYINDT = SMP(idx).CRYINDT
                    .CRYREST = SMP(idx).CRYREST
                    .CRYINDEP = SMP(idx).CRYINDEP
                    .CRYRESEP = SMP(idx).CRYRESEP
                End With
            Next m
            
#Else
            sql = " select "
            sql = sql & " V.E043CRYNUM, "
            sql = sql & " V.E043INGOTPOS, "
            sql = sql & " V.E043SMPKBN, "
            sql = sql & " V.E043HINBAN, "
            sql = sql & " V.E043REVNUM, "
            sql = sql & " V.E043FACTORY, "
            sql = sql & " V.E043OPECOND, "
            sql = sql & " V.E043CRYINDRS, "
            sql = sql & " V.E043CRYRESRS, "
            sql = sql & " V.E043CRYINDOI, "
            sql = sql & " V.E043CRYRESOI, "
            sql = sql & " V.E043CRYINDB1, "
            sql = sql & " V.E043CRYRESB1, "
            sql = sql & " V.E043CRYINDB2, "
            sql = sql & " V.E043CRYRESB2, "
            sql = sql & " V.E043CRYINDB3, "
            sql = sql & " V.E043CRYRESB3, "
            sql = sql & " V.E043CRYINDL1, "
            sql = sql & " V.E043CRYRESL1, "
            sql = sql & " V.E043CRYINDL2, "
            sql = sql & " V.E043CRYRESL2, "
            sql = sql & " V.E043CRYINDL3, "
            sql = sql & " V.E043CRYRESL3, "
            sql = sql & " V.E043CRYINDL4, "
            sql = sql & " V.E043CRYRESL4, "
            sql = sql & " V.E043CRYINDCS, "
            sql = sql & " V.E043CRYRESCS, "
            sql = sql & " V.E043CRYINDGD, "
            sql = sql & " V.E043CRYRESGD, "
            sql = sql & " V.E043CRYINDT, "
            sql = sql & " V.E043CRYREST, "
            sql = sql & " V.E043CRYINDEP, "
            sql = sql & " V.E043CRYRESEP "
            sql = sql & " from VECME010 V "
            sql = sql & " where E040CRYNUM = '" & .CRYNUM & "' "
            sql = sql & " and   E040INGOTPOS = '" & .IngotPos & "' "
            sql = sql & " order by E043INGOTPOS"
            
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            
            For m = 1 To 2
                DoEvents
                .SMP(m).CRYNUM = rs("E043CRYNUM")
                .SMP(m).IngotPos = rs("E043INGOTPOS")
                .SMP(m).SMPKBN = rs("E043SMPKBN")
                .SMP(m).HINBAN = rs("E043HINBAN")
                .SMP(m).REVNUM = rs("E043REVNUM")
                .SMP(m).factory = rs("E043FACTORY")
                .SMP(m).opecond = rs("E043OPECOND")
                .SMP(m).CRYINDRS = rs("E043CRYINDRS")
                .SMP(m).CRYRESRS = rs("E043CRYRESRS")
                .SMP(m).CRYINDOI = rs("E043CRYINDOI")
                .SMP(m).CRYRESOI = rs("E043CRYRESOI")
                .SMP(m).CRYINDB1 = rs("E043CRYINDB1")
                .SMP(m).CRYRESB1 = rs("E043CRYRESB1")
                .SMP(m).CRYINDB2 = rs("E043CRYINDB2")
                .SMP(m).CRYRESB2 = rs("E043CRYRESB2")
                .SMP(m).CRYINDB3 = rs("E043CRYINDB3")
                .SMP(m).CRYRESB3 = rs("E043CRYRESB3")
                .SMP(m).CRYINDL1 = rs("E043CRYINDL1")
                .SMP(m).CRYRESL1 = rs("E043CRYRESL1")
                .SMP(m).CRYINDL2 = rs("E043CRYINDL2")
                .SMP(m).CRYRESL2 = rs("E043CRYRESL2")
                .SMP(m).CRYINDL3 = rs("E043CRYINDL3")
                .SMP(m).CRYRESL3 = rs("E043CRYRESL3")
                .SMP(m).CRYINDL4 = rs("E043CRYINDL4")
                .SMP(m).CRYRESL4 = rs("E043CRYRESL4")
                .SMP(m).CRYINDCS = rs("E043CRYINDCS")
                .SMP(m).CRYRESCS = rs("E043CRYRESCS")
                .SMP(m).CRYINDGD = rs("E043CRYINDGD")
                .SMP(m).CRYRESGD = rs("E043CRYRESGD")
                .SMP(m).CRYINDT = rs("E043CRYINDT")
                .SMP(m).CRYREST = rs("E043CRYREST")
                .SMP(m).CRYINDEP = rs("E043CRYINDEP")
                .SMP(m).CRYRESEP = rs("E043CRYRESEP")
                
                rs.MoveNext
            Next m
            rs.Close
            Set rs = Nothing
#End If
            
'����������
'�i�Ԏd�l/Cs/EPD/LT�͂܂��u���b�N����SQL�𓊂��Ă���
'�������܂Ƃ߂Ă����΁A����5�b���x�k�ނ̂ł͂Ȃ����Ǝv����
'�������ACs/LT�ɂ��Ă͌��ʎ擾�̕��@���ς��̂ŁA���̌�̌������K�v
'������ɂ���A�Ώی����S�Ăɂ���Cs/LT/EPD�w���̂���T���v���𔲂��o���΂悢�͂�
            
            ' �i�Ԃ̎d�l���擾
            For m = 1 To 2
                If Trim$(.SMP(m).HINBAN) = "G" Or Trim$(.SMP(m).HINBAN) = "Z" Then
                    .SMP(m).HSXCNHWS = "S"
                    .SMP(m).HSXLTHWS = "S"
                    .SMP(m).EPD = "S"
                ElseIf Len(Trim$(.SMP(m).HINBAN)) Then
                    sql = " select "
                    sql = sql & " S.HSXCNHWS, "
                    sql = sql & " S.HSXLTHWS, "
                    sql = sql & " 'H' as EPD "
                    sql = sql & " from TBCME019 S "
                    sql = sql & " where S.HINBAN = '" & .SMP(m).HINBAN & "' "
                    sql = sql & " and S.MNOREVNO = " & .SMP(m).REVNUM & " "
                    sql = sql & " and S.FACTORY = '" & .SMP(m).factory & "' "
                    sql = sql & " and S.OPECOND = '" & .SMP(m).opecond & "' "
        
                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                    .SMP(m).HSXCNHWS = rs("HSXCNHWS")
                    .SMP(m).HSXLTHWS = rs("HSXLTHWS")
                    .SMP(m).EPD = rs("EPD")
                    
                    rs.Close
                    Set rs = Nothing
                Else
                    '��i�Ԃ̏ꍇ
                    .SMP(m).HSXCNHWS = " "
                    .SMP(m).HSXLTHWS = " "
                    .SMP(m).EPD = " "
                End If
            Next m
        
            ' �`�F�b�N
            For m = 1 To 2
                DoEvents
                ' CS�̃`�F�b�N
'                If (.SMP(m).HSXCNHWS = "H" Or .SMP(m).HSXCNHWS = "S") And .SMP(m).CRYINDCS = "0" Then  ' �Q�l�]���͂Ȃ��Ă��n�j
                If .SMP(m).HSXCNHWS = "H" And .SMP(m).CRYINDCS = "0" Then
                
                    sql = "select CRYRESCSCS as RES from XSDCS "
                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
                    sql = sql & "  and CRYINDCSCS<>'0'"
                    sql = sql & " order by INPOSCS"
                    
                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                    If rs.RecordCount Then
                        If rs("RES") = "0" Then LWD(l).flag = True
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                End If
                
                ' LT�̃`�F�b�N
'                If (.SMP(m).HSXLTHWS = "H" Or .SMP(m).HSXLTHWS = "S") And .SMP(m).CRYINDT = "0" And LWD(l).flag = False Then ' �Q�l�]���͂Ȃ��Ă��n�j
                If .SMP(m).HSXLTHWS = "H" And .SMP(m).CRYINDT = "0" And LWD(l).flag = False Then
                    
                    sql = "select CRYRESTCS as RES from XSDCS "
                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
                    sql = sql & "  and CRYINDTCS<>'0'"
                    sql = sql & " order by INPOSCS"
                    
                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                    If rs.RecordCount Then
                        If rs("RES") = "0" Then LWD(l).flag = True
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                End If
                
                ' EPD�̃`�F�b�N
'                If (.SMP(m).EPD = "H" Or .SMP(m).EPD = "S") And .SMP(m).CRYINDEP = "0" And LWD(l).flag = False Then ' S�͂��肦�Ȃ��Ǔ���
                If .SMP(m).EPD = "H" And .SMP(m).CRYINDEP = "0" And LWD(l).flag = False Then ' S�͂��肦�Ȃ��Ǔ���
                   
                    sql = "select CRYRESEPCS as RES from XSDCS "
                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
                    sql = sql & "  and CRYINDEP<>'0'"
                    sql = sql & " order by INPOSCS"
                    
                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                    If rs.RecordCount Then
                        If rs("RES") = "0" Then LWD(l).flag = True
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                End If
'                If LWD(l).flag = True Then
'                    Exit For
'                End If
            Next m
        End If
        
        End With    ' .Wd1()
        
    Next l
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    cmkc001b_DBDataCheck1 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

Public Function cmkc001b_DBDataCheck3(LWD() As cmkc001b_LockWait, _
                                 Wd3() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
    Dim c0 As Integer
    Dim c1 As Integer
    Dim c2 As Integer
    Dim MaxRec As Integer
    Dim recCount As Integer
    Dim EQFlag As Boolean
    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim GrpCount1 As Integer
    Dim GrpCount2 As Integer
    Dim ColorFlag As Boolean
    Dim TotalBlk As Integer
    Dim CheckPoint As Integer
    Dim CheckEnd As Integer
    Dim tempGrpFlag As String * 1
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp"

    cmkc001b_DBDataCheck3 = FUNCTION_RETURN_SUCCESS
    TotalBlk = UBound(Wd3())
    
Debug.Print " 1:" & Time
    
    'CC700�̃u���b�N�̌����ꗗ�����
    ReDim GrpInfo(1) As cmkc001b_Wait3
    GrpInfo(1).CRYNUM = vbNullString
    c1 = 0
    For c0 = 1 To TotalBlk
        DoEvents
        If c1 = 0 Then
            GrpInfo(1).CRYNUM = Wd3(c0).CRYNUM
        End If
        MaxRec = UBound(GrpInfo())
        EQFlag = False
        c1 = 1
        Do While c1 <= MaxRec
            DoEvents
            If Wd3(c0).CRYNUM = GrpInfo(c1).CRYNUM Then
                EQFlag = True
                Exit Do
            End If
            c1 = c1 + 1
        Loop
        If Not EQFlag Then
            ReDim Preserve GrpInfo(MaxRec + 1) As cmkc001b_Wait3
            GrpInfo(MaxRec + 1).CRYNUM = Wd3(c0).CRYNUM
        End If
    Next
Debug.Print " 2:" & Time
        
    '�����Ɋ܂܂��S�Ẵu���b�N�����߂�
    MaxRec = UBound(GrpInfo())
    For c0 = 1 To MaxRec
        sql = "select "
        sql = sql & "BLOCKID, "
        sql = sql & "INGOTPOS, "
        sql = sql & "LENGTH, "
        sql = sql & "NOWPROC, "
        sql = sql & "HOLDCLS "
        sql = sql & "from TBCME040 "
        sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
'2001/11/14 S.Sano        sql = sql & "and LSTATCLS='T' "
'2001/11/14 S.Sano        sql = sql & "and RSTATCLS='T' "
'2001/11/14 S.Sano        sql = sql & "and DELCLS='0' "
        'sql = sql & "and HOLDCLS='0' "
        sql = sql & "order by BLOCKID "
    
        
        '�f�[�^�𒊏o����
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        recCount = rs.RecordCount
        If recCount = 0 Then
            rs.Close
            GoTo proc_exit
        End If
        ReDim GrpInfo(c0).blkInfo(recCount) As cmkc001b_Wait3_BLK
        For c1 = 1 To recCount
            GrpInfo(c0).blkInfo(c1).BLOCKID = rs("BLOCKID")
            GrpInfo(c0).blkInfo(c1).IngotPos = rs("INGOTPOS")
            GrpInfo(c0).blkInfo(c1).LENGTH = rs("LENGTH")
            GrpInfo(c0).blkInfo(c1).NOWPROC = rs("NOWPROC")
            GrpInfo(c0).blkInfo(c1).HOLDCLS = rs("HOLDCLS")
            rs.MoveNext
        Next
        rs.Close
    Next

Debug.Print " 3:" & Time
    '�u���b�N�̏㉺�i�Ԃ����߂�
#If SPEEDUP Then   '���������� 02.1.28-2.15 �쑺
'����������
'�u���b�N�̏㉺�i�Ԃ����߂邾���Ȃ�A1���SQL�ł܂Ƃ߂ď����擾�ł���͂�
Dim blkID() As String
Dim topHin() As tFullHinban
Dim botHin() As tFullHinban
Dim idx As Integer
Dim rsCount As Integer
Dim found As Boolean

    sql = vbNullString
    sql = sql & "select"
    sql = sql & "  b.BLOCKID"
    sql = sql & ", TOP.HINBAN as THINBAN, TOP.REVNUM as TREVNUM, TOP.FACTORY as TFACTORY, TOP.OPECOND as TOPECOND"
    sql = sql & ", BOT.HINBAN as BHINBAN, BOT.REVNUM as BREVNUM, BOT.FACTORY as BFACTORY, BOT.OPECOND as BOPECOND "
    sql = sql & "from TBCME040 B, TBCME041 TOP, TBCME041 BOT "
    sql = sql & "Where b.CRYNUM = Top.CRYNUM"
    sql = sql & "  and B.CRYNUM=BOT.CRYNUM"
    sql = sql & "  and B.INGOTPOS>=0"
    sql = sql & "  and B.DELCLS='0'"
    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
    sql = sql & "  and B.RSTATCLS='T'"
    sql = sql & "  and B.HOLDCLS='0'"
    sql = sql & "  and B.INGOTPOS>=TOP.INGOTPOS"
    sql = sql & "  and B.INGOTPOS<TOP.INGOTPOS+TOP.LENGTH"
    sql = sql & "  and B.INGOTPOS+B.LENGTH>BOT.INGOTPOS"
    sql = sql & "  and B.INGOTPOS+B.LENGTH<=BOT.INGOTPOS+BOT.LENGTH "
    sql = sql & "order by B.BLOCKID"
    
    '�f�[�^�𒊏o����
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    rsCount = rs.RecordCount
    ReDim blkID(1 To rsCount)
    ReDim topHin(1 To rsCount)
    ReDim botHin(1 To rsCount)
    For c0 = 1 To rsCount
        blkID(c0) = rs!BLOCKID
        topHin(c0).HINBAN = rs!THINBAN
        topHin(c0).mnorevno = rs!TREVNUM
        topHin(c0).factory = rs!TFACTORY
        topHin(c0).opecond = rs!TOPECOND
        botHin(c0).HINBAN = rs!BHINBAN
        botHin(c0).mnorevno = rs!BREVNUM
        botHin(c0).factory = rs!BFACTORY
        botHin(c0).opecond = rs!BOPECOND
        rs.MoveNext
    Next
    rs.Close

    For c0 = 1 To MaxRec
        recCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To recCount
            found = False
            For idx = 1 To rsCount
                If blkID(idx) = GrpInfo(c0).blkInfo(c1).BLOCKID Then
                    found = True
                    Exit For
                ElseIf blkID(idx) > GrpInfo(c0).blkInfo(c1).BLOCKID Then
                    Exit For
                End If
            Next
        
            If found Then
                GrpInfo(c0).blkInfo(c1).topHin.HINBAN = topHin(idx).HINBAN
                GrpInfo(c0).blkInfo(c1).topHin.factory = topHin(idx).factory
                GrpInfo(c0).blkInfo(c1).topHin.opecond = topHin(idx).opecond
                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = topHin(idx).mnorevno
            Else
                GrpInfo(c0).blkInfo(c1).topHin.HINBAN = ""
                GrpInfo(c0).blkInfo(c1).topHin.factory = ""
                GrpInfo(c0).blkInfo(c1).topHin.opecond = ""
                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = 0
            End If
            
            If found Then
                GrpInfo(c0).blkInfo(c1).botHin.HINBAN = botHin(idx).HINBAN
                GrpInfo(c0).blkInfo(c1).botHin.factory = botHin(idx).factory
                GrpInfo(c0).blkInfo(c1).botHin.opecond = botHin(idx).opecond
                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = botHin(idx).mnorevno
            Else
                GrpInfo(c0).blkInfo(c1).botHin.HINBAN = ""
                GrpInfo(c0).blkInfo(c1).botHin.factory = ""
                GrpInfo(c0).blkInfo(c1).botHin.opecond = ""
                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = 0
            End If
        Next
    Next
#Else
    For c0 = 1 To MaxRec
        recCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To recCount
            sql = "select "
            sql = sql & "HINBAN, "
            sql = sql & "REVNUM, "
            sql = sql & "FACTORY, "
            sql = sql & "OPECOND "
            sql = sql & "from TBCME041 "
            sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
'2001/11/14 S.Sano            sql = sql & "and INGOTPOS <= " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
            sql = sql & "and INGOTPOS = " & GrpInfo(c0).blkInfo(c1).IngotPos & " " '2001/11/14 S.Sano
'2001/11/14 S.Sano            sql = sql & "and (INGOTPOS + LENGTH) > " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
            
            '�f�[�^�𒊏o����
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            recCount = rs.RecordCount
            If recCount = 0 Then
                GrpInfo(c0).blkInfo(c1).topHin.HINBAN = ""
                GrpInfo(c0).blkInfo(c1).topHin.factory = ""
                GrpInfo(c0).blkInfo(c1).topHin.opecond = ""
                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = 0
            Else
                GrpInfo(c0).blkInfo(c1).topHin.HINBAN = rs("HINBAN")
                GrpInfo(c0).blkInfo(c1).topHin.factory = rs("FACTORY")
                GrpInfo(c0).blkInfo(c1).topHin.opecond = rs("OPECOND")
                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = rs("REVNUM")
            End If
            rs.Close
        
            sql = "select "
            sql = sql & "HINBAN, "
            sql = sql & "REVNUM, "
            sql = sql & "FACTORY, "
            sql = sql & "OPECOND "
            sql = sql & "from TBCME041 "
            sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
            sql = sql & "and INGOTPOS < " & GrpInfo(c0).blkInfo(c1).IngotPos + GrpInfo(c0).blkInfo(c1).LENGTH & " "
            sql = sql & "and (INGOTPOS + LENGTH) >= " & GrpInfo(c0).blkInfo(c1).IngotPos + GrpInfo(c0).blkInfo(c1).LENGTH & " "
            
            '�f�[�^�𒊏o����
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            recCount = rs.RecordCount
            If recCount = 0 Then
                GrpInfo(c0).blkInfo(c1).botHin.HINBAN = ""
                GrpInfo(c0).blkInfo(c1).botHin.factory = ""
                GrpInfo(c0).blkInfo(c1).botHin.opecond = ""
                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = 0
            Else
                GrpInfo(c0).blkInfo(c1).botHin.HINBAN = rs("HINBAN")
                GrpInfo(c0).blkInfo(c1).botHin.factory = rs("FACTORY")
                GrpInfo(c0).blkInfo(c1).botHin.opecond = rs("OPECOND")
                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = rs("REVNUM")
            End If
            rs.Close
        Next
    Next
#End If
    
Debug.Print " 4:" & Time
    '���߂���񂩂�O���[�v�����߂�
    GrpCount1 = 0
    GrpCount2 = 0
    For c0 = 1 To MaxRec
        GrpCount1 = GrpCount1 + 1
        GrpCount2 = GrpCount2 + 1
        recCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To recCount
            '�u���b�N�؂�ڂŕi�Ԃ��ς��ΕʃO���[�v�Ɣ��f����
            Select Case c1
            Case 1
                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
            Case Else
                If (GrpInfo(c0).blkInfo(c1).topHin.factory <> GrpInfo(c0).blkInfo(c1 - 1).botHin.factory) Or _
                   (GrpInfo(c0).blkInfo(c1).topHin.HINBAN <> GrpInfo(c0).blkInfo(c1 - 1).botHin.HINBAN) Or _
                   (GrpInfo(c0).blkInfo(c1).topHin.opecond <> GrpInfo(c0).blkInfo(c1 - 1).botHin.opecond) Or _
                   (GrpInfo(c0).blkInfo(c1).topHin.REVNUM <> GrpInfo(c0).blkInfo(c1 - 1).botHin.REVNUM) Then
                    GrpCount1 = GrpCount1 + 1
                End If
                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
            End Select
            
            '����O���[�v���ŁA�H���Ⴂ�̃u���b�N�����݂����ꍇ�A����O���[�v����
            '���O���[�v�Ƃ��ăO���[�v��������B
            'CC710�ȊO�Ȃ�ΏۊO�Ƃ��O���[�v��������Ȃ�
            If GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_NUKISI_SIJI And GrpInfo(c0).blkInfo(c1).HOLDCLS = "0" Then
                Select Case c1
                Case 1
                    GrpInfo(c0).blkInfo(c1).GRPFLG2 = GrpCount2
                Case Else
                    If (GrpInfo(c0).blkInfo(c1).topHin.factory <> GrpInfo(c0).blkInfo(c1 - 1).botHin.factory) Or _
                       (GrpInfo(c0).blkInfo(c1).topHin.HINBAN <> GrpInfo(c0).blkInfo(c1 - 1).botHin.HINBAN) Or _
                       (GrpInfo(c0).blkInfo(c1).topHin.opecond <> GrpInfo(c0).blkInfo(c1 - 1).botHin.opecond) Or _
                       (GrpInfo(c0).blkInfo(c1).topHin.REVNUM <> GrpInfo(c0).blkInfo(c1 - 1).botHin.REVNUM) Then
                        GrpCount2 = GrpCount2 + 1
                    End If
                    GrpInfo(c0).blkInfo(c1).GRPFLG2 = GrpCount2
                End Select
            Else
                GrpCount2 = GrpCount2 + 1
                GrpInfo(c0).blkInfo(c1).GRPFLG2 = 0
            End If
        Next
    Next
Debug.Print " 5:" & Time
    '���߂���񂩂�\���F�����߂�
    For c0 = 1 To MaxRec
        recCount = UBound(GrpInfo(c0).blkInfo())
        ColorFlag = False
        CheckPoint = 0
        For c1 = 1 To recCount
            If CheckPoint > 0 Then
                If GrpInfo(c0).blkInfo(c1).GRPFLG1 <> GrpInfo(c0).blkInfo(CheckPoint).GRPFLG1 Then
                    For c2 = CheckPoint To c1 - 1
                        GrpInfo(c0).blkInfo(c2).COLORFLG = ColorFlag
                    Next
                    ColorFlag = False
                    CheckPoint = c1
                End If
            Else
                CheckPoint = c1
            End If
            If CheckPoint > 0 Then
                If (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_SETUDAN) Or _
                   (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_KESSYOU_SOUGOUHANTEI) Or _
                   (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_KESSYOU_SAISYUU_HARAIDASI) Or _
                   (GrpInfo(c0).blkInfo(c1).HOLDCLS = "1") Then
                    ColorFlag = True
                End If
            End If
        Next
        For c1 = CheckPoint To recCount
            GrpInfo(c0).blkInfo(c1).COLORFLG = ColorFlag
        Next
    Next
Debug.Print " 6:" & Time
    For c0 = 1 To MaxRec
        recCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To recCount
            For c2 = 1 To TotalBlk
                If Wd3(c2).BLOCKID = GrpInfo(c0).blkInfo(c1).BLOCKID Then
                    LWD(c2).flag = GrpInfo(c0).blkInfo(c1).COLORFLG
                    LWD(c2).Grp = GrpInfo(c0).blkInfo(c1).GRPFLG2
                    Exit For
                End If
            Next
        Next
    Next
'    Debug.Print Now

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    cmkc001b_DBDataCheck3 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'�w���P�����p
Private Function getKouBlock(records() As type_DBDRV_scmzc_fcmkc001b_Disp, NOWPROC As String) As FUNCTION_RETURN

    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long
    Dim motoRecCnt As Long
    Dim i As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function getKouBlock"

    getKouBlock = FUNCTION_RETURN_SUCCESS

    sql = " select "
    sql = sql & " B.BLOCKID, "
    sql = sql & " B.UPDDATE, "
    sql = sql & " B.HOLDCLS, "
    sql = sql & " K.HINBAN, "
    sql = sql & " K.MNOREVNO, "
    sql = sql & " K.FACTORY, "
    sql = sql & " K.OPECOND "
    sql = sql & " from TBCME040 B,TBCMG002 K "
    sql = sql & " where B.BLOCKID=K.CRYNUM "
    sql = sql & " and substr(B.BLOCKID,1,1)='8' "
    sql = sql & " and B.NOWPROC='" & NOWPROC & "' "
    sql = sql & " and B.LSTATCLS='T' "
    sql = sql & " and B.RSTATCLS='T' "
    sql = sql & " and B.DELCLS='0' "
    'sql = sql & " and B.HOLDCLS='0' " ' �z�[���h�u���b�N���擾
    sql = sql & " and K.TRANCNT=any(select max(TRANCNT) from TBCMG002 where CRYNUM=B.BLOCKID ) "
    sql = sql & " order by B.BLOCKID "

    
    '�f�[�^�𒊏o����
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        rs.Close
        GoTo proc_exit
    End If
    
    motoRecCnt = UBound(records)
    recCnt = rs.RecordCount
    ReDim Preserve records(UBound(records) + recCnt)
    
    For i = motoRecCnt + 1 To UBound(records)
        DoEvents
        ReDim records(i).hin(1)
        With records(i)
            .BLOCKID = rs("BLOCKID")     ' �u���b�NID
            .UPDDATE = rs("UPDDATE")     ' �X�V���t
            .HOLDCLS = rs("HOLDCLS")     ' �z�[���h�敪
            .hin(1).HINBAN = rs("HINBAN")       ' �i��
            .hin(1).mnorevno = rs("MNOREVNO")   ' ���i�ԍ������ԍ�
            .hin(1).factory = rs("FACTORY")     ' �H��
            .hin(1).opecond = rs("OPECOND")     ' ���Ə���
        End With
        rs.MoveNext
    Next
    rs.Close
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getKouBlock = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
    
End Function

'�����֐� �u���b�NID�A�X�V���t�擾�i���o�҂��A�����w���҂��p�j
Private Function getBlockID(records() As type_DBDRV_scmzc_fcmkc001b_Disp, _
                            NOWPROC As String) As FUNCTION_RETURN

    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim BlockIdBuf As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function getBlockID"

    getBlockID = FUNCTION_RETURN_SUCCESS

    sql = "select "
    sql = sql & " V.E040CRYNUM, "
    sql = sql & " V.E040BLOCKID, "
    sql = sql & " V.E040INGOTPOS, "
    sql = sql & " V.E040UPDDATE, "
    sql = sql & " V.E040HOLDCLS, "
    sql = sql & " V.E041HINBAN, "            ' �i��
    sql = sql & " V.E041REVNUM, "            ' ���i�ԍ������ԍ�
    sql = sql & " V.E041FACTORY, "           ' �H��
    sql = sql & " V.E041OPECOND, "           ' ���Ə���
    sql = sql & " S.HSXTYPE, "           ' �i�r�w�^�C�v
    sql = sql & " S.HSXCDIR "            ' �i�r�w�����ʕ���
    sql = sql & " from "
    sql = sql & " VECME009 V, TBCME018 S "
    sql = sql & " where "
    sql = sql & " V.E041HINBAN = S.HINBAN "
    sql = sql & " and V.E041REVNUM = S.MNOREVNO "
    sql = sql & " and V.E041FACTORY = S.FACTORY "
    sql = sql & " and V.E041OPECOND = S.OPECOND "
    sql = sql & " and V.E040NOWPROC='" & NOWPROC & "' "
    sql = sql & " and V.E040LSTATCLS='T' "
    sql = sql & " and V.E040RSTATCLS='T' "
    sql = sql & " and V.E040DELCLS='0' "
    'sql = sql & " and V.E040HOLDCLS='0' " ' �z�[���h�u���b�N���擾
    sql = sql & " order by V.E040BLOCKID, V.E041INGOTPOS "

    '�f�[�^�𒊏o����
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '���R�[�h���Ȃ��ꍇ����I��
    If rs.RecordCount = 0 Then
        rs.Close
        ReDim records(0)
        GoTo proc_exit
    End If
    
    BlockIdBuf = vbNullString
    recCnt = rs.RecordCount
    j = 0
    For i = 1 To recCnt
        DoEvents
        '�u���b�NID���̊i�[
        If rs("E040BLOCKID") <> BlockIdBuf Then
        
            j = j + 1
            ReDim Preserve records(j)
            
            With records(j)
                .CRYNUM = rs("E040CRYNUM")
                .IngotPos = rs("E040INGOTPOS")
                .BLOCKID = rs("E040BLOCKID")   ' �u���b�NID
                .UPDDATE = rs("E040UPDDATE")   ' �X�V���t
                .HOLDCLS = rs("E040HOLDCLS")   ' �z�[���h�敪
                BlockIdBuf = records(j).BLOCKID
                .HSXTYPE = rs("HSXTYPE")
                .HSXCDIR = rs("HSXCDIR")
                .Judg = " "
            End With
            
            k = 1
        End If
        
        '�i�Ԃ̊i�[
        ReDim Preserve records(j).hin(k)
        records(j).hin(k).HINBAN = rs("E041HINBAN")
        records(j).hin(k).mnorevno = rs("E041REVNUM")
        records(j).hin(k).factory = rs("E041FACTORY")
        records(j).hin(k).opecond = rs("E041OPECOND")
        k = k + 1
        rs.MoveNext
    Next i
    rs.Close
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getBlockID = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


