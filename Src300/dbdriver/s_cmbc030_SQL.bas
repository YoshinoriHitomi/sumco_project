Attribute VB_Name = "s_cmbc030_SQL"
Option Explicit

Type cmkc001b_LockWait
    flag        As Boolean
    Grp         As Integer
End Type

Type cmkc001b_Wait3_HINBAN
    hinban      As String * 8               ' �i��
    REVNUM      As Integer                  ' ���i�ԍ������ԍ�
    factory     As String * 1               ' �H��
    opecond     As String * 1               ' ���Ə���
End Type

Type cmkc001b_Wait3_BLK
    BLOCKID     As String * 12              ' �u���b�NID
    IngotPos    As Integer                  ' �������J�n�ʒu
    LENGTH      As Integer                  ' ����
    NOWPROC     As String * 5               ' ���ݍH��
    HOLDCLS     As String * 1               ' �z�[���h�敪 ---kuramoto �ǉ� 2001/09/19----
    GRPFLG1     As Integer                  ' �O���[�v���
    GRPFLG2     As Integer                  ' �O���[�v���
    COLORFLG    As Boolean
    topHin      As cmkc001b_Wait3_HINBAN
    botHin      As cmkc001b_Wait3_HINBAN
End Type

Type cmkc001b_Wait3
    CRYNUM      As String * 12              ' �����ԍ�
    blkInfo()   As cmkc001b_Wait3_BLK
End Type

Type type_cmkc001b_SmpMng
    CRYNUM      As String * 12
    IngotPos    As Integer
    SMPKBN      As String * 1
    
    hinban      As String * 8               ' �i��
    REVNUM      As Integer                  ' ���i�ԍ������ԍ�
    factory     As String * 1               ' �H��
    opecond     As String * 1               ' ���Ə���
        
    CRYINDRS    As String * 1
    CRYRESRS    As String * 1
    CRYINDOI    As String * 1
    CRYRESOI    As String * 1
    CRYINDB1    As String * 1
    CRYRESB1    As String * 1
    CRYINDB2    As String * 1
    CRYRESB2    As String * 1
    CRYINDB3    As String * 1
    CRYRESB3    As String * 1
    CRYINDL1    As String * 1
    CRYRESL1    As String * 1
    CRYINDL2    As String * 1
    CRYRESL2    As String * 1
    CRYINDL3    As String * 1
    CRYRESL3    As String * 1
    CRYINDL4    As String * 1
    CRYRESL4    As String * 1
    CRYINDCS    As String * 1
    CRYRESCS    As String * 1
    CRYINDGD    As String * 1
    CRYRESGD    As String * 1
    CRYINDT     As String * 1
    CRYREST     As String * 1
    CRYINDEP    As String * 1
    CRYRESEP    As String * 1
    
    HSXCNHWS    As String * 1               ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    HSXLTHWS    As String * 1               ' �i�r�w�k�^�C���ۏؕ��@�Q��
    EPD         As String * 1               ' EPD
End Type

'''''#If SPEEDUP Then   '���������� 02.1.28-2.15 �쑺
'''''Private Type tSmpMng
'''''    BLOCKID     As String * 12
'''''    TOPPOS      As Integer
'''''    BOTPOS      As Integer
'''''
'''''    CRYNUM      As String * 12
'''''    IngotPos    As Integer
'''''    SMPKBN      As String * 1
'''''
'''''    hinban      As String * 8           ' �i��
'''''    REVNUM      As Integer              ' ���i�ԍ������ԍ�
'''''    factory     As String * 1           ' �H��
'''''    opecond     As String * 1           ' ���Ə���
'''''
'''''    CRYINDRS    As String * 1
'''''    CRYRESRS    As String * 1
'''''    CRYINDOI    As String * 1
'''''    CRYRESOI    As String * 1
'''''    CRYINDB1    As String * 1
'''''    CRYRESB1    As String * 1
'''''    CRYINDB2    As String * 1
'''''    CRYRESB2    As String * 1
'''''    CRYINDB3    As String * 1
'''''    CRYRESB3    As String * 1
'''''    CRYINDL1    As String * 1
'''''    CRYRESL1    As String * 1
'''''    CRYINDL2    As String * 1
'''''    CRYRESL2    As String * 1
'''''    CRYINDL3    As String * 1
'''''    CRYRESL3    As String * 1
'''''    CRYINDL4    As String * 1
'''''    CRYRESL4    As String * 1
'''''    CRYINDCS    As String * 1
'''''    CRYRESCS    As String * 1
'''''    CRYINDGD    As String * 1
'''''    CRYRESGD    As String * 1
'''''    CRYINDT     As String * 1
'''''    CRYREST     As String * 1
'''''    CRYINDEP    As String * 1
'''''    CRYRESEP    As String * 1
'''''End Type
'''''#End If


'�҂��ꗗ

'�����\���p
Public Type type_DBDRV_scmzc_fcmkc001b_Disp
    CRYNUM      As String * 12              ' �����ԍ�
    IngotPos    As Integer                  ' �������J�n�ʒu
'   LENGTH      As Integer                  ' ����              '2001/11/8
    BLOCKID     As String * 12              ' �u���b�NID
    HSXTYPE     As String * 1               ' �i�r�w�^�C�v
    HSXCDIR     As String * 1               ' �i�r�w�����ʕ���
    UPDDATE     As Date                     ' �X�V���t
    Judg        As String                   ' ����
    hin()       As tFullHinban              ' �i��(full)
    HOLDCLS     As String * 1               ' �z�[���h�敪 ---kuramoto �ǉ� 2001/09/25----
    SMP()       As type_cmkc001b_SmpMng     ' �T���v���Ǘ�
End Type


''''''�i�ԁA�d�l�A���������擾�p (TOP,TAIL���łQ���R�[�h�擾)
'''''Public Type type_DBDRV_scmzc_fcmkc001c_Siyou
'''''    '�u���b�N�Ǘ�
'''''    CRYNUM As String * 12             ' �����ԍ�
'''''    INGOTPOS As Integer               ' �������J�n�ʒu
'''''    LENGTH As Integer                 ' ����
'''''    '�i�ԊǗ�
'''''    hin As tFullHinban                ' �i��(full)
'''''
'''''        '�������
'''''    PRODCOND As String * 4            ' �������
'''''    PGID As String * 8                ' �o�f�|�h�c
'''''    UPLENGTH As Integer               ' ���グ����
'''''    FREELENG As Integer               ' �t���[��
'''''    DIAMETER As Integer               ' ���a 2002/05/01 S.Sano
'''''    CHARGE As Double                  ' �`���[�W��
'''''    SEED As String * 4                ' �V�[�h
'''''    ADDDPPOS As Integer                 ' �ǉ��h�[�v�ʒu
'''''
'''''    '���i�d�l
'''''    HSXTYPE As String * 1             ' �i�r�w�^�C�v
'''''    HSXD1CEN As Double                ' �i�r�w���a�P���S
'''''    HSXCDIR As String * 1             ' �i�r�w�����ʕ���
'''''    HSXRMIN As Double                 ' �i�r�w���R����
'''''    HSXRMAX As Double                 ' �i�r�w���R���
'''''    HSXRAMIN As Double                ' �i�r�w���R���ω���
'''''    HSXRAMAX As Double                ' �i�r�w���R���Ϗ��
'''''    HSXRMCAL As String * 1            ' �i�r�w���R�ʓ��v�Z�@�@�@�@'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
'''''    HSXRMBNP As Double                ' �i�r�w���R�ʓ����z
'''''    HSXRSPOH As String * 1            ' �i�r�w���R����ʒu�Q��
'''''    HSXRSPOT As String * 1            ' �i�r�w���R����ʒu�Q�_
'''''    HSXRSPOI As String * 1            ' �i�r�w���R����ʒu�Q��
'''''    HSXRHWYT As String * 1            ' �i�r�w���R�ۏؕ��@�Q��
'''''    HSXRHWYS As String * 1            ' �i�r�w���R�ۏؕ��@�Q��
'''''
'''''    HSXONMIN As Double                ' �i�r�w�_�f�Z�x����
'''''    HSXONMAX As Double                ' �i�r�w�_�f�Z�x���
'''''    HSXONAMN As Double                ' �i�r�w�_�f�Z�x���ω���
'''''    HSXONAMX As Double                ' �i�r�w�_�f�Z�x���Ϗ��
'''''    HSXONMCL As String * 1            ' �i�r�w�_�f�Z�x�ʓ��v�Z�@�@�@�@'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
'''''    HSXONMBP As Double                ' �i�r�w�_�f�Z�x�ʓ����z
'''''    HSXONSPH As String * 1            ' �i�r�w�_�f�Z�x����ʒu�Q��
'''''    HSXONSPT As String * 1            ' �i�r�w�_�f�Z�x����ʒu�Q�_
'''''    HSXONSPI As String * 1            ' �i�r�w�_�f�Z�x����ʒu�Q��
'''''    HSXONHWT As String * 1            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
'''''    HSXONHWS As String * 1            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
'''''
'''''    HSXBM1AN As Double                ' �i�r�w�a�l�c�P���ω���
'''''    HSXBM1AX As Double                ' �i�r�w�a�l�c�P���Ϗ��
'''''    HSXBM2AN As Double                ' �i�r�w�a�l�c�Q���ω���
'''''    HSXBM2AX As Double                ' �i�r�w�a�l�c�Q���Ϗ��
'''''    HSXBM3AN As Double                ' �i�r�w�a�l�c�R���ω���
'''''    HSXBM3AX As Double                ' �i�r�w�a�l�c�R���Ϗ��
'''''    HSXBM1SH As String * 1            ' �i�r�w�a�l�c�P����ʒu�Q��
'''''    HSXBM1ST As String * 1            ' �i�r�w�a�l�c�P����ʒu�Q�_
'''''    HSXBM1SR As String * 1            ' �i�r�w�a�l�c�P����ʒu�Q��
'''''    HSXBM1HT As String * 1            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
'''''    HSXBM1HS As String * 1            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
'''''    HSXBM2SH As String * 1            ' �i�r�w�a�l�c�Q����ʒu�Q��
'''''    HSXBM2ST As String * 1            ' �i�r�w�a�l�c�Q����ʒu�Q�_
'''''    HSXBM2SR As String * 1            ' �i�r�w�a�l�c�Q����ʒu�Q��
'''''    HSXBM2HT As String * 1            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
'''''    HSXBM2HS As String * 1            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
'''''    HSXBM3SH As String * 1            ' �i�r�w�a�l�c�R����ʒu�Q��
'''''    HSXBM3ST As String * 1            ' �i�r�w�a�l�c�R����ʒu�Q�_
'''''    HSXBM3SR As String * 1            ' �i�r�w�a�l�c�R����ʒu�Q��
'''''    HSXBM3HT As String * 1            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
'''''    HSXBM3HS As String * 1            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
'''''
'''''    HSXOS1AX As Double                ' �i�r�w�n�r�e�P���Ϗ��
'''''    HSXOS1MX As Double                ' �i�r�w�n�r�e�P���
'''''    HSXOS2AX As Double                ' �i�r�w�n�r�e�Q���Ϗ��
'''''    HSXOS2MX As Double                ' �i�r�w�n�r�e�Q���
'''''    HSXOS3AX As Double                ' �i�r�w�n�r�e�R���Ϗ��
'''''    HSXOS3MX As Double                ' �i�r�w�n�r�e�R���
'''''    HSXOS4AX As Double                ' �i�r�w�n�r�e�S���Ϗ��
'''''    HSXOS4MX As Double                ' �i�r�w�n�r�e�S���
'''''    HSXOS1SH As String * 1            ' �i�r�w�n�r�e�P����ʒu�Q��
'''''    HSXOS1ST As String * 1            ' �i�r�w�n�r�e�P����ʒu�Q�_
'''''    HSXOS1SR As String * 1            ' �i�r�w�n�r�e�P����ʒu�Q��
'''''    HSXOS1HT As String * 1            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
'''''    HSXOS1HS As String * 1            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
'''''    HSXOS2SH As String * 1            ' �i�r�w�n�r�e�Q����ʒu�Q��
'''''    HSXOS2ST As String * 1            ' �i�r�w�n�r�e�Q����ʒu�Q�_
'''''    HSXOS2SR As String * 1            ' �i�r�w�n�r�e�Q����ʒu�Q��
'''''    HSXOS2HT As String * 1            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
'''''    HSXOS2HS As String * 1            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
'''''    HSXOS3SH As String * 1            ' �i�r�w�n�r�e�R����ʒu�Q��
'''''    HSXOS3ST As String * 1            ' �i�r�w�n�r�e�R����ʒu�Q�_
'''''    HSXOS3SR As String * 1            ' �i�r�w�n�r�e�R����ʒu�Q��
'''''    HSXOS3HT As String * 1            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
'''''    HSXOS3HS As String * 1            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
'''''    HSXOS4SH As String * 1            ' �i�r�w�n�r�e�S����ʒu�Q��
'''''    HSXOS4ST As String * 1            ' �i�r�w�n�r�e�S����ʒu�Q�_
'''''    HSXOS4SR As String * 1            ' �i�r�w�n�r�e�S����ʒu�Q��
'''''    HSXOS4HT As String * 1            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
'''''    HSXOS4HS As String * 1            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
'''''    HSXOS1NS As String * 2            ' �i�r�w�n�r�e�P�M�����@
'''''    HSXOS2NS As String * 2            ' �i�r�w�n�r�e�Q�M�����@
'''''    HSXOS3NS As String * 2            ' �i�r�w�n�r�e�R�M�����@
'''''    HSXOS4NS As String * 2            ' �i�r�w�n�r�e�S�M�����@
'''''    HSXBM1NS As String * 2            ' �i�r�w�a�l�c�P�M�����@
'''''    HSXBM2NS As String * 2            ' �i�r�w�a�l�c�Q�M�����@
'''''    HSXBM3NS As String * 2            ' �i�r�w�a�l�c�R�M�����@
'''''
'''''    HSXCNMIN As Double                ' �i�r�w�Y�f�Z�x����
'''''    HSXCNMAX As Double                ' �i�r�w�Y�f�Z�x���
'''''    HSXCNSPH As String * 1            ' �i�r�w�Y�f�Z�x����ʒu�Q��
'''''    HSXCNSPT As String * 1            ' �i�r�w�Y�f�Z�x����ʒu�Q�_
'''''    HSXCNSPI As String * 1            ' �i�r�w�Y�f�Z�x����ʒu�Q��
'''''    HSXCNHWT As String * 1            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
'''''    HSXCNHWS As String * 1            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
'''''
'''''    HSXDENMX As Integer               ' �i�r�w�c�������
'''''    HSXDENMN As Integer               ' �i�r�w�c��������
'''''    HSXLDLMX As Integer               ' �i�r�w�k�^�c�k���
'''''    HSXLDLMN As Integer               ' �i�r�w�k�^�c�k����
'''''    HSXDVDMX As Integer               ' �i�r�w�c�u�c�Q���
'''''    HSXDVDMN As Integer               ' �i�r�w�c�u�c�Q����
'''''    HSXDENHT As String * 1            ' �i�r�w�c�����ۏؕ��@�Q��
'''''    HSXDENHS As String * 1            ' �i�r�w�c�����ۏؕ��@�Q��
'''''    HSXLDLHT As String * 1            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
'''''    HSXLDLHS As String * 1            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
'''''    HSXDVDHT As String * 1            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
'''''    HSXDVDHS As String * 1            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
'''''    HSXDENKU As String * 1            ' �i�r�w�c���������L��
'''''    HSXDVDKU As String * 1            ' �i�r�w�c�u�c�Q�����L��
'''''    HSXLDLKU As String * 1            ' �i�r�w�k�^�c�k�����L��
'''''
'''''    HSXLTMIN As Integer               ' �i�r�w�k�^�C������
'''''    HSXLTMAX As Integer               ' �i�r�w�k�^�C�����
'''''    HSXLTSPH As String * 1            ' �i�r�w�k�^�C������ʒu�Q��
'''''    HSXLTSPT As String * 1            ' �i�r�w�k�^�C������ʒu�Q�_
'''''    HSXLTSPI As String * 1            ' �i�r�w�k�^�C������ʒu�Q��
'''''    HSXLTHWT As String * 1            ' �i�r�w�k�^�C���ۏؕ��@�Q��
'''''    HSXLTHWS As String * 1            ' �i�r�w�k�^�C���ۏؕ��@�Q��
'''''    '���������Ǘ�
'''''    EPDUP As Integer                  ' EPD�@���
'''''
'''''' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 start
'''''    TOPREG  As Integer                ' TOP�K��
'''''    TAILREG As Double                 ' TAIL�K��
'''''    BTMSPRT As Integer                ' �{�g���͏o�K��
'''''' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 end
'''''
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''    HSXOSF1PTK As String * 1          ' �i�r�w�n�r�e�P�p�^���敪
'''''    HSXOSF2PTK As String * 1          ' �i�r�w�n�r�e�Q�p�^���敪
'''''    HSXOSF3PTK As String * 1          ' �i�r�w�n�r�e�R�p�^���敪
'''''    HSXOSF4PTK As String * 1          ' �i�r�w�n�r�e�S�p�^���敪
'''''    HSXBMD1MBP As Double              ' �i�r�w�a�l�c�P�ʓ����z
'''''    HSXBMD2MBP As Double              ' �i�r�w�a�l�c�Q�ʓ����z
'''''    HSXBMD3MBP As Double              ' �i�r�w�a�l�c�R�ʓ����z
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''End Type


'''''' �����T���v���Ǘ��擾�p (TOP,TAIL���łQ���R�[�h�擾)
'''''Public Type type_DBDRV_scmzc_fcmkc001c_CrySmp
'''''    CRYNUM As String * 12             ' �����ԍ�
'''''    INGOTPOS As Integer               ' �������ʒu
'''''    LENGTH As Integer                 ' ����
'''''    BLOCKID As String * 12            ' �u���b�NID
'''''    SMPKBN As String * 1              ' �T���v���敪
'''''    SMPLNO As Integer                 ' �T���v��No
'''''    hinban As String * 12             ' �i��
'''''    REVNUM As Integer                 ' ���i�ԍ������ԍ�
'''''    factory As String * 1             ' �H��
'''''    opecond As String * 1             ' ���Ə���
'''''    KTKBN  As String * 1              ' �m��敪
'''''    CRYINDRS As String * 1            ' ���������w���iRs)
'''''    CRYINDOI As String * 1            ' ���������w���iOi)
'''''    CRYINDB1 As String * 1            ' ���������w���iB1)
'''''    CRYINDB2 As String * 1            ' ���������w���iB2�j
'''''    CRYINDB3 As String * 1            ' ���������w���iB3)
'''''    CRYINDL1 As String * 1            ' ���������w���iL1)
'''''    CRYINDL2 As String * 1            ' ���������w���iL2)
'''''    CRYINDL3 As String * 1            ' ���������w���iL3)
'''''    CRYINDL4 As String * 1            ' ���������w���iL4)
'''''    CRYINDCS As String * 1            ' ���������w���iCs)
'''''    CRYINDGD As String * 1            ' ���������w���iGD)
'''''    CRYINDT As String * 1             ' ���������w���iT)
'''''    CRYINDEP As String * 1            ' ���������w���iEPD)
'''''End Type


''''''������R����
'''''Public Type type_DBDRV_scmzc_fcmkc001c_CryR
'''''    CRYNUM As String * 12             ' �����ԍ�
'''''    POSITION As Integer               ' �ʒu
'''''    SMPKBN As String * 1              ' �T���v���敪
'''''    SMPLNO As Integer                 ' �T���v���m��
'''''    SMPLUMU As String * 1             ' �T���v���L��
'''''    TRANCOND As String * 1            ' ��������
'''''    MEAS1 As Double                   ' ����l�P
'''''    MEAS2 As Double                   ' ����l�Q
'''''    MEAS3 As Double                   ' ����l�R
'''''    MEAS4 As Double                   ' ����l�S
'''''    MEAS5 As Double                   ' ����l�T
'''''    RRG As Double                     ' �q�q�f
'''''    REGDATE As Date                   ' �o�^���t
'''''End Type


''''''Oi����
'''''Public Type type_DBDRV_scmzc_fcmkc001c_Oi
'''''    CRYNUM As String * 12             ' �����ԍ�
'''''    POSITION As Integer               ' �ʒu
'''''    SMPKBN As String * 1              ' �T���v���敪
'''''    SMPLNO As Integer                 ' �T���v���m��
'''''    SMPLUMU As String * 1             ' �T���v���L��
'''''    TRANCOND As String * 1            ' ��������
'''''    OIMEAS1 As Double                 ' �n������l�P
'''''    OIMEAS2 As Double                 ' �n������l�Q
'''''    OIMEAS3 As Double                 ' �n������l�R
'''''    OIMEAS4 As Double                 ' �n������l�S
'''''    OIMEAS5 As Double                 ' �n������l�T
'''''    ORGRES As Double                  ' �n�q�f����
'''''    AVE As Double                     ' �`�u�d
'''''    FTIRCONV As Double                ' �e�s�h�q���Z
'''''    INSPECTWAY As String * 2          ' �������@
'''''    REGDATE As Date                   ' �o�^���t
'''''End Type
'''''
'''''
''''''BMD1�`3����
'''''Public Type type_DBDRV_scmzc_fcmkc001c_BMD
'''''    CRYNUM As String * 12             ' �����ԍ�
'''''    POSITION As Integer               ' �ʒu
'''''    SMPKBN As String * 1              ' �T���v���敪
'''''    SMPLNO As Integer                 ' �T���v���m��
'''''    SMPLUMU As String * 1             ' �T���v���L��
'''''    HTPRC As String * 2               ' �M�������@
'''''    KKSP As String * 3                ' �������ב���ʒu
'''''    KKSET As String * 3               ' �������ב�������{�I��ET��@�@char(1)�{number(2)
'''''    TRANCOND As String * 1            ' ��������
'''''    MEAS1 As Double                   ' ����l�P
'''''    MEAS2 As Double                   ' ����l�Q
'''''    MEAS3 As Double                   ' ����l�R
'''''    MEAS4 As Double                   ' ����l�S
'''''    MEAS5 As Double                   ' ����l�T
'''''    Min As Double                     ' MIN
'''''    max As Double                     ' MAX
'''''    AVE As Double                     ' AVE
'''''    REGDATE As Date                   ' �o�^���t
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''    BMDMNBUNP As Double               ' BMD�ʓ����z
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''End Type
'''''
'''''
''''''OSF1�`4����
'''''Public Type type_DBDRV_scmzc_fcmkc001c_OSF
'''''    CRYNUM As String * 12             ' �����ԍ�
'''''    POSITION As Integer               ' �ʒu
'''''    SMPKBN As String * 1              ' �T���v���敪
'''''    SMPLNO As Integer                 ' �T���v���m��
'''''    SMPLUMU As String * 1             ' �T���v���L��
'''''    HTPRC As String * 2               ' �M�������@
'''''    KKSP As String * 3                ' �������ב���ʒu
'''''    KKSET As String * 3               ' �������ב�������{�I��ET��@�@char(1)�{number(2)
'''''    TRANCOND As String * 1            ' ��������
'''''    CALCMAX As Double                 ' �v�Z���� Max
'''''    CALCAVE As Double                 ' �v�Z���� Ave
'''''    MEAS1 As Double                   ' ����l�P
'''''    MEAS2 As Double                   ' ����l�Q
'''''    MEAS3 As Double                   ' ����l�R
'''''    MEAS4 As Double                   ' ����l�S
'''''    MEAS5 As Double                   ' ����l�T
'''''    MEAS6 As Double                   ' ����l�U
'''''    MEAS7 As Double                   ' ����l�V
'''''    MEAS8 As Double                   ' ����l�W
'''''    MEAS9 As Double                   ' ����l�X
'''''    MEAS10 As Double                  ' ����l�P�O
'''''    MEAS11 As Double                  ' ����l�P�P
'''''    MEAS12 As Double                  ' ����l�P�Q
'''''    MEAS13 As Double                  ' ����l�P�R
'''''    MEAS14 As Double                  ' ����l�P�S
'''''    MEAS15 As Double                  ' ����l�P�T
'''''    MEAS16 As Double                  ' ����l�P�U
'''''    MEAS17 As Double                  ' ����l�P�V
'''''    MEAS18 As Double                  ' ����l�P�W
'''''    MEAS19 As Double                  ' ����l�P�X
'''''    MEAS20 As Double                  ' ����l�Q�O
'''''    REGDATE As Date                   ' �o�^���t
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''    OSFPOS1    As Double              ' ����݋敪�P�ʒu
'''''    OSFWID1    As Double              ' ����݋敪�P��
'''''    OSFRD1     As String * 1          ' ����݋敪�PR/D
'''''    OSFPOS2    As Double              ' ����݋敪�Q�ʒu
'''''    OSFWID2    As Double              ' ����݋敪�Q��
'''''    OSFRD2     As String * 1          ' ����݋敪�QR/D
'''''    OSFPOS3    As Double              ' ����݋敪�R�ʒu
'''''    OSFWID3    As Double              ' ����݋敪�R��
'''''    OSFRD3     As String * 1          ' ����݋敪�RR/D
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''End Type
'''''
'''''
''''''CS����
'''''Public Type type_DBDRV_scmzc_fcmkc001c_CS
'''''    CRYNUM As String * 12             ' �����ԍ�
'''''    POSITION As Integer               ' �ʒu
'''''    SMPKBN As String * 1              ' �T���v���敪
'''''    SMPLNO As Integer                 ' �T���v���m��
'''''    SMPLUMU As String * 1             ' �T���v���L��
'''''    TRANCOND As String * 1            ' ��������
'''''    CSMEAS As Double                  ' Cs�����l
'''''    PRE70P As Double                  ' �V�O������l
'''''    REGDATE As Date                   ' �o�^���t
'''''End Type
'''''
'''''
''''''GD����
'''''Public Type type_DBDRV_scmzc_fcmkc001c_GD
'''''    CRYNUM As String * 12             ' �����ԍ�
'''''    POSITION As Integer               ' �ʒu
'''''    SMPKBN As String * 1              ' �T���v���敪
'''''    SMPLNO As Integer                 ' �T���v���m��
'''''    SMPLUMU As String * 1             ' �T���v���L��
'''''    TRANCOND As String * 1            ' ��������
'''''    MSRSDEN As Integer                ' ���茋�� Den
'''''    MSRSLDL As Integer                ' ���茋�� L/DL
'''''    MSRSDVD2 As Integer               ' ���茋�� DVD2
'''''    MS01LDL1 As Integer             ' ����l01 L/DL1
'''''    MS01LDL2 As Integer             ' ����l01 L/DL2
'''''    MS01LDL3 As Integer             ' ����l01 L/DL3
'''''    MS01LDL4 As Integer             ' ����l01 L/DL4
'''''    MS01LDL5 As Integer             ' ����l01 L/DL5
'''''    MS01DEN1 As Integer             ' ����l01 Den1
'''''    MS01DEN2 As Integer             ' ����l01 Den2
'''''    MS01DEN3 As Integer             ' ����l01 Den3
'''''    MS01DEN4 As Integer             ' ����l01 Den4
'''''    MS01DEN5 As Integer             ' ����l01 Den5
'''''    MS02LDL1 As Integer             ' ����l02 L/DL1
'''''    MS02LDL2 As Integer             ' ����l02 L/DL2
'''''    MS02LDL3 As Integer             ' ����l02 L/DL3
'''''    MS02LDL4 As Integer             ' ����l02 L/DL4
'''''    MS02LDL5 As Integer             ' ����l02 L/DL5
'''''    MS02DEN1 As Integer             ' ����l02 Den1
'''''    MS02DEN2 As Integer             ' ����l02 Den2
'''''    MS02DEN3 As Integer             ' ����l02 Den3
'''''    MS02DEN4 As Integer             ' ����l02 Den4
'''''    MS02DEN5 As Integer             ' ����l02 Den5
'''''    MS03LDL1 As Integer             ' ����l03 L/DL1
'''''    MS03LDL2 As Integer             ' ����l03 L/DL2
'''''    MS03LDL3 As Integer             ' ����l03 L/DL3
'''''    MS03LDL4 As Integer             ' ����l03 L/DL4
'''''    MS03LDL5 As Integer             ' ����l03 L/DL5
'''''    MS03DEN1 As Integer             ' ����l03 Den1
'''''    MS03DEN2 As Integer             ' ����l03 Den2
'''''    MS03DEN3 As Integer             ' ����l03 Den3
'''''    MS03DEN4 As Integer             ' ����l03 Den4
'''''    MS03DEN5 As Integer             ' ����l03 Den5
'''''    MS04LDL1 As Integer             ' ����l04 L/DL1
'''''    MS04LDL2 As Integer             ' ����l04 L/DL2
'''''    MS04LDL3 As Integer             ' ����l04 L/DL3
'''''    MS04LDL4 As Integer             ' ����l04 L/DL4
'''''    MS04LDL5 As Integer             ' ����l04 L/DL5
'''''    MS04DEN1 As Integer             ' ����l04 Den1
'''''    MS04DEN2 As Integer             ' ����l04 Den2
'''''    MS04DEN3 As Integer             ' ����l04 Den3
'''''    MS04DEN4 As Integer             ' ����l04 Den4
'''''    MS04DEN5 As Integer             ' ����l04 Den5
'''''    MS05LDL1 As Integer             ' ����l05 L/DL1
'''''    MS05LDL2 As Integer             ' ����l05 L/DL2
'''''    MS05LDL3 As Integer             ' ����l05 L/DL3
'''''    MS05LDL4 As Integer             ' ����l05 L/DL4
'''''    MS05LDL5 As Integer             ' ����l05 L/DL5
'''''    MS05DEN1 As Integer             ' ����l05 Den1
'''''    MS05DEN2 As Integer             ' ����l05 Den2
'''''    MS05DEN3 As Integer             ' ����l05 Den3
'''''    MS05DEN4 As Integer             ' ����l05 Den4
'''''    MS05DEN5 As Integer             ' ����l05 Den5
'''''    MS06LDL1 As Integer             ' ����l06 L/DL1
'''''    MS06LDL2 As Integer             ' ����l06 L/DL2
'''''    MS06LDL3 As Integer             ' ����l06 L/DL3
'''''    MS06LDL4 As Integer             ' ����l06 L/DL4
'''''    MS06LDL5 As Integer             ' ����l06 L/DL5
'''''    MS06DEN1 As Integer             ' ����l06 Den1
'''''    MS06DEN2 As Integer             ' ����l06 Den2
'''''    MS06DEN3 As Integer             ' ����l06 Den3
'''''    MS06DEN4 As Integer             ' ����l06 Den4
'''''    MS06DEN5 As Integer             ' ����l06 Den5
'''''    MS07LDL1 As Integer             ' ����l07 L/DL1
'''''    MS07LDL2 As Integer             ' ����l07 L/DL2
'''''    MS07LDL3 As Integer             ' ����l07 L/DL3
'''''    MS07LDL4 As Integer             ' ����l07 L/DL4
'''''    MS07LDL5 As Integer             ' ����l07 L/DL5
'''''    MS07DEN1 As Integer             ' ����l07 Den1
'''''    MS07DEN2 As Integer             ' ����l07 Den2
'''''    MS07DEN3 As Integer             ' ����l07 Den3
'''''    MS07DEN4 As Integer             ' ����l07 Den4
'''''    MS07DEN5 As Integer             ' ����l07 Den5
'''''    MS08LDL1 As Integer             ' ����l08 L/DL1
'''''    MS08LDL2 As Integer             ' ����l08 L/DL2
'''''    MS08LDL3 As Integer             ' ����l08 L/DL3
'''''    MS08LDL4 As Integer             ' ����l08 L/DL4
'''''    MS08LDL5 As Integer             ' ����l08 L/DL5
'''''    MS08DEN1 As Integer             ' ����l08 Den1
'''''    MS08DEN2 As Integer             ' ����l08 Den2
'''''    MS08DEN3 As Integer             ' ����l08 Den3
'''''    MS08DEN4 As Integer             ' ����l08 Den4
'''''    MS08DEN5 As Integer             ' ����l08 Den5
'''''    MS09LDL1 As Integer             ' ����l09 L/DL1
'''''    MS09LDL2 As Integer             ' ����l09 L/DL2
'''''    MS09LDL3 As Integer             ' ����l09 L/DL3
'''''    MS09LDL4 As Integer             ' ����l09 L/DL4
'''''    MS09LDL5 As Integer             ' ����l09 L/DL5
'''''    MS09DEN1 As Integer             ' ����l09 Den1
'''''    MS09DEN2 As Integer             ' ����l09 Den2
'''''    MS09DEN3 As Integer             ' ����l09 Den3
'''''    MS09DEN4 As Integer             ' ����l09 Den4
'''''    MS09DEN5 As Integer             ' ����l09 Den5
'''''    MS10LDL1 As Integer             ' ����l10 L/DL1
'''''    MS10LDL2 As Integer             ' ����l10 L/DL2
'''''    MS10LDL3 As Integer             ' ����l10 L/DL3
'''''    MS10LDL4 As Integer             ' ����l10 L/DL4
'''''    MS10LDL5 As Integer             ' ����l10 L/DL5
'''''    MS10DEN1 As Integer             ' ����l10 Den1
'''''    MS10DEN2 As Integer             ' ����l10 Den2
'''''    MS10DEN3 As Integer             ' ����l10 Den3
'''''    MS10DEN4 As Integer             ' ����l10 Den4
'''''    MS10DEN5 As Integer             ' ����l10 Den5
'''''    MS11LDL1 As Integer             ' ����l11 L/DL1
'''''    MS11LDL2 As Integer             ' ����l11 L/DL2
'''''    MS11LDL3 As Integer             ' ����l11 L/DL3
'''''    MS11LDL4 As Integer             ' ����l11 L/DL4
'''''    MS11LDL5 As Integer             ' ����l11 L/DL5
'''''    MS11DEN1 As Integer             ' ����l11 Den1
'''''    MS11DEN2 As Integer             ' ����l11 Den2
'''''    MS11DEN3 As Integer             ' ����l11 Den3
'''''    MS11DEN4 As Integer             ' ����l11 Den4
'''''    MS11DEN5 As Integer             ' ����l11 Den5
'''''    MS12LDL1 As Integer             ' ����l12 L/DL1
'''''    MS12LDL2 As Integer             ' ����l12 L/DL2
'''''    MS12LDL3 As Integer             ' ����l12 L/DL3
'''''    MS12LDL4 As Integer             ' ����l12 L/DL4
'''''    MS12LDL5 As Integer             ' ����l12 L/DL5
'''''    MS12DEN1 As Integer             ' ����l12 Den1
'''''    MS12DEN2 As Integer             ' ����l12 Den2
'''''    MS12DEN3 As Integer             ' ����l12 Den3
'''''    MS12DEN4 As Integer             ' ����l12 Den4
'''''    MS12DEN5 As Integer             ' ����l12 Den5
'''''    MS13LDL1 As Integer             ' ����l13 L/DL1
'''''    MS13LDL2 As Integer             ' ����l13 L/DL2
'''''    MS13LDL3 As Integer             ' ����l13 L/DL3
'''''    MS13LDL4 As Integer             ' ����l13 L/DL4
'''''    MS13LDL5 As Integer             ' ����l13 L/DL5
'''''    MS13DEN1 As Integer             ' ����l13 Den1
'''''    MS13DEN2 As Integer             ' ����l13 Den2
'''''    MS13DEN3 As Integer             ' ����l13 Den3
'''''    MS13DEN4 As Integer             ' ����l13 Den4
'''''    MS13DEN5 As Integer             ' ����l13 Den5
'''''    MS14LDL1 As Integer             ' ����l14 L/DL1
'''''    MS14LDL2 As Integer             ' ����l14 L/DL2
'''''    MS14LDL3 As Integer             ' ����l14 L/DL3
'''''    MS14LDL4 As Integer             ' ����l14 L/DL4
'''''    MS14LDL5 As Integer             ' ����l14 L/DL5
'''''    MS14DEN1 As Integer             ' ����l14 Den1
'''''    MS14DEN2 As Integer             ' ����l14 Den2
'''''    MS14DEN3 As Integer             ' ����l14 Den3
'''''    MS14DEN4 As Integer             ' ����l14 Den4
'''''    MS14DEN5 As Integer             ' ����l14 Den5
'''''    MS15LDL1 As Integer             ' ����l15 L/DL1
'''''    MS15LDL2 As Integer             ' ����l15 L/DL2
'''''    MS15LDL3 As Integer             ' ����l15 L/DL3
'''''    MS15LDL4 As Integer             ' ����l15 L/DL4
'''''    MS15LDL5 As Integer             ' ����l15 L/DL5
'''''    MS15DEN1 As Integer             ' ����l15 Den1
'''''    MS15DEN2 As Integer             ' ����l15 Den2
'''''    MS15DEN3 As Integer             ' ����l15 Den3
'''''    MS15DEN4 As Integer             ' ����l15 Den4
'''''    MS15DEN5 As Integer             ' ����l15 Den5
'''''    REGDATE As Date                   ' �o�^���t
'''''End Type
'''''
'''''
''''''���C�t�^�C�����ю擾�֐�
'''''Public Type type_DBDRV_scmzc_fcmkc001c_LT
'''''    CRYNUM As String * 12             ' �����ԍ�
'''''    POSITION As Integer               ' �ʒu
'''''    SMPKBN As String * 1              ' �T���v���敪
'''''    SMPLNO As Integer                 ' �T���v���m��
'''''    SMPLUMU As String * 1             ' �T���v���L��
'''''    MEAS1 As Integer                  ' ����l�P
'''''    MEAS2 As Integer                  ' ����l�Q
'''''    MEAS3 As Integer                  ' ����l�R
'''''    MEAS4 As Integer                  ' ����l�S
'''''    MEAS5 As Integer                  ' ����l�T
'''''    TRANCOND As String * 1            ' ��������
'''''    MEASPEAK As Integer               ' ����l �s�[�N�l
'''''    CALCMEAS As Integer               ' �v�Z����
'''''    REGDATE As Date                   ' �o�^���t
'''''    LTSPI As String                 '����ʒu�R�[�h
'''''End Type
'''''
'''''
''''''EPD���ю擾�֐�
'''''Public Type type_DBDRV_scmzc_fcmkc001c_EPD
'''''    CRYNUM As String * 12             ' �����ԍ�
'''''    POSITION As Integer               ' �ʒu
'''''    SMPKBN As String * 1              ' �T���v���敪
'''''    SMPLNO As Integer                 ' �T���v���m��
'''''    SMPLUMU As String * 1             ' �T���v���L��
'''''    TRANCOND As String * 1            ' ��������
'''''    MEASURE As Integer                ' ����l
'''''    REGDATE As Date                   ' �o�^���t
'''''End Type


''''''���т��܂Ƃ߂��\����
'''''Public Type type_DBDRV_scmzc_fcmkc001c_Zisseki
'''''    CRYRZ() As type_DBDRV_scmzc_fcmkc001c_CryR
'''''    OIZ() As type_DBDRV_scmzc_fcmkc001c_Oi
'''''    BMD1Z() As type_DBDRV_scmzc_fcmkc001c_BMD
'''''    BMD2Z() As type_DBDRV_scmzc_fcmkc001c_BMD
'''''    BMD3Z() As type_DBDRV_scmzc_fcmkc001c_BMD
'''''    OSF1Z() As type_DBDRV_scmzc_fcmkc001c_OSF
'''''    OSF2Z() As type_DBDRV_scmzc_fcmkc001c_OSF
'''''    OSF3Z() As type_DBDRV_scmzc_fcmkc001c_OSF
'''''    OSF4Z() As type_DBDRV_scmzc_fcmkc001c_OSF
'''''    csz() As type_DBDRV_scmzc_fcmkc001c_CS
'''''    GDZ() As type_DBDRV_scmzc_fcmkc001c_GD
'''''    LTZ() As type_DBDRV_scmzc_fcmkc001c_LT
'''''    EPDZ() As type_DBDRV_scmzc_fcmkc001c_EPD
'''''    SURSZ() As type_DBDRV_scmzc_fcmkc001c_CryR
'''''End Type


'�u���b�N�Ǘ��X�V�p�i���ݍH���A�ŏI�ʉߍH���j
Public Type type_DBDRV_scmzc_fcmkc001c_UpdBlock1
    CRYNUM      As String * 12          ' �����ԍ�
    IngotPos    As Integer              ' �������J�n�ʒu
    NOWPROC     As String * 5           ' ���ݍH��
    LASTPASS    As String * 5           ' �ŏI�ʉߍH��
End Type



'�u���b�N�Ǘ��X�V�p�i�N���X�^���J�^���O�A�������g�p�j
Public Type typ_DBDRV_fcmkc001c_UpdBlkCR
    CRYNUM      As String * 12          ' �����ԍ�
    IngotPos    As Integer              ' �������J�n�ʒu
    NOWPROC     As String * 5           ' ���ݍH��
'   LASTPASS    As String * 5           ' �ŏI�ʉߍH��
    DELCLS      As String * 1           ' �폜�敪
    BDCAUS      As String * 3           ' �s�Ǘ��R
    LSTATCLS    As String * 1           ' �ŏI��ԋ敪
    RSTATCLS    As String * 1           ' ������ԋ敪
End Type



'�����T���v���Ǘ��X�V�p
Public Type type_DBDRV_scmzc_fcmkc001c_UpdCrySmp
    CRYNUM      As String * 12          ' �����ԍ�
    IngotPos    As Integer              ' �������ʒu
    SMPKBN      As String * 1           ' �T���v���敪
End Type


''''''���茋�ʂ�J014�����v�ۍ\����
'''''Public Type Judg_Spec_Cry
'''''    Enable As Boolean           '�L���ȕi�Ԃł���
'''''    rs As Boolean               'Rs�͗v����
'''''    Oi As Boolean               'Oi�͗v����
'''''    B1 As Boolean               'BMD1�͗v����
'''''    B2 As Boolean               'BMD2�͗v����
'''''    B3 As Boolean               'BMD3�͗v����
'''''    L1 As Boolean               'OSF1�͗v����
'''''    L2 As Boolean               'OSF2�͗v����
'''''    L3 As Boolean               'OSF3�͗v����
'''''    L4 As Boolean               'OSF4�͗v����
'''''    Cs As Boolean               'Cs�͗v����
'''''    GD As Boolean               'GD�͗v����
'''''    Lt As Boolean               'LT�͗v����
'''''    EPD As Boolean              'EPD�͗v����
'''''End Type


'''''' �d�l�̎w���������Ă��锻�f�p
'''''Public Const SIJI = "H"
'''''Public Const SANKOU = "S"




''''''�����֐� �u���b�NID�A�X�V���t�擾�i���o�҂��A�����w���҂��p�j
'''''Private Function getBlockID(records() As type_DBDRV_scmzc_fcmkc001b_Disp, _
'''''                            NOWPROC As String) As FUNCTION_RETURN
'''''
'''''    Dim sql         As String       'SQL�S��
'''''    Dim rs          As OraDynaset   'RecordSet
'''''    Dim recCnt      As Long         '���R�[�h��
'''''    Dim i           As Long
'''''    Dim j           As Long
'''''    Dim k           As Long
'''''    Dim BlockIdBuf  As String
'''''
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function getBlockID"
'''''
'''''    getBlockID = FUNCTION_RETURN_SUCCESS
'''''
'''''    sql = "select V.E040CRYNUM, "
'''''    sql = sql & " V.E040BLOCKID, "
'''''    sql = sql & " V.E040INGOTPOS, "
'''''    sql = sql & " V.E040UPDDATE, "
'''''    sql = sql & " V.E040HOLDCLS, "
'''''    sql = sql & " V.E041HINBAN, "           ' �i��
'''''    sql = sql & " V.E041REVNUM, "           ' ���i�ԍ������ԍ�
'''''    sql = sql & " V.E041FACTORY, "          ' �H��
'''''    sql = sql & " V.E041OPECOND, "          ' ���Ə���
'''''    sql = sql & " S.HSXTYPE, "              ' �i�r�w�^�C�v
'''''    sql = sql & " S.HSXCDIR "               ' �i�r�w�����ʕ���
'''''    sql = sql & " from VECME009 V, TBCME018 S "
'''''    sql = sql & " where V.E041HINBAN  = S.HINBAN "
'''''    sql = sql & "   and V.E041REVNUM  = S.MNOREVNO "
'''''    sql = sql & "   and V.E041FACTORY = S.FACTORY "
'''''    sql = sql & "   and V.E041OPECOND = S.OPECOND "
'''''    sql = sql & "   and V.E040NOWPROC ='" & NOWPROC & "' "
'''''    sql = sql & "   and V.E040LSTATCLS='T' "
'''''    sql = sql & "   and V.E040RSTATCLS='T' "
'''''    sql = sql & "   and V.E040DELCLS  ='0' "
'''''   'sql = sql & "   and V.E040HOLDCLS ='0' " ' �z�[���h�u���b�N���擾
'''''    sql = sql & " order by V.E040BLOCKID, V.E041INGOTPOS "
'''''
'''''    '�f�[�^�𒊏o����
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    '���R�[�h���Ȃ��ꍇ����I��
'''''    If rs.RecordCount = 0 Then
'''''        rs.Close
'''''        ReDim records(0)
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    BlockIdBuf = vbNullString
'''''    recCnt = rs.RecordCount
'''''    j = 0
'''''    For i = 1 To recCnt
'''''        DoEvents
'''''        '�u���b�NID���̊i�[
'''''        If rs("E040BLOCKID") <> BlockIdBuf Then
'''''
'''''            j = j + 1
'''''            ReDim Preserve records(j)
'''''
'''''            With records(j)
'''''                .CRYNUM = rs("E040CRYNUM")
'''''                .IngotPos = rs("E040INGOTPOS")
'''''                .BLOCKID = rs("E040BLOCKID")   ' �u���b�NID
'''''                .UPDDATE = rs("E040UPDDATE")   ' �X�V���t
'''''                .HOLDCLS = rs("E040HOLDCLS")   ' �z�[���h�敪
'''''                BlockIdBuf = records(j).BLOCKID
'''''                .HSXTYPE = rs("HSXTYPE")
'''''                .HSXCDIR = rs("HSXCDIR")
'''''                .Judg = " "
'''''            End With
'''''
'''''            k = 1
'''''        End If
'''''
'''''        '�i�Ԃ̊i�[
'''''        ReDim Preserve records(j).hin(k)
'''''        records(j).hin(k).hinban = rs("E041HINBAN")
'''''        records(j).hin(k).mnorevno = rs("E041REVNUM")
'''''        records(j).hin(k).factory = rs("E041FACTORY")
'''''        records(j).hin(k).opecond = rs("E041OPECOND")
'''''        k = k + 1
'''''        rs.MoveNext
'''''    Next i
'''''    rs.Close
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    getBlockID = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''�w���P�����p
'''''Private Function getKouBlock(records() As type_DBDRV_scmzc_fcmkc001b_Disp, NOWPROC As String) As FUNCTION_RETURN
'''''
'''''    Dim sql         As String       'SQL�S��
'''''    Dim rs          As OraDynaset   'RecordSet
'''''    Dim recCnt      As Long
'''''    Dim motoRecCnt  As Long
'''''    Dim i           As Long
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function getKouBlock"
'''''
'''''    getKouBlock = FUNCTION_RETURN_SUCCESS
'''''
'''''    sql = " select "
'''''    sql = sql & " B.BLOCKID, "
'''''    sql = sql & " B.UPDDATE, "
'''''    sql = sql & " B.HOLDCLS, "
'''''    sql = sql & " K.HINBAN, "
'''''    sql = sql & " K.MNOREVNO, "
'''''    sql = sql & " K.FACTORY, "
'''''    sql = sql & " K.OPECOND "
'''''    sql = sql & " from  TBCME040 B,TBCMG002 K "
'''''    sql = sql & " where B.BLOCKID=K.CRYNUM "
'''''    sql = sql & "   and substr(B.BLOCKID,1,1)='8' "
'''''    sql = sql & "   and B.NOWPROC ='" & NOWPROC & "' "
'''''    sql = sql & "   and B.LSTATCLS='T' "
'''''    sql = sql & "   and B.RSTATCLS='T' "
'''''    sql = sql & "   and B.DELCLS  ='0' "
'''''   'sql = sql & "   and B.HOLDCLS ='0' " ' �z�[���h�u���b�N���擾
'''''    sql = sql & "   and K.TRANCNT =any(select max(TRANCNT) from TBCMG002 where CRYNUM=B.BLOCKID) "
'''''    sql = sql & " order by B.BLOCKID "
'''''
'''''
'''''    '�f�[�^�𒊏o����
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    If rs.RecordCount = 0 Then
'''''        rs.Close
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    motoRecCnt = UBound(records)
'''''    recCnt = rs.RecordCount
'''''    ReDim Preserve records(UBound(records) + recCnt)
'''''
'''''    For i = motoRecCnt + 1 To UBound(records)
'''''        DoEvents
'''''        ReDim records(i).HIN(1)
'''''        With records(i)
'''''            .BLOCKID = rs("BLOCKID")     ' �u���b�NID
'''''            .UPDDATE = rs("UPDDATE")     ' �X�V���t
'''''            .HOLDCLS = rs("HOLDCLS")     ' �z�[���h�敪
'''''            .HIN(1).hinban = rs("HINBAN")       ' �i��
'''''            .HIN(1).mnorevno = rs("MNOREVNO")   ' ���i�ԍ������ԍ�
'''''            .HIN(1).factory = rs("FACTORY")     ' �H��
'''''            .HIN(1).opecond = rs("OPECOND")     ' ���Ə���
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    getKouBlock = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''
'''''End Function



Public Function DBDRV_scmzc_fcmkc001b_Disp00(record0() As type_DBDRV_scmzc_fcmkc001b_Disp, _
                                             record1() As type_DBDRV_scmzc_fcmkc001b_Disp, _
                                             LWD() As cmkc001b_LockWait) As FUNCTION_RETURN

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim recCnt      As Long         '�u���b�N�Ǘ��̃��R�[�h��
    Dim i           As Long
    
    Dim j1          As Long
    Dim k1          As Long
    Dim j2          As Long
    Dim k2          As Long
    
    Dim BlockIdBuf1 As String
    Dim BlockIdBuf2 As String
    
    '<�����҂�>
    '<����҂�>
    
    '�u���b�N�Ǘ��e�[�u������u���b�NID�A�X�V���t�擾�i�������т��������̂��́j
    

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp1"
    DBDRV_scmzc_fcmkc001b_Disp00 = FUNCTION_RETURN_SUCCESS

    '�u���b�NID�A�X�V���t�̎擾
    sql = "select distinct "
    
    sql = sql & " X.XTALCA       as CRYNUM,"
    sql = sql & " B.INGOTPOS,"
    sql = sql & " X.CRYNUMCA     as BLOCKID,"
    sql = sql & " B.UPDDATE,"
    sql = sql & " B.HOLDCLS,"
    sql = sql & " X.HINBCA       as HINBAN,"        ' �i��
    sql = sql & " X.REVNUMCA     as REVNUM,"        ' ���i�ԍ������ԍ�
    sql = sql & " X.FACTORYCA    as FACTORY,"       ' �H��
    sql = sql & " X.OPECA        as OPECOND,"       ' ���Ə���
    sql = sql & " S.HSXTYPE,"                       ' �i�r�w�^�C�v
    sql = sql & " S.HSXCDIR,"                       ' �i�r�w�����ʕ���
    sql = sql & " X.INPOSCA,"
    
    sql = sql & " XT.CRYINDRSCS  as T_CRYINDRSCS,"  ' ��T���v��
    sql = sql & " XT.CRYINDOICS  as T_CRYINDOICS,"
    sql = sql & " XT.CRYINDB1CS  as T_CRYINDB1CS,"
    sql = sql & " XT.CRYINDB2CS  as T_CRYINDB2CS,"
    sql = sql & " XT.CRYINDB3CS  as T_CRYINDB3CS,"
    sql = sql & " XT.CRYINDL1CS  as T_CRYINDL1CS,"
    sql = sql & " XT.CRYINDL2CS  as T_CRYINDL2CS,"
    sql = sql & " XT.CRYINDL3CS  as T_CRYINDL3CS,"
    sql = sql & " XT.CRYINDL4CS  as T_CRYINDL4CS,"
    sql = sql & " XT.CRYINDCSCS  as T_CRYINDCSCS,"
    sql = sql & " XT.CRYINDGDCS  as T_CRYINDGDCS,"
    sql = sql & " XT.CRYINDTCS   as T_CRYINDT_CS,"
    sql = sql & " XT.CRYINDEPCS  as T_CRYINDEPCS,"
    sql = sql & " XT.CRYRESRS1CS as T_CRYRESR1CS,"
    sql = sql & " XT.CRYRESRS2CS as T_CRYRESR2CS,"
    sql = sql & " XT.CRYRESOICS  as T_CRYRESOICS,"
    sql = sql & " XT.CRYRESB1CS  as T_CRYRESB1CS,"
    sql = sql & " XT.CRYRESB2CS  as T_CRYRESB2CS,"
    sql = sql & " XT.CRYRESB3CS  as T_CRYRESB3CS,"
    sql = sql & " XT.CRYRESL1CS  as T_CRYRESL1CS,"
    sql = sql & " XT.CRYRESL2CS  as T_CRYRESL2CS,"
    sql = sql & " XT.CRYRESL3CS  as T_CRYRESL3CS,"
    sql = sql & " XT.CRYRESL4CS  as T_CRYRESL4CS,"
    sql = sql & " XT.CRYRESCSCS  as T_CRYRESCSCS,"
    sql = sql & " XT.CRYRESGDCS  as T_CRYRESGDCS,"
    sql = sql & " XT.CRYRESTCS   as T_CRYREST_CS,"
    sql = sql & " XT.CRYRESEPCS  as T_CRYRESEPCS,"
                                    
    sql = sql & " XB.CRYINDRSCS  as B_CRYINDRSCS,"  ' ���T���v��
    sql = sql & " XB.CRYINDOICS  as B_CRYINDOICS,"
    sql = sql & " XB.CRYINDB1CS  as B_CRYINDB1CS,"
    sql = sql & " XB.CRYINDB2CS  as B_CRYINDB2CS,"
    sql = sql & " XB.CRYINDB3CS  as B_CRYINDB3CS,"
    sql = sql & " XB.CRYINDL1CS  as B_CRYINDL1CS,"
    sql = sql & " XB.CRYINDL2CS  as B_CRYINDL2CS,"
    sql = sql & " XB.CRYINDL3CS  as B_CRYINDL3CS,"
    sql = sql & " XB.CRYINDL4CS  as B_CRYINDL4CS,"
    sql = sql & " XB.CRYINDCSCS  as B_CRYINDCSCS,"
    sql = sql & " XB.CRYINDGDCS  as B_CRYINDGDCS,"
    sql = sql & " XB.CRYINDTCS   as B_CRYINDT_CS,"
    sql = sql & " XB.CRYINDEPCS  as B_CRYINDEPCS,"
    sql = sql & " XB.CRYRESRS1CS as B_CRYRESR1CS,"
    sql = sql & " XB.CRYRESRS2CS as B_CRYRESR2CS,"
    sql = sql & " XB.CRYRESOICS  as B_CRYRESOICS,"
    sql = sql & " XB.CRYRESB1CS  as B_CRYRESB1CS,"
    sql = sql & " XB.CRYRESB2CS  as B_CRYRESB2CS,"
    sql = sql & " XB.CRYRESB3CS  as B_CRYRESB3CS,"
    sql = sql & " XB.CRYRESL1CS  as B_CRYRESL1CS,"
    sql = sql & " XB.CRYRESL2CS  as B_CRYRESL2CS,"
    sql = sql & " XB.CRYRESL3CS  as B_CRYRESL3CS,"
    sql = sql & " XB.CRYRESL4CS  as B_CRYRESL4CS,"
    sql = sql & " XB.CRYRESCSCS  as B_CRYRESCSCS,"
    sql = sql & " XB.CRYRESGDCS  as B_CRYRESGDCS,"
    sql = sql & " XB.CRYRESTCS   as B_CRYREST_CS,"
    sql = sql & " XB.CRYRESEPCS  as B_CRYRESEPCS,"
    
    sql = sql & " (select count(*) From XSDCS X2"                                   ' �w��������(1)�Ŏ��т�����(0)���P�J���ł�����Ό����҂�
    sql = sql & "   where X2.CRYNUMCS= X.CRYNUMCA"                                  '
    sql = sql & "     and X2.LIVKCS  ='0'"                                          ' �����敪
    sql = sql & "     and ((X2.CRYINDRSCS='1' and X2.CRYRESRS1CS='0')"              ' �����������сiRs)
    sql = sql & "      or  (X2.CRYINDOICS='1' and X2.CRYRESOICS ='0')"              ' �����������сiOi)
    sql = sql & "      or  (X2.CRYINDB1CS='1' and X2.CRYRESB1CS ='0')"              ' �����������сiB1)
    sql = sql & "      or  (X2.CRYINDB2CS='1' and X2.CRYRESB2CS ='0')"              ' �����������сiB2�j
    sql = sql & "      or  (X2.CRYINDB3CS='1' and X2.CRYRESB3CS ='0')"              ' �����������сiB3)
    sql = sql & "      or  (X2.CRYINDL1CS='1' and X2.CRYRESL1CS ='0')"              ' �����������сiL1)
    sql = sql & "      or  (X2.CRYINDL2CS='1' and X2.CRYRESL2CS ='0')"              ' �����������сiL2)
    sql = sql & "      or  (X2.CRYINDL3CS='1' and X2.CRYRESL3CS ='0')"              ' �����������сiL3)
    sql = sql & "      or  (X2.CRYINDL4CS='1' and X2.CRYRESL4CS ='0')"              ' �����������сiL4)
    sql = sql & "      or  (X2.CRYINDCSCS='1' and X2.CRYRESCSCS ='0')"              ' �����������сiCs)
    sql = sql & "      or  (X2.CRYINDGDCS='1' and X2.CRYRESGDCS ='0')"              ' �����������сiGD)
    sql = sql & "      or  (X2.CRYINDTCS ='1' and X2.CRYRESTCS  ='0')"              ' �����������сiT)
    sql = sql & "      or  (X2.CRYINDEPCS='1' and X2.CRYRESEPCS ='0')) ) as DTTYPE" ' �����������сiEPD)
    
    sql = sql & " from  XSDCA X, TBCME040 B, TBCME018 S, XSDCS XT, XSDCS XB"
    
    sql = sql & " where X.CRYNUMCA = B.BLOCKID "
    sql = sql & "   and X.GNWKNTCA ='CC600' "
    sql = sql & "   and X.LSTATBCA ='T' "
''''sql = sql & "   and X.RSTATBCA ='T' "           '' �i�グ�i�����ŕ\������Ȃ��Ȃ�i�H�j�̂ŃR�����g
    sql = sql & "   and X.LIVKCA   ='0' "
    sql = sql & "   and B.DELCLS   ='0' "
   'sql = sql & "   and B.HOLDCLS  ='0' "           ' �z�[���h�u���b�N���擾
    
    sql = sql & "   and X.HINBCA   = S.HINBAN "
    sql = sql & "   and X.REVNUMCA = S.MNOREVNO "
    sql = sql & "   and X.FACTORYCA= S.FACTORY "
    sql = sql & "   and X.OPECA    = S.OPECOND "
                                                    ' �T���v���͕K���㉺�P���Â��݂��鎖
    sql = sql & "   and XT.CRYNUMCS= X.CRYNUMCA"    ' ��T���v������
    sql = sql & "   and XT.TBKBNCS ='T'"
    sql = sql & "   and XT.LIVKCS  ='0'"
    sql = sql & "   and XB.CRYNUMCS= X.CRYNUMCA"    ' ���T���v������
    sql = sql & "   and XB.TBKBNCS ='B'"
    sql = sql & "   and XB.LIVKCS  ='0'"
    
    sql = sql & " order by X.CRYNUMCA, X.INPOSCA "

    '�f�[�^�𒊏o����
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '���R�[�h0�����͐���
    If rs.RecordCount = 0 Then
        rs.Close
        ReDim record0(0)
        ReDim record1(0)
        ReDim LWD(0)
    Else
        BlockIdBuf1 = vbNullString:  j1 = 0     ' �����҂�
        BlockIdBuf2 = vbNullString:  j2 = 0     ' ����҂�
        
        recCnt = rs.RecordCount
        For i = 1 To recCnt
            '�u���b�NID���̊i�[
            DoEvents
            
            If rs("DTTYPE") > 0 Then
            '<�����҂�>
                If rs("BLOCKID") <> BlockIdBuf1 Then
                    j1 = j1 + 1
                    ReDim Preserve record0(j1)
                    
                    With record0(j1)
                        .CRYNUM = rs("CRYNUM")
                        .IngotPos = rs("INGOTPOS")
                        .BLOCKID = rs("BLOCKID")   ' �u���b�NID
                        .UPDDATE = rs("UPDDATE")   ' �X�V���t
                        .HOLDCLS = rs("HOLDCLS")   ' �z�[���h�敪
                        .HSXTYPE = rs("HSXTYPE")
                        .HSXCDIR = rs("HSXCDIR")
                        .Judg = " "
                        BlockIdBuf1 = record0(j1).BLOCKID
                    End With
                    
                    k1 = 1
                End If
                
                '�i�Ԃ̊i�[
                ReDim Preserve record0(j1).hin(k1)
                With record0(j1).hin(k1)
                    .hinban = rs("HINBAN")
                    .mnorevno = rs("REVNUM")
                    .factory = rs("FACTORY")
                    .opecond = rs("OPECOND")
                End With
                k1 = k1 + 1
            
            Else
            '<����҂�>
                If rs("BLOCKID") <> BlockIdBuf2 Then
                    j2 = j2 + 1
                    ReDim Preserve record1(j2)
                    ReDim Preserve LWD(j2)
                    
                    With record1(j2)
                        .CRYNUM = rs("CRYNUM")
                        .IngotPos = rs("INGOTPOS")
                        .BLOCKID = rs("BLOCKID")   ' �u���b�NID
                        .UPDDATE = rs("UPDDATE")   ' �X�V���t
                        .HOLDCLS = rs("HOLDCLS")   ' �z�[���h�敪
                        .HSXTYPE = rs("HSXTYPE")
                        .HSXCDIR = rs("HSXCDIR")
                        .Judg = " "
                        BlockIdBuf2 = record1(j2).BLOCKID
                    End With
                    
                    ' ����(1)�ɑ΂�����т����ׂĐݒ肳��Ă���̂�
                    ' ���f(2)�A����(3)�ɑ΂��Ă����т����ׂĐݒ肳��Ă��邩�`�F�b�N
                    LWD(j2).flag = False
                    ' ��T���v��
                    If (rs("T_CRYINDRSCS") = "3" And rs("T_CRYRESR2CS") = "0") _
                    Or (rs("T_CRYINDRSCS") > "1" And rs("T_CRYRESR1CS") = "0") _
                    Or (rs("T_CRYINDOICS") > "1" And rs("T_CRYRESOICS") = "0") _
                    Or (rs("T_CRYINDB1CS") > "1" And rs("T_CRYRESB1CS") = "0") _
                    Or (rs("T_CRYINDB2CS") > "1" And rs("T_CRYRESB2CS") = "0") _
                    Or (rs("T_CRYINDB3CS") > "1" And rs("T_CRYRESB3CS") = "0") _
                    Or (rs("T_CRYINDL1CS") > "1" And rs("T_CRYRESL1CS") = "0") _
                    Or (rs("T_CRYINDL2CS") > "1" And rs("T_CRYRESL2CS") = "0") _
                    Or (rs("T_CRYINDL3CS") > "1" And rs("T_CRYRESL3CS") = "0") _
                    Or (rs("T_CRYINDL4CS") > "1" And rs("T_CRYRESL4CS") = "0") _
                    Or (rs("T_CRYINDCSCS") > "1" And rs("T_CRYRESCSCS") = "0") _
                    Or (rs("T_CRYINDGDCS") > "1" And rs("T_CRYRESGDCS") = "0") _
                    Or (rs("T_CRYINDT_CS") > "1" And rs("T_CRYREST_CS") = "0") _
                    Or (rs("T_CRYINDEPCS") > "1" And rs("T_CRYRESEPCS") = "0") Then
                        LWD(j2).flag = True
                    End If
                    ' ���T���v��
                    If (rs("B_CRYINDRSCS") = "3" And rs("B_CRYRESR2CS") = "0") _
                    Or (rs("B_CRYINDRSCS") > "1" And rs("B_CRYRESR1CS") = "0") _
                    Or (rs("B_CRYINDOICS") > "1" And rs("B_CRYRESOICS") = "0") _
                    Or (rs("B_CRYINDB1CS") > "1" And rs("B_CRYRESB1CS") = "0") _
                    Or (rs("B_CRYINDB2CS") > "1" And rs("B_CRYRESB2CS") = "0") _
                    Or (rs("B_CRYINDB3CS") > "1" And rs("B_CRYRESB3CS") = "0") _
                    Or (rs("B_CRYINDL1CS") > "1" And rs("B_CRYRESL1CS") = "0") _
                    Or (rs("B_CRYINDL2CS") > "1" And rs("B_CRYRESL2CS") = "0") _
                    Or (rs("B_CRYINDL3CS") > "1" And rs("B_CRYRESL3CS") = "0") _
                    Or (rs("B_CRYINDL4CS") > "1" And rs("B_CRYRESL4CS") = "0") _
                    Or (rs("B_CRYINDCSCS") > "1" And rs("B_CRYRESCSCS") = "0") _
                    Or (rs("B_CRYINDGDCS") > "1" And rs("B_CRYRESGDCS") = "0") _
                    Or (rs("B_CRYINDT_CS") > "1" And rs("B_CRYREST_CS") = "0") _
                    Or (rs("B_CRYINDEPCS") > "1" And rs("B_CRYRESEPCS") = "0") Then
                        LWD(j2).flag = True
                    End If
                    
                    ' ��type_DBDRV_scmzc_fcmkc001b_Disp�ɂ̓T���v������ێ�����\���̂���`����Ă��邪
                    ' �@�t���O�ݒ�̔���ɂ����g�p���Ă��Ȃ��̂Őݒ肵�Ȃ��ł���
                    ' �@�ݒ肷��ꍇ��cmkc001b_DBDataCheck1���Q�Ɓi�������Ԃ������̂Ŋ֐��͕s�g�p�̎��j
                    
                    k2 = 1
                End If
                
                '�i�Ԃ̊i�[
                ReDim Preserve record1(j2).hin(k2)
                With record1(j2).hin(k2)
                    .hinban = rs("HINBAN")
                    .mnorevno = rs("REVNUM")
                    .factory = rs("FACTORY")
                    .opecond = rs("OPECOND")
                End With
                k2 = k2 + 1
            End If
            
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
    DBDRV_scmzc_fcmkc001b_Disp00 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



''''''�T�v    :�҂��ꗗ �����\���p�c�a�h���C�o�i�����҂��j
''''''���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
''''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
''''''        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
''''''����    :
''''''����    :2001/07/06 ���{ �쐬
'''''Public Function DBDRV_scmzc_fcmkc001b_Disp1(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''
'''''    Dim sql         As String       'SQL�S��
'''''    Dim rs          As OraDynaset   'RecordSet
'''''    Dim recCnt      As Long         '�u���b�N�Ǘ��̃��R�[�h��
'''''    Dim i           As Long
'''''    Dim j           As Long
'''''    Dim k           As Long
'''''    Dim BlockIdBuf  As String
'''''
'''''    '<�����҂���
'''''    '�u���b�N�Ǘ��e�[�u������u���b�NID�A�X�V���t�擾�i�������т��������̂��́j
'''''
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp1"
'''''
'''''    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_SUCCESS
'''''
'''''    '�u���b�NID�A�X�V���t�̎擾
'''''    sql = "select distinct "
'''''
'''''    sql = sql & " X.XTALCA    as CRYNUM, "
'''''    sql = sql & " B.INGOTPOS, "
'''''    sql = sql & " X.CRYNUMCA  as BLOCKID, "
'''''    sql = sql & " B.UPDDATE, "
'''''    sql = sql & " B.HOLDCLS, "
'''''    sql = sql & " X.HINBCA    as HINBAN, "          ' �i��
'''''    sql = sql & " X.REVNUMCA  as REVNUM, "          ' ���i�ԍ������ԍ�
'''''    sql = sql & " X.FACTORYCA as FACTORY, "         ' �H��
'''''    sql = sql & " X.OPECA     as OPECOND, "         ' ���Ə���
'''''    sql = sql & " S.HSXTYPE, "                      ' �i�r�w�^�C�v
'''''    sql = sql & " S.HSXCDIR, "                      ' �i�r�w�����ʕ���
'''''    sql = sql & " X.INPOSCA   as INGOTPOS "
'''''
'''''
'''''    sql = sql & " from  XSDCA X, TBCME040 B, TBCME018 S, XSDCS X2 "
'''''
'''''    sql = sql & " where X.CRYNUMCA = B.BLOCKID "
'''''    sql = sql & "   and X.CRYNUMCA = X2.CRYNUMCS "
'''''    sql = sql & "   and X.HINBCA   = S.HINBAN "
'''''    sql = sql & "   and X.REVNUMCA = S.MNOREVNO "
'''''    sql = sql & "   and X.FACTORYCA= S.FACTORY "
'''''    sql = sql & "   and X.OPECA    = S.OPECOND "
'''''
'''''    sql = sql & "   and X.GNWKNTCA='CC600' "
'''''    sql = sql & "   and X.LSTATBCA='T' "
'''''    sql = sql & "   and X.RSTATBCA='T' "
'''''    sql = sql & "   and X.LIVKCA  ='0' "
'''''    sql = sql & "   and B.DELCLS  ='0' "
'''''   'sql = sql & "   and B.HOLDCLS ='0' " ' �z�[���h�u���b�N���擾
'''''
'''''    '�w����0�łȂ����т�0
'''''    sql = sql & " and ((X2.CRYINDRSCS<>'0' and X2.CRYRESRS1CS='0')"       ' �����������сiRs)
'''''    sql = sql & "   or (X2.CRYINDOICS<>'0' and X2.CRYRESOICS ='0')"       ' �����������сiOi)
'''''    sql = sql & "   or (X2.CRYINDB1CS<>'0' and X2.CRYRESB1CS ='0')"       ' �����������сiB1)
'''''    sql = sql & "   or (X2.CRYINDB2CS<>'0' and X2.CRYRESB2CS ='0')"       ' �����������сiB2�j
'''''    sql = sql & "   or (X2.CRYINDB3CS<>'0' and X2.CRYRESB3CS ='0')"       ' �����������сiB3)
'''''    sql = sql & "   or (X2.CRYINDL1CS<>'0' and X2.CRYRESL1CS ='0')"       ' �����������сiL1)
'''''    sql = sql & "   or (X2.CRYINDL2CS<>'0' and X2.CRYRESL2CS ='0')"       ' �����������сiL2)
'''''    sql = sql & "   or (X2.CRYINDL3CS<>'0' and X2.CRYRESL3CS ='0')"       ' �����������сiL3)
'''''    sql = sql & "   or (X2.CRYINDL4CS<>'0' and X2.CRYRESL4CS ='0')"       ' �����������сiL4)
'''''    sql = sql & "   or (X2.CRYINDCSCS<>'0' and X2.CRYRESCSCS ='0')"       ' �����������сiCs)
'''''    sql = sql & "   or (X2.CRYINDGDCS<>'0' and X2.CRYRESGDCS ='0')"       ' �����������сiGD)
'''''    sql = sql & "   or (X2.CRYINDTCS <>'0' and X2.CRYRESTCS  ='0')"       ' �����������сiT)
'''''    sql = sql & "   or (X2.CRYINDEPCS<>'0' and X2.CRYRESEPCS ='0'))"      ' �����������сiEPD)
'''''
'''''    sql = sql & " order by X.CRYNUMCA, X.INPOSCA "
'''''
'''''    '�f�[�^�𒊏o����
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    '���R�[�h0�����͐���
'''''    If rs.RecordCount = 0 Then
'''''        rs.Close
'''''        ReDim records(0)
'''''    Else
'''''        BlockIdBuf = vbNullString
'''''        recCnt = rs.RecordCount
'''''        j = 0
'''''        For i = 1 To recCnt
'''''            DoEvents
'''''        '�u���b�NID���̊i�[
'''''            If rs("BLOCKID") <> BlockIdBuf Then
'''''
'''''                j = j + 1
'''''                ReDim Preserve records(j)
'''''
'''''                With records(j)
'''''                    .CRYNUM = rs("CRYNUM")
'''''                    .IngotPos = rs("INGOTPOS")
'''''                    .BLOCKID = rs("BLOCKID")   ' �u���b�NID
'''''                    .UPDDATE = rs("UPDDATE")   ' �X�V���t
'''''                    .HOLDCLS = rs("HOLDCLS")   ' �z�[���h�敪
'''''                    BlockIdBuf = records(j).BLOCKID
'''''                    .HSXTYPE = rs("HSXTYPE")
'''''                    .HSXCDIR = rs("HSXCDIR")
'''''                    .Judg = " "
'''''                End With
'''''
'''''                k = 1
'''''            End If
'''''
'''''            '�i�Ԃ̊i�[
'''''            ReDim Preserve records(j).hin(k)
'''''            records(j).hin(k).hinban = rs("HINBAN")
'''''            records(j).hin(k).mnorevno = rs("REVNUM")
'''''            records(j).hin(k).factory = rs("FACTORY")
'''''            records(j).hin(k).opecond = rs("OPECOND")
'''''            k = k + 1
'''''            rs.MoveNext
'''''        Next i
'''''        rs.Close
'''''
'''''    End If
'''''
'''''
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''�T�v    :�҂��ꗗ �����\���p�c�a�h���C�o�i����҂��j
''''''���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
''''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
''''''        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
''''''����    :
''''''����    :2001/07/06 ���{ �쐬
'''''Public Function DBDRV_scmzc_fcmkc001b_Disp2(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''
'''''    '������҂���
'''''    '�����҂���������Ă���ꍇ�Ƌt�łO������Ȃ�����
'''''    Dim sql         As String       'SQL�S��
'''''    Dim rs          As OraDynaset   'RecordSet
'''''    Dim recCnt      As Long         '�u���b�N�Ǘ��̃��R�[�h��
'''''    Dim i           As Long
'''''    Dim j           As Long
'''''    Dim k           As Long
'''''    Dim BlockIdBuf  As String
'''''
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp2"
'''''
'''''    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_SUCCESS
'''''
'''''    sql = "select distinct "
'''''
'''''    sql = sql & " X.XTALCA    as CRYNUM,  "
'''''    sql = sql & " B.INGOTPOS  as ss, "
''''''   sql = sql & " B.LENGTH, "                   ' �����ǉ� 2001/11/8
'''''    sql = sql & " X.CRYNUMCA  as BLOCKID, "
'''''    sql = sql & " B.UPDDATE, "
'''''    sql = sql & " B.HOLDCLS, "
'''''    sql = sql & " X.HINBCA    as HINBAN,  "     ' �i��
'''''    sql = sql & " X.REVNUMCA  as REVNUM,  "     ' ���i�ԍ������ԍ�
'''''    sql = sql & " X.FACTORYCA as FACTORY, "     ' �H��
'''''    sql = sql & " X.OPECA     as OPECOND, "     ' ���Ə���
'''''    sql = sql & " S.HSXTYPE, "                  ' �i�r�w�^�C�v
'''''    sql = sql & " S.HSXCDIR, "                  ' �i�r�w�����ʕ���
'''''    sql = sql & " X.INPOSCA   as INGOTPOS "
'''''
''''''''''                '����NG�����邩�ǂ���
''''''''''    sql = sql & " (select count(*) from XSDCS X21 "
''''''''''    sql = sql & "  where   X21.CRYNUMCS=X.CRYNUMCA"
''''''''''    sql = sql & "    and ((X21.CRYINDRSCS<>'0' and X21.CRYRESRS1CS='2')"            ' �����������сiRs)
''''''''''    sql = sql & "      or (X21.CRYINDOICS<>'0' and X21.CRYRESOICS ='2')"            ' �����������сiOi)
''''''''''    sql = sql & "      or (X21.CRYINDB1CS<>'0' and X21.CRYRESB1CS ='2')"            ' �����������сiB1)
''''''''''    sql = sql & "      or (X21.CRYINDB2CS<>'0' and X21.CRYRESB2CS ='2')"            ' �����������сiB2�j
''''''''''    sql = sql & "      or (X21.CRYINDB3CS<>'0' and X21.CRYRESB3CS ='2')"            ' �����������сiB3)
''''''''''    sql = sql & "      or (X21.CRYINDL1CS<>'0' and X21.CRYRESL1CS ='2')"            ' �����������сiL1)
''''''''''    sql = sql & "      or (X21.CRYINDL2CS<>'0' and X21.CRYRESL2CS ='2')"            ' �����������сiL2)
''''''''''    sql = sql & "      or (X21.CRYINDL3CS<>'0' and X21.CRYRESL3CS ='2')"            ' �����������сiL3)
''''''''''    sql = sql & "      or (X21.CRYINDL4CS<>'0' and X21.CRYRESL4CS ='2')"            ' �����������сiL4)
''''''''''    sql = sql & "      or (X21.CRYINDCSCS<>'0' and X21.CRYRESCSCS ='2')"            ' �����������сiCs)
''''''''''    sql = sql & "      or (X21.CRYINDGDCS<>'0' and X21.CRYRESGDCS ='2')"            ' �����������сiGD)
''''''''''    sql = sql & "      or (X21.CRYINDTCS <>'0' and X21.CRYRESTCS  ='2')"            ' �����������сiT)
''''''''''    sql = sql & "      or (X21.CRYINDEPCS<>'0' and X21.CRYRESEPCS ='2')) ) as J "   ' �����������сiEPD)
'''''
'''''    sql = sql & " from  XSDCA X, TBCME040 B, TBCME018 S"
'''''    sql = sql & " where X.CRYNUMCA = B.BLOCKID "
'''''    sql = sql & "   and X.HINBCA   = S.HINBAN "
'''''    sql = sql & "   and X.REVNUMCA = S.MNOREVNO "
'''''    sql = sql & "   and X.FACTORYCA= S.FACTORY "
'''''    sql = sql & "   and X.OPECA    = S.OPECOND "
'''''
'''''    '�H���R�[�h�A��ԁA�敪�̏����w��
'''''
'''''    sql = sql & "   and X.GNWKNTCA='CC600' "
'''''    sql = sql & "   and X.LSTATBCA='T' "
'''''    sql = sql & "   and X.RSTATBCA='T' "
'''''    sql = sql & "   and X.LIVKCA  ='0' "
'''''    sql = sql & "   and B.DELCLS  ='0' "
'''''   'sql = sql & "   and B.HOLDCLS ='0' " ' �z�[���h�u���b�N���擾
'''''
'''''''                '�u���b�N���Ɋ܂܂��i�Ԃ�����
'''''''    sql = sql & " and (( B.INGOTPOS >= H.INGOTPOS "
'''''''    sql = sql & " and B.INGOTPOS < H.INGOTPOS + H.LENGTH ) "
'''''''    sql = sql & " or ( B.INGOTPOS + B.LENGTH > H.INGOTPOS "
'''''''    sql = sql & " and B.INGOTPOS + B.LENGTH < H.INGOTPOS + H.LENGTH  ) "
'''''''    sql = sql & " or ( H.INGOTPOS >= B.INGOTPOS "
'''''''    sql = sql & " and H.INGOTPOS < B.INGOTPOS + B.LENGTH ) "
'''''''    sql = sql & " or ( H.INGOTPOS + H.LENGTH > B.INGOTPOS "
'''''''    sql = sql & " and H.INGOTPOS + H.LENGTH < B.INGOTPOS + B.LENGTH )) "
'''''' �u���b�N���L�[�Ƀf�[�^���擾����̂Ŕ͈͊O�͖���
'''''
'''''                '�w����0�łȂ����т�0�łȂ��T���v�����㉺�Q�����邩
'''''    sql = sql & " and 2=(select count(*) From XSDCS X22"
'''''    sql = sql & "        where  X22.CRYNUMCS=X.CRYNUMCA"
'''''
'''''    sql = sql & "         and ((X22.CRYINDRSCS<>'0' and X22.CRYRESRS1CS<>'0')"          ' �����������сiRs)
'''''    sql = sql & "          or  (X22.CRYINDOICS<>'0' and X22.CRYRESOICS <>'0')"          ' �����������сiOi)
'''''    sql = sql & "          or  (X22.CRYINDB1CS<>'0' and X22.CRYRESB1CS <>'0')"          ' �����������сiB1)
'''''    sql = sql & "          or  (X22.CRYINDB2CS<>'0' and X22.CRYRESB2CS <>'0')"          ' �����������сiB2�j
'''''    sql = sql & "          or  (X22.CRYINDB3CS<>'0' and X22.CRYRESB3CS <>'0')"          ' �����������сiB3)
'''''    sql = sql & "          or  (X22.CRYINDL1CS<>'0' and X22.CRYRESL1CS <>'0')"          ' �����������сiL1)
'''''    sql = sql & "          or  (X22.CRYINDL2CS<>'0' and X22.CRYRESL2CS <>'0')"          ' �����������сiL2)
'''''    sql = sql & "          or  (X22.CRYINDL3CS<>'0' and X22.CRYRESL3CS <>'0')"          ' �����������сiL3)
'''''    sql = sql & "          or  (X22.CRYINDL4CS<>'0' and X22.CRYRESL4CS <>'0')"          ' �����������сiL4)
'''''    sql = sql & "          or  (X22.CRYINDCSCS<>'0' and X22.CRYRESCSCS <>'0')"          ' �����������сiCs)
'''''    sql = sql & "          or  (X22.CRYINDGDCS<>'0' and X22.CRYRESGDCS <>'0')"          ' �����������сiGD)
'''''    sql = sql & "          or  (X22.CRYINDTCS <>'0' and X22.CRYRESTCS  <>'0')"          ' �����������сiT)
'''''    sql = sql & "          or  (X22.CRYINDEPCS<>'0' and X22.CRYRESEPCS <>'0')) )"       ' �����������сiEPD)
''''''''''    sql = sql & "          and (X22.CRYINDRSCS='0' or X22.CRYRESRS1CS<>'0')"          ' �����������сiRs)
''''''''''    sql = sql & "          and (X22.CRYINDOICS='0' or X22.CRYRESOICS <>'0')"          ' �����������сiOi)
''''''''''    sql = sql & "          and (X22.CRYINDB1CS='0' or X22.CRYRESB1CS <>'0')"          ' �����������сiB1)
''''''''''    sql = sql & "          and (X22.CRYINDB2CS='0' or X22.CRYRESB2CS <>'0')"          ' �����������сiB2�j
''''''''''    sql = sql & "          and (X22.CRYINDB3CS='0' or X22.CRYRESB3CS <>'0')"          ' �����������сiB3)
''''''''''    sql = sql & "          and (X22.CRYINDL1CS='0' or X22.CRYRESL1CS <>'0')"          ' �����������сiL1)
''''''''''    sql = sql & "          and (X22.CRYINDL2CS='0' or X22.CRYRESL2CS <>'0')"          ' �����������сiL2)
''''''''''    sql = sql & "          and (X22.CRYINDL3CS='0' or X22.CRYRESL3CS <>'0')"          ' �����������сiL3)
''''''''''    sql = sql & "          and (X22.CRYINDL4CS='0' or X22.CRYRESL4CS <>'0')"          ' �����������сiL4)
''''''''''    sql = sql & "          and (X22.CRYINDCSCS='0' or X22.CRYRESCSCS <>'0')"          ' �����������сiCs)
''''''''''    sql = sql & "          and (X22.CRYINDGDCS='0' or X22.CRYRESGDCS <>'0')"          ' �����������сiGD)
''''''''''    sql = sql & "          and (X22.CRYINDTCS ='0' or X22.CRYRESTCS  <>'0')"          ' �����������сiT)
''''''''''    sql = sql & "          and (X22.CRYINDEPCS='0' or X22.CRYRESEPCS <>'0') )"        ' �����������сiEPD)
'''''
'''''''    sql = sql & " order by B.BLOCKID, H.INGOTPOS "
'''''    sql = sql & " order by X.CRYNUMCA, X.INPOSCA "
'''''
'''''    '�f�[�^�𒊏o����
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    '���R�[�h0�����͐���
'''''    If rs.RecordCount = 0 Then
'''''        rs.Close
'''''        ReDim records(0)
'''''    Else
'''''        BlockIdBuf = vbNullString
'''''        recCnt = rs.RecordCount
'''''        j = 0
'''''        For i = 1 To recCnt
'''''            DoEvents
'''''        '�u���b�NID���̊i�[
'''''            If rs("BLOCKID") <> BlockIdBuf Then
'''''
'''''                j = j + 1
'''''                ReDim Preserve records(j)
'''''
'''''                With records(j)
'''''                    .CRYNUM = rs("CRYNUM")
'''''                    .IngotPos = rs("ss")
''''''                   .LENGTH = rs("LENGTH")      ' ����
'''''                    .BLOCKID = rs("BLOCKID")    ' �u���b�NID
'''''                    .UPDDATE = rs("UPDDATE")    ' �X�V���t
'''''                    .HOLDCLS = rs("HOLDCLS")    ' �z�[���h�敪
'''''                    BlockIdBuf = records(j).BLOCKID
'''''                    .HSXTYPE = rs("HSXTYPE")
'''''                    .HSXCDIR = rs("HSXCDIR")
''''''                    If rs("J") > 0 Then
''''''
''''''                        .Judg = "2"
''''''                    Else
'''''                        .Judg = "1"
''''''                    End If
'''''
'''''                End With
'''''                k = 1
'''''            End If
'''''
'''''            '�i�Ԃ̊i�[
'''''            ReDim Preserve records(j).hin(k)
'''''            records(j).hin(k).hinban = rs("HINBAN")
'''''            records(j).hin(k).mnorevno = rs("REVNUM")
'''''            records(j).hin(k).factory = rs("FACTORY")
'''''            records(j).hin(k).opecond = rs("OPECOND")
'''''            k = k + 1
'''''            rs.MoveNext
'''''        Next i
'''''        rs.Close
'''''
'''''    End If
'''''
'''''
''''''''''    '�w���P�������ю擾
''''''''''    If getKouBlock(records(), "CC600") = FUNCTION_RETURN_FAILURE Then
''''''''''       DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
''''''''''       GoTo proc_exit
''''''''''    End If
'''''
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function



'''''�T�v    :�҂��ꗗ �����\���p�c�a�h���C�o�i���o�҂��j
'''''���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
'''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
'''''        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'''''����    :
'''''����    :2001/07/06 ���{ �쐬
''''Public Function DBDRV_scmzc_fcmkc001b_Disp3(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
''''
''''    '�����o�҂���
''''    'CC700�̂���
''''
''''    '�G���[�n���h���̐ݒ�
''''    On Error GoTo proc_err
''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp3"
''''
''''
''''    DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_SUCCESS
''''
''''    '�u���b�NID��X�V���t�A�i�ԓ��擾
''''    If getBlockID(records(), "CC700") = FUNCTION_RETURN_FAILURE Then
''''        DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
''''        GoTo proc_exit
''''    End If
''''
''''
'''''    '�w���P�������ю擾
'''''    If getKouBlock(records(), "CC700") = FUNCTION_RETURN_FAILURE Then
'''''       DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
'''''       GoTo proc_exit
'''''    End If
''''
''''proc_exit:
''''    '�I��
''''    gErr.Pop
''''    Exit Function
''''
''''proc_err:
''''    '�G���[�n���h��
''''    gErr.HandleError
''''    DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
''''    Resume proc_exit
''''End Function
''''
''''
''''
'''''�T�v    :�҂��ꗗ �����\���p�c�a�h���C�o�i�����w���҂��j
'''''���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
'''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
'''''        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'''''����    :
'''''����    :2001/07/06 ���{ �쐬
''''Public Function DBDRV_scmzc_fcmkc001b_Disp4(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
''''
''''    '�������w���҂���
''''    'CC710�̂���
''''
''''    '�u���b�NID��X�V���t�擾
''''
''''    '�G���[�n���h���̐ݒ�
''''    On Error GoTo proc_err
''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp4"
''''
''''    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_SUCCESS
''''
''''
''''    '�u���b�NID��X�V���t�A�i�ԓ��擾
''''    If getBlockID(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
''''        DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
''''        GoTo proc_exit
''''    End If
''''
'''''2000/08/24 S.Sano Start
'''''    '�w���P�������ю擾
'''''    If getKouBlock(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
'''''       DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
'''''       GoTo proc_exit
'''''    End If
'''''2000/08/24 S.Sano End
''''
''''
''''proc_exit:
''''    '�I��
''''    gErr.Pop
''''    Exit Function
''''
''''proc_err:
''''    '�G���[�n���h��
''''    gErr.HandleError
''''    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
''''    Resume proc_exit
''''End Function


'''''Public Function cmkc001b_DBDataCheck1(LWD() As cmkc001b_LockWait, Wd1() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
''''''    Dim typ_A As typ_AllTypes        '�S���\����
''''''    Dim c0 As Integer
''''''    Dim sErrMsg As String
''''''    Dim NothingFlag As Boolean
''''''    Dim FuncAns As FUNCTION_RETURN
''''''    For c0 = 1 To UBound(Wd1())
''''''        NothingFlag = False
''''''        FuncAns = DBDRV_scmzc_fcmkc001b_Disp(Wd1(c0).BLOCKID, typ_A.typ_si, typ_A.typ_cr, typ_A.typ_zi, sErrMsg, NothingFlag)
''''''        LWD(c0).flag = NothingFlag
''''''    Next
'''''
'''''
'''''    Dim l   As Long
'''''    Dim m   As Long
'''''    Dim sql As String
'''''    Dim rs  As OraDynaset    'RecordSet
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function cmkc001b_DBDataCheck1"
'''''
'''''
'''''    cmkc001b_DBDataCheck1 = FUNCTION_RETURN_SUCCESS
'''''
'''''    Set rs = Nothing
'''''
'''''#If SPEEDUP Then   '���������� 02.1.28-2.15 �쑺
''''''����������
''''''���ƂȂ�u���b�N�Ƃ��̗��[�T���v���ɂ��āA������Ԃ��܂Ƃ߂Ď擾
''''''SQL�̔��s�񐔂�}�����ă��������ł̏����ɐ؂芷����
'''''Dim SMP()   As tSmpMng
'''''Dim idx     As Integer
'''''Dim topIdx  As Integer
'''''Dim botIdx  As Integer
'''''
'''''Debug.Print " 1:" & Time
'''''    sql = vbNullString
''''''    sql = sql & "select"
''''''    sql = sql & "  B.BLOCKID, B.INGOTPOS as TOPPOS, B.INGOTPOS+LENGTH as BOTPOS"
''''''    sql = sql & ", S.CRYNUM, S.INGOTPOS, SMPKBN, HINBAN, REVNUM, FACTORY, OPECOND"
''''''    sql = sql & ", CRYINDRS, CRYRESRS, CRYINDOI, CRYRESOI"
''''''    sql = sql & ", CRYINDB1, CRYRESB1, CRYINDB2, CRYRESB2, CRYINDB3, CRYRESB3"
''''''    sql = sql & ", CRYINDL1, CRYRESL1, CRYINDL2, CRYRESL2, CRYINDL3, CRYRESL3, CRYINDL4, CRYRESL4"
''''''    sql = sql & ", CRYINDCS, CRYRESCS, CRYINDGD, CRYRESGD, CRYINDT, CRYREST, CRYINDEP, CRYRESEP "
''''''    sql = sql & "from TBCME043 S, TBCME040 B "
''''''    sql = sql & "where S.CRYNUM=B.CRYNUM"
''''''    sql = sql & "  and B.INGOTPOS>=0"
''''''    sql = sql & "  and B.DELCLS='0'"
''''''    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
''''''    sql = sql & "  and B.RSTATCLS='T'"
''''''    sql = sql & "  and B.HOLDCLS='0'"
''''''    sql = sql & "  and ((S.INGOTPOS=B.INGOTPOS) or (S.INGOTPOS=B.INGOTPOS+B.LENGTH)) "
''''''    sql = sql & "order by B.BLOCKID, S.INGOTPOS, S.SMPKBN"
'''''
'''''    sql = sql & "select "
'''''    sql = sql & "B.BLOCKID,  B.INGOTPOS as TOPPOS,    B.INGOTPOS+LENGTH as BOTPOS, "
'''''
'''''    sql = sql & "S.XTALCS,   S.INPOSCS,   SMPKBNCS,   HINBCS,     REVNUMCS,   FACTORYCS,  OPECS, "
'''''    sql = sql & "CRYINDRSCS, CRYRESRS1CS, CRYINDOICS, CRYRESOICS, CRYINDB1CS, CRYRESB1CS, CRYINDB2CS,"
'''''    sql = sql & "CRYRESB2CS, CRYINDB3CS,  CRYRESB3CS, CRYINDL1CS, CRYRESL1CS, CRYINDL2CS, CRYRESL2CS,"
'''''    sql = sql & "CRYINDL3CS, CRYRESL3CS,  CRYINDL4CS, CRYRESL4CS, CRYINDCSCS, CRYRESCSCS, CRYINDGDCS,"
'''''    sql = sql & "CRYRESGDCS, CRYINDTCS,   CRYRESTCS,  CRYINDEPCS, CRYRESEPCS "
'''''
'''''    sql = sql & "from  XSDCS S, TBCME040 B "
'''''
'''''    sql = sql & "where S.XTALCS  = B.CRYNUM"
'''''
'''''    sql = sql & "  and B.INGOTPOS>=0"
'''''    sql = sql & "  and B.DELCLS  = '0'"
'''''    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
'''''    sql = sql & "  and B.RSTATCLS='T'"
'''''    sql = sql & "  and B.HOLDCLS ='0'"
'''''    sql = sql & "  and ((S.INPOSCS=B.INGOTPOS) or (S.INPOSCS=B.INGOTPOS+B.LENGTH)) "
'''''
'''''    sql = sql & "order by B.BLOCKID, S.INPOSCS, S.SMPKBNCS"
'''''
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
'''''    ReDim SMP(rs.RecordCount)
'''''    With SMP(0)
'''''        .BLOCKID = " "
'''''        .CRYNUM = " "
'''''        .SMPKBN = " "
'''''        .hinban = " "
'''''        .factory = " "
'''''        .opecond = " "
'''''        .CRYINDRS = " "
'''''        .CRYRESRS = " "
'''''        .CRYINDOI = " "
'''''        .CRYRESOI = " "
'''''        .CRYINDB1 = " "
'''''        .CRYRESB1 = " "
'''''        .CRYINDB2 = " "
'''''        .CRYRESB2 = " "
'''''        .CRYINDB3 = " "
'''''        .CRYRESB3 = " "
'''''        .CRYINDL1 = " "
'''''        .CRYRESL1 = " "
'''''        .CRYINDL2 = " "
'''''        .CRYRESL2 = " "
'''''        .CRYINDL3 = " "
'''''        .CRYRESL3 = " "
'''''        .CRYINDL4 = " "
'''''        .CRYRESL4 = " "
'''''        .CRYINDCS = " "
'''''        .CRYRESCS = " "
'''''        .CRYINDGD = " "
'''''        .CRYRESGD = " "
'''''        .CRYINDT = " "
'''''        .CRYREST = " "
'''''        .CRYINDEP = " "
'''''        .CRYRESEP = " "
'''''    End With
'''''
'''''    For l = 1 To rs.RecordCount
'''''        With SMP(l)
'''''            .BLOCKID = rs("BLOCKID")
'''''            .TOPPOS = rs("TOPPOS")
'''''            .BOTPOS = rs("BOTPOS")
'''''            .CRYNUM = rs("XTALCS")
'''''            .IngotPos = rs("INPOSCS")
'''''            .SMPKBN = rs("SMPKBNCS")
'''''            .hinban = rs("HINBCS")
'''''            .REVNUM = rs("REVNUMCS")
'''''            .factory = rs("FACTORYCS")
'''''            .opecond = rs("OPECS")
'''''            .CRYINDRS = rs("CRYINDRSCS")
'''''            .CRYRESRS = rs("CRYRESRS1CS")
'''''            .CRYINDOI = rs("CRYINDOICS")
'''''            .CRYRESOI = rs("CRYRESOICS")
'''''            .CRYINDB1 = rs("CRYINDB1CS")
'''''            .CRYRESB1 = rs("CRYRESB1CS")
'''''            .CRYINDB2 = rs("CRYINDB2CS")
'''''            .CRYRESB2 = rs("CRYRESB2CS")
'''''            .CRYINDB3 = rs("CRYINDB3CS")
'''''            .CRYRESB3 = rs("CRYRESB3CS")
'''''            .CRYINDL1 = rs("CRYINDL1CS")
'''''            .CRYRESL1 = rs("CRYRESL1CS")
'''''            .CRYINDL2 = rs("CRYINDL2CS")
'''''            .CRYRESL2 = rs("CRYRESL2CS")
'''''            .CRYINDL3 = rs("CRYINDL3CS")
'''''            .CRYRESL3 = rs("CRYRESL3CS")
'''''            .CRYINDL4 = rs("CRYINDL4CS")
'''''            .CRYRESL4 = rs("CRYRESL4CS")
'''''            .CRYINDCS = rs("CRYINDCSCS")
'''''            .CRYRESCS = rs("CRYRESCSCS")
'''''            .CRYINDGD = rs("CRYINDGDCS")
'''''            .CRYRESGD = rs("CRYRESGDCS")
'''''            .CRYINDT = rs("CRYINDTCS")
'''''            .CRYREST = rs("CRYRESTCS")
'''''            .CRYINDEP = rs("CRYINDEPCS")
'''''            .CRYRESEP = rs("CRYRESEPCS")
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''    Set rs = Nothing
'''''Debug.Print " 2:" & Time
'''''#End If
'''''
'''''    For l = 1 To UBound(Wd1())
'''''        DoEvents
'''''        LWD(l).flag = False
''''''Debug.Print " " & l & ":" & Time
'''''
'''''        With Wd1(l)
'''''
'''''        ' �w���P�����̃u���b�N�͖������łn�j
'''''        If Mid$(.BLOCKID, 1, 1) <> "8" Then
'''''
'''''            ReDim .SMP(2)
'''''
'''''            ' �㉺�̃T���v�����擾
'''''#If SPEEDUP Then   '���������� 02.1.28-2.15 �쑺
''''''����������
''''''�ꊇ�擾����������Ԕz�񂩂�A�f�[�^���擾����悤�ɉ���
'''''            For m = 1 To 2
'''''                DoEvents
'''''
'''''                topIdx = 0
'''''                botIdx = 0
'''''                For idx = 1 To UBound(SMP)
'''''                    If (SMP(idx).BLOCKID = .BLOCKID) Then
'''''                        If (SMP(idx).SMPKBN = "T") Then
'''''                            topIdx = idx
'''''                        Else
'''''                            botIdx = idx
'''''                        End If
'''''                    ElseIf SMP(idx).BLOCKID > .BLOCKID Then
'''''                        Exit For
'''''                    End If
'''''                Next
'''''                If m = 1 Then
'''''                    If topIdx > 0 Then
'''''                        idx = topIdx
'''''                    Else
'''''                        idx = botIdx
'''''                    End If
'''''                Else
'''''                    If botIdx > 0 Then
'''''                        idx = botIdx
'''''                    Else
'''''                        idx = topIdx
'''''                    End If
'''''                End If
'''''
'''''                With .SMP(m)
'''''                    .CRYNUM = SMP(idx).CRYNUM
'''''                    .IngotPos = SMP(idx).IngotPos
'''''                    .SMPKBN = SMP(idx).SMPKBN
'''''                    .hinban = SMP(idx).hinban
'''''                    .REVNUM = SMP(idx).REVNUM
'''''                    .factory = SMP(idx).factory
'''''                    .opecond = SMP(idx).opecond
'''''                    .CRYINDRS = SMP(idx).CRYINDRS
'''''                    .CRYRESRS = SMP(idx).CRYRESRS
'''''                    .CRYINDOI = SMP(idx).CRYINDOI
'''''                    .CRYRESOI = SMP(idx).CRYRESOI
'''''                    .CRYINDB1 = SMP(idx).CRYINDB1
'''''                    .CRYRESB1 = SMP(idx).CRYRESB1
'''''                    .CRYINDB2 = SMP(idx).CRYINDB2
'''''                    .CRYRESB2 = SMP(idx).CRYRESB2
'''''                    .CRYINDB3 = SMP(idx).CRYINDB3
'''''                    .CRYRESB3 = SMP(idx).CRYRESB3
'''''                    .CRYINDL1 = SMP(idx).CRYINDL1
'''''                    .CRYRESL1 = SMP(idx).CRYRESL1
'''''                    .CRYINDL2 = SMP(idx).CRYINDL2
'''''                    .CRYRESL2 = SMP(idx).CRYRESL2
'''''                    .CRYINDL3 = SMP(idx).CRYINDL3
'''''                    .CRYRESL3 = SMP(idx).CRYRESL3
'''''                    .CRYINDL4 = SMP(idx).CRYINDL4
'''''                    .CRYRESL4 = SMP(idx).CRYRESL4
'''''                    .CRYINDCS = SMP(idx).CRYINDCS
'''''                    .CRYRESCS = SMP(idx).CRYRESCS
'''''                    .CRYINDGD = SMP(idx).CRYINDGD
'''''                    .CRYRESGD = SMP(idx).CRYRESGD
'''''                    .CRYINDT = SMP(idx).CRYINDT
'''''                    .CRYREST = SMP(idx).CRYREST
'''''                    .CRYINDEP = SMP(idx).CRYINDEP
'''''                    .CRYRESEP = SMP(idx).CRYRESEP
'''''                End With
'''''            Next m
'''''
'''''#Else
'''''            sql = "select "
'''''            sql = sql & " XTALCS, "
'''''            sql = sql & " INPOSCS, "
'''''            sql = sql & " SMPKBNCS, "
'''''            sql = sql & " HINBCS, "
'''''            sql = sql & " REVNUMCS, "
'''''            sql = sql & " FACTORYCS, "
'''''            sql = sql & " OPECS, "
'''''            sql = sql & " CRYINDRSCS, "
'''''            sql = sql & " CRYRESRSCS, "
'''''            sql = sql & " CRYINDOICS, "
'''''            sql = sql & " CRYRESOICS, "
'''''            sql = sql & " CRYINDB1CS, "
'''''            sql = sql & " CRYRESB1CS, "
'''''            sql = sql & " CRYINDB2CS, "
'''''            sql = sql & " CRYRESB2CS, "
'''''            sql = sql & " CRYINDB3CS, "
'''''            sql = sql & " CRYRESB3CS, "
'''''            sql = sql & " CRYINDL1CS, "
'''''            sql = sql & " CRYRESL1CS, "
'''''            sql = sql & " CRYINDL2CS, "
'''''            sql = sql & " CRYRESL2CS, "
'''''            sql = sql & " CRYINDL3CS, "
'''''            sql = sql & " CRYRESL3CS, "
'''''            sql = sql & " CRYINDL4CS, "
'''''            sql = sql & " CRYRESL4CS, "
'''''            sql = sql & " CRYINDCSCS, "
'''''            sql = sql & " CRYRESCSCS, "
'''''            sql = sql & " CRYINDGDCS, "
'''''            sql = sql & " CRYRESGDCS, "
'''''            sql = sql & " CRYINDTCS, "
'''''            sql = sql & " CRYRESTCS, "
'''''            sql = sql & " CRYINDEPCS, "
'''''            sql = sql & " CRYRESEPCS "
'''''
'''''''            sql = sql & " from VECME010 V "
'''''''            sql = sql & " where E040CRYNUM = '" & .Crynum & "' "
'''''''            sql = sql & " and   E040INGOTPOS = '" & .IngotPos & "' "
'''''''            sql = sql & " order by E043INPOSCS"
'''''
'''''            sql = sql & " from XSDCS "
'''''            sql = sql & " where CRYNUMCS = '" & .BLOCKID & "' "
'''''            sql = sql & " order by INPOSCS"
'''''
'''''            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''            For m = 1 To 2
'''''                DoEvents
'''''                .SMP(m).CRYNUM = rs("E043XTALCS")
'''''                .SMP(m).IngotPos = rs("E043INPOSCS")
'''''                .SMP(m).SMPKBN = rs("E043SMPKBNCS")
'''''                .SMP(m).hinban = rs("E043HINBCS")
'''''                .SMP(m).REVNUM = rs("E043REVNUMCS")
'''''                .SMP(m).factory = rs("E043FACTORYCS")
'''''                .SMP(m).opecond = rs("E043OPECS")
'''''                .SMP(m).CRYINDRS = rs("E043CRYINDRSCS")
'''''                .SMP(m).CRYRESRS = rs("E043CRYRESRS1CS")
'''''                .SMP(m).CRYINDOI = rs("E043CRYINDOICS")
'''''                .SMP(m).CRYRESOI = rs("E043CRYRESOICS")
'''''                .SMP(m).CRYINDB1 = rs("E043CRYINDB1CS")
'''''                .SMP(m).CRYRESB1 = rs("E043CRYRESB1CS")
'''''                .SMP(m).CRYINDB2 = rs("E043CRYINDB2CS")
'''''                .SMP(m).CRYRESB2 = rs("E043CRYRESB2CS")
'''''                .SMP(m).CRYINDB3 = rs("E043CRYINDB3CS")
'''''                .SMP(m).CRYRESB3 = rs("E043CRYRESB3CS")
'''''                .SMP(m).CRYINDL1 = rs("E043CRYINDL1CS")
'''''                .SMP(m).CRYRESL1 = rs("E043CRYRESL1CS")
'''''                .SMP(m).CRYINDL2 = rs("E043CRYINDL2CS")
'''''                .SMP(m).CRYRESL2 = rs("E043CRYRESL2CS")
'''''                .SMP(m).CRYINDL3 = rs("E043CRYINDL3CS")
'''''                .SMP(m).CRYRESL3 = rs("E043CRYRESL3CS")
'''''                .SMP(m).CRYINDL4 = rs("E043CRYINDL4CS")
'''''                .SMP(m).CRYRESL4 = rs("E043CRYRESL4CS")
'''''                .SMP(m).CRYINDCS = rs("E043CRYINDCSCS")
'''''                .SMP(m).CRYRESCS = rs("E043CRYRESCSCS")
'''''                .SMP(m).CRYINDGD = rs("E043CRYINDGDCS")
'''''                .SMP(m).CRYRESGD = rs("E043CRYRESGDCS")
'''''                .SMP(m).CRYINDT = rs("E043CRYINDTCS")
'''''                .SMP(m).CRYREST = rs("E043CRYRESTCS")
'''''                .SMP(m).CRYINDEP = rs("E043CRYINDEPCS")
'''''                .SMP(m).CRYRESEP = rs("E043CRYRESEPCS")
'''''
'''''                rs.MoveNext
'''''            Next m
'''''            rs.Close
'''''            Set rs = Nothing
'''''#End If
'''''
''''''����������
''''''�i�Ԏd�l/Cs/EPD/LT�͂܂��u���b�N����SQL�𓊂��Ă���
''''''�������܂Ƃ߂Ă����΁A����5�b���x�k�ނ̂ł͂Ȃ����Ǝv����
''''''�������ACs/LT�ɂ��Ă͌��ʎ擾�̕��@���ς��̂ŁA���̌�̌������K�v
''''''������ɂ���A�Ώی����S�Ăɂ���Cs/LT/EPD�w���̂���T���v���𔲂��o���΂悢�͂�
'''''
'''''            ' �i�Ԃ̎d�l���擾
'''''            For m = 1 To 2
'''''                If Trim$(.SMP(m).hinban) = "G" Or Trim$(.SMP(m).hinban) = "Z" Then
'''''                    .SMP(m).HSXCNHWS = "S"
'''''                    .SMP(m).HSXLTHWS = "S"
'''''                    .SMP(m).EPD = "S"
'''''                ElseIf Len(Trim$(.SMP(m).hinban)) Then
'''''                    sql = " select "
'''''                    sql = sql & " S.HSXCNHWS,"
'''''                    sql = sql & " S.HSXLTHWS,"
'''''                    sql = sql & " 'H' as EPD "
'''''                    sql = sql & " from  TBCME019 S "
'''''                    sql = sql & " where S.HINBAN   = '" & .SMP(m).hinban & "'"
'''''                    sql = sql & "   and S.MNOREVNO =  " & .SMP(m).REVNUM
'''''                    sql = sql & "   and S.FACTORY  = '" & .SMP(m).factory & "'"
'''''                    sql = sql & "   and S.OPECOND  = '" & .SMP(m).opecond & "'"
'''''
'''''                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                    .SMP(m).HSXCNHWS = rs("HSXCNHWS")
'''''                    .SMP(m).HSXLTHWS = rs("HSXLTHWS")
'''''                    .SMP(m).EPD = rs("EPD")
'''''
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                Else
'''''                    '��i�Ԃ̏ꍇ
'''''                    .SMP(m).HSXCNHWS = " "
'''''                    .SMP(m).HSXLTHWS = " "
'''''                    .SMP(m).EPD = " "
'''''                End If
'''''            Next m
'''''
'''''            ' �`�F�b�N
'''''            For m = 1 To 2
'''''                DoEvents
'''''                ' CS�̃`�F�b�N
''''''                If (.SMP(m).HSXCNHWS = "H" Or .SMP(m).HSXCNHWS = "S") And .SMP(m).CRYINDCS = "0" Then  ' �Q�l�]���͂Ȃ��Ă��n�j
'''''                If .SMP(m).HSXCNHWS = "H" And .SMP(m).CRYINDCS = "0" Then
'''''
''''''                    sql = "select CRYRESCS as RES from TBCME043 "
''''''                    sql = sql & "where CRYNUM = '" & .SMP(m).CRYNUM & "' "
''''''                    sql = sql & "  and INGOTPOS >= " & .SMP(m).INGOTPOS
''''''                    sql = sql & "  and CRYINDCS<>'0'"
''''''                    sql = sql & " order by INGOTPOS"
'''''                    sql = "select CRYRESCSCS as RES from XSDCS "
'''''                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
'''''                    sql = sql & "  and CRYINDCSCS<>'0'"
'''''                    sql = sql & " order by INPOSCS"
'''''
'''''                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                    If rs.RecordCount Then
'''''                        If rs("RES") = "0" Then LWD(l).flag = True
'''''                    End If
'''''
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''
'''''                ' LT�̃`�F�b�N
''''''                If (.SMP(m).HSXLTHWS = "H" Or .SMP(m).HSXLTHWS = "S") And .SMP(m).CRYINDT = "0" And LWD(l).flag = False Then ' �Q�l�]���͂Ȃ��Ă��n�j
'''''                If .SMP(m).HSXLTHWS = "H" And .SMP(m).CRYINDT = "0" And LWD(l).flag = False Then
'''''
''''''                    sql = "select CRYREST as RES from TBCME043 "
''''''                    sql = sql & "where CRYNUM = '" & .SMP(m).CRYNUM & "' "
''''''                    sql = sql & "  and INGOTPOS >= " & .SMP(m).INGOTPOS
''''''                    sql = sql & "  and CRYINDT<>'0'"
''''''                    sql = sql & " order by INGOTPOS"
'''''
'''''                    sql = "select CRYRESTCS as RES from XSDCS "
'''''                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
'''''                    sql = sql & "  and CRYINDTCS<>'0'"
'''''                    sql = sql & " order by INPOSCS"
'''''
'''''                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                    If rs.RecordCount Then
'''''                        If rs("RES") = "0" Then LWD(l).flag = True
'''''                    End If
'''''
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''
'''''                ' EPD�̃`�F�b�N
''''''                If (.SMP(m).EPD = "H" Or .SMP(m).EPD = "S") And .SMP(m).CRYINDEP = "0" And LWD(l).flag = False Then ' S�͂��肦�Ȃ��Ǔ���
'''''                If .SMP(m).EPD = "H" And .SMP(m).CRYINDEP = "0" And LWD(l).flag = False Then ' S�͂��肦�Ȃ��Ǔ���
'''''
''''''                    sql = "select CRYRESEP as RES from TBCME043 "
''''''                    sql = sql & "where CRYNUM = '" & .SMP(m).CRYNUM & "' "
''''''                    sql = sql & "  and INGOTPOS >= " & .SMP(m).INGOTPOS
''''''                    sql = sql & "  and CRYINDEP<>'0'"
''''''                    sql = sql & " order by INGOTPOS"
'''''
'''''                    sql = "select CRYRESEPCS as RES from XSDCS "
'''''                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
'''''                    sql = sql & "  and CRYINDEPCS<>'0'"
'''''                    sql = sql & " order by INPOSCS"
'''''
'''''                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                    If rs.RecordCount Then
'''''                        If rs("RES") = "0" Then LWD(l).flag = True
'''''                    End If
'''''
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
''''''                If LWD(l).flag = True Then
''''''                    Exit For
''''''                End If
'''''            Next m
'''''        End If
'''''
'''''        End With    ' .Wd1()
'''''
'''''    Next l
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''
'''''
'''''    gErr.HandleError
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    cmkc001b_DBDataCheck1 = FUNCTION_RETURN_FAILURE
'''''    Resume proc_exit
'''''End Function


'''''Public Function cmkc001b_DBDataCheck3(LWD() As cmkc001b_LockWait, _
'''''                                 Wd3() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''    Dim c0 As Integer
'''''    Dim c1 As Integer
'''''    Dim c2 As Integer
'''''    Dim MaxRec As Integer
'''''    Dim RecCount As Integer
'''''    Dim EQFlag As Boolean
'''''    Dim sql As String       'SQL�S��
'''''    Dim rs As OraDynaset    'RecordSet
'''''    Dim GrpCount1 As Integer
'''''    Dim GrpCount2 As Integer
'''''    Dim ColorFlag As Boolean
'''''    Dim TotalBlk As Integer
'''''    Dim CheckPoint As Integer
'''''    Dim CheckEnd As Integer
'''''    Dim tempGrpFlag As String * 1
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp"
'''''
'''''    cmkc001b_DBDataCheck3 = FUNCTION_RETURN_SUCCESS
'''''    TotalBlk = UBound(Wd3())
'''''
'''''Debug.Print " 1:" & Time
'''''
'''''    'CC700�̃u���b�N�̌����ꗗ�����
'''''    ReDim GrpInfo(1) As cmkc001b_Wait3
'''''    GrpInfo(1).CRYNUM = vbNullString
'''''    c1 = 0
'''''    For c0 = 1 To TotalBlk
'''''        DoEvents
'''''        If c1 = 0 Then
'''''            GrpInfo(1).CRYNUM = Wd3(c0).CRYNUM
'''''        End If
'''''        MaxRec = UBound(GrpInfo())
'''''        EQFlag = False
'''''        c1 = 1
'''''        Do While c1 <= MaxRec
'''''            DoEvents
'''''            If Wd3(c0).CRYNUM = GrpInfo(c1).CRYNUM Then
'''''                EQFlag = True
'''''                Exit Do
'''''            End If
'''''            c1 = c1 + 1
'''''        Loop
'''''        If Not EQFlag Then
'''''            ReDim Preserve GrpInfo(MaxRec + 1) As cmkc001b_Wait3
'''''            GrpInfo(MaxRec + 1).CRYNUM = Wd3(c0).CRYNUM
'''''        End If
'''''    Next
'''''Debug.Print " 2:" & Time
'''''
'''''    '�����Ɋ܂܂��S�Ẵu���b�N�����߂�
'''''    MaxRec = UBound(GrpInfo())
'''''    For c0 = 1 To MaxRec
'''''        sql = "select "
'''''        sql = sql & "BLOCKID, "
'''''        sql = sql & "INGOTPOS, "
'''''        sql = sql & "LENGTH, "
'''''        sql = sql & "NOWPROC, "
'''''        sql = sql & "HOLDCLS "
'''''        sql = sql & "from TBCME040 "
'''''        sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
''''''2001/11/14 S.Sano        sql = sql & "and LSTATCLS='T' "
''''''2001/11/14 S.Sano        sql = sql & "and RSTATCLS='T' "
''''''2001/11/14 S.Sano        sql = sql & "and DELCLS='0' "
'''''        'sql = sql & "and HOLDCLS='0' "
'''''        sql = sql & "order by BLOCKID "
'''''
'''''
'''''        '�f�[�^�𒊏o����
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''        RecCount = rs.RecordCount
'''''        If RecCount = 0 Then
'''''            rs.Close
'''''            GoTo proc_exit
'''''        End If
'''''        ReDim GrpInfo(c0).blkInfo(RecCount) As cmkc001b_Wait3_BLK
'''''        For c1 = 1 To RecCount
'''''            GrpInfo(c0).blkInfo(c1).BLOCKID = rs("BLOCKID")
'''''            GrpInfo(c0).blkInfo(c1).IngotPos = rs("INGOTPOS")
'''''            GrpInfo(c0).blkInfo(c1).LENGTH = rs("LENGTH")
'''''            GrpInfo(c0).blkInfo(c1).NOWPROC = rs("NOWPROC")
'''''            GrpInfo(c0).blkInfo(c1).HOLDCLS = rs("HOLDCLS")
'''''            rs.MoveNext
'''''        Next
'''''        rs.Close
'''''    Next
'''''
'''''Debug.Print " 3:" & Time
'''''    '�u���b�N�̏㉺�i�Ԃ����߂�
'''''#If SPEEDUP Then   '���������� 02.1.28-2.15 �쑺
''''''����������
''''''�u���b�N�̏㉺�i�Ԃ����߂邾���Ȃ�A1���SQL�ł܂Ƃ߂ď����擾�ł���͂�
'''''Dim BLKID() As String
'''''Dim topHin() As tFullHinban
'''''Dim botHin() As tFullHinban
'''''Dim idx As Integer
'''''Dim rsCount As Integer
'''''Dim found As Boolean
'''''
'''''    sql = vbNullString
'''''    sql = sql & "select"
'''''    sql = sql & "  b.BLOCKID"
'''''    sql = sql & ", TOP.HINBAN as THINBAN, TOP.REVNUM as TREVNUM, TOP.FACTORY as TFACTORY, TOP.OPECOND as TOPECOND"
'''''    sql = sql & ", BOT.HINBAN as BHINBAN, BOT.REVNUM as BREVNUM, BOT.FACTORY as BFACTORY, BOT.OPECOND as BOPECOND "
'''''    sql = sql & "from TBCME040 B, TBCME041 TOP, TBCME041 BOT "
'''''    sql = sql & "Where b.CRYNUM = Top.CRYNUM"
'''''    sql = sql & "  and B.CRYNUM=BOT.CRYNUM"
'''''    sql = sql & "  and B.INGOTPOS>=0"
'''''    sql = sql & "  and B.DELCLS='0'"
'''''    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
'''''    sql = sql & "  and B.RSTATCLS='T'"
'''''    sql = sql & "  and B.HOLDCLS='0'"
'''''    sql = sql & "  and B.INGOTPOS>=TOP.INGOTPOS"
'''''    sql = sql & "  and B.INGOTPOS<TOP.INGOTPOS+TOP.LENGTH"
'''''    sql = sql & "  and B.INGOTPOS+B.LENGTH>BOT.INGOTPOS"
'''''    sql = sql & "  and B.INGOTPOS+B.LENGTH<=BOT.INGOTPOS+BOT.LENGTH "
'''''    sql = sql & "order by B.BLOCKID"
'''''
'''''    '�f�[�^�𒊏o����
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    rsCount = rs.RecordCount
'''''    ReDim BLKID(1 To rsCount)
'''''    ReDim topHin(1 To rsCount)
'''''    ReDim botHin(1 To rsCount)
'''''    For c0 = 1 To rsCount
'''''        BLKID(c0) = rs!BLOCKID
'''''        topHin(c0).hinban = rs!THINBAN
'''''        topHin(c0).mnorevno = rs!TREVNUM
'''''        topHin(c0).factory = rs!TFACTORY
'''''        topHin(c0).opecond = rs!TOPECOND
'''''        botHin(c0).hinban = rs!BHINBAN
'''''        botHin(c0).mnorevno = rs!BREVNUM
'''''        botHin(c0).factory = rs!BFACTORY
'''''        botHin(c0).opecond = rs!BOPECOND
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    For c0 = 1 To MaxRec
'''''        RecCount = UBound(GrpInfo(c0).blkInfo())
'''''        For c1 = 1 To RecCount
'''''            found = False
'''''            For idx = 1 To rsCount
'''''                If BLKID(idx) = GrpInfo(c0).blkInfo(c1).BLOCKID Then
'''''                    found = True
'''''                    Exit For
'''''                ElseIf BLKID(idx) > GrpInfo(c0).blkInfo(c1).BLOCKID Then
'''''                    Exit For
'''''                End If
'''''            Next
'''''
'''''            If found Then
'''''                GrpInfo(c0).blkInfo(c1).topHin.hinban = topHin(idx).hinban
'''''                GrpInfo(c0).blkInfo(c1).topHin.factory = topHin(idx).factory
'''''                GrpInfo(c0).blkInfo(c1).topHin.opecond = topHin(idx).opecond
'''''                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = topHin(idx).mnorevno
'''''            Else
'''''                GrpInfo(c0).blkInfo(c1).topHin.hinban = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.factory = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.opecond = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = 0
'''''            End If
'''''
'''''            If found Then
'''''                GrpInfo(c0).blkInfo(c1).botHin.hinban = botHin(idx).hinban
'''''                GrpInfo(c0).blkInfo(c1).botHin.factory = botHin(idx).factory
'''''                GrpInfo(c0).blkInfo(c1).botHin.opecond = botHin(idx).opecond
'''''                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = botHin(idx).mnorevno
'''''            Else
'''''                GrpInfo(c0).blkInfo(c1).botHin.hinban = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.factory = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.opecond = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = 0
'''''            End If
'''''        Next
'''''    Next
'''''#Else
'''''    For c0 = 1 To MaxRec
'''''        RecCount = UBound(GrpInfo(c0).blkInfo())
'''''        For c1 = 1 To RecCount
'''''            sql = "select "
'''''            sql = sql & "HINBAN, "
'''''            sql = sql & "REVNUM, "
'''''            sql = sql & "FACTORY, "
'''''            sql = sql & "OPECOND "
'''''            sql = sql & "from TBCME041 "
'''''            sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
''''''2001/11/14 S.Sano            sql = sql & "and INGOTPOS <= " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
'''''            sql = sql & "and INGOTPOS = " & GrpInfo(c0).blkInfo(c1).IngotPos & " " '2001/11/14 S.Sano
''''''2001/11/14 S.Sano            sql = sql & "and (INGOTPOS + LENGTH) > " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
'''''
'''''            '�f�[�^�𒊏o����
'''''            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''            RecCount = rs.RecordCount
'''''            If RecCount = 0 Then
'''''                GrpInfo(c0).blkInfo(c1).topHin.hinban = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.factory = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.opecond = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = 0
'''''            Else
'''''                GrpInfo(c0).blkInfo(c1).topHin.hinban = rs("HINBAN")
'''''                GrpInfo(c0).blkInfo(c1).topHin.factory = rs("FACTORY")
'''''                GrpInfo(c0).blkInfo(c1).topHin.opecond = rs("OPECOND")
'''''                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = rs("REVNUM")
'''''            End If
'''''            rs.Close
'''''
'''''            sql = "select "
'''''            sql = sql & "HINBAN, "
'''''            sql = sql & "REVNUM, "
'''''            sql = sql & "FACTORY, "
'''''            sql = sql & "OPECOND "
'''''            sql = sql & "from TBCME041 "
'''''            sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
'''''            sql = sql & "and INGOTPOS < " & GrpInfo(c0).blkInfo(c1).IngotPos + GrpInfo(c0).blkInfo(c1).LENGTH & " "
'''''            sql = sql & "and (INGOTPOS + LENGTH) >= " & GrpInfo(c0).blkInfo(c1).IngotPos + GrpInfo(c0).blkInfo(c1).LENGTH & " "
'''''
'''''            '�f�[�^�𒊏o����
'''''            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''            RecCount = rs.RecordCount
'''''            If RecCount = 0 Then
'''''                GrpInfo(c0).blkInfo(c1).botHin.hinban = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.factory = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.opecond = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = 0
'''''            Else
'''''                GrpInfo(c0).blkInfo(c1).botHin.hinban = rs("HINBAN")
'''''                GrpInfo(c0).blkInfo(c1).botHin.factory = rs("FACTORY")
'''''                GrpInfo(c0).blkInfo(c1).botHin.opecond = rs("OPECOND")
'''''                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = rs("REVNUM")
'''''            End If
'''''            rs.Close
'''''        Next
'''''    Next
'''''#End If
'''''
'''''Debug.Print " 4:" & Time
'''''    '���߂���񂩂�O���[�v�����߂�
'''''    GrpCount1 = 0
'''''    GrpCount2 = 0
'''''    For c0 = 1 To MaxRec
'''''        GrpCount1 = GrpCount1 + 1
'''''        GrpCount2 = GrpCount2 + 1
'''''        RecCount = UBound(GrpInfo(c0).blkInfo())
'''''        For c1 = 1 To RecCount
'''''            '�u���b�N�؂�ڂŕi�Ԃ��ς��ΕʃO���[�v�Ɣ��f����
'''''            Select Case c1
'''''            Case 1
'''''                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
'''''            Case Else
'''''                If (GrpInfo(c0).blkInfo(c1).topHin.factory <> GrpInfo(c0).blkInfo(c1 - 1).botHin.factory) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).topHin.hinban <> GrpInfo(c0).blkInfo(c1 - 1).botHin.hinban) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).topHin.opecond <> GrpInfo(c0).blkInfo(c1 - 1).botHin.opecond) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).topHin.REVNUM <> GrpInfo(c0).blkInfo(c1 - 1).botHin.REVNUM) Then
'''''                    GrpCount1 = GrpCount1 + 1
'''''                End If
'''''                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
'''''            End Select
'''''
'''''            '����O���[�v���ŁA�H���Ⴂ�̃u���b�N�����݂����ꍇ�A����O���[�v����
'''''            '���O���[�v�Ƃ��ăO���[�v��������B
'''''            'CC710�ȊO�Ȃ�ΏۊO�Ƃ��O���[�v��������Ȃ�
'''''            If GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_NUKISI_SIJI And GrpInfo(c0).blkInfo(c1).HOLDCLS = "0" Then
'''''                Select Case c1
'''''                Case 1
'''''                    GrpInfo(c0).blkInfo(c1).GRPFLG2 = GrpCount2
'''''                Case Else
'''''                    If (GrpInfo(c0).blkInfo(c1).topHin.factory <> GrpInfo(c0).blkInfo(c1 - 1).botHin.factory) Or _
'''''                       (GrpInfo(c0).blkInfo(c1).topHin.hinban <> GrpInfo(c0).blkInfo(c1 - 1).botHin.hinban) Or _
'''''                       (GrpInfo(c0).blkInfo(c1).topHin.opecond <> GrpInfo(c0).blkInfo(c1 - 1).botHin.opecond) Or _
'''''                       (GrpInfo(c0).blkInfo(c1).topHin.REVNUM <> GrpInfo(c0).blkInfo(c1 - 1).botHin.REVNUM) Then
'''''                        GrpCount2 = GrpCount2 + 1
'''''                    End If
'''''                    GrpInfo(c0).blkInfo(c1).GRPFLG2 = GrpCount2
'''''                End Select
'''''            Else
'''''                GrpCount2 = GrpCount2 + 1
'''''                GrpInfo(c0).blkInfo(c1).GRPFLG2 = 0
'''''            End If
'''''        Next
'''''    Next
'''''Debug.Print " 5:" & Time
'''''    '���߂���񂩂�\���F�����߂�
'''''    For c0 = 1 To MaxRec
'''''        RecCount = UBound(GrpInfo(c0).blkInfo())
'''''        ColorFlag = False
'''''        CheckPoint = 0
'''''        For c1 = 1 To RecCount
'''''            If CheckPoint > 0 Then
'''''                If GrpInfo(c0).blkInfo(c1).GRPFLG1 <> GrpInfo(c0).blkInfo(CheckPoint).GRPFLG1 Then
'''''                    For c2 = CheckPoint To c1 - 1
'''''                        GrpInfo(c0).blkInfo(c2).COLORFLG = ColorFlag
'''''                    Next
'''''                    ColorFlag = False
'''''                    CheckPoint = c1
'''''                End If
'''''            Else
'''''                CheckPoint = c1
'''''            End If
'''''            If CheckPoint > 0 Then
'''''                If (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_SETUDAN) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_KESSYOU_SOUGOUHANTEI) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_KESSYOU_SAISYUU_HARAIDASI) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).HOLDCLS = "1") Then
'''''                    ColorFlag = True
'''''                End If
'''''            End If
'''''        Next
'''''        For c1 = CheckPoint To RecCount
'''''            GrpInfo(c0).blkInfo(c1).COLORFLG = ColorFlag
'''''        Next
'''''    Next
'''''Debug.Print " 6:" & Time
'''''    For c0 = 1 To MaxRec
'''''        RecCount = UBound(GrpInfo(c0).blkInfo())
'''''        For c1 = 1 To RecCount
'''''            For c2 = 1 To TotalBlk
'''''                If Wd3(c2).BLOCKID = GrpInfo(c0).blkInfo(c1).BLOCKID Then
'''''                    LWD(c2).flag = GrpInfo(c0).blkInfo(c1).COLORFLG
'''''                    LWD(c2).Grp = GrpInfo(c0).blkInfo(c1).GRPFLG2
'''''                    Exit For
'''''                End If
'''''            Next
'''''        Next
'''''    Next
''''''    Debug.Print Now
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    cmkc001b_DBDataCheck3 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function




'��������

'12���i�Ԃ�ommonDefine.BAS �Œ�`����Ă��� type tFullHinban ���g�p���܂���
'Public Type tFullHinban
'    hinban As String * 8            ' �i��
'    MNOREVNO As Integer             ' ���i�ԍ������ԍ�
'    FACTORY As String * 1           ' �H��
'    OPECOND As String * 1           ' ���Ə���
'End Type


'
'-*-*- �l���@20010807 �T����A�U���p���AEPD,Cs,LT�@�C��
' ������t���[��
' �d�l�ۏؕ��@�Q�� --+--�Ȃ�(H�ȊO) --���сi�Y���ʒu�j--�����Ă��Ȃ��Ă�����OK
'�@�@�@�@�@�@�@�@�@�@|
'                   +--����(H) --���сi�Y���ʒu) --+--���� -- ����`�F�b�N --+-- OK
'                                                 |                        |
'                                                 |                        +-- MG
'                                                 |
'                                                 +--�Ȃ� --+-- �����w���T�E�U�ȊO�̏ꍇ --+--EPD�ACs�ALT�̏ꍇ����T�� --+-- �Ȃ� -- NG
'                                                           |                            |                          �@ |
'                                                           |                            +--EPD,Cs�ALT�ȊO -- NG       +-- ���� -- ����`�F�b�N --+-- OK
'                                                           |                                                                                    |
'                                                           |                                                                                    +-- NG
'                                                           |
'                                                           +-- �����w���T�̏ꍇ (Rs, Cs) �Ȃ琄��Ȃ̂őS�̂�����т�T�� --+-- ����`�F�b�N --+-- OK
'                                                           |  �@�@�@�@�@�@�@�@�@                                          |                 |
'                                                           |                                                             |                 +-- NG
'                                                           +-- �����w���U�̏ꍇ TOP�Ȃ��ցATAIL�Ȃ牺�֎��т�T��       --+
'
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�i�����w���́A�w���𗧂Ă鑤������ɗ��ĂĂ���ƍl���Ă���j



''''''�T�v      :�����֐��@SQL�ɐ���ł̌���������t������
'''''Private Sub AddSQL_SUITEI(sql As String, Cry As String, Table As String, TorB As Integer, Optional subSQL As String = "")
'''''    sql = sql & " from " & Table & " T1 "
'''''    sql = sql & " where T1.CRYNUM='" & Cry & "' "
'''''    sql = sql & " and   T1.TRANCNT=ANY( select max(TRANCNT) from " & Table & " T2 "
'''''    sql = sql & "                       where T2.CRYNUM=T1.CRYNUM  and T2.POSITION=T1.POSITION and T2.SMPKBN=T1.SMPKBN " & subSQL & ") "
'''''    sql = sql & " and T1.SMPLUMU = '0' "
'''''    sql = sql & subSQL
'''''    If TorB = 1 Then
'''''        sql = sql & " order by T1.POSITION asc "     ' �P��ڂ͍ŏ�����
'''''    Else
'''''        sql = sql & " order by T1.POSITION desc "    ' �Q��ڂ͌�납��
'''''    End If
'''''End Sub
'''''
'''''
''''''�T�v      :�����֐��@SQL�Ɉ��p���ł̌���������t������
'''''Private Sub AddSQL_HIKITUGI(sql As String, Cry As String, pos As Integer, Table As String, TorB As Integer, Optional subSQL = "")
'''''    sql = sql & " from " & Table & " T1 "
'''''    sql = sql & " where T1.CRYNUM='" & Cry & "' "
'''''    sql = sql & " and   T1.TRANCNT=ANY( select max(TRANCNT) from " & Table & " T2 "
'''''    sql = sql & "                       where T2.CRYNUM=T1.CRYNUM  and T2.POSITION=T1.POSITION and T2.SMPKBN=T1.SMPKBN " & subSQL & ") "
'''''    sql = sql & " and T1.SMPLUMU = '0' "
'''''    sql = sql & subSQL
'''''    If TorB = 1 Then                        ' TOP���͏�ɒT��
'''''        sql = sql & " and T1.POSITION < " & CStr(pos)
'''''        sql = sql & " order by T1.POSITION asc, SMPKBN asc "
'''''    Else                                    ' BOT���͉��ɒT��
'''''        sql = sql & " and T1.POSITION > " & CStr(pos)
'''''        sql = sql & " order by T1.POSITION desc, SMPKBN desc "
'''''    End If
'''''End Sub
'''''
''''''�T�v      :�����֐��@SQL�Ɉ��p���ł̌���������t������
'''''Private Sub AddSQL_HIKITUGI2(sql As String, Cry As String, pos As Integer, Table As String, TorB As Integer, Optional subSQL = "")
'''''    sql = sql & " from " & Table & " T1 "
'''''    sql = sql & " where T1.CRYNUM='" & Cry & "' "
'''''    sql = sql & " and   T1.TRANCNT=ANY( select max(TRANCNT) from " & Table & " T2 "
'''''    sql = sql & "                       where T2.CRYNUM=T1.CRYNUM  and T2.POSITION=T1.POSITION and T2.SMPKBN=T1.SMPKBN " & subSQL & ") "
'''''    sql = sql & " and T1.SMPLUMU = '0' "
'''''    sql = sql & subSQL
'''''    If TorB = 1 Then                        ' TOP���͏�ɒT��
'''''        sql = sql & " and T1.POSITION < " & CStr(pos)
'''''        sql = sql & " order by T1.POSITION desc, SMPKBN desc "
'''''    Else                                    ' BOT���͉��ɒT��
'''''        sql = sql & " and T1.POSITION > " & CStr(pos)
'''''        sql = sql & " order by T1.POSITION asc, SMPKBN asc "
'''''    End If
'''''End Sub
'''''
''''''�T�v      :�����֐��@SQL�ɉ��Ɏ��т��������錟��������t������
'''''Private Sub AddSQL_Down(sql As String, Cry As String, pos As Integer, Table As String, Optional subSQL = "")
'''''    sql = sql & " from " & Table & " T1 "
'''''    sql = sql & " where T1.CRYNUM='" & Cry & "' "
'''''    sql = sql & " and   T1.TRANCNT=ANY( select max(TRANCNT) from " & Table & " T2 "
'''''    sql = sql & "                       where T2.CRYNUM=T1.CRYNUM  and T2.POSITION=T1.POSITION and T2.SMPKBN=T1.SMPKBN  " & subSQL & ") "
'''''    sql = sql & " and T1.SMPLUMU = '0' "
'''''    sql = sql & " and T1.POSITION > " & CStr(pos)
'''''    sql = sql & subSQL
'''''    sql = sql & " order by POSITION asc, SMPKBN asc "
'''''End Sub


'''''Private Sub AddSQL_Default(sql As String, Cry As String, pos As Integer, Spk As String, Table As String, Optional subSQL = "")
'''''    sql = sql & " from " & Table
'''''    sql = sql & " where CRYNUM='" & Cry & "' " & _
'''''                " and POSITION=" & pos & _
'''''                " and SMPKBN='" & Spk & "' " & _
'''''                subSQL & _
'''''                " and TRANCNT=ANY( select max(TRANCNT) from " & Table & _
'''''                                   " where CRYNUM='" & Cry & "' " & _
'''''                                   " and POSITION=" & pos & _
'''''                                   " and SMPKBN='" & Spk & "' " & _
'''''                                   subSQL & " ) "
'''''End Sub
'''''
'''''Private Sub AddSQL_Default2(sql As String, Cry As String, pos As Integer, Spk As String, Table As String, Optional subSQL = "")
'''''    sql = sql & " from " & Table
'''''    sql = sql & " where CRYNUM='" & Cry & "' " & _
'''''                " and POSITION=" & pos & _
'''''                subSQL & _
'''''                " and TRANCNT=ANY( select max(TRANCNT) from " & Table & _
'''''                                   " B where B.CRYNUM='" & Cry & "' " & _
'''''                                   " and B.POSITION=" & pos & _
'''''                                   " and B.SMPKBN=SMPKBN " & _
'''''                                   subSQL & " ) "
'''''    If (Spk = "T") Then
'''''        sql = sql & "order by SMPKBN desc "
'''''    Else
'''''        sql = sql & "order by SMPKBN "
'''''    End If
'''''End Sub


''''''�T�v      :�����֐� ������R���ю擾�p�I�u�W�F�N�g�R�s�[�֐�(���R�[�h�Z�b�g����̃R�s�[)
'''''Private Sub CryR_ObjCpy(CryR As type_DBDRV_scmzc_fcmkc001c_CryR, rs As OraDynaset)
'''''    With CryR
'''''        .CRYNUM = rs("CRYNUM")         ' �����ԍ�
'''''        .POSITION = rs("POSITION")     ' �ʒu
'''''        .SMPKBN = rs("SMPKBN")         ' �T���v���敪
'''''        .SMPLNO = rs("SMPLNO")         ' �T���v���m��
'''''        .SMPLUMU = rs("SMPLUMU")       ' �T���v���L��
'''''        .TRANCOND = rs("TRANCOND")     ' ��������
'''''        .MEAS1 = rs("MEAS1")           ' ����l�P
'''''        .MEAS2 = rs("MEAS2")           ' ����l�Q
'''''        .MEAS3 = rs("MEAS3")           ' ����l�R
'''''        .MEAS4 = rs("MEAS4")           ' ����l4
'''''        .MEAS5 = rs("MEAS5")           ' ����l�T
'''''        .RRG = rs("RRG")               ' RRG
'''''        .REGDATE = rs("REGDATE")       ' �o�^���t
'''''    End With
'''''End Sub
'''''
''''''�T�v      :�����֐� ������R���ю擾�p�x�[�XSQL
'''''Private Sub CryR_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "        ' �����ԍ�
'''''    sql = sql & "POSITION, "      ' �ʒu
'''''    sql = sql & "SMPKBN, "        ' �T���v���敪
'''''    sql = sql & "SMPLNO, "        ' �T���v���m��
'''''    sql = sql & "SMPLUMU, "       ' �T���v���L��
'''''    sql = sql & "TRANCOND, "      ' ��������
'''''    sql = sql & "MEAS1, "         ' ����l�P
'''''    sql = sql & "MEAS2, "         ' ����l�Q
'''''    sql = sql & "MEAS3, "         ' ����l�R
'''''    sql = sql & "MEAS4, "         ' ����l�S
'''''    sql = sql & "MEAS5, "         ' ����l�T
'''''    sql = sql & "RRG, "            ' RRG
'''''    sql = sql & "REGDATE "         '�@�o�^���t
'''''End Sub


''''''�T�v      :�����֐� ������R���ю擾�p
'''''Private Function CryR_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              CryR As type_DBDRV_scmzc_fcmkc001c_CryR, _
'''''                              SuCryR As type_DBDRV_scmzc_fcmkc001c_CryR, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim i As Long
'''''    Dim recCnt As Integer
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''    ' ������R���уe�[�u������l���擾
'''''    Dim Tname As String
'''''    Tname = "TBCMJ002"
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function CryR_Zisseki"
'''''
'''''    CryR_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Set rs = Nothing
'''''
'''''    ' �����w���ɂ�����炸���т��擾����@�i�i�Ԃ�������U��ւ����E�蓮�Ō����w���𗧂Ă��@�Ȃǂ̂��߁j
'''''    DoEvents
'''''    Call CryR_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname)
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then          ' ���т���������A������̗p����
'''''        DoEvents
'''''        Call CryR_ObjCpy(CryR, rs)
'''''        SuCryR = CryR   '2001/10/24 S.Sano�@����ł����т����݂����ꍇ�̏���
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
'''''        rs.Close
'''''        Set rs = Nothing
'''''
''''''----- 8/12 �쑺 �C���i���R�͎d�l�Ɋւ�炸���ʂ�\���������j
''''''��        If Siyou.HSXRHWYS = SIJI Then   ' �d�l�̎w���������Ă���
''''''        If Siyou.HSXRHWYS = SIJI Then   ' �d�l�̎w���������Ă���
''''''
'''''            If Samp.CRYINDRS = "5" Then       ' ����Ȃ�          ' �{���Ȃ�TOP�^BOT�łP��ł����͂� ---�P��ɂ���(���{)
''''''                For i = 1 To 2
'''''                DoEvents
'''''                Call CryR_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_SUITEI(sql, Samp.CRYNUM, Tname, TorB)
'''''                DoEvents
'''''
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''                If rs.RecordCount <> 0 Then
'''''                    DoEvents
'''''                    Call CryR_ObjCpy(SuCryR, rs)
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                rs.Close
'''''                Set rs = Nothing
''''''                Next i
'''''            ElseIf Samp.CRYINDRS = "6" Then       ' ���p���Ȃ�
'''''                DoEvents
'''''                Call CryR_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_HIKITUGI2(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname, TorB)
'''''                DoEvents
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''                If rs.RecordCount <> 0 Then
'''''                    recCnt = rs.RecordCount
'''''                    For i = 1 To recCnt
'''''                        If i = 1 Then                                     ' ���ڂ͕ێ�
'''''                            DoEvents
'''''                            Call CryR_ObjCpy(CryR, rs)
'''''                            Exit For    '�P���R�[�h�ڂ�����OK
'''''                        Else
'''''                            If CryR.POSITION = rs("POSITION") And CryR.REGDATE < rs("REGDATE") Then   ' �O�̈ʒu�Ɠ�����������o�^���t���V�������̂��Ƃ�
'''''                                DoEvents
'''''                                Call CryR_ObjCpy(CryR, rs)
'''''                            End If
'''''                            Exit For
'''''                        End If
'''''                        rs.MoveNext
'''''                    Next
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                If Not rs Is Nothing Then
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''            End If  ' �����w����5 or 6 �Ȃ�
''''''----- 8/12 �쑺 �C���i���R�͎d�l�Ɋւ�炸���ʂ�\���������j
''''''��        End If  ' �w���������Ă���
''''''        End If  ' �w���������Ă���
''''''-----
'''''    End If ' ���т�����
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    CryR_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''�T�v      :�����֐� Oi���ю擾�p�I�u�W�F�N�g�R�s�[�֐�(���R�[�h�Z�b�g����̃R�s�[)
'''''Private Sub Oi_ObjCpy(Oi As type_DBDRV_scmzc_fcmkc001c_Oi, rs As OraDynaset)
'''''    With Oi
'''''        .CRYNUM = rs("CRYNUM")         ' �����ԍ�
'''''        .POSITION = rs("POSITION")     ' �ʒu
'''''        .SMPKBN = rs("SMPKBN")         ' �T���v���敪
'''''        .SMPLNO = rs("SMPLNO")         ' �T���v���m��
'''''        .SMPLUMU = rs("SMPLUMU")       ' �T���v���L��
'''''        .TRANCOND = rs("TRANCOND")     ' ��������
'''''        .OIMEAS1 = rs("OIMEAS1")       ' �n������l�P
'''''        .OIMEAS2 = rs("OIMEAS2")       ' �n������l�Q
'''''        .OIMEAS3 = rs("OIMEAS3")       ' �n������l�R
'''''        .OIMEAS4 = rs("OIMEAS4")       ' �n������l�S
'''''        .OIMEAS5 = rs("OIMEAS5")       ' �n������l�T
'''''        .ORGRES = rs("ORGRES")         ' �n�q�f����
'''''        .AVE = rs("AVE")               ' �`�u�d
'''''        .FTIRCONV = rs("FTIRCONV")     ' �e�s�h�q���Z
'''''        .INSPECTWAY = rs("INSPECTWAY") ' �������@
'''''        .REGDATE = rs("REGDATE")       ' �o�^���t
'''''    End With
'''''End Sub
'''''
''''''�T�v      :�����֐� Oi���ю擾�p�x�[�XSQL
'''''Private Sub Oi_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "          ' �����ԍ�
'''''    sql = sql & "POSITION, "        ' �ʒu
'''''    sql = sql & "SMPKBN, "          ' �T���v���敪
'''''    sql = sql & "SMPLNO, "          ' �T���v���m��
'''''    sql = sql & "SMPLUMU, "         ' �T���v���L��
'''''    sql = sql & "TRANCOND, "        ' ��������
'''''    sql = sql & "OIMEAS1, "         ' �n������l�P
'''''    sql = sql & "OIMEAS2, "         ' �n������l�Q
'''''    sql = sql & "OIMEAS3, "         ' �n������l�R
'''''    sql = sql & "OIMEAS4, "         ' �n������l�S
'''''    sql = sql & "OIMEAS5, "         ' �n������l�T
'''''    sql = sql & "ORGRES, "          ' �n�q�f����
'''''    sql = sql & "AVE, "             ' �`�u�d
'''''    sql = sql & "FTIRCONV, "        ' �e�s�h�q���Z
'''''    sql = sql & "INSPECTWAY, "      ' �������@
'''''    sql = sql & "REGDATE "          ' �o�^���t
'''''End Sub


''''''�T�v      :�����֐� Oi���ю擾�p
'''''Private Function Oi_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              Oi As type_DBDRV_scmzc_fcmkc001c_Oi, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim i As Long
'''''    Dim recCnt As Integer
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function Oi_Zisseki"
'''''
'''''    Oi_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ003"
'''''    Set rs = Nothing
'''''
'''''    ' Oi���уe�[�u������l���擾
'''''    DoEvents
'''''    Call Oi_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname)
'''''
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then
'''''        DoEvents
'''''        Call Oi_ObjCpy(Oi, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
''''''----- 8/12 �쑺 �C���i���p���̂Ƃ��͌��ʂ�\���������j
''''''��        If Siyou.HSXONHWS = SIJI Then   ' �d�l�̎w���������Ă���
''''''-----
'''''            If Samp.CRYINDOI = "6" Then       ' ���p���Ȃ�
'''''                DoEvents
'''''                Call Oi_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_HIKITUGI2(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname, TorB)
'''''
'''''                DoEvents
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''
'''''                If rs.RecordCount <> 0 Then
'''''                    recCnt = rs.RecordCount
'''''                    For i = 1 To recCnt
'''''                        If i = 1 Then                                     ' ���ڂ͕ێ�
'''''                            DoEvents
'''''                            Call Oi_ObjCpy(Oi, rs)
'''''                            Exit For    '�P���R�[�h�ڂ�����OK
'''''                        Else
'''''                            If Oi.POSITION = rs("POSITION") And Oi.REGDATE < rs("REGDATE") Then   ' �O�̈ʒu�Ɠ�����������o�^���t���V�������̂��Ƃ�
'''''                                DoEvents
'''''                                Call Oi_ObjCpy(Oi, rs)
'''''                            End If
'''''                            Exit For
'''''                        End If
'''''
'''''                        rs.MoveNext
'''''                    Next
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                If Not rs Is Nothing Then
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''            End If  ' �����w���� 6 �Ȃ�
''''''----- 8/12 �쑺 �C���i���p���̂Ƃ��͌��ʂ�\���������j
''''''��        End If ' �d�l�̎w���������Ă���
''''''-----
'''''    End If ' ���т�����
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    Oi_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''�T�v      :�����֐� BMD���ю擾�p�I�u�W�F�N�g�R�s�[�֐�(���R�[�h�Z�b�g����̃R�s�[)
'''''Private Sub BMD_ObjCpy(BMD As type_DBDRV_scmzc_fcmkc001c_BMD, rs As OraDynaset)
'''''    With BMD
'''''        .CRYNUM = rs("CRYNUM")          ' �����ԍ�
'''''        .POSITION = rs("POSITION")      ' �ʒu
'''''        .SMPKBN = rs("SMPKBN")          ' �T���v���敪
'''''        .SMPLNO = rs("SMPLNO")          ' �T���v���m��
'''''        .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
'''''        .HTPRC = rs("HTPRC")            ' �M�������@
'''''        .KKSP = rs("KKSP")              ' �������ב���ʒu
'''''        .KKSET = rs("KKSET")            ' �������ב�������{�I��ET��
'''''        .TRANCOND = rs("TRANCOND")      ' ��������
'''''        .MEAS1 = rs("MEAS1")            ' ����l�P
'''''        .MEAS2 = rs("MEAS2")            ' ����l�Q
'''''        .MEAS3 = rs("MEAS3")            ' ����l�R
'''''        .MEAS4 = rs("MEAS4")            ' ����l�S
'''''        .MEAS5 = rs("MEAS5")            ' ����l�T
'''''        .Min = rs("MEASMIN")            ' MIN
'''''        .max = rs("MEASMAX")            ' MAX
'''''        .AVE = rs("MEASAVE")            ' AVE
'''''        .REGDATE = rs("REGDATE")        ' �o�^���t
'''''
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''         If IsNull(rs("BMDMNBUNP")) = False Then .BMDMNBUNP = rs("BMDMNBUNP")
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''    End With
'''''End Sub


''''''�T�v      :�����֐� BMD���ю擾�p�x�[�XSQL
'''''Private Sub BMD_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "              ' �����ԍ�
'''''    sql = sql & "POSITION, "            ' �ʒu
'''''    sql = sql & "SMPKBN, "              ' �T���v���敪
'''''    sql = sql & "SMPLNO, "              ' �T���v���m��
'''''    sql = sql & "SMPLUMU, "             ' �T���v���L��
'''''    sql = sql & "HTPRC,"                ' �M�������@"
'''''    sql = sql & "KKSP,"                 ' �������ב���ʒu"
'''''    sql = sql & "KKSET,"                ' �������ב�������{�I��ET��@�@char(1)�{number(2)"
'''''    sql = sql & "TRANCOND, "            ' ��������
'''''    sql = sql & "MEAS1, "               ' ����l�P
'''''    sql = sql & "MEAS2, "               ' ����l�Q
'''''    sql = sql & "MEAS3, "               ' ����l�R
'''''    sql = sql & "MEAS4, "               ' ����l�S
'''''    sql = sql & "MEAS5, "               ' ����l�T
'''''    sql = sql & "MEASMIN, "             ' MIN
'''''    sql = sql & "MEASMAX, "             ' MAX
'''''    sql = sql & "MEASAVE, "             ' AVE
'''''    sql = sql & "REGDATE,"              ' �o�^���t
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''    sql = sql & "BMDMNBUNP "            ' BMD�ʓ����z
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''End Sub



''''''�T�v      :�����֐� BMD���ю擾�p
'''''Private Function BMD_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              inTRANCOND As Integer, _
'''''                              BMD As type_DBDRV_scmzc_fcmkc001c_BMD, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''    ' BMD���уe�[�u������l���擾
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function BMD_Zisseki"
'''''
'''''    BMD_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ008"
'''''    Set rs = Nothing
'''''
'''''    DoEvents
'''''    Call BMD_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname, " and TRANCOND='" & inTRANCOND & "' ")
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then
'''''        DoEvents
'''''        Call BMD_ObjCpy(BMD, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
'''''        If (inTRANCOND = 1 And ((Siyou.HSXBM1HS = SIJI) Or (Siyou.HSXBM1HS = SANKOU))) _
'''''           Or (inTRANCOND = 2 And ((Siyou.HSXBM2HS = SIJI) Or (Siyou.HSXBM2HS = SANKOU))) _
'''''           Or (inTRANCOND = 3 And ((Siyou.HSXBM3HS = SIJI) Or (Siyou.HSXBM3HS = SANKOU))) Then           ' �d�l�̎w���������Ă���
'''''            If (inTRANCOND = 1 And Samp.CRYINDB1 = "6") _
'''''               Or (inTRANCOND = 2 And Samp.CRYINDB2 = "6") _
'''''               Or (inTRANCOND = 3 And Samp.CRYINDB3 = "6") Then       ' ���p���Ȃ�
'''''                DoEvents
'''''                Call BMD_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_HIKITUGI2(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname, TorB, " and TRANCOND='" & inTRANCOND & "' ")
'''''
'''''                DoEvents
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''
'''''                If rs.RecordCount <> 0 Then
'''''                    recCnt = rs.RecordCount
'''''                    For i = 1 To recCnt
'''''                        If i = 1 Then                                     ' ���ڂ͕ێ�
'''''                            DoEvents
'''''                            Call BMD_ObjCpy(BMD, rs)
'''''                            Exit For    '�P���R�[�h�ڂ�����OK
'''''                        Else
'''''                            If BMD.POSITION = rs("POSITION") And BMD.REGDATE < rs("REGDATE") Then   ' �O�̈ʒu�Ɠ�����������o�^���t���V�������̂��Ƃ�
'''''                                DoEvents
'''''                                Call BMD_ObjCpy(BMD, rs)
'''''                            End If
'''''                            Exit For
'''''                        End If
'''''
'''''                        rs.MoveNext
'''''                    Next
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                If Not rs Is Nothing Then
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''            End If  ' �����w����5 or 6 �Ȃ�
'''''        End If ' �w���������Ă���
'''''    End If ' ���т����邩�ǂ���
'''''
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    BMD_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


'''''' ���R�[�h�Z�b�g����̃R�s�[
''''''�T�v      :�����֐� BMD���ю擾�p�I�u�W�F�N�g�R�s�[�֐�(���R�[�h�Z�b�g����̃R�s�[)
'''''Private Sub OSF_ObjCpy(OSF As type_DBDRV_scmzc_fcmkc001c_OSF, rs As OraDynaset)
'''''    With OSF
'''''        .CRYNUM = rs("CRYNUM")         ' �����ԍ�
'''''        .POSITION = rs("POSITION")     ' �ʒu
'''''        .SMPKBN = rs("SMPKBN")         ' �T���v���敪
'''''        .SMPLNO = rs("SMPLNO")         ' �T���v���m��
'''''        .SMPLUMU = rs("SMPLUMU")       ' �T���v���L��
'''''        .HTPRC = rs("HTPRC")           ' �M�������@
'''''        .KKSP = rs("KKSP")             ' �������ב���ʒu
'''''        .KKSET = rs("KKSET")           ' �������ב�������{�I��ET��
'''''        .TRANCOND = rs("TRANCOND")     ' ��������
'''''        .CALCMAX = rs("CALCMAX")       ' �v�Z���� Max
'''''        .CALCAVE = rs("CALCAVE")       ' �v�Z���� Ave
'''''        .MEAS1 = rs("MEAS1")           ' ����l�P
'''''        .MEAS2 = rs("MEAS2")           ' ����l�Q
'''''        .MEAS3 = rs("MEAS3")           ' ����l�R
'''''        .MEAS4 = rs("MEAS4")           ' ����l�S
'''''        .MEAS5 = rs("MEAS5")           ' ����l�T
'''''        .MEAS6 = rs("MEAS6")           ' ����l�U
'''''        .MEAS7 = rs("MEAS7")           ' ����l�V
'''''        .MEAS8 = rs("MEAS8")           ' ����l�W
'''''        .MEAS9 = rs("MEAS9")           ' ����l�X
'''''        .MEAS10 = rs("MEAS10")         ' ����l�P�O
'''''        .MEAS11 = rs("MEAS11")         ' ����l�P�P
'''''        .MEAS12 = rs("MEAS12")         ' ����l�P�Q
'''''        .MEAS13 = rs("MEAS13")         ' ����l�P�R
'''''        .MEAS14 = rs("MEAS14")         ' ����l�P�S
'''''        .MEAS15 = rs("MEAS15")         ' ����l�P�T
'''''        .MEAS16 = rs("MEAS16")         ' ����l�P�U
'''''        .MEAS17 = rs("MEAS17")         ' ����l�P�V
'''''        .MEAS18 = rs("MEAS18")         ' ����l�P�W
'''''        .MEAS19 = rs("MEAS19")         ' ����l�P�X
'''''        .MEAS20 = rs("MEAS20")         ' ����l�Q�O
'''''        .REGDATE = rs("REGDATE")       ' �o�^���t
'''''
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''         If IsNull(rs("OSFPOS1")) = False Then .OSFPOS1 = rs("OSFPOS1")   '����݋敪�P�ʒu
'''''         If IsNull(rs("OSFWID1")) = False Then .OSFWID1 = rs("OSFWID1")   '����݋敪�P��
'''''         If IsNull(rs("OSFRD1")) = False Then .OSFRD1 = rs("OSFRD1")      '����݋敪�PR/D
'''''         If IsNull(rs("OSFPOS2")) = False Then .OSFPOS2 = rs("OSFPOS2")   '����݋敪�Q�ʒu
'''''         If IsNull(rs("OSFWID2")) = False Then .OSFWID2 = rs("OSFWID2")   '����݋敪�Q��
'''''         If IsNull(rs("OSFRD2")) = False Then .OSFRD2 = rs("OSFRD2")      '����݋敪�QR/D
'''''         If IsNull(rs("OSFPOS3")) = False Then .OSFPOS3 = rs("OSFPOS3")   '����݋敪�R�ʒu
'''''         If IsNull(rs("OSFWID3")) = False Then .OSFWID3 = rs("OSFWID3")   '����݋敪�R��
'''''         If IsNull(rs("OSFRD3")) = False Then .OSFRD3 = rs("OSFRD3")      '����݋敪�RR/D
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''    End With
'''''End Sub
'''''
''''''�T�v      :�����֐� BMD���ю擾�p�x�[�XSQL
'''''Private Sub OSF_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "        ' �����ԍ�
'''''    sql = sql & "POSITION, "      ' �ʒu
'''''    sql = sql & "SMPKBN, "        ' �T���v���敪
'''''    sql = sql & "SMPLNO, "        ' �T���v���m��
'''''    sql = sql & "SMPLUMU, "       ' �T���v���L��
'''''    sql = sql & "HTPRC,"          ' �M�������@"
'''''    sql = sql & "KKSP,"           ' �������ב���ʒu"
'''''    sql = sql & "KKSET,"          ' �������ב�������{�I��ET��@�@char(1)�{number(2)"
'''''    sql = sql & "TRANCOND, "      ' ��������
'''''    sql = sql & "CALCMAX, "       ' �v�Z���� Max
'''''    sql = sql & "CALCAVE, "       ' �v�Z���� Ave
'''''    sql = sql & "MEAS1, "         ' ����l�P
'''''    sql = sql & "MEAS2, "         ' ����l�Q
'''''    sql = sql & "MEAS3, "         ' ����l�R
'''''    sql = sql & "MEAS4, "         ' ����l�S
'''''    sql = sql & "MEAS5, "         ' ����l�T
'''''    sql = sql & "MEAS6, "         ' ����l�U
'''''    sql = sql & "MEAS7, "         ' ����l�V
'''''    sql = sql & "MEAS8, "         ' ����l�W
'''''    sql = sql & "MEAS9, "         ' ����l�X
'''''    sql = sql & "MEAS10, "        ' ����l�P�O
'''''    sql = sql & "MEAS11, "        ' ����l�P�P
'''''    sql = sql & "MEAS12, "        ' ����l�P�Q
'''''    sql = sql & "MEAS13, "        ' ����l�P�R
'''''    sql = sql & "MEAS14, "        ' ����l�P�S
'''''    sql = sql & "MEAS15, "        ' ����l�P�T
'''''    sql = sql & "MEAS16, "        ' ����l�P�U
'''''    sql = sql & "MEAS17, "        ' ����l�P�V
'''''    sql = sql & "MEAS18, "        ' ����l�P�W
'''''    sql = sql & "MEAS19, "        ' ����l�P�X
'''''    sql = sql & "MEAS20, "        ' ����l�Q�O
'''''    sql = sql & "REGDATE, "       ' �o�^���t
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''    sql = sql & "OSFPOS1, "       ' ����݋敪�P�ʒu
'''''    sql = sql & "OSFWID1, "       ' ����݋敪�P��
'''''    sql = sql & "OSFRD1, "        ' ����݋敪�PR/D
'''''    sql = sql & "OSFPOS2, "       ' ����݋敪�Q�ʒu
'''''    sql = sql & "OSFWID2, "       ' ����݋敪�Q��
'''''    sql = sql & "OSFRD2, "        ' ����݋敪�QR/D
'''''    sql = sql & "OSFPOS3, "       ' ����݋敪�R�ʒu
'''''    sql = sql & "OSFWID3, "       ' ����݋敪�R��
'''''    sql = sql & "OSFRD3 "         ' ����݋敪�RR/D
'''''' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'''''End Sub


''''''�T�v      :�����֐� OSF���ю擾�p
'''''Private Function OSF_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              inTRANCOND As Integer, _
'''''                              OSF As type_DBDRV_scmzc_fcmkc001c_OSF, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''
'''''    ' OSF���уe�[�u������l���擾
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function OSF_Zisseki"
'''''
'''''    OSF_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ005"
'''''    Set rs = Nothing
'''''
'''''    DoEvents
'''''    Call OSF_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname, " and TRANCOND='" & inTRANCOND & "' ")
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then
'''''        DoEvents
'''''        Call OSF_ObjCpy(OSF, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
'''''        rs.Close
'''''        Set rs = Nothing
'''''        If (inTRANCOND = 1 And ((Siyou.HSXOS1HS = SIJI) Or (Siyou.HSXOS1HS = SANKOU))) _
'''''           Or (inTRANCOND = 2 And ((Siyou.HSXOS2HS = SIJI) Or (Siyou.HSXOS2HS = SANKOU))) _
'''''           Or (inTRANCOND = 3 And ((Siyou.HSXOS3HS = SIJI) Or (Siyou.HSXOS3HS = SANKOU))) Then          ' �d�l�̎w���������Ă���
'''''           If (inTRANCOND = 1 And Samp.CRYINDL1 = "6") _
'''''              Or (inTRANCOND = 2 And Samp.CRYINDL2 = "6") _
'''''              Or (inTRANCOND = 3 And Samp.CRYINDL3 = "6") Then       ' ���p���Ȃ�
'''''                DoEvents
'''''                Call OSF_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_HIKITUGI2(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname, TorB, " and TRANCOND='" & inTRANCOND & "' ")
'''''                DoEvents
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''
'''''                If rs.RecordCount <> 0 Then
'''''                    recCnt = rs.RecordCount
'''''                    For i = 1 To recCnt
'''''                        If i = 1 Then                                     ' ���ڂ͕ێ�
'''''                            DoEvents
'''''                            Call OSF_ObjCpy(OSF, rs)
'''''                            Exit For    '�P���R�[�h�ڂ�����OK
'''''                        Else
'''''                            If OSF.POSITION = rs("POSITION") And OSF.REGDATE < rs("REGDATE") Then   ' �O�̈ʒu�Ɠ�����������o�^���t���V�������̂��Ƃ�
'''''                                DoEvents
'''''                                Call OSF_ObjCpy(OSF, rs)
'''''                            End If
'''''                        End If
'''''                        rs.MoveNext
'''''                    Next
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                If Not rs Is Nothing Then
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''            End If  ' �����w����5 or 6 �Ȃ�
'''''        End If  ' �w���������Ă���
'''''    End If ' ���т�����
'''''
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    OSF_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


'''''Private Sub Cs_ObjCpy(Cs As type_DBDRV_scmzc_fcmkc001c_CS, rs As OraDynaset)
'''''    With Cs
'''''        .CRYNUM = rs("CRYNUM")         ' �����ԍ�
'''''        .POSITION = rs("POSITION")     ' �ʒu
'''''        .SMPKBN = rs("SMPKBN")         ' �T���v���敪
'''''        .SMPLNO = rs("SMPLNO")         ' �T���v���m��
'''''        .SMPLUMU = rs("SMPLUMU")       ' �T���v���L��
'''''        .TRANCOND = rs("TRANCOND")     ' ��������
'''''        .CSMEAS = rs("CSMEAS")         ' Cs�����l
'''''        .PRE70P = rs("PRE70P")         ' �V�O������l
'''''        .REGDATE = rs("REGDATE")        ' �o�^���t
'''''    End With
'''''
'''''End Sub


'''''Private Sub Cs_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "        ' �����ԍ�
'''''    sql = sql & "POSITION, "      ' �ʒu
'''''    sql = sql & "SMPKBN, "        ' �T���v���敪
'''''    sql = sql & "SMPLNO, "        ' �T���v���m��
'''''    sql = sql & "SMPLUMU, "       ' �T���v���L��
'''''    sql = sql & "TRANCOND, "      ' ��������
'''''    sql = sql & "CSMEAS, "        ' Cs�����l
'''''    sql = sql & "PRE70P, "         ' �V�O������l
'''''    sql = sql & "REGDATE "        ' �o�^���t
'''''
'''''End Sub


''''''�����֐� Cs���ю擾�p
'''''Private Function CS_Zisseki(CRYNUM As String, Samp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, Cs() As type_DBDRV_scmzc_fcmkc001c_CS) As FUNCTION_RETURN
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''Dim recCnt As Integer
'''''Dim i As Long
'''''Dim Tname As String
'''''Dim jCs As String
'''''Dim jCsFromTo As String
'''''Dim tt As Integer
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function CS_Zisseki"
'''''    CS_Zisseki = FUNCTION_RETURN_FAILURE
'''''
'''''    '�u���b�N���̕i�Ԃɂ���Cs�d�l�̗L�����`�F�b�N����
'''''    If DBDRV_scmzc_fcmkc001c_CheckSpecCs(CRYNUM, Samp(1).INGOTPOS, Samp(2).INGOTPOS, jCs, jCsFromTo) = FUNCTION_RETURN_FAILURE Then
'''''        GoTo proc_err
'''''    End If
'''''
'''''    For tt = 1 To 2
'''''        With Cs(tt)
'''''            .CRYNUM = vbNullString
'''''            .CSMEAS = -1
'''''            .POSITION = Samp(tt).INGOTPOS
'''''            .PRE70P = -1
'''''            .SMPLNO = -1
'''''            .SMPLUMU = "0"
'''''        End With
'''''    Next
'''''    If (jCsFromTo = SIJI) Or (jCsFromTo = SANKOU) Then 'FromTo�d�l���܂ޕi�Ԃ����邽�߁A���p�s��
'''''        For tt = 1 To 2
'''''            Tname = "TBCMJ004"
'''''            Call Cs_SetBaseSQL(sql)
'''''            Call AddSQL_Default2(sql, CRYNUM, Samp(tt).INGOTPOS, Samp(tt).SMPKBN, Tname)
'''''            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''            If rs.RecordCount > 0 Then
'''''                Call Cs_ObjCpy(Cs(tt), rs)
'''''            End If
'''''            rs.Close
'''''            Set rs = Nothing
'''''        Next
'''''    ElseIf (jCs = SIJI) Or (jCs = SANKOU) Then         'Cs�d�l���܂ޕi�Ԃ�����BTail���̂݉���������ю擾
'''''        tt = 2
'''''        Tname = "TBCMJ004"
'''''        Call Cs_SetBaseSQL(sql)
'''''        sql = sql & " from TBCMJ004 T1"
'''''        sql = sql & " where T1.CRYNUM='" & CRYNUM & "'"
'''''        sql = sql & " and T1.TRANCNT=(select max(TRANCNT) from TBCMJ004 where CRYNUM=T1.CRYNUM and POSITION=T1.POSITION and SMPKBN=T1.SMPKBN)"
'''''        sql = sql & " and T1.SMPLUMU='0'"
'''''        sql = sql & " and T1.POSITION>=" & Samp(tt).INGOTPOS
'''''        sql = sql & " order by POSITION asc, SMPKBN asc"
'''''        sql = "select * from (" & sql & ") where rownum=1"
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''        If rs.RecordCount > 0 Then
'''''            Call Cs_ObjCpy(Cs(tt), rs)
'''''        End If
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''
'''''    CS_Zisseki = FUNCTION_RETURN_SUCCESS
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


'''''Private Sub GD_ObjCpy(GD As type_DBDRV_scmzc_fcmkc001c_GD, rs As OraDynaset)
'''''    With GD
'''''        .CRYNUM = rs("CRYNUM")         ' �����ԍ�
'''''        .POSITION = rs("POSITION")     ' �ʒu
'''''        .SMPKBN = rs("SMPKBN")         ' �T���v���敪
'''''        .SMPLNO = rs("SMPLNO")         ' �T���v���m��
'''''        .SMPLUMU = rs("SMPLUMU")       ' �T���v���L��
'''''        .TRANCOND = rs("TRANCOND")     ' ��������
'''''        .MSRSDEN = rs("MSRSDEN")       ' ���茋�� Den
'''''        .MSRSLDL = rs("MSRSLDL")       ' ���茋�� L/DL
'''''        .MSRSDVD2 = rs("MSRSDVD2")     ' ���茋�� DVD2
'''''        .MS01LDL1 = rs("MS01LDL1")            ' ����l01 L/DL1
'''''        .MS01LDL2 = rs("MS01LDL2")            ' ����l01 L/DL2
'''''        .MS01LDL3 = rs("MS01LDL3")            ' ����l01 L/DL3
'''''        .MS01LDL4 = rs("MS01LDL4")            ' ����l01 L/DL4
'''''        .MS01LDL5 = rs("MS01LDL5")            ' ����l01 L/DL5
'''''        .MS01DEN1 = rs("MS01DEN1")            ' ����l01 Den1
'''''        .MS01DEN2 = rs("MS01DEN2")            ' ����l01 Den2
'''''        .MS01DEN3 = rs("MS01DEN3")            ' ����l01 Den3
'''''        .MS01DEN4 = rs("MS01DEN4")            ' ����l01 Den4
'''''        .MS01DEN5 = rs("MS01DEN5")            ' ����l01 Den5
'''''        .MS02LDL1 = rs("MS02LDL1")            ' ����l02 L/DL1
'''''        .MS02LDL2 = rs("MS02LDL2")            ' ����l02 L/DL2
'''''        .MS02LDL3 = rs("MS02LDL3")            ' ����l02 L/DL3
'''''        .MS02LDL4 = rs("MS02LDL4")            ' ����l02 L/DL4
'''''        .MS02LDL5 = rs("MS02LDL5")            ' ����l02 L/DL5
'''''        .MS02DEN1 = rs("MS02DEN1")            ' ����l02 Den1
'''''        .MS02DEN2 = rs("MS02DEN2")            ' ����l02 Den2
'''''        .MS02DEN3 = rs("MS02DEN3")            ' ����l02 Den3
'''''        .MS02DEN4 = rs("MS02DEN4")            ' ����l02 Den4
'''''        .MS02DEN5 = rs("MS02DEN5")            ' ����l02 Den5
'''''        .MS03LDL1 = rs("MS03LDL1")            ' ����l03 L/DL1
'''''        .MS03LDL2 = rs("MS03LDL2")            ' ����l03 L/DL2
'''''        .MS03LDL3 = rs("MS03LDL3")            ' ����l03 L/DL3
'''''        .MS03LDL4 = rs("MS03LDL4")            ' ����l03 L/DL4
'''''        .MS03LDL5 = rs("MS03LDL5")            ' ����l03 L/DL5
'''''        .MS03DEN1 = rs("MS03DEN1")            ' ����l03 Den1
'''''        .MS03DEN2 = rs("MS03DEN2")            ' ����l03 Den2
'''''        .MS03DEN3 = rs("MS03DEN3")            ' ����l03 Den3
'''''        .MS03DEN4 = rs("MS03DEN4")            ' ����l03 Den4
'''''        .MS03DEN5 = rs("MS03DEN5")            ' ����l03 Den5
'''''        .MS04LDL1 = rs("MS04LDL1")            ' ����l04 L/DL1
'''''        .MS04LDL2 = rs("MS04LDL2")            ' ����l04 L/DL2
'''''        .MS04LDL3 = rs("MS04LDL3")            ' ����l04 L/DL3
'''''        .MS04LDL4 = rs("MS04LDL4")            ' ����l04 L/DL4
'''''        .MS04LDL5 = rs("MS04LDL5")            ' ����l04 L/DL5
'''''        .MS04DEN1 = rs("MS04DEN1")            ' ����l04 Den1
'''''        .MS04DEN2 = rs("MS04DEN2")            ' ����l04 Den2
'''''        .MS04DEN3 = rs("MS04DEN3")            ' ����l04 Den3
'''''        .MS04DEN4 = rs("MS04DEN4")            ' ����l04 Den4
'''''        .MS04DEN5 = rs("MS04DEN5")            ' ����l04 Den5
'''''        .MS05LDL1 = rs("MS05LDL1")            ' ����l05 L/DL1
'''''        .MS05LDL2 = rs("MS05LDL2")            ' ����l05 L/DL2
'''''        .MS05LDL3 = rs("MS05LDL3")            ' ����l05 L/DL3
'''''        .MS05LDL4 = rs("MS05LDL4")            ' ����l05 L/DL4
'''''        .MS05LDL5 = rs("MS05LDL5")            ' ����l05 L/DL5
'''''        .MS05DEN1 = rs("MS05DEN1")            ' ����l05 Den1
'''''        .MS05DEN2 = rs("MS05DEN2")            ' ����l05 Den2
'''''        .MS05DEN3 = rs("MS05DEN3")            ' ����l05 Den3
'''''        .MS05DEN4 = rs("MS05DEN4")            ' ����l05 Den4
'''''        .MS05DEN5 = rs("MS05DEN5")            ' ����l05 Den5
'''''        .MS06LDL1 = rs("MS06LDL1")            ' ����l06 L/DL1
'''''        .MS06LDL2 = rs("MS06LDL2")            ' ����l06 L/DL2
'''''        .MS06LDL3 = rs("MS06LDL3")            ' ����l06 L/DL3
'''''        .MS06LDL4 = rs("MS06LDL4")            ' ����l06 L/DL4
'''''        .MS06LDL5 = rs("MS06LDL5")            ' ����l06 L/DL5
'''''        .MS06DEN1 = rs("MS06DEN1")            ' ����l06 Den1
'''''        .MS06DEN2 = rs("MS06DEN2")            ' ����l06 Den2
'''''        .MS06DEN3 = rs("MS06DEN3")            ' ����l06 Den3
'''''        .MS06DEN4 = rs("MS06DEN4")            ' ����l06 Den4
'''''        .MS06DEN5 = rs("MS06DEN5")            ' ����l06 Den5
'''''        .MS07LDL1 = rs("MS07LDL1")            ' ����l07 L/DL1
'''''        .MS07LDL2 = rs("MS07LDL2")            ' ����l07 L/DL2
'''''        .MS07LDL3 = rs("MS07LDL3")            ' ����l07 L/DL3
'''''        .MS07LDL4 = rs("MS07LDL4")            ' ����l07 L/DL4
'''''        .MS07LDL5 = rs("MS07LDL5")            ' ����l07 L/DL5
'''''        .MS07DEN1 = rs("MS07DEN1")            ' ����l07 Den1
'''''        .MS07DEN2 = rs("MS07DEN2")            ' ����l07 Den2
'''''        .MS07DEN3 = rs("MS07DEN3")            ' ����l07 Den3
'''''        .MS07DEN4 = rs("MS07DEN4")            ' ����l07 Den4
'''''        .MS07DEN5 = rs("MS07DEN5")            ' ����l07 Den5
'''''        .MS08LDL1 = rs("MS08LDL1")            ' ����l08 L/DL1
'''''        .MS08LDL2 = rs("MS08LDL2")            ' ����l08 L/DL2
'''''        .MS08LDL3 = rs("MS08LDL3")            ' ����l08 L/DL3
'''''        .MS08LDL4 = rs("MS08LDL4")            ' ����l08 L/DL4
'''''        .MS08LDL5 = rs("MS08LDL5")            ' ����l08 L/DL5
'''''        .MS08DEN1 = rs("MS08DEN1")            ' ����l08 Den1
'''''        .MS08DEN2 = rs("MS08DEN2")            ' ����l08 Den2
'''''        .MS08DEN3 = rs("MS08DEN3")            ' ����l08 Den3
'''''        .MS08DEN4 = rs("MS08DEN4")            ' ����l08 Den4
'''''        .MS08DEN5 = rs("MS08DEN5")            ' ����l08 Den5
'''''        .MS09LDL1 = rs("MS09LDL1")            ' ����l09 L/DL1
'''''        .MS09LDL2 = rs("MS09LDL2")            ' ����l09 L/DL2
'''''        .MS09LDL3 = rs("MS09LDL3")            ' ����l09 L/DL3
'''''        .MS09LDL4 = rs("MS09LDL4")            ' ����l09 L/DL4
'''''        .MS09LDL5 = rs("MS09LDL5")            ' ����l09 L/DL5
'''''        .MS09DEN1 = rs("MS09DEN1")            ' ����l09 Den1
'''''        .MS09DEN2 = rs("MS09DEN2")            ' ����l09 Den2
'''''        .MS09DEN3 = rs("MS09DEN3")            ' ����l09 Den3
'''''        .MS09DEN4 = rs("MS09DEN4")            ' ����l09 Den4
'''''        .MS09DEN5 = rs("MS09DEN5")            ' ����l09 Den5
'''''        .MS10LDL1 = rs("MS10LDL1")            ' ����l10 L/DL1
'''''        .MS10LDL2 = rs("MS10LDL2")            ' ����l10 L/DL2
'''''        .MS10LDL3 = rs("MS10LDL3")            ' ����l10 L/DL3
'''''        .MS10LDL4 = rs("MS10LDL4")            ' ����l10 L/DL4
'''''        .MS10LDL5 = rs("MS10LDL5")            ' ����l10 L/DL5
'''''        .MS10DEN1 = rs("MS10DEN1")            ' ����l10 Den1
'''''        .MS10DEN2 = rs("MS10DEN2")            ' ����l10 Den2
'''''        .MS10DEN3 = rs("MS10DEN3")            ' ����l10 Den3
'''''        .MS10DEN4 = rs("MS10DEN4")            ' ����l10 Den4
'''''        .MS10DEN5 = rs("MS10DEN5")            ' ����l10 Den5
'''''        .MS11LDL1 = rs("MS11LDL1")            ' ����l11 L/DL1
'''''        .MS11LDL2 = rs("MS11LDL2")            ' ����l11 L/DL2
'''''        .MS11LDL3 = rs("MS11LDL3")            ' ����l11 L/DL3
'''''        .MS11LDL4 = rs("MS11LDL4")            ' ����l11 L/DL4
'''''        .MS11LDL5 = rs("MS11LDL5")            ' ����l11 L/DL5
'''''        .MS11DEN1 = rs("MS11DEN1")            ' ����l11 Den1
'''''        .MS11DEN2 = rs("MS11DEN2")            ' ����l11 Den2
'''''        .MS11DEN3 = rs("MS11DEN3")            ' ����l11 Den3
'''''        .MS11DEN4 = rs("MS11DEN4")            ' ����l11 Den4
'''''        .MS11DEN5 = rs("MS11DEN5")            ' ����l11 Den5
'''''        .MS12LDL1 = rs("MS12LDL1")            ' ����l12 L/DL1
'''''        .MS12LDL2 = rs("MS12LDL2")            ' ����l12 L/DL2
'''''        .MS12LDL3 = rs("MS12LDL3")            ' ����l12 L/DL3
'''''        .MS12LDL4 = rs("MS12LDL4")            ' ����l12 L/DL4
'''''        .MS12LDL5 = rs("MS12LDL5")            ' ����l12 L/DL5
'''''        .MS12DEN1 = rs("MS12DEN1")            ' ����l12 Den1
'''''        .MS12DEN2 = rs("MS12DEN2")            ' ����l12 Den2
'''''        .MS12DEN3 = rs("MS12DEN3")            ' ����l12 Den3
'''''        .MS12DEN4 = rs("MS12DEN4")            ' ����l12 Den4
'''''        .MS12DEN5 = rs("MS12DEN5")            ' ����l12 Den5
'''''        .MS13LDL1 = rs("MS13LDL1")            ' ����l13 L/DL1
'''''        .MS13LDL2 = rs("MS13LDL2")            ' ����l13 L/DL2
'''''        .MS13LDL3 = rs("MS13LDL3")            ' ����l13 L/DL3
'''''        .MS13LDL4 = rs("MS13LDL4")            ' ����l13 L/DL4
'''''        .MS13LDL5 = rs("MS13LDL5")            ' ����l13 L/DL5
'''''        .MS13DEN1 = rs("MS13DEN1")            ' ����l13 Den1
'''''        .MS13DEN2 = rs("MS13DEN2")            ' ����l13 Den2
'''''        .MS13DEN3 = rs("MS13DEN3")            ' ����l13 Den3
'''''        .MS13DEN4 = rs("MS13DEN4")            ' ����l13 Den4
'''''        .MS13DEN5 = rs("MS13DEN5")            ' ����l13 Den5
'''''        .MS14LDL1 = rs("MS14LDL1")            ' ����l14 L/DL1
'''''        .MS14LDL2 = rs("MS14LDL2")            ' ����l14 L/DL2
'''''        .MS14LDL3 = rs("MS14LDL3")            ' ����l14 L/DL3
'''''        .MS14LDL4 = rs("MS14LDL4")            ' ����l14 L/DL4
'''''        .MS14LDL5 = rs("MS14LDL5")            ' ����l14 L/DL5
'''''        .MS14DEN1 = rs("MS14DEN1")            ' ����l14 Den1
'''''        .MS14DEN2 = rs("MS14DEN2")            ' ����l14 Den2
'''''        .MS14DEN3 = rs("MS14DEN3")            ' ����l14 Den3
'''''        .MS14DEN4 = rs("MS14DEN4")            ' ����l14 Den4
'''''        .MS14DEN5 = rs("MS14DEN5")            ' ����l14 Den5
'''''        .MS15LDL1 = rs("MS15LDL1")            ' ����l15 L/DL1
'''''        .MS15LDL2 = rs("MS15LDL2")            ' ����l15 L/DL2
'''''        .MS15LDL3 = rs("MS15LDL3")            ' ����l15 L/DL3
'''''        .MS15LDL4 = rs("MS15LDL4")            ' ����l15 L/DL4
'''''        .MS15LDL5 = rs("MS15LDL5")            ' ����l15 L/DL5
'''''        .MS15DEN1 = rs("MS15DEN1")            ' ����l15 Den1
'''''        .MS15DEN2 = rs("MS15DEN2")            ' ����l15 Den2
'''''        .MS15DEN3 = rs("MS15DEN3")            ' ����l15 Den3
'''''        .MS15DEN4 = rs("MS15DEN4")            ' ����l15 Den4
'''''        .MS15DEN5 = rs("MS15DEN5")            ' ����l15 Den5
'''''        .REGDATE = rs("REGDATE")              ' �o�^���t
'''''    End With
'''''
'''''End Sub
'''''
'''''Private Sub GD_SetBaseSQL(sql As String)
'''''    ' GD���уe�[�u������l���擾
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "               ' �����ԍ�
'''''    sql = sql & "POSITION, "             ' �ʒu
'''''    sql = sql & "SMPKBN, "               ' �T���v���敪
'''''    sql = sql & "SMPLNO, "               ' �T���v���m��
'''''    sql = sql & "SMPLUMU, "              ' �T���v���L��
'''''    sql = sql & "TRANCOND, "             ' ��������
'''''    sql = sql & "MSRSDEN, "              ' ���茋�� Den
'''''    sql = sql & "MSRSLDL, "              ' ���茋�� L/DL
'''''    sql = sql & "MSRSDVD2, "             ' ���茋�� DVD2
'''''    sql = sql & "MS01LDL1, "             ' ����l01 L/DL1"
'''''    sql = sql & "MS01LDL2, "             ' ����l01 L/DL2"
'''''    sql = sql & "MS01LDL3, "             ' ����l01 L/DL3"
'''''    sql = sql & "MS01LDL4, "             ' ����l01 L/DL4"
'''''    sql = sql & "MS01LDL5, "             ' ����l01 L/DL5"
'''''    sql = sql & "MS01DEN1, "             ' ����l01 Den1"
'''''    sql = sql & "MS01DEN2, "             ' ����l01 Den2"
'''''    sql = sql & "MS01DEN3, "             ' ����l01 Den3"
'''''    sql = sql & "MS01DEN4, "             ' ����l01 Den4"
'''''    sql = sql & "MS01DEN5, "             ' ����l01 Den5"
'''''    sql = sql & "MS02LDL1, "             ' ����l02 L/DL1"
'''''    sql = sql & "MS02LDL2, "             ' ����l02 L/DL2"
'''''    sql = sql & "MS02LDL3, "             ' ����l02 L/DL3"
'''''    sql = sql & "MS02LDL4, "             ' ����l02 L/DL4"
'''''    sql = sql & "MS02LDL5, "             ' ����l02 L/DL5"
'''''    sql = sql & "MS02DEN1, "             ' ����l02 Den1"
'''''    sql = sql & "MS02DEN2, "             ' ����l02 Den2"
'''''    sql = sql & "MS02DEN3, "             ' ����l02 Den3"
'''''    sql = sql & "MS02DEN4, "             ' ����l02 Den4"
'''''    sql = sql & "MS02DEN5, "             ' ����l02 Den5"
'''''    sql = sql & "MS03LDL1, "             ' ����l03 L/DL1"
'''''    sql = sql & "MS03LDL2, "             ' ����l03 L/DL2"
'''''    sql = sql & "MS03LDL3, "             ' ����l03 L/DL3"
'''''    sql = sql & "MS03LDL4, "             ' ����l03 L/DL4"
'''''    sql = sql & "MS03LDL5, "             ' ����l03 L/DL5"
'''''    sql = sql & "MS03DEN1, "             ' ����l03 Den1"
'''''    sql = sql & "MS03DEN2, "             ' ����l03 Den2"
'''''    sql = sql & "MS03DEN3, "             ' ����l03 Den3"
'''''    sql = sql & "MS03DEN4, "             ' ����l03 Den4"
'''''    sql = sql & "MS03DEN5, "             ' ����l03 Den5"
'''''    sql = sql & "MS04LDL1, "             ' ����l04 L/DL1"
'''''    sql = sql & "MS04LDL2, "             ' ����l04 L/DL2"
'''''    sql = sql & "MS04LDL3, "             ' ����l04 L/DL3"
'''''    sql = sql & "MS04LDL4, "             ' ����l04 L/DL4"
'''''    sql = sql & "MS04LDL5, "             ' ����l04 L/DL5"
'''''    sql = sql & "MS04DEN1, "             ' ����l04 Den1"
'''''    sql = sql & "MS04DEN2, "             ' ����l04 Den2"
'''''    sql = sql & "MS04DEN3, "             ' ����l04 Den3"
'''''    sql = sql & "MS04DEN4, "             ' ����l04 Den4"
'''''    sql = sql & "MS04DEN5, "             ' ����l04 Den5"
'''''    sql = sql & "MS05LDL1, "             ' ����l05 L/DL1"
'''''    sql = sql & "MS05LDL2, "             ' ����l05 L/DL2"
'''''    sql = sql & "MS05LDL3, "             ' ����l05 L/DL3"
'''''    sql = sql & "MS05LDL4, "             ' ����l05 L/DL4"
'''''    sql = sql & "MS05LDL5, "             ' ����l05 L/DL5"
'''''    sql = sql & "MS05DEN1, "             ' ����l05 Den1"
'''''    sql = sql & "MS05DEN2, "             ' ����l05 Den2"
'''''    sql = sql & "MS05DEN3, "             ' ����l05 Den3"
'''''    sql = sql & "MS05DEN4, "             ' ����l05 Den4"
'''''    sql = sql & "MS05DEN5, "             ' ����l05 Den5"
'''''    sql = sql & "MS06LDL1, "             ' ����l06 L/DL1"
'''''    sql = sql & "MS06LDL2, "             ' ����l06 L/DL2"
'''''    sql = sql & "MS06LDL3, "             ' ����l06 L/DL3"
'''''    sql = sql & "MS06LDL4, "             ' ����l06 L/DL4"
'''''    sql = sql & "MS06LDL5, "             ' ����l06 L/DL5"
'''''    sql = sql & "MS06DEN1, "             ' ����l06 Den1"
'''''    sql = sql & "MS06DEN2, "             ' ����l06 Den2"
'''''    sql = sql & "MS06DEN3, "             ' ����l06 Den3"
'''''    sql = sql & "MS06DEN4, "             ' ����l06 Den4"
'''''    sql = sql & "MS06DEN5, "             ' ����l06 Den5"
'''''    sql = sql & "MS07LDL1, "             ' ����l07 L/DL1"
'''''    sql = sql & "MS07LDL2, "             ' ����l07 L/DL2"
'''''    sql = sql & "MS07LDL3, "             ' ����l07 L/DL3"
'''''    sql = sql & "MS07LDL4, "             ' ����l07 L/DL4"
'''''    sql = sql & "MS07LDL5, "             ' ����l07 L/DL5"
'''''    sql = sql & "MS07DEN1, "             ' ����l07 Den1"
'''''    sql = sql & "MS07DEN2, "             ' ����l07 Den2"
'''''    sql = sql & "MS07DEN3, "             ' ����l07 Den3"
'''''    sql = sql & "MS07DEN4, "             ' ����l07 Den4"
'''''    sql = sql & "MS07DEN5, "             ' ����l07 Den5"
'''''    sql = sql & "MS08LDL1, "             ' ����l08 L/DL1"
'''''    sql = sql & "MS08LDL2, "             ' ����l08 L/DL2"
'''''    sql = sql & "MS08LDL3, "             ' ����l08 L/DL3"
'''''    sql = sql & "MS08LDL4, "             ' ����l08 L/DL4"
'''''    sql = sql & "MS08LDL5, "             ' ����l08 L/DL5"
'''''    sql = sql & "MS08DEN1, "             ' ����l08 Den1"
'''''    sql = sql & "MS08DEN2, "             ' ����l08 Den2"
'''''    sql = sql & "MS08DEN3, "             ' ����l08 Den3"
'''''    sql = sql & "MS08DEN4, "             ' ����l08 Den4"
'''''    sql = sql & "MS08DEN5, "             ' ����l08 Den5"
'''''    sql = sql & "MS09LDL1, "             ' ����l09 L/DL1"
'''''    sql = sql & "MS09LDL2, "             ' ����l09 L/DL2"
'''''    sql = sql & "MS09LDL3, "             ' ����l09 L/DL3"
'''''    sql = sql & "MS09LDL4, "             ' ����l09 L/DL4"
'''''    sql = sql & "MS09LDL5, "             ' ����l09 L/DL5"
'''''    sql = sql & "MS09DEN1, "             ' ����l09 Den1"
'''''    sql = sql & "MS09DEN2, "             ' ����l09 Den2"
'''''    sql = sql & "MS09DEN3, "             ' ����l09 Den3"
'''''    sql = sql & "MS09DEN4, "             ' ����l09 Den4"
'''''    sql = sql & "MS09DEN5, "             ' ����l09 Den5"
'''''    sql = sql & "MS10LDL1, "             ' ����l10 L/DL1"
'''''    sql = sql & "MS10LDL2, "             ' ����l10 L/DL2"
'''''    sql = sql & "MS10LDL3, "             ' ����l10 L/DL3"
'''''    sql = sql & "MS10LDL4, "             ' ����l10 L/DL4"
'''''    sql = sql & "MS10LDL5, "             ' ����l10 L/DL5"
'''''    sql = sql & "MS10DEN1, "             ' ����l10 Den1"
'''''    sql = sql & "MS10DEN2, "             ' ����l10 Den2"
'''''    sql = sql & "MS10DEN3, "             ' ����l10 Den3"
'''''    sql = sql & "MS10DEN4, "             ' ����l10 Den4"
'''''    sql = sql & "MS10DEN5, "             ' ����l10 Den5"
'''''    sql = sql & "MS11LDL1, "             ' ����l11 L/DL1"
'''''    sql = sql & "MS11LDL2, "             ' ����l11 L/DL2"
'''''    sql = sql & "MS11LDL3, "             ' ����l11 L/DL3"
'''''    sql = sql & "MS11LDL4, "             ' ����l11 L/DL4"
'''''    sql = sql & "MS11LDL5, "             ' ����l11 L/DL5"
'''''    sql = sql & "MS11DEN1, "             ' ����l11 Den1"
'''''    sql = sql & "MS11DEN2, "             ' ����l11 Den2"
'''''    sql = sql & "MS11DEN3, "             ' ����l11 Den3"
'''''    sql = sql & "MS11DEN4, "             ' ����l11 Den4"
'''''    sql = sql & "MS11DEN5, "             ' ����l11 Den5"
'''''    sql = sql & "MS12LDL1, "             ' ����l12 L/DL1"
'''''    sql = sql & "MS12LDL2, "             ' ����l12 L/DL2"
'''''    sql = sql & "MS12LDL3, "             ' ����l12 L/DL3"
'''''    sql = sql & "MS12LDL4, "             ' ����l12 L/DL4"
'''''    sql = sql & "MS12LDL5, "             ' ����l12 L/DL5"
'''''    sql = sql & "MS12DEN1, "             ' ����l12 Den1"
'''''    sql = sql & "MS12DEN2, "             ' ����l12 Den2"
'''''    sql = sql & "MS12DEN3, "             ' ����l12 Den3"
'''''    sql = sql & "MS12DEN4, "             ' ����l12 Den4"
'''''    sql = sql & "MS12DEN5, "             ' ����l12 Den5"
'''''    sql = sql & "MS13LDL1, "             ' ����l13 L/DL1"
'''''    sql = sql & "MS13LDL2, "             ' ����l13 L/DL2"
'''''    sql = sql & "MS13LDL3, "             ' ����l13 L/DL3"
'''''    sql = sql & "MS13LDL4, "             ' ����l13 L/DL4"
'''''    sql = sql & "MS13LDL5, "             ' ����l13 L/DL5"
'''''    sql = sql & "MS13DEN1, "             ' ����l13 Den1"
'''''    sql = sql & "MS13DEN2, "             ' ����l13 Den2"
'''''    sql = sql & "MS13DEN3, "             ' ����l13 Den3"
'''''    sql = sql & "MS13DEN4, "             ' ����l13 Den4"
'''''    sql = sql & "MS13DEN5, "             ' ����l13 Den5"
'''''    sql = sql & "MS14LDL1, "             ' ����l14 L/DL1"
'''''    sql = sql & "MS14LDL2, "             ' ����l14 L/DL2"
'''''    sql = sql & "MS14LDL3, "             ' ����l14 L/DL3"
'''''    sql = sql & "MS14LDL4, "             ' ����l14 L/DL4"
'''''    sql = sql & "MS14LDL5, "             ' ����l14 L/DL5"
'''''    sql = sql & "MS14DEN1, "             ' ����l14 Den1"
'''''    sql = sql & "MS14DEN2, "             ' ����l14 Den2"
'''''    sql = sql & "MS14DEN3, "             ' ����l14 Den3"
'''''    sql = sql & "MS14DEN4, "             ' ����l14 Den4"
'''''    sql = sql & "MS14DEN5, "             ' ����l14 Den5"
'''''    sql = sql & "MS15LDL1, "             ' ����l15 L/DL1"
'''''    sql = sql & "MS15LDL2, "             ' ����l15 L/DL2"
'''''    sql = sql & "MS15LDL3, "             ' ����l15 L/DL3"
'''''    sql = sql & "MS15LDL4, "             ' ����l15 L/DL4"
'''''    sql = sql & "MS15LDL5, "             ' ����l15 L/DL5"
'''''    sql = sql & "MS15DEN1, "             ' ����l15 Den1"
'''''    sql = sql & "MS15DEN2, "             ' ����l15 Den2"
'''''    sql = sql & "MS15DEN3, "             ' ����l15 Den3"
'''''    sql = sql & "MS15DEN4, "             ' ����l15 Den4"
'''''    sql = sql & "MS15DEN5, "              ' ����l15 Den5"
'''''    sql = sql & "REGDATE "               ' �o�^���t
'''''
'''''End Sub


''''''�����֐� GD���ю擾�p
'''''Private Function GD_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              GD As type_DBDRV_scmzc_fcmkc001c_GD, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function GD_Zisseki"
'''''
'''''    GD_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ006"
'''''    Set rs = Nothing
'''''
'''''    DoEvents
'''''    Call GD_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname)
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then
'''''        DoEvents
'''''        Call GD_ObjCpy(GD, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
'''''        rs.Close
'''''        Set rs = Nothing
'''''        If Siyou.HSXDENKU = "1" Or Siyou.HSXDVDKU = "1" Or Siyou.HSXLDLKU = "1" Then
'''''           If Samp.CRYINDGD = "6" Then       ' ���p���Ȃ�
'''''                DoEvents
'''''                Call GD_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_HIKITUGI2(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname, TorB)
'''''
'''''                DoEvents
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''
'''''                If rs.RecordCount <> 0 Then
'''''                    recCnt = rs.RecordCount
'''''                    For i = 1 To recCnt
'''''                        If i = 1 Then                                     ' ���ڂ͕ێ�
'''''                            DoEvents
'''''                            Call GD_ObjCpy(GD, rs)
'''''                            Exit For    '�P���R�[�h�ڂ�����OK
'''''                        Else
'''''                            If GD.POSITION = rs("POSITION") And GD.REGDATE < rs("REGDATE") Then   ' �O�̈ʒu�Ɠ�����������o�^���t���V�������̂��Ƃ�
'''''                                DoEvents
'''''                                Call GD_ObjCpy(GD, rs)
'''''                            End If
'''''                        End If
'''''
'''''                        rs.MoveNext
'''''                    Next
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                If Not rs Is Nothing Then
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''            End If  ' �����w����5 or 6 �Ȃ�
'''''        End If  ' �w���������Ă���
'''''    End If ' ���т�����
'''''
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    GD_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function



'''''Private Sub LT_ObjCpy(Lt As type_DBDRV_scmzc_fcmkc001c_LT, rs As OraDynaset)
'''''    With Lt
'''''        .CRYNUM = rs("CRYNUM")         ' �����ԍ�
'''''        .POSITION = rs("POSITION")     ' �ʒu
'''''        .SMPKBN = rs("SMPKBN")         ' �T���v���敪
'''''        .SMPLNO = rs("SMPLNO")         ' �T���v���m��
'''''        .SMPLUMU = rs("SMPLUMU")       ' �T���v���L��
'''''        .MEAS1 = rs("MEAS1")           ' ����l�P
'''''        .MEAS2 = rs("MEAS2")           ' ����l�Q
'''''        .MEAS3 = rs("MEAS3")           ' ����l�R
'''''        .MEAS4 = rs("MEAS4")           ' ����l�S
'''''        .MEAS5 = rs("MEAS5")           ' ����l�T
'''''        .TRANCOND = rs("TRANCOND")     ' ��������
'''''        .MEASPEAK = rs("MEASPEAK")     ' ����l �s�[�N�l
'''''        .CALCMEAS = rs("CALCMEAS")     ' �v�Z����
'''''        .REGDATE = rs("REGDATE")        '�@�o�^���t
'''''        .LTSPI = rs("HSXLTSPI")            '����ʒu�R�[�h
'''''    End With
'''''
'''''End Sub
'''''
'''''Private Sub LT_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "        ' �����ԍ�
'''''    sql = sql & "POSITION, "      ' �ʒu
'''''    sql = sql & "SMPKBN, "        ' �T���v���敪
'''''    sql = sql & "SMPLNO, "        ' �T���v���m��
'''''    sql = sql & "SMPLUMU, "       ' �T���v���L��
'''''    sql = sql & "MEAS1,"          ' ����l�P"
'''''    sql = sql & "MEAS2,"          ' ����l�Q"
'''''    sql = sql & "MEAS3,"          ' ����l�R"
'''''    sql = sql & "MEAS4,"          ' ����l�S"
'''''    sql = sql & "MEAS5,"          ' ����l�T"
'''''    sql = sql & "TRANCOND, "      ' ��������
'''''    sql = sql & "MEASPEAK, "      ' ����l �s�[�N�l
'''''    sql = sql & "CALCMEAS, "       ' �v�Z����
'''''    sql = sql & "REGDATE "        ' �o�^���t
'''''
'''''End Sub


''''''�����֐� ���C�t�^�C�����ю擾�p
'''''Private Function LT_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              Lt As type_DBDRV_scmzc_fcmkc001c_LT, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''    Dim hin As tFullHinban
'''''    Dim LTSPI As String
'''''
'''''    NothingFlag = False
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function LT_Zisseki"
'''''
'''''    ' ���C�t�^�C�����уe�[�u������l���擾
'''''    LT_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ007"
'''''
'''''    If NothingFlagStr <> vbNullString Then NothingFlagStr = "0"
'''''    With Lt
'''''        .CRYNUM = vbNullString
'''''        .POSITION = -1
'''''        .SMPLNO = -1
'''''    End With
'''''
'''''    '�T���v���ʒu��LT���т�����΁A����� (B���D��A�Ō�̎���)
'''''    sql = "select * from ("
'''''    sql = sql & "  select CRYNUM, POSITION, SMPKBN, TRANCOND, SMPLNO, SMPLUMU"
'''''    sql = sql & "  , MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS, REGDATE"
'''''    sql = sql & "  , (select HSXLTSPI from TBCME019 where HINBAN=LT.HINBAN and MNOREVNO=LT.REVNUM and FACTORY=LT.FACTORY and OPECOND=LT.OPECOND) as HSXLTSPI"
'''''    sql = sql & "  from TBCMJ007 LT"
'''''    sql = sql & "  where CRYNUM='" & Samp.CRYNUM & "' and POSITION=" & Samp.INGOTPOS
'''''    sql = sql & "  order by SMPKBN, TRANCNT desc"
'''''    sql = sql & ") where rownum=1"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    If rs.RecordCount > 0 Then
'''''        Call LT_ObjCpy(Lt, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''        LT_Zisseki = FUNCTION_RETURN_SUCCESS
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    '�u���b�N���ōł�������LT�d�l�����i�ԂƁA���̑���ʒu�����߂�
'''''    If DBDRV_getLtHinbanInBlock(Samp.CRYNUM, Samp.INGOTPOS, hin, LTSPI) = FUNCTION_RETURN_FAILURE Then
'''''        LT_Zisseki = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''    If LTSPI = vbNullString Then    '�u���b�N����LT�d�l�Ȃ�
'''''        If NothingFlagStr <> vbNullString Then NothingFlagStr = "1"
'''''        LT_Zisseki = FUNCTION_RETURN_SUCCESS
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    'LT���茋�ʂ��������� (���ꑪ��ʒu������΂���A�Ȃ���΂�茵�����������ʂ��A�Ȃ�ׂ��߂�����������)
'''''    sql = "select * from ("
'''''    sql = sql & "  select LT.CRYNUM, POSITION, TRANCOND, SMPKBN, SMPLNO, SMPLUMU"
'''''    sql = sql & "  , MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS, LT.REGDATE"
'''''    sql = sql & "  , SIYO.HSXLTSPI"
'''''    sql = sql & "  from TBCMJ007 LT, TBCME019 SIYO"
'''''    sql = sql & "  where LT.HINBAN=SIYO.HINBAN and LT.REVNUM=SIYO.MNOREVNO and LT.FACTORY=SIYO.FACTORY and LT.OPECOND=SIYO.OPECOND"
'''''    sql = sql & "    and LT.CRYNUM='" & Samp.CRYNUM & "'"
'''''    sql = sql & "    and POSITION>=" & Samp.INGOTPOS
'''''    sql = sql & "    and decode(SIYO.HSXLTSPI,' ','ZZ',SIYO.HSXLTSPI)<='" & LTSPI & "'"
'''''    sql = sql & "  order by decode(HSXLTSPI,'" & LTSPI & "',1,0) desc, POSITION, SMPKBN,TRANCNT desc"
'''''    sql = sql & ") where rownum=1"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    If rs.RecordCount > 0 Then
'''''        Call LT_ObjCpy(Lt, rs)
'''''    Else
'''''        If NothingFlagStr <> vbNullString Then NothingFlagStr = "1"
'''''    End If
'''''    rs.Close
'''''    Set rs = Nothing
'''''
'''''    LT_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    LT_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''�T�v      :
'''''Private Sub EPD_ObjCpy(EPD As type_DBDRV_scmzc_fcmkc001c_EPD, rs As OraDynaset)
'''''    With EPD
'''''        .CRYNUM = rs("CRYNUM")      ' �����ԍ�
'''''        .POSITION = rs("POSITION")  ' �ʒu
'''''        .SMPKBN = rs("SMPKBN")      ' �T���v���敪
'''''        .SMPLNO = rs("SMPLNO")      ' �T���v���m��
'''''        .SMPLUMU = rs("SMPLUMU")    ' �T���v���L��
'''''        .TRANCOND = rs("TRANCOND")  ' ��������
'''''        .MEASURE = rs("MEASURE")    ' ����l
'''''        .REGDATE = rs("REGDATE")    ' �o�^���t
'''''    End With
'''''End Sub
'''''
''''''�T�v      :
'''''Private Sub EPD_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "          ' �����ԍ�
'''''    sql = sql & "POSITION, "        ' �ʒu
'''''    sql = sql & "SMPKBN, "          ' �T���v���敪
'''''    sql = sql & "SMPLNO, "          ' �T���v���m��
'''''    sql = sql & "SMPLUMU, "         ' �T���v���L��
'''''    sql = sql & "TRANCOND, "        ' ��������
'''''    sql = sql & "MEASURE, "         ' ����l
'''''    sql = sql & "REGDATE "          ' �o�^���t
'''''End Sub


''''''�T�v      :�����֐� EPD���ю擾�p
'''''Private Function EPD_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              EPD As type_DBDRV_scmzc_fcmkc001c_EPD, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''    ' EPD���уe�[�u������l���擾
'''''    ' VECME010�i�u���b�N�Ǘ����������A���̃u���b�N�ɑ΂���T���v����\������r���[�j����T���v���敪���擾�iwhere �����ԍ��A�ʒu�j
'''''    ' �����ԍ��A�ʒu�A�T���v���敪�A�����񐔍ő�����������Ƃ����уe�[�u������l���擾
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function EPD_Zisseki"
'''''
'''''    EPD_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ001"
'''''    Set rs = Nothing
'''''
'''''    DoEvents
'''''    Call EPD_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname)
'''''
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then
'''''        DoEvents
'''''        Call EPD_ObjCpy(EPD, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
'''''        rs.Close
'''''        Set rs = Nothing
'''''
'''''
'''''        ' �������Ɏ��т�T���A�����ʒu�ɂ������ꍇ�@�V�������t�̕����擾����
'''''        DoEvents
'''''        Call EPD_SetBaseSQL(sql)
'''''        DoEvents
'''''        Call AddSQL_Down(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname)
'''''
'''''        DoEvents
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''        DoEvents
'''''
'''''        If rs.RecordCount <> 0 Then
'''''            recCnt = rs.RecordCount
'''''            For i = 1 To recCnt
'''''                If i = 1 Then                                     ' ���ڂ͕ێ�
'''''                    DoEvents
'''''                    Call EPD_ObjCpy(EPD, rs)
'''''                Else
'''''                    If EPD.POSITION = rs("POSITION") And EPD.REGDATE < rs("REGDATE") Then   ' �O�̈ʒu�Ɠ�����������o�^���t���V�������̂��Ƃ�
'''''                        DoEvents
'''''                        Call EPD_ObjCpy(EPD, rs)
'''''                    End If
'''''                End If
'''''
'''''                rs.MoveNext
'''''            Next
'''''        Else
'''''            NothingFlag = True
'''''        End If
'''''        If Not rs Is Nothing Then
'''''            rs.Close
'''''            Set rs = Nothing
'''''        End If
'''''    End If
'''''
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    EPD_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''�T�v      :�����֐� �e�[�u���uTBCMG002�v��������ɂ��������R�[�h�𒊏o����
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :records       ,O  ,typ_TBCMG002 ,���o���R�[�h
''''''          :[sqlWhere]    ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
''''''          :[sqlOrder]    ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
''''''          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
''''''����      :
''''''����      :
'''''Private Function Kounyu_Zisseki(records As typ_TBCMG002, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
'''''Dim sql As String       'SQL�S��
'''''Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
'''''Dim rs As OraDynaset    'RecordSet
'''''Dim recCnt As Long      '���R�[�h��
'''''Dim i As Long
'''''
'''''    ''SQL��g�ݗ��Ă�
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function Kounyu_Zisseki"
'''''
'''''    Kounyu_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    sqlBase = "Select CRYNUM, TRANCNT, KRPROCCD, PROCCODE, HINBAN, MNOREVNO, FACTORY, OPECOND, REPCCL, RBATCHNO, DMTOP1, DMTOP2," & _
'''''              " DMTAIL1, DMTAIL2, NCHDPTH1, NCHDPTH2, UPLENGTH, SXLPOS, BLKLEN, BLKWGHT, CMPTOP1, CMPTOP2, CMPTOP3, CMPTOP4," & _
'''''              " CMPTOP5, CMPTOPR, CMPTAIL1, CMPTAIL2, CMPTAIL3, CMPTAIL4, CMPTAIL5, CMPTAILR, OITOP1, OITOP2, OITOP3, OITOP4," & _
'''''              " OITOP5, OITOPR, OITAIL1, OITAIL2, OITAIL3, OITAIL4, OITAIL5, OITAILR, CSTOP, CSTAIL, LD1TOPMX, LD1TOPAV, LD1TAILM," & _
'''''              " LD1TAILA, LD2TOPMM, LD2TOPAV, LD2TAILM, LD2TAILA, BMDTOPMX, BMDTOPAV, BMDTAILM, BMDTAILA, GD1TOP, GD1TAIL," & _
'''''              " GD2TOP, GD2TAIL, DIA1TOP, DIA1TAIL, DIA2TOP, DIA2TAIL, LTFTOP, LTFTAIL, EPD, TSTAFFID, REGDATE, KSTAFFID," & _
'''''              " UPDDATE, SENDFLAG, SENDDATE "
'''''    sqlBase = sqlBase & "From TBCMG002"
'''''    sql = sqlBase
'''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
'''''        sql = sql & sqlWhere & sqlOrder
'''''    End If
'''''
'''''    ''�f�[�^�𒊏o����
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    If rs.RecordCount = 0 Then
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    ''���o���ʂ��i�[����
'''''    With records
'''''        .CRYNUM = rs("CRYNUM")           ' �����ԍ�
'''''        .TRANCNT = rs("TRANCNT")         ' ������
'''''        .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
'''''        .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
'''''        .hinban = rs("HINBAN")           ' �i��
'''''        .mnorevno = rs("MNOREVNO")       ' ���i�ԍ������ԍ�
'''''        .factory = rs("FACTORY")         ' �H��
'''''        .opecond = rs("OPECOND")         ' ���Ə���
'''''        .REPCCL = rs("REPCCL")           ' �������敪
'''''        .RBATCHNO = rs("RBATCHNO")       ' �F�o�b�`�m��
'''''        .DMTOP1 = rs("DMTOP1")           ' ���a�s�n�o�P
'''''        .DMTOP2 = rs("DMTOP2")           ' ���a�s�n�o�Q
'''''        .DMTAIL1 = rs("DMTAIL1")         ' ���a�s�`�h�k�P
'''''        .DMTAIL2 = rs("DMTAIL2")         ' ���a�s�`�h�k�Q
'''''        .NCHDPTH1 = rs("NCHDPTH1")       ' �m�b�`�[���P
'''''        .NCHDPTH2 = rs("NCHDPTH2")       ' �m�b�`�[���Q
'''''        .UPLENGTH = rs("UPLENGTH")       ' ���グ��
'''''        .SXLPOS = rs("SXLPOS")           ' �r�w�k�ʒu
'''''        .BlkLen = rs("BLKLEN")           ' �u���b�N����
'''''        .BLKWGHT = rs("BLKWGHT")         ' �u���b�N�d��
'''''        .CMPTOP1 = rs("CMPTOP1")         ' ���RTOP�@�P
'''''        .CMPTOP2 = rs("CMPTOP2")         ' ���RTOP�@�Q
'''''        .CMPTOP3 = rs("CMPTOP3")         ' ���RTOP�@�R
'''''        .CMPTOP4 = rs("CMPTOP4")         ' ���RTOP�@�S
'''''        .CMPTOP5 = rs("CMPTOP5")         ' ���RTOP�@�T
'''''        .CMPTOPR = rs("CMPTOPR")         ' ���RTOP�@RRG
'''''        .CMPTAIL1 = rs("CMPTAIL1")       ' ���RTAIL�@�P
'''''        .CMPTAIL2 = rs("CMPTAIL2")       ' ���RTAIL�@�Q
'''''        .CMPTAIL3 = rs("CMPTAIL3")       ' ���RTAIL�@�R
'''''        .CMPTAIL4 = rs("CMPTAIL4")       ' ���RTAIL�@�S
'''''        .CMPTAIL5 = rs("CMPTAIL5")       ' ���RTAIL�@�T
'''''        .CMPTAILR = rs("CMPTAILR")       ' ���RTAIL�@RRG
'''''        .OITOP1 = rs("OITOP1")           ' Oi�@TOP�@�P
'''''        .OITOP2 = rs("OITOP2")           ' Oi�@TOP�@�Q
'''''        .OITOP3 = rs("OITOP3")           ' Oi�@TOP�@�R
'''''        .OITOP4 = rs("OITOP4")           ' Oi�@TOP�@�S
'''''        .OITOP5 = rs("OITOP5")           ' Oi�@TOP�@�T
'''''        .OITOPR = rs("OITOPR")           ' Oi�@TOP�@ROG
'''''        .OITAIL1 = rs("OITAIL1")         ' Oi�@TAIL�@�P
'''''        .OITAIL2 = rs("OITAIL2")         ' Oi�@TAIL�@�Q
'''''        .OITAIL3 = rs("OITAIL3")         ' Oi�@TAIL�@�R
'''''        .OITAIL4 = rs("OITAIL4")         ' Oi�@TAIL�@�S
'''''        .OITAIL5 = rs("OITAIL5")         ' Oi�@TAIL�@�T
'''''        .OITAILR = rs("OITAILR")         ' Oi�@TAIL�@ROG
'''''        .CSTOP = rs("CSTOP")             ' Cs�@TOP
'''''        .CSTAIL = rs("CSTAIL")           ' Cs�@TAIL
'''''        .LD1TOPMX = rs("LD1TOPMX")       ' LD-1�@TOP�@MAX
'''''        .LD1TOPAV = rs("LD1TOPAV")       ' LD-1�@TOP�@AVE
'''''        .LD1TAILM = rs("LD1TAILM")       ' LD-1�@TAIL�@MAX
'''''        .LD1TAILA = rs("LD1TAILA")       ' LD-1�@TAIL�@AVE
'''''        .LD2TOPMM = rs("LD2TOPMM")       ' LD-2�@TOP�@MAX
'''''        .LD2TOPAV = rs("LD2TOPAV")       ' LD-2�@TOP�@AVE
'''''        .LD2TAILM = rs("LD2TAILM")       ' LD-2�@TAIL�@MAX
'''''        .LD2TAILA = rs("LD2TAILA")       ' LD-2�@TAIL�@AVE
'''''        .BMDTOPMX = rs("BMDTOPMX")       ' BMD�@TOP�@MAX
'''''        .BMDTOPAV = rs("BMDTOPAV")       ' BMD�@TOP�@AVE
'''''        .BMDTAILM = rs("BMDTAILM")       ' BMD�@TAIL�@MAX
'''''        .BMDTAILA = rs("BMDTAILA")       ' BMD�@TAIL�@AVE
'''''        .GD1TOP = rs("GD1TOP")           ' GD1 TOP
'''''        .GD1TAIL = rs("GD1TAIL")         ' GD1 TAIL
'''''        .GD2TOP = rs("GD2TOP")           ' GD2 TOP
'''''        .GD2TAIL = rs("GD2TAIL")         ' GD2 TAIL
'''''        .DIA1TOP = rs("DIA1TOP")         ' DIA1 TOP
'''''        .DIA1TAIL = rs("DIA1TAIL")       ' DIA1 TAIL
'''''        .DIA2TOP = rs("DIA2TOP")         ' DIA2 TOP
'''''        .DIA2TAIL = rs("DIA2TAIL")       ' DIA2 TAIL
'''''        .LTFTOP = rs("LTFTOP")           ' LIFETIME from TOP
'''''        .LTFTAIL = rs("LTFTAIL")         ' LIFETIME from TAIL
'''''        .EPD = rs("EPD")                 ' EPD
'''''        .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
'''''        .REGDATE = rs("REGDATE")         ' �o�^���t
'''''        .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
'''''        .UPDDATE = rs("UPDDATE")         ' �X�V���t
'''''        .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
'''''        .SENDDATE = rs("SENDDATE")       ' ���M���t
'''''    End With
'''''    rs.Close
'''''
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    Kounyu_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function



'�T�v      :�����֐� �i�ԁA�d�l���擾����
Public Function getHinSiyou30(inBlockID As String, Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN

    Dim sql         As String
    Dim rs          As OraDynaset
    Dim recCnt      As Integer
    Dim i           As Long
    Dim Jiltuseki   As Judg_Kakou

    '�i�ԁASXL�d�l����f�[�^�̎擾
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 start

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function getHinSiyou"

    getHinSiyou30 = FUNCTION_RETURN_SUCCESS

    sql = "select "
    sql = sql & "distinct "                   ''�d���f�[�^�폜  2003/09/09 ooba
    sql = sql & "BH.E040CRYNUM, "             ' �����ԍ�
    sql = sql & "BH.E040INGOTPOS, "           ' �������J�n�ʒu
    sql = sql & "BH.E040LENGTH, "             ' ����
    sql = sql & "BH.E041HINBAN, "             ' �i��
    sql = sql & "BH.E041REVNUM, "             ' ���i�ԍ������ԍ�
    sql = sql & "BH.E041FACTORY, "            ' �H��
    sql = sql & "BH.E041OPECOND, "            ' ���Ə���

    sql = sql & "BH.E037PRODCOND, "           ' �������
    sql = sql & "BH.E037PGID, "               ' �o�f�|�h�c
    sql = sql & "BH.E037UPLENGTH, "           ' ���グ����
    sql = sql & "BH.E037FREELENG, "           ' �t���[��
    sql = sql & "BH.E037DIAMETER, "           ' ���a
    sql = sql & "BH.E037CHARGE, "             ' �`���[�W��
    sql = sql & "BH.E037SEED, "               ' �V�[�h
    sql = sql & "BH.E037ADDDPPOS, "           ' �ǉ��h�[�v�ʒu

    sql = sql & "S.E018HSXTYPE, "             ' �i�r�w�^�C�v
    sql = sql & "S.E018HSXD1CEN, "            ' �i�r�w���a�P���S
    sql = sql & "S.E018HSXCDIR, "             ' �i�r�w�����ʕ���

    sql = sql & "S.E018HSXRMIN, "             ' �i�r�w���R����
    sql = sql & "S.E018HSXRMAX, "             ' �i�r�w���R���
    sql = sql & "S.E018HSXRAMIN, "            ' �i�r�w���R���ω���
    sql = sql & "S.E018HSXRAMAX, "            ' �i�r�w���R���Ϗ��
    sql = sql & "S.E018HSXRMCAL, "            ' �i�r�w���R�ʓ��v�Z�@�@'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
    sql = sql & "S.E018HSXRMBNP, "            ' �i�r�w���R�ʓ����z
    sql = sql & "S.E018HSXRSPOH, "            ' �i�r�w���R����ʒu�Q��
    sql = sql & "S.E018HSXRSPOT, "            ' �i�r�w���R����ʒu�Q�_
    sql = sql & "S.E018HSXRSPOI, "            ' �i�r�w���R����ʒu�Q��
    sql = sql & "S.E018HSXRHWYT, "            ' �i�r�w���R�ۏؕ��@�Q��
    sql = sql & "S.E018HSXRHWYS, "            ' �i�r�w���R�ۏؕ��@�Q��

    sql = sql & "S.E019HSXONMIN, "            ' �i�r�w�_�f�Z�x����
    sql = sql & "S.E019HSXONMAX, "            ' �i�r�w�_�f�Z�x���
    sql = sql & "S.E019HSXONAMN, "            ' �i�r�w�_�f�Z�x���ω���
    sql = sql & "S.E019HSXONAMX, "            ' �i�r�w�_�f�Z�x���Ϗ��
    sql = sql & "S.E019HSXONMCL, "            ' �i�r�w�_�f�Z�x�ʓ��v�Z�@�@'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
    sql = sql & "S.E019HSXONMBP, "            ' �i�r�w�_�f�Z�x�ʓ����z
    sql = sql & "S.E019HSXONSPH, "            ' �i�r�w�_�f�Z�x����ʒu�Q��
    sql = sql & "S.E019HSXONSPT, "            ' �i�r�w�_�f�Z�x����ʒu�Q�_
    sql = sql & "S.E019HSXONSPI, "            ' �i�r�w�_�f�Z�x����ʒu�Q��
    sql = sql & "S.E019HSXONHWT, "            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    sql = sql & "S.E019HSXONHWS, "            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��

    sql = sql & "S.E020HSXBM1AN, "            ' �i�r�w�a�l�c�P���ω���
    sql = sql & "S.E020HSXBM1AX, "            ' �i�r�w�a�l�c�P���Ϗ��
    sql = sql & "S.E020HSXBM2AN, "            ' �i�r�w�a�l�c�Q���ω���
    sql = sql & "S.E020HSXBM2AX, "            ' �i�r�w�a�l�c�Q���Ϗ��
    sql = sql & "S.E020HSXBM3AN, "            ' �i�r�w�a�l�c�R���ω���
    sql = sql & "S.E020HSXBM3AX, "            ' �i�r�w�a�l�c�R���Ϗ��
    sql = sql & "S.E020HSXBM1SH, "            ' �i�r�w�a�l�c�P����ʒu�Q��
    sql = sql & "S.E020HSXBM1ST, "            ' �i�r�w�a�l�c�P����ʒu�Q�_
    sql = sql & "S.E020HSXBM1SR, "            ' �i�r�w�a�l�c�P����ʒu�Q��
    sql = sql & "S.E020HSXBM1HT, "            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    sql = sql & "S.E020HSXBM1HS, "            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    sql = sql & "S.E020HSXBM2SH, "            ' �i�r�w�a�l�c�Q����ʒu�Q��
    sql = sql & "S.E020HSXBM2ST, "            ' �i�r�w�a�l�c�Q����ʒu�Q�_
    sql = sql & "S.E020HSXBM2SR, "            ' �i�r�w�a�l�c�Q����ʒu�Q��
    sql = sql & "S.E020HSXBM2HT, "            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    sql = sql & "S.E020HSXBM2HS, "            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    sql = sql & "S.E020HSXBM3SH, "            ' �i�r�w�a�l�c�R����ʒu�Q��
    sql = sql & "S.E020HSXBM3ST, "            ' �i�r�w�a�l�c�R����ʒu�Q�_
    sql = sql & "S.E020HSXBM3SR, "            ' �i�r�w�a�l�c�R����ʒu�Q��
    sql = sql & "S.E020HSXBM3HT, "            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
    sql = sql & "S.E020HSXBM3HS, "            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��

    sql = sql & "S.E020HSXOF1AX, "            ' �i�r�w�n�r�e�P���Ϗ��
    sql = sql & "S.E020HSXOF1MX, "            ' �i�r�w�n�r�e�P���
    sql = sql & "S.E020HSXOF2AX, "            ' �i�r�w�n�r�e�Q���Ϗ��
    sql = sql & "S.E020HSXOF2MX, "            ' �i�r�w�n�r�e�Q���
    sql = sql & "S.E020HSXOF3AX, "            ' �i�r�w�n�r�e�R���Ϗ��
    sql = sql & "S.E020HSXOF3MX, "            ' �i�r�w�n�r�e�R���
    sql = sql & "S.E020HSXOF4AX, "            ' �i�r�w�n�r�e�S���Ϗ��
    sql = sql & "S.E020HSXOF4MX, "            ' �i�r�w�n�r�e�S���
    sql = sql & "S.E020HSXOF1SH, "            ' �i�r�w�n�r�e�P����ʒu�Q��
    sql = sql & "S.E020HSXOF1ST, "            ' �i�r�w�n�r�e�P����ʒu�Q�_
    sql = sql & "S.E020HSXOF1SR, "            ' �i�r�w�n�r�e�P����ʒu�Q��
    sql = sql & "S.E020HSXOF1HT, "            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF1HS, "            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF2SH, "            ' �i�r�w�n�r�e�Q����ʒu�Q��
    sql = sql & "S.E020HSXOF2ST, "            ' �i�r�w�n�r�e�Q����ʒu�Q�_
    sql = sql & "S.E020HSXOF2SR, "            ' �i�r�w�n�r�e�Q����ʒu�Q��
    sql = sql & "S.E020HSXOF2HT, "            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF2HS, "            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF3SH, "            ' �i�r�w�n�r�e�R����ʒu�Q��
    sql = sql & "S.E020HSXOF3ST, "            ' �i�r�w�n�r�e�R����ʒu�Q�_
    sql = sql & "S.E020HSXOF3SR, "            ' �i�r�w�n�r�e�R����ʒu�Q��
    sql = sql & "S.E020HSXOF3HT, "            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF3HS, "            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF4SH, "            ' �i�r�w�n�r�e�S����ʒu�Q��
    sql = sql & "S.E020HSXOF4ST, "            ' �i�r�w�n�r�e�S����ʒu�Q�_
    sql = sql & "S.E020HSXOF4SR, "            ' �i�r�w�n�r�e�S����ʒu�Q��
    sql = sql & "S.E020HSXOF4HT, "            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF4HS, "            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF1NS, "            ' �i�r�w�n�r�e�P�M�����@
    sql = sql & "S.E020HSXOF2NS, "            ' �i�r�w�n�r�e�Q�M�����@
    sql = sql & "S.E020HSXOF3NS, "            ' �i�r�w�n�r�e�R�M�����@
    sql = sql & "S.E020HSXOF4NS, "            ' �i�r�w�n�r�e�S�M�����@
    sql = sql & "S.E020HSXBM1NS, "            ' �i�r�w�a�l�c�P�M�����@
    sql = sql & "S.E020HSXBM2NS, "            ' �i�r�w�a�l�c�Q�M�����@
    sql = sql & "S.E020HSXBM3NS, "            ' �i�r�w�a�l�c�R�M�����@

    sql = sql & "S.E019HSXCNMIN, "            ' �i�r�w�Y�f�Z�x����
    sql = sql & "S.E019HSXCNMAX, "            ' �i�r�w�Y�f�Z�x���
    sql = sql & "S.E019HSXCNSPH, "            ' �i�r�w�Y�f�Z�x����ʒu�Q��
    sql = sql & "S.E019HSXCNSPT, "            ' �i�r�w�Y�f�Z�x����ʒu�Q�_
    sql = sql & "S.E019HSXCNSPI, "            ' �i�r�w�Y�f�Z�x����ʒu�Q��
    sql = sql & "S.E019HSXCNHWT, "            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    sql = sql & "S.E019HSXCNHWS, "            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��

    sql = sql & "S.E020HSXDENMX, "            ' �i�r�w�c�������
    sql = sql & "S.E020HSXDENMN, "            ' �i�r�w�c��������
    sql = sql & "S.E020HSXLDLMX, "            ' �i�r�w�k�^�c�k���
    sql = sql & "S.E020HSXLDLMN, "            ' �i�r�w�k�^�c�k����
    sql = sql & "S.E020HSXDVDMXN, "           ' �i�r�w�c�u�c�Q���   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sql = sql & "S.E020HSXDVDMNN, "           ' �i�r�w�c�u�c�Q����   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sql = sql & "S.E020HSXDENHT, "            ' �i�r�w�c�����ۏؕ��@�Q��
    sql = sql & "S.E020HSXDENHS, "            ' �i�r�w�c�����ۏؕ��@�Q��
    sql = sql & "S.E020HSXLDLHT, "            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    sql = sql & "S.E020HSXLDLHS, "            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    sql = sql & "S.E020HSXDVDHT, "            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    sql = sql & "S.E020HSXDVDHS, "            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    sql = sql & "S.E020HSXDENKU, "            ' �i�r�w�c���������L��
    sql = sql & "S.E020HSXDVDKU, "            ' �i�r�w�c�u�c�Q�����L��
    sql = sql & "S.E020HSXLDLKU, "            ' �i�r�w�k�^�c�k�����L��

    sql = sql & "S.E019HSXLTMIN, "            ' �i�r�w�k�^�C������
    sql = sql & "S.E019HSXLTMAX, "            ' �i�r�w�k�^�C�����
    sql = sql & "S.E019HSXLTSPH, "            ' �i�r�w�k�^�C������ʒu�Q��
    sql = sql & "S.E019HSXLTSPT, "            ' �i�r�w�k�^�C������ʒu�Q�_
    sql = sql & "S.E019HSXLTSPI, "            ' �i�r�w�k�^�C������ʒu�Q��
    sql = sql & "S.E019HSXLTHWT, "            ' �i�r�w�k�^�C���ۏؕ��@�Q��
    sql = sql & "S.E019HSXLTHWS, "            ' �i�r�w�k�^�C���ۏؕ��@�Q��
    sql = sql & "U.EPDUP, "                   ' EPD ���
    sql = sql & "U.TOPREG, "                  ' TOP�K��
    sql = sql & "U.TAILREG, "                 ' TAIL�K��
    sql = sql & "U.BTMSPRT, "                 ' �{�g���͏o�K��

' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
    sql = sql & "S.E020HSXOSF1PTK, "          ' �i�r�w�n�r�e�P�p�^���敪
    sql = sql & "S.E020HSXOSF2PTK, "          ' �i�r�w�n�r�e�Q�p�^���敪
    sql = sql & "S.E020HSXOSF3PTK, "          ' �i�r�w�n�r�e�R�p�^���敪
    sql = sql & "S.E020HSXOSF4PTK, "          ' �i�r�w�n�r�e�S�p�^���敪
    sql = sql & "S.E020HSXBMD1MBP, "          ' �i�r�w�a�l�c�P�ʓ����z
    sql = sql & "S.E020HSXBMD2MBP, "          ' �i�r�w�a�l�c�Q�ʓ����z
    sql = sql & "S.E020HSXBMD3MBP  "          ' �i�r�w�a�l�c�R�ʓ����z
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura


' �������ʒu�ŕi�ԁi�����i�Ԃ̏ꍇ�j���\�[�g���邽�߂Ɏ擾����K�v�L��
    sql = sql & ", BH.E041INGOTPOS "

    sql = sql & " from VECME009 BH, VECME001 S, TBCME036 U "
    ' �i��TOP�̎擾
    sql = sql & " where BH.E040BLOCKID='" & inBlockID & "' " & _
                "   and S.E018HINBAN  =BH.E041HINBAN  and S.E018MNOREVNO=BH.E041REVNUM " & _
                "   and S.E018FACTORY =BH.E041FACTORY and S.E018OPECOND =BH.E041OPECOND" & _
                "   and U.HINBAN      =BH.E041HINBAN  and U.MNOREVNO    =BH.E041REVNUM " & _
                "   and U.FACTORY     =BH.E041FACTORY and U.OPECOND     =BH.E041OPECOND"

' �������ʒu�ŕi�ԁi�����i�Ԃ̏ꍇ�j���\�[�g
    sql = sql & " order by BH.E041INGOTPOS ASC"


    ''-------------���R�����g���y�u���b�N���S�i�Ԏd�l�擾�z 2003/09/09 ooba-------------
'                " and BH.E041INGOTPOS=ANY(select min(E041INGOTPOS) from VECME009 where E040BLOCKID='" & inBlockID & "' ) "

'    sql = sql & " union all "
'    sql = sql & " select "
'    sql = sql & "BH.E040CRYNUM, "             ' �����ԍ�
'    sql = sql & "BH.E040INGOTPOS, "           ' �������J�n�ʒu
'    sql = sql & "BH.E040LENGTH, "             ' ����
'    sql = sql & "BH.E041HINBAN, "             ' �i��
'    sql = sql & "BH.E041REVNUM, "             ' ���i�ԍ������ԍ�
'    sql = sql & "BH.E041FACTORY, "            ' �H��
'    sql = sql & "BH.E041OPECOND, "            ' ���Ə���
'
'    sql = sql & "BH.E037PRODCOND, "           ' �������
'    sql = sql & "BH.E037PGID, "               ' �o�f�|�h�c
'    sql = sql & "BH.E037UPLENGTH, "           ' ���グ����
'    sql = sql & "BH.E037FREELENG, "           ' �t���[��
'    sql = sql & "BH.E037DIAMETER, "           ' ���a
'    sql = sql & "BH.E037CHARGE, "             ' �`���[�W��
'    sql = sql & "BH.E037SEED, "               ' �V�[�h
'    sql = sql & "BH.E037ADDDPPOS, "           ' �ǉ��h�[�v�ʒu
'
'    sql = sql & "S.E018HSXTYPE, "             ' �i�r�w�^�C�v
'    sql = sql & "S.E018HSXD1CEN, "            ' �i�r�w���a�P���S
'    sql = sql & "S.E018HSXCDIR, "             ' �i�r�w�����ʕ���
'
'    sql = sql & "S.E018HSXRMIN, "             ' �i�r�w���R����
'    sql = sql & "S.E018HSXRMAX, "             ' �i�r�w���R���
'    sql = sql & "S.E018HSXRAMIN, "            ' �i�r�w���R���ω���
'    sql = sql & "S.E018HSXRAMAX, "            ' �i�r�w���R���Ϗ��
'    sql = sql & "S.E018HSXRMCAL, "            ' �i�r�w���R�ʓ��v�Z    '' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
'    sql = sql & "S.E018HSXRMBNP, "            ' �i�r�w���R�ʓ����z
'    sql = sql & "S.E018HSXRSPOH, "            ' �i�r�w���R����ʒu�Q��
'    sql = sql & "S.E018HSXRSPOT, "            ' �i�r�w���R����ʒu�Q�_
'    sql = sql & "S.E018HSXRSPOI, "            ' �i�r�w���R����ʒu�Q��
'    sql = sql & "S.E018HSXRHWYT, "            ' �i�r�w���R�ۏؕ��@�Q��
'    sql = sql & "S.E018HSXRHWYS, "            ' �i�r�w���R�ۏؕ��@�Q��
'
'    sql = sql & "S.E019HSXONMIN, "            ' �i�r�w�_�f�Z�x����
'    sql = sql & "S.E019HSXONMAX, "            ' �i�r�w�_�f�Z�x���
'    sql = sql & "S.E019HSXONAMN, "            ' �i�r�w�_�f�Z�x���ω���
'    sql = sql & "S.E019HSXONAMX, "            ' �i�r�w�_�f�Z�x���Ϗ��
'    sql = sql & "S.E019HSXONMCL, "            ' �i�r�w�_�f�Z�x�ʓ��v�Z�@ '' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
'    sql = sql & "S.E019HSXONMBP, "            ' �i�r�w�_�f�Z�x�ʓ����z
'    sql = sql & "S.E019HSXONSPH, "            ' �i�r�w�_�f�Z�x����ʒu�Q��
'    sql = sql & "S.E019HSXONSPT, "            ' �i�r�w�_�f�Z�x����ʒu�Q�_
'    sql = sql & "S.E019HSXONSPI, "            ' �i�r�w�_�f�Z�x����ʒu�Q��
'    sql = sql & "S.E019HSXONHWT, "            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
'    sql = sql & "S.E019HSXONHWS, "            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
'
'    sql = sql & "S.E020HSXBM1AN, "            ' �i�r�w�a�l�c�P���ω���
'    sql = sql & "S.E020HSXBM1AX, "            ' �i�r�w�a�l�c�P���Ϗ��
'    sql = sql & "S.E020HSXBM2AN, "            ' �i�r�w�a�l�c�Q���ω���
'    sql = sql & "S.E020HSXBM2AX, "            ' �i�r�w�a�l�c�Q���Ϗ��
'    sql = sql & "S.E020HSXBM3AN, "            ' �i�r�w�a�l�c�R���ω���
'    sql = sql & "S.E020HSXBM3AX, "            ' �i�r�w�a�l�c�R���Ϗ��
'    sql = sql & "S.E020HSXBM1SH, "            ' �i�r�w�a�l�c�P����ʒu�Q��
'    sql = sql & "S.E020HSXBM1ST, "            ' �i�r�w�a�l�c�P����ʒu�Q�_
'    sql = sql & "S.E020HSXBM1SR, "            ' �i�r�w�a�l�c�P����ʒu�Q��
'    sql = sql & "S.E020HSXBM1HT, "            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXBM1HS, "            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXBM2SH, "            ' �i�r�w�a�l�c�Q����ʒu�Q��
'    sql = sql & "S.E020HSXBM2ST, "            ' �i�r�w�a�l�c�Q����ʒu�Q�_
'    sql = sql & "S.E020HSXBM2SR, "            ' �i�r�w�a�l�c�Q����ʒu�Q��
'    sql = sql & "S.E020HSXBM2HT, "            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXBM2HS, "            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXBM3SH, "            ' �i�r�w�a�l�c�R����ʒu�Q��
'    sql = sql & "S.E020HSXBM3ST, "            ' �i�r�w�a�l�c�R����ʒu�Q�_
'    sql = sql & "S.E020HSXBM3SR, "            ' �i�r�w�a�l�c�R����ʒu�Q��
'    sql = sql & "S.E020HSXBM3HT, "            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXBM3HS, "            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
'
'    sql = sql & "S.E020HSXOF1AX, "            ' �i�r�w�n�r�e�P���Ϗ��
'    sql = sql & "S.E020HSXOF1MX, "            ' �i�r�w�n�r�e�P���
'    sql = sql & "S.E020HSXOF2AX, "            ' �i�r�w�n�r�e�Q���Ϗ��
'    sql = sql & "S.E020HSXOF2MX, "            ' �i�r�w�n�r�e�Q���
'    sql = sql & "S.E020HSXOF3AX, "            ' �i�r�w�n�r�e�R���Ϗ��
'    sql = sql & "S.E020HSXOF3MX, "            ' �i�r�w�n�r�e�R���
'    sql = sql & "S.E020HSXOF4AX, "            ' �i�r�w�n�r�e�S���Ϗ��
'    sql = sql & "S.E020HSXOF4MX, "            ' �i�r�w�n�r�e�S���
'    sql = sql & "S.E020HSXOF1SH, "            ' �i�r�w�n�r�e�P����ʒu�Q��
'    sql = sql & "S.E020HSXOF1ST, "            ' �i�r�w�n�r�e�P����ʒu�Q�_
'    sql = sql & "S.E020HSXOF1SR, "            ' �i�r�w�n�r�e�P����ʒu�Q��
'    sql = sql & "S.E020HSXOF1HT, "            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXOF1HS, "            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXOF2SH, "            ' �i�r�w�n�r�e�Q����ʒu�Q��
'    sql = sql & "S.E020HSXOF2ST, "            ' �i�r�w�n�r�e�Q����ʒu�Q�_
'    sql = sql & "S.E020HSXOF2SR, "            ' �i�r�w�n�r�e�Q����ʒu�Q��
'    sql = sql & "S.E020HSXOF2HT, "            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXOF2HS, "            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXOF3SH, "            ' �i�r�w�n�r�e�R����ʒu�Q��
'    sql = sql & "S.E020HSXOF3ST, "            ' �i�r�w�n�r�e�R����ʒu�Q�_
'    sql = sql & "S.E020HSXOF3SR, "            ' �i�r�w�n�r�e�R����ʒu�Q��
'    sql = sql & "S.E020HSXOF3HT, "            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXOF3HS, "            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXOF4SH, "            ' �i�r�w�n�r�e�S����ʒu�Q��
'    sql = sql & "S.E020HSXOF4ST, "            ' �i�r�w�n�r�e�S����ʒu�Q�_
'    sql = sql & "S.E020HSXOF4SR, "            ' �i�r�w�n�r�e�S����ʒu�Q��
'    sql = sql & "S.E020HSXOF4HT, "            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXOF4HS, "            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXOF1NS, "            ' �i�r�w�n�r�e�P�M�����@
'    sql = sql & "S.E020HSXOF2NS, "            ' �i�r�w�n�r�e�Q�M�����@
'    sql = sql & "S.E020HSXOF3NS, "            ' �i�r�w�n�r�e�R�M�����@
'    sql = sql & "S.E020HSXOF4NS, "            ' �i�r�w�n�r�e�S�M�����@
'    sql = sql & "S.E020HSXBM1NS, "            ' �i�r�w�a�l�c�P�M�����@
'    sql = sql & "S.E020HSXBM2NS, "            ' �i�r�w�a�l�c�Q�M�����@
'    sql = sql & "S.E020HSXBM3NS, "            ' �i�r�w�a�l�c�R�M�����@
'
'    sql = sql & "S.E019HSXCNMIN, "            ' �i�r�w�Y�f�Z�x����
'    sql = sql & "S.E019HSXCNMAX, "            ' �i�r�w�Y�f�Z�x���
'    sql = sql & "S.E019HSXCNSPH, "            ' �i�r�w�Y�f�Z�x����ʒu�Q��
'    sql = sql & "S.E019HSXCNSPT, "            ' �i�r�w�Y�f�Z�x����ʒu�Q�_
'    sql = sql & "S.E019HSXCNSPI, "            ' �i�r�w�Y�f�Z�x����ʒu�Q��
'    sql = sql & "S.E019HSXCNHWT, "            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
'    sql = sql & "S.E019HSXCNHWS, "            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
'
'    sql = sql & "S.E020HSXDENMX, "            ' �i�r�w�c�������
'    sql = sql & "S.E020HSXDENMN, "            ' �i�r�w�c��������
'    sql = sql & "S.E020HSXLDLMX, "            ' �i�r�w�k�^�c�k���
'    sql = sql & "S.E020HSXLDLMN, "            ' �i�r�w�k�^�c�k����
'    sql = sql & "S.E020HSXDVDMXN, "           ' �i�r�w�c�u�c�Q���   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
'    sql = sql & "S.E020HSXDVDMNN, "           ' �i�r�w�c�u�c�Q����   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
'    sql = sql & "S.E020HSXDENHT, "            ' �i�r�w�c�����ۏؕ��@�Q��
'    sql = sql & "S.E020HSXDENHS, "            ' �i�r�w�c�����ۏؕ��@�Q��
'    sql = sql & "S.E020HSXLDLHT, "            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXLDLHS, "            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXDVDHT, "            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXDVDHS, "            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
'    sql = sql & "S.E020HSXDENKU, "            ' �i�r�w�c���������L��
'    sql = sql & "S.E020HSXDVDKU, "            ' �i�r�w�c�u�c�Q�����L��
'    sql = sql & "S.E020HSXLDLKU, "            ' �i�r�w�k�^�c�k�����L��
'
'    sql = sql & "S.E019HSXLTMIN, "            ' �i�r�w�k�^�C������
'    sql = sql & "S.E019HSXLTMAX, "            ' �i�r�w�k�^�C�����
'    sql = sql & "S.E019HSXLTSPH, "            ' �i�r�w�k�^�C������ʒu�Q��
'    sql = sql & "S.E019HSXLTSPT, "            ' �i�r�w�k�^�C������ʒu�Q�_
'    sql = sql & "S.E019HSXLTSPI, "            ' �i�r�w�k�^�C������ʒu�Q��
'    sql = sql & "S.E019HSXLTHWT, "            ' �i�r�w�k�^�C���ۏؕ��@�Q��
'    sql = sql & "S.E019HSXLTHWS, "            ' �i�r�w�k�^�C���ۏؕ��@�Q��
'    sql = sql & "U.EPDUP, "                   ' EPD ���
'    sql = sql & "U.TOPREG, "                  ' TOP�K��
'    sql = sql & "U.TAILREG, "                 ' TAIL�K��
'    sql = sql & "U.BTMSPRT, "                 ' �{�g���͏o�K��
'
'' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'    sql = sql & "S.E020HSXOSF1PTK, "          ' �i�r�w�n�r�e�P�p�^���敪
'    sql = sql & "S.E020HSXOSF2PTK, "          ' �i�r�w�n�r�e�Q�p�^���敪
'    sql = sql & "S.E020HSXOSF3PTK, "          ' �i�r�w�n�r�e�R�p�^���敪
'    sql = sql & "S.E020HSXOSF4PTK, "          ' �i�r�w�n�r�e�S�p�^���敪
'    sql = sql & "S.E020HSXBMD1MBP, "          ' �i�r�w�a�l�c�P�ʓ����z
'    sql = sql & "S.E020HSXBMD2MBP, "          ' �i�r�w�a�l�c�Q�ʓ����z
'    sql = sql & "S.E020HSXBMD3MBP  "          ' �i�r�w�a�l�c�R�ʓ����z
'' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'
'    sql = sql & " from VECME009 BH, VECME001 S, TBCME036 U "
'    '�i��TAIL�̎擾
'    sql = sql & " where BH.E040BLOCKID='" & inBlockID & "' " & _
'                " and BH.E041INGOTPOS=ANY(select max(E041INGOTPOS) from VECME009 where E040BLOCKID='" & inBlockID & "' ) " & _
'                " and S.E018HINBAN=BH.E041HINBAN and S.E018MNOREVNO=BH.E041REVNUM" & _
'                " and S.E018FACTORY=BH.E041FACTORY and S.E018OPECOND=BH.E041OPECOND " & _
'                " and U.HINBAN=BH.E041HINBAN and U.MNOREVNO=BH.E041REVNUM" & _
'                " and U.FACTORY=BH.E041FACTORY and U.OPECOND=BH.E041OPECOND "
    ''-------------���R�����g���y�u���b�N���S�i�Ԏd�l�擾�z 2003/09/09 ooba-------------

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
    End If

    recCnt = rs.RecordCount
    BlkHinCNT = rs.RecordCount             ''�d�l�\���i�Ԑ��擾  2003/09/09 ooba

    ReDim Siyou(recCnt)
    For i = 1 To recCnt

        With Siyou(i)
            .CRYNUM = rs("E040CRYNUM")                ' �����ԍ�
            .IngotPos = rs("E040INGOTPOS")            ' �������J�n�ʒu
            .LENGTH = rs("E040LENGTH")                ' ����
            .hin.hinban = rs("E041HINBAN")            ' �i��
            .hin.mnorevno = rs("E041REVNUM")          ' ���i�ԍ������ԍ�
            .hin.factory = rs("E041FACTORY")          ' �H��
            .hin.opecond = rs("E041OPECOND")          ' ���Ə���

            .PRODCOND = rs("E037PRODCOND")            ' �������
            .PGID = rs("E037PGID")                    ' �o�f�|�h�c
            .UPLENGTH = rs("E037UPLENGTH")            ' ���グ����
            .FREELENG = rs("E037FREELENG")            ' �t���[��
            .DIAMETER = rs("E037DIAMETER")            ' ���a
            .CHARGE = rs("E037CHARGE")                ' �`���[�W��
            .SEED = rs("E037SEED")                    ' �V�[�h
            .ADDDPPOS = rs("E037ADDDPPOS")            ' �ǉ��h�[�v�ʒu

            .HSXTYPE = rs("E018HSXTYPE")              ' �i�r�w�^�C�v"
            .HSXD1CEN = fncNullCheck(rs("E018HSXD1CEN"))            ' �i�r�w���a�P���S"
            .HSXCDIR = rs("E018HSXCDIR")              ' �i�r�w�����ʕ���"

            .HSXRMIN = fncNullCheck(rs("E018HSXRMIN"))              ' �i�r�w���R����  'NULL�Ή�
            .HSXRMAX = fncNullCheck(rs("E018HSXRMAX"))              ' �i�r�w���R����@'NULL�Ή�
            .HSXRAMIN = fncNullCheck(rs("E018HSXRAMIN"))            ' �i�r�w���R���ω����@'NULL�Ή�
            .HSXRAMAX = fncNullCheck(rs("E018HSXRAMAX"))            ' �i�r�w���R���Ϗ��  'NULL�Ή�
            .HSXRMCAL = rs("E018HSXRMCAL")            ' �i�r�w���R�ʓ��v�Z     '' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
            .HSXRMBNP = fncNullCheck(rs("E018HSXRMBNP"))            ' �i�r�w���R�ʓ����z  'NULL�Ή�
            .HSXRSPOH = rs("E018HSXRSPOH")            ' �i�r�w���R����ʒu�Q��
            .HSXRSPOT = rs("E018HSXRSPOT")            ' �i�r�w���R����ʒu�Q�_
            .HSXRSPOI = rs("E018HSXRSPOI")            ' �i�r�w���R����ʒu�Q��
            .HSXRHWYT = rs("E018HSXRHWYT")            ' �i�r�w���R�ۏؕ��@�Q��
            .HSXRHWYS = rs("E018HSXRHWYS")            ' �i�r�w���R�ۏؕ��@�Q��

            .HSXONMIN = fncNullCheck(rs("E019HSXONMIN"))            ' �i�r�w�_�f�Z�x����
            .HSXONMAX = fncNullCheck(rs("E019HSXONMAX"))            ' �i�r�w�_�f�Z�x���
            .HSXONAMN = fncNullCheck(rs("E019HSXONAMN"))            ' �i�r�w�_�f�Z�x���ω���
            .HSXONAMX = fncNullCheck(rs("E019HSXONAMX"))            ' �i�r�w�_�f�Z�x���Ϗ��
            .HSXONMCL = rs("E019HSXONMCL")            ' �i�r�w�_�f�Z�x�ʓ��v�Z   '' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
            .HSXONMBP = fncNullCheck(rs("E019HSXONMBP"))            ' �i�r�w�_�f�Z�x�ʓ����z
            .HSXONSPH = rs("E019HSXONSPH")            ' �i�r�w�_�f�Z�x����ʒu�Q��
            .HSXONSPT = rs("E019HSXONSPT")            ' �i�r�w�_�f�Z�x����ʒu�Q�_
            .HSXONSPI = rs("E019HSXONSPI")            ' �i�r�w�_�f�Z�x����ʒu�Q��
            .HSXONHWT = rs("E019HSXONHWT")            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
            .HSXONHWS = rs("E019HSXONHWS")            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��

           '.HSXBM1AN = rs("E020HSXBM1AN") * 10       ' �i�r�w�a�l�c�P���ω���
           '.HSXBM1AX = rs("E020HSXBM1AX") * 10       ' �i�r�w�a�l�c�P���Ϗ��
           '.HSXBM2AN = rs("E020HSXBM2AN") * 10       ' �i�r�w�a�l�c�Q���ω���
           '.HSXBM2AX = rs("E020HSXBM2AX") * 10       ' �i�r�w�a�l�c�Q���Ϗ��
           '.HSXBM3AN = rs("E020HSXBM3AN") * 10       ' �i�r�w�a�l�c�R���ω���
           '.HSXBM3AX = rs("E020HSXBM3AX") * 10       ' �i�r�w�a�l�c�R���Ϗ��
            'BMD�ׂ��搔�@�ύX�Ή��@2003/05/17 osawa
            .HSXBM1AN = fncNullCheck(rs("E020HSXBM1AN"))            ' �i�r�w�a�l�c�P���ω���
            .HSXBM1AX = fncNullCheck(rs("E020HSXBM1AX"))            ' �i�r�w�a�l�c�P���Ϗ��
            .HSXBM2AN = fncNullCheck(rs("E020HSXBM2AN"))            ' �i�r�w�a�l�c�Q���ω���
            .HSXBM2AX = fncNullCheck(rs("E020HSXBM2AX"))            ' �i�r�w�a�l�c�Q���Ϗ��
            .HSXBM3AN = fncNullCheck(rs("E020HSXBM3AN"))            ' �i�r�w�a�l�c�R���ω���
            .HSXBM3AX = fncNullCheck(rs("E020HSXBM3AX"))            ' �i�r�w�a�l�c�R���Ϗ��
            '
            .HSXBM1SH = rs("E020HSXBM1SH")            ' �i�r�w�a�l�c�P����ʒu�Q��
            .HSXBM1ST = rs("E020HSXBM1ST")            ' �i�r�w�a�l�c�P����ʒu�Q�_
            .HSXBM1SR = rs("E020HSXBM1SR")            ' �i�r�w�a�l�c�P����ʒu�Q��
            .HSXBM1HT = rs("E020HSXBM1HT")            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
            .HSXBM1HS = rs("E020HSXBM1HS")            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
            .HSXBM2SH = rs("E020HSXBM2SH")            ' �i�r�w�a�l�c�Q����ʒu�Q��
            .HSXBM2ST = rs("E020HSXBM2ST")            ' �i�r�w�a�l�c�Q����ʒu�Q�_
            .HSXBM2SR = rs("E020HSXBM2SR")            ' �i�r�w�a�l�c�Q����ʒu�Q��
            .HSXBM2HT = rs("E020HSXBM2HT")            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
            .HSXBM2HS = rs("E020HSXBM2HS")            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
            .HSXBM3SH = rs("E020HSXBM3SH")            ' �i�r�w�a�l�c�R����ʒu�Q��
            .HSXBM3ST = rs("E020HSXBM3ST")            ' �i�r�w�a�l�c�R����ʒu�Q�_
            .HSXBM3SR = rs("E020HSXBM3SR")            ' �i�r�w�a�l�c�R����ʒu�Q��
            .HSXBM3HT = rs("E020HSXBM3HT")            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
            .HSXBM3HS = rs("E020HSXBM3HS")            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��

            .HSXOF1AX = fncNullCheck(rs("E020HSXOF1AX"))            ' �i�r�w�n�r�e�P���Ϗ��
            .HSXOF1MX = fncNullCheck(rs("E020HSXOF1MX"))            ' �i�r�w�n�r�e�P���
            .HSXOF2AX = fncNullCheck(rs("E020HSXOF2AX"))            ' �i�r�w�n�r�e�Q���Ϗ��
            .HSXOF2MX = fncNullCheck(rs("E020HSXOF2MX"))            ' �i�r�w�n�r�e�Q���
            .HSXOF3AX = fncNullCheck(rs("E020HSXOF3AX"))            ' �i�r�w�n�r�e�R���Ϗ��
            .HSXOF3MX = fncNullCheck(rs("E020HSXOF3MX"))            ' �i�r�w�n�r�e�R���
            .HSXOF4AX = fncNullCheck(rs("E020HSXOF4AX"))            ' �i�r�w�n�r�e�S���Ϗ��
            .HSXOF4MX = fncNullCheck(rs("E020HSXOF4MX"))            ' �i�r�w�n�r�e�S���
            .HSXOF1SH = rs("E020HSXOF1SH")            ' �i�r�w�n�r�e�P����ʒu�Q��
            .HSXOF1ST = rs("E020HSXOF1ST")            ' �i�r�w�n�r�e�P����ʒu�Q�_
            .HSXOF1SR = rs("E020HSXOF1SR")            ' �i�r�w�n�r�e�P����ʒu�Q��
            .HSXOF1HT = rs("E020HSXOF1HT")            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
            .HSXOF1HS = rs("E020HSXOF1HS")            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
            .HSXOF2SH = rs("E020HSXOF2SH")            ' �i�r�w�n�r�e�Q����ʒu�Q��
            .HSXOF2ST = rs("E020HSXOF2ST")            ' �i�r�w�n�r�e�Q����ʒu�Q�_
            .HSXOF2SR = rs("E020HSXOF2SR")            ' �i�r�w�n�r�e�Q����ʒu�Q��
            .HSXOF2HT = rs("E020HSXOF2HT")            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
            .HSXOF2HS = rs("E020HSXOF2HS")            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
            .HSXOF3SH = rs("E020HSXOF3SH")            ' �i�r�w�n�r�e�R����ʒu�Q��
            .HSXOF3ST = rs("E020HSXOF3ST")            ' �i�r�w�n�r�e�R����ʒu�Q�_
            .HSXOF3SR = rs("E020HSXOF3SR")            ' �i�r�w�n�r�e�R����ʒu�Q��
            .HSXOF3HT = rs("E020HSXOF3HT")            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
            .HSXOF3HS = rs("E020HSXOF3HS")            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
            .HSXOF4SH = rs("E020HSXOF4SH")            ' �i�r�w�n�r�e�S����ʒu�Q��
            .HSXOF4ST = rs("E020HSXOF4ST")            ' �i�r�w�n�r�e�S����ʒu�Q�_
            .HSXOF4SR = rs("E020HSXOF4SR")            ' �i�r�w�n�r�e�S����ʒu�Q��
            .HSXOF4HT = rs("E020HSXOF4HT")            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
            .HSXOF4HS = rs("E020HSXOF4HS")            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
            .HSXOF1NS = rs("E020HSXOF1NS")            ' �i�r�w�n�r�e�P�M�����@
            .HSXOF2NS = rs("E020HSXOF2NS")            ' �i�r�w�n�r�e�Q�M�����@
            .HSXOF3NS = rs("E020HSXOF3NS")            ' �i�r�w�n�r�e�R�M�����@
            .HSXOF4NS = rs("E020HSXOF4NS")            ' �i�r�w�n�r�e�S�M�����@
            .HSXBM1NS = rs("E020HSXBM1NS")            ' �i�r�w�a�l�c�P�M�����@
            .HSXBM2NS = rs("E020HSXBM2NS")            ' �i�r�w�a�l�c�Q�M�����@
            .HSXBM3NS = rs("E020HSXBM3NS")            ' �i�r�w�a�l�c�R�M�����@

            .HSXCNMIN = fncNullCheck(rs("E019HSXCNMIN"))            ' �i�r�w�Y�f�Z�x����
            .HSXCNMAX = fncNullCheck(rs("E019HSXCNMAX"))            ' �i�r�w�Y�f�Z�x���
            .HSXCNSPH = rs("E019HSXCNSPH")            ' �i�r�w�Y�f�Z�x����ʒu�Q��
            .HSXCNSPT = rs("E019HSXCNSPT")            ' �i�r�w�Y�f�Z�x����ʒu�Q�_
            .HSXCNSPI = rs("E019HSXCNSPI")            ' �i�r�w�Y�f�Z�x����ʒu�Q��
            .HSXCNHWT = rs("E019HSXCNHWT")            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
            .HSXCNHWS = rs("E019HSXCNHWS")            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��

            .HSXDENMX = fncNullCheck(rs("E020HSXDENMX"))            ' �i�r�w�c�������
            .HSXDENMN = fncNullCheck(rs("E020HSXDENMN"))            ' �i�r�w�c��������
            .HSXLDLMX = fncNullCheck(rs("E020HSXLDLMX"))            ' �i�r�w�k�^�c�k���
            .HSXLDLMN = fncNullCheck(rs("E020HSXLDLMN"))            ' �i�r�w�k�^�c�k����
            .HSXDVDMX = fncNullCheck(rs("E020HSXDVDMXN"))           ' �i�r�w�c�u�c�Q���   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
            .HSXDVDMN = fncNullCheck(rs("E020HSXDVDMNN"))           ' �i�r�w�c�u�c�Q����   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
            .HSXDENHT = rs("E020HSXDENHT")            ' �i�r�w�c�����ۏؕ��@�Q��
            .HSXDENHS = rs("E020HSXDENHS")            ' �i�r�w�c�����ۏؕ��@�Q��
            .HSXLDLHT = rs("E020HSXLDLHT")            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
            .HSXLDLHS = rs("E020HSXLDLHS")            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
            .HSXDVDHT = rs("E020HSXDVDHT")            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
            .HSXDVDHS = rs("E020HSXDVDHS")            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
            .HSXDENKU = rs("E020HSXDENKU")            ' �i�r�w�c���������L��
            .HSXDVDKU = rs("E020HSXDVDKU")            ' �i�r�w�c�u�c�Q�����L��
            .HSXLDLKU = rs("E020HSXLDLKU")            ' �i�r�w�k�^�c�k�����L��

            .HSXLTMIN = fncNullCheck(rs("E019HSXLTMIN"))            ' �i�r�w�k�^�C������
            .HSXLTMAX = fncNullCheck(rs("E019HSXLTMAX"))            ' �i�r�w�k�^�C�����
            .HSXLTSPH = rs("E019HSXLTSPH")            ' �i�r�w�k�^�C������ʒu�Q��
            .HSXLTSPT = rs("E019HSXLTSPT")            ' �i�r�w�k�^�C������ʒu�Q�_
            .HSXLTSPI = rs("E019HSXLTSPI")            ' �i�r�w�k�^�C������ʒu�Q��
            .HSXLTHWT = rs("E019HSXLTHWT")            ' �i�r�w�k�^�C���ۏؕ��@�Q��
            .HSXLTHWS = rs("E019HSXLTHWS")            ' �i�r�w�k�^�C���ۏؕ��@�Q��
'''''       .EPDUP = rs("EPDUP")                      ' EPD���
'''''       .TOPREG = rs("TOPREG")                    ' TOP�K��
'''''       .TAILREG = rs("TAILREG")                  ' TAIL�K��
'''''       .BTMSPRT = rs("BTMSPRT")                  ' �{�g���͏o�K��
'''''       --TEST--
'''''       NULL�Ή����Z�b�g
            .EPDUP = IIf(IsNull(rs("EPDUP")) = True, 0, rs("EPDUP"))
            .TOPREG = IIf(IsNull(rs("TOPREG")) = True, 0, rs("TOPREG"))
            .TAILREG = IIf(IsNull(rs("TAILREG")) = True, 0, rs("TAILREG"))
            .BTMSPRT = IIf(IsNull(rs("BTMSPRT")) = True, 0, rs("BTMSPRT"))

' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
            If IsNull(rs("E020HSXOSF1PTK")) = False Then .HSXOSF1PTK = rs("E020HSXOSF1PTK")   ' �i�r�w�n�r�e�P�p�^���敪
            If IsNull(rs("E020HSXOSF2PTK")) = False Then .HSXOSF2PTK = rs("E020HSXOSF2PTK")   ' �i�r�w�n�r�e�Q�p�^���敪
            If IsNull(rs("E020HSXOSF3PTK")) = False Then .HSXOSF3PTK = rs("E020HSXOSF3PTK")   ' �i�r�w�n�r�e�R�p�^���敪
            If IsNull(rs("E020HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("E020HSXOSF4PTK")   ' �i�r�w�n�r�e�S�p�^���敪
            If IsNull(rs("E020HSXBMD1MBP")) = False Then .HSXBMD1MBP = rs("E020HSXBMD1MBP")   ' �i�r�w�a�l�c�P�ʓ����z
            If IsNull(rs("E020HSXBMD2MBP")) = False Then .HSXBMD2MBP = rs("E020HSXBMD2MBP")   ' �i�r�w�a�l�c�Q�ʓ����z
            If IsNull(rs("E020HSXBMD3MBP")) = False Then .HSXBMD3MBP = rs("E020HSXBMD3MBP")   ' �i�r�w�a�l�c�R�ʓ����z
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura

        End With
        rs.MoveNext
    Next

    If scmzc_getKakouJiltuseki(inBlockID, Jiltuseki) = FUNCTION_RETURN_FAILURE Then
        getHinSiyou30 = FUNCTION_RETURN_FAILURE
        ReDim Siyou(0)
        GoTo proc_exit
    End If
    For i = 1 To recCnt
        Siyou(i).DIAMETER = (Jiltuseki.Top(1) + Jiltuseki.Top(2) + Jiltuseki.TAIL(1) + Jiltuseki.TAIL(2)) / 4 ' ���a
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
    gErr.HandleError
    Resume proc_exit
End Function


''''''�T�v      :�����֐� �T���v���ԍ����擾����
'''''Private Function getCrySmp(inCryNum As String, inIngotPos, _
'''''                           CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp _
'''''                           ) As FUNCTION_RETURN
'''''
'''''
'''''    Dim sql     As String
'''''    Dim rs      As OraDynaset
'''''    Dim recCnt  As Integer
'''''    Dim i       As Long
'''''
'''''    '�T���v���ԍ��擾
'''''    'VECME010�i�u���b�N�Ǘ����������A���̃u���b�N�ɑ΂���T���v����\������r���[�j����l���擾
'''''
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function getCrySmp"
'''''
'''''    sql = "select "
'''''    sql = sql & "XTALCS, "              ' �����ԍ�
'''''    sql = sql & "INPOSCS, "             ' �������ʒu
'''''    sql = sql & "B.LENGTH, "            ' ����
'''''    sql = sql & "B.BLOCKID, "           ' �u���b�NID
'''''    sql = sql & "SMPKBNCS, "            ' �T���v���敪
'''''    sql = sql & "REPSMPLIDCS, "         ' �T���v��No
'''''    sql = sql & "HINBCS, "              ' �i��
'''''    sql = sql & "REVNUMCS, "            ' ���i�ԍ������ԍ�
'''''    sql = sql & "FACTORYCS, "           ' �H��
'''''    sql = sql & "OPECS, "               ' ���Ə���
'''''    sql = sql & "KTKBNCS, "             ' �m��敪
'''''    sql = sql & "CRYINDRSCS, "          ' ���������w���iRs)
'''''    sql = sql & "CRYINDOICS, "          ' ���������w���iOi)
'''''    sql = sql & "CRYINDB1CS, "          ' ���������w���iB1)
'''''    sql = sql & "CRYINDB2CS, "          ' ���������w���iB2�j
'''''    sql = sql & "CRYINDB3CS, "          ' ���������w���iB3)
'''''    sql = sql & "CRYINDL1CS, "          ' ���������w���iL1)
'''''    sql = sql & "CRYINDL2CS, "          ' ���������w���iL2)
'''''    sql = sql & "CRYINDL3CS, "          ' ���������w���iL3)
'''''    sql = sql & "CRYINDL4CS, "          ' ���������w���iL4)
'''''    sql = sql & "CRYINDCSCS, "          ' ���������w���iCs)
'''''    sql = sql & "CRYINDGDCS, "          ' ���������w���iGD)
'''''    sql = sql & "CRYINDTCS, "           ' ���������w���iT)
'''''    sql = sql & "CRYINDEPCS "           ' ���������w���iEPD)
'''''
'''''    sql = sql & " from  TBCME040 B, XSDCS X "
'''''    sql = sql & " where B.CRYNUM  ='" & inCryNum & "' "
'''''    sql = sql & "   and B.INGOTPOS= " & inIngotPos
'''''    sql = sql & "   and B.BLOCKID = X.CRYNUMCS "
'''''    sql = sql & " order by X.INPOSCS "  ' TOP TAIL��
'''''
'''''
''''''''''    sql = sql & " from VECME010"
''''''''''    sql = sql & " where E040CRYNUM='" & inCryNum & "' and E040INGOTPOS=" & inIngotPos
''''''''''    sql = sql & " order by E043INPOSCS " ' TOP TAIL��
'''''
'''''
'''''
'''''    ' SQL���s
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''Debug.Print sql
'''''
'''''    If rs.RecordCount = 0 Then
'''''        getCrySmp = FUNCTION_RETURN_FAILURE
'''''        ReDim CrySmp(0)
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    recCnt = rs.RecordCount
'''''    ReDim CrySmp(recCnt)
'''''    For i = 1 To recCnt
'''''        With CrySmp(i)
'''''
''''''            .CRYNUM = rs("XTALCS")              ' �����ԍ�
''''''            .INGOTPOS = rs("INPOSCS")           ' �������ʒu
''''''            .LENGTH = rs("LENGTH")              ' ����
''''''            .BLOCKID = rs("BLOCKID")            ' �u���b�NID
''''''            .SMPKBN = rs("SMPKBNCS")            ' �T���v���敪
''''''            .SMPLNO = rs("REPSMPLIDCS")         ' �T���v��No
''''''            .hinban = rs("HINBCS")              ' �i��
''''''            .REVNUM = rs("REVNUMCS")            ' ���i�ԍ������ԍ�
''''''            .factory = rs("FACTORYCS")          ' �H��
''''''            .opecond = rs("OPECS")              ' ���Ə���
''''''            .KTKBN = rs("KTKBNCS")              ' �m��敪
''''''            .CRYINDRS = rs("CRYINDRSCS")        ' ���������w���iRs)
''''''            .CRYINDOI = rs("CRYINDOICS")        ' ���������w���iOi)
''''''            .CRYINDB1 = rs("CRYINDB1CS")        ' ���������w���iB1)
''''''            .CRYINDB2 = rs("CRYINDB2CS")        ' ���������w���iB2)
''''''            .CRYINDB3 = rs("CRYINDB3CS")        ' ���������w���iB3)
''''''            .CRYINDL1 = rs("CRYINDL1CS")        ' ���������w���iL1)
''''''            .CRYINDL2 = rs("CRYINDL2CS")        ' ���������w���iL2)
''''''            .CRYINDL3 = rs("CRYINDL3CS")        ' ���������w���iL3)
''''''            .CRYINDL4 = rs("CRYINDL4CS")        ' ���������w���iL4)
''''''            .CRYINDCS = rs("CRYINDCSCS")        ' ���������w���iCs)
''''''            .CRYINDGD = rs("CRYINDGDCS")        ' ���������w���iGD)
''''''            .CRYINDT = rs("CRYINDTCS")          ' ���������w���iT)
''''''            .CRYINDEP = rs("CRYINDEPCS")        ' ���������w���iEPD)
'''''
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    getCrySmp = FUNCTION_RETURN_SUCCESS
'''''
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    getCrySmp = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


'�T�v      :�������� �\���p�c�a�h���C�o�i�u���b�N�h�c���đ�̏ꍇ�j
'���Ұ�    :�ϐ���        ,IO ,�^                                 ,����
'          :inBlockID     ,I  ,String                             ,�Ώۃu���b�NID
'          :Siyou()       ,O  ,type_DBDRV_scmzc_fcmkc001c_Siyou   ,�i�ԁA�d�l�A���������擾�p
'          :CrySmp()      ,O  ,type_DBDRV_scmzc_fcmkc001c_CrySmp  ,�����T���v���Ǘ��擾�p
'          :Zisseki       ,O  ,type_DBDRV_scmzc_fcmkc001c_Zisseki ,���їp
'          :sErrMsg       ,O  ,String                             ,
'          :�߂�l        ,O  ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'����      :
'����      :2001/06/26 ���{ �쐬
'''''Public Function DBDRV_scmzc_fcmkc001c_Disp(inBlockID As String, _
'''''                                           Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                                           CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                                           Zisseki As type_DBDRV_scmzc_fcmkc001c_Zisseki, _
'''''                                           sErrMsg As String) As FUNCTION_RETURN
Public Function DBDRV_scmzc_fcmkc001c_Disp(inBlockID As String, _
                                           Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                                           sErrMsg As String) As FUNCTION_RETURN
''''Dim i       As Integer
''''Dim recCnt  As Integer
    Dim sDbName As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_Disp"

    DBDRV_scmzc_fcmkc001c_Disp = FUNCTION_RETURN_FAILURE

    sDbName = "V011"
    '�i�ԁASXL�d�l����f�[�^�̎擾�i���R�[�h0���̏ꍇ���G���[�j
    If FUNCTION_RETURN_FAILURE = getHinSiyou30(inBlockID, Siyou()) Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


'''' -TEST-
'''''    sDbName = "V010"
'''''    '�����T���v���̎擾(���R�[�h0���̏ꍇ���G���[)
'''''    If FUNCTION_RETURN_FAILURE = getCrySmp(Siyou(1).CRYNUM, Siyou(1).INGOTPOS, CrySmp()) Then
'''''        sErrMsg = GetMsgStr("EGET2", sDbName)
'''''        DBDRV_scmzc_fcmkc001c_Disp = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''
'''''
'''''
'''''    With Zisseki
'''''        ReDim .CRYRZ(2)
'''''        ReDim .OIZ(2)
'''''        ReDim .BMD1Z(2)
'''''        ReDim .BMD2Z(2)
'''''        ReDim .BMD3Z(2)
'''''        ReDim .OSF1Z(2)
'''''        ReDim .OSF2Z(2)
'''''        ReDim .OSF3Z(2)
'''''        ReDim .OSF4Z(2)
'''''        ReDim .csz(2)
'''''        ReDim .GDZ(2)
'''''        ReDim .LTZ(2)
'''''        ReDim .EPDZ(2)
'''''        ReDim .SURSZ(2)
'''''    End With
'''''
'''''    'recCnt = UBound(CrySmp)
'''''    '�����T���v���̎w�������Ď��т����
'''''    For i = 1 To 2 'recCnt
'''''        '�T���v���Ǘ��̕i�Ԃ��L�ۂ݂ɂ��Ȃ�
'''''        CrySmp(i).hinban = Siyou(i).hin.hinban
'''''        CrySmp(i).REVNUM = Siyou(i).hin.mnorevno
'''''        CrySmp(i).factory = Siyou(i).hin.factory
'''''        CrySmp(i).opecond = Siyou(i).hin.opecond
'''''
'''''        sDbName = "J002"
'''''        If CryR_Zisseki(Siyou(i), CrySmp(i), Zisseki.CRYRZ(i), Zisseki.SURSZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J003"
'''''        If Oi_Zisseki(Siyou(i), CrySmp(i), Zisseki.OIZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J008"
'''''        If BMD_Zisseki(Siyou(i), CrySmp(i), "1", Zisseki.BMD1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J008"
'''''        If BMD_Zisseki(Siyou(i), CrySmp(i), "2", Zisseki.BMD2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J008"
'''''        If BMD_Zisseki(Siyou(i), CrySmp(i), "3", Zisseki.BMD3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J005"
'''''        If OSF_Zisseki(Siyou(i), CrySmp(i), "1", Zisseki.OSF1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J005"
'''''        If OSF_Zisseki(Siyou(i), CrySmp(i), "2", Zisseki.OSF2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J005"
'''''        If OSF_Zisseki(Siyou(i), CrySmp(i), "3", Zisseki.OSF3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J005"
'''''        If OSF_Zisseki(Siyou(i), CrySmp(i), "4", Zisseki.OSF4Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J006"
'''''        If GD_Zisseki(Siyou(i), CrySmp(i), Zisseki.GDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J007"
'''''        If LT_Zisseki(Siyou(i), CrySmp(i), Zisseki.LTZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J001"
'''''        If EPD_Zisseki(Siyou(i), CrySmp(i), Zisseki.EPDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''    Next
'''''    sDbName = "J004"
'''''    If CS_Zisseki(Siyou(1).CRYNUM, CrySmp, Zisseki.csz) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit

    sDbName = ""
    DBDRV_scmzc_fcmkc001c_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    If Trim$(sDbName) <> "" Then sErrMsg = GetMsgStr("EGET2", sDbName)
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_scmzc_fcmkc001c_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



'�T�v      :�������� �\���p�c�a�h���C�o�i�u���b�N�h�c���đ�ȊO�̏ꍇ�j
'���Ұ�    :�ϐ���        ,IO ,�^                                ,����
'          :inBlockID     ,I  ,String                            ,�Ώۃu���b�NID
'          :Siyou()       ,O  ,type_DBDRV_scmzc_fcmkc001c_Siyou  ,�i�ԁA�d�l�A���������擾�p
'          :CrySmp()      ,O  ,type_DBDRV_scmzc_fcmkc001c_CrySmp ,�����T���v���Ǘ��擾�p
'          :Zisseki       ,O  ,typ_TBCMG002                      ,�w���P����������їp
'          :sErrMsg       ,O  ,String                            ,
'          :�߂�l        ,O  ,FUNCTION_RETURN                   ,�ǂݍ��ݐ���
'����      :
'����      :2001/06/28 ���{ �쐬
'''''Public Function DBDRV_scmzc_fcmkc001c_Disp2(inBlockID As String, _
'''''                                           Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                                           CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                                           Zisseki As typ_TBCMG002, _
'''''                                           sErrMsg As String) As FUNCTION_RETURN
Public Function DBDRV_scmzc_fcmkc001c_Disp2(inBlockID As String, _
                                           Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                                           sErrMsg As String) As FUNCTION_RETURN

    Dim sDbName As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_Disp2"

    DBDRV_scmzc_fcmkc001c_Disp2 = FUNCTION_RETURN_SUCCESS

    sDbName = "V011"
    '�i�ԁASXL�d�l����f�[�^�̎擾�i���R�[�h0�����G���[�j
    If getHinSiyou30(inBlockID, Siyou()) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001c_Disp2 = FUNCTION_RETURN_FAILURE
        sErrMsg = GetMsgStr("EGET2", sDbName)
        GoTo proc_exit
    End If

'    sDBName = "V010"
'    '�����T���v���̎擾�i���R�[�h0�����G���[�j
'    If getCrySmp(Siyou(1).Crynum, Siyou(1).IngotPos, CrySmp()) = FUNCTION_RETURN_FAILURE Then
'        DBDRV_scmzc_fcmkc001c_Disp2 = FUNCTION_RETURN_FAILURE
'        sErrMsg = GetMsgStr("EGET2", sDBName)
'        GoTo proc_exit
'    End If

'''''    sDbName = "G002"
'''''    '�w���P����������ю擾�i���R�[�h0�����G���[�j
'''''    If Kounyu_Zisseki(Zisseki, " where CRYNUM = '" & inBlockID & _
'''''                               "' and TRANCNT=ANY(select MAX(TRANCNT) from TBCMG002 " & _
'''''                               " where CRYNUM = '" & inBlockID & "' )") = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001c_Disp2 = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_scmzc_fcmkc001c_Disp2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


''''''�T�v      :���p�̏ꍇ�ɂ́A�������R�[�h����K�v���ڂ��擾����
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :Soku          ,IO ,typ_TBCMJ014 ,����]�����уf�[�^
''''''����      :���т��L�^����Ă��Ȃ����ڂɂ��āA�����̋��p���R�[�h�ɒl������΂�����̗p����
''''''����      :2002/4/26 �쑺 �쐬
'''''Private Sub UpdateFromOrgJ014(Soku As typ_TBCMJ014)
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''
'''''    With Soku
'''''        '�����̋��p���R�[�h����l���擾����
'''''        '�u���p���R�[�h�v�̔��f�́A���̈ʒu�ɑ������葪��l��1���R�[�h�������邱�Ƃ������Ƃ���
'''''        '�� T/B�̗���������Ƃ��́A�����̒l�����Ȃ��Ă� Soku�ɓ����Ă�����̂��̗p���Ă悢
'''''        '�� ���p�T���v���ł��A�������R�[�h���Ȃ���Ζ��Ȃ�
'''''        sql = sql & "select POSITION,SMPKBN"
'''''        sql = sql & ",SXL_RS_SMPPOS,SXL_OI_SMPPOS,SXL_CS_SMPPOS,SXLOSF1_SMPPOS,SXLBMD_SMPPOS,SXLGD_SMPPOS,SXLT_SMPPOS"
'''''        sql = sql & ",SXLRS_MEAS1,SXLRS_MEAS2,SXLRS_MEAS3,SXLRS_MEAS4,SXLRS_MEAS5,SXLRS_EFEHS,SXLRS_RRG"
'''''        sql = sql & ",SXLOI_OIMEAS1,SXLOI_OIMEAS2,SXLOI_OIMEAS3,SXLOI_OIMEAS4,SXLOI_OIMEAS5,SXLOI_ORGRES,SXLOI_INSPECTWAY"
'''''        sql = sql & ",SXLCS_CSMEAS,SXLCS_70PPRE"
'''''        sql = sql & ",SXLOSF1_KKSP,SXLOSF1_NETU,SXLOSF1_KKSET,SXLOSF1_MEAS1,SXLOSF1_MEAS2,SXLOSF1_MEAS3,SXLOSF1_MEAS4,SXLOSF1_MEAS5,SXLOSF1_MEAS6,SXLOSF1_MEAS7,SXLOSF1_MEAS8,SXLOSF1_MEAS9,SXLOSF1_MEAS10"
'''''        sql = sql & ",SXLOSF1_MEAS11,SXLOSF1_MEAS12,SXLOSF1_MEAS13,SXLOSF1_MEAS14,SXLOSF1_MEAS15,SXLOSF1_MEAS16,SXLOSF1_MEAS17,SXLOSF1_MEAS18,SXLOSF1_MEAS19,SXLOSF1_MEAS20,SXLOSF1_CALCMAX,SXLOSF1_CALCAVE"
'''''        sql = sql & ",SXLOSF2_KKSP,SXLOSF2_NETU,SXLOSF2_KKSET,SXLOSF2_MEAS1,SXLOSF2_MEAS2,SXLOSF2_MEAS3,SXLOSF2_MEAS4,SXLOSF2_MEAS5,SXLOSF2_MEAS6,SXLOSF2_MEAS7,SXLOSF2_MEAS8,SXLOSF2_MEAS9,SXLOSF2_MEAS10"
'''''        sql = sql & ",SXLOSF2_MEAS11,SXLOSF2_MEAS12,SXLOSF2_MEAS13,SXLOSF2_MEAS14,SXLOSF2_MEAS15,SXLOSF2_MEAS16,SXLOSF2_MEAS17,SXLOSF2_MEAS18,SXLOSF2_MEAS19,SXLOSF2_MEAS20,SXLOSF2_CALCMAX,SXLOSF2_CALCAVE"
'''''        sql = sql & ",SXLOSF3_KKSP,SXLOSF3_NETU,SXLOSF3_KKSET,SXLOSF3_MEAS1,SXLOSF3_MEAS2,SXLOSF3_MEAS3,SXLOSF3_MEAS4,SXLOSF3_MEAS5,SXLOSF3_MEAS6,SXLOSF3_MEAS7,SXLOSF3_MEAS8,SXLOSF3_MEAS9,SXLOSF3_MEAS10"
'''''        sql = sql & ",SXLOSF3_MEAS11,SXLOSF3_MEAS12,SXLOSF3_MEAS13,SXLOSF3_MEAS14,SXLOSF3_MEAS15,SXLOSF3_MEAS16,SXLOSF3_MEAS17,SXLOSF3_MEAS18,SXLOSF3_MEAS19,SXLOSF3_MEAS20,SXLOSF3_CALCMAX,SXLOSF3_CALCAVE"
'''''        sql = sql & ",SXLOSF4_KKSP,SXLOSF4_NETU,SXLOSF4_KKSET,SXLOSF4_MEAS1,SXLOSF4_MEAS2,SXLOSF4_MEAS3,SXLOSF4_MEAS4,SXLOSF4_MEAS5,SXLOSF4_MEAS6,SXLOSF4_MEAS7,SXLOSF4_MEAS8,SXLOSF4_MEAS9,SXLOSF4_MEAS10"
'''''        sql = sql & ",SXLOSF4_MEAS11,SXLOSF4_MEAS12,SXLOSF4_MEAS13,SXLOSF4_MEAS14,SXLOSF4_MEAS15,SXLOSF4_MEAS16,SXLOSF4_MEAS17,SXLOSF4_MEAS18,SXLOSF4_MEAS19,SXLOSF4_MEAS20,SXLOSF4_CALCMAX,SXLOSF4_CALCAVE"
'''''        sql = sql & ",SXLBMD1_KKSP,SXLBMD1_NETU,SXLBMD1_KKSET,SXLBMD1_MEAS1,SXLBMD1_MEAS2,SXLBMD1_MEAS3,SXLBMD1_MEAS4,SXLBMD1_MEAS5,SXLBMD1_CALCMAX,SXLBMD1_CALCAVE"
'''''        sql = sql & ",SXLBMD2_KKSP,SXLBMD2_NETU,SXLBMD2_KKSET,SXLBMD2_MEAS1,SXLBMD2_MEAS2,SXLBMD2_MEAS3,SXLBMD2_MEAS4,SXLBMD2_MEAS5,SXLBMD2_CALCMAX,SXLBMD2_CALCAVE"
'''''        sql = sql & ",SXLBMD3_KKSP,SXLBMD3_NETU,SXLBMD3_KKSET,SXLBMD3_MEAS1,SXLBMD3_MEAS2,SXLBMD3_MEAS3,SXLBMD3_MEAS4,SXLBMD3_MEAS5,SXLBMD3_CALCMAX,SXLBMD3_CALCAVE"
'''''        sql = sql & ",SXLGD_MS01LDL1,SXLGD_MS01LDL2,SXLGD_MS01LDL3,SXLGD_MS01LDL4,SXLGD_MS01LDL5,SXLGD_MS01DEN1,SXLGD_MS01DEN2,SXLGD_MS01DEN3,SXLGD_MS01DEN4,SXLGD_MS01DEN5"
'''''        sql = sql & ",SXLGD_MS02LDL1,SXLGD_MS02LDL2,SXLGD_MS02LDL3,SXLGD_MS02LDL4,SXLGD_MS02LDL5,SXLGD_MS02DEN1,SXLGD_MS02DEN2,SXLGD_MS02DEN3,SXLGD_MS02DEN4,SXLGD_MS02DEN5"
'''''        sql = sql & ",SXLGD_MS03LDL1,SXLGD_MS03LDL2,SXLGD_MS03LDL3,SXLGD_MS03LDL4,SXLGD_MS03LDL5,SXLGD_MS03DEN1,SXLGD_MS03DEN2,SXLGD_MS03DEN3,SXLGD_MS03DEN4,SXLGD_MS03DEN5"
'''''        sql = sql & ",SXLGD_MS04LDL1,SXLGD_MS04LDL2,SXLGD_MS04LDL3,SXLGD_MS04LDL4,SXLGD_MS04LDL5,SXLGD_MS04DEN1,SXLGD_MS04DEN2,SXLGD_MS04DEN3,SXLGD_MS04DEN4,SXLGD_MS04DEN5"
'''''        sql = sql & ",SXLGD_MS05LDL1,SXLGD_MS05LDL2,SXLGD_MS05LDL3,SXLGD_MS05LDL4,SXLGD_MS05LDL5,SXLGD_MS05DEN1,SXLGD_MS05DEN2,SXLGD_MS05DEN3,SXLGD_MS05DEN4,SXLGD_MS05DEN5"
'''''        sql = sql & ",SXLGD_MS06LDL1,SXLGD_MS06LDL2,SXLGD_MS06LDL3,SXLGD_MS06LDL4,SXLGD_MS06LDL5,SXLGD_MS06DEN1,SXLGD_MS06DEN2,SXLGD_MS06DEN3,SXLGD_MS06DEN4,SXLGD_MS06DEN5"
'''''        sql = sql & ",SXLGD_MS07LDL1,SXLGD_MS07LDL2,SXLGD_MS07LDL3,SXLGD_MS07LDL4,SXLGD_MS07LDL5,SXLGD_MS07DEN1,SXLGD_MS07DEN2,SXLGD_MS07DEN3,SXLGD_MS07DEN4,SXLGD_MS07DEN5"
'''''        sql = sql & ",SXLGD_MS08LDL1,SXLGD_MS08LDL2,SXLGD_MS08LDL3,SXLGD_MS08LDL4,SXLGD_MS08LDL5,SXLGD_MS08DEN1,SXLGD_MS08DEN2,SXLGD_MS08DEN3,SXLGD_MS08DEN4,SXLGD_MS08DEN5"
'''''        sql = sql & ",SXLGD_MS09LDL1,SXLGD_MS09LDL2,SXLGD_MS09LDL3,SXLGD_MS09LDL4,SXLGD_MS09LDL5,SXLGD_MS09DEN1,SXLGD_MS09DEN2,SXLGD_MS09DEN3,SXLGD_MS09DEN4,SXLGD_MS09DEN5"
'''''        sql = sql & ",SXLGD_MS10LDL1,SXLGD_MS10LDL2,SXLGD_MS10LDL3,SXLGD_MS10LDL4,SXLGD_MS10LDL5,SXLGD_MS10DEN1,SXLGD_MS10DEN2,SXLGD_MS10DEN3,SXLGD_MS10DEN4,SXLGD_MS10DEN5"
'''''        sql = sql & ",SXLGD_MS11LDL1,SXLGD_MS11LDL2,SXLGD_MS11LDL3,SXLGD_MS11LDL4,SXLGD_MS11LDL5,SXLGD_MS11DEN1,SXLGD_MS11DEN2,SXLGD_MS11DEN3,SXLGD_MS11DEN4,SXLGD_MS11DEN5"
'''''        sql = sql & ",SXLGD_MS12LDL1,SXLGD_MS12LDL2,SXLGD_MS12LDL3,SXLGD_MS12LDL4,SXLGD_MS12LDL5,SXLGD_MS12DEN1,SXLGD_MS12DEN2,SXLGD_MS12DEN3,SXLGD_MS12DEN4,SXLGD_MS12DEN5"
'''''        sql = sql & ",SXLGD_MS13LDL1,SXLGD_MS13LDL2,SXLGD_MS13LDL3,SXLGD_MS13LDL4,SXLGD_MS13LDL5,SXLGD_MS13DEN1,SXLGD_MS13DEN2,SXLGD_MS13DEN3,SXLGD_MS13DEN4,SXLGD_MS13DEN5"
'''''        sql = sql & ",SXLGD_MS14LDL1,SXLGD_MS14LDL2,SXLGD_MS14LDL3,SXLGD_MS14LDL4,SXLGD_MS14LDL5,SXLGD_MS14DEN1,SXLGD_MS14DEN2,SXLGD_MS14DEN3,SXLGD_MS14DEN4,SXLGD_MS14DEN5"
'''''        sql = sql & ",SXLGD_MS15LDL1,SXLGD_MS15LDL2,SXLGD_MS15LDL3,SXLGD_MS15LDL4,SXLGD_MS15LDL5,SXLGD_MS15DEN1,SXLGD_MS15DEN2,SXLGD_MS15DEN3,SXLGD_MS15DEN4,SXLGD_MS15DEN5"
'''''        sql = sql & ",SXLGD_MSRSDEN,SXLGD_MSRSLDL,SXLGD_MSRSDVD2"
'''''        sql = sql & ",SXLLT_MEASPEAK,SXLLT_MEAS1,SXLLT_MEAS2,SXLLT_MEAS3,SXLLT_MEAS4,SXLLT_MEAS5,SXLLT_CALCMEAS"
'''''        sql = sql & ",SXLOSF1_POS1,SXLOSF1_WID1,SXLOSF1_RD1"
'''''        sql = sql & ",SXLOSF1_POS2,SXLOSF1_WID2,SXLOSF1_RD2"
'''''        sql = sql & ",SXLOSF1_POS3,SXLOSF1_WID3,SXLOSF1_RD3"
'''''        sql = sql & ",SXLOSF2_POS1,SXLOSF2_WID1,SXLOSF2_RD1"
'''''        sql = sql & ",SXLOSF2_POS2,SXLOSF2_WID2,SXLOSF2_RD2"
'''''        sql = sql & ",SXLOSF2_POS3,SXLOSF2_WID3,SXLOSF2_RD3"
'''''        sql = sql & ",SXLOSF3_POS1,SXLOSF3_WID1,SXLOSF3_RD1"
'''''        sql = sql & ",SXLOSF3_POS2,SXLOSF3_WID2,SXLOSF3_RD2"
'''''        sql = sql & ",SXLOSF3_POS3,SXLOSF3_WID3,SXLOSF3_RD3"
'''''        sql = sql & ",SXLOSF4_POS1,SXLOSF4_WID1,SXLOSF4_RD1"
'''''        sql = sql & ",SXLOSF4_POS2,SXLOSF4_WID2,SXLOSF4_RD2"
'''''        sql = sql & ",SXLOSF4_POS3,SXLOSF4_WID3,SXLOSF4_RD3"
'''''        sql = sql & ",SXLGD_MS01DVD2,SXLGD_MS02DVD2,SXLGD_MS03DVD2,SXLGD_MS04DVD2,SXLGD_MS05DVD2"
'''''        sql = sql & ",SXLBMD1_MNBCR,SXLBMD2_MNBCR,SXLBMD3_MNBCR"
'''''            '�����܂ł�SQL�� "select *" �̕����K�؂��Ǝv�� (nomura:�K�v�J�������قڑS�J�����̂���)
'''''        sql = sql & " from TBCMJ014"
'''''        sql = sql & " where CRYNUM='" & .CRYNUM & "' and POSITION=" & .POSITION & " and SMPKBN='" & .SMPKBN & "'"
'''''        sql = sql & " and 1=(select count(*) from TBCMJ014 where CRYNUM='" & .CRYNUM & "' and POSITION=" & .POSITION & ")"
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''        If rs.RecordCount > 0 Then
'''''            '�������ږ��ɁA���т��R�s�[���� (�������R�[�h�ɒl�������āA���u���b�N�Ɏ��т��Ȃ��ꍇ)
'''''            If (.SXL_RS_SMPPOS = -1) And (rs("SXL_RS_SMPPOS") <> -1) Then
'''''                'Rs���т��R�s�[
'''''                .SXL_RS_SMPPOS = rs("SXL_RS_SMPPOS")
'''''                .SXLRS_MEAS1 = rs("SXLRS_MEAS1")
'''''                .SXLRS_MEAS2 = rs("SXLRS_MEAS2")
'''''                .SXLRS_MEAS3 = rs("SXLRS_MEAS3")
'''''                .SXLRS_MEAS4 = rs("SXLRS_MEAS4")
'''''                .SXLRS_MEAS5 = rs("SXLRS_MEAS5")
'''''                .SXLRS_EFEHS = rs("SXLRS_EFEHS")
'''''                .SXLRS_RRG = rs("SXLRS_RRG")
'''''            End If
'''''            If (.SXL_OI_SMPPOS = -1) And (rs("SXL_OI_SMPPOS") <> -1) Then
'''''                'Oi���т��R�s�[
'''''                .SXL_OI_SMPPOS = rs("SXL_OI_SMPPOS")
'''''                .SXLOI_OIMEAS1 = rs("SXLOI_OIMEAS1")
'''''                .SXLOI_OIMEAS2 = rs("SXLOI_OIMEAS2")
'''''                .SXLOI_OIMEAS3 = rs("SXLOI_OIMEAS3")
'''''                .SXLOI_OIMEAS4 = rs("SXLOI_OIMEAS4")
'''''                .SXLOI_OIMEAS5 = rs("SXLOI_OIMEAS5")
'''''                .SXLOI_ORGRES = rs("SXLOI_ORGRES")
'''''                .SXLOI_INSPECTWAY = rs("SXLOI_INSPECTWAY")
'''''            End If
'''''            If (.SXLCS_CSMEAS = -1) And (rs("SXLCS_CSMEAS") <> -1) Then
'''''                'Cs���т��R�s�[
'''''                .SXL_CS_SMPPOS = rs("SXL_CS_SMPPOS")
'''''                .SXLCS_CSMEAS = rs("SXLCS_CSMEAS")
'''''                .SXLCS_70PPRE = rs("SXLCS_70PPRE")
'''''            End If
'''''            If (.SXLOSF1_SMPPOS = -1) And (rs("SXLOSF1_SMPPOS") <> -1) Then
'''''                'OSF���т��R�s�[
'''''                .SXLOSF1_SMPPOS = rs("SXLOSF1_SMPPOS")
'''''                .SXLOSF1_KKSP = rs("SXLOSF1_KKSP")
'''''                .SXLOSF1_NETU = rs("SXLOSF1_NETU")
'''''                .SXLOSF1_KKSET = rs("SXLOSF1_KKSET")
'''''                .SXLOSF1_MEAS1 = rs("SXLOSF1_MEAS1")
'''''                .SXLOSF1_MEAS2 = rs("SXLOSF1_MEAS2")
'''''                .SXLOSF1_MEAS3 = rs("SXLOSF1_MEAS3")
'''''                .SXLOSF1_MEAS4 = rs("SXLOSF1_MEAS4")
'''''                .SXLOSF1_MEAS5 = rs("SXLOSF1_MEAS5")
'''''                .SXLOSF1_MEAS6 = rs("SXLOSF1_MEAS6")
'''''                .SXLOSF1_MEAS7 = rs("SXLOSF1_MEAS7")
'''''                .SXLOSF1_MEAS8 = rs("SXLOSF1_MEAS8")
'''''                .SXLOSF1_MEAS9 = rs("SXLOSF1_MEAS9")
'''''                .SXLOSF1_MEAS10 = rs("SXLOSF1_MEAS10")
'''''                .SXLOSF1_MEAS11 = rs("SXLOSF1_MEAS11")
'''''                .SXLOSF1_MEAS12 = rs("SXLOSF1_MEAS12")
'''''                .SXLOSF1_MEAS13 = rs("SXLOSF1_MEAS13")
'''''                .SXLOSF1_MEAS14 = rs("SXLOSF1_MEAS14")
'''''                .SXLOSF1_MEAS15 = rs("SXLOSF1_MEAS15")
'''''                .SXLOSF1_MEAS16 = rs("SXLOSF1_MEAS16")
'''''                .SXLOSF1_MEAS17 = rs("SXLOSF1_MEAS17")
'''''                .SXLOSF1_MEAS18 = rs("SXLOSF1_MEAS18")
'''''                .SXLOSF1_MEAS19 = rs("SXLOSF1_MEAS19")
'''''                .SXLOSF1_MEAS20 = rs("SXLOSF1_MEAS20")
'''''                .SXLOSF1_CALCMAX = rs("SXLOSF1_CALCMAX")
'''''                .SXLOSF1_CALCAVE = rs("SXLOSF1_CALCAVE")
'''''                .SXLOSF2_KKSP = rs("SXLOSF2_KKSP")
'''''                .SXLOSF2_NETU = rs("SXLOSF2_NETU")
'''''                .SXLOSF2_KKSET = rs("SXLOSF2_KKSET")
'''''                .SXLOSF2_MEAS1 = rs("SXLOSF2_MEAS1")
'''''                .SXLOSF2_MEAS2 = rs("SXLOSF2_MEAS2")
'''''                .SXLOSF2_MEAS3 = rs("SXLOSF2_MEAS3")
'''''                .SXLOSF2_MEAS4 = rs("SXLOSF2_MEAS4")
'''''                .SXLOSF2_MEAS5 = rs("SXLOSF2_MEAS5")
'''''                .SXLOSF2_MEAS6 = rs("SXLOSF2_MEAS6")
'''''                .SXLOSF2_MEAS7 = rs("SXLOSF2_MEAS7")
'''''                .SXLOSF2_MEAS8 = rs("SXLOSF2_MEAS8")
'''''                .SXLOSF2_MEAS9 = rs("SXLOSF2_MEAS9")
'''''                .SXLOSF2_MEAS10 = rs("SXLOSF2_MEAS10")
'''''                .SXLOSF2_MEAS11 = rs("SXLOSF2_MEAS11")
'''''                .SXLOSF2_MEAS12 = rs("SXLOSF2_MEAS12")
'''''                .SXLOSF2_MEAS13 = rs("SXLOSF2_MEAS13")
'''''                .SXLOSF2_MEAS14 = rs("SXLOSF2_MEAS14")
'''''                .SXLOSF2_MEAS15 = rs("SXLOSF2_MEAS15")
'''''                .SXLOSF2_MEAS16 = rs("SXLOSF2_MEAS16")
'''''                .SXLOSF2_MEAS17 = rs("SXLOSF2_MEAS17")
'''''                .SXLOSF2_MEAS18 = rs("SXLOSF2_MEAS18")
'''''                .SXLOSF2_MEAS19 = rs("SXLOSF2_MEAS19")
'''''                .SXLOSF2_MEAS20 = rs("SXLOSF2_MEAS20")
'''''                .SXLOSF2_CALCMAX = rs("SXLOSF2_CALCMAX")
'''''                .SXLOSF2_CALCAVE = rs("SXLOSF2_CALCAVE")
'''''                .SXLOSF3_KKSP = rs("SXLOSF3_KKSP")
'''''                .SXLOSF3_NETU = rs("SXLOSF3_NETU")
'''''                .SXLOSF3_KKSET = rs("SXLOSF3_KKSET")
'''''                .SXLOSF3_MEAS1 = rs("SXLOSF3_MEAS1")
'''''                .SXLOSF3_MEAS2 = rs("SXLOSF3_MEAS2")
'''''                .SXLOSF3_MEAS3 = rs("SXLOSF3_MEAS3")
'''''                .SXLOSF3_MEAS4 = rs("SXLOSF3_MEAS4")
'''''                .SXLOSF3_MEAS5 = rs("SXLOSF3_MEAS5")
'''''                .SXLOSF3_MEAS6 = rs("SXLOSF3_MEAS6")
'''''                .SXLOSF3_MEAS7 = rs("SXLOSF3_MEAS7")
'''''                .SXLOSF3_MEAS8 = rs("SXLOSF3_MEAS8")
'''''                .SXLOSF3_MEAS9 = rs("SXLOSF3_MEAS9")
'''''                .SXLOSF3_MEAS10 = rs("SXLOSF3_MEAS10")
'''''                .SXLOSF3_MEAS11 = rs("SXLOSF3_MEAS11")
'''''                .SXLOSF3_MEAS12 = rs("SXLOSF3_MEAS12")
'''''                .SXLOSF3_MEAS13 = rs("SXLOSF3_MEAS13")
'''''                .SXLOSF3_MEAS14 = rs("SXLOSF3_MEAS14")
'''''                .SXLOSF3_MEAS15 = rs("SXLOSF3_MEAS15")
'''''                .SXLOSF3_MEAS16 = rs("SXLOSF3_MEAS16")
'''''                .SXLOSF3_MEAS17 = rs("SXLOSF3_MEAS17")
'''''                .SXLOSF3_MEAS18 = rs("SXLOSF3_MEAS18")
'''''                .SXLOSF3_MEAS19 = rs("SXLOSF3_MEAS19")
'''''                .SXLOSF3_MEAS20 = rs("SXLOSF3_MEAS20")
'''''                .SXLOSF3_CALCMAX = rs("SXLOSF3_CALCMAX")
'''''                .SXLOSF3_CALCAVE = rs("SXLOSF3_CALCAVE")
'''''                .SXLOSF4_KKSP = rs("SXLOSF4_KKSP")
'''''                .SXLOSF4_NETU = rs("SXLOSF4_NETU")
'''''                .SXLOSF4_KKSET = rs("SXLOSF4_KKSET")
'''''                .SXLOSF4_MEAS1 = rs("SXLOSF4_MEAS1")
'''''                .SXLOSF4_MEAS2 = rs("SXLOSF4_MEAS2")
'''''                .SXLOSF4_MEAS3 = rs("SXLOSF4_MEAS3")
'''''                .SXLOSF4_MEAS4 = rs("SXLOSF4_MEAS4")
'''''                .SXLOSF4_MEAS5 = rs("SXLOSF4_MEAS5")
'''''                .SXLOSF4_MEAS6 = rs("SXLOSF4_MEAS6")
'''''                .SXLOSF4_MEAS7 = rs("SXLOSF4_MEAS7")
'''''                .SXLOSF4_MEAS8 = rs("SXLOSF4_MEAS8")
'''''                .SXLOSF4_MEAS9 = rs("SXLOSF4_MEAS9")
'''''                .SXLOSF4_MEAS10 = rs("SXLOSF4_MEAS10")
'''''                .SXLOSF4_MEAS11 = rs("SXLOSF4_MEAS11")
'''''                .SXLOSF4_MEAS12 = rs("SXLOSF4_MEAS12")
'''''                .SXLOSF4_MEAS13 = rs("SXLOSF4_MEAS13")
'''''                .SXLOSF4_MEAS14 = rs("SXLOSF4_MEAS14")
'''''                .SXLOSF4_MEAS15 = rs("SXLOSF4_MEAS15")
'''''                .SXLOSF4_MEAS16 = rs("SXLOSF4_MEAS16")
'''''                .SXLOSF4_MEAS17 = rs("SXLOSF4_MEAS17")
'''''                .SXLOSF4_MEAS18 = rs("SXLOSF4_MEAS18")
'''''                .SXLOSF4_MEAS19 = rs("SXLOSF4_MEAS19")
'''''                .SXLOSF4_MEAS20 = rs("SXLOSF4_MEAS20")
'''''                .SXLOSF4_CALCMAX = rs("SXLOSF4_CALCMAX")
'''''                .SXLOSF4_CALCAVE = rs("SXLOSF4_CALCAVE")
'''''                If IsNull(rs("SXLOSF1_POS1")) = False Then .SXLOSF1_POS1 = rs("SXLOSF1_POS1")
'''''                If IsNull(rs("SXLOSF1_WID1")) = False Then .SXLOSF1_WID1 = rs("SXLOSF1_WID1")
'''''                If IsNull(rs("SXLOSF1_RD1")) = False Then .SXLOSF1_RD1 = rs("SXLOSF1_RD1")
'''''                If IsNull(rs("SXLOSF1_POS2")) = False Then .SXLOSF1_POS2 = rs("SXLOSF1_POS2")
'''''                If IsNull(rs("SXLOSF1_WID2")) = False Then .SXLOSF1_WID2 = rs("SXLOSF1_WID2")
'''''                If IsNull(rs("SXLOSF1_RD2")) = False Then .SXLOSF1_RD2 = rs("SXLOSF1_RD2")
'''''                If IsNull(rs("SXLOSF1_POS3")) = False Then .SXLOSF1_POS3 = rs("SXLOSF1_POS3")
'''''                If IsNull(rs("SXLOSF1_WID3")) = False Then .SXLOSF1_WID3 = rs("SXLOSF1_WID3")
'''''                If IsNull(rs("SXLOSF1_RD3")) = False Then .SXLOSF1_RD3 = rs("SXLOSF1_RD3")
'''''                If IsNull(rs("SXLOSF2_POS1")) = False Then .SXLOSF2_POS1 = rs("SXLOSF2_POS1")
'''''                If IsNull(rs("SXLOSF2_WID1")) = False Then .SXLOSF2_WID1 = rs("SXLOSF2_WID1")
'''''                If IsNull(rs("SXLOSF2_RD1")) = False Then .SXLOSF2_RD1 = rs("SXLOSF2_RD1")
'''''                If IsNull(rs("SXLOSF2_POS2")) = False Then .SXLOSF2_POS2 = rs("SXLOSF2_POS2")
'''''                If IsNull(rs("SXLOSF2_WID2")) = False Then .SXLOSF2_WID2 = rs("SXLOSF2_WID2")
'''''                If IsNull(rs("SXLOSF2_RD2")) = False Then .SXLOSF2_RD2 = rs("SXLOSF2_RD2")
'''''                If IsNull(rs("SXLOSF2_POS3")) = False Then .SXLOSF2_POS3 = rs("SXLOSF2_POS3")
'''''                If IsNull(rs("SXLOSF2_WID3")) = False Then .SXLOSF2_WID3 = rs("SXLOSF2_WID3")
'''''                If IsNull(rs("SXLOSF2_RD3")) = False Then .SXLOSF2_RD3 = rs("SXLOSF2_RD3")
'''''                If IsNull(rs("SXLOSF3_POS1")) = False Then .SXLOSF3_POS1 = rs("SXLOSF3_POS1")
'''''                If IsNull(rs("SXLOSF3_WID1")) = False Then .SXLOSF3_WID1 = rs("SXLOSF3_WID1")
'''''                If IsNull(rs("SXLOSF3_RD1")) = False Then .SXLOSF3_RD1 = rs("SXLOSF3_RD1")
'''''                If IsNull(rs("SXLOSF3_POS2")) = False Then .SXLOSF3_POS2 = rs("SXLOSF3_POS2")
'''''                If IsNull(rs("SXLOSF3_WID2")) = False Then .SXLOSF3_WID2 = rs("SXLOSF3_WID2")
'''''                If IsNull(rs("SXLOSF3_RD2")) = False Then .SXLOSF3_RD2 = rs("SXLOSF3_RD2")
'''''                If IsNull(rs("SXLOSF3_POS3")) = False Then .SXLOSF3_POS3 = rs("SXLOSF3_POS3")
'''''                If IsNull(rs("SXLOSF3_WID3")) = False Then .SXLOSF3_WID3 = rs("SXLOSF3_WID3")
'''''                If IsNull(rs("SXLOSF3_RD3")) = False Then .SXLOSF3_RD3 = rs("SXLOSF3_RD3")
'''''                If IsNull(rs("SXLOSF4_POS1")) = False Then .SXLOSF4_POS1 = rs("SXLOSF4_POS1")
'''''                If IsNull(rs("SXLOSF4_WID1")) = False Then .SXLOSF4_WID1 = rs("SXLOSF4_WID1")
'''''                If IsNull(rs("SXLOSF4_RD1")) = False Then .SXLOSF4_RD1 = rs("SXLOSF4_RD1")
'''''                If IsNull(rs("SXLOSF4_POS2")) = False Then .SXLOSF4_POS2 = rs("SXLOSF4_POS2")
'''''                If IsNull(rs("SXLOSF4_WID2")) = False Then .SXLOSF4_WID2 = rs("SXLOSF4_WID2")
'''''                If IsNull(rs("SXLOSF4_RD2")) = False Then .SXLOSF4_RD2 = rs("SXLOSF4_RD2")
'''''                If IsNull(rs("SXLOSF4_POS3")) = False Then .SXLOSF4_POS3 = rs("SXLOSF4_POS3")
'''''                If IsNull(rs("SXLOSF4_WID3")) = False Then .SXLOSF4_WID3 = rs("SXLOSF4_WID3")
'''''                If IsNull(rs("SXLOSF4_RD3")) = False Then .SXLOSF4_RD3 = rs("SXLOSF4_RD3")
'''''            End If
'''''            If (.SXLBMD_SMPPOS = -1) And (rs("SXLBMD_SMPPOS") <> -1) Then
'''''                'BMD���т��R�s�[
'''''                .SXLBMD_SMPPOS = rs("SXLBMD_SMPPOS")
'''''                .SXLBMD1_KKSP = rs("SXLBMD1_KKSP")
'''''                .SXLBMD1_NETU = rs("SXLBMD1_NETU")
'''''                .SXLBMD1_KKSET = rs("SXLBMD1_KKSET")
'''''                .SXLBMD1_MEAS1 = rs("SXLBMD1_MEAS1")
'''''                .SXLBMD1_MEAS2 = rs("SXLBMD1_MEAS2")
'''''                .SXLBMD1_MEAS3 = rs("SXLBMD1_MEAS3")
'''''                .SXLBMD1_MEAS4 = rs("SXLBMD1_MEAS4")
'''''                .SXLBMD1_MEAS5 = rs("SXLBMD1_MEAS5")
'''''                .SXLBMD1_CALCMAX = rs("SXLBMD1_CALCMAX")
'''''                .SXLBMD1_CALCAVE = rs("SXLBMD1_CALCAVE")
'''''                .SXLBMD2_KKSP = rs("SXLBMD2_KKSP")
'''''                .SXLBMD2_NETU = rs("SXLBMD2_NETU")
'''''                .SXLBMD2_KKSET = rs("SXLBMD2_KKSET")
'''''                .SXLBMD2_MEAS1 = rs("SXLBMD2_MEAS1")
'''''                .SXLBMD2_MEAS2 = rs("SXLBMD2_MEAS2")
'''''                .SXLBMD2_MEAS3 = rs("SXLBMD2_MEAS3")
'''''                .SXLBMD2_MEAS4 = rs("SXLBMD2_MEAS4")
'''''                .SXLBMD2_MEAS5 = rs("SXLBMD2_MEAS5")
'''''                .SXLBMD2_CALCMAX = rs("SXLBMD2_CALCMAX")
'''''                .SXLBMD2_CALCAVE = rs("SXLBMD2_CALCAVE")
'''''                .SXLBMD3_KKSP = rs("SXLBMD3_KKSP")
'''''                .SXLBMD3_NETU = rs("SXLBMD3_NETU")
'''''                .SXLBMD3_KKSET = rs("SXLBMD3_KKSET")
'''''                .SXLBMD3_MEAS1 = rs("SXLBMD3_MEAS1")
'''''                .SXLBMD3_MEAS2 = rs("SXLBMD3_MEAS2")
'''''                .SXLBMD3_MEAS3 = rs("SXLBMD3_MEAS3")
'''''                .SXLBMD3_MEAS4 = rs("SXLBMD3_MEAS4")
'''''                .SXLBMD3_MEAS5 = rs("SXLBMD3_MEAS5")
'''''                .SXLBMD3_CALCMAX = rs("SXLBMD3_CALCMAX")
'''''                .SXLBMD3_CALCAVE = rs("SXLBMD3_CALCAVE")
'''''                If IsNull(rs("SXLBMD1_MNBCR")) = False Then .SXLBMD1_MNBCR = rs("SXLBMD1_MNBCR")
'''''                If IsNull(rs("SXLBMD2_MNBCR")) = False Then .SXLBMD2_MNBCR = rs("SXLBMD2_MNBCR")
'''''                If IsNull(rs("SXLBMD3_MNBCR")) = False Then .SXLBMD3_MNBCR = rs("SXLBMD3_MNBCR")
'''''            End If
'''''            If (.SXLGD_SMPPOS = -1) And (rs("SXLGD_SMPPOS") <> -1) Then
'''''                'GD���т��R�s�[
'''''                .SXLGD_SMPPOS = rs("SXLGD_SMPPOS")
'''''                .SXLGD_MS01LDL1 = rs("SXLGD_MS01LDL1")
'''''                .SXLGD_MS01LDL2 = rs("SXLGD_MS01LDL2")
'''''                .SXLGD_MS01LDL3 = rs("SXLGD_MS01LDL3")
'''''                .SXLGD_MS01LDL4 = rs("SXLGD_MS01LDL4")
'''''                .SXLGD_MS01LDL5 = rs("SXLGD_MS01LDL5")
'''''                .SXLGD_MS01DEN1 = rs("SXLGD_MS01DEN1")
'''''                .SXLGD_MS01DEN2 = rs("SXLGD_MS01DEN2")
'''''                .SXLGD_MS01DEN3 = rs("SXLGD_MS01DEN3")
'''''                .SXLGD_MS01DEN4 = rs("SXLGD_MS01DEN4")
'''''                .SXLGD_MS01DEN5 = rs("SXLGD_MS01DEN5")
'''''                .SXLGD_MS02LDL1 = rs("SXLGD_MS02LDL1")
'''''                .SXLGD_MS02LDL2 = rs("SXLGD_MS02LDL2")
'''''                .SXLGD_MS02LDL3 = rs("SXLGD_MS02LDL3")
'''''                .SXLGD_MS02LDL4 = rs("SXLGD_MS02LDL4")
'''''                .SXLGD_MS02LDL5 = rs("SXLGD_MS02LDL5")
'''''                .SXLGD_MS02DEN1 = rs("SXLGD_MS02DEN1")
'''''                .SXLGD_MS02DEN2 = rs("SXLGD_MS02DEN2")
'''''                .SXLGD_MS02DEN3 = rs("SXLGD_MS02DEN3")
'''''                .SXLGD_MS02DEN4 = rs("SXLGD_MS02DEN4")
'''''                .SXLGD_MS02DEN5 = rs("SXLGD_MS02DEN5")
'''''                .SXLGD_MS03LDL1 = rs("SXLGD_MS03LDL1")
'''''                .SXLGD_MS03LDL2 = rs("SXLGD_MS03LDL2")
'''''                .SXLGD_MS03LDL3 = rs("SXLGD_MS03LDL3")
'''''                .SXLGD_MS03LDL4 = rs("SXLGD_MS03LDL4")
'''''                .SXLGD_MS03LDL5 = rs("SXLGD_MS03LDL5")
'''''                .SXLGD_MS03DEN1 = rs("SXLGD_MS03DEN1")
'''''                .SXLGD_MS03DEN2 = rs("SXLGD_MS03DEN2")
'''''                .SXLGD_MS03DEN3 = rs("SXLGD_MS03DEN3")
'''''                .SXLGD_MS03DEN4 = rs("SXLGD_MS03DEN4")
'''''                .SXLGD_MS03DEN5 = rs("SXLGD_MS03DEN5")
'''''                .SXLGD_MS04LDL1 = rs("SXLGD_MS04LDL1")
'''''                .SXLGD_MS04LDL2 = rs("SXLGD_MS04LDL2")
'''''                .SXLGD_MS04LDL3 = rs("SXLGD_MS04LDL3")
'''''                .SXLGD_MS04LDL4 = rs("SXLGD_MS04LDL4")
'''''                .SXLGD_MS04LDL5 = rs("SXLGD_MS04LDL5")
'''''                .SXLGD_MS04DEN1 = rs("SXLGD_MS04DEN1")
'''''                .SXLGD_MS04DEN2 = rs("SXLGD_MS04DEN2")
'''''                .SXLGD_MS04DEN3 = rs("SXLGD_MS04DEN3")
'''''                .SXLGD_MS04DEN4 = rs("SXLGD_MS04DEN4")
'''''                .SXLGD_MS04DEN5 = rs("SXLGD_MS04DEN5")
'''''                .SXLGD_MS05LDL1 = rs("SXLGD_MS05LDL1")
'''''                .SXLGD_MS05LDL2 = rs("SXLGD_MS05LDL2")
'''''                .SXLGD_MS05LDL3 = rs("SXLGD_MS05LDL3")
'''''                .SXLGD_MS05LDL4 = rs("SXLGD_MS05LDL4")
'''''                .SXLGD_MS05LDL5 = rs("SXLGD_MS05LDL5")
'''''                .SXLGD_MS05DEN1 = rs("SXLGD_MS05DEN1")
'''''                .SXLGD_MS05DEN2 = rs("SXLGD_MS05DEN2")
'''''                .SXLGD_MS05DEN3 = rs("SXLGD_MS05DEN3")
'''''                .SXLGD_MS05DEN4 = rs("SXLGD_MS05DEN4")
'''''                .SXLGD_MS05DEN5 = rs("SXLGD_MS05DEN5")
'''''                .SXLGD_MS06LDL1 = rs("SXLGD_MS06LDL1")
'''''                .SXLGD_MS06LDL2 = rs("SXLGD_MS06LDL2")
'''''                .SXLGD_MS06LDL3 = rs("SXLGD_MS06LDL3")
'''''                .SXLGD_MS06LDL4 = rs("SXLGD_MS06LDL4")
'''''                .SXLGD_MS06LDL5 = rs("SXLGD_MS06LDL5")
'''''                .SXLGD_MS06DEN1 = rs("SXLGD_MS06DEN1")
'''''                .SXLGD_MS06DEN2 = rs("SXLGD_MS06DEN2")
'''''                .SXLGD_MS06DEN3 = rs("SXLGD_MS06DEN3")
'''''                .SXLGD_MS06DEN4 = rs("SXLGD_MS06DEN4")
'''''                .SXLGD_MS06DEN5 = rs("SXLGD_MS06DEN5")
'''''                .SXLGD_MS07LDL1 = rs("SXLGD_MS07LDL1")
'''''                .SXLGD_MS07LDL2 = rs("SXLGD_MS07LDL2")
'''''                .SXLGD_MS07LDL3 = rs("SXLGD_MS07LDL3")
'''''                .SXLGD_MS07LDL4 = rs("SXLGD_MS07LDL4")
'''''                .SXLGD_MS07LDL5 = rs("SXLGD_MS07LDL5")
'''''                .SXLGD_MS07DEN1 = rs("SXLGD_MS07DEN1")
'''''                .SXLGD_MS07DEN2 = rs("SXLGD_MS07DEN2")
'''''                .SXLGD_MS07DEN3 = rs("SXLGD_MS07DEN3")
'''''                .SXLGD_MS07DEN4 = rs("SXLGD_MS07DEN4")
'''''                .SXLGD_MS07DEN5 = rs("SXLGD_MS07DEN5")
'''''                .SXLGD_MS08LDL1 = rs("SXLGD_MS08LDL1")
'''''                .SXLGD_MS08LDL2 = rs("SXLGD_MS08LDL2")
'''''                .SXLGD_MS08LDL3 = rs("SXLGD_MS08LDL3")
'''''                .SXLGD_MS08LDL4 = rs("SXLGD_MS08LDL4")
'''''                .SXLGD_MS08LDL5 = rs("SXLGD_MS08LDL5")
'''''                .SXLGD_MS08DEN1 = rs("SXLGD_MS08DEN1")
'''''                .SXLGD_MS08DEN2 = rs("SXLGD_MS08DEN2")
'''''                .SXLGD_MS08DEN3 = rs("SXLGD_MS08DEN3")
'''''                .SXLGD_MS08DEN4 = rs("SXLGD_MS08DEN4")
'''''                .SXLGD_MS08DEN5 = rs("SXLGD_MS08DEN5")
'''''                .SXLGD_MS09LDL1 = rs("SXLGD_MS09LDL1")
'''''                .SXLGD_MS09LDL2 = rs("SXLGD_MS09LDL2")
'''''                .SXLGD_MS09LDL3 = rs("SXLGD_MS09LDL3")
'''''                .SXLGD_MS09LDL4 = rs("SXLGD_MS09LDL4")
'''''                .SXLGD_MS09LDL5 = rs("SXLGD_MS09LDL5")
'''''                .SXLGD_MS09DEN1 = rs("SXLGD_MS09DEN1")
'''''                .SXLGD_MS09DEN2 = rs("SXLGD_MS09DEN2")
'''''                .SXLGD_MS09DEN3 = rs("SXLGD_MS09DEN3")
'''''                .SXLGD_MS09DEN4 = rs("SXLGD_MS09DEN4")
'''''                .SXLGD_MS09DEN5 = rs("SXLGD_MS09DEN5")
'''''                .SXLGD_MS10LDL1 = rs("SXLGD_MS10LDL1")
'''''                .SXLGD_MS10LDL2 = rs("SXLGD_MS10LDL2")
'''''                .SXLGD_MS10LDL3 = rs("SXLGD_MS10LDL3")
'''''                .SXLGD_MS10LDL4 = rs("SXLGD_MS10LDL4")
'''''                .SXLGD_MS10LDL5 = rs("SXLGD_MS10LDL5")
'''''                .SXLGD_MS10DEN1 = rs("SXLGD_MS10DEN1")
'''''                .SXLGD_MS10DEN2 = rs("SXLGD_MS10DEN2")
'''''                .SXLGD_MS10DEN3 = rs("SXLGD_MS10DEN3")
'''''                .SXLGD_MS10DEN4 = rs("SXLGD_MS10DEN4")
'''''                .SXLGD_MS10DEN5 = rs("SXLGD_MS10DEN5")
'''''                .SXLGD_MS11LDL1 = rs("SXLGD_MS11LDL1")
'''''                .SXLGD_MS11LDL2 = rs("SXLGD_MS11LDL2")
'''''                .SXLGD_MS11LDL3 = rs("SXLGD_MS11LDL3")
'''''                .SXLGD_MS11LDL4 = rs("SXLGD_MS11LDL4")
'''''                .SXLGD_MS11LDL5 = rs("SXLGD_MS11LDL5")
'''''                .SXLGD_MS11DEN1 = rs("SXLGD_MS11DEN1")
'''''                .SXLGD_MS11DEN2 = rs("SXLGD_MS11DEN2")
'''''                .SXLGD_MS11DEN3 = rs("SXLGD_MS11DEN3")
'''''                .SXLGD_MS11DEN4 = rs("SXLGD_MS11DEN4")
'''''                .SXLGD_MS11DEN5 = rs("SXLGD_MS11DEN5")
'''''                .SXLGD_MS12LDL1 = rs("SXLGD_MS12LDL1")
'''''                .SXLGD_MS12LDL2 = rs("SXLGD_MS12LDL2")
'''''                .SXLGD_MS12LDL3 = rs("SXLGD_MS12LDL3")
'''''                .SXLGD_MS12LDL4 = rs("SXLGD_MS12LDL4")
'''''                .SXLGD_MS12LDL5 = rs("SXLGD_MS12LDL5")
'''''                .SXLGD_MS12DEN1 = rs("SXLGD_MS12DEN1")
'''''                .SXLGD_MS12DEN2 = rs("SXLGD_MS12DEN2")
'''''                .SXLGD_MS12DEN3 = rs("SXLGD_MS12DEN3")
'''''                .SXLGD_MS12DEN4 = rs("SXLGD_MS12DEN4")
'''''                .SXLGD_MS12DEN5 = rs("SXLGD_MS12DEN5")
'''''                .SXLGD_MS13LDL1 = rs("SXLGD_MS13LDL1")
'''''                .SXLGD_MS13LDL2 = rs("SXLGD_MS13LDL2")
'''''                .SXLGD_MS13LDL3 = rs("SXLGD_MS13LDL3")
'''''                .SXLGD_MS13LDL4 = rs("SXLGD_MS13LDL4")
'''''                .SXLGD_MS13LDL5 = rs("SXLGD_MS13LDL5")
'''''                .SXLGD_MS13DEN1 = rs("SXLGD_MS13DEN1")
'''''                .SXLGD_MS13DEN2 = rs("SXLGD_MS13DEN2")
'''''                .SXLGD_MS13DEN3 = rs("SXLGD_MS13DEN3")
'''''                .SXLGD_MS13DEN4 = rs("SXLGD_MS13DEN4")
'''''                .SXLGD_MS13DEN5 = rs("SXLGD_MS13DEN5")
'''''                .SXLGD_MS14LDL1 = rs("SXLGD_MS14LDL1")
'''''                .SXLGD_MS14LDL2 = rs("SXLGD_MS14LDL2")
'''''                .SXLGD_MS14LDL3 = rs("SXLGD_MS14LDL3")
'''''                .SXLGD_MS14LDL4 = rs("SXLGD_MS14LDL4")
'''''                .SXLGD_MS14LDL5 = rs("SXLGD_MS14LDL5")
'''''                .SXLGD_MS14DEN1 = rs("SXLGD_MS14DEN1")
'''''                .SXLGD_MS14DEN2 = rs("SXLGD_MS14DEN2")
'''''                .SXLGD_MS14DEN3 = rs("SXLGD_MS14DEN3")
'''''                .SXLGD_MS14DEN4 = rs("SXLGD_MS14DEN4")
'''''                .SXLGD_MS14DEN5 = rs("SXLGD_MS14DEN5")
'''''                .SXLGD_MS15LDL1 = rs("SXLGD_MS15LDL1")
'''''                .SXLGD_MS15LDL2 = rs("SXLGD_MS15LDL2")
'''''                .SXLGD_MS15LDL3 = rs("SXLGD_MS15LDL3")
'''''                .SXLGD_MS15LDL4 = rs("SXLGD_MS15LDL4")
'''''                .SXLGD_MS15LDL5 = rs("SXLGD_MS15LDL5")
'''''                .SXLGD_MS15DEN1 = rs("SXLGD_MS15DEN1")
'''''                .SXLGD_MS15DEN2 = rs("SXLGD_MS15DEN2")
'''''                .SXLGD_MS15DEN3 = rs("SXLGD_MS15DEN3")
'''''                .SXLGD_MS15DEN4 = rs("SXLGD_MS15DEN4")
'''''                .SXLGD_MS15DEN5 = rs("SXLGD_MS15DEN5")
'''''                .SXLGD_MSRSDEN = rs("SXLGD_MSRSDEN")
'''''                .SXLGD_MSRSLDL = rs("SXLGD_MSRSLDL")
'''''                .SXLGD_MSRSDVD2 = rs("SXLGD_MSRSDVD2")
'''''                If IsNull(rs("SXLGD_MS01DVD2")) = False Then .SXLGD_MS01DVD2 = rs("SXLGD_MS01DVD2")
'''''                If IsNull(rs("SXLGD_MS02DVD2")) = False Then .SXLGD_MS02DVD2 = rs("SXLGD_MS02DVD2")
'''''                If IsNull(rs("SXLGD_MS03DVD2")) = False Then .SXLGD_MS03DVD2 = rs("SXLGD_MS03DVD2")
'''''                If IsNull(rs("SXLGD_MS04DVD2")) = False Then .SXLGD_MS04DVD2 = rs("SXLGD_MS04DVD2")
'''''                If IsNull(rs("SXLGD_MS05DVD2")) = False Then .SXLGD_MS05DVD2 = rs("SXLGD_MS05DVD2")
'''''            End If
'''''            If (.SXLT_SMPPOS = -1) And (rs("SXLT_SMPPOS") <> -1) Then
'''''                'LT���т��R�s�[
'''''                .SXLT_SMPPOS = rs("SXLT_SMPPOS")
'''''                .SXLLT_MEASPEAK = rs("SXLLT_MEASPEAK")
'''''                .SXLLT_MEAS1 = rs("SXLLT_MEAS1")
'''''                .SXLLT_MEAS2 = rs("SXLLT_MEAS2")
'''''                .SXLLT_MEAS3 = rs("SXLLT_MEAS3")
'''''                .SXLLT_MEAS4 = rs("SXLLT_MEAS4")
'''''                .SXLLT_MEAS5 = rs("SXLLT_MEAS5")
'''''                .SXLLT_CALCMEAS = rs("SXLLT_CALCMEAS")
'''''            End If
'''''        End If
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End With
'''''End Sub


''''''�T�v      :�������� �����������葪��l�}���p�h���C�o
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :BLOCKID       ,I  ,String       ,�u���b�NID
''''''          :YoneFlg       ,I  ,Boolean      ,�u���b�NID���đ򂩂ǂ����̃t���O�iTrue�͕đ�j
''''''          :Soku          ,I  ,typ_TBCMJ014 ,�����������葪��l�e�[�u���ւ̑}���p
''''''          :�߂�l        ,O  ,FUNCTION_RETURN,�ǂݍ��ݐ���
''''''����      :
''''''����      :2001/06/27 ���{ �쐬
'''''Public Function DBDRV_scmzc_fcmkc001c_InsSoku(BLOCKID As String, _
'''''                                           YoneFlg As Boolean, _
'''''                                           Soku As typ_TBCMJ014 _
'''''                                           ) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim PLUPDATE As String
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_InsSoku"
'''''
'''''    DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_SUCCESS
'''''
'''''
'''''    If YoneFlg = True Then
'''''        '���グ�I�����т��������t�擾 �����񐔂͂Ȃ��Ȃ�
'''''        sql = "select "
'''''        sql = sql & " to_char(REGDATE,'YYYYMMDDHH24MISS') as cDate "
'''''        sql = sql & " from TBCMH004 "
'''''        sql = sql & " where CRYNUM='" & Soku.CRYNUM & "' "
'''''
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''        If rs.RecordCount = 0 Then
'''''            DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_FAILURE
'''''            GoTo proc_exit
'''''        End If
'''''
'''''        PLUPDATE = rs("cDate")
'''''        rs.Close
'''''    Else
'''''        '�w���P�������т�����グ���t�擾
'''''        sql = "select "
'''''        sql = sql & " to_char(REGDATE,'YYYYMMDDHH24MISS') as cDate "
'''''        sql = sql & " from TBCMG002 "
'''''        sql = sql & " where CRYNUM='" & BLOCKID & "' "
'''''        sql = sql & " and TRANCNT=any(select max(TRANCNT) from TBCMG002 where CRYNUM='" & BLOCKID & "' ) "
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''        If rs.RecordCount = 0 Then
'''''            DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_FAILURE
'''''            GoTo proc_exit
'''''        End If
'''''
'''''        PLUPDATE = rs("cDate")
'''''        rs.Close
'''''    End If
'''''
'''''    '���p�̏ꍇ�ɂ́A�������R�[�h����K�v���ڂ��擾����
'''''    UpdateFromOrgJ014 Soku
'''''
'''''    '���p�̏ꍇ�Ɋ��Ƀ��R�[�h�����݂���ꍇ������̂ŁA�܂��폜����
'''''    With Soku
'''''        sql = "delete from TBCMJ014 "
'''''        sql = sql & "where (CRYNUM='" & .CRYNUM & "')"       ' �����ԍ�
'''''        sql = sql & " and (POSITION=" & .POSITION & ")"        ' �ʒu
'''''        sql = sql & " and (SMPKBN='" & .SMPKBN & "') "        ' �T���v���敪
'''''        If 0 > OraDB.ExecuteSQL(sql) Then
'''''            DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_FAILURE
'''''            GoTo proc_exit
'''''        End If
'''''    End With
'''''
'''''    '�����������葪��l�ւ̑}���iTBCMJ014�j
'''''    sql = "insert into TBCMJ014 ( "
'''''    sql = sql & "CRYNUM, "           ' �����ԍ�
'''''    sql = sql & "POSITION, "         ' �ʒu
'''''    sql = sql & "SMPKBN, "           ' �T���v���敪
'''''    sql = sql & "LENGTH, "           ' ����
'''''    sql = sql & "UBLOCKID, "         ' U�u���b�NID
'''''    sql = sql & "DBLOCKID, "         ' D�u���b�NID
'''''    sql = sql & "HINBAN, "           ' �i��
'''''    sql = sql & "REVNUM, "           ' ���i�ԍ������ԍ�
'''''    sql = sql & "FACTORY, "          ' �H��
'''''    sql = sql & "OPECOND, "          ' ���Ə���
'''''    sql = sql & "PRODCOND, "         ' �������
'''''    sql = sql & "PGID, "             ' �o�f�|�h�c
'''''    sql = sql & "UPLENGTH, "         ' ���グ����
'''''    sql = sql & "PLUPDATE, "         ' ������t
'''''    sql = sql & "FREELENG, "         ' �t���[��
'''''    sql = sql & "DIAMETER, "         ' ���a
'''''    sql = sql & "CHARGE, "           ' �`���[�W��
'''''    sql = sql & "SEED, "             ' �V�[�h
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXL_RS_SMPPOS, "   ' SXLRS����ّ���ʒu�iSXL������j
'''''    sql = sql & "SXLRS_MEAS1, "      ' SXLRS_����l�P
'''''    sql = sql & "SXLRS_MEAS2, "      ' SXLRS_����l�Q
'''''    sql = sql & "SXLRS_MEAS3, "      ' SXLRS_����l�R
'''''    sql = sql & "SXLRS_MEAS4, "      ' SXLRS_����l�S
'''''    sql = sql & "SXLRS_MEAS5, "      ' SXLRS_����l�T
'''''    sql = sql & "SXLRS_EFEHS, "      ' SXLRS_�����ΐ�
'''''    sql = sql & "SXLRS_RRG, "        ' SXLRS_�q�q�f
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXL_OI_SMPPOS, "    ' SXLOI����ّ���ʒu�iSXL������j
'''''    sql = sql & "SXLOI_OIMEAS1, "    ' SXLOI_�n������l�P
'''''    sql = sql & "SXLOI_OIMEAS2, "    ' SXLOI_�n������l�Q
'''''    sql = sql & "SXLOI_OIMEAS3, "    ' SXLOI_�n������l�R
'''''    sql = sql & "SXLOI_OIMEAS4, "    ' SXLOI_�n������l�S
'''''    sql = sql & "SXLOI_OIMEAS5, "    ' SXLOI_�n������l�T
'''''    sql = sql & "SXLOI_ORGRES, "     ' SXLOI_�n�q�f����
'''''    sql = sql & "SXLOI_INSPECTWAY, " ' SXLOI_�������@
'''''    sql = sql & "SXL_CS_SMPPOS, "   ' SXLCS����ّ���ʒu�iSXL������j
'''''    sql = sql & "SXLCS_CSMEAS, "     ' SXLCS_Cs�����l
'''''    sql = sql & "SXLCS_70PPRE, "     ' SXLCS_�V�O������l
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF1_SMPPOS, "   ' SXLOSF����ّ���ʒu�iSXL�ʒu���j
'''''    sql = sql & "SXLOSF1_KKSP, "     ' SXLOSF1�������ב���ʒu
'''''    sql = sql & "SXLOSF1_NETU, "     ' SXLOSF1�M�����@
'''''    sql = sql & "SXLOSF1_KKSET, "    ' SXLOSF1�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''    sql = sql & "SXLOSF1_MEAS1, "    ' SXLOSF1����_�P
'''''    sql = sql & "SXLOSF1_MEAS2, "    ' SXLOSF1����_2
'''''    sql = sql & "SXLOSF1_MEAS3, "    ' SXLOSF1����_3
'''''    sql = sql & "SXLOSF1_MEAS4, "    ' SXLOSF1����_4
'''''    sql = sql & "SXLOSF1_MEAS5, "    ' SXLOSF1����_5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF1_MEAS6, "    ' SXLOSF1����_6
'''''    sql = sql & "SXLOSF1_MEAS7, "    ' SXLOSF1����_7
'''''    sql = sql & "SXLOSF1_MEAS8, "    ' SXLOSF1����_8
'''''    sql = sql & "SXLOSF1_MEAS9, "    ' SXLOSF1����_9
'''''    sql = sql & "SXLOSF1_MEAS10, "   ' SXLOSF1����_10
'''''    sql = sql & "SXLOSF1_MEAS11, "   ' SXLOSF1����_11
'''''    sql = sql & "SXLOSF1_MEAS12, "   ' SXLOSF1����_12
'''''    sql = sql & "SXLOSF1_MEAS13, "   ' SXLOSF1����_13
'''''    sql = sql & "SXLOSF1_MEAS14, "   ' SXLOSF1����_14
'''''    sql = sql & "SXLOSF1_MEAS15, "   ' SXLOSF1����_15
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF1_MEAS16, "   ' SXLOSF1����_16
'''''    sql = sql & "SXLOSF1_MEAS17, "   ' SXLOSF1����_17
'''''    sql = sql & "SXLOSF1_MEAS18, "   ' SXLOSF1����_18
'''''    sql = sql & "SXLOSF1_MEAS19, "   ' SXLOSF1����_19
'''''    sql = sql & "SXLOSF1_MEAS20, "   ' SXLOSF1����_20
'''''    sql = sql & "SXLOSF1_CALCMAX, "  ' OSF1SXL�v�Z���� Max_1
'''''    sql = sql & "SXLOSF1_CALCAVE, "  ' OSF1SXL�v�Z���� Ave_1
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF2_KKSP, "    ' SXLOSF�Q�������ב���ʒu
'''''    sql = sql & "SXLOSF2_NETU, "     ' SXLOSF�Q�M�����@
'''''    sql = sql & "SXLOSF2_KKSET, "    ' SXLOSF�Q�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''    sql = sql & "SXLOSF2_MEAS1, "    ' SXLOSF2����_�P
'''''    sql = sql & "SXLOSF2_MEAS2, "    ' SXLOSF2����_2
'''''    sql = sql & "SXLOSF2_MEAS3, "    ' SXLOSF2����_3
'''''    sql = sql & "SXLOSF2_MEAS4, "    ' SXLOSF2����_4
'''''    sql = sql & "SXLOSF2_MEAS5, "    ' SXLOSF2����_5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF2_MEAS6, "    ' SXLOSF2����_6
'''''    sql = sql & "SXLOSF2_MEAS7, "    ' SXLOSF2����_7
'''''    sql = sql & "SXLOSF2_MEAS8, "    ' SXLOSF2����_8
'''''    sql = sql & "SXLOSF2_MEAS9, "    ' SXLOSF2����_9
'''''    sql = sql & "SXLOSF2_MEAS10, "   ' SXLOSF2����_10
'''''    sql = sql & "SXLOSF2_MEAS11, "   ' SXLOSF2����_11
'''''    sql = sql & "SXLOSF2_MEAS12, "   ' SXLOSF2����_12
'''''    sql = sql & "SXLOSF2_MEAS13, "   ' SXLOSF2����_13
'''''    sql = sql & "SXLOSF2_MEAS14, "   ' SXLOSF2����_14
'''''    sql = sql & "SXLOSF2_MEAS15, "   ' SXLOSF2����_15
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF2_MEAS16, "   ' SXLOSF2����_16
'''''    sql = sql & "SXLOSF2_MEAS17, "   ' SXLOSF2����_17
'''''    sql = sql & "SXLOSF2_MEAS18, "   ' SXLOSF2����_18
'''''    sql = sql & "SXLOSF2_MEAS19, "   ' SXLOSF2����_19
'''''    sql = sql & "SXLOSF2_MEAS20, "   ' SXLOSF2����_20
'''''    sql = sql & "SXLOSF2_CALCMAX, "  ' OSF�QSXL�v�Z���� Max_2
'''''    sql = sql & "SXLOSF2_CALCAVE, "  ' OSF�QSXL�v�Z���� Ave_2
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF3_KKSP, "   ' SXLOSF�R�������ב���ʒu
'''''    sql = sql & "SXLOSF3_NETU, "     ' SXLOSF�R�M�����@
'''''    sql = sql & "SXLOSF3_KKSET, "    ' SXLOSF�R�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''    sql = sql & "SXLOSF3_MEAS1, "    ' SXLOSF3����_�P
'''''    sql = sql & "SXLOSF3_MEAS2, "    ' SXLOSF3����_2
'''''    sql = sql & "SXLOSF3_MEAS3, "    ' SXLOSF3����_3
'''''    sql = sql & "SXLOSF3_MEAS4, "    ' SXLOSF3����_4
'''''    sql = sql & "SXLOSF3_MEAS5, "    ' SXLOSF3����_5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF3_MEAS6, "    ' SXLOSF3����_6
'''''    sql = sql & "SXLOSF3_MEAS7, "    ' SXLOSF3����_7
'''''    sql = sql & "SXLOSF3_MEAS8, "    ' SXLOSF3����_8
'''''    sql = sql & "SXLOSF3_MEAS9, "    ' SXLOSF3����_9
'''''    sql = sql & "SXLOSF3_MEAS10, "   ' SXLOSF3����_10
'''''    sql = sql & "SXLOSF3_MEAS11, "   ' SXLOSF3����_11
'''''    sql = sql & "SXLOSF3_MEAS12, "   ' SXLOSF3����_12
'''''    sql = sql & "SXLOSF3_MEAS13, "   ' SXLOSF3����_13
'''''    sql = sql & "SXLOSF3_MEAS14, "   ' SXLOSF3����_14
'''''    sql = sql & "SXLOSF3_MEAS15, "   ' SXLOSF3����_15
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF3_MEAS16, "   ' SXLOSF3����_16
'''''    sql = sql & "SXLOSF3_MEAS17, "   ' SXLOSF3����_17
'''''    sql = sql & "SXLOSF3_MEAS18, "   ' SXLOSF3����_18
'''''    sql = sql & "SXLOSF3_MEAS19, "   ' SXLOSF3����_19
'''''    sql = sql & "SXLOSF3_MEAS20, "   ' SXLOSF3����_20
'''''    sql = sql & "SXLOSF3_CALCMAX, "  ' OSF�RSXL�v�Z���� Max_3
'''''    sql = sql & "SXLOSF3_CALCAVE, "  ' OSF�RSXL�v�Z���� Ave_3
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF4_KKSP, "   ' SXLOSF�S�������ב���ʒu
'''''    sql = sql & "SXLOSF4_NETU, "     ' SXLOSF�S�M�����@
'''''    sql = sql & "SXLOSF4_KKSET, "    ' SXLOSF�S�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''    sql = sql & "SXLOSF4_MEAS1, "    ' SXLOSF4����_�P
'''''    sql = sql & "SXLOSF4_MEAS2, "    ' SXLOSF4����_2
'''''    sql = sql & "SXLOSF4_MEAS3, "    ' SXLOSF4����_3
'''''    sql = sql & "SXLOSF4_MEAS4, "    ' SXLOSF4����_4
'''''    sql = sql & "SXLOSF4_MEAS5, "    ' SXLOSF4����_5
'''''    sql = sql & "SXLOSF4_MEAS6, "    ' SXLOSF4����_6
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF4_MEAS7, "    ' SXLOSF4����_7
'''''    sql = sql & "SXLOSF4_MEAS8, "    ' SXLOSF4����_8
'''''    sql = sql & "SXLOSF4_MEAS9, "    ' SXLOSF4����_9
'''''    sql = sql & "SXLOSF4_MEAS10, "   ' SXLOSF4����_10
'''''    sql = sql & "SXLOSF4_MEAS11, "   ' SXLOSF4����_11
'''''    sql = sql & "SXLOSF4_MEAS12, "   ' SXLOSF4����_12
'''''    sql = sql & "SXLOSF4_MEAS13, "   ' SXLOSF4����_13
'''''    sql = sql & "SXLOSF4_MEAS14, "   ' SXLOSF4����_14
'''''    sql = sql & "SXLOSF4_MEAS15, "   ' SXLOSF4����_15
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF4_MEAS16, "   ' SXLOSF4����_16
'''''    sql = sql & "SXLOSF4_MEAS17, "   ' SXLOSF4����_17
'''''    sql = sql & "SXLOSF4_MEAS18, "   ' SXLOSF4����_18
'''''    sql = sql & "SXLOSF4_MEAS19, "   ' SXLOSF4����_19
'''''    sql = sql & "SXLOSF4_MEAS20, "   ' SXLOSF4����_20
'''''    sql = sql & "SXLOSF4_CALCMAX, "  ' OSF�SSXL�v�Z���� Max_4
'''''    sql = sql & "SXLOSF4_CALCAVE, "  ' OSF�SSXL�v�Z���� Ave_4
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLBMD_SMPPOS, "   ' SXLBMD����ّ���ʒu�iSXL�ʒu���j
'''''    sql = sql & "SXLBMD1_KKSP, "     ' SXLBMD1�������ב���ʒu
'''''    sql = sql & "SXLBMD1_NETU, "     ' SXLBMD1�M�����@
'''''    sql = sql & "SXLBMD1_KKSET, "    ' SXLBMD1�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''    sql = sql & "SXLBMD1_MEAS1, "    ' SXLBMD1����_�P
'''''    sql = sql & "SXLBMD1_MEAS2, "    ' SXLBMD1����_2
'''''    sql = sql & "SXLBMD1_MEAS3, "    ' SXLBMD1����_3
'''''    sql = sql & "SXLBMD1_MEAS4, "    ' SXLBMD1����_4
'''''    sql = sql & "SXLBMD1_MEAS5, "    ' SXLBMD1����_5
'''''    sql = sql & "SXLBMD1_CALCMAX, "  ' BMD1SXL�v�Z���� Max
'''''    sql = sql & "SXLBMD1_CALCAVE, "  ' BMD1SXL�v�Z���� Ave
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLBMD2_KKSP, "    ' SXLBMD�Q�������ב���ʒu
'''''    sql = sql & "SXLBMD2_NETU, "     ' SXLBMD�Q�M�����@
'''''    sql = sql & "SXLBMD2_KKSET, "    ' SXLBMD�Q�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''    sql = sql & "SXLBMD2_MEAS1, "    ' SXLBMD2����_�P
'''''    sql = sql & "SXLBMD2_MEAS2, "    ' SXLBMD2����_2
'''''    sql = sql & "SXLBMD2_MEAS3, "    ' SXLBMD2����_3
'''''    sql = sql & "SXLBMD2_MEAS4, "    ' SXLBMD2����_4
'''''    sql = sql & "SXLBMD2_MEAS5, "    ' SXLBMD2����_5
'''''    sql = sql & "SXLBMD2_CALCMAX, "  ' BMD�QSXL�v�Z���� Max
'''''    sql = sql & "SXLBMD2_CALCAVE, "  ' BMD�QSXL�v�Z���� Ave
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLBMD3_KKSP, "   ' SXLBMD�R�������ב���ʒu
'''''    sql = sql & "SXLBMD3_NETU, "     ' SXLBMD�R�M�����@
'''''    sql = sql & "SXLBMD3_KKSET, "    ' SXLBMD�R�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''    sql = sql & "SXLBMD3_MEAS1, "    ' SXLBMD3����_�P
'''''    sql = sql & "SXLBMD3_MEAS2, "    ' SXLBMD3����_2
'''''    sql = sql & "SXLBMD3_MEAS3, "    ' SXLBMD3����_3
'''''    sql = sql & "SXLBMD3_MEAS4, "    ' SXLBMD3����_4
'''''    sql = sql & "SXLBMD3_MEAS5, "    ' SXLBMD3����_5
'''''    sql = sql & "SXLBMD3_CALCMAX, "  ' BMD�RSXL�v�Z���� Max
'''''    sql = sql & "SXLBMD3_CALCAVE, "  ' BMD�RSXL�v�Z���� Ave
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_SMPPOS, "     ' SXLGD����ّ���ʒu�iSXL�ʒu���j
'''''    sql = sql & "SXLGD_MS01LDL1, "   ' SXLGD_����l01 L/DL1
'''''    sql = sql & "SXLGD_MS01LDL2, "   ' SXLGD_����l01 L/DL2
'''''    sql = sql & "SXLGD_MS01LDL3, "   ' SXLGD_����l01 L/DL3
'''''    sql = sql & "SXLGD_MS01LDL4, "   ' SXLGD_����l01 L/DL4
'''''    sql = sql & "SXLGD_MS01LDL5, "   ' SXLGD_����l01 L/DL5
'''''    sql = sql & "SXLGD_MS01DEN1, "   ' SXLGD_����l01 Den1
'''''    sql = sql & "SXLGD_MS01DEN2, "   ' SXLGD_����l01 Den2
'''''    sql = sql & "SXLGD_MS01DEN3, "   ' SXLGD_����l01 Den3
'''''    sql = sql & "SXLGD_MS01DEN4, "   ' SXLGD_����l01 Den4
'''''    sql = sql & "SXLGD_MS01DEN5, "   ' SXLGD_����l01 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS02LDL1, "   ' SXLGD_����l02 L/DL1
'''''    sql = sql & "SXLGD_MS02LDL2, "   ' SXLGD_����l02 L/DL2
'''''    sql = sql & "SXLGD_MS02LDL3, "   ' SXLGD_����l02 L/DL3
'''''    sql = sql & "SXLGD_MS02LDL4, "   ' SXLGD_����l02 L/DL4
'''''    sql = sql & "SXLGD_MS02LDL5, "   ' SXLGD_����l02 L/DL5
'''''    sql = sql & "SXLGD_MS02DEN1, "   ' SXLGD_����l02 Den1
'''''    sql = sql & "SXLGD_MS02DEN2, "   ' SXLGD_����l02 Den2
'''''    sql = sql & "SXLGD_MS02DEN3, "   ' SXLGD_����l02 Den3
'''''    sql = sql & "SXLGD_MS02DEN4, "   ' SXLGD_����l02 Den4
'''''    sql = sql & "SXLGD_MS02DEN5, "   ' SXLGD_����l02 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS03LDL1, "   ' SXLGD_����l03 L/DL1
'''''    sql = sql & "SXLGD_MS03LDL2, "   ' SXLGD_����l03 L/DL2
'''''    sql = sql & "SXLGD_MS03LDL3, "   ' SXLGD_����l03 L/DL3
'''''    sql = sql & "SXLGD_MS03LDL4, "   ' SXLGD_����l03 L/DL4
'''''    sql = sql & "SXLGD_MS03LDL5, "   ' SXLGD_����l03 L/DL5
'''''    sql = sql & "SXLGD_MS03DEN1, "   ' SXLGD_����l03 Den1
'''''    sql = sql & "SXLGD_MS03DEN2, "   ' SXLGD_����l03 Den2
'''''    sql = sql & "SXLGD_MS03DEN3, "   ' SXLGD_����l03 Den3
'''''    sql = sql & "SXLGD_MS03DEN4, "   ' SXLGD_����l03 Den4
'''''    sql = sql & "SXLGD_MS03DEN5, "  ' SXLGD_����l03 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS04LDL1, "   ' SXLGD_����l04 L/DL1
'''''    sql = sql & "SXLGD_MS04LDL2, "   ' SXLGD_����l04 L/DL2
'''''    sql = sql & "SXLGD_MS04LDL3, "   ' SXLGD_����l04 L/DL3
'''''    sql = sql & "SXLGD_MS04LDL4, "   ' SXLGD_����l04 L/DL4
'''''    sql = sql & "SXLGD_MS04LDL5, "   ' SXLGD_����l04 L/DL5
'''''    sql = sql & "SXLGD_MS04DEN1, "   ' SXLGD_����l04 Den1
'''''    sql = sql & "SXLGD_MS04DEN2, "   ' SXLGD_����l04 Den2
'''''    sql = sql & "SXLGD_MS04DEN3, "   ' SXLGD_����l04 Den3
'''''    sql = sql & "SXLGD_MS04DEN4, "   ' SXLGD_����l04 Den4
'''''    sql = sql & "SXLGD_MS04DEN5, "   ' SXLGD_����l04 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS05LDL1, "   ' SXLGD_����l05 L/DL1
'''''    sql = sql & "SXLGD_MS05LDL2, "   ' SXLGD_����l05 L/DL2
'''''    sql = sql & "SXLGD_MS05LDL3, "   ' SXLGD_����l05 L/DL3
'''''    sql = sql & "SXLGD_MS05LDL4, "   ' SXLGD_����l05 L/DL4
'''''    sql = sql & "SXLGD_MS05LDL5, "   ' SXLGD_����l05 L/DL5
'''''    sql = sql & "SXLGD_MS05DEN1, "   ' SXLGD_����l05 Den1
'''''    sql = sql & "SXLGD_MS05DEN2, "   ' SXLGD_����l05 Den2
'''''    sql = sql & "SXLGD_MS05DEN3, "   ' SXLGD_����l05 Den3
'''''    sql = sql & "SXLGD_MS05DEN4, "   ' SXLGD_����l05 Den4
'''''    sql = sql & "SXLGD_MS05DEN5, "   ' SXLGD_����l05 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS06LDL1, "   ' SXLGD_����l06 L/DL1
'''''    sql = sql & "SXLGD_MS06LDL2, "   ' SXLGD_����l06 L/DL2
'''''    sql = sql & "SXLGD_MS06LDL3, "   ' SXLGD_����l06 L/DL3
'''''    sql = sql & "SXLGD_MS06LDL4, "   ' SXLGD_����l06 L/DL4
'''''    sql = sql & "SXLGD_MS06LDL5, "   ' SXLGD_����l06 L/DL5
'''''    sql = sql & "SXLGD_MS06DEN1, "   ' SXLGD_����l06 Den1
'''''    sql = sql & "SXLGD_MS06DEN2, "   ' SXLGD_����l06 Den2
'''''    sql = sql & "SXLGD_MS06DEN3, "   ' SXLGD_����l06 Den3
'''''    sql = sql & "SXLGD_MS06DEN4, "   ' SXLGD_����l06 Den4
'''''    sql = sql & "SXLGD_MS06DEN5, "   ' SXLGD_����l06 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS07LDL1, "   ' SXLGD_����l07 L/DL1
'''''    sql = sql & "SXLGD_MS07LDL2, "   ' SXLGD_����l07 L/DL2
'''''    sql = sql & "SXLGD_MS07LDL3, "   ' SXLGD_����l07 L/DL3
'''''    sql = sql & "SXLGD_MS07LDL4, "   ' SXLGD_����l07 L/DL4
'''''    sql = sql & "SXLGD_MS07LDL5, "   ' SXLGD_����l07 L/DL5
'''''    sql = sql & "SXLGD_MS07DEN1, "   ' SXLGD_����l07 Den1
'''''    sql = sql & "SXLGD_MS07DEN2, "   ' SXLGD_����l07 Den2
'''''    sql = sql & "SXLGD_MS07DEN3, "   ' SXLGD_����l07 Den3
'''''    sql = sql & "SXLGD_MS07DEN4, "   ' SXLGD_����l07 Den4
'''''    sql = sql & "SXLGD_MS07DEN5, "   ' SXLGD_����l07 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS08LDL1, "   ' SXLGD_����l08 L/DL1
'''''    sql = sql & "SXLGD_MS08LDL2, "   ' SXLGD_����l08 L/DL2
'''''    sql = sql & "SXLGD_MS08LDL3, "   ' SXLGD_����l08 L/DL3
'''''    sql = sql & "SXLGD_MS08LDL4, "   ' SXLGD_����l08 L/DL4
'''''    sql = sql & "SXLGD_MS08LDL5, "   ' SXLGD_����l08 L/DL5
'''''    sql = sql & "SXLGD_MS08DEN1, "   ' SXLGD_����l08 Den1
'''''    sql = sql & "SXLGD_MS08DEN2, "   ' SXLGD_����l08 Den2
'''''    sql = sql & "SXLGD_MS08DEN3, "   ' SXLGD_����l08 Den3
'''''    sql = sql & "SXLGD_MS08DEN4, "   ' SXLGD_����l08 Den4
'''''    sql = sql & "SXLGD_MS08DEN5, "   ' SXLGD_����l08 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS09LDL1, "   ' SXLGD_����l09 L/DL1
'''''    sql = sql & "SXLGD_MS09LDL2, "   ' SXLGD_����l09 L/DL2
'''''    sql = sql & "SXLGD_MS09LDL3, "   ' SXLGD_����l09 L/DL3
'''''    sql = sql & "SXLGD_MS09LDL4, "   ' SXLGD_����l09 L/DL4
'''''    sql = sql & "SXLGD_MS09LDL5, "   ' SXLGD_����l09 L/DL5
'''''    sql = sql & "SXLGD_MS09DEN1, "   ' SXLGD_����l09 Den1
'''''    sql = sql & "SXLGD_MS09DEN2, "   ' SXLGD_����l09 Den2
'''''    sql = sql & "SXLGD_MS09DEN3, "   ' SXLGD_����l09 Den3
'''''    sql = sql & "SXLGD_MS09DEN4, "   ' SXLGD_����l09 Den4
'''''    sql = sql & "SXLGD_MS09DEN5, "   ' SXLGD_����l09 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS10LDL1, "   ' SXLGD_����l10 L/DL1
'''''    sql = sql & "SXLGD_MS10LDL2, "   ' SXLGD_����l10 L/DL2
'''''    sql = sql & "SXLGD_MS10LDL3, "   ' SXLGD_����l10 L/DL3
'''''    sql = sql & "SXLGD_MS10LDL4, "   ' SXLGD_����l10 L/DL4
'''''    sql = sql & "SXLGD_MS10LDL5, "   ' SXLGD_����l10 L/DL5
'''''    sql = sql & "SXLGD_MS10DEN1, "   ' SXLGD_����l10 Den1
'''''    sql = sql & "SXLGD_MS10DEN2, "   ' SXLGD_����l10 Den2
'''''    sql = sql & "SXLGD_MS10DEN3, "   ' SXLGD_����l10 Den3
'''''    sql = sql & "SXLGD_MS10DEN4, "   ' SXLGD_����l10 Den4
'''''    sql = sql & "SXLGD_MS10DEN5, "   ' SXLGD_����l10 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS11LDL1, "   ' SXLGD_����l11 L/DL1
'''''    sql = sql & "SXLGD_MS11LDL2, "   ' SXLGD_����l11 L/DL2
'''''    sql = sql & "SXLGD_MS11LDL3, "   ' SXLGD_����l11 L/DL3
'''''    sql = sql & "SXLGD_MS11LDL4, "   ' SXLGD_����l11 L/DL4
'''''    sql = sql & "SXLGD_MS11LDL5, "   ' SXLGD_����l11 L/DL5
'''''    sql = sql & "SXLGD_MS11DEN1, "   ' SXLGD_����l11 Den1
'''''    sql = sql & "SXLGD_MS11DEN2, "   ' SXLGD_����l11 Den2
'''''    sql = sql & "SXLGD_MS11DEN3, "   ' SXLGD_����l11 Den3
'''''    sql = sql & "SXLGD_MS11DEN4, "   ' SXLGD_����l11 Den4
'''''    sql = sql & "SXLGD_MS11DEN5, "   ' SXLGD_����l11 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS12LDL1, "   ' SXLGD_����l12 L/DL1
'''''    sql = sql & "SXLGD_MS12LDL2, "   ' SXLGD_����l12 L/DL2
'''''    sql = sql & "SXLGD_MS12LDL3, "   ' SXLGD_����l12 L/DL3
'''''    sql = sql & "SXLGD_MS12LDL4, "   ' SXLGD_����l12 L/DL4
'''''    sql = sql & "SXLGD_MS12LDL5, "   ' SXLGD_����l12 L/DL5
'''''    sql = sql & "SXLGD_MS12DEN1, "   ' SXLGD_����l12 Den1
'''''    sql = sql & "SXLGD_MS12DEN2, "   ' SXLGD_����l12 Den2
'''''    sql = sql & "SXLGD_MS12DEN3, "   ' SXLGD_����l12 Den3
'''''    sql = sql & "SXLGD_MS12DEN4, "   ' SXLGD_����l12 Den4
'''''    sql = sql & "SXLGD_MS12DEN5, "   ' SXLGD_����l12 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS13LDL1, "   ' SXLGD_����l13 L/DL1
'''''    sql = sql & "SXLGD_MS13LDL2, "   ' SXLGD_����l13 L/DL2
'''''    sql = sql & "SXLGD_MS13LDL3, "   ' SXLGD_����l13 L/DL3
'''''    sql = sql & "SXLGD_MS13LDL4, "   ' SXLGD_����l13 L/DL4
'''''    sql = sql & "SXLGD_MS13LDL5, "   ' SXLGD_����l13 L/DL5
'''''    sql = sql & "SXLGD_MS13DEN1, "   ' SXLGD_����l13 Den1
'''''    sql = sql & "SXLGD_MS13DEN2, "   ' SXLGD_����l13 Den2
'''''    sql = sql & "SXLGD_MS13DEN3, "   ' SXLGD_����l13 Den3
'''''    sql = sql & "SXLGD_MS13DEN4, "   ' SXLGD_����l13 Den4
'''''    sql = sql & "SXLGD_MS13DEN5, "   ' SXLGD_����l13 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS14LDL1, "   ' SXLGD_����l14 L/DL1
'''''    sql = sql & "SXLGD_MS14LDL2, "   ' SXLGD_����l14 L/DL2
'''''    sql = sql & "SXLGD_MS14LDL3, "   ' SXLGD_����l14 L/DL3
'''''    sql = sql & "SXLGD_MS14LDL4, "   ' SXLGD_����l14 L/DL4
'''''    sql = sql & "SXLGD_MS14LDL5, "   ' SXLGD_����l14 L/DL5
'''''    sql = sql & "SXLGD_MS14DEN1, "   ' SXLGD_����l14 Den1
'''''    sql = sql & "SXLGD_MS14DEN2, "   ' SXLGD_����l14 Den2
'''''    sql = sql & "SXLGD_MS14DEN3, "   ' SXLGD_����l14 Den3
'''''    sql = sql & "SXLGD_MS14DEN4, "   ' SXLGD_����l14 Den4
'''''    sql = sql & "SXLGD_MS14DEN5, "   ' SXLGD_����l14 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS15LDL1, "   ' SXLGD_����l15 L/DL1
'''''    sql = sql & "SXLGD_MS15LDL2, "   ' SXLGD_����l15 L/DL2
'''''    sql = sql & "SXLGD_MS15LDL3, "   ' SXLGD_����l15 L/DL3
'''''    sql = sql & "SXLGD_MS15LDL4, "   ' SXLGD_����l15 L/DL4
'''''    sql = sql & "SXLGD_MS15LDL5, "   ' SXLGD_����l15 L/DL5
'''''    sql = sql & "SXLGD_MS15DEN1, "   ' SXLGD_����l15 Den1
'''''    sql = sql & "SXLGD_MS15DEN2, "   ' SXLGD_����l15 Den2
'''''    sql = sql & "SXLGD_MS15DEN3, "   ' SXLGD_����l15 Den3
'''''    sql = sql & "SXLGD_MS15DEN4, "   ' SXLGD_����l15 Den4
'''''    sql = sql & "SXLGD_MS15DEN5, "   ' SXLGD_����l15 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MSRSDEN, "   ' SXLGD_���茋�� Den
'''''    sql = sql & "SXLGD_MSRSLDL, "    ' SXLGD_���茋�� L/DL
'''''    sql = sql & "SXLGD_MSRSDVD2, "   ' SXLGD_���茋�� DVD2
'''''    sql = sql & "SXLT_SMPPOS, "      ' SXLLT����ّ���ʒu�iSXL�ʒu���j
'''''    sql = sql & "SXLLT_MEASPEAK, "   ' SXLLT_����l �s�[�N�l
'''''    sql = sql & "SXLLT_MEAS1, "      ' SXLLT_����l1
'''''    sql = sql & "SXLLT_MEAS2, "      ' SXLLT_����l2
'''''    sql = sql & "SXLLT_MEAS3, "      ' SXLLT_����l3
'''''    sql = sql & "SXLLT_MEAS4, "      ' SXLLT_����l4
'''''    sql = sql & "SXLLT_MEAS5, "      ' SXLLT_����l5
'''''    sql = sql & "SXLLT_CALCMEAS, "   ' SXLLT_�v�Z����
'''''    sql = sql & "REGDATE, "          ' �o�^���t
'''''    sql = sql & "SENDFLAG, "         ' ���M�t���O
'''''    sql = sql & "SENDDATE,  "         ' ���M���t
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF1_POS1, "      ' OSF1����݋敪�P�ʒu
'''''    sql = sql & "SXLOSF1_WID1, "      ' OSF1����݋敪�P��
'''''    sql = sql & "SXLOSF1_RD1, "       ' OSF1����݋敪�PR/D
'''''    sql = sql & "SXLOSF1_POS2, "      ' OSF1����݋敪�Q�ʒu
'''''    sql = sql & "SXLOSF1_WID2, "      ' OSF1����݋敪�Q��
'''''    sql = sql & "SXLOSF1_RD2, "       ' OSF1����݋敪�QR/D
'''''    sql = sql & "SXLOSF1_POS3, "      ' OSF1����݋敪�R�ʒu
'''''    sql = sql & "SXLOSF1_WID3, "      ' OSF1����݋敪�R��
'''''    sql = sql & "SXLOSF1_RD3, "       ' OSF1����݋敪�RR/D
'''''    sql = sql & "SXLOSF2_POS1, "      ' OSF2����݋敪�P�ʒu
'''''    sql = sql & "SXLOSF2_WID1, "      ' OSF2����݋敪�P��
'''''    sql = sql & "SXLOSF2_RD1, "       ' OSF2����݋敪�PR/D
'''''    sql = sql & "SXLOSF2_POS2, "      ' OSF2����݋敪�Q�ʒu
'''''    sql = sql & "SXLOSF2_WID2, "      ' OSF2����݋敪�Q��
'''''    sql = sql & "SXLOSF2_RD2, "       ' OSF2����݋敪�QR/D
'''''    sql = sql & "SXLOSF2_POS3, "      ' OSF2����݋敪�R�ʒu
'''''    sql = sql & "SXLOSF2_WID3, "      ' OSF2����݋敪�R��
'''''    sql = sql & "SXLOSF2_RD3, "       ' OSF2����݋敪�RR/D
'''''    sql = sql & "SXLOSF3_POS1, "      ' OSF3����݋敪�P�ʒu
'''''    sql = sql & "SXLOSF3_WID1, "      ' OSF3����݋敪�P��
'''''    sql = sql & "SXLOSF3_RD1, "       ' OSF3����݋敪�PR/D
'''''    sql = sql & "SXLOSF3_POS2, "      ' OSF3����݋敪�Q�ʒu
'''''    sql = sql & "SXLOSF3_WID2, "      ' OSF3����݋敪�Q��
'''''    sql = sql & "SXLOSF3_RD2, "       ' OSF3����݋敪�QR/D
'''''    sql = sql & "SXLOSF3_POS3, "      ' OSF3����݋敪�R�ʒu
'''''    sql = sql & "SXLOSF3_WID3, "      ' OSF3����݋敪�R��
'''''    sql = sql & "SXLOSF3_RD3, "       ' OSF3����݋敪�RR/D
'''''    sql = sql & "SXLOSF4_POS1, "      ' OSF4����݋敪�P�ʒu
'''''    sql = sql & "SXLOSF4_WID1, "      ' OSF4����݋敪�P��
'''''    sql = sql & "SXLOSF4_RD1, "       ' OSF4����݋敪�PR/D
'''''    sql = sql & "SXLOSF4_POS2, "      ' OSF4����݋敪�Q�ʒu
'''''    sql = sql & "SXLOSF4_WID2, "      ' OSF4����݋敪�Q��
'''''    sql = sql & "SXLOSF4_RD2, "       ' OSF4����݋敪�QR/D
'''''    sql = sql & "SXLOSF4_POS3, "      ' OSF4����݋敪�R�ʒu
'''''    sql = sql & "SXLOSF4_WID3, "      ' OSF4����݋敪�R��
'''''    sql = sql & "SXLOSF4_RD3, "       ' OSF4����݋敪�RR/D
'''''    sql = sql & "SXLGD_MS01DVD2, "    ' DVD2���茋�ʒl�P
'''''    sql = sql & "SXLGD_MS02DVD2, "    ' DVD2���茋�ʒl�Q
'''''    sql = sql & "SXLGD_MS03DVD2, "    ' DVD2���茋�ʒl�R
'''''    sql = sql & "SXLGD_MS04DVD2, "    ' DVD2���茋�ʒl�S
'''''    sql = sql & "SXLGD_MS05DVD2, "    ' DVD2���茋�ʒl�T
'''''    sql = sql & "SXLBMD1_MNBCR, "     ' BMD1SXL�v�Z���ʖʓ����z
'''''    sql = sql & "SXLBMD2_MNBCR, "     ' BMD2SXL�v�Z���ʖʓ����z
'''''    sql = sql & "SXLBMD3_MNBCR ) "    ' BMD3SXL�v�Z���ʖʓ����z
'''''    With Soku
'''''        sql = sql & " values ( "
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " '" & .CRYNUM & "', "       ' �����ԍ�
'''''        sql = sql & " " & .POSITION & ", "        ' �ʒu
'''''        sql = sql & " '" & .SMPKBN & "', "        ' �T���v���敪
'''''        sql = sql & " " & .LENGTH & ", "          ' ����
'''''        sql = sql & " '" & .UBLOCKID & "', "      ' U�u���b�NID
'''''        sql = sql & " '" & .DBLOCKID & "', "      ' D�u���b�NID
'''''        sql = sql & " '" & .hinban & "', "        ' �i��
'''''        sql = sql & " " & .REVNUM & ", "          ' ���i�ԍ������ԍ�
'''''        sql = sql & " '" & .factory & "', "       ' �H��
'''''        sql = sql & " '" & .opecond & "', "       ' ���Ə���
'''''        sql = sql & " '" & .PRODCOND & "', "      ' �������
'''''        sql = sql & " '" & Mid(.PGID, 1, 8) & "', "        ' �o�f�|�h�c
'''''        sql = sql & " " & .UPLENGTH & ", "        ' ���グ����
'''''        sql = sql & " to_date('" & PLUPDATE & "','YYYYMMDDHH24MISS'), "    ' ������t
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .FREELENG & ", "        ' �t���[��
'''''        sql = sql & " " & .DIAMETER & ", "        ' ���a
'''''        sql = sql & " '" & .CHARGE & "', "        ' �`���[�W��
'''''        sql = sql & " '" & .SEED & "', "          ' �V�[�h
'''''        sql = sql & " " & .SXL_RS_SMPPOS & ", "  ' SXLRS����ّ���ʒu�iSXL������j
'''''        sql = sql & " " & .SXLRS_MEAS1 & ", "     ' SXLRS_����l�P
'''''        sql = sql & " " & .SXLRS_MEAS2 & ", "     ' SXLRS_����l�Q
'''''        sql = sql & " " & .SXLRS_MEAS3 & ", "     ' SXLRS_����l�R
'''''        sql = sql & " " & .SXLRS_MEAS4 & ", "     ' SXLRS_����l�S
'''''        sql = sql & " " & .SXLRS_MEAS5 & ", "     ' SXLRS_����l�T
'''''        sql = sql & " " & .SXLRS_EFEHS & ", "     ' SXLRS_�����ΐ�
'''''        sql = sql & " " & .SXLRS_RRG & ", "       ' SXLRS_�q�q�f
'''''        sql = sql & " " & .SXL_OI_SMPPOS & ", "   ' SXLOI����ّ���ʒu�iSXL������j
'''''        sql = sql & " " & .SXLOI_OIMEAS1 & ", "   ' SXLOI_�n������l�P
'''''        sql = sql & " " & .SXLOI_OIMEAS2 & ", "   ' SXLOI_�n������l�Q
'''''        sql = sql & " " & .SXLOI_OIMEAS3 & ", "   ' SXLOI_�n������l�R
'''''        sql = sql & " " & .SXLOI_OIMEAS4 & ", "   ' SXLOI_�n������l�S
'''''        sql = sql & " " & .SXLOI_OIMEAS5 & ", "   ' SXLOI_�n������l�T
'''''        sql = sql & " " & .SXLOI_ORGRES & ", "    ' SXLOI_�n�q�f����
'''''        sql = sql & " '" & .SXLOI_INSPECTWAY & "', " ' SXLOI_�������@
'''''        sql = sql & " " & .SXL_CS_SMPPOS & ", "      ' SXLCS����ّ���ʒu�iSXL������j
'''''        sql = sql & " " & .SXLCS_CSMEAS & ", "       ' SXLCS_Cs�����l
'''''        sql = sql & " " & .SXLCS_70PPRE & ", "       ' SXLCS_�V�O������l
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLOSF1_SMPPOS & ", "    ' SXLOSF����ّ���ʒu�iSXL�ʒu���j
'''''        sql = sql & " '" & .SXLOSF1_KKSP & "', "    ' SXLOSF1�������ב���ʒu
'''''        sql = sql & " '" & .SXLOSF1_NETU & "', "    ' SXLOSF1�M�����@
'''''        sql = sql & " '" & .SXLOSF1_KKSET & "', "   ' SXLOSF1�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''        sql = sql & " " & .SXLOSF1_MEAS1 & ", "   ' SXLOSF1����_�P
'''''        sql = sql & " " & .SXLOSF1_MEAS2 & ", "   ' SXLOSF1����_2
'''''        sql = sql & " " & .SXLOSF1_MEAS3 & ", "   ' SXLOSF1����_3
'''''        sql = sql & " " & .SXLOSF1_MEAS4 & ", "   ' SXLOSF1����_4
'''''        sql = sql & " " & .SXLOSF1_MEAS5 & ", "   ' SXLOSF1����_5
'''''        sql = sql & " " & .SXLOSF1_MEAS6 & ", "   ' SXLOSF1����_6
'''''        sql = sql & " " & .SXLOSF1_MEAS7 & ", "   ' SXLOSF1����_7
'''''        sql = sql & " " & .SXLOSF1_MEAS8 & ", "   ' SXLOSF1����_8
'''''        sql = sql & " " & .SXLOSF1_MEAS9 & ", "   ' SXLOSF1����_9
'''''        sql = sql & " " & .SXLOSF1_MEAS10 & ", " ' SXLOSF1����_10
'''''        sql = sql & " " & .SXLOSF1_MEAS11 & ", "  ' SXLOSF1����_11
'''''        sql = sql & " " & .SXLOSF1_MEAS12 & ", "  ' SXLOSF1����_12
'''''        sql = sql & " " & .SXLOSF1_MEAS13 & ", "  ' SXLOSF1����_13
'''''        sql = sql & " " & .SXLOSF1_MEAS14 & ", "  ' SXLOSF1����_14
'''''        sql = sql & " " & .SXLOSF1_MEAS15 & ", "  ' SXLOSF1����_15
'''''        sql = sql & " " & .SXLOSF1_MEAS16 & ", "  ' SXLOSF1����_16
'''''        sql = sql & " " & .SXLOSF1_MEAS17 & ", "  ' SXLOSF1����_17
'''''        sql = sql & " " & .SXLOSF1_MEAS18 & ", "  ' SXLOSF1����_18
'''''        sql = sql & " " & .SXLOSF1_MEAS19 & ", "  ' SXLOSF1����_19
'''''        sql = sql & " " & .SXLOSF1_MEAS20 & ", "  ' SXLOSF1����_20
'''''        sql = sql & " " & .SXLOSF1_CALCMAX & ", " ' OSF1SXL�v�Z���� Max_1
'''''        sql = sql & " " & .SXLOSF1_CALCAVE & ", "  ' OSF1SXL�v�Z���� Ave_1
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " '" & .SXLOSF2_KKSP & "', "    ' SXLOSF�Q�������ב���ʒu
'''''        sql = sql & " '" & .SXLOSF2_NETU & "', "    ' SXLOSF�Q�M�����@
'''''        sql = sql & " '" & .SXLOSF2_KKSET & "', "   ' SXLOSF�Q�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''        sql = sql & " " & .SXLOSF2_MEAS1 & ", "   ' SXLOSF2����_�P
'''''        sql = sql & " " & .SXLOSF2_MEAS2 & ", "   ' SXLOSF2����_2
'''''        sql = sql & " " & .SXLOSF2_MEAS3 & ", "   ' SXLOSF2����_3
'''''        sql = sql & " " & .SXLOSF2_MEAS4 & ", "   ' SXLOSF2����_4
'''''        sql = sql & " " & .SXLOSF2_MEAS5 & ", "   ' SXLOSF2����_5
'''''        sql = sql & " " & .SXLOSF2_MEAS6 & ", "   ' SXLOSF2����_6
'''''        sql = sql & " " & .SXLOSF2_MEAS7 & ", "   ' SXLOSF2����_7
'''''        sql = sql & " " & .SXLOSF2_MEAS8 & ", "   ' SXLOSF2����_8
'''''        sql = sql & " " & .SXLOSF2_MEAS9 & ", "   ' SXLOSF2����_9
'''''        sql = sql & " " & .SXLOSF2_MEAS10 & ", "  ' SXLOSF2����_10
'''''        sql = sql & " " & .SXLOSF2_MEAS11 & ", "  ' SXLOSF2����_11
'''''        sql = sql & " " & .SXLOSF2_MEAS12 & ", "  ' SXLOSF2����_12
'''''        sql = sql & " " & .SXLOSF2_MEAS13 & ", "  ' SXLOSF2����_13
'''''        sql = sql & " " & .SXLOSF2_MEAS14 & ", "  ' SXLOSF2����_14
'''''        sql = sql & " " & .SXLOSF2_MEAS15 & ", "  ' SXLOSF2����_15
'''''        sql = sql & " " & .SXLOSF2_MEAS16 & ", "  ' SXLOSF2����_16
'''''        sql = sql & " " & .SXLOSF2_MEAS17 & ", "  ' SXLOSF2����_17
'''''        sql = sql & " " & .SXLOSF2_MEAS18 & ", "  ' SXLOSF2����_18
'''''        sql = sql & " " & .SXLOSF2_MEAS19 & ", "  ' SXLOSF2����_19
'''''        sql = sql & " " & .SXLOSF2_MEAS20 & ", "  ' SXLOSF2����_20
'''''        sql = sql & " " & .SXLOSF2_CALCMAX & ", " ' OSF�QSXL�v�Z���� Max_2
'''''        sql = sql & " " & .SXLOSF2_CALCAVE & ", " ' OSF�QSXL�v�Z���� Ave_2
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " '" & .SXLOSF3_KKSP & "', "    ' SXLOSF�R�������ב���ʒu
'''''        sql = sql & " '" & .SXLOSF3_NETU & "', "    ' SXLOSF�R�M�����@
'''''        sql = sql & " '" & .SXLOSF3_KKSET & "', "   ' SXLOSF�R�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''        sql = sql & " " & .SXLOSF3_MEAS1 & ", "   ' SXLOSF3����_�P
'''''        sql = sql & " " & .SXLOSF3_MEAS2 & ", "   ' SXLOSF3����_2
'''''        sql = sql & " " & .SXLOSF3_MEAS3 & ", "   ' SXLOSF3����_3
'''''        sql = sql & " " & .SXLOSF3_MEAS4 & ", "   ' SXLOSF3����_4
'''''        sql = sql & " " & .SXLOSF3_MEAS5 & ", "   ' SXLOSF3����_5
'''''        sql = sql & " " & .SXLOSF3_MEAS6 & ", "   ' SXLOSF3����_6
'''''        sql = sql & " " & .SXLOSF3_MEAS7 & ", "   ' SXLOSF3����_7
'''''        sql = sql & " " & .SXLOSF3_MEAS8 & ", "   ' SXLOSF3����_8
'''''        sql = sql & " " & .SXLOSF3_MEAS9 & ", "   ' SXLOSF3����_9
'''''        sql = sql & " " & .SXLOSF3_MEAS10 & ", "  ' SXLOSF3����_10
'''''        sql = sql & " " & .SXLOSF3_MEAS11 & ", "  ' SXLOSF3����_11
'''''        sql = sql & " " & .SXLOSF3_MEAS12 & ", "  ' SXLOSF3����_12
'''''        sql = sql & " " & .SXLOSF3_MEAS13 & ", "  ' SXLOSF3����_13
'''''        sql = sql & " " & .SXLOSF3_MEAS14 & ", "  ' SXLOSF3����_14
'''''        sql = sql & " " & .SXLOSF3_MEAS15 & ", "  ' SXLOSF3����_15
'''''        sql = sql & " " & .SXLOSF3_MEAS16 & ", "  ' SXLOSF3����_16
'''''        sql = sql & " " & .SXLOSF3_MEAS17 & ", "  ' SXLOSF3����_17
'''''        sql = sql & " " & .SXLOSF3_MEAS18 & ", "  ' SXLOSF3����_18
'''''        sql = sql & " " & .SXLOSF3_MEAS19 & ", "  ' SXLOSF3����_19
'''''        sql = sql & " " & .SXLOSF3_MEAS20 & ", "  ' SXLOSF3����_20
'''''        sql = sql & " " & .SXLOSF3_CALCMAX & ", " ' OSF�RSXL�v�Z���� Max_3
'''''        sql = sql & " " & .SXLOSF3_CALCAVE & ", " ' OSF�RSXL�v�Z���� Ave_3
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " '" & .SXLOSF4_KKSP & "', "    ' SXLOSF�S�������ב���ʒu
'''''        sql = sql & " '" & .SXLOSF4_NETU & "', "    ' SXLOSF�S�M�����@
'''''        sql = sql & " '" & .SXLOSF4_KKSET & "', "   ' SXLOSF�S�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''        sql = sql & " " & .SXLOSF4_MEAS1 & ", "   ' SXLOSF4����_�P
'''''        sql = sql & " " & .SXLOSF4_MEAS2 & ", "   ' SXLOSF4����_2
'''''        sql = sql & " " & .SXLOSF4_MEAS3 & ", "   ' SXLOSF4����_3
'''''        sql = sql & " " & .SXLOSF4_MEAS4 & ", "   ' SXLOSF4����_4
'''''        sql = sql & " " & .SXLOSF4_MEAS5 & ", "   ' SXLOSF4����_5
'''''        sql = sql & " " & .SXLOSF4_MEAS6 & ", "   ' SXLOSF4����_6
'''''        sql = sql & " " & .SXLOSF4_MEAS7 & ", "   ' SXLOSF4����_7
'''''        sql = sql & " " & .SXLOSF4_MEAS8 & ", "   ' SXLOSF4����_8
'''''        sql = sql & " " & .SXLOSF4_MEAS9 & ", "   ' SXLOSF4����_9
'''''        sql = sql & " " & .SXLOSF4_MEAS10 & ", "  ' SXLOSF4����_10
'''''        sql = sql & " " & .SXLOSF4_MEAS11 & ", " ' SXLOSF4����_11
'''''        sql = sql & " " & .SXLOSF4_MEAS12 & ", "  ' SXLOSF4����_12
'''''        sql = sql & " " & .SXLOSF4_MEAS13 & ", "  ' SXLOSF4����_13
'''''        sql = sql & " " & .SXLOSF4_MEAS14 & ", "  ' SXLOSF4����_14
'''''        sql = sql & " " & .SXLOSF4_MEAS15 & ", "  ' SXLOSF4����_15
'''''        sql = sql & " " & .SXLOSF4_MEAS16 & ", "  ' SXLOSF4����_16
'''''        sql = sql & " " & .SXLOSF4_MEAS17 & ", "  ' SXLOSF4����_17
'''''        sql = sql & " " & .SXLOSF4_MEAS18 & ", "  ' SXLOSF4����_18
'''''        sql = sql & " " & .SXLOSF4_MEAS19 & ", "  ' SXLOSF4����_19
'''''        sql = sql & " " & .SXLOSF4_MEAS20 & ", "  ' SXLOSF4����_20
'''''        sql = sql & " " & .SXLOSF4_CALCMAX & ", " ' OSF�SSXL�v�Z���� Max_4
'''''        sql = sql & " " & .SXLOSF4_CALCAVE & ", " ' OSF�SSXL�v�Z���� Ave_4
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLBMD_SMPPOS & ", "   ' SXLBMD����ّ���ʒu�iSXL�ʒu���j
'''''        sql = sql & " '" & .SXLBMD1_KKSP & "', "    ' SXLBMD1�������ב���ʒu
'''''        sql = sql & " '" & .SXLBMD1_NETU & "', "    ' SXLBMD1�M�����@
'''''        sql = sql & " '" & .SXLBMD1_KKSET & "', "   ' SXLBMD1�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''        sql = sql & " " & .SXLBMD1_MEAS1 & ", "   ' SXLBMD1����_�P
'''''        sql = sql & " " & .SXLBMD1_MEAS2 & ", "   ' SXLBMD1����_2
'''''        sql = sql & " " & .SXLBMD1_MEAS3 & ", "   ' SXLBMD1����_3
'''''        sql = sql & " " & .SXLBMD1_MEAS4 & ", "   ' SXLBMD1����_4
'''''        sql = sql & " " & .SXLBMD1_MEAS5 & ", "   ' SXLBMD1����_5
'''''        sql = sql & " " & .SXLBMD1_CALCMAX & ", " ' BMD1SXL�v�Z���� Max
'''''        sql = sql & " " & .SXLBMD1_CALCAVE & ", " ' BMD1SXL�v�Z���� Ave
'''''        sql = sql & " '" & .SXLBMD2_KKSP & "', "    ' SXLBMD�Q�������ב���ʒu
'''''        sql = sql & " '" & .SXLBMD2_NETU & "', "    ' SXLBMD�Q�M�����@
'''''        sql = sql & " '" & .SXLBMD2_KKSET & "', "   ' SXLBMD�Q�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''        sql = sql & " " & .SXLBMD2_MEAS1 & ", "   ' SXLBMD2����_�P
'''''        sql = sql & " " & .SXLBMD2_MEAS2 & ", "   ' SXLBMD2����_2
'''''        sql = sql & " " & .SXLBMD2_MEAS3 & ", "   ' SXLBMD2����_3
'''''        sql = sql & " " & .SXLBMD2_MEAS4 & ", "   ' SXLBMD2����_4
'''''        sql = sql & " " & .SXLBMD2_MEAS5 & ", "   ' SXLBMD2����_5
'''''        sql = sql & " " & .SXLBMD2_CALCMAX & ", " ' BMD�QSXL�v�Z���� Max
'''''        sql = sql & " " & .SXLBMD2_CALCAVE & ", " ' BMD�QSXL�v�Z���� Ave
'''''        sql = sql & " '" & .SXLBMD3_KKSP & "', "    ' SXLBMD�R�������ב���ʒu
'''''        sql = sql & " '" & .SXLBMD3_NETU & "', "    ' SXLBMD�R�M�����@
'''''        sql = sql & " '" & .SXLBMD3_KKSET & "', "  ' SXLBMD�R�������ב�������{�I��ET��@�@char(1)�{number(2)
'''''        sql = sql & " " & .SXLBMD3_MEAS1 & ", "   ' SXLBMD3����_�P
'''''        sql = sql & " " & .SXLBMD3_MEAS2 & ", "   ' SXLBMD3����_2
'''''        sql = sql & " " & .SXLBMD3_MEAS3 & ", "   ' SXLBMD3����_3
'''''        sql = sql & " " & .SXLBMD3_MEAS4 & ", "   ' SXLBMD3����_4
'''''        sql = sql & " " & .SXLBMD3_MEAS5 & ", "   ' SXLBMD3����_5
'''''        sql = sql & " " & .SXLBMD3_CALCMAX & ", " ' BMD�RSXL�v�Z���� Max
'''''        sql = sql & " " & .SXLBMD3_CALCAVE & ", " ' BMD�RSXL�v�Z���� Ave
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_SMPPOS & ", "    ' SXLGD����ّ���ʒu�iSXL�ʒu���j
'''''        sql = sql & " " & .SXLGD_MS01LDL1 & ", "  ' SXLGD_����l01 L/DL1
'''''        sql = sql & " " & .SXLGD_MS01LDL2 & ", "  ' SXLGD_����l01 L/DL2
'''''        sql = sql & " " & .SXLGD_MS01LDL3 & ", "  ' SXLGD_����l01 L/DL3
'''''        sql = sql & " " & .SXLGD_MS01LDL4 & ", "  ' SXLGD_����l01 L/DL4
'''''        sql = sql & " " & .SXLGD_MS01LDL5 & ", "  ' SXLGD_����l01 L/DL5
'''''        sql = sql & " " & .SXLGD_MS01DEN1 & ", "  ' SXLGD_����l01 Den1
'''''        sql = sql & " " & .SXLGD_MS01DEN2 & ", "  ' SXLGD_����l01 Den2
'''''        sql = sql & " " & .SXLGD_MS01DEN3 & ", "  ' SXLGD_����l01 Den3
'''''        sql = sql & " " & .SXLGD_MS01DEN4 & ", "  ' SXLGD_����l01 Den4
'''''        sql = sql & " " & .SXLGD_MS01DEN5 & ", "  ' SXLGD_����l01 Den5
'''''        sql = sql & " " & .SXLGD_MS02LDL1 & ", "  ' SXLGD_����l02 L/DL1
'''''        sql = sql & " " & .SXLGD_MS02LDL2 & ", "  ' SXLGD_����l02 L/DL2
'''''        sql = sql & " " & .SXLGD_MS02LDL3 & ", "  ' SXLGD_����l02 L/DL3
'''''        sql = sql & " " & .SXLGD_MS02LDL4 & ", "  ' SXLGD_����l02 L/DL4
'''''        sql = sql & " " & .SXLGD_MS02LDL5 & ", "  ' SXLGD_����l02 L/DL5
'''''        sql = sql & " " & .SXLGD_MS02DEN1 & ", "  ' SXLGD_����l02 Den1
'''''        sql = sql & " " & .SXLGD_MS02DEN2 & ", "  ' SXLGD_����l02 Den2
'''''        sql = sql & " " & .SXLGD_MS02DEN3 & ", "  ' SXLGD_����l02 Den3
'''''        sql = sql & " " & .SXLGD_MS02DEN4 & ", "  ' SXLGD_����l02 Den4
'''''        sql = sql & " " & .SXLGD_MS02DEN5 & ", "  ' SXLGD_����l02 Den5
'''''        sql = sql & " " & .SXLGD_MS03LDL1 & ", "  ' SXLGD_����l03 L/DL1
'''''        sql = sql & " " & .SXLGD_MS03LDL2 & ", "  ' SXLGD_����l03 L/DL2
'''''        sql = sql & " " & .SXLGD_MS03LDL3 & ", "  ' SXLGD_����l03 L/DL3
'''''        sql = sql & " " & .SXLGD_MS03LDL4 & ", "  ' SXLGD_����l03 L/DL4
'''''        sql = sql & " " & .SXLGD_MS03LDL5 & ", "  ' SXLGD_����l03 L/DL5
'''''        sql = sql & " " & .SXLGD_MS03DEN1 & ", "  ' SXLGD_����l03 Den1
'''''        sql = sql & " " & .SXLGD_MS03DEN2 & ", "  ' SXLGD_����l03 Den2
'''''        sql = sql & " " & .SXLGD_MS03DEN3 & ", "  ' SXLGD_����l03 Den3
'''''        sql = sql & " " & .SXLGD_MS03DEN4 & ", "  ' SXLGD_����l03 Den4
'''''        sql = sql & " " & .SXLGD_MS03DEN5 & ", "  ' SXLGD_����l03 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_MS04LDL1 & ", "  ' SXLGD_����l04 L/DL1
'''''        sql = sql & " " & .SXLGD_MS04LDL2 & ", "  ' SXLGD_����l04 L/DL2
'''''        sql = sql & " " & .SXLGD_MS04LDL3 & ", "  ' SXLGD_����l04 L/DL3
'''''        sql = sql & " " & .SXLGD_MS04LDL4 & ", "  ' SXLGD_����l04 L/DL4
'''''        sql = sql & " " & .SXLGD_MS04LDL5 & ", "  ' SXLGD_����l04 L/DL5
'''''        sql = sql & " " & .SXLGD_MS04DEN1 & ", "  ' SXLGD_����l04 Den1
'''''        sql = sql & " " & .SXLGD_MS04DEN2 & ", "  ' SXLGD_����l04 Den2
'''''        sql = sql & " " & .SXLGD_MS04DEN3 & ", "  ' SXLGD_����l04 Den3
'''''        sql = sql & " " & .SXLGD_MS04DEN4 & ", "  ' SXLGD_����l04 Den4
'''''        sql = sql & " " & .SXLGD_MS04DEN5 & ", "  ' SXLGD_����l04 Den5
'''''        sql = sql & " " & .SXLGD_MS05LDL1 & ", "  ' SXLGD_����l05 L/DL1
'''''        sql = sql & " " & .SXLGD_MS05LDL2 & ", "  ' SXLGD_����l05 L/DL2
'''''        sql = sql & " " & .SXLGD_MS05LDL3 & ", "  ' SXLGD_����l05 L/DL3
'''''        sql = sql & " " & .SXLGD_MS05LDL4 & ", "  ' SXLGD_����l05 L/DL4
'''''        sql = sql & " " & .SXLGD_MS05LDL5 & ", "  ' SXLGD_����l05 L/DL5
'''''        sql = sql & " " & .SXLGD_MS05DEN1 & ", "  ' SXLGD_����l05 Den1
'''''        sql = sql & " " & .SXLGD_MS05DEN2 & ", "  ' SXLGD_����l05 Den2
'''''        sql = sql & " " & .SXLGD_MS05DEN3 & ", "  ' SXLGD_����l05 Den3
'''''        sql = sql & " " & .SXLGD_MS05DEN4 & ", "  ' SXLGD_����l05 Den4
'''''        sql = sql & " " & .SXLGD_MS05DEN5 & ", "  ' SXLGD_����l05 Den5
'''''        sql = sql & " " & .SXLGD_MS06LDL1 & ", "  ' SXLGD_����l06 L/DL1
'''''        sql = sql & " " & .SXLGD_MS06LDL2 & ", "  ' SXLGD_����l06 L/DL2
'''''        sql = sql & " " & .SXLGD_MS06LDL3 & ", "  ' SXLGD_����l06 L/DL3
'''''        sql = sql & " " & .SXLGD_MS06LDL4 & ", "  ' SXLGD_����l06 L/DL4
'''''        sql = sql & " " & .SXLGD_MS06LDL5 & ", "  ' SXLGD_����l06 L/DL5
'''''        sql = sql & " " & .SXLGD_MS06DEN1 & ", "  ' SXLGD_����l06 Den1
'''''        sql = sql & " " & .SXLGD_MS06DEN2 & ", "  ' SXLGD_����l06 Den2
'''''        sql = sql & " " & .SXLGD_MS06DEN3 & ", "  ' SXLGD_����l06 Den3
'''''        sql = sql & " " & .SXLGD_MS06DEN4 & ", "  ' SXLGD_����l06 Den4
'''''        sql = sql & " " & .SXLGD_MS06DEN5 & ", "  ' SXLGD_����l06 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_MS07LDL1 & ", "  ' SXLGD_����l07 L/DL1
'''''        sql = sql & " " & .SXLGD_MS07LDL2 & ", "  ' SXLGD_����l07 L/DL2
'''''        sql = sql & " " & .SXLGD_MS07LDL3 & ", "  ' SXLGD_����l07 L/DL3
'''''        sql = sql & " " & .SXLGD_MS07LDL4 & ", "  ' SXLGD_����l07 L/DL4
'''''        sql = sql & " " & .SXLGD_MS07LDL5 & ", "  ' SXLGD_����l07 L/DL5
'''''        sql = sql & " " & .SXLGD_MS07DEN1 & ", "  ' SXLGD_����l07 Den1
'''''        sql = sql & " " & .SXLGD_MS07DEN2 & ", "  ' SXLGD_����l07 Den2
'''''        sql = sql & " " & .SXLGD_MS07DEN3 & ", "  ' SXLGD_����l07 Den3
'''''        sql = sql & " " & .SXLGD_MS07DEN4 & ", "  ' SXLGD_����l07 Den4
'''''        sql = sql & " " & .SXLGD_MS07DEN5 & ", "  ' SXLGD_����l07 Den5
'''''        sql = sql & " " & .SXLGD_MS08LDL1 & ", "  ' SXLGD_����l08 L/DL1
'''''        sql = sql & " " & .SXLGD_MS08LDL2 & ", "  ' SXLGD_����l08 L/DL2
'''''        sql = sql & " " & .SXLGD_MS08LDL3 & ", "  ' SXLGD_����l08 L/DL3
'''''        sql = sql & " " & .SXLGD_MS08LDL4 & ", "  ' SXLGD_����l08 L/DL4
'''''        sql = sql & " " & .SXLGD_MS08LDL5 & ", "  ' SXLGD_����l08 L/DL5
'''''        sql = sql & " " & .SXLGD_MS08DEN1 & ", "  ' SXLGD_����l08 Den1
'''''        sql = sql & " " & .SXLGD_MS08DEN2 & ", "  ' SXLGD_����l08 Den2
'''''        sql = sql & " " & .SXLGD_MS08DEN3 & ", "  ' SXLGD_����l08 Den3
'''''        sql = sql & " " & .SXLGD_MS08DEN4 & ", "  ' SXLGD_����l08 Den4
'''''        sql = sql & " " & .SXLGD_MS08DEN5 & ", "  ' SXLGD_����l08 Den5
'''''        sql = sql & " " & .SXLGD_MS09LDL1 & ", "  ' SXLGD_����l09 L/DL1
'''''        sql = sql & " " & .SXLGD_MS09LDL2 & ", "  ' SXLGD_����l09 L/DL2
'''''        sql = sql & " " & .SXLGD_MS09LDL3 & ", "  ' SXLGD_����l09 L/DL3
'''''        sql = sql & " " & .SXLGD_MS09LDL4 & ", "  ' SXLGD_����l09 L/DL4
'''''        sql = sql & " " & .SXLGD_MS09LDL5 & ", "  ' SXLGD_����l09 L/DL5
'''''        sql = sql & " " & .SXLGD_MS09DEN1 & ", "  ' SXLGD_����l09 Den1
'''''        sql = sql & " " & .SXLGD_MS09DEN2 & ", " ' SXLGD_����l09 Den2
'''''        sql = sql & " " & .SXLGD_MS09DEN3 & ", "  ' SXLGD_����l09 Den3
'''''        sql = sql & " " & .SXLGD_MS09DEN4 & ", "  ' SXLGD_����l09 Den4
'''''        sql = sql & " " & .SXLGD_MS09DEN5 & ", "  ' SXLGD_����l09 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_MS10LDL1 & ", "  ' SXLGD_����l10 L/DL1
'''''        sql = sql & " " & .SXLGD_MS10LDL2 & ", "  ' SXLGD_����l10 L/DL2
'''''        sql = sql & " " & .SXLGD_MS10LDL3 & ", "  ' SXLGD_����l10 L/DL3
'''''        sql = sql & " " & .SXLGD_MS10LDL4 & ", "  ' SXLGD_����l10 L/DL4
'''''        sql = sql & " " & .SXLGD_MS10LDL5 & ", "  ' SXLGD_����l10 L/DL5
'''''        sql = sql & " " & .SXLGD_MS10DEN1 & ", "  ' SXLGD_����l10 Den1
'''''        sql = sql & " " & .SXLGD_MS10DEN2 & ", "  ' SXLGD_����l10 Den2
'''''        sql = sql & " " & .SXLGD_MS10DEN3 & ", "  ' SXLGD_����l10 Den3
'''''        sql = sql & " " & .SXLGD_MS10DEN4 & ", "  ' SXLGD_����l10 Den4
'''''        sql = sql & " " & .SXLGD_MS10DEN5 & ", "  ' SXLGD_����l10 Den5
'''''        sql = sql & " " & .SXLGD_MS11LDL1 & ", "  ' SXLGD_����l11 L/DL1
'''''        sql = sql & " " & .SXLGD_MS11LDL2 & ", "  ' SXLGD_����l11 L/DL2
'''''        sql = sql & " " & .SXLGD_MS11LDL3 & ", "  ' SXLGD_����l11 L/DL3
'''''        sql = sql & " " & .SXLGD_MS11LDL4 & ", " ' SXLGD_����l11 L/DL4
'''''        sql = sql & " " & .SXLGD_MS11LDL5 & ", "  ' SXLGD_����l11 L/DL5
'''''        sql = sql & " " & .SXLGD_MS11DEN1 & ", "  ' SXLGD_����l11 Den1
'''''        sql = sql & " " & .SXLGD_MS11DEN2 & ", "  ' SXLGD_����l11 Den2
'''''        sql = sql & " " & .SXLGD_MS11DEN3 & ", "  ' SXLGD_����l11 Den3
'''''        sql = sql & " " & .SXLGD_MS11DEN4 & ", "  ' SXLGD_����l11 Den4
'''''        sql = sql & " " & .SXLGD_MS11DEN5 & ", "  ' SXLGD_����l11 Den5
'''''        sql = sql & " " & .SXLGD_MS12LDL1 & ", "  ' SXLGD_����l12 L/DL1
'''''        sql = sql & " " & .SXLGD_MS12LDL2 & ", "  ' SXLGD_����l12 L/DL2
'''''        sql = sql & " " & .SXLGD_MS12LDL3 & ", "  ' SXLGD_����l12 L/DL3
'''''        sql = sql & " " & .SXLGD_MS12LDL4 & ", "  ' SXLGD_����l12 L/DL4
'''''        sql = sql & " " & .SXLGD_MS12LDL5 & ", "  ' SXLGD_����l12 L/DL5
'''''        sql = sql & " " & .SXLGD_MS12DEN1 & ", "  ' SXLGD_����l12 Den1
'''''        sql = sql & " " & .SXLGD_MS12DEN2 & ", "  ' SXLGD_����l12 Den2
'''''        sql = sql & " " & .SXLGD_MS12DEN3 & ", "  ' SXLGD_����l12 Den3
'''''        sql = sql & " " & .SXLGD_MS12DEN4 & ", "  ' SXLGD_����l12 Den4
'''''        sql = sql & " " & .SXLGD_MS12DEN5 & ", "  ' SXLGD_����l12 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_MS13LDL1 & ", "  ' SXLGD_����l13 L/DL1
'''''        sql = sql & " " & .SXLGD_MS13LDL2 & ", "  ' SXLGD_����l13 L/DL2
'''''        sql = sql & " " & .SXLGD_MS13LDL3 & ", "  ' SXLGD_����l13 L/DL3
'''''        sql = sql & " " & .SXLGD_MS13LDL4 & ", "  ' SXLGD_����l13 L/DL4
'''''        sql = sql & " " & .SXLGD_MS13LDL5 & ", "  ' SXLGD_����l13 L/DL5
'''''        sql = sql & " " & .SXLGD_MS13DEN1 & ", "  ' SXLGD_����l13 Den1
'''''        sql = sql & " " & .SXLGD_MS13DEN2 & ", "  ' SXLGD_����l13 Den2
'''''        sql = sql & " " & .SXLGD_MS13DEN3 & ", "  ' SXLGD_����l13 Den3
'''''        sql = sql & " " & .SXLGD_MS13DEN4 & ", "  ' SXLGD_����l13 Den4
'''''        sql = sql & " " & .SXLGD_MS13DEN5 & ", "  ' SXLGD_����l13 Den5
'''''        sql = sql & " " & .SXLGD_MS14LDL1 & ", "  ' SXLGD_����l14 L/DL1
'''''        sql = sql & " " & .SXLGD_MS14LDL2 & ", "  ' SXLGD_����l14 L/DL2
'''''        sql = sql & " " & .SXLGD_MS14LDL3 & ", "  ' SXLGD_����l14 L/DL3
'''''        sql = sql & " " & .SXLGD_MS14LDL4 & ", "  ' SXLGD_����l14 L/DL4
'''''        sql = sql & " " & .SXLGD_MS14LDL5 & ", "  ' SXLGD_����l14 L/DL5
'''''        sql = sql & " " & .SXLGD_MS14DEN1 & ", "  ' SXLGD_����l14 Den1
'''''        sql = sql & " " & .SXLGD_MS14DEN2 & ", "  ' SXLGD_����l14 Den2
'''''        sql = sql & " " & .SXLGD_MS14DEN3 & ", "  ' SXLGD_����l14 Den3
'''''        sql = sql & " " & .SXLGD_MS14DEN4 & ", "  ' SXLGD_����l14 Den4
'''''        sql = sql & " " & .SXLGD_MS14DEN5 & ", "  ' SXLGD_����l14 Den5
'''''        sql = sql & " " & .SXLGD_MS15LDL1 & ", "  ' SXLGD_����l15 L/DL1
'''''        sql = sql & " " & .SXLGD_MS15LDL2 & ", "  ' SXLGD_����l15 L/DL2
'''''        sql = sql & " " & .SXLGD_MS15LDL3 & ", "  ' SXLGD_����l15 L/DL3
'''''        sql = sql & " " & .SXLGD_MS15LDL4 & ", "  ' SXLGD_����l15 L/DL4
'''''        sql = sql & " " & .SXLGD_MS15LDL5 & ", "  ' SXLGD_����l15 L/DL5
'''''        sql = sql & " " & .SXLGD_MS15DEN1 & ", "  ' SXLGD_����l15 Den1
'''''        sql = sql & " " & .SXLGD_MS15DEN2 & ", "  ' SXLGD_����l15 Den2
'''''        sql = sql & " " & .SXLGD_MS15DEN3 & ", "  ' SXLGD_����l15 Den3
'''''        sql = sql & " " & .SXLGD_MS15DEN4 & ", "  ' SXLGD_����l15 Den4
'''''        sql = sql & " " & .SXLGD_MS15DEN5 & ", "  ' SXLGD_����l15 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_MSRSDEN & ", "   ' SXLGD_���茋�� Den
'''''        sql = sql & " " & .SXLGD_MSRSLDL & ", "   ' SXLGD_���茋�� L/DL
'''''        sql = sql & " " & .SXLGD_MSRSDVD2 & ", "  ' SXLGD_���茋�� DVD2
'''''        sql = sql & " " & .SXLT_SMPPOS & ", "     ' SXLLT����ّ���ʒu�iSXL�ʒu���j
'''''        sql = sql & " " & .SXLLT_MEASPEAK & ", "  ' SXLLT_����l �s�[�N�l
'''''        sql = sql & " " & .SXLLT_MEAS1 & ", "     ' SXLLT_����l1
'''''        sql = sql & " " & .SXLLT_MEAS2 & ", "     ' SXLLT_����l2
'''''        sql = sql & " " & .SXLLT_MEAS3 & ", "     ' SXLLT_����l3
'''''        sql = sql & " " & .SXLLT_MEAS4 & ", "     ' SXLLT_����l4
'''''        sql = sql & " " & .SXLLT_MEAS5 & ", "     ' SXLLT_����l5
'''''        sql = sql & " " & .SXLLT_CALCMEAS & ", "  ' SXLLT_�v�Z����
'''''        sql = sql & "sysdate, "
'''''        sql = sql & "'0', "
'''''        sql = sql & "sysdate, "
'''''        sql = sql & " " & .SXLOSF1_POS1 & ", "     'OSF1����݋敪�P�ʒu
'''''        sql = sql & " " & .SXLOSF1_WID1 & ", "     'OSF1����݋敪�P��
'''''        sql = sql & " '" & .SXLOSF1_RD1 & "', "      'OSF1����݋敪�PR/D
'''''        sql = sql & " " & .SXLOSF1_POS2 & ", "     'OSF1����݋敪�Q�ʒu
'''''        sql = sql & " " & .SXLOSF1_WID2 & ", "     'OSF1����݋敪�Q��
'''''        sql = sql & " '" & .SXLOSF1_RD2 & "', "      'OSF1����݋敪�QR/D
'''''        sql = sql & " " & .SXLOSF1_POS3 & ", "     'OSF1����݋敪�R�ʒu
'''''        sql = sql & " " & .SXLOSF1_WID3 & ", "     'OSF1����݋敪�R��
'''''        sql = sql & " '" & .SXLOSF1_RD3 & "', "      'OSF1����݋敪�RR/D
'''''        sql = sql & " " & .SXLOSF2_POS1 & ", "     'OSF2����݋敪�P�ʒu
'''''        sql = sql & " " & .SXLOSF2_WID1 & ", "     'OSF2����݋敪�P��
'''''        sql = sql & " '" & .SXLOSF2_RD1 & "', "      'OSF2����݋敪�PR/D
'''''        sql = sql & " " & .SXLOSF2_POS2 & ", "     'OSF2����݋敪�Q�ʒu
'''''        sql = sql & " " & .SXLOSF2_WID2 & ", "     'OSF2����݋敪�Q��
'''''        sql = sql & " '" & .SXLOSF2_RD2 & "', "      'OSF2����݋敪�QR/D
'''''        sql = sql & " " & .SXLOSF2_POS3 & ", "     'OSF2����݋敪�R�ʒu
'''''        sql = sql & " " & .SXLOSF2_WID3 & ", "     'OSF2����݋敪�R��
'''''        sql = sql & " '" & .SXLOSF2_RD3 & "', "      'OSF2����݋敪�RR/D
'''''        sql = sql & " " & .SXLOSF3_POS1 & ", "     'OSF3����݋敪�P�ʒu
'''''        sql = sql & " " & .SXLOSF3_WID1 & ", "     'OSF3����݋敪�P��
'''''        sql = sql & " '" & .SXLOSF3_RD1 & "', "      'OSF3����݋敪�PR/D
'''''        sql = sql & " " & .SXLOSF3_POS2 & ", "     'OSF3����݋敪�Q�ʒu
'''''        sql = sql & " " & .SXLOSF3_WID2 & ", "     'OSF3����݋敪�Q��
'''''        sql = sql & " '" & .SXLOSF3_RD2 & "', "      'OSF3����݋敪�QR/D
'''''        sql = sql & " " & .SXLOSF3_POS3 & ", "     'OSF3����݋敪�R�ʒu
'''''        sql = sql & " " & .SXLOSF3_WID3 & ", "     'OSF3����݋敪�R��
'''''        sql = sql & " '" & .SXLOSF3_RD3 & "', "      'OSF3����݋敪�RR/D
'''''        sql = sql & " " & .SXLOSF4_POS1 & ", "     'OSF4����݋敪�P�ʒu
'''''        sql = sql & " " & .SXLOSF4_WID1 & ", "     'OSF4����݋敪�P��
'''''        sql = sql & " '" & .SXLOSF4_RD1 & "', "      'OSF4����݋敪�PR/D
'''''        sql = sql & " " & .SXLOSF4_POS2 & ", "     'OSF4����݋敪�Q�ʒu
'''''        sql = sql & " " & .SXLOSF4_WID2 & ", "     'OSF4����݋敪�Q��
'''''        sql = sql & " '" & .SXLOSF4_RD2 & "', "      'OSF4����݋敪�QR/D
'''''        sql = sql & " " & .SXLOSF4_POS3 & ", "     'OSF4����݋敪�R�ʒu
'''''        sql = sql & " " & .SXLOSF4_WID3 & ", "     'OSF4����݋敪�R��
'''''        sql = sql & " '" & .SXLOSF4_RD3 & "', "      'OSF4����݋敪�RR/D
'''''        sql = sql & " " & .SXLGD_MS01DVD2 & ", "   'DVD2���茋�ʒl�P
'''''        sql = sql & " " & .SXLGD_MS02DVD2 & ", "   'DVD2���茋�ʒl�Q
'''''        sql = sql & " " & .SXLGD_MS03DVD2 & ", "   'DVD2���茋�ʒl�R
'''''        sql = sql & " " & .SXLGD_MS04DVD2 & ", "   'DVD2���茋�ʒl�S
'''''        sql = sql & " " & .SXLGD_MS05DVD2 & ", "   'DVD2���茋�ʒl�T
'''''        sql = sql & " " & .SXLBMD1_MNBCR & ", "    'BMD1SXL�v�Z���ʖʓ����z
'''''        sql = sql & " " & .SXLBMD2_MNBCR & ", "    'BMD2SXL�v�Z���ʖʓ����z
'''''        sql = sql & " " & .SXLBMD3_MNBCR & ") "    'BMD3SXL�v�Z���ʖʓ����z
'''''    End With
'''''#If PRNSQL > 0 Then
'''''Debug.Print sql
'''''Stop
'''''#End If
'''''
'''''    If 0 >= OraDB.ExecuteSQL(sql) Then
'''''        DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


'�T�v      :�������� ����������эX�V�p�h���C�o
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :Zis           ,I  ,typ_TBCMJ009 ,����������уe�[�u���ւ̑}���p
'          :�߂�l        ,O  ,FUNCTION_RETURN,�ǂݍ��ݐ���
'����      :
'����      :2001/06/27 ���{ �쐬
Public Function DBDRV_scmzc_fcmkc001c_InsZis(Zis As typ_TBCMJ009) As FUNCTION_RETURN

    Dim sql As String
                                          
    '����������тւ̑}���iTBCMJ009�j

    ' �����������

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_InsZis"
    
    sql = "insert into TBCMJ009 ( "
    sql = sql & "CRYNUM, "           ' �����ԍ�
    sql = sql & "INGOTPOS, "         ' �C���S�b�g���ʒu
    sql = sql & "TRANCNT, "          ' ������
    sql = sql & "LENGTH, "           ' ����
    sql = sql & "KRPROCCD, "         ' �Ǘ��H���R�[�h
    sql = sql & "PROCCODE, "         ' �H���R�[�h
    sql = sql & "CODE, "             ' �敪�R�[�h
    sql = sql & "TSTAFFID, "         ' �o�^�Ј�ID
    sql = sql & "REGDATE, "          ' �o�^���t
    sql = sql & "KSTAFFID, "         ' �X�V�Ј�ID
    sql = sql & "UPDDATE, "          ' �X�V���t
    sql = sql & "SENDFLAG, "         ' ���M�t���O
    sql = sql & "SENDDATE )"         ' ���M���t
    
    sql = sql & " select "
    With Zis
    '�H���R�[�h�ݒ胍�W�b�N�̓���@2002/11/28 hama
    '�H���R�[�h�ݒ�
        nextCd = .PROCCODE
        sql = sql & " '" & .CRYNUM & "', "           ' �����ԍ�
        sql = sql & " '" & .IngotPos & "', "         ' �C���S�b�g���ʒu
        sql = sql & " nvl(max(TRANCNT),0)+1, "       ' ������
        sql = sql & " " & .LENGTH & ", "             ' ����
        sql = sql & " '" & .KRPROCCD & "', "         ' �Ǘ��H���R�[�h
        sql = sql & " '" & nextCd & "', "            ' �H���R�[�h
        sql = sql & " '" & .CODE & "', "             ' �敪�R�[�h
        sql = sql & " '" & .TSTAFFID & "', "         ' �o�^�Ј�ID
        sql = sql & "sysdate, "
        sql = sql & " '" & .TSTAFFID & "', "         ' �X�V�Ј�ID
        sql = sql & "sysdate, "
        sql = sql & "'0', "
        sql = sql & "sysdate "
        sql = sql & " from TBCMJ009 "
        sql = sql & " where CRYNUM  ='" & Zis.CRYNUM & "'"
        sql = sql & "   and INGOTPOS= " & Zis.IngotPos
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmkc001c_InsZis = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmkc001c_InsZis = FUNCTION_RETURN_SUCCESS
    End If
    

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmkc001c_InsZis = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�������� �u���b�N�Ǘ��X�V�p�i���ݍH���A�ŏI�ʉߍH���j
'���Ұ�    :�ϐ���        ,IO ,�^                                   ,����
'          :Block         ,I  ,type_DBDRV_scmzc_fcmkc001c_UpdBlock1 ,�u���b�N�Ǘ��̌��݊Ǘ��H���A���ݍH���A�ŏI�ʉߊǗ��H���A�ŏI�ʉߍH���X�V�p
'          :�߂�l        ,O  ,FUNCTION_RETURN                      ,�ǂݍ��ݐ���
'����      :
'����      :2001/06/27 ���{ �쐬
Public Function DBDRV_scmzc_fcmkc001c_UpdBlock1(Block As type_DBDRV_scmzc_fcmkc001c_UpdBlock1) As FUNCTION_RETURN

    Dim sql As String

    ' �u���b�N�Ǘ��̍X�V

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_UpdBlock1"
    '�H���R�[�h�ݒ胍�W�b�N�̓���@2002/11/28 hama
    '�H���R�[�h�ݒ�
    nextCd = Block.NOWPROC
    nowCd = Block.LASTPASS
    
    sql = "update TBCME040 set "
    sql = sql & "  NOWPROC ='" & nextCd & "' "      '���ݍH��
    sql = sql & ", LASTPASS='" & nowCd & "' "       '�ŏI�ʉߍH��
    sql = sql & ", UPDDATE =sysdate "
    sql = sql & ", SENDFLAG='0' "
    
    sql = sql & " where CRYNUM  ='" & Block.CRYNUM & "'"
    sql = sql & "   and INGOTPOS= " & Block.IngotPos
        
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmkc001c_UpdBlock1 = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmkc001c_UpdBlock1 = FUNCTION_RETURN_SUCCESS
    End If


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmkc001c_UpdBlock1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



'�T�v      :�������� �����T���v���Ǘ��X�V�p�i�m��敪���P�ɍX�V�j
'���Ұ�    :�ϐ���        ,IO ,�^                                   ,����
'          :CrySmp()      ,I  ,type_DBDRV_scmzc_fcmkc001c_UpdCrySmp ,�����T���v���Ǘ��X�V�p
'          :�߂�l        ,O  ,FUNCTION_RETURN                      ,�ǂݍ��ݐ���
'����      :
'����      :2001/07/26 ���{ �쐬
Public Function DBDRV_scmzc_fcmkc001c_UpdCrySmp(CrySmp() As type_DBDRV_scmzc_fcmkc001c_UpdCrySmp) As FUNCTION_RETURN

    Dim sql As String
    Dim i As Long
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_UpdCrySmp"

    For i = 1 To UBound(CrySmp)
        ' �����T���v���Ǘ��̍X�V
'        sql = "update TBCME043 set "
'        sql = sql & "  KTKBN='1' "          '�m��敪
'        sql = sql & ", UPDDATE=sysdate "
'        sql = sql & ", SENDFLAG='0' "
'        sql = sql & " where CRYNUM='" & CrySmp(i).CRYNUM & "' "
'        sql = sql & " and INGOTPOS=" & CrySmp(i).INGOTPOS & " "
'        sql = sql & " and SMPKBN='" & CrySmp(i).SMPKBN & "' "

        sql = "update XSDCS set "
        sql = sql & "  KTKBNCS='1' "          '�m��敪
        sql = sql & ", KDAYCS =sysdate "
        sql = sql & ", SNDKCS ='0' "
        
        sql = sql & " where XTALCS  ='" & CrySmp(i).CRYNUM & "' "
        sql = sql & "   and INPOSCS = " & CrySmp(i).IngotPos & " "
        sql = sql & "   and SMPKBNCS='" & CrySmp(i).SMPKBN & "' "
        
        If 0 >= OraDB.ExecuteSQL(sql) Then
            DBDRV_scmzc_fcmkc001c_UpdCrySmp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
     Next
     
     DBDRV_scmzc_fcmkc001c_UpdCrySmp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmkc001c_UpdCrySmp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�������� �u���b�N�Ǘ��X�V�p�i�N���X�^���J�^���O�A�������g�p�j
'���Ұ�    :�ϐ���        ,IO ,�^                           ,����
'          :Block         ,I   ,type_DBDRV_scmzc_fcmkc001c_UpdBlock         ,�u���b�N�Ǘ��̌��݊Ǘ��H���A���ݍH���A�ŏI�ʉߊǗ��H���A�ŏI�ʉߍH���X�V�p
'          :�߂�l        ,O  ,FUNCTION_RETURN              ,
'����      :
'����      :2001/06/27 ���{ �쐬
Public Function DBDRV_fcmkc001c_UpdBlkCR(Block As typ_DBDRV_fcmkc001c_UpdBlkCR) As FUNCTION_RETURN

    Dim sql As String

    ' �u���b�N�Ǘ��̍X�V

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_fcmkc001c_UpdBlkCR"
    '�H���R�[�h�ݒ胍�W�b�N�̓���@2002/11/28 hama
    '�H���R�[�h�ݒ�
    nextCd = Block.NOWPROC
    nowCd = PROCD_KESSYOU_SOUGOUHANTEI

    sql = "update TBCME040 set "
    sql = sql & " NOWPROC  ='" & nextCd & "' "              '���ݍH��
    sql = sql & ", LASTPASS='" & nowCd & "' "               '�ŏI�ʉߍH��
    sql = sql & ", DELCLS  ='" & Block.DELCLS & "' "        '�폜�敪
    sql = sql & ", LSTATCLS='" & Block.LSTATCLS & "' "      '�ŏI��ԋ敪
    sql = sql & ", RSTATCLS='" & Block.RSTATCLS & "' "      '������ԋ敪
    sql = sql & ", BDCAUS  ='" & Block.BDCAUS & "' "        '�s�Ǘ��R
    sql = sql & ", UPDDATE =sysdate "
    sql = sql & ", SENDFLAG='0' "
    sql = sql & " where CRYNUM  ='" & Block.CRYNUM & "'"
    sql = sql & "   and INGOTPOS= " & Block.IngotPos
        
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_fcmkc001c_UpdBlkCR = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_fcmkc001c_UpdBlkCR = FUNCTION_RETURN_SUCCESS
    End If


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_fcmkc001c_UpdBlkCR = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�㉺�̕i�Ԃɂ��Č�������_���𒲂ׂ�
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :CRYNUM        ,I  ,String       ,�Ώۂ̌����ԍ�
''''''          :INGOTPOS      ,I  ,Integer      ,�Ώۂ̈ʒu
''''''          :smpShared     ,O  ,Boolean      ,�T���v���͋��p(�㉺�i�ԂƂ� Z/G/�� �łȂ��Ƃ��̂�True�ɂȂ肤��)
''''''          :WSpec(2)      ,O  ,Judg_Spec_Cry,�㉺�i�Ԃ̌�������_��(1:��i�� 2:���i��)
''''''����      :�������葪��l(TBCMJ014)�ɏ������ނׂ��������ڂ𒲂ׂ邽�߂ɗ��p����
''''''          :Z/G/��i�Ԃ̏ꍇ�͑S�Č����s�v�Ƃ���
''''''����      :2002/2/20 �쑺 �쐬
''''''�T�v      :
'''''Public Function GetHinbanSpec(CRYNUM As String, INGOTPOS As Integer, smpShared As Boolean, WSpec() As Judg_Spec_Cry) As FUNCTION_RETURN
'''''Dim sql$
'''''Dim rs As OraDynaset
'''''Dim i As Integer
'''''Dim loopTo As Integer
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    GetHinbanSpec = FUNCTION_RETURN_FAILURE
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function GetHinbanSpec"
'''''
'''''    '������
'''''    For i = 1 To 2
'''''        With WSpec(i)
'''''            .Enable = False
'''''            .rs = False
'''''            .Oi = False
'''''            .Cs = False
'''''            .Lt = False
'''''            .EPD = False
'''''            .B1 = False
'''''            .B2 = False
'''''            .B3 = False
'''''            .L1 = False
'''''            .L2 = False
'''''            .L3 = False
'''''            .L4 = False
'''''            .GD = False
'''''        End With
'''''    Next
'''''
'''''    '�㉺�i�Ԃ̌�������_���𒲂ׂ�
'''''    sql = "select HIN.INGOTPOS as HinFrom, HIN.INGOTPOS+HIN.LENGTH as HinTo"
'''''    sql = sql & ", E018.HSXRHWYS, E019.HSXONHWS, E019.HSXCNHWS, E019.HSXLTHWS"
'''''    sql = sql & ", E020.HSXOF1HS, E020.HSXOF2HS, E020.HSXOF3HS, E020.HSXOF4HS"
'''''    sql = sql & ", E020.HSXBM1HS, E020.HSXBM2HS, E020.HSXBM3HS"
'''''    sql = sql & ", E020.HSXDENHS, E020.HSXLDLHS, E020.HSXDVDHS "
'''''    sql = sql & "from TBCME041 HIN, TBCME018 E018, TBCME019 E019, TBCME020 E020 "
'''''    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
'''''    sql = sql & "  and HIN.INGOTPOS <= " & INGOTPOS
'''''    sql = sql & "  and HIN.INGOTPOS+HIN.LENGTH >= " & INGOTPOS
'''''    sql = sql & "  and HIN.HINBAN=E018.HINBAN and HIN.REVNUM=E018.MNOREVNO and HIN.FACTORY=E018.FACTORY and HIN.OPECOND=E018.OPECOND"
'''''    sql = sql & "  and HIN.HINBAN=E019.HINBAN and HIN.REVNUM=E019.MNOREVNO and HIN.FACTORY=E019.FACTORY and HIN.OPECOND=E019.OPECOND"
'''''    sql = sql & "  and HIN.HINBAN=E020.HINBAN and HIN.REVNUM=E020.MNOREVNO and HIN.FACTORY=E020.FACTORY and HIN.OPECOND=E020.OPECOND "
'''''    sql = sql & "order by HIN.INGOTPOS"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
'''''
'''''    Do While Not rs.EOF
'''''        If rs("HinFrom") < INGOTPOS Then
'''''            '��i�Ԃ̎d�l
'''''            With WSpec(1)
'''''                .Enable = True
'''''                If ((rs("HSXRHWYS") = SIJI) Or (rs("HSXRHWYS") = SANKOU)) Then .rs = True
'''''                If ((rs("HSXONHWS") = SIJI) Or (rs("HSXONHWS") = SANKOU)) Then .Oi = True
'''''                If ((rs("HSXOF1HS") = SIJI) Or (rs("HSXOF1HS") = SANKOU)) Then .L1 = True
'''''                If ((rs("HSXOF2HS") = SIJI) Or (rs("HSXOF2HS") = SANKOU)) Then .L2 = True
'''''                If ((rs("HSXOF3HS") = SIJI) Or (rs("HSXOF3HS") = SANKOU)) Then .L3 = True
'''''                If ((rs("HSXOF4HS") = SIJI) Or (rs("HSXOF4HS") = SANKOU)) Then .L4 = True
'''''                If ((rs("HSXBM1HS") = SIJI) Or (rs("HSXBM1HS") = SANKOU)) Then .B1 = True
'''''                If ((rs("HSXBM2HS") = SIJI) Or (rs("HSXBM2HS") = SANKOU)) Then .B2 = True
'''''                If ((rs("HSXBM3HS") = SIJI) Or (rs("HSXBM3HS") = SANKOU)) Then .B3 = True
'''''                If ((rs("HSXDENHS") = SIJI) Or (rs("HSXDENHS") = SANKOU)) Or _
'''''                   ((rs("HSXLDLHS") = SIJI) Or (rs("HSXLDLHS") = SANKOU)) Or _
'''''                   ((rs("HSXDVDHS") = SIJI) Or (rs("HSXDVDHS") = SANKOU)) Then .GD = True
'''''                If ((rs("HSXCNHWS") = SIJI) Or (rs("HSXCNHWS") = SANKOU)) Then .Cs = True
'''''                If ((rs("HSXLTHWS") = SIJI) Or (rs("HSXLTHWS") = SANKOU)) Then .Lt = True
'''''                .EPD = True
'''''            End With
'''''        End If
'''''
'''''        If rs("HinTo") > INGOTPOS Then
'''''            '���i�Ԃ̎d�l
'''''            With WSpec(2)
'''''                .Enable = True
'''''                If ((rs("HSXRHWYS") = SIJI) Or (rs("HSXRHWYS") = SANKOU)) Then .rs = True
'''''                If ((rs("HSXONHWS") = SIJI) Or (rs("HSXONHWS") = SANKOU)) Then .Oi = True
'''''                If ((rs("HSXOF1HS") = SIJI) Or (rs("HSXOF1HS") = SANKOU)) Then .L1 = True
'''''                If ((rs("HSXOF2HS") = SIJI) Or (rs("HSXOF2HS") = SANKOU)) Then .L2 = True
'''''                If ((rs("HSXOF3HS") = SIJI) Or (rs("HSXOF3HS") = SANKOU)) Then .L3 = True
'''''                If ((rs("HSXOF4HS") = SIJI) Or (rs("HSXOF4HS") = SANKOU)) Then .L4 = True
'''''                If ((rs("HSXBM1HS") = SIJI) Or (rs("HSXBM1HS") = SANKOU)) Then .B1 = True
'''''                If ((rs("HSXBM2HS") = SIJI) Or (rs("HSXBM2HS") = SANKOU)) Then .B2 = True
'''''                If ((rs("HSXBM3HS") = SIJI) Or (rs("HSXBM3HS") = SANKOU)) Then .B3 = True
'''''                If ((rs("HSXDENHS") = SIJI) Or (rs("HSXDENHS") = SANKOU)) Or _
'''''                   ((rs("HSXLDLHS") = SIJI) Or (rs("HSXLDLHS") = SANKOU)) Or _
'''''                   ((rs("HSXDVDHS") = SIJI) Or (rs("HSXDVDHS") = SANKOU)) Then .GD = True
'''''                If ((rs("HSXCNHWS") = SIJI) Or (rs("HSXCNHWS") = SANKOU)) Then .Cs = True
'''''                If ((rs("HSXLTHWS") = SIJI) Or (rs("HSXLTHWS") = SANKOU)) Then .Lt = True
'''''                .EPD = True
'''''            End With
'''''        End If
'''''
'''''        rs.MoveNext
'''''    Loop
'''''    rs.Close
'''''    Set rs = Nothing
'''''
'''''    '���p�T���v���ł��邩�𒲂ׂ�
'''''    If WSpec(1).Enable And WSpec(2).Enable Then
''''''       sql = "select count(*) as SMPCNT from TBCME043 where CRYNUM='" & CRYNUM & "' and INGOTPOS=" & INGOTPOS
'''''        sql = "select count(*) as SMPCNT from XSDCS where  XTALCS='" & CRYNUM & "' and INPOSCS=" & INGOTPOS
'''''
'''''        Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
'''''        If rs.RecordCount > 0 Then
'''''            If rs("SMPCNT") = 1 Then
'''''                smpShared = True
'''''            End If
'''''        End If
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''
'''''    GetHinbanSpec = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    GetHinbanSpec = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''�T�v      :GD�d�l���擾����
''''''���Ұ�    :�ϐ���        ,IO ,�^                               ,����
''''''          :Gd_Siyou()    ,   ,type_DBDRV_scmzc_fcmkc001c_Siyou ,
''''''          :BLOCKID       ,   ,String                           ,
''''''          :�߂�l        ,O  ,FUNCTION_RETURN                  ,
''''''����      :
''''''����      :
'''''Public Function getGDsiyou(Gd_Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                           BLOCKID As String) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function getGDsiyou"
'''''
'''''    getGDsiyou = FUNCTION_RETURN_SUCCESS
'''''
'''''    sql = sql & "select "
'''''    sql = sql & "E020.HINBAN,"                   ' �i��
'''''    sql = sql & "HSXDENMX,"                     ' �i�r�w�c�������
'''''    sql = sql & "HSXDENMN,"                     ' �i�r�w�c��������
'''''    sql = sql & "HSXLDLMX,"                     ' �i�r�w�k�^�c�k���
'''''    sql = sql & "HSXLDLMN,"                     ' �i�r�w�k�^�c�k����
'''''    sql = sql & "HSXDVDMXN,"                     ' �i�r�w�c�u�c�Q���   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
'''''    sql = sql & "HSXDVDMNN,"                     ' �i�r�w�c�u�c�Q����   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
'''''    sql = sql & "HSXDENHT, "                    ' �i�r�w�c�����ۏؕ��@�Q��
'''''    sql = sql & "HSXDENHS,"                     ' �i�r�w�c�����ۏؕ��@�Q��
'''''    sql = sql & "HSXLDLHT,"                     ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
'''''    sql = sql & "HSXLDLHS,"                     ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
'''''    sql = sql & "HSXDVDHT,"                     ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
'''''    sql = sql & "HSXDVDHS,"                     ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
'''''    sql = sql & "HSXDENKU,"                     ' �i�r�w�c���������L��
'''''    sql = sql & "HSXDVDKU,"                     ' �i�r�w�c�u�c�Q�����L��
'''''    sql = sql & "HSXLDLKU "                      ' �i�r�w�k�^�c�k�����L��
'''''    sql = sql & "from TBCME020 E020,TBCME041 E041,TBCME040  E040 "
'''''    sql = sql & "where E040.BLOCKID = '" & BLOCKID & "'"
'''''    sql = sql & "   and  E041.CRYNUM = E040.CRYNUM"
'''''    sql = sql & "   and E040.INGOTPOS<=E041.INGOTPOS"
'''''    sql = sql & "   and E041.INGOTPOS< E040.INGOTPOS+E040.LENGTH"
'''''    sql = sql & "   and E041.HINBAN = E020.HINBAN"
'''''    sql = sql & "   and E041.OPECOND = E020.OPECOND"
'''''
'''''
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    If rs.RecordCount = 0 Then
'''''        getGDsiyou = FUNCTION_RETURN_FAILURE
'''''        ReDim Gd_Siyou(0)
'''''        rs.Close
'''''        GoTo proc_exit
'''''    End If
'''''
'''''
'''''    recCnt = rs.RecordCount
'''''    ReDim Gd_Siyou(recCnt)
'''''
'''''    For i = 1 To recCnt
'''''        With Gd_Siyou(i)
'''''
'''''            .hin.hinban = rs("HINBAN")            ' �i��
'''''            .HSXDENMX = rs("HSXDENMX")            ' �i�r�w�c�������
'''''            .HSXDENMN = rs("HSXDENMN")             ' �i�r�w�c��������
'''''            .HSXLDLMX = rs("HSXLDLMX")            ' �i�r�w�k�^�c�k���
'''''            .HSXLDLMN = rs("HSXLDLMN")            ' �i�r�w�k�^�c�k����
'''''            .HSXDVDMX = rs("HSXDVDMXN")            ' �i�r�w�c�u�c�Q���   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
'''''            .HSXDVDMN = rs("HSXDVDMNN")            ' �i�r�w�c�u�c�Q����   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
'''''            .HSXDENHT = rs("HSXDENHT")            ' �i�r�w�c�����ۏؕ��@�Q��
'''''            .HSXDENHS = rs("HSXDENHS")            ' �i�r�w�c�����ۏؕ��@�Q��
'''''            .HSXLDLHT = rs("HSXLDLHT")            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
'''''            .HSXLDLHS = rs("HSXLDLHS")            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
'''''            .HSXDVDHT = rs("HSXDVDHT")            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
'''''            .HSXDVDHS = rs("HSXDVDHS")            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
'''''            .HSXDENKU = rs("HSXDENKU")            ' �i�r�w�c���������L��
'''''            .HSXDVDKU = rs("HSXDVDKU")            ' �i�r�w�c�u�c�Q�����L��
'''''            .HSXLDLKU = rs("HSXLDLKU")            ' �i�r�w�k�^�c�k�����L��
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''�T�v      :���H���є���ɍ\���̂ɒl���Z�b�g����
''''''���Ұ�    :�ϐ���        ,IO ,�^             ,����
''''''          :BLOCKID       ,   ,String         ,�u���b�NID
''''''          :Kakou         ,   ,type_KakouJudg ,���H���є���\����
''''''          :�߂�l        ,O  ,FUNCTION_RETURN,
''''''����      :�u���b�N���S�i�Ԃ̎d�l�Ǝ��т����߂�
''''''����      :2002/4/16 ���� �쐬
'''''Public Function DBDRV_scmzc_fcmkc001c_Kakou(BLOCKID As String, Kakou As type_KakouJudg) As FUNCTION_RETURN
'''''    Dim sql As String
'''''    Dim sql1 As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim c0 As Integer
'''''    Dim tHIN() As tFullHinban
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_Kakou"
'''''
'''''    DBDRV_scmzc_fcmkc001c_Kakou = FUNCTION_RETURN_FAILURE
'''''
'''''    '�u���b�N���̑S�i�Ԃ����߂�
'''''    sql = "select HINBAN, REVNUM, FACTORY, OPECOND from TBCME040 E40, TBCME041 E41 "
'''''    sql = sql & "Where E41.CRYNUM = E40.CRYNUM and "
'''''    sql = sql & "E40.BLOCKID = '" & BLOCKID & "' and "
'''''    sql = sql & "E40.INGOTPOS < E41.INGOTPOS+E41.LENGTH and "
'''''    sql = sql & "E40.INGOTPOS+E40.LENGTH > E41.INGOTPOS"
'''''
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    recCnt = rs.RecordCount
'''''    ReDim tHIN(recCnt)
'''''    If recCnt = 0 Then
'''''        rs.Close
'''''        GoTo PROC_EXIT
'''''    End If
'''''    For c0 = 1 To recCnt
'''''        tHIN(c0).hinban = rs("HINBAN")
'''''        tHIN(c0).mnorevno = rs("REVNUM")
'''''        tHIN(c0).factory = rs("FACTORY")
'''''        tHIN(c0).opecond = rs("OPECOND")
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    '���߂��S�i�Ԃ̉��H�d�l�����߂�
'''''    If scmzc_getKakouSpec(tHIN(), Kakou.Spec()) = FUNCTION_RETURN_FAILURE Then
'''''        GoTo PROC_EXIT
'''''    End If
'''''
'''''    '�Ώۃu���b�N�̉��H���т����߂�
'''''    If scmzc_getKakouJiltuseki(BLOCKID, Kakou.Jiltuseki) = FUNCTION_RETURN_FAILURE Then
'''''        GoTo PROC_EXIT
'''''    End If
'''''
'''''    DBDRV_scmzc_fcmkc001c_Kakou = FUNCTION_RETURN_SUCCESS
'''''
'''''PROC_EXIT:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function


''''''�T�v      :�u���b�N���̕i�Ԃɂ���Cs�d�l�̗L�����`�F�b�N����
''''''���Ұ�    :�ϐ���        ,IO ,�^        ,����
''''''          :crynum        ,I  ,String    ,�����ԍ�
''''''          :blkFrom       ,I  ,Integer   ,�u���b�N�J�n�ʒu
''''''          :blkTo         ,I  ,Integer   ,�u���b�N�I���ʒu
''''''          :hasCs         ,O  ,String    ,�u���b�N���̕i�Ԃ�Cs�d�l�������̂����邩(='H':�ۏ؂��� ='S':�Q�l���� ��:�d�l�Ȃ�)
''''''          :hasCsFromTo   ,O  ,String    ,�u���b�N���̕i�Ԃ�FromTo��Cs�d�l�������̂����邩(='H':�ۏ؂��� ='S':�Q�l���� ��:�d�l�Ȃ�)
''''''          :�߂�l        ,O  ,FUNCTION_RETURN,
''''''����      :�����ʂ�Top��/Bot���ɂ��ĕ\���E������s�����ǂ��������肷�邽�߂ɗ��p����
''''''����      :2002/4/16 �쑺 �쐬
'''''Public Function DBDRV_scmzc_fcmkc001c_CheckSpecCs(CRYNUM As String, BlkFrom As Integer, BlkTo As Integer, jCs As String, jCsFromTo As String) As FUNCTION_RETURN
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''Dim HSXCNHWS As String
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_CheckSpecCs"
'''''
'''''    DBDRV_scmzc_fcmkc001c_CheckSpecCs = FUNCTION_RETURN_FAILURE
'''''
'''''    jCs = " "
'''''    jCsFromTo = " "
'''''    sql = "select HSXCNHWS, HSXCNMIN "
'''''    sql = sql & "from TBCME041 HIN, TBCME019 SPEC "
'''''    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
'''''    sql = sql & "  and HIN.INGOTPOS<" & BlkTo
'''''    sql = sql & "  and HIN.INGOTPOS+LENGTH>" & BlkFrom
'''''    sql = sql & "  and SPEC.HINBAN=HIN.HINBAN"
'''''    sql = sql & "  and SPEC.MNOREVNO=HIN.REVNUM"
'''''    sql = sql & "  and SPEC.FACTORY=HIN.FACTORY"
'''''    sql = sql & "  and SPEC.OPECOND=HIN.OPECOND"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    '�i�Ԗ���Cs�d�l�ECsFromTo�d�l��'H','S'�𒲂ׂ�
'''''    Do While rs.EOF = False
'''''        HSXCNHWS = rs("HSXCNHWS")
'''''        If HSXCNHWS = SIJI Then
'''''            jCs = HSXCNHWS
'''''            If rs("HSXCNMIN") > 0# Then
'''''                jCsFromTo = HSXCNHWS
'''''            End If
'''''        ElseIf HSXCNHWS = SANKOU Then
'''''            If jCs <> SIJI Then jCs = HSXCNHWS
'''''            If rs("HSXCNMIN") > 0# Then
'''''                If jCsFromTo <> SIJI Then jCsFromTo = HSXCNHWS
'''''            End If
'''''        End If
'''''        rs.MoveNext
'''''    Loop
'''''    rs.Close
'''''    Set rs = Nothing
'''''    DBDRV_scmzc_fcmkc001c_CheckSpecCs = FUNCTION_RETURN_SUCCESS
'''''
'''''PROC_EXIT:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function


''''''�T�v      :�u���b�N���̕i�Ԃɂ���Cs�d�l���擾����('H'or'S'�̂��̂̂�)
''''''���Ұ�    :�ϐ���        ,IO ,�^        ,����
''''''          :crynum        ,   ,String    ,
''''''          :blkFrom       ,   ,Integer   ,
''''''          :blkTo         ,   ,Integer   ,
''''''          :SpecCs()      ,   ,C_Cs      ,
''''''          :�߂�l        ,O  ,FUNCTION_R,
''''''����      :
''''''����      :
'''''Public Function DBDRV_scmzc_fcmkc001c_GetSpecCs(CRYNUM As String, BlkFrom As Integer, BlkTo As Integer, SpecCs() As C_Cs) As FUNCTION_RETURN
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''Dim recCnt As Long
'''''Dim i As Long
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_GetSpecCs"
'''''
'''''    DBDRV_scmzc_fcmkc001c_GetSpecCs = FUNCTION_RETURN_FAILURE
'''''    sql = "select HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT,HSXCNHWS, HSXCNMIN, HSXCNMAX "
'''''    sql = sql & "from TBCME041 HIN, TBCME019 SPEC "
'''''    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
'''''    sql = sql & "  and HIN.INGOTPOS<" & BlkTo
''''''�쑺���̎w���ɂ��ύX
''''''2002/05/11 S.Sano    sql = sql & "  and HIN.INGOTPOS+LENGTH>=" & BlkFrom
'''''    sql = sql & "  and HIN.INGOTPOS+LENGTH>" & BlkFrom '2002/05/11 S.Sano
'''''    sql = sql & "  and SPEC.HINBAN=HIN.HINBAN"
'''''    sql = sql & "  and SPEC.MNOREVNO=HIN.REVNUM"
'''''    sql = sql & "  and SPEC.FACTORY=HIN.FACTORY"
'''''    sql = sql & "  and SPEC.OPECOND=HIN.OPECOND"
'''''    sql = sql & "  and SPEC.HSXCNHWS in ('H','S')"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    recCnt = rs.RecordCount
'''''    If recCnt = 0 Then
'''''        ReDim SpecCs(0)
'''''    Else
'''''        ReDim SpecCs(1 To recCnt)
'''''        For i = 1 To recCnt
'''''            With SpecCs(i)
'''''                .GuaranteeCs.cMeth = rs("HSXCNSPH")
'''''                .GuaranteeCs.cCount = rs("HSXCNSPT")
'''''                .GuaranteeCs.cPos = rs("HSXCNSPI")
'''''                .GuaranteeCs.cObj = rs("HSXCNHWT")
'''''                .GuaranteeCs.cJudg = rs("HSXCNHWS")
'''''                .SpecCsMin = rs("HSXCNMIN")
'''''                .SpecCsMax = rs("HSXCNMAX")
'''''            End With
'''''            rs.MoveNext
'''''        Next
'''''    End If
'''''    rs.Close
'''''    Set rs = Nothing
'''''
'''''    DBDRV_scmzc_fcmkc001c_GetSpecCs = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''�T�v      :�u���b�N���̕i�Ԃɂ���Lt�d�l���擾����('H'or'S'�̂��̂̂�)
''''''���Ұ�    :�ϐ���        ,IO ,�^        ,����
''''''          :crynum        ,   ,String    ,
''''''          :blkFrom       ,   ,Integer   ,
''''''          :blkTo         ,   ,Integer   ,
''''''          :SpecLt()      ,   ,C_Lt      ,
''''''          :�߂�l        ,O  ,FUNCTION_R,
''''''����      :
''''''����      :
'''''Public Function DBDRV_scmzc_fcmkc001c_GetSpecLt(CRYNUM As String, BlkFrom As Integer, BlkTo As Integer, SpecLt() As C_LT) As FUNCTION_RETURN
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''Dim recCnt As Long
'''''Dim i As Long
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_GetSpecLt"
'''''
'''''    DBDRV_scmzc_fcmkc001c_GetSpecLt = FUNCTION_RETURN_FAILURE
'''''    sql = "select HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT,HSXLTHWS, HSXLTMIN, HSXLTMAX "
'''''    sql = sql & "from TBCME041 HIN, TBCME019 SPEC "
'''''    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
'''''    sql = sql & "  and HIN.INGOTPOS<" & BlkTo
'''''    sql = sql & "  and HIN.INGOTPOS+LENGTH>=" & BlkFrom
'''''    sql = sql & "  and SPEC.HINBAN=HIN.HINBAN"
'''''    sql = sql & "  and SPEC.MNOREVNO=HIN.REVNUM"
'''''    sql = sql & "  and SPEC.FACTORY=HIN.FACTORY"
'''''    sql = sql & "  and SPEC.OPECOND=HIN.OPECOND"
'''''    sql = sql & "  and SPEC.HSXLTHWS in ('H','S')"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    recCnt = rs.RecordCount
'''''    If recCnt = 0 Then
'''''        ReDim SpecLt(0)
'''''    Else
'''''        ReDim SpecLt(1 To recCnt)
'''''        For i = 1 To recCnt
'''''            With SpecLt(i)
'''''                .GuaranteeLt.cMeth = rs("HSXLTSPH")
'''''                .GuaranteeLt.cCount = rs("HSXLTSPT")
'''''                .GuaranteeLt.cPos = rs("HSXLTSPI")
'''''                .GuaranteeLt.cObj = rs("HSXLTHWT")
'''''                .GuaranteeLt.cJudg = rs("HSXLTHWS")
'''''                .SpecLtMin = rs("HSXLTMIN")
'''''                .SpecLtMax = rs("HSXLTMAX")
'''''            End With
'''''            rs.MoveNext
'''''        Next
'''''    End If
'''''    rs.Close
'''''    Set rs = Nothing
'''''
'''''    DBDRV_scmzc_fcmkc001c_GetSpecLt = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''�T�v      :�u���b�N����LT�u�ۏ؁v�͂��邩�H�Ȃ���΁u�Q�l�v�͂��邩�H
''''''���Ұ�    :�ϐ���        ,IO ,�^          ,����
''''''          :Crynum        ,I  ,String      ,�����ԍ�
''''''          :BlkFrom       ,I  ,Integer     ,�u���b�N�̊J�n�ʒu
''''''          :BlkTo         ,I  ,Integer     ,�u���b�N�̏I���ʒu
''''''          :�߂�l        ,O  ,String      ,'H':�ۏ؂��� 'S':�Q�l���� vbNullString:�Ȃ�
''''''����      :
''''''����      :2002/10/10 �쑺 �쐬
'''''Public Function DBDRV_getLtGuaranteeInBlock(CRYNUM As String, BlkFrom As Integer, BlkTo As Integer) As String
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_getLtGuaranteeInBlock"
'''''    DBDRV_getLtGuaranteeInBlock = vbNullString
'''''
'''''    sql = "select SIYO.HSXLTHWS "
'''''    sql = sql & "from TBCME041 HIN, TBCME019 SIYO "
'''''    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
'''''    sql = sql & "  and HIN.INGOTPOS<" & BlkTo & " and HIN.INGOTPOS+HIN.LENGTH>" & BlkFrom
'''''    sql = sql & "  and SIYO.HINBAN=HIN.HINBAN and SIYO.MNOREVNO=HIN.REVNUM and SIYO.FACTORY=HIN.FACTORY and SIYO.OPECOND=HIN.OPECOND"
'''''    sql = sql & "  and SIYO.HSXLTHWS in ('H','S') "
'''''    sql = sql & "order by SIYO.HSXLTHWS"
'''''    sql = "select * from (" & sql & ") where rownum=1"
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    If rs.RecordCount > 0 Then
'''''        DBDRV_getLtGuaranteeInBlock = rs("HSXLTHWS")
'''''    Else
'''''        DBDRV_getLtGuaranteeInBlock = vbNullString
'''''    End If
'''''    rs.Close
'''''
'''''PROC_EXIT:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function
'''''============================================================================================================================


'�T�v      :����Z�o�f�[�^�̓o�^
'���Ұ�    :�ϐ���        ,IO ,�^                       ,����
'          :RS            ,I  ,typ_TBCMJ002             ,������R���уe�[�u���ւ̑}���p
'          :�߂�l        ,O  ,FUNCTION_RETURN          ,����
'����      :
'����      :2001/06/27 ���{ �쐬
Public Function DBDRV_SuiteiZis_InsRS(rs() As typ_TBCMJ002) As FUNCTION_RETURN

    Dim lcnt    As Integer
    Dim sql     As String
                                          
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_SuiteiZis_InsRS"
    
    '����Z�o�f�[�^��������R���уe�[�u��(TBCMJ002)�֓o�^����B
    For lcnt = 1 To UBound(rs)
        With rs(lcnt)
            sql = "insert into TBCMJ002 ("
            sql = sql & "CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, "
            sql = sql & "HINBAN, REVNUM, factory, opecond, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, "
            sql = sql & "EFEHS, RRG, JudgData, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, SUIFLG) "
            
            sql = sql & "values ('"
            sql = sql & .CRYNUM & "', "             '�����ԍ�
            sql = sql & .POSITION & ", '"           '�ʒu
            sql = sql & .SMPKBN & "', '"            '�T���v���敪"
            sql = sql & .TRANCOND & "', "           '��������
            sql = sql & .TRANCNT & ", "             '������
            sql = sql & .SMPLNO & ", '"             '�T���v����
            sql = sql & .SMPLUMU & "', '"           '�T���v���L��
            sql = sql & .KRPROCCD & "', '"          '�Ǘ��H���R�[�h
            sql = sql & .PROCCODE & "', '"          '�H���R�[�h
            sql = sql & .hinban & "', "             '�i��
            sql = sql & .REVNUM & ", '"             '���i�ԍ������ԍ�
            sql = sql & .factory & "', '"           '�H��
            sql = sql & .opecond & "', '"           '���Ə���
            sql = sql & .GOUKI & "', '"             '���@
            sql = sql & .Type & "', "               '�^�C�v
            sql = sql & .MEAS1 & ", "               '����l�P
            sql = sql & .MEAS2 & ", "               '����l�Q
            sql = sql & .MEAS3 & ", "               '����l�R
            sql = sql & .MEAS4 & ", "               '����l�S
            sql = sql & .MEAS5 & ", "               '����l�T
            sql = sql & .EFEHS & ", "               '���s�ΐ�
            sql = sql & .RRG & ", "                 '�q�q�f
            sql = sql & .JudgData & ", '"           '�����Ώےl
            sql = sql & .TSTAFFID & "', "           '�o�^�Ј�ID
            sql = sql & "sysdate, '"                '�o�^���t
            sql = sql & .KSTAFFID & "', "           '�X�V�Ј�ID
            sql = sql & "sysdate, '"                '�X�V���t
            sql = sql & .SENDFLAG & "', "           '���M�t���O
            sql = sql & "sysdate, '"                '���M���t
            sql = sql & .SUIFLG & "')"              '����FLG"
        End With
        
        '' WriteDBLog sql
        If OraDB.ExecuteSQL(sql) <= 0 Then GoTo proc_err
    
    Next lcnt

    DBDRV_SuiteiZis_InsRS = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_SuiteiZis_InsRS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

