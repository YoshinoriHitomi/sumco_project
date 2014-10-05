Attribute VB_Name = "s_cmbc032_SQL"
Option Explicit

'�����ŏI���o����


Type cmkc001b_LockWait
    flag                As Boolean
    Grp                 As Integer
End Type

Type cmkc001b_Wait3_HINBAN
    hinban              As String * 8                           ' �i��
    REVNUM              As Integer                              ' ���i�ԍ������ԍ�
    factory             As String * 1                           ' �H��
    opecond             As String * 1                           ' ���Ə���
End Type

Type cmkc001b_Wait3_BLK
    BLOCKID             As String * 12                          ' �u���b�NID
    INGOTPOS            As Integer                              ' �������J�n�ʒu
    LENGTH              As Integer                              ' ����
    NOWPROC             As String * 5                           ' ���ݍH��
    HOLDCLS             As String * 1                           ' �z�[���h�敪 ---kuramoto �ǉ� 2001/09/19----
    GRPFLG1             As Integer                              ' �O���[�v���
    GRPFLG2             As Integer                              ' �O���[�v���
    COLORFLG            As Boolean
    topHin              As cmkc001b_Wait3_HINBAN
    botHin              As cmkc001b_Wait3_HINBAN
End Type

Type cmkc001b_Wait3
    CRYNUM              As String * 12           ' �����ԍ�
    blkInfo()           As cmkc001b_Wait3_BLK
End Type

'�u���b�N�Ǘ�
Public Type typ_cmkc001f_Block
    'E040 �u���b�N�Ǘ�
    INGOTPOS            As Integer              ' �������J�n�ʒu
    LENGTH              As Integer              ' ����
    REALLEN             As Integer              ' ������
    KRPROCCD            As String * 5           ' ���݊Ǘ��H��
    NOWPROC             As String * 5           ' ���ݍH��
    LPKRPROCCD          As String * 5           ' �ŏI�ʉߊǗ��H��
    LASTPASS            As String * 5           ' �ŏI�ʉߍH��
    DELCLS              As String * 1           ' �폜�敪
    RSTATCLS            As String * 1           ' ������ԋ敪
    LSTATCLS            As String * 1           ' �ŏI��ԋ敪 */
    'E037 �������Ǘ�
    SEED                As String               'SEED
End Type


'�d�l�擾�p
Public Type typ_cmkc001f_Disp
    '�i�ԊǗ�
    hinban              As String * 8            ' �i��
    INGOTPOS            As Integer               ' �������J�n�ʒu
    REVNUM              As Integer               ' ���i�ԍ������ԍ�
    factory             As String * 1            ' �H��
    opecond             As String * 1            ' ���Ə���
    LENGTH              As Integer               ' ����
    '���i�d�lSXL�f�[�^
    HSXD1CEN            As Double                ' �i�r�w���a�P���S
    HSXRMIN             As Double                ' �i�r�w���R����
    HSXRMAX             As Double                ' �i�r�w���R���
    HSXRMBNP            As Double                ' �i�r�w���R�ʓ����z
    HSXRHWYS            As String * 1            ' �i�r�w���R�ۏؕ��@�Q��
    HSXONMIN            As Double                ' �i�r�w�_�f�Z�x����
    HSXONMAX            As Double                ' �i�r�w�_�f�Z�x���
    HSXONMBP            As Double                ' �i�r�w�_�f�Z�x�ʓ����z
    HSXONHWS            As String * 1            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    HSXCNMIN            As Double                ' �i�r�w�Y�f�Z�x����
    HSXCNMAX            As Double                ' �i�r�w�Y�f�Z�x���
    HSXCNHWS            As String * 1            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    HSXTMMAX            As Double                ' �i�r�w�]�ʖ��x���         ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    HSXBMnAN(1 To 3)    As Double                ' �i�r�w�a�l�cn ���ω���
    HSXBMnAX(1 To 3)    As Double                ' �i�r�w�a�l�cn ���Ϗ��
    HSXBMnHS(1 To 3)    As String * 1            ' �i�r�w�a�l�cn �ۏؕ��@�Q��
    HSXOFnAX(1 To 4)    As Double                ' �i�r�w�n�r�en���Ϗ��
    HSXOFnMX(1 To 4)    As Double                ' �i�r�w�n�r�en���
    HSXOFnHS(1 To 4)    As String * 1            ' �i�r�w�n�r�en �ۏؕ��@�Q��
    HSXDENMX            As Integer               ' �i�r�w�c�������
    HSXDENMN            As Integer               ' �i�r�w�c��������
    HSXDENHS            As String * 1            ' �i�r�w�c�����ۏؕ��@�Q��
    HSXDVDMX            As Integer               ' �i�r�w�c�u�c�Q���
    HSXDVDMN            As Integer               ' �i�r�w�c�u�c�Q����
    HSXDVDHS            As String * 1            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    HSXLDLMX            As Integer               ' �i�r�w�k�^�c�k���
    HSXLDLMN            As Integer               ' �i�r�w�k�^�c�k����
    HSXLDLHS            As String * 1            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    HSXLTMIN            As Integer               ' �i�r�w�k�^�C������
    HSXLTMAX            As Integer               ' �i�r�w�k�^�C�����
    HSXLTHWS            As String * 1            ' �i�r�w�k�^�C���ۏؕ��@�Q��
    HSXDPDIR            As String * 2            ' �i�r�w�a�ʒu����
    HSXDPDRC            As String * 1            ' �i�r�w�a�ʒu����
    HSXDWMIN            As Double                ' �i�r�w�a�Љ���
    HSXDWMAX            As Double                ' �i�r�w�a�Џ��
    HSXDDMIN            As Double                ' �i�r�w�a�[����
    HSXDDMAX            As Double                ' �i�r�w�a�[���
    HSXD1MIN            As Double                ' �i�r�w���a�P����
    HSXD1MAX            As Double                ' �i�r�w���a�P���
    HSXCTCEN            As Double                ' �i�r�w�����ʌX�c���S
    HSXCYCEN            As Double                ' �i�r�w�����ʌX�����S
    EPDUP               As Integer               ' ���������Ǘ� EPD�@���
End Type


'���s�����͗p
Public Type typ_cmkc001f_ExecCryIn
    CRYNUM              As String * 12       ' �����ԍ�(IN)
    INGOTPOS            As Integer         ' �C���S�b�g���ʒu(IN)
End Type


'�����ŏI����
Public Type typ_cmkc001f_ExecFts
    LENGTH              As Integer           ' ����
    KRPROCCD            As String * 5      ' �Ǘ��H���R�[�h
    PROCCODE            As String * 5      ' �H���R�[�h
    PAYCLASS            As String * 1      ' �����o���敪
    OUTLENGTH           As Integer        ' �o�ג���
    PART(1 To 5)        As Integer     ' ����n
    BDLEN(1 To 5)       As Integer    ' ����n �s�ǒ���
    BDCAUS(1 To 5)      As String * 3 ' ����n �s�Ǘ��R
    TSTAFFID            As String * 8      ' �o�^�Ј�ID
End Type


'�N���X�^���J�^���O���
Public Type typ_cmkc001f_ExecCatalog
    CRYNUM              As String * 12       ' �����ԍ�
    KRPROCCD            As String * 5      ' �Ǘ��H���R�[�h
    PROCCODE            As String * 5      ' �H���R�[�h
    BDCODE              As String * 3        ' �s�Ǘ��R�R�[�h
    PALTNUM             As String * 4       ' �p���b�g�ԍ�
    TSTAFFID            As String * 8      ' �o�^�Ј�ID
End Type

'------------------------------------------------------------------------
Type type_cmkc001b_SmpMng
    CRYNUM              As String * 12
    INGOTPOS            As Integer
    SMPKBN              As String * 1
    
    hinban              As String * 8            ' �i��
    REVNUM              As Integer               ' ���i�ԍ������ԍ�
    factory             As String * 1           ' �H��
    opecond             As String * 1           ' ���Ə���
    
    
    CRYINDRS            As String * 1
    CRYRESRS            As String * 1
    CRYINDOI            As String * 1
    CRYRESOI            As String * 1
    CRYINDB1            As String * 1
    CRYRESB1            As String * 1
    CRYINDB2            As String * 1
    CRYRESB2            As String * 1
    CRYINDB3            As String * 1
    CRYRESB3            As String * 1
    CRYINDL1            As String * 1
    CRYRESL1            As String * 1
    CRYINDL2            As String * 1
    CRYRESL2            As String * 1
    CRYINDL3            As String * 1
    CRYRESL3            As String * 1
    CRYINDL4            As String * 1
    CRYRESL4            As String * 1
    CRYINDCS            As String * 1
    CRYRESCS            As String * 1
    CRYINDGD            As String * 1
    CRYRESGD            As String * 1
    CRYINDT             As String * 1
    CRYREST             As String * 1
    CRYINDEP            As String * 1
    CRYRESEP            As String * 1
    
    HSXCNHWS            As String * 1          ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    HSXLTHWS            As String * 1          ' �i�r�w�k�^�C���ۏؕ��@�Q��
    EPD                 As String * 1               ' EPD
End Type

Private Type tSmpMng
    BLOCKID             As String * 12
    TOPPOS              As Integer
    BOTPOS              As Integer
    
    CRYNUM              As String * 12
    INGOTPOS            As Integer
    SMPKBN              As String * 1
    
    hinban              As String * 8            ' �i��
    REVNUM              As Integer               ' ���i�ԍ������ԍ�
    factory             As String * 1           ' �H��
    opecond             As String * 1           ' ���Ə���
    
    CRYINDRS            As String * 1
    CRYRESRS            As String * 1
    CRYINDOI            As String * 1
    CRYRESOI            As String * 1
    CRYINDB1            As String * 1
    CRYRESB1            As String * 1
    CRYINDB2            As String * 1
    CRYRESB2            As String * 1
    CRYINDB3            As String * 1
    CRYRESB3            As String * 1
    CRYINDL1            As String * 1
    CRYRESL1            As String * 1
    CRYINDL2            As String * 1
    CRYRESL2            As String * 1
    CRYINDL3            As String * 1
    CRYRESL3            As String * 1
    CRYINDL4            As String * 1
    CRYRESL4            As String * 1
    CRYINDCS            As String * 1
    CRYRESCS            As String * 1
    CRYINDGD            As String * 1
    CRYRESGD            As String * 1
    CRYINDT             As String * 1
    CRYREST             As String * 1
    CRYINDEP            As String * 1
    CRYRESEP            As String * 1
End Type

'�҂��ꗗ
Public Type typ_HinMap      '2006/02
    HIN         As tFullHinban                  ' �i��
    LENGTH      As Integer                      ' ����
    Weight      As Double                       ' �d��
End Type


'�����\���p
Public Type type_DBDRV_scmzc_fcmkc001b_Disp
    CRYNUM              As String * 12          ' �����ԍ�
    INGOTPOS            As Integer              ' �������J�n�ʒu
    INPOS               As Integer              ' TOP�ʒu
    LENGTH              As Integer              ' ����              '2001/11/8
    BLOCKID             As String * 12          ' �u���b�NID
    HSXTYPE             As String * 1           ' �i�r�w�^�C�v
    HSXCDIR             As String * 1           ' �i�r�w�����ʕ���
    UPDDATE             As Date                 ' �X�V���t
    Judg                As String               ' ����
    hinM()              As typ_HinMap           ' �i��(full)   '2006/02
    HOLDCLS             As String * 1           ' �z�[���h�敪 ---kuramoto �ǉ� 2001/09/25----
    SMP()               As type_cmkc001b_SmpMng ' �T���v���Ǘ�
    PUPTN               As String               ' ���������   ---kubota �ǉ� 2004/12/21----
    HOLDB               As String               ' 2005/08
    HOLDC               As String               ' 2005/08
    HOLDKT              As String               ' 2005/08
    LBLFLG              As String               ' ���x�����s�t���O  2005/11 ADD
    DIA                 As Integer              ' ���a 2006/02
    KIKBN               As String               '�����ʋ敪 2006/11/14 SETsw kubota
    AGRSTATUS           As String               ' ���F�m�F�敪 add SETkimizuka
    STOP                As String               ' ��~     add SETkimizuka
    CAUSE               As String               ' ��~���R add SETkimizuka
    PRINTNO             As String               ' ��s�]�� add SETkimizuka
End Type


'�ꗗ�\���p(�o�בO�ꗗ)    2008/05/26 SHINDOH
Public Type type_DBDRV_scmzc_fcmkc001b_Disp52
    PLANT               As String               ' ����
    CRYNUM              As String * 12          ' �����ԍ�
    INGOTPOS            As Integer              ' �������J�n�ʒu
    INPOS               As Integer              ' TOP�ʒu
    LENGTH              As Integer              ' ����
    BLOCKID             As String * 12          ' �u���b�NID
    HSXTYPE             As String * 1           ' �i�r�w�^�C�v
    HSXCDIR             As String * 1           ' �i�r�w�����ʕ���
    UPDDATE             As Date                 ' �X�V���t
    hinM()              As typ_HinMap           ' �i��(full)
    HOLDCLS             As String * 1           ' �z�[���h�敪
    PUPTN               As String               ' ���������
    HOLDB               As String               ' 2005/08
    HOLDC               As String               ' 2005/08
    HOLDKT              As String               ' 2005/08
End Type
'�����\���p         2008/05/26 SHINDOH
Public Type type_DBDRV_scmzc_fcmkc001b_Disp5
    CRYNUM              As String * 12                          ' �����ԍ�
    INGOTPOS            As Integer                              ' �������J�n�ʒu
'   LENGTH              As Integer                              ' ����              '2001/11/8
    BLOCKID             As String * 12                          ' �u���b�NID
    HSXTYPE             As String * 1                           ' �i�r�w�^�C�v
    HSXCDIR             As String * 1                           ' �i�r�w�����ʕ���
    UPDDATE             As Date                                 ' �X�V���t
    Judg                As String                               ' ����
    HIN()               As tFullHinban                          ' �i��(full)
    HOLDCLS             As String * 1                           ' �z�[���h�敪 ---kuramoto �ǉ� 2001/09/25----
    SMP()               As type_cmkc001b_SmpMng                 ' �T���v���Ǘ�
    PUPTN               As String                               ' ���������    ---kubota �ǉ� 2004/12/08----
    WFCUTT              As Integer                              ' WF��ĒP�ʁ@05/04/19 ooba
    BLOCKHFLAG          As String * 1                           ' ��ۯ��P�ʕۏ��׸ށ@05/04/19 ooba
    HOLDBCA             As String * 1                           ' ΰ��ދ敪(XSDCA)�@05/04/19 ooba
    HOLDB               As String               '2006/01
    HOLDC               As String               '2006/01
    HOLDKT              As String               '2006/01
    AGRSTATUS           As String           ' ���F�m�F�敪 add SETkimizuka
    STOP                As String           ' ��~     add SETkimizuka
    CAUSE               As String           ' ��~���R add SETkimizuka
    PRINTNO             As String           ' ��s�]�� add SETkimizuka
End Type
'�u���b�N�̏�� '2008/05/28 SHINDOH
Public Type typ_BlkData
    CRYNUM              As String * 12      ' �����ԍ�
    BLOCKID             As String * 12      ' �u���b�NID
    INGOTPOS            As Integer          ' �C���S�b�g���ʒu
    LENGTH              As Integer          ' �u���b�N����
    REALLEN             As Integer          ' �u���b�N������
    sBlockId            As String * 12      ' ���o�擪�u���b�NID
    BLOCKORDER          As Integer          ' �u���b�N����
    DIAMETER            As Double           ' ���a 2002/05/01 S.Sano
    WFINDDATE           As String * 10      ' �ŏI�������t
    HOLDCLS             As String * 1       ' �z�[���h���
End Type

'�u���b�N���i�ԏ�� '2008/05/28 SHINDOH
Public Type typ_BlkHinMap
    BLOCKID             As String * 12      ' �u���b�NID
    HIN                 As tFullHinban      ' �i��
    REALLEN             As Integer          ' �i�Ԏ�����
    HinLen              As Integer          ' ���i��
    PASSFLAG            As String * 1       ' �ʉ߃t���O
    INPOSCA             As Integer          ' �������J�n�ʒu�@--- 2007/07/17 shindo �ǉ� ---
    PLANTCATCA          As String           ' ���� 2007/09/12 SPK Tsutsumi Add
End Type

' �u���b�N�ꗗ  '2008/05/28 SHINDOH
Public Type typ_BlkMap
    BLOCKID             As String * 12      ' �u���b�NID
    HIN(1 To 5)         As tFullHinban      ' �i��
    WFINDDATE           As String * 10      ' �ŏI�������t
    CRYNUM              As String * 12      ' �����ԍ�
    INGOTPOS            As Integer          ' �C���S�b�g���ʒu
    LENGTH              As Integer          ' �u���b�N����
    REALLEN             As Integer          ' �u���b�N������
    HINREALLEN(1 To 5)  As Integer          ' �i�Ԏ�����
    HinLen(1 To 5)      As Integer          ' �i�Ԓ���
    DIAMETER            As Double           ' ���a 2002/05/01 S.Sano
    sBlockId            As String * 12      ' �擪�u���b�NID
    BLOCKORDER          As Integer          ' �u���b�N����
    HOLDCLS             As String * 1       ' �z�[���h���  --- 2001/09/19 kuramoto �ǉ� ---
    PASSFLAG            As String * 1       ' �ʉ߃t���O�@�@--- 200/04/16 Yam
End Type
''�u���b�N���i�ԏ��(�\���i�Ԏ擾�p)�@�@--- '2008/05/28 SHINDOH
Public Type typ_WkBlkMap
    BLOCKID             As String * 12      ' �u���b�NID
    HINCNT As Integer
    HIN()         As tFullHinban      ' �i��
    HINREALLEN()  As Integer          ' �i�Ԏ�����
    HinLen()      As Integer          ' �i�Ԓ���
    INPOSCA() As Integer '�������J�n�ʒu
End Type

'�i�ԏ��--- '2008/05/28 SHINDOH
Public Wk_tblBlkMap() As typ_WkBlkMap


'�i�ԁA�d�l�A���������擾�p (TOP,TAIL���łQ���R�[�h�擾)
Public Type type_DBDRV_scmzc_fcmkc001c_Siyou
    '�u���b�N�Ǘ�
    CRYNUM              As String * 12          ' �����ԍ�
    INGOTPOS            As Integer              ' �������J�n�ʒu
    LENGTH              As Integer              ' ����
    '�i�ԊǗ�
    HIN                 As tFullHinban          ' �i��(full)
        
    '�������
    PRODCOND            As String * 4           ' �������
    PGID                As String * 8           ' �o�f�|�h�c
    UPLENGTH            As Integer              ' ���グ����
    FREELENG            As Integer              ' �t���[��
    DIAMETER            As Integer              ' ���a 2002/05/01 S.Sano
    CHARGE              As Double               ' �`���[�W��
    SEED                As String * 4           ' �V�[�h
    ADDDPPOS            As Integer              ' �ǉ��h�[�v�ʒu

    '���i�d�l
    HSXTYPE             As String * 1           ' �i�r�w�^�C�v
    HSXD1CEN            As Double               ' �i�r�w���a�P���S
    HSXCDIR             As String * 1           ' �i�r�w�����ʕ���
    HSXRMIN             As Double               ' �i�r�w���R����
    HSXRMAX             As Double               ' �i�r�w���R���
    HSXRAMIN            As Double               ' �i�r�w���R���ω���
    HSXRAMAX            As Double               ' �i�r�w���R���Ϗ��
    HSXRMBNP            As Double               ' �i�r�w���R�ʓ����z
    HSXRSPOH            As String * 1           ' �i�r�w���R����ʒu�Q��
    HSXRSPOT            As String * 1           ' �i�r�w���R����ʒu�Q�_
    HSXRSPOI            As String * 1           ' �i�r�w���R����ʒu�Q��
    HSXRHWYT            As String * 1           ' �i�r�w���R�ۏؕ��@�Q��
    HSXRHWYS            As String * 1           ' �i�r�w���R�ۏؕ��@�Q��

    HSXONMIN            As Double               ' �i�r�w�_�f�Z�x����
    HSXONMAX            As Double               ' �i�r�w�_�f�Z�x���
    HSXONAMN            As Double               ' �i�r�w�_�f�Z�x���ω���
    HSXONAMX            As Double               ' �i�r�w�_�f�Z�x���Ϗ��
    HSXONMBP            As Double               ' �i�r�w�_�f�Z�x�ʓ����z
    HSXONSPH            As String * 1           ' �i�r�w�_�f�Z�x����ʒu�Q��
    HSXONSPT            As String * 1           ' �i�r�w�_�f�Z�x����ʒu�Q�_
    HSXONSPI            As String * 1           ' �i�r�w�_�f�Z�x����ʒu�Q��
    HSXONHWT            As String * 1           ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    HSXONHWS            As String * 1           ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��

    HSXBM1AN            As Double               ' �i�r�w�a�l�c�P���ω���
    HSXBM1AX            As Double               ' �i�r�w�a�l�c�P���Ϗ��
    HSXBM2AN            As Double               ' �i�r�w�a�l�c�Q���ω���
    HSXBM2AX            As Double               ' �i�r�w�a�l�c�Q���Ϗ��
    HSXBM3AN            As Double               ' �i�r�w�a�l�c�R���ω���
    HSXBM3AX            As Double               ' �i�r�w�a�l�c�R���Ϗ��
    HSXBM1SH            As String * 1           ' �i�r�w�a�l�c�P����ʒu�Q��
    HSXBM1ST            As String * 1           ' �i�r�w�a�l�c�P����ʒu�Q�_
    HSXBM1SR            As String * 1           ' �i�r�w�a�l�c�P����ʒu�Q��
    HSXBM1HT            As String * 1           ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    HSXBM1HS            As String * 1           ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    HSXBM2SH            As String * 1           ' �i�r�w�a�l�c�Q����ʒu�Q��
    HSXBM2ST            As String * 1           ' �i�r�w�a�l�c�Q����ʒu�Q�_
    HSXBM2SR            As String * 1           ' �i�r�w�a�l�c�Q����ʒu�Q��
    HSXBM2HT            As String * 1           ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    HSXBM2HS            As String * 1           ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    HSXBM3SH            As String * 1           ' �i�r�w�a�l�c�R����ʒu�Q��
    HSXBM3ST            As String * 1           ' �i�r�w�a�l�c�R����ʒu�Q�_
    HSXBM3SR            As String * 1           ' �i�r�w�a�l�c�R����ʒu�Q��
    HSXBM3HT            As String * 1           ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
    HSXBM3HS            As String * 1           ' �i�r�w�a�l�c�R�ۏؕ��@�Q��

    HSXOS1AX            As Double               ' �i�r�w�n�r�e�P���Ϗ��
    HSXOS1MX            As Double               ' �i�r�w�n�r�e�P���
    HSXOS2AX            As Double               ' �i�r�w�n�r�e�Q���Ϗ��
    HSXOS2MX            As Double               ' �i�r�w�n�r�e�Q���
    HSXOS3AX            As Double               ' �i�r�w�n�r�e�R���Ϗ��
    HSXOS3MX            As Double               ' �i�r�w�n�r�e�R���
    HSXOS4AX            As Double               ' �i�r�w�n�r�e�S���Ϗ��
    HSXOS4MX            As Double               ' �i�r�w�n�r�e�S���
    HSXOS1SH            As String * 1           ' �i�r�w�n�r�e�P����ʒu�Q��
    HSXOS1ST            As String * 1           ' �i�r�w�n�r�e�P����ʒu�Q�_
    HSXOS1SR            As String * 1           ' �i�r�w�n�r�e�P����ʒu�Q��
    HSXOS1HT            As String * 1           ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    HSXOS1HS            As String * 1           ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    HSXOS2SH            As String * 1           ' �i�r�w�n�r�e�Q����ʒu�Q��
    HSXOS2ST            As String * 1           ' �i�r�w�n�r�e�Q����ʒu�Q�_
    HSXOS2SR            As String * 1           ' �i�r�w�n�r�e�Q����ʒu�Q��
    HSXOS2HT            As String * 1           ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    HSXOS2HS            As String * 1           ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    HSXOS3SH            As String * 1           ' �i�r�w�n�r�e�R����ʒu�Q��
    HSXOS3ST            As String * 1           ' �i�r�w�n�r�e�R����ʒu�Q�_
    HSXOS3SR            As String * 1           ' �i�r�w�n�r�e�R����ʒu�Q��
    HSXOS3HT            As String * 1           ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    HSXOS3HS            As String * 1           ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    HSXOS4SH            As String * 1           ' �i�r�w�n�r�e�S����ʒu�Q��
    HSXOS4ST            As String * 1           ' �i�r�w�n�r�e�S����ʒu�Q�_
    HSXOS4SR            As String * 1           ' �i�r�w�n�r�e�S����ʒu�Q��
    HSXOS4HT            As String * 1           ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    HSXOS4HS            As String * 1           ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    HSXOS1NS            As String * 2           ' �i�r�w�n�r�e�P�M�����@
    HSXOS2NS            As String * 2           ' �i�r�w�n�r�e�Q�M�����@
    HSXOS3NS            As String * 2           ' �i�r�w�n�r�e�R�M�����@
    HSXOS4NS            As String * 2           ' �i�r�w�n�r�e�S�M�����@
    HSXBM1NS            As String * 2           ' �i�r�w�a�l�c�P�M�����@
    HSXBM2NS            As String * 2           ' �i�r�w�a�l�c�Q�M�����@
    HSXBM3NS            As String * 2           ' �i�r�w�a�l�c�R�M�����@

    HSXCNMIN            As Double               ' �i�r�w�Y�f�Z�x����
    HSXCNMAX            As Double               ' �i�r�w�Y�f�Z�x���
    HSXCNSPH            As String * 1           ' �i�r�w�Y�f�Z�x����ʒu�Q��
    HSXCNSPT            As String * 1           ' �i�r�w�Y�f�Z�x����ʒu�Q�_
    HSXCNSPI            As String * 1           ' �i�r�w�Y�f�Z�x����ʒu�Q��
    HSXCNHWT            As String * 1           ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    HSXCNHWS            As String * 1           ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��

    HSXDENMX            As Integer              ' �i�r�w�c�������
    HSXDENMN            As Integer              ' �i�r�w�c��������
    HSXLDLMX            As Integer              ' �i�r�w�k�^�c�k���
    HSXLDLMN            As Integer              ' �i�r�w�k�^�c�k����
    HSXDVDMX            As Integer              ' �i�r�w�c�u�c�Q���
    HSXDVDMN            As Integer              ' �i�r�w�c�u�c�Q����
    HSXDENHT            As String * 1           ' �i�r�w�c�����ۏؕ��@�Q��
    HSXDENHS            As String * 1           ' �i�r�w�c�����ۏؕ��@�Q��
    HSXLDLHT            As String * 1           ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    HSXLDLHS            As String * 1           ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    HSXDVDHT            As String * 1           ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    HSXDVDHS            As String * 1           ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    HSXDENKU            As String * 1           ' �i�r�w�c���������L��
    HSXDVDKU            As String * 1           ' �i�r�w�c�u�c�Q�����L��
    HSXLDLKU            As String * 1           ' �i�r�w�k�^�c�k�����L��

    HSXLTMIN            As Integer              ' �i�r�w�k�^�C������
    HSXLTMAX            As Integer              ' �i�r�w�k�^�C�����
    HSXLTSPH            As String * 1           ' �i�r�w�k�^�C������ʒu�Q��
    HSXLTSPT            As String * 1           ' �i�r�w�k�^�C������ʒu�Q�_
    HSXLTSPI            As String * 1           ' �i�r�w�k�^�C������ʒu�Q��
    HSXLTHWT            As String * 1           ' �i�r�w�k�^�C���ۏؕ��@�Q��
    HSXLTHWS            As String * 1           ' �i�r�w�k�^�C���ۏؕ��@�Q��
    '���������Ǘ�
    EPDUP               As Integer              ' EPD�@���
End Type


' �����T���v���Ǘ��擾�p (TOP,TAIL���łQ���R�[�h�擾)
Public Type type_DBDRV_scmzc_fcmkc001c_CrySmp
    CRYNUM              As String * 12          ' �����ԍ�
    INGOTPOS            As Integer              ' �������ʒu
    LENGTH              As Integer              ' ����
    BLOCKID             As String * 12          ' �u���b�NID
    SMPKBN              As String * 1           ' �T���v���敪
    SMPLNO              As Integer              ' �T���v��No
    hinban              As String * 12          ' �i��
    REVNUM              As Integer              ' ���i�ԍ������ԍ�
    factory             As String * 1           ' �H��
    opecond             As String * 1           ' ���Ə���
    KTKBN               As String * 1           ' �m��敪
    CRYINDRS            As String * 1           ' ���������w���iRs)
    CRYINDOI            As String * 1           ' ���������w���iOi)
    CRYINDB1            As String * 1           ' ���������w���iB1)
    CRYINDB2            As String * 1           ' ���������w���iB2�j
    CRYINDB3            As String * 1           ' ���������w���iB3)
    CRYINDL1            As String * 1           ' ���������w���iL1)
    CRYINDL2            As String * 1           ' ���������w���iL2)
    CRYINDL3            As String * 1           ' ���������w���iL3)
    CRYINDL4            As String * 1           ' ���������w���iL4)
    CRYINDCS            As String * 1           ' ���������w���iCs)
    CRYINDGD            As String * 1           ' ���������w���iGD)
    CRYINDT             As String * 1           ' ���������w���iT)
    CRYINDEP            As String * 1           ' ���������w���iEPD)
End Type


'������R����
Public Type type_DBDRV_scmzc_fcmkc001c_CryR
    CRYNUM              As String * 12          ' �����ԍ�
    POSITION            As Integer              ' �ʒu
    SMPKBN              As String * 1           ' �T���v���敪
    SMPLNO              As Integer              ' �T���v���m��
    SMPLUMU             As String * 1           ' �T���v���L��
    TRANCOND            As String * 1           ' ��������
    MEAS1               As Double               ' ����l�P
    MEAS2               As Double               ' ����l�Q
    MEAS3               As Double               ' ����l�R
    MEAS4               As Double               ' ����l�S
    MEAS5               As Double               ' ����l�T
    RRG                 As Double               ' �q�q�f
    REGDATE             As Date                 ' �o�^���t
End Type


'Oi����
Public Type type_DBDRV_scmzc_fcmkc001c_Oi
    CRYNUM              As String * 12          ' �����ԍ�
    POSITION            As Integer              ' �ʒu
    SMPKBN              As String * 1           ' �T���v���敪
    SMPLNO              As Integer              ' �T���v���m��
    SMPLUMU             As String * 1           ' �T���v���L��
    TRANCOND            As String * 1           ' ��������
    OIMEAS1             As Double               ' �n������l�P
    OIMEAS2             As Double               ' �n������l�Q
    OIMEAS3             As Double               ' �n������l�R
    OIMEAS4             As Double               ' �n������l�S
    OIMEAS5             As Double               ' �n������l�T
    ORGRES              As Double               ' �n�q�f����
    AVE                 As Double               ' �`�u�d
    FTIRCONV            As Double               ' �e�s�h�q���Z
    INSPECTWAY          As String * 2           ' �������@
    REGDATE             As Date                 ' �o�^���t
End Type


'BMD1�`3����
Public Type type_DBDRV_scmzc_fcmkc001c_BMD
    CRYNUM              As String * 12          ' �����ԍ�
    POSITION            As Integer              ' �ʒu
    SMPKBN              As String * 1           ' �T���v���敪
    SMPLNO              As Integer              ' �T���v���m��
    SMPLUMU             As String * 1           ' �T���v���L��
    HTPRC               As String * 2           ' �M�������@
    KKSP                As String * 3           ' �������ב���ʒu
    KKSET               As String * 3           ' �������ב�������{�I��ET��@�@char(1)�{number(2)
    TRANCOND            As String * 1           ' ��������
    MEAS1               As Double               ' ����l�P
    MEAS2               As Double               ' ����l�Q
    MEAS3               As Double               ' ����l�R
    MEAS4               As Double               ' ����l�S
    MEAS5               As Double               ' ����l�T
    Min                 As Double               ' MIN
    max                 As Double               ' MAX
    AVE                 As Double               ' AVE
    REGDATE             As Date                 ' �o�^���t
End Type


'OSF1�`4����
Public Type type_DBDRV_scmzc_fcmkc001c_OSF
    CRYNUM              As String * 12          ' �����ԍ�
    POSITION            As Integer              ' �ʒu
    SMPKBN              As String * 1           ' �T���v���敪
    SMPLNO              As Integer              ' �T���v���m��
    SMPLUMU             As String * 1           ' �T���v���L��
    HTPRC               As String * 2           ' �M�������@
    KKSP                As String * 3           ' �������ב���ʒu
    KKSET               As String * 3           ' �������ב�������{�I��ET��@�@char(1)�{number(2)
    TRANCOND            As String * 1           ' ��������
    CALCMAX             As Double               ' �v�Z���� Max
    CALCAVE             As Double               ' �v�Z���� Ave
    MEAS1               As Double               ' ����l�P
    MEAS2               As Double               ' ����l�Q
    MEAS3               As Double               ' ����l�R
    MEAS4               As Double               ' ����l�S
    MEAS5               As Double               ' ����l�T
    MEAS6               As Double               ' ����l�U
    MEAS7               As Double               ' ����l�V
    MEAS8               As Double               ' ����l�W
    MEAS9               As Double               ' ����l�X
    MEAS10              As Double               ' ����l�P�O
    MEAS11              As Double               ' ����l�P�P
    MEAS12              As Double               ' ����l�P�Q
    MEAS13              As Double               ' ����l�P�R
    MEAS14              As Double               ' ����l�P�S
    MEAS15              As Double               ' ����l�P�T
    MEAS16              As Double               ' ����l�P�U
    MEAS17              As Double               ' ����l�P�V
    MEAS18              As Double               ' ����l�P�W
    MEAS19              As Double               ' ����l�P�X
    MEAS20              As Double               ' ����l�Q�O
    REGDATE             As Date                 ' �o�^���t
End Type


'CS����
Public Type type_DBDRV_scmzc_fcmkc001c_CS
    CRYNUM              As String * 12          ' �����ԍ�
    POSITION            As Integer              ' �ʒu
    SMPKBN              As String * 1           ' �T���v���敪
    SMPLNO              As Integer              ' �T���v���m��
    SMPLUMU             As String * 1           ' �T���v���L��
    TRANCOND            As String * 1           ' ��������
    CSMEAS              As Double               ' Cs�����l
    PRE70P              As Double               ' �V�O������l
    REGDATE             As Date                 ' �o�^���t
End Type


'GD����
Public Type type_DBDRV_scmzc_fcmkc001c_GD
    CRYNUM              As String * 12          ' �����ԍ�
    POSITION            As Integer              ' �ʒu
    SMPKBN              As String * 1           ' �T���v���敪
    SMPLNO              As Integer              ' �T���v���m��
    SMPLUMU             As String * 1           ' �T���v���L��
    TRANCOND            As String * 1           ' ��������
    MSRSDEN             As Integer              ' ���茋�� Den
    MSRSLDL             As Integer              ' ���茋�� L/DL
    MSRSDVD2            As Integer              ' ���茋�� DVD2
    MS01LDL1            As Integer              ' ����l01 L/DL1
    MS01LDL2            As Integer              ' ����l01 L/DL2
    MS01LDL3            As Integer              ' ����l01 L/DL3
    MS01LDL4            As Integer              ' ����l01 L/DL4
    MS01LDL5            As Integer              ' ����l01 L/DL5
    MS01DEN1            As Integer              ' ����l01 Den1
    MS01DEN2            As Integer              ' ����l01 Den2
    MS01DEN3            As Integer              ' ����l01 Den3
    MS01DEN4            As Integer              ' ����l01 Den4
    MS01DEN5            As Integer              ' ����l01 Den5
    MS02LDL1            As Integer              ' ����l02 L/DL1
    MS02LDL2            As Integer              ' ����l02 L/DL2
    MS02LDL3            As Integer              ' ����l02 L/DL3
    MS02LDL4            As Integer              ' ����l02 L/DL4
    MS02LDL5            As Integer              ' ����l02 L/DL5
    MS02DEN1            As Integer              ' ����l02 Den1
    MS02DEN2            As Integer              ' ����l02 Den2
    MS02DEN3            As Integer              ' ����l02 Den3
    MS02DEN4            As Integer              ' ����l02 Den4
    MS02DEN5            As Integer              ' ����l02 Den5
    MS03LDL1            As Integer              ' ����l03 L/DL1
    MS03LDL2            As Integer              ' ����l03 L/DL2
    MS03LDL3            As Integer              ' ����l03 L/DL3
    MS03LDL4            As Integer              ' ����l03 L/DL4
    MS03LDL5            As Integer              ' ����l03 L/DL5
    MS03DEN1            As Integer              ' ����l03 Den1
    MS03DEN2            As Integer              ' ����l03 Den2
    MS03DEN3            As Integer              ' ����l03 Den3
    MS03DEN4            As Integer              ' ����l03 Den4
    MS03DEN5            As Integer              ' ����l03 Den5
    MS04LDL1            As Integer              ' ����l04 L/DL1
    MS04LDL2            As Integer              ' ����l04 L/DL2
    MS04LDL3            As Integer              ' ����l04 L/DL3
    MS04LDL4            As Integer              ' ����l04 L/DL4
    MS04LDL5            As Integer              ' ����l04 L/DL5
    MS04DEN1            As Integer              ' ����l04 Den1
    MS04DEN2            As Integer              ' ����l04 Den2
    MS04DEN3            As Integer              ' ����l04 Den3
    MS04DEN4            As Integer              ' ����l04 Den4
    MS04DEN5            As Integer              ' ����l04 Den5
    MS05LDL1            As Integer              ' ����l05 L/DL1
    MS05LDL2            As Integer              ' ����l05 L/DL2
    MS05LDL3            As Integer              ' ����l05 L/DL3
    MS05LDL4            As Integer              ' ����l05 L/DL4
    MS05LDL5            As Integer              ' ����l05 L/DL5
    MS05DEN1            As Integer              ' ����l05 Den1
    MS05DEN2            As Integer              ' ����l05 Den2
    MS05DEN3            As Integer              ' ����l05 Den3
    MS05DEN4            As Integer              ' ����l05 Den4
    MS05DEN5            As Integer              ' ����l05 Den5
    MS06LDL1            As Integer              ' ����l06 L/DL1
    MS06LDL2            As Integer              ' ����l06 L/DL2
    MS06LDL3            As Integer              ' ����l06 L/DL3
    MS06LDL4            As Integer              ' ����l06 L/DL4
    MS06LDL5            As Integer              ' ����l06 L/DL5
    MS06DEN1            As Integer              ' ����l06 Den1
    MS06DEN2            As Integer              ' ����l06 Den2
    MS06DEN3            As Integer              ' ����l06 Den3
    MS06DEN4            As Integer              ' ����l06 Den4
    MS06DEN5            As Integer              ' ����l06 Den5
    MS07LDL1            As Integer              ' ����l07 L/DL1
    MS07LDL2            As Integer              ' ����l07 L/DL2
    MS07LDL3            As Integer              ' ����l07 L/DL3
    MS07LDL4            As Integer              ' ����l07 L/DL4
    MS07LDL5            As Integer              ' ����l07 L/DL5
    MS07DEN1            As Integer              ' ����l07 Den1
    MS07DEN2            As Integer              ' ����l07 Den2
    MS07DEN3            As Integer              ' ����l07 Den3
    MS07DEN4            As Integer              ' ����l07 Den4
    MS07DEN5            As Integer              ' ����l07 Den5
    MS08LDL1            As Integer              ' ����l08 L/DL1
    MS08LDL2            As Integer              ' ����l08 L/DL2
    MS08LDL3            As Integer              ' ����l08 L/DL3
    MS08LDL4            As Integer              ' ����l08 L/DL4
    MS08LDL5            As Integer              ' ����l08 L/DL5
    MS08DEN1            As Integer              ' ����l08 Den1
    MS08DEN2            As Integer              ' ����l08 Den2
    MS08DEN3            As Integer              ' ����l08 Den3
    MS08DEN4            As Integer              ' ����l08 Den4
    MS08DEN5            As Integer              ' ����l08 Den5
    MS09LDL1            As Integer              ' ����l09 L/DL1
    MS09LDL2            As Integer              ' ����l09 L/DL2
    MS09LDL3            As Integer              ' ����l09 L/DL3
    MS09LDL4            As Integer              ' ����l09 L/DL4
    MS09LDL5            As Integer              ' ����l09 L/DL5
    MS09DEN1            As Integer              ' ����l09 Den1
    MS09DEN2            As Integer              ' ����l09 Den2
    MS09DEN3            As Integer              ' ����l09 Den3
    MS09DEN4            As Integer              ' ����l09 Den4
    MS09DEN5            As Integer              ' ����l09 Den5
    MS10LDL1            As Integer              ' ����l10 L/DL1
    MS10LDL2            As Integer              ' ����l10 L/DL2
    MS10LDL3            As Integer              ' ����l10 L/DL3
    MS10LDL4            As Integer              ' ����l10 L/DL4
    MS10LDL5            As Integer              ' ����l10 L/DL5
    MS10DEN1            As Integer              ' ����l10 Den1
    MS10DEN2            As Integer              ' ����l10 Den2
    MS10DEN3            As Integer              ' ����l10 Den3
    MS10DEN4            As Integer              ' ����l10 Den4
    MS10DEN5            As Integer              ' ����l10 Den5
    MS11LDL1            As Integer              ' ����l11 L/DL1
    MS11LDL2            As Integer              ' ����l11 L/DL2
    MS11LDL3            As Integer              ' ����l11 L/DL3
    MS11LDL4            As Integer              ' ����l11 L/DL4
    MS11LDL5            As Integer              ' ����l11 L/DL5
    MS11DEN1            As Integer              ' ����l11 Den1
    MS11DEN2            As Integer              ' ����l11 Den2
    MS11DEN3            As Integer              ' ����l11 Den3
    MS11DEN4            As Integer              ' ����l11 Den4
    MS11DEN5            As Integer              ' ����l11 Den5
    MS12LDL1            As Integer              ' ����l12 L/DL1
    MS12LDL2            As Integer              ' ����l12 L/DL2
    MS12LDL3            As Integer              ' ����l12 L/DL3
    MS12LDL4            As Integer              ' ����l12 L/DL4
    MS12LDL5            As Integer              ' ����l12 L/DL5
    MS12DEN1            As Integer              ' ����l12 Den1
    MS12DEN2            As Integer              ' ����l12 Den2
    MS12DEN3            As Integer              ' ����l12 Den3
    MS12DEN4            As Integer              ' ����l12 Den4
    MS12DEN5            As Integer              ' ����l12 Den5
    MS13LDL1            As Integer              ' ����l13 L/DL1
    MS13LDL2            As Integer              ' ����l13 L/DL2
    MS13LDL3            As Integer              ' ����l13 L/DL3
    MS13LDL4            As Integer              ' ����l13 L/DL4
    MS13LDL5            As Integer              ' ����l13 L/DL5
    MS13DEN1            As Integer              ' ����l13 Den1
    MS13DEN2            As Integer              ' ����l13 Den2
    MS13DEN3            As Integer              ' ����l13 Den3
    MS13DEN4            As Integer              ' ����l13 Den4
    MS13DEN5            As Integer              ' ����l13 Den5
    MS14LDL1            As Integer              ' ����l14 L/DL1
    MS14LDL2            As Integer              ' ����l14 L/DL2
    MS14LDL3            As Integer              ' ����l14 L/DL3
    MS14LDL4            As Integer              ' ����l14 L/DL4
    MS14LDL5            As Integer              ' ����l14 L/DL5
    MS14DEN1            As Integer              ' ����l14 Den1
    MS14DEN2            As Integer              ' ����l14 Den2
    MS14DEN3            As Integer              ' ����l14 Den3
    MS14DEN4            As Integer              ' ����l14 Den4
    MS14DEN5            As Integer              ' ����l14 Den5
    MS15LDL1            As Integer              ' ����l15 L/DL1
    MS15LDL2            As Integer              ' ����l15 L/DL2
    MS15LDL3            As Integer              ' ����l15 L/DL3
    MS15LDL4            As Integer              ' ����l15 L/DL4
    MS15LDL5            As Integer              ' ����l15 L/DL5
    MS15DEN1            As Integer              ' ����l15 Den1
    MS15DEN2            As Integer              ' ����l15 Den2
    MS15DEN3            As Integer              ' ����l15 Den3
    MS15DEN4            As Integer              ' ����l15 Den4
    MS15DEN5            As Integer              ' ����l15 Den5
    REGDATE             As Date                 ' �o�^���t
End Type


'���C�t�^�C�����ю擾�֐�
Public Type type_DBDRV_scmzc_fcmkc001c_LT
    CRYNUM              As String * 12          ' �����ԍ�
    POSITION            As Integer              ' �ʒu
    SMPKBN              As String * 1           ' �T���v���敪
    SMPLNO              As Integer              ' �T���v���m��
    SMPLUMU             As String * 1           ' �T���v���L��
    MEAS1               As Integer              ' ����l�P
    MEAS2               As Integer              ' ����l�Q
    MEAS3               As Integer              ' ����l�R
    MEAS4               As Integer              ' ����l�S
    MEAS5               As Integer              ' ����l�T
    TRANCOND            As String * 1           ' ��������
    MEASPEAK            As Integer              ' ����l �s�[�N�l
    CALCMEAS            As Integer              ' �v�Z����
    REGDATE             As Date                 ' �o�^���t
    LTSPI               As String               ' ����ʒu�R�[�h
End Type


'EPD���ю擾�֐�
Public Type type_DBDRV_scmzc_fcmkc001c_EPD
    CRYNUM              As String * 12          ' �����ԍ�
    POSITION            As Integer              ' �ʒu
    SMPKBN              As String * 1           ' �T���v���敪
    SMPLNO              As Integer              ' �T���v���m��
    SMPLUMU             As String * 1           ' �T���v���L��
    TRANCOND            As String * 1           ' ��������
    MEASURE             As Integer              ' ����l
    REGDATE             As Date                 ' �o�^���t
End Type


'���т��܂Ƃ߂��\����
Public Type type_DBDRV_scmzc_fcmkc001c_Zisseki
    CRYRZ()             As type_DBDRV_scmzc_fcmkc001c_CryR
    OIZ()               As type_DBDRV_scmzc_fcmkc001c_Oi
    BMD1Z()             As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD2Z()             As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD3Z()             As type_DBDRV_scmzc_fcmkc001c_BMD
    OSF1Z()             As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF2Z()             As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF3Z()             As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF4Z()             As type_DBDRV_scmzc_fcmkc001c_OSF
    csz()               As type_DBDRV_scmzc_fcmkc001c_CS
    GDZ()               As type_DBDRV_scmzc_fcmkc001c_GD
    LTZ()               As type_DBDRV_scmzc_fcmkc001c_LT
    EPDZ()              As type_DBDRV_scmzc_fcmkc001c_EPD
    SURSZ()             As type_DBDRV_scmzc_fcmkc001c_CryR
End Type

'*** UPDATE START T.TERAUCHI 2004/12/19 CX�쐬�p�\���̐ݒ�
Public Type type_DBDRV_fcmkc001c_InsXodcx
    BLOCKID     As String               ' �u���b�NID
    CRYNUM      As String * 12          ' �����ԍ�
    INGOTPOS    As Integer              ' �������J�n�ʒu
    LASTPASS    As String * 5           ' �ŏI�ʉߍH��
    LENGTH      As Integer              ' ����
    Weight      As Double               ' �d��
    STAFFID     As String               ' �S���҃R�[�h
    STAFFNAME   As String               ' �S���Җ�
End Type
'*** UPDATE END   T.TERAUCHI 2004/12/19

' 2007/08/30 SPK Tsutsumi Add Start
Public Type typ_Mukesaki
    sMukeCode As String     '' ����R�[�h
    sMukeName As String     '' ���於
End Type

Public s_CmbMukesaki() As typ_Mukesaki
Public s_Mukesaki() As typ_Mukesaki
' 2007/08/30 SPK Tsutsumi Add End
'2008/05/30 SHINDOH----------------------------------------------
' �z�񏉊����l
Public Const DEF_PARAM_VALUE_LT = -1
' ���C�t�^�C������_���i�V�f�[�^�͂P�O�_�Œ�j
Public Const SS_SOKUETI_TENSU = 10
' ���C�t�^�C������_���i���f�[�^�͂T�_�Œ�j
Public Const SS_SOKUETI_TENSU_OLD = 5
'2008/05/30 SHINDOH----------------------------------------------


Public Type typ_TBCMX011
BLOCKID As String
FROMTOKBN As String
TRANCNT As Integer
STCID As String
hinban As String
REVNUM As String
factory As String
opecond As String
STCKNNUM As String
CRYNUM As String
CRYDECDATE As Date
PLUPDATE As Date
UPLENGTH As Integer
FREELENG As Integer
INGOTPOS As Integer
BlkLen As Integer
BLKWGHT As Long
LENGTH As Integer
Weight As Long
MCNO As String
PGID As String
DM1 As Integer
DM2 As Integer
NCHDPTH As Integer
CHARGE As Double
SEED As String
SXL_RS_SMPPOS As Integer
SXLRS_MEAS1 As Double
SXLRS_MEAS2 As Double
SXLRS_MEAS3 As Double
SXLRS_MEAS4 As Double
SXLRS_MEAS5 As Double
SXLRS_EFEHS As Double
SXLRS_RRG As Double
SXL_OI_SMPPOS As Integer
SXLOI_OIMEAS1 As Double
SXLOI_OIMEAS2 As Double
SXLOI_OIMEAS3 As Double
SXLOI_OIMEAS4 As Double
SXLOI_OIMEAS5 As Double
SXLOI_ORGRES As Double
SXLOI_INSPECTWAY As String
SXL_CS_SMPPOS As Integer
SXLCS_CSMEAS As Double
SXLCS_70PPRE As Double
SXLCS_BSUIME As Double
SXLOSF_SMPPOS As Integer
SXLOSF1_KKSP As String
SXLOSF1_NETU As String
SXLOSF1_KKSET As String
SXLOSF1_CALCMAX As Double
SXLOSF1_CALCAVE As Double
SXLOSF2_KKSP As String
SXLOSF2_NETU As String
SXLOSF2_KKSET As String
SXLOSF2_CALCMAX As Double
SXLOSF2_CALCAVE As Double
SXLOSF3_KKSP As String
SXLOSF3_NETU As String
SXLOSF3_KKSET As String
SXLOSF3_CALCMAX As Double
SXLOSF3_CALCAVE As Double
SXLOSF4_KKSP As String
SXLOSF4_NETU As String
SXLOSF4_KKSET As String
SXLOSF4_CALCMAX As Double
SXLOSF4_CALCAVE As Double
SXLBMD_SMPPOS As Integer
SXLBMD1_KKSP As String
SXLBMD1_NETU As String
SXLBMD1_KKSET As String
SXLBMD1_CALCMAX As Double
SXLBMD1_CALCAVE As Double
SXLBMD1_CALCMIN As Double
SXLBMD1_CALCMB As Double
SXLBMD2_KKSP As String
SXLBMD2_NETU As String
SXLBMD2_KKSET As String
SXLBMD2_CALCMAX As Double
SXLBMD2_CALCAVE As Double
SXLBMD2_CALCMIN As Double
SXLBMD2_CALCMB As Double
SXLBMD3_KKSP As String
SXLBMD3_NETU As String
SXLBMD3_KKSET As String
SXLBMD3_CALCMAX As Double
SXLBMD3_CALCAVE As Double
SXLBMD3_CALCMIN As Double
SXLBMD3_CALCMB As Double
SXLGD_SMPPOS As Integer
SXLGD_MSRSDEN As Integer
SXLGD_MSRSLDL As Integer
SXLGD_MSRSDVD2 As Integer
SXLLT_SMPPOS As Integer
SXLLT_MEASPEAK As Integer
SXLLT_CALCMEAS As Integer
REGDATE As Date
SENDFLAG As String
SENDDATE As Date
SNDKDWH As String
SDAYDWH As Date
SNDKSPC As String
SDAYSPC As Date
End Type


Public Type typ_TBCMX012

BLOCKID As String
FROMTOKBN As String
STCID As String
hinban As String
REVNUM As String
factory As String
opecond As String
STCKNNUM As String
CRYNUM As String
SXLOSF1_SMPPOS As Integer
SXLOSF1_KKSP As String
SXLOSF1_NETU As String
SXLOSF1_KKSET As String
SXLOSF1_MEAS1 As Double
SXLOSF1_MEAS2 As Double
SXLOSF1_MEAS3 As Double
SXLOSF1_MEAS4 As Double
SXLOSF1_MEAS5 As Double
SXLOSF1_MEAS6 As Double
SXLOSF1_MEAS7 As Double
SXLOSF1_MEAS8 As Double
SXLOSF1_MEAS9 As Double
SXLOSF1_MEAS10 As Double
SXLOSF1_MEAS11 As Double
SXLOSF1_MEAS12 As Double
SXLOSF1_MEAS13 As Double
SXLOSF1_MEAS14 As Double
SXLOSF1_MEAS15 As Double
SXLOSF1_MEAS16 As Double
SXLOSF1_MEAS17 As Double
SXLOSF1_MEAS18 As Double
SXLOSF1_MEAS19 As Double
SXLOSF1_MEAS20 As Double
SXLOSF2_KKSP As String
SXLOSF2_NETU As String
SXLOSF2_KKSET As String
SXLOSF2_MEAS1 As Double
SXLOSF2_MEAS2 As Double
SXLOSF2_MEAS3 As Double
SXLOSF2_MEAS4 As Double
SXLOSF2_MEAS5 As Double
SXLOSF2_MEAS6 As Double
SXLOSF2_MEAS7 As Double
SXLOSF2_MEAS8 As Double
SXLOSF2_MEAS9 As Double
SXLOSF2_MEAS10 As Double
SXLOSF2_MEAS11 As Double
SXLOSF2_MEAS12 As Double
SXLOSF2_MEAS13 As Double
SXLOSF2_MEAS14 As Double
SXLOSF2_MEAS15 As Double
SXLOSF2_MEAS16 As Double
SXLOSF2_MEAS17 As Double
SXLOSF2_MEAS18 As Double
SXLOSF2_MEAS19 As Double
SXLOSF2_MEAS20 As Double
SXLOSF3_KKSP As String
SXLOSF3_NETU As String
SXLOSF3_KKSET As String
SXLOSF3_MEAS1 As Double
SXLOSF3_MEAS2 As Double
SXLOSF3_MEAS3 As Double
SXLOSF3_MEAS4 As Double
SXLOSF3_MEAS5 As Double
SXLOSF3_MEAS6 As Double
SXLOSF3_MEAS7 As Double
SXLOSF3_MEAS8 As Double
SXLOSF3_MEAS9 As Double
SXLOSF3_MEAS10 As Double
SXLOSF3_MEAS11 As Double
SXLOSF3_MEAS12 As Double
SXLOSF3_MEAS13 As Double
SXLOSF3_MEAS14 As Double
SXLOSF3_MEAS15 As Double
SXLOSF3_MEAS16 As Double
SXLOSF3_MEAS17 As Double
SXLOSF3_MEAS18 As Double
SXLOSF3_MEAS19 As Double
SXLOSF3_MEAS20 As Double
SXLOSF4_KKSP As String
SXLOSF4_NETU As String
SXLOSF4_KKSET As String
SXLOSF4_MEAS1  As Double
SXLOSF4_MEAS2 As Double
SXLOSF4_MEAS3 As Double
SXLOSF4_MEAS4 As Double
SXLOSF4_MEAS5 As Double
SXLOSF4_MEAS6 As Double
SXLOSF4_MEAS7 As Double
SXLOSF4_MEAS8 As Double
SXLOSF4_MEAS9 As Double
SXLOSF4_MEAS10 As Double
SXLOSF4_MEAS11 As Double
SXLOSF4_MEAS12 As Double
SXLOSF4_MEAS13 As Double
SXLOSF4_MEAS14 As Double
SXLOSF4_MEAS15 As Double
SXLOSF4_MEAS16 As Double
SXLOSF4_MEAS17 As Double
SXLOSF4_MEAS18 As Double
SXLOSF4_MEAS19 As Double
SXLOSF4_MEAS20 As Double
SXLBMD_SMPPOS As Integer
SXLBMD1_KKSP As String
SXLBMD1_NETU As String
SXLBMD1_KKSET As String
SXLBMD1_MEAS1 As Double
SXLBMD1_MEAS2 As Double
SXLBMD1_MEAS3 As Double
SXLBMD1_MEAS4 As Double
SXLBMD1_MEAS5 As Double
SXLBMD2_KKSP As String
SXLBMD2_NETU As String
SXLBMD2_KKSET As String
SXLBMD2_MEAS1 As Double
SXLBMD2_MEAS2 As Double
SXLBMD2_MEAS3 As Double
SXLBMD2_MEAS4 As Double
SXLBMD2_MEAS5 As Double
SXLBMD3_KKSP As String
SXLBMD3_NETU As String
SXLBMD3_KKSET As String
SXLBMD3_MEAS1 As Double
SXLBMD3_MEAS2 As Double
SXLBMD3_MEAS3 As Double
SXLBMD3_MEAS4 As Double
SXLBMD3_MEAS5 As Double
SXLT_SMPPOS As Integer
SXLLT_MEASPEAK As Integer
SXLLT_MEAS1 As Integer
SXLLT_MEAS2 As Integer
SXLLT_MEAS3 As Integer
SXLLT_MEAS4 As Integer
SXLLT_MEAS5 As Integer
REGDATE As Date
SENDFLAG As String
SENDDATE As Date
SNDKDWH As String
SDAYDWH As Date
SNDKSPC As String
SDAYSPC As Date
End Type

Public Type typ_TBCMX013
BLOCKID As String
FROMTOKBN As String
STCID As String
hinban As String
REVNUM As String
factory As String
opecond As String
STCKNNUM As String
CRYNUM As String
SXLGD_SMPPOS As Integer
SXLGD_MS01DEN1 As Integer
SXLGD_MS02DEN1 As Integer
SXLGD_MS03DEN1 As Integer
SXLGD_MS04DEN1 As Integer
SXLGD_MS05DEN1 As Integer
SXLGD_MS06DEN1 As Integer
SXLGD_MS07DEN1 As Integer
SXLGD_MS08DEN1 As Integer
SXLGD_MS09DEN1 As Integer
SXLGD_MS10DEN1 As Integer
SXLGD_MS11DEN1 As Integer
SXLGD_MS12DEN1 As Integer
SXLGD_MS13DEN1 As Integer
SXLGD_MS14DEN1 As Integer
SXLGD_MS15DEN1 As Integer
SXLGD_MS01DEN2 As Integer
SXLGD_MS02DEN2 As Integer
SXLGD_MS03DEN2 As Integer
SXLGD_MS04DEN2 As Integer
SXLGD_MS05DEN2 As Integer
SXLGD_MS06DEN2 As Integer
SXLGD_MS07DEN2 As Integer
SXLGD_MS08DEN2 As Integer
SXLGD_MS09DEN2 As Integer
SXLGD_MS10DEN2 As Integer
SXLGD_MS11DEN2 As Integer
SXLGD_MS12DEN2 As Integer
SXLGD_MS13DEN2 As Integer
SXLGD_MS14DEN2 As Integer
SXLGD_MS15DEN2 As Integer
SXLGD_MS01DEN3 As Integer
SXLGD_MS02DEN3 As Integer
SXLGD_MS03DEN3 As Integer
SXLGD_MS04DEN3 As Integer
SXLGD_MS05DEN3 As Integer
SXLGD_MS06DEN3 As Integer
SXLGD_MS07DEN3 As Integer
SXLGD_MS08DEN3 As Integer
SXLGD_MS09DEN3 As Integer
SXLGD_MS10DEN3 As Integer
SXLGD_MS11DEN3 As Integer
SXLGD_MS12DEN3 As Integer
SXLGD_MS13DEN3 As Integer
SXLGD_MS14DEN3 As Integer
SXLGD_MS15DEN3 As Integer
SXLGD_MS01DEN4 As Integer
SXLGD_MS02DEN4 As Integer
SXLGD_MS03DEN4 As Integer
SXLGD_MS04DEN4 As Integer
SXLGD_MS05DEN4 As Integer
SXLGD_MS06DEN4 As Integer
SXLGD_MS07DEN4 As Integer
SXLGD_MS08DEN4 As Integer
SXLGD_MS09DEN4 As Integer
SXLGD_MS10DEN4 As Integer
SXLGD_MS11DEN4 As Integer
SXLGD_MS12DEN4 As Integer
SXLGD_MS13DEN4 As Integer
SXLGD_MS14DEN4 As Integer
SXLGD_MS15DEN4 As Integer
SXLGD_MS01DEN5 As Integer
SXLGD_MS02DEN5 As Integer
SXLGD_MS03DEN5 As Integer
SXLGD_MS04DEN5 As Integer
SXLGD_MS05DEN5 As Integer
SXLGD_MS06DEN5 As Integer
SXLGD_MS07DEN5 As Integer
SXLGD_MS08DEN5 As Integer
SXLGD_MS09DEN5 As Integer
SXLGD_MS10DEN5 As Integer
SXLGD_MS11DEN5 As Integer
SXLGD_MS12DEN5 As Integer
SXLGD_MS13DEN5 As Integer
SXLGD_MS14DEN5 As Integer
SXLGD_MS15DEN5 As Integer
SXLGD_MS01LDL1 As Integer
SXLGD_MS02LDL1 As Integer
SXLGD_MS03LDL1 As Integer
SXLGD_MS04LDL1 As Integer
SXLGD_MS05LDL1 As Integer
SXLGD_MS06LDL1 As Integer
SXLGD_MS07LDL1 As Integer
SXLGD_MS08LDL1 As Integer
SXLGD_MS09LDL1 As Integer
SXLGD_MS10LDL1 As Integer
SXLGD_MS11LDL1 As Integer
SXLGD_MS12LDL1 As Integer
SXLGD_MS13LDL1 As Integer
SXLGD_MS14LDL1 As Integer
SXLGD_MS15LDL1 As Integer
SXLGD_MS01LDL2 As Integer
SXLGD_MS02LDL2 As Integer
SXLGD_MS03LDL2 As Integer
SXLGD_MS04LDL2 As Integer
SXLGD_MS05LDL2 As Integer
SXLGD_MS06LDL2 As Integer
SXLGD_MS07LDL2 As Integer
SXLGD_MS08LDL2 As Integer
SXLGD_MS09LDL2 As Integer
SXLGD_MS10LDL2 As Integer
SXLGD_MS11LDL2 As Integer
SXLGD_MS12LDL2 As Integer
SXLGD_MS13LDL2 As Integer
SXLGD_MS14LDL2 As Integer
SXLGD_MS15LDL2 As Integer
SXLGD_MS01LDL3 As Integer
SXLGD_MS02LDL3 As Integer
SXLGD_MS03LDL3 As Integer
SXLGD_MS04LDL3 As Integer
SXLGD_MS05LDL3 As Integer
SXLGD_MS06LDL3 As Integer
SXLGD_MS07LDL3 As Integer
SXLGD_MS08LDL3 As Integer
SXLGD_MS09LDL3 As Integer
SXLGD_MS10LDL3 As Integer
SXLGD_MS11LDL3 As Integer
SXLGD_MS12LDL3 As Integer
SXLGD_MS13LDL3 As Integer
SXLGD_MS14LDL3 As Integer
SXLGD_MS15LDL3 As Integer
SXLGD_MS01LDL4 As Integer
SXLGD_MS02LDL4 As Integer
SXLGD_MS03LDL4 As Integer
SXLGD_MS04LDL4 As Integer
SXLGD_MS05LDL4 As Integer
SXLGD_MS06LDL4 As Integer
SXLGD_MS07LDL4 As Integer
SXLGD_MS08LDL4 As Integer
SXLGD_MS09LDL4 As Integer
SXLGD_MS10LDL4 As Integer
SXLGD_MS11LDL4 As Integer
SXLGD_MS12LDL4 As Integer
SXLGD_MS13LDL4 As Integer
SXLGD_MS14LDL4 As Integer
SXLGD_MS15LDL4 As Integer
SXLGD_MS01LDL5 As Integer
SXLGD_MS02LDL5 As Integer
SXLGD_MS03LDL5 As Integer
SXLGD_MS04LDL5 As Integer
SXLGD_MS05LDL5 As Integer
SXLGD_MS06LDL5 As Integer
SXLGD_MS07LDL5 As Integer
SXLGD_MS08LDL5 As Integer
SXLGD_MS09LDL5 As Integer
SXLGD_MS10LDL5 As Integer
SXLGD_MS11LDL5 As Integer
SXLGD_MS12LDL5 As Integer
SXLGD_MS13LDL5 As Integer
SXLGD_MS14LDL5 As Integer
SXLGD_MS15LDL5 As Integer
SXLGD_MS01DVD21 As Integer
SXLGD_MS01DVD22 As Integer
SXLGD_MS01DVD23 As Integer
SXLGD_MS01DVD24 As Integer
SXLGD_MS01DVD25 As Integer
REGDATE As Date
SENDFLAG As String
SENDDATE As Date
SNDKDWH As String
SDAYDWH As Date
SNDKSPC As String
SDAYSPC As Date
End Type

Public recX011() As typ_TBCMX011
Public recX012() As typ_TBCMX012
Public recX013() As typ_TBCMX013




'�T�v      :�����ŏI���o���� �\���p�c�a�h���C�o
'���Ұ��@�@:�ϐ���       ,IO ,�^                   ,����
'      �@�@:BlockID_in�@ ,I  ,String               ,�u���b�NID
'      �@�@:blkInfo�@�@�@,O  ,typ_cmkc001f_Block   ,�u���b�N���
'      �@�@:records�@�@�@,O  ,typ_cmkc001f_Disp    ,���i�d�l�擾�p
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN      ,�ǂݍ��݂̐���
Public Function DBDRV_fcmkc001f_Disp(BlockID_in As String, blkInfo As typ_cmkc001f_Block, records() As typ_cmkc001f_Disp) As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Integer
    Dim i As Long
    Dim n As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_fcmkc001f_Disp"
    
    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_SUCCESS
    
    ''�u���b�N���𓾂�
    sql = "Select BLK.INGOTPOS, BLK.LENGTH, BLK.REALLEN, BLK.KRPROCCD, BLK.NOWPROC, BLK.LPKRPROCCD, " & _
          "BLK.LASTPASS, BLK.DELCLS, BLK.RSTATCLS, BLK.LSTATCLS, CRY.SEED " & _
          "From TBCME040 BLK, TBCME037 CRY " & _
          "Where (BLOCKID='" & BlockID_in & "') and (BLK.CRYNUM=CRY.CRYNUM)"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If
    With blkInfo
        .INGOTPOS = rs("INGOTPOS")          ' �������J�n�ʒu
        .LENGTH = rs("LENGTH")              ' ����
        .REALLEN = rs("REALLEN")            ' ������
        .KRPROCCD = rs("KRPROCCD")          ' ���݊Ǘ��H��
        .NOWPROC = rs("NOWPROC")            ' ���ݍH��
        .LPKRPROCCD = rs("LPKRPROCCD")      ' �ŏI�ʉߊǗ��H��
        .LASTPASS = rs("LASTPASS")          ' �ŏI�ʉߍH��
        .DELCLS = rs("DELCLS")              ' �폜�敪
        .RSTATCLS = rs("RSTATCLS")          ' ������ԋ敪
        .LSTATCLS = rs("LSTATCLS")          ' �ŏI��ԋ敪
        .SEED = rs("SEED")                  ' SEED
    End With
    rs.Close
    
    
    
    ''���i�d�l�𓾂�
    sql = "select "
    sql = sql & "BH.E041HINBAN, "           ' �i��
    sql = sql & "BH.E041INGOTPOS, "         ' �������J�n�ʒu
    sql = sql & "BH.E041REVNUM, "           ' ���i�ԍ������ԍ�
    sql = sql & "BH.E041FACTORY, "          ' �H��
    sql = sql & "BH.E041OPECOND, "          ' ���Ə���
    sql = sql & "BH.E041LENGTH, "           ' ����
    '���i�d�lSXL�f�[�^
    sql = sql & "S.E018HSXD1CEN, "          ' �i�r�w���a�P���S
    sql = sql & "S.E018HSXRMIN, "           ' �i�r�w���R����
    sql = sql & "S.E018HSXRMAX, "           ' �i�r�w���R���
    sql = sql & "S.E018HSXRMBNP, "          ' �i�r�w���R�ʓ����z
    sql = sql & "S.E018HSXRHWYS, "          ' �i�r�w���R�ۏؕ��@�Q��
    sql = sql & "S.E019HSXONMIN, "          ' �i�r�w�_�f�Z�x����
    sql = sql & "S.E019HSXONMAX, "          ' �i�r�w�_�f�Z�x���
    sql = sql & "S.E019HSXONMBP, "          ' �i�r�w�_�f�Z�x�ʓ����z
    sql = sql & "S.E019HSXONHWS, "          ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    sql = sql & "S.E019HSXCNMIN, "          ' �i�r�w�Y�f�Z�x����
    sql = sql & "S.E019HSXCNMAX, "          ' �i�r�w�Y�f�Z�x���
    sql = sql & "S.E019HSXCNHWS, "          ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    sql = sql & "S.E019HSXTMMAXN, "         ' �i�r�w�]�ʖ��x���        ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sql = sql & "S.E020HSXBM1AN, "          ' �i�r�w�a�l�c�P���ω���
    sql = sql & "S.E020HSXBM1AX, "          ' �i�r�w�a�l�c�P���Ϗ��
    sql = sql & "S.E020HSXBM1HS, "          ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    sql = sql & "S.E020HSXBM2AN, "          ' �i�r�w�a�l�c�Q���ω���
    sql = sql & "S.E020HSXBM2AX, "          ' �i�r�w�a�l�c�Q���Ϗ��
    sql = sql & "S.E020HSXBM2HS, "          ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    sql = sql & "S.E020HSXBM3AN, "          ' �i�r�w�a�l�c�R���ω���
    sql = sql & "S.E020HSXBM3AX, "          ' �i�r�w�a�l�c�R���Ϗ��
    sql = sql & "S.E020HSXBM3HS, "          ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF1AX, "          ' �i�r�w�n�r�e�P���Ϗ��
    sql = sql & "S.E020HSXOF1MX, "          ' �i�r�w�n�r�e�P���
    sql = sql & "S.E020HSXOF1HS, "          ' �i�r�w�n�r�e�P �ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF2AX, "          ' �i�r�w�n�r�e�Q���Ϗ��
    sql = sql & "S.E020HSXOF2MX, "          ' �i�r�w�n�r�e�Q���
    sql = sql & "S.E020HSXOF2HS, "          ' �i�r�w�n�r�e�Q �ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF3AX, "          ' �i�r�w�n�r�e�R���Ϗ��
    sql = sql & "S.E020HSXOF3MX, "          ' �i�r�w�n�r�e�R���
    sql = sql & "S.E020HSXOF3HS, "          ' �i�r�w�n�r�e�R �ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF4AX, "          ' �i�r�w�n�r�e�S���Ϗ��
    sql = sql & "S.E020HSXOF4MX, "          ' �i�r�w�n�r�e�S���
    sql = sql & "S.E020HSXOF4HS, "          ' �i�r�w�n�r�e�S �ۏؕ��@�Q��
    sql = sql & "S.E020HSXDENMX, "          ' �i�r�w�c�������
    sql = sql & "S.E020HSXDENMN, "          ' �i�r�w�c��������
    sql = sql & "S.E020HSXDENHS, "          ' �i�r�w�c�����ۏؕ��@�Q��
    sql = sql & "S.E020HSXDVDMXN, "         ' �i�r�w�c�u�c�Q���       ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sql = sql & "S.E020HSXDVDMNN, "         ' �i�r�w�c�u�c�Q����       ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sql = sql & "S.E020HSXDVDHS, "          ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    sql = sql & "S.E020HSXLDLMX, "          ' �i�r�w�k�^�c�k���
    sql = sql & "S.E020HSXLDLMN, "          ' �i�r�w�k�^�c�k����
    sql = sql & "S.E020HSXLDLHS, "          ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    sql = sql & "S.E019HSXLTMIN, "          ' �i�r�w�k�^�C������
    sql = sql & "S.E019HSXLTMAX, "          ' �i�r�w�k�^�C�����
    sql = sql & "S.E019HSXLTHWS, "          ' �i�r�w�k�^�C���ۏؕ��@�Q��
    sql = sql & "S.E018HSXDPDIR, "          ' �i�r�w�a�ʒu����
    sql = sql & "S.E018HSXDPDRC, "          ' �i�r�w�a�ʒu����
    sql = sql & "S.E018HSXDWMIN, "          ' �i�r�w�a�Љ���
    sql = sql & "S.E018HSXDWMAX, "          ' �i�r�w�a�Џ��
    sql = sql & "S.E018HSXDDMIN, "          ' �i�r�w�a�[����
    sql = sql & "S.E018HSXDDMAX, "          ' �i�r�w�a�[���
    sql = sql & "S.E018HSXD1MIN, "          ' �i�r�w���a�P����
    sql = sql & "S.E018HSXD1MAX, "          ' �i�r�w���a�P���
    sql = sql & "S.E018HSXCTCEN, "          ' �i�r�w�����ʌX�c���S
    sql = sql & "S.E018HSXCYCEN, "          ' �i�r�w�����ʌX�����S
    sql = sql & "U.EPDUP "                  ' ���������Ǘ� EPD�@���
    sql = sql & " from VECME009 BH, VECME001 S, TBCME036 U "
    sql = sql & " where BH.E040BLOCKID='" & BlockID_in & "' "
    sql = sql & " and S.E018HINBAN=BH.E041HINBAN "
    sql = sql & " and S.E018MNOREVNO=BH.E041REVNUM "
    sql = sql & " and S.E018FACTORY=BH.E041FACTORY "
    sql = sql & " and S.E018OPECOND=BH.E041OPECOND "
    sql = sql & " and U.HINBAN=BH.E041HINBAN "
    sql = sql & " and U.MNOREVNO=BH.E041REVNUM "
    sql = sql & " and U.FACTORY=BH.E041FACTORY "
    sql = sql & " and U.OPECOND=BH.E041OPECOND "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        ReDim records(0)
        rs.Close
        GoTo proc_exit
    End If
    
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            '�i�ԊǗ�
            .hinban = rs("E041HINBAN")                          ' �i��
            .INGOTPOS = rs("E041INGOTPOS")                      ' �������J�n�ʒu
            .REVNUM = rs("E041REVNUM")                          ' ���i�ԍ������ԍ�
            .factory = rs("E041FACTORY")                        ' �H��
            .opecond = rs("E041OPECOND")                        ' ���Ə���
            .LENGTH = rs("E041LENGTH")                          ' ����
            '���i�d�lSXL�f�[�^
            .HSXD1CEN = fncNullCheck(rs("E018HSXD1CEN"))                      ' �i�r�w���a�P���S
            .HSXRMIN = fncNullCheck(rs("E018HSXRMIN"))                        ' �i�r�w���R����
            .HSXRMAX = fncNullCheck(rs("E018HSXRMAX"))                        ' �i�r�w���R���
            .HSXRMBNP = fncNullCheck(rs("E018HSXRMBNP"))                      ' �i�r�w���R�ʓ����z
            .HSXRHWYS = rs("E018HSXRHWYS")                      ' �i�r�w���R�ۏؕ��@�Q��
            .HSXONMIN = fncNullCheck(rs("E019HSXONMIN"))                      ' �i�r�w�_�f�Z�x����  'NULL�Ή�
            .HSXONMAX = fncNullCheck(rs("E019HSXONMAX"))                      ' �i�r�w�_�f�Z�x���
            .HSXONMBP = fncNullCheck(rs("E019HSXONMBP"))                      ' �i�r�w�_�f�Z�x�ʓ����z  'NULL�Ή�
            .HSXONHWS = rs("E019HSXONHWS")                      ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
            .HSXCNMIN = fncNullCheck(rs("E019HSXCNMIN"))                      ' �i�r�w�Y�f�Z�x����  'NULL�Ή�
            .HSXCNMAX = fncNullCheck(rs("E019HSXCNMAX"))                      ' �i�r�w�Y�f�Z�x���  'NULL�Ή�
            .HSXCNHWS = rs("E019HSXCNHWS")                      ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
            .HSXTMMAX = rs("E019HSXTMMAXN")                     ' �i�r�w�]�ʖ��x���       ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
            For n = 1 To 3 'NULL�Ή�
                If IsNull(rs("E020HSXBM" & n & "AN")) = False Then
                    .HSXBMnAN(n) = rs("E020HSXBM" & n & "AN") * 10  ' �i�r�w�a�l�cn ���ω���
                Else
                    .HSXBMnAN(n) = -1
                End If
                
                If IsNull(rs("E020HSXBM" & n & "AX")) = False Then
                    .HSXBMnAX(n) = rs("E020HSXBM" & n & "AX") * 10 ' �i�r�w�a�l�cn ���Ϗ��
                Else
                    .HSXBMnAX(n) = -1
                End If
                .HSXBMnHS(n) = rs("E020HSXBM" & n & "HS")       ' �i�r�w�a�l�cn �ۏؕ��@�Q��
            Next
            For n = 1 To 4
                .HSXOFnAX(n) = fncNullCheck(rs("E020HSXOF" & n & "AX"))       ' �i�r�w�n�r�en ���Ϗ��  'NULL�Ή�
                .HSXOFnMX(n) = fncNullCheck(rs("E020HSXOF" & n & "MX"))       ' �i�r�w�n�r�en ���      'NULL�Ή�
                .HSXOFnHS(n) = rs("E020HSXOF" & n & "HS")       ' �i�r�w�n�r�en �ۏؕ��@�Q��
            Next
            .HSXDENMX = fncNullCheck(rs("E020HSXDENMX"))                      ' �i�r�w�c�������    'NULL�Ή�
            .HSXDENMN = fncNullCheck(rs("E020HSXDENMN"))                      ' �i�r�w�c��������    'NULL�Ή�
            .HSXDENHS = rs("E020HSXDENHS")                      ' �i�r�w�c�����ۏؕ��@�Q��
            .HSXDVDMX = fncNullCheck(rs("E020HSXDVDMXN"))                     ' �i�r�w�c�u�c�Q���      ���ڒǉ��C�C���Ή� 2003.05.20 yakimura 'NULL�Ή�
            .HSXDVDMN = fncNullCheck(rs("E020HSXDVDMNN"))                     ' �i�r�w�c�u�c�Q����      ���ڒǉ��C�C���Ή� 2003.05.20 yakimura  'NULL�Ή�
            .HSXDVDHS = rs("E020HSXDVDHS")                      ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
            .HSXLDLMX = fncNullCheck(rs("E020HSXLDLMX"))                      ' �i�r�w�k�^�c�k���  'NULL�Ή�
            .HSXLDLMN = fncNullCheck(rs("E020HSXLDLMN"))                      ' �i�r�w�k�^�c�k����  'NULL�Ή�
            .HSXLDLHS = rs("E020HSXLDLHS")                      ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
            .HSXLTMIN = fncNullCheck(rs("E019HSXLTMIN"))                      ' �i�r�w�k�^�C������  'NULL�Ή�
            .HSXLTMAX = fncNullCheck(rs("E019HSXLTMAX"))                      ' �i�r�w�k�^�C�����  'NULL�Ή�
            .HSXLTHWS = rs("E019HSXLTHWS")                      ' �i�r�w�k�^�C���ۏؕ��@�Q��
            .HSXDPDIR = rs("E018HSXDPDIR")                      ' �i�r�w�a�ʒu����
            .HSXDPDRC = rs("E018HSXDPDRC")                      ' �i�r�w�a�ʒu����
            .HSXDWMIN = fncNullCheck(rs("E018HSXDWMIN"))                      ' �i�r�w�a�Љ���  'NULL�Ή�
            .HSXDWMAX = fncNullCheck(rs("E018HSXDWMAX"))                      ' �i�r�w�a�Џ��  'NULL�Ή�
            .HSXDDMIN = fncNullCheck(rs("E018HSXDDMIN"))                      ' �i�r�w�a�[����  'NULL�Ή�
            .HSXDDMAX = fncNullCheck(rs("E018HSXDDMAX"))                      ' �i�r�w�a�[���  'NULL�Ή�
            .HSXD1MIN = fncNullCheck(rs("E018HSXD1MIN"))                      ' �i�r�w���a�P����    'NULL�Ή�
            .HSXD1MAX = fncNullCheck(rs("E018HSXD1MAX"))                      ' �i�r�w���a�P���    'NULL�Ή�
            .HSXCTCEN = fncNullCheck(rs("E018HSXCTCEN"))                      ' �i�r�w�����ʌX�c���S    'NULL�Ή�
            .HSXCYCEN = fncNullCheck(rs("E018HSXCYCEN"))                      ' �i�r�w�����ʌX�����S    'NULL�Ή�
            .EPDUP = fncNullCheck(rs("EPDUP"))                                ' ���������Ǘ� EPD�@���  'NULL�Ή�
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
    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


'�����ŏI�����ւ̑}���i�����֐��j
Private Function fcmkc001f_ExecFts(CryIn As typ_cmkc001f_ExecCryIn, CryFTest As typ_cmkc001f_ExecFts) As FUNCTION_RETURN
Dim sql As String
Dim n As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Sub fcmkc001f_ExecFts"

    fcmkc001f_ExecFts = FUNCTION_RETURN_SUCCESS
        
    '�����ŏI�����ւ̑}��
    With CryFTest
        sql = "insert into TBCMJ010 ( "
        sql = sql & "CRYNUM, "                      ' �����ԍ�
        sql = sql & "INGOTPOS, "                    ' �C���S�b�g���ʒu
        sql = sql & "TRANCNT, "                     ' ������
        sql = sql & "LENGTH, "                      ' ����
        sql = sql & "KRPROCCD, "                    ' �Ǘ��H���R�[�h
        sql = sql & "PROCCODE, "                    ' �H���R�[�h
        sql = sql & "PAYCLASS, "                    ' �����o���敪
        sql = sql & "OUTLENGTH, "                   ' �o�ג���
        For n = 1 To 5
            sql = sql & "PART" & n & ", "           ' ����n
            sql = sql & "P" & n & "BDLEN, "         ' ����n �s�ǒ���
            sql = sql & "P" & n & "BDCAUS, "        ' ����n �s�Ǘ��R
        Next
        sql = sql & "TSTAFFID, "                    ' �o�^�Ј�ID
        sql = sql & "REGDATE, "                     ' �o�^���t
        sql = sql & "KSTAFFID, "                    ' �X�V�Ј�ID
        sql = sql & "UPDDATE, "                     ' �X�V���t
        sql = sql & "SUMMITSENDFLAG, "              ' SUMMIT���M�t���O
        sql = sql & "SENDFLAG, "                    ' ���M�t���O
        sql = sql & "SENDDATE ) "                   ' ���M���t
        
        sql = sql & "select "
        sql = sql & " '" & CryIn.CRYNUM & "', "     ' �����ԍ�
        sql = sql & CryIn.INGOTPOS & ", "           ' �C���S�b�g���ʒu
        sql = sql & "nvl(max(TRANCNT),0)+1, "       ' ������
        sql = sql & .LENGTH & ", "                  ' ����
        sql = sql & " '" & .KRPROCCD & "', "        ' �Ǘ��H���R�[�h
        sql = sql & " '" & .PROCCODE & "', "        ' �H���R�[�h
        sql = sql & " '" & .PAYCLASS & "', "        ' �����o���敪
        sql = sql & .OUTLENGTH & ", "               ' �o�ג���
        
        For n = 1 To 5
            sql = sql & .PART(n) & ", "             ' ����n
            sql = sql & .BDLEN(n) & ", "            ' ����n �s�ǒ���
            sql = sql & " '" & .BDCAUS(n) & "', "   ' ����n �s�Ǘ��R
        Next
        sql = sql & " '" & .TSTAFFID & "', "        ' �o�^�Ј�ID
        sql = sql & "sysdate, "                     ' �o�^���t
        sql = sql & " '" & .TSTAFFID & "', "        ' �X�V�Ј�ID
        sql = sql & "sysdate, "                     ' �X�V���t
        sql = sql & "'0', "                         ' SUMMIT���M�t���O
        sql = sql & "'0', "                         ' ���M�t���O
        sql = sql & "sysdate "                      ' ���M���t
        sql = sql & " From TBCMJ010 "
        sql = sql & " where CRYNUM='" & CryIn.CRYNUM & "' and INGOTPOS=" & CryIn.INGOTPOS
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        fcmkc001f_ExecFts = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    fcmkc001f_ExecFts = FUNCTION_RETURN_FAILURE
    Debug.Print "==== ERROR"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


'�N���X�^���J�^���O������тւ̑}���i�����֐��j
Private Function fcmkc001f_ExecCatalog(CryIn As typ_cmkc001f_ExecCryIn, CryCatalog As typ_cmkc001f_ExecCatalog) As FUNCTION_RETURN
Dim sql As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function fcmkc001f_ExecCatalog"

    fcmkc001f_ExecCatalog = FUNCTION_RETURN_SUCCESS


    ' �N���X�^���J�^���O������тւ̑}��
    sql = "insert into TBCMG007 ( "
    sql = sql & "CRYNUM, "            ' �����ԍ�
    sql = sql & "TRANCNT, "           ' ������
    sql = sql & "KRPROCCD, "          ' �Ǘ��H���R�[�h
    sql = sql & "PROCCODE, "          ' �H���R�[�h
    sql = sql & "BDCODE, "            ' �s�Ǘ��R�R�[�h
    sql = sql & "PALTNUM, "           ' �p���b�g�ԍ�
    sql = sql & "TSTAFFID, "          ' �o�^�Ј�ID
    sql = sql & "REGDATE, "           ' �o�^���t
    sql = sql & "KSTAFFID, "          ' �X�V�Ј�ID
    sql = sql & "UPDDATE, "           ' �X�V���t
    sql = sql & "SENDFLAG, "          ' ���M�t���O
    sql = sql & "SENDDATE) "          ' ���M���t

    With CryCatalog
        sql = sql & "Select "
        sql = sql & " '" & .CRYNUM & "', "                          ' �����ԍ�
        sql = sql & "nvl(max(TRANCNT),0)+1, "                       ' ������
        sql = sql & " '" & MGPRCD_KESSYOU_SAISYUU_HARAIDASI & "', " ' �Ǘ��H���R�[�h
        sql = sql & " '" & PROCD_KESSYOU_SAISYUU_HARAIDASI & "', "  ' �H���R�[�h
        sql = sql & " '" & .BDCODE & "', "                          ' �s�Ǘ��R�R�[�h
        sql = sql & " '" & .PALTNUM & "', "                         ' �p���b�g�ԍ�
        sql = sql & " '" & .TSTAFFID & "', "                        ' �o�^�Ј�ID
        sql = sql & "sysdate, "                                     ' �o�^���t
        sql = sql & " '" & .TSTAFFID & "', "                        ' �X�V�Ј�ID
        sql = sql & "sysdate, "                                     ' �X�V���t
        sql = sql & "'0', "                                         ' ���M�t���O
        sql = sql & "sysdate "                                      ' ���M���t
        sql = sql & "From TBCMG007 " & _
              "Where (CRYNUM='" & .CRYNUM & "')"
    End With

    If 0 >= OraDB.ExecuteSQL(sql) Then
        fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


' �u���b�N�Ǘ��̍X�V�i�����֐��j
Private Function fcmkc001f_ExecBlock(CryIn As typ_cmkc001f_ExecCryIn, BlockMan As typ_cmkc001f_Block, Optional BDCAUS$ = vbNullString) As FUNCTION_RETURN
Dim sql As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function fcmkc001f_ExecBlock"

    fcmkc001f_ExecBlock = FUNCTION_RETURN_SUCCESS

    ' �u���b�N�Ǘ��̍X�V
    With BlockMan
        sql = "update TBCME040 set "
        sql = sql & "REALLEN=" & .REALLEN & ", "            ' ������
        sql = sql & "KRPROCCD='" & .KRPROCCD & "', "        ' ���݊Ǘ��H��
        sql = sql & "NOWPROC='" & .NOWPROC & "', "          ' ���ݍH��
        sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "    ' �ŏI�ʉߊǗ��H��
        sql = sql & "LASTPASS='" & .LASTPASS & "', "        ' �ŏI�ʉߍH��
        sql = sql & "DELCLS='" & .DELCLS & "', "            ' �폜�敪
        sql = sql & "RSTATCLS='" & .RSTATCLS & "', "        ' ������ԋ敪
        sql = sql & "LSTATCLS='" & .LSTATCLS & "', "        ' �ŏI��ԋ敪
        If BDCAUS <> vbNullString Then
            sql = sql & "BDCAUS='" & BDCAUS & "', "         ' �ŏI��ԋ敪
        End If
        sql = sql & "UPDDATE=SYSDATE, "                     ' �X�V��
        sql = sql & "SENDFLAG='0' "
        sql = sql & " where  "
        sql = sql & "CRYNUM='" & CryIn.CRYNUM & "' and INGOTPOS=" & CryIn.INGOTPOS
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        fcmkc001f_ExecBlock = FUNCTION_RETURN_FAILURE
    End If
    

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    fcmkc001f_ExecBlock = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


'2008/06/01 SHINDOH--------------------------------------
'Public Function fcmkc001f_Exec_037(sCryNum As String) As FUNCTION_RETURN
Public Function fcmkc001f_Exec_037(sCrynum As String, sProcCd As String) As FUNCTION_RETURN
'2008/06/01 SHINDOH--------------------------------------
    Dim sDbName As String
    Dim sql     As String
    Dim sErrMsg As String
    
    
    '' �������̍X�V
    sDbName = "E037"
    sql = "update TBCME037 set "
    sql = sql & "KRPROCCD  ='" & MGPRCD_WFC_HARAIDASI & "', "
'2008/06/01 SHINDOH--------------------------------------
'    sql = sql & "PROCCD    ='" & PROCD_WFC_HARAIDASI & "', "
    sql = sql & "PROCCD    ='" & sProcCd & "', "
'2008/06/01 SHINDOH--------------------------------------
    sql = sql & "LPKRPROCCD='" & MGPRCD_KESSYOU_SAISYUU_HARAIDASI & "', "
    sql = sql & "LASTPASS  ='" & PROCD_KESSYOU_SAISYUU_HARAIDASI & "', "
    sql = sql & "UPDDATE   = sysdate, "
    sql = sql & "SENDFLAG  ='0'"
    sql = sql & " where CRYNUM='" & sCrynum & "'"
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        fcmkc001f_Exec_037 = FUNCTION_RETURN_FAILURE
    Else
        fcmkc001f_Exec_037 = FUNCTION_RETURN_SUCCESS
    End If
End Function


'���s�����C��
Public Function DBDRV_fcmkc001f_Exec( _
  CryIn As typ_cmkc001f_ExecCryIn, CryFTest As typ_cmkc001f_ExecFts, _
  BlockMan As typ_cmkc001f_Block, blkID As String, STAFFID As String) As FUNCTION_RETURN

Dim skipNukishi As Boolean
Dim sql$
Dim sqlWhere$
Dim INGOTPOS%
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001f_SQL.bas -- Function DBDRV_fcmkc001f_Exec"

    DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_SUCCESS
    BlockMan.RSTATCLS = "T" '???

'�v�e�T���v�������ύX 2003.05.20 yakimura
'    '(P+/N+��)�����w���H�����΂����ǂ����̃t���O��ݒ�
'    If CryFTest.PROCCODE = PROCD_WFC_HARAIDASI Then
'        skipNukishi = True
'    Else
'        skipNukishi = False
'    End If
'�v�e�T���v�������ύX 2003.05.20 yakimura
    
    '�u���b�N�Ǘ��e�[�u���̍X�V(�Ȃ�������}��???)
    With BlockMan
        .KRPROCCD = CryFTest.KRPROCCD
        .NOWPROC = CryFTest.PROCCODE

'�v�e�T���v�������ύX 2003.05.20 yakimura
'        If skipNukishi Then
'            .LPKRPROCCD = MGPRCD_NUKISI_SIJI
'            .LASTPASS = PROCD_NUKISI_SIJI
'        Else
            .LPKRPROCCD = MGPRCD_KESSYOU_SAISYUU_HARAIDASI
            .LASTPASS = PROCD_KESSYOU_SAISYUU_HARAIDASI
'        End If
'�v�e�T���v�������ύX 2003.05.20 yakimura
        
        .REALLEN = CryFTest.OUTLENGTH
    End With
        
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecBlock(CryIn, BlockMan) Then
        DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_FAILURE
    End If
    
    '�����ŏI�����փC���T�[�g
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecFts(CryIn, CryFTest) Then
        DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_FAILURE
    End If
    
'�v�e�T���v�������ύX 2003.05.20 yakimura
    '(P+/N+��)�����w���H�����΂��ꍇ�̒ǉ�����
'    If skipNukishi Then
'        sqlWhere = " where CRYNUM='" & CryIn.CRYNUM & "' and INGOTPOS=" & BlockMan.INGOTPOS
        
        'SXL�Ǘ������
'        sql = "insert into TBCME042 ("
'        sql = sql & "CRYNUM,INGOTPOS,LENGTH,SXLID,KRPROCCD,NOWPROC,LPKRPROCCD,LASTPASS,DELCLS,LSTATCLS,HOLDCLS"
'        sql = sql & ",HINBAN,REVNUM,FACTORY,OPECOND,BDCAUS,COUNT"
'        sql = sql & ",REGDATE,UPDDATE,SUMMITSENDFLAG,SENDFLAG,SENDDATE,PASSFLAG"
'        sql = sql & ") select"
'        sql = sql & " CRYNUM, HINFROM, HINTO-HINFROM"
'        sql = sql & ", substr(BLOCKID,1,10) || substr('0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ',HINFROM/100+1,1) || to_char(mod(HINFROM,100),'FM00') as SXLID"
'        sql = sql & ", ' ', 'CC720', ' ', 'CC710', '0', 'T', '0'"
'        sql = sql & ", HINBAN, REVNUM, FACTORY, OPECOND"
'        sql = sql & ", ' ', 0, sysdate, sysdate, '0', '0', sysdate, ' ' "
'        sql = sql & "from"
'        sql = sql & "("
'        sql = sql & " select BLK.CRYNUM, BLK.BLOCKID, HIN.HINBAN, HIN.REVNUM, HIN.FACTORY, HIN.OPECOND"
'        sql = sql & " , greatest(BLK.INGOTPOS,HIN.INGOTPOS) as HINFROM"
'        sql = sql & " , least(BLK.INGOTPOS+BLK.LENGTH,HIN.INGOTPOS+HIN.LENGTH) as HINTO"
'        sql = sql & " from TBCME041 HIN, TBCME040 BLK"
'        sql = sql & " where BLK.CRYNUM='" & CryIn.CRYNUM & "' and BLK.INGOTPOS=" & BlockMan.INGOTPOS
'        sql = sql & "  and HIN.CRYNUM=BLK.CRYNUM"
'        sql = sql & "  and HIN.INGOTPOS<BLK.INGOTPOS+BLK.LENGTH"
'        sql = sql & "  and HIN.INGOTPOS+HIN.LENGTH>BLK.INGOTPOS"
'        sql = sql & ") HINS"
'        If (OraDB.ExecuteSQL(sql) < 1) Then
'            DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
        
'        '�����w�����т����
'        sql = "insert into TBCMW001 ("
'        sql = sql & " CRYNUM, INGOTPOS, TRANCNT, CRYLEN, KRPROCCD, PROCCODE,"
'        sql = sql & " BLOCKID, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE"
'        sql = sql & ") values ("
'        sql = sql & "'" & CryIn.CRYNUM & "', " & BlockMan.INGOTPOS & ","
'        sql = sql & " (select nvl(max(trancnt),0)+1 from TBCMW001" & sqlWhere & "), "
'        sql = sql & BlockMan.LENGTH & ","
'        sql = sql & " '" & MGPRCD_KESSYOU_SAISYUU_HARAIDASI & "', '" & PROCD_KESSYOU_SAISYUU_HARAIDASI & "','"
'        sql = sql & blkID & "', "
'        sql = sql & "'" & STAFFID & "', sysdate, ' ', sysdate, '0', sysdate)"
'        If (OraDB.ExecuteSQL(sql) < 1) Then
'            DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'BAR�o�ׂ̂Ƃ�
Public Function DBDRV_fcmkc001f_ExecBar(CryIn As typ_cmkc001f_ExecCryIn, CryFTest As typ_cmkc001f_ExecFts, BlockMan As typ_cmkc001f_Block) As FUNCTION_RETURN

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001f_SQL.bas -- Function DBDRV_fcmkc001f_ExecBar"

    DBDRV_fcmkc001f_ExecBar = FUNCTION_RETURN_SUCCESS
'OraDB.BeginTrans
    '�u���b�N�Ǘ��e�[�u���̍X�V(�Ȃ�������}��???)
    With BlockMan
        .DELCLS = "1"
        .LSTATCLS = "B"
        .KRPROCCD = "     "
        .NOWPROC = "CC705"
        .LPKRPROCCD = MGPRCD_KESSYOU_SAISYUU_HARAIDASI
        .LASTPASS = PROCD_KESSYOU_SAISYUU_HARAIDASI
        .REALLEN = CryFTest.OUTLENGTH
    End With
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecBlock(CryIn, BlockMan) Then
        DBDRV_fcmkc001f_ExecBar = FUNCTION_RETURN_FAILURE
    End If
    
    '�����ŏI�����փC���T�[�g
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecFts(CryIn, CryFTest) Then
        DBDRV_fcmkc001f_ExecBar = FUNCTION_RETURN_FAILURE
    End If
    
'OraDB.CommitTrans
    
    '�u���b�N�Ǘ��e�[�u���̍X�V(�Ȃ�������}��???)(�u���b�N�Ǘ�???)
'    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecBlock(CryIn, BlockMan) Then
'        DBDRV_fcmkc001f_ExecBar = FUNCTION_RETURN_FAILURE
'    End If


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_fcmkc001f_ExecBar = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�N���X�^���J�^���O�i��
Public Function DBDRV_fcmkc001f_ExecCatalog(CryIn As typ_cmkc001f_ExecCryIn, CryFTest As typ_cmkc001f_ExecFts, _
                                           BlockMan As typ_cmkc001f_Block, CryCatalog As typ_cmkc001f_ExecCatalog) As FUNCTION_RETURN
Dim HIN As tFullHinban

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001f_SQL.bas -- Function DBDRV_fcmkc001f_ExecCatalog"

    DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_SUCCESS
    CryFTest.PAYCLASS = "2"
    BlockMan.RSTATCLS = "G" '???
    BlockMan.NOWPROC = PROCD_KAKUAGE
    BlockMan.KRPROCCD = MGPRCD_KAKUAGE

    '�u���b�N�Ǘ��e�[�u���̍X�V(�Ȃ�������}��???)
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecBlock(CryIn, BlockMan, CryCatalog.BDCODE) Then
        DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    End If

    '�i�ԊǗ��e�[�u���̍X�V
    With HIN
        .hinban = "G"
        .mnorevno = 0
        .factory = " "
        .opecond = " "
    End With
    With BlockMan
        If ChangeAreaHinban(CryIn.CRYNUM, .INGOTPOS, .LENGTH, HIN) = FUNCTION_RETURN_FAILURE Then
            DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
        End If
    End With

    '�����ŏI�����փC���T�[�g
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecFts(CryIn, CryFTest) Then
        DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    End If

    '�N���X�^���J�^���O������тփC���T�[�g
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecCatalog(CryIn, CryCatalog) Then
        DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�������g
Public Function DBDRV_fcmkc001f_ExecRemelt(CryIn As typ_cmkc001f_ExecCryIn, CryFTest As typ_cmkc001f_ExecFts, _
                                           BlockMan As typ_cmkc001f_Block) As FUNCTION_RETURN
Dim HIN As tFullHinban

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001f_SQL.bas -- Function DBDRV_fcmkc001f_ExecRemelt"

    DBDRV_fcmkc001f_ExecRemelt = FUNCTION_RETURN_SUCCESS
    CryFTest.PAYCLASS = "3" '???
    BlockMan.RSTATCLS = "M"
    BlockMan.NOWPROC = PROCD_RIMERUTO_UKEIRE
    BlockMan.KRPROCCD = MGPRCD_RIMERUTO_UKEIRE

    '�u���b�N�Ǘ��e�[�u���̍X�V(�Ȃ�������}��???)
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecBlock(CryIn, BlockMan) Then
        DBDRV_fcmkc001f_ExecRemelt = FUNCTION_RETURN_FAILURE
    End If
    
    '�i�ԊǗ��e�[�u���̍X�V
    With HIN
        .hinban = "Z"
        .mnorevno = 0
        .factory = " "
        .opecond = " "
    End With
    With BlockMan
        If ChangeAreaHinban(CryIn.CRYNUM, .INGOTPOS, .LENGTH, HIN) = FUNCTION_RETURN_FAILURE Then
            DBDRV_fcmkc001f_ExecRemelt = FUNCTION_RETURN_FAILURE
        End If
    End With

    '�����ŏI�����փC���T�[�g
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecFts(CryIn, CryFTest) Then
        DBDRV_fcmkc001f_ExecRemelt = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_fcmkc001f_ExecRemelt = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


''''''(2002/07 s_cmzcF_cmhc001d_SQL.bas���ړ�)
'''''Private Function AreaStr(cnd$, v1, v2, fmt$) As String
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmzcF_cmhc001d_SQL.bas -- Function AreaStr"
'''''
'''''    If Trim$(cnd) = vbNullString Then                           ''�ۏؕ��@���󗓂Ȃ�K�i�Ȃ�
'''''        AreaStr = vbNullString
'''''    Else
'''''        AreaStr = Format$(v1, fmt) & " - " & Format$(v2, fmt)   ''�w��̏����Ŕ͈͕�������쐬
'''''    End If
'''''
'''''PROC_EXIT:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    '�G���[�n���h��
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function


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
'''''    Dim l       As Long
'''''    Dim m       As Long
'''''    Dim sql     As String
'''''    Dim rs      As OraDynaset    'RecordSet
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo PROC_ERR
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
'''''    sql = sql & "select"
'''''    sql = sql & "  B.BLOCKID, B.INGOTPOS as TOPPOS, B.INGOTPOS+LENGTH as BOTPOS"
'''''    sql = sql & ", S.XTALCS, S.INPOSCS, SMPKBNCS, HINBCS, REVNUMCS, FACTORYCS, OPECS"
'''''    sql = sql & ", CRYINDRSCS, CRYRESRS1CS, CRYINDOICS, CRYRESOICS"
'''''    sql = sql & ", CRYINDB1CS, CRYRESB1CS, CRYINDB2CS, CRYRESB2CS, CRYINDB3CS, CRYRESB3CS"
'''''    sql = sql & ", CRYINDL1CS, CRYRESL1CS, CRYINDL2CS, CRYRESL2CS, CRYINDL3CS, CRYRESL3CS, CRYINDL4CS, CRYRESL4CS"
'''''    sql = sql & ", CRYINDCSCS, CRYRESCSCS, CRYINDGDCS, CRYRESGDCS, CRYINDTCS, CRYRESTCS, CRYINDEPCS, CRYRESEPCS "
'''''    sql = sql & "from XSDCS S, TBCME040 B "
'''''    sql = sql & "where S.XTALCS=B.CRYNUM"
'''''    sql = sql & "  and B.INGOTPOS>=0"
'''''    sql = sql & "  and B.DELCLS='0'"
'''''    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
'''''    sql = sql & "  and B.RSTATCLS='T'"
'''''    sql = sql & "  and B.HOLDCLS='0'"
'''''    sql = sql & "  and ((S.INPOSCS=B.INGOTPOS) or (S.INPOSCS=B.INGOTPOS+B.LENGTH)) "
'''''    sql = sql & "order by B.BLOCKID, S.INPOSCS, S.SMPKBNCS"
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
'''''            .CRYNUM = rs("CRYNUM")
'''''            .INGOTPOS = rs("INGOTPOS")
'''''            .SMPKBN = rs("SMPKBN")
'''''            .hinban = rs("HINBAN")
'''''            .REVNUM = rs("REVNUM")
'''''            .factory = rs("FACTORY")
'''''            .opecond = rs("OPECOND")
'''''            .CRYINDRS = rs("CRYINDRS")
'''''            .CRYRESRS = rs("CRYRESRS")
'''''            .CRYINDOI = rs("CRYINDOI")
'''''            .CRYRESOI = rs("CRYRESOI")
'''''            .CRYINDB1 = rs("CRYINDB1")
'''''            .CRYRESB1 = rs("CRYRESB1")
'''''            .CRYINDB2 = rs("CRYINDB2")
'''''            .CRYRESB2 = rs("CRYRESB2")
'''''            .CRYINDB3 = rs("CRYINDB3")
'''''            .CRYRESB3 = rs("CRYRESB3")
'''''            .CRYINDL1 = rs("CRYINDL1")
'''''            .CRYRESL1 = rs("CRYRESL1")
'''''            .CRYINDL2 = rs("CRYINDL2")
'''''            .CRYRESL2 = rs("CRYRESL2")
'''''            .CRYINDL3 = rs("CRYINDL3")
'''''            .CRYRESL3 = rs("CRYRESL3")
'''''            .CRYINDL4 = rs("CRYINDL4")
'''''            .CRYRESL4 = rs("CRYRESL4")
'''''            .CRYINDCS = rs("CRYINDCS")
'''''            .CRYRESCS = rs("CRYRESCS")
'''''            .CRYINDGD = rs("CRYINDGD")
'''''            .CRYRESGD = rs("CRYRESGD")
'''''            .CRYINDT = rs("CRYINDT")
'''''            .CRYREST = rs("CRYREST")
'''''            .CRYINDEP = rs("CRYINDEP")
'''''            .CRYRESEP = rs("CRYRESEP")
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
'''''                    .INGOTPOS = SMP(idx).INGOTPOS
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
'''''            sql = " select "
'''''            sql = sql & " V.E043CRYNUM, "
'''''            sql = sql & " V.E043INGOTPOS, "
'''''            sql = sql & " V.E043SMPKBN, "
'''''            sql = sql & " V.E043HINBAN, "
'''''            sql = sql & " V.E043REVNUM, "
'''''            sql = sql & " V.E043FACTORY, "
'''''            sql = sql & " V.E043OPECOND, "
'''''            sql = sql & " V.E043CRYINDRS, "
'''''            sql = sql & " V.E043CRYRESRS, "
'''''            sql = sql & " V.E043CRYINDOI, "
'''''            sql = sql & " V.E043CRYRESOI, "
'''''            sql = sql & " V.E043CRYINDB1, "
'''''            sql = sql & " V.E043CRYRESB1, "
'''''            sql = sql & " V.E043CRYINDB2, "
'''''            sql = sql & " V.E043CRYRESB2, "
'''''            sql = sql & " V.E043CRYINDB3, "
'''''            sql = sql & " V.E043CRYRESB3, "
'''''            sql = sql & " V.E043CRYINDL1, "
'''''            sql = sql & " V.E043CRYRESL1, "
'''''            sql = sql & " V.E043CRYINDL2, "
'''''            sql = sql & " V.E043CRYRESL2, "
'''''            sql = sql & " V.E043CRYINDL3, "
'''''            sql = sql & " V.E043CRYRESL3, "
'''''            sql = sql & " V.E043CRYINDL4, "
'''''            sql = sql & " V.E043CRYRESL4, "
'''''            sql = sql & " V.E043CRYINDCS, "
'''''            sql = sql & " V.E043CRYRESCS, "
'''''            sql = sql & " V.E043CRYINDGD, "
'''''            sql = sql & " V.E043CRYRESGD, "
'''''            sql = sql & " V.E043CRYINDT, "
'''''            sql = sql & " V.E043CRYREST, "
'''''            sql = sql & " V.E043CRYINDEP, "
'''''            sql = sql & " V.E043CRYRESEP "
'''''            sql = sql & " from VECME010 V "
'''''            sql = sql & " where E040CRYNUM = '" & .CRYNUM & "' "
'''''            sql = sql & " and   E040INGOTPOS = '" & .INGOTPOS & "' "
'''''            sql = sql & " order by E043INGOTPOS"
'''''
'''''            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''            For m = 1 To 2
'''''                DoEvents
'''''                .SMP(m).CRYNUM = rs("E043CRYNUM")
'''''                .SMP(m).INGOTPOS = rs("E043INGOTPOS")
'''''                .SMP(m).SMPKBN = rs("E043SMPKBN")
'''''                .SMP(m).hinban = rs("E043HINBAN")
'''''                .SMP(m).REVNUM = rs("E043REVNUM")
'''''                .SMP(m).factory = rs("E043FACTORY")
'''''                .SMP(m).opecond = rs("E043OPECOND")
'''''                .SMP(m).CRYINDRS = rs("E043CRYINDRS")
'''''                .SMP(m).CRYRESRS = rs("E043CRYRESRS")
'''''                .SMP(m).CRYINDOI = rs("E043CRYINDOI")
'''''                .SMP(m).CRYRESOI = rs("E043CRYRESOI")
'''''                .SMP(m).CRYINDB1 = rs("E043CRYINDB1")
'''''                .SMP(m).CRYRESB1 = rs("E043CRYRESB1")
'''''                .SMP(m).CRYINDB2 = rs("E043CRYINDB2")
'''''                .SMP(m).CRYRESB2 = rs("E043CRYRESB2")
'''''                .SMP(m).CRYINDB3 = rs("E043CRYINDB3")
'''''                .SMP(m).CRYRESB3 = rs("E043CRYRESB3")
'''''                .SMP(m).CRYINDL1 = rs("E043CRYINDL1")
'''''                .SMP(m).CRYRESL1 = rs("E043CRYRESL1")
'''''                .SMP(m).CRYINDL2 = rs("E043CRYINDL2")
'''''                .SMP(m).CRYRESL2 = rs("E043CRYRESL2")
'''''                .SMP(m).CRYINDL3 = rs("E043CRYINDL3")
'''''                .SMP(m).CRYRESL3 = rs("E043CRYRESL3")
'''''                .SMP(m).CRYINDL4 = rs("E043CRYINDL4")
'''''                .SMP(m).CRYRESL4 = rs("E043CRYRESL4")
'''''                .SMP(m).CRYINDCS = rs("E043CRYINDCS")
'''''                .SMP(m).CRYRESCS = rs("E043CRYRESCS")
'''''                .SMP(m).CRYINDGD = rs("E043CRYINDGD")
'''''                .SMP(m).CRYRESGD = rs("E043CRYRESGD")
'''''                .SMP(m).CRYINDT = rs("E043CRYINDT")
'''''                .SMP(m).CRYREST = rs("E043CRYREST")
'''''                .SMP(m).CRYINDEP = rs("E043CRYINDEP")
'''''                .SMP(m).CRYRESEP = rs("E043CRYRESEP")
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
'''''                    sql = sql & " S.HSXCNHWS, "
'''''                    sql = sql & " S.HSXLTHWS, "
'''''                    sql = sql & " 'H' as EPD "
'''''                    sql = sql & " from TBCME019 S "
'''''                    sql = sql & " where S.HINBAN = '" & .SMP(m).hinban & "' "
'''''                    sql = sql & " and S.MNOREVNO = " & .SMP(m).REVNUM & " "
'''''                    sql = sql & " and S.FACTORY = '" & .SMP(m).factory & "' "
'''''                    sql = sql & " and S.OPECOND = '" & .SMP(m).opecond & "' "
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
'''''                    sql = "select CRYRESCSCS as RES from XSDCS "
'''''                    sql = sql & "where CRYNUM = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).INGOTPOS
'''''                    sql = sql & "  and CRYINDCSCS<>'0'"
'''''                    sql = sql & " order by INPOSCS"
'''''
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
'''''                    sql = "select CRYRESTCS as RES from XSDCS "
'''''                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).INGOTPOS
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
'''''                    sql = "select CRYRESEPCS as RES from XSDCS "
'''''                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).INGOTPOS
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
'''''PROC_EXIT:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    '�G���[�n���h��
'''''    gErr.HandleError
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    cmkc001b_DBDataCheck1 = FUNCTION_RETURN_FAILURE
'''''    Resume PROC_EXIT
'''''End Function


'''''Public Function cmkc001b_DBDataCheck3(LWD() As cmkc001b_LockWait, _
'''''                                 Wd3() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''    Dim c0              As Integer
'''''    Dim c1              As Integer
'''''    Dim c2              As Integer
'''''    Dim MaxRec          As Integer
'''''    Dim RecCount        As Integer
'''''    Dim EQFlag          As Boolean
'''''    Dim sql             As String       'SQL�S��
'''''    Dim rs              As OraDynaset    'RecordSet
'''''    Dim GrpCount1       As Integer
'''''    Dim GrpCount2       As Integer
'''''    Dim ColorFlag       As Boolean
'''''    Dim TotalBlk        As Integer
'''''    Dim CheckPoint      As Integer
'''''    Dim CheckEnd        As Integer
'''''    Dim tempGrpFlag     As String * 1
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo PROC_ERR
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
'''''
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
'''''            GoTo PROC_EXIT
'''''        End If
'''''        ReDim GrpInfo(c0).blkInfo(RecCount) As cmkc001b_Wait3_BLK
'''''        For c1 = 1 To RecCount
'''''            GrpInfo(c0).blkInfo(c1).BLOCKID = rs("BLOCKID")
'''''            GrpInfo(c0).blkInfo(c1).INGOTPOS = rs("INGOTPOS")
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
'''''Dim blkID() As String
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
'''''    ReDim blkID(1 To rsCount)
'''''    ReDim topHin(1 To rsCount)
'''''    ReDim botHin(1 To rsCount)
'''''    For c0 = 1 To rsCount
'''''        blkID(c0) = rs!BLOCKID
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
'''''                If blkID(idx) = GrpInfo(c0).blkInfo(c1).BLOCKID Then
'''''                    found = True
'''''                    Exit For
'''''                ElseIf blkID(idx) > GrpInfo(c0).blkInfo(c1).BLOCKID Then
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
'''''            sql = sql & "and INGOTPOS = " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " " '2001/11/14 S.Sano
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
'''''            sql = sql & "and INGOTPOS < " & GrpInfo(c0).blkInfo(c1).INGOTPOS + GrpInfo(c0).blkInfo(c1).LENGTH & " "
'''''            sql = sql & "and (INGOTPOS + LENGTH) >= " & GrpInfo(c0).blkInfo(c1).INGOTPOS + GrpInfo(c0).blkInfo(c1).LENGTH & " "
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
'''''
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
'''''
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
'''''PROC_EXIT:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    '�G���[�n���h��
'''''    cmkc001b_DBDataCheck3 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function


''''''�T�v    :�҂��ꗗ �����\���p�c�a�h���C�o�i�����҂��j
''''''���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
''''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
''''''        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
''''''����    :
''''''����    :2001/07/06 ���{ �쐬
'''''Public Function DBDRV_scmzc_fcmkc001b_Disp1(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''
'''''    Dim sql As String       'SQL�S��
'''''    Dim rs As OraDynaset    'RecordSet
'''''    Dim recCnt As Long      '�u���b�N�Ǘ��̃��R�[�h��
'''''    Dim i As Long
'''''    Dim j As Long
'''''    Dim k As Long
'''''    Dim BlockIdBuf As String
'''''
'''''    '<�����҂���
'''''    '�u���b�N�Ǘ��e�[�u������u���b�NID�A�X�V���t�擾�i�������т��������̂��́j
'''''
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp1"
'''''
'''''    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_SUCCESS
'''''
'''''    '�u���b�NID�A�X�V���t�̎擾
'''''    sql = "select distinct "
'''''    sql = sql & " V.E040CRYNUM, "
'''''    sql = sql & " V.E040INGOTPOS, "
'''''    sql = sql & " V.E040BLOCKID, "
'''''    sql = sql & " V.E040UPDDATE, "
'''''    sql = sql & " V.E040HOLDCLS, "
'''''    sql = sql & " H.HINBAN, "            ' �i��
'''''    sql = sql & " H.REVNUM, "            ' ���i�ԍ������ԍ�
'''''    sql = sql & " H.FACTORY, "           ' �H��
'''''    sql = sql & " H.OPECOND, "           ' ���Ə���
'''''    sql = sql & " S.HSXTYPE, "           ' �i�r�w�^�C�v
'''''    sql = sql & " S.HSXCDIR, "            ' �i�r�w�����ʕ���
'''''    sql = sql & " H.INGOTPOS "
'''''    sql = sql & " from "
'''''    sql = sql & " VECME010 V, TBCME041 H, TBCME018 S "
'''''    sql = sql & " where "
'''''    sql = sql & " V.E040CRYNUM = H.CRYNUM "
'''''    sql = sql & " and H.HINBAN = S.HINBAN "
'''''    sql = sql & " and H.REVNUM = S.MNOREVNO "
'''''    sql = sql & " and H.FACTORY = S.FACTORY "
'''''    sql = sql & " and H.OPECOND = S.OPECOND "
'''''                '�u���b�N���̕i�Ԍ���
'''''    sql = sql & " and (( V.E040INGOTPOS >= H.INGOTPOS "
'''''    sql = sql & " and V.E040INGOTPOS < H.INGOTPOS + H.LENGTH ) "
'''''    sql = sql & " or ( V.E040INGOTPOS + V.E040LENGTH > H.INGOTPOS "
'''''    sql = sql & " and V.E040INGOTPOS + V.E040LENGTH < H.INGOTPOS + H.LENGTH  ) "
'''''    sql = sql & " or ( H.INGOTPOS >= V.E040INGOTPOS "
'''''    sql = sql & " and H.INGOTPOS < V.E040INGOTPOS + V.E040LENGTH ) "
'''''    sql = sql & " or ( H.INGOTPOS + H.LENGTH > V.E040INGOTPOS "
'''''    sql = sql & " and H.INGOTPOS + H.LENGTH < V.E040INGOTPOS + V.E040LENGTH )) "
'''''                '�H���R�[�h�A��ԁA�敪�̏����w��
'''''    sql = sql & " and V.E040NOWPROC='CC600' "
'''''    sql = sql & " and V.E040LSTATCLS='T' "
'''''    sql = sql & " and V.E040RSTATCLS='T' "
'''''    sql = sql & " and V.E040DELCLS='0' "
'''''    'sql = sql & " and V.E040HOLDCLS='0' " ' �z�[���h�u���b�N���擾
'''''                '�w����0�łȂ����т�0
'''''    sql = sql & " and ((V.E043CRYINDRS<>'0' and V.E043CRYRESRS='0') "         ' �����������сiRs)
'''''    sql = sql & " or (V.E043CRYINDOI<>'0' and V.E043CRYRESOI='0') "         ' �����������сiOi)
'''''    sql = sql & " or (V.E043CRYINDB1<>'0' and V.E043CRYRESB1='0')"          ' �����������сiB1)
'''''    sql = sql & " or (V.E043CRYINDB2<>'0' and V.E043CRYRESB2='0') "         ' �����������сiB2�j
'''''    sql = sql & " or (V.E043CRYINDB3<>'0' and V.E043CRYRESB3='0') "         ' �����������сiB3)
'''''    sql = sql & " or (V.E043CRYINDL1<>'0' and V.E043CRYRESL1='0') "         ' �����������сiL1)
'''''    sql = sql & " or (V.E043CRYINDL2<>'0' and V.E043CRYRESL2='0') "         ' �����������сiL2)
'''''    sql = sql & " or (V.E043CRYINDL3<>'0' and V.E043CRYRESL3='0') "         ' �����������сiL3)
'''''    sql = sql & " or (V.E043CRYINDL4<>'0' and V.E043CRYRESL4='0') "         ' �����������сiL4)
'''''    sql = sql & " or (V.E043CRYINDCS<>'0' and V.E043CRYRESCS='0') "         ' �����������сiCs)
'''''    sql = sql & " or (V.E043CRYINDGD<>'0' and V.E043CRYRESGD='0') "         ' �����������сiGD)
'''''    sql = sql & " or (V.E043CRYINDT<>'0' and V.E043CRYREST='0') "           ' �����������сiT)
'''''    sql = sql & " or (V.E043CRYINDEP<>'0' and V.E043CRYRESEP='0')) "         ' �����������сiEPD)
'''''    sql = sql & " order by V.E040BLOCKID, H.INGOTPOS "
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
'''''            If rs("E040BLOCKID") <> BlockIdBuf Then
'''''
'''''                j = j + 1
'''''                ReDim Preserve records(j)
'''''
'''''                With records(j)
'''''                    .CRYNUM = rs("E040CRYNUM")
'''''                    .INGOTPOS = rs("E040INGOTPOS")
'''''                    .BLOCKID = rs("E040BLOCKID")   ' �u���b�NID
'''''                    .UPDDATE = rs("E040UPDDATE")   ' �X�V���t
'''''                    .HOLDCLS = rs("E040HOLDCLS")   ' �z�[���h�敪
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
'''''            ReDim Preserve records(j).HIN(k)
'''''            records(j).HIN(k).hinban = rs("HINBAN")
'''''            records(j).HIN(k).mnorevno = rs("REVNUM")
'''''            records(j).HIN(k).factory = rs("FACTORY")
'''''            records(j).HIN(k).opecond = rs("OPECOND")
'''''            k = k + 1
'''''            rs.MoveNext
'''''        Next i
'''''        rs.Close
'''''
'''''    End If
'''''
'''''
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
'''''    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
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
'''''    Dim sql As String       'SQL�S��
'''''    Dim rs As OraDynaset    'RecordSet
'''''    Dim recCnt As Long      '�u���b�N�Ǘ��̃��R�[�h��
'''''    Dim i As Long
'''''    Dim j As Long
'''''    Dim k As Long
'''''    Dim BlockIdBuf As String
'''''
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp2"
'''''
'''''    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_SUCCESS
'''''
'''''    sql = "select distinct "
'''''    sql = sql & " B.CRYNUM, "
'''''    sql = sql & " B.INGOTPOS as ss, "
''''''    sql = sql & " B.LENGTH, "             ' �����ǉ� 2001/11/8
'''''    sql = sql & " B.BLOCKID, "
'''''    sql = sql & " B.UPDDATE, "
'''''    sql = sql & " B.HOLDCLS, "
'''''    sql = sql & " H.HINBAN, "            ' �i��
'''''    sql = sql & " H.REVNUM, "            ' ���i�ԍ������ԍ�
'''''    sql = sql & " H.FACTORY, "           ' �H��
'''''    sql = sql & " H.OPECOND, "           ' ���Ə���
'''''    sql = sql & " S.HSXTYPE, "           ' �i�r�w�^�C�v
'''''    sql = sql & " S.HSXCDIR, "            ' �i�r�w�����ʕ���
'''''    sql = sql & " H.INGOTPOS, "
'''''                '����NG�����邩�ǂ���
'''''    sql = sql & " (select count(*) from VECME010 V1 "
'''''    sql = sql & "  where V1.E040BLOCKID=B.BLOCKID "
'''''    sql = sql & "  and ((V1.E043CRYINDRS<>'0' and V1.E043CRYRESRS='2') "         ' �����������сiRs)
'''''    sql = sql & "  or (V1.E043CRYINDOI<>'0' and V1.E043CRYRESOI='2') "         ' �����������сiOi)
'''''    sql = sql & "  or (V1.E043CRYINDB1<>'0' and V1.E043CRYRESB1='2')"          ' �����������сiB1)
'''''    sql = sql & "  or (V1.E043CRYINDB2<>'0' and V1.E043CRYRESB2='2') "         ' �����������сiB2�j
'''''    sql = sql & "  or (V1.E043CRYINDB3<>'0' and V1.E043CRYRESB3='2') "         ' �����������сiB3)
'''''    sql = sql & "  or (V1.E043CRYINDL1<>'0' and V1.E043CRYRESL1='2') "         ' �����������сiL1)
'''''    sql = sql & "  or (V1.E043CRYINDL2<>'0' and V1.E043CRYRESL2='2') "         ' �����������сiL2)
'''''    sql = sql & "  or (V1.E043CRYINDL3<>'0' and V1.E043CRYRESL3='2') "         ' �����������сiL3)
'''''    sql = sql & "  or (V1.E043CRYINDL4<>'0' and V1.E043CRYRESL4='2') "         ' �����������сiL4)
'''''    sql = sql & "  or (V1.E043CRYINDCS<>'0' and V1.E043CRYRESCS='2') "         ' �����������сiCs)
'''''    sql = sql & "  or (V1.E043CRYINDGD<>'0' and V1.E043CRYRESGD='2') "         ' �����������сiGD)
'''''    sql = sql & "  or (V1.E043CRYINDT<>'0' and V1.E043CRYREST='2') "           ' �����������сiT)
'''''    sql = sql & "  or (V1.E043CRYINDEP<>'0' and V1.E043CRYRESEP='2')) ) as J "         ' �����������сiEPD)
'''''    sql = sql & " from "
'''''    sql = sql & " TBCME040 B, TBCME041 H, TBCME018 S"
'''''    sql = sql & " where "
'''''    sql = sql & " B.CRYNUM = H.CRYNUM "
'''''    sql = sql & " and H.HINBAN = S.HINBAN "
'''''    sql = sql & " and H.REVNUM = S.MNOREVNO "
'''''    sql = sql & " and H.FACTORY = S.FACTORY "
'''''    sql = sql & " and H.OPECOND = S.OPECOND "
'''''
'''''                '�H���R�[�h�A��ԁA�敪�̏����w��
'''''    sql = sql & " and B.NOWPROC='CC600' "
'''''    sql = sql & " and B.LSTATCLS='T' "
'''''    sql = sql & " and B.RSTATCLS='T' "
'''''    sql = sql & " and B.DELCLS='0' "
'''''    'sql = sql & " and B.HOLDCLS='0' " ' �z�[���h�u���b�N���擾
'''''                '�u���b�N���Ɋ܂܂��i�Ԃ�����
'''''    sql = sql & " and (( B.INGOTPOS >= H.INGOTPOS "
'''''    sql = sql & " and B.INGOTPOS < H.INGOTPOS + H.LENGTH ) "
'''''    sql = sql & " or ( B.INGOTPOS + B.LENGTH > H.INGOTPOS "
'''''    sql = sql & " and B.INGOTPOS + B.LENGTH < H.INGOTPOS + H.LENGTH  ) "
'''''    sql = sql & " or ( H.INGOTPOS >= B.INGOTPOS "
'''''    sql = sql & " and H.INGOTPOS < B.INGOTPOS + B.LENGTH ) "
'''''    sql = sql & " or ( H.INGOTPOS + H.LENGTH > B.INGOTPOS "
'''''    sql = sql & " and H.INGOTPOS + H.LENGTH < B.INGOTPOS + B.LENGTH )) "
'''''                '�w����0�łȂ����т�0�łȂ��T���v�����㉺�Q�����邩
'''''    sql = sql & " and 2=( select count(*) "
'''''    sql = sql & "  from VECME010 V2 "
'''''    sql = sql & "  where "
'''''    sql = sql & "  B.BLOCKID=V2.E040BLOCKID"
'''''    sql = sql & "  and (V2.E043CRYINDRS='0' or V2.E043CRYRESRS<>'0') "         ' �����������сiRs)
'''''    sql = sql & "  and (V2.E043CRYINDOI='0' or V2.E043CRYRESOI<>'0') "         ' �����������сiOi)
'''''    sql = sql & "  and (V2.E043CRYINDB1='0' or V2.E043CRYRESB1<>'0')"          ' �����������сiB1)
'''''    sql = sql & "  and (V2.E043CRYINDB2='0' or V2.E043CRYRESB2<>'0') "         ' �����������сiB2�j
'''''    sql = sql & "  and (V2.E043CRYINDB3='0' or V2.E043CRYRESB3<>'0') "         ' �����������сiB3)
'''''    sql = sql & "  and (V2.E043CRYINDL1='0' or V2.E043CRYRESL1<>'0') "         ' �����������сiL1)
'''''    sql = sql & "  and (V2.E043CRYINDL2='0' or V2.E043CRYRESL2<>'0') "         ' �����������сiL2)
'''''    sql = sql & "  and (V2.E043CRYINDL3='0' or V2.E043CRYRESL3<>'0') "         ' �����������сiL3)
'''''    sql = sql & "  and (V2.E043CRYINDL4='0' or V2.E043CRYRESL4<>'0') "         ' �����������сiL4)
'''''    sql = sql & "  and (V2.E043CRYINDCS='0' or V2.E043CRYRESCS<>'0') "         ' �����������сiCs)
'''''    sql = sql & "  and (V2.E043CRYINDGD='0' or V2.E043CRYRESGD<>'0') "         ' �����������сiGD)
'''''    sql = sql & "  and (V2.E043CRYINDT='0' or V2.E043CRYREST<>'0') "           ' �����������сiT)
'''''    sql = sql & "  and (V2.E043CRYINDEP='0' or V2.E043CRYRESEP<>'0') )"         ' �����������сiEPD)
'''''    sql = sql & " order by B.BLOCKID, H.INGOTPOS "
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
'''''                    .INGOTPOS = rs("ss")
''''''                    .LENGTH = rs("LENGTH")      ' ����
'''''                    .BLOCKID = rs("BLOCKID")   ' �u���b�NID
'''''                    .UPDDATE = rs("UPDDATE")   ' �X�V���t
'''''                    .HOLDCLS = rs("HOLDCLS")   ' �z�[���h�敪
'''''                    BlockIdBuf = records(j).BLOCKID
'''''                    .HSXTYPE = rs("HSXTYPE")
'''''                    .HSXCDIR = rs("HSXCDIR")
'''''                    If rs("J") > 0 Then
'''''
'''''                        .Judg = "2"
'''''                    Else
'''''                        .Judg = "1"
'''''                    End If
'''''
'''''                End With
'''''                k = 1
'''''            End If
'''''
'''''            '�i�Ԃ̊i�[
'''''            ReDim Preserve records(j).HIN(k)
'''''            records(j).HIN(k).hinban = rs("HINBAN")
'''''            records(j).HIN(k).mnorevno = rs("REVNUM")
'''''            records(j).HIN(k).factory = rs("FACTORY")
'''''            records(j).HIN(k).opecond = rs("OPECOND")
'''''            k = k + 1
'''''            rs.MoveNext
'''''        Next i
'''''        rs.Close
'''''
'''''    End If
'''''
'''''
'''''    '�w���P�������ю擾
'''''    If getKouBlock(records(), "CC600") = FUNCTION_RETURN_FAILURE Then
'''''       DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
'''''       GoTo PROC_EXIT
'''''    End If
'''''
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
'''''    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function



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
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp3"


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



''''''�T�v    :�҂��ꗗ �����\���p�c�a�h���C�o�i�����w���҂��j
''''''���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
''''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
''''''        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
''''''����    :
''''''����    :2001/07/06 ���{ �쐬
'''''Public Function DBDRV_scmzc_fcmkc001b_Disp4(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''
'''''    '�������w���҂���
'''''    'CC710�̂���
'''''
'''''    '�u���b�NID��X�V���t�擾
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp4"
'''''
'''''    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_SUCCESS
'''''
'''''
'''''    '�u���b�NID��X�V���t�A�i�ԓ��擾
'''''    If getBlockID(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
'''''        GoTo PROC_EXIT
'''''    End If
'''''
''''''2000/08/24 S.Sano Start
''''''    '�w���P�������ю擾
''''''    If getKouBlock(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
''''''       DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
''''''       GoTo proc_exit
''''''    End If
''''''2000/08/24 S.Sano End
'''''
'''''
'''''PROC_EXIT:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    '�G���[�n���h��
'''''    gErr.HandleError
'''''    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
'''''    Resume PROC_EXIT
'''''End Function



''''''�w���P�����p
'''''Private Function getKouBlock(records() As type_DBDRV_scmzc_fcmkc001b_Disp, NOWPROC As String) As FUNCTION_RETURN
'''''
'''''    Dim sql As String       'SQL�S��
'''''    Dim rs As OraDynaset    'RecordSet
'''''    Dim recCnt As Long
'''''    Dim motoRecCnt As Long
'''''    Dim i As Long
'''''
'''''    '�G���[�n���h���̐ݒ�
'''''    On Error GoTo PROC_ERR
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
'''''    sql = sql & " from TBCME040 B,TBCMG002 K "
'''''    sql = sql & " where B.BLOCKID=K.CRYNUM "
'''''    sql = sql & " and substr(B.BLOCKID,1,1)='8' "
'''''    sql = sql & " and B.NOWPROC='" & NOWPROC & "' "
'''''    sql = sql & " and B.LSTATCLS='T' "
'''''    sql = sql & " and B.RSTATCLS='T' "
'''''    sql = sql & " and B.DELCLS='0' "
'''''    'sql = sql & " and B.HOLDCLS='0' " ' �z�[���h�u���b�N���擾
'''''    sql = sql & " and K.TRANCNT=any(select max(TRANCNT) from TBCMG002 where CRYNUM=B.BLOCKID ) "
'''''    sql = sql & " order by B.BLOCKID "
'''''
'''''
'''''    '�f�[�^�𒊏o����
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    If rs.RecordCount = 0 Then
'''''        rs.Close
'''''        GoTo PROC_EXIT
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
'''''PROC_EXIT:
'''''    '�I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    '�G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    getKouBlock = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''
'''''End Function



'�����֐� �u���b�NID�A�X�V���t�擾�i���o�҂��A�����w���҂��p�j
Private Function getBlockID(records() As type_DBDRV_scmzc_fcmkc001b_Disp, _
                            NOWPROC As String) As FUNCTION_RETURN

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim recCnt      As Long         '���R�[�h��
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim BlockIdBuf  As String
    Dim lLp         As Long         '2007/08/30 SPK Tsutsumi Add
    Dim sBakPos     As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function getBlockID"

    getBlockID = FUNCTION_RETURN_SUCCESS

    'sql = "select "
    'sql = sql & " X.INPOSC2, "
    'sql = sql & " X.GNLC2, "
    'sql = sql & " X.HOLDBC2, "      '2005/08
    'sql = sql & " X.HOLDCC2, "      '2005/08
    'sql = sql & " X.HOLDKTC2, "      '2005/08
    'sql = sql & " X.LBLFLGC2,"       '2005/11
    'sql = sql & " V.E040CRYNUM, "
    'sql = sql & " V.E040BLOCKID, "
    'sql = sql & " V.E040INGOTPOS, "
    'sql = sql & " V.E040UPDDATE, "
    'sql = sql & " V.E040HOLDCLS, "
    'sql = sql & " V.E041HINBAN, "            ' �i��
    'sql = sql & " V.E041REVNUM, "            ' ���i�ԍ������ԍ�
    'sql = sql & " V.E041FACTORY, "           ' �H��
    'sql = sql & " V.E041OPECOND, "           ' ���Ə���
    'sql = sql & " V.E041LENGTH, "            ' ����  2006/02
    'sql = sql & " V.E037DIAMETER, "          ' ���a  2006/02
    'sql = sql & " S.HSXTYPE, "           ' �i�r�w�^�C�v
    'sql = sql & " S.HSXCDIR "            ' �i�r�w�����ʕ���
    'sql = sql & ",XC1.PUPTNC1 "          ' ���������(2004/12/21) kubota
    'sql = sql & " from "
    ''sql = sql & " VECME009 V, TBCME018 S "
    'sql = sql & " VECME009 V, TBCME018 S , XSDC2 X "
    
    'sql = sql & ",XSDC1 XC1 "                       '��������ݒǉ��Ή�(2004/12/21) kubota
    
    'sql = sql & " where "
    'sql = sql & " V.E040BLOCKID = X.CRYNUMC2 "
    'sql = sql & " and V.E041HINBAN = S.HINBAN "
    'sql = sql & " and V.E041REVNUM = S.MNOREVNO "
    'sql = sql & " and V.E041FACTORY = S.FACTORY "
    'sql = sql & " and V.E041OPECOND = S.OPECOND "
    'sql = sql & " and V.E040NOWPROC='" & NOWPROC & "' "
    'sql = sql & " and V.E040LSTATCLS='T' "
    'sql = sql & " and V.E040RSTATCLS='T' "
    'sql = sql & " and V.E040DELCLS='0' "
    ''sql = sql & " and V.E040HOLDCLS='0' " ' �z�[���h�u���b�N���擾
    'sql = sql & " and X.XTALC2 = XC1.XTALC1(+) "    '��������ݒǉ��Ή�(2004/12/21) kubota
    'sql = sql & " order by V.E040BLOCKID, V.E041INGOTPOS "
    
    'VIEW --> CA �ύX  2006/02
    sql = "select "
    sql = sql & " X.XTALC2, "
    sql = sql & " X.INPOSC2, "
    sql = sql & " X.GNLC2, "
    sql = sql & " X.HOLDBC2, "      '2005/08
    sql = sql & " X.HOLDCC2, "      '2005/08
    sql = sql & " X.HOLDKTC2, "      '2005/08
    sql = sql & " X.LBLFLGC2,"       '2005/11
    sql = sql & " CA.CRYNUMCA, "
    sql = sql & " CA.INPOSCA, "
    sql = sql & " CA.KDAYCA, "
    sql = sql & " CA.HOLDBCA, "
    sql = sql & " CA.HINBCA, "           ' �i��
    sql = sql & " CA.REVNUMCA, "         ' ���i�ԍ������ԍ�
    sql = sql & " CA.FACTORYCA, "        ' �H��
    sql = sql & " CA.OPECA, "            ' ���Ə���
    sql = sql & " CA.GNLCA, "            ' ����  2006/02
    sql = sql & " CA.GNWCA, "            ' �d��  2006/02
    sql = sql & " S.HSXTYPE, "           ' �i�r�w�^�C�v
    sql = sql & " S.HSXCDIR "            ' �i�r�w�����ʕ���
    sql = sql & ",XC1.PUPTNC1 "          ' ���������(2004/12/21) kubota
    sql = sql & ",X.KIKBNC2 "            ' �����ʋ敪 2006/11/14 SETsw kubota
    sql = sql & ",X.PLANTCATC2 "         ' ���� 2007/08/30 SPK Tsutsumi Add
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/07/01
    '' ������~���ڒǉ� add SETkimizuka Start  09/03/26
    'sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUS "
    'sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
    'sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSE "
    'sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNO "
    '' ������~���ڒǉ� add SETkimizuka End    09/03/26
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUS "
    sql = sql & " , NVL(TO_CHAR(Y4.STOPY4),' ') as STOP "
    sql = sql & " , DECODE(TRIM(Y4.CAUSEY4),NULL,' ',TRIM(Y4.CAUSEY4) || ':' || A9.NAMEJA9) as CAUSE "
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNO "
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/07/01
    sql = sql & " from "
    'sql = sql & " VECME009 V, TBCME018 S , XSDC2 X "
    sql = sql & " XSDCA CA, TBCME018 S , XSDC2 X "
    
    sql = sql & ",XSDC1 XC1 "                       '��������ݒǉ��Ή�(2004/12/21) kubota
    
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/07/01
    sql = sql & "    ,XODY3 Y3,XODY4 Y4,KODA9 A9 "
    '' ������~���ڒǉ� add SETkimizuka Start  09/03/26
    'sql = sql & "    ,(SELECT XTALNOY3 as XTALNO,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
    'sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
    'sql = sql & "      FROM XODY3  "
    'sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND STOPY4 <> '2'  AND WKKTY4 in " & CreateWkktSQL(WATCH_PROCCD) & ") "
    'sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
    'sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
    'sql = sql & "                FROM XODY3,XODY4,XODY5 "
    'sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
    'sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
    'sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
    'sql = sql & "      WHERE  "
    'sql = sql & "       LIVKY3    = '0' "
    'sql = sql & "       GROUP BY XTALNOY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9) Y4 "
    '' ������~���ڒǉ� add SETkimizuka End  09/03/26
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/07/01
    sql = sql & " where "
    sql = sql & " CA.CRYNUMCA = X.CRYNUMC2 "
    sql = sql & " and CA.HINBCA = S.HINBAN "
    sql = sql & " and CA.REVNUMCA = S.MNOREVNO "
    sql = sql & " and CA.FACTORYCA = S.FACTORY "
    sql = sql & " and CA.OPECA = S.OPECOND "
    sql = sql & " and CA.GNWKNTCA ='" & NOWPROC & "' "
    sql = sql & " and CA.LIVKCA='0' "
    sql = sql & " and X.XTALC2 = XC1.XTALC1(+) "    '��������ݒǉ��Ή�(2004/12/21) kubota
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/07/01
    'sql = sql & " AND CA.CRYNUMCA    = Y4.XTALNO(+) "            'add 09/03/26 SETkimizuka
    sql = sql & " AND CA.CRYNUMCA = Y3.XTALNOY3(+) "
    sql = sql & " AND Y3.LIVKY3(+) = '0' "
    sql = sql & " AND Y4.LIVKY4(+) = '0' "
    sql = sql & " AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & " AND Y3.RCNTY3 = Y4.RCNTY4(+) "
    sql = sql & " AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/07/01

    sql = sql & " order by CA.CRYNUMCA, CA.INPOSCA "

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
        If rs("CRYNUMCA") <> BlockIdBuf Then
        
            j = j + 1
            ReDim Preserve records(j)
            
            With records(j)
                '.CRYNUM = rs("E040CRYNUM")
                '.INGOTPOS = rs("E040INGOTPOS")
                '.BLOCKID = rs("E040BLOCKID")   ' �u���b�NID
                '.UPDDATE = rs("E040UPDDATE")   ' �X�V���t
                '.HOLDCLS = rs("E040HOLDCLS")   ' �z�[���h�敪
                .CRYNUM = rs("XTALC2")
                .INGOTPOS = rs("INPOSCA")
                .BLOCKID = rs("CRYNUMCA")   ' �u���b�NID
                .UPDDATE = rs("KDAYCA")   ' �X�V���t
                .HOLDCLS = rs("HOLDBCA")   ' �z�[���h�敪
                BlockIdBuf = records(j).BLOCKID
                .HSXTYPE = rs("HSXTYPE")
                .HSXCDIR = rs("HSXCDIR")
                .Judg = " "
                .INPOS = rs("INPOSC2")
                .LENGTH = rs("GNLC2")
                .PUPTN = rs("PUPTNC1")         ' ���������(2004/12/21) kubota
                .HOLDB = rs("HOLDBC2")
                .HOLDC = rs("HOLDCC2")
                If IsNull(rs("HOLDKTC2")) = False Then .HOLDKT = rs("HOLDKTC2")    '2005/08
                If IsNull(rs("LBLFLGC2")) = False Then .LBLFLG = rs("LBLFLGC2")    '2005/11
                '.DIA = rs("E037DIAMETER")   '2006/02
                If IsNull(rs("KIKBNC2")) = False Then .KIKBN = rs("KIKBNC2")       ' �����ʋ敪 2006/11/14 SETsw kubota
            End With
            
            k = 1
            sBakPos = ""    'add 09/03/26 SETkimizuka
        End If
        
        
        If InStr(sBakPos, Trim(rs("INPOSCA"))) = 0 Then '�����Ď����ڒǉ��ɔ����C�� add
            '�i�Ԃ̊i�[
            ReDim Preserve records(j).hinM(k)
            records(j).hinM(k).HIN.hinban = rs("HINBCA")
            records(j).hinM(k).HIN.mnorevno = rs("REVNUMCA")
            records(j).hinM(k).HIN.factory = rs("FACTORYCA")
            records(j).hinM(k).HIN.opecond = rs("OPECA")
            records(j).hinM(k).LENGTH = rs("GNLCA")
            records(j).hinM(k).Weight = rs("GNWCA")
            
            ' ���� 2007/08/30 SPK Tsutsumi Start
            If IsNull(rs("PLANTCATC2")) = False Then
                For lLp = 0 To UBound(s_Mukesaki)
                    If rs("PLANTCATC2") = s_Mukesaki(lLp).sMukeCode Then
                        records(j).hinM(k).HIN.sMukesaki = s_Mukesaki(lLp).sMukeName
                    End If
                Next lLp
            End If
            sBakPos = sBakPos & Trim(rs("INPOSCA")) & " "
            k = k + 1
        End If
'        '�i�Ԃ̊i�[
'        ReDim Preserve records(j).hinM(k)
'        records(j).hinM(k).HIN.hinban = rs("HINBCA")
'        records(j).hinM(k).HIN.mnorevno = rs("REVNUMCA")
'        records(j).hinM(k).HIN.Factory = rs("FACTORYCA")
'        records(j).hinM(k).HIN.OpeCond = rs("OPECA")
'        records(j).hinM(k).LENGTH = rs("GNLCA")
'        records(j).hinM(k).Weight = rs("GNWCA")
        
        ' ���� 2007/08/30 SPK Tsutsumi End

        ' �����Ď�SQL�C�� upd SETkimizuka Start  09/07/01
        ' ������~���ڒǉ� add SETkimizuka Start  09/03/26
        'records(j).STOP = rs("STOP")                   '��~�敪
        'records(j).AGRSTATUS = rs("AGRSTATUS")       '���F�m�F�敪
        'If Trim(rs("CAUSE")) <> "" And InStr(records(j).CAUSE, Trim(rs("CAUSE"))) = 0 Then
        '    records(j).CAUSE = records(j).CAUSE & rs("CAUSE") & vbTab       '��~���R
        'End If
        If rs("STOP") <> "2" And rs("WKKTY4") = "CC700" Then
           If Trim(records(j).AGRSTATUS) = "" Or (rs("AGRSTATUS") < records(j).AGRSTATUS) Then
                records(j).STOP = rs("STOP")                   '��~�敪
                records(j).AGRSTATUS = rs("AGRSTATUS")       '���F�m�F�敪
           End If
            If Trim(rs("CAUSE")) <> "" And InStr(records(j).CAUSE, Trim(rs("CAUSE"))) = 0 Then
                records(j).CAUSE = records(j).CAUSE & rs("CAUSE") & vbTab       '��~���R
            End If
        End If
        If Trim(rs("PRINTNO")) <> "" And InStr(records(j).PRINTNO, Trim(rs("PRINTNO"))) = 0 Then
            records(j).PRINTNO = records(j).PRINTNO & rs("PRINTNO") & vbTab       '��s�]��
        End If
        ' ������~���ڒǉ� add SETkimizuka End    09/03/26

'        k = k + 1
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

'---- ADD [���������Ǘ��A�����H�����э쐬�����ǉ�] �ȉ��ǉ��֐��@START ---- TCS)T.TERAUCHI

'�T�v      :���������Ǘ��쐬�p
'���Ұ�    :�ϐ���        ,IO ,�^                           ,����
'          :Xodcx         ,I  ,type_DBDRV_fcmkc001c_InsXodcx,�u���b�N�Ǘ��̌��݊Ǘ��H���A���ݍH���A�ŏI�ʉߊǗ��H���A�ŏI�ʉߍH���X�V�p
'          :�߂�l        ,O  ,FUNCTION_RETURN              ,
'����      :
Public Function DBDRV_fcmkc001c_InsXODCX(Xodcx As type_DBDRV_fcmkc001c_InsXodcx) As FUNCTION_RETURN

    Dim sSql        As String
    Dim objDS       As Object
    Dim sDbName     As String
    Dim sErrMsg     As String
    Dim dCyokkei    As Double
    Dim sDopType    As String
    Dim sCSDop      As String       'CS�h�[�v�L��
    Dim sNDop       As String       '���f�h�[�v�L��
    Dim sLTUmu      As String
    Dim sSCNTRL     As String       '���ʺ��۰ٺ��� ADD 2011/03/24 TSMC�i���ʑΉ�
    
'*** UPDATE START TAGAWA 2004/12/16
    Dim sFlag       As String
'*** UPDATE END  TAGAWA 2004/12/16

'�G���[�n���h���̐ݒ�
On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_fcmkc001c_InsXODCX"
    
    DBDRV_fcmkc001c_InsXODCX = FUNCTION_RETURN_FAILURE
    
    '****** �o�^���̎擾 ******
    '' ����������{���擾SQL�̍쐬
    sDbName = "XSDC1"
    Call GetAssistSQL_300(sSql, Xodcx.CRYNUM)
    If DynSet2(objDS, sSql) = False Then
        If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
        sErrMsg = GetMsgStr("EGET", sDbName)
        Exit Function
    End If
    '�Y���f�[�^�����̏ꍇ
    If objDS.EOF = True Then
        If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
        sErrMsg = GetMsgStr("ESM04", sDbName)
        Exit Function
    End If
    
    '' �d�ʎ擾
    dCyokkei = objDS.Fields("PRODMCX").Value                'MID���㒼�a
    Xodcx.Weight = WeightOfCylinder(dCyokkei, Xodcx.LENGTH) '�u���b�N�d��
    
    '' CS�h�[�v�L���A���f�h�[�v�L���ݒ�
    sDopType = UCase(NulltoStr(objDS.Fields("DTYPEC1").Value))
'*** UPDATE START TAGAWA 2004/12/16**************************
''    If sDopType = " " Or sDopType = "P" Then
''        sCSDop = "2"
''        sNDop = "1"
''    ElseIf sDopType = "N" Then
''        sCSDop = "1"
''        sNDop = "2"
''    Else
''        sCSDop = " "
''        sNDop = " "
''    End If
        ''�����h�[�v�擾
        sFlag = UCase(Trim(NulltoStr(objDS.Fields("DPNTCLS").Value)))
        ''Cs�h�[�v�̎�
        If sFlag = "C" Then
            sCSDop = "2"
            sNDop = "1"
        ''���f�h�[�v�̎�
        ElseIf sFlag = "N" Then
            sCSDop = "1"
            sNDop = "2"
        ''�v�h�[�v�̎�
        ElseIf sFlag = "M" Then
            sCSDop = "2"
            sNDop = "2"
        ''���̑�
        Else
            sCSDop = "1"
            sNDop = "1"
        End If
'*** UPDATE END TAGAWA 2004/12/16**************************

    ''���C�t�^�C���d�l�L��
    If objDS.Fields("HSXLTHWS").Value = "H" Then
        ''�L��
        sLTUmu = "2"
    Else
        ''����
        sLTUmu = "1"
    End If
    '*** UPDATE START Marushita 2011/03/24 TSMC�i���ʑΉ�
    ''���������`�F�b�N�t���O�̔��f
    If NulltoStr(objDS.Fields("MTRLCHKFLG").Value) = "1" Then
        ''�i��NULL���̎��ʃR���g���[���R�[�h�Z�b�g(��3��)
        If NulltoStr(objDS.Fields("HINBCX").Value) = "" Then
            sSCNTRL = "   "
        Else
            ''���ʃR���g���[���R�[�h�Z�b�g(�i��3��)
            sSCNTRL = left(objDS.Fields("HINBCX").Value, 3)
        End If
    Else
        ''���ʃR���g���[���R�[�h�Z�b�g(��3��)
        sSCNTRL = "   "
    End If
    '*** UPDATE END   Marushita 2011/03/24
    
    '****** ���������Ǘ��쐬���� ******
    sSql = ""
    sSql = sSql & "INSERT INTO xodcx(" & vbLf
    sSql = sSql & "crynumcx" & vbLf    ''�u���b�NID
    sSql = sSql & ",mtrlnumcx" & vbLf   ''������
    sSql = sSql & ",wkktcx" & vbLf      ''�H���R�[�h
    sSql = sSql & ",workcx" & vbLf      ''�H��R�[�h
    sSql = sSql & ",hdaycx" & vbLf      ''��������
    sSql = sSql & ",weightcx" & vbLf    ''�d��
    sSql = sSql & ",htkbncx" & vbLf     ''�p��/�K���敪
    sSql = sSql & ",divumucx" & vbLf    ''�����L��
    sSql = sSql & ",toworkcx" & vbLf    ''���o��H��R�[�h
    sSql = sSql & ",frworkcx" & vbLf    ''�����H��R�[�h
    sSql = sSql & ",hinbcx" & vbLf      ''�i��
    sSql = sSql & ",typecx" & vbLf      ''�^�C�v
    sSql = sSql & ",dptypecx" & vbLf    ''�h�[�v�^�C�v
    sSql = sSql & ",tposcx" & vbLf      ''�ʒuL(�g�b�v��)
    sSql = sSql & ",lencx" & vbLf       ''�u���b�N����
    sSql = sSql & ",siweightcx" & vbLf  ''�d���ݏd��
    sSql = sSql & ",updmcx" & vbLf      ''����AV�a
    sSql = sSql & ",prodmcx" & vbLf     ''���i�a
    sSql = sSql & ",tdopposcx" & vbLf   ''�ǉ��h�[�v�����ʒuL
    sSql = sSql & ",wdopumucx" & vbLf   ''W�h�[�v(P/N����)�L��
    sSql = sSql & ",csdopumucx" & vbLf  ''CS�h�[�v�L��
    sSql = sSql & ",ndopumucx" & vbLf   ''���f�h�[�v�L��
    sSql = sSql & ",ltspecumucx" & vbLf ''���C�t�^�C���d�l�L��
    sSql = sSql & ",csspecumucx" & vbLf ''CS�d�l�L��
    sSql = sSql & ",topwcx" & vbLf      ''�g�b�vWT
    sSql = sSql & ",dmkcx" & vbLf       ''���a�敪
    sSql = sSql & ",xtalcx" & vbLf      ''�����ԍ�
    sSql = sSql & ",livkcx" & vbLf      ''�����敪
    sSql = sSql & ",unifgcx" & vbLf     ''����FLG
    sSql = sSql & ",twarifgcx" & vbLf   ''�c��FLG
    sSql = sSql & ",refusefgcx" & vbLf  ''�����FLG
    sSql = sSql & ",tstafidcx" & vbLf   ''�o�^�Ј�ID
    sSql = sSql & ",tdaycx" & vbLf      ''�o�^���t
    sSql = sSql & ",kstafidcx" & vbLf   ''�X�V��
    sSql = sSql & ",kdaycx" & vbLf      ''�X�V����
    sSql = sSql & ",crydopcx" & vbLf    ''�����h�[�v
    sSql = sSql & ",crydopvlcx" & vbLf  ''�����h�[�v��
    sSql = sSql & ",bkformcx" & vbLf    ''�u���b�N�`��
    sSql = sSql & ",pgidcx" & vbLf      ''PG-ID
    sSql = sSql & ",blktypcx" & vbLf    ''�u���b�N���
    sSql = sSql & ",tkacutwcx" & vbLf   ''T�T���v���O�d��
'*** UPDATE START TAGAWA 2004/12/16***************
    sSql = sSql & ",denflgcx" & vbLf     ''�d�ɍރt���O
'*** UPDATE END   TAGAWA 2004/12/16***************
    sSql = sSql & ",toptwcx" & vbLf     ''į�ߎ�o��WT
'*** UPDATE START Marushita 2011/03/24 TSMC�i���ʑΉ�
    sSql = sSql & ",scntrlcx" & vbLf    ''���ʺ��۰ٺ���
'*** UPDATE END   Marushita 2011/03/24
    sSql = sSql & ")values(" & vbLf
    sSql = sSql & "'" & Xodcx.BLOCKID & "0" & "'" & vbLf                    ''��ۯ�ID
    sSql = sSql & ",' '" & vbLf                                             ''�����ԍ�
    sSql = sSql & ",'" & Right(PROCD_KOUNYU_TAN_KESSYOU, 4) & "'" & vbLf    ''�H���R�[�h('B410')
    sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''�H�꺰��
    sSql = sSql & ",sysdate" & vbLf                                         ''��������
    sSql = sSql & "," & Xodcx.Weight & vbLf                                 ''�d��
    sSql = sSql & ",'1'" & vbLf                                             ''�p���E�K���敪
    sSql = sSql & ",'1'" & vbLf                                             ''�����L��
    sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''���o��H�꺰��
    sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''�����H�꺰��
    sSql = sSql & ",'" & objDS.Fields("HINBCX").Value & "'" & vbLf          ''�i��
    sSql = sSql & ",'" & objDS.Fields("HSXTYPE").Value & "'  " & vbLf       ''�^�C�v
    sSql = sSql & ",'" & sDopType & "'" & vbLf                              ''�h�[�v�^�C�v
    sSql = sSql & ", " & Xodcx.INGOTPOS & vbLf                              ''�u���b�N�Ǘ���������J�n�ʒu
    sSql = sSql & ", " & Xodcx.LENGTH & vbLf                                ''�u���b�N�Ǘ������
    sSql = sSql & ", " & ConvNum(objDS.Fields("SUICHARGE").Value) & vbLf    ''�d���ݏd��
    sSql = sSql & ", " & ConvNum(objDS.Fields("UPDMCX").Value) & vbLf       ''����AV�a
    sSql = sSql & ", " & ConvNum(objDS.Fields("PRODMCX").Value) & vbLf      ''���i�a
    sSql = sSql & ", " & ConvNum(objDS.Fields("ADDOPPC1").Value) & vbLf     ''�ǉ��h�[�v�����ʒuL
    sSql = sSql & ",'1'" & vbLf                                             ''W�h�[�v(P/N����)�L��
    sSql = sSql & ",'" & sCSDop & "'" & vbLf                                ''CS�h�[�v�L��
    sSql = sSql & ",'" & sNDop & "'" & vbLf                                 ''���f�h�[�v�L��
    sSql = sSql & ",'" & sLTUmu & "'" & vbLf                                ''���C�t�^�C���g�p�L��
    sSql = sSql & ",'2'" & vbLf                                             ''CS�g�p�L��
    sSql = sSql & "," & ConvNum(objDS.Fields("CTR01A9").Value) & vbLf       '�f�g�b�vWT
    sSql = sSql & ",'300'" & vbLf                                           ''���a�敪
    sSql = sSql & ",'" & Xodcx.CRYNUM & "'" & vbLf                          ''�����ԍ�
    sSql = sSql & ",'0'" & vbLf                                             ''�����敪
    sSql = sSql & ",'0'" & vbLf                                             ''����FLG
    sSql = sSql & ",'0'" & vbLf                                             ''�c��FLG
    sSql = sSql & ",'0'" & vbLf                                             ''�����FLG
    sSql = sSql & ",'" & Xodcx.STAFFID & "'" & vbLf                         ''�o�^�Ј�ID
    sSql = sSql & ",sysdate" & vbLf                                         ''�o�^���t
    sSql = sSql & ",'" & Xodcx.STAFFID & "'" & vbLf                         ''�X�V�Ј�ID
    sSql = sSql & ",sysdate" & vbLf                                         ''�X�V���t
    sSql = sSql & ",'" & objDS.Fields("DPNTCLS").Value & "'" & vbLf         ''�����h�[�v
    sSql = sSql & "," & ConvNum(objDS.Fields("DOPANT").Value) & vbLf        ''�����h�[�v��
    sSql = sSql & ",'3'" & vbLf                                             ''�u���b�N�`��
    sSql = sSql & ",'" & objDS.Fields("PGID").Value & "'" & vbLf            ''PG-ID
    sSql = sSql & ",'A'" & vbLf                                             ''�u���b�N���
    sSql = sSql & ",0" & vbLf                                               ''T�T���v���O�d��
'*** UPDATE START TAGAWA 2004/12/16***************
    sSql = sSql & ",'1'" & vbLf                                             ''�d�ɍރt���O
'*** UPDATE END   TAGAWA 2004/12/16***************
    sSql = sSql & "," & ConvNum(objDS.Fields("PUTCUTWC1").Value) & vbLf     ''į�ߎ�o��WT
'*** UPDATE START Marushita 2011/03/24 TSMC�i���ʑΉ�
    sSql = sSql & ",'" & sSCNTRL & "'" & vbLf                               ''���ʺ��۰ٺ���
'*** UPDATE END   Marushita 2011/03/24
    
    sSql = sSql & ")"
        
    If 0 >= OraDB.ExecuteSQL(sSql) Then
        DBDRV_fcmkc001c_InsXODCX = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_fcmkc001c_InsXODCX = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    DBDRV_fcmkc001c_InsXODCX = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�����H�����э쐬�p
'���Ұ�    :�ϐ���        ,IO ,�^                           ,����
'          :Xodcx         ,I   ,type_DBDRV_scmzc_fcmkc001c_InsXodcx         ,�u���b�N�Ǘ��̌��݊Ǘ��H���A���ݍH���A�ŏI�ʉߊǗ��H���A�ŏI�ʉߍH���X�V�p
'          :�߂�l        ,O  ,FUNCTION_RETURN              ,
'����      :
'����      :2004/12/04 �V�K�쐬 TCS)T.TERAUCHI
Public Function DBDRV_fcmkc001c_InsXODB3(Xodcx As type_DBDRV_fcmkc001c_InsXodcx) As FUNCTION_RETURN
    Dim objDS       As Object
    Dim sSql        As String
    Dim iRenban     As Integer
    Dim sYear       As String
    Dim sMonth      As String
    Dim sDay        As String
    Dim sHour       As String
    Dim sMin        As String
    Dim sNowdate    As String
    Dim sCyoku      As String

'' �G���[�n���h���̐ݒ�
On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_fcmkc001c_InsXODB3"

    DBDRV_fcmkc001c_InsXODB3 = FUNCTION_RETURN_FAILURE
    
    '***** �o�^���̎擾 *****
    '' �H���R�[�h�ݒ�
    nowCd = Xodcx.LASTPASS
    
    '' �V�X�e�����t�A���ѓ��t���̐ݒ�
    If Not GetSysdate Then
        GoTo proc_exit
    End If
    sNowdate = gsSysdate
    
    '�T�[�o�[�V�X�e�����t�����ѓ��ɕύX
    sNowdate = GetJITUDATE(Format(sNowdate, "yyyymmddhhmmss"))
    
    '���ѓ���蒼�敪�𔻒�
    sCyoku = GetCYOKU(gsSysdate)
    
    '���ѓ�����؂���
    sYear = Mid(sNowdate, 1, 4)     '�N
    sMonth = Mid(sNowdate, 5, 2)    '��
    sDay = Mid(sNowdate, 7, 2)      '��
    sHour = Mid(sNowdate, 9, 2)     '��
    sMin = Mid(sNowdate, 11, 2)     '��

    iRenban = 0

    '' �H���A�Ԃ̎擾
    sSql = ""
    sSql = sSql & " SELECT NVL(MAX(kcntb3),0) maxcnt        " & vbLf   '�H���A��
    sSql = sSql & " FROM   xodb3                            " & vbLf
    sSql = sSql & " WHERE  polnob3 = '" & Xodcx.BLOCKID & "0" & "'" & vbLf
    
    'SQL�����s
    If DynSet2(objDS, sSql) = False Then
        If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
        GoTo proc_exit
    End If
    
    '�擾�����f�[�^���i�[
    iRenban = objDS.Fields("maxcnt").Value + 1
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing

    '****** �����H������(XODB3)�X�V ******
    sSql = ""
    sSql = sSql & "insert into XODB3(                       " & vbLf
    sSql = sSql & "             POLNOB3                     " & vbLf   '�����ԍ�
    sSql = sSql & "            ,KCNTB3                      " & vbLf   '�H���A��
    sSql = sSql & "            ,CRSEQB3                     " & vbLf   '�����A��
    sSql = sSql & "            ,TDAYB3                      " & vbLf   '�o�^���t
    sSql = sSql & "            ,RDAYB3                      " & vbLf   '�C�����t
    sSql = sSql & "            ,SDAYB3                      " & vbLf   '���M���t
    sSql = sSql & "            ,SNDKB3                      " & vbLf   '���M�敪
    sSql = sSql & "            ,SAKJB3                      " & vbLf   '�폜�敪
    sSql = sSql & "            ,POKUBB3                     " & vbLf   '�����敪
    sSql = sSql & "            ,POKIDCB3                    " & vbLf   '������ރR�[�h
    sSql = sSql & "            ,POLTNB3                     " & vbLf   '�������b�gNo
    sSql = sSql & "            ,MODKBB3                     " & vbLf   '�ԍ��敪
    sSql = sSql & "            ,SUMKBB3                     " & vbLf   '�W�v�敪
    sSql = sSql & "            ,WKKTB3                      " & vbLf   '�H���R�[�h
    sSql = sSql & "            ,PLACB3                      " & vbLf   '���C���R�[�h
    sSql = sSql & "            ,FRWB3                       " & vbLf   '����d��
    sSql = sSql & "            ,TOWB3                       " & vbLf   '���o�d��
    sSql = sSql & "            ,LOSWB3                      " & vbLf   '���X�d��
    sSql = sSql & "            ,FRWKKTB3                    " & vbLf   '����H���R�[�h
    sSql = sSql & "            ,TOWKKTB3                    " & vbLf   '���o�H���R�[�h
    sSql = sSql & "            ,TOWKKBB3                    " & vbLf   '���o�敪
    sSql = sSql & "            ,TOWORKB3                    " & vbLf   '���o�H��R�[�h
    sSql = sSql & "            ,TOPLACB3                    " & vbLf   '���o���C���R�[�h
    sSql = sSql & "            ,CHGNB3                      " & vbLf   '�`���[�WNo
    sSql = sSql & "            ,EYYB3                       " & vbLf   '���ѓ��t(�N)
    sSql = sSql & "            ,EMMB3                       " & vbLf   '���ѓ��t(��)
    sSql = sSql & "            ,EDDB3                       " & vbLf   '���ѓ��t(��)
    sSql = sSql & "            ,ECYOKB3                     " & vbLf   '���敪
    sSql = sSql & "            ,EHHB3                       " & vbLf   '���ю���(��)
    sSql = sSql & "            ,EMIB3�@                     " & vbLf   '���ю���(��)
    sSql = sSql & "            ,MANB3                       " & vbLf   '�S����
    sSql = sSql & "            ,MANJB3                      " & vbLf   '�S���Җ�
    sSql = sSql & "            ,DENKB3                      " & vbLf   '�Z�x�敪
    sSql = sSql & "            ,DENSITYB3                   " & vbLf   '�Z�x�l
    sSql = sSql & "            ,GSNDFLGB3                   " & vbLf   '�������M�t���O
    sSql = sSql & "            ,HFLGB3                      " & vbLf   '�����t���O
    sSql = sSql & "            ,htkbnb3                     " & vbLf   '���i�敪
    sSql = sSql & "            ,plworkb3                    " & vbLf   '�g�p�\��H��
    sSql = sSql & "            ,mdensityb3                  " & vbLf   '���Z�x�l
    sSql = sSql & "            ,gsdayb3                     " & vbLf
    sSql = sSql & ")VALUES(                                 " & vbLf
    sSql = sSql & " '" & Xodcx.BLOCKID & "0" & "'                 " & vbLf  ' �����ԍ�
    sSql = sSql & "," & iRenban & "                         " & vbLf   '�H���A��
    sSql = sSql & ",1                                       " & vbLf   '�����A��
    sSql = sSql & ",to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf '�o�^���t
    sSql = sSql & ",null                                    " & vbLf   '�C�����t
    sSql = sSql & ",null                                    " & vbLf   '���M���t
    sSql = sSql & ",' '                                     " & vbLf   '���M�敪
    sSql = sSql & ",'0'                                     " & vbLf   '�폜�敪
    sSql = sSql & ",'2'                                     " & vbLf   '�����敪
    sSql = sSql & ",'888'                                   " & vbLf   '������ރR�[�h
    sSql = sSql & ",' '                                     " & vbLf   '�������b�g�ԍ�
    sSql = sSql & ",' '                                     " & vbLf   '�ԍ��敪
    sSql = sSql & ",' '                                     " & vbLf   '�W�v�敪
    sSql = sSql & ",'" & Right(nowCd, 4) & "'               " & vbLf   '�H���R�[�h
    sSql = sSql & ",' '                                     " & vbLf   '���C���R�[�h
    sSql = sSql & "," & Xodcx.Weight & "                    " & vbLf   '����d��
    sSql = sSql & "," & Xodcx.Weight & "                    " & vbLf   '���o�d��
    sSql = sSql & ",0                                       " & vbLf   '���X�d��
    sSql = sSql & ",'" & Right(nowCd, 4) & "'               " & vbLf   '����H���R�[�h
    sSql = sSql & ",'" & Right(PROCD_KOUNYU_TAN_KESSYOU, 4) & "'" & vbLf '���o�H���R�[�h('B410')
    sSql = sSql & ",' '                                     " & vbLf   '���o�敪
    sSql = sSql & ",'" & gsFactryCd & "'                    " & vbLf   '���o�H��R�[�h
    sSql = sSql & ",' '                                     " & vbLf   '���o���C���R�[�h
    sSql = sSql & ",' '                                     " & vbLf   '�`���[�WNo
    sSql = sSql & ",'" & sYear & "'                         " & vbLf   '���ѓ��t(�N)
    sSql = sSql & ",'" & sMonth & "'                        " & vbLf   '���ѓ��t(��)
    sSql = sSql & ",'" & sDay & "'                          " & vbLf   '���ѓ��t(��)
    sSql = sSql & ",'" & sCyoku & "'                        " & vbLf   '���敪
    sSql = sSql & ",'" & sHour & "'                         " & vbLf   '���ю���(��)
    sSql = sSql & ",'" & sMin & "'                          " & vbLf   '���ю���(��)
    sSql = sSql & ",'" & Xodcx.STAFFID & "'                 " & vbLf   '�S����
    sSql = sSql & ",'" & Xodcx.STAFFNAME & "'               " & vbLf   '�S���Җ�
    sSql = sSql & ",' '                                     " & vbLf   '�Z�x�敪
    sSql = sSql & ",NULL                                    " & vbLf   '�Z�x�l
    sSql = sSql & ",'7'                                     " & vbLf   '�������M�t���O
    sSql = sSql & ",'0'                                     " & vbLf   '�����t���O
    sSql = sSql & ",'1'                                     " & vbLf   '���i�敪
    sSql = sSql & ",'" & gsFactryCd & "'                    " & vbLf   '�g�p�\��H��
    sSql = sSql & ",NULL                                    " & vbLf   '���Z�x
    sSql = sSql & ",NULL                                    " & vbLf '
    sSql = sSql & ")"
    
    If SqlExec2(sSql) = -1 Then
        GoTo proc_exit
    End If
    
    DBDRV_fcmkc001c_InsXODB3 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    DBDRV_fcmkc001c_InsXODB3 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

' @(f)
' �@�\      : SQL���l�ϊ��֐�
'
' �Ԃ�l    : <���͐��l> or NULL
'
' ������    : �ϊ��Ώې��l
'
' �@�\����  : �n���ꂽ���l��NULL�ł����"NULL"�������łȂ���΂��̂܂܏o�͂���
Private Function ConvNum(vinput) As String
    If IsNull(vinput) Or vinput = "NULL" Then
        vinput = ""
    End If
    
    If vinput = "" Then
        ConvNum = "NULL"
    Else
        ConvNum = vinput
    End If
End Function

' @(f)
' �@�\      : ���敪����
'
' �Ԃ�l    : ���敪
'
' ������    : nowdate  -  ���t�f�[�^
'
' �@�\����  : �n���ꂽ���t�f�[�^�̎������璼�敪�𔻒肷��
'
Private Function GetCYOKU(nowdate As String) As String
    
    Dim jitutime As String     '����p�̎������i�[����ϐ�

    '����p�Ɏ�����؂�o��
    jitutime = Format(nowdate, "hhnnss")

    '���敪��ݒ肷��
    '3�� 00:00����07:59
    If jitutime >= "000000" And jitutime < "080000" Then
        GetCYOKU = "3"
    '1�� 08:00����15:59
    ElseIf jitutime >= "080000" And jitutime < "160000" Then
        GetCYOKU = "1"
    '2�� 16:00����23:59
    ElseIf jitutime >= "160000" And jitutime < "240000" Then
        GetCYOKU = "2"
    End If
End Function
'---- ADD [���������Ǘ��A�����H�����э쐬�����ǉ�] �ȏ�ǉ��֐��@END ---- TCS)T.TERAUCHI
Public Function DBDRV_SELECT_HOLD(gTblDispData As typ_TBCMJ012, pCrynum As String) As FUNCTION_RETURN

    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDC1_SQL.bas -- Function DBDRV_SELECT_HOLD"
     ''SQL��g�ݗ��Ă�
     'sql = "SELECT PROCCODE, HLDTRCLS, HLDCAUSE, HLDCMNT, UPDDATE, KSTAFFID, HOLDKT FROM TBCMJ012,XSDC2 "
     sql = "SELECT HLDCMNT FROM TBCMJ012,XSDC2 "
     sql = sql & " WHERE CRYNUMC2 = '" & pCrynum & "'"
     sql = sql & " AND   XTALC2 = CRYNUM   "
     sql = sql & " AND   INPOSC2 = INGOTPOS   "
     'sql = sql & " AND TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ012 WHERE CRYNUM = '" & pCrynum & "')"
     sql = sql & " ORDER BY TRANCNT"
    ''�f�[�^�𒊏o����
     Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
     
     If rs Is Nothing Then
         DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
         Exit Function
     End If
     With gTblDispData
        If rs.RecordCount > 0 Then
           rs.MoveLast
           If IsNull(rs("HLDCMNT")) = False Then .HLDCMNT = rs("HLDCMNT")
        End If
    End With
    rs.Close

    DBDRV_SELECT_HOLD = FUNCTION_RETURN_SUCCESS
    

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v   �F���ُo�͏���
'����   �F�u�o�͍ρv�ɂ����ꍇ�X�g�b�J�H���̍Ō�̍H���Ɂu�o�ׁv��ǉ�����
'����   �FstrCrynum     I   �����ԍ�
'       �FstrStaffID    I   �S���Җ�
'�Ԃ�l �FTRUE  ����
'       �FFALSE �ُ�
'����   �F2006/10/27 SETsw ����@�V�K�쐬
'���l   �F�g�����U�N�V�����͌Ăяo�����ł����Ă�������
Public Function StockerShip(StrCryNum As String, StrStaffId As String) As Boolean

    '�֐����Ȃ�ׂ��Ɨ������邽�߂ɒ�`�������Ɏ�������
    Const SHIP As String = "08"         '�o��
    Const SGEN As String = "09"         '�����������o
    Const KAKUSITA As String = "10"     '�i���i�o��
    Const STOCKER As String = "11"      '�X�g�b�J�[
    Const OLDSTOCKER As String = "12"   '�����X�g�b�J�[
    Const DELETE As String = "13"       '�폜
    Const HAIKI As String = "14"        '�p��
    
    Dim sSql As String
    Dim rs As OraDynaset    'RecordSet
    
    Dim sProcNum As String  '�H���ԍ�
    Dim sProcKbn As String  '�H���敪
    
    Dim sLastProcNum As String  '�ŏI�H���ԍ�
    
    Dim sTranNum As String      '������
    Dim sNowProcNum As String   '���ݍH���ԍ�
    
    Dim bUpdateFlg As Boolean   'UPDATE���s������
    
    StockerShip = False
    
    '���s����
    '�ŏI�H�����u�X�g�b�J�[�v�̏ꍇ�ɂ́u�X�g�b�J�[�v���u�o�ׁv�ɕύX����
    '�ŏI�H�����u�X�g�b�J�[�v�ȊO�̏ꍇ�ɂ͍ŏI�H���Ɂu�o�ׁv��o�^����
    '�������A�u�����������o�v�u�i���i�o�Ɂv�u�����X�g�b�J�[�v�u�폜�v�u�p���v�̏ꍇ�̓G���[�Ƃ���
    'TBCMF005�ɑΏۂ̌����ԍ����Ȃ��ꍇ�͉�����������I���Ƃ���B
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "�������.bas -- Function StockerShip"
    
    '�O����
        '�����񐔁A���ݍH���ԍ����擾����
        sSql = ""
        sSql = sSql & "SELECT TRANNUM, PROCNUM AS NOWPROCNUM FROM TBCMF005 WHERE CRYNUM = '" & StrCryNum & "' "
        sSql = sSql & "AND DELCLS = '0'"
        ''�f�[�^�𒊏o����
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
        If rs Is Nothing Then
            GoTo proc_exit
        End If
        ''���o���ʂ��i�[����
        sTranNum = NulltoStr(rs("TRANNUM"))
        sNowProcNum = NulltoStr(rs("NOWPROCNUM"))
        
        '1�����擾�ł��Ȃ��ꍇ�͐���I��
        If rs.RecordCount = 0 Then
            StockerShip = True
            Exit Function
        End If
        
        rs.Close
        
        '���ݍH���ԍ����擾�ł��Ȃ��ꍇ�A���ݍH����"0"�ُ͈�I��
        If sNowProcNum = "" Or sNowProcNum = "0" Then Exit Function

        '������
        Set rs = Nothing
    
        
        '�����ԍ��ɑ΂���폜���ꂽ�H�����܂߂��ŏI�H���ԍ����擾����
        sSql = ""
        sSql = sSql & "SELECT NVL(MAX(PROCNUM),0) AS LASTPROCNUM FROM TBCMF006 WHERE CRYNUM = '" & StrCryNum & "' "
        ''�f�[�^�𒊏o����
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
        If rs Is Nothing Then
            GoTo proc_exit
        End If
        ''���o���ʂ��i�[����
        sLastProcNum = NulltoStr(rs("LASTPROCNUM"))
        
        '�d�|���擾�ł��Ă��čH����1�����擾�ł��Ȃ��ꍇ�͏I��
        If rs.RecordCount = 0 Then
            StockerShip = True
            Exit Function
        End If
        rs.Close
        
        '�ŏI�H���ԍ����擾�ł��Ȃ��ꍇ�͏I��
        If sLastProcNum = "0" Then
            StockerShip = True
            Exit Function
        End If

        '������
        Set rs = Nothing
    
    '���@
    '1. �����ԍ�����ŏI�H���ԍ�(���폜)�ƍH���敪������o��
    sSql = ""
    sSql = sSql & "SELECT PROCKBN, PROCNUM FROM TBCMF006 "
    sSql = sSql & "WHERE "
    sSql = sSql & "CRYNUM = '" & StrCryNum & "' "
    sSql = sSql & "AND PROCNUM = ("
        sSql = sSql & "SELECT MAX(PROCNUM) FROM TBCMF006 WHERE CRYNUM = '" & StrCryNum & "' "
        sSql = sSql & "AND DELCLS = '0' "
        sSql = sSql & ") "
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    sProcKbn = NulltoStr(rs("PROCKBN"))
    sProcNum = NulltoStr(rs("PROCNUM"))
        
    '1�����擾�ł��Ȃ��ꍇ�͏I��
    If rs.RecordCount = 0 Then
        StockerShip = True
        Exit Function
    End If
    
    rs.Close
    
    '������
    Set rs = Nothing
    
    '2. �H���敪�𔻒f����
    '   2.1 �ŏI�H��(���폜)���u�����������o�v�u�i���i�o�Ɂv�u�����X�g�b�J�[�v�u�폜�v�u�p���v�ꍇ�ɂ̓G���[
        If sProcKbn = SGEN Or sProcKbn = KAKUSITA Or sProcKbn = OLDSTOCKER Or sProcKbn = DELETE Or sProcKbn = HAIKI Then
            Exit Function
    '   2.2 �ŏI�H��(���폜)���u�X�g�b�J�[�v�̏ꍇ�ɂ͌����ԍ��Ɗ���o�����H���ԍ���p����UPDATE
        ElseIf sProcKbn = STOCKER Then
        '�����ł�SQL���쐬���邾���B���s��if���𔲂�����B
            sSql = ""
            sSql = sSql & "UPDATE TBCMF006 SET "
            sSql = sSql & "PROCKBN = '" & SHIP & "', "
            sSql = sSql & "KSTAFFID = '" & StrStaffId & "', "
            sSql = sSql & "UPDDATE = SYSDATE "
            sSql = sSql & "WHERE "
            sSql = sSql & "CRYNUM = '" & StrCryNum & "' "
            sSql = sSql & "AND PROCNUM = " & sProcNum
            
            'UPDATE���s�t���O�𗧂ĂĂ���
            bUpdateFlg = True
    '   2.3 2.1,2.2�ȊO�̏ꍇ�A�����ԍ��Ɗ���o�����ŏI�H���ԍ�+1��p����INSERT
        '�����ł�SQL���쐬���邾���B���s��if���𔲂�����B
        Else
            sSql = ""
            sSql = sSql & "INSERT INTO TBCMF006("
            sSql = sSql & "CRYNUM,"
            sSql = sSql & "PROCNUM,"
            sSql = sSql & "PROCKBN,"
            sSql = sSql & "PROCSTAT,"
            sSql = sSql & "HOLDFLG,"
            sSql = sSql & "PRIORITY,"
            sSql = sSql & "KSIYOUFLG,"
            sSql = sSql & "DELIVFLG,"
            sSql = sSql & "STOCKFLG,"
            sSql = sSql & "RESLTFLG,"
            sSql = sSql & "INSPCTFLG,"
            sSql = sSql & "DELCLS,"
            sSql = sSql & "TSTAFFID,"
            sSql = sSql & "REGDATE) "
            sSql = sSql & "VALUES("
            sSql = sSql & "'" & StrCryNum & "',"
            sSql = sSql & CInt(sLastProcNum) + 1 & ","
            sSql = sSql & "'" & SHIP & "',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'1',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'" & StrStaffId & "',"
            sSql = sSql & "SYSDATE)"
        End If
    '���s
    OraDB.ExecuteSQL (sSql)

    'UPDATE�̏ꍇ�AUPDATE�����H�������ݍH���ԍ��ł������ꍇ�͉��H�d�|�e�[�u�����X�V����
    If bUpdateFlg Then
        If sProcNum = sNowProcNum Then
            sSql = ""
            sSql = sSql & "UPDATE TBCMF005 SET "
            sSql = sSql & "PROCKBN = '" & SHIP & "', "
            sSql = sSql & "KSTAFFID = '" & StrStaffId & "', "
            sSql = sSql & "UPDDATE = SYSDATE "
            sSql = sSql & "WHERE "
            sSql = sSql & "CRYNUM = '" & StrCryNum & "' "
            sSql = sSql & "AND TRANNUM = " & sTranNum
            
            '���s
            OraDB.ExecuteSQL (sSql)
        End If
    End If
    
    StockerShip = True
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    
    Resume proc_exit

End Function

' 2007/08/30 SPK Tsutsumi Add Start
Public Function GetMukeCode() As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Long      '���R�[�h��
    Dim i  As Long
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmbc016_0.frm -- Function Getstaffauthority"
    
    GetMukeCode = FUNCTION_RETURN_FAILURE
    
    sql = "Select CODEA9,NAMEJA9 "
    sql = sql & "from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '20' "
    sql = sql & "and (CODEA9 = '14' "
    sql = sql & "or CODEA9 = '15' "
    sql = sql & "or CODEA9 = '16' "
'2008/05/26 SHINDOH UPD
'    sql = sql & "or CODEA9 = 'ZZ') "
'------------------------------------
    sql = sql & "or CODEA9 = 'ZZ' "
    sql = sql & "or CODEA9 = 'ZX') "
'------------------------------------
    sql = sql & "order by CODEA9 "      '����s��Ή� 2009/01/04 SETsw kubota

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If
    
    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim s_Mukesaki(recCnt)
    
    If recCnt = 0 Then
        Exit Function
    End If
    
    For i = 1 To recCnt
        With s_Mukesaki(i)
            If IsNull(rs.Fields("CODEA9")) = False Then .sMukeCode = rs.Fields("CODEA9")    ' ����R�[�h
            If IsNull(rs.Fields("NAMEJA9")) = False Then .sMukeName = rs.Fields("NAMEJA9")  ' ���於
        End With
        rs.MoveNext
    Next
    rs.Close

    GetMukeCode = FUNCTION_RETURN_SUCCESS
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
'2007/08/30 SPK Tsutsumi Add End

'2008/01/25 SETsw kubota Add Start
Public Function GetSiyoHaraiLen(ByRef sHSXCLMIN As String _
                              , ByRef sHSXCLMAX As String _
                              ) As Boolean
    
    Dim sql As String
    Dim rs As OraDynaset
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL -- Function GetSiyoHaraiLen"
    
    GetSiyoHaraiLen = False
    
    ''�����l�̍ő�(�ł��������d�l)���擾
    sql = "Select max(HSXCLMIN) HSXCLMIN"   '�����l�̍ő�
    sql = sql & "  from TBCME020,XSDCA"
    sql = sql & " where CRYNUMCA = '" & f_cmbc032_2.txtBlkID & "'"
    sql = sql & "   and HINBAN   = HINBCA"
    sql = sql & "   and MNOREVNO = REVNUMCA"
    sql = sql & "   and FACTORY  = FACTORYCA"
    sql = sql & "   and OPECOND  = OPECA"
    sql = sql & "   and HSXCLMIN <> 0"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.EOF = True Then
        GoTo proc_exit
    End If
    sHSXCLMIN = NulltoStr(rs.Fields("HSXCLMIN"))
    
    rs.Close
    
    ''����l�̍ŏ�(�ł��������d�l)���擾
    sql = "Select min(HSXCLMAX) HSXCLMAX"   '����l�̍ŏ�
    sql = sql & "  from TBCME020,XSDCA"
    sql = sql & " where CRYNUMCA = '" & f_cmbc032_2.txtBlkID & "'"
    sql = sql & "   and HINBAN   = HINBCA"
    sql = sql & "   and MNOREVNO = REVNUMCA"
    sql = sql & "   and FACTORY  = FACTORYCA"
    sql = sql & "   and OPECOND  = OPECA"
    sql = sql & "   and HSXCLMAX <> 0"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.EOF = True Then
        GoTo proc_exit
    End If
    sHSXCLMAX = NulltoStr(rs.Fields("HSXCLMAX"))
    
    rs.Close

    GetSiyoHaraiLen = True

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
'2008/01/25 SETsw kubota Add End

''***********************************************************************************************
''SHINDOH ADD
''
''***********************************************************************************************
'------------------------------------------------------------------------------------------------------------(���蒼��STR)
'�T�v    :WF�o�ב҂��ꗗ �����\���p�c�a�h���C�o
'���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
'        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'����    :
'����    :
'@'Public Function DBDRV_scmzc_fcmkc001b_Disp5(records() As type_DBDRV_scmzc_fcmkc001b_Disp5) As FUNCTION_RETURN
'@'
    '��WF�Z���^���o�҂���
    'CC720�̂���
    '�G���[�n���h���̐ݒ�
'@'    On Error GoTo proc_err
'@'    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp5"

'@'    DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_SUCCESS


    '�u���b�NID��X�V���t�A�i�ԓ��擾
'@'     If GetListData(records(), "CC720") = FUNCTION_RETURN_FAILURE Then
'@'        DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_FAILURE
'@'        GoTo proc_exit
'@'    End If


'@'proc_exit:
    '�I��
'@'    gErr.Pop
'@'    Exit Function

'@'proc_err:
    '�G���[�n���h��
'@'    gErr.HandleError
'@'    DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_FAILURE
'@'    Resume proc_exit
'@'End Function
Public Function DBDRV_scmzc_fcmkc001b_Disp5(records() As type_DBDRV_scmzc_fcmkc001b_Disp5, tmpBlkData() As typ_BlkData) As FUNCTION_RETURN

    '�u���b�NID��X�V���t�擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp"

    DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_SUCCESS

    '�u���b�NID��X�V���t�A�i�ԓ��擾
    If getBlockID2(records(), tmpBlkData(), "CC720", 1, 0, "") = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'------------------------------------------------------------------------------------------------------------(���蒼��STR)
'�T�v    :SXL�o�ב҂��ꗗ �����\���p�c�a�h���C�o
'���Ұ�  :�ϐ���       ,IO  ,�^                                 ,����
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,�����\���p
'        :��ؒl        ,O   ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'����    :
'����    :
Public Function DBDRV_scmzc_fcmkc001b_Disp6(records() As type_DBDRV_scmzc_fcmkc001b_Disp5, tmpBlkData() As typ_BlkData, scmbKosei As Integer, stxtblk As String) As FUNCTION_RETURN

    '��SXL�o�בO��
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp6"

    DBDRV_scmzc_fcmkc001b_Disp6 = FUNCTION_RETURN_SUCCESS

    '�u���b�NID��X�V���t�A�i�ԓ��擾
     If getBlockID2(records(), tmpBlkData(), "CC705", 2, scmbKosei, stxtblk) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001b_Disp6 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    DBDRV_scmzc_fcmkc001b_Disp6 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'�����֐� �u���b�NID�A�X�V���t�擾�i���o�҂��A�����w���҂��p�j
Private Function getBlockID2(records() As type_DBDRV_scmzc_fcmkc001b_Disp5, _
                            pBlkData() As typ_BlkData, _
                            NOWPROC As String, formNum As Integer, scmbKosei As Integer, stxtblk As String) As FUNCTION_RETURN

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim recCnt      As Long         '���R�[�h��
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim BlockIdBuf  As String
    Dim sBlkID      As String
    Dim blkOrder    As Integer
    Dim Jiltuseki   As Judg_Kakou
    Dim nowtime     As Date         '���ݓ��t
    Dim sBakPos     As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function getBlockID2"

    getBlockID2 = FUNCTION_RETURN_SUCCESS

'    sql = "select X.XTALCA    as CRYNUM, "
'    sql = sql & " X.CRYNUMCA  as BLOCKID, "
'    sql = sql & " B.INGOTPOS, "
'    sql = sql & " XC2.KDAYC2, "
'    sql = sql & " B.HOLDCLS, "
'    sql = sql & " X.HINBCA    as HINBAN, "      ' �i��
'    sql = sql & " X.REVNUMCA  as REVNUM, "      ' ���i�ԍ������ԍ�
'    sql = sql & " X.FACTORYCA as FACTORY, "     ' �H��
'    sql = sql & " X.OPECA     as OPECOND, "     ' ���Ə���
'    sql = sql & " S.HSXTYPE, "                  ' �i�r�w�^�C�v
'    sql = sql & " S.HSXCDIR, "                  ' �i�r�w�����ʕ���
'    sql = sql & " B.REALLEN, "
'    sql = sql & " XC2.GNLC2 as LEN, "
'    sql = sql & " B2.BLOCKID  as SBLOCKID, "
'    sql = sql & " nvl("
'    sql = sql & "    (select DMTOP1 from TBCMI002 I2"
'    sql = sql & "     where CRYNUM=B.CRYNUM"
'    sql = sql & "       and INGOTPOS=(select max(INGOTPOS) from TBCMI002 where CRYNUM=B.CRYNUM  and INGOTPOS<=B.INGOTPOS)"
'    sql = sql & "       and TRANCNT =(select max(TRANCNT)  from TBCMI002 where CRYNUM=I2.CRYNUM and INGOTPOS=I2.INGOTPOS)"
'    sql = sql & "    )"
'    sql = sql & "    , (select DIAMETER from TBCME037 where CRYNUM=B.CRYNUM)"
'    sql = sql & "  ) as DIAM, "
'    sql = sql & " (select max(UPDDATE) from TBCMW001 where CRYNUM=B2.CRYNUM and INGOTPOS=B2.INGOTPOS) as NUKISHI_AT "
'    sql = sql & ",XC1.PUPTNC1 as PUPTN "            '��������ݒǉ��Ή�
'    sql = sql & ",X.HOLDBCA "                       'ΰ��ދ敪(XSDCA)
'    sql = sql & ",E36.WFCUTT "                      'WF��ĒP��
'    sql = sql & ",E36.BLOCKHFLAG "                  '��ۯ��P�ʕۏ��׸�
'    sql = sql & ",X.HOLDCCA "                       'ΰ��ޗ��R
'    sql = sql & ",X.HOLDKTCA "                       'ΰ��ލH��
'    sql = sql & ",X.PLANTCATCA "                    '����
'    sql = sql & " From  XSDCA X, TBCME018 S, TBCMJ010 J, TBCME040 B, TBCME040 B2 "
'    sql = sql & "     , XSDC1 XC1 "                 '��������ݒǉ��Ή�
'    sql = sql & "     , TBCME036 E36 "
'    sql = sql & "     , XSDC2 XC2 "
'    sql = sql & " where X.CRYNUMCA = B.BLOCKID "
'    sql = sql & "   and X.HINBCA   = S.HINBAN "
'    sql = sql & "   and X.REVNUMCA = S.MNOREVNO "
'    sql = sql & "   and X.FACTORYCA= S.FACTORY "
'    sql = sql & "   and X.OPECA    = S.OPECOND "
'    sql = sql & "   and X.HINBCA   = E36.HINBAN "
'    sql = sql & "   and X.REVNUMCA = E36.MNOREVNO "
'    sql = sql & "   and X.FACTORYCA= E36.FACTORY "
'    sql = sql & "   and X.OPECA    = E36.OPECOND "
'    sql = sql & "   and X.GNWKNTCA = '" & NOWPROC & "' "
'    sql = sql & "   and B.LSTATCLS = 'T' "
'    sql = sql & "   and B.RSTATCLS = 'T' "
'    sql = sql & "   and X.LIVKCA   = '0' "
'    sql = sql & "   and B.DELCLS   = '0' "
'    sql = sql & "   and J.CRYNUM   = B.CRYNUM "
'    sql = sql & "   and J.INGOTPOS = B.INGOTPOS"
'    sql = sql & "   and J.TRANCNT  = (select max(TRANCNT) from TBCMJ010 where CRYNUM=J.CRYNUM and INGOTPOS=J.INGOTPOS)"
'    sql = sql & "   and B2.CRYNUM  = B.CRYNUM"
'    sql = sql & "   and B2.INGOTPOS = B.INGOTPOS"
'    sql = sql & "   and X.XTALCA   = XC1.XTALC1(+) "   '��������ݒǉ��Ή�
'    sql = sql & "   and XC2.CRYNUMC2 = X.CRYNUMCA "
    sql = "select  X.XTALCA  as CRYNUM,"
    sql = sql & " X.CRYNUMCA  as BLOCKID,"
    sql = sql & " XC2.INPOSC2,"
    sql = sql & " XC2.KDAYC2,"
    sql = sql & " XC2.HOLDBC2,"
    sql = sql & " X.HINBCA    as HINBAN,"
    sql = sql & " X.REVNUMCA  as REVNUM,"
    sql = sql & " X.FACTORYCA as FACTORY,"
    sql = sql & " X.OPECA     as OPECOND,"
    sql = sql & " S.HSXTYPE,"
    sql = sql & " S.HSXCDIR,"
    sql = sql & " XC2.REALLC2,"
    sql = sql & " XC2.GNLC2 as LEN,"
    sql = sql & " XC1.PUPTNC1 as PUPTN ,"
    sql = sql & " X.HOLDBCA ,"
    sql = sql & " X.HOLDCCA ,"
    sql = sql & " X.HOLDKTCA ,"
    sql = sql & " X.PLANTCATCA"
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/07/01
    '' ������~���ڒǉ� add SETkimizuka Start  09/03/26
    'sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUS "
    'sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
    'sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSE "
    'sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNO "
    '' ������~���ڒǉ� add SETkimizuka End    09/03/26
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUS "
    sql = sql & " , NVL(TO_CHAR(Y4.STOPY4),' ') as STOP "
    sql = sql & " , DECODE(TRIM(Y4.CAUSEY4),NULL,' ',TRIM(Y4.CAUSEY4) || ':' || A9.NAMEJA9) as CAUSE "
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNO "
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/07/01
    sql = sql & " From  XSDCA X, TBCME018 S,XSDC1 XC1  ,XSDC2 XC2"
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/07/01
    sql = sql & "    ,XODY3 Y3,XODY4 Y4,KODA9 A9 "
    '' ������~���ڒǉ� add SETkimizuka Start  09/03/26
    'sql = sql & "    ,(SELECT XTALNOY3 as XTALNO,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
    'sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
    'sql = sql & "      FROM XODY3  "
    'sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND STOPY4 <> '2'  AND WKKTY4 in " & IIf(formNum = 1, CreateWkktSQL(WATCH_PROCCD_WF), IIf(formNum = 2, CreateWkktSQL(WATCH_PROCCD_BAR), "(' ')")) & ")"
    'sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
    'sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
    'sql = sql & "                FROM XODY3,XODY4,XODY5 "
    'sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
    'sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
    'sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
    'sql = sql & "      WHERE  "
    'sql = sql & "       LIVKY3    = '0' "
    'sql = sql & "       GROUP BY XTALNOY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9) Y4 "
    ' ������~���ڒǉ� add SETkimizuka End  09/03/26
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/07/01
    sql = sql & " Where"
    sql = sql & " x.HINBCA = s.hinban"
    sql = sql & " and X.REVNUMCA = S.MNOREVNO"
    sql = sql & " and X.FACTORYCA= S.FACTORY"
    sql = sql & " and X.OPECA    = S.OPECOND"
    sql = sql & " and X.XTALCA   = XC1.XTALC1(+)"
    If formNum = 1 Then
        sql = sql & " and XC2.GNWKNTC2 = 'CC720'"
    '    sql = sql & " and XC2.LSTATBC2 = 'T'"
    ElseIf formNum = 2 Then
        sql = sql & " and XC2.GNWKNTC2 = 'CC705'"
        sql = sql & " and XC2.LSTATBC2 = 'B'"
    End If
    sql = sql & " and XC2.RSTATBC2 = 'T'"
    sql = sql & " and XC2.LIVKC2   = '0'"
    sql = sql & " and X.LIVKCA     = '0'"
    sql = sql & " and XC2.CRYNUMC2 = X.CRYNUMCA"
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/07/01
    'sql = sql & " AND X.CRYNUMCA    = Y4.XTALNO(+) "            'add 09/03/26 SETkimizuka
    sql = sql & " AND X.CRYNUMCA = Y3.XTALNOY3(+) "
    sql = sql & " AND Y3.LIVKY3(+) = '0' "
    sql = sql & " AND Y4.LIVKY4(+) = '0' "
    sql = sql & " AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & " AND Y3.RCNTY3 = Y4.RCNTY4(+) "
    sql = sql & " AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/07/01

    If formNum = 2 Then
        sql = sql & "   and XC2.CRYNUMC2 like '" & stxtblk & "' "
        If scmbKosei = 0 Then
            sql = sql & "   and trim(XC2.GNWKKBC2) is null"
        Else
            sql = sql & "   and XC2.GNWKKBC2 ='" & scmbKosei & "'"
        End If
    End If

    sql = sql & " order by X.CRYNUMCA , X.INPOSCA"

Debug.Print "GetBlk " & sql
    '�f�[�^�𒊏o����
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    '���R�[�h���Ȃ��ꍇ����I��
    If rs.RecordCount = 0 Then
        rs.Close
        ReDim records(0)
        ReDim pBlkData(0)
        GoTo proc_exit
    End If

    BlockIdBuf = vbNullString
    recCnt = rs.RecordCount

    ReDim pBlkData(1 To recCnt)
    sBlkID = vbNullString
    blkOrder = 0
    j = 0
    For i = 1 To recCnt
        DoEvents
        '�u���b�NID���̊i�[
        If rs("BLOCKID") <> BlockIdBuf Then

            j = j + 1
            ReDim Preserve records(j)
            With records(j)
                .CRYNUM = rs("CRYNUM")
                .INGOTPOS = rs("INPOSC2")
                .BLOCKID = rs("BLOCKID")   ' �u���b�NID
                .UPDDATE = rs("KDAYC2")   ' �X�V���t
                .HOLDCLS = rs("HOLDBC2")   ' �z�[���h�敪
                BlockIdBuf = records(j).BLOCKID
                .HSXTYPE = rs("HSXTYPE")
                .HSXCDIR = rs("HSXCDIR")
                .Judg = " "
                .PUPTN = rs("PUPTN")
                If IsNull(rs("HOLDBCA")) = False Then .HOLDBCA = rs("HOLDBCA") Else .HOLDBCA = " "  'ΰ��ދ敪(XSDCA)
                If IsNull(rs("HOLDCCA")) = False Then .HOLDC = rs("HOLDCCA") Else .HOLDC = " "  'ΰ��ޗ��R
                If IsNull(rs("HOLDKTCA")) = False Then .HOLDKT = rs("HOLDKTCA") Else .HOLDKT = " "  'ΰ��ލH��
            End With

            k = 1
            sBakPos = ""    'add 09/03/26 SETkimizuka
        End If


        With pBlkData(i)
            .CRYNUM = rs("CRYNUM")
            .BLOCKID = rs("BLOCKID")
            .INGOTPOS = rs("INPOSC2")
            .LENGTH = rs("LEN")
            .REALLEN = rs("REALLC2")
            '.sBlockID = rs("sBLOCKID")
            .sBlockId = rs("BLOCKID")
            If sBlkID <> .sBlockId Then
                sBlkID = .sBlockId
                blkOrder = 1
            Else
                blkOrder = blkOrder + 1
            End If
            .BLOCKORDER = blkOrder
            '.DIAMETER = rs("DIAM")

            '�ŏI�������t�Ɍ��ݓ��t��\��
            nowtime = getSvrTime()
            .WFINDDATE = Format$(nowtime, "YYYY/MM/DD")
            .HOLDCLS = rs("HOLDBC2")
        End With


        If InStr(sBakPos, Trim(rs("INPOSC2"))) = 0 Then '�����Ď����ڒǉ��ɔ����C�� upd
            '�i�Ԃ̊i�[
            ReDim Preserve records(j).HIN(k)
            records(j).HIN(k).hinban = rs("HINBAN")
            records(j).HIN(k).mnorevno = rs("REVNUM")
            records(j).HIN(k).factory = rs("FACTORY")
            records(j).HIN(k).opecond = rs("OPECOND")
            
            If IsNull(rs("PLANTCATCA")) = False Then
                records(j).HIN(k).sMukesaki = rs("PLANTCATCA")
            End If
            sBakPos = sBakPos & Trim(rs("INPOSC2")) & " "
            k = k + 1
        End If

        ''�i�Ԃ̊i�[
        'ReDim Preserve records(j).HIN(k)
        'records(j).HIN(k).hinban = rs("HINBAN")
        'records(j).HIN(k).mnorevno = rs("REVNUM")
        'records(j).HIN(k).Factory = rs("FACTORY")
        'records(j).HIN(k).OpeCond = rs("OPECOND")
        
        'If IsNull(rs("PLANTCATCA")) = False Then
        '    records(j).HIN(k).sMukesaki = rs("PLANTCATCA")
        'End If

'        If k = 1 Then
'            '�i�ԂPWF��ĒP��
'            If IsNull(rs("WFCUTT")) = False Then
'                records(j).WFCUTT = rs("WFCUTT")
'            Else
'                records(j).WFCUTT = -1
'            End If
'            '�i�ԂP��ۯ��P�ʕۏ��׸�
'            If IsNull(rs("BLOCKHFLAG")) = False Then
'                records(j).BLOCKHFLAG = rs("BLOCKHFLAG")
'            Else
'                records(j).BLOCKHFLAG = " "
'            End If
'        End If
'

        ' �����Ď�SQL�C�� upd SETkimizuka Start  09/07/01
        '' ������~���ڒǉ� add SETkimizuka Start  09/03/26
        'records(j).STOP = rs("STOP")                   '��~�敪
        'records(j).AGRSTATUS = rs("AGRSTATUS")       '���F�m�F�敪
        'If Trim(rs("CAUSE")) <> "" And InStr(records(j).CAUSE, Trim(rs("CAUSE"))) = 0 Then
        '    records(j).CAUSE = records(j).CAUSE & rs("CAUSE") & vbTab       '��~���R
        'End If
        
        'IIf(formNum = 1, CreateWkktSQL(CC720), IIf(formNum = 2, CreateWkktSQL(CC705)
        
        If rs("STOP") <> "2" And rs("WKKTY4") = IIf(formNum = 1, "CC720", IIf(formNum = 2, "CC705", "")) Then
           If Trim(records(j).AGRSTATUS) = "" Or (rs("AGRSTATUS") < records(j).AGRSTATUS) Then
                records(j).STOP = rs("STOP")                   '��~�敪
                records(j).AGRSTATUS = rs("AGRSTATUS")         '���F�m�F�敪
           End If
            If Trim(rs("CAUSE")) <> "" And InStr(records(j).CAUSE, Trim(rs("CAUSE"))) = 0 Then
                records(j).CAUSE = records(j).CAUSE & rs("CAUSE") & vbTab       '��~���R
            End If
        End If
        ' �����Ď�SQL�C�� upd SETkimizuka End  09/07/01
        If Trim(rs("PRINTNO")) <> "" And InStr(records(j).PRINTNO, Trim(rs("PRINTNO"))) = 0 Then
            records(j).PRINTNO = records(j).PRINTNO & rs("PRINTNO") & vbTab       '��s�]��
        End If
        ' ������~���ڒǉ� add SETkimizuka End    09/03/26

        'k = k + 1
        rs.MoveNext
    Next i
    rs.Close

    For i = 1 To recCnt
        With pBlkData(i)
            If scmzc_getKakouJiltuseki(.BLOCKID, Jiltuseki) = FUNCTION_RETURN_SUCCESS Then
                .DIAMETER = (Jiltuseki.TAIL(1) + Jiltuseki.TAIL(2) + Jiltuseki.top(1) + Jiltuseki.top(2)) / 4
            End If
        End With
    Next

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getBlockID2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

Public Function cmkc001b_DBDataCheck5(LWD() As cmkc001b_LockWait, _
                                      Wd3() As type_DBDRV_scmzc_fcmkc001b_Disp5) As FUNCTION_RETURN
    Dim c0          As Integer
    Dim c1          As Integer
    Dim c2          As Integer
    Dim MaxRec      As Integer
    Dim RecCount    As Integer
    Dim EQFlag      As Boolean
    Dim sql         As String       ' SQL�S��
    Dim rs          As OraDynaset   ' RecordSet
    Dim GrpCount1   As Integer
    Dim GrpCount2   As Integer
    Dim ColorFlag   As Boolean
    Dim TotalBlk    As Integer
    Dim CheckPoint  As Integer
    Dim CheckEnd    As Integer
    Dim tempGrpFlag As String * 1

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function cmkc001b_DBDataCheck5"

    cmkc001b_DBDataCheck5 = FUNCTION_RETURN_SUCCESS
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
    'Z�i�ԃu���b�N�i���R������'$'���܂ނ��́j�͏���

    MaxRec = UBound(GrpInfo())
    For c0 = 1 To MaxRec
        sql = "select BLOCKID, INGOTPOS, LENGTH, NOWPROC, HOLDCLS "
        sql = sql & "from  TBCME040 "
        sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
        sql = sql & "  and 0     = INSTR(BLOCKID,'$',10,1)"
        sql = sql & "order by INGOTPOS, BLOCKID "

        '�f�[�^�𒊏o����
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        RecCount = rs.RecordCount
        If RecCount = 0 Then
            rs.Close
            GoTo proc_exit
        End If
        ReDim GrpInfo(c0).blkInfo(RecCount) As cmkc001b_Wait3_BLK
        For c1 = 1 To RecCount
            GrpInfo(c0).blkInfo(c1).BLOCKID = rs("BLOCKID")
            GrpInfo(c0).blkInfo(c1).INGOTPOS = rs("INGOTPOS")
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
Dim blkID()     As String
Dim topHin()    As tFullHinban
Dim botHin()    As tFullHinban
Dim idx         As Integer
Dim rsCount     As Integer
Dim found       As Boolean

    sql = vbNullString
    sql = sql & "select"
    sql = sql & "  b.BLOCKID"
    sql = sql & ", TOP.HINBAN as THINBAN, TOP.REVNUM as TREVNUM, TOP.FACTORY as TFACTORY, TOP.OPECOND as TOPECOND"
    sql = sql & ", BOT.HINBAN as BHINBAN, BOT.REVNUM as BREVNUM, BOT.FACTORY as BFACTORY, BOT.OPECOND as BOPECOND "

    sql = sql & "from TBCME040 B, TBCME041 TOP, TBCME041 BOT "

    sql = sql & "Where B.CRYNUM            =  TOP.CRYNUM"
    sql = sql & "  and B.CRYNUM            =  BOT.CRYNUM"

    sql = sql & "  and B.INGOTPOS          >=  0"
    sql = sql & "  and B.DELCLS            =  '0'"

    sql = sql & "  and B.NOWPROC           in ('CC600','CC700', 'CC710', 'CC720')"
    sql = sql & "  and B.RSTATCLS          =  'T'"
    sql = sql & "  and B.HOLDCLS           =  '0'"

    sql = sql & "  and B.INGOTPOS          >= TOP.INGOTPOS"
    sql = sql & "  and B.INGOTPOS          <  TOP.INGOTPOS+TOP.LENGTH"
    sql = sql & "  and B.INGOTPOS+B.LENGTH >  BOT.INGOTPOS"
    sql = sql & "  and B.INGOTPOS+B.LENGTH <= BOT.INGOTPOS+BOT.LENGTH "

    sql = sql & "order by B.BLOCKID"

    '�f�[�^�𒊏o����
    ' �㉺�i�Ԃ��t���i�ԂŎ擾
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    rsCount = rs.RecordCount

    ReDim blkID(1 To rsCount)
    ReDim topHin(1 To rsCount)
    ReDim botHin(1 To rsCount)
    For c0 = 1 To rsCount
        blkID(c0) = rs!BLOCKID

        topHin(c0).hinban = rs!THINBAN
        topHin(c0).mnorevno = rs!TREVNUM
        topHin(c0).factory = rs!TFACTORY
        topHin(c0).opecond = rs!TOPECOND

        botHin(c0).hinban = rs!BHINBAN
        botHin(c0).mnorevno = rs!BREVNUM
        botHin(c0).factory = rs!BFACTORY
        botHin(c0).opecond = rs!BOPECOND
        rs.MoveNext
    Next
    rs.Close

    For c0 = 1 To MaxRec
        RecCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To RecCount
            With GrpInfo(c0).blkInfo(c1)
                found = False
                For idx = 1 To rsCount
                    If blkID(idx) = .BLOCKID Then
                        found = True
                        Exit For
                    ElseIf blkID(idx) > .BLOCKID Then
                        Exit For
                    End If
                Next

                If found Then
                    .topHin.hinban = topHin(idx).hinban
                    .topHin.factory = topHin(idx).factory
                    .topHin.opecond = topHin(idx).opecond
                    .topHin.REVNUM = topHin(idx).mnorevno
                Else
                    .topHin.hinban = ""
                    .topHin.factory = ""
                    .topHin.opecond = ""
                    .topHin.REVNUM = 0
                End If

                If found Then
                    .botHin.hinban = botHin(idx).hinban
                    .botHin.factory = botHin(idx).factory
                    .botHin.opecond = botHin(idx).opecond
                    .botHin.REVNUM = botHin(idx).mnorevno
                Else
                    .botHin.hinban = ""
                    .botHin.factory = ""
                    .botHin.opecond = ""
                    .botHin.REVNUM = 0
                End If
            End With
        Next
    Next
#Else
' #IF �ŏ�������Ȃ��̂ŃR�����g�Ƃ��Ă����i�R�[�h����UP�̂��߁j
'    For c0 = 1 To MaxRec
'        RecCount = UBound(GrpInfo(c0).blkInfo())
'        For c1 = 1 To RecCount
'            sql = "select "
'            sql = sql & "HINBAN, "
'            sql = sql & "REVNUM, "
'            sql = sql & "FACTORY, "
'            sql = sql & "OPECOND "
'            sql = sql & "from TBCME041 "
'            sql = sql & "where CRYNUM='" & GrpInfo(c0).Crynum & "' "
''2001/11/14 S.Sano            sql = sql & "and INGOTPOS <= " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
'            sql = sql & "and INGOTPOS = " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " " '2001/11/14 S.Sano
''2001/11/14 S.Sano            sql = sql & "and (INGOTPOS + LENGTH) > " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
'
'            '�f�[�^�𒊏o����
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                GrpInfo(c0).blkInfo(c1).topHin.hinban = ""
'                GrpInfo(c0).blkInfo(c1).topHin.factory = ""
'                GrpInfo(c0).blkInfo(c1).topHin.opecond = ""
'                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = 0
'            Else
'                GrpInfo(c0).blkInfo(c1).topHin.hinban = rs("HINBAN")
'                GrpInfo(c0).blkInfo(c1).topHin.factory = rs("FACTORY")
'                GrpInfo(c0).blkInfo(c1).topHin.opecond = rs("OPECOND")
'                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = rs("REVNUM")
'            End If
'            rs.Close
'
'
'            sql = "select HINBAN, REVNUM, FACTORY, OPECOND "
'            sql = sql & "from  TBCME041 "
'            sql = sql & "where CRYNUM='" & GrpInfo(c0).Crynum & "' "
'            sql = sql & "  and INGOTPOS            <  " & GrpInfo(c0).blkInfo(c1).INGOTPOS + GrpInfo(c0).blkInfo(c1).LENGTH & " "
'            sql = sql & "  and (INGOTPOS + LENGTH) >= " & GrpInfo(c0).blkInfo(c1).INGOTPOS + GrpInfo(c0).blkInfo(c1).LENGTH & " "
'
'            '�f�[�^�𒊏o����
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                GrpInfo(c0).blkInfo(c1).botHin.hinban = ""
'                GrpInfo(c0).blkInfo(c1).botHin.factory = ""
'                GrpInfo(c0).blkInfo(c1).botHin.opecond = ""
'                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = 0
'            Else
'                GrpInfo(c0).blkInfo(c1).botHin.hinban = rs("HINBAN")
'                GrpInfo(c0).blkInfo(c1).botHin.factory = rs("FACTORY")
'                GrpInfo(c0).blkInfo(c1).botHin.opecond = rs("OPECOND")
'                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = rs("REVNUM")
'            End If
'            rs.Close
'        Next
'    Next
#End If

Debug.Print " 4:" & Time

    '���߂���񂩂�O���[�v�����߂�
    GrpCount1 = 0
    GrpCount2 = 0
    For c0 = 1 To MaxRec
        GrpCount1 = GrpCount1 + 1
        GrpCount2 = GrpCount2 + 1
        RecCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To RecCount
            Dim wTopHin As cmkc001b_Wait3_HINBAN
            Dim wBotHin As cmkc001b_Wait3_HINBAN

            wTopHin = GrpInfo(c0).blkInfo(c1).topHin
            wBotHin = GrpInfo(c0).blkInfo(c1 - 1).botHin

            '�u���b�N�؂�ڂŕi�Ԃ��ς��ΕʃO���[�v�Ɣ��f����
            Select Case c1
            Case 1
                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
            Case Else
                If (wTopHin.factory <> wBotHin.factory) Or (wTopHin.hinban <> wBotHin.hinban) Or _
                   (wTopHin.opecond <> wBotHin.opecond) Or (wTopHin.REVNUM <> wBotHin.REVNUM) Then
                    GrpCount1 = GrpCount1 + 1
                End If
                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
            End Select

            '����O���[�v���ŁA�H���Ⴂ�̃u���b�N�����݂����ꍇ�A����O���[�v����
            '���O���[�v�Ƃ��ăO���[�v��������B
            'CC710�ȊO�Ȃ�ΏۊO�Ƃ��O���[�v��������Ȃ�
            If GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_WFC_HARAIDASI And GrpInfo(c0).blkInfo(c1).HOLDCLS = "0" Then
                Select Case c1
                Case 1
                    GrpInfo(c0).blkInfo(c1).GRPFLG2 = GrpCount2
                Case Else
                    If (wTopHin.factory <> wBotHin.factory) Or (wTopHin.hinban <> wBotHin.hinban) Or _
                       (wTopHin.opecond <> wBotHin.opecond) Or (wTopHin.REVNUM <> wBotHin.REVNUM) Then
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
        RecCount = UBound(GrpInfo(c0).blkInfo())
        ColorFlag = False
        CheckPoint = 0
        For c1 = 1 To RecCount
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
        For c1 = CheckPoint To RecCount
            GrpInfo(c0).blkInfo(c1).COLORFLG = ColorFlag
        Next
    Next

Debug.Print " 6:" & Time

    For c0 = 1 To MaxRec
        RecCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To RecCount
            For c2 = 1 To TotalBlk
                If Wd3(c2).BLOCKID = GrpInfo(c0).blkInfo(c1).BLOCKID Then
                    LWD(c2).flag = GrpInfo(c0).blkInfo(c1).COLORFLG
                    LWD(c2).Grp = GrpInfo(c0).blkInfo(c1).GRPFLG2
                    Exit For
                End If
            Next
        Next
    Next

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    cmkc001b_DBDataCheck5 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



Public Function GetListData(records() As type_DBDRV_scmzc_fcmkc001b_Disp52, NOWPROC As String) As FUNCTION_RETURN

'�����֐� �u���b�NID�A�X�V���t�擾�i���o�҂��A�����w���҂��p�j

    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim recCnt      As Long         '���R�[�h��
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim BlockIdBuf  As String
    Dim lLp         As Long         '2007/08/30 SPK Tsutsumi Add
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function GetListData"

    GetListData = FUNCTION_RETURN_SUCCESS

    sql = "select "
    sql = sql & " XC2.XTALC2, "
    sql = sql & " XC2.INPOSC2, "
    sql = sql & " XC2.GNLC2, "
    sql = sql & " XC2.HOLDBC2, "
    sql = sql & " XC2.HOLDCC2, "
    sql = sql & " XC2.HOLDKTC2, "
    sql = sql & " XC2.LBLFLGC2,"
    sql = sql & " XCA.CRYNUMCA, "
    sql = sql & " XCA.INPOSCA, "
    sql = sql & " XCA.KDAYCA, "
    sql = sql & " XCA.HOLDBCA, "
    sql = sql & " XCA.HINBCA, "             ' �i��
    sql = sql & " XCA.REVNUMCA, "           ' ���i�ԍ������ԍ�
    sql = sql & " XCA.FACTORYCA, "          ' �H��
    sql = sql & " XCA.OPECA, "              ' ���Ə���
    sql = sql & " XCA.GNLCA, "              ' ����
    sql = sql & " XCA.GNWCA, "              ' �d��
    sql = sql & " S.HSXTYPE, "              ' �i�r�w�^�C�v
    sql = sql & " S.HSXCDIR "               ' �i�r�w�����ʕ���
    sql = sql & ",XC1.PUPTNC1 "             ' ���������
    sql = sql & ",XC2.KIKBNC2 "             ' �����ʋ敪
    sql = sql & ",XC2.PLANTCATC2 "          ' ����
    sql = sql & " from "
    sql = sql & " XSDCA XCA, TBCME018 S , XSDC2 XC2 "
    sql = sql & ",XSDC1 XC1 "                       '��������ݒǉ��Ή�
    sql = sql & " where "
    sql = sql & " XCA.CRYNUMCA = XC2.CRYNUMC2 "
    sql = sql & " and XCA.HINBCA = S.HINBAN "
    sql = sql & " and XCA.REVNUMCA = S.MNOREVNO "
    sql = sql & " and XCA.FACTORYCA = S.FACTORY "
    sql = sql & " and XCA.OPECA = S.OPECOND "
    sql = sql & " and XCA.GNWKNTCA ='" & NOWPROC & "' "
    sql = sql & " and XCA.LIVKCA='0' "
    sql = sql & " and XC2.XTALC2 = XC1.XTALC1(+) "    '��������ݒǉ��Ή�
    
    If NOWPROC = "CC705" Then
    
    End If
    sql = sql & " order by XCA.CRYNUMCA, XCA.INPOSCA "

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
        If rs("CRYNUMCA") <> BlockIdBuf Then
        
            j = j + 1
            ReDim Preserve records(j)
            With records(j)
                .CRYNUM = rs("XTALC2")
                .INGOTPOS = rs("INPOSCA")
                .BLOCKID = rs("CRYNUMCA")   ' �u���b�NID
                .UPDDATE = rs("KDAYCA")     ' �X�V���t
                .HOLDCLS = rs("HOLDBCA")    ' �z�[���h�敪
                BlockIdBuf = records(j).BLOCKID
                .HSXTYPE = rs("HSXTYPE")
                .HSXCDIR = rs("HSXCDIR")
                .INPOS = rs("INPOSC2")
                .LENGTH = rs("GNLC2")
                .PUPTN = rs("PUPTNC1")      ' ���������
                .HOLDB = rs("HOLDBC2")
                .HOLDC = rs("HOLDCC2")
                If IsNull(rs("HOLDKTC2")) = False Then .HOLDKT = rs("HOLDKTC2")
            End With
            
            k = 1
        End If
        
        '�i�Ԃ̊i�[
        ReDim Preserve records(j).hinM(1)
        records(j).hinM(1).HIN.hinban = rs("HINBCA")
        records(j).hinM(1).HIN.mnorevno = rs("REVNUMCA")
        records(j).hinM(1).HIN.factory = rs("FACTORYCA")
        records(j).hinM(1).HIN.opecond = rs("OPECA")
        records(j).hinM(1).LENGTH = rs("GNLCA")
        records(j).hinM(1).Weight = rs("GNWCA")
        
        ' ����
        If IsNull(rs("PLANTCATC2")) = False Then
            For lLp = 0 To UBound(s_Mukesaki)
                If rs("PLANTCATC2") = s_Mukesaki(lLp).sMukeCode Then
                    records(j).hinM(1).HIN.sMukesaki = s_Mukesaki(lLp).sMukeName
                End If
            Next lLp
        End If

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
    GetListData = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'-------------------------------------------------------------------------CC720����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pBlkHinMap() ,O  ,typ_BlkHinMap    ,�u���b�N�i�ԏ��
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :
'����      :2002/04/22 �쐬 �쑺
Public Function DBDRV_scmzc_fcmkc001h_Disp22(pBlkHinMap() As typ_BlkHinMap, formNum As Integer) As FUNCTION_RETURN
Dim sql     As String
Dim rs      As OraDynaset
Dim recCnt  As Long
Dim i       As Long
Dim j       As Long ' 2007/09/12 SPK Tsutsumi Add
Dim sBuff   As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_scmzc_fcmkc001h_Disp2"

    ''�u���b�N���̕i�ԍ\�����擾���� (�u���b�NID, �i��, ���i��, ������)
'    sql = "select X.CRYNUMCA  as BLOCKID, "
'    sql = sql & " B.PASSFLAG, "
'    sql = sql & " X.HINBCA    as HINBAN, "
'    sql = sql & " X.REVNUMCA  as REVNUM, "
'    sql = sql & " X.FACTORYCA as FACTORY, "
'    sql = sql & " X.OPECA     as OPECOND, "
'    sql = sql & " X.GNLCA     as HINLEN, "
'    sql = sql & " X.GNLCA + case when X.INPOSCA = (select max(INPOSCA) from XSDCA where CRYNUMCA = B.BLOCKID and LIVKCA = '0')"
'    sql = sql & "           then (J.P1BDLEN+J.P2BDLEN+J.P3BDLEN+J.P4BDLEN+J.P5BDLEN) "
'    sql = sql & "           else 0 end as REALLEN ,"
'    sql = sql & " X.INPOSCA as INPOSCA"
'    sql = sql & " ,X.PLANTCATCA as PLANTCATCA"
'    sql = sql & " from  XSDCA X, TBCME040 B, TBCMJ010 J "
'
'    sql = sql & " where X.CRYNUMCA = B.BLOCKID "
'    If formNum = 1 Then
'        sql = sql & "   and X.GNWKNTCA = 'CC720' "
'    ElseIf formNum = 2 Then
'        sql = sql & "   and X.NEWKNTCA = 'CC705' "
'    End If
'    sql = sql & "   and X.LIVKCA   = '0' "
'    sql = sql & "   and B.DELCLS   = '0'"
'    sql = sql & "   and J.CRYNUM   = B.CRYNUM "
'    sql = sql & "   and J.INGOTPOS = B.INGOTPOS"
'    sql = sql & "   and J.TRANCNT  = (select max(TRANCNT) from TBCMJ010 where CRYNUM=J.CRYNUM and INGOTPOS=J.INGOTPOS)"
'    If formNum = 1 Then
'        sql = sql & " and (X.PLANTCATCA =14  or X.PLANTCATCA =15 or X.PLANTCATCA =16)"
'    ElseIf formNum = 2 Then
'        sql = sql & " and (X.PLANTCATCA ='ZZ'  or X.PLANTCATCA ='ZX')"
'    End If
'    sql = sql & " order by B.BLOCKID, X.INPOSCA"

    sql = "select X.CRYNUMCA  as BLOCKID, "
    sql = sql & " X.HINBCA    as HINBAN, "
    sql = sql & " X.REVNUMCA  as REVNUM, "
    sql = sql & " X.FACTORYCA as FACTORY, "
    sql = sql & " X.OPECA     as OPECOND, "
    sql = sql & " X.GNLCA     as HINLEN, "
    sql = sql & " X.INPOSCA as INPOSCA, "
    sql = sql & " X.PLANTCATCA as PLANTCATCA"
    sql = sql & " from  XSDCA X, XSDC2 B"
    sql = sql & " where X.CRYNUMCA = B.CRYNUMC2 "
    If formNum = 1 Then
        sql = sql & "   and X.GNWKNTCA = 'CC720' "
    ElseIf formNum = 2 Then
        sql = sql & "   and X.NEWKNTCA = 'CC700' "
    End If
    sql = sql & "   and X.LIVKCA   = '0' "
    If formNum = 1 Then
'        sql = sql & " and (X.PLANTCATCA =14  or X.PLANTCATCA =15 or X.PLANTCATCA =16)"
        sql = sql & " and (X.PLANTCATCA ='14'  or X.PLANTCATCA ='15' or X.PLANTCATCA ='16')"
    ElseIf formNum = 2 Then
        sql = sql & " and (X.PLANTCATCA ='ZZ'  or X.PLANTCATCA ='ZX')"
    End If
    sql = sql & " order by X.CRYNUMCA, X.INPOSCA"

Debug.Print "Disp22 " & sql
    Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
    recCnt = rs.RecordCount
    If recCnt <= 0 Then
        ReDim pBlkHinMap(0)
    Else
        ReDim pBlkHinMap(1 To recCnt)
        For i = 1 To recCnt
            With pBlkHinMap(i)
                .BLOCKID = rs("BLOCKID")
                .HIN.hinban = rs("HINBAN")
                .HIN.mnorevno = rs("REVNUM")
                .HIN.factory = rs("FACTORY")
                .HIN.opecond = rs("OPECOND")
                .HinLen = rs("HINLEN")
                '.REALLEN = rs("REALLEN")
                .INPOSCA = rs("INPOSCA")

'                If IsNull(rs("PASSFLAG")) = True Then
'                    sBuff = ""
'                Else
'                    sBuff = rs("PASSFLAG")
'                End If
'                .PASSFLAG = vbNullString & sBuff
                
                If IsNull(rs("PLANTCATCA")) = True Then
                    .PLANTCATCA = ""
                Else
                    For j = 0 To UBound(s_Mukesaki)
                        If s_Mukesaki(j).sMukeCode = rs("PLANTCATCA") Then
                            .PLANTCATCA = ""
                            Exit For
                        End If
                    Next j
                End If
            End With
            rs.MoveNext
        Next
    End If
    rs.Close

    DBDRV_scmzc_fcmkc001h_Disp22 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmkc001h_Disp22 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'------------------------------------------------------------------------
''2008/05/30 SHINDOH 40�̏���
'------------------------------------------------------------------------

'�S�ʕύX 2003/10/17 SystemBrain
'����    :2008/03/31 �� �@���sSXL�m����s������SXL�}�b�v��M���ɑ��M���ݸޕύX����B
'                         �A�T���v��ID�𔽉f���̃T���v��ID(�����T���v��ID�܂�)�ł͂Ȃ��A��\�T���v��ID�ɕύX����B
'                         �BGB7/GB8/GB9��SXL�m����t����v������B
Public Function WriteX01n(ByVal DoProc%, ByVal blkID$, ByVal WfCnt%, errmsg$) As FUNCTION_RETURN
    Dim recX001(1 To 2)     As c_cmzcrec
    Dim recX002(1 To 2)     As c_cmzcrec
    Dim recX003(1 To 2)     As c_cmzcrec        'GD��������_�f�[�^
    Dim recX004(1 To 2)     As c_cmzcrec        'EP������
    Dim recX005(1 To 2)     As c_cmzcrec        'EP����_�ް�
    Dim i                   As Integer
    Dim j                   As Integer
    Dim rs                  As OraDynaset
    Dim sql                 As String
    Dim XlSmpPos(1 To 2)    As Integer
    Dim CRYNUM              As String
    Dim sBlkID(1 To 2)      As String       'XSDCW BLOCKID�i�[
    Dim smpId(2)            As String
    Dim HIN                 As tFullHinban
    Dim iX011cnt            As Integer      '08/09/12 ooba
    Dim recXSDCS(1 To 2)    As c_cmzcrec        '�V����يǗ�(��ۯ�)
    Dim recXSDCW(1 To 2)    As c_cmzcrec        '�V����يǗ�(SXL)
    Dim recE037             As c_cmzcrec        '�������
    Dim recXSDC1            As c_cmzcrec        '��������
    Dim recXSDC2            As c_cmzcrec
    Dim recX011(1 To 2)     As c_cmzcrec
    Dim recX012(1 To 2)     As c_cmzcrec
    Dim recX013(1 To 2)     As c_cmzcrec
    Dim Jiltuseki           As Judg_Kakou
    Dim sKMGCSHN            As String

    Dim RsHIN       As tFullHinban  '���R(Rs)�d�l�擾�i��
    Dim sRsData(10) As String       '���R(Rs)�ް�
'    Dim sRsPtn      As String       '���R�ް��擾�����
    Dim sRsPtn(2)   As String       '���R�ް��擾�����
    Dim sPos        As String       'SXL�ʒu(TOP/BOT)
    Dim gSmpID(2)   As String       'TBCMX003�p�T���v��ID
    Dim sErrMsg     As String       '�װү����
    Dim nowtime     As Date  '�BGB7/GB8/GB9��SXL�m����t����v������B
    
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function WriteX00n"
    
    WriteX01n = FUNCTION_RETURN_FAILURE
    
    '�BGB7/GB8/GB9��SXL�m����t����v������B
    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾

    ''SXL�̕i�Ԃ��擾����
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   HINBCA as HINBAN"           ''�i��
    sql = sql & "  ,REVNUMCA as REVNUM"         ''���i�ԍ������ԍ�
    sql = sql & "  ,FACTORYCA as FACTORY"       ''�H��
    sql = sql & "  ,OPECA as OPECOND"           ''���Ə���
    sql = sql & "  ,PLANTCATCA as PLANTCAT"     ''����  2007/09/04 SPK Tsutsumi Add
    sql = sql & " FROM"
    sql = sql & "   XSDCA"
    sql = sql & " WHERE CRYNUMCA = '" & blkID$ & "'"
    sql = sql & " and LIVKCA = '0'"             '08/07/22 ooba
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount < 1 Then
        errmsg = "XSDCA:" & rs.RecordCount
        rs.Close
        GoTo proc_exit
    End If
    HIN.hinban = rs!hinban
    HIN.mnorevno = rs!REVNUM
    HIN.factory = rs!factory
    HIN.opecond = rs!opecond
    HIN.sMukesaki = rs!PLANTCAT
    Set rs = Nothing

    
    '-------------------- XSDCS�̓ǂݍ��� ----------------------------------------
    For j = 1 To 2
        If j = 1 Then
            '�߂�XL����ʒu(FROM)�����߂�
            sql = "select * from XSDCS where CRYNUMCS = '" & blkID & "' and "
            sql = sql & "TBKBNCS = 'T' and LIVKCS = '0'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount <> 1 Then
                errmsg = "XSDCS:From"
                Set rs = Nothing
                GoTo proc_exit
            End If
            Set recXSDCS(1) = New c_cmzcrec
            recXSDCS(1).CopyFromRs "XSDCS", rs
            Set rs = Nothing
            XlSmpPos(1) = recXSDCS(1)("INPOSCS").Value
        ElseIf j = 2 Then
            '�߂�XL����ʒu(TO)�����߂�
            sql = "select * from XSDCS where CRYNUMCS = '" & blkID & "' and "
            sql = sql & "TBKBNCS = 'B' and LIVKCS = '0'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount <> 1 Then
                errmsg = "XSDCS:To"
                Set rs = Nothing
                GoTo proc_exit
            End If
            Set recXSDCS(2) = New c_cmzcrec
            recXSDCS(2).CopyFromRs "XSDCS", rs
            Set rs = Nothing
            XlSmpPos(2) = recXSDCS(2)("INPOSCS").Value
        End If
    Next j

    '-------------------- TBCME037�̓ǂݍ��� ----------------------------------------
    CRYNUM = left$(blkID, 9) & "000"        ' �����ԍ�
    sql = "select * from TBCME037 where (CRYNUM='" & CRYNUM & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "TBCME037"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recE037 = New c_cmzcrec
    recE037.CopyFromRs "TBCME037", rs
    Set rs = Nothing

    '-------------------- XSDC1�̓ǂݍ��� ----------------------------------------
    sql = "select * from XSDC1 where (XTALC1='" & CRYNUM & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "XSDC1"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recXSDC1 = New c_cmzcrec
    recXSDC1.CopyFromRs "XSDC1", rs
    Set rs = Nothing
    
    '-------------------- XSDC2�̓ǂݍ��� ----------------------------------------
    sql = "select * from XSDC2 where (CRYNUMC2='" & blkID & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "XSDC2"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recXSDC2 = New c_cmzcrec
    recXSDC2.CopyFromRs "XSDC2", rs
    Set rs = Nothing

    '-------------------- TBCME001�̓ǂݍ��� ----------------------------------------
    sql = "SELECT KMGCSHN FROM TBCME001 WHERE HINBAN = '" & HIN.hinban & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "TBCME001"
        Set rs = Nothing
        GoTo proc_exit
    Else
        sKMGCSHN = rs.Fields("KMGCSHN")
    End If
    Set rs = Nothing
                    
    '-------------------- TBCMI002�̓ǂݍ��� ----------------------------------------
    If scmzc_getKakouJiltuseki(blkID, Jiltuseki) = FUNCTION_RETURN_FAILURE Then
        errmsg = "TBCMI002"
        GoTo proc_exit
    End If

    'TBCMX011�����񐔎擾�@08/09/12 ooba
    If GetTBCMX011cnt(blkID, iX011cnt) = FUNCTION_RETURN_FAILURE Then
        errmsg = "TBCMX011"
        GoTo proc_exit
    End If
    
    '==============================================
    '�@�e����уf�[�^�̎擾�E�ݒ�
    '==============================================
    For i = 1 To 2
        '-------------------- TBCMX011�Œ���f�[�^�ݒ� ----------------------------------------
        Set recX011(i) = New c_cmzcrec
        recX011(i).TABLENAME = "TBCMX011"
        recX011(i).SetRecDefault
        
        With recX011(i)
            .Fields("BLOCKID").Value = blkID                                'BLOCKID
            .Fields("FROMTOKBN").Value = CStr(i)                            'FROMTO�敪
''            If DoProc = 0 Then
''                .Fields("TRANCNT").Value = 1                                '���ѓ��͂͂P�Œ�
''            Else
''                .Fields("TRANCNT").Value = "(SELECT NVL(MAX(TRANCNT),0) + 1 FROM TBCMX011" & _
''                                              " WHERE BLOCKID = '" & .Fields("BLOCKID").Value & "'" & _
''                                                " AND FROMTOKBN = '" & .Fields("FROMTOKBN").Value & "')"
''            End If
            .Fields("TRANCNT").Value = iX011cnt     '08/09/12 ooba
            .Fields("STCID").Value = IIf(Trim(f_cmbc032_2.lblSTCID.Caption) = "", "", f_cmbc032_2.lblSTCID.Caption)
            .Fields("HINBAN").Value = recXSDCS(i)("HINBCS").Value
            .Fields("REVNUM").Value = recXSDCS(i)("REVNUMCS").Value
            .Fields("FACTORY").Value = recXSDCS(i)("FACTORYCS").Value
            .Fields("OPECOND").Value = recXSDCS(i)("OPECS").Value
            If Trim(f_cmbc032_2.lblSTCID.Caption) = "" Then
                .Fields("STCKNNUM").Value = ""
            Else
                .Fields("STCKNNUM").Value = sKMGCSHN
            End If
            .Fields("CRYNUM").Value = recXSDCS(i)("XTALCS").Value
            .Fields("CRYDECDATE").Value = nowtime
            .nowtime = nowtime
            .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value
            .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value
            .Fields("FREELENG").Value = recE037("FREELENG").Value
            .Fields("INGOTPOS").Value = recXSDCS(i)("INPOSCS").Value
            .Fields("BLKLEN").Value = recXSDC2("REALLC2").Value
            .Fields("BLKWGHT").Value = recXSDC2("REALWC2").Value
            .Fields("LENGTH").Value = recXSDC2("GNLC2").Value
            .Fields("WEIGHT").Value = recXSDC2("GNWC2").Value
            .Fields("MCNO").Value = recE037("PRODCOND").Value
            .Fields("PGID").Value = recE037("PGID").Value
            If i = 1 Then
                .Fields("DM1").Value = Jiltuseki.top(1)
                .Fields("DM2").Value = Jiltuseki.top(2)
            Else
                .Fields("DM1").Value = Jiltuseki.TAIL(1)
                .Fields("DM2").Value = Jiltuseki.TAIL(2)
            End If
            .Fields("NCHDPTH").Value = Jiltuseki.DPTH(1)
'            .Fields("CHARGE").Value = recE037("CHARGE").Value / 1000
            .Fields("CHARGE").Value = recXSDC1("SUICHARGE").Value / 1000
            .Fields("SEED").Value = recE037("SEED").Value
            .Fields("REGDATE").Value = "SYSDATE"
            .Fields("SENDFLAG").Value = 0
        End With
        
        '-------------------- TBCMX012�Œ���f�[�^�ݒ� ----------------------------------------
        Set recX012(i) = New c_cmzcrec
        recX012(i).TABLENAME = "TBCMX012"
        recX012(i).SetRecDefault
        
        With recX012(i)
            .Fields("BLOCKID").Value = blkID                                'BLOCKID
            .Fields("FROMTOKBN").Value = CStr(i)                            'FROMTO�敪
            .Fields("TRANCNT").Value = iX011cnt                             '08/09/12 ooba
            .Fields("STCID").Value = IIf(Trim(f_cmbc032_2.lblSTCID.Caption) = "", "", f_cmbc032_2.lblSTCID.Caption)
            .Fields("HINBAN").Value = recXSDCS(i)("HINBCS").Value
            .Fields("REVNUM").Value = recXSDCS(i)("REVNUMCS").Value
            .Fields("FACTORY").Value = recXSDCS(i)("FACTORYCS").Value
            .Fields("OPECOND").Value = recXSDCS(i)("OPECS").Value
            If Trim(f_cmbc032_2.lblSTCID.Caption) = "" Then
                .Fields("STCKNNUM").Value = ""
            Else
                .Fields("STCKNNUM").Value = sKMGCSHN
            End If
            .Fields("CRYNUM").Value = recXSDCS(i)("XTALCS").Value
            .Fields("REGDATE").Value = "SYSDATE"
            .Fields("SENDFLAG").Value = 0
        End With
        
        '-------------------- TBCMX013�Œ���f�[�^�ݒ� ----------------------------------------
        Set recX013(i) = New c_cmzcrec
        recX013(i).TABLENAME = "TBCMX013"
        recX013(i).SetRecDefault
        
        With recX013(i)
            .Fields("BLOCKID").Value = blkID                                'BLOCKID
            .Fields("FROMTOKBN").Value = CStr(i)                            'FROMTO�敪
            .Fields("TRANCNT").Value = iX011cnt                             '08/09/12 ooba
            .Fields("STCID").Value = IIf(Trim(f_cmbc032_2.lblSTCID.Caption) = "", "", f_cmbc032_2.lblSTCID.Caption)
            .Fields("HINBAN").Value = recXSDCS(i)("HINBCS").Value
            .Fields("REVNUM").Value = recXSDCS(i)("REVNUMCS").Value
            .Fields("FACTORY").Value = recXSDCS(i)("FACTORYCS").Value
            .Fields("OPECOND").Value = recXSDCS(i)("OPECS").Value
            If Trim(f_cmbc032_2.lblSTCID.Caption) = "" Then
                .Fields("STCKNNUM").Value = ""
            Else
                .Fields("STCKNNUM").Value = sKMGCSHN
            End If
            .Fields("CRYNUM").Value = recXSDCS(i)("XTALCS").Value
            .Fields("REGDATE").Value = "SYSDATE"
            .Fields("SENDFLAG").Value = 0
        End With
                
        If i = 1 Then sPos = "TOP" Else sPos = "BOT"
        
        '-------------------- (����Rs)������R����(TBCMJ002)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMJ002(CRYNUM, recXSDCS(), i, HIN, recX011(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J002:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (����Oi)����Oi����(TBCMJ003)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMJ003(CRYNUM, recXSDCS(i), HIN, recX011(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J003:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (Cs)Cs����(TBCMJ004)�f�[�^�擾�ݒ� ----------------------------------------
        '�i�ԁ��װү���ޒǉ�
        If getTBCMJ004(CRYNUM, recXSDCS(i), HIN, recX011(i), sErrMsg) = FUNCTION_RETURN_FAILURE Then
            If sErrMsg = "" Then
                errmsg = "J004:" & XlSmpPos(i)
            Else
                errmsg = sErrMsg
            End If
            GoTo proc_exit
        End If

        '-------------------- (����OSF1�`4)����OSF����(TBCMJ005)�f�[�^�擾�ݒ� ----------------------------------------
        For j = 1 To 4
            If getTBCMJ005(CRYNUM, recXSDCS(i), j, recX011(i), recX012(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J005-" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (����BMD1�`3)����BMD����(TBCMJ008)�f�[�^�擾�ݒ� ----------------------------------------
        For j = 1 To 3
            If getTBCMJ008(CRYNUM, recXSDCS(i), j, recX011(i), recX012(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J008-" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (GD)GD����(TBCMJ006)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMJ006(CRYNUM, recXSDCS(i), recX011(i), recX013(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J006:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (LT)LT����(TBCMJ007)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMJ007(CRYNUM, recXSDCS(i), HIN, i, recX011(i), recX012(i)) = FUNCTION_RETURN_FAILURE Then  '05/12/05 ooba
            errmsg = "J007:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '==============================================
        '�@TBCMX011 �ɏ�������
        '==============================================
        With recX011(i)
            sql = .SqlInsert
            
            If 0 >= OraDB.ExecuteSQL(sql) Then
                WriteX01n = FUNCTION_RETURN_FAILURE
            End If
            
        End With

        '==============================================
        '�@TBCMX012 �ɏ�������
        '==============================================
        '�ύX�����o�^�@08/09/12 ooba
''        If DoProc = 0 Then
          With recX012(i)
            sql = .SqlInsert
            
            If 0 >= OraDB.ExecuteSQL(sql) Then
                WriteX01n = FUNCTION_RETURN_FAILURE
            End If
          End With
''        End If
            
            
        '==============================================
        '�@TBCMX013 �ɏ�������
        '==============================================
        '�ύX�����o�^�@08/09/12 ooba
''        If DoProc = 0 Then
          With recX013(i)
            sql = .SqlInsert
            
            If 0 >= OraDB.ExecuteSQL(sql) Then
                WriteX01n = FUNCTION_RETURN_FAILURE
            End If
          End With
''        End If
    
    Next
    
    WriteX01n = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    WriteX01n = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :���L�T���v���`�F�b�N����
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :inSXLID         , I  ,String            , SXL-ID
'          :inSMPLID        , I  ,String            , �����ID
'          :outSMPLID       , O  ,String            , ���L�����ID(���L�łȂ��ꍇ�AinSMPLID��Ԃ�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :�w�肳�ꂽ�����ID���S���L���ǂ������������A�S���L�̏ꍇ�A���L�����ID���擾���Ԃ�
'����      :��
Private Function chkComSAMPL(inSXLID As String, inSMPLID As String, outSMPLID As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim wXTALCW     As String
    Dim wINPOSCW    As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function chkComSAMPL"
    
    chkComSAMPL = FUNCTION_RETURN_SUCCESS
    outSMPLID = inSMPLID
    
    '-------------------- �S���L�m�F(XSDCW) ----------------------------------------
    sql = "select XTALCW, INPOSCW from XSDCW "
    sql = sql & "where SXLIDCW = '" & inSXLID & "' and "
    sql = sql & "      REPSMPLIDCW = '" & inSMPLID & "' and "
    sql = sql & "      (WFINDRSCW = '2' or WFINDRSCW = '0' or WFINDRSCW = ' ' or WFINDRSCW is null) and "
    sql = sql & "      (WFINDOICW = '2' or WFINDOICW = '0' or WFINDOICW = ' ' or WFINDOICW is null) and "
    sql = sql & "      (WFINDB1CW = '2' or WFINDB1CW = '0' or WFINDB1CW = ' ' or WFINDB1CW is null) and "
    sql = sql & "      (WFINDB2CW = '2' or WFINDB2CW = '0' or WFINDB2CW = ' ' or WFINDB2CW is null) and "
    sql = sql & "      (WFINDB2CW = '2' or WFINDB3CW = '0' or WFINDB3CW = ' ' or WFINDB3CW is null) and "
    sql = sql & "      (WFINDL1CW = '2' or WFINDL1CW = '0' or WFINDL1CW = ' ' or WFINDL1CW is null) and "
    sql = sql & "      (WFINDL2CW = '2' or WFINDL2CW = '0' or WFINDL2CW = ' ' or WFINDL2CW is null) and "
    sql = sql & "      (WFINDL3CW = '2' or WFINDL3CW = '0' or WFINDL3CW = ' ' or WFINDL3CW is null) and "
    sql = sql & "      (WFINDL4CW = '2' or WFINDL4CW = '0' or WFINDL4CW = ' ' or WFINDL4CW is null) and "
    sql = sql & "      (WFINDDSCW = '2' or WFINDDSCW = '0' or WFINDDSCW = ' ' or WFINDDSCW is null) and "
    sql = sql & "      (WFINDDZCW = '2' or WFINDDZCW = '0' or WFINDDZCW = ' ' or WFINDDZCW is null) and "
    sql = sql & "      (WFINDSPCW = '2' or WFINDSPCW = '0' or WFINDSPCW = ' ' or WFINDSPCW is null) and "
    sql = sql & "      (WFINDDO1CW = '2' or WFINDDO1CW = '0' or WFINDDO1CW = ' ' or WFINDDO1CW is null) and "
    sql = sql & "      (WFINDDO2CW = '2' or WFINDDO2CW = '0' or WFINDDO2CW = ' ' or WFINDDO2CW is null) and "
    sql = sql & "      (WFINDDO3CW = '2' or WFINDDO3CW = '0' or WFINDDO3CW = ' ' or WFINDDO3CW is null) and "
    sql = sql & "      (WFINDOT1CW = '2' or WFINDOT1CW = '0' or WFINDOT1CW = ' ' or WFINDOT1CW is null) and "
    sql = sql & "      (WFINDOT2CW = '2' or WFINDOT2CW = '0' or WFINDOT2CW = ' ' or WFINDOT2CW is null) and "
    sql = sql & "      (WFINDAOICW = '2' or WFINDAOICW = '0' or WFINDAOICW = ' ' or WFINDAOICW is null) and "
    sql = sql & "      (((WFINDGDCW = '2' or WFINDGDCW = '0' or WFINDGDCW = ' ' or WFINDGDCW is null) and WFHSGDCW = '0') or WFHSGDCW = '1') "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    wXTALCW = rs("XTALCW")      '�����ԍ�
    wINPOSCW = rs("INPOSCW")    '�������ʒu
    Set rs = Nothing
    
    '-------------------- ���L�����ID�̎擾(XSDCW) ----------------------------------------
    sql = "select REPSMPLIDCW from XSDCW "
    sql = sql & "where XTALCW = '" & wXTALCW & "' and "
    sql = sql & "      INPOSCW = '" & wINPOSCW & "' and "
    sql = sql & "      SXLIDCW != '" & inSXLID & "' and "
    sql = sql & "      REPSMPLIDCW != '" & inSMPLID & "' "
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    outSMPLID = rs("REPSMPLIDCW")       '��\�����ID(���L)
    Set rs = Nothing

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    chkComSAMPL = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :������R����(TBCMJ002)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS()      , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :i               , I  ,Integer           , Top/Bot���(1:Top, 2:Bot)
'          :hin             , I  ,tFullHinban       , �i��(�S�i�ԍ\����)
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :������R����(TBCMJ002)�����ް����擾���ASXL�������\���̂ɾ�Ă���
'����      :
Private Function getTBCMJ002(CRYNUM As String, recXSDCS() As c_cmzcrec, i As Integer, HIN As tFullHinban, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim k           As Integer
    Dim wMeas1(2)   As Double
    Dim wgtCharge   As Long                 '�ΐ͌v�Z�p�p�����[�^
    Dim wgtTop      As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim wgtTopCut   As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim DM          As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim cc          As type_Coefficient
    Dim CRes        As C_RES                '����RS����\����
    Dim wComp       As Double
    Dim wHSXRHWYS   As String               '�ۏؕ��@�Q��
    Dim RET As FUNCTION_RETURN
    Dim wStaff      As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ002"
    
    getTBCMJ002 = FUNCTION_RETURN_FAILURE

    With recX001
        .Fields("SXL_RS_SMPPOS").Value = -1                 'SXLRS�T���v������ʒu(SXL������)
        .Fields("SXLRS_MEAS1").Value = -1                   'SXLRS_����l1
        .Fields("SXLRS_MEAS2").Value = -1                   'SXLRS_����l2
        .Fields("SXLRS_MEAS3").Value = -1                   'SXLRS_����l3
        .Fields("SXLRS_MEAS4").Value = -1                   'SXLRS_����l4
        .Fields("SXLRS_MEAS5").Value = -1                   'SXLRS_����l5
        .Fields("SXLRS_EFEHS").Value = -1                   'SXLRS_�����ΐ�
        .Fields("SXLRS_RRG").Value = -1                     'SXLRS_RRG
    
        '-------------------- TBCMJ002�̓ǂݍ���(Rs) ----------------------------------------
        If (recXSDCS(i)("CRYINDRSCS").Value <> "0") And (recXSDCS(i)("CRYRESRS1CS").Value <> "0") Then
            '�����ΐ͎Z�o�ׁ̈ATop/Bot�̗������擾
            For k = 1 To 2
                sql = "select * from TBCMJ002 "
                sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
                sql = sql & "      SMPLNO = " & recXSDCS(k)("CRYSMPLIDRSCS").Value & " "
                sql = sql & "order by TRANCNT desc"
                sql = "select * from (" & sql & ") where rownum = 1"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                If k = i Then
                    .Fields("SXL_RS_SMPPOS").Value = rs("POSITION")             'SXLRS�T���v������ʒu(SXL������)
                    .Fields("SXLRS_MEAS1").Value = rs("MEAS1")                  'SXLRS_����l1
                    .Fields("SXLRS_MEAS2").Value = rs("MEAS2")                  'SXLRS_����l2
                    .Fields("SXLRS_MEAS3").Value = rs("MEAS3")                  'SXLRS_����l3
                    .Fields("SXLRS_MEAS4").Value = rs("MEAS4")                  'SXLRS_����l4
                    .Fields("SXLRS_MEAS5").Value = rs("MEAS5")                  'SXLRS_����l5
                    wStaff = rs("KSTAFFID")                                     '---TEST2004/10
                End If
                wMeas1(k) = rs("MEAS1")                             '�����ΐ͎Z�o�p
                Set rs = Nothing
            Next k
            
            'SXLRS_EFEHS
            If GetCoeffParams(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
            cc.DUNMENSEKI = AreaOfCircle(DM)
            cc.TOPSMPLPOS = recXSDCS(1)("INPOSCS").Value
            cc.BOTSMPLPOS = recXSDCS(2)("INPOSCS").Value
            cc.CHARGEWEIGHT = wgtCharge
            cc.TOPWEIGHT = wgtTop + wgtTopCut
            cc.TOPRES = wMeas1(1)
            cc.BOTRES = wMeas1(2)
            wComp = CoefficientCalculation(cc)
        
            If wComp = -9999 Then
                wComp = 0                                       'SXLRS_�����ΐ�
            End If
            .Fields("SXLRS_EFEHS").Value = wComp                'SXLRS_�����ΐ�
            
            'SXLRS_RRG
            'Cng Start 2011/10/13 Y.Hitomi
            'Cng Start 2011/09/19 Y.Hitomi
            sql = "select HSXRHWYS, HSXRSPOH, HSXRSPOT, HSXRSPOI,HSXRMCAL,HSXRHWYT from TBCME018 where "
'            sql = "select HSXRHWYS, HSXRSPOH, HSXRSPOT, HSXRSPOI,HSXRMCAL from TBCME018 where "
'            sql = "select HSXRHWYS, HSXRSPOH, HSXRSPOT, HSXRSPOI from TBCME018 where "
            'Cng End 2011/09/19 Y.Hitomi
            'Cng End 2011/10/13 Y.Hitomi
            
            sql = sql & " HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " FACTORY = '" & HIN.factory & "' and "
            sql = sql & " OPECOND = '" & HIN.opecond & "' "
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
                
            'Cng Start 2011/09/19 Y.Hitomi
            CRes.GuaranteeRes.cBunp = rs("HSXRMCAL")                     ' �i�r�w���R���z�v�Z
            'CRes.GuaranteeRes.cBunp = rs("HSXRSPOH")                    ' �i�r�w���R����ʒu�Q��
            'Cng End   2011/09/19 Y.Hitomi
            CRes.GuaranteeRes.cCount = rs("HSXRSPOT")                   ' �i�r�w���R����ʒu�Q�_
            CRes.GuaranteeRes.cPos = rs("HSXRSPOI")                     ' �i�r�w���R����ʒu�Q��
            wHSXRHWYS = rs("HSXRHWYS")                                  ' �i�r�w���R�ۏؕ��@�Q��
            'Add Start 2011/10/12 Y.Hitomi
            CRes.GuaranteeRes.cObj = rs("HSXRHWYT")                     ' �i�r�w���R�ۏؕ��@�Q��
            'Add End 2011/10/12 Y.Hitomi
            Set rs = Nothing
            
            CRes.Res(0) = NtoZ2(.Fields("SXLRS_MEAS1").Value)           'Rs����l1
            CRes.Res(1) = NtoZ2(.Fields("SXLRS_MEAS2").Value)           'Rs����l2
            CRes.Res(2) = NtoZ2(.Fields("SXLRS_MEAS3").Value)           'Rs����l3
            CRes.Res(3) = NtoZ2(.Fields("SXLRS_MEAS4").Value)           'Rs����l4
            CRes.Res(4) = NtoZ2(.Fields("SXLRS_MEAS5").Value)           'Rs����l5
            
            ''-----> 2006/06 ����ʒu�ɂ��v�Z�͕K�v�Ȃ��߃R�����g���O�����菇�Ƀf�[�^��߂�������ǉ�����
            If Trim(wStaff) <> KSTAFF_J002 Then   '�V����f�[�^�̏ꍇ������������
                RET = Set_Rs_Ichi(CRes.GuaranteeRes.cCount, CRes.GuaranteeRes.cPos, CRes.Res(0), CRes.Res(1), CRes.Res(2), _
                               CRes.Res(3), CRes.Res(4))
            End If
            
            .Fields("SXLRS_RRG").Value = CryRES_Judg(CRes.Res(), CRes.GuaranteeRes)     'SXLRS_RRG
            
            CRes.Res(0) = NtoZ2(.Fields("SXLRS_MEAS1").Value)           'Rs����l1
            CRes.Res(1) = NtoZ2(.Fields("SXLRS_MEAS2").Value)           'Rs����l2
            CRes.Res(2) = NtoZ2(.Fields("SXLRS_MEAS3").Value)           'Rs����l3
            CRes.Res(3) = NtoZ2(.Fields("SXLRS_MEAS4").Value)           'Rs����l4
            CRes.Res(4) = NtoZ2(.Fields("SXLRS_MEAS5").Value)           'Rs����l5

'Cng Start 2011/10/25 Y.Hitomi
            '�ۏؕ��@="H"�A���ASXLRS_RRG�v�Z���ʂ�-2(=���z�v�Z����`�j�̏ꍇ�A�G���[�Ƃ���B
            If (wHSXRHWYS = "H") And (.Fields("SXLRS_RRG").Value = -2) Then GoTo proc_exit
                        
'            '�ۏؕ��@="H"�A���ASXLRS_RRG�v�Z���ʂ�-1�̏ꍇ�A�G���[�Ƃ���
'            'Cng Start 2011/10/12 Y.Hitomi
'            'If (wHSXRHWYS = "H") And (.Fields("SXLRS_RRG").Value = -1) Then GoTo proc_exit
'            If (wHSXRHWYS = "H") And (.Fields("SXLRS_RRG").Value = -1) And CRes.GuaranteeRes.cObj <> "1" Then GoTo proc_exit
'            'Cng End 2011/10/12 Y.Hitomi
'Cng End 2011/10/25 Y.Hitomi
        End If
    End With

    getTBCMJ002 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ002 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :����Oi����(TBCMJ003)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :hin             , I  ,tFullHinban       , �i��(�S�i�ԍ\����)
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :����Oi����(TBCMJ003)�����ް����擾���ASXL�������\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ003(CRYNUM As String, recXSDCS As c_cmzcrec, HIN As tFullHinban, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim COi         As C_Oi                 '����Oi����\����
    Dim wHSXONHWS   As String               '�ۏؕ��@�Q��
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ003"
    
    getTBCMJ003 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    With recX001
        .Fields("SXL_OI_SMPPOS").Value = -1                 'SXLOI�T���v������ʒu(SXL������)
        .Fields("SXLOI_OIMEAS1").Value = -1                 'SXLOI_Oi����l1
        .Fields("SXLOI_OIMEAS2").Value = -1                 'SXLOI_Oi����l2
        .Fields("SXLOI_OIMEAS3").Value = -1                 'SXLOI_Oi����l3
        .Fields("SXLOI_OIMEAS4").Value = -1                 'SXLOI_Oi����l4
        .Fields("SXLOI_OIMEAS5").Value = -1                 'SXLOI_Oi����l5
        .Fields("SXLOI_ORGRES").Value = -1                  'SXLOI_ORG����
        .Fields("SXLOI_INSPECTWAY").Value = -1              'SXLOI�������@
    
        '-------------------- TBCMJ003�̓ǂݍ���(Oi) ----------------------------------------
        If (recXSDCS("CRYINDOICS").Value <> "0") And (recXSDCS("CRYRESOICS").Value <> "0") Then
            sql = "select * from TBCMJ003 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDOICS").Value & " "
            sql = sql & "  and TRANCOND = 0 "   'GFA��FTIR���Z�l�擾�ُ�Ή� 2011/02/28 SETsw kubota
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("SXL_OI_SMPPOS").Value = rs("POSITION")             'SXLOI�T���v������ʒu(SXL������)
''''            .Fields("SXLOI_OIMEAS1").Value = rs("OIMEAS1")              'SXLOI_Oi����l1
''''            .Fields("SXLOI_OIMEAS2").Value = rs("OIMEAS2")              'SXLOI_Oi����l2
''''            .Fields("SXLOI_OIMEAS3").Value = rs("OIMEAS3")              'SXLOI_Oi����l3
''''            .Fields("SXLOI_OIMEAS4").Value = rs("OIMEAS4")              'SXLOI_Oi����l4
''''            .Fields("SXLOI_OIMEAS5").Value = rs("OIMEAS5")              'SXLOI_Oi����l5
            'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
            If IsNull(rs("OIMEAS1")) = False Then .Fields("SXLOI_OIMEAS1").Value = rs("OIMEAS1") Else .Fields("SXLOI_OIMEAS1").Value = -1  'SXLOI_Oi����l1
            If IsNull(rs("OIMEAS2")) = False Then .Fields("SXLOI_OIMEAS2").Value = rs("OIMEAS2") Else .Fields("SXLOI_OIMEAS2").Value = -1  'SXLOI_Oi����l2
            If IsNull(rs("OIMEAS3")) = False Then .Fields("SXLOI_OIMEAS3").Value = rs("OIMEAS3") Else .Fields("SXLOI_OIMEAS3").Value = -1  'SXLOI_Oi����l3
            If IsNull(rs("OIMEAS4")) = False Then .Fields("SXLOI_OIMEAS4").Value = rs("OIMEAS4") Else .Fields("SXLOI_OIMEAS4").Value = -1  'SXLOI_Oi����l4
            If IsNull(rs("OIMEAS5")) = False Then .Fields("SXLOI_OIMEAS5").Value = rs("OIMEAS5") Else .Fields("SXLOI_OIMEAS5").Value = -1  'SXLOI_Oi����l5
            'OI_NULL�Ή��@2005/03/08 TUKU END   --------------------------------------------------
            .Fields("SXLOI_INSPECTWAY").Value = rs("INSPECTWAY")        'SXLOI�������@
            Set rs = Nothing
        
            'SXLOI_ORG
            'Cng Start 2011/10/13 Y.Hitomi
            'Cng Start 2011/09/19 Y.Hitomi
            sql = "select HSXONHWS, HSXONSPH, HSXONSPT, HSXONSPI,HSXONMCL,HSXONHWT from TBCME019 where "
'            sql = "select HSXONHWS, HSXONSPH, HSXONSPT, HSXONSPI,HSXONMCL from TBCME019 where "
            'sql = "select HSXONHWS, HSXONSPH, HSXONSPT, HSXONSPI from TBCME019 where "
            'Cng End   2011/09/19 Y.Hitomi
            'Cng Start 2011/10/13 Y.Hitomi
            sql = sql & " HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " FACTORY = '" & HIN.factory & "' and "
            sql = sql & " OPECOND = '" & HIN.opecond & "' "
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            
            ReDim COi.Oi(4) As Double
            'Cng Start 2011/09/19 Y.Hitomi
            COi.GuaranteeOi.cBunp = rs("HSXONMCL")                      ' �i�r�w�_�f�Z�x���z�v�Z
            'COi.GuaranteeOi.cBunp = rs("HSXONSPH")                      ' �i�r�w�_�f�Z�x����ʒu�Q��
            'Cng End   2011/09/19 Y.Hitomi
            
            COi.GuaranteeOi.cCount = rs("HSXONSPT")                     ' �i�r�w�_�f�Z�x����ʒu�Q�_
            COi.GuaranteeOi.cPos = rs("HSXONSPI")                       ' �i�r�w�_�f�Z�x����ʒu�Q��
            wHSXONHWS = rs("HSXONHWS")                                  ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
            'Add Start 2011/10/12 Y.Hitomi
            COi.GuaranteeOi.cObj = rs("HSXONHWT")                       ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
            'Add End 2011/10/12 Y.Hitomi
            Set rs = Nothing

            COi.Oi(0) = NtoZ2(.Fields("SXLOI_OIMEAS1").Value)           'Oi����l1
            COi.Oi(1) = NtoZ2(.Fields("SXLOI_OIMEAS2").Value)           'Oi����l2
            COi.Oi(2) = NtoZ2(.Fields("SXLOI_OIMEAS3").Value)           'Oi����l3
            COi.Oi(3) = NtoZ2(.Fields("SXLOI_OIMEAS4").Value)           'Oi����l4
            COi.Oi(4) = NtoZ2(.Fields("SXLOI_OIMEAS5").Value)           'Oi����l5
            
            .Fields("SXLOI_ORGRES").Value = CryOi_Judg(COi.Oi(), COi.GuaranteeOi)       'SXLOI_ORG����
            
'Cng Start 2011/10/25 Y.Hitomi
            '�ۏؕ��@="H"�A���ASXLOI_ORG�v�Z���ʂ�-2(=���z�v�Z����`�j�̏ꍇ�A�G���[�Ƃ���B
            If (wHSXONHWS = "H") And (.Fields("SXLOI_ORGRES").Value = -2) Then GoTo proc_exit
            
'            '�ۏؕ��@="H"�A���ASXLOI_ORG�v�Z���ʂ�-1�̏ꍇ�A�G���[�Ƃ���B2003/11/21 SystemBrain
''            If (wHSXONHWS = "H") And (.Fields("SXLOI_ORGRES").Value = -1) Then GoTo proc_exit
'            'Cng Start 2011/10/12 Y.Hitomi
'            If (wHSXONHWS = "H") And (.Fields("SXLOI_ORGRES").Value = -1) And COi.GuaranteeOi.cObj <> "1" Then GoTo proc_exit
'            'If (wHSXONHWS = "H") And (.Fields("SXLOI_ORGRES").Value = -1) Then GoTo proc_exit
'            'Cng End 2011/10/12 Y.Hitomi
'Cng End 2011/10/25 Y.Hitomi

        End If
    End With

    getTBCMJ003 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ003 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :Cs����(TBCMJ004)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :HIN             , I  ,tFullHinban       , �i�ԁ@06/04/20 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :sErrMsg         , O  ,String            , �װү���ށ@06/04/20 ooba
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :Cs����(TBCMJ004)�����ް����擾���ASXL�������\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ004(CRYNUM As String, recXSDCS As c_cmzcrec, HIN As tFullHinban, _
                             recX001 As c_cmzcrec, sErrMsg As String) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    

    Dim rs2         As OraDynaset
    Dim dCmax       As Double           '�d�l(����l)
    Dim dCmin       As Double           '�d�l(�����l)
    Dim iSmpNo      As Long             '���茳�����No
    Dim tCsSuitei   As CS_SUITEI_TYPE   'CS����v�Z�p�\����
    Dim dCsSuitei   As Double           'Cs����l
    
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ004"
    
    getTBCMJ004 = FUNCTION_RETURN_FAILURE

    sErrMsg = ""        '06/04/20 ooba
    
    '-------------------- �����ر ----------------------------------------
    With recX001
        .Fields("SXL_CS_SMPPOS").Value = -1                 'SXLCS�T���v������ʒu(SXL������)
        .Fields("SXLCS_CSMEAS").Value = -1                  'SXLCS_Cs�����l
        .Fields("SXLCS_70PPRE").Value = -1                  'SXLCS_70%����l
        .Fields("SXLCS_BSUIMEAS").Value = -1                'SXLCS_Cs��ۯ�����l
    
        '-------------------- TBCMJ004�̓ǂݍ���(Cs) ----------------------------------------
        If (recXSDCS("CRYINDCSCS").Value <> "0") And (recXSDCS("CRYRESCSCS").Value <> "0") Then
            sql = "select * from TBCMJ004 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDCSCS").Value & " "
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("SXL_CS_SMPPOS").Value = rs("POSITION")             'SXLCS�T���v������ʒu(SXL������)
''''            .Fields("SXLCS_CSMEAS").Value = rs("CSMEAS")                'SXLCS_Cs�����l
''''            .Fields("SXLCS_70PPRE").Value = rs("PRE70P")                'SXLCS_70%����l
            'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
            If IsNull(rs("CSMEAS")) = False Then .Fields("SXLCS_CSMEAS").Value = rs("CSMEAS") Else .Fields("SXLCS_CSMEAS").Value = -1  'SXLCS_Cs�����l
            If IsNull(rs("PRE70P")) = False Then .Fields("SXLCS_70PPRE").Value = rs("PRE70P") Else .Fields("SXLCS_70PPRE").Value = -1  'SXLCS_70%����l
            'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
            
            Set rs = Nothing
            
            ''Cs��ۯ�����l�v�Z�Ή��@06/04/20 ooba START ======================================>
        
            '�����̏ꍇ�͢��ۯ�����l�������l�
            If recXSDCS("CRYINDCSCS").Value = "1" Then
                .Fields("SXLCS_BSUIMEAS").Value = .Fields("SXLCS_CSMEAS").Value
            Else
                '�@����ʒu
                tCsSuitei.sInfPos = CStr(recXSDCS("INPOSCS").Value)
                
                '�A����وʒu
                '�B����ّ���l
                '���茳�����No�擾
                iSmpNo = recXSDCS("CRYSMPLIDCSCS").Value
                
                '����وʒu������ّ���l�擾
                sql = "select POSITION, CSMEAS from TBCMJ004 "
                sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
                sql = sql & "      SMPLNO = " & iSmpNo & " "
                sql = sql & "order by TRANCNT desc"
                
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <����ّ���l> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sSamplePos = rs("POSITION")       '����وʒu
                tCsSuitei.sResCs = rs("CSMEAS")             '����ّ���l
                Set rs = Nothing
                
                '�C����ޗ�
                '�DTOP�d��
                sql = "select SUICHARGE, WGHTTOC1, PUTCUTWC1 from XSDC1 "
                sql = sql & "where XTALC1 = '" & CRYNUM & "' "
                
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <����ޗ�> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                '�ް��s��
                If (IsNull(rs("SUICHARGE")) Or IsNull(rs("WGHTTOC1")) Or IsNull(rs("PUTCUTWC1"))) Then
                    sErrMsg = GetMsgStr("ECLC3") & " <����ޗ�> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                
                tCsSuitei.sSiWeight = rs("SUICHARGE")       '��������ޗ�
                tCsSuitei.sTopWT = CLng(rs("WGHTTOC1")) + CLng(rs("PUTCUTWC1"))     'TOP�d��
                Set rs = Nothing
                '���������ޗ�=0�or���������ޗʁ�TOP�d�ʣ�̏ꍇ�ʹװ�Ƃ���
                If CLng(tCsSuitei.sSiWeight) = 0 Or _
                   (CLng(tCsSuitei.sSiWeight) <= CLng(tCsSuitei.sTopWT)) Then
                    sErrMsg = GetMsgStr("ECLC3") & " <����ޗ�> "
                    GoTo proc_exit
                End If
                
                '�E���a
                sql = "select HSXD1CEN from TBCME018 "
                sql = sql & "where HINBAN = '" & HIN.hinban & "' "
                sql = sql & "and MNOREVNO = " & HIN.mnorevno & " "
                sql = sql & "and FACTORY = '" & HIN.factory & "' "
                sql = sql & "and OPECOND = '" & HIN.opecond & "' "
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <���a> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sUpDm = rs("HSXD1CEN")            '�iSX���a1���S
                
                '�F����ݕΐ͌W��
                sql = "select CTR01A9 from KODA9 "
                sql = sql & "where SYSCA9 = 'K' "
                sql = sql & "and SHUCA9 = 'AP' "
                sql = sql & "and CODEA9 = '1' "
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <����ݕΐ͌W��> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sCsHenseki = rs("CTR01A9")        '����ݕΐ͌W��
                
                '�GCs��ۯ�����l�v�Z
                If Not GetCsSuiteiMain(tCsSuitei, dCsSuitei) Then
                    sErrMsg = GetMsgStr("ECLC3")
                    GoTo proc_exit
                End If
                .Fields("SXLCS_BSUIMEAS").Value = dCsSuitei
            End If
            ''Cs��ۯ�����l�v�Z�Ή��@06/04/20 ooba END ========================================>
        End If
    End With

    getTBCMJ004 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ004 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :����OSF����(TBCMJ005)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :j               , I  ,Integer           , OSF No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :����OSF����(TBCMJ005)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ005(CRYNUM As String, recXSDCS As c_cmzcrec, j As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ005"
    
    getTBCMJ005 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        If j = 1 Then
            .Fields("SXLOSF_SMPPOS").Value = -1             'OSF�T���v������ʒu(SXL������)
        End If
        .Fields("SXLOSF" & j & "_KKSP").Value = ""          'OSFx�������ב���ʒu
        .Fields("SXLOSF" & j & "_NETU").Value = ""          'OSFx�M�����@
        .Fields("SXLOSF" & j & "_KKSET").Value = ""         'OSFx�������ב������+�I��ET��
        .Fields("SXLOSF" & j & "_CALCMAX").Value = -1       'OSFxSXL�v�Z���� Max_x
        .Fields("SXLOSF" & j & "_CALCAVE").Value = -1       'OSFxSXL�v�Z���� Ave_x
    End With
        
    'TBCMX002
    With recX002
        If j = 1 Then
            .Fields("SXLOSF1_SMPPOS").Value = -1            'SXLOSF�T���v������ʒu(SXL�ʒu���)
        End If
        .Fields("SXLOSF" & j & "_KKSP").Value = ""          'SXLOSFx�������׊m��ʒu
        .Fields("SXLOSF" & j & "_NETU").Value = ""          'SXLOSFx�M�����@
        .Fields("SXLOSF" & j & "_KKSET").Value = ""         'SXLOSFx�������ב������+�I��ET��
        .Fields("SXLOSF" & j & "_MEAS1").Value = -1         'SXLOSFx����_1
        .Fields("SXLOSF" & j & "_MEAS2").Value = -1         'SXLOSFx����_2
        .Fields("SXLOSF" & j & "_MEAS3").Value = -1         'SXLOSFx����_3
        .Fields("SXLOSF" & j & "_MEAS4").Value = -1         'SXLOSFx����_4
        .Fields("SXLOSF" & j & "_MEAS5").Value = -1         'SXLOSFx����_5
        .Fields("SXLOSF" & j & "_MEAS6").Value = -1         'SXLOSFx����_6
        .Fields("SXLOSF" & j & "_MEAS7").Value = -1         'SXLOSFx����_7
        .Fields("SXLOSF" & j & "_MEAS8").Value = -1         'SXLOSFx����_8
        .Fields("SXLOSF" & j & "_MEAS9").Value = -1         'SXLOSFx����_9
        .Fields("SXLOSF" & j & "_MEAS10").Value = -1        'SXLOSFx����_10
        .Fields("SXLOSF" & j & "_MEAS11").Value = -1        'SXLOSFx����_11
        .Fields("SXLOSF" & j & "_MEAS12").Value = -1        'SXLOSFx����_12
        .Fields("SXLOSF" & j & "_MEAS13").Value = -1        'SXLOSFx����_13
        .Fields("SXLOSF" & j & "_MEAS14").Value = -1        'SXLOSFx����_14
        .Fields("SXLOSF" & j & "_MEAS15").Value = -1        'SXLOSFx����_15
        .Fields("SXLOSF" & j & "_MEAS16").Value = -1        'SXLOSFx����_16
        .Fields("SXLOSF" & j & "_MEAS17").Value = -1        'SXLOSFx����_17
        .Fields("SXLOSF" & j & "_MEAS18").Value = -1        'SXLOSFx����_18
        .Fields("SXLOSF" & j & "_MEAS19").Value = -1        'SXLOSFx����_19
        .Fields("SXLOSF" & j & "_MEAS20").Value = -1        'SXLOSFx����_20
    End With
    
    '-------------------- TBCMJ005�̓ǂݍ���(OSF1�`4) ----------------------------------------
    If (recXSDCS("CRYINDL" & j & "CS").Value <> "0") And (recXSDCS("CRYRESL" & j & "CS").Value <> "0") Then
        sql = "select * from TBCMJ005 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDL" & j & "CS").Value & " and "
        sql = sql & "      TRANCOND = '" & j & "' "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
    
        'TBCMX001
        With recX001
            If .Fields("SXLOSF_SMPPOS").Value = -1 Then
                .Fields("SXLOSF_SMPPOS").Value = rs("POSITION")         'OSF�T���v������ʒu(SXL������)
            End If
            .Fields("SXLOSF" & j & "_KKSP").Value = rs("KKSP")          'OSFx�������ב���ʒu
            .Fields("SXLOSF" & j & "_NETU").Value = rs("HTPRC")         'OSFx�M�����@
            .Fields("SXLOSF" & j & "_KKSET").Value = rs("KKSET")        'OSFx�������ב������+�I��ET��
            .Fields("SXLOSF" & j & "_CALCMAX").Value = rs("CALCMAX")    'OSFxSXL�v�Z���� Max_x
            .Fields("SXLOSF" & j & "_CALCAVE").Value = rs("CALCAVE")    'OSFxSXL�v�Z���� Ave_x
        End With
            
        'TBCMX002
        With recX002
            If .Fields("SXLOSF1_SMPPOS").Value = -1 Then
                .Fields("SXLOSF1_SMPPOS").Value = rs("POSITION")        'SXLOSF�T���v������ʒu(SXL�ʒu���)
            End If
            .Fields("SXLOSF" & j & "_KKSP").Value = rs("KKSP")          'SXLOSFx�������׊m��ʒu
            .Fields("SXLOSF" & j & "_NETU").Value = rs("HTPRC")         'SXLOSFx�M�����@
            .Fields("SXLOSF" & j & "_KKSET").Value = rs("KKSET")        'SXLOSFx�������ב������+�I��ET��
            .Fields("SXLOSF" & j & "_MEAS1").Value = rs("MEAS1")        'SXLOSFx����_1
            .Fields("SXLOSF" & j & "_MEAS2").Value = rs("MEAS2")        'SXLOSFx����_2
            .Fields("SXLOSF" & j & "_MEAS3").Value = rs("MEAS3")        'SXLOSFx����_3
            .Fields("SXLOSF" & j & "_MEAS4").Value = rs("MEAS4")        'SXLOSFx����_4
            .Fields("SXLOSF" & j & "_MEAS5").Value = rs("MEAS5")        'SXLOSFx����_5
            .Fields("SXLOSF" & j & "_MEAS6").Value = rs("MEAS6")        'SXLOSFx����_6
            .Fields("SXLOSF" & j & "_MEAS7").Value = rs("MEAS7")        'SXLOSFx����_7
            .Fields("SXLOSF" & j & "_MEAS8").Value = rs("MEAS8")        'SXLOSFx����_8
            .Fields("SXLOSF" & j & "_MEAS9").Value = rs("MEAS9")        'SXLOSFx����_9
            .Fields("SXLOSF" & j & "_MEAS10").Value = rs("MEAS10")      'SXLOSFx����_10
            .Fields("SXLOSF" & j & "_MEAS11").Value = rs("MEAS11")      'SXLOSFx����_11
            .Fields("SXLOSF" & j & "_MEAS12").Value = rs("MEAS12")      'SXLOSFx����_12
            .Fields("SXLOSF" & j & "_MEAS13").Value = rs("MEAS13")      'SXLOSFx����_13
            .Fields("SXLOSF" & j & "_MEAS14").Value = rs("MEAS14")      'SXLOSFx����_14
            .Fields("SXLOSF" & j & "_MEAS15").Value = rs("MEAS15")      'SXLOSFx����_15
            .Fields("SXLOSF" & j & "_MEAS16").Value = rs("MEAS16")      'SXLOSFx����_16
            .Fields("SXLOSF" & j & "_MEAS17").Value = rs("MEAS17")      'SXLOSFx����_17
            .Fields("SXLOSF" & j & "_MEAS18").Value = rs("MEAS18")      'SXLOSFx����_18
            .Fields("SXLOSF" & j & "_MEAS19").Value = rs("MEAS19")      'SXLOSFx����_19
            .Fields("SXLOSF" & j & "_MEAS20").Value = rs("MEAS20")      'SXLOSFx����_20
        End With
        Set rs = Nothing
    End If

    getTBCMJ005 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ005 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :����BMD����(TBCMJ008)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :j               , I  ,Integer           , BMD No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :����BMD����(TBCMJ008)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :
Private Function getTBCMJ008(CRYNUM As String, recXSDCS As c_cmzcrec, j As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim dMeas(9)    As Double
    Dim strMeasPos  As String
    Dim iRet        As Integer
    Dim wComp       As Double
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ008"
    
    getTBCMJ008 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        If j = 1 Then
            .Fields("SXLBMD_SMPPOS").Value = -1             'BMD�T���v������ʒu(SXL�ʒu���)
        End If
        .Fields("SXLBMD" & j & "_KKSP").Value = ""          'BMDx�������ב���ʒu
        .Fields("SXLBMD" & j & "_NETU").Value = ""          'BMDx�M�����@
        .Fields("SXLBMD" & j & "_KKSET").Value = ""         'BMDx�������ב�������{�I��ET��
        .Fields("SXLBMD" & j & "_CALCMAX").Value = -1       'BMDxSXL�v�Z���� Max
        .Fields("SXLBMD" & j & "_CALCAVE").Value = -1       'BMDxSXL�v�Z���� Ave
        .Fields("SXLBMD" & j & "_CALCMIN").Value = -1       'BMDxSXL�v�Z���� Min
        .Fields("SXLBMD" & j & "_CALCMB").Value = -1        'BMDxSXL�v�Z���� �ʓ����z
    End With
        
    'TBCMX002
    With recX002
        If j = 1 Then
            .Fields("SXLBMD_SMPPOS").Value = -1             'SXLBMD�T���v������ʒu(SXL�ʒu���)
        End If
        .Fields("SXLBMD" & j & "_KKSP").Value = ""          'SXLBMD1�������ב���ʒu
        .Fields("SXLBMD" & j & "_NETU").Value = ""          'SXLBMD1�M�����@
        .Fields("SXLBMD" & j & "_KKSET").Value = ""         'SXLBMD1�������ב������+�I��ET��
        .Fields("SXLBMD" & j & "_MEAS1").Value = -1         'SXLBMD1����_1
        .Fields("SXLBMD" & j & "_MEAS2").Value = -1         'SXLBMD1����_2
        .Fields("SXLBMD" & j & "_MEAS3").Value = -1         'SXLBMD1����_3
        .Fields("SXLBMD" & j & "_MEAS4").Value = -1         'SXLBMD1����_4
        .Fields("SXLBMD" & j & "_MEAS5").Value = -1         'SXLBMD1����_5
    End With
    
    '-------------------- TBCMJ008�̓ǂݍ���(BMD1�`3) ----------------------------------------
    If (recXSDCS("CRYINDB" & j & "CS").Value <> "0") And (recXSDCS("CRYRESB" & j & "CS").Value <> "0") Then
        sql = "select * from TBCMJ008 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDB" & j & "CS").Value & " and "
        sql = sql & "      TRANCOND = '" & j & "' "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            If .Fields("SXLBMD_SMPPOS").Value = -1 Then
                .Fields("SXLBMD_SMPPOS").Value = rs("POSITION")         'BMD�T���v������ʒu(SXL�ʒu���)
            End If
            .Fields("SXLBMD" & j & "_KKSP").Value = rs("KKSP")          'BMDx�������ב���ʒu
            .Fields("SXLBMD" & j & "_NETU").Value = rs("HTPRC")         'BMDx�M�����@
            .Fields("SXLBMD" & j & "_KKSET").Value = rs("KKSET")        'BMDx�������ב�������{�I��ET��
            .Fields("SXLBMD" & j & "_CALCMAX").Value = rs("MEASMAX")    'BMDxSXL�v�Z���� Max
            .Fields("SXLBMD" & j & "_CALCAVE").Value = rs("MEASAVE")    'BMDxSXL�v�Z���� Ave
'            .Fields("SXLBMD" & j & "_CALCMB").Value = rs("BMDMNBUNP")   'BMDxSXL�v�Z���� �ʓ����z
            If IsNull(rs("BMDMNBUNP")) = False Then .Fields("SXLBMD" & j & "_CALCMB").Value = rs("BMDMNBUNP")   'BMDxSXL�v�Z���� �ʓ����z
        End With
            
        'TBCMX002
        With recX002
            If .Fields("SXLBMD_SMPPOS").Value = -1 Then
                .Fields("SXLBMD_SMPPOS").Value = rs("POSITION")         'SXLBMD�T���v������ʒu(SXL�ʒu���)
            End If
            .Fields("SXLBMD" & j & "_KKSP").Value = rs("KKSP")          'SXLBMDx�������ב���ʒu
            .Fields("SXLBMD" & j & "_NETU").Value = rs("HTPRC")         'SXLBMDx�M�����@
            .Fields("SXLBMD" & j & "_KKSET").Value = rs("KKSET")        'SXLBMDx�������ב������+�I��ET��
            .Fields("SXLBMD" & j & "_MEAS1").Value = rs("MEAS1")        'SXLBMDx����_1
            .Fields("SXLBMD" & j & "_MEAS2").Value = rs("MEAS2")        'SXLBMDx����_2
            .Fields("SXLBMD" & j & "_MEAS3").Value = rs("MEAS3")        'SXLBMDx����_3
            .Fields("SXLBMD" & j & "_MEAS4").Value = rs("MEAS4")        'SXLBMDx����_4
            .Fields("SXLBMD" & j & "_MEAS5").Value = rs("MEAS5")        'SXLBMDx����_5
        End With
        Set rs = Nothing
    
        'BMD�ŏ��l�̎擾 2003/05/31 tuku                START
        dMeas(0) = recX002.Fields("SXLBMD" & j & "_MEAS1").Value
        dMeas(1) = recX002.Fields("SXLBMD" & j & "_MEAS2").Value
        dMeas(2) = recX002.Fields("SXLBMD" & j & "_MEAS3").Value
        dMeas(3) = recX002.Fields("SXLBMD" & j & "_MEAS4").Value
        dMeas(4) = recX002.Fields("SXLBMD" & j & "_MEAS5").Value
        ''�������ב���ʒu�R�[�h
        strMeasPos = Trim(recX002.Fields("SXLBMD" & j & "_KKSP").Value)
        ''�ŏ��l���v�Z����B
        iRet = getSXLBMDMIN(wComp, strMeasPos, dMeas)
        ''�v�Z���ʂ��i�[����
        recX001.Fields("SXLBMD" & j & "_CALCMIN").Value = wComp
    End If

    getTBCMJ008 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ008 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :GD����(TBCMJ006)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :GD����(TBCMJ006)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :
Private Function getTBCMJ006(CRYNUM As String, recXSDCS As c_cmzcrec, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ006"
    
    getTBCMJ006 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("SXLGD_SMPPOS").Value = -1                  'GD�T���v������ʒu(SXL�ʒu���)
        .Fields("SXLGD_MSRSDEN").Value = -1                 'SXLGD_���茋�� Den
        .Fields("SXLGD_MSRSLDL").Value = -1                 'SXLGD_���茋�� L/DL
        .Fields("SXLGD_MSRSDVD2").Value = -1                'SXLGD_���茋�� DVD2
    End With
        
    'TBCMX002
    With recX002
        .Fields("SXLGD_SMPPOS").Value = -1                                  'SXLGD�T���v������ʒu(SXL�ʒu���)
        For i = 1 To 15
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL1").Value = -1       'SXLGD_����lxx L/DL1
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL2").Value = -1       'SXLGD_����lxx L/DL2
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL3").Value = -1       'SXLGD_����lxx L/DL3
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL4").Value = -1       'SXLGD_����lxx L/DL4
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL5").Value = -1       'SXLGD_����lxx L/DL5
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN1").Value = -1       'SXLGD_����lxx Den1
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN2").Value = -1       'SXLGD_����lxx Den2
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN3").Value = -1       'SXLGD_����lxx Den3
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN4").Value = -1       'SXLGD_����lxx Den4
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN5").Value = -1       'SXLGD_����lxx Den5
        Next
        For i = 1 To 5
            .Fields("SXLGD_MS01DVD2" & i).Value = -1                        'SXLGD_����lxx DVD2
        Next
    End With
        
    '-------------------- TBCMJ006�̓ǂݍ���(GD) ----------------------------------------
    If (recXSDCS("CRYINDGDCS").Value <> "0") And (recXSDCS("CRYRESGDCS").Value <> "0") Then
        sql = "select * from TBCMJ006 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDGDCS").Value & " "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("SXLGD_SMPPOS").Value = rs("POSITION")              'GD�T���v������ʒu(SXL�ʒu���)
            .Fields("SXLGD_MSRSDEN").Value = rs("MSRSDEN")              'SXLGD_���茋�� Den
            .Fields("SXLGD_MSRSLDL").Value = rs("MSRSLDL")              'SXLGD_���茋�� L/DL
            .Fields("SXLGD_MSRSDVD2").Value = rs("MSRSDVD2")            'SXLGD_���茋�� DVD2
        End With
            
        'TBCMX002
        With recX002
            .Fields("SXLGD_SMPPOS").Value = rs("POSITION")                                                      'SXLGD�T���v������ʒu(SXL�ʒu���)
            For i = 1 To 15
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'SXLGD_����lxx L/DL1
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'SXLGD_����lxx L/DL2
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'SXLGD_����lxx L/DL3
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'SXLGD_����lxx L/DL4
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'SXLGD_����lxx L/DL5
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'SXLGD_����lxx Den1
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'SXLGD_����lxx Den2
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'SXLGD_����lxx Den3
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'SXLGD_����lxx Den4
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'SXLGD_����lxx Den5
            Next
            
'--------------- 208/06/24 INSERT START  By Systech ---------------
            For i = 1 To 5
                If rs("MS0" & i & "DVD2") <> -1 Then
                    .Fields("SXLGD_MS01DVD2" & i).Value = rs("MS0" & i & "DVD2")                                'SXLGD_����lxx DVD2
                End If
            Next
'--------------- 208/06/24 INSERT  END   By Systech ---------------
        
        End With
        Set rs = Nothing
    End If

    getTBCMJ006 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ006 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :LT����(TBCMJ007)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :ChkHin          , I  ,tFullHinban       , LT�d�l�擾�p�i�ԁ@05/12/05 ooba
'          :i               , I  ,Integer           , BMD No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :LT����(TBCMJ007)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :
Private Function getTBCMJ007(CRYNUM As String, recXSDCS As c_cmzcrec, ChkHin As tFullHinban, i As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim j           As Integer      '
    Dim rs2         As OraDynaset   '
    Dim sql2        As String       '
    Dim iRet        As Integer      '
    Dim iTmpMes(9)  As Integer      'LT�����ް�(1�`10)
    Dim iCalcMeas   As Integer      'LT�v�Z����
    Dim sIchi       As String       '�iSXL��ё���ʒu_��
    Dim iOldFlg     As Integer      '���ް��׸�
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ007"
    
    getTBCMJ007 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("SXLLT_SMPPOS").Value = -1                  'LT�T���v������ʒu(SXL�ʒu���)
        .Fields("SXLLT_MEASPEAK").Value = -1                'SXLLT_����l �s�[�N�l
        .Fields("SXLLT_CALCMEAS").Value = -1                'SXLLT_�v�Z����
    End With
        
    'TBCMX002
    With recX002
        .Fields("SXLT_SMPPOS").Value = -1                   'SXLLT�T���v������ʒu(SXL�ʒu���)
        .Fields("SXLLT_MEASPEAK").Value = -1                'SXLLT_����l �s�[�N�l
        .Fields("SXLLT_MEAS1").Value = -1                   'SXLLT_����l1
        .Fields("SXLLT_MEAS2").Value = -1                   'SXLLT_����l2
        .Fields("SXLLT_MEAS3").Value = -1                   'SXLLT_����l3
        .Fields("SXLLT_MEAS4").Value = -1                   'SXLLT_����l4
        .Fields("SXLLT_MEAS5").Value = -1                   'SXLLT_����l5
    End With
        
    'BOT���̂��ް��擾
    If i <> 1 Then
        '-------------------- TBCMJ007�̓ǂݍ���(LT) ----------------------------------------
        If (recXSDCS("CRYINDTCS").Value <> "0") And (recXSDCS("CRYRESTCS").Value <> "0") Then
            sql = "select * from TBCMJ007 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDTCS").Value & " "
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            
            If IsNull(rs("LTSPIFLG")) Then iOldFlg = 1 Else iOldFlg = 0
            
            '������
            iCalcMeas = -1
            For j = 0 To 9
                iTmpMes(j) = -1
            Next j
            
            If Not IsNull(rs("MEAS1")) Then iTmpMes(0) = rs("MEAS1")
            If Not IsNull(rs("MEAS2")) Then iTmpMes(1) = rs("MEAS2")
            If Not IsNull(rs("MEAS3")) Then iTmpMes(2) = rs("MEAS3")
            If Not IsNull(rs("MEAS4")) Then iTmpMes(3) = rs("MEAS4")
            If Not IsNull(rs("MEAS5")) Then iTmpMes(4) = rs("MEAS5")
            If Not IsNull(rs("MEAS6")) Then iTmpMes(5) = rs("MEAS6")
            If Not IsNull(rs("MEAS7")) Then iTmpMes(6) = rs("MEAS7")
            If Not IsNull(rs("MEAS8")) Then iTmpMes(7) = rs("MEAS8")
            If Not IsNull(rs("MEAS9")) Then iTmpMes(8) = rs("MEAS9")
            If Not IsNull(rs("MEAS10")) Then iTmpMes(9) = rs("MEAS10")
            
            '10�_����̏ꍇ
            If iOldFlg = 0 Then
                sql2 = "select HSXLTSPI from TBCME019"
                sql2 = sql2 & " where HINBAN = '" & ChkHin.hinban & "'"
                sql2 = sql2 & " and MNOREVNO = " & ChkHin.mnorevno
                sql2 = sql2 & " and FACTORY = '" & ChkHin.factory & "'"
                sql2 = sql2 & " and OPECOND = '" & ChkHin.opecond & "'"
                Set rs2 = OraDB.DBCreateDynaset(sql2, ORADYN_NO_BLANKSTRIP)
                If rs2.RecordCount = 0 Then
                    Set rs2 = Nothing
                    GoTo proc_exit
                End If
                If Not IsNull(rs2("HSXLTSPI")) Then sIchi = rs2("HSXLTSPI") Else sIchi = ""
                Set rs2 = Nothing
            End If
            
            '�v�Z���ʎ擾
            iRet = KNS_CalculateMeasResult_LT(iCalcMeas, iTmpMes(), sIchi, iOldFlg)

            
            'TBCMX001
            With recX001
                .Fields("SXLLT_SMPPOS").Value = rs("POSITION")          'LT�T���v������ʒu(SXL�ʒu���)
                .Fields("SXLLT_MEASPEAK").Value = rs("MEASPEAK")        'SXLLT_����l �s�[�N�l

                .Fields("SXLLT_CALCMEAS").Value = iCalcMeas             'SXLLT_�v�Z����
            End With
                
            'TBCMX002
            With recX002
                .Fields("SXLT_SMPPOS").Value = rs("POSITION")           'SXLLT�T���v������ʒu(SXL�ʒu���)
                .Fields("SXLLT_MEASPEAK").Value = rs("MEASPEAK")        'SXLLT_����l �s�[�N�l

                '���ް�
                If iOldFlg = 1 Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(1)           'SXLLT_����l2
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(2)           'SXLLT_����l3
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(3)           'SXLLT_����l4
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(4)           'SXLLT_����l5
                '3:CE,Inside3mm
                ElseIf sIchi = "3" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(7)           'SXLLT_����l8
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(8)           'SXLLT_����l9
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(9)           'SXLLT_����l10
                '5:CE,Inside5mm
                ElseIf sIchi = "5" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(4)           'SXLLT_����l5
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(5)           'SXLLT_����l6
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(6)           'SXLLT_����l7
                'A:CE,Inside10mm
                ElseIf sIchi = "A" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(1)           'SXLLT_����l2
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(2)           'SXLLT_����l3
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(3)           'SXLLT_����l4
                '���̑�
                Else
                    '���̑��̏ꍇ�͢A:CE,Inside10mm��Ƃ���
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(1)           'SXLLT_����l2
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(2)           'SXLLT_����l3
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(3)           'SXLLT_����l4
                End If
            End With
            Set rs = Nothing
        End If
    End If

    getTBCMJ007 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ007 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



'�T�v      :����GD����(TBCMJ006)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :recX003         , O  ,c_cmzcrec         , TBCMX003�\����(GD��������_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :����GD����(TBCMJ006)�����ް����擾���AGD��������_�ް��\���̂ɾ�Ă���
'          :����GD����(TBCMJ006)�̑���f�[�^�̏����l�ł���-1��NULL�ɕύX����TBCMX003�ɓo�^����B
'����      :2005/02/15 ffc)tanabe
Private Function getTBCMJ006GD(CRYNUM As String, recXSDCW As c_cmzcrec, recX003 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ006GD"
    
    getTBCMJ006GD = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
        
    'TBCMX003
    With recX003
            .Fields("SXLGD_HSFLG").Value = vbNullString                              'SXLGDGD���茋�ʕۏ؃t���O
            .Fields("SXLGD_SMPPOS").Value = vbNullString                             'SXLGDGD�T���v������ʒu(SXL�ʒu���)
            .Fields("SXLGD_MSRSDEN").Value = vbNullString                            'SXLGDGD_���茋�� Den
            .Fields("SXLGD_MSRSLDL").Value = vbNullString                            'SXLGDGD_���茋�� L/DL
            .Fields("SXLGD_MSRSDVD2").Value = vbNullString                           'SXLGDGD_���茋�� DVD2
            .Fields("WFGD_HSFLG").Value = vbNullString                               'WFGD���茋�ʕۏ؃t���O
            .Fields("WFGD_SMPPOS").Value = vbNullString                              'WFGD�T���v������ʒu(SXL�ʒu���)
            .Fields("WFGD_MSRSDEN").Value = vbNullString                             'WFGD_���茋�� Den
            .Fields("WFGD_MSRSLDL").Value = vbNullString                             'WFGD_���茋�� L/DL
            .Fields("WFGD_MSRSDVD2").Value = vbNullString                            'WFGD_���茋�� DVD2
            
        For i = 1 To 15
            .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = vbNullString       'WFGD_����lxx L/DL1
            .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = vbNullString       'WFGD_����lxx L/DL2
            .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = vbNullString       'WFGD_����lxx L/DL3
            .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = vbNullString       'WFGD_����lxx L/DL4
            .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = vbNullString       'WFGD_����lxx L/DL5
            .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = vbNullString       'WFGD_����lxx Den1
            .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = vbNullString       'WFGD_����lxx Den2
            .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = vbNullString       'WFGD_����lxx Den3
            .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = vbNullString       'WFGD_����lxx Den4
            .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = vbNullString       'WFGD_����lxx Den5
        Next
        
        For i = 1 To 5
            .Fields("WFGD_MS01DVD2" & i).Value = vbNullString                        'WFGD_����lxx DVD2
        Next
        
    End With
        
    '-------------------- TBCMJ006�̓ǂݍ���(GD) ----------------------------------------
    sql = "select * from TBCMJ006 "
    sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
    sql = sql & "      SMPLNO = " & Trim(recXSDCW("WFSMPLIDGDCW").Value)
    sql = sql & " order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    
    'TBCMX003
    With recX003
        .Fields("SXLGD_HSFLG").Value = "1"                          'SXLGD���茋�ʕۏ؃t���O
        .Fields("SXLGD_SMPPOS").Value = rs("POSITION")              'SXLGD�T���v������ʒu(SXL�ʒu���)
        If rs("MSRSDEN") <> -1 Then
            .Fields("SXLGD_MSRSDEN").Value = rs("MSRSDEN")          'SXLGD_���茋�� Den
        End If
        If rs("MSRSLDL") <> -1 Then
            .Fields("SXLGD_MSRSLDL").Value = rs("MSRSLDL")          'SXLGD_���茋�� L/DL
        End If
        If rs("MSRSDVD2") <> -1 Then
            .Fields("SXLGD_MSRSDVD2").Value = rs("MSRSDVD2")        'SXLGD_���茋�� DVD2
        End If
        
        For i = 1 To 15
            If rs("MS" & Format(i, "00") & "DEN1") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'SXLGD_����lxx Den1
            End If
            If rs("MS" & Format(i, "00") & "DEN2") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'SXLGD_����lxx Den2
            End If
            If rs("MS" & Format(i, "00") & "DEN3") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'SXLGD_����lxx Den3
            End If
            If rs("MS" & Format(i, "00") & "DEN4") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'SXLGD_����lxx Den4
            End If
            If rs("MS" & Format(i, "00") & "DEN5") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'SXLGD_����lxx Den5
            End If
            If rs("MS" & Format(i, "00") & "LDL1") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'SXLGD_����lxx L/DL1
            End If
            If rs("MS" & Format(i, "00") & "LDL2") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'SXLGD_����lxx L/DL2
            End If
            If rs("MS" & Format(i, "00") & "LDL3") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'SXLGD_����lxx L/DL3
            End If
            If rs("MS" & Format(i, "00") & "LDL4") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'SXLGD_����lxx L/DL4
            End If
            If rs("MS" & Format(i, "00") & "LDL5") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'SXLGD_����lxx L/DL5
            End If
        Next
        
        For i = 1 To 5
            If rs("MS0" & i & "DVD2") <> -1 Then
                .Fields("WFGD_MS01DVD2" & i).Value = rs("MS0" & i & "DVD2")         'SXLGD_����lxx DVD2
            End If
        Next
        
    End With
    Set rs = Nothing

    getTBCMJ006GD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ006GD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFGD����(TBCMJ015)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :recX003         , O  ,c_cmzcrec         , TBCMX003�\����(GD��������_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFGD����(TBCMJ015)�����ް����擾���AGD��������_�ް��\���̂ɾ�Ă���
'����      :2005/02/15 ffc)tanabe
Private Function getTBCMJ015WFGD(CRYNUM As String, recXSDCW As c_cmzcrec, recX003 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ015WFGD"
    
    getTBCMJ015WFGD = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
        
    'TBCMX003
    With recX003
            .Fields("SXLGD_HSFLG").Value = vbNullString                              'SXLGDGD���茋�ʕۏ؃t���O
            .Fields("SXLGD_SMPPOS").Value = vbNullString                             'SXLGDGD�T���v������ʒu(SXL�ʒu���)
            .Fields("SXLGD_MSRSDEN").Value = vbNullString                            'SXLGDGD_���茋�� Den
            .Fields("SXLGD_MSRSLDL").Value = vbNullString                            'SXLGDGD_���茋�� L/DL
            .Fields("SXLGD_MSRSDVD2").Value = vbNullString                           'SXLGDGD_���茋�� DVD2
            .Fields("WFGD_HSFLG").Value = vbNullString                               'WFGD���茋�ʕۏ؃t���O
            .Fields("WFGD_SMPPOS").Value = vbNullString                              'WFGD�T���v������ʒu(SXL�ʒu���)
            .Fields("WFGD_MSRSDEN").Value = vbNullString                             'WFGD_���茋�� Den
            .Fields("WFGD_MSRSLDL").Value = vbNullString                             'WFGD_���茋�� L/DL
            .Fields("WFGD_MSRSDVD2").Value = vbNullString                            'WFGD_���茋�� DVD2
            
        For i = 1 To 15
            .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = vbNullString       'WFGD_����lxx L/DL1
            .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = vbNullString       'WFGD_����lxx L/DL2
            .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = vbNullString       'WFGD_����lxx L/DL3
            .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = vbNullString       'WFGD_����lxx L/DL4
            .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = vbNullString       'WFGD_����lxx L/DL5
            .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = vbNullString       'WFGD_����lxx Den1
            .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = vbNullString       'WFGD_����lxx Den2
            .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = vbNullString       'WFGD_����lxx Den3
            .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = vbNullString       'WFGD_����lxx Den4
            .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = vbNullString       'WFGD_����lxx Den5
        Next
        
        For i = 1 To 5
            .Fields("WFGD_MS01DVD2" & i).Value = vbNullString                        'WFGD_����lxx DVD2
        Next
        
    End With
        
    '-------------------- TBCMJ015�̓ǂݍ���(GD) ----------------------------------------
    sql = "select * from TBCMJ015 "
    sql = sql & " where CRYNUM = '" & CRYNUM & "'"
    sql = sql & " and   SMPLNO = '" & recXSDCW("WFSMPLIDGDCW").Value & "'"
    sql = sql & " and   HSFLG = '1'"
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    
    'TBCMX003
    With recX003
        .Fields("WFGD_HSFLG").Value = "1"                                                                 'WFGD���茋�ʕۏ؃t���O
        .Fields("WFGD_SMPPOS").Value = rs("POSITION")                                                     'WFGD�T���v������ʒu(SXL�ʒu���)
        .Fields("WFGD_MSRSDEN").Value = rs("MSRSDEN")                                                     'WFGD_���茋�� Den
        .Fields("WFGD_MSRSLDL").Value = rs("MSRSLDL")                                                     'WFGD_���茋�� L/DL
        .Fields("WFGD_MSRSDVD2").Value = rs("MSRSDVD2")                                                   'WFGD_���茋�� DVD2
        
        For i = 1 To 15
            .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'WFGD_����lxx Den1
            .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'WFGD_����lxx Den2
            .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'WFGD_����lxx Den3
            .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'WFGD_����lxx Den4
            .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'WFGD_����lxx Den5
            .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'WFGD_����lxx L/DL1
            .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'WFGD_����lxx L/DL2
            .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'WFGD_����lxx L/DL3
            .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'WFGD_����lxx L/DL4
            .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'WFGD_����lxx L/DL5
        Next
        
        For i = 1 To 5
            .Fields("WFGD_MS01DVD2" & i).Value = rs("MS0" & i & "DVD2")                                    'WFGD_����lxx DVD2
        Next
        
    End With
    Set rs = Nothing

    getTBCMJ015WFGD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ015WFGD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :SXL�m��w��(TBCMY007)ð��قɾ�Ă���SXL�̔��R�ް����擾����B
'���Ұ�    :�ϐ���          ,IO  ,�^                :����
'          :SXLID          ,I   ,String            ,SXLID
'�@�@      :sPos  �@�@�@    ,I   ,String �@         ,SXL�ʒu(TOP/BOT)   04/04/15 ooba
'          :sPattern       ,I   ,String            ,���R�ް��擾�����
'                                                   �������A : WF�����ް��擾
'                                                   �������B : ���������ް��擾
'                                                   �������C : �擾�ް��Ȃ�
'          :mesdata()      ,O   ,String            ,���R�ް�
'          :�߂�l          ,O   ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :04/02/12 ooba�@�쐬
Public Function cmbc040_GetSxlRsData(SXLID As String, sPos As String, sPattern As String, mesdata() As String) As FUNCTION_RETURN
    
    Dim sTBkbn As String        'T/B�敪
    Dim i As Integer
    Dim j As Integer
    Dim sSql As String
    Dim rs As OraDynaset
    Dim dTmpData(10) As Double   '���R(Rs)�ް�
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function cmbc040_GetSxlRsData"
    cmbc040_GetSxlRsData = FUNCTION_RETURN_FAILURE
    
    If sPos = "TOP" Then sTBkbn = "T" Else sTBkbn = "B"  '04/04/15 ooba
    
    '���R�ް��擾����݂��wA�x�̏ꍇ�AWF�����ް�(TBCMY013)���擾����B
    If sPattern = "A" Then
'''        For i = 1 To 2
'''            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"
        '�Y��SXL���A�V����يǗ�-WF<XSDCW>�̻����ID_Rs���擾�B
        '�����ID_Rs����A����]������<TBCMY013>�̔��R�����ް�(TOP��/BOT��)���擾����B
        sSql = "select MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5 "
        sSql = sSql & "from TBCMY013 "
        sSql = sSql & "where OSITEM = 'RES' "
        sSql = sSql & "and SAMPLEID in ( "
        sSql = sSql & "         select WFSMPLIDRSCW from XSDCW "
        sSql = sSql & "         where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "         and SXLIDCW = '" & SXLID & "') "
        
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount > 0 Then
            'TOP�������ް�
            If sTBkbn = "T" Then
                mesdata(1) = rs("MESDATA1")
                mesdata(2) = rs("MESDATA2")
                mesdata(3) = rs("MESDATA3")
                mesdata(4) = rs("MESDATA4")
                mesdata(5) = rs("MESDATA5")
            'BOT�������ް�
            ElseIf sTBkbn = "B" Then
                mesdata(6) = rs("MESDATA1")
                mesdata(7) = rs("MESDATA2")
                mesdata(8) = rs("MESDATA3")
                mesdata(9) = rs("MESDATA4")
                mesdata(10) = rs("MESDATA5")
            End If
        Else
            '�����ް����Ȃ��ꍇ�ʹװ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
'''        Next
    '���R�ް��擾����݂��wB�x�̏ꍇ�A���������ް�(TBCMJ002)���擾����B
    ElseIf sPattern = "B" Then
'''        For i = 1 To 2
'''            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"
        '�Y��SXL���A�V����يǗ�-WF<XSDCW>��T/B�敪�A�������ۯ�ID���擾�B
        'T/B�敪�A�������ۯ�ID����A�V����يǗ�-��ۯ�<XSDCS>�̌����ԍ��A�����ID_Rs���擾�B
        '�����ԍ��A�����ID_Rs����A������R����<TBCMJ002>�̔��R�����ް�(TOP��/BOT��)���擾����B
        sSql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 "
        sSql = sSql & "from TBCMJ002 "
        sSql = sSql & "where (CRYNUM, SMPLNO) in ( "
        sSql = sSql & "         select XTALCS, CRYSMPLIDRSCS "
        sSql = sSql & "         from XSDCS "
        sSql = sSql & "         where (TBKBNCS, CRYNUMCS) in ( "
        sSql = sSql & "                  select TBKBNCW, SMCRYNUMCW "
        sSql = sSql & "                  from XSDCW "
        sSql = sSql & "                  where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "                  and SXLIDCW = '" & SXLID & "')) "
        sSql = sSql & "and TRANCNT = ( "
        sSql = sSql & "         select max(TRANCNT) "
        sSql = sSql & "         from TBCMJ002 "
        sSql = sSql & "         where (CRYNUM, SMPLNO) in ( "
        sSql = sSql & "                  select XTALCS, CRYSMPLIDRSCS "
        sSql = sSql & "                  from XSDCS "
        sSql = sSql & "                  where (TBKBNCS, CRYNUMCS) in ( "
        sSql = sSql & "                           select TBKBNCW, SMCRYNUMCW "
        sSql = sSql & "                           from XSDCW "
        sSql = sSql & "                           where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "                           and SXLIDCW = '" & SXLID & "'))) "
    
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    
        If rs.RecordCount > 0 Then
            'TOP�������ް�
            If sTBkbn = "T" Then
                dTmpData(1) = rs("MEAS1")
                dTmpData(2) = rs("MEAS2")
                dTmpData(3) = rs("MEAS3")
                dTmpData(4) = rs("MEAS4")
                dTmpData(5) = rs("MEAS5")
                '�^�ϊ�
                For j = 1 To 5
                    mesdata(j) = CStr(dTmpData(j))
                Next
            'BOT�������ް�
            ElseIf sTBkbn = "B" Then
                dTmpData(6) = rs("MEAS1")
                dTmpData(7) = rs("MEAS2")
                dTmpData(8) = rs("MEAS3")
                dTmpData(9) = rs("MEAS4")
                dTmpData(10) = rs("MEAS5")
                '�^�ϊ�
                For j = 6 To 10
                    mesdata(j) = CStr(dTmpData(j))
                Next
            End If
        Else
            '�����ް����Ȃ��ꍇ�ʹװ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
'''        Next
    '���R�ް��擾����݂��wC�x�̏ꍇ�A�擾�����ް��Ȃ��B
    ElseIf sPattern = "C" Then
    
    End If
    
    '�擾�ް�����/-1/NULL�̎��ͽ�߰���Ă���B
'''    For i = 1 To 10
'''        If mesdata(i) = "" Or mesdata(i) = "-1" Or mesdata(i) = vbNullString Then
'''            mesdata(i) = " "
'''        End If
'''    Next
    For i = 1 To 5
        If sTBkbn = "T" Then j = i Else j = i + 5
        If mesdata(j) = "" Or mesdata(j) = "-1" Or mesdata(j) = vbNullString Then
            mesdata(j) = " "
        End If
    Next
    
    cmbc040_GetSxlRsData = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    cmbc040_GetSxlRsData = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


Private Function NtoZ2(strWk As String) As Double
    If Trim(strWk) = "" Then
        NtoZ2 = -1
        Exit Function
    End If
    NtoZ2 = CDbl(strWk)
End Function


Private Function CryRES_Judg(CRs() As Double, GarRes As Guarantee) As Double
    Dim pt As Integer

    ''RRG����
    Select Case GarRes.cPos
      Case "B", "C", "D", "E", "F", "K", "S", "Y"
          Select Case GarRes.cBunp
          Case "A", "B", "C", "M"
             ''RRG�v�Z
             CryRES_Judg = MENNAI_Cal(RES_JUDG, CRs(), GarRes, GarRes.cBunp)

          Case "", " "                                          '��߰��ǉ��@05/07/05 ooba
'Add Start 2011/10/13 Y.Hitomi
                CryRES_Judg = -1
                Exit Function
'Add End 2011/10/13 Y.Hitomi

'Del Start 2011/10/13 Y.Hitomi
''Cng Start 2011/09/19 Y.Hitomi
'            If GarRes.cCount = "1" Then
'                CryRES_Judg = 0  '�P�_����̏ꍇ�݂̂O�Ƃ���
'            Else
'                CryRES_Judg = -1
'                Exit Function
'            End If
''                CryRES_Judg = -1
''                Exit Function
''Cng End   2011/09/19 Y.Hitomi
'Del End 2011/10/13 Y.Hitomi

          Case Else
'Cng Start 2011/10/25 Y.Hitomi
                CryRES_Judg = -2
                Exit Function
'             ''RRG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
'             If Trim(GarRes.cCount) = "" Then
'                pt = 3
'             Else
'                pt = val(GarRes.cCount)
'             End If
'             CryRES_Judg = RoundUp((RGCal(CRs(), pt)), 4)
'Cng End 2011/10/25 Y.Hitomi
         End Select
      Case Else
         Select Case GarRes.cBunp
         Case "A", "B", "C", "D", "E", "M", "N"
             ''RRG�v�Z
             CryRES_Judg = MENNAI_Cal(RES_JUDG, CRs(), GarRes, GarRes.cBunp)

         Case "", " " '��߰��ǉ��@05/07/05 ooba
'Add Start 2011/10/13 Y.Hitomi
                CryRES_Judg = -1
                Exit Function
'Add End 2011/10/13 Y.Hitomi

'Del Start 2011/10/13 Y.Hitomi
''Cng Start 2011/09/19 Y.Hitomi
'            If GarRes.cCount = "1" Then
'                CryRES_Judg = 0  '�P�_����̏ꍇ�݂̂O�Ƃ���
'            Else
'                CryRES_Judg = -1
'                Exit Function
'            End If
''                CryRES_Judg = -1
''                Exit Function
''Cng End   2011/09/19 Y.Hitomi
'Del End 2011/10/13 Y.Hitomi
         Case Else
'Cng Start 2011/10/25 Y.Hitomi
                CryRES_Judg = -2
                Exit Function
'             ''RRG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
'             If Trim(GarRes.cCount) = "" Then
'                pt = 3
'             Else
'                pt = val(GarRes.cCount)
'             End If
'             CryRES_Judg = RoundUp((RGCal(CRs(), pt)), 4)
'Cng End 2011/10/25 Y.Hitomi
         End Select
    End Select
Cal_Escp:
        
End Function

Private Function CryOi_Judg(COi() As Double, GarOi As Guarantee) As Double
    Dim pt As Integer
    ReDim JData(UBound(COi())) As Double
    
    ''ORG����
    
    Select Case GarOi.cPos
      Case "B", "C", "D", "E", "F", "K", "Y"
          Select Case GarOi.cBunp
          Case "A", "B", "C"
             ''ORG�v�Z
             CryOi_Judg = MENNAI_Cal(OI_JUDG, COi(), GarOi, GarOi.cBunp)

          Case "", " "                                              '��߰��ǉ��@05/07/05 ooba
'Add Start 2011/10/13 Y.Hitomi
                CryOi_Judg = -1
                Exit Function
'Add End 2011/10/13 Y.Hitomi
             
             ''�v�Z�敪���X�y�[�X�̏ꍇ�́A�v�Z�C������s��Ȃ�
'             If GarOi.cBunp = "" Or GarOi.cBunp = " " Then         '�����ĉ��@05/07/05 ooba
'                    GoTo Cal_Escp
'Del Start 2011/10/13 Y.Hitomi
''Cng Start 2011/09/19 Y.Hitomi
'            If GarOi.cCount = "1" Then
'                CryOi_Judg = 0  '�P�_����̏ꍇ�݂̂O�Ƃ���
'            Else
'                CryOi_Judg = -1
'                Exit Function
'            End If
''                CryOi_Judg = -1
''                Exit Function
''Cng End   2011/09/19 Y.Hitomi
'Del End 2011/10/13 Y.Hitomi

'             End If                                                '�����ĉ��@05/07/05 ooba

          Case Else
'Cng Start 2011/10/25 Y.Hitomi
                CryOi_Judg = -2
                Exit Function
'             ''ORG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
'             If Trim(GarOi.cCount) = "" Then
'                pt = 3
'             Else
'                pt = val(GarOi.cCount)
'             End If
'             CryOi_Judg = RoundUp((RGCal(COi(), pt)), 4)
'Cng End 2011/10/25 Y.Hitomi

         End Select

      Case Else

         Select Case GarOi.cBunp
         Case "A", "B", "C", "D", "E", "N"
             ''ORG�v�Z
             CryOi_Judg = MENNAI_Cal(OI_JUDG, COi(), GarOi, GarOi.cBunp)

         Case "", " "                                               '��߰��ǉ��@05/07/05 ooba
'Add Start 2011/10/13 Y.Hitomi
                CryOi_Judg = -1
                Exit Function
'Add End 2011/10/13 Y.Hitomi
             
             ''�v�Z�敪���X�y�[�X�̏ꍇ�́A�v�Z�C������s��Ȃ�
'             If GarOi.cBunp = "" Or GarOi.cBunp = " " Then         '�����ĉ��@05/07/05 ooba
'                    GoTo Cal_Escp
'Del Start 2011/10/13 Y.Hitomi
''Cng Start 2011/09/19 Y.Hitomi
'            If GarOi.cCount = "1" Then
'                CryOi_Judg = 0  '�P�_����̏ꍇ�݂̂O�Ƃ���
'            Else
'                CryOi_Judg = -1
'                Exit Function
'            End If
''                CryOi_Judg = -1
''                Exit Function
''Cng End   2011/09/19 Y.Hitomi
'Del End 2011/10/13 Y.Hitomi
'Cng End   2011/09/19 Y.Hitomi
'             End If                                                '�����ĉ��@05/07/05 ooba

         Case Else
'Cng Start 2011/10/25 Y.Hitomi
                CryOi_Judg = -2
                Exit Function
'             ''ORG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
'             If Trim(GarOi.cCount) = "" Then
'                pt = 3
'             Else
'                pt = val(GarOi.cCount)
'             End If
'             CryOi_Judg = RoundUp((RGCal(COi(), pt)), 4)
'Cng End 2011/10/25 Y.Hitomi

         End Select
    End Select
Cal_Escp:

End Function

'�T�v      :BMD���т�Min�l���v�Z����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dMin          ,O   ,Double    ,Min�l
'          :strMeasPos    ,I   ,String    ,�������ב���ʒu�R�[�h�i3byte�j
'          :dMeas()       ,I   ,Double    ,����ʒu�z��
'          :�߂�l        ,O   ,Integer     ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Private Function getSXLBMDMIN(dMin As Double, strMeasPos As String, dMeas() As Double) As Integer
    Dim dConv       As Double
    Dim iMeasNum    As Integer
    Dim Index       As Integer
    Dim dForMin()   As Double
    Dim strParam    As String

    On Error GoTo Err
    getSXLBMDMIN = FUNCTION_RETURN_FAILURE

    If strMeasPos = "" Then
        dMin = -1
        getSXLBMDMIN = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    '' �������ב���ʒu�i������@�j��芷�Z�W�����擾
    strParam = GetCodeField("GP", "01", Mid(strMeasPos, 1, 1), "INFO8")
    If strParam = vbNullString Then strParam = "1"
    dConv = val(strParam)

    '' �������ב���ʒu�i����_�j�̎擾
    iMeasNum = GetMeasureNum(Mid(strMeasPos, 2, 1), 1)
    If iMeasNum < 1 Then Exit Function

    '' Min�l�v�Z
    ReDim dForMin(iMeasNum - 1)
    For Index = 0 To UBound(dForMin)
        dForMin(Index) = dMeas(Index)
    Next Index
    dMin = GetMin(dForMin) * dConv / 10000

    getSXLBMDMIN = FUNCTION_RETURN_SUCCESS
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
End Function



'�T�v      :���茋�ʂ��v�Z����i���C�t�^�C�����сj
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :iResult       ,O   ,Integer   ,�v�Z����
'          :iParam()      ,I   ,Integer   ,����l�z��
'          :sHsxLtspi     ,I   ,String    ,����ʒu         (�V�f�[�^[10�_����]��3,5,A�̂ǂꂩ��ݒ肷��)
'          :iOldFlg       ,I   ,Integer   ,���f�[�^�t���O   (���f�[�^[5�_����]��1��ݒ肷��)
'          :�߂�l        ,O   ,Integer   ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :2005/11/07 �q�� �ύX�@10�_����Ή�
Public Function KNS_CalculateMeasResult_LT(iResult As Integer, iParam() As Integer, _
                    sHsxLtspi As String, iOldFlg As Integer) As Integer
    Dim Index   As Integer
    Dim iAve    As Integer

    On Error GoTo Err
    KNS_CalculateMeasResult_LT = FUNCTION_RETURN_FAILURE
    
    '' ���f�[�^�̏ꍇ�i�T�_����j
    If iOldFlg = 1 Then
        '' �p�����[�^���̓`�F�b�N
        For Index = 0 To KNS_GetMeasureNum_LT(iOldFlg) - 1
            If iParam(Index) = DEF_PARAM_VALUE_LT Then
                Exit Function
            End If
        Next Index
        ''�R�C�S�C�T�_�̑���_��AVE�����߂�
        iAve = RoundDown((iParam(2) + iParam(3) + iParam(4)) / 3#, 0)

        '' ����_�Q��AVE�l���r�A�l�̏��������𑪒茋�ʂƂ���
        If iAve < iParam(1) Then
            iResult = iAve
        Else
            iResult = iParam(1)
        End If

    '' �V�f�[�^�̏ꍇ�i�P�O�_����j
    Else
        '' �p�����[�^���̓`�F�b�N
        For Index = 0 To KNS_GetMeasureNum_LT(iOldFlg) - 1
            If iParam(Index) = DEF_PARAM_VALUE_LT Then
                Exit Function
            End If
        Next Index

        ''' [A:Ce,Inside3mm]�̏ꍇ
        If Trim(sHsxLtspi) = "3" Then
            ''�W�C�X�C�P�O�_�̑���_��AVE�����߂�
            iAve = RoundDown((iParam(7) + iParam(8) + iParam(9)) / 3#, 0)

        ''' [A:Ce,Inside5mm]�̏ꍇ
        ElseIf Trim(sHsxLtspi) = "5" Then
            ''�T�C�U�C�V�_�̑���_��AVE�����߂�
            iAve = RoundDown((iParam(4) + iParam(5) + iParam(6)) / 3#, 0)

        ''' [A:Ce,Inside10mm]�̏ꍇ
        ElseIf Trim(sHsxLtspi) = "A" Then
            ''�Q�C�R�C�S�_�̑���_��AVE�����߂�
            iAve = RoundDown((iParam(1) + iParam(2) + iParam(3)) / 3#, 0)

        ''' ���̑��̏ꍇ��[A:Ce,Inside10mm]�̎d�l�Ƃ���
        Else
            ''�Q�C�R�C�S�_�̑���_��AVE�����߂�
            iAve = RoundDown((iParam(1) + iParam(2) + iParam(3)) / 3#, 0)

        End If
    
        '' ����_�P��AVE�l���r�A�l�̏��������𑪒茋�ʂƂ���
        If iAve < iParam(0) Then
            iResult = iAve
        Else
            iResult = iParam(0)
        End If
    End If

    KNS_CalculateMeasResult_LT = FUNCTION_RETURN_SUCCESS
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
End Function

'2008/05/30 SHINDOH------------------------------------------------------------------
'���ԍ������̐ԕ����쐬��
    ' �u���b�N�Ǘ��̍X�V�i�����֐��j
Public Function DBDRV_redXSDC3(records As typ_XSDC3_Update) As FUNCTION_RETURN

    Dim sql As String
    Dim strErrMsg As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_redXSDC3"

    DBDRV_redXSDC3 = FUNCTION_RETURN_SUCCESS

    ' �u���b�N�Ǘ��̍X�V
    With records
        .KCNTC3 = .KCNTC3 + 1
        .MODKBC3 = 1
        .FRWC3 = .FRWC3 * (-1)
        .FRLC3 = .FRLC3 * (-1)
        .FUWC3 = .FUWC3 * (-1)
        .FULC3 = .FULC3 * (-1)
        .TOLC3 = .TOLC3 * (-1)
        .TOWC3 = .TOWC3 * (-1)
        .SNDKC3 = 0
        .SUMITBC3 = 0
    End With
    If CreateXSDC3(records, strErrMsg) = FUNCTION_RETURN_FAILURE Then
        DBDRV_redXSDC3 = FUNCTION_RETURN_FAILURE
    End If
   
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_redXSDC3 = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'-----------------------------------------------------------------------------
'�T�v      :�e�[�u���uXSDC3�v��������ɐԍ������p���R�[�h�̒��o
'����      :
'-----------------------------------------------------------------------------
Public Function DBDRV_GetredXSDC3(records As typ_XSDC3_Update, t_CryNum As String, t_INPOS As Integer) As FUNCTION_RETURN
    
    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long
     '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_GetredXSDC3"

    sql = ""
    sql = sql & "select * from XSDC3"
    sql = sql & " Where CRYNUMC3='" & t_CryNum & "'"
    sql = sql & " and INPOSC3=" & t_INPOS
    sql = sql & " and KCNTC3= ( SELECT MAX(KCNTC3) FROM XSDC3 WHERE "
    sql = sql & " CRYNUMC3='" & t_CryNum & "'"
    sql = sql & " and INPOSC3=" & t_INPOS
    sql = sql & " Group by CRYNUMC3, INPOSC3 )"
    
Debug.Print "getredXSDC3 " & sql
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        DBDRV_GetredXSDC3 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        Exit Function
    End If
    With records
        If IsNull(rs.Fields("CRYNUMC3")) = False Then .CRYNUMC3 = rs.Fields("CRYNUMC3")
        If IsNull(rs.Fields("INPOSC3")) = False Then .INPOSC3 = rs.Fields("INPOSC3")
        If IsNull(rs.Fields("KCNTC3")) = False Then .KCNTC3 = rs.Fields("KCNTC3")
        If IsNull(rs.Fields("HINBC3")) = False Then .HINBC3 = rs.Fields("HINBC3")
        If IsNull(rs.Fields("REVNUMC3")) = False Then .REVNUMC3 = rs.Fields("REVNUMC3")
        If IsNull(rs.Fields("FACTORYC3")) = False Then .FACTORYC3 = rs.Fields("FACTORYC3")
        If IsNull(rs.Fields("OPEC3")) = False Then .OPEC3 = rs.Fields("OPEC3")
        If IsNull(rs.Fields("LENC3")) = False Then .LENC3 = rs.Fields("LENC3")
        If IsNull(rs.Fields("XTALC3")) = False Then .XTALC3 = rs.Fields("XTALC3")
        If IsNull(rs.Fields("SXLIDC3")) = False Then .SXLIDC3 = rs.Fields("SXLIDC3")
        If IsNull(rs.Fields("KNKTC3")) = False Then .KNKTC3 = rs.Fields("KNKTC3")
        If IsNull(rs.Fields("WKKTC3")) = False Then .WKKTC3 = rs.Fields("WKKTC3")
        If IsNull(rs.Fields("WKKBC3")) = False Then .WKKBC3 = rs.Fields("WKKBC3")
        If IsNull(rs.Fields("MACOC3")) = False Then .MACOC3 = rs.Fields("MACOC3")
        If IsNull(rs.Fields("MODKBC3")) = False Then .MODKBC3 = rs.Fields("MODKBC3")
        If IsNull(rs.Fields("SUMKBC3")) = False Then .SUMKBC3 = rs.Fields("SUMKBC3")
        If IsNull(rs.Fields("FRKNKTC3")) = False Then .FRKNKTC3 = rs.Fields("FRKNKTC3")
        If IsNull(rs.Fields("FRWKKTC3")) = False Then .FRWKKTC3 = rs.Fields("FRWKKTC3")
        If IsNull(rs.Fields("FRWKKBC3")) = False Then .FRWKKBC3 = rs.Fields("FRWKKBC3")
        If IsNull(rs.Fields("FRMACOC3")) = False Then .FRMACOC3 = rs.Fields("FRMACOC3")
        If IsNull(rs.Fields("TOWNKTC3")) = False Then .TOWNKTC3 = rs.Fields("TOWNKTC3")
        If IsNull(rs.Fields("TOWKKTC3")) = False Then .TOWKKTC3 = rs.Fields("TOWKKTC3")
        If IsNull(rs.Fields("TOMACOC3")) = False Then .TOMACOC3 = rs.Fields("TOMACOC3")
        If IsNull(rs.Fields("FRLC3")) = False Then .FRLC3 = rs.Fields("FRLC3")
        If IsNull(rs.Fields("FRWC3")) = False Then .FRWC3 = rs.Fields("FRWC3")
        If IsNull(rs.Fields("FRMC3")) = False Then .FRMC3 = rs.Fields("FRMC3")
        If IsNull(rs.Fields("FULC3")) = False Then .FULC3 = rs.Fields("FULC3")
        If IsNull(rs.Fields("FUWC3")) = False Then .FUWC3 = rs.Fields("FUWC3")
        If IsNull(rs.Fields("FUMC3")) = False Then .FUMC3 = rs.Fields("FUMC3")
        If IsNull(rs.Fields("LOSWC3")) = False Then .LOSWC3 = rs.Fields("LOSWC3")
        If IsNull(rs.Fields("LOSLC3")) = False Then .LOSLC3 = rs.Fields("LOSLC3")
        If IsNull(rs.Fields("LOSMC3")) = False Then .LOSMC3 = rs.Fields("LOSMC3")
        If IsNull(rs.Fields("TOLC3")) = False Then .TOLC3 = rs.Fields("TOLC3")
        If IsNull(rs.Fields("TOWC3")) = False Then .TOWC3 = rs.Fields("TOWC3")
        If IsNull(rs.Fields("TOMC3")) = False Then .TOMC3 = rs.Fields("TOMC3")
        If IsNull(rs.Fields("SUMITLC3")) = False Then .SUMITLC3 = rs.Fields("SUMITLC3")
        If IsNull(rs.Fields("SUMITWC3")) = False Then .SUMITWC3 = rs.Fields("SUMITWC3")
        If IsNull(rs.Fields("SUMITMC3")) = False Then .SUMITMC3 = rs.Fields("SUMITMC3")
        If IsNull(rs.Fields("MOTHINC3")) = False Then .MOTHINC3 = rs.Fields("MOTHINC3")
        If IsNull(rs.Fields("XTWORKC3")) = False Then .XTWORKC3 = rs.Fields("XTWORKC3")
        If IsNull(rs.Fields("WFWORKC3")) = False Then .WFWORKC3 = rs.Fields("WFWORKC3")
        If IsNull(rs.Fields("STATIMEC3")) = False Then .STATIMEC3 = rs.Fields("STATIMEC3")
        If IsNull(rs.Fields("STOTIMEC3")) = False Then .STOTIMEC3 = rs.Fields("STOTIMEC3")
        If IsNull(rs.Fields("ETIMEC3")) = False Then .ETIMEC3 = rs.Fields("ETIMEC3")
        If IsNull(rs.Fields("HOLDCC3")) = False Then .HOLDCC3 = rs.Fields("HOLDCC3")
        If IsNull(rs.Fields("HOLDBC3")) = False Then .HOLDBC3 = rs.Fields("HOLDBC3")
        If IsNull(rs.Fields("LDFRCC3")) = False Then .LDFRCC3 = rs.Fields("LDFRCC3")
        If IsNull(rs.Fields("LDFRBC3")) = False Then .LDFRBC3 = rs.Fields("LDFRBC3")
        If IsNull(rs.Fields("TSTAFFC3")) = False Then .TSTAFFC3 = rs.Fields("TSTAFFC3")
        If IsNull(rs.Fields("TDAYC3")) = False Then .TDAYC3 = rs.Fields("TDAYC3")
        If IsNull(rs.Fields("KSTAFFC3")) = False Then .KSTAFFC3 = rs.Fields("KSTAFFC3")
        If IsNull(rs.Fields("KDAYC3")) = False Then .KDAYC3 = rs.Fields("KDAYC3")
        If IsNull(rs.Fields("SUMITBC3")) = False Then .SUMITBC3 = rs.Fields("SUMITBC3")
        If IsNull(rs.Fields("SNDKC3")) = False Then .SNDKC3 = rs.Fields("SNDKC3")
        If IsNull(rs.Fields("SNDDAYC3")) = False Then .SNDDAYC3 = rs.Fields("SNDDAYC3")
        If IsNull(rs.Fields("SUMDAYC3")) = False Then .SUMDAYC3 = rs.Fields("SUMDAYC3")
        If IsNull(rs.Fields("PAYCLASSC3")) = False Then .PAYCLASSC3 = rs.Fields("PAYCLASSC3")
        If IsNull(rs.Fields("CUTCNTC3")) = False Then .CUTCNTC3 = rs.Fields("CUTCNTC3")
        If IsNull(rs.Fields("HINBFLGC3")) = False Then .HINBFLGC3 = rs.Fields("HINBFLGC3")
        If IsNull(rs.Fields("RPCRYNUMC3")) = False Then .RPCRYNUMC3 = rs.Fields("RPCRYNUMC3")
        If IsNull(rs.Fields("PLANTCATC3")) = False Then .PLANTCATC3 = rs.Fields("PLANTCATC3")
        End With
        rs.MoveNext
    rs.Close

    DBDRV_GetredXSDC3 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_GetredXSDC3 = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'�s�ǒ��������@08/06/27 ooba
Public Function chkFuryoLen(lJitsuLen As Long, lSeiLen As Long, lMaxLen As Long) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function chkFuryoLen"
    
    chkFuryoLen = FUNCTION_RETURN_FAILURE
    lMaxLen = 0
    
    sql = "select CTR01A9 from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '19' "
    sql = sql & "and CODEA9 = 'HARAISXL' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '�ް��Ȃ�
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    If IsNull(rs("CTR01A9")) = False Then lMaxLen = rs("CTR01A9")       '�s�ǒ������
    
    rs.Close
    
    
    '���͕s�ǒ���(�������|���i����)������l�ȉ��̏ꍇ��OK�B
    If lMaxLen = 0 Or (lJitsuLen - lSeiLen) <= lMaxLen Then
        chkFuryoLen = FUNCTION_RETURN_SUCCESS
    End If
    
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'�����C�����͈̔������@08/06/30 ooba
Public Function chkRange(lUkeLen As Long, lHaraiLen As Long, _
                                    lMaxLen As Long, lMinLen As Long) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function chkRange"
    
    chkRange = FUNCTION_RETURN_FAILURE
    lMaxLen = 0
    lMinLen = 0
    
    sql = "select CTR01A9,CTR02A9 from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '19' "
    sql = sql & "and CODEA9 = 'HARAIALT' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '�ް��Ȃ�
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    If IsNull(rs("CTR01A9")) = False Then lMaxLen = rs("CTR01A9")       '���
    If IsNull(rs("CTR02A9")) = False Then lMinLen = rs("CTR02A9")       '����
    
    rs.Close
    
    
    '���o�������Ǝ���������̍�������^�����Ɋ܂܂��ꍇ��OK�B
    '�������
    If lMaxLen = 0 Or (lHaraiLen - lUkeLen) <= lMaxLen Then
        '��������
        If lMinLen = 0 Or (lUkeLen - lHaraiLen) <= lMinLen Then
            chkRange = FUNCTION_RETURN_SUCCESS
        End If
    End If
    
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'TBCST001ں��ތ����擾���X�V�@08/06/30 ooba
Public Function GetTBCST001cnt(sBlkID As String, iCnt As Integer) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function GetTBCST001cnt"
    
    GetTBCST001cnt = FUNCTION_RETURN_FAILURE
    iCnt = 0
    
    sql = "select BLOCKID from TBCST001 "
    sql = sql & "where BLOCKID = '" & sBlkID & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    'ں��ތ���
    iCnt = rs.RecordCount
    
    rs.Close
    
    '�����ް��𑗐M�ΏۊO�Ƃ���
    If iCnt > 0 Then
        sql = "update TBCST001 "
        sql = sql & "set SENDFLAG = '5' "
        sql = sql & "where BLOCKID = '" & sBlkID & "' "
        
        If OraDB.ExecuteSQL(sql) <= 0 Then
            GoTo proc_exit
        End If
    End If
    
    GetTBCST001cnt = FUNCTION_RETURN_SUCCESS
    
    
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'TBCMX011�����񐔎擾�@08/09/12 ooba
Public Function GetTBCMX011cnt(sBlkID As String, iCnt As Integer) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function GetTBCMX011cnt"
    
    GetTBCMX011cnt = FUNCTION_RETURN_FAILURE
    
    iCnt = 1
    
    sql = "select MAX(TRANCNT) TRANCNT from TBCMX011 "
    sql = sql & "where BLOCKID = '" & sBlkID & "' "
    sql = sql & "group by BLOCKID "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount > 0 Then
        '������
        iCnt = rs("TRANCNT") + 1
    End If
    
    rs.Close

    
    GetTBCMX011cnt = FUNCTION_RETURN_SUCCESS
    
    
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'Add Start 2012/01/31 Y.Hitomi
'WF�����o�������`�F�b�N
Public Function fncChkWFLen(lUkeLen As Long, lHaraiLen As Long, _
                                    lMaxLen As Long, lMinLen As Long) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function chkRange"
    
    fncChkWFLen = FUNCTION_RETURN_FAILURE
    lMaxLen = 0
    lMinLen = 0
    
    sql = "select CTR01A9,CTR02A9 from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '19' "
    sql = sql & "and CODEA9 = 'HARAIWFLEN' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '�ް��Ȃ�
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    If IsNull(rs("CTR01A9")) = False Then lMinLen = rs("CTR01A9")       '����
    If IsNull(rs("CTR02A9")) = False Then lMaxLen = rs("CTR02A9")       '���
    
    rs.Close
    
    '���o�������Ə㉺���`�F�b�N
    If lHaraiLen >= lMinLen And lHaraiLen <= lMaxLen Then
        fncChkWFLen = FUNCTION_RETURN_SUCCESS
    End If
    
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
'SIRD��s�]���u���b�NSXL�o�ח����`�F�b�N
Public Function FncChkSird(sBlockId As String) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function chkRange"
    
    FncChkSird = FUNCTION_RETURN_FAILURE
    
    sql = "select SIRDKBNY3 from XODY3 "
    sql = sql & "where XTALNOY3 = '" & sBlockId & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '�ް��Ȃ�
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    
    '�Ώۃu���b�N���ASIRD�]���u���b�N�łȂ���΃`�F�b�NOK
    If rs("SIRDKBNY3") <> "2" Then
        FncChkSird = FUNCTION_RETURN_SUCCESS
    End If
    
    rs.Close
    

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
'Add End 2012/01/31 Y.Hitomi

