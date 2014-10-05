Attribute VB_Name = "s_cmzcCrystal"
Option Explicit
'                                     2001/05/31
'================================================
' �N���X�E���[�U��`�^�̕ϊ��v���V�[�W��
' ��`���e: 060207_�����Ǘ�
'================================================


'------ �e�[�u����:TBCME037    ---- �������

' ���[�U��`�^���N���X�ɕϊ�����(�������)
Public Function u2c_TBCME037(data As typ_TBCME037) As c_TBCME037
Dim cls As New c_TBCME037       '�ϊ���N���X

    cls.CRYNUM = data.CRYNUM                '�����ԍ�
    cls.KRPROCCD = data.KRPROCCD            '�Ǘ��H���R�[�h
    cls.PROCCD = data.PROCCD                '�H���R�[�h
    cls.RPHINBAN = data.RPHINBAN            '�˂炢�i��
    cls.RPREVNUM = data.RPREVNUM            '�˂炢�i�Ԑ��i�ԍ������ԍ�
    cls.RPFACT = data.RPFACT                '�˂炢�i�ԍH��
    cls.RPOPCOND = data.RPOPCOND            '�˂炢�i�ԑ��Ə���
    cls.PRODCOND = data.PRODCOND            '�������
    cls.PGID = data.PGID                    '�o�f�|�h�c
    cls.UPLENGTH = data.UPLENGTH            '���グ����
    cls.TOPLENG = data.TOPLENG              '�s�n�o����
    cls.BODYLENG = data.BODYLENG            '��������
    cls.BOTLENG = data.BOTLENG              '�a�n�s����
    cls.FREELENG = data.FREELENG            '�t���[��
    cls.DIAMETER = data.DIAMETER            '���a
    cls.CHARGE = data.CHARGE                '�`���[�W��
    cls.SEED = data.SEED                    '�V�[�h
    cls.ADDDPCLS = data.ADDDPCLS            '�ǉ��h�[�v���
    cls.ADDDPPOS = data.ADDDPPOS            '�ǉ��h�[�v�ʒu
    cls.ADDDPVAL = data.ADDDPVAL            '�ǉ��h�[�v��
    'cls.REGDATE = data.REGDATE              '�o�^���t
    'cls.UPDDATE = data.UPDDATE              '�X�V���t
    'cls.SENDFLAG = data.SENDFLAG            '���M�t���O
    'cls.SENDDATE = data.SENDDATE            '���M���t

    Set u2c_TBCME037 = cls
End Function


' �N���X�����[�U��`�^�ɕϊ�����(�������)
Public Function c2u_TBCME037(cls As c_TBCME037) As typ_TBCME037
Dim data As typ_TBCME037        '�ϊ��惆�[�U��`�^

    data.CRYNUM = cls.CRYNUM                '�����ԍ�
    data.KRPROCCD = cls.KRPROCCD            '�Ǘ��H���R�[�h
    data.PROCCD = cls.PROCCD                '�H���R�[�h
    data.RPHINBAN = cls.RPHINBAN            '�˂炢�i��
    data.RPREVNUM = cls.RPREVNUM            '�˂炢�i�Ԑ��i�ԍ������ԍ�
    data.RPFACT = cls.RPFACT                '�˂炢�i�ԍH��
    data.RPOPCOND = cls.RPOPCOND            '�˂炢�i�ԑ��Ə���
    data.PRODCOND = cls.PRODCOND            '�������
    data.PGID = cls.PGID                    '�o�f�|�h�c
    data.UPLENGTH = cls.UPLENGTH            '���グ����
    data.TOPLENG = cls.TOPLENG              '�s�n�o����
    data.BODYLENG = cls.BODYLENG            '��������
    data.BOTLENG = cls.BOTLENG              '�a�n�s����
    data.FREELENG = cls.FREELENG            '�t���[��
    data.DIAMETER = cls.DIAMETER            '���a
    data.CHARGE = cls.CHARGE                '�`���[�W��
    data.SEED = cls.SEED                    '�V�[�h
    data.ADDDPCLS = cls.ADDDPCLS            '�ǉ��h�[�v���
    data.ADDDPPOS = cls.ADDDPPOS            '�ǉ��h�[�v�ʒu
    data.ADDDPVAL = cls.ADDDPVAL            '�ǉ��h�[�v��
    'data.REGDATE = cls.REGDATE              '�o�^���t
    'data.UPDDATE = cls.UPDDATE              '�X�V���t
    'data.SENDFLAG = cls.SENDFLAG            '���M�t���O
    'data.SENDDATE = cls.SENDDATE            '���M���t

    c2u_TBCME037 = data
End Function


'------ �e�[�u����:TBCME038    ---- �u���b�N�݌v

' ���[�U��`�^���N���X�ɕϊ�����(�������)
Public Function u2c_TBCME038(data As typ_TBCME038) As c_TBCME038
Dim cls As New c_TBCME038       '�ϊ���N���X

    cls.CRYNUM = data.CRYNUM                '�����ԍ�
    cls.IngotPos = data.IngotPos            '�������J�n�ʒu
    cls.LENGTH = data.LENGTH                '����
    cls.USECLASS = data.USECLASS            '�g�p�敪
    'cls.REGDATE = data.REGDATE              '�o�^���t
    'cls.UPDDATE = data.UPDDATE              '�X�V���t
    'cls.SENDFLAG = data.SENDFLAG            '���M�t���O
    'cls.SENDDATE = data.SENDDATE            '���M���t

    Set u2c_TBCME038 = cls
End Function


' �N���X�����[�U��`�^�ɕϊ�����(�u���b�N�݌v)
Public Function c2u_TBCME038(cls As c_TBCME038) As typ_TBCME038
Dim data As typ_TBCME038        '�ϊ��惆�[�U��`�^

    data.CRYNUM = cls.CRYNUM                '�����ԍ�
    data.IngotPos = cls.IngotPos            '�������J�n�ʒu
    data.LENGTH = cls.LENGTH                '����
    data.USECLASS = cls.USECLASS            '�g�p�敪
    'data.REGDATE = cls.REGDATE              '�o�^���t
    'data.UPDDATE = cls.UPDDATE              '�X�V���t
    'data.SENDFLAG = cls.SENDFLAG            '���M�t���O
    'data.SENDDATE = cls.SENDDATE            '���M���t

    c2u_TBCME038 = data
End Function


'------ �e�[�u����:TBCME039    ---- �i�Ԑ݌v

' ���[�U��`�^���N���X�ɕϊ�����(�������)
Public Function u2c_TBCME039(data As typ_TBCME039) As c_TBCME039
Dim cls As New c_TBCME039       '�ϊ���N���X

    cls.CRYNUM = data.CRYNUM                '�����ԍ�
    cls.IngotPos = data.IngotPos            '�������J�n�ʒu
    cls.HINBAN = data.HINBAN                '�i��
    cls.REVNUM = data.REVNUM                '�����ԍ�
    cls.FACT = data.FACT                    '�H��
    cls.OPCOND = data.OPCOND                '���Ə���
    cls.LENGTH = data.LENGTH                '����
    cls.USECLASS = data.USECLASS            '�g�p�敪
    'cls.REGDATE = data.REGDATE              '�o�^���t
    'cls.UPDDATE = data.UPDDATE              '�X�V���t
    'cls.SENDFLAG = data.SENDFLAG            '���M�t���O
    'cls.SENDDATE = data.SENDDATE            '���M���t

    Set u2c_TBCME039 = cls
End Function


' �N���X�����[�U��`�^�ɕϊ�����(�i�Ԑ݌v)
Public Function c2u_TBCME039(cls As c_TBCME039) As typ_TBCME039
Dim data As typ_TBCME039        '�ϊ��惆�[�U��`�^

    data.CRYNUM = cls.CRYNUM                '�����ԍ�
    data.IngotPos = cls.IngotPos            '�������J�n�ʒu
    data.HINBAN = cls.HINBAN                '�i��
    data.REVNUM = cls.REVNUM                '�����ԍ�
    data.FACT = cls.FACT                    '�H��
    data.OPCOND = cls.OPCOND                '���Ə���
    data.LENGTH = cls.LENGTH                '����
    data.USECLASS = cls.USECLASS            '�g�p�敪
    'data.REGDATE = cls.REGDATE              '�o�^���t
    'data.UPDDATE = cls.UPDDATE              '�X�V���t
    'data.SENDFLAG = cls.SENDFLAG            '���M�t���O
    'data.SENDDATE = cls.SENDDATE            '���M���t

    c2u_TBCME039 = data
End Function


'------ �e�[�u����:TBCME040    ---- �u���b�N�Ǘ�

' ���[�U��`�^���N���X�ɕϊ�����(�������)
Public Function u2c_cmzc001b(data As typ_TBCME040) As c_cmzc001b
Dim cls As New c_cmzc001b       '�ϊ���N���X

    cls.CRYNUM = data.CRYNUM                '�����ԍ�
    cls.IngotPos = data.IngotPos            '�������J�n�ʒu
    cls.LENGTH = data.LENGTH                '����
    cls.BLOCKID = data.BLOCKID              '�u���b�NID
    cls.KRPROCCD = data.KRPROCCD            '���݊Ǘ��H��
    cls.NOWPROC = data.NOWPROC              '���ݍH��
    cls.LPKRPROCCD = data.LPKRPROCCD        '�ŏI�ʉߊǗ��H��
    cls.LASTPASS = data.LASTPASS            '�ŏI�ʉߍH��
    cls.DELCLS = data.DELCLS                '�폜�敪
    cls.LSTATCLS = data.LSTATCLS            '�ŏI��ԋ敪
    cls.RSTATCLS = data.RSTATCLS            '������ԋ敪
    cls.HOLDCLS = data.HOLDCLS              '�z�[���h�敪
    cls.BDCAUS = data.BDCAUS                '�s�Ǘ��R
    'cls.REGDATE = data.REGDATE              '�o�^���t
    'cls.UPDDATE = data.UPDDATE              '�X�V���t
    'cls.SUMMITSENDFLAG = data.SUMMITSENDFLAG 'SUMMIT���M�t���O
    'cls.SENDFLAG = data.SENDFLAG            '���M�t���O
    'cls.SENDDATE = data.SENDDATE            '���M���t

    Set u2c_cmzc001b = cls
End Function


' �N���X�����[�U��`�^�ɕϊ�����(�u���b�N�Ǘ�)
Public Function c2u_TBCME040(cls As c_cmzc001b) As typ_TBCME040
Dim data As typ_TBCME040        '�ϊ��惆�[�U��`�^

    data.CRYNUM = cls.CRYNUM                '�����ԍ�
    data.IngotPos = cls.IngotPos            '�������J�n�ʒu
    data.LENGTH = cls.LENGTH                '����
    data.BLOCKID = cls.BLOCKID              '�u���b�NID
    data.KRPROCCD = cls.KRPROCCD            '���݊Ǘ��H��
    data.NOWPROC = cls.NOWPROC              '���ݍH��
    data.LPKRPROCCD = cls.LPKRPROCCD        '�ŏI�ʉߊǗ��H��
    data.LASTPASS = cls.LASTPASS            '�ŏI�ʉߍH��
    data.DELCLS = cls.DELCLS                '�폜�敪
    data.LSTATCLS = cls.LSTATCLS            '�ŏI��ԋ敪
    data.RSTATCLS = cls.RSTATCLS            '������ԋ敪
    data.HOLDCLS = cls.HOLDCLS              '�z�[���h�敪
    data.BDCAUS = cls.BDCAUS                '�s�Ǘ��R
    'data.REGDATE = cls.REGDATE              '�o�^���t
    'data.UPDDATE = cls.UPDDATE              '�X�V���t
    'data.SUMMITSENDFLAG = cls.SUMMITSENDFLAG 'SUMMIT���M�t���O
    'data.SENDFLAG = cls.SENDFLAG            '���M�t���O
    'data.SENDDATE = cls.SENDDATE            '���M���t

    c2u_TBCME040 = data
End Function


'------ �e�[�u����:TBCME041    ---- �i�ԊǗ�

' ���[�U��`�^���N���X�ɕϊ�����(�������)
Public Function u2c_cmzc001d(data As typ_TBCME041) As c_cmzc001d
Dim cls As New c_cmzc001d       '�ϊ���N���X

    cls.CRYNUM = data.CRYNUM                '�����ԍ�
    cls.IngotPos = data.IngotPos            '�������J�n�ʒu
    cls.HINBAN = data.HINBAN                '�i��
    cls.REVNUM = data.REVNUM                '���i�ԍ������ԍ�
    cls.factory = data.factory              '�H��
    cls.opecond = data.opecond              '���Ə���
    cls.LENGTH = data.LENGTH                '����
    'cls.REGDATE = data.REGDATE              '�o�^���t
    'cls.UPDDATE = data.UPDDATE              '�X�V���t
    'cls.SENDFLAG = data.SENDFLAG            '���M�t���O
    'cls.SENDDATE = data.SENDDATE            '���M���t

    Set u2c_cmzc001d = cls
End Function


' �N���X�����[�U��`�^�ɕϊ�����(�i�ԊǗ�)
Public Function c2u_TBCME041(cls As c_cmzc001d) As typ_TBCME041
Dim data As typ_TBCME041        '�ϊ��惆�[�U��`�^

    data.CRYNUM = cls.CRYNUM                '�����ԍ�
    data.IngotPos = cls.IngotPos            '�������J�n�ʒu
    data.HINBAN = cls.HINBAN                '�i��
    data.REVNUM = cls.REVNUM                '���i�ԍ������ԍ�
    data.factory = cls.factory              '�H��
    data.opecond = cls.opecond              '���Ə���
    data.LENGTH = cls.LENGTH                '����
    'data.REGDATE = cls.REGDATE              '�o�^���t
    'data.UPDDATE = cls.UPDDATE              '�X�V���t
    'data.SENDFLAG = cls.SENDFLAG            '���M�t���O
    'data.SENDDATE = cls.SENDDATE            '���M���t

    c2u_TBCME041 = data
End Function


'------ �e�[�u����:TBCME042    ---- SXL�Ǘ�

' ���[�U��`�^���N���X�ɕϊ�����(�������)
Public Function u2c_TBCME042(data As typ_TBCME042) As c_TBCME042
Dim cls As New c_TBCME042       '�ϊ���N���X

    cls.CRYNUM = data.CRYNUM                '�����ԍ�
    cls.IngotPos = data.IngotPos            '�������J�n�ʒu
    cls.LENGTH = data.LENGTH                '����
    cls.SXLID = data.SXLID                  'SXLID
    cls.KRPROCCD = data.KRPROCCD            '�Ǘ��H��
    cls.NOWPROC = data.NOWPROC              '���ݍH��
    cls.LPKRPROCCD = data.LPKRPROCCD        '�ŏI�ʉߊǗ��H��
    cls.LASTPASS = data.LASTPASS            '�ŏI�ʉߍH��
    cls.DELCLS = data.DELCLS                '�폜�敪
    cls.LSTATCLS = data.LSTATCLS            '�ŏI��ԋ敪
    cls.HOLDCLS = data.HOLDCLS              '�z�[���h�敪
    cls.HINBAN = data.HINBAN                '�i��
    cls.REVNUM = data.REVNUM                '���i�ԍ������ԍ�
    cls.factory = data.factory              '�H��
    cls.opecond = data.opecond              '���Ə���
    cls.Count = data.Count                  '����
    'cls.REGDATE = data.REGDATE              '�o�^���t
    'cls.UPDDATE = data.UPDDATE              '�X�V���t
    'cls.SUMMITSENDFLAG = data.SUMMITSENDFLAG 'SUMMIT���M�t���O
    'cls.SENDFLAG = data.SENDFLAG            '���M�t���O
    'cls.SENDDATE = data.SENDDATE            '���M���t

    Set u2c_TBCME042 = cls
End Function


' �N���X�����[�U��`�^�ɕϊ�����(SXL�Ǘ�)
Public Function c2u_TBCME042(cls As c_TBCME042) As typ_TBCME042
Dim data As typ_TBCME042        '�ϊ��惆�[�U��`�^

    data.CRYNUM = cls.CRYNUM                '�����ԍ�
    data.IngotPos = cls.IngotPos            '�������J�n�ʒu
    data.LENGTH = cls.LENGTH                '����
    data.SXLID = cls.SXLID                  'SXLID
    data.KRPROCCD = cls.KRPROCCD            '�Ǘ��H��
    data.NOWPROC = cls.NOWPROC              '���ݍH��
    data.LPKRPROCCD = cls.LPKRPROCCD        '�ŏI�ʉߊǗ��H��
    data.LASTPASS = cls.LASTPASS            '�ŏI�ʉߍH��
    data.DELCLS = cls.DELCLS                '�폜�敪
    data.LSTATCLS = cls.LSTATCLS            '�ŏI��ԋ敪
    data.HOLDCLS = cls.HOLDCLS              '�z�[���h�敪
    data.HINBAN = cls.HINBAN                '�i��
    data.REVNUM = cls.REVNUM                '���i�ԍ������ԍ�
    data.factory = cls.factory              '�H��
    data.opecond = cls.opecond              '���Ə���
    data.Count = cls.Count                  '����
    'data.REGDATE = cls.REGDATE              '�o�^���t
    'data.UPDDATE = cls.UPDDATE              '�X�V���t
    'data.SUMMITSENDFLAG = cls.SUMMITSENDFLAG 'SUMMIT���M�t���O
    'data.SENDFLAG = cls.SENDFLAG            '���M�t���O
    'data.SENDDATE = cls.SENDDATE            '���M���t

    c2u_TBCME042 = data
End Function


'------ �e�[�u����:XSDCS    ---- �V�T���v���Ǘ��i�u���b�N�j

' ���[�U��`�^���N���X�ɕϊ�����(�������)
Public Function u2c_cmzc001e(data As typ_XSDCS) As c_cmzc001e
'Public Function u2c_cmzc001e(data As typ_TBCME043) As c_cmzc001e
Dim cls As New c_cmzc001e       '�ϊ���N���X
    
    cls.CRYNUMCS = data.CRYNUMCS                '�u���b�NID
    cls.SMPKBNCS = data.SMPKBNCS                '�T���v���敪
    cls.TBKBNCS = data.TBKBNCS                  'T/B�敪
    cls.REPSMPLIDCS = data.REPSMPLIDCS          '�T���v��No
    cls.XTALCS = data.XTALCS                    '�����ԍ�
    cls.INPOSCS = data.INPOSCS                  '�������ʒu
    cls.HINBCS = data.HINBCS                    '�i��
    cls.REVNUMCS = data.REVNUMCS                '���i�ԍ������ԍ�
    cls.FACTORYCS = data.FACTORYCS              '�H��
    cls.OPECS = data.OPECS                      '���Ə���
    cls.KTKBNCS = data.KTKBNCS                  '�m��敪
    cls.SMPLUMU = data.SMPLUMU                  '�T���v���L���敪
    cls.BLKKTFLAGCS = data.BLKKTFLAGCS          '�u���b�N�m��t���O
    cls.CRYSMPLIDRSCS = data.CRYSMPLIDRSCS      '�T���v��ID
    cls.CRYSMPLIDRS1CS = data.CRYSMPLIDRS1CS    '����T���v��ID1
    cls.CRYSMPLIDRS2CS = data.CRYSMPLIDRS2CS    '����T���v��ID2
    cls.CRYINDRSCS = data.CRYINDRSCS            '���FLG(Rs)
    cls.CRYRESRS1CS = data.CRYRESRS1CS          '����FLG1(Rs)
    cls.CRYRESRS2CS = data.CRYRESRS2CS          '����FLG2(Rs)
    cls.CRYSMPLIDOICS = data.CRYSMPLIDOICS      '�T���v��ID(Oi)
    cls.CRYINDOICS = data.CRYINDOICS            '���FLG(Oi)
    cls.CRYRESOICS = data.CRYRESOICS            '����FLG(Oi)
    cls.CRYSMPLIDB1CS = data.CRYSMPLIDB1CS      '�T���v��ID(B1)
    cls.CRYINDB1CS = data.CRYINDB1CS            '���FLG(B1)
    cls.CRYRESB1CS = data.CRYRESB1CS            '����FLG(B1)
    cls.CRYSMPLIDB2CS = data.CRYSMPLIDB2CS      '�T���v��ID(B2)
    cls.CRYINDB2CS = data.CRYINDB2CS            '���FLG(B2)
    cls.CRYRESB2CS = data.CRYRESB2CS            '����FLG(B2)
    cls.CRYSMPLIDB3CS = data.CRYSMPLIDB3CS      '�T���v��ID(B3)
    cls.CRYINDB3CS = data.CRYINDB3CS            '���FLG(B3)
    cls.CRYRESB3CS = data.CRYRESB3CS            '����FLG(B3)
    cls.CRYSMPLIDL1CS = data.CRYSMPLIDL1CS      '�T���v��ID(L1)
    cls.CRYINDL1CS = data.CRYINDL1CS            '���FLG(L1)
    cls.CRYRESL1CS = data.CRYRESL1CS            '����FLG(L1)
    cls.CRYSMPLIDL2CS = data.CRYSMPLIDL2CS      '�T���v��ID(L2)
    cls.CRYINDL2CS = data.CRYINDL2CS            '���FLG(L2)
    cls.CRYRESL2CS = data.CRYRESL2CS            '����FLG(L2)
    cls.CRYSMPLIDL3CS = data.CRYSMPLIDL3CS      '�T���v��ID(L3)
    cls.CRYINDL3CS = data.CRYINDL3CS            '���FLG(L3)
    cls.CRYRESL3CS = data.CRYRESL3CS            '����FLG(L3)
    cls.CRYSMPLIDL4CS = data.CRYSMPLIDL4CS      '�T���v��ID(L4)
    cls.CRYINDL4CS = data.CRYINDL4CS            '���FLG(L4)
    cls.CRYRESL4CS = data.CRYRESL4CS            '����FLG(L4)
    cls.CRYSMPLIDCSCS = data.CRYSMPLIDCSCS      '�T���v��ID(Cs)
    cls.CRYINDCSCS = data.CRYINDCSCS            '���FLG(Cs)
    cls.CRYRESCSCS = data.CRYRESCSCS            '����FLG(Cs)
    cls.CRYSMPLIDGDCS = data.CRYSMPLIDGDCS      '�T���v��ID(GD)
    cls.CRYINDGDCS = data.CRYINDGDCS            '���FLG(GD)
    cls.CRYRESGDCS = data.CRYRESGDCS            '����FLG(GD)
    cls.CRYSMPLIDTCS = data.CRYSMPLIDTCS        '�T���v��ID(T)
    cls.CRYINDTCS = data.CRYINDTCS              '���FLG(T)
    cls.CRYRESTCS = data.CRYRESTCS              '����FLG(T)
    cls.CRYSMPLIDEPCS = data.CRYSMPLIDEPCS      '�T���v��ID(EPD)
    cls.CRYINDEPCS = data.CRYINDEPCS            '���FLG(EPD)
    cls.CRYRESEPCS = data.CRYRESEPCS            '����FLG(EPD)
    cls.SMPLNUMCS = data.SMPLNUMCS              '�T���v������
    cls.SMPLPATCS = data.SMPLPATCS              '�T���v���p�^�[��
    'cls.REGDATE = data.REGDATE              '�o�^���t
    'cls.UPDDATE = data.UPDDATE              '�X�V���t
    'cls.SENDFLAG = data.SENDFLAG            '���M�t���O
    'cls.SENDDATE = data.SENDDATE            '���M���t

    Set u2c_cmzc001e = cls
End Function


' �N���X�����[�U��`�^�ɕϊ�����(�����T���v���Ǘ�)
'Public Function c2u_TBCME043(cls As c_cmzc001e) As typ_TBCME043
'Dim data As typ_TBCME043        '�ϊ��惆�[�U��`�^
'
'    data.CRYNUM = cls.CRYNUM                '�����ԍ�
'    data.IngotPos = cls.IngotPos            '�������ʒu
'    data.SMPKBN = cls.SMPKBN                '�T���v���敪
'    data.SMPLNO = cls.SMPLNO                '�T���v��No
'    data.HINBAN = cls.HINBAN                '�i��
'    data.REVNUM = cls.REVNUM                '���i�ԍ������ԍ�
'    data.factory = cls.factory              '�H��
'    data.opecond = cls.opecond              '���Ə���
'    data.SMPLUMU = cls.SMPLUMU              '�T���v���L���敪
'    data.CRYINDRS = cls.CRYINDRS            '���������w���iRs)
'    data.CRYINDOI = cls.CRYINDOI            '���������w���iOi)
'    data.CRYINDB1 = cls.CRYINDB1            '���������w���iB1)
'    data.CRYINDB2 = cls.CRYINDB2            '���������w���iB2�j
'    data.CRYINDB3 = cls.CRYINDB3            '���������w���iB3)
'    data.CRYINDL1 = cls.CRYINDL1            '���������w���iL1)
'    data.CRYINDL2 = cls.CRYINDL2            '���������w���iL2)
'    data.CRYINDL3 = cls.CRYINDL3            '���������w���iL3)
'    data.CRYINDL4 = cls.CRYINDL4            '���������w���iL4)
'    data.CRYINDCS = cls.CRYINDCS            '���������w���iCs)
'    data.CRYINDGD = cls.CRYINDGD            '���������w���iGD)
'    data.CRYINDT = cls.CRYINDT              '���������w���iT)
'    data.CRYINDEP = cls.CRYINDEP            '���������w���iEPD)
'    data.CRYRESRS = cls.CRYRESRS            '�����������сiRs)
'    data.CRYRESOI = cls.CRYRESOI            '�����������сiOi)
'    data.CRYRESB1 = cls.CRYRESB1            '�����������сiB1)
'    data.CRYRESB2 = cls.CRYRESB2            '�����������сiB2�j
'    data.CRYRESB3 = cls.CRYRESB3            '�����������сiB3)
'    data.CRYRESL1 = cls.CRYRESL1            '�����������сiL1)
'    data.CRYRESL2 = cls.CRYRESL2            '�����������сiL2)
'    data.CRYRESL3 = cls.CRYRESL3            '�����������сiL3)
'    data.CRYRESL4 = cls.CRYRESL4            '�����������сiL4)
'    data.CRYRESCS = cls.CRYRESCS            '�����������сiCs)
'    data.CRYRESGD = cls.CRYRESGD            '�����������сiGD)
'    data.CRYREST = cls.CRYREST              '�����������сiT)
'    data.CRYRESEP = cls.CRYRESEP            '�����������сiEPD)
'    data.SMPLNUM = cls.SMPLNUM              '�T���v������
'    data.SMPLPAT = cls.SMPLPAT              '�T���v���p�^�[��
'    'data.REGDATE = cls.REGDATE              '�o�^���t
'    'data.UPDDATE = cls.UPDDATE              '�X�V���t
'    'data.SENDFLAG = cls.SENDFLAG            '���M�t���O
'    'data.SENDDATE = cls.SENDDATE            '���M���t
'
'    c2u_TBCME043 = data
'End Function

Public Function c2u_XSDCS(cls As c_cmzc001e) As typ_XSDCS
Dim data As typ_XSDCS        '�ϊ��惆�[�U��`�^

    data.CRYNUMCS = cls.CRYNUMCS                '�u���b�NID
    data.SMPKBNCS = cls.SMPKBNCS                '�T���v���敪
    data.TBKBNCS = cls.TBKBNCS                  'T/B�敪
    data.REPSMPLIDCS = cls.REPSMPLIDCS          '�T���v��No
    data.XTALCS = cls.XTALCS                    '�����ԍ�
    data.INPOSCS = cls.INPOSCS                  '�������ʒu
    data.HINBCS = cls.HINBCS                    '�i��
    data.REVNUMCS = cls.REVNUMCS                '���i�ԍ������ԍ�
    data.FACTORYCS = cls.FACTORYCS              '�H��
    data.OPECS = cls.OPECS                      '���Ə���
    data.KTKBNCS = cls.KTKBNCS                  '�m��敪
    data.SMPLUMU = cls.SMPLUMU                  '�T���v���L���敪
    data.BLKKTFLAGCS = cls.BLKKTFLAGCS          '�u���b�N�m��t���O
    data.CRYSMPLIDRSCS = cls.CRYSMPLIDRSCS      '�T���v��ID
    data.CRYSMPLIDRS1CS = cls.CRYSMPLIDRS1CS    '����T���v��ID1
    data.CRYSMPLIDRS2CS = cls.CRYSMPLIDRS2CS    '����T���v��ID2
    data.CRYINDRSCS = cls.CRYINDRSCS            '���FLG(Rs)
    data.CRYRESRS1CS = cls.CRYRESRS1CS          '����FLG1(Rs)
    data.CRYRESRS2CS = cls.CRYRESRS2CS          '����FLG2(Rs)
    data.CRYSMPLIDOICS = cls.CRYSMPLIDOICS      '�T���v��ID(Oi)
    data.CRYINDOICS = cls.CRYINDOICS            '���FLG(Oi)
    data.CRYRESOICS = cls.CRYRESOICS            '����FLG(Oi)
    data.CRYSMPLIDB1CS = cls.CRYSMPLIDB1CS      '�T���v��ID(B1)
    data.CRYINDB1CS = cls.CRYINDB1CS            '���FLG(B1)
    data.CRYRESB1CS = cls.CRYRESB1CS            '����FLG(B1)
    data.CRYSMPLIDB2CS = cls.CRYSMPLIDB2CS      '�T���v��ID(B2)
    data.CRYINDB2CS = cls.CRYINDB2CS            '���FLG(B2)
    data.CRYRESB2CS = cls.CRYRESB2CS            '����FLG(B2)
    data.CRYSMPLIDB3CS = cls.CRYSMPLIDB3CS      '�T���v��ID(B3)
    data.CRYINDB3CS = cls.CRYINDB3CS            '���FLG(B3)
    data.CRYRESB3CS = cls.CRYRESB3CS            '����FLG(B3)
    data.CRYSMPLIDL1CS = cls.CRYSMPLIDL1CS      '�T���v��ID(L1)
    data.CRYINDL1CS = cls.CRYINDL1CS            '���FLG(L1)
    data.CRYRESL1CS = cls.CRYRESL1CS            '����FLG(L1)
    data.CRYSMPLIDL2CS = cls.CRYSMPLIDL2CS      '�T���v��ID(L2)
    data.CRYINDL2CS = cls.CRYINDL2CS            '���FLG(L2)
    data.CRYRESL2CS = cls.CRYRESL2CS            '����FLG(L2)
    data.CRYSMPLIDL3CS = cls.CRYSMPLIDL3CS      '�T���v��ID(L3)
    data.CRYINDL3CS = cls.CRYINDL3CS            '���FLG(L3)
    data.CRYRESL3CS = cls.CRYRESL3CS            '����FLG(L3)
    data.CRYSMPLIDL4CS = cls.CRYSMPLIDL4CS      '�T���v��ID(L4)
    data.CRYINDL4CS = cls.CRYINDL4CS            '���FLG(L4)
    data.CRYRESL4CS = cls.CRYRESL4CS            '����FLG(L4)
    data.CRYSMPLIDCSCS = cls.CRYSMPLIDCSCS      '�T���v��ID(Cs)
    data.CRYINDCSCS = cls.CRYINDCSCS            '���FLG(Cs)
    data.CRYRESCSCS = cls.CRYRESCSCS            '����FLG(Cs)
    data.CRYSMPLIDGDCS = cls.CRYSMPLIDGDCS      '�T���v��ID(GD)
    data.CRYINDGDCS = cls.CRYINDGDCS            '���FLG(GD)
    data.CRYRESGDCS = cls.CRYRESGDCS            '����FLG(GD)
    data.CRYSMPLIDTCS = cls.CRYSMPLIDTCS        '�T���v��ID(T)
    data.CRYINDTCS = cls.CRYINDTCS              '���FLG(T)
    data.CRYRESTCS = cls.CRYRESTCS              '����FLG(T)
    data.CRYSMPLIDEPCS = cls.CRYSMPLIDEPCS      '�T���v��ID(EPD)
    data.CRYINDEPCS = cls.CRYINDEPCS            '���FLG(EPD)
    data.CRYRESEPCS = cls.CRYRESEPCS            '����FLG(EPD)
    data.SMPLNUMCS = cls.SMPLNUMCS              '�T���v������
    data.SMPLPATCS = cls.SMPLPATCS              '�T���v���p�^�[��
    'data.REGDATE = cls.REGDATE                 '�o�^���t
    'data.UPDDATE = cls.UPDDATE                 '�X�V���t
    'data.SENDFLAG = cls.SENDFLAG               '���M�t���O
    'data.SENDDATE = cls.SENDDATE               '���M���t

    c2u_XSDCS = data
End Function

''2003/09/02 SystemBrain �T���v���Ǘ��ύX
''------ �e�[�u����:TBCME044    ---- WF�T���v���Ǘ�
'
'' ���[�U��`�^���N���X�ɕϊ�����(�������)
'Public Function u2c_cmzc001f(data As typ_TBCME044) As c_cmzc001f
'Dim cls As New c_cmzc001f       '�ϊ���N���X
'
'    cls.CRYNUM = data.CRYNUM                '�����ԍ�
'    cls.INGOTPOS = data.INGOTPOS            '�������ʒu
'    cls.SMPKBN = data.SMPKBN                '�T���v���敪
'    cls.SMPLID = data.SMPLID                '�T���v��ID
'    cls.hinban = data.hinban                '�i��
'    cls.REVNUM = data.REVNUM                '���i�ԍ������ԍ�
'    cls.FACTORY = data.FACTORY              '�H��
'    cls.OPECOND = data.OPECOND              '���Ə���
'    cls.SMPLUMU = data.SMPLUMU              '�T���v���L���敪
'    cls.WFINDRS = data.WFINDRS              'WF�����w���iRs)
'    cls.WFINDOI = data.WFINDOI              'WF�����w���iOi)
'    cls.WFINDB1 = data.WFINDB1              'WF�����w���iB1)
'    cls.WFINDB2 = data.WFINDB2              'WF�����w���iB2�j
'    cls.WFINDB3 = data.WFINDB3              'WF�����w���iB3)
'    cls.WFINDL1 = data.WFINDL1              'WF�����w���iL1)
'    cls.WFINDL2 = data.WFINDL2              'WF�����w���iL2)
'    cls.WFINDL3 = data.WFINDL3              'WF�����w���iL3)
'    cls.WFINDL4 = data.WFINDL4              'WF�����w���iL4)
'    cls.WFINDDS = data.WFINDDS              'WF�����w���iDS)
'    cls.WFINDDZ = data.WFINDDZ              'WF�����w���iDZ)
'    cls.WFINDSP = data.WFINDSP              'WF�����w���iSP)
'    cls.WFINDDO1 = data.WFINDDO1            'WF�����w���iDO1)
'    cls.WFINDDO2 = data.WFINDDO2            'WF�����w���iDO2)
'    cls.WFINDDO3 = data.WFINDDO3            'WF�����w���iDO3)
'    cls.WFRESRS = data.WFRESRS              'WF�������сiRs)
'    cls.WFRESOI = data.WFRESOI              'WF�������сiOi)
'    cls.WFRESB1 = data.WFRESB1              'WF�������сiB1)
'    cls.WFRESB2 = data.WFRESB2              'WF�������сiB2�j
'    cls.WFRESB3 = data.WFRESB3              'WF�������сiB3)
'    cls.WFRESL1 = data.WFRESL1              'WF�������сiL1)
'    cls.WFRESL2 = data.WFRESL2              'WF�������сiL2)
'    cls.WFRESL3 = data.WFRESL3              'WF�������сiL3)
'    cls.WFRESL4 = data.WFRESL4              'WF�������сiL4)
'    cls.WFRESDS = data.WFRESDS              'WF�������сiDS)
'    cls.WFRESDZ = data.WFRESDZ              'WF�������сiDZ)
'    cls.WFRESSP = data.WFRESSP              'WF�������сiSP)
'    cls.WFRESDO1 = data.WFRESDO1            'WF�������сiDO1)
'    cls.WFRESDO2 = data.WFRESDO2            'WF�������сiDO2)
'    cls.WFRESDO3 = data.WFRESDO3            'WF�������сiDO3)
'    'cls.REGDATE = data.REGDATE              '�o�^���t
'    'cls.UPDDATE = data.UPDDATE              '�X�V���t
'    'cls.SENDFLAG = data.SENDFLAG            '���M�t���O
'    'cls.SENDDATE = data.SENDDATE            '���M���t
'
'    Set u2c_cmzc001f = cls
'End Function

'------ �e�[�u����:XSDCW    ---- �V�T���v���Ǘ�(SXL)

' ���[�U��`�^���N���X�ɕϊ�����(�������)
Public Function u2c_cmzc001f(data As typ_XSDCW) As c_cmzc001f
Dim cls As New c_cmzc001f       '�ϊ���N���X

    cls.SXLIDCW = data.SXLIDCW              'SXLID
    cls.SMPKBNCW = data.SMPKBNCW            '�T���v���敪
    cls.TBKBNCW = data.TBKBNCW              'T/B�敪
    cls.REPSMPLIDCW = data.REPSMPLIDCW      '��\�T���v��ID
    cls.XTALCW = data.XTALCW                '�����ԍ�
    cls.INPOSCW = data.INPOSCW              '�������ʒu
    cls.HINBCW = data.HINBCW                '�i��
    cls.REVNUMCW = data.REVNUMCW            '���i�ԍ������ԍ�
    cls.FACTORYCW = data.FACTORYCW          '�H��
    cls.OPECW = data.OPECW                  '���Ɣԍ�
    cls.KTKBNCW = data.KTKBNCW              '�m��敪
    cls.SMPLUMU = data.SMPLUMU              '�T���v���L���敪
    cls.SMCRYNUMCW = data.SMCRYNUMCW        '�T���v���u���b�NID
    cls.WFSMPLIDRSCW = data.WFSMPLIDRSCW    '�T���v��ID(Rs)
    cls.WFSMPLIDRS1CW = data.WFSMPLIDRS1CW  '����T���v��ID1(Rs)
    cls.WFSMPLIDRS2CW = data.WFSMPLIDRS2CW  '����T���v��ID2(Rs)
    cls.WFINDRSCW = data.WFINDRSCW          '���FLG(Rs)
    cls.WFRESRS1CW = data.WFRESRS1CW        '����FLG1(Rs)
    cls.WFRESRS2CW = data.WFRESRS2CW        '����FLG2(Rs)
    cls.WFSMPLIDOICW = data.WFSMPLIDOICW    '�T���v��ID(Oi)
    cls.WFINDOICW = data.WFINDOICW          '���FLG(Oi)
    cls.WFRESOICW = data.WFRESOICW          '����FLG(Oi)
    cls.WFSMPLIDB1CW = data.WFSMPLIDB1CW    '�T���v��ID(B1)
    cls.WFINDB1CW = data.WFINDB1CW          '���FLG(B1)
    cls.WFRESB1CW = data.WFRESB1CW          '����FLG(B1)
    cls.WFSMPLIDB2CW = data.WFSMPLIDB2CW    '�T���v��ID(B2)
    cls.WFINDB2CW = data.WFINDB2CW          '���FLG(B2)
    cls.WFRESB2CW = data.WFRESB2CW          '����FLG(B2)
    cls.WFSMPLIDB3CW = data.WFSMPLIDB3CW    '�T���v��ID(B3)
    cls.WFINDB3CW = data.WFINDB3CW          '���FLG(B3)
    cls.WFRESB3CW = data.WFRESB3CW          '����FLG(B3)
    cls.WFSMPLIDL1CW = data.WFSMPLIDL1CW    '�T���v��ID(L1)
    cls.WFINDL1CW = data.WFINDL1CW          '���FLG(L1)
    cls.WFRESL1CW = data.WFRESL1CW          '����FLG(L1)
    cls.WFSMPLIDL2CW = data.WFSMPLIDL2CW    '�T���v��ID(L2)
    cls.WFINDL2CW = data.WFINDL2CW          '���FLG(L2)
    cls.WFRESL2CW = data.WFRESL2CW          '����FLG(L2)
    cls.WFSMPLIDL3CW = data.WFSMPLIDL3CW    '�T���v��ID(L3)
    cls.WFINDL3CW = data.WFINDL3CW          '���FLG(L3)
    cls.WFRESL3CW = data.WFRESL3CW          '����FLG(L3)
    cls.WFSMPLIDL4CW = data.WFSMPLIDL4CW    '�T���v��ID(L4)
    cls.WFINDL4CW = data.WFINDL4CW          '���FLG(L4)
    cls.WFRESL4CW = data.WFRESL4CW          '����FLG(L4)
    cls.WFSMPLIDDSCW = data.WFSMPLIDDSCW    '�T���v��ID(DS)
    cls.WFINDDSCW = data.WFINDDSCW          '���FLG(DS)
    cls.WFRESDSCW = data.WFRESDSCW          '����FLG(DS)
    cls.WFSMPLIDDZCW = data.WFSMPLIDDZCW    '�T���v��ID(DZ)
    cls.WFINDDZCW = data.WFINDDZCW          '���FLG(DZ)
    cls.WFRESDZCW = data.WFRESDZCW          '����FLG(DZ)
    cls.WFSMPLIDSPCW = data.WFSMPLIDSPCW    '�T���v��ID(SP)
    cls.WFINDSPCW = data.WFINDSPCW          '���FLG(SP)
    cls.WFRESSPCW = data.WFRESSPCW          '����FLG(SP)
    cls.WFSMPLIDDO1CW = data.WFSMPLIDDO1CW  '�T���v��ID(DO1)
    cls.WFINDDO1CW = data.WFINDDO1CW        '���FLG(DO1)
    cls.WFRESDO1CW = data.WFRESDO1CW        '����FLG(DO1)
    cls.WFSMPLIDDO2CW = data.WFSMPLIDDO2CW  '�T���v��ID(DO2)
    cls.WFINDDO2CW = data.WFINDDO2CW        '���FLG(DO2)
    cls.WFRESDO2CW = data.WFRESDO2CW        '����FLG(DO2)
    cls.WFSMPLIDDO3CW = data.WFSMPLIDDO3CW  '�T���v��ID(DO3)
    cls.WFINDDO3CW = data.WFINDDO3CW        '���FLG(DO3)
    cls.WFRESDO3CW = data.WFRESDO3CW        '����FLG(DO3)

    'cls.REGDATE = data.REGDATE              '�o�^���t
    'cls.UPDDATE = data.UPDDATE              '�X�V���t
    'cls.SENDFLAG = data.SENDFLAG            '���M�t���O
    'cls.SENDDATE = data.SENDDATE            '���M���t

    Set u2c_cmzc001f = cls
End Function

''2003/09/02 SystemBrain �T���v���Ǘ��ύX
' �N���X�����[�U��`�^�ɕϊ�����(WF�T���v���Ǘ�)
'Public Function c2u_TBCME044(cls As c_cmzc001f) As typ_TBCME044
'Dim data As typ_TBCME044        '�ϊ��惆�[�U��`�^
'
'    data.CRYNUM = cls.CRYNUM                '�����ԍ�
'    data.INGOTPOS = cls.INGOTPOS            '�������ʒu
'    data.SMPKBN = cls.SMPKBN                '�T���v���敪
'    data.SMPLID = cls.SMPLID                '�T���v��ID
'    data.hinban = cls.hinban                '�i��
'    data.REVNUM = cls.REVNUM                '���i�ԍ������ԍ�
'    data.FACTORY = cls.FACTORY              '�H��
'    data.OPECOND = cls.OPECOND              '���Ə���
'    data.SMPLUMU = cls.SMPLUMU              '�T���v���L���敪
'    data.WFINDRS = cls.WFINDRS              'WF�����w���iRs)
'    data.WFINDOI = cls.WFINDOI              'WF�����w���iOi)
'    data.WFINDB1 = cls.WFINDB1              'WF�����w���iB1)
'    data.WFINDB2 = cls.WFINDB2              'WF�����w���iB2�j
'    data.WFINDB3 = cls.WFINDB3              'WF�����w���iB3)
'    data.WFINDL1 = cls.WFINDL1              'WF�����w���iL1)
'    data.WFINDL2 = cls.WFINDL2              'WF�����w���iL2)
'    data.WFINDL3 = cls.WFINDL3              'WF�����w���iL3)
'    data.WFINDL4 = cls.WFINDL4              'WF�����w���iL4)
'    data.WFINDDS = cls.WFINDDS              'WF�����w���iDS)
'    data.WFINDDZ = cls.WFINDDZ              'WF�����w���iDZ)
'    data.WFINDSP = cls.WFINDSP              'WF�����w���iSP)
'    data.WFINDDO1 = cls.WFINDDO1            'WF�����w���iDO1)
'    data.WFINDDO2 = cls.WFINDDO2            'WF�����w���iDO2)
'    data.WFINDDO3 = cls.WFINDDO3            'WF�����w���iDO3)
'    data.WFRESRS = cls.WFRESRS              'WF�������сiRs)
'    data.WFRESOI = cls.WFRESOI              'WF�������сiOi)
'    data.WFRESB1 = cls.WFRESB1              'WF�������сiB1)
'    data.WFRESB2 = cls.WFRESB2              'WF�������сiB2�j
'    data.WFRESB3 = cls.WFRESB3              'WF�������сiB3)
'    data.WFRESL1 = cls.WFRESL1              'WF�������сiL1)
'    data.WFRESL2 = cls.WFRESL2              'WF�������сiL2)
'    data.WFRESL3 = cls.WFRESL3              'WF�������сiL3)
'    data.WFRESL4 = cls.WFRESL4              'WF�������сiL4)
'    data.WFRESDS = cls.WFRESDS              'WF�������сiDS)
'    data.WFRESDZ = cls.WFRESDZ              'WF�������сiDZ)
'    data.WFRESSP = cls.WFRESSP              'WF�������сiSP)
'    data.WFRESDO1 = cls.WFRESDO1            'WF�������сiDO1)
'    data.WFRESDO2 = cls.WFRESDO2            'WF�������сiDO2)
'    data.WFRESDO3 = cls.WFRESDO3            'WF�������сiDO3)
'    'data.REGDATE = cls.REGDATE              '�o�^���t
'    'data.UPDDATE = cls.UPDDATE              '�X�V���t
'    'data.SENDFLAG = cls.SENDFLAG            '���M�t���O
'    'data.SENDDATE = cls.SENDDATE            '���M���t
'
'    c2u_TBCME044 = data
'End Function
Public Function c2u_XSDCW(cls As c_cmzc001f) As typ_XSDCW
Dim data As typ_XSDCW        '�ϊ��惆�[�U��`�^

    data.SXLIDCW = cls.SXLIDCW              'SXLID
    data.SMPKBNCW = cls.SMPKBNCW            '�T���v���敪
    data.TBKBNCW = cls.TBKBNCW              'T/B�敪
    data.REPSMPLIDCW = cls.REPSMPLIDCW      '��\�T���v��ID
    data.XTALCW = cls.XTALCW                '�����ԍ�
    data.INPOSCW = cls.INPOSCW              '�������ʒu
    data.HINBCW = cls.HINBCW                '�i��
    data.REVNUMCW = cls.REVNUMCW            '���i�ԍ������ԍ�
    data.FACTORYCW = cls.FACTORYCW          '�H��
    data.OPECW = cls.OPECW                  '���Ɣԍ�
    data.KTKBNCW = cls.KTKBNCW              '�m��敪
    data.SMPLUMU = cls.SMPLUMU              '�T���v���L���敪
    data.SMCRYNUMCW = cls.SMCRYNUMCW        '�T���v���u���b�NID
    data.WFSMPLIDRSCW = cls.WFSMPLIDRSCW    '�T���v��ID(Rs)
    data.WFSMPLIDRS1CW = cls.WFSMPLIDRS1CW  '����T���v��ID1(Rs)
    data.WFSMPLIDRS2CW = cls.WFSMPLIDRS2CW  '����T���v��ID2(Rs)
    data.WFINDRSCW = cls.WFINDRSCW          '���FLG(Rs)
    data.WFRESRS1CW = cls.WFRESRS1CW        '����FLG1(Rs)
    data.WFRESRS2CW = cls.WFRESRS2CW        '����FLG2(Rs)
    data.WFSMPLIDOICW = cls.WFSMPLIDOICW    '�T���v��ID(Oi)
    data.WFINDOICW = cls.WFINDOICW          '���FLG(Oi)
    data.WFRESOICW = cls.WFRESOICW          '����FLG(Oi)
    data.WFSMPLIDB1CW = cls.WFSMPLIDB1CW    '�T���v��ID(B1)
    data.WFINDB1CW = cls.WFINDB1CW          '���FLG(B1)
    data.WFRESB1CW = cls.WFRESB1CW          '����FLG(B1)
    data.WFSMPLIDB2CW = cls.WFSMPLIDB2CW    '�T���v��ID(B2)
    data.WFINDB2CW = cls.WFINDB2CW          '���FLG(B2)
    data.WFRESB2CW = cls.WFRESB2CW          '����FLG(B2)
    data.WFSMPLIDB3CW = cls.WFSMPLIDB3CW    '�T���v��ID(B3)
    data.WFINDB3CW = cls.WFINDB3CW          '���FLG(B3)
    data.WFRESB3CW = cls.WFRESB3CW          '����FLG(B3)
    data.WFSMPLIDL1CW = cls.WFSMPLIDL1CW    '�T���v��ID(L1)
    data.WFINDL1CW = cls.WFINDL1CW          '���FLG(L1)
    data.WFRESL1CW = cls.WFRESL1CW          '����FLG(L1)
    data.WFSMPLIDL2CW = cls.WFSMPLIDL2CW    '�T���v��ID(L2)
    data.WFINDL2CW = cls.WFINDL2CW          '���FLG(L2)
    data.WFRESL2CW = cls.WFRESL2CW          '����FLG(L2)
    data.WFSMPLIDL3CW = cls.WFSMPLIDL3CW    '�T���v��ID(L3)
    data.WFINDL3CW = cls.WFINDL3CW          '���FLG(L3)
    data.WFRESL3CW = cls.WFRESL3CW          '����FLG(L3)
    data.WFSMPLIDL4CW = cls.WFSMPLIDL4CW    '�T���v��ID(L4)
    data.WFINDL4CW = cls.WFINDL4CW          '���FLG(L4)
    data.WFRESL4CW = cls.WFRESL4CW          '����FLG(L4)
    data.WFSMPLIDDSCW = cls.WFSMPLIDDSCW    '�T���v��ID(DS)
    data.WFINDDSCW = cls.WFINDDSCW          '���FLG(DS)
    data.WFRESDSCW = cls.WFRESDSCW          '����FLG(DS)
    data.WFSMPLIDDZCW = cls.WFSMPLIDDZCW    '�T���v��ID(DZ)
    data.WFINDDZCW = cls.WFINDDZCW          '���FLG(DZ)
    data.WFRESDZCW = cls.WFRESDZCW          '����FLG(DZ)
    data.WFSMPLIDSPCW = cls.WFSMPLIDSPCW    '�T���v��ID(SP)
    data.WFINDSPCW = cls.WFINDSPCW          '���FLG(SP)
    data.WFRESSPCW = cls.WFRESSPCW          '����FLG(SP)
    data.WFSMPLIDDO1CW = cls.WFSMPLIDDO1CW  '�T���v��ID(DO1)
    data.WFINDDO1CW = cls.WFINDDO1CW        '���FLG(DO1)
    data.WFRESDO1CW = cls.WFRESDO1CW        '����FLG(DO1)
    data.WFSMPLIDDO2CW = cls.WFSMPLIDDO2CW  '�T���v��ID(DO2)
    data.WFINDDO2CW = cls.WFINDDO2CW        '���FLG(DO2)
    data.WFRESDO2CW = cls.WFRESDO2CW        '����FLG(DO2)
    data.WFSMPLIDDO3CW = cls.WFSMPLIDDO3CW  '�T���v��ID(DO3)
    data.WFINDDO3CW = cls.WFINDDO3CW        '���FLG(DO3)
    data.WFRESDO3CW = cls.WFRESDO3CW        '����FLG(DO3)

    'data.REGDATE = cls.REGDATE              '�o�^���t
    'data.UPDDATE = cls.UPDDATE              '�X�V���t
    'data.SENDFLAG = cls.SENDFLAG            '���M�t���O
    'data.SENDDATE = cls.SENDDATE            '���M���t

    c2u_XSDCW = data
End Function


'------ �e�[�u����:TBCME045    ---- �ؒf�w��

' ���[�U��`�^���N���X�ɕϊ�����(�������)
Public Function u2c_cmzc001c(data As typ_TBCME045) As c_cmzc001c
Dim cls As New c_cmzc001c       '�ϊ���N���X

    cls.CRYNUM = data.CRYNUM                '�����ԍ�
    cls.IngotPos = data.IngotPos            '�������J�n�ʒu
    cls.TRANCNT = data.TRANCNT              '������
    cls.LENGTH = data.LENGTH                '����
    cls.PROCCODE = data.PROCCODE            '�H���R�[�h
    cls.StaffID = data.StaffID              '�Ј�ID
    cls.HINBAN = data.HINBAN                '��i��
    cls.REVNUM = data.REVNUM                '��i�Ԑ��i�ԍ������ԍ�
    cls.factory = data.factory              '��i�ԍH��
    cls.opecond = data.opecond              '��i�ԑ��Ə���
    cls.BDCAUS = data.BDCAUS                '�敪�R�[�h
    cls.STATCLS = data.STATCLS              '��ԋ敪
    cls.BLOCKID = data.BLOCKID              '�u���b�NID
    'cls.REGDATE = data.REGDATE              '�o�^���t
    'cls.UPDDATE = data.UPDDATE              '�X�V���t
    'cls.SENDFLAG = data.SENDFLAG            '���M�t���O
    'cls.SENDDATE = data.SENDDATE            '���M���t

    Set u2c_cmzc001c = cls
End Function


' �N���X�����[�U��`�^�ɕϊ�����(�ؒf�w��)
Public Function c2u_TBCME045(cls As c_cmzc001c) As typ_TBCME045
Dim data As typ_TBCME045        '�ϊ��惆�[�U��`�^

    data.CRYNUM = cls.CRYNUM                '�����ԍ�
    data.IngotPos = cls.IngotPos            '�������J�n�ʒu
    data.TRANCNT = cls.TRANCNT              '������
    data.LENGTH = cls.LENGTH                '����
    data.PROCCODE = cls.PROCCODE            '�H���R�[�h
    data.StaffID = cls.StaffID              '�Ј�ID
    data.HINBAN = cls.HINBAN                '��i��
    data.REVNUM = cls.REVNUM                '��i�Ԑ��i�ԍ������ԍ�
    data.factory = cls.factory              '��i�ԍH��
    data.opecond = cls.opecond              '��i�ԑ��Ə���
    data.BDCAUS = cls.BDCAUS                '�敪�R�[�h
    data.STATCLS = cls.STATCLS              '��ԋ敪
    data.BLOCKID = cls.BLOCKID              '�u���b�NID
    'data.REGDATE = cls.REGDATE              '�o�^���t
    'data.UPDDATE = cls.UPDDATE              '�X�V���t
    'data.SENDFLAG = cls.SENDFLAG            '���M�t���O
    'data.SENDDATE = cls.SENDDATE            '���M���t

    c2u_TBCME045 = data
End Function
