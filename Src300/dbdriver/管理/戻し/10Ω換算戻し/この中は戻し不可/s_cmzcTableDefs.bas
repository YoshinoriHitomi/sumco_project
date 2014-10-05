Attribute VB_Name = "s_cmzcTableDefs"
Option Explicit
'7/30

Public STAFFIDBUFF  As String
Public spread_Col As Long
Public spread_Row As Long
Public MaxLine As Long

''�G���[���b�Z�[�W
Public Const ESTAF = "ESTAF" ''�S���҃R�[�h�������ł��
Public Const EIE00 = "EIE00" ''�S�Ẵf�[�^���͂��������Ă��܂���
Public Const EIE01 = "EBLK1" ''�u���b�NID�̌������Ԉ���Ă��܂��
Public Const EIM00 = "EIM00" ''�w���P����������с@�₢���킹���B
Public Const EGET = "EGET" ''DB����̓Ǎ��Ɏ��s���܂����B
Public Const EAPLY = "EAPLY" ''DB�ւ̏����Ɏ��s���܂����B
Public Const EMAT1 = "EMAT1" '' �����ԍ��̌������Ԉ���Ă��܂��
Public Const EMAT2 = "EMAT2" '' �w�肵�������ԍ��͖��o�^�ł��B
Public Const KIE00 = "EBLK0" ''���͂��ꂽ�u���b�NID�ͤ���݂��܂���
'Public Const KDE01 = "KDE01" ''�w���P�����́A�C���[�W�\���ł��܂���B  ??????
Public Const KDE01 = "EKDE1" ''�w���P�����́A�C���[�W�\���ł��܂���B
Public Const PWAIT = "PWAIT" ''���X���҂�������
Public Const KC001 = "EKC01" ''�N���X�^���J�^���O���������s���܂����I
Public Const TJE01 = "PJE01" ''��������NG�ł��B
Public Const ESXL0 = "ESXL0" ''���͂��ꂽSXLID�́A���݂��܂���B"
Public Const ESXL1 = "ESXL1" ''SXLID�̌������Ԉ���Ă��܂��B"
Public Const PWCC0 = "PWCC0" ''�N���X�^���J�^���O �������B
Public Const E0001 = "E0001" ''�u���b�N���I������Ă��܂���B
Public Const EGB01 = "EGB01" ''����d�ʂ��������܂��B
Public Const EGB02 = "EGB02" ''����d�ʂ�����������܂���B
Public Const EINPM = "EINPM" ''���͒l���s���ł��
Public Const PIN16 = "PLBL1" ''���x�����Ĕ��s���܂��B��낵���ł����H
Public Const POK06 = "PLBL2" ''���x�����Ĕ��s���܂����B
Public Const PLBL3 = "PLBL3" ''���x��������
Public Const ELB00 = "ELBL0" ''���x������G���[�
Public Const EHIN1 = "EHIN1" ''�i�Ԃ̌������Ԉ���Ă��܂��B"
Public Const EHIN0 = "EHIN0" ''�w��̕i�Ԃ͖��o�^�ł��B"
Public Const EBLK5 = "EBLK5" ''�u���b�NID���A����������܂���B

Public Const E0002 = "E0002" ''�I�������u���b�N�͘A�����Ă��܂���
Public Const WGD01 = "WGD01" ''GD�G���[�ƂȂ�i�Ԃ�����܂�
Public Const EKDE2 = "EKDE2" ''�w���P�����ͤ���ύX�ł��܂���
Public Const EGET2 = "EGET2" ''DB����̓Ǎ��Ɏ��s���܂���(%s)�B


Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public EndFlag As Boolean
'                                     2001/08/24
'================================================
' ���[�U��`�^�̐錾
' ��`���e: 060200_�S�e�[�u��
'================================================


' SXL������
Public Type typ_TBCMX001
    SXLID As String * 13            ' SXLID
    FROMTOKBN As String * 1         ' FROMTO�敪
    SAMPLE_FROM As String * 16      ' �T���v��ID (From)
    SAMPLE_TO As String * 16        ' �T���v��ID (To)
    BLOCKID As String * 12          ' �u���b�NID
    CRYNUM As String * 12           ' �����ԍ�
    SXLDECDATE As Date              ' SXL-ID�m����t
    PLUPDATE As Date                ' ������t
    INGOTPOS As Integer             ' �������J�n�ʒu
    hinban As String * 12           ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    PRODCOND As String * 10         ' �������
    PGID As String * 8              ' �o�f�|�h�c
    UPLENGTH As Integer             ' ���グ����
    SXLPOS As Integer               ' SXL�ʒu
    SXLLENGTH As Integer            ' SXL-ID�m�蒷��
    SXLWAFERCNT As Integer          ' SXL-ID�m�莞�̖���
    FREELENG As Integer             ' �t���[��

    DIAMETER As Integer             ' ���a
    CHARGE As Long                  ' �`���[�W��
    SEED As String * 4              ' �V�[�h
    SAMPID As String * 16           ' �T���v��ID
    SXL_RS_SMPPOS As Integer        ' SXLRS����ّ���ʒu�iSXL������j
    SXLRS_MEAS1 As Double           ' SXLRS_����l�P
    SXLRS_MEAS2 As Double           ' SXLRS_����l�Q
    SXLRS_MEAS3 As Double           ' SXLRS_����l�R
    SXLRS_MEAS4 As Double           ' SXLRS_����l�S
    SXLRS_MEAS5 As Double           ' SXLRS_����l�T
    SXLRS_EFEHS As Double           ' SXLRS_�����ΐ�
    SXLRS_RRG As Double             ' SXLRS_�q�q�f
    SXL_OI_SMPPOS As Integer        ' SXLOI����ّ���ʒu�iSXL������j
    SXLOI_OIMEAS1 As Double         ' SXLOI_�n������l�P
    SXLOI_OIMEAS2 As Double         ' SXLOI_�n������l�Q
    SXLOI_OIMEAS3 As Double         ' SXLOI_�n������l�R
    SXLOI_OIMEAS4 As Double         ' SXLOI_�n������l�S
    SXLOI_OIMEAS5 As Double         ' SXLOI_�n������l�T
    SXLOI_ORGRES As Double          ' SXLOI_�n�q�f����
    SXLOI_INSPECTWAY As String * 2  ' SXLOI_�������@
    SXL_CS_SMPPOS As Integer        ' SXLCS����ّ���ʒu�iSXL������j
    SXLCS_CSMEAS As Double          ' SXLCS_Cs�����l
    SXLCS_70PPRE As Double          ' SXLCS_�V�O������l
    SXLOSF_SMPPOS As Integer        ' OSF����ّ���ʒu�iSXL�ʒu���j
    SXLOSF1_KKSP As String * 3      ' OSF1�������ב���ʒu
    SXLOSF1_NETU As String * 2      ' OSF1�M�����@
    SXLOSF1_KKSET As String * 3     ' OSF1�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF1_CALCMAX As Double       ' OSF1SXL�v�Z���� Max_1
    SXLOSF1_CALCAVE As Double       ' OSF1SXL�v�Z���� Ave_1
    SXLOSF2_KKSP As String * 3      ' OSF�Q�������ב���ʒu
    SXLOSF2_NETU As String * 2      ' OSF�Q�M�����@
    SXLOSF2_KKSET As String * 3     ' OSF�Q�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF2_CALCMAX As Double       ' OSF�QSXL�v�Z���� Max_2
    SXLOSF2_CALCAVE As Double       ' OSF�QSXL�v�Z���� Ave_2
    SXLOSF3_KKSP As String * 3      ' OSF�R�������ב���ʒu
    SXLOSF3_NETU As String * 2      ' OSF�R�M�����@
    SXLOSF3_KKSET As String * 3     ' OSF�R�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF3_CALCMAX As Double       ' OSF�RSXL�v�Z���� Max_3
    SXLOSF3_CALCAVE As Double       ' OSF�RSXL�v�Z���� Ave_3
    SXLOSF4_KKSP As String * 3      ' OSF�S�������ב���ʒu
    SXLOSF4_NETU As String * 2      ' OSF�S�M�����@
    SXLOSF4_KKSET As String * 3     ' OSF�S�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF4_CALCMAX As Double       ' OSF�SSXL�v�Z���� Max_4
    SXLOSF4_CALCAVE As Double       ' OSF�SSXL�v�Z���� Ave_4
    SXLBMD_SMPPOS As Integer        ' BMD����ّ���ʒu�iSXL�ʒu���j
    SXLBMD1_KKSP As String * 3      ' BMD1�������ב���ʒu
    SXLBMD1_NETU As String * 2      ' BMD1�M�����@
    SXLBMD1_KKSET As String * 3     ' BMD1�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLBMD1_CALCMAX As Double       ' BMD1SXL�v�Z���� Max
    SXLBMD1_CALCAVE As Double       ' BMD1SXL�v�Z���� Ave
    SXLBMD2_KKSP As String * 3      ' BMD�Q�������ב���ʒu
    SXLBMD2_NETU As String * 2      ' BMD�Q�M�����@
    SXLBMD2_KKSET As String * 3     ' BMD�Q�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLBMD2_CALCMAX As Double       ' BMD�QSXL�v�Z���� Max
    SXLBMD2_CALCAVE As Double       ' BMD�QSXL�v�Z���� Ave
    SXLBMD3_KKSP As String * 3      ' BMD�R�������ב���ʒu
    SXLBMD3_NETU As String * 2      ' BMD�R�M�����@
    SXLBMD3_KKSET As String * 3     ' BMD�R�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLBMD3_CALCMAX As Double       ' BMD�RSXL�v�Z���� Max
    SXLBMD3_CALCAVE As Double       ' BMD�RSXL�v�Z���� Ave
    SXLGD_SMPPOS As Integer         ' GD����ّ���ʒu�iSXL�ʒu���j
    SXLGD_MSRSDEN As Integer        ' SXLGD_���茋�� Den
    SXLGD_MSRSLDL As Integer        ' SXLGD_���茋�� L/DL
    SXLGD_MSRSDVD2 As Integer       ' SXLGD_���茋�� DVD2
    SXLLT_SMPPOS As Integer         ' LT����ّ���ʒu�iSXL�ʒu���j
    SXLLT_MEASPEAK As Integer       ' SXLLT_����l �s�[�N�l
    SXLLT_CALCMEAS As Integer       ' SXLLT_�v�Z����
    WFOI_SMPPOS As Integer          ' WFOI�����-ID����ʒu�iSXL�ʒu���j
    WFOI_NETSU As String * 2        ' WFOI_�M��������
    WFOI_ET As String * 3           ' WFOI_�G�b�`���O����
    WFOI_MES As String * 3          ' WFOI_�v�����@
    WFOI_MESDATA1 As Double         ' WFOI_����f�[�^���̂P
    WFOI_MESDATA2 As Double         ' WFOI_����f�[�^���̂Q
    WFOI_MESDATA3 As Double         ' WFOI_����f�[�^���̂R
    WFOI_MESDATA4 As Double         ' WFOI_����f�[�^���̂S
    WFOI_MESDATA5 As Double         ' WFOI_����f�[�^���̂T
    WFOI_MESDATA6 As Double         ' WFOI_����f�[�^���̂U
    WFOI_MESDATA7 As Double         ' WFOI_����f�[�^���̂V
    WFOI_MESDATA8 As Double         ' WFOI_����f�[�^���̂W
    WFOI_MESDATA9 As Double         ' WFOI_����f�[�^���̂X
    WFOI_MESDATA10 As Double        ' WFOI_����f�[�^���̂P�O
    WFOI_ORG As Double              ' WFOI_ORG�v�Z����
    WFRS_SMPPOS As Integer          ' WFRS�����-ID����ʒu�iSXL�ʒu���j
    WFRS_NETSU As String * 2        ' WFRS_�M��������
    WFRS_ET As String * 3           ' WFRS_�G�b�`���O����
    WFRS_MES As String * 3          ' WFRS_�v�����@
    WFRS_MESDATA1 As Double         ' WFRS_����f�[�^���̂P
    WFRS_MESDATA2 As Double         ' WFRS_����f�[�^���̂Q
    WFRS_MESDATA3 As Double         ' WFRS_����f�[�^���̂R
    WFRS_MESDATA4 As Double         ' WFRS_����f�[�^���̂S
    WFRS_MESDATA5 As Double         ' WFRS_����f�[�^���̂T
    WFRS_RRG As Double              ' WFRS_RRG�v�Z����
    WFDOI_SMPPOS As Integer         ' WFDOI�����-ID����ʒu�iSXL�ʒu���j�@number(4)
    WFDOI_NETU_1 As String * 2      ' WFDOI_�M��������_1
    WFDOI_MES_1 As String * 3       ' WFDOI_�v�����@_1
    WFDOI_MESDATA1_1 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi�j�P_1
    WFDOI_MESDATA2_1 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�Q_1
    WFDOI_MESDATA3_1 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�R_1
    WFDOI_NETU_2 As String * 2      ' WFDOI_�M��������_�Q
    WFDOI_MES_2 As String * 3       ' WFDOI_�v�����@_�Q
    WFDOI_MESDATA1_2 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi�j�P_�Q
    WFDOI_MESDATA2_2 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�Q_�Q
    WFDOI_MESDATA3_2 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�R_�Q
    WFDOI_NETU_3 As String * 2      ' WFDOI_�M��������_�R
    WFDOI_MES_3 As String * 3       ' WFDOI_�v�����@_�R
    WFDOI_MESDATA1_3 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi�j�P_�R
    WFDOI_MESDATA2_3 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�Q_�R
    WFDOI_MESDATA3_3 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�R_�R
    WFOSF1_SMPPOS As Integer        ' WFOSF1�����-ID����ʒu�iSXL�ʒu���j
    WFOSF1_NETSU As String * 2      ' WFOSF1_�M��������
    WFOSF1_ET As String * 3         ' WFOSF1_�G�b�`���O����
    WFOSF1_MES As String * 3        ' WFOSF1_�v�����@
    WFOSF1_MAX As Double            ' WFOSF1_���莞��MAX�l_1
    WFOSF1_AVE As Double            ' WFOSF1_���莞��AVE�l_1
    WFOSF2_SMPPOS As Integer        ' WFOSF�Q�����-ID����ʒu�iSXL�ʒu���j�@number(4)
    WFOSF2_NETSU As String * 2      ' WFOSF2_�M��������_�Q
    WFOSF2_ET As String * 3         ' WFOSF2_�G�b�`���O����_�Q
    WFOSF2_MES As String * 3        ' WFOSF2_�v�����@_�Q
    WFOSF2_MAX As Double            ' WFOSF2_���莞��MAX�l_�Q
    WFOSF2_AVE As Double            ' WFOSF2_���莞��AVE�l_�Q
    WFOSF3_SMPPOS As Integer        ' WFOSF�R�����-ID����ʒu�iSXL�ʒu���j
    WFOSF3_NETSU As String * 2      ' WFOSF3_�M��������_�R
    WFOSF3_ET As String * 3         ' WFOSF3_�G�b�`���O����_�R
    WFOSF3_MES As String * 3        ' WFOSF3_�v�����@_�R
    WFOSF3_MAX As Double            ' WFOSF3_���莞��MAX�l_�R
    WFOSF3_AVE As Double            ' WFOSF3_���莞��AVE�l_�R
    WFOSF4_SMPPOS As Integer        ' WFOSF�S�����-ID����ʒu�iSXL�ʒu���j
    WFOSF4_NETSU As String * 2      ' WFOSF4_�M��������_�S
    WFOSF4_ET As String * 3         ' WFOSF4_�G�b�`���O����_�S
    WFOSF4_MES As String * 3        ' WFOSF4_�v�����@_�S
    WFOSF4_MAX As Double            ' WFOSF4_���莞��MAX�l_�S
    WFOSF4_AVE As Double            ' WFOSF4_���莞��AVE�l_�S
    WFBMD1_SMPPOS As Integer        ' WFBMD1�����-ID����ʒu�iSXL�ʒu���j
    WFBMD1_NETSU As String * 2      ' WFBMD1_�M��������_1
    WFBMD1_ET As String * 3         ' WFBMD1_�G�b�`���O����_1
    WFBMD1_MES As String * 3        ' WFBMD1_�v�����@_1
    WFBMD1_MAX As Double            ' WFBMD1_���莞��MAX�l_1
    WFBMD1_AVE As Double            ' WFBMD1_���莞��AVE�l_1
    WFBMD2_SMPPOS As Integer        ' WFBMD�Q�����-ID����ʒu�iSXL�ʒu���j
    WFBMD2_NETSU As String * 2      ' WFBMD2_�M��������_�Q
    WFBMD2_ET As String * 3         ' WFBMD2_�G�b�`���O����_�Q
    WFBMD2_MES As String * 3        ' WFBMD2_�v�����@_�Q
    WFBMD2_MAX As Double            ' WFBMD2_���莞��MAX�l_�Q
    WFBMD2_AVE As Double            ' WFBMD2_���莞��AVE�l_�Q
    WFBMD3_SMPPOS As Integer        ' WFBMD�R�����-ID����ʒu�iSXL�ʒu���j
    WFBMD3_NETSU As String * 2      ' WFBMD3_�M��������_�R
    WFBMD3_ET As String * 3         ' WFBMD3_�G�b�`���O����_�R
    WFBMD3_MES As String * 3        ' WFBMD3_�v�����@_�R
    WFBMD3_MAX As Double            ' WFBMD3_���莞��MAX�l_�R
    WFBMD3_AVE As Double            ' WFBMD3_���莞��AVE�l_�R
    WFDSOD_SMPPOS As Integer        ' WFDSOD�����-ID����ʒu�iSXL�ʒu���j
    WFDSOD_NETSU As String * 2      ' WFDSOD_�M��������
    WFDSOD_ET As String * 3         ' WFDSOD_�G�b�`���O����
    WFDSOD_MES As String * 3        ' WFDSOD_�v�����@
    WFDSOD_TOTAL As Integer         ' WFDSOD_���莞��TOTAL�l
    WFSPV_SMPPOS As Integer         ' WFSPV�����-ID����ʒu�iSXL�ʒu���j
    WFSPV_NETSU As String * 2       ' WFSVP_�M��������
    WFSPV_ET As String * 3          ' WFSPV_�G�b�`���O����
    WFSPV_MES As String * 3         ' WFSPV_�v�����@
    WFSPV_KST_MAX As Double         ' WFSPV_�g�U�����莞�̂l�`�w�l
    WFSPV_KST_AVE As Double         ' WFSPV_�g�U�����莞��AVE�l
    WFSPV_KST_MIN As Double         ' WFSPV_�g�U�����莞��MIN�l
    WFSPV_FE_MAX As Double          ' WFSPV_Fe�Z�x���莞��MAX�l
    WFSPV_FE_AVE As Double          ' WFSPV_Fe�Z�x���莞��AVE�l
    WFSPV_FE_MIN As Double          ' WFSPV_Fe�Z�x���莞��MIN�l
    WFDZ_SMPPOS As Integer          ' WFDZ�����-ID����ʒu�iSXL�ʒu���j
    WFDZ_NETSU As String * 2        ' WFDZ_�M��������
    WFDZ_ET As String * 3           ' WFDZ_�G�b�`���O����
    WFDZ_MES As String * 3          ' WFDZ_�v�����@
    WFDZ_MAX As Double              ' WFDZ_���莞��MAX�l_
    WFDZ_AVE As Double              ' WFDZ_���莞��AVE�l
    REGDATE As Date                 ' �o�^���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    
    SXLOSF1_PTNJUDGRES  As String * 1   ' OSF1�p�^�[�����茋��  '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
End Type


' SXL����_�f�|�^
Public Type typ_TBCMX002
    SXLID As String * 13            ' SXLID
    FROMTOKBN As String * 1         ' FROMTO�敪
    SAMPLE_FROM As String * 16      ' �T���v��ID (From)
    SAMPLE_TO As String * 16        ' �T���v��ID (To)
    BLOCKID As String * 12          ' �u���b�NID
    CRYNUM As String * 12           ' �����ԍ�
    SXLDECDATE As Date              ' SXL-ID�m����t
    PLUPDATE As Date                ' ������t
    INGOTPOS As Integer             ' �������J�n�ʒu
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    UPLENGTH As Integer             ' ���グ����
    SXLPOS As Integer               ' SXL�ʒu
    SXLLENGTH As Integer            ' SXL-ID�m�蒷��
    SXLWAFERCNT As Integer          ' SXL-ID�m�莞�̖���
    FREELENG As Integer             ' �t���[��
    SAMPID_1 As String * 16         ' �T���v��ID 1
    SXLOSF1_SMPPOS As Integer       ' SXLOSF����ّ���ʒu�iSXL�ʒu���j
    SXLOSF1_KKSP As String * 3      ' SXLOSF1�������ב���ʒu
    SXLOSF1_NETU As String * 2      ' SXLOSF1�M�����@
    SXLOSF1_KKSET As String * 3     ' SXLOSF1�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF1_MEAS1 As Integer        ' SXLOSF1����_�P
    SXLOSF1_MEAS2 As Integer        ' SXLOSF1����_2
    SXLOSF1_MEAS3 As Integer        ' SXLOSF1����_3
    SXLOSF1_MEAS4 As Integer        ' SXLOSF1����_4
    SXLOSF1_MEAS5 As Integer        ' SXLOSF1����_5
    SXLOSF1_MEAS6 As Integer        ' SXLOSF1����_6
    SXLOSF1_MEAS7 As Integer        ' SXLOSF1����_7
    SXLOSF1_MEAS8 As Integer        ' SXLOSF1����_8
    SXLOSF1_MEAS9 As Integer        ' SXLOSF1����_9
    SXLOSF1_MEAS10 As Integer       ' SXLOSF1����_10
    SXLOSF1_MEAS11 As Integer       ' SXLOSF1����_11
    SXLOSF1_MEAS12 As Integer       ' SXLOSF1����_12
    SXLOSF1_MEAS13 As Integer       ' SXLOSF1����_13
    SXLOSF1_MEAS14 As Integer       ' SXLOSF1����_14
    SXLOSF1_MEAS15 As Integer       ' SXLOSF1����_15
    SXLOSF1_MEAS16 As Integer       ' SXLOSF1����_16
    SXLOSF1_MEAS17 As Integer       ' SXLOSF1����_17
    SXLOSF1_MEAS18 As Integer       ' SXLOSF1����_18
    SXLOSF1_MEAS19 As Integer       ' SXLOSF1����_19
    SXLOSF1_MEAS20 As Integer       ' SXLOSF1����_20
    SXLOSF2_KKSP As String * 3      ' SXLOSF�Q�������ב���ʒu
    SXLOSF2_NETU As String * 2      ' SXLOSF�Q�M�����@
    SXLOSF2_KKSET As String * 3     ' SXLOSF�Q�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF2_MEAS1 As Integer        ' SXLOSF2����_�P
    SXLOSF2_MEAS2 As Integer        ' SXLOSF2����_2
    SXLOSF2_MEAS3 As Integer        ' SXLOSF2����_3
    SXLOSF2_MEAS4 As Integer        ' SXLOSF2����_4
    SXLOSF2_MEAS5 As Integer        ' SXLOSF2����_5
    SXLOSF2_MEAS6 As Integer        ' SXLOSF2����_6
    SXLOSF2_MEAS7 As Integer        ' SXLOSF2����_7
    SXLOSF2_MEAS8 As Integer        ' SXLOSF2����_8
    SXLOSF2_MEAS9 As Integer        ' SXLOSF2����_9
    SXLOSF2_MEAS10 As Integer       ' SXLOSF2����_10
    SXLOSF2_MEAS11 As Integer       ' SXLOSF2����_11
    SXLOSF2_MEAS12 As Integer       ' SXLOSF2����_12
    SXLOSF2_MEAS13 As Integer       ' SXLOSF2����_13
    SXLOSF2_MEAS14 As Integer       ' SXLOSF2����_14
    SXLOSF2_MEAS15 As Integer       ' SXLOSF2����_15
    SXLOSF2_MEAS16 As Integer       ' SXLOSF2����_16
    SXLOSF2_MEAS17 As Integer       ' SXLOSF2����_17
    SXLOSF2_MEAS18 As Integer       ' SXLOSF2����_18
    SXLOSF2_MEAS19 As Integer       ' SXLOSF2����_19
    SXLOSF2_MEAS20 As Integer       ' SXLOSF2����_20
    SXLOSF3_KKSP As String * 3      ' SXLOSF�R�������ב���ʒu
    SXLOSF3_NETU As String * 2      ' SXLOSF�R�M�����@
    SXLOSF3_KKSET As String * 3     ' SXLOSF�R�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF3_MEAS1 As Integer        ' SXLOSF3����_�P
    SXLOSF3_MEAS2 As Integer        ' SXLOSF3����_2
    SXLOSF3_MEAS3 As Integer        ' SXLOSF3����_3
    SXLOSF3_MEAS4 As Integer        ' SXLOSF3����_4
    SXLOSF3_MEAS5 As Integer        ' SXLOSF3����_5
    SXLOSF3_MEAS6 As Integer        ' SXLOSF3����_6
    SXLOSF3_MEAS7 As Integer        ' SXLOSF3����_7
    SXLOSF3_MEAS8 As Integer        ' SXLOSF3����_8
    SXLOSF3_MEAS9 As Integer        ' SXLOSF3����_9
    SXLOSF3_MEAS10 As Integer       ' SXLOSF3����_10
    SXLOSF3_MEAS11 As Integer       ' SXLOSF3����_11
    SXLOSF3_MEAS12 As Integer       ' SXLOSF3����_12
    SXLOSF3_MEAS13 As Integer       ' SXLOSF3����_13
    SXLOSF3_MEAS14 As Integer       ' SXLOSF3����_14
    SXLOSF3_MEAS15 As Integer       ' SXLOSF3����_15
    SXLOSF3_MEAS16 As Integer       ' SXLOSF3����_16
    SXLOSF3_MEAS17 As Integer       ' SXLOSF3����_17
    SXLOSF3_MEAS18 As Integer       ' SXLOSF3����_18
    SXLOSF3_MEAS19 As Integer       ' SXLOSF3����_19
    SXLOSF3_MEAS20 As Integer       ' SXLOSF3����_20
    SXLOSF4_KKSP As String * 3      ' SXLOSF�S�������ב���ʒu
    SXLOSF4_NETU As String * 2      ' SXLOSF�S�M�����@
    SXLOSF4_KKSET As String * 3     ' SXLOSF�S�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF4_MEAS1 As Integer        ' SXLOSF4����_�P
    SXLOSF4_MEAS2 As Integer        ' SXLOSF4����_2
    SXLOSF4_MEAS3 As Integer        ' SXLOSF4����_3
    SXLOSF4_MEAS4 As Integer        ' SXLOSF4����_4
    SXLOSF4_MEAS5 As Integer        ' SXLOSF4����_5
    SXLOSF4_MEAS6 As Integer        ' SXLOSF4����_6
    SXLOSF4_MEAS7 As Integer        ' SXLOSF4����_7
    SXLOSF4_MEAS8 As Integer        ' SXLOSF4����_8
    SXLOSF4_MEAS9 As Integer        ' SXLOSF4����_9
    SXLOSF4_MEAS10 As Integer       ' SXLOSF4����_10
    SXLOSF4_MEAS11 As Integer       ' SXLOSF4����_11
    SXLOSF4_MEAS12 As Integer       ' SXLOSF4����_12
    SXLOSF4_MEAS13 As Integer       ' SXLOSF4����_13
    SXLOSF4_MEAS14 As Integer       ' SXLOSF4����_14
    SXLOSF4_MEAS15 As Integer       ' SXLOSF4����_15
    SXLOSF4_MEAS16 As Integer       ' SXLOSF4����_16
    SXLOSF4_MEAS17 As Integer       ' SXLOSF4����_17
    SXLOSF4_MEAS18 As Integer       ' SXLOSF4����_18
    SXLOSF4_MEAS19 As Integer       ' SXLOSF4����_19
    SXLOSF4_MEAS20 As Integer       ' SXLOSF4����_20
    SXLBMD_SMPPOS As Integer        ' SXLBMD����ّ���ʒu�iSXL�ʒu���j
    SXLBMD1_KKSP As String * 3      ' SXLBMD1�������ב���ʒu
    SXLBMD1_NETU As String * 2      ' SXLBMD1�M�����@
    SXLBMD1_KKSET As String * 3     ' SXLBMD1�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLBMD1_MEAS1 As Integer        ' SXLBMD1����_�P
    SXLBMD1_MEAS2 As Integer        ' SXLBMD1����_2
    SXLBMD1_MEAS3 As Integer        ' SXLBMD1����_3
    SXLBMD1_MEAS4 As Integer        ' SXLBMD1����_4
    SXLBMD1_MEAS5 As Integer        ' SXLBMD1����_5
    SXLBMD2_KKSP As String * 3      ' SXLBMD�Q�������ב���ʒu
    SXLBMD2_NETU As String * 2      ' SXLBMD�Q�M�����@
    SXLBMD2_KKSET As String * 3     ' SXLBMD�Q�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLBMD2_MEAS1 As Integer        ' SXLBMD2����_�P
    SXLBMD2_MEAS2 As Integer        ' SXLBMD2����_2
    SXLBMD2_MEAS3 As Integer        ' SXLBMD2����_3
    SXLBMD2_MEAS4 As Integer        ' SXLBMD2����_4
    SXLBMD2_MEAS5 As Integer        ' SXLBMD2����_5
    SXLBMD3_KKSP As String * 3      ' SXLBMD�R�������ב���ʒu
    SXLBMD3_NETU As String * 2      ' SXLBMD�R�M�����@
    SXLBMD3_KKSET As String * 3     ' SXLBMD�R�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLBMD3_MEAS1 As Integer        ' SXLBMD3����_�P
    SXLBMD3_MEAS2 As Integer        ' SXLBMD3����_2
    SXLBMD3_MEAS3 As Integer        ' SXLBMD3����_3
    SXLBMD3_MEAS4 As Integer        ' SXLBMD3����_4
    SXLBMD3_MEAS5 As Integer        ' SXLBMD3����_5
    SXLGD_SMPPOS As Integer         ' SXLGD����ّ���ʒu�iSXL�ʒu���j
    SXLGD_MS01LDL1 As Integer       ' SXLGD_����l01 L/DL1
    SXLGD_MS01LDL2 As Integer       ' SXLGD_����l01 L/DL2
    SXLGD_MS01LDL3 As Integer       ' SXLGD_����l01 L/DL3
    SXLGD_MS01LDL4 As Integer       ' SXLGD_����l01 L/DL4
    SXLGD_MS01LDL5 As Integer       ' SXLGD_����l01 L/DL5
    SXLGD_MS01DEN1 As Integer       ' SXLGD_����l01 Den1
    SXLGD_MS01DEN2 As Integer       ' SXLGD_����l01 Den2
    SXLGD_MS01DEN3 As Integer       ' SXLGD_����l01 Den3
    SXLGD_MS01DEN4 As Integer       ' SXLGD_����l01 Den4
    SXLGD_MS01DEN5 As Integer       ' SXLGD_����l01 Den5
    SXLGD_MS02LDL1 As Integer       ' SXLGD_����l02 L/DL1
    SXLGD_MS02LDL2 As Integer       ' SXLGD_����l02 L/DL2
    SXLGD_MS02LDL3 As Integer       ' SXLGD_����l02 L/DL3
    SXLGD_MS02LDL4 As Integer       ' SXLGD_����l02 L/DL4
    SXLGD_MS02LDL5 As Integer       ' SXLGD_����l02 L/DL5
    SXLGD_MS02DEN1 As Integer       ' SXLGD_����l02 Den1
    SXLGD_MS02DEN2 As Integer       ' SXLGD_����l02 Den2
    SXLGD_MS02DEN3 As Integer       ' SXLGD_����l02 Den3
    SXLGD_MS02DEN4 As Integer       ' SXLGD_����l02 Den4
    SXLGD_MS02DEN5 As Integer       ' SXLGD_����l02 Den5
    SXLGD_MS03LDL1 As Integer       ' SXLGD_����l03 L/DL1
    SXLGD_MS03LDL2 As Integer       ' SXLGD_����l03 L/DL2
    SXLGD_MS03LDL3 As Integer       ' SXLGD_����l03 L/DL3
    SXLGD_MS03LDL4 As Integer       ' SXLGD_����l03 L/DL4
    SXLGD_MS03LDL5 As Integer       ' SXLGD_����l03 L/DL5
    SXLGD_MS03DEN1 As Integer       ' SXLGD_����l03 Den1
    SXLGD_MS03DEN2 As Integer       ' SXLGD_����l03 Den2
    SXLGD_MS03DEN3 As Integer       ' SXLGD_����l03 Den3
    SXLGD_MS03DEN4 As Integer       ' SXLGD_����l03 Den4
    SXLGD_MS03DEN5 As Integer       ' SXLGD_����l03 Den5
    SXLGD_MS04LDL1 As Integer       ' SXLGD_����l04 L/DL1
    SXLGD_MS04LDL2 As Integer       ' SXLGD_����l04 L/DL2
    SXLGD_MS04LDL3 As Integer       ' SXLGD_����l04 L/DL3
    SXLGD_MS04LDL4 As Integer       ' SXLGD_����l04 L/DL4
    SXLGD_MS04LDL5 As Integer       ' SXLGD_����l04 L/DL5
    SXLGD_MS04DEN1 As Integer       ' SXLGD_����l04 Den1
    SXLGD_MS04DEN2 As Integer       ' SXLGD_����l04 Den2
    SXLGD_MS04DEN3 As Integer       ' SXLGD_����l04 Den3
    SXLGD_MS04DEN4 As Integer       ' SXLGD_����l04 Den4
    SXLGD_MS04DEN5 As Integer       ' SXLGD_����l04 Den5
    SXLGD_MS05LDL1 As Integer       ' SXLGD_����l05 L/DL1
    SXLGD_MS05LDL2 As Integer       ' SXLGD_����l05 L/DL2
    SXLGD_MS05LDL3 As Integer       ' SXLGD_����l05 L/DL3
    SXLGD_MS05LDL4 As Integer       ' SXLGD_����l05 L/DL4
    SXLGD_MS05LDL5 As Integer       ' SXLGD_����l05 L/DL5
    SXLGD_MS05DEN1 As Integer       ' SXLGD_����l05 Den1
    SXLGD_MS05DEN2 As Integer       ' SXLGD_����l05 Den2
    SXLGD_MS05DEN3 As Integer       ' SXLGD_����l05 Den3
    SXLGD_MS05DEN4 As Integer       ' SXLGD_����l05 Den4
    SXLGD_MS05DEN5 As Integer       ' SXLGD_����l05 Den5
    SXLGD_MS06LDL1 As Integer       ' SXLGD_����l06 L/DL1
    SXLGD_MS06LDL2 As Integer       ' SXLGD_����l06 L/DL2
    SXLGD_MS06LDL3 As Integer       ' SXLGD_����l06 L/DL3
    SXLGD_MS06LDL4 As Integer       ' SXLGD_����l06 L/DL4
    SXLGD_MS06LDL5 As Integer       ' SXLGD_����l06 L/DL5
    SXLGD_MS06DEN1 As Integer       ' SXLGD_����l06 Den1
    SXLGD_MS06DEN2 As Integer       ' SXLGD_����l06 Den2
    SXLGD_MS06DEN3 As Integer       ' SXLGD_����l06 Den3
    SXLGD_MS06DEN4 As Integer       ' SXLGD_����l06 Den4
    SXLGD_MS06DEN5 As Integer       ' SXLGD_����l06 Den5
    SXLGD_MS07LDL1 As Integer       ' SXLGD_����l07 L/DL1
    SXLGD_MS07LDL2 As Integer       ' SXLGD_����l07 L/DL2
    SXLGD_MS07LDL3 As Integer       ' SXLGD_����l07 L/DL3
    SXLGD_MS07LDL4 As Integer       ' SXLGD_����l07 L/DL4
    SXLGD_MS07LDL5 As Integer       ' SXLGD_����l07 L/DL5
    SXLGD_MS07DEN1 As Integer       ' SXLGD_����l07 Den1
    SXLGD_MS07DEN2 As Integer       ' SXLGD_����l07 Den2
    SXLGD_MS07DEN3 As Integer       ' SXLGD_����l07 Den3
    SXLGD_MS07DEN4 As Integer       ' SXLGD_����l07 Den4
    SXLGD_MS07DEN5 As Integer       ' SXLGD_����l07 Den5
    SXLGD_MS08LDL1 As Integer       ' SXLGD_����l08 L/DL1
    SXLGD_MS08LDL2 As Integer       ' SXLGD_����l08 L/DL2
    SXLGD_MS08LDL3 As Integer       ' SXLGD_����l08 L/DL3
    SXLGD_MS08LDL4 As Integer       ' SXLGD_����l08 L/DL4
    SXLGD_MS08LDL5 As Integer       ' SXLGD_����l08 L/DL5
    SXLGD_MS08DEN1 As Integer       ' SXLGD_����l08 Den1
    SXLGD_MS08DEN2 As Integer       ' SXLGD_����l08 Den2
    SXLGD_MS08DEN3 As Integer       ' SXLGD_����l08 Den3
    SXLGD_MS08DEN4 As Integer       ' SXLGD_����l08 Den4
    SXLGD_MS08DEN5 As Integer       ' SXLGD_����l08 Den5
    SXLGD_MS09LDL1 As Integer       ' SXLGD_����l09 L/DL1
    SXLGD_MS09LDL2 As Integer       ' SXLGD_����l09 L/DL2
    SXLGD_MS09LDL3 As Integer       ' SXLGD_����l09 L/DL3
    SXLGD_MS09LDL4 As Integer       ' SXLGD_����l09 L/DL4
    SXLGD_MS09LDL5 As Integer       ' SXLGD_����l09 L/DL5
    SXLGD_MS09DEN1 As Integer       ' SXLGD_����l09 Den1
    SXLGD_MS09DEN2 As Integer       ' SXLGD_����l09 Den2
    SXLGD_MS09DEN3 As Integer       ' SXLGD_����l09 Den3
    SXLGD_MS09DEN4 As Integer       ' SXLGD_����l09 Den4
    SXLGD_MS09DEN5 As Integer       ' SXLGD_����l09 Den5
    SXLGD_MS10LDL1 As Integer       ' SXLGD_����l10 L/DL1
    SXLGD_MS10LDL2 As Integer       ' SXLGD_����l10 L/DL2
    SXLGD_MS10LDL3 As Integer       ' SXLGD_����l10 L/DL3
    SXLGD_MS10LDL4 As Integer       ' SXLGD_����l10 L/DL4
    SXLGD_MS10LDL5 As Integer       ' SXLGD_����l10 L/DL5
    SXLGD_MS10DEN1 As Integer       ' SXLGD_����l10 Den1
    SXLGD_MS10DEN2 As Integer       ' SXLGD_����l10 Den2
    SXLGD_MS10DEN3 As Integer       ' SXLGD_����l10 Den3
    SXLGD_MS10DEN4 As Integer       ' SXLGD_����l10 Den4
    SXLGD_MS10DEN5 As Integer       ' SXLGD_����l10 Den5
    SXLGD_MS11LDL1 As Integer       ' SXLGD_����l11 L/DL1
    SXLGD_MS11LDL2 As Integer       ' SXLGD_����l11 L/DL2
    SXLGD_MS11LDL3 As Integer       ' SXLGD_����l11 L/DL3
    SXLGD_MS11LDL4 As Integer       ' SXLGD_����l11 L/DL4
    SXLGD_MS11LDL5 As Integer       ' SXLGD_����l11 L/DL5
    SXLGD_MS11DEN1 As Integer       ' SXLGD_����l11 Den1
    SXLGD_MS11DEN2 As Integer       ' SXLGD_����l11 Den2
    SXLGD_MS11DEN3 As Integer       ' SXLGD_����l11 Den3
    SXLGD_MS11DEN4 As Integer       ' SXLGD_����l11 Den4
    SXLGD_MS11DEN5 As Integer       ' SXLGD_����l11 Den5
    SXLGD_MS12LDL1 As Integer       ' SXLGD_����l12 L/DL1
    SXLGD_MS12LDL2 As Integer       ' SXLGD_����l12 L/DL2
    SXLGD_MS12LDL3 As Integer       ' SXLGD_����l12 L/DL3
    SXLGD_MS12LDL4 As Integer       ' SXLGD_����l12 L/DL4
    SXLGD_MS12LDL5 As Integer       ' SXLGD_����l12 L/DL5
    SXLGD_MS12DEN1 As Integer       ' SXLGD_����l12 Den1
    SXLGD_MS12DEN2 As Integer       ' SXLGD_����l12 Den2
    SXLGD_MS12DEN3 As Integer       ' SXLGD_����l12 Den3
    SXLGD_MS12DEN4 As Integer       ' SXLGD_����l12 Den4
    SXLGD_MS12DEN5 As Integer       ' SXLGD_����l12 Den5
    SXLGD_MS13LDL1 As Integer       ' SXLGD_����l13 L/DL1
    SXLGD_MS13LDL2 As Integer       ' SXLGD_����l13 L/DL2
    SXLGD_MS13LDL3 As Integer       ' SXLGD_����l13 L/DL3
    SXLGD_MS13LDL4 As Integer       ' SXLGD_����l13 L/DL4
    SXLGD_MS13LDL5 As Integer       ' SXLGD_����l13 L/DL5
    SXLGD_MS13DEN1 As Integer       ' SXLGD_����l13 Den1
    SXLGD_MS13DEN2 As Integer       ' SXLGD_����l13 Den2
    SXLGD_MS13DEN3 As Integer       ' SXLGD_����l13 Den3
    SXLGD_MS13DEN4 As Integer       ' SXLGD_����l13 Den4
    SXLGD_MS13DEN5 As Integer       ' SXLGD_����l13 Den5
    SXLGD_MS14LDL1 As Integer       ' SXLGD_����l14 L/DL1
    SXLGD_MS14LDL2 As Integer       ' SXLGD_����l14 L/DL2
    SXLGD_MS14LDL3 As Integer       ' SXLGD_����l14 L/DL3
    SXLGD_MS14LDL4 As Integer       ' SXLGD_����l14 L/DL4
    SXLGD_MS14LDL5 As Integer       ' SXLGD_����l14 L/DL5
    SXLGD_MS14DEN1 As Integer       ' SXLGD_����l14 Den1
    SXLGD_MS14DEN2 As Integer       ' SXLGD_����l14 Den2
    SXLGD_MS14DEN3 As Integer       ' SXLGD_����l14 Den3
    SXLGD_MS14DEN4 As Integer       ' SXLGD_����l14 Den4
    SXLGD_MS14DEN5 As Integer       ' SXLGD_����l14 Den5
    SXLGD_MS15LDL1 As Integer       ' SXLGD_����l15 L/DL1
    SXLGD_MS15LDL2 As Integer       ' SXLGD_����l15 L/DL2
    SXLGD_MS15LDL3 As Integer       ' SXLGD_����l15 L/DL3
    SXLGD_MS15LDL4 As Integer       ' SXLGD_����l15 L/DL4
    SXLGD_MS15LDL5 As Integer       ' SXLGD_����l15 L/DL5
    SXLGD_MS15DEN1 As Integer       ' SXLGD_����l15 Den1
    SXLGD_MS15DEN2 As Integer       ' SXLGD_����l15 Den2
    SXLGD_MS15DEN3 As Integer       ' SXLGD_����l15 Den3
    SXLGD_MS15DEN4 As Integer       ' SXLGD_����l15 Den4
    SXLGD_MS15DEN5 As Integer       ' SXLGD_����l15 Den5
    SXLT_SMPPOS As Integer          ' SXLLT����ّ���ʒu�iSXL�ʒu���j
    SXLLT_MEASPEAK As Integer       ' SXLLT_����l �s�[�N�l
    SXLLT_MEAS1 As Integer          ' SXLLT_����l1
    SXLLT_MEAS2 As Integer          ' SXLLT_����l2
    SXLLT_MEAS3 As Integer          ' SXLLT_����l3
    SXLLT_MEAS4 As Integer          ' SXLLT_����l4
    SXLLT_MEAS5 As Integer          ' SXLLT_����l5
    WFDOI_SMPPOS As Integer         ' WFDOI�����-ID����ʒu�iSXL�ʒu���j�@number(4)
    WFDOI1_NETSU As String * 2      ' WFDOI-1_�M��������
    WFDOI1_MES As String * 3        ' WFDOI-1_�v�����@
    WFDOI1_MESDATA1 As String * 10  ' WFDOI-1_����l�P
    WFDOI1_MESDATA2 As String * 10  ' WFDOI-1_����l2
    WFDOI1_MESDATA3 As String * 10  ' WFDOI-1_����l3
    WFDOI1_MESDATA4 As String * 10  ' WFDOI-1_����l4
    WFDOI1_MESDATA5 As String * 10  ' WFDOI-1_����l5
    WFDOI1_MESDATA6 As String * 10  ' WFDOI-1_����l6
    WFDOI1_MESDATA7 As String * 10  ' WFDOI-1_����l7
    WFDOI1_MESDATA8 As String * 10  ' WFDOI-1_����l8
    WFDOI1_MESDATA9 As String * 10  ' WFDOI-1_����l9
    WFDOI1_MESDATA10 As String * 10 ' WFDOI-1_����l10
    WFDOI1_MESDATA11 As String * 10 ' WFDOI-1_����l11
    WFDOI1_MESDATA12 As String * 10 ' WFDOI-1_����l12
    WFDOI1_MESDATA13 As String * 10 ' WFDOI-1_����l13
    WFDOI1_MESDATA14 As String * 10 ' WFDOI-1_����l14
    WFDOI1_MESDATA15 As String * 10 ' WFDOI-1_����l15
    WFDOI2_NETSU As String * 2      ' WFDOI-2_�M��������
    WFDOI2_MES As String * 3        ' WFDOI-2_�v�����@
    WFDOI2_MESDATA1 As String * 10  ' WFDOI-2_����l�P
    WFDOI2_MESDATA2 As String * 10  ' WFDOI-2_����l2
    WFDOI2_MESDATA3 As String * 10  ' WFDOI-2_����l3
    WFDOI2_MESDATA4 As String * 10  ' WFDOI-2_����l4
    WFDOI2_MESDATA5 As String * 10  ' WFDOI-2_����l5
    WFDOI2_MESDATA6 As String * 10  ' WFDOI-2_����l6
    WFDOI2_MESDATA7 As String * 10  ' WFDOI-2_����l7
    WFDOI2_MESDATA8 As String * 10  ' WFDOI-2_����l8
    WFDOI2_MESDATA9 As String * 10  ' WFDOI-2_����l9
    WFDOI2_MESDATA10 As String * 10 ' WFDOI-2_����l10
    WFDOI2_MESDATA11 As String * 10 ' WFDOI-2_����l11
    WFDOI2_MESDATA12 As String * 10 ' WFDOI-2_����l12
    WFDOI2_MESDATA13 As String * 10 ' WFDOI-2_����l13
    WFDOI2_MESDATA14 As String * 10 ' WFDOI-2_����l14
    WFDOI2_MESDATA15 As String * 10 ' WFDOI-2_����l15
    WFDOI3_NETSU As String * 2      ' WFDOI-3_�M��������
    WFDOI3_MES As String * 3        ' WFDOI-3_�v�����@
    WFDOI3_MESDATA1 As String * 10  ' WFDOI-3_����l�P
    WFDOI3_MESDATA2 As String * 10  ' WFDOI-3_����l2
    WFDOI3_MESDATA3 As String * 10  ' WFDOI-3_����l3
    WFDOI3_MESDATA4 As String * 10  ' WFDOI-3_����l4
    WFDOI3_MESDATA5 As String * 10  ' WFDOI-3_����l5
    WFDOI3_MESDATA6 As String * 10  ' WFDOI-3_����l6
    WFDOI3_MESDATA7 As String * 10  ' WFDOI-3_����l7
    WFDOI3_MESDATA8 As String * 10  ' WFDOI-3_����l8
    WFDOI3_MESDATA9 As String * 10  ' WFDOI-3_����l9
    WFDOI3_MESDATA10 As String * 10 ' WFDOI-3_����l10
    WFDOI3_MESDATA11 As String * 10 ' WFDOI-3_����l11
    WFDOI3_MESDATA12 As String * 10 ' WFDOI-3_����l12
    WFDOI3_MESDATA13 As String * 10 ' WFDOI-3_����l13
    WFDOI3_MESDATA14 As String * 10 ' WFDOI-3_����l14
    WFDOI3_MESDATA15 As String * 10 ' WFDOI-3_����l15
    WFOSF1_SMPPOS As Integer        ' WFOSF1�����-ID����ʒu�iSXL�ʒu���j
    WFOSF1_NETSU As String * 2      ' WFOSF1_�M��������
    WFOSF1_ET As String * 3         ' WFOSF1_�G�b�`���O����
    WFOSF1_MES As String * 3        ' WFOSF1_�v�����@
    WFOSF1_DKAN As String * 10      ' WFOSF1_�c�j�A�j�[������
    WFOSF1_MESDATA1 As String * 10  ' WFOSF1����_�P
    WFOSF1_MESDATA2 As String * 10  ' WFOSF1����_2
    WFOSF1_MESDATA3 As String * 10  ' WFOSF1����_3
    WFOSF1_MESDATA4 As String * 10  ' WFOSF1����_4
    WFOSF1_MESDATA5 As String * 10  ' WFOSF1����_5
    WFOSF1_MESDATA6 As String * 10  ' WFOSF1����_6
    WFOSF1_MESDATA7 As String * 10  ' WFOSF1����_7
    WFOSF1_MESDATA8 As String * 10  ' WFOSF1����_8
    WFOSF1_MESDATA9 As String * 10  ' WFOSF1����_9
    WFOSF1_MESDATA10 As String * 10 ' WFOSF1����_10
    WFOSF1_MESDATA11 As String * 10 ' WFOSF1����_11
    WFOSF1_MESDATA12 As String * 10 ' WFOSF1����_12
    WFOSF1_MESDATA13 As String * 10 ' WFOSF1����_13
    WFOSF1_MESDATA14 As String * 10 ' WFOSF1����_14
    WFOSF1_MESDATA15 As String * 10 ' WFOSF1����_15
    WFOSF2_SMPPOS As Integer        ' WFOSF�Q�����-ID����ʒu�iSXL�ʒu���j�@number(4)
    WFOSF2_NETSU As String * 2      ' WFOSF2_�M��������
    WFOSF2_ET As String * 3         ' WFOSF2_�G�b�`���O����
    WFOSF2_MES As String * 3        ' WFOSF2_�v�����@
    WFOSF2_DKAN As String * 10      ' WFOSF2_�c�j�A�j�[������
    WFOSF2_MESDATA1 As String * 10  ' WFOSF2����_�P
    WFOSF2_MESDATA2 As String * 10  ' WFOSF2����_2
    WFOSF2_MESDATA3 As String * 10  ' WFOSF2����_3
    WFOSF2_MESDATA4 As String * 10  ' WFOSF2����_4
    WFOSF2_MESDATA5 As String * 10  ' WFOSF2����_5
    WFOSF2_MESDATA6 As String * 10  ' WFOSF2����_6
    WFOSF2_MESDATA7 As String * 10  ' WFOSF2����_7
    WFOSF2_MESDATA8 As String * 10  ' WFOSF2����_8
    WFOSF2_MESDATA9 As String * 10  ' WFOSF2����_9
    WFOSF2_MESDATA10 As String * 10 ' WFOSF2����_10
    WFOSF2_MESDATA11 As String * 10 ' WFOSF2����_11
    WFOSF2_MESDATA12 As String * 10 ' WFOSF2����_12
    WFOSF2_MESDATA13 As String * 10 ' WFOSF2����_13
    WFOSF2_MESDATA14 As String * 10 ' WFOSF2����_14
    WFOSF2_MESDATA15 As String * 10 ' WFOSF2����_15
    WFOSF3_SMPPOS As Integer        ' WFOSF�R�����-ID����ʒu�iSXL�ʒu���j
    WFOSF3_NETSU As String * 2      ' WFOSF3_�M��������
    WFOSF3_ET As String * 3         ' WFOSF3_�G�b�`���O����
    WFOSF3_MES As String * 3        ' WFOSF3_�v�����@
    WFOSF3_DKAN As String * 10      ' WFOSF3_�c�j�A�j�[������
    WFOSF3_MESDATA1 As String * 10  ' WFOSF3����_�P
    WFOSF3_MESDATA2 As String * 10  ' WFOSF3����_2
    WFOSF3_MESDATA3 As String * 10  ' WFOSF3����_3
    WFOSF3_MESDATA4 As String * 10  ' WFOSF3����_4
    WFOSF3_MESDATA5 As String * 10  ' WFOSF3����_5
    WFOSF3_MESDATA6 As String * 10  ' WFOSF3����_6
    WFOSF3_MESDATA7 As String * 10  ' WFOSF3����_7
    WFOSF3_MESDATA8 As String * 10  ' WFOSF3����_8
    WFOSF3_MESDATA9 As String * 10  ' WFOSF3����_9
    WFOSF3_MESDATA10 As String * 10 ' WFOSF3����_10
    WFOSF3_MESDATA11 As String * 10 ' WFOSF3����_11
    WFOSF3_MESDATA12 As String * 10 ' WFOSF3����_12
    WFOSF3_MESDATA13 As String * 10 ' WFOSF3����_13
    WFOSF3_MESDATA14 As String * 10 ' WFOSF3����_14
    WFOSF3_MESDATA15 As String * 10 ' WFOSF3����_15
    WFOSF4_SMPPOS As Integer        ' WFOSF�S�����-ID����ʒu�iSXL�ʒu���j
    WFOSF4_NETSU As String * 2      ' WFOSF4_�M��������
    WFOSF4_ET As String * 3         ' WFOSF4_�G�b�`���O����
    WFOSF4_MES As String * 3        ' WFOSF4_�v�����@
    WFOSF4_DKAN As String * 10      ' WFOSF4_�c�j�A�j�[������
    WFOSF4_MESDATA1 As String * 10  ' WFOSF4����_�P
    WFOSF4_MESDATA2 As String * 10  ' WFOSF4����_2
    WFOSF4_MESDATA3 As String * 10  ' WFOSF4����_3
    WFOSF4_MESDATA4 As String * 10  ' WFOSF4����_4
    WFOSF4_MESDATA5 As String * 10  ' WFOSF4����_5
    WFOSF4_MESDATA6 As String * 10  ' WFOSF4����_6
    WFOSF4_MESDATA7 As String * 10  ' WFOSF4����_7
    WFOSF4_MESDATA8 As String * 10  ' WFOSF4����_8
    WFOSF4_MESDATA9 As String * 10  ' WFOSF4����_9
    WFOSF4_MESDATA10 As String * 10 ' WFOSF4����_10
    WFOSF4_MESDATA11 As String * 10 ' WFOSF4����_11
    WFOSF4_MESDATA12 As String * 10 ' WFOSF4����_12
    WFOSF4_MESDATA13 As String * 10 ' WFOSF4����_13
    WFOSF4_MESDATA14 As String * 10 ' WFOSF4����_14
    WFOSF4_MESDATA15 As String * 10 ' WFOSF4����_15
    WFBMD1_SMPPOS As Integer        ' WFBMD1�����-ID����ʒu�iSXL�ʒu���j
    WFBMD1_NETSU As String * 2      ' WFBMD1_�M��������
    WFBMD1_ET As String * 3         ' WFBMD1_�G�b�`���O����
    WFBMD1_MES As String * 3        ' WFBMD1_�v�����@
    WFBMD1_DKAN As String * 10      ' WFBMD1_�c�j�A�j�[������
    WFBMD1_MESDATA1 As String * 10  ' WFBMD1����_�P
    WFBMD1_MESDATA2 As String * 10  ' WFBMD1����_2
    WFBMD1_MESDATA3 As String * 10  ' WFBMD1����_3
    WFBMD1_MESDATA4 As String * 10  ' WFBMD1����_4
    WFBMD1_MESDATA5 As String * 10  ' WFBMD1����_5
    WFBMD1_MESDATA6 As String * 10  ' WFBMD1����_6
    WFBMD1_MESDATA7 As String * 10  ' WFBMD1����_7
    WFBMD1_MESDATA8 As String * 10  ' WFBMD1����_8
    WFBMD1_MESDATA9 As String * 10  ' WFBMD1����_9
    WFBMD1_MESDATA10 As String * 10 ' WFBMD1����_10
    WFBMD1_MESDATA11 As String * 10 ' WFBMD1����_11
    WFBMD1_MESDATA12 As String * 10 ' WFBMD1����_12
    WFBMD1_MESDATA13 As String * 10 ' WFBMD1����_13
    WFBMD1_MESDATA14 As String * 10 ' WFBMD1����_14
    WFBMD1_MESDATA15 As String * 10 ' WFBMD1����_15
    WFBMD2_SMPPOS As Integer        ' WFBMD�Q�����-ID����ʒu�iSXL�ʒu���j
    WFBMD2_NETSU As String * 2      ' WFBMD2_�M��������
    WFBMD2_ET As String * 3         ' WFBMD2_�G�b�`���O����
    WFBMD2_MES As String * 3        ' WFBMD2_�v�����@
    WFBMD2_DKAN As String * 10      ' WFBMD2_�c�j�A�j�[������
    WFBMD2_MESDATA1 As String * 10  ' WFBMD2����_�P
    WFBMD2_MESDATA2 As String * 10  ' WFBMD2����_2
    WFBMD2_MESDATA3 As String * 10  ' WFBMD2����_3
    WFBMD2_MESDATA4 As String * 10  ' WFBMD2����_4
    WFBMD2_MESDATA5 As String * 10  ' WFBMD2����_5
    WFBMD2_MESDATA6 As String * 10  ' WFBMD2����_6
    WFBMD2_MESDATA7 As String * 10  ' WFBMD2����_7
    WFBMD2_MESDATA8 As String * 10  ' WFBMD2����_8
    WFBMD2_MESDATA9 As String * 10  ' WFBMD2����_9
    WFBMD2_MESDATA10 As String * 10 ' WFBMD2����_10
    WFBMD2_MESDATA11 As String * 10 ' WFBMD2����_11
    WFBMD2_MESDATA12 As String * 10 ' WFBMD2����_12
    WFBMD2_MESDATA13 As String * 10 ' WFBMD2����_13
    WFBMD2_MESDATA14 As String * 10 ' WFBMD2����_14
    WFBMD2_MESDATA15 As String * 10 ' WFBMD2����_15
    WFBMD3_SMPPOS As Integer        ' WFBMD�R�����-ID����ʒu�iSXL�ʒu���j
    WFBMD3_NETSU As String * 2      ' WFBMD3_�M��������
    WFBMD3_ET As String * 3         ' WFBMD3_�G�b�`���O����
    WFBMD3_MES As String * 3        ' WFBMD3_�v�����@
    WFBMD3_DKAN As String * 10      ' WFBMD3_�c�j�A�j�[������
    WFBMD3_MESDATA1 As String * 10  ' WFBMD3����_�P
    WFBMD3_MESDATA2 As String * 10  ' WFBMD3����_2
    WFBMD3_MESDATA3 As String * 10  ' WFBMD3����_3
    WFBMD3_MESDATA4 As String * 10  ' WFBMD3����_4
    WFBMD3_MESDATA5 As String * 10  ' WFBMD3����_5
    WFBMD3_MESDATA6 As String * 10  ' WFBMD3����_6
    WFBMD3_MESDATA7 As String * 10  ' WFBMD3����_7
    WFBMD3_MESDATA8 As String * 10  ' WFBMD3����_8
    WFBMD3_MESDATA9 As String * 10  ' WFBMD3����_9
    WFBMD3_MESDATA10 As String * 10 ' WFBMD3����_10
    WFBMD3_MESDATA11 As String * 10 ' WFBMD3����_11
    WFBMD3_MESDATA12 As String * 10 ' WFBMD3����_12
    WFBMD3_MESDATA13 As String * 10 ' WFBMD3����_13
    WFBMD3_MESDATA14 As String * 10 ' WFBMD3����_14
    WFBMD3_MESDATA15 As String * 10 ' WFBMD3����_15
    WFDSOD_SMPPOS As Integer        ' WFDSOD�����-ID����ʒu�iSXL�ʒu���j
    WFDSOD_NETSU As String * 2      ' WFDSOD_�M��������
    WFDSOD_ET As String * 3         ' WFDSOD_�G�b�`���O����
    WFDSOD_MES As String * 3        ' WFDSOD_�v�����@
    WFDSOD_DKAN As String * 10      ' WFDSOD_�c�j�A�j�[������
    WFDSOD_MESDATA1 As String * 10  ' WFDSOD����_�P
    WFDSOD_MESDATA2 As String * 10  ' WFDSOD����_2
    WFDSOD_MESDATA3 As String * 10  ' WFDSOD����_3
    WFDSOD_MESDATA4 As String * 10  ' WFDSOD����_4
    WFDSOD_MESDATA5 As String * 10  ' WFDSOD����_5
    WFDSOD_MESDATA6 As String * 10  ' WFDSOD����_6
    WFDSOD_MESDATA7 As String * 10  ' WFDSOD����_7
    WFDSOD_MESDATA8 As String * 10  ' WFDSOD����_8
    WFDSOD_MESDATA9 As String * 10  ' WFDSOD����_9
    WFDSOD_MESDATA10 As String * 10 ' WFDSOD����_10
    WFDSOD_MESDATA11 As String * 10 ' WFDSOD����_11
    WFDSOD_MESDATA12 As String * 10 ' WFDSOD����_12
    WFDSOD_MESDATA13 As String * 10 ' WFDSOD����_13
    WFDSOD_MESDATA14 As String * 10 ' WFDSOD����_14
    WFDSOD_MESDATA15 As String * 10 ' WFDSOD����_15
    WFSPV_SMPPOS As Integer         ' WFSPV�����-ID����ʒu�iSXL�ʒu���j
    WFSPV_NETSU As String * 2       ' WFSVP_�M��������
    WFSPV_ET As String * 3          ' WFSPV_�G�b�`���O����
    WFSPV_MES As String * 3         ' WFSPV_�v�����@
    WFSPV_DKAN As String * 10       ' WFSPV_�c�j�A�j�[������
    WFSPV_MESDATA1 As String * 10   ' WFSPV����_�P
    WFSPV_MESDATA2 As String * 10   ' WFSPV����_2
    WFSPV_MESDATA3 As String * 10   ' WFSPV����_3
    WFSPV_MESDATA4 As String * 10   ' WFSPV����_4
    WFSPV_MESDATA5 As String * 10   ' WFSPV����_5
    WFSPV_MESDATA6 As String * 10   ' WFSPV����_6
    WFSPV_MESDATA7 As String * 10   ' WFSPV����_7
    WFSPV_MESDATA8 As String * 10   ' WFSPV����_8
    WFSPV_MESDATA9 As String * 10   ' WFSPV����_9
    WFSPV_MESDATA10 As String * 10  ' WFSPV����_10
    WFSPV_MESDATA11 As String * 10  ' WFSPV����_11
    WFSPV_MESDATA12 As String * 10  ' WFSPV����_12
    WFSPV_MESDATA13 As String * 10  ' WFSPV����_13
    WFSPV_MESDATA14 As String * 10  ' WFSPV����_14
    WFSPV_MESDATA15 As String * 10  ' WFSPV����_15
    WFDZ_SMPPOS As Integer          ' WFDZ�����-ID����ʒu�iSXL�ʒu���j
    WFDZ_NETSU As String * 2        ' WFDZ_�M��������
    WFDZ_ET As String * 3           ' WFDZ_�G�b�`���O����
    WFDZ_MES As String * 3          ' WFDZ_�v�����@
    WFDZ_DKAN As String * 10        ' WFDZ_�c�j�A�j�[������
    WFDZ_MESDATA1 As String * 10    ' WFDZ����_�P
    WFDZ_MESDATA2 As String * 10    ' WFDZ����_2
    WFDZ_MESDATA3 As String * 10    ' WFDZ����_3
    WFDZ_MESDATA4 As String * 10    ' WFDZ����_4
    WFDZ_MESDATA5 As String * 10    ' WFDZ����_5
    WFDZ_MESDATA6 As String * 10    ' WFDZ����_6
    WFDZ_MESDATA7 As String * 10    ' WFDZ����_7
    WFDZ_MESDATA8 As String * 10    ' WFDZ����_8
    WFDZ_MESDATA9 As String * 10    ' WFDZ����_9
    WFDZ_MESDATA10 As String * 10   ' WFDZ����_10
    WFDZ_MESDATA11 As String * 10   ' WFDZ����_11
    WFDZ_MESDATA12 As String * 10   ' WFDZ����_12
    WFDZ_MESDATA13 As String * 10   ' WFDZ����_13
    WFDZ_MESDATA14 As String * 10   ' WFDZ����_14
    WFDZ_MESDATA15 As String * 10   ' WFDZ����_15
    REGDATE As Date                 ' �o�^���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lWF�ް��P
Public Type typ_TBCME008
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    CONFLAG As String * 1           ' �m�F�t���O
    REINFLAG As String * 1          ' �ĕt�^�t���O
    KPRFACES As String * 2          ' �w���i�\�ʎd�グ
    KPRBACKS As String * 2          ' �w���i���d�グ
    KPRBACK2 As String * 2          ' �w���i���d�グ�Q
    KPRBDSWY As String * 2          ' �w���i�a�c�������@
    KPRFKBWK As String * 1          ' �w���i�\�ʋ敪���@�Q��
    KPRFKBWS As String * 1          ' �w���i�\�ʋ敪���@�Q�w
    KPRTYPE As String * 1           ' �w���i�^�C�v
    KPRTYPKB As String * 1          ' �w���i�^�C�v�����敪
    KPRTYPKW As String * 1          ' �w���i�^�C�v�������@
    KPRDOP As String * 1            ' �w���i�h�[�p���g
    KPRRMIN As Double               ' �w���i���R����
    KPRRMAX As Double               ' �w���i���R���
    KPRRUNIT As String * 1          ' �w���i���R�P��
    KPRRSPOH As String * 1          ' �w���i���R����ʒu�Q��
    KPRRSPOT As String * 1          ' �w���i���R����ʒu�Q�_
    KPRRSPOI As String * 1          ' �w���i���R����ʒu�Q��
    KPRRHWYT As String * 1          ' �w���i���R�ۏؕ��@�Q��
    KPRRHWYS As String * 1          ' �w���i���R�ۏؕ��@�Q��
    KPRRKKBN As String * 1          ' �w���i���R�����敪
    KPRRKWAY As String * 2          ' �w���i���R�������@
    KPRRKHNM As String * 1          ' �w���i���R�����p�x�Q��
    KPRRKHNN As String * 1          ' �w���i���R�����p�x�Q��
    KPRRKHNH As String * 1          ' �w���i���R�����p�x�Q��
    KPRRKHNU As String * 1          ' �w���i���R�����p�x�Q�E
    KPRRSDEV As Double              ' �w���i���R�W���΍�
    KPRRAMIN As Double              ' �w���i���R���ω���
    KPRRAMAX As Double              ' �w���i���R���Ϗ��
    KPRRMBNP As Double              ' �w���i���R�ʓ����z
    KPRRMCAL As String * 1          ' �w���i���R�ʓ��v�Z
    KPRRMBP2 As Double              ' �w���i���R�ʓ����z�Q
    KPRRMCL2 As String * 1          ' �w���i���R�ʓ��v�Z�Q
    KPRRKBSH As String * 1          ' �w���i���R�U�敪����ʒu�Q��
    KPRRKBST As String * 1          ' �w���i���R�U�敪����ʒu�Q�_
    KPRRKBSI As String * 1          ' �w���i���R�U�敪����ʒu�Q��
    KPRRKBHT As String * 1          ' �w���i���R�U�敪�ۏؕ��@�Q��
    KPRRKBHS As String * 1          ' �w���i���R�U�敪�ۏؕ��@�Q��
    KPRSTMAX As Double              ' �w���i�X�g���G���
    KPRSTSPH As String * 1          ' �w���i�X�g���G����ʒu�Q��
    KPRSTSPT As String * 1          ' �w���i�X�g���G����ʒu�Q�_
    KPRSTSPI As String * 1          ' �w���i�X�g���G����ʒu�Q��
    KPRSTHWT As String * 1          ' �w���i�X�g���G�ۏؕ��@�Q��
    KPRSTHWS As String * 1          ' �w���i�X�g���G�ۏؕ��@�Q��
    KPRSTKBN As String * 1          ' �w���i�X�g���G�����敪
    KPRSTKWY As String * 2          ' �w���i�X�g���G�������@
    KPRSTKHM As String * 1          ' �w���i�X�g���G�����p�x�Q��
    KPRSTKHN As String * 1          ' �w���i�X�g���G�����p�x�Q��
    KPRSTKHH As String * 1          ' �w���i�X�g���G�����p�x�Q��
    KPRSTKHU As String * 1          ' �w���i�X�g���G�����p�x�Q�E
    KPRRHCAL As String * 2          ' �w���i���R�␳�v�Z
    KPRRMINH As Double              ' �w���i���R�����␳
    KPRRMAXH As Double              ' �w���i���R����␳
    KPRACEN As Double               ' �w���i�����S
    KPRAMIN As Double               ' �w���i������
    KPRAMAX As Double               ' �w���i�����
    KPRAUNIT As String * 1          ' �w���i���P��
    KPRASPOH As String * 1          ' �w���i������ʒu�Q��
    KPRASPOT As String * 1          ' �w���i������ʒu�Q�_
    KPRASPOI As String * 1          ' �w���i������ʒu�Q��
    KPRAHWYT As String * 1          ' �w���i���ۏؕ��@�Q��
    KPRAHWYS As String * 1          ' �w���i���ۏؕ��@�Q��
    KPRAKKBN As String * 1          ' �w���i�������敪
    KPRAKWAY As String * 1          ' �w���i���������@
    KPRAKHNM As String * 1          ' �w���i�������p�x�Q��
    KPRAKHNN As String * 1          ' �w���i�������p�x�Q��
    KPRAKHNH As String * 1          ' �w���i�������p�x�Q��
    KPRAKHNU As String * 1          ' �w���i�������p�x�Q�E
    KPRASDEV As Double              ' �w���i���W���΍�
    KPRAAMIN As Double              ' �w���i�����ω���
    KPRAAMAX As Double              ' �w���i�����Ϗ��
    KPRAMBNP As Double              ' �w���i���ʓ����z
    KPRAMCAL As String * 1          ' �w���i���ʓ��v�Z
    KPRALTBP As Double              ' �w���i���k�s���z
    KPRALTCL As String * 1          ' �w���i���k�s�v�Z
    KPRALTRA As Double              ' �w���i���k�s�͈�
    KPRAMRAN As Double              ' �w���i���ʓ��͈�
    KPRAKBSH As String * 1          ' �w���i���U�敪����ʒu�Q��
    KPRAKBST As String * 1          ' �w���i���U�敪����ʒu�Q�_
    KPRAKBSI As String * 1          ' �w���i���U�敪����ʒu�Q��
    KPRAKBHT As String * 1          ' �w���i���U�敪�ۏؕ��@�Q��
    KPRAKBHS As String * 1          ' �w���i���U�敪�ۏؕ��@�Q��
    KPRWFORM As String * 1          ' �w���i�E�F�[�n�`��
    KPRD1CEN As Double              ' �w���i���a�P���S
    KPRD1MIN As Double              ' �w���i���a�P����
    KPRD1MAX As Double              ' �w���i���a�P���
    KPRD1KBN As String * 1          ' �w���i���a�P�����敪
    KPRD2CEN As Double              ' �w���i���a�Q���S
    KPRD2MIN As Double              ' �w���i���a�Q����
    KPRD2MAX As Double              ' �w���i���a�Q���
    KPRD2KBN As String * 1          ' �w���i���a�Q�����敪
    KPRDUNIT As String * 1          ' �w���i���a�P��
    KPRDKHNM As String * 1          ' �w���i���a�����p�x�Q��
    KPRDKHNN As String * 1          ' �w���i���a�����p�x�Q��
    KPRDKHNH As String * 1          ' �w���i���a�����p�x�Q��
    KPRDKHNU As String * 1          ' �w���i���a�����p�x�Q�E
    KPRLPMNP As Integer             ' �w���i�k�o���ŏ����H��
    KPRSGMNP As Integer             ' �w���i�r�f���ŏ����H��
    KPRETMNP As Integer             ' �w���i�d�s���ŏ����H��
    KPRMPMNP As Integer             ' �w���i�l�o���ŏ����H��
    KPRLPKS1 As String * 1          ' �w���i�k�o�����ގ�P
    KPRLPKS2 As String * 1          ' �w���i�k�o�����ގ�Q
    KPRLPKZ1 As String * 1          ' �w���i�k�o�����ޗ��x��P
    KPRLPKZ2 As String * 1          ' �w���i�k�o�����ޗ��x��Q
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lWF�ް��Q
Public Type typ_TBCME009
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KPRCDIR As String * 1           ' �w���i�����ʕ���
    KPRCSCEN As Double              ' �w���i�����ʌX���S
    KPRCSMIN As Double              ' �w���i�����ʌX����
    KPRCSMAX As Double              ' �w���i�����ʌX���
    KPRCSDIS As String * 1          ' �w���i�����ʌX���ʎw��
    KPRCSDIR As String * 2          ' �w���i�����ʌX����
    KPRCKKBN As String * 1          ' �w���i�����ʌ����敪
    KPRCKWAY As String * 2          ' �w���i�����ʌ������@
    KPRCKHNM As String * 1          ' �w���i�����ʌ����p�x�Q��
    KPRCKHNN As String * 1          ' �w���i�����ʌ����p�x�Q��
    KPRCKHNH As String * 1          ' �w���i�����ʌ����p�x�Q��
    KPRCKHNU As String * 1          ' �w���i�����ʌ����p�x�Q�E
    KPRCTDIR As String * 2          ' �w���i�����ʌX�c����
    KPRCTCEN As Double              ' �w���i�����ʌX�c���S
    KPRCTMIN As Double              ' �w���i�����ʌX�c����
    KPRCTMAX As Double              ' �w���i�����ʌX�c���
    KPRCYDIR As String * 2          ' �w���i�����ʌX������
    KPRCYCEN As Double              ' �w���i�����ʌX�����S
    KPRCYMIN As Double              ' �w���i�����ʌX������
    KPRCYMAX As Double              ' �w���i�����ʌX�����
    KPRCSDSC As Double              ' �w���i�����ʌX���ʌX���S
    KPRCSDSN As Double              ' �w���i�����ʌX���ʌX����
    KPRCSDSX As Double              ' �w���i�����ʌX���ʌX���
    KPROFPKM As String * 1          ' �w���i�n�e�ʒu�����p�x�Q��
    KPROFPKN As String * 1          ' �w���i�n�e�ʒu�����p�x�Q��
    KPROFPKH As String * 1          ' �w���i�n�e�ʒu�����p�x�Q��
    KPROFPKU As String * 1          ' �w���i�n�e�ʒu�����p�x�Q�E
    KPROFLKM As String * 1          ' �w���i�n�e�������p�x�Q��
    KPROFLKN As String * 1          ' �w���i�n�e�������p�x�Q��
    KPROFLKH As String * 1          ' �w���i�n�e�������p�x�Q��
    KPROFLKU As String * 1          ' �w���i�n�e�������p�x�Q�E
    KPROF1PD As String * 2          ' �w���i�n�e�P�ʒu����
    KPROF1PN As Double              ' �w���i�n�e�P�ʒu����
    KPROF1PX As Double              ' �w���i�n�e�P�ʒu���
    KPROF1PK As String * 1          ' �w���i�n�e�P�ʒu�����敪
    KPROF1PW As String * 2          ' �w���i�n�e�P�ʒu�������@
    KPROF1LC As Double              ' �w���i�n�e�P�����S
    KPROF1LN As Double              ' �w���i�n�e�P������
    KPROF1LX As Double              ' �w���i�n�e�P�����
    KPROF1LK As String * 1          ' �w���i�n�e�P�������敪
    KPROF1RF As String * 1          ' �w���i�n�e�P���[�q�`��
    KPROFRRC As Double              ' �w���i�n�e���[�q�E���S
    KPROFRRN As Double              ' �w���i�n�e���[�q�E����
    KPROFRRX As Double              ' �w���i�n�e���[�q�E���
    KPROFRLC As Double              ' �w���i�n�e���[�q�����S
    KPROFRLN As Double              ' �w���i�n�e���[�q������
    KPROFRLX As Double              ' �w���i�n�e���[�q�����
    KPROFRKB As String * 1          ' �w���i�n�e���[�q�����敪
    KPROF1DC As Double              ' �w���i�n�e�P���a���S
    KPROF1DN As Double              ' �w���i�n�e�P���a����
    KPROF1DX As Double              ' �w���i�n�e�P���a���
    KPROF1DK As String * 1          ' �w���i�n�e�P���a�����敪
    KPRDFORM As String * 1          ' �w���i�a�`��
    KPRDFKBN As String * 1          ' �w���i�a�`�󌟍��敪
    KPRDFKHM As String * 1          ' �w���i�a�`�󌟍��p�x�Q��
    KPRDFKHN As String * 1          ' �w���i�a�`�󌟍��p�x�Q��
    KPRDFKHH As String * 1          ' �w���i�a�`�󌟍��p�x�Q��
    KPRDFKHU As String * 1          ' �w���i�a�`�󌟍��p�x�Q�E
    KPRDPDRC As String * 1          ' �w���i�a�ʒu����
    KPRDPACN As Integer             ' �w���i�a�ʒu�p�x���S
    KPRDPAMN As Integer             ' �w���i�a�ʒu�p�x����
    KPRDPAMX As Integer             ' �w���i�a�ʒu�p�x���
    KPRDPDIR As String * 2          ' �w���i�a�ʒu����
    KPRDPMIN As Double              ' �w���i�a�ʒu����
    KPRDPMAX As Double              ' �w���i�a�ʒu���
    KPRDPKBN As String * 1          ' �w���i�a�ʒu�����敪
    KPRDPKWY As String * 2          ' �w���i�a�ʒu�������@
    KPRDPKHM As String * 1          ' �w���i�a�ʒu�����p�x�Q��
    KPRDPKHB As String * 1          ' �w���i�a�ʒu�����p�x�Q��
    KPRDPKHH As String * 1          ' �w���i�a�ʒu�����p�x�Q��
    KPRDPKHU As String * 1          ' �w���i�a�ʒu�����p�x�Q�E
    KPRDACEN As Double              ' �w���i�a�p�x���S
    KPRDAMIN As Double              ' �w���i�a�p�x����
    KPRDAMAX As Double              ' �w���i�a�p�x���
    KPRDAKBN As String * 1          ' �w���i�a�p�x�����敪
    KPRDWCEN As Double              ' �w���i�a�В��S
    KPRDWMIN As Double              ' �w���i�a�Љ���
    KPRDWMAX As Double              ' �w���i�a�Џ��
    KPRDWKBN As String * 1          ' �w���i�a�Ќ����敪
    KPRDDCEN As Double              ' �w���i�a�[���S
    KPRDDMIN As Double              ' �w���i�a�[����
    KPRDDMAX As Double              ' �w���i�a�[���
    KPRDDKBN As String * 1          ' �w���i�a�[�����敪
    KPRDBRCN As Double              ' �w���i�a��q���S
    KPRDBRMN As Double              ' �w���i�a��q����
    KPRDBRMX As Double              ' �w���i�a��q���
    KPRDBRKB As String * 1          ' �w���i�a��q�����敪
    KPRDRRCN As Double              ' �w���i�a�E�q���S
    KPRDRRMN As Double              ' �w���i�a�E�q����
    KPRDRRMX As Double              ' �w���i�a�E�q���
    KPRDLRCN As Double              ' �w���i�a���q���S
    KPRDLRMN As Double              ' �w���i�a���q����
    KPRDLRMX As Double              ' �w���i�a���q���
    KPRDRRKB As String * 1          ' �w���i�a���[�q�����敪
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lWF�ް��R
Public Type typ_TBCME010
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KPRMFORM As String * 1          ' �w���i�ʎ�`��
    KPRMM As String * 1             ' �w���i�ʎ�ʑe
    KPRMFKBN As String * 1          ' �w���i�ʎ�`�󌟍��敪
    KPRMFKHM As String * 1          ' �w���i�ʎ�`�󌟍��p�x�Q��
    KPRMFKHN As String * 1          ' �w���i�ʎ�`�󌟍��p�x�Q��
    KPRMFKHH As String * 1          ' �w���i�ʎ�`�󌟍��p�x�Q��
    KPRMFKHU As String * 1          ' �w���i�ʎ�`�󌟍��p�x�Q�E
    KPRMMKBN As String * 1          ' �w���i�ʎ�ʑe�����敪
    KPRMACEN As Double              ' �w���i�ʎ�p�x���S
    KPRMAMIN As Double              ' �w���i�ʎ�p�x����
    KPRMAMAX As Double              ' �w���i�ʎ�p�x���
    KPRMAKBN As String * 1          ' �w���i�ʎ�p�x�����敪
    KPRMWFCN As Integer             ' �w���i�ʎ�Е\���S
    KPRMWFMN As Integer             ' �w���i�ʎ�Е\����
    KPRMWFMX As Integer             ' �w���i�ʎ�Е\���
    KPRMWBCN As Integer             ' �w���i�ʎ�З����S
    KPRMWBMN As Integer             ' �w���i�ʎ�З�����
    KPRMWBMX As Integer             ' �w���i�ʎ�З����
    KPRMHKBN As String * 1          ' �w���i�ʎ捂�����敪
    KPRMHCEN As Integer             ' �w���i�ʎ捂���S
    KPRMHMIN As Integer             ' �w���i�ʎ捂����
    KPRMHMAX As Integer             ' �w���i�ʎ捂���
    KPRMWKBN As String * 1          ' �w���i�ʎ�Ќ����敪
    KPRMPWCN As Integer             ' �w���i�ʎ��[�В��S
    KPRMPWMN As Integer             ' �w���i�ʎ��[�Љ���
    KPRMPWMX As Integer             ' �w���i�ʎ��[�Џ��
    KPRMPWKB As String * 1          ' �w���i�ʎ��[�Ќ����敪
    KPRMPRCN As Double              ' �w���i�ʎ��[�q���S
    KPRMPRMN As Double              ' �w���i�ʎ��[�q����
    KPRMPRMX As Double              ' �w���i�ʎ��[�q���
    KPRMPRKB As String * 1          ' �w���i�ʎ��[�q�����敪
    KPRDMFRM As String * 1          ' �w���i�a�ʎ�`��
    KPRDMM As String * 1            ' �w���i�a�ʎ�ʑe
    KPRDMPRC As Double              ' �w���i�a�ʎ��[�q���S
    KPRDMACN As Double              ' �w���i�a�ʎ�p�x���S
    KPRIDSTA As String * 2          ' �w���i�h�c�K�i
    KPRIDWAY As String * 1          ' �w���i�h�c���@
    KPRIDPRI As String * 1          ' �w���i�h�c�󎚎��
    KPRIDKND As String * 1          ' �w���i�h�c���
    KPRIDDIR As String * 1          ' �w���i�h�c����
    KPRIDFAC As String * 1          ' �w���i�h�c��
    KPRCSIZE As String * 1          ' �w���i�����T�C�Y
    KPRIDPBS As String * 1          ' �w���i�h�c�ʒu����
    KPRIDFIG As Integer             ' �w���i�h�c����
    KPRIDCON As String              ' �w���i�h�c���e
    KPRIDZAR As Double              ' �w���i�h�c���O�̈�
    KPRIDPAP As String * 1          ' �w���i�h�c�󎚘A�Ԏw��
    KPRIDDCN As Integer             ' �w���i�h�c�h�b�g�[���S
    KPRIDDMX As Integer             ' �w���i�h�c�h�b�g�[���
    KPRIDDMN As Integer             ' �w���i�h�c�h�b�g�[����
    KPRIDSCN As Integer             ' �w���i�h�c�h�b�g�r���S
    KPRIDSMX As Integer             ' �w���i�h�c�h�b�g�r���
    KPRIDSMN As Integer             ' �w���i�h�c�h�b�g�r����
    KPRBDPRS As Double              ' �w���i�a�c����
    KPRBDTIM As Integer             ' �w���i�a�c��
    KPRETWAY As String * 2          ' �w���i�d�s���@
    KPRMPFIN As String * 1          ' �w���i�l�o�d�グ
    KPRLWASW As String * 1          ' �w���i�ŏI�����@
    KPRCDOP As String * 1           ' �w���i�����h�[�v
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lWF�ް��S
Public Type typ_TBCME011
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KPRM1S As String * 1            ' �w���i���P��
    KPRM1H As String * 1            ' �w���i���P�t��
    KPRM2S As String * 1            ' �w���i���Q��
    KPRM2H As String * 1            ' �w���i���Q�t��
    KPRNJSUM As String * 1          ' �w���i�m�W���[�������L��
    KPRNJSMX As Double              ' �w���i�m�W���[�������Џ��
    KPRNJSMN As Double              ' �w���i�m�W���[�������Љ���
    KPROXCEN As Long                ' �w���i�_���������S
    KPROXMIN As Long                ' �w���i�_����������
    KPROXMAX As Long                ' �w���i�_���������
    KPROXUNT As String * 1          ' �w���i�_�������P��
    KPROXSPH As String * 1          ' �w���i�_����������ʒu�Q��
    KPROXSPT As String * 1          ' �w���i�_����������ʒu�Q�_
    KPROXSPI As String * 1          ' �w���i�_����������ʒu�Q��
    KPROXHWT As String * 1          ' �w���i�_�������ۏؕ��@�Q��
    KPROXHWS As String * 1          ' �w���i�_�������ۏؕ��@�Q��
    KPROXHWY As String * 2          ' �w���i�_�������������@
    KPROXNPO As String * 1          ' �w���i�_����������ʒu
    KPROXKHM As String * 1          ' �w���i�_�����������p�x�Q��
    KPROXKHN As String * 1          ' �w���i�_�����������p�x�Q��
    KPROXKHH As String * 1          ' �w���i�_�����������p�x�Q��
    KPROXKHU As String * 1          ' �w���i�_�����������p�x�Q�E
    KPROXZAR As Integer             ' �w���i�_�������O�̈�
    KPROXMBP As Double              ' �w���i�_�������ʓ����z
    KPROXMCL As String * 1          ' �w���i�_�������ʓ��v�Z
    KPROXMRA As Integer             ' �w���i�_�������ʓ��͈�
    KPROXLTB As Double              ' �w���i�_�������k�s���z
    KPROXLTC As String * 1          ' �w���i�_�������k�s�v�Z
    KPROXLTR As Integer             ' �w���i�_�������k�s�͈�
    KPRPSCEN As Double              ' �w���i�|���V�������S
    KPRPSMIN As Double              ' �w���i�|���V��������
    KPRPSMAX As Double              ' �w���i�|���V�������
    KPRPSUNT As String * 1          ' �w���i�|���V�������P��
    KPRPSSPH As String * 1          ' �w���i�|���V��������ʒu�Q��
    KPRPSSPT As String * 1          ' �w���i�|���V��������ʒu�Q�_
    KPRPSSPI As String * 1          ' �w���i�|���V��������ʒu�Q��
    KPRPSHWT As String * 1          ' �w���i�|���V�����ۏؕ��@�Q��
    KPRPSHWS As String * 1          ' �w���i�|���V�����ۏؕ��@�Q��
    KPRPSKWY As String * 2          ' �w���i�|���V�����������@
    KPRPSNPS As String * 1          ' �w���i�|���V��������ʒu
    KPRPSKHM As String * 1          ' �w���i�|���V���������p�x�Q��
    KPRPSKHN As String * 1          ' �w���i�|���V���������p�x�Q��
    KPRPSKHH As String * 1          ' �w���i�|���V���������p�x�Q��
    KPRPSKHU As String * 1          ' �w���i�|���V���������p�x�Q�E
    KPRPSMBP As Double              ' �w���i�|���V�����ʓ����z
    KPRPSMCL As String * 1          ' �w���i�|���V�����ʓ��v�Z
    KPRPSMRA As Double              ' �w���i�|���V�����ʓ��͈�
    KPRNOXCN As Long                ' �w���i�����������S
    KPRNOXMN As Long                ' �w���i������������
    KPRNOXMX As Long                ' �w���i�����������
    KPRNOXUN As String * 1          ' �w���i���������P��
    KPRNOXSH As String * 1          ' �w���i������������ʒu�Q��
    KPRNOXST As String * 1          ' �w���i������������ʒu�Q�_
    KPRNOXSI As String * 1          ' �w���i������������ʒu�Q��
    KPRNOXHT As String * 1          ' �w���i���������ۏؕ��@�Q��
    KPRNOXHS As String * 1          ' �w���i���������ۏؕ��@�Q��
    KPRNOXHW As String * 2          ' �w���i���������������@
    KPRNOXNP As String * 1          ' �w���i������������ʒu
    KPRNOXKM As String * 1          ' �w���i�������������p�x�Q��
    KPRNOXKN As String * 1          ' �w���i�������������p�x�Q��
    KPRNOXKH As String * 1          ' �w���i�������������p�x�Q��
    KPRNOXKU As String * 1          ' �w���i�������������p�x�Q�E
    KPRNOXMB As Double              ' �w���i���������ʓ����z
    KPRNOXMC As String * 1          ' �w���i���������ʓ��v�Z
    KPRNOXMR As Integer             ' �w���i���������ʓ��͈�
    KPRMKMIN As Double              ' �w���i�����בw����
    KPRMKMAX As Double              ' �w���i�����בw���
    KPRMKSPH As String * 1          ' �w���i�����בw����ʒu�Q��
    KPRMKSPT As String * 1          ' �w���i�����בw����ʒu�Q�_
    KPRMKSPR As String * 1          ' �w���i�����בw����ʒu�Q��
    KPRMKHWT As String * 1          ' �w���i�����בw�ۏؕ��@�Q��
    KPRMKHWS As String * 1          ' �w���i�����בw�ۏؕ��@�Q��
    KPRMKSZY As String * 1          ' �w���i�����בw�������
    KPRMKKHM As String * 1          ' �w���i�����בw�����p�x�Q��
    KPRMKKHN As String * 1          ' �w���i�����בw�����p�x�Q��
    KPRMKKHH As String * 1          ' �w���i�����בw�����p�x�Q��
    KPRMKKHU As String * 1          ' �w���i�����בw�����p�x�Q�E
    KPRMKNSW As String * 2          ' �w���i�����בw�M�����@
    KPRMKFGS As String * 1          ' �w���i�����בw���͋C�K�X
    KPRMKCET As Integer             ' �w���i�����בw�I���d�s��
    KPRDZSWY As String * 1          ' �w���i�c�y�������@
    KPRD1STO As Integer             ' �w���i�c�y�P�r�s���x
    KPRD1STT As Integer             ' �w���i�c�y�P�r�s����
    KPRD1STG As String * 1          ' �w���i�c�y�P�r�s�K�X����
    KPRD2NDO As Integer             ' �w���i�c�y�Q�m�c���x
    KPRD2NDC As Integer             ' �w���i�c�y�Q�m�c���x���
    KPRD2NDT As Integer             ' �w���i�c�y�Q�m�c����
    KPRD3RDO As Integer             ' �w���i�c�y�R�q�c���x
    KPRD3RDT As Integer             ' �w���i�c�y�R�q�c����
    KPRDZMPS As String * 1          ' �w���i�c�y�l�o�����敪
    KPRH2ANO As Integer             ' �w���i�g�Q�`�m���x
    KPRH2ANT As Integer             ' �w���i�g�Q�`�m����
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lWF�ް��T
Public Type typ_TBCME012
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KPRTMMAX As Long                ' �w���i�]�ʖ��x���
    KPRTMSPH As String * 1          ' �w���i�]�ʖ��x����ʒu�Q��
    KPRTMSPT As String * 1          ' �w���i�]�ʖ��x����ʒu�Q�_
    KPRTMSPR As String * 1          ' �w���i�]�ʖ��x����ʒu�Q��
    KPRTMKBN As String * 1          ' �w���i�]�ʖ��x�����敪
    KPRTMKHM As String * 1          ' �w���i�]�ʖ��x�����p�x�Q��
    KPRTMKHN As String * 1          ' �w���i�]�ʖ��x�����p�x�Q��
    KPRTMKHH As String * 1          ' �w���i�]�ʖ��x�����p�x�Q��
    KPRTMKHU As String * 1          ' �w���i�]�ʖ��x�����p�x�Q�E
    KPRLTMIN As Integer             ' �w���i�k�^�C������
    KPRLTMAX As Integer             ' �w���i�k�^�C�����
    KPRLTUNT As String * 1          ' �w���i�k�^�C���P��
    KPRLTSPH As String * 1          ' �w���i�k�^�C������ʒu�Q��
    KPRLTSPT As String * 1          ' �w���i�k�^�C������ʒu�Q�_
    KPRLTSPI As String * 1          ' �w���i�k�^�C������ʒu�Q��
    KPRLTHWT As String * 1          ' �w���i�k�^�C���ۏؕ��@�Q��
    KPRLTHWS As String * 1          ' �w���i�k�^�C���ۏؕ��@�Q��
    KPRLTNSW As String * 2          ' �w���i�k�^�C���M�����@
    KPRLTKBN As String * 1          ' �w���i�k�^�C�������敪
    KPRLTKWY As String * 2          ' �w���i�k�^�C���������@
    KPRLTKHM As String * 1          ' �w���i�k�^�C�������p�x�Q��
    KPRLTKHN As String * 1          ' �w���i�k�^�C�������p�x�Q��
    KPRLTKHH As String * 1          ' �w���i�k�^�C�������p�x�Q��
    KPRLTKHU As String * 1          ' �w���i�k�^�C�������p�x�Q�E
    KPRLTMBP As Double              ' �w���i�k�^�C���ʓ����z
    KPRLTMCL As String * 1          ' �w���i�k�^�C���ʓ��v�Z
    KPRCNMIN As Double              ' �w���i�Y�f�Z�x����
    KPRCNMAX As Double              ' �w���i�Y�f�Z�x���
    KPRCNUNT As String * 1          ' �w���i�Y�f�Z�x�P��
    KPRCNIND As String * 2          ' �w���i�Y�f�Z�x�w��
    KPRCNSPH As String * 1          ' �w���i�Y�f�Z�x����ʒu�Q��
    KPRCNSPT As String * 1          ' �w���i�Y�f�Z�x����ʒu�Q�_
    KPRCNSPI As String * 1          ' �w���i�Y�f�Z�x����ʒu�Q��
    KPRCNHWT As String * 1          ' �w���i�Y�f�Z�x�ۏؕ��@�Q��
    KPRCNHWS As String * 1          ' �w���i�Y�f�Z�x�ۏؕ��@�Q��
    KPRCNKBN As String * 1          ' �w���i�Y�f�Z�x�����敪
    KPRCNKWY As String * 2          ' �w���i�Y�f�Z�x�������@
    KPRONMIN As Double              ' �w���i�_�f�Z�x����
    KPRONMAX As Double              ' �w���i�_�f�Z�x���
    KPRONUNT As String * 1          ' �w���i�_�f�Z�x�P��
    KPRONIND As String * 2          ' �w���i�_�f�Z�x�w��
    KPRONSPH As String * 1          ' �w���i�_�f�Z�x����ʒu�Q��
    KPRONSPT As String * 1          ' �w���i�_�f�Z�x����ʒu�Q�_
    KPRONSPI As String * 1          ' �w���i�_�f�Z�x����ʒu�Q��
    KPRONHWT As String * 1          ' �w���i�_�f�Z�x�ۏؕ��@�Q��
    KPRONHWS As String * 1          ' �w���i�_�f�Z�x�ۏؕ��@�Q��
    KPRONKBN As String * 1          ' �w���i�_�f�Z�x�����敪
    KPRONKWY As String * 2          ' �w���i�_�f�Z�x�������@
    KPRONKHM As String * 1          ' �w���i�_�f�Z�x�����p�x�Q��
    KPRONKHN As String * 1          ' �w���i�_�f�Z�x�����p�x�Q��
    KPRONKHH As String * 1          ' �w���i�_�f�Z�x�����p�x�Q��
    KPRONKHU As String * 1          ' �w���i�_�f�Z�x�����p�x�Q�E
    KPRONMBP As Double              ' �w���i�_�f�Z�x�ʓ����z
    KPRONMCL As String * 1          ' �w���i�_�f�Z�x�ʓ��v�Z
    KPRONLTB As Double              ' �w���i�_�f�Z�x�k�s���z
    KPRONLTC As String * 1          ' �w���i�_�f�Z�x�k�s�v�Z
    KPRONSDV As Double              ' �w���i�_�f�Z�x�W���΍�
    KPRONAMN As Double              ' �w���i�_�f�Z�x���ω���
    KPRONAMX As Double              ' �w���i�_�f�Z�x���Ϗ��
    KPRONAST As String * 1          ' �w���i�_�f�Z�x�`�r�s�l�V��
    KPRONHCL As String * 2          ' �w���i�_�f�Z�x�␳�v�Z
    KPRONMNH As Double              ' �w���i�_�f�Z�x�����␳
    KPRONMXH As Double              ' �w���i�_�f�Z�x����␳
    KPROKBSH As String * 1          ' �w���i�_�f�U�敪����ʒu�Q��
    KPROKBST As String * 1          ' �w���i�_�f�U�敪����ʒu�Q�_
    KPROKBSI As String * 1          ' �w���i�_�f�U�敪����ʒu�Q��
    KPROKBHT As String * 1          ' �w���i�_�f�U�敪�ۏؕ��@�Q��
    KPROKBHS As String * 1          ' �w���i�_�f�U�敪�ۏؕ��@�Q��
    KPROS1MN As Double              ' �w���i�_�f�͏o�P����
    KPROS1MX As Double              ' �w���i�_�f�͏o�P���
    KPROS1NS As String * 2          ' �w���i�_�f�͏o�P�M�����@
    KPROS1SH As String * 1          ' �w���i�_�f�͏o�P����ʒu�Q��
    KPROS1ST As String * 1          ' �w���i�_�f�͏o�P����ʒu�Q�_
    KPROS1SI As String * 1          ' �w���i�_�f�͏o�P����ʒu�Q��
    KPROS1HT As String * 1          ' �w���i�_�f�͏o�P�ۏؕ��@�Q��
    KPROS1HS As String * 1          ' �w���i�_�f�͏o�P�ۏؕ��@�Q��
    KPROS1HM As String * 1          ' �w���i�_�f�͏o�P�����p�x�Q��
    KPROS1KN As String * 1          ' �w���i�_�f�͏o�P�����p�x�Q��
    KPROS1KH As String * 1          ' �w���i�_�f�͏o�P�����p�x�Q��
    KPROS1KU As String * 1          ' �w���i�_�f�͏o�P�����p�x�Q�E
    KPROS2MN As Double              ' �w���i�_�f�͏o�Q����
    KPROS2MX As Double              ' �w���i�_�f�͏o�Q���
    KPROS2NS As String * 2          ' �w���i�_�f�͏o�Q�M�����@
    KPROS2SH As String * 1          ' �w���i�_�f�͏o�Q����ʒu�Q��
    KPROS2ST As String * 1          ' �w���i�_�f�͏o�Q����ʒu�Q�_
    KPROS2SI As String * 1          ' �w���i�_�f�͏o�Q����ʒu�Q��
    KPROS2HT As String * 1          ' �w���i�_�f�͏o�Q�ۏؕ��@�Q��
    KPROS2HS As String * 1          ' �w���i�_�f�͏o�Q�ۏؕ��@�Q��
    KPROS2KM As String * 1          ' �w���i�_�f�͏o�Q�����p�x�Q��
    KPROS2KN As String * 1          ' �w���i�_�f�͏o�Q�����p�x�Q��
    KPROS2KH As String * 1          ' �w���i�_�f�͏o�Q�����p�x�Q��
    KPROS2KU As String * 1          ' �w���i�_�f�͏o�Q�����p�x�Q�E
    KPROS3MN As Double              ' �w���i�_�f�͏o�R����
    KPROS3MX As Double              ' �w���i�_�f�͏o�R���
    KPROS3NS As String * 2          ' �w���i�_�f�͏o�R�M�����@
    KPROS3SH As String * 1          ' �w���i�_�f�͏o�R����ʒu�Q��
    KPROS3ST As String * 1          ' �w���i�_�f�͏o�R����ʒu�Q�_
    KPROS3SI As String * 1          ' �w���i�_�f�͏o�R����ʒu�Q��
    KPROS3HT As String * 1          ' �w���i�_�f�͏o�R�ۏؕ��@�Q��
    KPROS3HS As String * 1          ' �w���i�_�f�͏o�R�ۏؕ��@�Q��
    KPROS3KM As String * 1          ' �w���i�_�f�͏o�R�����p�x�Q��
    KPROS3KN As String * 1          ' �w���i�_�f�͏o�R�����p�x�Q��
    KPROS3KH As String * 1          ' �w���i�_�f�͏o�R�����p�x�Q��
    KPROS3KU As String * 1          ' �w���i�_�f�͏o�R�����p�x�Q�E
    KPRANTNP As Integer             ' �w���i�`�m���x
    KPRANTIM As Integer             ' �w���i�`�m����
    KPRANTMN As Integer             ' �w���i�`�m���ԉ���
    KPRANTMX As Integer             ' �w���i�`�m���ԏ��
    KPRZOMIN As Double              ' �w���i�c���_�f����
    KPRZOMAX As Double              ' �w���i�c���_�f���
    KPRZOSPH As String * 1          ' �w���i�c���_�f����ʒu�Q��
    KPRZOSPT As String * 1          ' �w���i�c���_�f����ʒu�Q�_
    KPRZOSPI As String * 1          ' �w���i�c���_�f����ʒu�Q��
    KPRZOHWT As String * 1          ' �w���i�c���_�f�ۏؕ��@�Q��
    KPRZOHWS As String * 1          ' �w���i�c���_�f�ۏؕ��@�Q��
    KPRZONSW As String * 2          ' �w���i�c���_�f�M�����@
    KPRZOKWY As String * 2          ' �w���i�c���_�f�������@
    KPRZOKHM As String * 1          ' �w���i�c���_�f�����p�x�Q��
    KPRZOKHN As String * 1          ' �w���i�c���_�f�����p�x�Q��
    KPRZOKHH As String * 1          ' �w���i�c���_�f�����p�x�Q��
    KPRZOKHU As String * 1          ' �w���i�c���_�f�����p�x�Q�E
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lWF�ް��U
Public Type typ_TBCME013
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KPRBDOMN As Integer             ' �w���i�a�c�n�r�e����
    KPRBDOMX As Integer             ' �w���i�a�c�n�r�e���
    KPRBDOSH As String * 1          ' �w���i�a�c�n�r�e����ʒu�Q��
    KPRBDOST As String * 1          ' �w���i�a�c�n�r�e����ʒu�Q�_
    KPRBDOSR As String * 1          ' �w���i�a�c�n�r�e����ʒu�Q��
    KPRBDOHT As String * 1          ' �w���i�a�c�n�r�e�ۏؕ��@�Q��
    KPRBDOHS As String * 1          ' �w���i�a�c�n�r�e�ۏؕ��@�Q��
    KPRBDOSZ As String * 1          ' �w���i�a�c�n�r�e�������
    KPRBDONS As String * 2          ' �w���i�a�c�n�r�e�M�����@
    KPRBDOKM As String * 1          ' �w���i�a�c�n�r�e�����p�x�Q��
    KPRBDOKN As String * 1          ' �w���i�a�c�n�r�e�����p�x�Q��
    KPRBDOKH As String * 1          ' �w���i�a�c�n�r�e�����p�x�Q��
    KPRBDOKU As String * 1          ' �w���i�a�c�n�r�e�����p�x�Q�E
    KPRBDOET As Integer             ' �w���i�a�c�n�r�e�I���d�s��
    KPRBDSMN As Integer             ' �w���i�a�c�r�s�Չ���
    KPRBDSMX As Integer             ' �w���i�a�c�r�s�Տ��
    KPRBDSSH As String * 1          ' �w���i�a�c�r�s�Ց���ʒu�Q��
    KPRBDSST As String * 1          ' �w���i�a�c�r�s�Ց���ʒu�Q�_
    KPRBDSSR As String * 1          ' �w���i�a�c�r�s�Ց���ʒu�Q��
    KPRBDSHT As String * 1          ' �w���i�a�c�r�s�Օۏؕ��@�Q��
    KPRBDSHS As String * 1          ' �w���i�a�c�r�s�Օۏؕ��@�Q��
    KPRBDSSZ As String * 1          ' �w���i�a�c�r�s�Ց������
    KPRBDSKM As String * 1          ' �w���i�a�c�r�s�Ռ����p�x�Q��
    KPRBDSKN As String * 1          ' �w���i�a�c�r�s�Ռ����p�x�Q��
    KPRBDSKH As String * 1          ' �w���i�a�c�r�s�Ռ����p�x�Q��
    KPRBDSKU As String * 1          ' �w���i�a�c�r�s�Ռ����p�x�Q�E
    KPRBDSET As Integer             ' �w���i�a�c�r�s�ՑI���d�s��
    KPRRNFMX As Double              ' �w���i���t�l�X�\���
    KPRRNFKB As String * 1          ' �w���i���t�l�X�\�����敪
    KPRRNFKW As String * 2          ' �w���i���t�l�X�\�������@
    KPRRNFZA As Integer             ' �w���i���t�l�X�\���O�̈�
    KPRRNBMX As Double              ' �w���i���t�l�X�����
    KPRRNBKB As String * 1          ' �w���i���t�l�X�������敪
    KPRRNBKW As String * 2          ' �w���i���t�l�X���������@
    KPRRNBZA As Integer             ' �w���i���t�l�X�����O�̈�
    KPRDENKU As String * 1          ' �w���i�c���������L��
    KPRDENMX As Integer             ' �w���i�c�������
    KPRDENMN As Integer             ' �w���i�c��������
    KPRDENHT As String * 1          ' �w���i�c�����ۏؕ��@�Q��
    KPRDENHS As String * 1          ' �w���i�c�����ۏؕ��@�Q��
    KPRLDLKU As String * 1          ' �w���i�k�^�c�k�����L��
    KPRLDLMX As Integer             ' �w���i�k�^�c�k���
    KPRLDLMN As Integer             ' �w���i�k�^�c�k����
    KPRLDLHT As String * 1          ' �w���i�k�^�c�k�ۏؕ��@�Q��
    KPRLDLHS As String * 1          ' �w���i�k�^�c�k�ۏؕ��@�Q��
    KPRDVDKU As String * 1          ' �w���i�c�u�c�Q�����L��
    KPRDVDMX As Integer             ' �w���i�c�u�c�Q���
    KPRDVDMN As Integer             ' �w���i�c�u�c�Q����
    KPRDVDHT As String * 1          ' �w���i�c�u�c�Q�ۏؕ��@�Q��
    KPRDVDHS As String * 1          ' �w���i�c�u�c�Q�ۏؕ��@�Q��
    KPRGDSPH As String * 1          ' �w���i�f�c����ʒu�Q��
    KPRGDSPT As String * 1          ' �w���i�f�c����ʒu�Q�_
    KPRGDSPR As String * 1          ' �w���i�f�c����ʒu�Q��
    KPRGDSZY As String * 1          ' �w���i�f�c�������
    KPRGDZAR As Integer             ' �w���i�f�c���O�̈�
    KPRGDKHM As String * 1          ' �w���i�f�c�����p�x�Q��
    KPRGDKHN As String * 1          ' �w���i�f�c�����p�x�Q��
    KPRGDKHH As String * 1          ' �w���i�f�c�����p�x�Q��
    KPRGDKHU As String * 1          ' �w���i�f�c�����p�x�Q�E
    KPRDSOKE As String * 1          ' �w���i�c�r�n�c����
    KPRDSOMX As Long                ' �w���i�c�r�n�c���
    KPRDSOMN As Long                ' �w���i�c�r�n�c����
    KPRDSOAX As Integer             ' �w���i�c�r�n�c�̈���
    KPRDSOAN As Integer             ' �w���i�c�r�n�c�̈扺��
    KPRDSOHT As String * 1          ' �w���i�c�r�n�c�ۏؕ��@�Q��
    KPRDSOHS As String * 1          ' �w���i�c�r�n�c�ۏؕ��@�Q��
    KPRDSOKM As String * 1          ' �w���i�c�r�n�c�����p�x�Q��
    KPRDSOKN As String * 1          ' �w���i�c�r�n�c�����p�x�Q��
    KPRDSOKH As String * 1          ' �w���i�c�r�n�c�����p�x�Q��
    KPRDSOKU As String * 1          ' �w���i�c�r�n�c�����p�x�Q�E
    KPRNTPUM As String * 1          ' �w���i���R�i�m�g�|�L��
    KPRNTPK1 As Double              ' �w���i���R�i�m�g�|�K�i�P
    KPRNTPP1 As Double              ' �w���i���R�i�m�g�|�o�t�`�P
    KPRNTPS1 As Double              ' �w���i���R�i�m�g�|�T�C�g�P
    KPRNTPK2 As Double              ' �w���i���R�i�m�g�|�K�i�Q
    KPRNTPP2 As Double              ' �w���i���R�i�m�g�|�o�t�`�Q
    KPRNTPS2 As Double              ' �w���i���R�i�m�g�|�T�C�g�Q
    KPRNTPK3 As Double              ' �w���i���R�i�m�g�|�K�i�R
    KPRNTPP3 As Double              ' �w���i���R�i�m�g�|�o�t�`�R
    KPRNTPS3 As Double              ' �w���i���R�i�m�g�|�T�C�g�R
    KPRNTPZA As Integer             ' �w���i���R�i�m�g�|���O�̈�
    KPRNTPHT As String * 1          ' �w���i���R�i�m�g�|�ۏؕ��@�Q��
    KPRNTPHS As String * 1          ' �w���i���R�i�m�g�|�ۏؕ��@�Q��
    KPRNTPKM As String * 1          ' �w���i���R�i�m�g�|�����p�x�Q��
    KPRNTPKN As String * 1          ' �w���i���R�i�m�g�|�����p�x�Q��
    KPRNTPKH As String * 1          ' �w���i���R�i�m�g�|�����p�x�Q��
    KPRNTPKU As String * 1          ' �w���i���R�i�m�g�|�����p�x�Q�E
    KPRCRSSK As String * 1          ' �w���i���R�N���X�r�r����
    KPRMDCEN As Double              ' �w���i���R�ʃ_�����፷���S
    KPRMDMAX As Double              ' �w���i���R�ʃ_�����፷���
    KPRMDMIN As Double              ' �w���i���R�ʃ_�����፷����
    KPRMDSPH As String * 3          ' �w���i���R�ʃ_������ʒu�Q��
    KPRMDSPT As String * 3          ' �w���i���R�ʃ_������ʒu�Q�_
    KPRMDSPI As String * 3          ' �w���i���R�ʃ_������ʒu�Q��
    KPRMDHWT As String * 2          ' �w���i���R�ʃ_���ۏؕ��@�Q��
    KPRMDHWS As String * 2          ' �w���i���R�ʃ_���ۏؕ��@�Q��
    KPRMDKHM As String * 4          ' �w���i���R�ʃ_�������p�x�Q��
    KPRMDKHN As String * 4          ' �w���i���R�ʃ_�������p�x�Q��
    KPRMDKHH As String * 4          ' �w���i���R�ʃ_�������p�x�Q��
    KPRMDKHU As String * 4          ' �w���i���R�ʃ_�������p�x�Q�E
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lWF�ް��V
Public Type typ_TBCME014
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KPRSMIN As Double               ' �w���i���艺��
    KPRSMAX As Double               ' �w���i������
    KPRSHWYT As String * 1          ' �w���i����ۏؕ��@�Q��
    KPRSHWYS As String * 1          ' �w���i����ۏؕ��@�Q��
    KPRSKKBN As String * 1          ' �w���i���茟���敪
    KPRSKWAY As String * 2          ' �w���i���茟�����@
    KPRSSZYO As String * 1          ' �w���i���葪�����
    KPRSZARA As Integer             ' �w���i���菜�O�̈�
    KPRSSDEV As Double              ' �w���i����W���΍�
    KPRSAMIN As Double              ' �w���i���蕽�ω���
    KPRSAMAX As Double              ' �w���i���蕽�Ϗ��
    KPRSSREC As String * 1          ' �w���i���葪���
    KPRSBO1 As Double               ' �w���i���苫�E�P
    KPRSBO1B As Integer             ' �w���i���苫�E�P��
    KPRSBO2 As Double               ' �w���i���苫�E�Q
    KPRSBO2B As Integer             ' �w���i���苫�E�Q��
    KPRSBO3 As Double               ' �w���i���苫�E�R
    KPRSBO3B As Integer             ' �w���i���苫�E�R��
    KPRWARMX As Double              ' �w���i�v�`�q�o���
    KPRWARSZ As String * 1          ' �w���i�v�`�q�o�������
    KPRWARHT As String * 1          ' �w���i�v�`�q�o�ۏؕ��@�Q��
    KPRWARHS As String * 1          ' �w���i�v�`�q�o�ۏؕ��@�Q��
    KPRWARKB As String * 1          ' �w���i�v�`�q�o�����敪
    KPRWARKW As String * 2          ' �w���i�v�`�q�o�������@
    KPRWARZA As Integer             ' �w���i�v�`�q�o���O�̈�
    KPRWARSR As String * 1          ' �w���i�v�`�q�o�����
    KPRWAB1 As Double               ' �w���i�v�`�q�o���E�P
    KPRWAB1B As Integer             ' �w���i�v�`�q�o���E�P��
    KPRWAB2 As Double               ' �w���i�v�`�q�o���E�Q
    KPRWAB2B As Integer             ' �w���i�v�`�q�o���E�Q��
    KPRWAB3 As Double               ' �w���i�v�`�q�o���E�R
    KPRWAB3B As Integer             ' �w���i�v�`�q�o���E�R��
    KPRFKKBN As String * 1          ' �w���i���R�����敪
    KPRFSZYO As String * 1          ' �w���i���R�������
    KPRFSREC As String * 1          ' �w���i���R�����
    KPRGBMAX As Double              ' �w���i���R�f�a���
    KPRGBPUG As Double              ' �w���i���R�f�a�o�t�`��
    KPRGBPUR As Integer             ' �w���i���R�f�a�o�t�`��
    KPRGBHWT As String * 1          ' �w���i���R�f�a�ۏؕ��@�Q��
    KPRGBHWS As String * 1          ' �w���i���R�f�a�ۏؕ��@�Q��
    KPRGBKW As String * 4           ' �w���i���R�f�a�������@
    KPRGBKWO As String * 4          ' �w���i���R�f�a�������@��
    KPRGBZAR As Integer             ' �w���i���R�f�a���O�̈�
    KPRGBB1 As Double               ' �w���i���R�f�a���E�P
    KPRGBB1B As Integer             ' �w���i���R�f�a���E�P��
    KPRGBB2 As Double               ' �w���i���R�f�a���E�Q
    KPRGBB2B As Integer             ' �w���i���R�f�a���E�Q��
    KPRGBB3 As Double               ' �w���i���R�f�a���E�R
    KPRGBB3B As Integer             ' �w���i���R�f�a���E�R��
    KPRGFDMX As Double              ' �w���i���R�f�e�c���
    KPRGFDPG As Double              ' �w���i���R�f�e�c�o�t�`��
    KPRGFDPR As Integer             ' �w���i���R�f�e�c�o�t�`��
    KPRGFDHT As String * 1          ' �w���i���R�f�e�c�ۏؕ��@�Q��
    KPRGFDHS As String * 1          ' �w���i���R�f�e�c�ۏؕ��@�Q��
    KPRGFDBM As String * 1          ' �w���i���R�f�e�c���
    KPRGFDKW As String * 4          ' �w���i���R�f�e�c�������@
    KPRGFDKO As String * 4          ' �w���i���R�f�e�c�������@��
    KPRGFDZA As Integer             ' �w���i���R�f�e�c���O�̈�
    KPRGDB1 As Double               ' �w���i���R�f�e�c���E�P
    KPRGDB1B As Integer             ' �w���i���R�f�e�c���E�P��
    KPRGDB2 As Double               ' �w���i���R�f�e�c���E�Q
    KPRGDB2B As Integer             ' �w���i���R�f�e�c���E�Q��
    KPRGDB3 As Double               ' �w���i���R�f�e�c���E�R
    KPRGDB3B As Integer             ' �w���i���R�f�e�c���E�R��
    KPRGFRMX As Double              ' �w���i���R�f�e�q���
    KPRGFRPG As Double              ' �w���i���R�f�e�q�o�t�`��
    KPRGFRPR As Integer             ' �w���i���R�f�e�q�o�t�`��
    KPRGFRHT As String * 1          ' �w���i���R�f�e�q�ۏؕ��@�Q��
    KPRGFRHS As String * 1          ' �w���i���R�f�e�q�ۏؕ��@�Q��
    KPRGFRBM As String * 1          ' �w���i���R�f�e�q���
    KPRGFRKW As String * 4          ' �w���i���R�f�e�q�������@
    KPRGFRKO As String * 4          ' �w���i���R�f�e�q�������@��
    KPRGFRZA As Integer             ' �w���i���R�f�e�q���O�̈�
    KPRGRB1 As Double               ' �w���i���R�f�e�q���E�P
    KPRGRB1B As Integer             ' �w���i���R�f�e�q���E�P��
    KPRGRB2 As Double               ' �w���i���R�f�e�q���E�Q
    KPRGRB2B As Integer             ' �w���i���R�f�e�q���E�Q��
    KPRGRB3 As Double               ' �w���i���R�f�e�q���E�R
    KPRGRB3B As Integer             ' �w���i���R�f�e�q���E�R��
    KPRSBMAX As Double              ' �w���i���R�r�a���
    KPRSBPUG As Double              ' �w���i���R�r�a�o�t�`��
    KPRSBPUR As Integer             ' �w���i���R�r�a�o�t�`��
    KPRSBSZX As Double              ' �w���i���R�r�a�T�C�Y�w
    KPRSBSZY As Double              ' �w���i���R�r�a�T�C�Y�x
    KPRSBHWT As String * 1          ' �w���i���R�r�a�ۏؕ��@�Q��
    KPRSBHWS As String * 1          ' �w���i���R�r�a�ۏؕ��@�Q��
    KPRSBBM As String * 1           ' �w���i���R�r�a���
    KPRSBKW As String * 4           ' �w���i���R�r�a�������@
    KPRSBKWO As String * 4          ' �w���i���R�r�a�������@��
    KPRSBZAR As Integer             ' �w���i���R�r�a���O�̈�
    KPRSBB1 As Double               ' �w���i���R�r�a���E�P
    KPRSBB1B As Integer             ' �w���i���R�r�a���E�P��
    KPRSBB2 As Double               ' �w���i���R�r�a���E�Q
    KPRSBB2B As Integer             ' �w���i���R�r�a���E�Q��
    KPRSBB3 As Double               ' �w���i���R�r�a���E�R
    KPRSBB3B As Integer             ' �w���i���R�r�a���E�R��
    KPRSFMAX As Double              ' �w���i���R�r�e���
    KPRSFPUG As Double              ' �w���i���R�r�e�o�t�`��
    KPRSFPUR As Integer             ' �w���i���R�r�e�o�t�`��
    KPRSFSZX As Double              ' �w���i���R�r�e�T�C�Y�w
    KPRSFSZY As Double              ' �w���i���R�r�e�T�C�Y�x
    KPRSFHWT As String * 1          ' �w���i���R�r�e�ۏؕ��@�Q��
    KPRSFHWS As String * 1          ' �w���i���R�r�e�ۏؕ��@�Q��
    KPRSFBM As String * 1           ' �w���i���R�r�e���
    KPRSFKW As String * 4           ' �w���i���R�r�e�������@
    KPRSFKWO As String * 4          ' �w���i���R�r�e�������@��
    KPRSFZAR As Integer             ' �w���i���R�r�e���O�̈�
    KPRSFB1 As Double               ' �w���i���R�r�e���E�P
    KPRSFB1B As Integer             ' �w���i���R�r�e���E�P��
    KPRSFB2 As Double               ' �w���i���R�r�e���E�Q
    KPRSFB2B As Integer             ' �w���i���R�r�e���E�Q��
    KPRSFB3 As Double               ' �w���i���R�r�e���E�R
    KPRSFB3B As Integer             ' �w���i���R�r�e���E�R��
    KPRFSXOF As Double              ' �w���i���R�T�C�g�w�n�e
    KPRFSYOF As Double              ' �w���i���R�T�C�g�x�n�e
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lWF�ް��W
Public Type typ_TBCME015
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KPRMK1SI As Double              ' �w���i�ʌ����ׂP�T�C�Y
    KPRMK1MX As Integer             ' �w���i�ʌ����ׂP���
    KPRMK1SZ As String * 1          ' �w���i�ʌ����ׂP�������
    KPRMK1ZA As Integer             ' �w���i�ʌ����ׂP���O�̈�
    KPRMK1HT As String * 1          ' �w���i�ʌ����ׂP�ۏؕ��@�Q��
    KPRMK1HS As String * 1          ' �w���i�ʌ����ׂP�ۏؕ��@�Q��
    KPRM1B1 As Integer              ' �w���i�ʌ����ׂP���E�P
    KPRM1B1B As Integer             ' �w���i�ʌ����ׂP���E�P��
    KPRM1B2 As Integer              ' �w���i�ʌ����ׂP���E�Q
    KPRM1B2B As Integer             ' �w���i�ʌ����ׂP���E�Q��
    KPRM1B3 As Integer              ' �w���i�ʌ����ׂP���E�R
    KPRM1B3B As Integer             ' �w���i�ʌ����ׂP���E�R��
    KPRMK2SI As Double              ' �w���i�ʌ����ׂQ�T�C�Y
    KPRMK2MX As Integer             ' �w���i�ʌ����ׂQ���
    KPRMK2HT As String * 1          ' �w���i�ʌ����ׂQ�ۏؕ��@�Q��
    KPRMK2HS As String * 1          ' �w���i�ʌ����ׂQ�ۏؕ��@�Q��
    KPRM2B1 As Integer              ' �w���i�ʌ����ׂQ���E�P
    KPRM2B1B As Integer             ' �w���i�ʌ����ׂQ���E�P��
    KPRM2B2 As Integer              ' �w���i�ʌ����ׂQ���E�Q
    KPRM2B2B As Integer             ' �w���i�ʌ����ׂQ���E�Q��
    KPRM2B3 As Integer              ' �w���i�ʌ����ׂQ���E�R
    KPRM2B3B As Integer             ' �w���i�ʌ����ׂQ���E�R��
    KPRMK3SI As Double              ' �w���i�ʌ����ׂR�T�C�Y
    KPRMK3MX As Integer             ' �w���i�ʌ����ׂR���
    KPRMK3HT As String * 1          ' �w���i�ʌ����ׂR�ۏؕ��@�Q��
    KPRMK3HS As String * 1          ' �w���i�ʌ����ׂR�ۏؕ��@�Q��
    KPRM3B1 As Integer              ' �w���i�ʌ����ׂR���E�P
    KPRM3B1B As Integer             ' �w���i�ʌ����ׂR���E�P��
    KPRM3B2 As Integer              ' �w���i�ʌ����ׂR���E�Q
    KPRM3B2B As Integer             ' �w���i�ʌ����ׂR���E�Q��
    KPRM3B3 As Integer              ' �w���i�ʌ����ׂR���E�R
    KPRM3B3B As Integer             ' �w���i�ʌ����ׂR���E�R��
    KPRMK4SI As Double              ' �w���i�ʌ����ׂS�T�C�Y
    KPRMK4MX As Integer             ' �w���i�ʌ����ׂS���
    KPRMK4HT As String * 1          ' �w���i�ʌ����ׂS�ۏؕ��@�Q��
    KPRMK4HS As String * 1          ' �w���i�ʌ����ׂS�ۏؕ��@�Q��
    KPRM4B1 As Integer              ' �w���i�ʌ����ׂS���E�P
    KPRM4B1B As Integer             ' �w���i�ʌ����ׂS���E�P��
    KPRM4B2 As Integer              ' �w���i�ʌ����ׂS���E�Q
    KPRM4B2B As Integer             ' �w���i�ʌ����ׂS���E�Q��
    KPRM4B3 As Integer              ' �w���i�ʌ����ׂS���E�R
    KPRM4B3B As Integer             ' �w���i�ʌ����ׂS���E�R��
    KPRMB1SI As Double              ' �w���i�ʌ����ח��P�T�C�Y
    KPRMB1MX As Integer             ' �w���i�ʌ����ח��P���
    KPRMB1SZ As String * 1          ' �w���i�ʌ����ח��P�������
    KPRMB1ZA As Integer             ' �w���i�ʌ����ח��P���O�̈�
    KPRMB1HT As String * 1          ' �w���i�ʌ����ח��P�ۏؕ��@�Q��
    KPRMB1HS As String * 1          ' �w���i�ʌ����ח��P�ۏؕ��@�Q��
    KPRMB2SI As Double              ' �w���i�ʌ����ח��Q�T�C�Y
    KPRMB2MX As Integer             ' �w���i�ʌ����ח��Q���
    KPRMB2SZ As String * 1          ' �w���i�ʌ����ח��Q�������
    KPRMB2ZA As Integer             ' �w���i�ʌ����ח��Q���O�̈�
    KPRMB2HT As String * 1          ' �w���i�ʌ����ח��Q�ۏؕ��@�Q��
    KPRMB2HS As String * 1          ' �w���i�ʌ����ח��Q�ۏؕ��@�Q��
    KPRMKSRE As String * 1          ' �w���i�ʌ����ב����
    KPRMPIPT As String * 1          ' �w���i�ʌ����ׂo�h�o����
    KPRMPIPK As Integer             ' �w���i�ʌ����ׂo�h�o��
    KPRMPISH As String * 1          ' �w���i�ʌ��o�h�o����ʒu�Q��
    KPRMPIST As String * 1          ' �w���i�ʌ��o�h�o����ʒu�Q�_
    KPRMPISI As String * 1          ' �w���i�ʌ��o�h�o����ʒu�Q��
    KPRMPIKM As String * 1          ' �w���i�ʌ��o�h�o�����p�x�Q��
    KPRMPIKN As String * 1          ' �w���i�ʌ��o�h�o�����p�x�Q��
    KPRMPIKH As String * 1          ' �w���i�ʌ��o�h�o�����p�x�Q��
    KPRMPIKU As String * 1          ' �w���i�ʌ��o�h�o�����p�x�Q�E
    KPRMNIND As String * 2          ' �w���i�����Z�x�w��
    KPRMNMAX As Double              ' �w���i�����Z�x���
    KPRMNALX As Double              ' �w���i�����Z�x�`�k���
    KPRMNCAX As Double              ' �w���i�����Z�x�b�`���
    KPRMNCRX As Double              ' �w���i�����Z�x�b�q���
    KPRMNCUX As Double              ' �w���i�����Z�x�b�t���
    KPRMNFEX As Double              ' �w���i�����Z�x�e�d���
    KPRMNKMX As Double              ' �w���i�����Z�x�j���
    KPRMNMGX As Double              ' �w���i�����Z�x�l�f���
    KPRMNNAX As Double              ' �w���i�����Z�x�m�`���
    KPRMNNIX As Double              ' �w���i�����Z�x�m�h���
    KPRMNZNX As Double              ' �w���i�����Z�x�y�m���
    KPRMNKWY As String * 2          ' �w���i�����Z�x�������@
    KPRMNZAR As Integer             ' �w���i�����Z�x���O�̈�
    KPRMNKHM As String * 1          ' �w���i�����Z�x�����p�x�Q��
    KPRMNKHN As String * 1          ' �w���i�����Z�x�����p�x�Q��
    KPRMNKHH As String * 1          ' �w���i�����Z�x�����p�x�Q��
    KPRMNKHU As String * 1          ' �w���i�����Z�x�����p�x�Q�E
    KPRSPVMX As Double              ' �w���i�r�o�u�e�d���
    KPRSPVKM As String * 1          ' �w���i�r�o�u�e�d�����p�x�Q��
    KPRSPVKN As String * 1          ' �w���i�r�o�u�e�d�����p�x�Q��
    KPRSPVKH As String * 1          ' �w���i�r�o�u�e�d�����p�x�Q��
    KPRSPVKU As String * 1          ' �w���i�r�o�u�e�d�����p�x�Q�E
    KPRSPVIN As String * 2          ' �w���i�r�o�u�e�d�w��
    KPRDLMIN As Integer             ' �w���i�g�U������
    KPRDLMAX As Integer             ' �w���i�g�U�����
    KPRDLSPH As String * 1          ' �w���i�g�U������ʒu�Q��
    KPRDLSPT As String * 1          ' �w���i�g�U������ʒu�Q�_
    KPRDLSPI As String * 1          ' �w���i�g�U������ʒu�Q��
    KPRDLHWT As String * 1          ' �w���i�g�U���ۏؕ��@�Q��
    KPRDLHWS As String * 1          ' �w���i�g�U���ۏؕ��@�Q��
    KPRDLKHM As String * 1          ' �w���i�g�U�������p�x�Q��
    KPRDLKHN As String * 1          ' �w���i�g�U�������p�x�Q��
    KPRDLKHH As String * 1          ' �w���i�g�U�������p�x�Q��
    KPRDLKHU As String * 1          ' �w���i�g�U�������p�x�Q�E
    KPROTMIN As Double              ' �w���i�_�����ψ�����
    KPROTSPH As String * 1          ' �w���i�_�����ψ�����ʒu�Q��
    KPROTSPT As String * 1          ' �w���i�_�����ψ�����ʒu�Q�_
    KPROTSPI As String * 1          ' �w���i�_�����ψ�����ʒu�Q��
    KPROTKWY As String * 2          ' �w���i�_�����ψ��������@
    KPROTZAR As Integer             ' �w���i�_�����ψ����O�̈�
    KPROTKHM As String * 1          ' �w���i�_�����ψ������p�x�Q��
    KPROTKHN As String * 1          ' �w���i�_�����ψ������p�x�Q��
    KPROTKHH As String * 1          ' �w���i�_�����ψ������p�x�Q��
    KPROTKHU As String * 1          ' �w���i�_�����ψ������p�x�Q�E
    KPROTMX1 As Double              ' �w���i�_�����ψ�����P
    KPROTMX2 As Double              ' �w���i�_�����ψ�����Q
    KPROTKW1 As String * 2          ' �w���i�_�����ψ��������@�P
    KPROTKW2 As String * 2          ' �w���i�_�����ψ��������@�Q
    KPROTHWT As String * 1          ' �w���i�_�����ψ��ۏؕ��@�Q��
    KPROTHWS As String * 1          ' �w���i�_�����ψ��ۏؕ��@�Q��
    KPRLTDCX As Double              ' �w���i�k�s�c�Z�x�b�t���
    KPRLTDIN As String * 2          ' �w���i�k�s�c�Z�x�w��
    KPRLTDKW As String * 2          ' �w���i�k�s�c�Z�x�������@
    KPRLTDSH As String * 1          ' �w���i�k�s�c�Z�x����ʒu�Q��
    KPRLTDST As String * 1          ' �w���i�k�s�c�Z�x����ʒu�Q�_
    KPRLTDSI As String * 1          ' �w���i�k�s�c�Z�x����ʒu�Q��
    KPRLTDHT As String * 1          ' �w���i�k�s�c�Z�x�ۏؕ��@�Q��
    KPRLTDHS As String * 1          ' �w���i�k�s�c�Z�x�ۏؕ��@�Q��
    KPRLTDKM As String * 1          ' �w���i�k�s�c�Z�x�����p�x�Q��
    KPRLTDKN As String * 1          ' �w���i�k�s�c�Z�x�����p�x�Q��
    KPRLTDKH As String * 1          ' �w���i�k�s�c�Z�x�����p�x�Q��
    KPRLTDKU As String * 1          ' �w���i�k�s�c�Z�x�����p�x�Q�E
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lWF�ް��X
Public Type typ_TBCME016
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KPROS1AX As Double              ' �w���i�n�r�e�P���Ϗ��
    KPROS1MX As Double              ' �w���i�n�r�e�P���
    KPROS1O1 As Integer             ' �w���i�n�r�e�P�������x�P
    KPROS1T1 As Integer             ' �w���i�n�r�e�P�������ԂP
    KPROS1GS As String * 1          ' �w���i�n�r�e�P���͋C�K�X
    KPROS1ET As Integer             ' �w���i�n�r�e�P�I���d�s��
    KPROS1NS As String * 2          ' �w���i�n�r�e�P�M�����@
    KPROS1SZ As String * 1          ' �w���i�n�r�e�P�������
    KPROS1SH As String * 1          ' �w���i�n�r�e�P����ʒu�Q��
    KPROS1ST As String * 1          ' �w���i�n�r�e�P����ʒu�Q�_
    KPROS1SR As String * 1          ' �w���i�n�r�e�P����ʒu�Q��
    KPROS1HT As String * 1          ' �w���i�n�r�e�P�ۏؕ��@�Q��
    KPROS1HS As String * 1          ' �w���i�n�r�e�P�ۏؕ��@�Q��
    KPROS1KB As String * 1          ' �w���i�n�r�e�P�����敪
    KPROS1KM As String * 1          ' �w���i�n�r�e�P�����p�x�Q��
    KPROS1KN As String * 1          ' �w���i�n�r�e�P�����p�x�Q��
    KPROS1KH As String * 1          ' �w���i�n�r�e�P�����p�x�Q��
    KPROS1KU As String * 1          ' �w���i�n�r�e�P�����p�x�Q�E
    KPROS2AX As Double              ' �w���i�n�r�e�Q���Ϗ��
    KPROS2MX As Double              ' �w���i�n�r�e�Q���
    KPROS2O1 As Integer             ' �w���i�n�r�e�Q�������x�P
    KPROS2T1 As Integer             ' �w���i�n�r�e�Q�������ԂP
    KPROS2GS As String * 1          ' �w���i�n�r�e�Q���͋C�K�X
    KPROS2ET As Integer             ' �w���i�n�r�e�Q�I���d�s��
    KPROS2NS As String * 2          ' �w���i�n�r�e�Q�M�����@
    KPROS2SZ As String * 1          ' �w���i�n�r�e�Q�������
    KPROS2SH As String * 1          ' �w���i�n�r�e�Q����ʒu�Q��
    KPROS2ST As String * 1          ' �w���i�n�r�e�Q����ʒu�Q�_
    KPROS2SR As String * 1          ' �w���i�n�r�e�Q����ʒu�Q��
    KPROS2HT As String * 1          ' �w���i�n�r�e�Q�ۏؕ��@�Q��
    KPROS2HS As String * 1          ' �w���i�n�r�e�Q�ۏؕ��@�Q��
    KPROS2KB As String * 1          ' �w���i�n�r�e�Q�����敪
    KPROS2KM As String * 1          ' �w���i�n�r�e�Q�����p�x�Q��
    KPROS2KN As String * 1          ' �w���i�n�r�e�Q�����p�x�Q��
    KPROS2KH As String * 1          ' �w���i�n�r�e�Q�����p�x�Q��
    KPROS2KU As String * 1          ' �w���i�n�r�e�Q�����p�x�Q�E
    KPROS3AX As Double              ' �w���i�n�r�e�R���Ϗ��
    KPROS3MX As Double              ' �w���i�n�r�e�R���
    KPROS3O1 As Integer             ' �w���i�n�r�e�R�������x�P
    KPROS3T1 As Integer             ' �w���i�n�r�e�R�������ԂP
    KPROS3GS As String * 1          ' �w���i�n�r�e�R���͋C�K�X
    KPROS3ET As Integer             ' �w���i�n�r�e�R�I���d�s��
    KPROS3NS As String * 2          ' �w���i�n�r�e�R�M�����@
    KPROS3SZ As String * 1          ' �w���i�n�r�e�R�������
    KPROS3SH As String * 1          ' �w���i�n�r�e�R����ʒu�Q��
    KPROS3ST As String * 1          ' �w���i�n�r�e�R����ʒu�Q�_
    KPROS3SR As String * 1          ' �w���i�n�r�e�R����ʒu�Q��
    KPROS3HT As String * 1          ' �w���i�n�r�e�R�ۏؕ��@�Q��
    KPROS3HS As String * 1          ' �w���i�n�r�e�R�ۏؕ��@�Q��
    KPROS3KB As String * 1          ' �w���i�n�r�e�R�����敪
    KPROS3KM As String * 1          ' �w���i�n�r�e�R�����p�x�Q��
    KPROS3KN As String * 1          ' �w���i�n�r�e�R�����p�x�Q��
    KPROS3KH As String * 1          ' �w���i�n�r�e�R�����p�x�Q��
    KPROS3KU As String * 1          ' �w���i�n�r�e�R�����p�x�Q�E
    KPROS4AX As Double              ' �w���i�n�r�e�S���Ϗ��
    KPROS4MX As Double              ' �w���i�n�r�e�S���
    KPROS4O1 As Integer             ' �w���i�n�r�e�S�������x�P
    KPROS4T1 As Integer             ' �w���i�n�r�e�S�������ԂP
    KPROS4GS As String * 1          ' �w���i�n�r�e�S���͋C�K�X
    KPROS4ET As Integer             ' �w���i�n�r�e�S�I���d�s��
    KPROS4NS As String * 2          ' �w���i�n�r�e�S�M�����@
    KPROS4SZ As String * 1          ' �w���i�n�r�e�S�������
    KPROS4SH As String * 1          ' �w���i�n�r�e�S����ʒu�Q��
    KPROS4ST As String * 1          ' �w���i�n�r�e�S����ʒu�Q�_
    KPROS4SR As String * 1          ' �w���i�n�r�e�S����ʒu�Q��
    KPROS4HT As String * 1          ' �w���i�n�r�e�S�ۏؕ��@�Q��
    KPROS4HS As String * 1          ' �w���i�n�r�e�S�ۏؕ��@�Q��
    KPROS4KB As String * 1          ' �w���i�n�r�e�S�����敪
    KPROS4KM As String * 1          ' �w���i�n�r�e�S�����p�x�Q��
    KPROS4KN As String * 1          ' �w���i�n�r�e�S�����p�x�Q��
    KPROS4KH As String * 1          ' �w���i�n�r�e�S�����p�x�Q��
    KPROS4KU As String * 1          ' �w���i�n�r�e�S�����p�x�Q�E
    KPRBM1AN As Double              ' �w���i�a�l�c�P���ω���
    KPRBM1AX As Double              ' �w���i�a�l�c�P���Ϗ��
    KPRBM1GS As String * 1          ' �w���i�a�l�c�P���͋C�K�X
    KPRBM1ET As Integer             ' �w���i�a�l�c�P�I���d�s��
    KPRBM1NS As String * 2          ' �w���i�a�l�c�P�M�����@
    KPRBM1SZ As String * 1          ' �w���i�a�l�c�P�������
    KPRBM1SH As String * 1          ' �w���i�a�l�c�P����ʒu�Q��
    KPRBM1ST As String * 1          ' �w���i�a�l�c�P����ʒu�Q�_
    KPRBM1SR As String * 1          ' �w���i�a�l�c�P����ʒu�Q��
    KPRBM1HT As String * 1          ' �w���i�a�l�c�P�ۏؕ��@�Q��
    KPRBM1HS As String * 1          ' �w���i�a�l�c�P�ۏؕ��@�Q��
    KPRBM1KB As String * 1          ' �w���i�a�l�c�P�����敪
    KPRBM1KM As String * 1          ' �w���i�a�l�c�P�����p�x�Q��
    KPRBM1KN As String * 1          ' �w���i�a�l�c�P�����p�x�Q��
    KPRBM1KH As String * 1          ' �w���i�a�l�c�P�����p�x�Q��
    KPRBM1KU As String * 1          ' �w���i�a�l�c�P�����p�x�Q�E
    KPRBM2AN As Double              ' �w���i�a�l�c�Q���ω���
    KPRBM2AX As Double              ' �w���i�a�l�c�Q���Ϗ��
    KPRBM2GS As String * 1          ' �w���i�a�l�c�Q���͋C�K�X
    KPRBM2ET As Integer             ' �w���i�a�l�c�Q�I���d�s��
    KPRBM2NS As String * 2          ' �w���i�a�l�c�Q�M�����@
    KPRBM2SZ As String * 1          ' �w���i�a�l�c�Q�������
    KPRBM2SH As String * 1          ' �w���i�a�l�c�Q����ʒu�Q��
    KPRBM2ST As String * 1          ' �w���i�a�l�c�Q����ʒu�Q�_
    KPRBM2SR As String * 1          ' �w���i�a�l�c�Q����ʒu�Q��
    KPRBM2HT As String * 1          ' �w���i�a�l�c�Q�ۏؕ��@�Q��
    KPRBM2HS As String * 1          ' �w���i�a�l�c�Q�ۏؕ��@�Q��
    KPRBM2KB As String * 1          ' �w���i�a�l�c�Q�����敪
    KPRBM2KM As String * 1          ' �w���i�a�l�c�Q�����p�x�Q��
    KPRBM2KN As String * 1          ' �w���i�a�l�c�Q�����p�x�Q��
    KPRBM2KH As String * 1          ' �w���i�a�l�c�Q�����p�x�Q��
    KPRBM2KU As String * 1          ' �w���i�a�l�c�Q�����p�x�Q�E
    KPRBM3AN As Double              ' �w���i�a�l�c�R���ω���
    KPRBM3AX As Double              ' �w���i�a�l�c�R���Ϗ��
    KPRBM3GS As String * 1          ' �w���i�a�l�c�R���͋C�K�X
    KPRBM3ET As Integer             ' �w���i�a�l�c�R�I���d�s��
    KPRBM3NS As String * 2          ' �w���i�a�l�c�R�M�����@
    KPRBM3SZ As String * 1          ' �w���i�a�l�c�R�������
    KPRBM3SH As String * 1          ' �w���i�a�l�c�R����ʒu�Q��
    KPRBM3ST As String * 1          ' �w���i�a�l�c�R����ʒu�Q�_
    KPRBM3SR As String * 1          ' �w���i�a�l�c�R����ʒu�Q��
    KPRBM3HT As String * 1          ' �w���i�a�l�c�R�ۏؕ��@�Q��
    KPRBM3HS As String * 1          ' �w���i�a�l�c�R�ۏؕ��@�Q��
    KPRBM3KB As String * 1          ' �w���i�a�l�c�R�����敪
    KPRBM3KM As String * 1          ' �w���i�a�l�c�R�����p�x�Q��
    KPRBM3KN As String * 1          ' �w���i�a�l�c�R�����p�x�Q��
    KPRBM3KH As String * 1          ' �w���i�a�l�c�R�����p�x�Q��
    KPRBM3KU As String * 1          ' �w���i�a�l�c�R�����p�x�Q�E
    KPRBMDVO As String * 1          ' �w���i�a�l�c�̐ϊ��Z�L��
    KPROSPAX As Integer             ' �w���i�n�r�o���Ϗ��
    KPROSPMX As Integer             ' �w���i�n�r�o���
    KPROSPSH As String * 1          ' �w���i�n�r�o����ʒu�Q��
    KPROSPST As String * 1          ' �w���i�n�r�o����ʒu�Q�_
    KPROSPSR As String * 1          ' �w���i�n�r�o����ʒu�Q��
    KPROSPHT As String * 1          ' �w���i�n�r�o�ۏؕ��@�Q��
    KPROSPHS As String * 1          ' �w���i�n�r�o�ۏؕ��@�Q��
    KPROSPNS As String * 2          ' �w���i�n�r�o�M�����@
    KPROSPSZ As String * 1          ' �w���i�n�r�o�������
    KPROSPKM As String * 1          ' �w���i�n�r�o�����p�x�Q��
    KPROSPKN As String * 1          ' �w���i�n�r�o�����p�x�Q��
    KPROSPKH As String * 1          ' �w���i�n�r�o�����p�x�Q��
    KPROSPKU As String * 1          ' �w���i�n�r�o�����p�x�Q�E
    KPROSPET As Integer             ' �w���i�n�r�o�I���d�s��
    KPRTSPHM As String * 1          ' �w���i�g���X�T���v���p�x�Q��
    KPRTSPHN As String * 1          ' �w���i�g���X�T���v���p�x�Q��
    KPRTSPHH As String * 1          ' �w���i�g���X�T���v���p�x�Q��
    KPRTSPHU As String * 1          ' �w���i�g���X�T���v���p�x�Q�E
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ���i�d�l�Ǘ�
Public Type typ_TBCME017
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    SHNAME As String * 11           ' �Г��i��
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGSXSNO As String * 6          ' �i�Ǘ��r�w���i�ԍ�
    HMGSXSNE As Integer             ' �i�Ǘ��r�w���i�ԍ��}��
    HMGWFSNO As String * 6          ' �i�Ǘ��v�e���i�ԍ�
    HMGWFSNE As Integer             ' �i�Ǘ��v�e���i�ԍ��}��
    HMGEPSNO As String * 6          ' �i�Ǘ��d�o���i�ԍ�
    HMGEPSNE As Integer             ' �i�Ǘ��d�o���i�ԍ��}��
    SPECMWAY As String * 1          ' �d�l�쐬���@
    UNIFLAG As String * 1           ' �����t���O
    CONFLAG As String * 1           ' �m�F�t���O
    REINFLAG As String * 1          ' �ĕt�^�t���O
    HMGRDIAM As Integer             ' �i�Ǘ���\���a
    HMGWKBN As String * 2           ' �i�Ǘ����@�敪
    HMGPMKBN As String * 1          ' �i�Ǘ��݌v�Ǘ��敪
    HSXSKBN As String * 1           ' �i�r�w���i�敪
    HSXNCKBN As String * 1          ' �i�r�w�m�b�`�敪
    HWFSKBN As String * 1           ' �i�v�e���i�敪
    HWFRKBNK As String * 1          ' �i�v�e���R�敪���
    HWFAKBUM As String * 1          ' �i�v�e���݋敪�L��
    HWFOXKBN As String * 1          ' �i�v�e�_�f�敪
    HWFIGKBN As String * 1          ' �i�v�e�h�f�敪
    HWFNCKBN As String * 1          ' �i�v�e�m�b�`�敪
    HWFCMPKU As String * 1          ' �i�v�e�b�l�o���H�L��
    HWFSZKBN As String * 1          ' �i�v�e�x���ޗ��敪
    HWFSZMUM As String * 1          ' �i�v�e�x���ޗ��ʎ�L��
    HWFKZKBN As String * 1          ' �i�v�e�w���ޗ��敪
    HWFHGRAD As String * 7          ' �i�v�e�i���O���[�h
    HEPSKBN As String * 1           ' �i�d�o���i�敪
    HEPRKBNU As String * 1          ' �i�d�o���R�敪�L��
    HEPAKBUM As String * 1          ' �i�d�o���݋敪�L��
    HEPSZKBN As String * 1          ' �i�d�o�x���ޗ��敪
    HEPKZKBN As String * 1          ' �i�d�o�w���ޗ��敪
    HMGTRKSI As String * 1          ' �i�Ǘ��s�q�j���w��
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ���i�d�lSXL�ް��P
Public Type typ_TBCME018
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGSXSNO As String * 6          ' �i�Ǘ��r�w���i�ԍ�
    HMGSXSNE As Integer             ' �i�Ǘ��r�w���i�ԍ��}��
    CONFLAG As String * 1           ' �m�F�t���O
    REINFLAG As String * 1          ' �ĕt�^�t���O
    HSXTRWKB As String * 1          ' �i�r�w�����ۋ敪
    HSXTYPE As String * 1           ' �i�r�w�^�C�v
    KSXTYPKW As String * 1          ' �i�r�w�^�C�v�������@
    HSXDOP As String * 1            ' �i�r�w�h�[�p���g
    HSXRMIN As Double               ' �i�r�w���R����
    HSXRMAX As Double               ' �i�r�w���R���
    HSXRSPOH As String * 1          ' �i�r�w���R����ʒu�Q��
    HSXRSPOT As String * 1          ' �i�r�w���R����ʒu�Q�_
    HSXRSPOI As String * 1          ' �i�r�w���R����ʒu�Q��
    HSXRHWYT As String * 1          ' �i�r�w���R�ۏؕ��@�Q��
    HSXRHWYS As String * 1          ' �i�r�w���R�ۏؕ��@�Q��
    HSXRKWAY As String * 2          ' �i�r�w���R�������@
    HSXRKHNM As String * 1          ' �i�r�w���R�����p�x�Q��
    HSXRKHNI As String * 1          ' �i�r�w���R�����p�x�Q��
    HSXRKHNH As String * 1          ' �i�r�w���R�����p�x�Q��
    HSXRKHNS As String * 1          ' �i�r�w���R�����p�x�Q��
    HSXRMCAL As String * 1          ' �i�r�w���R�ʓ��v�Z
    HSXRMBNP As Double              ' �i�r�w���R�ʓ����z
    HSXRMCL2 As String * 1          ' �i�r�w���R�ʓ��v�Z�Q
    HSXRMBP2 As Double              ' �i�r�w���R�ʓ����z�Q
    HSXRSDEV As Double              ' �i�r�w���R�W���΍�
    HSXRAMIN As Double              ' �i�r�w���R���ω���
    HSXRAMAX As Double              ' �i�r�w���R���Ϗ��
    HSXFORM As String * 1           ' �i�r�w�`��
    HSXD1CEN As Double              ' �i�r�w���a�P���S
    HSXD1MIN As Double              ' �i�r�w���a�P����
    HSXD1MAX As Double              ' �i�r�w���a�P���
    HSXD2CEN As Double              ' �i�r�w���a�Q���S
    HSXD2MIN As Double              ' �i�r�w���a�Q����
    HSXD2MAX As Double              ' �i�r�w���a�Q���
    HSXCDIR As String * 1           ' �i�r�w�����ʕ���
    HSXCSCEN As Double              ' �i�r�w�����ʌX���S
    HSXCSMIN As Double              ' �i�r�w�����ʌX����
    HSXCSMAX As Double              ' �i�r�w�����ʌX���
    HSXCKWAY As String * 2          ' �i�r�w�����ʌ������@
    HSXCKHNM As String * 1          ' �i�r�w�����ʌ����p�x�Q��
    HSXCKHNI As String * 1          ' �i�r�w�����ʌ����p�x�Q��
    HSXCKHNH As String * 1          ' �i�r�w�����ʌ����p�x�Q��
    HSXCKHNS As String * 1          ' �i�r�w�����ʌ����p�x�Q��
    HSXCSDIR As String * 2          ' �i�r�w�����ʌX����
    HSXCSDIS As String * 1          ' �i�r�w�����ʌX���ʎw��
    HSXCTDIR As String * 2          ' �i�r�w�����ʌX�c����
    HSXCTCEN As Double              ' �i�r�w�����ʌX�c���S
    HSXCTMIN As Double              ' �i�r�w�����ʌX�c����
    HSXCTMAX As Double              ' �i�r�w�����ʌX�c���
    HSXCYDIR As String * 2          ' �i�r�w�����ʌX������
    HSXCYCEN As Double              ' �i�r�w�����ʌX�����S
    HSXCYMIN As Double              ' �i�r�w�����ʌX������
    HSXCYMAX As Double              ' �i�r�w�����ʌX�����
    HSXOF1PD As String * 2          ' �i�r�w�n�e�P�ʒu����
    HSXOF1PN As Double              ' �i�r�w�n�e�P�ʒu����
    HSXOF1PX As Double              ' �i�r�w�n�e�P�ʒu���
    HSXOF1PW As String * 2          ' �i�r�w�n�e�P�ʒu�������@
    HSXOF1LC As Double              ' �i�r�w�n�e�P�����S
    HSXOF1LN As Double              ' �i�r�w�n�e�P������
    HSXOF1LX As Double              ' �i�r�w�n�e�P�����
    HSXOF1DC As Double              ' �i�r�w�n�e�P���a���S
    HSXOF1DN As Double              ' �i�r�w�n�e�P���a����
    HSXOF1DX As Double              ' �i�r�w�n�e�P���a���
    HSXDFORM As String * 1          ' �i�r�w�a�`��
    HSXDPDRC As String * 1          ' �i�r�w�a�ʒu����
    HSXDPACN As Integer             ' �i�r�w�a�ʒu�p�x���S
    HSXDPAMN As Integer             ' �i�r�w�a�ʒu�p�x����
    HSXDPAMX As Integer             ' �i�r�w�a�ʒu�p�x���
    HSXDPKWY As String * 2          ' �i�r�w�a�ʒu�������@
    HSXDPDIR As String * 2          ' �i�r�w�a�ʒu����
    HSXDPMIN As Double              ' �i�r�w�a�ʒu����
    HSXDPMAX As Double              ' �i�r�w�a�ʒu���
    HSXDWCEN As Double              ' �i�r�w�a�В��S
    HSXDWMIN As Double              ' �i�r�w�a�Љ���
    HSXDWMAX As Double              ' �i�r�w�a�Џ��
    HSXDDCEN As Double              ' �i�r�w�a�[���S
    HSXDDMIN As Double              ' �i�r�w�a�[����
    HSXDDMAX As Double              ' �i�r�w�a�[���
    HSXDACEN As Double              ' �i�r�w�a�p�x���S
    HSXDAMIN As Double              ' �i�r�w�a�p�x����
    HSXDAMAX As Double              ' �i�r�w�a�p�x���
    MCNO As String * 10             ' �������Ɠ��������
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ���i�d�lSXL�ް��Q
Public Type typ_TBCME019
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGSXSNO As String * 6          ' �i�Ǘ��r�w���i�ԍ�
    HMGSXSNE As Integer             ' �i�Ǘ��r�w���i�ԍ��}��
    HSXTMMAX As Double              ' �i�r�w�]�ʖ��x���     ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    HSXTMSPH As String * 1          ' �i�r�w�]�ʖ��x����ʒu�Q��
    HSXTMSPT As String * 1          ' �i�r�w�]�ʖ��x����ʒu�Q�_
    HSXTMSPR As String * 1          ' �i�r�w�]�ʖ��x����ʒu�Q��
    HSXTMKHM As String * 1          ' �i�r�w�]�ʖ��x�����p�x�Q��
    HSXTMKHI As String * 1          ' �i�r�w�]�ʖ��x�����p�x�Q��
    HSXTMKHH As String * 1          ' �i�r�w�]�ʖ��x�����p�x�Q��
    HSXTMKHS As String * 1          ' �i�r�w�]�ʖ��x�����p�x�Q��
    HSXLTMIN As Integer             ' �i�r�w�k�^�C������
    HSXLTMAX As Integer             ' �i�r�w�k�^�C�����
    HSXLTSPH As String * 1          ' �i�r�w�k�^�C������ʒu�Q��
    HSXLTSPT As String * 1          ' �i�r�w�k�^�C������ʒu�Q�_
    HSXLTSPI As String * 1          ' �i�r�w�k�^�C������ʒu�Q��
    HSXLTHWT As String * 1          ' �i�r�w�k�^�C���ۏؕ��@�Q��
    HSXLTHWS As String * 1          ' �i�r�w�k�^�C���ۏؕ��@�Q��
    HSXLTKWY As String * 2          ' �i�r�w�k�^�C���������@
    HSXLTNSW As String * 2          ' �i�r�w�k�^�C���M�����@
    HSXLTKHM As String * 1          ' �i�r�w�k�^�C�������p�x�Q��
    HSXLTKHI As String * 1          ' �i�r�w�k�^�C�������p�x�Q��
    HSXLTKHH As String * 1          ' �i�r�w�k�^�C�������p�x�Q��
    HSXLTKHS As String * 1          ' �i�r�w�k�^�C�������p�x�Q��
    HSXLTMBP As Double              ' �i�r�w�k�^�C���ʓ����z
    HSXLTMCL As String * 1          ' �i�r�w�k�^�C���ʓ��v�Z
    HSXCNMIN As Double              ' �i�r�w�Y�f�Z�x����
    HSXCNMAX As Double              ' �i�r�w�Y�f�Z�x���
    HSXCNSPH As String * 1          ' �i�r�w�Y�f�Z�x����ʒu�Q��
    HSXCNSPT As String * 1          ' �i�r�w�Y�f�Z�x����ʒu�Q�_
    HSXCNSPI As String * 1          ' �i�r�w�Y�f�Z�x����ʒu�Q��
    HSXCNHWT As String * 1          ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    HSXCNHWS As String * 1          ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    HSXCNKWY As String * 2          ' �i�r�w�Y�f�Z�x�������@
    HSXCNKHM As String * 1          ' �i�r�w�Y�f�Z�x�����p�x�Q��
    HSXCNKHI As String * 1          ' �i�r�w�Y�f�Z�x�����p�x�Q��
    HSXCNKHH As String * 1          ' �i�r�w�Y�f�Z�x�����p�x�Q��
    HSXCNKHS As String * 1          ' �i�r�w�Y�f�Z�x�����p�x�Q��
    HSXONMIN As Double              ' �i�r�w�_�f�Z�x����
    HSXONMAX As Double              ' �i�r�w�_�f�Z�x���
    HSXONSPH As String * 1          ' �i�r�w�_�f�Z�x����ʒu�Q��
    HSXONSPT As String * 1          ' �i�r�w�_�f�Z�x����ʒu�Q�_
    HSXONSPI As String * 1          ' �i�r�w�_�f�Z�x����ʒu�Q��
    HSXONHWT As String * 1          ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    HSXONHWS As String * 1          ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    HSXONKWY As String * 2          ' �i�r�w�_�f�Z�x�������@
    HSXONKHM As String * 1          ' �i�r�w�_�f�Z�x�����p�x�Q��
    HSXONKHI As String * 1          ' �i�r�w�_�f�Z�x�����p�x�Q��
    HSXONKHH As String * 1          ' �i�r�w�_�f�Z�x�����p�x�Q��
    HSXONKHS As String * 1          ' �i�r�w�_�f�Z�x�����p�x�Q��
    HSXONMBP As Double              ' �i�r�w�_�f�Z�x�ʓ����z
    HSXONMCL As String * 1          ' �i�r�w�_�f�Z�x�ʓ��v�Z
    HSXONLTB As Double              ' �i�r�w�_�f�Z�x�k�s���z
    HSXONLTC As String * 1          ' �i�r�w�_�f�Z�x�k�s�v�Z
    HSXONSDV As Double              ' �i�r�w�_�f�Z�x�W���΍�
    HSXONAMN As Double              ' �i�r�w�_�f�Z�x���ω���
    HSXONAMX As Double              ' �i�r�w�_�f�Z�x���Ϗ��
    HSXOS1MN As Double              ' �i�r�w�_�f�͏o�P����
    HSXOS1MX As Double              ' �i�r�w�_�f�͏o�P���
    HSXOS1NS As String * 2          ' �i�r�w�_�f�͏o�P�M�����@
    HSXOS1SH As String * 1          ' �i�r�w�_�f�͏o�P����ʒu�Q��
    HSXOS1ST As String * 1          ' �i�r�w�_�f�͏o�P����ʒu�Q�_
    HSXOS1SI As String * 1          ' �i�r�w�_�f�͏o�P����ʒu�Q��
    HSXOS1HT As String * 1          ' �i�r�w�_�f�͏o�P�ۏؕ��@�Q��
    HSXOS1HS As String * 1          ' �i�r�w�_�f�͏o�P�ۏؕ��@�Q��
    HSXOS1HM As String * 1          ' �i�r�w�_�f�͏o�P�����p�x�Q��
    HSXOS1KI As String * 1          ' �i�r�w�_�f�͏o�P�����p�x�Q��
    HSXOS1KH As String * 1          ' �i�r�w�_�f�͏o�P�����p�x�Q��
    HSXOS1KS As String * 1          ' �i�r�w�_�f�͏o�P�����p�x�Q��
    HSXOS2MN As Double              ' �i�r�w�_�f�͏o�Q����
    HSXOS2MX As Double              ' �i�r�w�_�f�͏o�Q���
    HSXOS2NS As String * 2          ' �i�r�w�_�f�͏o�Q�M�����@
    HSXOS2SH As String * 1          ' �i�r�w�_�f�͏o�Q����ʒu�Q��
    HSXOS2ST As String * 1          ' �i�r�w�_�f�͏o�Q����ʒu�Q�_
    HSXOS2SI As String * 1          ' �i�r�w�_�f�͏o�Q����ʒu�Q��
    HSXOS2HT As String * 1          ' �i�r�w�_�f�͏o�Q�ۏؕ��@�Q��
    HSXOS2HS As String * 1          ' �i�r�w�_�f�͏o�Q�ۏؕ��@�Q��
    HSXOS2KM As String * 1          ' �i�r�w�_�f�͏o�Q�����p�x�Q��
    HSXOS2KN As String * 1          ' �i�r�w�_�f�͏o�Q�����p�x�Q��
    HSXOS2KH As String * 1          ' �i�r�w�_�f�͏o�Q�����p�x�Q��
    HSXOS2KU As String * 1          ' �i�r�w�_�f�͏o�Q�����p�x�Q��
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
' �ǉ� 2003/09.11 SystemBrain Start
    HSXTMMAXN As Double             ' �i�r�w�]�ʖ��x���
' �ǉ� 2003/09.11 SystemBrain End
End Type


' ���i�d�lSXL�ް��R
Public Type typ_TBCME020
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGSXSNO As String * 6          ' �i�Ǘ��r�w���i�ԍ�
    HMGSXSNE As Integer             ' �i�Ǘ��r�w���i�ԍ��}��
    HSXDENKU As String * 1          ' �i�r�w�c���������L��
    HSXDENMX As Integer             ' �i�r�w�c�������
    HSXDENMN As Integer             ' �i�r�w�c��������
    HSXDENHT As String * 1          ' �i�r�w�c�����ۏؕ��@�Q��
    HSXDENHS As String * 1          ' �i�r�w�c�����ۏؕ��@�Q��
    HSXDVDKU As String * 1          ' �i�r�w�c�u�c�Q�����L��
    HSXDVDMX As Integer             ' �i�r�w�c�u�c�Q���
    HSXDVDMN As Integer             ' �i�r�w�c�u�c�Q����
    HSXDVDHT As String * 1          ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    HSXDVDHS As String * 1          ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    HSXLDLKU As String * 1          ' �i�r�w�k�^�c�k�����L��
    HSXLDLMX As Integer             ' �i�r�w�k�^�c�k���
    HSXLDLMN As Integer             ' �i�r�w�k�^�c�k����
    HSXLDLHT As String * 1          ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    HSXLDLHS As String * 1          ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    HSXGDSZY As String * 1          ' �i�r�w�f�c�������
    HSXGDSPH As String * 1          ' �i�r�w�f�c����ʒu�Q��
    HSXGDSPT As String * 1          ' �i�r�w�f�c����ʒu�Q�_
    HSXGDSPR As String * 1          ' �i�r�w�f�c����ʒu�Q��
    HSXGDZAR As Integer             ' �i�r�w�f�c���O�̈�
    HSXGDKHM As String * 1          ' �i�r�w�f�c�����p�x�Q��
    HSXGDKHI As String * 1          ' �i�r�w�f�c�����p�x�Q��
    HSXGDKHH As String * 1          ' �i�r�w�f�c�����p�x�Q��
    HSXGDKHS As String * 1          ' �i�r�w�f�c�����p�x�Q��
    HSXDSOKE As String * 1          ' �i�r�w�c�r�n�c����
    HSXDSOMX As Long                ' �i�r�w�c�r�n�c���
    HSXDSOMN As Long                ' �i�r�w�c�r�n�c����
    HSXDSOAX As Integer             ' �i�r�w�c�r�n�c�̈���
    HSXDSOAN As Integer             ' �i�r�w�c�r�n�c�̈扺��
    HSXDSOHT As String * 1          ' �i�r�w�c�r�n�c�ۏؕ��@�Q��
    HSXDSOHS As String * 1          ' �i�r�w�c�r�n�c�ۏؕ��@�Q��
    HSXDSOKM As String * 1          ' �i�r�w�c�r�n�c�����p�x�Q��
    HSXDSOKI As String * 1          ' �i�r�w�c�r�n�c�����p�x�Q��
    HSXDSOKH As String * 1          ' �i�r�w�c�r�n�c�����p�x�Q��
    HSXDSOKS As String * 1          ' �i�r�w�c�r�n�c�����p�x�Q��
    HSXLIFTW As String * 2          ' �i�r�w������@
    HSXSDSLP As String * 1          ' �i�r�w�V�[�h�X
    HSXGKKNO As String * 6          ' �i�r�w�O�ϋK�i�m��
    HSXCDOP As String * 1           ' �i�r�w�����h�[�v
    HSXCDOPN As Double              ' �i�r�w�����h�[�v�Z�x
    HSXCDPNI As String * 2          ' �i�r�w�����h�[�v�Z�x�w��
    HSXGSFIN As String * 1          ' �i�r�w�O���d�グ
    HSXCLMIN As Integer             ' �i�r�w����������
    HSXCLMAX As Integer             ' �i�r�w���������
    HSXCLPMN As Integer             ' �i�r�w���������e����
    HSXCLPR As Double               ' �i�r�w���������e�䗦
    HSXWFWAR As String * 1          ' �i�r�w�v�e�v�����������N
    HSXOF1AX As Double              ' �i�r�w�n�r�e�P���Ϗ��
    HSXOF1MX As Double              ' �i�r�w�n�r�e�P���
    HSXOF1SH As String * 1          ' �i�r�w�n�r�e�P����ʒu�Q��
    HSXOF1ST As String * 1          ' �i�r�w�n�r�e�P����ʒu�Q�_
    HSXOF1SR As String * 1          ' �i�r�w�n�r�e�P����ʒu�Q��
    HSXOF1HT As String * 1          ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    HSXOF1HS As String * 1          ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    HSXOF1SZ As String * 1          ' �i�r�w�n�r�e�P�������
    HSXOF1KM As String * 1          ' �i�r�w�n�r�e�P�����p�x�Q��
    HSXOF1KI As String * 1          ' �i�r�w�n�r�e�P�����p�x�Q��
    HSXOF1KH As String * 1          ' �i�r�w�n�r�e�P�����p�x�Q��
    HSXOF1KS As String * 1          ' �i�r�w�n�r�e�P�����p�x�Q��
    HSXOF1NS As String * 2          ' �i�r�w�n�r�e�P�M�����@
    HSXOF1ET As Integer             ' �i�r�w�n�r�e�P�I���d�s��
    HSXOF2AX As Double              ' �i�r�w�n�r�e�Q���Ϗ��
    HSXOF2MX As Double              ' �i�r�w�n�r�e�Q���
    HSXOF2SH As String * 1          ' �i�r�w�n�r�e�Q����ʒu�Q��
    HSXOF2ST As String * 1          ' �i�r�w�n�r�e�Q����ʒu�Q�_
    HSXOF2SR As String * 1          ' �i�r�w�n�r�e�Q����ʒu�Q��
    HSXOF2HT As String * 1          ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    HSXOF2HS As String * 1          ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    HSXOF2SZ As String * 1          ' �i�r�w�n�r�e�Q�������
    HSXOF2KM As String * 1          ' �i�r�w�n�r�e�Q�����p�x�Q��
    HSXOF2KI As String * 1          ' �i�r�w�n�r�e�Q�����p�x�Q��
    HSXOF2KH As String * 1          ' �i�r�w�n�r�e�Q�����p�x�Q��
    HSXOF2KS As String * 1          ' �i�r�w�n�r�e�Q�����p�x�Q��
    HSXOF2NS As String * 2          ' �i�r�w�n�r�e�Q�M�����@
    HSXOF2ET As Integer             ' �i�r�w�n�r�e�Q�I���d�s��
    HSXOF3AX As Double              ' �i�r�w�n�r�e�R���Ϗ��
    HSXOF3MX As Double              ' �i�r�w�n�r�e�R���
    HSXOF3SH As String * 1          ' �i�r�w�n�r�e�R����ʒu�Q��
    HSXOF3ST As String * 1          ' �i�r�w�n�r�e�R����ʒu�Q�_
    HSXOF3SR As String * 1          ' �i�r�w�n�r�e�R����ʒu�Q��
    HSXOF3HT As String * 1          ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    HSXOF3HS As String * 1          ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    HSXOF3SZ As String * 1          ' �i�r�w�n�r�e�R�������
    HSXOF3KM As String * 1          ' �i�r�w�n�r�e�R�����p�x�Q��
    HSXOF3KI As String * 1          ' �i�r�w�n�r�e�R�����p�x�Q��
    HSXOF3KH As String * 1          ' �i�r�w�n�r�e�R�����p�x�Q��
    HSXOF3KS As String * 1          ' �i�r�w�n�r�e�R�����p�x�Q��
    HSXOF3NS As String * 2          ' �i�r�w�n�r�e�R�M�����@
    HSXOF3ET As Integer             ' �i�r�w�n�r�e�R�I���d�s��
    HSXOF4AX As Double              ' �i�r�w�n�r�e�S���Ϗ��
    HSXOF4MX As Double              ' �i�r�w�n�r�e�S���
    HSXOF4SH As String * 1          ' �i�r�w�n�r�e�S����ʒu�Q��
    HSXOF4ST As String * 1          ' �i�r�w�n�r�e�S����ʒu�Q�_
    HSXOF4SR As String * 1          ' �i�r�w�n�r�e�S����ʒu�Q��
    HSXOF4HT As String * 1          ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    HSXOF4HS As String * 1          ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    HSXOF4SZ As String * 1          ' �i�r�w�n�r�e�S�������
    HSXOF4KM As String * 1          ' �i�r�w�n�r�e�S�����p�x�Q��
    HSXOF4KI As String * 1          ' �i�r�w�n�r�e�S�����p�x�Q��
    HSXOF4KH As String * 1          ' �i�r�w�n�r�e�S�����p�x�Q��
    HSXOF4KS As String * 1          ' �i�r�w�n�r�e�S�����p�x�Q��
    HSXOF4NS As String * 2          ' �i�r�w�n�r�e�S�M�����@
    HSXOF4ET As Integer             ' �i�r�w�n�r�e�S�I���d�s��
    HSXBM1AN As Double              ' �i�r�w�a�l�c�P���ω���
    HSXBM1AX As Double              ' �i�r�w�a�l�c�P���Ϗ��
    HSXBM1SH As String * 1          ' �i�r�w�a�l�c�P����ʒu�Q��
    HSXBM1ST As String * 1          ' �i�r�w�a�l�c�P����ʒu�Q�_
    HSXBM1SR As String * 1          ' �i�r�w�a�l�c�P����ʒu�Q��
    HSXBM1HT As String * 1          ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    HSXBM1HS As String * 1          ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    HSXBM1SZ As String * 1          ' �i�r�w�a�l�c�P�������
    HSXBM1KM As String * 1          ' �i�r�w�a�l�c�P�����p�x�Q��
    HSXBM1KI As String * 1          ' �i�r�w�a�l�c�P�����p�x�Q��
    HSXBM1KH As String * 1          ' �i�r�w�a�l�c�P�����p�x�Q��
    HSXBM1KS As String * 1          ' �i�r�w�a�l�c�P�����p�x�Q��
    HSXBM1NS As String * 2          ' �i�r�w�a�l�c�P�M�����@
    HSXBM1ET As Integer             ' �i�r�w�a�l�c�P�I���d�s��
    HSXBM2AN As Double              ' �i�r�w�a�l�c�Q���ω���
    HSXBM2AX As Double              ' �i�r�w�a�l�c�Q���Ϗ��
    HSXBM2SH As String * 1          ' �i�r�w�a�l�c�Q����ʒu�Q��
    HSXBM2ST As String * 1          ' �i�r�w�a�l�c�Q����ʒu�Q�_
    HSXBM2SR As String * 1          ' �i�r�w�a�l�c�Q����ʒu�Q��
    HSXBM2HT As String * 1          ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    HSXBM2HS As String * 1          ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    HSXBM2SZ As String * 1          ' �i�r�w�a�l�c�Q�������
    HSXBM2KM As String * 1          ' �i�r�w�a�l�c�Q�����p�x�Q��
    HSXBM2KI As String * 1          ' �i�r�w�a�l�c�Q�����p�x�Q��
    HSXBM2KH As String * 1          ' �i�r�w�a�l�c�Q�����p�x�Q��
    HSXBM2KS As String * 1          ' �i�r�w�a�l�c�Q�����p�x�Q��
    HSXBM2NS As String * 2          ' �i�r�w�a�l�c�Q�M�����@
    HSXBM2ET As Integer             ' �i�r�w�a�l�c�Q�I���d�s��
    HSXBM3AN As Double              ' �i�r�w�a�l�c�R���ω���
    HSXBM3AX As Double              ' �i�r�w�a�l�c�R���Ϗ��
    HSXBM3SH As String * 1          ' �i�r�w�a�l�c�R����ʒu�Q��
    HSXBM3ST As String * 1          ' �i�r�w�a�l�c�R����ʒu�Q�_
    HSXBM3SR As String * 1          ' �i�r�w�a�l�c�R����ʒu�Q��
    HSXBM3HT As String * 1          ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
    HSXBM3HS As String * 1          ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
    HSXBM3SZ As String * 1          ' �i�r�w�a�l�c�R�������
    HSXBM3KM As String * 1          ' �i�r�w�a�l�c�R�����p�x�Q��
    HSXBM3KI As String * 1          ' �i�r�w�a�l�c�R�����p�x�Q��
    HSXBM3KH As String * 1          ' �i�r�w�a�l�c�R�����p�x�Q��
    HSXBM3KS As String * 1          ' �i�r�w�a�l�c�R�����p�x�Q��
    HSXBM3NS As String * 2          ' �i�r�w�a�l�c�R�M�����@
    HSXBM3ET As Integer             ' �i�r�w�a�l�c�R�I���d�s��
    HSXNOTE As String               ' �i�r�w���L
    HSXRS1N As String               ' �i�r�w�\���P�Q��
    HSXRS1Y As String               ' �i�r�w�\���P�Q�p
    HSXRS2N As String               ' �i�r�w�\���Q�Q��
    HSXRS2Y As String               ' �i�r�w�\���Q�Q�p
    HSXRS3N As String               ' �i�r�w�\���R�Q��
    HSXRS3Y As String               ' �i�r�w�\���R�Q�p
    HSXRS4N As String               ' �i�r�w�\���S�Q��
    HSXRS4Y As String               ' �i�r�w�\���S�Q�p
    HSXRS5N As String               ' �i�r�w�\���T�Q��
    HSXRS5Y As String               ' �i�r�w�\���T�Q�p
    HSXRS6N As String               ' �i�r�w�\���U�Q��
    HSXRS6Y As String               ' �i�r�w�\���U�Q�p
    HSXRS7N As String               ' �i�r�w�\���V�Q��
    HSXRS7Y As String               ' �i�r�w�\���V�Q�p
    HSXRS8N As String               ' �i�r�w�\���W�Q��
    HSXRS8Y As String               ' �i�r�w�\���W�Q�p
    HSXRS9N As String               ' �i�r�w�\���X�Q��
    HSXRS9Y As String               ' �i�r�w�\���X�Q�p
    HSXRS10N As String              ' �i�r�w�\���P�O�Q��
    HSXRS10Y As String              ' �i�r�w�\���P�O�Q�p
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
' �ǉ� 2003/09.11 SystemBrain Start
    HSXDVDMXN As Integer            ' �i�r�w�c�u�c�Q���
    HSXDVDMNN As Integer            ' �i�r�w�c�u�c�Q����
    HSXDSONS As String * 2          ' �i�r�w�c�r�n�c�M�����@
    HSXCDOPMX As Double             ' �i�r�w�����h�[�v�Z�x����
    HSXCDOPMN As Double             ' �i�r�w�����h�[�v�Z�x���
' �ǉ� 2003/09.11 SystemBrain End
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
    HSXOSF1PTK As String * 1        ' �i�r�w�n�r�e�P�p�^���敪
    HSXOSF2PTK As String * 1        ' �i�r�w�n�r�e�Q�p�^���敪
    HSXOSF3PTK As String * 1        ' �i�r�w�n�r�e�R�p�^���敪
    HSXOSF4PTK As String * 1        ' �i�r�w�n�r�e�S�p�^���敪
    HSXBMD1MBP As Double            ' �i�r�w�a�l�c�P�ʓ����z
    HSXBMD2MBP As Double            ' �i�r�w�a�l�c�Q�ʓ����z
    HSXBMD3MBP As Double            ' �i�r�w�a�l�c�R�ʓ����z
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
' �ǉ� 2003/09.11 SystemBrain Start
    HSXBMD1MCL As String * 2        ' �iSXBMD1�ʓ��v�Z
    HSXBMD2MCL As String * 2        ' �iSXBMD2�ʓ��v�Z
    HSXBMD3MCL As String * 2        ' �iSXBMD3�ʓ��v�Z
' �ǉ� 2003/09.11 SystemBrain End

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    HSXGDPTK As String * 1          ' �i�r�w�f�c�p�^���敪
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

    'Add Start 2011/01/26 SMPK Miyata
    HSXCPK      As String * 1       '�i�r�w�b�p�^�[���敪
    HSXCSZ      As String * 1       '�i�r�w�b�������
    HSXCHT      As String * 1       '�i�r�w�b�ۏؕ��@�Q��
    HSXCHS      As String * 1       '�i�r�w�b�ۏؕ��@�Q��
    HSXCJPK     As String * 1       '�i�r�w�b�i�p�^�[���敪
    HSXCJNS     As String * 2       '�i�r�w�b�i�M�����@
    HSXCJHT     As String * 1       '�i�r�w�b�i�ۏؕ��@�Q��
    HSXCJHS     As String * 1       '�i�r�w�b�i�ۏؕ��@�Q��
    HSXCJLTPK   As String * 1       '�i�r�w�b�i�k�s�p�^�[���敪
    HSXCJLTNS   As String * 2       '�i�r�w�b�i�k�s�M�����@
    HSXCJLTHT   As String * 1       '�i�r�w�b�i�k�s�ۏؕ��@�Q��
    HSXCJLTHS   As String * 1       '�i�r�w�b�i�k�s�ۏؕ��@�Q��
    HSXCJ2PK    As String * 1       '�i�r�w�b�i�Q�p�^�[���敪
    HSXCJ2NS    As String * 2       '�i�r�w�b�i�Q�M�����@
    HSXCJ2HT    As String * 1       '�i�r�w�b�i�Q�ۏؕ��@�Q��
    HSXCJ2HS    As String * 1       '�i�r�w�b�i�Q�ۏؕ��@�Q��
    'Add End   2011/01/26 SMPK Miyata
    
    'Add Start 2011/02/17 Y.Hitomi
    HSXCOSF3NS As String * 2        '�i�r�w�b�n�r�e�R�M�����@
    'Add End   2011/02/17 Y.Hitomi

End Type


' ���i�d�lWF�ް��P
Public Type typ_TBCME021
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGWFSNO As String * 6          ' �i�Ǘ��v�e���i�ԍ�
    HMGWFSNE As Integer             ' �i�Ǘ��v�e���i�ԍ��}��
    CONFLAG As String * 1           ' �m�F�t���O
    REINFLAG As String * 1          ' �ĕt�^�t���O
    HWFTRWKB As String * 1          ' �i�v�e�����ۋ敪
    HWFFACES As String * 2          ' �i�v�e�\�ʎd�グ
    HWFBACKS As String * 2          ' �i�v�e���d�グ
    HWFBDSWY As String * 2          ' �i�v�e�a�c�������@
'    HSXTYPE As String * 1           ' �i�v�e�^�C�v
    HWFTYPE As String * 1           ' �i�v�e�^�C�v�@05/03/01 ooba
    HWFTYPKW As String * 1          ' �i�v�e�^�C�v�������@
    HWFDOP As String * 1            ' �i�v�e�h�[�p���g
    HWFFKBWK As String * 1          ' �i�v�e�\�ʋ敪���@�Q��
    HWFFKBWS As String * 1          ' �i�v�e�\�ʋ敪���@�Q�w
    HWFRMIN As Double               ' �i�v�e���R����
    HWFRMAX As Double               ' �i�v�e���R���
    HWFRSPOH As String * 1          ' �i�v�e���R����ʒu�Q��
    HWFRSPOT As String * 1          ' �i�v�e���R����ʒu�Q�_
    HWFRSPOI As String * 1          ' �i�v�e���R����ʒu�Q��
    HWFRHWYT As String * 1          ' �i�v�e���R�ۏؕ��@�Q��
    HWFRHWYS As String * 1          ' �i�v�e���R�ۏؕ��@�Q��
    HWFRKWAY As String * 2          ' �i�v�e���R�������@
    HWFRKHNM As String * 1          ' �i�v�e���R�����p�x�Q��
    HWFRKHNN As String * 1          ' �i�v�e���R�����p�x�Q��
    HWFRKHNH As String * 1          ' �i�v�e���R�����p�x�Q��
    HWFRKHNU As String * 1          ' �i�v�e���R�����p�x�Q�E
    HWFRSDEV As Double              ' �i�v�e���R�W���΍�
    HWFRAMIN As Double              ' �i�v�e���R���ω���
    HWFRAMAX As Double              ' �i�v�e���R���Ϗ��
    HWFRMBNP As Double              ' �i�v�e���R�ʓ����z
    HWFRMCAL As String * 1          ' �i�v�e���R�ʓ��v�Z
    HWFRMBP2 As Double              ' �i�v�e���R�ʓ����z�Q
    HWFRMCL2 As String * 1          ' �i�v�e���R�ʓ��v�Z�Q
    HWFRKBSH As String * 1          ' �i�v�e���R�U�敪����ʒu�Q��
    HWFRKBST As String * 1          ' �i�v�e���R�U�敪����ʒu�Q�_
    HWFRKBSI As String * 1          ' �i�v�e���R�U�敪����ʒu�Q��
    HWFRKBHT As String * 1          ' �i�v�e���R�U�敪�ۏؕ��@�Q��
    HWFRKBHS As String * 1          ' �i�v�e���R�U�敪�ۏؕ��@�Q��
    HWFSTMAX As Double              ' �i�v�e�X�g���G���
    HWFSTSPH As String * 1          ' �i�v�e�X�g���G����ʒu�Q��
    HWFSTSPT As String * 1          ' �i�v�e�X�g���G����ʒu�Q�_
    HWFSTSPI As String * 1          ' �i�v�e�X�g���G����ʒu�Q��
    HWFSTHWT As String * 1          ' �i�v�e�X�g���G�ۏؕ��@�Q��
    HWFSTHWS As String * 1          ' �i�v�e�X�g���G�ۏؕ��@�Q��
    HWFSTKWY As String * 2          ' �i�v�e�X�g���G�������@
    HWFSTKHM As String * 1          ' �i�v�e�X�g���G�����p�x�Q��
    HWFSTKHN As String * 1          ' �i�v�e�X�g���G�����p�x�Q��
    HWFSTKHH As String * 1          ' �i�v�e�X�g���G�����p�x�Q��
    HWFSTKHU As String * 1          ' �i�v�e�X�g���G�����p�x�Q�E
    HWFACEN As Double               ' �i�v�e�����S
    HWFAMIN As Double               ' �i�v�e������
    HWFAMAX As Double               ' �i�v�e�����
    HWFASPOH As String * 1          ' �i�v�e������ʒu�Q��
    HWFASPOT As String * 1          ' �i�v�e������ʒu�Q�_
    HWFASPOI As String * 1          ' �i�v�e������ʒu�Q��
    HWFAHWYT As String * 1          ' �i�v�e���ۏؕ��@�Q��
    HWFAHWYS As String * 1          ' �i�v�e���ۏؕ��@�Q��
    HWFAKWAY As String * 1          ' �i�v�e���������@
    HWFAKHNM As String * 1          ' �i�v�e�������p�x�Q��
    HWFAKHNN As String * 1          ' �i�v�e�������p�x�Q��
    HWFAKHNH As String * 1          ' �i�v�e�������p�x�Q��
    HWFAKHNU As String * 1          ' �i�v�e�������p�x�Q�E
    HWFASDEV As Double              ' �i�v�e���W���΍�
    HWFAAMIN As Double              ' �i�v�e�����ω���
    HWFAAMAX As Double              ' �i�v�e�����Ϗ��
    HWFAMBNP As Double              ' �i�v�e���ʓ����z
    HWFAMCAL As String * 1          ' �i�v�e���ʓ��v�Z
    HWFALTBP As Double              ' �i�v�e���k�s���z
    HWFALTCL As String * 1          ' �i�v�e���k�s�v�Z
    HWFALTRA As Double              ' �i�v�e���k�s�͈�
    HWFAMRAN As Double              ' �i�v�e���ʓ��͈�
    HWFDIVS As Integer              ' �i�v�e������
    HWFAKBSH As String * 1          ' �i�v�e���U�敪����ʒu�Q��
    HWFAKBST As String * 1          ' �i�v�e���U�敪����ʒu�Q�_
    HWFAKBSI As String * 1          ' �i�v�e���U�敪����ʒu�Q��
    HWFAKBHT As String * 1          ' �i�v�e���U�敪�ۏؕ��@�Q��
    HWFAKBHS As String * 1          ' �i�v�e���U�敪�ۏؕ��@�Q��
    HWFWFORM As String * 1          ' �i�v�e�E�F�[�n�`��
    HWFD1CEN As Double              ' �i�v�e���a�P���S
    HWFD1MIN As Double              ' �i�v�e���a�P����
    HWFD1MAX As Double              ' �i�v�e���a�P���
    HWFD2CEN As Double              ' �i�v�e���a�Q���S
    HWFD2MIN As Double              ' �i�v�e���a�Q����
    HWFD2MAX As Double              ' �i�v�e���a�Q���
    HWFDKHNM As String * 1          ' �i�v�e���a�����p�x�Q��
    HWFDKHNN As String * 1          ' �i�v�e���a�����p�x�Q��
    HWFDKHNH As String * 1          ' �i�v�e���a�����p�x�Q��
    HWFDKHNU As String * 1          ' �i�v�e���a�����p�x�Q�E
    HWFLPMNP As Integer             ' �i�v�e�k�o���ŏ����H��
    HWFSGMNP As Integer             ' �i�v�e�r�f���ŏ����H��
    HWFETMNP As Integer             ' �i�v�e�d�s���ŏ����H��
    HWFMPMNP As Integer             ' �i�v�e�l�o���ŏ����H��
    HWFLPKS1 As String * 1          ' �i�v�e�k�o�����ގ�P
    HWFLPKS2 As String * 1          ' �i�v�e�k�o�����ގ�Q
    HWFLPKZ1 As String * 1          ' �i�v�e�k�o�����ޗ��x��P
    HWFLPKZ2 As String * 1          ' �i�v�e�k�o�����ޗ��x��Q
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ���i�d�lWF�ް��Q
Public Type typ_TBCME022
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGWFSNO As String * 6          ' �i�Ǘ��v�e���i�ԍ�
    HMGWFSNE As Integer             ' �i�Ǘ��v�e���i�ԍ��}��
    HWFCDIR As String * 1           ' �i�v�e�����ʕ���
    HWFCSCEN As Double              ' �i�v�e�����ʌX���S
    HWFCSMIN As Double              ' �i�v�e�����ʌX����
    HWFCSMAX As Double              ' �i�v�e�����ʌX���
    HWFCSDIS As String * 1          ' �i�v�e�����ʌX���ʎw��
    HWFCSDIR As String * 2          ' �i�v�e�����ʌX����
    HWFCKWAY As String * 2          ' �i�v�e�����ʌ������@
    HWFCKHNM As String * 1          ' �i�v�e�����ʌ����p�x�Q��
    HWFCKHNN As String * 1          ' �i�v�e�����ʌ����p�x�Q��
    HWFCKHNH As String * 1          ' �i�v�e�����ʌ����p�x�Q��
    HWFCKHNU As String * 1          ' �i�v�e�����ʌ����p�x�Q�E
    HWFCTDIR As String * 2          ' �i�v�e�����ʌX�c����
    HWFCTCEN As Double              ' �i�v�e�����ʌX�c���S
    HWFCTMIN As Double              ' �i�v�e�����ʌX�c����
    HWFCTMAX As Double              ' �i�v�e�����ʌX�c���
    HWFCYDIR As String * 2          ' �i�v�e�����ʌX������
    HWFCYCEN As Double              ' �i�v�e�����ʌX�����S
    HWFCYMIN As Double              ' �i�v�e�����ʌX������
    HWFCYMAX As Double              ' �i�v�e�����ʌX�����
    HWFKPTNN As String * 3          ' �i�v�e�����p�^����
    HWFOFPKM As String * 1          ' �i�v�e�n�e�ʒu�����p�x�Q��
    HWFOFPKN As String * 1          ' �i�v�e�n�e�ʒu�����p�x�Q��
    HWFOFPKH As String * 1          ' �i�v�e�n�e�ʒu�����p�x�Q��
    HWFOFPKU As String * 1          ' �i�v�e�n�e�ʒu�����p�x�Q�E
    HWFOFLKM As String * 1          ' �i�v�e�n�e�������p�x�Q��
    HWFOFLKN As String * 1          ' �i�v�e�n�e�������p�x�Q��
    HWFOFLKH As String * 1          ' �i�v�e�n�e�������p�x�Q��
    HWFOFLKU As String * 1          ' �i�v�e�n�e�������p�x�Q�E
    HWFOF1PD As String * 2          ' �i�v�e�n�e�P�ʒu����
    HWFOF1PN As Double              ' �i�v�e�n�e�P�ʒu����
    HWFOF1PX As Double              ' �i�v�e�n�e�P�ʒu���
    HWFOF1PW As String * 2          ' �i�v�e�n�e�P�ʒu�������@
    HWFOF1LC As Double              ' �i�v�e�n�e�P�����S
    HWFOF1LN As Double              ' �i�v�e�n�e�P������
    HWFOF1LX As Double              ' �i�v�e�n�e�P�����
    HWFOF1RF As String * 1          ' �i�v�e�n�e�P���[�q�`��
    HWFOFRRC As Double              ' �i�v�e�n�e���[�q�E���S
    HWFOFRRN As Double              ' �i�v�e�n�e���[�q�E����
    HWFOFRRX As Double              ' �i�v�e�n�e���[�q�E���
    HWFOFRLC As Double              ' �i�v�e�n�e���[�q�����S
    HWFOFRLN As Double              ' �i�v�e�n�e���[�q������
    HWFOFRLX As Double              ' �i�v�e�n�e���[�q�����
    HWFOF1DC As Double              ' �i�v�e�n�e�P���a���S
    HWFOF1DN As Double              ' �i�v�e�n�e�P���a����
    HWFOF1DX As Double              ' �i�v�e�n�e�P���a���
    HWFZFORM As String * 1          ' �i�v�e�ޗ��`��
    HWFD3CEN As Double              ' �i�v�e���a�R���S
    HWFD3MIN As Double              ' �i�v�e���a�R����
    HWFD3MAX As Double              ' �i�v�e���a�R���
    HWFDFKJ As String * 1           ' �i�v�e�a�`��
    HWFDFKHM As String * 1          ' �i�v�e�a�`�󌟍��p�x�Q��
    HWFDFKHN As String * 1          ' �i�v�e�a�`�󌟍��p�x�Q��
    HWFDFKHH As String * 1          ' �i�v�e�a�`�󌟍��p�x�Q��
    HWFDFKHU As String * 1          ' �i�v�e�a�`�󌟍��p�x�Q�E
    HWFDPDRC As String * 1          ' �i�v�e�a�ʒu����
    HWFDPACN As Integer             ' �i�v�e�a�ʒu�p�x���S
    HWFDPAMN As Integer             ' �i�v�e�a�ʒu�p�x����
    HWFDPAMX As Integer             ' �i�v�e�a�ʒu�p�x���
    HWFDPDIR As String * 2          ' �i�v�e�a�ʒu����
    HWFDPMIN As Double              ' �i�v�e�a�ʒu����
    HWFDPMAX As Double              ' �i�v�e�a�ʒu���
    HWFDPKWY As String * 2          ' �i�v�e�a�ʒu�������@
    HWFDPKHM As String * 1          ' �i�v�e�a�ʒu�����p�x�Q��
    HWFDPKHB As String * 1          ' �i�v�e�a�ʒu�����p�x�Q��
    HWFDPKHH As String * 1          ' �i�v�e�a�ʒu�����p�x�Q��
    HWFDPKHU As String * 1          ' �i�v�e�a�ʒu�����p�x�Q�E
    HWFDACEN As Double              ' �i�v�e�a�p�x���S
    HWFDAMIN As Double              ' �i�v�e�a�p�x����
    HWFDAMAX As Double              ' �i�v�e�a�p�x���
    HWFDWCEN As Double              ' �i�v�e�a�В��S
    HWFDWMIN As Double              ' �i�v�e�a�Љ���
    HWFDWMAX As Double              ' �i�v�e�a�Џ��
    HWFDDCEN As Double              ' �i�v�e�a�[���S
    HWFDDMIN As Double              ' �i�v�e�a�[����
    HWFDDMAX As Double              ' �i�v�e�a�[���
    HWFDBRCN As Double              ' �i�v�e�a��q���S
    HWFDBRMN As Double              ' �i�v�e�a��q����
    HWFDBRMX As Double              ' �i�v�e�a��q���
    HWFDRRCN As Double              ' �i�v�e�a�E�q���S
    HWFDRRMN As Double              ' �i�v�e�a�E�q����
    HWFDRRMX As Double              ' �i�v�e�a�E�q���
    HWFDLRCN As Double              ' �i�v�e�a���q���S
    HWFDLRMN As Double              ' �i�v�e�a���q����
    HWFDLRMX As Double              ' �i�v�e�a���q���
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ���i�d�lWF�ް��R
Public Type typ_TBCME023
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGWFSNO As String * 6          ' �i�Ǘ��v�e���i�ԍ�
    HMGWFSNE As Integer             ' �i�Ǘ��v�e���i�ԍ��}��
    HWFMFORM As String * 1          ' �i�v�e�ʎ�`��
    KWFMM As String * 1             ' �i�v�e�ʎ�ʑe
    HWFMFKHM As String * 1          ' �i�v�e�ʎ�`�󌟍��p�x�Q��
    HWFMFKHN As String * 1          ' �i�v�e�ʎ�`�󌟍��p�x�Q��
    HWFMFKHH As String * 1          ' �i�v�e�ʎ�`�󌟍��p�x�Q��
    HWFMFKHU As String * 1          ' �i�v�e�ʎ�`�󌟍��p�x�Q�E
    HWFMACEN As Double              ' �i�v�e�ʎ�p�x���S
    HWFMAMIN As Double              ' �i�v�e�ʎ�p�x����
    HWFMAMAX As Double              ' �i�v�e�ʎ�p�x���
    HWFMWFCN As Integer             ' �i�v�e�ʎ�Е\���S
    HWFMWFMN As Integer             ' �i�v�e�ʎ�Е\����
    HWFMWFMX As Integer             ' �i�v�e�ʎ�Е\���
    HWFMWBCN As Integer             ' �i�v�e�ʎ�З����S
    HWFMWBMN As Integer             ' �i�v�e�ʎ�З�����
    HWFMWBMX As Integer             ' �i�v�e�ʎ�З����
    HWFMHCEN As Integer             ' �i�v�e�ʎ捂���S
    HWFMHMIN As Integer             ' �i�v�e�ʎ捂����
    HWFMHMAX As Integer             ' �i�v�e�ʎ捂���
    HWFMPWCN As Integer             ' �i�v�e�ʎ��[�В��S
    HWFMPWMN As Integer             ' �i�v�e�ʎ��[�Љ���
    HWFMPWMX As Integer             ' �i�v�e�ʎ��[�Џ��
    HWFMPRCN As Double              ' �i�v�e�ʎ��[�q���S
    HWFMPRMN As Double              ' �i�v�e�ʎ��[�q����
    HWFMPRMX As Double              ' �i�v�e�ʎ��[�q���
    HWFMBACEN As Double              ' �i�v�e�ʎ无�p�x���S�@6/22 Yam
    HWFMBAMIN As Double              ' �i�v�e�ʎ无�p�x����
    HWFMBAMAX As Double              ' �i�v�e�ʎ无�p�x���
    HWFDMFRM As String * 1          ' �i�v�e�a�ʎ�`��
    HWFDMM As String * 1            ' �i�v�e�a�ʎ�ʑe
    HWFDMACN As Double              ' �i�v�e�a�ʎ�p�x���S
    HWFDMPRC As Double              ' �i�v�e�a�ʎ��[�q���S
    HWFIDKBU As String * 1          ' �i�v�e�h�c�敪�L��
    HWFIDWAY As String * 1          ' �i�v�e�h�c���@
    HWFIDPRI As String * 1          ' �i�v�e�h�c�󎚎��
    HWFIDKND As String * 1          ' �i�v�e�h�c���
    HWFIDDIR As String * 1          ' �i�v�e�h�c����
    HWFIDFAC As String * 1          ' �i�v�e�h�c��
    HWFCSIZE As String * 1          ' �i�v�e�����T�C�Y
    HWFIDPBS As String * 1          ' �i�v�e�h�c�ʒu����
    HWFIDFIG As Integer             ' �i�v�e�h�c����
    HWFIDCON As String              ' �i�v�e�h�c���e
    HWFIDZAR As Double              ' �i�v�e�h�c���O�̈�
    HWFIDPAP As String * 1          ' �i�v�e�h�c�󎚘A�Ԏw��
    HWFIDDCN As Integer             ' �i�v�e�h�c�h�b�g�[���S
    HWFIDDMX As Integer             ' �i�v�e�h�c�h�b�g�[���
    HWFIDDMN As Integer             ' �i�v�e�h�c�h�b�g�[����
    HWFIDSCN As Integer             ' �i�v�e�h�c�h�b�g�r���S
    HWFIDSMX As Integer             ' �i�v�e�h�c�h�b�g�r���
    HWFIDSMN As Integer             ' �i�v�e�h�c�h�b�g�r����
    HWFIDBCZ As String * 3          ' �i�v�e�h�c�a�b�ڍא}��
    HWFIDZNO As Long                ' �i�v�e�h�c�}�ԍ�
    HWFBDPRS As Double              ' �i�v�e�a�c����
    HWFBDTIM As Integer             ' �i�v�e�a�c��
    HWFETWAY As String * 2          ' �i�v�e�d�s���@
    HWFMPFIN As String * 1          ' �i�v�e�l�o�d�グ
    HWFLWASW As String * 1          ' �i�v�e�ŏI�����@
    HWFCDOP As String * 1           ' �i�v�e�����h�[�v
    HWFCDOPN As Double              ' �i�v�e�����h�[�v�Z�x
    HWFCDPNI As String * 2          ' �i�v�e�����h�[�v�Z�x�w��
    HWFCMPUL As String * 1          ' �i�v�e�b�l�o�E�l�����x��
    HWFTPROC As String * 1          ' �i�v�e�ψ����H
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ���i�d�lWF�ް��S
Public Type typ_TBCME024
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGWFSNO As String * 6          ' �i�Ǘ��v�e���i�ԍ�
    HMGWFSNE As Integer             ' �i�Ǘ��v�e���i�ԍ��}��
    HWFM1S As String * 1            ' �i�v�e���P��
    HWFM1H As String * 1            ' �i�v�e���P�t��
    HWFM2S As String * 1            ' �i�v�e���Q��
    HWFM2H As String * 1            ' �i�v�e���Q�t��
    HWFNJSUM As String * 1          ' �i�v�e�m�W���[�������L��
    HWFNJSMX As Double              ' �i�v�e�m�W���[�������Џ��
    HWFNJSMN As Double              ' �i�v�e�m�W���[�������Љ���
    HWFOXCEN As Long                ' �i�v�e�_���������S
    HWFOXMIN As Long                ' �i�v�e�_����������
    HWFOXMAX As Long                ' �i�v�e�_���������
    HWFOXSPH As String * 1          ' �i�v�e�_����������ʒu�Q��
    HWFOXSPT As String * 1          ' �i�v�e�_����������ʒu�Q�_
    HWFOXSPI As String * 1          ' �i�v�e�_����������ʒu�Q��
    HWFOXHWT As String * 1          ' �i�v�e�_�������ۏؕ��@�Q��
    HWFOXHWS As String * 1          ' �i�v�e�_�������ۏؕ��@�Q��
    HWFOXHWY As String * 2          ' �i�v�e�_�������������@
    HWFOXNPO As String * 1          ' �i�v�e�_����������ʒu
    HWFOXKHM As String * 1          ' �i�v�e�_�����������p�x�Q��
    HWFOXKHN As String * 1          ' �i�v�e�_�����������p�x�Q��
    HWFOXKHH As String * 1          ' �i�v�e�_�����������p�x�Q��
    HWFOXKHU As String * 1          ' �i�v�e�_�����������p�x�Q�E
    HWFOXZAR As Integer             ' �i�v�e�_�������O�̈�
    HWFOXMBP As Double              ' �i�v�e�_�������ʓ����z
    HWFOXMCL As String * 1          ' �i�v�e�_�������ʓ��v�Z
    HWFOXMRA As Integer             ' �i�v�e�_�������ʓ��͈�
    HWFOXLTB As Double              ' �i�v�e�_�������k�s���z
    HWFOXLTC As String * 1          ' �i�v�e�_�������k�s�v�Z
    HWFOXLTR As Integer             ' �i�v�e�_�������k�s�͈�
    HWFPSCEN As Long                ' �i�v�e�|���V�������S
    HWFPSMIN As Long                ' �i�v�e�|���V��������
    HWFPSMAX As Long                ' �i�v�e�|���V�������
    HWFPSSPH As String * 1          ' �i�v�e�|���V��������ʒu�Q��
    HWFPSSPT As String * 1          ' �i�v�e�|���V��������ʒu�Q�_
    HWFPSSPI As String * 1          ' �i�v�e�|���V��������ʒu�Q��
    HWFPSHWT As String * 1          ' �i�v�e�|���V�����ۏؕ��@�Q��
    HWFPSHWS As String * 1          ' �i�v�e�|���V�����ۏؕ��@�Q��
    HWFPSKWY As String * 2          ' �i�v�e�|���V�����������@
    HWFPSNPS As String * 1          ' �i�v�e�|���V��������ʒu
    HWFPSKHM As String * 1          ' �i�v�e�|���V���������p�x�Q��
    HWFPSKHN As String * 1          ' �i�v�e�|���V���������p�x�Q��
    HWFPSKHH As String * 1          ' �i�v�e�|���V���������p�x�Q��
    HWFPSKHU As String * 1          ' �i�v�e�|���V���������p�x�Q�E
    HWFPSMBP As Double              ' �i�v�e�|���V�����ʓ����z
    HWFPSMCL As String * 1          ' �i�v�e�|���V�����ʓ��v�Z
    HWFPSMRA As Integer             ' �i�v�e�|���V�����ʓ��͈�
    HWFNOXCN As Long                ' �i�v�e�����������S
    HWFNOXMN As Long                ' �i�v�e������������
    HWFNOXMX As Long                ' �i�v�e�����������
    HWFNOXSH As String * 1          ' �i�v�e������������ʒu�Q��
    HWFNOXST As String * 1          ' �i�v�e������������ʒu�Q�_
    HWFNOXSI As String * 1          ' �i�v�e������������ʒu�Q��
    HWFNOXHT As String * 1          ' �i�v�e���������ۏؕ��@�Q��
    HWFNOXHS As String * 1          ' �i�v�e���������ۏؕ��@�Q��
    HWFNOXHW As String * 2          ' �i�v�e���������������@
    HWFNOXNP As String * 1          ' �i�v�e������������ʒu
    HWFNOXKM As String * 1          ' �i�v�e�������������p�x�Q��
    HWFNOXKN As String * 1          ' �i�v�e�������������p�x�Q��
    HWFNOXKH As String * 1          ' �i�v�e�������������p�x�Q��
    HWFNOXKU As String * 1          ' �i�v�e�������������p�x�Q�E
    HWFNOXMB As Double              ' �i�v�e���������ʓ����z
    HWFNOXMC As String * 1          ' �i�v�e���������ʓ��v�Z
    HWFNOXMR As Integer             ' �i�v�e���������ʓ��͈�
    HWFMKMIN As Double              ' �i�v�e�����בw����
    HWFMKMAX As Double              ' �i�v�e�����בw���
    HWFMKSPH As String * 1          ' �i�v�e�����בw����ʒu�Q��
    HWFMKSPT As String * 1          ' �i�v�e�����בw����ʒu�Q�_
    HWFMKSPR As String * 1          ' �i�v�e�����בw����ʒu�Q��
    HWFMKHWT As String * 1          ' �i�v�e�����בw�ۏؕ��@�Q��
    HWFMKHWS As String * 1          ' �i�v�e�����בw�ۏؕ��@�Q��
    HWFMKSZY As String * 1          ' �i�v�e�����בw�������
    HWFMKKHM As String * 1          ' �i�v�e�����בw�����p�x�Q��
    HWFMKKHN As String * 1          ' �i�v�e�����בw�����p�x�Q��
    HWFMKKHH As String * 1          ' �i�v�e�����בw�����p�x�Q��
    HWFMKKHU As String * 1          ' �i�v�e�����בw�����p�x�Q�E
    HWFMKNSW As String * 2          ' �i�v�e�����בw�M�����@
    HWFMKCET As Integer             ' �i�v�e�����בw�I���d�s��
    HWFDZSWY As String * 1          ' �i�v�e�c�y�������@
    HWFD1STO As Integer             ' �i�v�e�c�y�P�r�s���x
    HWFD1STT As Integer             ' �i�v�e�c�y�P�r�s����
    HWFD1STG As String * 1          ' �i�v�e�c�y�P�r�s�K�X����
    HWFD2NDO As Integer             ' �i�v�e�c�y�Q�m�c���x
    HWFD2NDC As Integer             ' �i�v�e�c�y�Q�m�c���x���
    HWFD2NDT As Integer             ' �i�v�e�c�y�Q�m�c����
    HWFD3RDO As Integer             ' �i�v�e�c�y�R�q�c���x
    HWFD3RDT As Integer             ' �i�v�e�c�y�R�q�c����
    HWFDZMPS As String * 1          ' �i�v�e�c�y�l�o�����敪
    HWFH2ANO As Integer             ' �i�v�e�g�Q�`�m���x
    HWFH2ANT As Integer             ' �i�v�e�g�Q�`�m����
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
' �ǉ� 2003/09.11 SystemBrain Start
    HWFANGZY As String * 1          ' �i�v�e�����`�m�K�X����
' �ǉ� 2003/09.11 SystemBrain End
End Type


' ���i�d�lWF�ް��T
Public Type typ_TBCME025
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGWFSNO As String * 6          ' �i�Ǘ��v�e���i�ԍ�
    HMGWFSNE As Integer             ' �i�Ǘ��v�e���i�ԍ��}��
    HWFTMMAX As Long                ' �i�v�e�]�ʖ��x���
    HWFTMSPH As String * 1          ' �i�v�e�]�ʖ��x����ʒu�Q��
    HWFTMSPT As String * 1          ' �i�v�e�]�ʖ��x����ʒu�Q�_
    HWFTMSPR As String * 1          ' �i�v�e�]�ʖ��x����ʒu�Q��
    HWFTMKHM As String * 1          ' �i�v�e�]�ʖ��x�����p�x�Q��
    HWFTMKHN As String * 1          ' �i�v�e�]�ʖ��x�����p�x�Q��
    HWFTMKHH As String * 1          ' �i�v�e�]�ʖ��x�����p�x�Q��
    HWFTMKHU As String * 1          ' �i�v�e�]�ʖ��x�����p�x�Q�E
    HWFLTMIN As Integer             ' �i�v�e�k�^�C������
    HWFLTMAX As Integer             ' �i�v�e�k�^�C�����
    HWFLTSPH As String * 1          ' �i�v�e�k�^�C������ʒu�Q��
    HWFLTSPT As String * 1          ' �i�v�e�k�^�C������ʒu�Q�_
    HWFLTSPI As String * 1          ' �i�v�e�k�^�C������ʒu�Q��
    HWFLTHWT As String * 1          ' �i�v�e�k�^�C���ۏؕ��@�Q��
    HWFLTHWS As String * 1          ' �i�v�e�k�^�C���ۏؕ��@�Q��
    HWFLTNSW As String * 2          ' �i�v�e�k�^�C���M�����@
    HWFLTKWY As String * 2          ' �i�v�e�k�^�C���������@
    HWFLTKHM As String * 1          ' �i�v�e�k�^�C�������p�x�Q��
    HWFLTKHN As String * 1          ' �i�v�e�k�^�C�������p�x�Q��
    HWFLTKHH As String * 1          ' �i�v�e�k�^�C�������p�x�Q��
    HWFLTKHU As String * 1          ' �i�v�e�k�^�C�������p�x�Q�E
    HWFLTMBP As Double              ' �i�v�e�k�^�C���ʓ����z
    HWFLTMCL As String * 1          ' �i�v�e�k�^�C���ʓ��v�Z
    HWFCNMIN As Double              ' �i�v�e�Y�f�Z�x����
    HWFCNMAX As Double              ' �i�v�e�Y�f�Z�x���
    HWFCNSPH As String * 1          ' �i�v�e�Y�f�Z�x����ʒu�Q��
    HWFCNSPT As String * 1          ' �i�v�e�Y�f�Z�x����ʒu�Q�_
    HWFCNSPI As String * 1          ' �i�v�e�Y�f�Z�x����ʒu�Q��
    HWFCNHWT As String * 1          ' �i�v�e�Y�f�Z�x�ۏؕ��@�Q��
    HWFCNHWS As String * 1          ' �i�v�e�Y�f�Z�x�ۏؕ��@�Q��
    HWFCNKWY As String * 2          ' �i�v�e�Y�f�Z�x�������@
    HWFCNKHM As String * 1          ' �i�v�e�Y�f�Z�x�����p�x�Q��
    HWFCNKHN As String * 1          ' �i�v�e�Y�f�Z�x�����p�x�Q��
    HWFCNKHH As String * 1          ' �i�v�e�Y�f�Z�x�����p�x�Q��
    HWFCNKHU As String * 1          ' �i�v�e�Y�f�Z�x�����p�x�Q�E
    HWFONMIN As Double              ' �i�v�e�_�f�Z�x����
    HWFONMAX As Double              ' �i�v�e�_�f�Z�x���
    HWFONSPH As String * 1          ' �i�v�e�_�f�Z�x����ʒu�Q��
    HWFONSPT As String * 1          ' �i�v�e�_�f�Z�x����ʒu�Q�_
    HWFONSPI As String * 1          ' �i�v�e�_�f�Z�x����ʒu�Q��
    HWFONHWT As String * 1          ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
    HWFONHWS As String * 1          ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
    HWFONKWY As String * 2          ' �i�v�e�_�f�Z�x�������@
    HWFONKHM As String * 1          ' �i�v�e�_�f�Z�x�����p�x�Q��
    HWFONKHN As String * 1          ' �i�v�e�_�f�Z�x�����p�x�Q��
    HWFONKHH As String * 1          ' �i�v�e�_�f�Z�x�����p�x�Q��
    HWFONKHU As String * 1          ' �i�v�e�_�f�Z�x�����p�x�Q�E
    HWFONMBP As Double              ' �i�v�e�_�f�Z�x�ʓ����z
    HWFONMCL As String * 1          ' �i�v�e�_�f�Z�x�ʓ��v�Z
    HWFONLTB As Double              ' �i�v�e�_�f�Z�x�k�s���z
    HWFONLTC As String * 1          ' �i�v�e�_�f�Z�x�k�s�v�Z
    HWFONSDV As Double              ' �i�v�e�_�f�Z�x�W���΍�
    HWFONAMN As Double              ' �i�v�e�_�f�Z�x���ω���
    HWFONAMX As Double              ' �i�v�e�_�f�Z�x���Ϗ��
    HWFOKBSH As String * 1          ' �i�v�e�_�f�U�敪����ʒu�Q��
    HWFOKBST As String * 1          ' �i�v�e�_�f�U�敪����ʒu�Q�_
    HWFOKBSI As String * 1          ' �i�v�e�_�f�U�敪����ʒu�Q��
    HWFOKBHT As String * 1          ' �i�v�e�_�f�U�敪�ۏؕ��@�Q��
    HWFOKBHS As String * 1          ' �i�v�e�_�f�U�敪�ۏؕ��@�Q��
    HWFOS1MN As Double              ' �i�v�e�_�f�͏o�P����
    HWFOS1MX As Double              ' �i�v�e�_�f�͏o�P���
    HWFOS1NS As String * 2          ' �i�v�e�_�f�͏o�P�M�����@
    HWFOS1SH As String * 1          ' �i�v�e�_�f�͏o�P����ʒu�Q��
    HWFOS1ST As String * 1          ' �i�v�e�_�f�͏o�P����ʒu�Q�_
    HWFOS1SI As String * 1          ' �i�v�e�_�f�͏o�P����ʒu�Q��
    HWFOS1HT As String * 1          ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
    HWFOS1HS As String * 1          ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
    HWFOS1HM As String * 1          ' �i�v�e�_�f�͏o�P�����p�x�Q��
    HWFOS1KN As String * 1          ' �i�v�e�_�f�͏o�P�����p�x�Q��
    HWFOS1KH As String * 1          ' �i�v�e�_�f�͏o�P�����p�x�Q��
    HWFOS1KU As String * 1          ' �i�v�e�_�f�͏o�P�����p�x�Q�E
    HWFOS2MN As Double              ' �i�v�e�_�f�͏o�Q����
    HWFOS2MX As Double              ' �i�v�e�_�f�͏o�Q���
    HWFOS2NS As String * 2          ' �i�v�e�_�f�͏o�Q�M�����@
    HWFOS2SH As String * 1          ' �i�v�e�_�f�͏o�Q����ʒu�Q��
    HWFOS2ST As String * 1          ' �i�v�e�_�f�͏o�Q����ʒu�Q�_
    HWFOS2SI As String * 1          ' �i�v�e�_�f�͏o�Q����ʒu�Q��
    HWFOS2HT As String * 1          ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
    HWFOS2HS As String * 1          ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
    HWFOS2KM As String * 1          ' �i�v�e�_�f�͏o�Q�����p�x�Q��
    HWFOS2KN As String * 1          ' �i�v�e�_�f�͏o�Q�����p�x�Q��
    HWFOS2KH As String * 1          ' �i�v�e�_�f�͏o�Q�����p�x�Q��
    HWFOS2KU As String * 1          ' �i�v�e�_�f�͏o�Q�����p�x�Q�E
    HWFOS3MN As Double              ' �i�v�e�_�f�͏o�R����
    HWFOS3MX As Double              ' �i�v�e�_�f�͏o�R���
    HWFOS3NS As String * 2          ' �i�v�e�_�f�͏o�R�M�����@
    HWFOS3SH As String * 1          ' �i�v�e�_�f�͏o�R����ʒu�Q��
    HWFOS3ST As String * 1          ' �i�v�e�_�f�͏o�R����ʒu�Q�_
    HWFOS3SI As String * 1          ' �i�v�e�_�f�͏o�R����ʒu�Q��
    HWFOS3HT As String * 1          ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��
    HWFOS3HS As String * 1          ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��
    HWFOS3KM As String * 1          ' �i�v�e�_�f�͏o�R�����p�x�Q��
    HWFOS3KN As String * 1          ' �i�v�e�_�f�͏o�R�����p�x�Q��
    HWFOS3KH As String * 1          ' �i�v�e�_�f�͏o�R�����p�x�Q��
    HWFOS3KU As String * 1          ' �i�v�e�_�f�͏o�R�����p�x�Q�E
    HWFANTNP As Integer             ' �i�v�e�`�m���x
    HWFANTIM As Integer             ' �i�v�e�`�m����
    HWFANTMN As Integer             ' �i�v�e�`�m���ԉ���
    HWFANTMX As Integer             ' �i�v�e�`�m���ԏ��
    HWFZOMIN As Double              ' �i�v�e�c���_�f����
    HWFZOMAX As Double              ' �i�v�e�c���_�f���
    HWFZOSPH As String * 1          ' �i�v�e�c���_�f����ʒu�Q��
    HWFZOSPT As String * 1          ' �i�v�e�c���_�f����ʒu�Q�_
    HWFZOSPI As String * 1          ' �i�v�e�c���_�f����ʒu�Q��
    HWFZOHWT As String * 1          ' �i�v�e�c���_�f�ۏؕ��@�Q��
    HWFZOHWS As String * 1          ' �i�v�e�c���_�f�ۏؕ��@�Q��
    HWFZONSW As String * 2          ' �i�v�e�c���_�f�M�����@
    HWFZOKWY As String * 2          ' �i�v�e�c���_�f�������@
    HWFZOKHM As String * 1          ' �i�v�e�c���_�f�����p�x�Q��
    HWFZOKHN As String * 1          ' �i�v�e�c���_�f�����p�x�Q��
    HWFZOKHH As String * 1          ' �i�v�e�c���_�f�����p�x�Q��
    HWFZOKHU As String * 1          ' �i�v�e�c���_�f�����p�x�Q�E
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
' �ǉ� 2003/09.11 SystemBrain Start
    HWFTMMAXN As Double             ' �i�v�e�]�ʖ��x���
    HWFANTTAN As String * 1         ' �i�v�e�`�m���ԒP��
' �ǉ� 2003/09.11 SystemBrain End
End Type


' ���i�d�lWF�ް��U
Public Type typ_TBCME026
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGWFSNO As String * 6          ' �i�Ǘ��v�e���i�ԍ�
    HMGWFSNE As Integer             ' �i�Ǘ��v�e���i�ԍ��}��
    HWFBDOMN As Integer             ' �i�v�e�a�c�n�r�e����
    HWFBDOMX As Integer             ' �i�v�e�a�c�n�r�e���
    HWFBDOSH As String * 1          ' �i�v�e�a�c�n�r�e����ʒu�Q��
    HWFBDOST As String * 1          ' �i�v�e�a�c�n�r�e����ʒu�Q�_
    HWFBDOSR As String * 1          ' �i�v�e�a�c�n�r�e����ʒu�Q��
    HWFBDOHT As String * 1          ' �i�v�e�a�c�n�r�e�ۏؕ��@�Q��
    HWFBDOHS As String * 1          ' �i�v�e�a�c�n�r�e�ۏؕ��@�Q��
    HWFBDOSZ As String * 1          ' �i�v�e�a�c�n�r�e�������
    HWFBDONS As String * 2          ' �i�v�e�a�c�n�r�e�M�����@
    HWFBDOKM As String * 1          ' �i�v�e�a�c�n�r�e�����p�x�Q��
    HWFBDOKN As String * 1          ' �i�v�e�a�c�n�r�e�����p�x�Q��
    HWFBDOKH As String * 1          ' �i�v�e�a�c�n�r�e�����p�x�Q��
    HWFBDOKU As String * 1          ' �i�v�e�a�c�n�r�e�����p�x�Q�E
    HWFBDOET As Integer             ' �i�v�e�a�c�n�r�e�I���d�s��
    HWFBDSMN As Integer             ' �i�v�e�a�c�r�s�Չ���
    HWFBDSMX As Integer             ' �i�v�e�a�c�r�s�Տ��
    HWFBDSSH As String * 1          ' �i�v�e�a�c�r�s�Ց���ʒu�Q��
    HWFBDSST As String * 1          ' �i�v�e�a�c�r�s�Ց���ʒu�Q�_
    HWFBDSSR As String * 1          ' �i�v�e�a�c�r�s�Ց���ʒu�Q��
    HWFBDSHT As String * 1          ' �i�v�e�a�c�r�s�Օۏؕ��@�Q��
    HWFBDSHS As String * 1          ' �i�v�e�a�c�r�s�Օۏؕ��@�Q��
    HWFBDSSZ As String * 1          ' �i�v�e�a�c�r�s�Ց������
    HWFBDSNS As String * 2          ' �i�v�e�a�c�r�s�ՔM�����@
    HWFBDSKM As String * 1          ' �i�v�e�a�c�r�s�Ռ����p�x�Q��
    HWFBDSKN As String * 1          ' �i�v�e�a�c�r�s�Ռ����p�x�Q��
    HWFBDSKH As String * 1          ' �i�v�e�a�c�r�s�Ռ����p�x�Q��
    HWFBDSKU As String * 1          ' �i�v�e�a�c�r�s�Ռ����p�x�Q�E
    HWFBDSET As Integer             ' �i�v�e�a�c�r�s�ՑI���d�s��
    HWFRNFMX As Double              ' �i�v�e���t�l�X�\���
    HWFRNFSH As String * 1          ' �i�v�e���t�l�X�\����ʒu�Q��
    HWFRNFST As String * 1          ' �i�v�e���t�l�X�\����ʒu�Q�_
    HWFRNFSI As String * 1          ' �i�v�e���t�l�X�\����ʒu�Q��
    HWFRNFKW As String * 2          ' �i�v�e���t�l�X�\�������@
    HWFRNFZA As Integer             ' �i�v�e���t�l�X�\���O�̈�
    HWFRNBMX As Double              ' �i�v�e���t�l�X�����
    HWFRNBSH As String * 1          ' �i�v�e���t�l�X������ʒu�Q��
    HWFRNBST As String * 1          ' �i�v�e���t�l�X������ʒu�Q�_
    HWFRNBSI As String * 1          ' �i�v�e���t�l�X������ʒu�Q��
    HWFRNBKW As String * 2          ' �i�v�e���t�l�X���������@
    HWFRNBZA As Integer             ' �i�v�e���t�l�X�����O�̈�
    HWFDENKU As String * 1          ' �i�v�e�c���������L��
    HWFDENMX As Integer             ' �i�v�e�c�������
    HWFDENMN As Integer             ' �i�v�e�c��������
    HWFDENHT As String * 1          ' �i�v�e�c�����ۏؕ��@�Q��
    HWFDENHS As String * 1          ' �i�v�e�c�����ۏؕ��@�Q��
    HWFDVDKU As String * 1          ' �i�v�e�c�u�c�Q�����L��
    HWFDVDMX As Integer             ' �i�v�e�c�u�c�Q���
    HWFDVDMN As Integer             ' �i�v�e�c�u�c�Q����
    HWFDVDHT As String * 1          ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��
    HWFDVDHS As String * 1          ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��
    HWFLDLKU As String * 1          ' �i�v�e�k�^�c�k�����L��
    HWFLDLMX As Integer             ' �i�v�e�k�^�c�k���
    HWFLDLMN As Integer             ' �i�v�e�k�^�c�k����
    HWFLDLHT As String * 1          ' �i�v�e�k�^�c�k�ۏؕ��@�Q��
    HWFLDLHS As String * 1          ' �i�v�e�k�^�c�k�ۏؕ��@�Q��
    HWFGDSPH As String * 1          ' �i�v�e�f�c����ʒu�Q��
    HWFGDSPT As String * 1          ' �i�v�e�f�c����ʒu�Q�_
    HWFGDSPR As String * 1          ' �i�v�e�f�c����ʒu�Q��
    HWFGDSZY As String * 1          ' �i�v�e�f�c�������
    HWFGDZAR As Integer             ' �i�v�e�f�c���O�̈�
    HWFGDKHM As String * 1          ' �i�v�e�f�c�����p�x�Q��
    HWFGDKHN As String * 1          ' �i�v�e�f�c�����p�x�Q��
    HWFGDKHH As String * 1          ' �i�v�e�f�c�����p�x�Q��
    HWFGDKHU As String * 1          ' �i�v�e�f�c�����p�x�Q�E
    HWFDSOKE As String * 1          ' �i�v�e�c�r�n�c����
    HWFDSOMX As Long                ' �i�v�e�c�r�n�c���
    HWFDSOMN As Long                ' �i�v�e�c�r�n�c����
    HWFDSOAX As Integer             ' �i�v�e�c�r�n�c�̈���
    HWFDSOAN As Integer             ' �i�v�e�c�r�n�c�̈扺��
    HWFDSOHT As String * 1          ' �i�v�e�c�r�n�c�ۏؕ��@�Q��
    HWFDSOHS As String * 1          ' �i�v�e�c�r�n�c�ۏؕ��@�Q��
    HWFDSOKM As String * 1          ' �i�v�e�c�r�n�c�����p�x�Q��
    HWFDSOKN As String * 1          ' �i�v�e�c�r�n�c�����p�x�Q��
    HWFDSOKH As String * 1          ' �i�v�e�c�r�n�c�����p�x�Q��
    HWFDSOKU As String * 1          ' �i�v�e�c�r�n�c�����p�x�Q�E
    HWFNTPUM As String * 1          ' �i�v�e���R�i�m�g�|�L��
    HWFNTPK1 As Double              ' �i�v�e���R�i�m�g�|�K�i�P
    HWFNTPP1 As Double              ' �i�v�e���R�i�m�g�|�o�t�`�P
    HWFNTPS1 As Double              ' �i�v�e���R�i�m�g�|�T�C�g�P
    HWFNTPK2 As Double              ' �i�v�e���R�i�m�g�|�K�i�Q
    HWFNTPP2 As Double              ' �i�v�e���R�i�m�g�|�o�t�`�Q
    HWFNTPS2 As Double              ' �i�v�e���R�i�m�g�|�T�C�g�Q
    HWFNTPK3 As Double              ' �i�v�e���R�i�m�g�|�K�i�R
    HWFNTPP3 As Double              ' �i�v�e���R�i�m�g�|�o�t�`�R
    HWFNTPS3 As Double              ' �i�v�e���R�i�m�g�|�T�C�g�R
    HWFNTPZA As Integer             ' �i�v�e���R�i�m�g�|���O�̈�
    HWFNTPHT As String * 1          ' �i�v�e���R�i�m�g�|�ۏؕ��@�Q��
    HWFNTPHS As String * 1          ' �i�v�e���R�i�m�g�|�ۏؕ��@�Q��
    HWFNTPKM As String * 1          ' �i�v�e���R�i�m�g�|�����p�x�Q��
    HWFNTPKN As String * 1          ' �i�v�e���R�i�m�g�|�����p�x�Q��
    HWFNTPKH As String * 1          ' �i�v�e���R�i�m�g�|�����p�x�Q��
    HWFNTPKU As String * 1          ' �i�v�e���R�i�m�g�|�����p�x�Q�E
    HWFCRSSK As String * 1          ' �i�v�e���R�N���X�r�r����
    HWFMDCEN As Double              ' �i�v�e���R�ʃ_�����፷���S
    HWFMDMAX As Double              ' �i�v�e���R�ʃ_�����፷���
    HWFMDMIN As Double              ' �i�v�e���R�ʃ_�����፷����
    HWFMDSPH As String * 1          ' �i�v�e���R�ʃ_������ʒu�Q��
    HWFMDSPT As String * 1          ' �i�v�e���R�ʃ_������ʒu�Q�_
    HWFMDSPI As String * 1          ' �i�v�e���R�ʃ_������ʒu�Q��
    HWFMDHWT As String * 1          ' �i�v�e���R�ʃ_���ۏؕ��@�Q��
    HWFMDHWS As String * 1          ' �i�v�e���R�ʃ_���ۏؕ��@�Q��
    HWFMDKHM As String * 1          ' �i�v�e���R�ʃ_�������p�x�Q��
    HWFMDKHN As String * 1          ' �i�v�e���R�ʃ_�������p�x�Q��
    HWFMDKHH As String * 1          ' �i�v�e���R�ʃ_�������p�x�Q��
    HWFMDKHU As String * 1          ' �i�v�e���R�ʃ_�������p�x�Q�E
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
' �ǉ� 2003/09.11 SystemBrain Start
    HWFDVDMXN As Integer            ' �iWFDVD2���
    HWFDVDMNN As Integer            ' �iWFDVD2����
    HWFDSONWY As String * 2         ' �iWFDSOD�M�����@
    HWFMSUMX As Integer             ' �iWFM�X�N���b�`���
    HWFMSUZY As String * 1          ' �iWFM�X�N���b�`�������
    HWFMSUKW As String * 1          ' �iWFM�X�N���b�`�������@
    HWFMSUSZ As Double              ' �iWFM�X�N���b�`�T�C�Y
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    sStaffID As String * 8          ' ���F�Ј�ID
    SYNFLAG As String * 1           ' ���F�t���O
    SYNDATE As Date                 ' ���F���t
    HWFNP1AR As Double              ' �iWF�i�m�g�|1�G���A
    HWFNP1MAX As Double             ' �iWF�i�m�g�|1���
    HWFNP2AR As Double              ' �iWF�i�m�g�|2�G���A
    HWFNP2MAX As Double             ' �iWF�i�m�g�|2���
' �ǉ� 2003/09.11 SystemBrain End
End Type


' ���i�d�lWF�ް��V
Public Type typ_TBCME027
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGWFSNO As String * 6          ' �i�Ǘ��v�e���i�ԍ�
    HMGWFSNE As Integer             ' �i�Ǘ��v�e���i�ԍ��}��
    HWFSMIN As Double               ' �i�v�e���艺��
    HWFSMAX As Double               ' �i�v�e������
    HWFSHWYT As String * 1          ' �i�v�e����ۏؕ��@�Q��
    HWFSHWYS As String * 1          ' �i�v�e����ۏؕ��@�Q��
    HWFSKWAY As String * 2          ' �i�v�e���茟�����@
    HWFSKHM As String * 1           ' �i�v�e���茟���p�x�Q��
    HWFSKHN As String * 1           ' �i�v�e���茟���p�x�Q��
    HWFSKHH As String * 1           ' �i�v�e���茟���p�x�Q��
    HWFSKHU As String * 1           ' �i�v�e���茟���p�x�Q�E
    HWFSSZYO As String * 1          ' �i�v�e���葪�����
    'HWFSZARA As Integer             ' �i�v�e���菜�O�̈�
    HWFSZARAN As Double             ' �i�v�e���菜�O�̈� 6/22 Yam
    HWFSSDEV As Double              ' �i�v�e����W���΍�
    HWFSAMIN As Double              ' �i�v�e���蕽�ω���
    HWFSAMAX As Double              ' �i�v�e���蕽�Ϗ��
    HWFSSREC As String * 1          ' �i�v�e���葪���
    HWFSBO1 As Double               ' �i�v�e���苫�E�P
    HWFSBO1B As Integer             ' �i�v�e���苫�E�P��
    HWFSBO2 As Double               ' �i�v�e���苫�E�Q
    HWFSBO2B As Integer             ' �i�v�e���苫�E�Q��
    HWFSBO3 As Double               ' �i�v�e���苫�E�R
    HWFSBO3B As Integer             ' �i�v�e���苫�E�R��
    HWFWARMX As Double              ' �i�v�e�v�`�q�o���
    HWFWARSZ As String * 1          ' �i�v�e�v�`�q�o�������
    HWFWARHT As String * 1          ' �i�v�e�v�`�q�o�ۏؕ��@�Q��
    HWFWARHS As String * 1          ' �i�v�e�v�`�q�o�ۏؕ��@�Q��
    HWFWARKW As String * 2          ' �i�v�e�v�`�q�o�������@
    'HWFWARZA As Integer             ' �i�v�e�v�`�q�o���O�̈�
    HWFWARZAN As Double             ' �i�v�e�v�`�q�o���O�̈� 6/22 Yam
    HWFWARKM As String * 1          ' �i�v�e�v�`�q�o�����p�x�Q��
    HWFWARKN As String * 1          ' �i�v�e�v�`�q�o�����p�x�Q��
    HWFWARKH As String * 1          ' �i�v�e�v�`�q�o�����p�x�Q��
    HWFWARKU As String * 1          ' �i�v�e�v�`�q�o�����p�x�Q�E
    HWFWARSR As String * 1          ' �i�v�e�v�`�q�o�����
    HWFWAB1 As Double               ' �i�v�e�v�`�q�o���E�P
    HWFWAB1B As Integer             ' �i�v�e�v�`�q�o���E�P��
    HWFWAB2 As Double               ' �i�v�e�v�`�q�o���E�Q
    HWFWAB2B As Integer             ' �i�v�e�v�`�q�o���E�Q��
    HWFWAB3 As Double               ' �i�v�e�v�`�q�o���E�R
    HWFWAB3B As Integer             ' �i�v�e�v�`�q�o���E�R��
    HWFWARPR As String * 1          ' �i�v�e�v�����������N
    HWFFSZYO As String * 1          ' �i�v�e���R�������
    HWFFSREC As String * 1          ' �i�v�e���R�����
    HWFGBMAX As Double              ' �i�v�e���R�f�a���
    HWFGBPUG As Double              ' �i�v�e���R�f�a�o�t�`��
    HWFGBPUR As Integer             ' �i�v�e���R�f�a�o�t�`��
    HWFGBHWT As String * 1          ' �i�v�e���R�f�a�ۏؕ��@�Q��
    HWFGBHWS As String * 1          ' �i�v�e���R�f�a�ۏؕ��@�Q��
    HWFGBKW As String * 4           ' �i�v�e���R�f�a�������@
    'HWFGBZAR As Integer             ' �i�v�e���R�f�a���O�̈�
    HWFGBZARN As Double             ' �i�v�e���R�f�a���O�̈�
    HWFGBKHM As String * 1          ' �i�v�e���R�f�a�����p�x�Q��
    HWFGBKHN As String * 1          ' �i�v�e���R�f�a�����p�x�Q��
    HWFGBKHH As String * 1          ' �i�v�e���R�f�a�����p�x�Q��
    HWFGBKHU As String * 1          ' �i�v�e���R�f�a�����p�x�Q�E
    HWFGBB1 As Double               ' �i�v�e���R�f�a���E�P
    HWFGBB1B As Integer             ' �i�v�e���R�f�a���E�P��
    HWFGBB2 As Double               ' �i�v�e���R�f�a���E�Q
    HWFGBB2B As Integer             ' �i�v�e���R�f�a���E�Q��
    HWFGBB3 As Double               ' �i�v�e���R�f�a���E�R
    HWFGBB3B As Integer             ' �i�v�e���R�f�a���E�R��
    HWFGFDMX As Double              ' �i�v�e���R�f�e�c���
    HWFGFDPG As Double              ' �i�v�e���R�f�e�c�o�t�`��
    HWFGFDPR As Integer             ' �i�v�e���R�f�e�c�o�t�`��
    HWFGFDHT As String * 1          ' �i�v�e���R�f�e�c�ۏؕ��@�Q��
    HWFGFDHS As String * 1          ' �i�v�e���R�f�e�c�ۏؕ��@�Q��
    HWFGFDKW As String * 4          ' �i�v�e���R�f�e�c�������@
    'HWFGFDZA As Integer             ' �i�v�e���R�f�e�c���O�̈�
    HWFGFDZAN As Double             ' �i�v�e���R�f�e�c���O�̈�
    HWFGFDKM As String * 1          ' �i�v�e���R�f�e�c�����p�x�Q��
    HWFGFDKN As String * 1          ' �i�v�e���R�f�e�c�����p�x�Q��
    HWFGFDKH As String * 1          ' �i�v�e���R�f�e�c�����p�x�Q��
    HWFGFDKU As String * 1          ' �i�v�e���R�f�e�c�����p�x�Q�E
    HWFGDB1 As Double               ' �i�v�e���R�f�e�c���E�P
    HWFGDB1B As Integer             ' �i�v�e���R�f�e�c���E�P��
    HWFGDB2 As Double               ' �i�v�e���R�f�e�c���E�Q
    HWFGDB2B As Integer             ' �i�v�e���R�f�e�c���E�Q��
    HWFGDB3 As Double               ' �i�v�e���R�f�e�c���E�R
    HWFGDB3B As Integer             ' �i�v�e���R�f�e�c���E�R��
    HWFGFRMX As Double              ' �i�v�e���R�f�e�q���
    HWFGFRPG As Double              ' �i�v�e���R�f�e�q�o�t�`��
    HWFGFRPR As Integer             ' �i�v�e���R�f�e�q�o�t�`��
    HWFGFRHT As String * 1          ' �i�v�e���R�f�e�q�ۏؕ��@�Q��
    HWFGFRHS As String * 1          ' �i�v�e���R�f�e�q�ۏؕ��@�Q��
    HWFGFRKW As String * 4          ' �i�v�e���R�f�e�q�������@
    'HWFGFRZA As Integer             ' �i�v�e���R�f�e�q���O�̈�
    HWFGFRZAN As Double             ' �i�v�e���R�f�e�q���O�̈�
    HWFGFRKM As String * 1          ' �i�v�e���R�f�e�q�����p�x�Q��
    HWFGFRKN As String * 1          ' �i�v�e���R�f�e�q�����p�x�Q��
    HWFGFRKH As String * 1          ' �i�v�e���R�f�e�q�����p�x�Q��
    HWFGFRKU As String * 1          ' �i�v�e���R�f�e�q�����p�x�Q�E
    HWFGRB1 As Double               ' �i�v�e���R�f�e�q���E�P
    HWFGRB1B As Integer             ' �i�v�e���R�f�e�q���E�P��
    HWFGRB2 As Double               ' �i�v�e���R�f�e�q���E�Q
    HWFGRB2B As Integer             ' �i�v�e���R�f�e�q���E�Q��
    HWFGRB3 As Double               ' �i�v�e���R�f�e�q���E�R
    HWFGRB3B As Integer             ' �i�v�e���R�f�e�q���E�R��
    HWFSBMAX As Double              ' �i�v�e���R�r�a���
    HWFSBPUG As Double              ' �i�v�e���R�r�a�o�t�`��
    HWFSBPUR As Integer             ' �i�v�e���R�r�a�o�t�`��
    HWFSBSZX As Double              ' �i�v�e���R�r�a�T�C�Y�w
    HWFSBSZY As Double              ' �i�v�e���R�r�a�T�C�Y�x
    HWFSBHWT As String * 1          ' �i�v�e���R�r�a�ۏؕ��@�Q��
    HWFSBHWS As String * 1          ' �i�v�e���R�r�a�ۏؕ��@�Q��
    HWFSBKW As String * 4           ' �i�v�e���R�r�a�������@
    'HWFSBZAR As Integer             ' �i�v�e���R�r�a���O�̈�
    HWFSBZARN As Double             ' �i�v�e���R�r�a���O�̈�
    HWFSBKHM As String * 1          ' �i�v�e���R�r�a�����p�x�Q��
    HWFSBKHN As String * 1          ' �i�v�e���R�r�a�����p�x�Q��
    HWFSBKHH As String * 1          ' �i�v�e���R�r�a�����p�x�Q��
    HWFSBKHU As String * 1          ' �i�v�e���R�r�a�����p�x�Q�E
    HWFSBB1 As Double               ' �i�v�e���R�r�a���E�P
    HWFSBB1B As Integer             ' �i�v�e���R�r�a���E�P��
    HWFSBB2 As Double               ' �i�v�e���R�r�a���E�Q
    HWFSBB2B As Integer             ' �i�v�e���R�r�a���E�Q��
    HWFSBB3 As Double               ' �i�v�e���R�r�a���E�R
    HWFSBB3B As Integer             ' �i�v�e���R�r�a���E�R��
    HWFSFMAX As Double              ' �i�v�e���R�r�e���
    HWFSFPUG As Double              ' �i�v�e���R�r�e�o�t�`��
    HWFSFPUR As Integer             ' �i�v�e���R�r�e�o�t�`��
    HWFSFSZX As Double              ' �i�v�e���R�r�e�T�C�Y�w
    HWFSFSZY As Double              ' �i�v�e���R�r�e�T�C�Y�x
    HWFSFHWT As String * 1          ' �i�v�e���R�r�e�ۏؕ��@�Q��
    HWFSFHWS As String * 1          ' �i�v�e���R�r�e�ۏؕ��@�Q��
    HWFSFKW As String * 4           ' �i�v�e���R�r�e�������@
    'HWFSFZAR As Integer             ' �i�v�e���R�r�e���O�̈�
    HWFSFZARN As Double             ' �i�v�e���R�r�e���O�̈�
    HWFSFKHM As String * 1          ' �i�v�e���R�r�e�����p�x�Q��
    HWFSFKHN As String * 1          ' �i�v�e���R�r�e�����p�x�Q��
    HWFSFKHH As String * 1          ' �i�v�e���R�r�e�����p�x�Q��
    HWFSFKHU As String * 1          ' �i�v�e���R�r�e�����p�x�Q�E
    HWFSFB1 As Double               ' �i�v�e���R�r�e���E�P
    HWFSFB1B As Integer             ' �i�v�e���R�r�e���E�P��
    HWFSFB2 As Double               ' �i�v�e���R�r�e���E�Q
    HWFSFB2B As Integer             ' �i�v�e���R�r�e���E�Q��
    HWFSFB3 As Double               ' �i�v�e���R�r�e���E�R
    HWFSFB3B As Integer             ' �i�v�e���R�r�e���E�R��
    HWFFSXOF As Double              ' �i�v�e���R�T�C�g�w�n�e
    HWFFSYOF As Double              ' �i�v�e���R�T�C�g�x�n�e
    HWFFPSUM As String * 1          ' �i�v�e���R�o�T�C�g�L��
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
' �ǉ� 2003/09.11 SystemBrain Start
    HWFSBPUAGN As Double            ' �iWF���RSBPUA��
    HWFSBMAXN As Double             ' �iWF���RSB���
    HWFSBB1N As Double              ' �iWF���RSB���E1
    HWFSBB2N As Double              ' �iWF���RSB���E2
    HWFSBB3N As Double              ' �iWF���RSB���E3
    HWFSFPUAGN As Double            ' �iWF���RSFPUA��
    HWFSFMAXN As Double             ' �iWF���RSF���
    HWFSFB1N As Double              ' �iWF���RSF���E1
    HWFSFB2N As Double              ' �iWF���RSF���E2
    HWFSFB3N As Double              ' �iWF���RSF���E3
' �ǉ� 2003/09.11 SystemBrain End
End Type


' ���i�d�lWF�ް��W
Public Type typ_TBCME028
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGWFSNO As String * 6          ' �i�Ǘ��v�e���i�ԍ�
    HMGWFSNE As Integer             ' �i�Ǘ��v�e���i�ԍ��}��
    HWFMK1SI As Double              ' �i�v�e�ʌ����ׂP�T�C�Y
    HWFMK1MX As Integer             ' �i�v�e�ʌ����ׂP���
    HWFMK1SZ As String * 1          ' �i�v�e�ʌ����ׂP�������
    HWFMK1ZA As Integer             ' �i�v�e�ʌ����ׂP���O�̈�
    HWFMK1HT As String * 1          ' �i�v�e�ʌ����ׂP�ۏؕ��@�Q��
    HWFMK1HS As String * 1          ' �i�v�e�ʌ����ׂP�ۏؕ��@�Q��
    HWFMK1KM As String * 1          ' �i�v�e�ʌ����ׂP�����p�x�Q��
    HWFMK1KN As String * 1          ' �i�v�e�ʌ����ׂP�����p�x�Q��
    HWFMK1KH As String * 1          ' �i�v�e�ʌ����ׂP�����p�x�Q��
    HWFMK1KU As String * 1          ' �i�v�e�ʌ����ׂP�����p�x�Q�E
    HWFM1B1 As Integer              ' �i�v�e�ʌ����ׂP���E�P
    HWFM1B1B As Integer             ' �i�v�e�ʌ����ׂP���E�P��
    HWFM1B2 As Integer              ' �i�v�e�ʌ����ׂP���E�Q
    HWFM1B2B As Integer             ' �i�v�e�ʌ����ׂP���E�Q��
    HWFM1B3 As Integer              ' �i�v�e�ʌ����ׂP���E�R
    HWFM1B3B As Integer             ' �i�v�e�ʌ����ׂP���E�R��
    HWFMK2SI As Double              ' �i�v�e�ʌ����ׂQ�T�C�Y
    HWFMK2MX As Integer             ' �i�v�e�ʌ����ׂQ���
    HWFMK2HT As String * 1          ' �i�v�e�ʌ����ׂQ�ۏؕ��@�Q��
    HWFMK2HS As String * 1          ' �i�v�e�ʌ����ׂQ�ۏؕ��@�Q��
    HWFMK2KM As String * 1          ' �i�v�e�ʌ����ׂQ�����p�x�Q��
    HWFMK2KN As String * 1          ' �i�v�e�ʌ����ׂQ�����p�x�Q��
    HWFMK2KH As String * 1          ' �i�v�e�ʌ����ׂQ�����p�x�Q��
    HWFMK2KU As String * 1          ' �i�v�e�ʌ����ׂQ�����p�x�Q�E
    HWFM2B1 As Integer              ' �i�v�e�ʌ����ׂQ���E�P
    HWFM2B1B As Integer             ' �i�v�e�ʌ����ׂQ���E�P��
    HWFM2B2 As Integer              ' �i�v�e�ʌ����ׂQ���E�Q
    HWFM2B2B As Integer             ' �i�v�e�ʌ����ׂQ���E�Q��
    HWFM2B3 As Integer              ' �i�v�e�ʌ����ׂQ���E�R
    HWFM2B3B As Integer             ' �i�v�e�ʌ����ׂQ���E�R��
    HWFMK3SI As Double              ' �i�v�e�ʌ����ׂR�T�C�Y
    HWFMK3MX As Integer             ' �i�v�e�ʌ����ׂR���
    HWFMK3HT As String * 1          ' �i�v�e�ʌ����ׂR�ۏؕ��@�Q��
    HWFMK3HS As String * 1          ' �i�v�e�ʌ����ׂR�ۏؕ��@�Q��
    HWFMK3KM As String * 1          ' �i�v�e�ʌ����ׂR�����p�x�Q��
    HWFMK3KN As String * 1          ' �i�v�e�ʌ����ׂR�����p�x�Q��
    HWFMK3KH As String * 1          ' �i�v�e�ʌ����ׂR�����p�x�Q��
    HWFMK3KU As String * 1          ' �i�v�e�ʌ����ׂR�����p�x�Q�E
    HWFM3B1 As Integer              ' �i�v�e�ʌ����ׂR���E�P
    HWFM3B1B As Integer             ' �i�v�e�ʌ����ׂR���E�P��
    HWFM3B2 As Integer              ' �i�v�e�ʌ����ׂR���E�Q
    HWFM3B2B As Integer             ' �i�v�e�ʌ����ׂR���E�Q��
    HWFM3B3 As Integer              ' �i�v�e�ʌ����ׂR���E�R
    HWFM3B3B As Integer             ' �i�v�e�ʌ����ׂR���E�R��
    HWFMK4SI As Double              ' �i�v�e�ʌ����ׂS�T�C�Y
    HWFMK4MX As Integer             ' �i�v�e�ʌ����ׂS���
    HWFMK4HT As String * 1          ' �i�v�e�ʌ����ׂS�ۏؕ��@�Q��
    HWFMK4HS As String * 1          ' �i�v�e�ʌ����ׂS�ۏؕ��@�Q��
    HWFMK4KM As String * 1          ' �i�v�e�ʌ����ׂS�����p�x�Q��
    HWFMK4KN As String * 1          ' �i�v�e�ʌ����ׂS�����p�x�Q��
    HWFMK4KH As String * 1          ' �i�v�e�ʌ����ׂS�����p�x�Q��
    HWFMK4KU As String * 1          ' �i�v�e�ʌ����ׂS�����p�x�Q�E
    HWFM4B1 As Integer              ' �i�v�e�ʌ����ׂS���E�P
    HWFM4B1B As Integer             ' �i�v�e�ʌ����ׂS���E�P��
    HWFM4B2 As Integer              ' �i�v�e�ʌ����ׂS���E�Q
    HWFM4B2B As Integer             ' �i�v�e�ʌ����ׂS���E�Q��
    HWFM4B3 As Integer              ' �i�v�e�ʌ����ׂS���E�R
    HWFM4B3B As Integer             ' �i�v�e�ʌ����ׂS���E�R��
    HWFMB1SI As Double              ' �i�v�e�ʌ����ח��P�T�C�Y
    HWFMB1MX As Integer             ' �i�v�e�ʌ����ח��P���
    HWFMB1SZ As String * 1          ' �i�v�e�ʌ����ח��P�������
    HWFMB1ZA As Integer             ' �i�v�e�ʌ����ח��P���O�̈�
    HWFMB1HT As String * 1          ' �i�v�e�ʌ����ח��P�ۏؕ��@�Q��
    HWFMB1HS As String * 1          ' �i�v�e�ʌ����ח��P�ۏؕ��@�Q��
    HWFMB1KM As String * 1          ' �i�v�e�ʌ����ח��P�����p�x�Q��
    HWFMB1KN As String * 1          ' �i�v�e�ʌ����ח��P�����p�x�Q��
    HWFMB1KH As String * 1          ' �i�v�e�ʌ����ח��P�����p�x�Q��
    HWFMB1KU As String * 1          ' �i�v�e�ʌ����ח��P�����p�x�Q�E
    HWFMB2SI As Double              ' �i�v�e�ʌ����ח��Q�T�C�Y
    HWFMB2MX As Integer             ' �i�v�e�ʌ����ח��Q���
    HWFMB2SZ As String * 1          ' �i�v�e�ʌ����ח��Q�������
    HWFMB2ZA As Integer             ' �i�v�e�ʌ����ח��Q���O�̈�
    HWFMB2HT As String * 1          ' �i�v�e�ʌ����ח��Q�ۏؕ��@�Q��
    HWFMB2HS As String * 1          ' �i�v�e�ʌ����ח��Q�ۏؕ��@�Q��
    HWFMB2KM As String * 1          ' �i�v�e�ʌ����ח��Q�����p�x�Q��
    HWFMB2KN As String * 1          ' �i�v�e�ʌ����ח��Q�����p�x�Q��
    HWFMB2KH As String * 1          ' �i�v�e�ʌ����ח��Q�����p�x�Q��
    HWFMB2KU As String * 1          ' �i�v�e�ʌ����ח��Q�����p�x�Q�E
    HWFMKSRE As String * 1          ' �i�v�e�ʌ����ב����
    HWFMKKW As String * 1           ' �i�v�e�ʌ����׌������@
    HWFMPIPT As String * 1          ' �i�v�e�ʌ����ׂo�h�o����
    HWFMPIPK As Integer             ' �i�v�e�ʌ����ׂo�h�o��
    HWFMPISH As String * 1          ' �i�v�e�ʌ��o�h�o����ʒu�Q��
    HWFMPIST As String * 1          ' �i�v�e�ʌ��o�h�o����ʒu�Q�_
    HWFMPISI As String * 1          ' �i�v�e�ʌ��o�h�o����ʒu�Q��
    HWFMPIKM As String * 1          ' �i�v�e�ʌ��o�h�o�����p�x�Q��
    HWFMPIKN As String * 1          ' �i�v�e�ʌ��o�h�o�����p�x�Q��
    HWFMPIKH As String * 1          ' �i�v�e�ʌ��o�h�o�����p�x�Q��
    HWFMPIKU As String * 1          ' �i�v�e�ʌ��o�h�o�����p�x�Q�E
    HWFMNMAX As Double              ' �i�v�e�����Z�x���
    HWFMNALX As Double              ' �i�v�e�����Z�x�`�k���
    HWFMNCAX As Double              ' �i�v�e�����Z�x�b�`���
    HWFMNCRX As Double              ' �i�v�e�����Z�x�b�q���
    HWFMNCUX As Double              ' �i�v�e�����Z�x�b�t���
    HWFMNFEX As Double              ' �i�v�e�����Z�x�e�d���
    HWFMNKMX As Double              ' �i�v�e�����Z�x�j���
    HWFMNMGX As Double              ' �i�v�e�����Z�x�l�f���
    HWFMNNAX As Double              ' �i�v�e�����Z�x�m�`���
    HWFMNNIX As Double              ' �i�v�e�����Z�x�m�h���
    HWFMNZNX As Double              ' �i�v�e�����Z�x�y�m���
    HWFMNKWY As String * 2          ' �i�v�e�����Z�x�������@
    HWFMNSPH As String * 1          ' �i�v�e�����Z�x����ʒu�Q��
    HWFMNSPT As String * 1          ' �i�v�e�����Z�x����ʒu�Q�_
    HWFMNSPI As String * 1          ' �i�v�e�����Z�x����ʒu�Q��
    HWFMNHWT As String * 1          ' �i�v�e�����Z�x�ۏؕ��@�Q��
    HWFMNHWS As String * 1          ' �i�v�e�����Z�x�ۏؕ��@�Q��
    HWFMNKHM As String * 1          ' �i�v�e�����Z�x�����p�x�Q��
    HWFMNKHN As String * 1          ' �i�v�e�����Z�x�����p�x�Q��
    HWFMNKHH As String * 1          ' �i�v�e�����Z�x�����p�x�Q��
    HWFMNKHU As String * 1          ' �i�v�e�����Z�x�����p�x�Q�E
    HWFSPVMX As Double              ' �i�v�e�r�o�u�e�d���
'    HWFSPVMXN As Double              ' �i�v�e�r�o�u�e�d���  6/22 Yam
    HWFSPVKM As String * 1          ' �i�v�e�r�o�u�e�d�����p�x�Q��
    HWFSPVKN As String * 1          ' �i�v�e�r�o�u�e�d�����p�x�Q��
    HWFSPVKH As String * 1          ' �i�v�e�r�o�u�e�d�����p�x�Q��
    HWFSPVKU As String * 1          ' �i�v�e�r�o�u�e�d�����p�x�Q�E
    HWFSPVSH As String * 1          ' �i�v�e�r�o�u�e�d����ʒu�Q��
    HWFSPVST As String * 1          ' �i�v�e�r�o�u�e�d����ʒu�Q�_
    HWFSPVSI As String * 1          ' �i�v�e�r�o�u�e�d����ʒu�Q��
    HWFSPVHT As String * 1          ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
    HWFSPVHS As String * 1          ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
    HWFDLMIN As Integer             ' �i�v�e�g�U������
    HWFDLMAX As Integer             ' �i�v�e�g�U�����
    HWFDLKHM As String * 1          ' �i�v�e�g�U�������p�x�Q��
    HWFDLKHN As String * 1          ' �i�v�e�g�U�������p�x�Q��
    HWFDLKHH As String * 1          ' �i�v�e�g�U�������p�x�Q��
    HWFDLKHU As String * 1          ' �i�v�e�g�U�������p�x�Q�E
    HWFDLSPH As String * 1          ' �i�v�e�g�U������ʒu�Q��
    HWFDLSPT As String * 1          ' �i�v�e�g�U������ʒu�Q�_
    HWFDLSPI As String * 1          ' �i�v�e�g�U������ʒu�Q��
    HWFDLHWT As String * 1          ' �i�v�e�g�U���ۏؕ��@�Q��
    HWFDLHWS As String * 1          ' �i�v�e�g�U���ۏؕ��@�Q��
    HWFGKNO1 As String * 6          ' �i�v�e�O�ϋK�i�m���P
    HWFGKNO2 As String * 6          ' �i�v�e�O�ϋK�i�m���Q
    HWFOTMIN As Double              ' �i�v�e�_�����ψ�����
    HWFOTMX1 As Double              ' �i�v�e�_�����ψ�����P
    HWFOTMX2 As Double              ' �i�v�e�_�����ψ�����Q
    HWFOTSPH As String * 1          ' �i�v�e�_�����ψ�����ʒu�Q��
    HWFOTSPT As String * 1          ' �i�v�e�_�����ψ�����ʒu�Q�_
    HWFOTSPI As String * 1          ' �i�v�e�_�����ψ�����ʒu�Q��
    HWFOTHWT As String * 1          ' �i�v�e�_�����ψ��ۏؕ��@�Q��
    HWFOTHWS As String * 1          ' �i�v�e�_�����ψ��ۏؕ��@�Q��
    HWFOTKWY As String * 2          ' �i�v�e�_�����ψ��������@
    HWFOTKW1 As String * 2          ' �i�v�e�_�����ψ��������@�P
    HWFOTKW2 As String * 2          ' �i�v�e�_�����ψ��������@�Q
    HWFOTKHM As String * 1          ' �i�v�e�_�����ψ������p�x�Q��
    HWFOTKHN As String * 1          ' �i�v�e�_�����ψ������p�x�Q��
    HWFOTKHH As String * 1          ' �i�v�e�_�����ψ������p�x�Q��
    HWFOTKHU As String * 1          ' �i�v�e�_�����ψ������p�x�Q�E
    HWFTSPHM As String * 1          ' �i�v�e�g���X�T���v���p�x�Q��
    HWFTSPHN As String * 1          ' �i�v�e�g���X�T���v���p�x�Q��
    HWFTSPHH As String * 1          ' �i�v�e�g���X�T���v���p�x�Q��
    HWFTSPHU As String * 1          ' �i�v�e�g���X�T���v���p�x�Q�E
    HWFLTDCX As Double              ' �i�v�e�k�s�c�Z�x�b�t���
    HWFLTDIN As String * 2          ' �i�v�e�k�s�c�Z�x�w��
    HWFLTDKW As String * 2          ' �i�v�e�k�s�c�Z�x�������@
    HWFLTDSH As String * 1          ' �i�v�e�k�s�c�Z�x����ʒu�Q��
    HWFLTDST As String * 1          ' �i�v�e�k�s�c�Z�x����ʒu�Q�_
    HWFLTDSI As String * 1          ' �i�v�e�k�s�c�Z�x����ʒu�Q��
    HWFLTDHT As String * 1          ' �i�v�e�k�s�c�Z�x�ۏؕ��@�Q��
    HWFLTDHS As String * 1          ' �i�v�e�k�s�c�Z�x�ۏؕ��@�Q��
    HWFLTDKM As String * 1          ' �i�v�e�k�s�c�Z�x�����p�x�Q��
    HWFLTDKN As String * 1          ' �i�v�e�k�s�c�Z�x�����p�x�Q��
    HWFLTDKH As String * 1          ' �i�v�e�k�s�c�Z�x�����p�x�Q��
    HWFLTDKU As String * 1          ' �i�v�e�k�s�c�Z�x�����p�x�Q�E
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
' �ǉ� 2003/09.11 SystemBrain Start
    HWFSPVAM As Double              ' �iWFSPVFE����
'    HWFSPVAMN As Double              ' �iWFSPVFE���� 6/22
    HWFMK1MC As String * 1          ' �iWF�ʌ�����1�ʎw��
    HWFMK2MC As String * 1          ' �iWF�ʌ�����2�ʎw��
    HWFMK3MC As String * 1          ' �iWF�ʌ�����3�ʎw��
    HWFMK4MC As String * 1          ' �iWF�ʌ�����4�ʎw��
    HWFMK5MC As String * 1          ' �iWF�ʌ�����5�ʎw��
    HWFMK6MC As String * 1          ' �iWF�ʌ�����6�ʎw��
    HWFMK2SZ As String * 1          ' �iWF�ʌ�����2�������
    HWFMK3SZ As String * 1          ' �iWF�ʌ�����3�������
    HWFMK4SZ As String * 1          ' �iWF�ʌ�����4�������
    HWFMK2ZAR As Integer            ' �iWF�ʌ�����2���O�̈�
    HWFMK3ZAR As Integer            ' �iWF�ʌ�����3���O�̈�
    HWFMK4ZAR As Integer            ' �iWF�ʌ�����4���O�̈�
    HWFMK5B1 As Integer             ' �iWF�ʌ�����5���E1
    HWFMK5B1B As Integer            ' �iWF�ʌ�����5���E1��
    HWFMK5B2 As Integer             ' �iWF�ʌ�����5���E2
    HWFMK5B2B As Integer            ' �iWF�ʌ�����5���E2��
    HWFMK5B3 As Integer             ' �iWF�ʌ�����5���E3
    HWFMK5B3B As Integer            ' �iWF�ʌ�����5���E3��
    HWFMK6B1 As Integer             ' �iWF�ʌ�����6���E1
    HWFMK6B1B As Integer            ' �iWF�ʌ�����6���E1��
    HWFMK6B2 As Integer             ' �iWF�ʌ�����6���E2
    HWFMK6B2B As Integer            ' �iWF�ʌ�����6���E2��
    HWFMK6B3 As Integer             ' �iWF�ʌ�����6���E3
    HWFMK6B3B As Integer            ' �iWF�ʌ�����6���E3��
' �ǉ� 2003/09.11 SystemBrain End
' �ǉ� 2005/06/16 ffc)tanabe start
    HWFMK7MC    As String * 1       '�i�v�e�ʌ����ׂV�ʎw��
    HWFMK7SI    As Double           '�i�v�e�ʌ����ׂV�T�C�Y
    HWFMK7MX    As Integer          '�i�v�e�ʌ����ׂV���
    HWFMK7SZ    As String * 1       '�i�v�e�ʌ����ׂV�������
    HWFMK7ZA    As Integer          '�i�v�e�ʌ����ׂV���O�̈�
    HWFMK7HT    As String * 1       '�i�v�e�ʌ����ׂV�ۏؕ��@�Q��
    HWFMK7HS    As String * 1       '�i�v�e�ʌ����ׂV�ۏؕ��@�Q��
    HWFMK8MC    As String * 1       '�i�v�e�ʌ����ׂW�ʎw��
    HWFMK8SI    As Double           '�i�v�e�ʌ����ׂW�T�C�Y
    HWFMK8MX    As Integer          '�i�v�e�ʌ����ׂW���
    HWFMK8SZ    As String * 1       '�i�v�e�ʌ����ׂW�������
    HWFMK8ZA    As Integer          '�i�v�e�ʌ����ׂW���O�̈�
    HWFMK8HT    As String * 1       '�i�v�e�ʌ����ׂW�ۏؕ��@�Q��
    HWFMK8HS    As String * 1       '�i�v�e�ʌ����ׂW�ۏؕ��@�Q��
    HWFMK9MC    As String * 1       '�i�v�e�ʌ����ׂX�ʎw��
    HWFMK9SI    As Double           '�i�v�e�ʌ����ׂX�T�C�Y
    HWFMK9MX    As Integer          '�i�v�e�ʌ����ׂX���
    HWFMK9SZ    As String * 1       '�i�v�e�ʌ����ׂX�������
    HWFMK9ZA    As Integer          '�i�v�e�ʌ����ׂX���O�̈�
    HWFMK9HT    As String * 1       '�i�v�e�ʌ����ׂX�ۏؕ��@�Q��
    HWFMK9HS    As String * 1       '�i�v�e�ʌ����ׂX�ۏؕ��@�Q��
    HWFMK10MC   As String * 1       '�i�v�e�ʌ����ׂP�O�ʎw��
    HWFMK10SI   As Double           '�i�v�e�ʌ����ׂP�O�T�C�Y
    HWFMK10MX   As Integer          '�i�v�e�ʌ����ׂP�O���
    HWFMK10SZ   As String * 1       '�i�v�e�ʌ����ׂP�O�������
    HWFMK10ZA   As Integer          '�i�v�e�ʌ����ׂP�O���O�̈�
    HWFMK10HT   As String * 1       '�i�v�e�ʌ����ׂP�O�ۏؕ��@�Q��
    HWFMK10HS   As String * 1       '�i�v�e�ʌ����ׂP�O�ۏؕ��@�Q��
    HWFMK11MC   As String * 1       '�i�v�e�ʌ����ׂP�P�ʎw��
    HWFMK11SI   As Double           '�i�v�e�ʌ����ׂP�P�T�C�Y
    HWFMK11MX   As Integer          '�i�v�e�ʌ����ׂP�P���
    HWFMK11SZ   As String * 1       '�i�v�e�ʌ����ׂP�P�������
    HWFMK11ZA   As Integer          '�i�v�e�ʌ����ׂP�P���O�̈�
    HWFMK11HT   As String * 1       '�i�v�e�ʌ����ׂP�P�ۏؕ��@�Q��
    HWFMK11HS   As String * 1       '�i�v�e�ʌ����ׂP�P�ۏؕ��@�Q��
    HWFMK12MC   As String * 1       '�i�v�e�ʌ����ׂP�Q�ʎw��
    HWFMK12SI   As Double           '�i�v�e�ʌ����ׂP�Q�T�C�Y
    HWFMK12MX   As Integer          '�i�v�e�ʌ����ׂP�Q���
    HWFMK12SZ   As String * 1       '�i�v�e�ʌ����ׂP�Q�������
    HWFMK12ZA   As Integer          '�i�v�e�ʌ����ׂP�Q���O�̈�
    HWFMK12HT   As String * 1       '�i�v�e�ʌ����ׂP�Q�ۏؕ��@�Q��
    HWFMK12HS   As String * 1       '�i�v�e�ʌ����ׂP�Q�ۏؕ��@�Q��
    HWFMK13MC   As String * 1       '�i�v�e�ʌ����ׂP�R�ʎw��
    HWFMK13SI   As Double           '�i�v�e�ʌ����ׂP�R�T�C�Y
    HWFMK13MX   As Integer          '�i�v�e�ʌ����ׂP�R���
    HWFMK13SZ   As String * 1       '�i�v�e�ʌ����ׂP�R�������
    HWFMK13ZA   As Integer          '�i�v�e�ʌ����ׂP�R���O�̈�
    HWFMK13HT   As String * 1       '�i�v�e�ʌ����ׂP�R�ۏؕ��@�Q��
    HWFMK13HS   As String * 1       '�i�v�e�ʌ����ׂP�R�ۏؕ��@�Q��
    HWFMK14MC   As String * 1       '�i�v�e�ʌ����ׂP�S�ʎw��
    HWFMK14SI   As Double           '�i�v�e�ʌ����ׂP�S�T�C�Y
    HWFMK14MX   As Integer          '�i�v�e�ʌ����ׂP�S���
    HWFMK14SZ   As String * 1       '�i�v�e�ʌ����ׂP�S�������
    HWFMK14ZA   As Integer          '�i�v�e�ʌ����ׂP�S���O�̈�
    HWFMK14HT   As String * 1       '�i�v�e�ʌ����ׂP�S�ۏؕ��@�Q��
    HWFMK14HS   As String * 1       '�i�v�e�ʌ����ׂP�S�ۏؕ��@�Q��
    HWFMK15MC   As String * 1       '�i�v�e�ʌ����ׂP�T�ʎw��
    HWFMK15SI   As Double           '�i�v�e�ʌ����ׂP�T�T�C�Y
    HWFMK15MX   As Integer          '�i�v�e�ʌ����ׂP�T���
    HWFMK15SZ   As String * 1       '�i�v�e�ʌ����ׂP�T�������
    HWFMK15ZA   As Integer          '�i�v�e�ʌ����ׂP�T���O�̈�
    HWFMK15HT   As String * 1       '�i�v�e�ʌ����ׂP�T�ۏؕ��@�Q��
    HWFMK15HS   As String * 1       '�i�v�e�ʌ����ׂP�T�ۏؕ��@�Q��
    HWFSPVMXN   As Double           '�i�v�e�r�o�u�e�d���
    HWFSPVAMN   As Double           '�i�v�e�r�o�u�e�d����
' �ǉ� 2005/06/16 ffc)tanabe end
End Type


' ���i�d�lWF�ް��X
Public Type typ_TBCME029
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    HMGSTFNO As String * 8          ' �i�Ǘ��Ј��m��
    HMGWFSNO As String * 6          ' �i�Ǘ��v�e���i�ԍ�
    HMGWFSNE As Integer             ' �i�Ǘ��v�e���i�ԍ��}��
    HWFOF1AX As Double              ' �i�v�e�n�r�e�P���Ϗ��
    HWFOF1MX As Double              ' �i�v�e�n�r�e�P���
    HWFOF1ET As Integer             ' �i�v�e�n�r�e�P�I���d�s��
    HWFOF1NS As String * 2          ' �i�v�e�n�r�e�P�M�����@
    HWFOF1SZ As String * 1          ' �i�v�e�n�r�e�P�������
    HWFOF1SH As String * 1          ' �i�v�e�n�r�e�P����ʒu�Q��
    HWFOF1ST As String * 1          ' �i�v�e�n�r�e�P����ʒu�Q�_
    HWFOF1SR As String * 1          ' �i�v�e�n�r�e�P����ʒu�Q��
    HWFOF1HT As String * 1          ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    HWFOF1HS As String * 1          ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    HWFOF1KM As String * 1          ' �i�v�e�n�r�e�P�����p�x�Q��
    HWFOF1KN As String * 1          ' �i�v�e�n�r�e�P�����p�x�Q��
    HWFOF1KH As String * 1          ' �i�v�e�n�r�e�P�����p�x�Q��
    HWFOF1KU As String * 1          ' �i�v�e�n�r�e�P�����p�x�Q�E
    HWFOF2AX As Double              ' �i�v�e�n�r�e�Q���Ϗ��
    HWFOF2MX As Double              ' �i�v�e�n�r�e�Q���
    HWFOF2ET As Integer             ' �i�v�e�n�r�e�Q�I���d�s��
    HWFOF2NS As String * 2          ' �i�v�e�n�r�e�Q�M�����@
    HWFOF2SZ As String * 1          ' �i�v�e�n�r�e�Q�������
    HWFOF2SH As String * 1          ' �i�v�e�n�r�e�Q����ʒu�Q��
    HWFOF2ST As String * 1          ' �i�v�e�n�r�e�Q����ʒu�Q�_
    HWFOF2SR As String * 1          ' �i�v�e�n�r�e�Q����ʒu�Q��
    HWFOF2HT As String * 1          ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    HWFOF2HS As String * 1          ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    HWFOF2KM As String * 1          ' �i�v�e�n�r�e�Q�����p�x�Q��
    HWFOF2KN As String * 1          ' �i�v�e�n�r�e�Q�����p�x�Q��
    HWFOF2KH As String * 1          ' �i�v�e�n�r�e�Q�����p�x�Q��
    HWFOF2KU As String * 1          ' �i�v�e�n�r�e�Q�����p�x�Q�E
    HWFOF3AX As Double              ' �i�v�e�n�r�e�R���Ϗ��
    HWFOF3MX As Double              ' �i�v�e�n�r�e�R���
    HWFOF3ET As Integer             ' �i�v�e�n�r�e�R�I���d�s��
    HWFOF3NS As String * 2          ' �i�v�e�n�r�e�R�M�����@
    HWFOF3SZ As String * 1          ' �i�v�e�n�r�e�R�������
    HWFOF3SH As String * 1          ' �i�v�e�n�r�e�R����ʒu�Q��
    HWFOF3ST As String * 1          ' �i�v�e�n�r�e�R����ʒu�Q�_
    HWFOF3SR As String * 1          ' �i�v�e�n�r�e�R����ʒu�Q��
    HWFOF3HT As String * 1          ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    HWFOF3HS As String * 1          ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    HWFOF3KM As String * 1          ' �i�v�e�n�r�e�R�����p�x�Q��
    HWFOF3KN As String * 1          ' �i�v�e�n�r�e�R�����p�x�Q��
    HWFOF3KH As String * 1          ' �i�v�e�n�r�e�R�����p�x�Q��
    HWFOF3KU As String * 1          ' �i�v�e�n�r�e�R�����p�x�Q�E
    HWFOF4AX As Double              ' �i�v�e�n�r�e�S���Ϗ��
    HWFOF4MX As Double              ' �i�v�e�n�r�e�S���
    HWFOF4ET As Integer             ' �i�v�e�n�r�e�S�I���d�s��
    HWFOF4NS As String * 2          ' �i�v�e�n�r�e�S�M�����@
    HWFOF4SZ As String * 1          ' �i�v�e�n�r�e�S�������
    HWFOF4SH As String * 1          ' �i�v�e�n�r�e�S����ʒu�Q��
    HWFOF4ST As String * 1          ' �i�v�e�n�r�e�S����ʒu�Q�_
    HWFOF4SR As String * 1          ' �i�v�e�n�r�e�S����ʒu�Q��
    HWFOF4HT As String * 1          ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    HWFOF4HS As String * 1          ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    HWFOF4KM As String * 1          ' �i�v�e�n�r�e�S�����p�x�Q��
    HWFOF4KN As String * 1          ' �i�v�e�n�r�e�S�����p�x�Q��
    HWFOF4KH As String * 1          ' �i�v�e�n�r�e�S�����p�x�Q��
    HWFOF4KU As String * 1          ' �i�v�e�n�r�e�S�����p�x�Q�E
    HWFBM1AN As Double              ' �i�v�e�a�l�c�P���ω���
    HWFBM1AX As Double              ' �i�v�e�a�l�c�P���Ϗ��
    HWFBM1ET As Integer             ' �i�v�e�a�l�c�P�I���d�s��
    HWFBM1NS As String * 2          ' �i�v�e�a�l�c�P�M�����@
    HWFBM1SZ As String * 1          ' �i�v�e�a�l�c�P�������
    HWFBM1SH As String * 1          ' �i�v�e�a�l�c�P����ʒu�Q��
    HWFBM1ST As String * 1          ' �i�v�e�a�l�c�P����ʒu�Q�_
    HWFBM1SR As String * 1          ' �i�v�e�a�l�c�P����ʒu�Q��
    HWFBM1HT As String * 1          ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    HWFBM1HS As String * 1          ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    HWFBM1KM As String * 1          ' �i�v�e�a�l�c�P�����p�x�Q��
    HWFBM1KN As String * 1          ' �i�v�e�a�l�c�P�����p�x�Q��
    HWFBM1KH As String * 1          ' �i�v�e�a�l�c�P�����p�x�Q��
    HWFBM1KU As String * 1          ' �i�v�e�a�l�c�P�����p�x�Q�E
    HWFBM2AN As Double              ' �i�v�e�a�l�c�Q���ω���
    HWFBM2AX As Double              ' �i�v�e�a�l�c�Q���Ϗ��
    HWFBM2ET As Integer             ' �i�v�e�a�l�c�Q�I���d�s��
    HWFBM2NS As String * 2          ' �i�v�e�a�l�c�Q�M�����@
    HWFBM2SZ As String * 1          ' �i�v�e�a�l�c�Q�������
    HWFBM2SH As String * 1          ' �i�v�e�a�l�c�Q����ʒu�Q��
    HWFBM2ST As String * 1          ' �i�v�e�a�l�c�Q����ʒu�Q�_
    HWFBM2SR As String * 1          ' �i�v�e�a�l�c�Q����ʒu�Q��
    HWFBM2HT As String * 1          ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    HWFBM2HS As String * 1          ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    HWFBM2KM As String * 1          ' �i�v�e�a�l�c�Q�����p�x�Q��
    HWFBM2KN As String * 1          ' �i�v�e�a�l�c�Q�����p�x�Q��
    HWFBM2KH As String * 1          ' �i�v�e�a�l�c�Q�����p�x�Q��
    HWFBM2KU As String * 1          ' �i�v�e�a�l�c�Q�����p�x�Q�E
    HWFBM3AN As Double              ' �i�v�e�a�l�c�R���ω���
    HWFBM3AX As Double              ' �i�v�e�a�l�c�R���Ϗ��
    HWFBM3ET As Integer             ' �i�v�e�a�l�c�R�I���d�s��
    HWFBM3NS As String * 2          ' �i�v�e�a�l�c�R�M�����@
    HWFBM3SZ As String * 1          ' �i�v�e�a�l�c�R�������
    HWFBM3SH As String * 1          ' �i�v�e�a�l�c�R����ʒu�Q��
    HWFBM3ST As String * 1          ' �i�v�e�a�l�c�R����ʒu�Q�_
    HWFBM3SR As String * 1          ' �i�v�e�a�l�c�R����ʒu�Q��
    HWFBM3HT As String * 1          ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    HWFBM3HS As String * 1          ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    HWFBM3KM As String * 1          ' �i�v�e�a�l�c�R�����p�x�Q��
    HWFBM3KN As String * 1          ' �i�v�e�a�l�c�R�����p�x�Q��
    HWFBM3KH As String * 1          ' �i�v�e�a�l�c�R�����p�x�Q��
    HWFBM3KU As String * 1          ' �i�v�e�a�l�c�R�����p�x�Q�E
    HWFOSPAX As Integer             ' �i�v�e�n�r�o���Ϗ��
    HWFOSPMX As Integer             ' �i�v�e�n�r�o���
    HWFOSPSH As String * 1          ' �i�v�e�n�r�o����ʒu�Q��
    HWFOSPST As String * 1          ' �i�v�e�n�r�o����ʒu�Q�_
    HWFOSPSR As String * 1          ' �i�v�e�n�r�o����ʒu�Q��
    HWFOSPHT As String * 1          ' �i�v�e�n�r�o�ۏؕ��@�Q��
    HWFOSPHS As String * 1          ' �i�v�e�n�r�o�ۏؕ��@�Q��
    HWFOSPNS As String * 2          ' �i�v�e�n�r�o�M�����@
    HWFOSPSZ As String * 1          ' �i�v�e�n�r�o�������
    HWFOSPKM As String * 1          ' �i�v�e�n�r�o�����p�x�Q��
    HWFOSPKN As String * 1          ' �i�v�e�n�r�o�����p�x�Q��
    HWFOSPKH As String * 1          ' �i�v�e�n�r�o�����p�x�Q��
    HWFOSPKU As String * 1          ' �i�v�e�n�r�o�����p�x�Q�E
    HWFOSPET As Integer             ' �i�v�e�n�r�o�I���d�s��
    HWFNOTE As String               ' �i�v�e���L
    HWFRS1N As String               ' �i�v�e�\���P�Q��
    HWFRS1Y As String               ' �i�v�e�\���P�Q�p
    HWFRS2N As String               ' �i�v�e�\���Q�Q��
    HWFRS2Y As String               ' �i�v�e�\���Q�Q�p
    HWFRS3N As String               ' �i�v�e�\���R�Q��
    HWFRS3Y As String               ' �i�v�e�\���R�Q�p
    HWFRS4N As String               ' �i�v�e�\���S�Q��
    HWFRS4Y As String               ' �i�v�e�\���S�Q�p
    HWFRS5N As String               ' �i�v�e�\���T�Q��
    HWFRS5Y As String               ' �i�v�e�\���T�Q�p
    HWFRS6N As String               ' �i�v�e�\���U�Q��
    HWFRS6Y As String               ' �i�v�e�\���U�Q�p
    HWFRS7N As String               ' �i�v�e�\���V�Q��
    HWFRS7Y As String               ' �i�v�e�\���V�Q�p
    HWFRS8N As String               ' �i�v�e�\���W�Q��
    HWFRS8Y As String               ' �i�v�e�\���W�Q�p
    HWFRS9N As String               ' �i�v�e�\���X�Q��
    HWFRS9Y As String               ' �i�v�e�\���X�Q�p
    HWFRS10N As String              ' �i�v�e�\���P�O�Q��
    HWFRS10Y As String              ' �i�v�e�\���P�O�Q�p
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
' �ǉ� 2003/09.11 SystemBrain Start
    HWFOSF1PTK As String * 1        ' �iWFOSF1�p�^���敪
    HWFOSF2PTK As String * 1        ' �iWFOSF2�p�^���敪
    HWFOSF3PTK As String * 1        ' �iWFOSF3�p�^���敪
    HWFOSF4PTK As String * 1        ' �iWFOSF4�p�^���敪
    HWFBM1MBP As Double             ' �iWFBMD1�ʓ����z
    HWFBM2MBP As Double             ' �iWFBMD2�ʓ����z
    HWFBM3MBP As Double             ' �iWFBMD3�ʓ����z
    HWFBM1MCL As String * 2         ' �iWFBMD1�ʓ��v�Z
    HWFBM2MCL As String * 2         ' �iWFBMD2�ʓ��v�Z
    HWFBM3MCL As String * 2         ' �iWFBMD3�ʓ��v�Z
' �ǉ� 2003/09.11 SystemBrain End
End Type


' SXL��������ް�
Public Type typ_TBCME030
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    SSXLIFTW As String * 2          ' ���r�w������@
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �r�w�k��������t�^���
Public Type typ_TBCME031
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    STAFFNO As String * 8           ' �Ј���
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �g�p�J�n
Public Type typ_TBCME032
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    STAFFNO As String * 8           ' �Ј���
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ėp����Ͻ�
Public Type typ_TBCME033
    codeNo As String * 12           ' �R�[�h�m�n
    CODE As String * 5              ' �R�[�h
    codeCont As String              ' �R�[�h���e
    INDORDER As Long                ' �\����
    codename As String              ' �R�[�h����
    KUBUN As String                 ' �敪
    READTIME As Double              ' ���[�h�^�C��
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ���R�␳�v�Z
Public Type typ_TBCME034
    RESIHCAL As String * 2          ' ���R�␳�v�Z
    RESIHINA As Double              ' ���R�␳�W���`
    RESIHINB As Double              ' ���R�␳�W���a
    CSGROUP As String * 3           ' �ڋq�O���[�v
    CSCODE As String * 8            ' �ڋq�R�[�h
    CSNAME As String                ' �ڋq��
    NOTE As String                  ' ���L
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �_�f�Z�x�␳�v�Z
Public Type typ_TBCME035
    OXYNHCAL As String * 2          ' �_�f�Z�x�␳�v�Z
    OXYNHINA As Double              ' �_�f�Z�x�␳�W���`
    OXYNHINB As Double              ' �_�f�Z�x�␳�W���a
    CSGROUP As String * 3           ' �ڋq�O���[�v
    CSCODE As String * 8            ' �ڋq�R�[�h
    CSNAME As String                ' �ڋq��
    NOTE As String                  ' ���L
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ���������Ǘ�
Public Type typ_TBCME036
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    EPDSETCH As String * 1          ' EPD�@�I���G�b�`
    EPDUP As Integer                ' EPD�@���
    CUTUNIT As Integer              ' �J�b�g�P��
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 start
    TOPREG As Integer               ' TOP�K��
    TAILREG As Double               ' TAIL�K��
    BTMSPRT As Integer              ' �{�g���͏o�K��
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 end
' �ǉ� 2003/09.11 SystemBrain Start
    OTHER1 As String * 1            '
    OTHER2 As String * 1            '
    OTHERTIME As Date               '
    DCHYUUBU As String * 1          ' �h���[�`���[�u
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    sStaffID As String * 8          ' ���F�Ј�ID
    SYNFLAG As String * 1           ' ���F�t���O
    SYNDATE As Date                 ' ���F���t
    SNOTE As String * 255           ' ���i�d�l���L
    JNOTE As String * 255           ' ����������L
    BLOCKHFLAG As String * 1        ' �u���b�N�P�ʕۏؕi�ԃt���O
' �ǉ� 2003/09.11 SystemBrain End
' WF�J�b�g�P�ʋ@�\�ǉ� 2005/04/12 ffc)tanabe start
    WFCUTUNIT As String * 4         'WF�J�b�g�P��
' WF�J�b�g�P�ʋ@�\�ǉ� 2005/04/12 ffc)tanabe end
'*** UPDATE �� Y.SIMIZU 2005/10/1 GDײݐ�
    HSXGDLINE   As Single           '�iSXGDײݐ�
    HWFGDLINE   As Single           '�iWFGDײݐ�
'*** UPDATE �� Y.SIMIZU 2005/10/1 GDײݐ�

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP    As String * 1       '�iSXDK���x
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    HSXLDLRMN   As Integer          ' �iSXL/DL�A��0����
    HSXLDLRMX   As Integer          ' �iSXL/DL�A��0���
    HWFLDLRMN   As Integer          ' �iWFL/DL�A��0����
    HWFLDLRMX   As Integer          ' �iWFL/DL�A��0���
    HSXOF1ARPTK As String * 1       ' �iSXOSF1(ArAN)�p�^���敪
    HSXOFARMIN  As Double           ' �iSXOSF(ArAN)����
    HSXOFARMAX  As Double           ' �iSXOSF(ArAN)���
    HSXOFARMHMX As Double           ' �iSXOSF(ArAN)�ʓ�����
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    'Add Start 2011/01/27 SMPK Miyata
    HSXCJLTBND  As Integer          ' �iSXL/CJLT�o���h��
    'Add End   2011/01/27 SMPK Miyata

End Type


' �u���b�N�V�K���i�����w���t�j
Public Type typ_TBCMY001
    BLOCKID As String * 12          ' �u���b�NID
    BLOCKLEN As String * 3          ' �u���b�N�̒���
    MAINHINBAN As String * 10       ' ��\�i��
    PNTYPE As String * 1            ' �^�C�v
    ROUP As String * 8              ' ���R����l
    ROLOW As String * 8             ' ���R�����l
    OIUP As String * 5              ' �_�f�Z�x����l
    OILOW As String * 5             ' �_�f�Z�x�����l
    TANMEN As String * 3            ' �[�ʊp�x
    WARPRANK As String * 1          ' ���[�v�����N
    CRYSTALMEN As String * 3        ' ������
    SLPCEN As String * 4            ' �X���S
    SLPLOW As String * 5            ' �X����
    SLPUP As String * 5             ' �X���
    INSPMETH As String * 2          ' �������@
    INSPFREQ As String * 4          ' �����p�x
    SLPDRC As String * 2            ' �X����
    SLPDRCAPP As String * 1         ' �X���ʎw��
    SLPHEIDRC As String * 2         ' �X�c����
    SLPHEICEN As String * 5         ' �X�c���S
    SLPHEILOW As String * 5         ' �X�c����
    SLPHEIUP As String * 5          ' �X�c���
    SLPWIDDRC As String * 2         ' �X������
    SLPWIDCEN As String * 5         ' �X�����S
    SLPWIDLOW As String * 5         ' �X������
    SLPWIDUP As String * 5          ' �X�����
    SEED As String * 1              ' ���㎞�g�p�����V�|�h�X��
    TXID As String * 6              ' �g�����U�N�V����ID
    sBlockId As String * 12         ' �擪�u���b�NID
    BLOCKORDER As Integer           ' �u���b�N����
    REGDATE As Date                 ' �o�^���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
'2007/07/17 UPDATE_STR �}���`�u���b�N�Ή� SHINDOH
    HINCNT As Integer               ' �\���i�Ԑ�
    MULUTIHINBAN1 As String * 10    ' �\���i�Ԃ��̂P�i��
    TOPICHI1 As Integer             ' �\���i�Ԃ��̂PTop�ʒu(mm)
    TAILICHI1 As Integer            ' �\���i�Ԃ��̂PTail�ʒu(mm)
    HINBANLEN1 As Integer           ' �\���i�Ԃ��̂P����(mm)
    MULUTIHINBAN2 As String * 10    ' �\���i�Ԃ��̂Q�i��
    TOPICHI2 As Integer             ' �\���i�Ԃ��̂QTop�ʒu(mm)
    TAILICHI2 As Integer            ' �\���i�Ԃ��̂QTail�ʒu(mm)
    HINBANLEN2 As Integer           ' �\���i�Ԃ��̂Q����(mm)
    MULUTIHINBAN3 As String * 10    ' �\���i�Ԃ��̂R�i��
    TOPICHI3 As Integer             ' �\���i�Ԃ��̂RTop�ʒu(mm)
    TAILICHI3 As Integer            ' �\���i�Ԃ��̂RTail�ʒu(mm)
    HINBANLEN3 As Integer           ' �\���i�Ԃ��̂R����(mm)
    MULUTIHINBAN4 As String * 10    ' �\���i�Ԃ��̂S�i��
    TOPICHI4 As Integer             ' �\���i�Ԃ��̂STop�ʒu(mm)
    TAILICHI4 As Integer            ' �\���i�Ԃ��̂STail�ʒu(mm)
    HINBANLEN4 As Integer           ' �\���i�Ԃ��̂S����(mm)
    MULUTIHINBAN5 As String * 10    ' �\���i�Ԃ��̂T�i��
    TOPICHI5 As Integer             ' �\���i�Ԃ��̂TTop�ʒu(mm)
    TAILICHI5 As Integer            ' �\���i�Ԃ��̂TTail�ʒu(mm)
    HINBANLEN5 As Integer           ' �\���i�Ԃ��̂T����(mm)
'2007/07/17 UPDATE_END �}���`�u���b�N�Ή� SHINDOH
End Type


' �u���b�N�V�K���ԓ�
Public Type typ_TBCMY002
    BLOCKID As String * 12          ' �u���b�NID
    RET As String * 6               ' ���^�[���R�[�h
    TXID As String * 6              ' �g�����U�N�V����ID
    TXIDRET As String * 6           ' �g�����U�N�V����ID ���^�[���R�[�h
    BLKIDRET As String * 6          ' �u���b�NID�̃��^�[���R�[�h
    REGDATE As Date                 ' �o�^���t
    CHECKFLG As String * 1          ' �`�F�b�N�t���O
End Type


' ����]�����@�w��
Public Type typ_TBCMY003
    SAMPLEID As String * 16         ' �T���v��ID
    OSITEM As String * 4            ' �]������
    TRANCNT As Integer              ' ������
    SAMPLEKB As String * 1          ' �T���v���敪
    MAISU As String * 1             ' �]������
    Spec As String * 10             ' �K�i�l
    NETSU As String * 2             ' �M��������
    ET As String * 3                ' �G�b�`���O����
    MES As String * 3               ' �v�����@
    DKAN As String * 10             ' �c�j�A�j�[������
    TXID As String * 6              ' �g�����U�N�V����ID
    REGDATE As Date                 ' �o�^���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    FEPUA       As String * 10      ' SPV_Fe_PUA�l (number(5,2))    06/06/08 ooba START ======>
    FEPUAPCT    As String * 10      ' SPV_Fe_PUA���l (number(6,3))
    FESTD       As String * 10      ' SPV_Fe_STD (number(6,3))
    DIFFPUA     As String * 10      ' SPV_�g�U��_PUA�l (number(5,1))
    DIFFPUAPCT  As String * 10      ' SPV_�g�U��_PUA���l (number(6,3))
    NRPUA       As String * 10      ' SPV_NR_PUA�l (number(5,2))
    NRPUAPCT    As String * 10      ' SPV_NR_PUA%�l (number(6,3))
    NRSTD       As String * 10      ' SPV_NR_STD (number(6,3))      06/06/08 ooba END ========>
    MUKESAKI As String              ' 07/09/05 SPK Tsutsumi Add
End Type

' �G�s����]�����@�w��  2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
Public Type typ_TBCMY020
    SAMPLEID As String * 16         ' �T���v��ID
    OSITEM As String * 4            ' �]������
    TRANCNT As Integer              ' ������
    SAMPLEKB As String * 1          ' �T���v���敪
    MAISU As String * 1             ' �]������
    Spec As String * 10             ' �K�i�l
    NETSU As String * 2             ' �M��������
    ET As String * 3                ' �G�b�`���O����
    MES As String * 3               ' �v�����@
    DKAN As String * 10             ' �c�j�A�j�[������
    TXID As String * 6              ' �g�����U�N�V����ID
    REGDATE As Date                 ' �o�^���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    MUKESAKI As String              ' 07/09/05 SPK Tsutsumi Add
End Type


' ����]�����@�w���ԓ�
Public Type typ_TBCMY004
    SAMPLEID As String * 16         ' �T���v��ID
    TRANCNT As Integer              ' ������
    TXID As String * 6              ' �g�����U�N�V����ID
    RET As String * 6               ' ���^�[���R�[�h
    REGDATE As Date                 ' �o�^���t
    CHECKFLG As String * 1          ' �`�F�b�N�t���O
End Type


' �u���b�N�ύX���
Public Type typ_TBCMY005
    BLOCKID As String * 12          ' �u���b�NID
    TRANCNT As Integer              ' ������
    DELFLG As String * 1            ' �폜�w��
    BLOCKLEN As String * 3          ' �u���b�N�̒���
    MAINHINBAN As String * 10       ' ��\�i��
    PNTYPE As String * 1            ' �^�C�v
    ROUP As String * 8              ' ���R����l
    ROLOW As String * 8             ' ���R�����l
    OIUP As String * 5              ' �_�f�Z�x����l
    OILOW As String * 5             ' �_�f�Z�x�����l
    TANMEN As String * 3            ' �[�ʊp�x
    WARPRANK As String * 1          ' ���[�v�����N
    CRYSTALMEN As String * 3        ' ������
    SLPCEN As String * 4            ' �X���S
    SLPLOW As String * 5            ' �X����
    SLPUP As String * 5             ' �X���
    INSPMETH As String * 2          ' �������@
    INSPFREQ As String * 4          ' �����p�x
    SLPDRC As String * 2            ' �X����
    SLPDRCAPP As String * 1         ' �X���ʎw��
    SLPHEIDRC As String * 2         ' �X�c����
    SLPHEICEN As String * 5         ' �X�c���S
    SLPHEILOW As String * 5         ' �X�c����
    SLPHEIUP As String * 5          ' �X�c���
    SLPWIDDRC As String * 2         ' �X������
    SLPWIDCEN As String * 5         ' �X�����S
    SLPWIDLOW As String * 5         ' �X������
    SLPWIDUP As String * 5          ' �X�����
    SEED As String * 1              ' ���㎞�g�p�����V�|�h�X��
    TXID As String * 6              ' �g�����U�N�V����ID
    REGDATE As Date                 ' �o�^���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �u���b�N�ύX���ԓ�
Public Type typ_TBCMY006
    BLOCKID As String * 12          ' �u���b�NID
    TRANCNT As Integer              ' ������
    RET As String * 6               ' ���^�[���R�[�h
    TXID As String * 6              ' �g�����U�N�V����ID
    TXIDRET As String * 6           ' �g�����U�N�V����ID ���^�[���R�[�h
    BLKIDRET As String * 6          ' �u���b�NID�̃��^�[���R�[�h
    REGDATE As Date                 ' �o�^���t
    CHECKFLG As String * 1          ' �`�F�b�N�t���O
End Type


' �r�����m��w��
Public Type typ_TBCMY007
    SXL_ID As String * 13           ' SXL-ID
    SAMPLE_FROM As String * 16      ' �T���v��ID (From)
    SAMPLE_TO As String * 16        ' �T���v��ID (To)
    BLOCKID As String * 12          ' �u���b�N�h�c
    hinban As String * 10           ' �m��i��
    KUBUN As String * 2             ' �敪�R�[�h
    TXID As String * 6              ' �g�����U�N�V����ID
    REGDATE As Date                 ' �o�^���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    MESDATA1TOP As String * 10      ' ����l�P(Top)  center         '04/02/12 ooba START =======>
    MESDATA2TOP As String * 10      ' ����l�Q(Top)  R/2
    MESDATA3TOP As String * 10      ' ����l�R(Top)  Inside 10mm
    MESDATA4TOP As String * 10      ' ����l�S(Top)  Inside   6mm
    MESDATA5TOP As String * 10      ' ����l�T(Top)  Inside   3mm
    MESDATA1BOT As String * 10      ' ����l�P(Tail)  center
    MESDATA2BOT As String * 10      ' ����l�Q(Tail)  R/2
    MESDATA3BOT As String * 10      ' ����l�R(Tail)  Inside 10mm
    MESDATA4BOT As String * 10      ' ����l�S(Tail)  Inside   6mm
    MESDATA5BOT As String * 10      ' ����l�T(Tail)  Inside   3mm  '04/02/12 ooba END =========>
End Type


' �r�����m��w���ԓ�
Public Type typ_TBCMY008
    SXL_ID As String * 13           ' SXL-ID
    TXID As String * 6              ' �g�����U�N�V����ID
    RET As String * 6               ' ���^�[���R�[�h
    REGDATE As Date                 ' �o�^���t
    CHECKFLG As String * 1          ' �`�F�b�N�t���O
End Type


' �������
Public Type typ_TBCMY009
    LOTID As String * 12            ' �u���b�NID
    STRDTM As Date                  ' �������
    STRUSER_ID As String * 10       ' �����
    TXID As String * 6              ' �g�����U�N�V����ID
    REGDATE As Date                 ' �o�^���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ��ƊJ�n�E�I��
Public Type typ_TBCMY010
    LOTID As String * 12            ' �u���b�NID
    TRANCNT As Integer              ' ������
    ROUTE_ID As String * 10         ' ���|�g�h�c
    ROUTE_VER As String * 3         ' ���|�gID�o�[�W����
    OPE_ID As String * 6            ' �H��ID
    EQPID As String * 8             ' ���uID
    STRDTM As Date                  ' ��ƊJ�n����
    STRUSER_ID As String * 10       ' ��ƊJ�n��
    CMPDTM As Date                  ' ��ƏI������
    CMPUSER_ID As String * 10       ' ��ƏI����
    CURRWPCS As Integer             ' �E�F�n�[����
    TXID As String * 6              ' �g�����U�N�V����ID
    REGDATE As Date                 ' �o�^���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �E�F�n�|�Z���^�|���ɏ��
Public Type typ_TBCMY011
    LOTID As String * 12            ' �u���b�NID
    BLOCKSEQ As Integer             ' �u���b�N���A��
    INDTM As Date                   ' �E�F�n�[�Z���^�[���ɓ���
    BASKETID As String * 6          ' �o�X�P�b�gID
    SLOTNO As Integer               ' �X���b�gNo
    CURRWPCS As Integer             ' �E�F�n�[����
    EXISTFLG As String * 1          ' ���݃t���O
    TOP_POS As Integer              ' �u���b�N��Top����� �ʒu
    REJCAT As String * 1            ' �������R
    TXID As String * 6              ' �g�����U�N�V����ID
    REGDATE As Date                 ' �o�^���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �������
Public Type typ_TBCMY012
    LOTID As String * 12            ' �u���b�NID
    BLOCKSEQ As Integer             ' �u���b�N���A��
    REJPCS As Integer               ' �s�ǖ���
    TOP_POS As Integer              ' �u���b�N��Top����� �ʒu
    REJCAT As String * 1            ' �������R
    REJDTTM As Date                 ' ������
    REJPROC As String * 12          ' ���������H��
    ALLSCRAP As String * 1          ' �S���X�N���b�v
    LENFROM As Integer              ' �����@FROM
    LENTO As Integer                ' �����@TO
    TXID As String * 6              ' �g�����U�N�V����ID
    CHKFLG As String * 1            ' �`�F�b�N�t���O
    REGDATE As Date                 ' �o�^���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type

' ����]������
Public Type typ_TBCMY013
    SAMPLEID As String * 16         ' �T���v��ID
    OSITEM As String * 4            ' �]������
    MAISU As Integer                ' �]������
    Spec As String * 10             ' �K�i�l
    NETSU As String * 2             ' �M��������
    ET As String * 3                ' �G�b�`���O����
    MES As String * 3               ' �v�����@
    DKAN As String * 10             ' �c�j�A�j�[������
    MESDATA1 As String * 10         ' ����f�[�^���̂P
    MESDATA2 As String * 10         ' ����f�[�^���̂Q
    MESDATA3 As String * 10         ' ����f�[�^���̂R
    MESDATA4 As String * 10         ' ����f�[�^���̂S
    MESDATA5 As String * 10         ' ����f�[�^���̂T
    MESDATA6 As String * 10         ' ����f�[�^���̂U
    MESDATA7 As String * 10         ' ����f�[�^���̂V
    MESDATA8 As String * 10         ' ����f�[�^���̂W
    MESDATA9 As String * 10         ' ����f�[�^���̂X
    MESDATA10 As String * 10        ' ����f�[�^���̂P�O
    MESDATA11 As String * 10        ' ����f�[�^����1�P
    MESDATA12 As String * 10        ' ����f�[�^����1�Q
    MESDATA13 As String * 10        ' ����f�[�^����1�R
    MESDATA14 As String * 10        ' ����f�[�^����1�S
    MESDATA15 As String * 10        ' ����f�[�^����1�T
    TXID As String * 6              ' �g�����U�N�V����ID
    REGDATE As Date                 ' �o�^���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type

' �G�s����]�����@�w��  2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
' �G�s��s����]������
Public Type typ_TBCMY022
    SAMPLEID As String * 16         ' �T���v��ID
    OSITEM As String * 4            ' �]������
    MAISU As Integer                ' �]������
    Spec As String * 10             ' �K�i�l
    NETSU As String * 2             ' �M��������
    ET As String * 3                ' �G�b�`���O����
    MES As String * 3               ' �v�����@
    DKAN As String * 10             ' �c�j�A�j�[������
    MESDATA1 As String * 10         ' ����f�[�^���̂P
    MESDATA2 As String * 10         ' ����f�[�^���̂Q
    MESDATA3 As String * 10         ' ����f�[�^���̂R
    MESDATA4 As String * 10         ' ����f�[�^���̂S
    MESDATA5 As String * 10         ' ����f�[�^���̂T
    MESDATA6 As String * 10         ' ����f�[�^���̂U
    MESDATA7 As String * 10         ' ����f�[�^���̂V
    MESDATA8 As String * 10         ' ����f�[�^���̂W
    MESDATA9 As String * 10         ' ����f�[�^���̂X
    MESDATA10 As String * 10        ' ����f�[�^���̂P�O
    MESDATA11 As String * 10        ' ����f�[�^����1�P
    MESDATA12 As String * 10        ' ����f�[�^����1�Q
    MESDATA13 As String * 10        ' ����f�[�^����1�R
    MESDATA14 As String * 10        ' ����f�[�^����1�S
    MESDATA15 As String * 10        ' ����f�[�^����1�T
    TXID As String * 6              ' �g�����U�N�V����ID
    REGDATE As Date                 ' �o�^���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type

' �u���b�N�m����
Public Type typ_TBCMY014
    LOTID As String * 12            ' �u���b�NID
    BLOCKSEQ As Integer             ' �u���b�N���A��
    CURRWPCS As Integer             ' �E�F�n�[����
    EXISTFLG As String * 1          ' ���݃t���O
    SXL_ID As String * 13           ' �V���O��ID
    TOP_POS As String * 3           ' �u���b�N��Top����� �ʒu
    REJCAT As String * 1            ' �������R
    TXID As String * 6              ' �g�����U�N�V����ID
    REGDATE As Date                 ' �o�^���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �V���O���}�b�v���
Public Type typ_TBCMY015
    SXL_ID As String * 13           ' �V���O��ID
    SXLSEQ As Integer               ' �V���O�����A��
    SXLWPCS As Integer              ' �E�F�n�[����
    BLOCKID As String * 12          ' �u���b�NID
    BLOCKSEQ As Integer             ' �u���b�N���A��
    EXISTFLG As String * 1          ' ���݃t���O
    REJCAT As String * 1            ' �������R
    REGDATE As Date                 ' �o�^���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �����ُ�ԓ�
Public Type typ_TBCMY016
    SAMPLEID As String * 16         ' SAMPLEID
    RPTDATE As Date                 ' �񍐓���
    RET As String * 6               ' ��������
    TXID As String * 6              ' �g�����U�N�V����ID
    REGDATE As Date                 ' �o�^���t
End Type

' �֘A�u���b�N�R�t�R�؁@07/08/06 ooba
Public Type typ_TBCMY023
    CRYNUM As String * 12           '�����ԍ�
    TRANCNT As Integer              '������
    BLOCKID As String * 12          '��ۯ�ID
    PROCCAT As String * 1           '�����敪
    TXID As String * 6              '��ݻ޸���ID
End Type

'--------------- 2008/07/25 INSERT START  By Systech ---------------
' �u���b�N�i�ԐU�֗v��
Public Type typ_TBCMY027
    CRYNUM      As String * 12      '�����ԍ�
    TRANCNT     As String * 2       '������
    BLOCKID     As String * 12      '�u���b�NID
    TXID        As String * 6       '�g�����U�N�V����ID
    REQDATE     As Date             '�U�֗v������
    USER_ID     As String * 10      '��Ǝ҃R�[�h
    FRMAINHIN   As String * 10      '�U�֌���\�i��
    TOMAINHIN   As String * 10      '�U�֐��\�i��
    HINCNT      As Integer          '�\���i�Ԑ�
    FRHIN1      As String * 10      '�U�֌��\���i�ԂP
    TOHIN1      As String * 10      '�U�֐�\���i�ԂP
    FRHIN2      As String * 10      '�U�֌��\���i�ԂQ
    TOHIN2      As String * 10      '�U�֐�\���i�ԂQ
    FRHIN3      As String * 10      '�U�֌��\���i�ԂR
    TOHIN3      As String * 10      '�U�֐�\���i�ԂR
    FRHIN4      As String * 10      '�U�֌��\���i�ԂS
    TOHIN4      As String * 10      '�U�֐�\���i�ԂS
    FRHIN5      As String * 10      '�U�֌��\���i�ԂT
    TOHIN5      As String * 10      '�U�֐�\���i�ԂT
    REGDATE     As Date             '�o�^����
    CHECKFLG    As String * 1       '�`�F�b�N�t���O
    SNDKDWH     As String * 1       'DWH���M�t���O
    SDAYDWH     As Date             'DWH���M���t
    SNDKSPC     As String * 1       'SPC���M�t���O
    SDAYSPC     As Date             'SPC���M���t
    PLANTCAT    As String * 2       '���Ə��敪
End Type

' �u���b�N�i�ԐU�֗v������
Public Type typ_TBCMY028
    CRYNUM      As String * 12      '�����ԍ�
    TRANCNT     As String * 2       '������
    BLOCKID     As String * 12      '�u���b�NID
    TXID        As String * 6       '�g�����U�N�V����ID
    ALLJUDGRES  As String * 1       '�������茋��
    JUDGDATE    As Date             '�������
    HINCNT      As Integer          '�\���i�Ԑ�
    FRHIN1      As String * 10      '�U�֌��\���i�ԂP
    TOHIN1      As String * 10      '�U�֐�\���i�ԂP
    JudgRes1    As String * 1       '���茋�ʂP
    ERRCODE1    As String * 5       '�G���[�R�[�h�P
    FRHIN2      As String * 10      '�U�֌��\���i�ԂQ
    TOHIN2      As String * 10      '�U�֐�\���i�ԂQ
    JUDGRES2    As String * 1       '���茋�ʂQ
    ERRCODE2    As String * 5       '�G���[�R�[�h�Q
    FRHIN3      As String * 10      '�U�֌��\���i�ԂR
    TOHIN3      As String * 10      '�U�֐�\���i�ԂR
    JUDGRES3    As String * 1       '���茋�ʂR
    ERRCODE3    As String * 5       '�G���[�R�[�h�R
    FRHIN4      As String * 10      '�U�֌��\���i�ԂS
    TOHIN4      As String * 10      '�U�֐�\���i�ԂS
    JUDGRES4    As String * 1       '���茋�ʂS
    ERRCODE4    As String * 5       '�G���[�R�[�h�S
    FRHIN5      As String * 10      '�U�֌��\���i�ԂT
    TOHIN5      As String * 10      '�U�֐�\���i�ԂT
    JUDGRES5    As String * 1       '���茋�ʂT
    ERRCODE5    As String * 5       '�G���[�R�[�h�T
    REGDATE     As Date             '�o�^���t
    SENDFLAG    As String * 1       '���M�t���O
    SENDDATE    As Date             '���M���t
    SNDKDWH     As String * 1       'DWH���M�t���O
    SDAYDWH     As Date             'DWH���M���t
    SNDKSPC     As String * 1       'SPC���M�t���O
    SDAYSPC     As Date             'SPC���M���t
    PLANTCAT    As String * 2       '���Ə��敪
End Type
'--------------- 2008/07/25 INSERT  END   By Systech ---------------

' �Ј��}�X�^�[
Public Type typ_TBCMB001
    StaffID As String * 8           ' �Ј�ID
    PASSWD As String * 8            ' �p�X���[�h
    JFMLNAME As String              ' ���{�ꖼ�i���j
    JFSTNAME As String              ' ���{�ꖼ�i���j
    RFMLNAME As String              ' ���[�}�����i���j
    RFSTNAME As String              ' ���[�}�����i���j
    EXECODE As String * 4           ' ���s�����R�[�h
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
End Type


' �H���R�[�h�}�X�^�[
Public Type typ_TBCMB002
    KRPROCID As String * 5          ' �Ǘ��H��ID
    PROCCODE As String * 5          ' �H���R�[�h
    JPNNAME As String               ' ���{�ꖼ
    PROCSEQ As Integer              ' �H��������
    NOTE As String                  ' ���l
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
End Type


' �e�[�u�����}�X�^�[
Public Type typ_TBCMB018
    TABLENAME As String             ' �e�[�u����
    COLUM As String                 ' �J������
    NO As Integer                   ' �J������
    TYPE As String * 16             ' �^
    PKEY As String * 1              ' ��L�[
    BASETYPE As String * 16         ' ��{�^
    SIZE1 As Integer                ' �^�T�C�Y�P
    SIZE2 As Integer                ' �^�T�C�Y�Q
    MQBYTE As Long                  ' �l�p�o�C�g��
    JTABLE As String                ' ���{��e�[�u����
    JCOLUM As String                ' ���{��J������
    REF1 As String                  ' ���l�P
    REF2 As String                  ' ���l�Q
    TBLKBN As String * 1            ' �e�[�u�����
    REGDATE As Date                 ' �o�^���t
End Type


' ���b�Z�[�W�}�X�^�[
Public Type typ_TBCMB003
    MsgID As String * 5             ' ���b�Z�[�WID
    FORMINFO As String              ' Format���
    USEPRCID As String * 5          ' ���p����ID
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
End Type


' �����}�X�^�[
Public Type typ_TBCMB004
    AUTHCODE As String * 4          ' ���s�����R�[�h
    TRANID As String * 5            ' ����ID
    PWCHECK As String * 1           ' �p�X���[�h�`�F�b�N
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
End Type


' �R�[�h�}�X�^�[
Public Type typ_TBCMB005
    SYSCLASS As String * 2          ' �V�X�e���敪
    Class As String * 2             ' �敪
    CODE As String * 5              ' �R�[�h
    INFO1 As String                 ' ���P
    INFO2 As String                 ' ���Q
    INFO3 As String                 ' ���R
    INFO4 As String                 ' ���S
    INFO5 As String                 ' ���T
    INFO6 As String                 ' ���U
    INFO7 As String                 ' ���V
    INFO8 As String                 ' ���W
    INFO9 As String                 ' ���X
    NOTE As String                  ' ���l
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
End Type


' ���s�Ǘ��}�X�^�[
Public Type typ_TBCMB006
    ProcID As String * 12           ' ����ID
    EXENAME As String               ' ���s�t�@�C����
    BIKOU As String                 ' ���l
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
End Type


' �ϋ敪�}�X�^�[
Public Type typ_TBCMB007
    RCLSCODE As String * 3          ' �ϋ敪�R�[�h
    TYPE As String * 1              ' �^�C�v
    MINRESIST As Double             ' MIN�@��R�l
    MINMOVAL As Double              ' MIN�@MO�l
    MINFVAL As Double               ' MIN�@F�l
    MAXRESIST As Double             ' MAX�@��R�l
    MAXMOVAL As Double              ' MAX�@MO�l
    MAXFVAL As Double               ' MAX�@F�l
    REPRESIST As Double             ' ��\�@��R�l
    REPMOVAL As Double              ' ��\�@MO�l
    REPFVAL As Double               ' ��\�@F�l
    IonDensity As Double            ' ��\�C�I���Z�x
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
End Type


' ���������Z�x�}�X�^�[
Public Type typ_TBCMB008
    MltType As String * 3           ' �^�C�v
    MINRESIST As Double             ' MIN�@��R�l
    MAXRESIST As Double             ' MAX�@��R�l
    IonDensity As Double            ' ��\�C�I���Z�x
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
End Type


' �h�[�p���g�Z�x�}�X�^�[
Public Type typ_TBCMB009
    DopKind As String * 4           ' �h�[�p���g���
    IonDensity As Double            ' �C�I���Z�x
    CoreCoeff As Integer            ' �␳�W��
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
End Type


' �h�[�p���g�v�Z�W���}�X�^�[
Public Type typ_TBCMB010
    TYPE As String * 1              ' �^�C�v
    ResFrom As Double               ' ��RFrom
    ResTo As Double                 ' ��RTo
    FixNumA As Double               ' �萔A
    FizNumB As Double               ' �萔B
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
End Type


' PG-ID�Ǘ�
Public Type typ_TBCMB011
    PGID As String * 10             ' PG-ID
    HZPART As String * 4            ' HZ�p�[�c
    HZPTRN As String * 2            ' HZ�p�^�[��
    SPACER As String * 5            ' �X�y�[�T
    UPRING As String * 5            ' �A�b�p�[�����O
    CHARGE As Long                  ' �`���[�W��
    RTBPOS As Integer               ' ���c�{�ʒu
    RTBSIZE As String * 2           ' ���c�{�T�C�Y
    GAP As Integer                  ' �M���b�v
    UPDM As Integer                 ' ���㒼�a
    UPLENGTH As Integer             ' ���㒷�i�S���j
    UPRC As Integer                 ' ����iRC�j
    RFRNEED As String * 1           ' ���t���N�^�v��
    UPSPIN As String * 10           ' �㎲��]��
    DOWNSPIN As String * 10         ' ������]��
    ROPRESS As String * 8           ' �F����
    ARUGON As String * 7            ' �A���S����
    AIMOIMIN As Double              ' �˂炢Oi�iMIN)
    AIMOIMAX As Double              ' �˂炢Oi�iMAX)
    HCCLASS As String * 7           ' HC���
    HC As String                    ' HC
    AVEUPSPD As Double              ' ���ψ��㑬�x
    UPCNTL As String * 1            ' ���㐧��
    BTMSHAPE As String * 1          ' �{�g���`��
    MAGSTR As Double                ' ���ꋭ�x
    MAGPOS As Long                  ' ����ʒu
    CONDGRT As String * 10          ' �����ۏؓo�^
    MODEL As String * 4             ' �@��
    UPMETHOD As String * 4          ' ������@
    UPCLASS As String * 2           ' ����敪
    UPNUM As String * 1             ' ����{��
    OPETIME As Long                 ' �^�]����
    WTRCOOL As String * 1           ' ����Ǘv��
    PGID2 As String * 10            ' PG-ID�i��{���j
    RCPT1 As String * 3             ' �Ή����V�sNo�iT1)
    RCPT2 As String * 3             ' �Ή����V�sNo�iT2)
    RCPT3 As String * 3             ' �Ή����V�sNo�iT3)
    RCPT4 As String * 3             ' �Ή����V�sNo�iT4)
    RCPT5 As String * 3             ' �Ή����V�sNo�iT5)
    RCPT6 As String * 3             ' �Ή����V�sNo�iT6)
    CNTL1 As String * 1             ' �������ځi1�j
    CNTL2 As String * 1             ' �������ځi2�j
    CNTL3 As String * 1             ' �������ځi3�j
    CNTL4 As String * 1             ' �������ځi4�j
    CNTL5 As String * 1             ' �������ځi5�j
    CNTL6 As String * 1             ' �������ځi6�j
    CNTL7 As String * 1             ' �������ځi7�j
    CNTL8 As String * 1             ' �������ځi8�j
    CNTL9 As String * 1             ' �������ځi9�j
    CNTL10 As String * 1            ' �������ځi10�j
    CNTL11 As String * 1            ' �������ځi11�j
    CNTL12 As String * 1            ' �������ځi12�j
    CNTL13 As String * 1            ' �������ځi13�j
    CNTL14 As String * 1            ' �������ځi14�j
    CNTL15 As String * 1            ' �������ځi15�j
    DRDOP  As String                ' �h�[�v     4/30
    DRAR3  As String                ' �A���S�����R����
    RUNCOND1 As String              ' �^�]�����P
    RUNCOND2 As String              ' �^�]�����Q
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �������
Public Type typ_TBCMB012
    MKCONDNO As String * 12         ' �������No.
    MODEL As String * 1             ' �@��
    RTBSIZE As String * 1           ' ���c�{�T�C�Y
    CHARGE As String * 1            ' �`���[�W��
    HZTYPE As String * 1            ' HZ�^�C�v
    UPSPDTYP As String * 1          ' ���グ���x�^�C�v
    MAGTYPE As String * 1           ' ����^�C�v
    USECLS As String * 1            ' �g�p�敪
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M����
End Type


' �������PG-ID�Ή�
Public Type typ_TBCMB013
    MKCONDNO As String * 12         ' �������No.
    PGIDNO As String * 10           ' PG-IDNo
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' GFA�Z�����
Public Type typ_TBCMB014
    GOUKI As String * 3             ' ���@
    INPDATE As Date                 ' ���t
    FTIRFZI As Double               ' FTIR�iFZ)
    FTIRCZH As Double               ' FTIR�iCZ���j
    FTIRCZC As Double               ' FTIR�iCZ���j
    MS1FZ As Double                 ' ����T���v��1�iFZ)
    MS1CZ1 As Double                ' ����T���v��1�iCZ-1)
    MS1CZ2 As Double                ' ����T���v��1�iCZ-2)
    MS2FZ As Double                 ' ����T���v��2�iFZ)
    MS2CZ1 As Double                ' ����T���v��2�iCZ-1)
    MS2CZ2 As Double                ' ����T���v��2�iCZ-2)
    MS3FZ As Double                 ' ����T���v��3�iFZ)
    MS3CZ1 As Double                ' ����T���v��3�iCZ-1)
    MS3CZ2 As Double                ' ����T���v��3�iCZ-2)
    MS4FZ As Double                 ' ����T���v��4�iFZ)
    MS4CZ1 As Double                ' ����T���v��4�iCZ-1)
    MS4CZ2 As Double                ' ����T���v��4�iCZ-2)
    MS5FZ As Double                 ' ����T���v��5�iFZ)
    MS5CZ1 As Double                ' ����T���v��5�iCZ-1)
    MS5CZ2 As Double                ' ����T���v��5�iCZ-2)
    MSAVEFZ As Double               ' ���蕽�ρiFZ�j
    MSAVECZ1 As Double              ' ���蕽�ρiCZ-1�j
    MSAVECZ2 As Double              ' ���蕽�ρiCZ-2�j
    MSSGFZ As Double                ' ����ЁiFZ�j
    MSSGCZ1 As Double               ' ����ЁiCZ-1�j
    MSSGCZ2 As Double               ' ����ЁiCZ-2�j
    MSPSGFZ As Double               ' ����AVE+�ЁiFZ�j
    MSPSGCZ1 As Double              ' ����AVE+�ЁiCZ-1�j
    MSPSGCZ2 As Double              ' ����AVE+�ЁiCZ-2�j
    MSNSGFZ As Double               ' ����AVE-�ЁiFZ�j
    MSNSGCZ1 As Double              ' ����AVE-�ЁiCZ-1�j
    MSNSGCZ2 As Double              ' ����AVE-�ЁiCZ-2�j
    MINFZ As Double                 ' MIN�iFZ�j
    MINCZ1 As Double                ' MIN�iCZ-1�j
    MINCZ2 As Double                ' MIN�iCZ-2�j
    MAXFZ As Double                 ' MAX�iFZ�j
    MAXCZ1 As Double                ' MAX�iCZ-1�j
    MAXCZ2 As Double                ' MAX�iCZ-2�j
    SGCK1FZ As Double               ' ��ck�T���v��1�iFZ)
    SGCK1CZ1 As Double              ' ��ck�T���v��1�iCZ-1)
    SGCK1CZ2 As Double              ' ��ck�T���v��1�iCZ-2)
    SGCK2FZ As Double               ' ��ck�T���v��2�iFZ)
    SGCK2CZ1 As Double              ' ��ck�T���v��2�iCZ-1)
    SGCK2CZ2 As Double              ' ��ck�T���v��2�iCZ-2)
    SGCK3FZ As Double               ' ��ck�T���v��3�iFZ)
    SGCK3CZ1 As Double              ' ��ck�T���v��3�iCZ-1)
    SGCK3CZ2 As Double              ' ��ck�T���v��3�iCZ-2)
    SGCK4FZ As Double               ' ��ck�T���v��4�iFZ)
    SGCK4CZ1 As Double              ' ��ck�T���v��4�iCZ-1)
    SGCK4CZ2 As Double              ' ��ck�T���v��4�iCZ-2)
    SGCK5FZ As Double               ' ��ck�T���v��5�iFZ)
    SGCK5CZ1 As Double              ' ��ck�T���v��5�iCZ-1)
    SGCK5CZ2 As Double              ' ��ck�T���v��5�iCZ-2)
    SGCKDFZ As Double               ' ��ck�f�[�^���iFZ�j
    SGCKDCZ1 As Double              ' ��ck�f�[�^���iCZ-1�j
    SGCKDCZ2 As Double              ' ��ck�f�[�^���iCZ-2�j
    SGCKAFZ As Double               ' ��ck���ρiFZ�j
    SGCKAACZ1 As Double             ' ��ck���ρiCZ-1�j
    SGCKACZ2 As Double              ' ��ck���ρiCZ-2�j
    SGNFZ As Double                 ' ��ck�ЁiFZ�j
    SGNCZ1 As Double                ' ��ck�� CZ-1�j
    SGNCZ2 As Double                ' ��ck�ЁiCZ-2�j
    FTIRFZ As Double                ' FTIR���Z�iFZ�j
    FTIRCZ1 As Double               ' FTIR���Z�iCZ-1�j
    FTIRCZ2 As Double               ' FTIR���Z�iCZ-2�j
    EFFECTTM As Integer             ' �L������
    YCOEF As Double                 ' �e�s�h�q���Z���i�x�ؕЁj
    XCOEF As Double                 ' �e�s�h�q���Z���i�w�W���j
    RSQUARE As Double               ' �q�Q��
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
'2006/05/22�ǉ�
    SGCKST      As Double           ' �Д���
    SGCKFZ      As String * 1       ' �Д���(FZ)
    SGCKCZ1     As String * 1       ' �Д���(CZ-1)
    SGCKCZ2     As String * 1       ' �Д���(CZ-2)
    FTIRCKST    As Double           ' FTIR���Z����
    FTIRCKFZ    As String * 1       ' FTIR���Z����(FZ)
    FTIRCKCZ1   As String * 1       ' FTIR���Z����(CZ-1)
    FTIRCKCZ2   As String * 1       ' FTIR���Z����(CZ-2)
'2010/02/09�ǉ� SETsw kubota
    MS6FZ As Double                 ' ����T���v��6�iFZ)
    MS6CZ1 As Double                ' ����T���v��6�iCZ-1)
    MS6CZ2 As Double                ' ����T���v��6�iCZ-2)
    SGCK6FZ As Double               ' ��ck�T���v��6�iFZ)
    SGCK6CZ1 As Double              ' ��ck�T���v��6�iCZ-1)
    SGCK6CZ2 As Double              ' ��ck�T���v��6�iCZ-2)
    CVFZ As Double                  ' CV(%)�iFZ�j
    CVCZ1 As Double                 ' CV(%)�iCZ-1�j
    CVCZ2 As Double                 ' CV(%)�iCZ-2�j
End Type


' �A�ԊǗ�
Public Type typ_TBCMB015
    CNTMNGCD As String * 4          ' �A�Ԏ�ʊǗ��R�[�h
    CNTNUMCD As String * 4          ' �A�Ԏ�ʃR�[�h
    CONTNUM As Long                 ' �A��
    MAXFIG As Integer               ' �ő包��
    NUMUNIT As String * 1           ' �A�ԒP�ʋ敪
    NUMNAME As String               ' �A�Ԗ�
    CLRDATE As Date                 ' �N���A���t
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
End Type


' �o�[�W�����Ǘ�
Public Type typ_TBCMB016
    MACHINE As String * 8           ' �}�V����
    EXENAME As String               ' ���s�t�@�C����
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
End Type


' ���x���v�����^�v��
Public Type typ_TBCMB017
    QUEDATE As Date                 ' �L���[���t
    PRINTKIND As String * 4         ' ������
    ENDFLG As String * 1            ' �����敪
    STATUS As String * 4            ' �I���X�e�[�^�X
    PrintInfo1 As String            ' ������P
    PrintInfo2 As String            ' ������Q
    PrintInfo3 As String            ' ������R
    PrintInfo4 As String            ' ������S
    PrintInfo5 As String            ' ������T
    PrintInfo6 As String            ' ������U
    PrintInfo7 As String            ' ������V
    PrintInfo8 As String            ' ������W
    PrintInfo9 As String            ' ������X
    PrintInfo10 As String           ' ������P�O
    StaffID As String * 8           ' �v���S����ID
    MACHINE As String * 8           ' �v���}�V����
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
End Type


'Add Start 2011/03/30 SMPK H.Ohkubo
' FRS�Z�����
Public Type typ_TBCMB019
    GOUKI       As String * 3       ' ���@
    INPDATE     As Date             ' ���t
    FTIROIL     As Double           ' FTIR�iOi��)
    FTIROIM     As Double           ' FTIR�iOi���j
    FTIROIH     As Double           ' FTIR�iOi���j
    MS1OIL      As Double           ' ����T���v��1�iOi��)
    MS1OIM      As Double           ' ����T���v��1�iOi��)
    MS1OIH      As Double           ' ����T���v��1�iOi��)
    MS2OIL      As Double           ' ����T���v��2�iOi��)
    MS2OIM      As Double           ' ����T���v��2�iOi��)
    MS2OIH      As Double           ' ����T���v��2�iOi��)
    MS3OIL      As Double           ' ����T���v��3�iOi��)
    MS3OIM      As Double           ' ����T���v��3�iOi��)
    MS3OIH      As Double           ' ����T���v��3�iOi��)
    MS4OIL      As Double           ' ����T���v��4�iOi��)
    MS4OIM      As Double           ' ����T���v��4�iOi��)
    MS4OIH      As Double           ' ����T���v��4�iOi��)
    MS5OIL      As Double           ' ����T���v��5�iOi��)
    MS5OIM      As Double           ' ����T���v��5�iOi��)
    MS5OIH      As Double           ' ����T���v��5�iOi��)
    MSAVEOIL    As Double           ' ���蕽�ρiOi��j
    MSAVEOIM    As Double           ' ���蕽�ρiOi���j
    MSAVEOIH    As Double           ' ���蕽�ρiOi���j
    MSSGOIL     As Double           ' ����ЁiOi��j
    MSSGOIM     As Double           ' ����ЁiOi���j
    MSSGOIH     As Double           ' ����ЁiOi���j
    MSPSGOIL    As Double           ' ����AVE+�ЁiOi��j
    MSPSGOIM    As Double           ' ����AVE+�ЁiOi���j
    MSPSGOIH    As Double           ' ����AVE+�ЁiOi���j
    MSNSGOIL    As Double           ' ����AVE-�ЁiOi��j
    MSNSGOIM    As Double           ' ����AVE-�ЁiOi���j
    MSNSGOIH    As Double           ' ����AVE-�ЁiOi���j
    MINOIL      As Double           ' MIN�iOi��j
    MINOIM      As Double           ' MIN�iOi���j
    MINOIH      As Double           ' MIN�iOi���j
    MAXOIL      As Double           ' MAX�iOi��j
    MAXOIM      As Double           ' MAX�iOi���j
    MAXOIH      As Double           ' MAX�iOi���j
    SGCK1OIL    As Double           ' ��ck�T���v��1�iOi��)
    SGCK1OIM    As Double           ' ��ck�T���v��1�iOi��)
    SGCK1OIH    As Double           ' ��ck�T���v��1�iOi��)
    SGCK2OIL    As Double           ' ��ck�T���v��2�iOi��)
    SGCK2OIM    As Double           ' ��ck�T���v��2�iOi��)
    SGCK2OIH    As Double           ' ��ck�T���v��2�iOi��)
    SGCK3OIL    As Double           ' ��ck�T���v��3�iOi��)
    SGCK3OIM    As Double           ' ��ck�T���v��3�iOi��)
    SGCK3OIH    As Double           ' ��ck�T���v��3�iOi��)
    SGCK4OIL    As Double           ' ��ck�T���v��4�iOi��)
    SGCK4OIM    As Double           ' ��ck�T���v��4�iOi��)
    SGCK4OIH    As Double           ' ��ck�T���v��4�iOi��)
    SGCK5OIL    As Double           ' ��ck�T���v��5�iOi��)
    SGCK5OIM    As Double           ' ��ck�T���v��5�iOi��)
    SGCK5OIH    As Double           ' ��ck�T���v��5�iOi��)
    SGCKDOIL    As Double           ' ��ck�f�[�^���iOi��j
    SGCKDOIM    As Double           ' ��ck�f�[�^���iOi���j
    SGCKDOIH    As Double           ' ��ck�f�[�^���iOi���j
    SGCKAOIL    As Double           ' ��ck���ρiOi��j
    SGCKAAOIM   As Double           ' ��ck���ρiOi���j
    SGCKAOIH    As Double           ' ��ck���ρiOi���j
    SGNOIL      As Double           ' ��ck�ЁiOi��j
    SGNOIM      As Double           ' ��ck�ЁiOi���j
    SGNOIH      As Double           ' ��ck�ЁiOi���j
    FTIRKOIL    As Double           ' FTIR���Z�iOi��j
    FTIRKOIM    As Double           ' FTIR���Z�iOi���j
    FTIRKOIH    As Double           ' FTIR���Z�iOi���j
    EFFECTTM    As Integer          ' �L������
    YCOEF       As Double           ' �e�s�h�q���Z���i�x�ؕЁj
    XCOEF       As Double           ' �e�s�h�q���Z���i�w�W���j
    RSQUARE     As Double           ' �q�Q��
    SGCKST      As Double           ' �Д���
    SGCKOIL     As String * 1       ' �Д���(Oi��)
    SGCKOIM     As String * 1       ' �Д���(Oi��)
    SGCKOIH     As String * 1       ' �Д���(Oi��)
    FTIRCKST    As Double           ' FTIR���Z����
    FTIRCKOIL   As String * 1       ' FTIR���Z����(Oi��)
    FTIRCKOIM   As String * 1       ' FTIR���Z����(Oi��)
    FTIRCKOIH   As String * 1       ' FTIR���Z����(Oi��)
    MS6OIL      As Double           ' ����T���v��6�iOi��)
    MS6OIM      As Double           ' ����T���v��6�iOi��)
    MS6OIH      As Double           ' ����T���v��6�iOi��)
    SGCK6OIL    As Double           ' ��ck�T���v��6�iOi��)
    SGCK6OIM    As Double           ' ��ck�T���v��6�iOi��)
    SGCK6OIH    As Double           ' ��ck�T���v��6�iOi��)
    CVOIL       As Double           ' CV(%)�iOi��j
    CVOIM       As Double           ' CV(%)�iOi���j
    CVOIH       As Double           ' CV(%)�iOi���j
    TSTAFFID    As String * 8       ' �o�^�Ј�ID
    REGDATE     As Date             ' �o�^���t
    KSTAFFID    As String * 8       ' �X�V�Ј�ID
    UPDDATE     As Date             ' �X�V���t
    SENDFLAG    As String * 1       ' ���M�t���O
    SENDDATE    As Date             ' ���M���t
End Type
'Add End 2011/03/30 SMPK H.Ohkubo

' �������
Public Type typ_TBCME037
    CRYNUM As String * 12           ' �����ԍ�
    DELCLS As String * 1            ' �폜�敪
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCD As String * 5            ' �H���R�[�h
    LPKRPROCCD As String * 5        ' �ŏI�ʉߊǗ��H��
    LASTPASS As String * 5          ' �ŏI�ʉߍH��
    RPHINBAN As String * 8          ' �˂炢�i��
    RPREVNUM As Integer             ' �˂炢�i�Ԑ��i�ԍ������ԍ�
    RPFACT As String * 1            ' �˂炢�i�ԍH��
    RPOPCOND As String * 1          ' �˂炢�i�ԑ��Ə���
    PRODCOND As String * 12         ' �������
    PGID As String * 8              ' �o�f�|�h�c
    UPLENGTH As Integer             ' ���グ����
    TOPLENG As Integer              ' �s�n�o����
    BODYLENG As Integer             ' ��������
    BOTLENG As Integer              ' �a�n�s����
    FREELENG As Integer             ' �t���[��
    DIAMETER As Integer             ' ���a
    CHARGE As Long                  ' �`���[�W��
    SEED As String * 4              ' �V�[�h
    ADDDPCLS As String * 4          ' �ǉ��h�[�v���
    ADDDPPOS As Integer             ' �ǉ��h�[�v�ʒu
    ADDDPVAL As Double              ' �ǉ��h�[�v��
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �u���b�N�݌v
Public Type typ_TBCME038
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �������J�n�ʒu
    Length As Integer               ' ����
    USECLASS As String * 1          ' �g�p�敪
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �i�Ԑ݌v
Public Type typ_TBCME039
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �������J�n�ʒu
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' �����ԍ�
    FACT As String * 1              ' �H��
    OPCOND As String * 1            ' ���Ə���
    Length As Integer               ' ����
    USECLASS As String * 1          ' �g�p�敪
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �u���b�N�Ǘ�
Public Type typ_TBCME040
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �������J�n�ʒu
    Length As Integer               ' ����
    REALLEN As Integer              ' ������
    BLOCKID As String * 12          ' �u���b�NID
    KRPROCCD As String * 5          ' ���݊Ǘ��H��
    NOWPROC As String * 5           ' ���ݍH��
    LPKRPROCCD As String * 5        ' �ŏI�ʉߊǗ��H��
    LASTPASS As String * 5          ' �ŏI�ʉߍH��
    DELCLS As String * 1            ' �폜�敪
    LSTATCLS As String * 1          ' �ŏI��ԋ敪
    RSTATCLS As String * 1          ' ������ԋ敪
    HOLDCLS As String * 1           ' �z�[���h�敪
    BDCAUS As String * 3            ' �s�Ǘ��R
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    PASSFLAG As String * 1          ' �ʉ߃t���O�@�@'7/5�@hama
End Type


' �i�ԊǗ�
Public Type typ_TBCME041
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �������J�n�ʒu
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    Length As Integer               ' ����
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' SXL�Ǘ�
Public Type typ_TBCME042
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �������J�n�ʒu
    Length As Integer               ' ����
    SXLID As String * 13            ' SXLID
    KRPROCCD As String * 5          ' �Ǘ��H��
    NOWPROC As String * 5           ' ���ݍH��
    LPKRPROCCD As String * 5        ' �ŏI�ʉߊǗ��H��
    LASTPASS As String * 5          ' �ŏI�ʉߍH��
    DELCLS As String * 1            ' �폜�敪
    LSTATCLS As String * 1          ' �ŏI��ԋ敪
    HOLDCLS As String * 1           ' �z�[���h�敪
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    BDCAUS As String * 3            ' �s�Ǘ��R
    COUNT As Integer                ' ����
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    PASSFLAG As String * 1          ' �ʉ߃t���O�@�@'4/5�@Yam
End Type

' �����T���v���Ǘ�
Public Type typ_TBCME043
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �������ʒu
    SMPKBN As String * 1            ' �T���v���敪
    SMPLNO As Long                  ' �T���v��No    Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KTKBN As String * 1             ' �m��敪
    CRYINDRS As String * 1          ' ���������w���iRs)
    CRYINDOI As String * 1          ' ���������w���iOi)
    CRYINDB1 As String * 1          ' ���������w���iB1)
    CRYINDB2 As String * 1          ' ���������w���iB2�j
    CRYINDB3 As String * 1          ' ���������w���iB3)
    CRYINDL1 As String * 1          ' ���������w���iL1)
    CRYINDL2 As String * 1          ' ���������w���iL2)
    CRYINDL3 As String * 1          ' ���������w���iL3)
    CRYINDL4 As String * 1          ' ���������w���iL4)
    CRYINDCS As String * 1          ' ���������w���iCs)
    CRYINDGD As String * 1          ' ���������w���iGD)
    CRYINDT As String * 1           ' ���������w���iT)
    CRYINDEP As String * 1          ' ���������w���iEPD)
    CRYRESRS As String * 1          ' �����������сiRs)
    CRYRESOI As String * 1          ' �����������сiOi)
    CRYRESB1 As String * 1          ' �����������сiB1)
    CRYRESB2 As String * 1          ' �����������сiB2�j
    CRYRESB3 As String * 1          ' �����������сiB3)
    CRYRESL1 As String * 1          ' �����������сiL1)
    CRYRESL2 As String * 1          ' �����������сiL2)
    CRYRESL3 As String * 1          ' �����������сiL3)
    CRYRESL4 As String * 1          ' �����������сiL4)
    CRYRESCS As String * 1          ' �����������сiCs)
    CRYRESGD As String * 1          ' �����������сiGD)
    CRYREST As String * 1           ' �����������сiT)
    CRYRESEP As String * 1          ' �����������сiEPD)
    SMPLNUM As Integer              ' �T���v������
    SMPLPAT As String * 1           ' �T���v���p�^�[��
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    '2003/09/29 �ǉ��@KURO
    SMPLNOOI As Long                ' �T���v��No(OI)    Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLNOCS As Long                ' �T���v��No(CS)    Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    XTALCS   As String * 12         ' �����ԍ�
    '' 2007/11/26 y.hosokawa Update 15���Ή�
    BLOCKCS  As String * 15         ' �u���b�NID
    'BLOCKCS  As String * 12         ' �u���b�NID
    CRYSMPLIDRS1CS As Long          '                   Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYSMPLIDRS2CS As Long          '                   Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYRESRS1CS    As String * 1
    CRYRESRS2CS    As String * 1
    CRYSMPLIDB1CS  As Long          '                   Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYSMPLIDB2CS  As Long          '                   Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYSMPLIDB3CS  As Long          '                   Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYSMPLIDL1CS  As Long          '                   Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYSMPLIDL2CS  As Long          '                   Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYSMPLIDL3CS  As Long          '                   Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYSMPLIDL4CS  As Long          '                   Integer��Long   �T���v����6���Ή� 2007/05/28 SETsw kubota
    QCKBNCS As String * 1           ' �Ǘ��敪   2009/11/05 SETsw kubota
End Type

'' �V�T���v���Ǘ�(��ۯ�)
Public Type typ_XSDCS
    CRYNUMCS As String * 12         '�u���b�NID
    SMPKBNCS As String * 1          '�T���v���敪
    TBKBNCS As String * 1           'T/B�敪
    REPSMPLIDCS As Long             '��\�T���v��ID     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    XTALCS As String * 12           '�����ԍ�
    INPOSCS As Integer              '�������ʒu
    HINBCS As String * 8            '�i��
    REVNUMCS As Integer             '���i�ԍ������ԍ�
    FACTORYCS As String * 1         '�H��
    OPECS As String * 1             '���Ə���
    KTKBNCS As String * 1           '�m��敪
    BLKKTFLAGCS As String * 1       '�u���b�N�m��t���O
    CRYSMPLIDRSCS As Long           '�T���v��ID(Rs)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYSMPLIDRS1CS As Long          '����T���v��ID1(Rs)    Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYSMPLIDRS2CS As Long          '����T���v��ID2(Rs)    Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDRSCS As String * 1        '���FLG(Rs)
    CRYRESRS1CS As String * 1       '����FLG1(Rs)
    CRYRESRS2CS As String * 1       '����FLG2(Rs)
    CRYSMPLIDOICS As Long           '�T���v��ID(Oi)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDOICS As String * 1        '���FLG(Oi)
    CRYRESOICS As String * 1        '����FLG(Oi)
    CRYSMPLIDB1CS As Long           '�T���v��ID(B1)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDB1CS As String * 1        '���FLG(B1)
    CRYRESB1CS As String * 1        '����FLG(B1)
    CRYSMPLIDB2CS As Long           '�T���v��ID(B2)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDB2CS As String * 1        '���FLG(B2)
    CRYRESB2CS As String * 1        '����FLG(B2)
    CRYSMPLIDB3CS As Long           '�T���v��ID(B3)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDB3CS As String * 1        '���FLG(B3)
    CRYRESB3CS As String * 1        '����FLG(B3)
    CRYSMPLIDL1CS As Long           '�T���v��ID(L1)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDL1CS As String * 1        '���FLG(L1)
    CRYRESL1CS As String * 1        '����FLG(L1)
    CRYSMPLIDL2CS As Long           '�T���v��ID(L2)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDL2CS As String * 1        '���FLG(L2)
    CRYRESL2CS As String * 1        '����FLG(L2)
    CRYSMPLIDL3CS As Long           '�T���v��ID(L3)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDL3CS As String * 1        '���FLG(L3)
    CRYRESL3CS As String * 1        '����FLG(L3)
    CRYSMPLIDL4CS As Long           '�T���v��ID(L4)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDL4CS As String * 1        '���FLG(L4)
    CRYRESL4CS As String * 1        '����FLG(L4)
    CRYSMPLIDCSCS As Long           '�T���v��ID(Cs)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDCSCS As String * 1        '���FLG(Cs)
    CRYRESCSCS As String * 1        '����FLG(Cs)
    CRYSMPLIDGDCS As Long           '�T���v��ID(GD)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDGDCS As String * 1        '���FLG(GD)
    CRYRESGDCS As String * 1        '����FLG(GD)
    CRYSMPLIDTCS As Long            '�T���v��ID(T)      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDTCS As String * 1         '���FLG(T)
    CRYRESTCS As String * 1         '����FLG(T)
    CRYSMPLIDEPCS As Long           '�T���v��ID(EPD)    Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    CRYINDEPCS As String * 1        '���FLG(EPD)
    CRYRESEPCS As String * 1        '����FLG(EPD)
    SMPLNUMCS As Integer            '�T���v������
    SMPLPATCS As String * 1         '�T���v���p�^�[��
    LIVKCS As String * 1            '�����敪
    TSTAFFCS As String * 8          '�o�^�Ј�ID
    TDAYCS As Date                  '�o�^���t
    KSTAFFCS As String * 8          '�X�V�Ј�ID
    KDAYCS As Date                  '�X�V���t
    SNDKCS As String * 1            '���M�t���O
    SNDDAYCS As Date                '���M���t
    RPCRYNUMCS As String * 12       '�e�u���b�NID�@05/10/17 ooba
    CUTFLGCS As String * 1          '�ؒf�t���O�@05/10/17 ooba
    
    '2009/08 SUMCO Akizuki X��������с@���ڒǉ�
    CRYSMPLIDXCS As Long            '�T���v��ID(X��)
    CRYINDXCS As String * 1         '���FLG(X��)
    CRYRESXCS As String * 1         '����FLG(X��)
    
    QCKBNCS As String * 1           '�Ǘ��敪   2009/11/05 SETsw kubota

    'Add Start 2010/12/13 SMPK Miyata
    CRYSMPLIDCCS    As Long         '�T���v��ID(C)
    CRYINDCCS       As String * 1   '���FLG(C)
    CRYRESCCS       As String * 1   '����FLG(C)
    CRYSMPLIDCJCS   As Long         '�T���v��ID(CJ)
    CRYINDCJCS      As String * 1   '���FLG(CJ)
    CRYRESCJCS      As String * 1   '����FLG(CJ)
    CRYSMPLIDCJLTCS As Long         '�T���v��ID(CJLT)
    CRYINDCJLTCS    As String * 1   '���FLG(CJLT)
    CRYRESCJLTCS    As String * 1   '����FLG(CJLT)
    CRYSMPLIDCJ2CS  As Long         '�T���v��ID(CJ2)
    CRYINDCJ2CS     As String * 1   '���FLG(CJ2)
    CRYRESCJ2CS     As String * 1   '����FLG(CJ2)
    'Add End   2010/12/13 SMPK Miyata

End Type

'2003/09/1 ���ı�� SystemBrain
' WF�T���v���Ǘ�
'Public Type typ_TBCME044
'    CRYNUM As String * 12           ' �����ԍ�
'    IngotPos As Integer             ' �������ʒu
'    SMPKBN As String * 1            ' �T���v���敪
'    SMPLID As String * 16           ' �T���v��ID
'    BKSMPLID As String * 16         ' �ύX�O�T���v��ID  'add 2003/05/06 hitec)matsumoto
'    hinban As String * 8            ' �i��
'    REVNUM As Integer               ' ���i�ԍ������ԍ�
'    factory As String * 1           ' �H��
'    opecond As String * 1           ' ���Ə���
'    KTKBN As String * 1             ' �m��敪
'    WFINDRS As String * 1           ' WF�����w���iRs)
'    WFINDOI As String * 1           ' WF�����w���iOi)
'    WFINDB1 As String * 1           ' WF�����w���iB1)
'    WFINDB2 As String * 1           ' WF�����w���iB2�j
'    WFINDB3 As String * 1           ' WF�����w���iB3)
'    WFINDL1 As String * 1           ' WF�����w���iL1)
'    WFINDL2 As String * 1           ' WF�����w���iL2)
'    WFINDL3 As String * 1           ' WF�����w���iL3)
'    WFINDL4 As String * 1           ' WF�����w���iL4)
'    WFINDDS As String * 1           ' WF�����w���iDS)
'    WFINDDZ As String * 1           ' WF�����w���iDZ)
'    WFINDSP As String * 1           ' WF�����w���iSP)
'    WFINDDO1 As String * 1          ' WF�����w���iDO1)
'    WFINDDO2 As String * 1          ' WF�����w���iDO2)
'    WFINDDO3 As String * 1          ' WF�����w���iDO3)
'    WFINDOT1 As String * 1          ' WF�����w���iOT1)  'Add.03/05/20
'    WFINDOT2 As String * 1          ' WF�����w���iOT2)  'Add.03/05/20
'    WFRESRS As String * 1           ' WF�������сiRs)
'    WFRESOI As String * 1           ' WF�������сiOi)
'    WFRESB1 As String * 1           ' WF�������сiB1)
'    WFRESB2 As String * 1           ' WF�������сiB2�j
'    WFRESB3 As String * 1           ' WF�������сiB3)
'    WFRESL1 As String * 1           ' WF�������сiL1)
'    WFRESL2 As String * 1           ' WF�������сiL2)
'    WFRESL3 As String * 1           ' WF�������сiL3)
'    WFRESL4 As String * 1           ' WF�������сiL4)
'    WFRESDS As String * 1           ' WF�������сiDS)
'    WFRESDZ As String * 1           ' WF�������сiDZ)
'    WFRESSP As String * 1           ' WF�������сiSP)
'    WFRESDO1 As String * 1          ' WF�������сiDO1)
'    WFRESDO2 As String * 1          ' WF�������сiDO2)
'    WFRESDO3 As String * 1          ' WF�������сiDO3)
'    WFRESOT1 As String * 1          ' WF�������сiOT1)  'Add.03/05/20
'    WFRESOT2 As String * 1          ' WF�������сiOT2)  'Add.03/05/20
'    REGDATE As Date                 ' �o�^���t
'    UPDDATE As Date                 ' �X�V���t
'    SENDFLAG As String * 1          ' ���M�t���O
'    SENDDATE As Date                ' ���M���t
'    BkIngotPos  As Integer          ' add 2003/03/28 hitec)matsumoto
'End Type

''�V�T���v���Ǘ�(SXL)
Public Type typ_XSDCW
    SXLIDCW As String * 13          'SXLID
    SMPKBNCW As String * 1          '�T���v���敪
    TBKBNCW As String * 1           'T/B�敪
    REPSMPLIDCW As String * 16      '��\�T���v��ID
    XTALCW As String * 12           '�����ԍ�
    INPOSCW As Integer              '�������ʒu
    HINBCW As String * 8            '�i��
    REVNUMCW As Integer             '���i�ԍ������ԍ�
    FACTORYCW As String * 1         '�H��
    OPECW As String * 1             '���Ə���
    KTKBNCW As String * 1           '�m��敪
    SMCRYNUMCW As String * 12       '�T���v���u���b�NID
    WFSMPLIDRSCW As String * 16     '�T���v��ID(Rs)
    WFSMPLIDRS1CW As String * 16    '����T���v��ID1(Rs)
    WFSMPLIDRS2CW As String * 16    '����T���v��ID2(Rs)
    WFINDRSCW As String * 1         '���FLG(Rs)
    WFRESRS1CW As String * 1        '����FLG1(Rs)
    WFRESRS2CW As String * 1        '����FLG2(Rs)
    WFSMPLIDOICW As String * 16     '�T���v��ID(Oi)
    WFINDOICW As String * 1         '���FLG(Oi)
    WFRESOICW As String * 1         '����FLG(Oi)
    WFSMPLIDB1CW As String * 16     '�T���v��ID(B1)
    WFINDB1CW As String * 1         '���FLG(B1)
    WFRESB1CW As String * 1         '����FLG(B1)
    WFSMPLIDB2CW As String * 16     '�T���v��ID(B2)
    WFINDB2CW As String * 1         '���FLG(B2)
    WFRESB2CW As String * 1         '����FLG(B2)
    WFSMPLIDB3CW As String * 16     '�T���v��ID(B3)
    WFINDB3CW As String * 1         '���FLG(B3)
    WFRESB3CW As String * 1         '����FLG(B3)
    WFSMPLIDL1CW As String * 16     '�T���v��ID(L1)
    WFINDL1CW As String * 1         '���FLG(L1)
    WFRESL1CW As String * 1         '����FLG(L1)
    WFSMPLIDL2CW As String * 16     '�T���v��ID(L2)
    WFINDL2CW As String * 1         '���FLG(L2)
    WFRESL2CW As String * 1         '����FLG(L2)
    WFSMPLIDL3CW As String * 16     '�T���v��ID(L3)
    WFINDL3CW As String * 1         '���FLG(L3)
    WFRESL3CW As String * 1         '����FLG(L3)
    WFSMPLIDL4CW As String * 16     '�T���v��ID(L4)
    WFINDL4CW As String * 1         '���FLG(L4)
    WFRESL4CW As String * 1         '����FLG(L4)
    WFSMPLIDDSCW As String * 16     '�T���v��ID(DS)
    WFINDDSCW As String * 1         '���FLG(DS)
    WFRESDSCW As String * 1         '����FLG(DS)
    WFSMPLIDDZCW As String * 16     '�T���v��ID(DZ)
    WFINDDZCW As String * 1         '���FLG(DZ)
    WFRESDZCW As String * 1         '����FLG(DZ)
    WFSMPLIDSPCW As String * 16     '�T���v��ID(SP)
    WFINDSPCW As String * 1         '���FLG(SP)
    WFRESSPCW As String * 1         '����FLG(SP)
    WFSMPLIDDO1CW As String * 16    '�T���v��ID(DO1)
    WFINDDO1CW As String * 1        '���FLG(DO1)
    WFRESDO1CW As String * 1        '����FLG(DO1)
    WFSMPLIDDO2CW As String * 16    '�T���v��ID(DO2)
    WFINDDO2CW As String * 1        '���FLG(DO2)
    WFRESDO2CW As String * 1        '����FLG(DO2)
    WFSMPLIDDO3CW As String * 16    '�T���v��ID(DO3)
    WFINDDO3CW As String * 1        '���FLG(DO3)
    WFRESDO3CW As String * 1        '����FLG(DO3)
    WFSMPLIDOT1CW As String * 16    '�T���v��ID(OT1)
    WFINDOT1CW As String * 1        '���FLG(OT1)
    WFRESOT1CW As String * 1        '����FLG(OT1)
    WFSMPLIDOT2CW As String * 16    '�T���v��ID(OT2)
    WFINDOT2CW As String * 1        '���FLG(OT2)
    WFRESOT2CW As String * 1        '����FLG(OT2)
    WFSMPLIDAOICW As String * 16    '�T���v��ID(AOi)
    WFINDAOICW As String * 1        '���FLG(AOi)
    WFRESAOICW As String * 1        '����FLG(AOi)
    SMPLNUMCW As Integer            '�T���v������
    SMPLPATCW As String * 1         '�T���v���p�^�[��
    LIVKCW As String * 1            '�����敪�@'�ǉ� 2003/10/04
    TSTAFFCW As String * 8          '�o�^�Ј�ID
    TDAYCW As Date                  '�o�^���t
    KSTAFFCW As String * 8          '�X�V�Ј�ID
    KDAYCW As Date                  '�X�V���t
    SNDKCW As String * 1            '���M�t���O
    SNDDAYCW As Date                '���M���t
    WFSMPLIDGDCW As String * 16     '�T���v��ID(GD)     '05/01/17 ooba START =======>
    WFINDGDCW As String * 1         '���FLG(GD)
    WFRESGDCW As String * 1         '����FLG(GD)
    WFHSGDCW As String * 1          '�ۏ�FLG(GD)        '05/01/17 ooba END =========>
'    BkIngotPos  As Integer          ' add 2003/03/28 hitec)matsumoto
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    EPSMPLIDB1CW As String * 16     '�T���v��ID(BMD1)
    EPINDB1CW As String * 1         '���FLG(BMD1)
    EPRESB1CW As String * 1         '����FLG(BMD1)
    EPSMPLIDB2CW As String * 16     '�T���v��ID(BMD2)
    EPINDB2CW As String * 1         '���FLG(BMD2)
    EPRESB2CW As String * 1         '����FLG(BMD2)
    EPSMPLIDB3CW As String * 16     '�T���v��ID(BMD3)
    EPINDB3CW As String * 1         '���FLG(BMD3)
    EPRESB3CW As String * 1         '����FLG(BMD3)
    EPSMPLIDL1CW As String * 16     '�T���v��ID(OSF1)
    EPINDL1CW As String * 1         '���FLG(OSF1)
    EPRESL1CW As String * 1         '����FLG(OSF1)
    EPSMPLIDL2CW As String * 16     '�T���v��ID(OSF2)
    EPINDL2CW As String * 1         '���FLG(OSF2)
    EPRESL2CW As String * 1         '����FLG(OSF2)
    EPSMPLIDL3CW As String * 16     '�T���v��ID(OSF3)
    EPINDL3CW As String * 1         '���FLG(OSF3)
    EPRESL3CW As String * 1         '����FLG(OSF3)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
End Type


' �ؒf�w��
Public Type typ_TBCME045
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �������J�n�ʒu
    TRANCNT As Integer              ' ������
    Length As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    StaffID As String * 8           ' �Ј�ID
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' �i�Ԑ��i�ԍ������ԍ�
    FACTORY As String * 1           ' �i�ԍH��
    OPECOND As String * 1           ' �i�ԑ��Ə���
    BDCAUS As String * 3            ' �敪�R�[�h
    STATCLS As String * 1           ' ��ԋ敪
    BLOCKID As String * 12          ' �u���b�NID
    CRYINDRS As String * 1          ' ���FLG�iRs)
    CRYINDOI As String * 1          ' ���FLG�iOi)
    CRYINDB1 As String * 1          ' ���FLG�iB1)
    CRYINDB2 As String * 1          ' ���FLG�iB2�j
    CRYINDB3 As String * 1          ' ���FLG�iB3)
    CRYINDL1 As String * 1          ' ���FLG�iL1)
    CRYINDL2 As String * 1          ' ���FLG�iL2)
    CRYINDL3 As String * 1          ' ���FLG�iL3)
    CRYINDL4 As String * 1          ' ���FLG�iL4)
    CRYINDCS As String * 1          ' ���FLG�iCs)
    CRYINDGD As String * 1          ' ���FLG�iGD)
    CRYINDT As String * 1           ' ���FLG�iT)
    CRYINDEP As String * 1          ' ���FLG�iEPD)
    PRIORITY As String * 1          ' �D��x
    PALTNUM As String * 4           ' �p���b�g�ԍ�
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �������������
Public Type typ_TBCMG001
    MTRLNUM As String * 10          ' �����ԍ�
    JDATE As Date                   ' ���t
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    MTRLTYPE As String * 3          ' �������
    MAKERNO As String * 6           ' ���[�J�Ǘ�No
    RVWEIGHT As Long                ' ����w���d��
    CRYCOMMENT As String            ' �R�����g
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј��h�c
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �w���P�����������
Public Type typ_TBCMG002
    CRYNUM As String * 12           ' �����ԍ�
    TRANCNT As Integer              ' ������
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    REPCCL As String * 1            ' �������敪
    RBATCHNO As String * 10         ' �F�o�b�`�m��
    DMTOP1 As Integer               ' ���a�s�n�o�P
    DMTOP2 As Integer               ' ���a�s�n�o�Q
    DMTAIL1 As Integer              ' ���a�s�`�h�k�P
    DMTAIL2 As Integer              ' ���a�s�`�h�k�Q
    NCHDPTH1 As Integer             ' �m�b�`�[���P
    NCHDPTH2 As Integer             ' �m�b�`�[���Q
    UPLENGTH As Integer             ' ���グ��
    SXLPOS As Integer               ' �r�w�k�ʒu
    BlkLen As Integer               ' �u���b�N����
    BLKWGHT As Long                 ' �u���b�N�d��
    CMPTOP1 As Double               ' ���RTOP�@�P
    CMPTOP2 As Double               ' ���RTOP�@�Q
    CMPTOP3 As Double               ' ���RTOP�@�R
    CMPTOP4 As Double               ' ���RTOP�@�S
    CMPTOP5 As Double               ' ���RTOP�@�T
    CMPTOPR As Double               ' ���RTOP�@RRG
    CMPTAIL1 As Double              ' ���RTAIL�@�P
    CMPTAIL2 As Double              ' ���RTAIL�@�Q
    CMPTAIL3 As Double              ' ���RTAIL�@�R
    CMPTAIL4 As Double              ' ���RTAIL�@�S
    CMPTAIL5 As Double              ' ���RTAIL�@�T
    CMPTAILR As Double              ' ���RTAIL�@RRG
    OITOP1 As Double                ' Oi�@TOP�@�P
    OITOP2 As Double                ' Oi�@TOP�@�Q
    OITOP3 As Double                ' Oi�@TOP�@�R
    OITOP4 As Double                ' Oi�@TOP�@�S
    OITOP5 As Double                ' Oi�@TOP�@�T
    OITOPR As Double                ' Oi�@TOP�@ROG
    OITAIL1 As Double               ' Oi�@TAIL�@�P
    OITAIL2 As Double               ' Oi�@TAIL�@�Q
    OITAIL3 As Double               ' Oi�@TAIL�@�R
    OITAIL4 As Double               ' Oi�@TAIL�@�S
    OITAIL5 As Double               ' Oi�@TAIL�@�T
    OITAILR As Double               ' Oi�@TAIL�@ROG
    CSTOP As Double                 ' Cs�@TOP
    CSTAIL As Double                ' Cs�@TAIL
    LD1TOPMX As Double              ' LD-1�@TOP�@MAX
    LD1TOPAV As Double              ' LD-1�@TOP�@AVE
    LD1TAILM As Double              ' LD-1�@TAIL�@MAX
    LD1TAILA As Double              ' LD-1�@TAIL�@AVE
    LD2TOPMM As Double              ' LD-2�@TOP�@MAX
    LD2TOPAV As Double              ' LD-2�@TOP�@AVE
    LD2TAILM As Double              ' LD-2�@TAIL�@MAX
    LD2TAILA As Double              ' LD-2�@TAIL�@AVE
    BMDTOPMX As Double              ' BMD�@TOP�@MAX
    BMDTOPAV As Double              ' BMD�@TOP�@AVE
    BMDTAILM As Double              ' BMD�@TAIL�@MAX
    BMDTAILA As Double              ' BMD�@TAIL�@AVE
    GD1TOP As Integer               ' GD1 TOP
    GD1TAIL As Integer              ' GD1 TAIL
    GD2TOP As Integer               ' GD2 TOP
    GD2TAIL As Integer              ' GD2 TAIL
    DIA1TOP As Integer              ' DIA1 TOP
    DIA1TAIL As Integer             ' DIA1 TAIL
    DIA2TOP As Integer              ' DIA2 TOP
    DIA2TAIL As Integer             ' DIA2 TAIL
    LTFTOP As Integer               ' LIFETIME from TOP
    LTFTAIL As Integer              ' LIFETIME from TAIL
    EPD As Integer                  ' EPD
    HCNO As String * 10             ' ����No
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �������g����ؒf����
Public Type typ_TBCMG003
    CRYNUM As String * 12           ' �����ԍ�
    ROCLASS As String * 3           ' �ϋ敪
    HRCLASS As String * 1           ' �p���E���X�敪
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    Weight As Long                  ' �d��
    RMSHAPE As String * 1           ' �������g�`��
    RMMTRLNUM As String * 10        ' �������g�����ԍ�
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �������g��򕥏o����
Public Type typ_TBCMG004
    MTRLNUM As String * 10          ' �����ԍ�
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    DRWEIGHT As Long                ' ������d��
    LSWEIGHT As Long                ' ���X�d��
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SUMITSENDFLAG As String * 1     ' SUMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �����݌ɊǗ�
Public Type typ_TBCMG005
    MTRLNUM As String * 10          ' �����ԍ�
    USABLCLS As String * 1          ' �g�p�\�敪
    Weight As Long                  ' �d��
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
End Type


' �����݌Ɏ���
Public Type typ_TBCMG006
    MTRLNUM As String * 10          ' �����ԍ�
    TRANCNT As Long                 ' ������
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    Class As String * 1             ' �敪
    INWEIGHT As Long                ' ���͏d��
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �N���X�^���J�^���O�������
Public Type typ_TBCMG007
    CRYNUM As String * 12           ' �����ԍ�
    TRANCNT As Integer              ' ������
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    BDCODE As String * 3            ' �s�Ǘ��R�R�[�h
    PALTNUM As String * 4           ' �p���b�g�ԍ�
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �i�グ����
Public Type typ_TBCMG008
    CRYNUM As String * 12           ' �����ԍ��i�i�グ�j
    TRANCNT As Integer              ' ������
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    NHINBAN As String * 8           ' �V�i��
    NMNOREVNO As Integer            ' �V���i�ԍ������ԍ�
    NFACTORY As String * 1          ' �V�H��
    NOPECOND As String * 1          ' �V���Ə���
    OHINBAN As String * 8           ' ���i��
    OMNOREVNO As Integer            ' �����i�ԍ������ԍ�
    OFACTORY As String * 1          ' ���H��
    OOPECOND As String * 1          ' �����Ə���
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ���グ�w������
Public Type typ_TBCMH001
    UPINDNO As String * 9           ' ���グ�w��No.
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    MODEL As String * 4             ' �@��
    GOUKI As String * 3             ' ���@
    PGID As String * 10             ' PG-ID
    CPORGIND As String * 12         ' ���ʌ��w��No
    hinban As String * 8            ' �i��
    NMNOREVNO As Integer            ' ���i�ԍ������ԍ�
    NFACTORY As String * 1          ' �H��
    NOPECOND As String * 1          ' ���Ə���
    NUMNOTE1 As String              ' �i�Ԕ��l�P
    NUMNOTE2 As String              ' �i�Ԕ��l�Q
    SEED As String * 4              ' �V�[�h
    SEKIERTB As String * 7          ' �Ήp���c�{
    DPNTCLS As String * 7           ' �h�[�p���g���
    DOPANT As Double                ' �h�[�p���g��
    AMRESIST As Double              ' �˂炢��R
    CRYDOPCL As String * 7          ' �����h�[�v���
    CRYDOPVL As Double              ' �����h�[�v��
    UPBTCHNM As Integer             ' ���グ�o�b�`��
    ADDDOPCL As String * 7          ' �ǉ��h�[�p���g���
    ADDDOPVL As Double              ' �ǉ��h�[�p���g��
    ADDDOPPT As Integer             ' �ǉ��h�[�p���g�ʒu
    BCNT1COD As String * 3          ' �o�b�`���l1�i�R�[�h�j
    BCNT1CMT As String              ' �o�b�`���l1�i���āj
    BCNT2COD As String * 3          ' �o�b�`���l2�i�R�[�h�j
    BCNT2CMT As String              ' �o�b�`���l2�i���āj
    MTCLS1 As String * 3            ' �������1
    MTWGHT1 As Long                 ' �����d��1
    ESWGHT1 As Long                 ' ����c�d��1
    MTCLS2 As String * 3            ' �������2
    MTWGHT2 As Long                 ' �����d��2
    ESWGHT2 As Long                 ' ����c�d��2
    MTCLS3 As String * 3            ' �������3
    MTWGHT3 As Long                 ' �����d��3
    ESWGHT3 As Long                 ' ����c�d��3
    MTCLS4 As String * 3            ' �������4
    MTWGHT4 As Long                 ' �����d��4
    ESWGHT4 As Long                 ' ����c�d��4
    MTCLS5 As String * 3            ' �������5
    MTWGHT5 As Long                 ' �����d��5
    ESWGHT5 As Long                 ' ����c�d��5
    MTCLS6 As String * 3            ' �������6
    MTWGHT6 As Long                 ' �����d��6
    ESWGHT6 As Long                 ' ����c�d��6
    MTCLS7 As String * 3            ' �������7
    MTWGHT7 As Long                 ' �����d��7
    ESWGHT7 As Long                 ' ����c�d��7
    MTCLS8 As String * 3            ' �������8
    MTWGHT8 As Long                 ' �����d��8
    ESWGHT8 As Long                 ' ����c�d��8
    MTCLS9 As String * 3            ' �������9
    MTWGHT9 As Long                 ' �����d��9
    ESWGHT9 As Long                 ' ����c�d��9
    MTCLS10 As String * 3           ' �������10
    MTWGHT10 As Long                ' �����d��10
    ESWGHT10 As Long                ' ����c�d��10
    MTCLS11 As String * 3           ' �������11
    MTWGHT11 As Long                ' �����d��11
    ESWGHT11 As Long                ' ����c�d��11
    MTCLS12 As String * 3           ' �������12
    MTWGHT12 As Long                ' �����d��12
    ESWGHT12 As Long                ' ����c�d��12
    MTCLS13 As String * 3           ' �������13
    MTWGHT13 As Long                ' �����d��13
    ESWGHT13 As Long                ' ����c�d��13
    MTCLS14 As String * 3           ' �������14
    MTWGHT14 As Long                ' �����d��14
    ESWGHT14 As Long                ' ����c�d��14
    MTCLS15 As String * 3           ' �������15
    MTWGHT15 As Long                ' �����d��15
    ESWGHT15 As Long                ' ����c�d��15
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ���グ��������
Public Type typ_TBCMH002
    UPINDNO As String * 9           ' ���グ�w��No
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    TRANCLS As String * 1           ' �����敪
    PGID As String * 8              ' PG-ID
    TYPE As String * 2              ' �^�C�v
    SEKIERTB As String * 10         ' �Ήp���c�{
    DPNTCLS As String * 7           ' �h�[�p���g���
    DOPANT As Double                ' �h�[�p���g��
    CRYDOP As String * 1            ' �����h�[�v
    CRYDOPVL As Double              ' �����h�[�v��
    SEED As String * 4              ' �V�[�h
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ���グ������������
Public Type typ_TBCMH003
    UPINDNO As String * 9           ' ���グ�w��No
    MTRLNUM As String * 10          ' �����ԍ�
    Weight As Long                  ' �d��
    ESWEIGHT As Long                ' ����c�d��
End Type


' ���グ�I������
Public Type typ_TBCMH004
    CRYNUM As String * 12           ' �����ԍ�
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    LENGTOP As Integer              ' �����iTOP�j
    LENGTKDO As Integer             ' �����i�����j
    LENGTAIL As Integer             ' �����iTAIL�j
    LENGFREE As Integer             ' �t���[����
    DM1 As Double                   ' �������a�P
    DM2 As Double                   ' �������a�Q
    DM3 As Double                   ' �������a�R
    WGHTTOP As Long                 ' �d�ʁiTOP�j
    WGHTTKDO As Long                ' �d�ʁi�����j
    WGHTTAIL As Long                ' �d�ʁiTAIL)
    WGHTFREE As Long                ' �d�ʁi�t���[�����j
    WGTOPCUT As Long                ' �g�b�v�J�b�g�d��
    UPWEIGHT As Long                ' ���グ�d��
    CHARGE As Long                  ' �`���[�W��
    SEED As String * 4              ' �V�[�h
    STATCLS As String * 3           ' BOT�󋵋敪
    JDGECODE As String * 3          ' ����R�[�h
    PWTIME As Double                ' �p���[����
    ADDDPPOS As Integer             ' �ǉ��h�[�v�ʒu
    ADDDPCLS As String * 7          ' �ǉ��h�[�p���g���
    ADDDPVAL As Double              ' �ǉ��h�[�v��
    ADDDPNAM As String              ' �ǉ��h�[�v��
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    PULENTKC1 As Long               ' ���㒼������  2010/09/06 add Kameda
    PUWGHTTKC1 As Long              ' ���㒼���d��  2010/09/06 add Kameda
End Type


' ���グ�I���c�d�ʎ���
Public Type typ_TBCMH005
    RSCRYNUM As String * 12         ' �c�d�ʌ����ԍ�
    CRYNUM As String * 9            ' �������ԍ�
    RSWEIGHT As Long                ' �c�d��
End Type


' ����@�퍆�@���Ǘ�
Public Type typ_TBCMH006
    MODEL As String * 4             ' �@��
    GOUKI As String * 3             ' ���@
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    PROCDATE As Date                ' ���t
    CRYNUM As String * 12           ' �����ԍ�
    hinban As String * 8            ' �i��
    PGID As String * 10             ' PG-ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
End Type


' ���H���o����
Public Type typ_TBCMI001
    CRYNUM As String * 12           ' �����ԍ�
    TRANCNT As Integer              ' ������
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    UPLENGTH As Integer             ' ���グ����
    FREELENG As Integer             ' �t���[��
    UPWEIGHT As Long                ' ���グ�d��
    SEED As String * 4              ' �V�[�h
    PRCMCN As String * 1            ' ����@
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    TOPSFTQTY As String * 4         ' �g�b�v�V�t�g��  add 06/11/13 SET/Miyazaki
    BOTSFTQTY As String * 4         ' �{�g���V�t�g��  add 06/11/13 SET/Miyazaki
End Type


' ������H����
Public Type typ_TBCMI002
    CRYNUM As String * 12           ' �����ԍ�
    TRANCNT As Integer              ' ������
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    DMTOP1 As Double                ' ���aTOP�P
    DMTOP2 As Double                ' ���aTOP�Q
    DMTAIL1 As Double               ' ���aTAIL�P
    DMTAIL2 As Double               ' ���aTAIL�Q
    NCHPOS As String * 2            ' �m�b�`�ʒu
    NCHDPTH As Double               ' �m�b�`�[��
    NCHWIDTH As Double              ' �m�b�`��
    BDLNTOP As Integer              ' �s�ǒ����iTOP�j
    BDCDTOP As String * 3           ' �s�ǔ���R�[�h�iTOP�j
    BDLNTAIL As Integer             ' �s�ǒ����iTAIL)
    BDCDTAIL As String * 3          ' �s�ǔ���R�[�h�iTAIL�j
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    INGOTPOS As Integer             ' �C���S�b�g���ʒu
    Length As Integer               ' ����
    GOUKI As String * 5             ' ���@             2003/06/12 osawa
    NCHWTAIL As Double              ' �m�b�`�[��(TAIL) 2004/05/25
    BLOCKID As String * 12          ' ��ۯ�ID          2006/02/01 tuku
    CYGRTIM As String * 6           ' ���펞��         2006/11/09 SETsw J.W
    NCHLENGTH As String * 4         ' �m�b�`����       2006/11/09 SETsw J.W
    SOPROCTIM As String * 6         ' ���H����(�e��)   2006/11/20 SETsw Y.M
    SEIPROCTIM As String * 6        ' ���H����(����)   2006/11/20 SETsw Y.M
    NCHANGLE As Double              ' �m�b�`�p�x�@     2009/09    SUMCO Akizuki
End Type


' �ؒf����
Public Type typ_TBCMI003
    CRYNUM As String * 12           ' �����ԍ�
    TRANCNT As Integer              ' ������
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    GOUKI As String                 ' ���@�@�@3/13 Yam
    BLOCKID As String * 12          ' ��ۯ�ID  2006/02/01 tuku
    cutNum As String * 2            ' �J�b�g��
    CUTNUM2 As String * 4           ' �J�b�g��2 2006/11/02 SETsw
    PROCTIM As String * 6           ' ���H����  2006/11/02 SETsw
End Type


' EPD����
Public Type typ_TBCMJ001
    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
    TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    GOUKI As String * 3             ' ���@
    MEASURE As Integer              ' ����l
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' ������R����
Public Type typ_TBCMJ002
    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
    TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    GOUKI As String * 3             ' ���@
    TYPE As String * 1              ' �^�C�v
    MEAS1 As Double                 ' ����l�P
    MEAS2 As Double                 ' ����l�Q
    MEAS3 As Double                 ' ����l�R
    MEAS4 As Double                 ' ����l�S
    MEAS5 As Double                 ' ����l�T
    JMEAS1 As Double                ' ����l�P
    JMEAS2 As Double                ' ����l�Q
    JMEAS3 As Double                ' ����l�R
    JMEAS4 As Double                ' ����l�S
    JMEAS5 As Double                ' ����l�T
    EFEHS As Double                 ' �����ΐ�
    RRG As Double                   ' �q�q�f
    JudgData As Double              ' �����Ώےl
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    SUIFLG  As String               ' ����FLG
    LTDATA As String                ' �����ʂ�LT�����l
    KANSANCHI As String             ' 10�����Z�l
End Type


' �n������
Public Type typ_TBCMJ003
    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
    TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    OIMEAS1 As Double               ' �n������l�P
    OIMEAS2 As Double               ' �n������l�Q
    OIMEAS3 As Double               ' �n������l�R
    OIMEAS4 As Double               ' �n������l�S
    OIMEAS5 As Double               ' �n������l�T
    ORGRES As Double                ' �n�q�f����
    SETDTM As Date                  ' �ݒ����
    EFFECTTM As Integer             ' �L������
    FTIRMETH As String              ' �e�s�h�q���֎�
    YCOEF As Double                 ' �e�s�h�q���Z���i�x�ؕЁj
    XCOEF As Double                 ' �e�s�h�q���Z���i�w�W���j
    AVE As Double                   ' �`�u�d
    SIGMA As Double                 ' �Ёi�V�O�}�j
    FTIRCONV As Double              ' �e�s�h�q���Z
    INSPECTWAY As String * 2        ' �������@
    JudgData As Double              ' �����Ώےl
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' Cs����
Public Type typ_TBCMJ004
    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
    TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    CSMEAS As Double                ' Cs�����l
    PRE70P As Double                ' �V�O������l
    INSPECTWAY As String * 2        ' �������@
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �n�r�e����
Public Type typ_TBCMJ005
    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
    TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    MEASMETH As String * 1          ' ������@
    MEASSPOT As Integer             ' ����_
    MAG As String * 4               ' �{��
    HTPRC As String * 2             ' �M�������@
    KKSP As String * 3              ' �������ב���ʒu
    KKSET As String * 3             ' �������ב�������{�I��ET��@�@char(1)�{number(2)
    CALCMAX As Double               ' �v�Z���� Max
    CALCAVE As Double               ' �v�Z���� Ave
    MEAS1 As Integer                ' ����l�P
    MEAS2 As Integer                ' ����l�Q
    MEAS3 As Integer                ' ����l�R
    MEAS4 As Integer                ' ����l�S
    MEAS5 As Integer                ' ����l�T
    MEAS6 As Integer                ' ����l�U
    MEAS7 As Integer                ' ����l�V
    MEAS8 As Integer                ' ����l�W
    MEAS9 As Integer                ' ����l�X
    MEAS10 As Integer               ' ����l�P�O
    MEAS11 As Integer               ' ����l�P�P
    MEAS12 As Integer               ' ����l�P�Q
    MEAS13 As Integer               ' ����l�P�R
    MEAS14 As Integer               ' ����l�P�S
    MEAS15 As Integer               ' ����l�P�T
    MEAS16 As Integer               ' ����l�P�U
    MEAS17 As Integer               ' ����l�P�V
    MEAS18 As Integer               ' ����l�P�W
    MEAS19 As Integer               ' ����l�P�X
    MEAS20 As Integer               ' ����l�Q�O
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
    OSFPOS1 As Double               ' ����݋敪�P�ʒu
    OSFWID1 As Double               ' ����݋敪�P��
    OSFRD1  As String               ' ����݋敪�PR/D
    OSFPOS2 As Double               ' ����݋敪�Q�ʒu
    OSFWID2 As Double               ' ����݋敪�Q��
    OSFRD2  As String               ' ����݋敪�QR/D
    OSFPOS3 As Double               ' ����݋敪�R�ʒu
    OSFWID3 As Double               ' ����݋敪�R��
    OSFRD3  As String               ' ����݋敪�RR/D
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
'�X�|�b�gMAX�ǉ� K.Goto 2006/03/31
    SPOTMAX As Long              ' �X�|�b�gMAX
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    CALCMH  As Double               ' �ʓ���(MAX/MIN)
    PTNJUDGRES  As String * 1       ' �p�^�[�����茋��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
End Type


' �f�c����
Public Type typ_TBCMJ006
    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
    TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    MSRSDEN As Integer              ' ���茋�� Den
    MSRSLDL As Integer              ' ���茋�� L/DL
    MSRSDVD2 As Integer             ' ���茋�� DVD2
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
    MS01DVD2 As Integer             ' ����l01 DVD   2002/7/02 tuku
    MS02DVD2 As Integer             ' ����l02 DVD
    MS03DVD2 As Integer             ' ����l03 DVD
    MS04DVD2 As Integer             ' ����l04 DVD
    MS05DVD2 As Integer             ' ����l05 DVD
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    MSZEROMN    As Integer          ' L/DL0�A�����ŏ��l
    MSZEROMX    As Integer          ' L/DL0�A�����ő�l
    PTNJUDGRES  As String * 1       ' �p�^�[�����茋��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
End Type


' ���C�t�^�C��
Public Type typ_TBCMJ007
    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
    TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    MEAS1 As Integer                ' ����l�P
    MEAS2 As Integer                ' ����l�Q
    MEAS3 As Integer                ' ����l�R
    MEAS4 As Integer                ' ����l�S
    MEAS5 As Integer                ' ����l�T
    MEASPEAK As Integer             ' ����l �s�[�N�l
    CALCMEAS As Integer             ' �v�Z����
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
''add 2005/11/11 ����->
''->�X�v���b�h�ɋ󕶎����\�����邽��
    MEAS6 As Integer                ' ����l�U
    MEAS7 As Integer                ' ����l�V
    MEAS8 As Integer                ' ����l�W
    MEAS9 As Integer                ' ����l�X
    MEAS10 As Integer               ' ����l�P�O
    MEASFILE As String              ' ����f�[�^�t�@�C����
    RESVAL As String                ' ������R
    INCVAL As String                ' �X��
    CUTVAL As String                ' �ؕ�
    SETVAL As String                ' �ݒ�l
    CONVAL As String                ' 10�����Z�l
    MEAS1DAT1 As String             ' ����l�P�@���f�[�^�P
    MEAS1DAT2 As String             ' ����l�P�@���f�[�^�Q
    MEAS1DAT3 As String             ' ����l�P�@���f�[�^�R
    MEAS1DAT4 As String             ' ����l�P�@���f�[�^�S
    MEAS1DAT5 As String             ' ����l�P�@���f�[�^�T
    MEAS2DAT1 As String             ' ����l�Q�@���f�[�^�P
    MEAS2DAT2 As String             ' ����l�Q�@���f�[�^�Q
    MEAS2DAT3 As String             ' ����l�Q�@���f�[�^�R
    MEAS2DAT4 As String             ' ����l�Q�@���f�[�^�S
    MEAS2DAT5 As String             ' ����l�Q�@���f�[�^�T
    MEAS3DAT1 As String             ' ����l�R�@���f�[�^�P
    MEAS3DAT2 As String             ' ����l�R�@���f�[�^�Q
    MEAS3DAT3 As String             ' ����l�R�@���f�[�^�R
    MEAS3DAT4 As String             ' ����l�R�@���f�[�^�S
    MEAS3DAT5 As String             ' ����l�R�@���f�[�^�T
    MEAS4DAT1 As String             ' ����l�S�@���f�[�^�P
    MEAS4DAT2 As String             ' ����l�S�@���f�[�^�Q
    MEAS4DAT3 As String             ' ����l�S�@���f�[�^�R
    MEAS4DAT4 As String             ' ����l�S�@���f�[�^�S
    MEAS4DAT5 As String             ' ����l�S�@���f�[�^�T
    MEAS5DAT1 As String             ' ����l�T�@���f�[�^�P
    MEAS5DAT2 As String             ' ����l�T�@���f�[�^�Q
    MEAS5DAT3 As String             ' ����l�T�@���f�[�^�R
    MEAS5DAT4 As String             ' ����l�T�@���f�[�^�S
    MEAS5DAT5 As String             ' ����l�T�@���f�[�^�T
    MEAS6DAT1 As String             ' ����l�U�@���f�[�^�P
    MEAS6DAT2 As String             ' ����l�U�@���f�[�^�Q
    MEAS6DAT3 As String             ' ����l�U�@���f�[�^�R
    MEAS6DAT4 As String             ' ����l�U�@���f�[�^�S
    MEAS6DAT5 As String             ' ����l�U�@���f�[�^�T
    MEAS7DAT1 As String             ' ����l�V�@���f�[�^�P
    MEAS7DAT2 As String             ' ����l�V�@���f�[�^�Q
    MEAS7DAT3 As String             ' ����l�V�@���f�[�^�R
    MEAS7DAT4 As String             ' ����l�V�@���f�[�^�S
    MEAS7DAT5 As String             ' ����l�V�@���f�[�^�T
    MEAS8DAT1 As String             ' ����l�W�@���f�[�^�P
    MEAS8DAT2 As String             ' ����l�W�@���f�[�^�Q
    MEAS8DAT3 As String             ' ����l�W�@���f�[�^�R
    MEAS8DAT4 As String             ' ����l�W�@���f�[�^�S
    MEAS8DAT5 As String             ' ����l�W�@���f�[�^�T
    MEAS9DAT1 As String             ' ����l�X�@���f�[�^�P
    MEAS9DAT2 As String             ' ����l�X�@���f�[�^�Q
    MEAS9DAT3 As String             ' ����l�X�@���f�[�^�R
    MEAS9DAT4 As String             ' ����l�X�@���f�[�^�S
    MEAS9DAT5 As String             ' ����l�X�@���f�[�^�T
    MEAS10DAT1 As String            ' ����l�P�O�@���f�[�^�P
    MEAS10DAT2 As String            ' ����l�P�O�@���f�[�^�Q
    MEAS10DAT3 As String            ' ����l�P�O�@���f�[�^�R
    MEAS10DAT4 As String            ' ����l�P�O�@���f�[�^�S
    MEAS10DAT5 As String            ' ����l�P�O�@���f�[�^�T
    LTSPIFLG As String              ' ����ʒu����t���O
''add 2005/11/11 ����->
End Type


' �a�l�c����
Public Type typ_TBCMJ008
    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
    TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    MEASMETH As String * 1          ' ������@
    MEASSPOT As Integer             ' ����_
    MAG As String * 4               ' �{��
    HTPRC As String * 2             ' �M�������@
    KKSP As String * 3              ' �������ב���ʒu
    KKSET As String * 3             ' �������ב�������{�I��ET��@�@char(1)�{number(2)
    MEAS1 As Integer                ' ����l�P
    MEAS2 As Integer                ' ����l�Q
    MEAS3 As Integer                ' ����l�R
    MEAS4 As Integer                ' ����l�S
    MEAS5 As Integer                ' ����l�T
    MEASMIN As Double               ' MIN
    MEASMAX As Double               ' MAX
    MEASAVE As Double               ' AVE
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
    BMDMNBUNP As Double             ' �a�l�c�ʓ����z
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
End Type


' �����������
Public Type typ_TBCMJ009
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g���ʒu
    TRANCNT As Integer              ' ������
    Length As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    CODE As String * 1              ' �敪�R�[�h
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �����ŏI����
Public Type typ_TBCMJ010
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g���ʒu
    TRANCNT As Integer              ' ������
    Length As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    PAYCLASS As String * 1          ' �����o���敪
    OUTLENGTH As Integer            ' �o�ג���
    PART1 As Integer                ' ���ʂP
    P1BDLEN As Integer              ' ���ʂP�s�ǒ���
    P1BDCAUS As String * 3          ' ���ʂP�s�Ǘ��R
    PART2 As Integer                ' ���ʂQ
    P2BDLEN As Integer              ' ���ʂQ�s�ǒ���
    P2BDCAUS As String * 3          ' ���ʂQ�s�Ǘ��R
    PART3 As Integer                ' ���ʂR
    P3BDLEN As Integer              ' ���ʂR�s�ǒ���
    P3BDCAUS As String * 3          ' ���ʂR�s�Ǘ��R
    PART4 As Integer                ' ���ʂS
    P4BDLEN As Integer              ' ���ʂS�s�ǒ���
    P4BDCAUS As String * 3          ' ���ʂS�s�Ǘ��R
    PART5 As Integer                ' ���ʂT
    P5BDLEN As Integer              ' ���ʂT�s�ǒ���
    P5BDCAUS As String * 3          ' ���ʂT�s�Ǘ��R
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �v�e���o����
Public Type typ_TBCMJ011
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g���ʒu
    Length As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    BLOCKID As String * 12          ' �u���b�NID
    sBlockId As String * 12         ' �擪�u���b�NID
    BLOCKORDER As Integer           ' �u���b�N����
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �z�[���h�i�����j����
Public Type typ_TBCMJ012
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g���ʒu
    TRANCNT As Integer              ' ������
    Length As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    HLDTRCLS As String * 1          ' �z�[���h�����敪
    HLDCAUSE As String * 3          ' �z�[���h���R
    HLDCMNT As String               ' �z�[���h�R�����g
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    HOLDKT As String * 5            ' ΰ��ލH��  2005/07
End Type


' �]�p����
Public Type typ_TBCMJ013
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g���ʒu
    TRANCNT As Integer              ' ������
    Length As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    DUNWNUM As String * 12          ' �]�p��i��
    DUNWREV As Integer              ' �]�p��i�� ���i�ԍ������ԍ�
    DUNWFACT As String * 1          ' �]�p��i�� �H��
    DUNWOPCD As String * 1          ' �]�p��i�� ���Ə���
    DUOGNUM As String * 12          ' �]�p���i��
    DUOGREV As Integer              ' �]�p���i�� ���i�ԍ������ԍ�
    DUOGFACT As String * 1          ' �]�p���i�� �H��
    DUOGOPCD As String * 1          ' �]�p���i�� ���Ə���
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
End Type


' �����������葪��l
Public Type typ_TBCMJ014
    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    Length As Integer               ' ����
    UBLOCKID As String * 12         ' U�u���b�NID
    DBLOCKID As String * 12         ' D�u���b�NID
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    PRODCOND As String * 10         ' �������
    PGID As String * 8              ' �o�f�|�h�c
    UPLENGTH As Integer             ' ���グ����
    PLUPDATE As Date                ' ������t
    FREELENG As Integer             ' �t���[��
    DIAMETER As Integer             ' ���a
    CHARGE As Long                  ' �`���[�W��
    SEED As String * 4              ' �V�[�h
    SXL_RS_SMPPOS As Integer        ' SXLRS����ّ���ʒu�iSXL������j
    SXLRS_MEAS1 As Double           ' SXLRS_����l�P
    SXLRS_MEAS2 As Double           ' SXLRS_����l�Q
    SXLRS_MEAS3 As Double           ' SXLRS_����l�R
    SXLRS_MEAS4 As Double           ' SXLRS_����l�S
    SXLRS_MEAS5 As Double           ' SXLRS_����l�T
    SXLRS_EFEHS As Double           ' SXLRS_�����ΐ�
    SXLRS_RRG As Double             ' SXLRS_�q�q�f
    SXL_OI_SMPPOS As Integer        ' SXLOI����ّ���ʒu�iSXL������j
    SXLOI_OIMEAS1 As Double         ' SXLOI_�n������l�P
    SXLOI_OIMEAS2 As Double         ' SXLOI_�n������l�Q
    SXLOI_OIMEAS3 As Double         ' SXLOI_�n������l�R
    SXLOI_OIMEAS4 As Double         ' SXLOI_�n������l�S
    SXLOI_OIMEAS5 As Double         ' SXLOI_�n������l�T
    SXLOI_ORGRES As Double          ' SXLOI_�n�q�f����
    SXLOI_INSPECTWAY As String * 2  ' SXLOI_�������@
    SXL_CS_SMPPOS As Integer        ' SXLCS����ّ���ʒu�iSXL������j
    SXLCS_CSMEAS As Double          ' SXLCS_Cs�����l
    SXLCS_70PPRE As Double          ' SXLCS_�V�O������l
    SXLOSF1_SMPPOS As Integer       ' SXLOSF����ّ���ʒu�iSXL�ʒu���j
    SXLOSF1_KKSP As String * 3      ' SXLOSF1�������ב���ʒu
    SXLOSF1_NETU As String * 2      ' SXLOSF1�M�����@
    SXLOSF1_KKSET As String * 3     ' SXLOSF1�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF1_MEAS1 As Integer        ' SXLOSF1����_�P
    SXLOSF1_MEAS2 As Integer        ' SXLOSF1����_2
    SXLOSF1_MEAS3 As Integer        ' SXLOSF1����_3
    SXLOSF1_MEAS4 As Integer        ' SXLOSF1����_4
    SXLOSF1_MEAS5 As Integer        ' SXLOSF1����_5
    SXLOSF1_MEAS6 As Integer        ' SXLOSF1����_6
    SXLOSF1_MEAS7 As Integer        ' SXLOSF1����_7
    SXLOSF1_MEAS8 As Integer        ' SXLOSF1����_8
    SXLOSF1_MEAS9 As Integer        ' SXLOSF1����_9
    SXLOSF1_MEAS10 As Integer       ' SXLOSF1����_10
    SXLOSF1_MEAS11 As Integer       ' SXLOSF1����_11
    SXLOSF1_MEAS12 As Integer       ' SXLOSF1����_12
    SXLOSF1_MEAS13 As Integer       ' SXLOSF1����_13
    SXLOSF1_MEAS14 As Integer       ' SXLOSF1����_14
    SXLOSF1_MEAS15 As Integer       ' SXLOSF1����_15
    SXLOSF1_MEAS16 As Integer       ' SXLOSF1����_16
    SXLOSF1_MEAS17 As Integer       ' SXLOSF1����_17
    SXLOSF1_MEAS18 As Integer       ' SXLOSF1����_18
    SXLOSF1_MEAS19 As Integer       ' SXLOSF1����_19
    SXLOSF1_MEAS20 As Integer       ' SXLOSF1����_20
    SXLOSF1_CALCMAX As Double       ' OSF1SXL�v�Z���� Max_1
    SXLOSF1_CALCAVE As Double       ' OSF1SXL�v�Z���� Ave_1
    SXLOSF2_KKSP As String * 3      ' SXLOSF�Q�������ב���ʒu
    SXLOSF2_NETU As String * 2      ' SXLOSF�Q�M�����@
    SXLOSF2_KKSET As String * 3     ' SXLOSF�Q�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF2_MEAS1 As Integer        ' SXLOSF2����_�P
    SXLOSF2_MEAS2 As Integer        ' SXLOSF2����_2
    SXLOSF2_MEAS3 As Integer        ' SXLOSF2����_3
    SXLOSF2_MEAS4 As Integer        ' SXLOSF2����_4
    SXLOSF2_MEAS5 As Integer        ' SXLOSF2����_5
    SXLOSF2_MEAS6 As Integer        ' SXLOSF2����_6
    SXLOSF2_MEAS7 As Integer        ' SXLOSF2����_7
    SXLOSF2_MEAS8 As Integer        ' SXLOSF2����_8
    SXLOSF2_MEAS9 As Integer        ' SXLOSF2����_9
    SXLOSF2_MEAS10 As Integer       ' SXLOSF2����_10
    SXLOSF2_MEAS11 As Integer       ' SXLOSF2����_11
    SXLOSF2_MEAS12 As Integer       ' SXLOSF2����_12
    SXLOSF2_MEAS13 As Integer       ' SXLOSF2����_13
    SXLOSF2_MEAS14 As Integer       ' SXLOSF2����_14
    SXLOSF2_MEAS15 As Integer       ' SXLOSF2����_15
    SXLOSF2_MEAS16 As Integer       ' SXLOSF2����_16
    SXLOSF2_MEAS17 As Integer       ' SXLOSF2����_17
    SXLOSF2_MEAS18 As Integer       ' SXLOSF2����_18
    SXLOSF2_MEAS19 As Integer       ' SXLOSF2����_19
    SXLOSF2_MEAS20 As Integer       ' SXLOSF2����_20
    SXLOSF2_CALCMAX As Double       ' OSF�QSXL�v�Z���� Max_2
    SXLOSF2_CALCAVE As Double       ' OSF�QSXL�v�Z���� Ave_2
    SXLOSF3_KKSP As String * 3      ' SXLOSF�R�������ב���ʒu
    SXLOSF3_NETU As String * 2      ' SXLOSF�R�M�����@
    SXLOSF3_KKSET As String * 3     ' SXLOSF�R�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF3_MEAS1 As Integer        ' SXLOSF3����_�P
    SXLOSF3_MEAS2 As Integer        ' SXLOSF3����_2
    SXLOSF3_MEAS3 As Integer        ' SXLOSF3����_3
    SXLOSF3_MEAS4 As Integer        ' SXLOSF3����_4
    SXLOSF3_MEAS5 As Integer        ' SXLOSF3����_5
    SXLOSF3_MEAS6 As Integer        ' SXLOSF3����_6
    SXLOSF3_MEAS7 As Integer        ' SXLOSF3����_7
    SXLOSF3_MEAS8 As Integer        ' SXLOSF3����_8
    SXLOSF3_MEAS9 As Integer        ' SXLOSF3����_9
    SXLOSF3_MEAS10 As Integer       ' SXLOSF3����_10
    SXLOSF3_MEAS11 As Integer       ' SXLOSF3����_11
    SXLOSF3_MEAS12 As Integer       ' SXLOSF3����_12
    SXLOSF3_MEAS13 As Integer       ' SXLOSF3����_13
    SXLOSF3_MEAS14 As Integer       ' SXLOSF3����_14
    SXLOSF3_MEAS15 As Integer       ' SXLOSF3����_15
    SXLOSF3_MEAS16 As Integer       ' SXLOSF3����_16
    SXLOSF3_MEAS17 As Integer       ' SXLOSF3����_17
    SXLOSF3_MEAS18 As Integer       ' SXLOSF3����_18
    SXLOSF3_MEAS19 As Integer       ' SXLOSF3����_19
    SXLOSF3_MEAS20 As Integer       ' SXLOSF3����_20
    SXLOSF3_CALCMAX As Double       ' OSF�RSXL�v�Z���� Max_3
    SXLOSF3_CALCAVE As Double       ' OSF�RSXL�v�Z���� Ave_3
    SXLOSF4_KKSP As String * 3      ' SXLOSF�S�������ב���ʒu
    SXLOSF4_NETU As String * 2      ' SXLOSF�S�M�����@
    SXLOSF4_KKSET As String * 3     ' SXLOSF�S�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLOSF4_MEAS1 As Integer        ' SXLOSF4����_�P
    SXLOSF4_MEAS2 As Integer        ' SXLOSF4����_2
    SXLOSF4_MEAS3 As Integer        ' SXLOSF4����_3
    SXLOSF4_MEAS4 As Integer        ' SXLOSF4����_4
    SXLOSF4_MEAS5 As Integer        ' SXLOSF4����_5
    SXLOSF4_MEAS6 As Integer        ' SXLOSF4����_6
    SXLOSF4_MEAS7 As Integer        ' SXLOSF4����_7
    SXLOSF4_MEAS8 As Integer        ' SXLOSF4����_8
    SXLOSF4_MEAS9 As Integer        ' SXLOSF4����_9
    SXLOSF4_MEAS10 As Integer       ' SXLOSF4����_10
    SXLOSF4_MEAS11 As Integer       ' SXLOSF4����_11
    SXLOSF4_MEAS12 As Integer       ' SXLOSF4����_12
    SXLOSF4_MEAS13 As Integer       ' SXLOSF4����_13
    SXLOSF4_MEAS14 As Integer       ' SXLOSF4����_14
    SXLOSF4_MEAS15 As Integer       ' SXLOSF4����_15
    SXLOSF4_MEAS16 As Integer       ' SXLOSF4����_16
    SXLOSF4_MEAS17 As Integer       ' SXLOSF4����_17
    SXLOSF4_MEAS18 As Integer       ' SXLOSF4����_18
    SXLOSF4_MEAS19 As Integer       ' SXLOSF4����_19
    SXLOSF4_MEAS20 As Integer       ' SXLOSF4����_20
    SXLOSF4_CALCMAX As Double       ' OSF�SSXL�v�Z���� Max_4
    SXLOSF4_CALCAVE As Double       ' OSF�SSXL�v�Z���� Ave_4
    SXLBMD_SMPPOS As Integer        ' SXLBMD����ّ���ʒu�iSXL�ʒu���j
    SXLBMD1_KKSP As String * 3      ' SXLBMD1�������ב���ʒu
    SXLBMD1_NETU As String * 2      ' SXLBMD1�M�����@
    SXLBMD1_KKSET As String * 3     ' SXLBMD1�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLBMD1_MEAS1 As Integer        ' SXLBMD1����_�P
    SXLBMD1_MEAS2 As Integer        ' SXLBMD1����_2
    SXLBMD1_MEAS3 As Integer        ' SXLBMD1����_3
    SXLBMD1_MEAS4 As Integer        ' SXLBMD1����_4
    SXLBMD1_MEAS5 As Integer        ' SXLBMD1����_5
    SXLBMD1_CALCMAX As Double       ' BMD1SXL�v�Z���� Max
    SXLBMD1_CALCAVE As Double       ' BMD1SXL�v�Z���� Ave
    SXLBMD2_KKSP As String * 3      ' SXLBMD�Q�������ב���ʒu
    SXLBMD2_NETU As String * 2      ' SXLBMD�Q�M�����@
    SXLBMD2_KKSET As String * 3     ' SXLBMD�Q�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLBMD2_MEAS1 As Integer        ' SXLBMD2����_�P
    SXLBMD2_MEAS2 As Integer        ' SXLBMD2����_2
    SXLBMD2_MEAS3 As Integer        ' SXLBMD2����_3
    SXLBMD2_MEAS4 As Integer        ' SXLBMD2����_4
    SXLBMD2_MEAS5 As Integer        ' SXLBMD2����_5
    SXLBMD2_CALCMAX As Double       ' BMD�QSXL�v�Z���� Max
    SXLBMD2_CALCAVE As Double       ' BMD�QSXL�v�Z���� Ave
    SXLBMD3_KKSP As String * 3      ' SXLBMD�R�������ב���ʒu
    SXLBMD3_NETU As String * 2      ' SXLBMD�R�M�����@
    SXLBMD3_KKSET As String * 3     ' SXLBMD�R�������ב�������{�I��ET��@�@char(1)�{number(2)
    SXLBMD3_MEAS1 As Integer        ' SXLBMD3����_�P
    SXLBMD3_MEAS2 As Integer        ' SXLBMD3����_2
    SXLBMD3_MEAS3 As Integer        ' SXLBMD3����_3
    SXLBMD3_MEAS4 As Integer        ' SXLBMD3����_4
    SXLBMD3_MEAS5 As Integer        ' SXLBMD3����_5
    SXLBMD3_CALCMAX As Double       ' BMD�RSXL�v�Z���� Max
    SXLBMD3_CALCAVE As Double       ' BMD�RSXL�v�Z���� Ave
    SXLGD_SMPPOS As Integer         ' SXLGD����ّ���ʒu�iSXL�ʒu���j
    SXLGD_MS01LDL1 As Integer       ' SXLGD_����l01 L/DL1
    SXLGD_MS01LDL2 As Integer       ' SXLGD_����l01 L/DL2
    SXLGD_MS01LDL3 As Integer       ' SXLGD_����l01 L/DL3
    SXLGD_MS01LDL4 As Integer       ' SXLGD_����l01 L/DL4
    SXLGD_MS01LDL5 As Integer       ' SXLGD_����l01 L/DL5
    SXLGD_MS01DEN1 As Integer       ' SXLGD_����l01 Den1
    SXLGD_MS01DEN2 As Integer       ' SXLGD_����l01 Den2
    SXLGD_MS01DEN3 As Integer       ' SXLGD_����l01 Den3
    SXLGD_MS01DEN4 As Integer       ' SXLGD_����l01 Den4
    SXLGD_MS01DEN5 As Integer       ' SXLGD_����l01 Den5
    SXLGD_MS02LDL1 As Integer       ' SXLGD_����l02 L/DL1
    SXLGD_MS02LDL2 As Integer       ' SXLGD_����l02 L/DL2
    SXLGD_MS02LDL3 As Integer       ' SXLGD_����l02 L/DL3
    SXLGD_MS02LDL4 As Integer       ' SXLGD_����l02 L/DL4
    SXLGD_MS02LDL5 As Integer       ' SXLGD_����l02 L/DL5
    SXLGD_MS02DEN1 As Integer       ' SXLGD_����l02 Den1
    SXLGD_MS02DEN2 As Integer       ' SXLGD_����l02 Den2
    SXLGD_MS02DEN3 As Integer       ' SXLGD_����l02 Den3
    SXLGD_MS02DEN4 As Integer       ' SXLGD_����l02 Den4
    SXLGD_MS02DEN5 As Integer       ' SXLGD_����l02 Den5
    SXLGD_MS03LDL1 As Integer       ' SXLGD_����l03 L/DL1
    SXLGD_MS03LDL2 As Integer       ' SXLGD_����l03 L/DL2
    SXLGD_MS03LDL3 As Integer       ' SXLGD_����l03 L/DL3
    SXLGD_MS03LDL4 As Integer       ' SXLGD_����l03 L/DL4
    SXLGD_MS03LDL5 As Integer       ' SXLGD_����l03 L/DL5
    SXLGD_MS03DEN1 As Integer       ' SXLGD_����l03 Den1
    SXLGD_MS03DEN2 As Integer       ' SXLGD_����l03 Den2
    SXLGD_MS03DEN3 As Integer       ' SXLGD_����l03 Den3
    SXLGD_MS03DEN4 As Integer       ' SXLGD_����l03 Den4
    SXLGD_MS03DEN5 As Integer       ' SXLGD_����l03 Den5
    SXLGD_MS04LDL1 As Integer       ' SXLGD_����l04 L/DL1
    SXLGD_MS04LDL2 As Integer       ' SXLGD_����l04 L/DL2
    SXLGD_MS04LDL3 As Integer       ' SXLGD_����l04 L/DL3
    SXLGD_MS04LDL4 As Integer       ' SXLGD_����l04 L/DL4
    SXLGD_MS04LDL5 As Integer       ' SXLGD_����l04 L/DL5
    SXLGD_MS04DEN1 As Integer       ' SXLGD_����l04 Den1
    SXLGD_MS04DEN2 As Integer       ' SXLGD_����l04 Den2
    SXLGD_MS04DEN3 As Integer       ' SXLGD_����l04 Den3
    SXLGD_MS04DEN4 As Integer       ' SXLGD_����l04 Den4
    SXLGD_MS04DEN5 As Integer       ' SXLGD_����l04 Den5
    SXLGD_MS05LDL1 As Integer       ' SXLGD_����l05 L/DL1
    SXLGD_MS05LDL2 As Integer       ' SXLGD_����l05 L/DL2
    SXLGD_MS05LDL3 As Integer       ' SXLGD_����l05 L/DL3
    SXLGD_MS05LDL4 As Integer       ' SXLGD_����l05 L/DL4
    SXLGD_MS05LDL5 As Integer       ' SXLGD_����l05 L/DL5
    SXLGD_MS05DEN1 As Integer       ' SXLGD_����l05 Den1
    SXLGD_MS05DEN2 As Integer       ' SXLGD_����l05 Den2
    SXLGD_MS05DEN3 As Integer       ' SXLGD_����l05 Den3
    SXLGD_MS05DEN4 As Integer       ' SXLGD_����l05 Den4
    SXLGD_MS05DEN5 As Integer       ' SXLGD_����l05 Den5
    SXLGD_MS06LDL1 As Integer       ' SXLGD_����l06 L/DL1
    SXLGD_MS06LDL2 As Integer       ' SXLGD_����l06 L/DL2
    SXLGD_MS06LDL3 As Integer       ' SXLGD_����l06 L/DL3
    SXLGD_MS06LDL4 As Integer       ' SXLGD_����l06 L/DL4
    SXLGD_MS06LDL5 As Integer       ' SXLGD_����l06 L/DL5
    SXLGD_MS06DEN1 As Integer       ' SXLGD_����l06 Den1
    SXLGD_MS06DEN2 As Integer       ' SXLGD_����l06 Den2
    SXLGD_MS06DEN3 As Integer       ' SXLGD_����l06 Den3
    SXLGD_MS06DEN4 As Integer       ' SXLGD_����l06 Den4
    SXLGD_MS06DEN5 As Integer       ' SXLGD_����l06 Den5
    SXLGD_MS07LDL1 As Integer       ' SXLGD_����l07 L/DL1
    SXLGD_MS07LDL2 As Integer       ' SXLGD_����l07 L/DL2
    SXLGD_MS07LDL3 As Integer       ' SXLGD_����l07 L/DL3
    SXLGD_MS07LDL4 As Integer       ' SXLGD_����l07 L/DL4
    SXLGD_MS07LDL5 As Integer       ' SXLGD_����l07 L/DL5
    SXLGD_MS07DEN1 As Integer       ' SXLGD_����l07 Den1
    SXLGD_MS07DEN2 As Integer       ' SXLGD_����l07 Den2
    SXLGD_MS07DEN3 As Integer       ' SXLGD_����l07 Den3
    SXLGD_MS07DEN4 As Integer       ' SXLGD_����l07 Den4
    SXLGD_MS07DEN5 As Integer       ' SXLGD_����l07 Den5
    SXLGD_MS08LDL1 As Integer       ' SXLGD_����l08 L/DL1
    SXLGD_MS08LDL2 As Integer       ' SXLGD_����l08 L/DL2
    SXLGD_MS08LDL3 As Integer       ' SXLGD_����l08 L/DL3
    SXLGD_MS08LDL4 As Integer       ' SXLGD_����l08 L/DL4
    SXLGD_MS08LDL5 As Integer       ' SXLGD_����l08 L/DL5
    SXLGD_MS08DEN1 As Integer       ' SXLGD_����l08 Den1
    SXLGD_MS08DEN2 As Integer       ' SXLGD_����l08 Den2
    SXLGD_MS08DEN3 As Integer       ' SXLGD_����l08 Den3
    SXLGD_MS08DEN4 As Integer       ' SXLGD_����l08 Den4
    SXLGD_MS08DEN5 As Integer       ' SXLGD_����l08 Den5
    SXLGD_MS09LDL1 As Integer       ' SXLGD_����l09 L/DL1
    SXLGD_MS09LDL2 As Integer       ' SXLGD_����l09 L/DL2
    SXLGD_MS09LDL3 As Integer       ' SXLGD_����l09 L/DL3
    SXLGD_MS09LDL4 As Integer       ' SXLGD_����l09 L/DL4
    SXLGD_MS09LDL5 As Integer       ' SXLGD_����l09 L/DL5
    SXLGD_MS09DEN1 As Integer       ' SXLGD_����l09 Den1
    SXLGD_MS09DEN2 As Integer       ' SXLGD_����l09 Den2
    SXLGD_MS09DEN3 As Integer       ' SXLGD_����l09 Den3
    SXLGD_MS09DEN4 As Integer       ' SXLGD_����l09 Den4
    SXLGD_MS09DEN5 As Integer       ' SXLGD_����l09 Den5
    SXLGD_MS10LDL1 As Integer       ' SXLGD_����l10 L/DL1
    SXLGD_MS10LDL2 As Integer       ' SXLGD_����l10 L/DL2
    SXLGD_MS10LDL3 As Integer       ' SXLGD_����l10 L/DL3
    SXLGD_MS10LDL4 As Integer       ' SXLGD_����l10 L/DL4
    SXLGD_MS10LDL5 As Integer       ' SXLGD_����l10 L/DL5
    SXLGD_MS10DEN1 As Integer       ' SXLGD_����l10 Den1
    SXLGD_MS10DEN2 As Integer       ' SXLGD_����l10 Den2
    SXLGD_MS10DEN3 As Integer       ' SXLGD_����l10 Den3
    SXLGD_MS10DEN4 As Integer       ' SXLGD_����l10 Den4
    SXLGD_MS10DEN5 As Integer       ' SXLGD_����l10 Den5
    SXLGD_MS11LDL1 As Integer       ' SXLGD_����l11 L/DL1
    SXLGD_MS11LDL2 As Integer       ' SXLGD_����l11 L/DL2
    SXLGD_MS11LDL3 As Integer       ' SXLGD_����l11 L/DL3
    SXLGD_MS11LDL4 As Integer       ' SXLGD_����l11 L/DL4
    SXLGD_MS11LDL5 As Integer       ' SXLGD_����l11 L/DL5
    SXLGD_MS11DEN1 As Integer       ' SXLGD_����l11 Den1
    SXLGD_MS11DEN2 As Integer       ' SXLGD_����l11 Den2
    SXLGD_MS11DEN3 As Integer       ' SXLGD_����l11 Den3
    SXLGD_MS11DEN4 As Integer       ' SXLGD_����l11 Den4
    SXLGD_MS11DEN5 As Integer       ' SXLGD_����l11 Den5
    SXLGD_MS12LDL1 As Integer       ' SXLGD_����l12 L/DL1
    SXLGD_MS12LDL2 As Integer       ' SXLGD_����l12 L/DL2
    SXLGD_MS12LDL3 As Integer       ' SXLGD_����l12 L/DL3
    SXLGD_MS12LDL4 As Integer       ' SXLGD_����l12 L/DL4
    SXLGD_MS12LDL5 As Integer       ' SXLGD_����l12 L/DL5
    SXLGD_MS12DEN1 As Integer       ' SXLGD_����l12 Den1
    SXLGD_MS12DEN2 As Integer       ' SXLGD_����l12 Den2
    SXLGD_MS12DEN3 As Integer       ' SXLGD_����l12 Den3
    SXLGD_MS12DEN4 As Integer       ' SXLGD_����l12 Den4
    SXLGD_MS12DEN5 As Integer       ' SXLGD_����l12 Den5
    SXLGD_MS13LDL1 As Integer       ' SXLGD_����l13 L/DL1
    SXLGD_MS13LDL2 As Integer       ' SXLGD_����l13 L/DL2
    SXLGD_MS13LDL3 As Integer       ' SXLGD_����l13 L/DL3
    SXLGD_MS13LDL4 As Integer       ' SXLGD_����l13 L/DL4
    SXLGD_MS13LDL5 As Integer       ' SXLGD_����l13 L/DL5
    SXLGD_MS13DEN1 As Integer       ' SXLGD_����l13 Den1
    SXLGD_MS13DEN2 As Integer       ' SXLGD_����l13 Den2
    SXLGD_MS13DEN3 As Integer       ' SXLGD_����l13 Den3
    SXLGD_MS13DEN4 As Integer       ' SXLGD_����l13 Den4
    SXLGD_MS13DEN5 As Integer       ' SXLGD_����l13 Den5
    SXLGD_MS14LDL1 As Integer       ' SXLGD_����l14 L/DL1
    SXLGD_MS14LDL2 As Integer       ' SXLGD_����l14 L/DL2
    SXLGD_MS14LDL3 As Integer       ' SXLGD_����l14 L/DL3
    SXLGD_MS14LDL4 As Integer       ' SXLGD_����l14 L/DL4
    SXLGD_MS14LDL5 As Integer       ' SXLGD_����l14 L/DL5
    SXLGD_MS14DEN1 As Integer       ' SXLGD_����l14 Den1
    SXLGD_MS14DEN2 As Integer       ' SXLGD_����l14 Den2
    SXLGD_MS14DEN3 As Integer       ' SXLGD_����l14 Den3
    SXLGD_MS14DEN4 As Integer       ' SXLGD_����l14 Den4
    SXLGD_MS14DEN5 As Integer       ' SXLGD_����l14 Den5
    SXLGD_MS15LDL1 As Integer       ' SXLGD_����l15 L/DL1
    SXLGD_MS15LDL2 As Integer       ' SXLGD_����l15 L/DL2
    SXLGD_MS15LDL3 As Integer       ' SXLGD_����l15 L/DL3
    SXLGD_MS15LDL4 As Integer       ' SXLGD_����l15 L/DL4
    SXLGD_MS15LDL5 As Integer       ' SXLGD_����l15 L/DL5
    SXLGD_MS15DEN1 As Integer       ' SXLGD_����l15 Den1
    SXLGD_MS15DEN2 As Integer       ' SXLGD_����l15 Den2
    SXLGD_MS15DEN3 As Integer       ' SXLGD_����l15 Den3
    SXLGD_MS15DEN4 As Integer       ' SXLGD_����l15 Den4
    SXLGD_MS15DEN5 As Integer       ' SXLGD_����l15 Den5
    SXLGD_MSRSDEN As Integer        ' SXLGD_���茋�� Den
    SXLGD_MSRSLDL As Integer        ' SXLGD_���茋�� L/DL
    SXLGD_MSRSDVD2 As Integer       ' SXLGD_���茋�� DVD2
    SXLT_SMPPOS As Integer          ' SXLLT����ّ���ʒu�iSXL�ʒu���j
    SXLLT_MEASPEAK As Integer       ' SXLLT_����l �s�[�N�l
    SXLLT_MEAS1 As Integer          ' SXLLT_����l1
    SXLLT_MEAS2 As Integer          ' SXLLT_����l2
    SXLLT_MEAS3 As Integer          ' SXLLT_����l3
    SXLLT_MEAS4 As Integer          ' SXLLT_����l4
    SXLLT_MEAS5 As Integer          ' SXLLT_����l5
    SXLLT_CALCMEAS As Integer       ' SXLLT_�v�Z����
    REGDATE As Date                 ' �o�^���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    SXLOSF1_POS1  As Double         'OSF1����݋敪�P�ʒu
    SXLOSF1_WID1  As Double         'OSF1����݋敪�P��
    SXLOSF1_RD1   As String * 1     'OSF1����݋敪�PR/D
    SXLOSF1_POS2  As Double         'OSF1����݋敪�Q�ʒu
    SXLOSF1_WID2  As Double         'OSF1����݋敪�Q��
    SXLOSF1_RD2   As String * 1     'OSF1����݋敪�QR/D
    SXLOSF1_POS3  As Double         'OSF1����݋敪�R�ʒu
    SXLOSF1_WID3  As Double         'OSF1����݋敪�R��
    SXLOSF1_RD3   As String * 1     'OSF1����݋敪�RR/D
    SXLOSF2_POS1  As Double         'OSF2����݋敪�P�ʒu
    SXLOSF2_WID1  As Double         'OSF2����݋敪�P��
    SXLOSF2_RD1   As String * 1     'OSF2����݋敪�PR/D
    SXLOSF2_POS2  As Double         'OSF2����݋敪�Q�ʒu
    SXLOSF2_WID2  As Double         'OSF2����݋敪�Q��
    SXLOSF2_RD2   As String * 1     'OSF2����݋敪�QR/D
    SXLOSF2_POS3  As Double         'OSF2����݋敪�R�ʒu
    SXLOSF2_WID3  As Double         'OSF2����݋敪�R��
    SXLOSF2_RD3   As String * 1     'OSF2����݋敪�RR/D
    SXLOSF3_POS1  As Double         'OSF3����݋敪�P�ʒu
    SXLOSF3_WID1  As Double         'OSF3����݋敪�P��
    SXLOSF3_RD1   As String * 1     'OSF3����݋敪�PR/D
    SXLOSF3_POS2  As Double         'OSF3����݋敪�Q�ʒu
    SXLOSF3_WID2  As Double         'OSF3����݋敪�Q��
    SXLOSF3_RD2   As String * 1     'OSF3����݋敪�QR/D
    SXLOSF3_POS3  As Double         'OSF3����݋敪�R�ʒu
    SXLOSF3_WID3  As Double         'OSF3����݋敪�R��
    SXLOSF3_RD3   As String * 1     'OSF3����݋敪�RR/D
    SXLOSF4_POS1  As Double         'OSF4����݋敪�P�ʒu
    SXLOSF4_WID1  As Double         'OSF4����݋敪�P��
    SXLOSF4_RD1   As String * 1     'OSF4����݋敪�PR/D
    SXLOSF4_POS2  As Double         'OSF4����݋敪�Q�ʒu
    SXLOSF4_WID2  As Double         'OSF4����݋敪�Q��
    SXLOSF4_RD2   As String * 1     'OSF4����݋敪�QR/D
    SXLOSF4_POS3  As Double         'OSF4����݋敪�R�ʒu
    SXLOSF4_WID3  As Double         'OSF4����݋敪�R��
    SXLOSF4_RD3   As String * 1     'OSF4����݋敪�RR/D
    SXLGD_MS01DVD2 As Integer       'DVD2���茋�ʒl�P
    SXLGD_MS02DVD2 As Integer       'DVD2���茋�ʒl�Q
    SXLGD_MS03DVD2 As Integer       'DVD2���茋�ʒl�R
    SXLGD_MS04DVD2 As Integer       'DVD2���茋�ʒl�S
    SXLGD_MS05DVD2 As Integer       'DVD2���茋�ʒl�T
    SXLBMD1_MNBCR As Double         'BMD1SXL�v�Z���ʖʓ����z
    SXLBMD2_MNBCR As Double         'BMD2SXL�v�Z���ʖʓ����z
    SXLBMD3_MNBCR As Double         'BMD3SXL�v�Z���ʖʓ����z
End Type


' GD����(WF)�@05/01/31 ooba
Public Type typ_TBCMJ015
    CRYNUM      As String * 12          ' �����ԍ�
    POSITION    As Integer              ' �ʒu
    SMPKBN      As String * 1           ' �T���v���敪
    TRANCOND    As String * 1           ' ��������
    TRANCNT     As String * 1           ' ������
    HSFLG       As String * 1           ' �ۏ؃t���O
    SMPLNO      As String * 16          ' �T���v���m��
    SMPLUMU     As String * 1           ' �T���v���L��
    hinban      As String * 8           ' �i��
    REVNUM      As Integer              ' ���i�ԍ������ԍ�
    FACTORY     As String * 1           ' �H��
    OPECOND     As String * 1           ' ���Ə���
    SXLID       As String * 13          ' SXLID
    KRPROCCD    As String * 5           ' �Ǘ��H���R�[�h
    PROCCODE    As String * 5           ' �H���R�[�h
    GOUKI       As String * 3           ' ���@
    OSITEM      As String * 4           ' �]������
    MAISU       As Integer              ' �]������
    Spec        As String * 10          ' �K�i�l
    NETSU       As String * 2           ' �M��������
    ET          As String * 3           ' �G�b�`���O����
    MES         As String * 3           ' �v�����@
    DKAN        As String * 10          ' �c�j�A�j�[������
    ETMAE_RYO01 As Double               ' ET�O�d��01
    ETATO_RYO01 As Double               ' ET��d��01
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
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    MSZEROMN    As Integer              ' L/DL0�A�����ŏ��l
    MSZEROMX    As Integer              ' L/DL0�A�����ő�l
    PTNJUDGRES  As String * 1           ' �p�^�[�����茋��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    TSTAFFID    As String * 8           ' �o�^�Ј�ID
    REGDATE     As Date                 ' �o�^���t
    KSTAFFID    As String * 8           ' �X�V�Ј�ID
    UPDDATE     As Date                 ' �X�V���t
    SENDFLAG    As String * 1           ' ���M�t���O
    SENDDATE    As Date                 ' ���M���t
End Type

''Upd start 2005/06/21 (TCS)T.Terauchi  SPV9�_�Ή�  SPV����ð���
Public Type typ_TBCMJ016
    CRYNUM          As String * 12          ' �����ԍ�
    POSITION        As Integer              ' �ʒu
    SMPKBN          As String * 1           ' �T���v���敪
    TRANCOND        As String * 1           ' ��������
    TRANCNT         As Integer              ' ������
    HSFLG           As String * 1           ' �ۏ؃t���O
    SMPLNO          As String * 16          ' �T���v���m��
    SMPLUMU         As String * 1           ' �T���v���L��
    hinban          As String * 8           ' �i��
    REVNUM          As Integer              ' ���i�ԍ������ԍ�
    FACTORY         As String * 1           ' �H��
    OPECOND         As String * 1           ' ���Ə���
    SXLID           As String * 13          ' SXLID
    KRPROCCD        As String * 5           ' �Ǘ��H���R�[�h
    PROCCODE        As String * 5           ' �H���R�[�h
    GOUKI           As String * 3           ' ���@
    OSITEM          As String * 4           ' �]������
    MAISU           As Integer              ' �]������
    Spec            As String * 10          ' �K�i�l
    NETSU           As String * 2           ' �M��������
    ET              As String * 3           ' �G�b�`���O����
    MES             As String * 3           ' �v�����@
    DKAN            As String * 10          ' �c�j�A�j�[������
    SPV_Fe_MAX      As Double               ' SPV_Fe_MAX
    SPV_Fe_AVE      As Double               ' SPV_Fe_AVE
    SPV_Fe_MIN      As Double               ' SPV_Fe_MIN
    ms01_SPV_Fe     As Double               ' ����l01 SPV_Fe
    ms02_SPV_Fe     As Double               ' ����l02 SPV_Fe
    ms03_SPV_Fe     As Double               ' ����l03 SPV_Fe
    ms04_SPV_Fe     As Double               ' ����l04 SPV_Fe
    ms05_SPV_Fe     As Double               ' ����l05 SPV_Fe
    ms06_SPV_Fe     As Double               ' ����l06 SPV_Fe
    ms07_SPV_Fe     As Double               ' ����l07 SPV_Fe
    ms08_SPV_Fe     As Double               ' ����l08 SPV_Fe
    ms09_SPV_Fe     As Double               ' ����l09 SPV_Fe
    SPV_Diff_MAX    As Double               ' SPV_�g�U��_MAX
    SPV_Diff_AVE    As Double               ' SPV_�g�U��_AVE
    SPV_Diff_MIN    As Double               ' SPV_�g�U��_MIN
    ms01_SPV_Diff   As Double               ' ����l01 SPV_�g�U��
    ms02_SPV_Diff   As Double               ' ����l02 SPV_�g�U��
    ms03_SPV_Diff   As Double               ' ����l03 SPV_�g�U��
    ms04_SPV_Diff   As Double               ' ����l04 SPV_�g�U��
    ms05_SPV_Diff   As Double               ' ����l05 SPV_�g�U��
    ms06_SPV_Diff   As Double               ' ����l06 SPV_�g�U��
    ms07_SPV_Diff   As Double               ' ����l07 SPV_�g�U��
    ms08_SPV_Diff   As Double               ' ����l08 SPV_�g�U��
    ms09_SPV_Diff   As Double               ' ����l09 SPV_�g�U��
    TSTAFFID        As String * 8           ' �o�^�Ј�ID
    REGDATE         As Date                 ' �o�^���t
    KSTAFFID        As String * 8           ' �X�V�Ј�ID
    UPDDATE         As Date                 ' �X�V���t
    SENDFLAG        As String * 1           ' ���M�t���O
    SENDDATE        As Date                 ' ���M���t
    MAX_FE          As Double               ' FE�Z�x�@�ő�l(�\���A����p)
    MIN_FE          As Double               ' FE�Z�x�@�ŏ��l(�\���A����p)
    AVE_FE          As Double               ' FE�Z�x�@����(�\���A����p)
    CENTER_FE       As Double               ' FE�Z�x�@���S(�\���A����p)
    MAX_DIFF        As Double               ' �g�U���@�ő�l(�\���A����p)
    MIN_DIFF        As Double               ' �g�U���@�ŏ��l(�\���A����p)
    AVE_DIFF        As Double               ' �g�U���@����(�\���A����p)
    CENTER_DIFF     As Double               ' �g�U���@���S(�\���A����p)
    ''==SPV����@20060529 SMP����
    SPV_Fe_PUA      As Double               'SPV_Fe PUA�l
    SPV_Fe_PUAP     As Double               'SPV_Fe PUA���l
    SPV_Fe_STD      As Double               'SPV_Fe STD
    SPV_Diff_PUA    As Double               'SPV_�g�U�� PUA�l
    SPV_Diff_PUAP   As Double               'SPV_�g�U�� PUA���l
    SPV_Nr_MAX      As Double               'SPV_OtherRecords_MAX
    SPV_Nr_AVE      As Double               'SPV_OtherRecords_AVE
    SPV_Nr_STD      As Double               'SPV_OtherRecords_STD
    SPV_Nr_PUA      As Double               'SPV_OtherRecords_PUA�l
    SPV_Nr_PUAP     As Double               'SPV_OtherRecords_PUA���l
    ''==============================
    ''==SPV����@20060612 SMP)kondoh
    PUA_FE      As Double                   ' FE�Z�x  PUA�l(�\���A����p)
    PUAP_FE     As Double                   ' FE�Z�x  PUA���l(�\���A����p)
    STD_FE      As Double                   ' FE�Z�x  STD(�\���A����p)
    PUA_DIFF    As Double                   ' �g�U��  PUA�l(�\���A����p)
    PUAP_DIFF   As Double                   ' �g�U��  PUA���l(�\���A����p)
    MAX_NR      As Double                   ' NR�Z�x  �ő�l(�\���A����p)
    MIN_NR      As Double                   ' NR�Z�x  �ŏ��l(�\���A����p)
    AVE_NR      As Double                   ' NR�Z�x  ����(�\���A����p)
    CENTER_NR   As Double                   ' NR�Z�x  ���S(�\���A����p)
    PUA_NR      As Double                   ' NR�Z�x  PUA�l(�\���A����p)
    PUAP_NR     As Double                   ' NR�Z�x  PUA���l(�\���A����p)
    STD_NR      As Double                   ' NR�Z�x  STD(�\���A����p)
    ''==============================
End Type
''Upd end   2005/06/21 (TCS)T.Terauchi  SPV9�_�Ή�  SPV����ð���


' X�����с@2009/08 SUMCO Akizuki�ǉ�
Public Type typ_TBCMJ021
    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
    TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��
    SMPLUMU As String * 1           ' �T���v���L��
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    XRAYX As Single                 ' �����ʌX�� ������(X)
    XRAYY As Single                 ' �����ʌX�� �c����(Y)
    XRAYXY As Single                ' �����ʌX�� ����(����)
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type

''��Add 2010/01/12 SIRD�Ή� Y.Hitomi
Public Type typ_TBCMJ022
    CRYNUM          As String * 12          ' �����ԍ�
    POSITION        As Integer              ' �ʒu
    SMPKBN          As String * 1           ' �T���v���敪
    TRANCOND        As String * 1           ' ��������
    TRANCNT         As Integer              ' ������
    HSFLG           As String * 1           ' �ۏ؃t���O
    SMPLNO          As String * 16          ' �T���v���m��
    SMPLUMU         As String * 1           ' �T���v���L��
    BLOCKID         As String * 12          ' �u���b�NID
    SXLID           As String * 13          ' SXLID
    hinban          As String * 8           ' �i��
    REVNUM          As Integer              ' ���i�ԍ������ԍ�
    FACTORY         As String * 1           ' �H��
    OPECOND         As String * 1           ' ���Ə���
    KRPROCCD        As String * 5           ' �Ǘ��H���R�[�h
    PROCCODE        As String * 5           ' �H���R�[�h
    GOUKI           As String * 3           ' ���@
    OSITEM          As String * 4           ' �]������
    MAISU           As Integer              ' �]������
    Spec            As String * 10          ' �K�i�l
    NETSU           As String * 2           ' �M��������
    ET              As String * 3           ' �G�b�`���O����
    MES             As String * 3           ' �v�����@
    DKAN            As String * 10          ' �c�j�A�j�[������
    SIRDCNT         As Integer              ' �ʓ���
    PLANTCAT        As String * 2           ' ���Ə��敪
    OSWAFID         As String * 6           ' OS�E�F�n�[ID
    TSTAFFID        As String * 8           ' �o�^�Ј�ID
    REGDATE         As Date                 ' �o�^���t
    KSTAFFID        As String * 8           ' �X�V�Ј�ID
    UPDDATE         As Date                 ' �X�V���t
    SENDFLAG        As String * 1           ' ���M�t���O
    SENDDATE        As Date                 ' ���M���t
End Type
''��Add 2010/01/12 SIRD�Ή� Y.Hitomi

'Add Start 2010/12/17 SMPK Miyata
Public Type typ_TBCMJ023
    CRYNUM          As String * 12          ' �����ԍ�
    POSITION        As Integer              ' �ʒu
    SMPKBN          As String * 1           ' �T���v���敪
    TRANCOND        As String * 1           ' ��������
    TRANCNT         As Integer              ' ������
    SMPLNO          As Long                 ' �T���v���m��
    SMPLUMUC        As String               ' �T���v���L��(C)
    SMPLUMUCJ       As String               ' �T���v���L��(CJ)
    SMPLUMUCJLT     As String               ' �T���v���L��(CJLT)
    SMPLUMUCJ2      As String               ' �T���v���L��(CJ2)
    hinban          As String * 8           ' �i��
    REVNUM          As Integer              ' ���i�ԍ������ԍ�
    factory         As String * 1           ' �H��
    opecond         As String * 1           ' ���Ə���
    KRPROCCD        As String * 5           ' �Ǘ��H���R�[�h
    PROCCODE        As String * 5           ' �H���R�[�h
    GOUKI           As String * 3           ' ���@
    CPTNJSK         As String * 1           ' C �p�^�[������
    CDISKJSK        As Integer              ' C Disk���a����
    CRINGNKJSK      As Integer              ' C Ring���a����
    CRINGGKJSK      As Integer              ' C Ring�O�a����
    C_SZ            As String * 1           ' C �������
    CHANTEI         As String               ' C ���茋��
    CJPTNJSK        As String               ' CJ �p�^�[������
    CJDISKJSK       As Integer              ' CJ Disk���a����
    CJRINGNKJSK     As Integer              ' CJ Ring���a����
    CJRINGGKJSK     As Integer              ' CJ Ring�O�a����
    CJBANDNKJSK     As Integer              ' CJ Band���a����
    CJBANDGKJSK     As Integer              ' CJ Band�O�a����
    CJRINGCALC      As Integer              ' CJ Ring���v�Z
    CJPICALC        As Integer              ' CJ Pi���v�Z
    CJ_NETU         As String * 2           ' CJ �M�����@
    CJHANTEI        As String               ' CJ ���茋��
    CJBUIUMU        As String               ' CJ ���ʕʔ���L��
    CJDMAXPIC5      As Integer              ' CJ Disk�̂݃p�^�[�� Pi������l
    CJRMAXPIC5      As Integer              ' CJ Ring�̂݃p�^�[�� Pi������l
    CJDRMAXPIC5     As Integer              ' CJ DiskRing�p�^�[�� Pi������l
    CJALLMAXDIC5    As Integer              ' CJ ����Disk���a����l
    CJALLMINRINC5   As Integer              ' CJ ����Ring���a�����l
    CJALLMAXRIGC5   As Integer              ' CJ ����Ring�O�a����l
    CJLTPTNJSK      As String               ' CJ(LT) �p�^�[������
    CJLTDISKJSK     As Integer              ' CJ(LT) Disk���a����
    CJLTRINGNKJSK   As Integer              ' CJ(LT) Ring���a����
    CJLTRINGGKJSK   As Integer              ' CJ(LT) Ring�O�a����
    CJLTBANDNKJSK   As Integer              ' CJ(LT) Band���a����
    CJLTBANDGKJSK   As Integer              ' CJ(LT) Band�O�a����
    CJLTRINGCALC    As Integer              ' CJ(LT) Ring���v�Z
    CJLTPICALC      As Integer              ' CJ(LT) Pi���v�Z
    CJLTBANDCALC    As Integer              ' CJ(LT) Band���v�Z
    HSXCJLTBND      As Integer              ' CJ(LT) Band������l
    CJLT_NETU       As String * 2           ' CJ(LT) �M�����@
    CJLTHANTEI      As String               ' CJ(LT) ���茋��
    CJ2PTNJSK       As String               ' CJ2 �p�^�[������
    CJ2DISKJSK      As Integer              ' CJ2 Disk���a����
    CJ2RINGNKJSK    As Integer              ' CJ2 Ring���a����
    CJ2RINGGKJSK    As Integer              ' CJ2 Ring�O�a����
    CJ2PICALC       As Integer              ' CJ2 Pi���v�Z
    CJ2_NETU        As String * 2           ' CJ2 �M�����@
    CJ2HANTEI       As String               ' CJ2 ���茋��
    CJ2BUIUMU       As String               ' CJ2 ���ʕʔ���L��
    CJ2DMAXPIC5     As Integer              ' CJ2 Disk�̂݃p�^�[�� Pi������l
    CJ2RMAXPIC5     As Integer              ' CJ2 Ring�̂݃p�^�[�� Pi������l
    CJ2RMINRINC5    As Integer              ' CJ2 Ring�̂݃p�^�[�� Ring���a�����l
    CJ2RMAXRIGC5    As Integer              ' CJ2 Ring�̂݃p�^�[�� Ring�O�a����l
    CJ2DRMAXPIC5    As Integer              ' CJ2 DiskRing�p�^�[�� Pi������l
    CJ2DRMINRINC5   As Integer              ' CJ2 DiskRing�p�^�[�� Ring���a�����l
    CJ2DRMAXRIGC5   As Integer              ' CJ2 DiskRing�p�^�[�� Ring�O�a����l
    TSTAFFID        As String               ' �o�^�Ј�ID
    REGDATE         As Date                 ' �o�^���t
    TSTAFFIDC       As String               ' �o�^�Ј�ID (C)
    REGDATEC        As String               ' �o�^���t   (C)
    TSTAFFIDCJ      As String               ' �o�^�Ј�ID (CJ)
    REGDATECJ       As String               ' �o�^���t   (CJ)
    TSTAFFIDCJLT    As String               ' �o�^�Ј�ID (CJLT)
    REGDATECJLT     As String               ' �o�^���t   (CJLT)
    TSTAFFIDCJ2     As String               ' �o�^�Ј�ID (CJ2)
    REGDATECJ2      As String               ' �o�^���t   (CJ2)
    KSTAFFID        As String               ' �X�V�Ј�ID
    UPDDATE         As String               ' �X�V���t
    KSTAFFIDC       As String               ' �X�V�Ј�ID (C)
    UPDDATEC        As String               ' �X�V���t   (C)
    KSTAFFIDCJ      As String               ' �X�V�Ј�ID (CJ)
    UPDDATECJ       As String               ' �X�V���t   (CJ)
    KSTAFFIDCJLT    As String               ' �X�V�Ј�ID (CJLT)
    UPDDATECJLT     As String               ' �X�V���t   (CJLT)
    KSTAFFIDCJ2     As String               ' �X�V�Ј�ID (CJ2)
    UPDDATECJ2      As String               ' �X�V���t   (CJ2)
    SENDFLAG        As String               ' ���M�t���O
    SENDDATE        As Date                 ' ���M���t
End Type
'Add End   2010/12/17 SMPK Miyata

' �����w������
Public Type typ_TBCMW001
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g�ʒu
    TRANCNT As Integer              ' ������
    CRYLEN As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    BLOCKID As String * 12          ' �u���b�NID
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �������ύX
Public Type typ_TBCMW002
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g�ʒu
    TRANCNT As Integer              ' ������
    CRYLEN As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    BLOCKID As String * 12          ' �u���b�NID
    DELFLG As String * 1            ' �폜�敪
    TOPBDLN As Integer              ' TOP�s�ǒ���
    TOPBDCS As String * 3           ' TOP�s�Ǘ��R
    TAILBDLN As Integer             ' TAIL�s�ǒ���
    TAILBDCS As String * 3          ' TAIL�s�Ǘ��R
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SUMMITSENDFLAG As String * 1    ' SUMMIT���M�t���O
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �����ύX�w������
Public Type typ_TBCMW003
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g�ʒu
    TRANCNT As Integer              ' ������
    CRYLEN As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    BLOCKID As String * 12          ' �u���b�NID
    DELFLG As String * 1            ' �폜�敪
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �Ĕ����w������
Public Type typ_TBCMW004
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g�ʒu
    TRANCNT As Integer              ' ������
    CRYLEN As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    BLOCKID As String * 12          ' �u���b�NID
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' WF�����������
Public Type typ_TBCMW005
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g�ʒu
    TRANCNT As Integer              ' ������
    CRYLEN As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    SXLID As String * 13            ' SXLID
    CODE As String * 1              ' �敪�R�[�h
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �U�֔p������
Public Type typ_TBCMW006
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g�ʒu
    TRANCNT As Integer              ' ������
    CRYLEN As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    TRANCLS As String * 1           ' �����敪
    DUNWNUM As String * 8           ' �]�p��i��
    DUNWREV As Integer              ' �]�p��i�� ���i�ԍ������ԍ�
    DUNWFACT As String * 1          ' �]�p��i�� �H��
    DUNWOPCD As String * 1          ' �]�p��i�� ���Ə���
    DUOGNUM As String * 8           ' �]�p���i��
    DUOGREV As Integer              ' �]�p���i�� ���i�ԍ������ԍ�
    DUOGFACT As String * 1          ' �]�p���i�� �H��
    DUOGOPCD As String * 1          ' �]�p���i�� ���Ə���
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    MUKESAKI As String              ' 07/09/05 SPK Tsutsumi Add
End Type


' �V���O���m�����
Public Type typ_TBCMW007
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g�ʒu
    CRYLEN As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    SXLID As String * 13            ' �V���O��ID
    SAMPLE_FROM As String * 16      ' �T���v��ID (From)
    SAMPLE_TO As String * 16        ' �T���v��ID (To)
    BLOCKID As String * 12          ' �u���b�NID
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' WF�z�[���h�i�����j����
Public Type typ_TBCMW008
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �C���S�b�g�ʒu
    TRANCNT As Integer              ' ������
    CRYLEN As Integer               ' ����
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    SNGLID As String * 13           ' �V���O��ID
    HLDCLASS As String * 1          ' �z�[���h�����敪
    HLDCAUSE As String * 3          ' �z�[���h���R
    HLDCMNT As String               ' �z�[���h�R�����g
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' WF�Z���^�[�������葪��l
Public Type typ_TBCMW009
    SXLID As String * 13            ' SXLID
    FROMTOKBN As String * 1         ' FROMTO�敪
    SAMPLE_FROM As String * 16      ' �T���v��ID (From)
    SAMPLE_TO As String * 16        ' �T���v��ID (To)
    BLOCKID As String * 12          ' �u���b�NID
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    WFOI_SMPPOS As Integer          ' WFOI�����-ID����ʒu�iSXL�ʒu���j
    WFOI_NETSU As String * 2        ' WFOI_�M��������
    WFOI_ET As String * 3           ' WFOI_�G�b�`���O����
    WFOI_MES As String * 3          ' WFOI_�v�����@
    WFOI_MESDATA1 As Double         ' WFOI_����f�[�^���̂P
    WFOI_MESDATA2 As Double         ' WFOI_����f�[�^���̂Q
    WFOI_MESDATA3 As Double         ' WFOI_����f�[�^���̂R
    WFOI_MESDATA4 As Double         ' WFOI_����f�[�^���̂S
    WFOI_MESDATA5 As Double         ' WFOI_����f�[�^���̂T
    WFOI_MESDATA6 As Double         ' WFOI_����f�[�^���̂U
    WFOI_MESDATA7 As Double         ' WFOI_����f�[�^���̂V
    WFOI_MESDATA8 As Double         ' WFOI_����f�[�^���̂W
    WFOI_MESDATA9 As Double         ' WFOI_����f�[�^���̂X
    WFOI_MESDATA10 As Double        ' WFOI_����f�[�^���̂P�O
    WFOI_ORG As Double              ' WFOI_ORG�v�Z����
    WFRS_SMPPOS As Integer          ' WFRS�����-ID����ʒu�iSXL�ʒu���j
    WFRS_NETSU As String * 2        ' WFRS_�M��������
    WFRS_ET As String * 3           ' WFRS_�G�b�`���O����
    WFRS_MES As String * 3          ' WFRS_�v�����@
    WFRS_MESDATA1 As Double         ' WFRS_����f�[�^���̂P
    WFRS_MESDATA2 As Double         ' WFRS_����f�[�^���̂Q
    WFRS_MESDATA3 As Double         ' WFRS_����f�[�^���̂R
    WFRS_MESDATA4 As Double         ' WFRS_����f�[�^���̂S
    WFRS_MESDATA5 As Double         ' WFRS_����f�[�^���̂T
    WFRS_RRG As Double              ' WFRS_RRG�v�Z����
    WFDOI_SMPPOS As Integer         ' WFDOI�����-ID����ʒu�iSXL�ʒu���j�@number(4)
    WFDOI_NETU_1 As String * 2      ' WFDOI_�M��������_1
    WFDOI_MES_1 As String * 3       ' WFDOI_�v�����@_1
    WFDOI_MESDATA1_1 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi�j�P_1
    WFDOI_MESDATA2_1 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�Q_1
    WFDOI_MESDATA3_1 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�R_1
    WFDOI_NETU_2 As String * 2      ' WFDOI_�M��������_�Q
    WFDOI_MES_2 As String * 3       ' WFDOI_�v�����@_�Q
    WFDOI_MESDATA1_2 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi�j�P_�Q
    WFDOI_MESDATA2_2 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�Q_�Q
    WFDOI_MESDATA3_2 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�R_�Q
    WFDOI_NETU_3 As String * 2      ' WFDOI_�M��������_�R
    WFDOI_MES_3 As String * 3       ' WFDOI_�v�����@_�R
    WFDOI_MESDATA1_3 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi�j�P_�R
    WFDOI_MESDATA2_3 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�Q_�R
    WFDOI_MESDATA3_3 As Double      ' WFDOI_(�Ƽ��Oi-AfterOi)�R_�R
    WFOSF1_SMPPOS As Integer        ' WFOSF1�����-ID����ʒu�iSXL�ʒu���j
    WFOSF1_NETSU As String * 2      ' WFOSF1_�M��������
    WFOSF1_ET As String * 3         ' WFOSF1_�G�b�`���O����
    WFOSF1_MES As String * 3        ' WFOSF1_�v�����@
    WFOSF1_MAX As Double            ' WFOSF1_���莞��MAX�l_1
    WFOSF1_AVE As Double            ' WFOSF1_���莞��AVE�l_1
    WFOSF2_SMPPOS As Integer        ' WFOSF�Q�����-ID����ʒu�iSXL�ʒu���j�@number(4)
    WFOSF2_NETSU As String * 2      ' WFOSF2_�M��������_�Q
    WFOSF2_ET As String * 3         ' WFOSF2_�G�b�`���O����_�Q
    WFOSF2_MES As String * 3        ' WFOSF2_�v�����@_�Q
    WFOSF2_MAX As Double            ' WFOSF2_���莞��MAX�l_�Q
    WFOSF2_AVE As Double            ' WFOSF2_���莞��AVE�l_�Q
    WFOSF3_SMPPOS As Integer        ' WFOSF�R�����-ID����ʒu�iSXL�ʒu���j
    WFOSF3_NETSU As String * 2      ' WFOSF3_�M��������_�R
    WFOSF3_ET As String * 3         ' WFOSF3_�G�b�`���O����_�R
    WFOSF3_MES As String * 3        ' WFOSF3_�v�����@_�R
    WFOSF3_MAX As Double            ' WFOSF3_���莞��MAX�l_�R
    WFOSF3_AVE As Double            ' WFOSF3_���莞��AVE�l_�R
    WFOSF4_SMPPOS As Integer        ' WFOSF�S�����-ID����ʒu�iSXL�ʒu���j
    WFOSF4_NETSU As String * 2      ' WFOSF4_�M��������_�S
    WFOSF4_ET As String * 3         ' WFOSF4_�G�b�`���O����_�S
    WFOSF4_MES As String * 3        ' WFOSF4_�v�����@_�S
    WFOSF4_MAX As Double            ' WFOSF4_���莞��MAX�l_�S
    WFOSF4_AVE As Double            ' WFOSF4_���莞��AVE�l_�S
    WFBMD1_SMPPOS As Integer        ' WFBMD1�����-ID����ʒu�iSXL�ʒu���j
    WFBMD1_NETSU As String * 2      ' WFBMD1_�M��������_1
    WFBMD1_ET As String * 3         ' WFBMD1_�G�b�`���O����_1
    WFBMD1_MES As String * 3        ' WFBMD1_�v�����@_1
    WFBMD1_MAX As Double            ' WFBMD1_���莞��MAX�l_1
    WFBMD1_AVE As Double            ' WFBMD1_���莞��AVE�l_1
    WFBMD2_SMPPOS As Integer        ' WFBMD�Q�����-ID����ʒu�iSXL�ʒu���j
    WFBMD2_NETSU As String * 2      ' WFBMD2_�M��������_�Q
    WFBMD2_ET As String * 3         ' WFBMD2_�G�b�`���O����_�Q
    WFBMD2_MES As String * 3        ' WFBMD2_�v�����@_�Q
    WFBMD2_MAX As Double            ' WFBMD2_���莞��MAX�l_�Q
    WFBMD2_AVE As Double            ' WFBMD2_���莞��AVE�l_�Q
    WFBMD3_SMPPOS As Integer        ' WFBMD�R�����-ID����ʒu�iSXL�ʒu���j
    WFBMD3_NETSU As String * 2      ' WFBMD3_�M��������_�R
    WFBMD3_ET As String * 3         ' WFBMD3_�G�b�`���O����_�R
    WFBMD3_MES As String * 3        ' WFBMD3_�v�����@_�R
    WFBMD3_MAX As Double            ' WFBMD3_���莞��MAX�l_�R
    WFBMD3_AVE As Double            ' WFBMD3_���莞��AVE�l_�R
    WFDSOD_SMPPOS As Integer        ' WFDSOD�����-ID����ʒu�iSXL�ʒu���j
    WFDSOD_NETSU As String * 2      ' WFDSOD_�M��������
    WFDSOD_ET As String * 3         ' WFDSOD_�G�b�`���O����
    WFDSOD_MES As String * 3        ' WFDSOD_�v�����@
    WFDSOD_TOTAL As Integer         ' WFDSOD_���莞��TOTAL�l
    WFSPV_SMPPOS As Integer         ' WFSPV�����-ID����ʒu�iSXL�ʒu���j
    WFSPV_NETSU As String * 2       ' WFSVP_�M��������
    WFSPV_ET As String * 3          ' WFSPV_�G�b�`���O����
    WFSPV_MES As String * 3         ' WFSPV_�v�����@
    WFSPV_KST_MAX As Double         ' WFSPV_�g�U�����莞��MAX�l
    WFSPV_KST_AVE As Double         ' WFSPV_�g�U�����莞��AVE�l
    WFSPV_KST_MIN As Double         ' WFSPV_�g�U�����莞��MIN�l
    WFSPV_FE_MAX As Double          ' WFSPV_Fe�Z�x���莞��MAX�l
    WFSPV_FE_AVE As Double          ' WFSPV_Fe�Z�x���莞��AVE�l
    WFSPV_FE_MIN As Double          ' WFSPV_Fe�Z�x���莞��MIN�l
    WFDZ_SMPPOS As Integer          ' WFDZ�����-ID����ʒu�iSXL�ʒu���j
    WFDZ_NETSU As String * 2        ' WFDZ_�M��������
    WFDZ_ET As String * 3           ' WFDZ_�G�b�`���O����
    WFDZ_MES As String * 3          ' WFDZ_�v�����@
    WFDZ_MAX As Double              ' WFDZ_���莞��MAX�l_
    WFDZ_AVE As Double              ' WFDZ_���莞��AVE�l
    TSTAFFID As String * 8          ' �o�^�Ј�ID
    REGDATE As Date                 ' �o�^���t
    KSTAFFID As String * 8          ' �X�V�Ј�ID
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
    WFBMD1_MIN As Double            ' WFBMD1_���莞��MIN�l_1�@��2003/05/14 ooba
    WFBMD1_MBNP As Double           ' WFBMD1_���莞�̖ʓ����z
    WFBMD2_MIN As Double            ' WFBMD2_���莞��MIN�l_2
    WFBMD2_MBNP As Double           ' WFBMD2_���莞�̖ʓ����z
    WFBMD3_MIN As Double            ' WFBMD3_���莞��MIN�l_3
    WFBMD3_MBNP As Double           ' WFBMD3_���莞�̖ʓ����z
    WFDZ_MIN As Double              ' WFDZ_���莞��MIN�l
    WFOSF1_PATKBNP1 As Double       ' WF_OSF1_�p�^�[���敪�P�ʒu
    WFOSF1_PATKBNWID1 As Double     ' WF_OSF1_�p�^�[���敪�P��
    WFOSF1_PATKBNRD1 As String * 1  ' WF_OSF1_�p�^�[���敪�PRing/Disk
    WFOSF1_PATKBNP2 As Double       ' WF_OSF1_�p�^�[���敪�Q�ʒu
    WFOSF1_PATKBNWID2 As Double     ' WF_OSF1_�p�^�[���敪�Q��
    WFOSF1_PATKBNRD2 As String * 1  ' WF_OSF1_�p�^�[���敪�QRing/Disk
    WFOSF1_PATKBNP3 As Double       ' WF_OSF1_�p�^�[���敪�R�ʒu
    WFOSF1_PATKBNWID3 As Double     ' WF_OSF1_�p�^�[���敪�R��
    WFOSF1_PATKBNRD3 As String * 1  ' WF_OSF1_�p�^�[���敪�RRing/Disk
    WFOSF2_PATKBNP1 As Double       ' WF_OSF2_�p�^�[���敪�P�ʒu
    WFOSF2_PATKBNWID1 As Double     ' WF_OSF2_�p�^�[���敪�P��
    WFOSF2_PATKBNRD1 As String * 1  ' WF_OSF2_�p�^�[���敪�PRing/Disk
    WFOSF2_PATKBNP2 As Double       ' WF_OSF2_�p�^�[���敪�Q�ʒu
    WFOSF2_PATKBNWID2 As Double     ' WF_OSF2_�p�^�[���敪�Q��
    WFOSF2_PATKBNRD2 As String * 1  ' WF_OSF2_�p�^�[���敪�QRing/Disk
    WFOSF2_PATKBNP3 As Double       ' WF_OSF2_�p�^�[���敪�R�ʒu
    WFOSF2_PATKBNWID3 As Double     ' WF_OSF2_�p�^�[���敪�R��
    WFOSF2_PATKBNRD3 As String * 1  ' WF_OSF2_�p�^�[���敪�RRing/Disk
    WFOSF3_PATKBNP1 As Double       ' WF_OSF3_�p�^�[���敪�P�ʒu
    WFOSF3_PATKBNWID1 As Double     ' WF_OSF3_�p�^�[���敪�P��
    WFOSF3_PATKBNRD1 As String * 1  ' WF_OSF3_�p�^�[���敪�PRing/Disk
    WFOSF3_PATKBNP2 As Double       ' WF_OSF3_�p�^�[���敪�Q�ʒu
    WFOSF3_PATKBNWID2 As Double     ' WF_OSF3_�p�^�[���敪�Q��
    WFOSF3_PATKBNRD2 As String * 1  ' WF_OSF3_�p�^�[���敪�QRing/Disk
    WFOSF3_PATKBNP3 As Double       ' WF_OSF3_�p�^�[���敪�R�ʒu
    WFOSF3_PATKBNWID3 As Double     ' WF_OSF3_�p�^�[���敪�R��
    WFOSF3_PATKBNRD3 As String * 1  ' WF_OSF3_�p�^�[���敪�RRing/Disk
    WFOSF4_PATKBNP1 As Double       ' WF_OSF4_�p�^�[���敪�P�ʒu
    WFOSF4_PATKBNWID1 As Double     ' WF_OSF4_�p�^�[���敪�P��
    WFOSF4_PATKBNRD1 As String * 1  ' WF_OSF4_�p�^�[���敪�PRing/Disk
    WFOSF4_PATKBNP2 As Double       ' WF_OSF4_�p�^�[���敪�Q�ʒu
    WFOSF4_PATKBNWID2 As Double     ' WF_OSF4_�p�^�[���敪�Q��
    WFOSF4_PATKBNRD2 As String * 1  ' WF_OSF4_�p�^�[���敪�QRing/Disk
    WFOSF4_PATKBNP3 As Double       ' WF_OSF4_�p�^�[���敪�R�ʒu
    WFOSF4_PATKBNWID3 As Double     ' WF_OSF4_�p�^�[���敪�R��
    WFOSF4_PATKBNRD3 As String * 1  ' WF_OSF4_�p�^�[���敪�RRing/Disk�@��2003/05/14 ooba
End Type


' �ڋq�d�l�Ǘ�
Public Type typ_TBCME001
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KMGSHN As String * 7            ' �w�Ǘ��Г��i��
    KMGSNRNO As Integer             ' �w�Ǘ��Г��i�������ԍ�
    KMGSTFNO As String * 8          ' �w�Ǘ��Ј��m��
    KMGCSGRP As String * 3          ' �w�Ǘ��ڋq�O���[�v
    KMGCSCOD As String * 8          ' �w�Ǘ��ڋq�R�[�h
    KMGCSHN As String               ' �w�Ǘ��ڋq�i��
    KMGKBNNO As String * 3          ' �w�Ǘ��敪�m��
    COPYMSHN As String * 9          ' �R�s�[���Г��i��
    CONFLAG As String * 1           ' �m�F�t���O
    REINFLAG As String * 1          ' �ĕt�^�t���O
    KMGCSGNO As String              ' �w�Ǘ��ڋq��ʎd�l�ԍ�
    KMGCSKNO As String              ' �w�Ǘ��ڋq�ʎd�l�ԍ�
    KMGCSBNO As String              ' �w�Ǘ��ڋq���i�ԍ�
    KMGTTBNO As Integer             ' �w�Ǘ��Ή��e�[�u���}��
    DIHYSHN As String * 7           ' ��\�Г��i��
    KMGSHKBN As String * 1          ' �w�Ǘ����i�敪
    KMGWKBN As String * 2           ' �w�Ǘ����@�敪
    KMGYKBN As String * 6           ' �w�Ǘ��p�r�敪
    KMGDMKBN As String * 1          ' �w�Ǘ����a�敪
    KMGRSKBN As String * 1          ' �w�Ǘ��q�r�敪
    KMGSNBEF As String * 7          ' �w�Ǘ��Г��i���O��
    KMGSDATE As Date                ' �w�Ǘ��K�p�J�n��
    KMGEDATE As Date                ' �w�Ǘ��K�p�I����
    KMGTDATE As Date                ' �w�Ǘ��o�^����
    KMGRKBNK As String * 1          ' �w�Ǘ����R�敪���
    KMGAKBUM As String * 1          ' �w�Ǘ����݋敪�L��
    KMGOXKBN As String * 1          ' �w�Ǘ��_�f�敪
    KMGIDKBU As String * 1          ' �w�Ǘ��h�c�敪�L��
    KMGIGKBN As String * 1          ' �w�Ǘ��h�f�敪
    KMGWRBKU As String * 1          ' �w�Ǘ��v�`�q�o���z�K�i�L��
    KMGSBKUM As String * 1          ' �w�Ǘ����蕪�z�K�i�L��
    KMGFBKUM As String * 1          ' �w�Ǘ����R���z�K�i�L��
    KMGFPSUM As String * 1          ' �w�Ǘ����R�o�T�C�g�L��
    KMGFOFUM As String * 1          ' �w�Ǘ����R�I�t�Z�b�g�L��
    KMGNCKBN As String * 1          ' �w�Ǘ��m�b�`�敪
    KMGMKBKU As String * 1          ' �w�Ǘ��ʌ����ו��z�K�i�L��
    KMGCMPKU As String * 1          ' �w�Ǘ��b�l�o���H�L��
    KMGSZKBN As String * 1          ' �w�Ǘ��x���ޗ��敪
    KMGSZMUM As String * 1          ' �w�Ǘ��x���ޗ��ʎ�L��
    KMGEPKBN As String              ' �w�Ǘ��d�o���
    KMGEPSKN As String              ' �w�Ǘ��d�o���i�^��
    KMGEPSKB As String * 2          ' �w�Ǘ��d�o�d��敪
    KMGEPRKU As String * 1          ' �w�Ǘ��d�o���R�敪�L��
    KMGEPAKU As String * 1          ' �w�Ǘ��d�o���݋敪�L��
    KMGEPIKU As String * 1          ' �w�Ǘ��d�o�h�c�敪�L��
    KMGKZKBN As String * 1          ' �w�Ǘ��w���ޗ��敪
    KMGSKBN As String * 1           ' �w�Ǘ�����敪
    KMGTRKSI As String * 1          ' �w�Ǘ��s�q�j���w��
    KMGHNCDS As String * 1          ' �w�Ǘ��i���R�[�h�Q��
    KMGHNCDT As String * 1          ' �w�Ǘ��i���R�[�h�Q��
    KMGHNCDD As String * 1          ' �w�Ǘ��i���R�[�h�Q�h
    KMGHNCDK As String * 1          ' �w�Ǘ��i���R�[�h�Q��
    KMGHNCDF As String * 1          ' �w�Ǘ��i���R�[�h�Q�t
    KMGHNCDN As String * 1          ' �w�Ǘ��i���R�[�h�Q��
    KMGHNCDH As String * 1          ' �w�Ǘ��i���R�[�h�Q��
    KMGNOTE As String               ' �w�Ǘ����L
    KMGRS1N As String               ' �w�Ǘ��\���P�Q��
    KMGRS1Y As String               ' �w�Ǘ��\���P�Q�p
    KMGRS2N As String               ' �w�Ǘ��\���Q�Q��
    KMGRS2Y As String               ' �w�Ǘ��\���Q�Q�p
    KMGRS3N As String               ' �w�Ǘ��\���R�Q��
    KMGRS3Y As String               ' �w�Ǘ��\���R�Q�p
    KMGRS4N As String               ' �w�Ǘ��\���S�Q��
    KMGRS4Y As String               ' �w�Ǘ��\���S�Q�p
    KMGRS5N As String               ' �w�Ǘ��\���T�Q��
    KMGRS5Y As String               ' �w�Ǘ��\���T�Q�p
    KMGRS6N As String               ' �w�Ǘ��\���U�Q��
    KMGRS6Y As String               ' �w�Ǘ��\���U�Q�p
    KMGRS7N As String               ' �w�Ǘ��\���V�Q��
    KMGRS7Y As String               ' �w�Ǘ��\���V�Q�p
    KMGRS8N As String               ' �w�Ǘ��\���W�Q��
    KMGRS8Y As String               ' �w�Ǘ��\���W�Q�p
    KMGRS9N As String               ' �w�Ǘ��\���X�Q��
    KMGRS9Y As String               ' �w�Ǘ��\���X�Q�p
    KMGRS10N As String              ' �w�Ǘ��\���P�O�Q��
    KMGRS10Y As String              ' �w�Ǘ��\���P�O�Q�p
    KMGWFLVS As Integer             ' �w�Ǘ��v�e������
    KMGWFLVN As String * 2          ' �w�Ǘ��v�e������
    KMGWFLVC As String              ' �w�Ǘ��v�e�������e
    KMGEPLVS As Integer             ' �w�Ǘ��d�o������
    KMGEPLVN As String * 2          ' �w�Ǘ��d�o������
    KMGEPLVY As String              ' �w�Ǘ��d�o�������e
    KMGSSREC As String * 3          ' �w�Ǘ��d�l�ݒ�L�^
    SSMGKBN As String * 1           ' ���Y�Ǘ��m�F�敪
    HSHSKBN As String * 1           ' �i���ۏ؊m�F�敪
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�l�[���ް�
Public Type typ_TBCME002
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KNURSMAX As Integer             ' �w�[�����b�g�T�C�Y���
    KNURSMIN As Integer             ' �w�[�����b�g�T�C�Y����
    KNUKRMAX As Integer             ' �w�[���\�����b�g���
    KNUNPACK As String * 1          ' �w�[�������p�b�N�Q��
    KNUNPACT As String * 1          ' �w�[�������p�b�N�Q��
    KNUNPACS As String * 1          ' �w�[�������p�b�N�Q�A
    KNUNPACR As String * 1          ' �w�[�������p�b�N�Q��
    KNUNZWAY As String * 1          ' �w�[�������[�U���@�Q��
    KNUNZWYH As String * 1          ' �w�[�������[�U���@�Q�[
    KNUNZWYT As String * 1          ' �w�[�������[�U���@�Q�P
    KNUNZWYW As String * 1          ' �w�[�������[�U���@�Q�v
    KNUNPAC As String * 1           ' �w�[�������
    KNUPACKZ As Integer             ' �w�[���p�b�N�[�U��
    KNUSEALA As String * 1          ' �w�[���V�[���\�t
    KNUSEALK As String * 1          ' �w�[���V�[�����
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lSXL�ް��Y�t
Public Type typ_TBCME003
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSHN As String * 7            ' �w�Ǘ��Г��i��
    KMGSNRNO As Integer             ' �w�Ǘ��Г��i�������ԍ�
    KNUSXDTS As String * 1          ' �w�[���r�w�f�[�^�Y�t�`���Q�Y
    KNUSXDTT As String * 1          ' �w�[���r�w�f�[�^�Y�t�`���Q��
    KNUSXDTY As String * 1          ' �w�[���r�w�f�[�^�Y�t�`���Q�\
    KTSXRTK As String * 1           ' �w�Y�r�w���R�l�Q�R
    KTSXRTD As String * 1           ' �w�Y�r�w���R�l�Q�f
    KTSXRTS As String * 1           ' �w�Y�r�w���R�l�Q�T
    KTSXRMK As String * 1           ' �w�Y�r�w���R�ʓ����z�Q�R
    KTSXRMD As String * 1           ' �w�Y�r�w���R�ʓ����z�Q�f
    KTSXRMS As String * 1           ' �w�Y�r�w���R�ʓ����z�Q�T
    KTSXRM2K As String * 1          ' �w�Y�r�w���R�ʓ����z�Q�Q�R
    KTSXRM2D As String * 1          ' �w�Y�r�w���R�ʓ����z�Q�Q�f
    KTSXRM2S As String * 1          ' �w�Y�r�w���R�ʓ����z�Q�Q�T
    KTSXDIMK As String * 1          ' �w�Y�r�w���a�Q�R
    KTSXDIMD As String * 1          ' �w�Y�r�w���a�Q�f
    KTSXDIMS As String * 1          ' �w�Y�r�w���a�Q�T
    KTSXTMK As String * 1           ' �w�Y�r�w�]�ʖ��x�Q�R
    KTSXTMD As String * 1           ' �w�Y�r�w�]�ʖ��x�Q�f
    KTSXTMS As String * 1           ' �w�Y�r�w�]�ʖ��x�Q�T
    KTSXLTK As String * 1           ' �w�Y�r�w���C�t�^�C���Q�R
    KTSXLTD As String * 1           ' �w�Y�r�w���C�t�^�C���Q�f
    KTSXLTS As String * 1           ' �w�Y�r�w���C�t�^�C���Q�T
    KTSXCNK As String * 1           ' �w�Y�r�w�Y�f�Z�x�Q�R
    KTSXCND As String * 1           ' �w�Y�r�w�Y�f�Z�x�Q�f
    KTSXCNS As String * 1           ' �w�Y�r�w�Y�f�Z�x�Q�T
    KTSXONK As String * 1           ' �w�Y�r�w�_�f�Z�x�Q�R
    KTSXOND As String * 1           ' �w�Y�r�w�_�f�Z�x�Q�f
    KTSXONS As String * 1           ' �w�Y�r�w�_�f�Z�x�Q�T
    KTSXOS1K As String * 1          ' �w�Y�r�w�n�r�e�P�Q�R
    KTSXOS1D As String * 1          ' �w�Y�r�w�n�r�e�P�Q�f
    KTSXOS1S As String * 1          ' �w�Y�r�w�n�r�e�P�Q�T
    KTSXOS2K As String * 1          ' �w�Y�r�w�n�r�e�Q�Q�R
    KTSXOS2D As String * 1          ' �w�Y�r�w�n�r�e�Q�Q�f
    KTSXOS2S As String * 1          ' �w�Y�r�w�n�r�e�Q�Q�T
    KTSXBM1K As String * 1          ' �w�Y�r�w�a�l�c�P�Q�R
    KTSXBM1D As String * 1          ' �w�Y�r�w�a�l�c�P�Q�f
    KTSXBM1S As String * 1          ' �w�Y�r�w�a�l�c�P�Q�T
    KTSXBM2K As String * 1          ' �w�Y�r�w�a�l�c�Q�Q�R
    KTSXBM2D As String * 1          ' �w�Y�r�w�a�l�c�Q�Q�f
    KTSXBM2S As String * 1          ' �w�Y�r�w�a�l�c�Q�Q�T
    KTSXDSOK As String * 1          ' �w�Y�r�w�c�r�n�c�Q�R
    KTSXDSOD As String * 1          ' �w�Y�r�w�c�r�n�c�Q�f
    KTSXDSOS As String * 1          ' �w�Y�r�w�c�r�n�c�Q�T
    KTSXFPDK As String * 1          ' �w�Y�r�w�e�o�c�Q�R
    KTSXFPDD As String * 1          ' �w�Y�r�w�e�o�c�Q�f
    KTSXFPDS As String * 1          ' �w�Y�r�w�e�o�c�Q�T
    KTSXSRK As String * 1           ' �w�Y�r�w�r�q�Q�R
    KTSXSRD As String * 1           ' �w�Y�r�w�r�q�Q�f
    KTSXSRS As String * 1           ' �w�Y�r�w�r�q�Q�T
    KTSXBNK As String * 1           ' �w�Y�r�w�a�Z�x�Q�R
    KTSXBND As String * 1           ' �w�Y�r�w�a�Z�x�Q�f
    KTSXBNS As String * 1           ' �w�Y�r�w�a�Z�x�Q�T
    KTSXSMP As String * 1           ' �w�Y�r�w���i�ʒu�Q�R
    KTSXMPD As String * 1           ' �w�Y�r�w���i�ʒu�Q�f
    KTSXMPS As String * 1           ' �w�Y�r�w���i�ʒu�Q�T
    KTSXODNK As String * 1          ' �w�Y�r�w�_�f�͏o�Z�x�Q�R
    KTSXODND As String * 1          ' �w�Y�r�w�_�f�͏o�Z�x�Q�f
    KTSXODNS As String * 1          ' �w�Y�r�w�_�f�͏o�Z�x�Q�T
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lWF�ް��Y�t
Public Type typ_TBCME004
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSHN As String * 7            ' �w�Ǘ��Г��i��
    KMGSNRNO As Integer             ' �w�Ǘ��Г��i�������ԍ�
    KNWFDTFS As String * 1          ' �w�[���v�e�f�[�^�Y�t�`���Q�Y
    KNWFDTFT As String * 1          ' �w�[���v�e�f�[�^�Y�t�`���Q��
    KNWFDTFY As String * 1          ' �w�[���v�e�f�[�^�Y�t�`���Q�\
    KTWFRTK As String * 1           ' �w�Y�v�e���R�l�Q�R
    KTWFRTD As String * 1           ' �w�Y�v�e���R�l�Q�f
    KTWFRTS As String * 1           ' �w�Y�v�e���R�l�Q�T
    KTWFRMK As String * 1           ' �w�Y�v�e���R�ʓ����z�Q�R
    KTWFRMD As String * 1           ' �w�Y�v�e���R�ʓ����z�Q�f
    KTWFRMS As String * 1           ' �w�Y�v�e���R�ʓ����z�Q�T
    KTWFRM2K As String * 1          ' �w�Y�v�e���R�ʓ����z�Q�Q�R
    KTWFRM2D As String * 1          ' �w�Y�v�e���R�ʓ����z�Q�Q�f
    KTWFRM2S As String * 1          ' �w�Y�v�e���R�ʓ����z�Q�Q�T
    KTWFDIMK As String * 1          ' �w�Y�v�e���a�Q�R
    KTWFDIMD As String * 1          ' �w�Y�v�e���a�Q�f
    KTWFDIMS As String * 1          ' �w�Y�v�e���a�Q�T
    KTWFSAK As String * 1           ' �w�Y�v�e�d����Q�R
    KTWFSAD As String * 1           ' �w�Y�v�e�d����Q�f
    KTWFSAS As String * 1           ' �w�Y�v�e�d����Q�T
    KTWFARK As String * 1           ' �w�Y�v�e���ʓ��͈́Q�R
    KTWFARD As String * 1           ' �w�Y�v�e���ʓ��͈́Q�f
    KTWFARS As String * 1           ' �w�Y�v�e���ʓ��͈́Q�T
    KTWFWARK As String * 1          ' �w�Y�v�e�v�`�q�o�Q�R
    KTWFWARD As String * 1          ' �w�Y�v�e�v�`�q�o�Q�f
    KTWFWARS As String * 1          ' �w�Y�v�e�v�`�q�o�Q�T
    KTWFSK As String * 1            ' �w�Y�v�e����Q�R
    KTWFSD As String * 1            ' �w�Y�v�e����Q�f
    KTWFSS As String * 1            ' �w�Y�v�e����Q�T
    KTWFGBK As String * 1           ' �w�Y�v�e���R�f�a�Q�R
    KTWFGBD As String * 1           ' �w�Y�v�e���R�f�a�Q�f
    KTWFGBS As String * 1           ' �w�Y�v�e���R�f�a�Q�T
    KTWFGFRK As String * 1          ' �w�Y�v�e���R�f�e�q�Q�R
    KTWFGFRD As String * 1          ' �w�Y�v�e���R�f�e�q�Q�f
    KTWFGFRS As String * 1          ' �w�Y�v�e���R�f�e�q�Q�T
    KTWFGFDK As String * 1          ' �w�Y�v�e���R�f�e�c�Q�R
    KTWFGFDD As String * 1          ' �w�Y�v�e���R�f�e�c�Q�f
    KTWFGFDS As String * 1          ' �w�Y�v�e���R�f�e�c�Q�T
    KTWFSBK As String * 1           ' �w�Y�v�e���R�r�a�Q�R
    KTWFSBD As String * 1           ' �w�Y�v�e���R�r�a�Q�f
    KTWFSBS As String * 1           ' �w�Y�v�e���R�r�a�Q�T
    KTWFSFK As String * 1           ' �w�Y�v�e���R�r�e�Q�R
    KTWFSFD As String * 1           ' �w�Y�v�e���R�r�e�Q�f
    KTWFSFS As String * 1           ' �w�Y�v�e���R�r�e�Q�T
    KTWFGBPK As String * 1          ' �w�Y�v�e���R�f�a�o�t�`�Q�R
    KTWFGBPD As String * 1          ' �w�Y�v�e���R�f�a�o�t�`�Q�f
    KTWFGBPS As String * 1          ' �w�Y�v�e���R�f�a�o�t�`�Q�T
    KTWFGFPK As String * 1          ' �w�Y�v�e���R�f�e�q�o�t�`�Q�R
    KTWFGFPD As String * 1          ' �w�Y�v�e���R�f�e�q�o�t�`�Q�f
    KTWFGFPS As String * 1          ' �w�Y�v�e���R�f�e�q�o�t�`�Q�T
    KTWFGDPK As String * 1          ' �w�Y�v�e���R�f�e�c�o�t�`�Q�R
    KTWFGDPD As String * 1          ' �w�Y�v�e���R�f�e�c�o�t�`�Q�f
    KTWFGDPS As String * 1          ' �w�Y�v�e���R�f�e�c�o�t�`�Q�T
    KTWFSBPK As String * 1          ' �w�Y�v�e���R�r�a�o�t�`�Q�R
    KTWFSBPD As String * 1          ' �w�Y�v�e���R�r�a�o�t�`�Q�f
    KTWFSBPS As String * 1          ' �w�Y�v�e���R�r�a�o�t�`�Q�T
    KTWFSFPK As String * 1          ' �w�Y�v�e���R�r�e�o�t�`�Q�R
    KTWFSFPD As String * 1          ' �w�Y�v�e���R�r�e�o�t�`�Q�f
    KTWFSFPS As String * 1          ' �w�Y�v�e���R�r�e�o�t�`�Q�T
    KTWFBDK As String * 1           ' �w�Y�v�e�a�c�Q�R
    KTWFBDD As String * 1           ' �w�Y�v�e�a�c�Q�f
    KTWFBDS As String * 1           ' �w�Y�v�e�a�c�Q�T
    KTWFMKK As String * 1           ' �w�Y�v�e�ʌ����ׁQ�R
    KTWFMKD As String * 1           ' �w�Y�v�e�ʌ����ׁQ�f
    KTWFMKS As String * 1           ' �w�Y�v�e�ʌ����ׁQ�T
    KTWFOTAK As String * 1          ' �w�Y�v�e�_�����ψ��Q�R
    KTWFOTAD As String * 1          ' �w�Y�v�e�_�����ψ��Q�f
    KTWFOTAS As String * 1          ' �w�Y�v�e�_�����ψ��Q�T
    KTWFARAK As String * 1          ' �w�Y�v�e�\�ʑe���Q�R
    KTWFARAD As String * 1          ' �w�Y�v�e�\�ʑe���Q�f
    KTWFARAS As String * 1          ' �w�Y�v�e�\�ʑe���Q�T
    KTWFLTK As String * 1           ' �w�Y�v�e���C�t�^�C���Q�R
    KTWFLTD As String * 1           ' �w�Y�v�e���C�t�^�C���Q�f
    KTWFLTS As String * 1           ' �w�Y�v�e���C�t�^�C���Q�T
    KTWFCNK As String * 1           ' �w�Y�v�e�Y�f�Z�x�Q�R
    KTWFCND As String * 1           ' �w�Y�v�e�Y�f�Z�x�Q�f
    KTWFCNS As String * 1           ' �w�Y�v�e�Y�f�Z�x�Q�T
    KTWFONK As String * 1           ' �w�Y�v�e�_�f�Z�x�Q�R
    KTWFOND As String * 1           ' �w�Y�v�e�_�f�Z�x�Q�f
    KTWFONS As String * 1           ' �w�Y�v�e�_�f�Z�x�Q�T
    KTWFOBK As String * 1           ' �w�Y�v�e�_�f�ʓ����z�Q�R
    KTWFOBD As String * 1           ' �w�Y�v�e�_�f�ʓ����z�Q�f
    KTWFOBS As String * 1           ' �w�Y�v�e�_�f�ʓ����z�Q�T
    KTWFOS1K As String * 1          ' �w�Y�v�e�n�r�e�P�Q�R
    KTWFOS1D As String * 1          ' �w�Y�v�e�n�r�e�P�Q�f
    KTWFOS1S As String * 1          ' �w�Y�v�e�n�r�e�P�Q�T
    KTWFOS2K As String * 1          ' �w�Y�v�e�n�r�e�Q�Q�R
    KTWFOS2D As String * 1          ' �w�Y�v�e�n�r�e�Q�Q�f
    KTWFOS2S As String * 1          ' �w�Y�v�e�n�r�e�Q�Q�T
    KTWFOS3K As String * 1          ' �w�Y�v�e�n�r�e�R�Q�R
    KTWFOS3D As String * 1          ' �w�Y�v�e�n�r�e�R�Q�f
    KTWFOS3S As String * 1          ' �w�Y�v�e�n�r�e�R�Q�T
    KTWFOS4K As String * 1          ' �w�Y�v�e�n�r�e�S�Q�R
    KTWFOS4D As String * 1          ' �w�Y�v�e�n�r�e�S�Q�f
    KTWFOS4S As String * 1          ' �w�Y�v�e�n�r�e�S�Q�T
    KTWFBM1K As String * 1          ' �w�Y�v�e�a�l�c�P�Q�R
    KTWFBM1D As String * 1          ' �w�Y�v�e�a�l�c�P�Q�f
    KTWFBM1S As String * 1          ' �w�Y�v�e�a�l�c�P�Q�T
    KTWFBM2K As String * 1          ' �w�Y�v�e�a�l�c�Q�Q�R
    KTWFBM2D As String * 1          ' �w�Y�v�e�a�l�c�Q�Q�f
    KTWFBM2S As String * 1          ' �w�Y�v�e�a�l�c�Q�Q�T
    KTWFBM3K As String * 1          ' �w�Y�v�e�a�l�c�R�Q�R
    KTWFBM3D As String * 1          ' �w�Y�v�e�a�l�c�R�Q�f
    KTWFBM3S As String * 1          ' �w�Y�v�e�a�l�c�R�Q�T
    KTWFOSPK As String * 1          ' �w�Y�v�e�n�r�o�Q�R
    KTWFOSPD As String * 1          ' �w�Y�v�e�n�r�o�Q�f
    KTWFOSPS As String * 1          ' �w�Y�v�e�n�r�o�Q�T
    KTWFDZOK As String * 1          ' �w�Y�v�e�c�y�͏o�_�f�Z�x�Q�R
    KTWFDZOD As String * 1          ' �w�Y�v�e�c�y�͏o�_�f�Z�x�Q�f
    KTWFDZOS As String * 1          ' �w�Y�v�e�c�y�͏o�_�f�Z�x�Q�T
    KTWFKMHK As String * 1          ' �w�Y�v�e�����ʌX�����Q�R
    KTWFKMHD As String * 1          ' �w�Y�v�e�����ʌX�����Q�f
    KTWFKMHS As String * 1          ' �w�Y�v�e�����ʌX�����Q�T
    KTWFOFKL As String * 1          ' �w�Y�v�e�n�e�P�����Q�R
    KTWFOSDL As String * 1          ' �w�Y�v�e�n�e�P�����Q�f
    KTWFOFSL As String * 1          ' �w�Y�v�e�n�e�P�����Q�T
    KTWFMWK As String * 1           ' �w�Y�v�e�ʎ�ЁQ�R
    KTWFMWD As String * 1           ' �w�Y�v�e�ʎ�ЁQ�f
    KTWFMWS As String * 1           ' �w�Y�v�e�ʎ�ЁQ�T
    KTWFMKWK As String * 1          ' �w�Y�v�e�����בw�ЁQ�R
    KTWFMKWD As String * 1          ' �w�Y�v�e�����בw�ЁQ�f
    KTWFMKWS As String * 1          ' �w�Y�v�e�����בw�ЁQ�T
    KTWFNOXK As String * 1          ' �w�Y�v�e�M�_�������Q�R
    KTWFNOXD As String * 1          ' �w�Y�v�e�M�_�������Q�f
    KTWFNOXS As String * 1          ' �w�Y�v�e�M�_�������Q�T
    KTWFPSK As String * 1           ' �w�Y�v�e�|���V�����Q�R
    KTWFPSD As String * 1           ' �w�Y�v�e�|���V�����Q�f
    KTWFPSS As String * 1           ' �w�Y�v�e�|���V�����Q�T
    KTWFCVDK As String * 1          ' �w�Y�v�e�b�u�c���Q�R
    KTWFCVDD As String * 1          ' �w�Y�v�e�b�u�c���Q�f
    KTWFCVDS As String * 1          ' �w�Y�v�e�b�u�c���Q�T
    KTWFSONK As String * 1          ' �w�Y�v�e�͏o�_�f�Z�x�Q�R
    KTWFSOND As String * 1          ' �w�Y�v�e�͏o�_�f�Z�x�Q�f
    KTWFSONS As String * 1          ' �w�Y�v�e�͏o�_�f�Z�x�Q�T
    KTWFOFPK As String * 1          ' �w�Y�v�e�n�e�P�ʒu�Q�R
    KTWFOFPD As String * 1          ' �w�Y�v�e�n�e�P�ʒu�Q�f
    KTWFOFPS As String * 1          ' �w�Y�v�e�n�e�P�ʒu�Q�T
    KTWFGDK As String * 1           ' �w�Y�v�e�f�c�Q�R
    KTWFGDD As String * 1           ' �w�Y�v�e�f�c�Q�f
    KTWFGDS As String * 1           ' �w�Y�v�e�f�c�Q�T
    KTWFDSOK As String * 1          ' �w�Y�v�e�c�r�n�c�Q�R
    KTWFDSOD As String * 1          ' �w�Y�v�e�c�r�n�c�Q�f
    KTWFDSOS As String * 1          ' �w�Y�v�e�c�r�n�c�Q�T
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lSXL�ް��P
Public Type typ_TBCME005
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    CONFLAG As String * 1           ' �m�F�t���O
    REINFLAG As String * 1          ' �ĕt�^�t���O
    KSXTYPE As String * 1           ' �w�r�w�^�C�v
    KSXTYPKW As String * 1          ' �w�r�w�^�C�v�������@
    KSXTYPKB As String * 1          ' �w�r�w�^�C�v�����敪
    KSXDOP As String * 1            ' �w�r�w�h�[�p���g
    KSXRMIN As Double               ' �w�r�w���R����
    KSXRMAX As Double               ' �w�r�w���R���
    KSXRUNIT As String * 1          ' �w�r�w���R�P��
    KSXRSPOH As String * 1          ' �w�r�w���R����ʒu�Q��
    KSXRSPOT As String * 1          ' �w�r�w���R����ʒu�Q�_
    KSXRSPOI As String * 1          ' �w�r�w���R����ʒu�Q��
    KSXRHWYT As String * 1          ' �w�r�w���R�ۏؕ��@�Q��
    KSXRHWYS As String * 1          ' �w�r�w���R�ۏؕ��@�Q��
    KSXRKKBN As String * 1          ' �w�r�w���R�����敪
    KSXRKWAY As String * 2          ' �w�r�w���R�������@
    KSXRKHNM As String * 1          ' �w�r�w���R�����p�x�Q��
    KSXRKHNI As String * 1          ' �w�r�w���R�����p�x�Q��
    KSXRKHNH As String * 1          ' �w�r�w���R�����p�x�Q��
    KSXRKHNS As String * 1          ' �w�r�w���R�����p�x�Q��
    KSXRMCAL As String * 1          ' �w�r�w���R�ʓ��v�Z
    KSXRMBNP As Double              ' �w�r�w���R�ʓ����z
    KSXRMCL2 As String * 1          ' �w�r�w���R�ʓ��v�Z�Q
    KSXRMBP2 As Double              ' �w�r�w���R�ʓ����z�Q
    KSXRSDEV As Double              ' �w�r�w���R�W���΍�
    KSXRAMIN As Double              ' �w�r�w���R���ω���
    KSXRAMAX As Double              ' �w�r�w���R���Ϗ��
    KSXFORM As String * 1           ' �w�r�w�`��
    KSXD1CEN As Double              ' �w�r�w���a�P���S
    KSXD1MIN As Double              ' �w�r�w���a�P����
    KSXD1MAX As Double              ' �w�r�w���a�P���
    KSXD1KBN As String * 1          ' �w�r�w���a�P�����敪
    KSXD2CEN As Double              ' �w�r�w���a�Q���S
    KSXD2MIN As Double              ' �w�r�w���a�Q����
    KSXD2MAX As Double              ' �w�r�w���a�Q���
    KSXD2KBN As String * 1          ' �w�r�w���a�Q�����敪
    KSXDUNIT As String * 1          ' �w�r�w���a�P��
    KSXCDIR As String * 1           ' �w�r�w�����ʕ���
    KSXCSCEN As Double              ' �w�r�w�����ʌX���S
    KSXCSMIN As Double              ' �w�r�w�����ʌX����
    KSXCSMAX As Double              ' �w�r�w�����ʌX���
    KSXCKWAY As String * 2          ' �w�r�w�����ʌ������@
    KSXCKHNM As String * 1          ' �w�r�w�����ʌ����p�x�Q��
    KSXCKHNI As String * 1          ' �w�r�w�����ʌ����p�x�Q��
    KSXCKHNH As String * 1          ' �w�r�w�����ʌ����p�x�Q��
    KSXCKHNS As String * 1          ' �w�r�w�����ʌ����p�x�Q��
    KSXCSDIR As String * 2          ' �w�r�w�����ʌX����
    KSXCSDIS As String * 1          ' �w�r�w�����ʌX���ʎw��
    KSXCTDIR As String * 2          ' �w�r�w�����ʌX�c����
    KSXCTCEN As Double              ' �w�r�w�����ʌX�c���S
    KSXCTMIN As Double              ' �w�r�w�����ʌX�c����
    KSXCTMAX As Double              ' �w�r�w�����ʌX�c���
    KSXCYDIR As String * 2          ' �w�r�w�����ʌX������
    KSXCYCEN As Double              ' �w�r�w�����ʌX�����S
    KSXCYMIN As Double              ' �w�r�w�����ʌX������
    KSXCYMAX As Double              ' �w�r�w�����ʌX�����
    KSXOF1PD As String * 2          ' �w�r�w�n�e�P�ʒu����
    KSXOF1PN As Double              ' �w�r�w�n�e�P�ʒu����
    KSXOF1PX As Double              ' �w�r�w�n�e�P�ʒu���
    KSXOF1PK As String * 1          ' �w�r�w�n�e�P�ʒu�����敪
    KSXOF1PW As String * 2          ' �w�r�w�n�e�P�ʒu�������@
    KSXOF1LC As Double              ' �w�r�w�n�e�P�����S
    KSXOF1LN As Double              ' �w�r�w�n�e�P������
    KSXOF1LX As Double              ' �w�r�w�n�e�P�����
    KSXOF1LK As String * 1          ' �w�r�w�n�e�P�������敪
    KSXOF1DC As Double              ' �w�r�w�n�e�P���a���S
    KSXOF1DN As Double              ' �w�r�w�n�e�P���a����
    KSXOF1DX As Double              ' �w�r�w�n�e�P���a���
    KSXOF1DK As String * 1          ' �w�r�w�n�e�P���a�����敪
    KSXDFORM As String * 1          ' �w�r�w�a�`��
    KSXDFKBN As String * 1          ' �w�r�w�a�`�󌟍��敪
    KSXDPDRC As String * 1          ' �w�r�w�a�ʒu����
    KSXDPACN As Integer             ' �w�r�w�a�ʒu�p�x���S
    KSXDPAMN As Integer             ' �w�r�w�a�ʒu�p�x����
    KSXDPAMX As Integer             ' �w�r�w�a�ʒu�p�x���
    KSXDPKWY As String * 2          ' �w�r�w�a�ʒu�������@
    KSXDPKBN As String * 1          ' �w�r�w�a�ʒu�����敪
    KSXDPDIR As String * 2          ' �w�r�w�a�ʒu����
    KSXDPMIN As Double              ' �w�r�w�a�ʒu����
    KSXDPMAX As Double              ' �w�r�w�a�ʒu���
    KSXDWCEN As Double              ' �w�r�w�a�В��S
    KSXDWMIN As Double              ' �w�r�w�a�Љ���
    DSXDWMAX As Double              ' �w�r�w�a�Џ��
    KSXDWKBN As String * 1          ' �w�r�w�a�Ќ����敪
    KSXDDCEN As Double              ' �w�r�w�a�[���S
    KSXDDMIN As Double              ' �w�r�w�a�[����
    KSXDDMAX As Double              ' �w�r�w�a�[���
    KSXDDKBN As String * 1          ' �w�r�w�a�[�����敪
    KSXDACEN As Double              ' �w�r�w�a�p�x���S
    KSXDAMIN As Double              ' �w�r�w�a�p�x����
    KSXDAMAX As Double              ' �w�r�w�a�p�x���
    KSXDAKBN As String * 1          ' �w�r�w�a�p�x�����敪
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lSXL�ް��Q
Public Type typ_TBCME006
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KSXTMMAX As Long                ' �w�r�w�]�ʖ��x���
    KSXTMSPH As String * 1          ' �w�r�w�]�ʖ��x����ʒu�Q��
    KSXTMSPT As String * 1          ' �w�r�w�]�ʖ��x����ʒu�Q�_
    KSXTMSPR As String * 1          ' �w�r�w�]�ʖ��x����ʒu�Q��
    KSXTMKBN As String * 1          ' �w�r�w�]�ʖ��x�����敪
    KSXTMKHM As String * 1          ' �w�r�w�]�ʖ��x�����p�x�Q��
    KSXTMKHI As String * 1          ' �w�r�w�]�ʖ��x�����p�x�Q��
    KSXTMKHH As String * 1          ' �w�r�w�]�ʖ��x�����p�x�Q��
    KSXTMKHS As String * 1          ' �w�r�w�]�ʖ��x�����p�x�Q��
    KSXLTMIN As Integer             ' �w�r�w�k�^�C������
    KSXLTMAX As Integer             ' �w�r�w�k�^�C�����
    KSXLTUNT As String * 1          ' �w�r�w�k�^�C���P��
    KSXLTSPH As String * 1          ' �w�r�w�k�^�C������ʒu�Q��
    KSXLTSPT As String * 1          ' �w�r�w�k�^�C������ʒu�Q�_
    KSXLTSPI As String * 1          ' �w�r�w�k�^�C������ʒu�Q��
    KSXLTHWT As String * 1          ' �w�r�w�k�^�C���ۏؕ��@�Q��
    KSXLTHWS As String * 1          ' �w�r�w�k�^�C���ۏؕ��@�Q��
    KSXLTKWY As String * 2          ' �w�r�w�k�^�C���������@
    KSXLTNSW As String * 2          ' �w�r�w�k�^�C���M�����@
    KSXLTKBN As String * 1          ' �w�r�w�k�^�C�������敪
    KSXLTKHM As String * 1          ' �w�r�w�k�^�C�������p�x�Q��
    KSXLTKHI As String * 1          ' �w�r�w�k�^�C�������p�x�Q��
    KSXLTKHH As String * 1          ' �w�r�w�k�^�C�������p�x�Q��
    KSXLTKHS As String * 1          ' �w�r�w�k�^�C�������p�x�Q��
    KSXLTMBP As Double              ' �w�r�w�k�^�C���ʓ����z
    KSXLTMCL As String * 1          ' �w�r�w�k�^�C���ʓ��v�Z
    KSXCNMIN As Double              ' �w�r�w�Y�f�Z�x����
    KSXCNMAX As Double              ' �w�r�w�Y�f�Z�x���
    KSXCNIND As String * 2          ' �w�r�w�Y�f�Z�x�w��
    KSXCNUNT As String * 1          ' �w�r�w�Y�f�Z�x�P��
    KSXCNSPH As String * 1          ' �w�r�w�Y�f�Z�x����ʒu�Q��
    KSXCNSPT As String * 1          ' �w�r�w�Y�f�Z�x����ʒu�Q�_
    KSXCNSPI As String * 1          ' �w�r�w�Y�f�Z�x����ʒu�Q��
    KSXCNHWT As String * 1          ' �w�r�w�Y�f�Z�x�ۏؕ��@�Q��
    KSXCNHWS As String * 1          ' �w�r�w�Y�f�Z�x�ۏؕ��@�Q��
    KSXCNKWY As String * 2          ' �w�r�w�Y�f�Z�x�������@
    KSXCNKBN As String * 1          ' �w�r�w�Y�f�Z�x�����敪
    KSXONMIN As Double              ' �w�r�w�_�f�Z�x����
    KSXONMAX As Double              ' �w�r�w�_�f�Z�x���
    KSXONIND As String * 2          ' �w�r�w�_�f�Z�x�w��
    KSXONUNT As String * 1          ' �w�r�w�_�f�Z�x�P��
    KSXONSPH As String * 1          ' �w�r�w�_�f�Z�x����ʒu�Q��
    KSXONSPT As String * 1          ' �w�r�w�_�f�Z�x����ʒu�Q�_
    KSXONSPI As String * 1          ' �w�r�w�_�f�Z�x����ʒu�Q��
    KSXONHWT As String * 1          ' �w�r�w�_�f�Z�x�ۏؕ��@�Q��
    KSXONHWS As String * 1          ' �w�r�w�_�f�Z�x�ۏؕ��@�Q��
    KSXONKWY As String * 2          ' �w�r�w�_�f�Z�x�������@
    KSXONKBN As String * 1          ' �w�r�w�_�f�Z�x�����敪
    KSXONKHM As String * 1          ' �w�r�w�_�f�Z�x�����p�x�Q��
    KSXONKHI As String * 1          ' �w�r�w�_�f�Z�x�����p�x�Q��
    KSXONKHH As String * 1          ' �w�r�w�_�f�Z�x�����p�x�Q��
    KSXONKHS As String * 1          ' �w�r�w�_�f�Z�x�����p�x�Q��
    KSXONMBP As Double              ' �w�r�w�_�f�Z�x�ʓ����z
    KSXONMCL As String * 1          ' �w�r�w�_�f�Z�x�ʓ��v�Z
    KSXONLTB As Double              ' �w�r�w�_�f�Z�x�k�s���z
    KSXONLTC As String * 1          ' �w�r�w�_�f�Z�x�k�s�v�Z
    KSXONSDV As Double              ' �w�r�w�_�f�Z�x�W���΍�
    KSXONAMN As Double              ' �w�r�w�_�f�Z�x���ω���
    KSXONAMX As Double              ' �w�r�w�_�f�Z�x���Ϗ��
    KSXONMNH As Double              ' �w�r�w�_�f�Z�x�����␳
    KSXONMXH As Double              ' �w�r�w�_�f�Z�x����␳
    KSXONHCL As String * 2          ' �w�r�w�_�f�Z�x�␳�v�Z
    KSXGSFIN As String * 1          ' �w�r�w�O���d�グ
    KSXCLMIN As Integer             ' �w�r�w����������
    KSXCLMAX As Integer             ' �w�r�w���������
    KSXCLPMN As Integer             ' �w�r�w���������e����
    KSXCLPR As Double               ' �w�r�w���������e�䗦
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type


' �ڋq�d�lSXL�ް��R
Public Type typ_TBCME007
    hinban As String * 8            ' �i��
    mnorevno As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KMGSRRNO As String * 9          ' �w�Ǘ��d�l�o�^�˗��ԍ�
    KSXOF1MAX As Double             ' �w�r�w�n�r�e���
    KSXOF1AMX As Double             ' �w�r�w�n�r�e���Ϗ��
    KSXOF1SPH As String * 1         ' �w�r�w�n�r�e����ʒu�Q��
    KSXOF1SPT As String * 1         ' �w�r�w�n�r�e����ʒu�Q�_
    KSXOF1SPR As String * 1         ' �w�r�w�n�r�e����ʒu�Q��
    KSXOF1HWT As String * 1         ' �w�r�w�n�r�e�ۏؕ��@�Q��
    KSXOF1HWS As String * 1         ' �w�r�w�n�r�e�ۏؕ��@�Q��
    KSXOF1SZY As String * 1         ' �w�r�w�n�r�e�������
    KSXOF1KBN As String * 1         ' �w�r�w�n�r�e�����敪
    KSXOF1KHM As String * 1         ' �w�r�w�n�r�e�����p�x�Q��
    KSXOF1KHI As String * 1         ' �w�r�w�n�r�e�����p�x�Q��
    KSXOF1KHH As String * 1         ' �w�r�w�n�r�e�����p�x�Q��
    KSXOF1KHS As String * 1         ' �w�r�w�n�r�e�����p�x�Q��
    KSXOF1FGS As String * 1         ' �w�r�w�n�r�e���͋C�K�X
    KSXOF1CET As Integer            ' �w�r�w�n�r�e�I���d�s��
    KSXOF1NSW As String * 2         ' �w�r�w�n�r�e�M�����@
    KSXOF1SO1 As Integer            ' �w�r�w�n�r�e�������x�P
    KSXOF1ST1 As Integer            ' �w�r�w�n�r�e�������ԂP
    KSXOF2MX As Double              ' �w�r�w�n�r�e�Q���
    KSXOF2AX As Double              ' �w�r�w�n�r�e�Q���Ϗ��
    KSXOF2SH As String * 1          ' �w�r�w�n�r�e�Q����ʒu�Q��
    KSXOF2ST As String * 1          ' �w�r�w�n�r�e�Q����ʒu�Q�_
    KSXOF2SR As String * 1          ' �w�r�w�n�r�e�Q����ʒu�Q��
    KSXOF2HT As String * 1          ' �w�r�w�n�r�e�Q�ۏؕ��@�Q��
    KSXOF2HS As String * 1          ' �w�r�w�n�r�e�Q�ۏؕ��@�Q��
    KSXOF2SZ As String * 1          ' �w�r�w�n�r�e�Q�������
    KSXOF2KB As String * 1          ' �w�r�w�n�r�e�Q�����敪
    KSXOF2KM As String * 1          ' �w�r�w�n�r�e�Q�����p�x�Q��
    KSXOF2KI As String * 1          ' �w�r�w�n�r�e�Q�����p�x�Q��
    KSXOF2KH As String * 1          ' �w�r�w�n�r�e�Q�����p�x�Q��
    KSXOF2KS As String * 1          ' �w�r�w�n�r�e�Q�����p�x�Q��
    KSXOF2GS As String * 1          ' �w�r�w�n�r�e�Q���͋C�K�X
    KSXOF2ET As Integer             ' �w�r�w�n�r�e�Q�I���d�s��
    KSXOF2NS As String * 2          ' �w�r�w�n�r�e�Q�M�����@
    KSXOF2O1 As Integer             ' �w�r�w�n�r�e�Q�������x�P
    KSXOF2T1 As Integer             ' �w�r�w�n�r�e�Q�������ԂP
    KSXBMMAX As Double              ' �w�r�w�a�l�c���ω���
    KSXBMMIN As Double              ' �w�r�w�a�l�c���Ϗ��
    KSXBMSPH As String * 1          ' �w�r�w�a�l�c����ʒu�Q��
    KSXBMSPT As String * 1          ' �w�r�w�a�l�c����ʒu�Q�_
    KSXBMSPR As String * 1          ' �w�r�w�a�l�c����ʒu�Q��
    KSXBMHWT As String * 1          ' �w�r�w�a�l�c�ۏؕ��@�Q��
    KSXBMHWS As String * 1          ' �w�r�w�a�l�c�ۏؕ��@�Q��
    KSXBMSZY As String * 1          ' �w�r�w�a�l�c�������
    KSXBMKBN As String * 1          ' �w�r�w�a�l�c�����敪
    KSXBMKHM As String * 1          ' �w�r�w�a�l�c�����p�x�Q��
    KSXBMKHI As String * 1          ' �w�r�w�a�l�c�����p�x�Q��
    KSXBMKHH As String * 1          ' �w�r�w�a�l�c�����p�x�Q��
    KSXBMKHS As String * 1          ' �w�r�w�a�l�c�����p�x�Q��
    KSXBMFGS As String * 1          ' �w�r�w�a�l�c���͋C�K�X
    KSXBMCET As Integer             ' �w�r�w�a�l�c�I���d�s��
    KSXBMNS As String * 2           ' �w�r�w�a�l�c�M�����@
    KSXBM2AN As Double              ' �w�r�w�a�l�c�Q���ω���
    KSXBM2AX As Double              ' �w�r�w�a�l�c�Q���Ϗ��
    KSXBM2SH As String * 1          ' �w�r�w�a�l�c�Q����ʒu�Q��
    KSXBM2ST As String * 1          ' �w�r�w�a�l�c�Q����ʒu�Q�_
    KSXBM2SR As String * 1          ' �w�r�w�a�l�c�Q����ʒu�Q��
    KSXBM2HT As String * 1          ' �w�r�w�a�l�c�Q�ۏؕ��@�Q��
    KSXBM2HS As String * 1          ' �w�r�w�a�l�c�Q�ۏؕ��@�Q��
    KSXBM2SZ As String * 1          ' �w�r�w�a�l�c�Q�������
    KSXBM2KB As String * 1          ' �w�r�w�a�l�c�Q�����敪
    KSXBM2KM As String * 1          ' �w�r�w�a�l�c�Q�����p�x�Q��
    KSXBM2KI As String * 1          ' �w�r�w�a�l�c�Q�����p�x�Q��
    KSXBM2KH As String * 1          ' �w�r�w�a�l�c�Q�����p�x�Q��
    KSXBM2KS As String * 1          ' �w�r�w�a�l�c�Q�����p�x�Q��
    KSXBM2GS As String * 1          ' �w�r�w�a�l�c�Q���͋C�K�X
    KSXBM2ET As Integer             ' �w�r�w�a�l�c�Q�I���d�s��
    KSXBM2NS As String * 2          ' �w�r�w�a�l�c�Q�M�����@
    KSXDENKU As String * 1          ' �w�r�w�c���������L��
    KSXDENMX As Integer             ' �w�r�w�c�������
    KSXDENMN As Integer             ' �w�r�w�c��������
    KSXDENHT As String * 1          ' �w�r�w�c�����ۏؕ��@�Q��
    KSXDENHS As String * 1          ' �w�r�w�c�����ۏؕ��@�Q��
    KSXLDLKU As String * 1          ' �w�r�w�k�^�c�k�����L��
    KSXLDLMX As Integer             ' �w�r�w�k�^�c�k���
    KSXLDLMN As Integer             ' �w�r�w�k�^�c�k����
    KSXLDLHT As String * 1          ' �w�r�w�k�^�c�k�ۏؕ��@�Q��
    KSXLDLHS As String * 1          ' �w�r�w�k�^�c�k�ۏؕ��@�Q��
    KSXDVDKU As String * 1          ' �w�r�w�c�u�c�Q�����L��
    KSXDVDMX As Integer             ' �w�r�w�c�u�c�Q���
    KSXDVDMN As Integer             ' �w�r�w�c�u�c�Q����
    KSXDVDHT As String * 1          ' �w�r�w�c�u�c�Q�ۏؕ��@�Q��
    KSXDVDHS As String * 1          ' �w�r�w�c�u�c�Q�ۏؕ��@�Q��
    KSXGDSPH As String * 1          ' �w�r�w�f�c����ʒu�Q��
    KSXGDSPT As String * 1          ' �w�r�w�f�c����ʒu�Q�_
    KSXGDSPR As String * 1          ' �w�r�w�f�c����ʒu�Q��
    KSXGDSZY As String * 1          ' �w�r�w�f�c�������
    KSXGDZAR As Integer             ' �w�r�w�f�c���O�̈�
    KSXGDKHM As String * 1          ' �w�r�w�f�c�����p�x�Q��
    KSXGDKHI As String * 1          ' �w�r�w�f�c�����p�x�Q��
    KSXGDKHH As String * 1          ' �w�r�w�f�c�����p�x�Q��
    KSXGDKHS As String * 1          ' �w�r�w�f�c�����p�x�Q��
    KSXDSOKE As String * 1          ' �w�r�w�c�r�n�c����
    KSXDSOMX As Long                ' �w�r�w�c�r�n�c���
    KSXDSOMN As Long                ' �w�r�w�c�r�n�c����
    KSXDSOAX As Integer             ' �w�r�w�c�r�n�c�̈���
    KSXDSOAN As Integer             ' �w�r�w�c�r�n�c�̈扺��
    KSXDSOHT As String * 1          ' �w�r�w�c�r�n�c�ۏؕ��@�Q��
    KSXDSOHS As String * 1          ' �w�r�w�c�r�n�c�ۏؕ��@�Q��
    KSXDSOKM As String * 1          ' �w�r�w�c�r�n�c�����p�x�Q��
    KSXDSOKI As String * 1          ' �w�r�w�c�r�n�c�����p�x�Q��
    KSXDSOKH As String * 1          ' �w�r�w�c�r�n�c�����p�x�Q��
    KSXDSOKS As String * 1          ' �w�r�w�c�r�n�c�����p�x�Q��
    KSXCDOP As String * 1           ' �w�r�w�����h�[�v
    IFKBN As String * 4             ' �h�^�e�敪
    SYORIKBN As String * 1          ' �����敪
    SPECRRNO As String * 9          ' �d�l�o�^�˗��ԍ�
    SXLMCNO As String * 12          ' �r�w�k��������ԍ�
    WFMCNO As String * 12           ' �v�e��������ԍ�
    StaffID As String * 8           ' �Ј�ID
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
    SENDFLAG As String * 1          ' ���M�t���O
    SENDDATE As Date                ' ���M���t
End Type

' �����u���b�N�ꗗ
Public Type typ_LackBlk
    BLOCKID As String * 12      ' �u���b�NID
    INGOTPOS As Integer         ' �������J�n�ʒu
    REALLEN As Integer          ' ����
    REJDTTM As Date             ' ������
    NYUKO As Integer            ' ���̃u���b�N�����ɍς�(1:���ɍ� 0:������)
    sBlockId As String * 12     ' ���o���P�ʂ̐擪�u���b�NID
    MINYUKO As Integer          ' ���o���P�ʂł̖����Ƀu���b�N��
    PUPTN As String             ' ���������     2004/12/08 �ǉ�
    HOLDFLG As String * 1       ' ΰ��ދ敪�@05/01/31 ooba
    WFHOLDFLG As String * 1     ' WFΰ��ދ敪�@05/01/31 ooba
    WFHUFLG As String * 1       ' WF�U��FLG�@06/02/06 ooba
    MUKESAKI As String          ' ���� 07/09/03 SPK Tsutsumi Add
    Koutei As String * 5        ' �H��(XSDCB)�@08/01/31 ooba
    KANREN As String * 1        ' �֘A��ۯ��L���@08/01/31 ooba
    AGRSTATUS As String             ' ���F�m�F�敪 add SETkimizuka
    STOP    As String               ' ��~ add SETkimizuka
    CAUSE   As String               ' ��~���R add SETkimizuka
    PRINTNO As String               ' ��s�]�� add SETkimizuka
    'Add Start 2010/07/08 SMPK Nakamura
    HINCNT  As String           '�u���b�N���i�Ԑ�
    hinban  As String           '�u���b�N���i��
    CW740STS As String          'CW740�X�e�[�^�X
    'Add End 2010/07/08 SMPK Nakamura
End Type

'add start 2003/04/18 hitec)�㓡�@--------
' �����u���b�N�ꗗ(�\���p�j
Public Type tbl_DispLack
'Chg Start 2010/07/08 SMPK Nakamura
'    BLOCKID As String * 12      ' �u���b�NID
    SELECTED As Long            ' �I������
    BLOCKID As String           ' �u���b�NID
'Chg End 2010/07/08 SMPK Nakamura
    INGOTPOS As Integer         ' �������J�n�ʒu
    REALLEN As Integer          ' ����
    REJDTTM As String             ' ������
    PUPTN As String             ' ���������     2004/12/08 �ǉ�
    HOLDFLG As String * 1       ' ΰ��ދ敪�@05/01/31 ooba
    WFHOLDFLG As String * 1     ' WFΰ��ދ敪�@05/01/31 ooba
    WFHUFLG As String * 1       ' WF�U��FLG�@06/02/06 ooba
    MUKESAKI As String          ' ���� 07/09/03 SPK Tsutsumi Add
    Koutei As String * 5        ' �H��(XSDCB)�@08/01/31 ooba
    KANREN As String * 1        ' �֘A��ۯ��L���@08/01/31 ooba
    AGRSTATUS As String             ' ���F�m�F�敪 add SETkimizuka
    STOP    As String               ' ��~ add SETkimizuka
    CAUSE   As String               ' ��~���R add SETkimizuka
    PRINTNO As String               ' ��s�]�� add SETkimizuka
    'Add Start 2010/07/08 SMPK Nakamura
    hinban  As String           '�u���b�N���i��
    WAITBLOCK As String         '�֘A�҂��u���b�N�L��
    CW740STS As String          'CW740�X�e�[�^�X
    'Add End 2010/07/08 SMPK Nakamura
End Type
Public tblDispLack() As tbl_DispLack
'add end 2003/04/18 hitec)�㓡�@--------

'add 2005/11/11 ����->
'10�����Z�l�擾�\����
Public Type typ_OumConvSet
    CTR01A9 As String          '�X��
    CTR02A9 As String          '�ؕ�
    CTR03A9 As String          '�ݒ�l
End Type
'add 2005/11/11 ����<-

'add 2009/07/22 SETsw Nakada -->
Public Type typ_TBCMJ020
    CRYNUM     As String * 12       ' �����ԍ�
    POSITION   As Integer           ' �ʒu
    SMPKBN     As String * 1        ' �T���v���敪
    TRANCNT    As Integer           ' ������
    TRANCOND   As String * 1        ' ��������
    BLOCKID    As String * 12       ' �u���b�NID
    SMPLNO     As Long              ' �T���v���m��
    SMPLUMU    As String * 1        ' �T���v���L��
    KRPROCCD   As String * 5        ' �Ǘ��H���R�[�h
    PROCCODE   As String * 5        ' �H���R�[�h
    N2NOUDO    As Double            ' �m�Q�Z�x����
    N2NI       As Integer           ' �m�Q�Z�x�w��
    TSTAFFID   As String * 8        ' �o�^�Ј�ID
    REGDATE    As Date              ' �o�^���t
    KSTAFFID   As String * 8        ' �X�V�Ј�ID
    UPDDATE    As Date              ' �X�V���t
    SENDFLAG   As String * 1        ' ���M�t���O
    SENDDATE   As Date              ' ���M���t
End Type
'add 2009/07/22 SETsw Nakada <--

