Attribute VB_Name = "s_cmbc039_SQL"
Option Explicit

'' WF�Z���^�[��������҂��ꗗ

' SXL�Ǘ�
Public Type DBDRV_scmzc_fcmlc001b_SXL039
    CRYNUMCA As String * 12        ' �����ԍ�
    INPOSCA As Integer             ' �������J�n�ʒu
    GNLCA As Integer               ' ���ݒ���
    SXLIDCA As String * 13         ' SXLID
    GNKKNTCA As String * 5         ' ���݊Ǘ��H��(���g�p)
    GNWKNTCA As String * 5         ' ���ݍH��
    NOWPROC As String * 5          ' ���ݍH��
    NEKKNTCA As String * 5         ' �ŏI�ʉߊǗ��H��(���g�p)
    NEWKNTCA As String * 5         ' �ŏI�ʉߍH��
    SAKJCA As String * 1           ' �폜�敪
    LSTATBCA As String * 1         ' �ŏI��ԋ敪
    HOLDBCA As String * 1          ' �z�[���h�敪
    HINBCA As String * 8           ' �i��
    REVNUMCA As Integer            ' ���i�ԍ������ԍ�
    FACTORYCA As String * 1        ' �H��
    OPECA As String * 1            ' ���Ə���
    MAICB As Integer               ' ����
    TDAYCB As Date                 ' �o�^���t
    KDAYCA As Date                 ' �X�V���t
    HOLDBCB As String * 1          ' ΰ��ދ敪�@06/02/08 ooba
    WFHOLDFLGCB As String * 1      ' WFΰ��ދ敪�@06/02/08 ooba
    KETURAKU As Boolean            ' �������L���t���O
    WFSMP() As typ_XSDCW           ' �T���v���Ǘ��iTOP�ATAIL�� �Q���R�[�h�j
    PLANTCAT As String             ' ���� 07/09/04 SPK Tsutsumi Add
    KANREN As String * 1            ' �֘A��ۯ��L���@08/01/31 ooba
    AGRSTATUS  As String            ' ���F�m�F�敪 add SETkimizuka
    STOP    As String               ' ��~ add SETkimizuka
    CAUSE   As String               ' ��~���R add SETkimizuka
    PRINTNO As String               ' ��s�]�� add SETkimizuka
End Type

'WF�Z���^�[��������

'���͗p
Public Type type_DBDRV_scmzc_fcmlc001c_In039
    HIN As tFullHinban             ' �i��(full)
    SAMPLEID As String * 16        ' �T���v��ID
    SXLID As String * 13           ' SXLID
End Type

'WF���i�d�l�擾�p
Public Type type_DBDRV_scmzc_fcmlc001c_Siyou039
    HWFTYPE As String * 1          ' �i�v�e�^�C�v
    HWFCDIR As String * 1          ' �i�v�e�����ʕ�
    HWFCDOP As String * 1          ' �i�v�e�����h�[�v

    HWFRMIN As Double              ' �i�v�e���R����
    HWFRMAX As Double              ' �i�v�e���R���
    HWFRSPOH As String * 1         ' �i�v�e���R����ʒu�Q��
    HWFRSPOT As String * 1         ' �i�v�e���R����ʒu�Q�_
    HWFRSPOI As String * 1         ' �i�v�e���R����ʒu�Q��
    HWFRHWYT As String * 1         ' �i�v�e���R�ۏؕ��@�Q��
    HWFRHWYS As String * 1         ' �i�v�e���R�ۏؕ��@�Q��
    HWFRMCAL As String * 1         ' �i�v�e���R�ʓ��v�Z 2001/11/08 S.Sano
    HWFRAMIN As Double             ' �i�v�e���R���ω���
    HWFRAMAX As Double             ' �i�v�e���R���Ϗ��
    HWFRMBNP As Double             ' �i�v�e���R�ʓ����z

    HWFMKMIN As Double             ' �i�v�e�����בw����
    HWFMKMAX As Double             ' �i�v�e�����בw���
    HWFMKSPH As String * 1         ' �i�v�e�����בw����ʒu�Q��
    HWFMKSPT As String * 1         ' �i�v�e�����בw����ʒu�Q�_
    HWFMKSPR As String * 1         ' �i�v�e�����בw����ʒu�Q��
    HWFMKHWT As String * 1         ' �i�v�e�����בw�ۏؕ��@�Q��
    HWFMKHWS As String * 1         ' �i�v�e�����בw�ۏؕ��@�Q��

    HWFONMIN As Double             ' �i�v�e�_�f�Z�x����
    HWFONMAX As Double             ' �i�v�e�_�f�Z�x���
    HWFONSPH As String * 1         ' �i�v�e�_�f�Z�x����ʒu�Q��
    HWFONSPT As String * 1         ' �i�v�e�_�f�Z�x����ʒu�Q�_
    HWFONSPI As String * 1         ' �i�v�e�_�f�Z�x����ʒu�Q��
    HWFONHWT As String * 1         ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
    HWFONHWS As String * 1         ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
    HWFONMCL As String * 1         ' �i�v�e�_�f�Z�x�ʓ��v�Z 2001/11/08 S.Sano
    HWFONMBP As Double             ' �i�v�e�_�f�Z�x�ʓ����z
    HWFONAMN As Double             ' �i�v�e�_�f�Z�x���ω���
    HWFONAMX As Double             ' �i�v�e�_�f�Z�x���Ϗ��

    HWFOS1MN As Double             ' �i�v�e�_�f�͏o�P����
    HWFOS1MX As Double             ' �i�v�e�_�f�͏o�P���
    HWFOS1SH As String * 1         ' �i�v�e�_�f�͏o�P����ʒu�Q��
    HWFOS1ST As String * 1         ' �i�v�e�_�f�͏o�P����ʒu�Q�_
    HWFOS1SI As String * 1         ' �i�v�e�_�f�͏o�P����ʒu�Q��
    HWFOS1HT As String * 1         ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
    HWFOS1HS As String * 1         ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
    HWFOS2SH As String * 1         ' �i�v�e�_�f�͏o�Q����ʒu�Q��
    HWFOS2ST As String * 1         ' �i�v�e�_�f�͏o�Q����ʒu�Q�_
    HWFOS2SI As String * 1         ' �i�v�e�_�f�͏o�Q����ʒu�Q��
    HWFOS2MN As Double             ' �i�v�e�_�f�͏o�Q����
    HWFOS2MX As Double             ' �i�v�e�_�f�͏o�Q���
    HWFOS2HT As String * 1         ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
    HWFOS2HS As String * 1         ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
    HWFOS3MN As Double             ' �i�v�e�_�f�͏o�R����
    HWFOS3MX As Double             ' �i�v�e�_�f�͏o�R���
    HWFOS3SH As String * 1         ' �i�v�e�_�f�͏o�R����ʒu�Q��
    HWFOS3ST As String * 1         ' �i�v�e�_�f�͏o�R����ʒu�Q�_
    HWFOS3SI As String * 1         ' �i�v�e�_�f�͏o�R����ʒu�Q��
    HWFOS3HT As String * 1         ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��
    HWFOS3HS As String * 1         ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��

    HWFDSOMX As Double             ' �i�v�e�c�r�n�c���              '2003/11/17 SystemBrain Integer �� Double
    HWFDSOMN As Double             ' �i�v�e�c�r�n�c����              '2003/11/17 SystemBrain Integer �� Double
    HWFDSOAX As Integer            ' �i�v�e�c�r�n�c�̈���
    HWFDSOAN As Integer            ' �i�v�e�c�r�n�c�̈扺��
    HWFDSOHT As String * 1         ' �i�v�e�c�r�n�c�ۏؕ��@�Q��
    HWFDSOHS As String * 1         ' �i�v�e�c�r�n�c�ۏؕ��@�Q��

    HWFSPVMX As Double             ' �i�v�e�r�o�u�e�d���
    HWFSPVSH As String * 1         ' �i�v�e�r�o�u�e�d����ʒu�Q��
    HWFSPVST As String * 1         ' �i�v�e�r�o�u�e�d����ʒu�Q�_
    HWFSPVSI As String * 1         ' �i�v�e�r�o�u�e�d����ʒu�Q��
    HWFSPVHT As String * 1         ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
    HWFSPVHS As String * 1         ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
    HWFDLSPH As String * 1         ' �i�v�e�g�U������ʒu�Q��
    HWFDLSPT As String * 1         ' �i�v�e�g�U������ʒu�Q�_
    HWFDLSPI As String * 1         ' �i�v�e�g�U������ʒu�Q��
    HWFDLHWT As String * 1         ' �i�v�e�g�U���ۏؕ��@�Q��
    HWFDLHWS As String * 1         ' �i�v�e�g�U���ۏؕ��@�Q��
    HWFDLMIN As Integer            ' �i�v�e�g�U������
    HWFDLMAX As Integer            ' �i�v�e�g�U�����

    HWFOF1AX As Double             ' �i�v�e�n�r�e�P���Ϗ��
    HWFOF1MX As Double             ' �i�v�e�n�r�e�P���
    HWFOF1SH As String * 1         ' �i�v�e�n�r�e�P����ʒu�Q��
    HWFOF1ST As String * 1         ' �i�v�e�n�r�e�P����ʒu�Q�_
    HWFOF1SR As String * 1         ' �i�v�e�n�r�e�P����ʒu�Q��
    HWFOF1HT As String * 1         ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    HWFOF1HS As String * 1         ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    HWFOF2AX As Double             ' �i�v�e�n�r�e�Q���Ϗ��
    HWFOF2MX As Double             ' �i�v�e�n�r�e�Q���
    HWFOF2SH As String * 1         ' �i�v�e�n�r�e�Q����ʒu�Q��
    HWFOF2ST As String * 1         ' �i�v�e�n�r�e�Q����ʒu�Q�_
    HWFOF2SR As String * 1         ' �i�v�e�n�r�e�Q����ʒu�Q��
    HWFOF2HT As String * 1         ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    HWFOF2HS As String * 1         ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    HWFOF3AX As Double             ' �i�v�e�n�r�e�R���Ϗ��
    HWFOF3MX As Double             ' �i�v�e�n�r�e�R���
    HWFOF3SH As String * 1         ' �i�v�e�n�r�e�R����ʒu�Q��
    HWFOF3ST As String * 1         ' �i�v�e�n�r�e�R����ʒu�Q�_
    HWFOF3SR As String * 1         ' �i�v�e�n�r�e�R����ʒu�Q��
    HWFOF3HT As String * 1         ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    HWFOF3HS As String * 1         ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    HWFOF4AX As Double             ' �i�v�e�n�r�e�S���Ϗ��
    HWFOF4MX As Double             ' �i�v�e�n�r�e�S���
    HWFOF4SH As String * 1         ' �i�v�e�n�r�e�S����ʒu�Q��
    HWFOF4ST As String * 1         ' �i�v�e�n�r�e�S����ʒu�Q�_
    HWFOF4SR As String * 1         ' �i�v�e�n�r�e�S����ʒu�Q��
    HWFOF4HT As String * 1         ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    HWFOF4HS As String * 1         ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    HWFOSF1PTK As String * 1       ' �i�v�e�n�r�e�P�p�^���敪�@��2003/05/14 ooba
    HWFOSF2PTK As String * 1       ' �i�v�e�n�r�e�Q�p�^���敪
    HWFOSF3PTK As String * 1       ' �i�v�e�n�r�e�R�p�^���敪
    HWFOSF4PTK As String * 1       ' �i�v�e�n�r�e�S�p�^���敪�@��2003/05/14 ooba

    HWFBM1AN As Double             ' �i�v�e�a�l�c�P���ω���
    HWFBM1AX As Double             ' �i�v�e�a�l�c�P���Ϗ��
    HWFBM1SH As String * 1         ' �i�v�e�a�l�c�P����ʒu�Q��
    HWFBM1ST As String * 1         ' �i�v�e�a�l�c�P����ʒu�Q�_
    HWFBM1SR As String * 1         ' �i�v�e�a�l�c�P����ʒu�Q��
    HWFBM1HT As String * 1         ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    HWFBM1HS As String * 1         ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    HWFBM2AN As Double             ' �i�v�e�a�l�c�Q���ω���
    HWFBM2AX As Double             ' �i�v�e�a�l�c�Q���Ϗ��
    HWFBM2SH As String * 1         ' �i�v�e�a�l�c�Q����ʒu�Q��
    HWFBM2ST As String * 1         ' �i�v�e�a�l�c�Q����ʒu�Q�_
    HWFBM2SR As String * 1         ' �i�v�e�a�l�c�Q����ʒu�Q��
    HWFBM2HT As String * 1         ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    HWFBM2HS As String * 1         ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    HWFBM3AN As Double             ' �i�v�e�a�l�c�R���ω���
    HWFBM3AX As Double             ' �i�v�e�a�l�c�R���Ϗ��
    HWFBM3SH As String * 1         ' �i�v�e�a�l�c�R����ʒu�Q��
    HWFBM3ST As String * 1         ' �i�v�e�a�l�c�R����ʒu�Q�_
    HWFBM3SR As String * 1         ' �i�v�e�a�l�c�R����ʒu�Q��
    HWFBM3HT As String * 1         ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    HWFBM3HS As String * 1         ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    HWFBM1MBP As Double            ' �i�v�e�a�l�c�P�ʓ����z�@��2003/05/14 ooba
    HWFBM2MBP As Double            ' �i�v�e�a�l�c�Q�ʓ����z
    HWFBM3MBP As Double            ' �i�v�e�a�l�c�R�ʓ����z
    HWFBM1MCL As String * 2        ' �i�v�e�a�l�c�P�ʓ��v�Z
    HWFBM2MCL As String * 2        ' �i�v�e�a�l�c�Q�ʓ��v�Z
    HWFBM3MCL As String * 2        ' �i�v�e�a�l�c�R�ʓ��v�Z�@��2003/05/14 ooba

    HWFOS1NS As String * 2         ' �i�v�e�_�f�͏o�P�M�����@
    HWFOS2NS As String * 2         ' �i�v�e�_�f�͏o�Q�M�����@
    HWFOS3NS As String * 2         ' �i�v�e�_�f�͏o�R�M�����@
    HWFOF1NS As String * 2         ' �i�v�e�n�r�e�P�M�����@
    HWFOF2NS As String * 2         ' �i�v�e�n�r�e�Q�M�����@
    HWFOF3NS As String * 2         ' �i�v�e�n�r�e�R�M�����@
    HWFOF4NS As String * 2         ' �i�v�e�n�r�e�S�M�����@
    HWFBM1NS As String * 2         ' �i�v�e�a�l�c�P�M�����@
    HWFBM2NS As String * 2         ' �i�v�e�a�l�c�Q�M�����@
    HWFBM3NS As String * 2         ' �i�v�e�a�l�c�R�M�����@

    HWFANTIM As Integer            ' �i�v�e�`�m����
    HWFANTNP As Integer            ' �i�v�e�`�m���x

    HWFOF1ET As Integer            ' �i�v�e�n�r�e�P�I���d�s��
    HWFOF2ET As Integer            ' �i�v�e�n�r�e�Q�I���d�s��
    HWFOF3ET As Integer            ' �i�v�e�n�r�e�R�I���d�s��
    HWFOF4ET As Integer            ' �i�v�e�n�r�e�S�I���d�s��
    HWFBM1ET As Integer            ' �i�v�e�a�l�c�P�I���d�s��
    HWFBM2ET As Integer            ' �i�v�e�a�l�c�Q�I���d�s��
    HWFBM3ET As Integer            ' �i�v�e�a�l�c�R�I���d�s��

    HWFOF1SZ As String * 1         ' �i�v�e�n�r�e�P�������
    HWFOF2SZ As String * 1         ' �i�v�e�n�r�e�Q�������
    HWFOF3SZ As String * 1         ' �i�v�e�n�r�e�R�������
    HWFOF4SZ As String * 1         ' �i�v�e�n�r�e�S�������
    HWFBM1SZ As String * 1         ' �i�v�e�a�l�c�P�������
    HWFBM2SZ As String * 1         ' �i�v�e�a�l�c�Q�������
    HWFBM3SZ As String * 1         ' �i�v�e�a�l�c�R�������

    BLOCKID() As String * 12       ' �u���b�NID
End Type

'SXL�Ǘ��X�V�p�i���ݍH���A�ŏI�ʉߍH���j
Public Type type_DBDRV_scmzc_fcmlc001c_UpdSXL1
    CRYNUM As String * 12          ' �����ԍ�
    INGOTPOS As Integer            ' �������J�n�ʒu
    NOWPROC As String * 5          ' ���ݍH��
    LASTPASS As String * 5         ' �ŏI�ʉߍH��
End Type

'SXL�Ǘ��X�V�p�i�폜�敪�A�ŏI��ԋ敪�j
Public Type type_DBDRV_scmzc_fcmlc001c_UpdSXL2
    CRYNUM As String * 12          ' �����ԍ�
    INGOTPOS As Integer            ' �������J�n�ʒu
    DELCLS As String * 1           ' �폜�敪
    LSTATCLS As String * 1         ' �ŏI��ԋ敪
End Type

'WF�T���v���Ǘ��X�V�p
Public Type type_DBDRV_scmzc_fcmlc001c_UpdWfCrySmp
    CRYNUM As String * 12          ' �����ԍ�
    INGOTPOS As Integer            ' �������ʒu
    SMPKBN As String * 1           ' �T���v���敪
End Type


' �Ĕ����w��
'���͗p
Type type_DBDRV_scmzc_fcmlc001d_In
    CRYNUM As String * 12          '�����ԍ�
    HIN As tFullHinban             '�i��
    LENGHT As Integer
End Type

'WF�d�l�擾�p
Public Type type_DBDRV_scmzc_fcmlc001d_WfSiyou
    HWFRMIN As Double              ' �i�v�e���R����
    HWFRMAX As Double              ' �i�v�e���R���
    HWFRHWYS As String * 1         ' �i�v�e���R�ۏؕ��@�Q��(Rs)
    HWFONHWS As String * 1         ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��(Oi)
    HWFBM1HS As String * 1         ' �i�v�e�a�l�c�P�ۏؕ��@�Q��(B1)
    HWFBM2HS As String * 1         ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��(B2)
    HWFBM3HS As String * 1         ' �i�v�e�a�l�c�R�ۏؕ��@�Q��(B3)
    HWFOF1HS As String * 1         ' �i�v�e�n�r�e�P�ۏؕ��@�Q��(L1)
    HWFOF2HS As String * 1         ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��(L2)
    HWFOF3HS As String * 1         ' �i�v�e�n�r�e�R�ۏؕ��@�Q��(L3)
    HWFOF4HS As String * 1         ' �i�v�e�n�r�e�S�ۏؕ��@�Q��(L4)
    HWFDSOHS As String * 1         ' �i�v�e�c�r�n�c�ۏؕ��@�Q��(DS)
    HWFMKHWS As String * 1         ' �i�v�e�����בw�ۏؕ��@�Q��(DZ)
    HWFSPVHS As String * 1         ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��(SP)
    HWFDLHWS As String * 1         ' �i�v�e�g�U���ۏؕ��@�Q��(KL)�@06/06/08 ooba
    HWFNRHS  As String * 1         ' �i�v�e�r�o�u�m�q�ۏؕ��@�Q��(NR)�@06/06/08 ooba
    HWFOS1HS As String * 1         ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��(D1)
    HWFOS2HS As String * 1         ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��(D2)
    HWFOS3HS As String * 1         ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��(D3)
    HWFZOHWS As String * 1         ' �i�v�e�c���_�f�ۏؕ��@�Q��(AO)    ''�ǉ��@03/12/15 ooba
    HWFDENHS As String * 1         ' �i�v�e�c�����ۏؕ��@�Q��(GD)      '�ǉ��@05/02/17 ooba START ====>
    HWFDVDHS As String * 1         ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��(GD)
    HWFLDLHS As String * 1         ' �i�v�e�k�^�c�k�ۏؕ��@�Q��(GD)    '�ǉ��@05/02/17 ooba END ======>
    HWFOT1   As String * 1         ' 03/05/26
    HWFOT2   As String * 1         ' 03/05/26
    KEIKAKUL As Integer            ' �v�撷
    HWFMAI1   As String * 1        ' 04/07/16
    HWFMAI2   As String * 1        ' 04/07/16
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    HEPOF1HS As String * 1         ' �����L��(OSF1E)
    HEPOF2HS As String * 1         ' �����L��(OSF2E)
    HEPOF3HS As String * 1         ' �����L��(OSF3E)
    HEPBM1HS As String * 1         ' �����L��(BMD1E)
    HEPBM2HS As String * 1         ' �����L��(BMD2E)
    HEPBM3HS As String * 1         ' �����L��(BMD3E)
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
' ��10/01/06 Add SIRD�Ή� Y.Hitomi
    HWFSIRDHS As String * 1         ' �����L��(SIRD)
' ��10/01/06 Add SIRD�Ή� Y.Hitomi
    CHUTAN   As Integer            ' ���Ԕ����P��(��)
    CHUKYO   As Integer            ' ���Ԕ������e�l
    CHUFLG   As String             ' ���Ԕ����t���O
End Type

'WF�T���v���Ǘ��iTOP�ATAIL �Q���R�[�h�j
Public Type type_DBDRV_scmzc_fcmlc001d_WfSmp
    INGOTPOS As Integer            ' �������ʒu
    SMPLID As String * 16          ' �T���v��ID
    hinban As String * 8           ' �i��
    REVNUM As Integer              ' ���i�ԍ������ԍ�
    factory As String * 1          ' �H��
    opecond As String * 1          ' ���Ə���
    WFINDRS As String * 1          ' ���FLG�iRs)
    WFINDOI As String * 1          ' ���FLG�iOi)
    WFINDB1 As String * 1          ' ���FLG�iB1)
    WFINDB2 As String * 1          ' ���FLG�iB2�j
    WFINDB3 As String * 1          ' ���FLG�iB3)
    WFINDL1 As String * 1          ' ���FLG�iL1)
    WFINDL2 As String * 1          ' ���FLG�iL2)
    WFINDL3 As String * 1          ' ���FLG�iL3)
    WFINDL4 As String * 1          ' ���FLG�iL4)
    WFINDDS As String * 1          ' ���FLG�iDS)
    WFINDDZ As String * 1          ' ���FLG�iDZ)
    WFINDSP As String * 1          ' ���FLG�iSP)
    WFINDDO1 As String * 1         ' ���FLG�iDO1)
    WFINDDO2 As String * 1         ' ���FLG�iDO2)
    WFINDDO3 As String * 1         ' ���FLG�iDO3)
    WFINDOTHER1 As String * 1      ' �����L��(OT2) ''Add.03/05/20 �㓡
    WFINDOTHER2 As String * 1      ' �����L��(OT1) ''Add.03/05/20
    WFINDAOI As String * 1         ' ���FLG (AOi)     '�c���_�f�ǉ��@03/12/15 ooba
    WFINDGD As String * 1          ' ���FLG (GD)      'GD�ǉ��@05/02/17 ooba
    WFHSGD As String * 1           ' �ۏ�FLG (GD)      'GD�ǉ��@05/02/17 ooba
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    EPINDB1CW As String * 1        '���FLG(BMD1)
    EPINDB2CW As String * 1        '���FLG(BMD2)
    EPINDB3CW As String * 1        '���FLG(BMD3)
    EPINDL1CW As String * 1        '���FLG(OSF1)
    EPINDL2CW As String * 1        '���FLG(OSF2)
    EPINDL3CW As String * 1        '���FLG(OSF3)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
End Type

Public Type typ_WfSampleGr
    BLOCKID As String
    blockp As Integer
    WFSMP As typ_XSDCW
    HINUP As tFullHinban           '��i��
    HINDN As tFullHinban           '���i��
    ERRDNFLG As Boolean            '���i�ԃG���[�t���O
End Type

' �����E�F�n�[���
Public Type typ_LackWaf
    BLOCKID As String * 12         '�u���b�NID
    WAFERNO As Integer             '�E�F�n�[�A��
    TOP_POS As Integer             '�E�F�n�[�J�n�ʒu
    TAIL_POS As Integer            '�E�F�n�[�I���ʒu
End Type
'���R
Public Type NoTest_RES
    HWFRHWYS As String * 1         '�iWF���R�ۏؕ��@�Q��
End Type
'�_�f�Z�x
Public Type NoTest_OI
    HWFONHWS As String * 1         '�iWF�_�f�Z�x�ۏؕ��@�Q��
End Type
'BMDx
Public Type NoTest_BMD
    HWFBMxHS As String * 1         '�iWFBMDx�ۏؕ��@�Q��
    HWFBMxET As Integer            '�iWFBMD1�I��ET��
    HWFBMxNS As String * 2         '�iWFBMD1�M�����@
    HWFBMxSZ As String * 1         '�iWFBMD1�������
    HWFBMxSH As String * 1         '�iWFBMD1����ʒu_��
    HWFBMxST As String * 1         '�iWFBMD1����ʒu_�_
    HWFBMxSR As String * 1         '�iWFBMD1����ʒu_��
End Type
'OSFx
Public Type NoTest_OSF
    HWFOFxHS As String * 1         '�iWFOSFx�ۏؕ��@�Q��
    HWFOFxET As Integer            '�iWFOSF1�I��ET��
    HWFOFxNS As String * 2         '�iWFOSF1�M�����@
    HWFOFxSZ As String * 1         '�iWFOSF1�������
    HWFOFxSH As String * 1         '�iWFOSF1����ʒu_��
    HWFOFxST As String * 1         '�iWFOSF1����ʒu_�_
    HWFOFxSR As String * 1         '�iWFOSF1����ʒu_��
End Type
'DSOD
Public Type NoTest_DSOD
    HWFDSOHS As String * 1         '�iWFDSOD�ۏؕ��@�Q��
    HWFDSOKE As String * 1         '�iWFDSOD����
End Type
'DZ
Public Type NoTest_DZ
    HWFMKHWS As String * 1         '�iWF�����בw�ۏؕ��@�Q��
    HWFMKSZY As String * 1         '�iWF�����בw�������
    HWFMKSPH As String * 1         '�iWF�����בw����ʒu�Q��
    HWFMKSPT As String * 1         '�iWF�����בw����ʒu�Q�_
    HWFMKSPR As String * 1         '�iWF�����בw����ʒu�Q��
End Type
'SPVFE
Public Type NoTest_SPVFE
    HWFSPVHS As String * 1         '�iWFSPVFE�ۏؕ��@�Q��
    HWFSPVSH As String * 1         '�iWFSPVFE����ʒu�Q��
    HWFSPVST As String * 1         '�iWFSPVFE����ʒu�Q�_
    HWFSPVSI As String * 1         '�iWFSPVFE����ʒu�Q��
End Type
'�g�U��
Public Type NoTest_SPV
    HWFDLHWS As String * 1         '�iWF�g�U���ۏؕ��@�Q��
    HWFDLSPH As String * 1         '�iWF�g�U������ʒu�Q��
    HWFDLSPT As String * 1         '�iWF�g�U������ʒu�Q�_
    HWFDLSPI As String * 1         '�iWF�g�U������ʒu�Q��
End Type
'��Oix
Public Type NoTest_DOI
    HWFOSxHS As String * 1         '�iWF�_�f�͏ox�ۏؕ��@�Q��
    HWFOSxNS As String * 2         '�iWF�_�f�͏o1�M�����@
    HWFOSxSH As String * 1         '�iWF�_�f�͏o1����ʒu�Q��
    HWFOSxST As String * 1         '�iWF�_�f�͏o1����ʒu�Q�_
    HWFOSxSI As String * 1         '�iWF�_�f�͏o1����ʒu�Q��
End Type

Public Type NoTest_Info
    Res As NoTest_RES
    Oi As NoTest_OI
    BMD(2) As NoTest_BMD
    OSF(3) As NoTest_OSF
    Dsod As NoTest_DSOD
    DZ As NoTest_DZ
    SpvFe As NoTest_SPVFE
    Spv As NoTest_SPV
    Doi(2) As NoTest_DOI
End Type

' WF�T���v���d�l(*�͖��`�F�b�N�̃p�����[�^)
Public Type typ_SpWFSamp
    HIN As tFullHinban             ' �i��

    HWFRHWYS As String * 1         ' �������@(Rs)
    HWFRSPOH As String * 1         ' ������@(Rs)*
    HWFRSPOT As String * 1         ' ����_��(Rs) -> Heavy
    HWFRSPOI As String * 1         ' ����ʒu(Rs)*

    HWFONHWS As String * 1         ' �������@(Oi)
    HWFONKWY As String * 2         ' �������@(Oi)
    HWFONSPH As String * 1         ' ������@(Oi)
    HWFONSPT As String * 1         ' ����_��(Oi) -> Heavy
    HWFONSPI As String * 1         ' ����ʒu(Oi)

    HWFBM1HS As String * 1         ' �������@(B1)
    HWFBM1SH As String * 1         ' ������@(B1)
    HWFBM1ST As String * 1         ' ����_��(B1)
    HWFBM1SR As String * 1         ' ���O�̈�(B1)
    HWFBM1NS As String * 2         ' �M�����@(B1)
    HWFBM1SZ As String * 1         ' �������(B1)
    HWFBM1ET As Integer            ' �I���G�b�`(B1)

    HWFBM2HS As String * 1         ' �������@(B2)
    HWFBM2SH As String * 1         ' ������@(B2)
    HWFBM2ST As String * 1         ' ����_��(B2)
    HWFBM2SR As String * 1         ' ���O�̈�(B2)
    HWFBM2NS As String * 2         ' �M�����@(B2)
    HWFBM2SZ As String * 1         ' �������(B2)
    HWFBM2ET As Integer            ' �I���G�b�`(B2)

    HWFBM3HS As String * 1         ' �������@(B3)
    HWFBM3SH As String * 1         ' ������@(B3)
    HWFBM3ST As String * 1         ' ����_��(B3)
    HWFBM3SR As String * 1         ' ���O�̈�(B3)
    HWFBM3NS As String * 2         ' �M�����@(B3)
    HWFBM3SZ As String * 1         ' �������(B3)
    HWFBM3ET As Integer            ' �I���G�b�`(B3)

    HWFOF1HS As String * 1         ' �������@(L1)
    HWFOF1SH As String * 1         ' ������@(L1)
    HWFOF1ST As String * 1         ' ����_��(L1)
    HWFOF1SR As String * 1         ' ���O�̈�(L1)
    HWFOF1NS As String * 2         ' �M�����@(L1)
    HWFOF1SZ As String * 1         ' �������(L1)
    HWFOF1ET As Integer            ' �I���G�b�`(L1)

    HWFOF2HS As String * 1         ' �������@(L2)
    HWFOF2SH As String * 1         ' ������@(L2)
    HWFOF2ST As String * 1         ' ����_��(L2)
    HWFOF2SR As String * 1         ' ���O�̈�(L2)
    HWFOF2NS As String * 2         ' �M�����@(L2)
    HWFOF2SZ As String * 1         ' �������(L2)
    HWFOF2ET As Integer            ' �I���G�b�`(L2)

    HWFOF3HS As String * 1         ' �������@(L3)
    HWFOF3SH As String * 1         ' ������@(L3)
    HWFOF3ST As String * 1         ' ����_��(L3)
    HWFOF3SR As String * 1         ' ���O�̈�(L3)
    HWFOF3NS As String * 2         ' �M�����@(L3)
    HWFOF3SZ As String * 1         ' �������(L3)
    HWFOF3ET As Integer            ' �I���G�b�`(L3)

'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
'''    HWFOF4HS As String * 1         ' �������@(L4)
'''    HWFOF4SH As String * 1         ' ������@(L4)
'''    HWFOF4ST As String * 1         ' ����_��(L4)
'''    HWFOF4SR As String * 1         ' ���O�̈�(L4)
'''    HWFOF4NS As String * 2         ' �M�����@(L4)
'''    HWFOF4SZ As String * 1         ' �������(L4)
'''    HWFOF4ET As Integer            ' �I���G�b�`(L4)
    
    HWFSIRDMX As Integer       '����]�ʏ��(SIRD)
    HWFSIRDSZ As String * 1    '����]�ʑ������(SIRD)
    HWFSIRDHT As String * 1    '����]�ʕۏؕ��@�Q��(SIRD)
    HWFSIRDHS As String * 1    '����]�ʕۏؕ��@�Q��(SIRD)
    HWFSIRDKM As String * 1    '����]�ʌ����p�x�Q��(SIRD)
    HWFSIRDKH As String * 1    '����]�ʌ����p�x�Q��(SIRD)
    HWFSIRDKU As String * 1    '����]�ʌ����p�x�Q�E(SIRD)
    HWFSIRDPS As String * 2    '����]��TB�ۏ؈ʒu(SIRD)
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)

    HWFDSOHS As String * 1         ' �������@(DS)

    HWFMKHWS As String * 1         ' �������@(DZ)
    HWFMKSPH As String * 1         ' ������@(DZ)
    HWFMKSPT As String * 1         ' ����_��(DZ)
    HWFMKSPR As String * 1         ' ���O�̈�(DZ)
    HWFMKNSW As String * 2         ' �M�����@(DZ)
    HWFMKSZY As String * 1         ' �������(DZ)
    HWFMKCET As Integer            ' �I���G�b�`(DZ)

    HWFSPVHS As String * 1         ' �������@(SP/Fe�Z�x)
    HWFSPVSH As String * 1         ' ������@(SP/Fe�Z�x)*
    HWFSPVST As String * 1         ' ����_��(SP/Fe�Z�x)*
    HWFSPVSI As String * 1         ' ����ʒu(SP/Fe�Z�x)*
    HWFDLHWS As String * 1         ' �������@(SP/�g�U��)
    HWFDLSPH As String * 1         ' ������@(SP/�g�U��)*
    HWFDLSPT As String * 1         ' ����_��(SP/�g�U��)*
    HWFDLSPI As String * 1         ' ����ʒu(SP/�g�U��)*
    HWFNRHS  As String * 1         ' �������@(SP/Nr�Z�x)               06/06/08 ooba START ======>
    HWFNRSH  As String * 1         ' ������@(SP/Nr�Z�x)*
    HWFNRST  As String * 1         ' ����_��(SP/Nr�Z�x)*
    HWFNRSI  As String * 1         ' ����ʒu(SP/Nr�Z�x)*
    HWFSPVPUG   As String * 10     ' PUA��(SP/Fe�Z�x)*
    HWFSPVPUR   As String * 10     ' PUA��(SP/Fe�Z�x)*
    HWFSPVSTD   As String * 10     ' �W���΍�(SP/Fe�Z�x)*
    HWFDLPUG    As String * 10     ' PUA��(SP/�g�U��)*
    HWFDLPUR    As String * 10     ' PUA��(SP/�g�U��)*
    HWFNRPUG    As String * 10     ' PUA��(SP/Nr�Z�x)*
    HWFNRPUR    As String * 10     ' PUA��(SP/Nr�Z�x)*
    HWFNRSTD    As String * 10     ' �W���΍�(SP/Nr�Z�x)*      06/06/08 ooba END ========>

    HWFOS1HS As String * 1         ' �������@(D1)
    HWFOS1SH As String * 1         ' ������@(D1)*
    HWFOS1ST As String * 1         ' ����_��(D1)*
    HWFOS1SI As String * 1         ' ����ʒu(D1)*
    HWFOS1NS As String * 2         ' �M�����@(D1)

    HWFOS2HS As String * 1         ' �������@(D2)
    HWFOS2SH As String * 1         ' ������@(D2)*
    HWFOS2ST As String * 1         ' ����_��(D2)*
    HWFOS2SI As String * 1         ' ����ʒu(D2)*
    HWFOS2NS As String * 2         ' �M�����@(D2)

    HWFOS3HS As String * 1         ' �������@(D3)
    HWFOS3SH As String * 1         ' ������@(D3)*
    HWFOS3ST As String * 1         ' ����_��(D3)*
    HWFOS3SI As String * 1         ' ����ʒu(D3)*
    HWFOS3NS As String * 2         ' �M�����@(D3)

    HWFZOHWS As String * 1         ' �������@(AO)  ''�ǉ� 03/12/15 ooba START ======>
    HWFZOSPH As String * 1         ' ������@(AO)*
    HWFZOSPT As String * 1         ' ����_��(AO)*
    HWFZOSPI As String * 1         ' ����ʒu(AO)*
    HWFZONSW As String * 2         ' �M�����@(AO)  ''�ǉ� 03/12/15 ooba END ========>

    HWFDENHS As String * 1         ' �������@(GD/DEN)  '�ǉ��@05/02/18 ooba START ====>
    HWFLDLHS As String * 1         ' �������@(GD/LDL)
    HWFDVDHS As String * 1         ' �������@(GD/DVD2) '�ǉ��@05/02/18 ooba END ======>
    HWFGDSPH As String * 1         ' ������@(GD)�@    '05/10/25 ooba
    HWFGDSPT As String * 1         ' ����_��(GD)�@    '05/10/25 ooba
    HWFGDZAR As String * 1         ' ���O�̈�(GD)�@    '05/10/25 ooba

    HWFRKHNN As String * 1         ' �����p�x_��(Rs)   '�ǉ��@04/04/12 ooba START ====>
    HWFONKHN As String * 1         ' �����p�x_��(Oi)
    HWFOF1KN As String * 1         ' �����p�x_��(L1)
    HWFOF2KN As String * 1         ' �����p�x_��(L2)
    HWFOF3KN As String * 1         ' �����p�x_��(L3)
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
''    HWFOF4KN As String * 1         ' �����p�x_��(L4)
    HWFSIRDKN As String * 1  ' �����p�x_��(SIRD)
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
    HWFBM1KN As String * 1         ' �����p�x_��(B1)
    HWFBM2KN As String * 1         ' �����p�x_��(B2)
    HWFBM3KN As String * 1         ' �����p�x_��(B3)
    HWFOS1KN As String * 1         ' �����p�x_��(D1)
    HWFOS2KN As String * 1         ' �����p�x_��(D2)
    HWFOS3KN As String * 1         ' �����p�x_��(D3)
    HWFDSOKN As String * 1         ' �����p�x_��(DS)
    HWFMKKHN As String * 1         ' �����p�x_��(DZ)
    HWFSPVKN As String * 1         ' �����p�x_��(SP/Fe�Z�x)
    HWFDLKHN As String * 1         ' �����p�x_��(SP/�g�U��)
    HWFZOKHN As String * 1         ' �����p�x_��(AO)   '�ǉ��@04/04/12 ooba END ======>
    HWFGDKHN As String * 1         ' �����p�x_��(GD)�@05/02/18 ooba
    HWFNRKN  As String * 1         ' �����p�x_��(SP/Nr�Z�x)  06/06/08 ooba

    HWFIGKBN As String * 1         ' IG�敪
    HWFANTNP As Integer            ' DK�A�j�[������(���x)
    HWFANTIM As Integer            ' DK�A�j�[������(����)
    HWFANGZY As String * 1         ' DK�A�j�[������(�K�X)�@04/07/29 ooba
    HWOTHER1 As String * 1         ' �����L��(OT2) ''Add.03/05/20 �㓡
    HWOTHER2 As String * 1         ' �����L��(OT1) ''Add.03/05/20
    HWOTHER1MAI As String * 1      ' 04/07/16
    HWOTHER2MAI As String * 1      ' 04/07/16

''Upd Start (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�
    HWFGDLINE   As String * 3      '�iWFGDײݐ�(TBCME036)

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    HEPOF1NS As String * 2         ' �i�M�����@(OSF1E)
    HEPOF1SZ As String * 1         ' �i�������(OSF1E)
    HEPOF1ET As Integer            ' �i�I��ET��(OSF1E)
    HEPOF1HS As String * 1         ' �i�ۏؕ��@_��(OSF1E)
    HEPOF1SH As String * 1         ' �i����ʒu_��(OSF1E)
    HEPOF1ST As String * 1         ' �i����ʒu_�_(OSF1E)
    HEPOF1SR As String * 1         ' �i����ʒu_��(OSF1E)
    HEPOF1KN As String * 1         ' �i�����p�x_��(OSF1E)
    HEPOF2NS As String * 2         ' �i�M�����@(OSF2E)
    HEPOF2SZ As String * 1         ' �i�������(OSF2E)
    HEPOF2ET As Integer            ' �i�I��ET��(OSF2E)
    HEPOF2HS As String * 1         ' �i�ۏؕ��@_��(OSF2E)
    HEPOF2SH As String * 1         ' �i����ʒu_��(OSF2E)
    HEPOF2ST As String * 1         ' �i����ʒu_�_(OSF2E)
    HEPOF2SR As String * 1         ' �i����ʒu_��(OSF2E)
    HEPOF2KN As String * 1         ' �i�����p�x_��(OSF2E)
    HEPOF3NS As String * 2         ' �i�M�����@(OSF3E)
    HEPOF3SZ As String * 1         ' �i�������(OSF3E)
    HEPOF3ET As Integer            ' �i�I��ET��(OSF3E)
    HEPOF3HS As String * 1         ' �i�ۏؕ��@_��(OSF3E)
    HEPOF3SH As String * 1         ' �i����ʒu_��(OSF3E)
    HEPOF3ST As String * 1         ' �i����ʒu_�_(OSF3E)
    HEPOF3SR As String * 1         ' �i����ʒu_��(OSF3E)
    HEPOF3KN As String * 1         ' �i�����p�x_��(OSF3E)
    HEPBM1NS As String * 2         ' �i�M�����@(BMD1E)
    HEPBM1SZ As String * 1         ' �i�������(BMD1E)
    HEPBM1ET As Integer            ' �i�I��ET��(BMD1E)
    HEPBM1HS As String * 1         ' �i�ۏؕ��@_��(BMD1E)
    HEPBM1SH As String * 1         ' �i����ʒu_��(BMD1E)
    HEPBM1ST As String * 1         ' �i����ʒu_�_(BMD1E)
    HEPBM1SR As String * 1         ' �i����ʒu_��(BMD1E)
    HEPBM1KN As String * 1         ' �i�����p�x_��(BMD1E)
    HEPBM2NS As String * 2         ' �i�M�����@(BMD2E)
    HEPBM2SZ As String * 1         ' �i�������(BMD2E)
    HEPBM2ET As Integer            ' �i�I��ET��(BMD2E)
    HEPBM2HS As String * 1         ' �i�ۏؕ��@_��(BMD2E)
    HEPBM2SH As String * 1         ' �i����ʒu_��(BMD2E)
    HEPBM2ST As String * 1         ' �i����ʒu_�_(BMD2E)
    HEPBM2SR As String * 1         ' �i����ʒu_��(BMD2E)
    HEPBM2KN As String * 1         ' �i�����p�x_��(BMD2E)
    HEPBM3NS As String * 2         ' �i�M�����@(BMD3E)
    HEPBM3SZ As String * 1         ' �i�������(BMD3E)
    HEPBM3ET As Integer            ' �i�I��ET��(BMD3E)
    HEPBM3HS As String * 1         ' �i�ۏؕ��@_��(BMD3E)
    HEPBM3SH As String * 1         ' �i����ʒu_��(BMD3E)
    HEPBM3ST As String * 1         ' �i����ʒu_�_(BMD3E)
    HEPBM3SR As String * 1         ' �i����ʒu_��(BMD3E)
    HEPBM3KN As String * 1         ' �i�����p�x_��(BMD3E)
    HEPACEN  As Double             ' �iE1�����S
    HEPANTNP As Integer            ' �iEPAN���x
    HEPANTIM As Integer            ' �iEPAN����
    HEPIGKBN As String * 1         ' �iEPIG�敪
    HEPANGZY As String * 1         ' �iEP����AN�K�X����
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga START   ---
    HWFGDSZY As String * 1         ' �i�v�e�f�c�������

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP As String * 1         ' DK���x
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
End Type

' WF�T���v���e�[�u��
Public Type typ_WFSample
    CRYINDRS As String * 1         ' ��������(Rs)
    CRYINDOI As String * 1         ' ��������(Oi)
    CRYINDB1 As String * 1         ' ��������(B1)
    CRYINDB2 As String * 1         ' ��������(B2�j
    CRYINDB3 As String * 1         ' ��������(B3)
    CRYINDL1 As String * 1         ' ��������(L1)
    CRYINDL2 As String * 1         ' ��������(L2)
    CRYINDL3 As String * 1         ' ��������(L3)
    CRYINDL4 As String * 1         ' ��������(L4)
    CRYINDDS As String * 1         ' ��������(DS)
    CRYINDDZ As String * 1         ' ��������(DZ)
    CRYINDSP As String * 1         ' ��������(SP)
    CRYINDD1 As String * 1         ' ��������(D1)
    CRYINDD2 As String * 1         ' ��������(D2)
    CRYINDD3 As String * 1         ' ��������(D3)
    CRYOTHER1 As String * 1        ' �����L��(OT2) ''Add.03/05/20 �㓡
    CRYOTHER2 As String * 1        ' �����L��(OT1) ''Add.03/05/20
    CRYINDAO As String * 1         ' ��������(AO)      ''�ǉ��@03/12/15 ooba
    CRYINDGD As String * 1         ' �����L��(GD)      '�ǉ� 05/01/18 ooba
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga START   ---
    CRYINDGD2 As String * 1        ' �����L��(GD��������p)
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga END     ---
    WFHSGD As String * 1           ' �ۏ�FLG(GD)       '�ǉ� 05/01/18 ooba
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    EPIINDL1 As String * 1         ' �����L��(OSF1E)
    EPIINDL2 As String * 1         ' �����L��(OSF2E)
    EPIINDL3 As String * 1         ' �����L��(OSF3E)
    EPIINDB1 As String * 1         ' �����L��(BMD1E)
    EPIINDB2 As String * 1         ' �����L��(BMD2E)
    EPIINDB3 As String * 1         ' �����L��(BMD3E)
End Type

'2002/09/11 ADD hitec)N.MATSUMOTO Start
Public strBlockID()    As String
Public Const PROCD_WFC_SAINUKISI = "CW760"  'WF�Z���^�[�Ĕ���
Public Const PROCD_SXL_MAP = "TX860"        '�V���O���}�b�v
Public Const WF_HANTEI_FORM As Integer = 1  '��ʂ̔���iWF�Z���^�[��������j
Public Const SAINUKISI_FORM As Integer = 2  '��ʂ̔���i�Ĕ����w���j


'2002/09/11 ADD hitec)N.MATSUMOTO  End


'=================================
'2003/02/28 ADD HITEC)okazaki start

Public Type type_DBDRV_Nukisi
    LOTID       As String * 12     ' �u���b�NID
    SXLID       As String * 13     ' SXLID
    MinMax      As Integer         ' 0:MIN 1:MAX
    BLOCKSEQ    As String * 3      ' �u���b�N���A��
    WFSTA       As String * 1      ' WF���
    hinban      As String * 8      ' �i��
    RTOP_POS    As Double          ' �_���u���b�N���ʒu
    RITOP_POS   As Double          ' �_���������ʒu
    SMPLEID     As String * 16     ' �����ʒu
    SHAFLAG     As String * 1      ' �T���v���t���O
    INDTM       As Date
    BASKETID    As String * 6
    SLOTNO      As Integer
    CURRWPCS    As Integer
    EXISTFLG    As String * 1
    TOP_POS     As Integer
    REJCAT      As String * 1
    TXID        As String * 6
    REGDATE     As Date
    SUMMITSENDFLAG As String * 1
    SENDFLAG    As String * 1
    SENDDATE    As Date
    HREJCODE    As String * 4
    UPDPROC     As String * 5
    UPDDATE     As Date
    REVNUM      As Integer
    factory     As String * 1
    opecond     As String * 1
    KANKBN      As String * 1
    NREJCODE    As String * 6
    SMPLEFLG    As String
End Type
Public Type type_DBDRV_LOTSXL
    LOTID       As String * 12     ' �u���b�NID
    SXLID       As String * 13     ' SXLID
End Type
'2003/02/28 Add HITEC)okazaki end

'2003/02/28 Hitec)okazaki add start
Public tExamine() As type_DBDRV_Nukisi  '��ʕ\����
                                        '�E�F�n�[�Z���^�[���ɏ��e�[�u��
Public tKeturaku() As typ_TBCMY012

Public tSXLID() As type_DBDRV_LOTSXL
'2003/02/28 Hitec)okazaki add end

'add  2003/03/15 hitec)matsumoto ---------------
Public bWfmapView As Boolean

Public CngSmpID_UD()    As String       ' UD��TB�ύX�p�@2004/01/29 ooba
Public bMotoGDcpyFlg(2) As Boolean      ' �����s�̌���GD���p���L���@05/08/04 ooba

'add 2003/03/25 hitec)matsumoto ��۰��ي֐��Ƃ��Ďg�������̂ŁAf_cmbc039_3.frm���ړ�----------------
Public SIngotP As Integer              ' �C���S�b�g�㑤�ʒu
Public EIngotP As Integer              ' �C���S�b�g�����ʒu
'add 2003/03/25 hitec)matsumoto ------------------------------

Public tblSXL As DBDRV_scmzc_fcmlc001b_SXL ' SXL�Ǘ��i�҂��ꗗ����j    'upd 2003/04/27 hitec)matsumoto f_cmbc039_3���ړ�

'WF�T���v������FLG�X�V�Ώ��������ʍ\���́@04/02/06 tuku
Public Type type_chkUP
    rs As String * 1               ' ��MFLG�iRs)
    Oi As String * 1               ' ��MFLG�iOi)
    B1 As String * 1               ' ��MFLG�iB1)
    B2 As String * 1               ' ��MFLG�iB2�j
    B3 As String * 1               ' ��MFLG�iB3)
    L1 As String * 1               ' ��MFLG�iL1)
    L2 As String * 1               ' ��MFLG�iL2)
    L3 As String * 1               ' ��MFLG�iL3)
    L4 As String * 1               ' ��MFLG�iL4)
    DS As String * 1               ' ��MFLG�iDS)
    DZ As String * 1               ' ��MFLG�iDZ)
    sp As String * 1               ' ��MFLG�iSP)
    DO1 As String * 1              ' ��MFLG�iDO1)
    DO2 As String * 1              ' ��MFLG�iDO2)
    DO3 As String * 1              ' ��MFLG�iDO3)
    OT1 As String * 1              ' ��MFLG (OT1)
    OT2 As String * 1              ' ��MFLG (OT2)
    AOI As String * 1              ' ��MFLG (AOi)
    GD As String * 1               ' ��MFLG (GD)   '05/02/04 ooba
    B1E As String * 1              ' ��MFLG�iB1E)
    B2E As String * 1              ' ��MFLG�iB2E�j
    B3E As String * 1              ' ��MFLG�iB3E)
    L1E As String * 1              ' ��MFLG�iL1E)
    L2E As String * 1              ' ��MFLG�iL2E)
    L3E As String * 1              ' ��MFLG�iL3E)
End Type

'**************************************************************************************
'*    �֐���        : KeturakuInfo
'*
'*    �����T�v      : 1.�����L���擾
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                            ,����
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL�Ǘ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'**************************************************************************************
Private Function KeturakuInfo(sxl() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim i           As Long
    Dim j           As Long
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim intSXLCnt   As Integer
    Dim sSXLID      As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_sSql.bas -- Function KeturakuInfo"

    KeturakuInfo = FUNCTION_RETURN_SUCCESS

    intSXLCnt = UBound(sxl)

#If True Then   'New Version  2002.1.24
    sSQL = "select distinct SXL.sSXLID "
    sSQL = sSQL & "from TBCME042 SXL, TBCME040 BLK, TBCMY012 REJ "
    sSQL = sSQL & "where"
    sSQL = sSQL & "  REJ.LOTID=BLK.BLOCKID"
    sSQL = sSQL & "  and SXL.CRYNUM=BLK.CRYNUM"
    sSQL = sSQL & "  and SXL.DELCLS<>'1'"
    sSQL = sSQL & "  and ("
    sSQL = sSQL & "    ("
    sSQL = sSQL & "      REJ.ALLSCRAP='Y'"
    sSQL = sSQL & "      and SXL.INGOTPOS<BLK.INGOTPOS+BLK.LENGTH"
    sSQL = sSQL & "      and SXL.INGOTPOS+SXL.LENGTH>BLK.INGOTPOS"
    sSQL = sSQL & "    ) or ("
    sSQL = sSQL & "      REJ.ALLSCRAP='N'"
    sSQL = sSQL & "      and REJ.REJCAT='A'"
    sSQL = sSQL & "      and (SXL.INGOTPOS < BLK.INGOTPOS + REJ.LENTO)"
    sSQL = sSQL & "      and (SXL.INGOTPOS + SXL.LENGTH > BLK.INGOTPOS + REJ.LENFROM)"
    sSQL = sSQL & "    ) or ("
    sSQL = sSQL & "      REJ.REJCAT='B'"
    sSQL = sSQL & "      and BLK.INGOTPOS + REJ.TOP_POS/10.0 between SXL.INGOTPOS and SXL.INGOTPOS + SXL.LENGTH"
    sSQL = sSQL & "    )"
    sSQL = sSQL & "  )"
#Else
    sSQL = "select "
    sSQL = sSQL & " distinct sSXLID "
    sSQL = sSQL & " from "
    sSQL = sSQL & " VECMW002 K, XSDCA A, TBCME040 B "
    sSQL = sSQL & " where "
    sSQL = sSQL & " A.CRYNUMCA = B.CRYNUM "
    sSQL = sSQL & " and B.BLOCKID = K.BLOCKID "
    sSQL = sSQL & " and ( "
    sSQL = sSQL & " ((B.INGOTPOS + K.TOP_POS) >= A.INPOSCA and (B.INGOTPOS + K.TOP_POS) < (A.INPOSCA + A.GNLCA)) "
    sSQL = sSQL & " or ((B.INGOTPOS + K.TAIL_POS) > A.INPOSCA and (B.INGOTPOS + K.TAIL_POS) < (A.INPOSCA + A.GNLCA)) "
    sSQL = sSQL & " or (A.INPOSCA >= (B.INGOTPOS + K.TOP_POS)  and A.INPOSCA < (B.INGOTPOS + K.TAIL_POS)) "
    sSQL = sSQL & " or ((A.INPOSCA + A.GNLCA) > (B.INGOTPOS + K.TOP_POS) and (A.INPOSCA + A.GNLCA) < (B.INGOTPOS + K.TAIL_POS)) "
    sSQL = sSQL & " and S.sSXLID in ("
    For i = 1 To intSXLCnt
        If i = intSXLCnt Then
            sSQL = sSQL & "'" & sxl(i).sSXLID & "' "
        Else
            sSQL = sSQL & "'" & sxl(i).sSXLID & "', "
        End If
    Next
    sSQL = sSQL & ") "
#End If
    Debug.Print sSQL
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    lngRecCnt = rs.RecordCount

    '������
    For i = 1 To intSXLCnt
        sxl(i).KETURAKU = False
    Next

    'sSql���ʂ�sSXLID�����������sSXLID
    For i = 1 To lngRecCnt
        sSXLID = rs("sSXLID")
        For j = 1 To intSXLCnt
            If sSXLID = sxl(j).CRYNUMCA Then
                sxl(j).KETURAKU = True
            End If
        Next
        rs.MoveNext
    Next
    rs.Close

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error sSql ======"
    Debug.Print sSQL
    KeturakuInfo = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************
'*    �֐���        : GetMaisu
'*
'*    �����T�v      : 1.WF�����擾
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                            ,����
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL�Ǘ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************
Private Function GetMaisu(sxl() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim i           As Long
    Dim j           As Long
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim intSXLCnt   As Integer
    Dim sSXLID      As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_sSql.bas -- Function getMaisu"

    GetMaisu = FUNCTION_RETURN_SUCCESS

    intSXLCnt = UBound(sxl)

    sSQL = sSQL & "SELECT sSXLIDCB,MAICB "
    sSQL = sSQL & "FROM XSDCB "
    sSQL = sSQL & "GROUP BY sSXLIDCB,MAICB"

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    lngRecCnt = rs.RecordCount

    '������
    For i = 1 To intSXLCnt
        sxl(i).MAICB = 0
    Next

    '�����i�[
    For i = 1 To lngRecCnt
        sSXLID = rs("sSXLIDCB")
        For j = 1 To intSXLCnt
            If sSXLID = sxl(j).SXLIDCA Then
                sxl(j).MAICB = rs("MAICB")
                Exit For
            End If
        Next
        rs.MoveNext
    Next
    rs.Close

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error sSql ======"
    Debug.Print sSQL
    GetMaisu = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************************
'*    �֐���        : GetsSXLIDINBlkid
'*
'*    �����T�v      : 1.SXL�̑S�u���b�N���Ƀ`�F�b�N
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                            ,����
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL�Ǘ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Private Function GetsSXLIDINBlkid(sxl() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim i           As Long
    Dim j           As Long
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim intSXLCnt   As Integer
    Dim sSXLID      As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_sSql.bas -- Function GetsSXLIDINBlkid"

    GetsSXLIDINBlkid = FUNCTION_RETURN_SUCCESS

    intSXLCnt = UBound(sxl)
    ReDim WFJudgExecOkFlag(intSXLCnt) As Boolean    'WF����������s�\�t���O

    '���ɑ҂��u���b�N���܂�SXL�̓O���[�\��
    sSQL = "select distinct SXL.sSXLID "
    sSQL = sSQL & "from TBCME042 SXL, TBCME040 BLK "
    sSQL = sSQL & "where SXL.DELCLS='0' and SXL.NOWPROC='CW750'"
    sSQL = sSQL & "  and BLK.CRYNUM=SXL.CRYNUM and BLK.INGOTPOS>=0"
    sSQL = sSQL & "  and SXL.INGOTPOS<BLK.INGOTPOS+BLK.LENGTH"
    sSQL = sSQL & "  and SXL.INGOTPOS+SXL.LENGTH>BLK.INGOTPOS"
    sSQL = sSQL & "  and not exists (select LOTID from TBCMY011 where LOTID=BLK.BLOCKID)"
    sSQL = sSQL & "  and not exists (select LOTID from TBCMY012 where LOTID=BLK.BLOCKID and ALLSCRAP='Y')"

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)
    lngRecCnt = rs.RecordCount

    For i = 1 To intSXLCnt
        WFJudgExecOkFlag(i) = True
    Next

    For i = 1 To lngRecCnt
        sSXLID = rs("sSXLID")
        For j = 1 To intSXLCnt
            If sSXLID = sxl(j).SXLIDCA Then
                WFJudgExecOkFlag(j) = False
                Exit For
            End If
        Next
        rs.MoveNext
    Next
    rs.Close
    Set rs = Nothing
    Debug.Print sSQL

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error sSql ======"
    Debug.Print sSQL
    GetsSXLIDINBlkid = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***************************************************************************************
'*    �֐���        : MeasRsltCheck
'*
'*    �����T�v      : 1.����]�����ʎ�M�m�F
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                            ,����
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL�Ǘ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************
Private Function MeasRsltCheck(sxl() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim intSXLCnt       As Integer
    Dim udtSokutei()    As typ_TBCMY013
    Dim intGDcnt        As Integer           'GD����ں��ސ�
''SPV9�_�Ή�
    Dim intSPVCnt       As Integer          'SPV���у��R�[�h��
    Dim c0              As Integer
    Dim c1              As Integer
    Dim c2              As Integer
    Dim blChangeFlag    As Boolean
    Dim blPassFlg       As Boolean
    Dim sSqlWhere       As String
#If SPEEDUP Then
    Dim sChkWF()        As String
    Dim i               As Integer
#End If
    Dim udtChkUp        As type_chkUP

    '�G���[�n���h���̐ݒ�
    'On Error GoTo proc_err
'    gErr.Push "s_cmbc039_sSql.bas -- Function MeasRsltCheck"

    MeasRsltCheck = FUNCTION_RETURN_SUCCESS

    '����]�����ʎ擾sSql�ύX
    intSXLCnt = UBound(sxl)
    If intSXLCnt = 0 Then GoTo proc_exit

    sSQL = sSQL & "select "
    sSQL = sSQL & "SXLIDCW, "
    sSQL = sSQL & "REPSMPLIDCW, "
    sSQL = sSQL & "decode(trim(RES_SPEC),'RES','1','0') RES, "
    sSQL = sSQL & "decode(trim(OI_SPEC),'OI','1','0') OI, "
    sSQL = sSQL & "decode(trim(BMD1_SPEC),'BMD1','1','0') BMD1, "
    sSQL = sSQL & "decode(trim(BMD2_SPEC),'BMD2','1','0') BMD2, "
    sSQL = sSQL & "decode(trim(BMD3_SPEC),'BMD3','1','0') BMD3, "
    sSQL = sSQL & "decode(trim(OSF1_SPEC),'OSF1','1','0') OSF1, "
    sSQL = sSQL & "decode(trim(OSF2_SPEC),'OSF2','1','0') OSF2, "
    sSQL = sSQL & "decode(trim(OSF3_SPEC),'OSF3','1','0') OSF3, "
'    sSql = sSql & "decode(trim(OSF4_SPEC),'OSF4','1','0') OSF4, " SIRD�Ή�
'    sSQL = sSQL & "decode(trim(SIRD_SPEC),'SIRD','1','0') SIRD, " 2010/05/19 REP Y.Hitomi
    sSQL = sSQL & "decode(trim(SIRD_SPEC),'TENI','1','0') SIRD, "
    sSQL = sSQL & "decode(trim(DSOD_SPEC),'DSOD','1','0') DSOD, "
    sSQL = sSQL & "decode(trim(DZ_SPEC),'DZ','1','0') DZ, "
    sSQL = sSQL & "decode(trim(DOI1_SPEC),'DOI1','1','0') DOI1, "
    sSQL = sSQL & "decode(trim(DOI2_SPEC),'DOI2','1','0') DOI2, "
    sSQL = sSQL & "decode(trim(DOI3_SPEC),'DOI3','1','0') DOI3, "
    sSQL = sSQL & "decode(trim(AOI_SPEC),'AOI','1','0') AOI, "
    sSQL = sSQL & "decode(trim(GD_SPEC),'GD','1','0') GD, "
    sSQL = sSQL & "decode(trim(SPV_SPEC),'SPV','1','0') SPV, "
    sSQL = sSQL & "decode(trim(IJO),'1','1','0') IJO "

    '�G�s��s�]���ǉ��Ή�
    sSQL = sSQL & ",decode(trim(BMD1E_SPEC),'BMD1','1','0') BMD1E, "
    sSQL = sSQL & "decode(trim(BMD2E_SPEC),'BMD2','1','0') BMD2E, "
    sSQL = sSQL & "decode(trim(BMD3E_SPEC),'BMD3','1','0') BMD3E, "
    sSQL = sSQL & "decode(trim(OSF1E_SPEC),'OSF1','1','0') OSF1E, "
    sSQL = sSQL & "decode(trim(OSF2E_SPEC),'OSF2','1','0') OSF2E, "
    sSQL = sSQL & "decode(trim(OSF3E_SPEC),'OSF3','1','0') OSF3E "
    sSQL = sSQL & "from "
    sSQL = sSQL & "(select SXLIDCW,REPSMPLIDCW,INPOSCW,WFSMPLIDL4CW from XSDCW where SXLIDCW in ("

    For c0 = 1 To intSXLCnt
        sSQL = sSQL & "'" & sxl(c0).SXLIDCA & "'"
        If c0 <> intSXLCnt Then sSQL = sSQL & ", " Else sSQL = sSQL & ") "
    Next c0

    sSQL = sSQL & "and LIVKCW = '0'), "
    sSQL = sSQL & "(select SAMPLEID,SPEC RES_SPEC from TBCMY013 where SPEC = 'RES') RES, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OI_SPEC from TBCMY013 where SPEC = 'OI') OI, "
    sSQL = sSQL & "(select SAMPLEID,SPEC BMD1_SPEC from TBCMY013 where SPEC = 'BMD1') BMD1, "
    sSQL = sSQL & "(select SAMPLEID,SPEC BMD2_SPEC from TBCMY013 where SPEC = 'BMD2') BMD2, "
    sSQL = sSQL & "(select SAMPLEID,SPEC BMD3_SPEC from TBCMY013 where SPEC = 'BMD3') BMD3, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF1_SPEC from TBCMY013 where SPEC = 'OSF1') OSF1, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF2_SPEC from TBCMY013 where SPEC = 'OSF2') OSF2, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF3_SPEC from TBCMY013 where SPEC = 'OSF3') OSF3, "
'    sSql = sSql & "(select SAMPLEID,SPEC OSF4_SPEC from TBCMY013 where SPEC = 'OSF4') OSF4, "  'SIRD�Ή�
'    sSQL = sSQL & "(select SMPLNO,SPEC SIRD_SPEC from TBCMJ022 where SPEC = 'SIRD') SIRD, "    '2010/05/19 REP Y.Hitomi
    sSQL = sSQL & "(select SMPLNO,SPEC SIRD_SPEC from TBCMJ022 where SPEC = 'TENI') SIRD, "
    sSQL = sSQL & "(select SAMPLEID,SPEC DSOD_SPEC from TBCMY013 where SPEC = 'DSOD') DSOD, "
    sSQL = sSQL & "(select SAMPLEID,SPEC DZ_SPEC from TBCMY013 where SPEC = 'DZ') DZ, "
    sSQL = sSQL & "(select SAMPLEID,SPEC DOI1_SPEC from TBCMY013 where SPEC = 'DOI1') DOI1, "
    sSQL = sSQL & "(select SAMPLEID,SPEC DOI2_SPEC from TBCMY013 where SPEC = 'DOI2') DOI2, "
    sSQL = sSQL & "(select SAMPLEID,SPEC DOI3_SPEC from TBCMY013 where SPEC = 'DOI3') DOI3, "
    sSQL = sSQL & "(select SAMPLEID,SPEC AOI_SPEC from TBCMY013 where SPEC = 'AOI') AOI, "
    sSQL = sSQL & "(select SMPLNO,'GD' GD_SPEC from TBCMJ015 where HSFLG = '1') GD, "
    sSQL = sSQL & "(select SMPLNO,'SPV' SPV_SPEC from TBCMJ016 where HSFLG = '1') SPV, "
    sSQL = sSQL & "(select SAMPLEID,'1' IJO from TBCMY016) Y16 "

    '�G�s��s�]���ǉ��Ή�
    sSQL = sSQL & ",(select SAMPLEID,SPEC BMD1E_SPEC from TBCMY022 where SPEC = 'BMD1') BMD1E, "
    sSQL = sSQL & "(select SAMPLEID,SPEC BMD2E_SPEC from TBCMY022 where SPEC = 'BMD2') BMD2E, "
    sSQL = sSQL & "(select SAMPLEID,SPEC BMD3E_SPEC from TBCMY022 where SPEC = 'BMD3') BMD3E, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF1E_SPEC from TBCMY022 where SPEC = 'OSF1') OSF1E, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF2E_SPEC from TBCMY022 where SPEC = 'OSF2') OSF2E, "
    sSQL = sSQL & "(select SAMPLEID,SPEC OSF3E_SPEC from TBCMY022 where SPEC = 'OSF3') OSF3E "
    sSQL = sSQL & "where REPSMPLIDCW = RES.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OI.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = BMD1.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = BMD2.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = BMD3.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF1.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF2.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF3.SAMPLEID(+) "
'    sSql = sSql & "and REPSMPLIDCW = OSF4.SAMPLEID(+) " SIRD_Y.Hitomi
    sSQL = sSQL & "and WFSMPLIDL4CW = SIRD.SMPLNO(+) "
    sSQL = sSQL & "and REPSMPLIDCW = DSOD.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = DZ.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = DOI1.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = DOI2.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = DOI3.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = AOI.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = GD.SMPLNO(+) "
    sSQL = sSQL & "and REPSMPLIDCW = SPV.SMPLNO(+) "
    sSQL = sSQL & "and REPSMPLIDCW = Y16.SAMPLEID(+) "

    '�G�s��s�]���ǉ��Ή�
    sSQL = sSQL & "and REPSMPLIDCW = BMD1E.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = BMD2E.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = BMD3E.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF1E.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF2E.SAMPLEID(+) "
    sSQL = sSQL & "and REPSMPLIDCW = OSF3E.SAMPLEID(+) "
    sSQL = sSQL & "order by SXLIDCW,INPOSCW "

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)
    For c0 = 1 To rs.RecordCount
        For c1 = 1 To intSXLCnt
            For c2 = 1 To UBound(sxl(c1).WFSMP())
                If sxl(c1).WFSMP(c2).REPSMPLIDCW = rs("REPSMPLIDCW") Then
                    '��M����FLG�̏�����
                    With udtChkUp
                        .rs = "0"           ' ��MFLG�iRs)
                        .Oi = "0"           ' ��MFLG�iOi)
                        .B1 = "0"           ' ��MFLG�iB1)
                        .B2 = "0"           ' ��MFLG�iB2�j
                        .B3 = "0"           ' ��MFLG�iB3)
                        .L1 = "0"           ' ��MFLG�iL1)
                        .L2 = "0"           ' ��MFLG�iL2)
                        .L3 = "0"           ' ��MFLG�iL3)
                        .L4 = "0"           ' ��MFLG�iL4)
                        .DS = "0"           ' ��MFLG�iDS)
                        .DZ = "0"           ' ��MFLG�iDZ)
                        .sp = "0"           ' ��MFLG�iSP)
                        .DO1 = "0"          ' ��MFLG�iDO1)
                        .DO2 = "0"          ' ��MFLG�iDO2)
                        .DO3 = "0"          ' ��MFLG�iDO3)
                        .OT1 = "0"          ' ��MFLG (OT2)
                        .OT2 = "0"          ' ��MFLG (OT1)
                        .AOI = "0"          ' ��MFLG (AOi)
                        .GD = "0"           ' ��MFLG (GD)

                        '�G�s��s�]���ǉ��Ή�
                        .B1E = "0"           ' ��MFLG�iB1E)
                        .B2E = "0"           ' ��MFLG�iB2E�j
                        .B3E = "0"           ' ��MFLG�iB3E)
                        .L1E = "0"           ' ��MFLG�iL1E)
                        .L2E = "0"           ' ��MFLG�iL2E)
                        .L3E = "0"           ' ��MFLG�iL3E)
                    End With

                    blChangeFlag = False
                    With sxl(c1).WFSMP(c2)
                        If rs("RES") = "1" Then         'RES
                            If (.REPSMPLIDCW = .WFSMPLIDRSCW) And (.WFRESRS1CW = "0") Then
                                .WFRESRS1CW = "1"
                                blChangeFlag = True
                                udtChkUp.rs = "1"
                            End If
                        End If
                        If rs("OI") = "1" Then          'OI
                            If (.REPSMPLIDCW = .WFSMPLIDOICW) And (.WFRESOICW = "0") Then
                                .WFRESOICW = "1"
                                blChangeFlag = True
                                udtChkUp.Oi = "1"
                            End If
                        End If
                        If rs("BMD1") = "1" Then        'BMD1
                            If (.REPSMPLIDCW = .WFSMPLIDB1CW) And (.WFRESB1CW = "0") Then
                                .WFRESB1CW = "1"
                                blChangeFlag = True
                                udtChkUp.B1 = "1"
                            End If
                        End If
                        If rs("BMD2") = "1" Then        'BMD2
                            If (.REPSMPLIDCW = .WFSMPLIDB2CW) And (.WFRESB2CW = "0") Then
                                .WFRESB2CW = "1"
                                blChangeFlag = True
                                udtChkUp.B2 = "1"
                            End If
                        End If
                        If rs("BMD3") = "1" Then        'BMD3
                            If (.REPSMPLIDCW = .WFSMPLIDB3CW) And (.WFRESB3CW = "0") Then
                                .WFRESB3CW = "1"
                                blChangeFlag = True
                                udtChkUp.B3 = "1"
                            End If
                        End If
                        If rs("OSF1") = "1" Then        'OSF1
                            If (.REPSMPLIDCW = .WFSMPLIDL1CW) And (.WFRESL1CW = "0") Then
                                .WFRESL1CW = "1"
                                blChangeFlag = True
                                udtChkUp.L1 = "1"
                            End If
                        End If
                        If rs("OSF2") = "1" Then        'OSF2
                            If (.REPSMPLIDCW = .WFSMPLIDL2CW) And (.WFRESL2CW = "0") Then
                                .WFRESL2CW = "1"
                                blChangeFlag = True
                                udtChkUp.L2 = "1"
                            End If
                        End If
                        If rs("OSF3") = "1" Then        'OSF3
                            If (.REPSMPLIDCW = .WFSMPLIDL3CW) And (.WFRESL3CW = "0") Then
                                .WFRESL3CW = "1"
                                blChangeFlag = True
                                udtChkUp.L3 = "1"
                            End If
                        End If
'                        If rs("OSF4") = "1" Then        'OSF4 SIRD_Y.Hitomi
'                            If (.REPSMPLIDCW = .WFSMPLIDL4CW) And (.WFRESL4CW = "0") Then
'                                .WFRESL4CW = "1"
'                                blChangeFlag = True
'                                udtChkUp.L4 = "1"
'                            End If
'                        End If
                        If rs("SIRD") = "1" Then        'SIRD
                            If (.WFRESL4CW = "0") Then
                                .WFRESL4CW = "1"
                                blChangeFlag = True
                                udtChkUp.L4 = "1"
                            End If
                        End If
                        If rs("DSOD") = "1" Then        'DSOD
                            If (.REPSMPLIDCW = .WFSMPLIDDSCW) And (.WFRESDSCW = "0") Then
                                .WFRESDSCW = "1"
                                blChangeFlag = True
                                udtChkUp.DS = "1"
                            End If
                        End If
                        If rs("DZ") = "1" Then          'DZ
                            If (.REPSMPLIDCW = .WFSMPLIDDZCW) And (.WFRESDZCW = "0") Then
                                .WFRESDZCW = "1"
                                blChangeFlag = True
                                udtChkUp.DZ = "1"
                            End If
                        End If
                        If rs("DOI1") = "1" Then        'DOI1
                            If (.REPSMPLIDCW = .WFSMPLIDDO1CW) And (.WFRESDO1CW = "0") Then
                                .WFRESDO1CW = "1"
                                blChangeFlag = True
                                udtChkUp.DO1 = "1"
                            End If
                        End If
                        If rs("DOI2") = "1" Then        'DOI2
                            If (.REPSMPLIDCW = .WFSMPLIDDO2CW) And (.WFRESDO2CW = "0") Then
                                .WFRESDO2CW = "1"
                                blChangeFlag = True
                                udtChkUp.DO2 = "1"
                            End If
                        End If
                        If rs("DOI3") = "1" Then        'DOI3
                            If (.REPSMPLIDCW = .WFSMPLIDDO3CW) And (.WFRESDO3CW = "0") Then
                                .WFRESDO3CW = "1"
                                blChangeFlag = True
                                udtChkUp.DO3 = "1"
                            End If
                        End If
                        If rs("AOI") = "1" Then         'AOI
                            If (.REPSMPLIDCW = .WFSMPLIDAOICW) And (.WFRESAOICW = "0") Then
                                .WFRESAOICW = "1"
                                blChangeFlag = True
                                udtChkUp.AOI = "1"
                            End If
                        End If
                        If rs("GD") = "1" Then          'GD
                            If (.REPSMPLIDCW = .WFSMPLIDGDCW) And (.WFRESGDCW = "0") Then
                                .WFRESGDCW = "1"
                                blChangeFlag = True
                                udtChkUp.GD = "1"
                            End If
                        End If
                        If rs("SPV") = "1" Then         'SPV
                            If (.REPSMPLIDCW = .WFSMPLIDSPCW) And (.WFRESSPCW = "0") Then
                                .WFRESSPCW = "1"
                                blChangeFlag = True
                                udtChkUp.sp = "1"
                            End If
                        End If

                        '�G�s��s�]���ǉ��Ή�
                        If rs("BMD1E") = "1" Then        'BMD1E
                            If (.REPSMPLIDCW = .EPSMPLIDB1CW) And (.EPRESB1CW = "0") Then
                                .EPRESB1CW = "1"
                                blChangeFlag = True
                                udtChkUp.B1E = "1"
                            End If
                        End If
                        If rs("BMD2E") = "1" Then        'BMD2E
                            If (.REPSMPLIDCW = .EPSMPLIDB2CW) And (.EPRESB2CW = "0") Then
                                .EPRESB2CW = "1"
                                blChangeFlag = True
                                udtChkUp.B2E = "1"
                            End If
                        End If
                        If rs("BMD3E") = "1" Then        'BMD3E
                            If (.REPSMPLIDCW = .EPSMPLIDB3CW) And (.EPRESB3CW = "0") Then
                                .EPRESB3CW = "1"
                                blChangeFlag = True
                                udtChkUp.B3E = "1"
                            End If
                        End If
                        If rs("OSF1E") = "1" Then        'OSF1E
                            If (.REPSMPLIDCW = .EPSMPLIDL1CW) And (.EPRESL1CW = "0") Then
                                .EPRESL1CW = "1"
                                blChangeFlag = True
                                udtChkUp.L1E = "1"
                            End If
                        End If
                        If rs("OSF2E") = "1" Then        'OSF2E
                            If (.REPSMPLIDCW = .EPSMPLIDL2CW) And (.EPRESL2CW = "0") Then
                                .EPRESL2CW = "1"
                                blChangeFlag = True
                                udtChkUp.L2E = "1"
                            End If
                        End If
                        If rs("OSF3E") = "1" Then        'OSF3E
                            If (.REPSMPLIDCW = .EPSMPLIDL3CW) And (.EPRESL3CW = "0") Then
                                .EPRESL3CW = "1"
                                blChangeFlag = True
                                udtChkUp.L3E = "1"
                            End If
                        End If
                        If blChangeFlag Then
                            '���T���v��ID�̎��ѕt��
                            If WfSmp_Upd_SmplID(.XTALCW, .REPSMPLIDCW, udtChkUp) = FUNCTION_RETURN_FAILURE Then
                                MeasRsltCheck = FUNCTION_RETURN_FAILURE
                            End If
                            'Add 2010/01/21 SIRD�Ή� Y.Hitomi
                            If WfSmp_Upd_SmplID_SD(.XTALCW, .WFSMPLIDL4CW, udtChkUp) = FUNCTION_RETURN_FAILURE Then
                                MeasRsltCheck = FUNCTION_RETURN_FAILURE
                            End If
                        '�����ُ킪�o�^����Ă���
                        ElseIf rs("IJO") = "1" Then
                            '�G�s��s�]���ǉ��Ή�
                            If (.WFINDRSCW <> "0" And .WFRESRS1CW <> "2") Or _
                               (.WFINDOICW <> "0" And .WFRESOICW <> "2") Or _
                               (.WFINDB1CW <> "0" And .WFRESB1CW <> "2") Or _
                               (.WFINDB2CW <> "0" And .WFRESB2CW <> "2") Or _
                               (.WFINDB3CW <> "0" And .WFRESB3CW <> "2") Or _
                               (.WFINDL1CW <> "0" And .WFRESL1CW <> "2") Or _
                               (.WFINDL2CW <> "0" And .WFRESL2CW <> "2") Or _
                               (.WFINDL3CW <> "0" And .WFRESL3CW <> "2") Or _
                               (.WFINDL4CW <> "0" And .WFRESL4CW <> "2") Or _
                               (.WFINDDSCW <> "0" And .WFRESDSCW <> "2") Or _
                               (.WFINDDZCW <> "0" And .WFRESDZCW <> "2") Or _
                               (.WFINDSPCW <> "0" And .WFRESSPCW <> "2") Or _
                               (.WFINDDO1CW <> "0" And .WFRESDO1CW <> "2") Or _
                               (.WFINDDO2CW <> "0" And .WFRESDO2CW <> "2") Or _
                               (.WFINDDO3CW <> "0" And .WFRESDO3CW <> "2") Or _
                               (.WFINDAOICW <> "0" And .WFRESAOICW <> "2") Or _
                               (.WFINDGDCW <> "0" And .WFHSGDCW <> "1" And .WFRESGDCW <> "2") Or _
                               (.EPINDB1CW <> "0" And .EPRESB1CW <> "2") Or _
                               (.EPINDB2CW <> "0" And .EPRESB2CW <> "2") Or _
                               (.EPINDB3CW <> "0" And .EPRESB3CW <> "2") Or _
                               (.EPINDL1CW <> "0" And .EPRESL1CW <> "2") Or _
                               (.EPINDL2CW <> "0" And .EPRESL2CW <> "2") Or _
                               (.EPINDL3CW <> "0" And .EPRESL3CW <> "2") Then

                                If .WFINDRSCW <> "0" Then .WFRESRS1CW = "2"
                                If .WFINDOICW <> "0" Then .WFRESOICW = "2"
                                If .WFINDB1CW <> "0" Then .WFRESB1CW = "2"
                                If .WFINDB2CW <> "0" Then .WFRESB2CW = "2"
                                If .WFINDB3CW <> "0" Then .WFRESB3CW = "2"
                                If .WFINDL1CW <> "0" Then .WFRESL1CW = "2"
                                If .WFINDL2CW <> "0" Then .WFRESL2CW = "2"
                                If .WFINDL3CW <> "0" Then .WFRESL3CW = "2"
                                If .WFINDL4CW <> "0" Then .WFRESL4CW = "2"
                                If .WFINDDSCW <> "0" Then .WFRESDSCW = "2"
                                If .WFINDDZCW <> "0" Then .WFRESDZCW = "2"
                                If .WFINDSPCW <> "0" Then .WFRESSPCW = "2"
                                If .WFINDDO1CW <> "0" Then .WFRESDO1CW = "2"
                                If .WFINDDO2CW <> "0" Then .WFRESDO2CW = "2"
                                If .WFINDDO3CW <> "0" Then .WFRESDO3CW = "2"
                                If .WFINDAOICW <> "0" Then .WFRESAOICW = "2"
                                If .WFINDGDCW <> "0" And .WFHSGDCW <> "1" Then .WFRESGDCW = "2"
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                                If .EPINDB1CW <> "0" Then .EPRESB1CW = "2"
                                If .EPINDB2CW <> "0" Then .EPRESB2CW = "2"
                                If .EPINDB3CW <> "0" Then .EPRESB3CW = "2"
                                If .EPINDL1CW <> "0" Then .EPRESL1CW = "2"
                                If .EPINDL2CW <> "0" Then .EPRESL2CW = "2"
                                If .EPINDL3CW <> "0" Then .EPRESL3CW = "2"
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                                If WfSmp_Upd(sxl(c1).WFSMP(c2)) = FUNCTION_RETURN_FAILURE Then
                                    MeasRsltCheck = FUNCTION_RETURN_FAILURE
                                End If
                            End If
                        End If
                    End With
                    GoTo LoopNext
                End If
            Next c2
        Next c1
LoopNext:
        rs.MoveNext
    Next c0
    rs.Close
    '����]�����ʎ擾sSql�ύX�@06/02/07 ooba END =============================================>

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error sSql ======"
    Debug.Print sSQL
    MeasRsltCheck = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add Start 2011/06/16 Y.Hitomi
'***************************************************************************************
'*    �֐���        : MeasRsltCheck1
'*
'*    �����T�v      : 1.����]�����ʎ�M�m�F
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                            ,����
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL�Ǘ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************
Private Function MeasRsltCheck1(sxl() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim intSXLCnt       As Integer
    Dim udtSokutei()    As typ_TBCMY013
    Dim intGDcnt        As Integer           'GD����ں��ސ�
''SPV9�_�Ή�
    Dim intSPVCnt       As Integer          'SPV���у��R�[�h��
    Dim c0              As Integer
    Dim c1              As Integer
    Dim c2              As Integer
    Dim blChangeFlag    As Boolean
    Dim blPassFlg       As Boolean
    Dim sSqlWhere       As String
#If SPEEDUP Then
    Dim sChkWF()        As String
    Dim i               As Integer
#End If
    Dim udtChkUp        As type_chkUP

    '�G���[�n���h���̐ݒ�
    'On Error GoTo proc_err
'    gErr.Push "s_cmbc039_sSql.bas -- Function MeasRsltCheck"

    MeasRsltCheck1 = FUNCTION_RETURN_SUCCESS

    '����]�����ʎ擾sSql�ύX
    intSXLCnt = UBound(sxl)
    If intSXLCnt = 0 Then GoTo proc_exit

    sSQL = sSQL & "select "
    sSQL = sSQL & "SXLIDCW, "
    sSQL = sSQL & "REPSMPLIDCW, "
    sSQL = sSQL & "decode(trim(SIRD_SPEC),'TENI','1','0') SIRD "
    sSQL = sSQL & "from "
    sSQL = sSQL & "(select SXLIDCW,REPSMPLIDCW,INPOSCW,WFSMPLIDL4CW from XSDCW where SXLIDCW in ("

    For c0 = 1 To intSXLCnt
        sSQL = sSQL & "'" & sxl(c0).SXLIDCA & "'"
        If c0 <> intSXLCnt Then sSQL = sSQL & ", " Else sSQL = sSQL & ") "
    Next c0

    sSQL = sSQL & "and LIVKCW = '0'), "
    sSQL = sSQL & "(select SMPLNO,SPEC SIRD_SPEC from TBCMJ022 where SPEC = 'TENI') SIRD "
    sSQL = sSQL & "where "

    sSQL = sSQL & "WFSMPLIDL4CW = SIRD.SMPLNO(+) "
    sSQL = sSQL & "order by SXLIDCW,INPOSCW "

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)
    For c0 = 1 To rs.RecordCount
        For c1 = 1 To intSXLCnt
            For c2 = 1 To UBound(sxl(c1).WFSMP())
                If sxl(c1).WFSMP(c2).REPSMPLIDCW = rs("REPSMPLIDCW") Then
                    '��M����FLG�̏�����
                    With udtChkUp
                        .rs = "0"           ' ��MFLG�iRs)
                        .Oi = "0"           ' ��MFLG�iOi)
                        .B1 = "0"           ' ��MFLG�iB1)
                        .B2 = "0"           ' ��MFLG�iB2�j
                        .B3 = "0"           ' ��MFLG�iB3)
                        .L1 = "0"           ' ��MFLG�iL1)
                        .L2 = "0"           ' ��MFLG�iL2)
                        .L3 = "0"           ' ��MFLG�iL3)
                        .L4 = "0"           ' ��MFLG�iL4)
                        .DS = "0"           ' ��MFLG�iDS)
                        .DZ = "0"           ' ��MFLG�iDZ)
                        .sp = "0"           ' ��MFLG�iSP)
                        .DO1 = "0"          ' ��MFLG�iDO1)
                        .DO2 = "0"          ' ��MFLG�iDO2)
                        .DO3 = "0"          ' ��MFLG�iDO3)
                        .OT1 = "0"          ' ��MFLG (OT2)
                        .OT2 = "0"          ' ��MFLG (OT1)
                        .AOI = "0"          ' ��MFLG (AOi)
                        .GD = "0"           ' ��MFLG (GD)

                        '�G�s��s�]���ǉ��Ή�
                        .B1E = "0"           ' ��MFLG�iB1E)
                        .B2E = "0"           ' ��MFLG�iB2E�j
                        .B3E = "0"           ' ��MFLG�iB3E)
                        .L1E = "0"           ' ��MFLG�iL1E)
                        .L2E = "0"           ' ��MFLG�iL2E)
                        .L3E = "0"           ' ��MFLG�iL3E)
                    End With

                    blChangeFlag = False
                    With sxl(c1).WFSMP(c2)
                        If rs("SIRD") = "1" Then        'SIRD
                            If (.WFRESL4CW = "0") Then
                                .WFRESL4CW = "1"
                                blChangeFlag = True
                                udtChkUp.L4 = "1"
                            End If
                        End If
                        If blChangeFlag Then
                            If WfSmp_Upd_SmplID_SD(.XTALCW, .WFSMPLIDL4CW, udtChkUp) = FUNCTION_RETURN_FAILURE Then
                                MeasRsltCheck1 = FUNCTION_RETURN_FAILURE
                            End If
                        End If
                    End With
                    GoTo LoopNext
                End If
            Next c2
        Next c1
LoopNext:
        rs.MoveNext
    Next c0
    rs.Close
    '����]�����ʎ擾sSql�ύX�@06/02/07 ooba END =============================================>

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error sSql ======"
    Debug.Print sSQL
    MeasRsltCheck1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add End 2011/06/16 Y.Hitomi

'*******************************************************************************
'*    �֐���        : WfSmp_Upd
'*
'*    �����T�v      : 1.WF�T���v���Ǘ��A�b�v�f�[�g
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^              ,����
'*                    WFSMP         ,I  ,typ_XSDCW       ,�V�T���v���Ǘ��iSXL�j
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function WfSmp_Upd(WFSMP As typ_XSDCW) As FUNCTION_RETURN
    Dim sSQL As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function WfSmp_Upd"

    WfSmp_Upd = FUNCTION_RETURN_SUCCESS

    With WFSMP
        sSQL = sSQL & "update XSDCW set "
        sSQL = sSQL & "REPSMPLIDCW='" & .REPSMPLIDCW & "',"           '�T���v��ID
        sSQL = sSQL & "HINBCW='" & .HINBCW & "',"                     '�i��
        sSQL = sSQL & "REVNUMCW=" & .REVNUMCW & ","                   '���i�ԍ������ԍ�
        sSQL = sSQL & "FACTORYCW='" & .FACTORYCW & "',"               '�H��
        sSQL = sSQL & "OPECW='" & .OPECW & "',"                       '���Ə���
        sSQL = sSQL & "KTKBNCW='" & .KTKBNCW & "',"                   '�m��敪
        sSQL = sSQL & "WFINDRSCW='" & .WFINDRSCW & "',"               '���FLG(Rs)
        sSQL = sSQL & "WFRESRS1CW='" & .WFRESRS1CW & "',"                '����FLG(Rs)
        sSQL = sSQL & "WFINDOICW='" & .WFINDOICW & "',"               '���FLG(Oi)
        sSQL = sSQL & "WFRESOICW='" & .WFRESOICW & "',"                 '����FLG(Oi)
        sSQL = sSQL & "WFINDB1CW='" & .WFINDB1CW & "',"               '���FLG(B1)
        sSQL = sSQL & "WFRESB1CW='" & .WFRESB1CW & "',"                 '����FLG(B1)
        sSQL = sSQL & "WFINDB2CW='" & .WFINDB2CW & "',"               '���FLG(B2)
        sSQL = sSQL & "WFRESB2CW='" & .WFRESB2CW & "',"                 '����FLG(B2)
        sSQL = sSQL & "WFINDB3CW='" & .WFINDB3CW & "',"               '���FLG(B3)
        sSQL = sSQL & "WFRESB3CW='" & .WFRESB3CW & "',"                 '����FLG(B3)
        sSQL = sSQL & "WFINDL1CW='" & .WFINDL1CW & "',"               '���FLG(L1)
        sSQL = sSQL & "WFRESL1CW='" & .WFRESL1CW & "',"                 '����FLG(L1)
        sSQL = sSQL & "WFINDL2CW='" & .WFINDL2CW & "',"               '���FLG(L2)
        sSQL = sSQL & "WFRESL2CW='" & .WFRESL2CW & "',"                 '����FLG(L2)
        sSQL = sSQL & "WFINDL3CW='" & .WFINDL3CW & "',"               '���FLG(L3)
        sSQL = sSQL & "WFRESL3CW='" & .WFRESL3CW & "',"                 '����FLG(L3)
        sSQL = sSQL & "WFINDL4CW='" & .WFINDL4CW & "',"               '���FLG(L4)
        sSQL = sSQL & "WFRESL4CW='" & .WFRESL4CW & "',"                 '����FLG(L4)
        sSQL = sSQL & "WFINDDSCW='" & .WFINDDSCW & "',"               '���FLG(DS)
        sSQL = sSQL & "WFRESDSCW='" & .WFRESDSCW & "',"                 '����FLG(DS)
        sSQL = sSQL & "WFINDDZCW='" & .WFINDDZCW & "',"               '���FLG(DZ)
        sSQL = sSQL & "WFRESDZCW='" & .WFRESDZCW & "',"                 '����FLG(DZ)
        sSQL = sSQL & "WFINDSPCW='" & .WFINDSPCW & "',"               '���FLG(SP)
        sSQL = sSQL & "WFRESSPCW='" & .WFRESSPCW & "',"                 '����FLG(SP)
        sSQL = sSQL & "WFINDDO1CW='" & .WFINDDO1CW & "',"             '���FLG(DO1)
        sSQL = sSQL & "WFRESDO1CW='" & .WFRESDO1CW & "', "              '����FLG(DO1)
        sSQL = sSQL & "WFINDDO2CW='" & .WFINDDO2CW & "',"             '���FLG(DO2)
        sSQL = sSQL & "WFRESDO2CW='" & .WFRESDO2CW & "',"               '����FLG(DO2)
        sSQL = sSQL & "WFINDDO3CW='" & .WFINDDO3CW & "',"             '���FLG(DO3)
        sSQL = sSQL & "WFRESDO3CW='" & .WFRESDO3CW & "',"               '����FLG(DO3)
        ''�c���_�f�ǉ��@03/12/15 ooba
        sSQL = sSQL & "WFINDAOICW='" & .WFINDAOICW & "',"             '���FLG(AOi)
        sSQL = sSQL & "WFRESAOICW='" & .WFRESAOICW & "',"               '����FLG(AOi)
        'GD�ǉ��@05/02/04 ooba
        sSQL = sSQL & "WFINDGDCW='" & .WFINDGDCW & "',"               '���FLG(GD)
        sSQL = sSQL & "WFRESGDCW='" & .WFRESGDCW & "',"               '����FLG(GD)
        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
        sSQL = sSQL & "EPINDB1CW='" & .EPINDB1CW & "',"               '���FLG(B1E)
        sSQL = sSQL & "EPRESB1CW='" & .EPRESB1CW & "',"                 '����FLG(B1E)
        sSQL = sSQL & "EPINDB2CW='" & .EPINDB2CW & "',"               '���FLG(B2E)
        sSQL = sSQL & "EPRESB2CW='" & .EPRESB2CW & "',"                 '����FLG(B2E)
        sSQL = sSQL & "EPINDB3CW='" & .EPINDB3CW & "',"               '���FLG(B3E)
        sSQL = sSQL & "EPRESB3CW='" & .EPRESB3CW & "',"                 '����FLG(B3E)
        sSQL = sSQL & "EPINDL1CW='" & .EPINDL1CW & "',"               '���FLG(L1E)
        sSQL = sSQL & "EPRESL1CW='" & .EPRESL1CW & "',"                 '����FLG(L1E)
        sSQL = sSQL & "EPINDL2CW='" & .EPINDL2CW & "',"               '���FLG(L2E)
        sSQL = sSQL & "EPRESL2CW='" & .EPRESL2CW & "',"                 '����FLG(L2E)
        sSQL = sSQL & "EPINDL3CW='" & .EPINDL3CW & "',"               '���FLG(L3E)
        sSQL = sSQL & "EPRESL3CW='" & .EPRESL3CW & "',"                 '����FLG(L3E)
        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
        sSQL = sSQL & "KDAYCW=sysdate,"
        sSQL = sSQL & "SNDKCW='0'"
        sSQL = sSQL & "where XTALCW='" & .XTALCW & "'"
        sSQL = sSQL & "and INPOSCW=" & .INPOSCW & ""
        sSQL = sSQL & "and SMPKBNCW='" & .SMPKBNCW & "'"
    End With

    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        WfSmp_Upd = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    WfSmp_Upd = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'****************************************************************************************************
'*    �֐���        : WfSmp_Upd_SmplID
'*
'*    �����T�v      : 1.�e�[�u���uXSDCW�v�̏����ɂ��������R�[�h���X�V����(�����ID�̎����׸�)
'*
'*    �p�����[�^    : �ϐ���       ,IO  ,�^                      ,����
'*                  :XTAL          ,I   ,String                  ,�����ԍ�
'*                  :WFSMPID       ,I   ,String                  ,�T���v��ID
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'****************************************************************************************************
Private Function WfSmp_Upd_SmplID(xtal As String, WFSMPID As String, chkUp As type_chkUP) As Integer
    Dim sSQL    As String
    Dim rs      As OraDynaset    'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function WfSmp_Upd_SmplID"

    WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE

    ''�T���v���Ǘ��X�V���@�������@2004/02/06 TUKU START ===============================================>
    ''�������Ƃ̍X�VFLG�̌��ʂŌ������ƂɎ���FLG�̍X�V���s���悤�ɕύX

    ' �����ID(Rs)�̍X�V
    If chkUp.rs = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESRS1CW='1', "                              ' ����FLG1(Rs)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDRSCW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(Oi)�̍X�V
    If chkUp.Oi = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESOICW='1', "                               ' ����FLG(Oi)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDOICW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(B1)�̍X�V
    If chkUp.B1 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESB1CW='1', "                               ' ����FLG(B1)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDB1CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(B2)�̍X�V
    If chkUp.B2 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESB2CW='1', "                               ' ����FLG(B2)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDB2CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(B3)�̍X�V
    If chkUp.B3 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESB3CW='1', "                               ' ����FLG(B3)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDB3CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(L1)�̍X�V
    If chkUp.L1 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESL1CW='1', "                               ' ����FLG(L1)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDL1CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(L2)�̍X�V
    If chkUp.L2 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESL2CW='1', "                               ' ����FLG(L2)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDL2CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(L3)�̍X�V
    If chkUp.L3 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESL3CW='1', "                               ' ����FLG(L3)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDL3CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
'Del 2010/01/21 SIRD�Ή� Y.Hitomi
'    ' �����ID(L4)�̍X�V
'    If chkUp.L4 = "1" Then
'        sSql = "update XSDCW set "
'        sSql = sSql & "WFRESL4CW='1', "                               ' ����FLG(L4)
'        sSql = sSql & "KDAYCW=sysdate "                               ' �X�V���t
'        sSql = sSql & "WHERE XTALCW = '" & xtal & "'"
'        sSql = sSql & "      WFSMPLIDL4CW = '" & WFSMPID & "'"
'
'        If OraDB.ExecuteSQL(sSql) <= 0 Then
'            rs.Close
'            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    End If

    ' �����ID(DS)�̍X�V
    If chkUp.DS = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESDSCW='1', "                               ' ����FLG(DS)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDDSCW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(DZ)�̍X�V
    If chkUp.DZ = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESDZCW='1', "                               ' ����FLG(DZ)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDDZCW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(SP)�̍X�V
    If chkUp.sp = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESSPCW='1', "                               ' ����FLG(SP)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDSPCW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(DO1)�̍X�V
    If chkUp.DO1 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESDO1CW='1', "                              ' ����FLG(DO1)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDDO1CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(DO2)�̍X�V
    If chkUp.DO2 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESDO2CW='1', "                              ' ����FLG(DO2)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDDO2CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(DO3)�̍X�V
    If chkUp.DO3 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESDO3CW='1', "                              ' ����FLG(DO3)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDDO3CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ''�c���_�f�ǉ��@03/12/15 ooba START ===============================================>
    ' �����ID(AOi)�̍X�V
    If chkUp.AOI = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESAOICW='1', "                              ' ����FLG(AOi)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDAOICW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''�c���_�f�ǉ��@03/12/15 ooba END =================================================>

    ''GD�ǉ��@05/02/04 ooba START =====================================================>
    If chkUp.GD = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESGDCW='1', "                               ' ����FLG(GD)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      WFSMPLIDGDCW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''GD�ǉ��@05/02/04 ooba END =======================================================>

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    ' �����ID(B1E)�̍X�V
    If chkUp.B1E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESB1CW='1', "                               ' ����FLG(B1)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDB1CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(B2E)�̍X�V
    If chkUp.B2E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESB2CW='1', "                               ' ����FLG(B2)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDB2CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(B3E)�̍X�V
    If chkUp.B3E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESB3CW='1', "                               ' ����FLG(B3)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDB3CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(L1E)�̍X�V
    If chkUp.L1E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESL1CW='1', "                               ' ����FLG(L1)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDL1CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(L2E)�̍X�V
    If chkUp.L2E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESL2CW='1', "                               ' ����FLG(L2)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDL2CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    ' �����ID(L3E)�̍X�V
    If chkUp.L3E = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "EPRESL3CW='1', "                               ' ����FLG(L3)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "' and "
        sSQL = sSQL & "      EPSMPLIDL3CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    ''�T���v���Ǘ��X�V���@�������@2004/02/06 TUKU END ===============================================>

    WfSmp_Upd_SmplID = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    WfSmp_Upd_SmplID = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
''Add 2010/01/21 SIRD�]���Ή� Y.Hitomi
'****************************************************************************************************
'*    �֐���        : WfSmp_Upd_SmplID_SD
'*
'*    �����T�v      : 1.�e�[�u���uXSDCW�v�̏����ɂ��������R�[�h���X�V����(�����ID�̎����׸�) SIRD�]���p
'*
'*    �p�����[�^    : �ϐ���       ,IO  ,�^                      ,����
'*                  :XTAL          ,I   ,String                  ,�����ԍ�
'*                  :WFSMPID       ,I   ,String                  ,�T���v��ID
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'****************************************************************************************************
Private Function WfSmp_Upd_SmplID_SD(xtal As String, WFSMPID As String, chkUp As type_chkUP) As Integer
    Dim sSQL    As String
    Dim rs      As OraDynaset    'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    WfSmp_Upd_SmplID_SD = FUNCTION_RETURN_FAILURE


    If chkUp.L4 = "1" Then
        sSQL = "update XSDCW set "
        sSQL = sSQL & "WFRESL4CW='1', "                               ' ����FLG(SIRD)
        sSQL = sSQL & "KDAYCW=sysdate "                               ' �X�V���t
        sSQL = sSQL & "WHERE XTALCW = '" & xtal & "'"
        sSQL = sSQL & " and WFSMPLIDL4CW = '" & WFSMPID & "'"

        If OraDB.ExecuteSQL(sSQL) <= 0 Then
            rs.Close
            WfSmp_Upd_SmplID_SD = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If


    WfSmp_Upd_SmplID_SD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    WfSmp_Upd_SmplID_SD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'*******************************************************************************
'*    �֐���        : TBCMY016Check
'*
'*    �����T�v      : 1.�����ُ�`�F�b�N
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    WFSMPID       ,O  ,String   ,SXL�Ǘ�
'*
'*    �߂�l        : �g�p���Ă��Ȃ�
'*
'*******************************************************************************
Private Function TBCMY016Check(WFSMPID As String) As Integer
    Dim sSQL    As String
    Dim rs      As OraDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function TBCMY016Check"

    sSQL = "select "
    sSQL = sSQL & " SAMPLEID "
    sSQL = sSQL & " from "
    sSQL = sSQL & " TBCMY016 "
    sSQL = sSQL & " where "
    sSQL = sSQL & " SAMPLEID = '" & WFSMPID & "' "
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    TBCMY016Check = rs.RecordCount
    rs.Close

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    TBCMY016Check = -1
    gErr.HandleError
    Resume proc_exit
End Function

'************************************************************************************
'*    �֐���        : DBDRV_scmzc_fcmlc001b_Disp
'*
'*    �����T�v      : 1.WF�������� �҂��ꗗ �\���p�c�a�h���C�o
'*
'*    �p�����[�^    : �ϐ���       ,IO  ,�^                            ,����
'*              �@�@:inNowPorc     ,I   ,String                        ,���͗p(�H��)
'*              �@�@:SXL           ,O   ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL�Ǘ��p
'*              �@�@:sErrMsg �@�@�@,O   ,String                        ,�G���[���b�Z�[�W
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Public Function DBDRV_scmzc_fcmlc001b_Disp(inNowPorc As String, _
                                            sxl() As DBDRV_scmzc_fcmlc001b_SXL039, _
                                            sErrMsg As String _
                                            ) As FUNCTION_RETURN
    Dim udtWKSXL()      As DBDRV_scmzc_fcmlc001b_SXL039
    Dim sSQL            As String
    Dim sDBName         As String
    Dim rs              As OraDynaset
    Dim lngRecCnt       As Long
    Dim sCryNumBuf      As String
    Dim iIngotPosBuf    As Integer
    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim l               As Long
    Dim lngNullCnt      As Long
    Dim sNullSXLID      As String
    Dim intSxlCount     As Integer
    Dim rs2             As OraDynaset
    Dim lngCmpCnt       As Long             '�T���v����     'Add 2011/03/07 SMPK Miyata
    
    '�G���[�n���h���̐ݒ�
    'On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001b_Disp"
Debug.Print "1 " & Now & " SXL�Ǘ��AXSDCB���擾 SQL���s"
    DBDRV_scmzc_fcmlc001b_Disp = FUNCTION_RETURN_SUCCESS

    ' SXL�Ǘ��AXSDCB���擾�i�r���[���g�p���Ȃ��悤�ɕύX)�@TUKU  2003/10/9
''���Ӂj�҂��ꗗ�p�ׁ̈A���C���ł��B
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP���{
    sDBName = "(XSDCB)"

    '�d�|SXL�擾SQL�ύX�@06/02/07 ooba START ===============================================>
Debug.Print "2 " & Now & " SXL�ɑ΂���V����ُ����Z�b�g"
    sSQL = sSQL & "select "
    sSQL = sSQL & "CRYNUM, "
    sSQL = sSQL & "INGOTPOS, "
    sSQL = sSQL & "LENGTH, "
    sSQL = sSQL & "SXLID, "
    sSQL = sSQL & "KRPROCCD, "
    sSQL = sSQL & "NOWPROC, "
    sSQL = sSQL & "LPKRPROCCD, "
    sSQL = sSQL & "LASTPASS, "
    sSQL = sSQL & "DELCLS, "
    sSQL = sSQL & "LSTATCLS, "
    sSQL = sSQL & "HOLDCLS, "
    sSQL = sSQL & "HINBAN, "
    sSQL = sSQL & "REVNUM, "
    sSQL = sSQL & "FACTORY, "
    sSQL = sSQL & "OPECOND, "
    sSQL = sSQL & "MAICB, "
    sSQL = sSQL & "REGDATE, "
    sSQL = sSQL & "UPDDATE, "
    sSQL = sSQL & "HOLDBCB, "             'ΰ��ދ敪�@06/02/08 ooba
    sSQL = sSQL & "WFHOLDFLGCB, "         'WFΰ��ދ敪�@06/02/08 ooba
    sSQL = sSQL & "KBLKFLGCB, "           '�֘A��ۯ��׸ށ@08/01/31 ooba
    sSQL = sSQL & "PLANTCAT, "            '���� 07/09/05 SPK Tsutsumi Add
    sSQL = sSQL & "XTALCW, "
    sSQL = sSQL & "INPOSCW, "
    sSQL = sSQL & "nvl(TBKBNCW,'T') as TBKBNCW, "
    sSQL = sSQL & "SMPKBNCW, "
    sSQL = sSQL & "REPSMPLIDCW, "
    sSQL = sSQL & "HINBCW, "
    sSQL = sSQL & "REVNUMCW, "
    sSQL = sSQL & "FACTORYCW, "
    sSQL = sSQL & "OPECW, "
    sSQL = sSQL & "KTKBNCW, "
    sSQL = sSQL & "WFINDRSCW, "
    sSQL = sSQL & "WFINDOICW, "
    sSQL = sSQL & "WFINDB1CW, "
    sSQL = sSQL & "WFINDB2CW, "
    sSQL = sSQL & "WFINDB3CW, "
    sSQL = sSQL & "WFINDL1CW, "
    sSQL = sSQL & "WFINDL2CW, "
    sSQL = sSQL & "WFINDL3CW, "
    sSQL = sSQL & "WFINDL4CW, "
    sSQL = sSQL & "WFINDDSCW, "
    sSQL = sSQL & "WFINDDZCW, "
    sSQL = sSQL & "WFINDSPCW, "
    sSQL = sSQL & "WFINDDO1CW, "
    sSQL = sSQL & "WFINDDO2CW, "
    sSQL = sSQL & "WFINDDO3CW, "
    sSQL = sSQL & "WFINDAOICW, "
    sSQL = sSQL & "WFINDGDCW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    sSQL = sSQL & "EPINDB1CW, "
    sSQL = sSQL & "EPINDB2CW, "
    sSQL = sSQL & "EPINDB3CW, "
    sSQL = sSQL & "EPINDL1CW, "
    sSQL = sSQL & "EPINDL2CW, "
    sSQL = sSQL & "EPINDL3CW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    sSQL = sSQL & "WFRESRS1CW, "
    sSQL = sSQL & "WFRESOICW, "
    sSQL = sSQL & "WFRESB1CW, "
    sSQL = sSQL & "WFRESB2CW, "
    sSQL = sSQL & "WFRESB3CW, "
    sSQL = sSQL & "WFRESL1CW, "
    sSQL = sSQL & "WFRESL2CW, "
    sSQL = sSQL & "WFRESL3CW, "
    sSQL = sSQL & "WFRESL4CW, "
    sSQL = sSQL & "WFRESDSCW, "
    sSQL = sSQL & "WFRESDZCW, "
    sSQL = sSQL & "WFRESSPCW, "
    sSQL = sSQL & "WFRESDO1CW, "
    sSQL = sSQL & "WFRESDO2CW, "
    sSQL = sSQL & "WFRESDO3CW, "
    sSQL = sSQL & "WFRESAOICW, "
    sSQL = sSQL & "WFRESGDCW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    sSQL = sSQL & "EPRESB1CW, "
    sSQL = sSQL & "EPRESB2CW, "
    sSQL = sSQL & "EPRESB3CW, "
    sSQL = sSQL & "EPRESL1CW, "
    sSQL = sSQL & "EPRESL2CW, "
    sSQL = sSQL & "EPRESL3CW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    sSQL = sSQL & "WFSMPLIDRSCW, "
    sSQL = sSQL & "WFSMPLIDOICW, "
    sSQL = sSQL & "WFSMPLIDB1CW, "
    sSQL = sSQL & "WFSMPLIDB2CW, "
    sSQL = sSQL & "WFSMPLIDB3CW, "
    sSQL = sSQL & "WFSMPLIDL1CW, "
    sSQL = sSQL & "WFSMPLIDL2CW, "
    sSQL = sSQL & "WFSMPLIDL3CW, "
    sSQL = sSQL & "WFSMPLIDL4CW, "
    sSQL = sSQL & "WFSMPLIDDSCW, "
    sSQL = sSQL & "WFSMPLIDDZCW, "
    sSQL = sSQL & "WFSMPLIDSPCW, "
    sSQL = sSQL & "WFSMPLIDDO1CW, "
    sSQL = sSQL & "WFSMPLIDDO2CW, "
    sSQL = sSQL & "WFSMPLIDDO3CW, "
    sSQL = sSQL & "WFSMPLIDAOICW, "
    sSQL = sSQL & "WFSMPLIDGDCW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    sSQL = sSQL & "EPSMPLIDB1CW, "
    sSQL = sSQL & "EPSMPLIDB2CW, "
    sSQL = sSQL & "EPSMPLIDB3CW, "
    sSQL = sSQL & "EPSMPLIDL1CW, "
    sSQL = sSQL & "EPSMPLIDL2CW, "
    sSQL = sSQL & "EPSMPLIDL3CW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    sSQL = sSQL & "WFHSGDCW, "
    sSQL = sSQL & "TDAYCW, "
    sSQL = sSQL & "KDAYCW "
    sSQL = sSQL & "from "

    sSQL = sSQL & "(select "
    sSQL = sSQL & "xtalcb as CRYNUM, "
    sSQL = sSQL & "inposcb as INGOTPOS, "
'    sSql = sSql & "rlencb as LENGTH, "
    sSQL = sSQL & "LENCB as LENGTH, "         '���_�����������@06/11/09 ooba
    sSQL = sSQL & "sxlidcb as SXLID, "
    sSQL = sSQL & "' ' as KRPROCCD, "
    sSQL = sSQL & "gnwkntcb as NOWPROC, "
    sSQL = sSQL & "' ' as LPKRPROCCD, "
    sSQL = sSQL & "newkntcb as LASTPASS, "
    sSQL = sSQL & "livkcb as DELCLS, "
    sSQL = sSQL & "lstccb as LSTATCLS, "
    sSQL = sSQL & "sholdclscb HOLDCLS, "
    sSQL = sSQL & "hinbcb as HINBAN, "
    sSQL = sSQL & "revnumcb as REVNUM, "
    sSQL = sSQL & "factorycb as FACTORY, "
    sSQL = sSQL & "opecb as OPECOND, "
    sSQL = sSQL & "MAICB, "
    sSQL = sSQL & "tdaycb as REGDATE, "
    sSQL = sSQL & "kdaycb as UPDDATE, "
    sSQL = sSQL & "HOLDBCB, "
    sSQL = sSQL & "WFHOLDFLGCB, "
    sSQL = sSQL & "KBLKFLGCB, "           '�֘A��ۯ��׸ށ@08/01/31 ooba
    sSQL = sSQL & "PLANTCATCB  as PLANTCAT "
    sSQL = sSQL & "from XSDCB "
    sSQL = sSQL & "where GNWKNTCB = '" & inNowPorc & "' "
    sSQL = sSQL & "and livkcb = '0' "

    If sCmbMukesaki <> "ALL" Then
        sSQL = sSQL & "   AND PLANTCATCB      = '" & sCmbMukesaki & "'"
    End If
    sSQL = sSQL & "), "

    sSQL = sSQL & "(select "
    sSQL = sSQL & "SXLIDCW, "
    sSQL = sSQL & "XTALCW, "
    sSQL = sSQL & "INPOSCW, "
    sSQL = sSQL & "TBKBNCW, "
    sSQL = sSQL & "SMPKBNCW, "
    sSQL = sSQL & "REPSMPLIDCW, "
    sSQL = sSQL & "HINBCW, "
    sSQL = sSQL & "REVNUMCW, "
    sSQL = sSQL & "FACTORYCW, "
    sSQL = sSQL & "OPECW, "
    sSQL = sSQL & "KTKBNCW, "
    sSQL = sSQL & "WFINDRSCW, "
    sSQL = sSQL & "WFINDOICW, "
    sSQL = sSQL & "WFINDB1CW, "
    sSQL = sSQL & "WFINDB2CW, "
    sSQL = sSQL & "WFINDB3CW, "
    sSQL = sSQL & "WFINDL1CW, "
    sSQL = sSQL & "WFINDL2CW, "
    sSQL = sSQL & "WFINDL3CW, "
    sSQL = sSQL & "WFINDL4CW, "
    sSQL = sSQL & "WFINDDSCW, "
    sSQL = sSQL & "WFINDDZCW, "
    sSQL = sSQL & "WFINDSPCW, "
    sSQL = sSQL & "WFINDDO1CW, "
    sSQL = sSQL & "WFINDDO2CW, "
    sSQL = sSQL & "WFINDDO3CW, "
    sSQL = sSQL & "WFINDAOICW, "
    sSQL = sSQL & "WFINDGDCW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    sSQL = sSQL & "EPINDB1CW, "
    sSQL = sSQL & "EPINDB2CW, "
    sSQL = sSQL & "EPINDB3CW, "
    sSQL = sSQL & "EPINDL1CW, "
    sSQL = sSQL & "EPINDL2CW, "
    sSQL = sSQL & "EPINDL3CW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    sSQL = sSQL & "WFRESRS1CW, "
    sSQL = sSQL & "WFRESOICW, "
    sSQL = sSQL & "WFRESB1CW, "
    sSQL = sSQL & "WFRESB2CW, "
    sSQL = sSQL & "WFRESB3CW, "
    sSQL = sSQL & "WFRESL1CW, "
    sSQL = sSQL & "WFRESL2CW, "
    sSQL = sSQL & "WFRESL3CW, "
    sSQL = sSQL & "WFRESL4CW, "
    sSQL = sSQL & "WFRESDSCW, "
    sSQL = sSQL & "WFRESDZCW, "
    sSQL = sSQL & "WFRESSPCW, "
    sSQL = sSQL & "WFRESDO1CW, "
    sSQL = sSQL & "WFRESDO2CW, "
    sSQL = sSQL & "WFRESDO3CW, "
    sSQL = sSQL & "WFRESAOICW, "
    sSQL = sSQL & "WFRESGDCW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    sSQL = sSQL & "EPRESB1CW, "
    sSQL = sSQL & "EPRESB2CW, "
    sSQL = sSQL & "EPRESB3CW, "
    sSQL = sSQL & "EPRESL1CW, "
    sSQL = sSQL & "EPRESL2CW, "
    sSQL = sSQL & "EPRESL3CW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    sSQL = sSQL & "WFSMPLIDRSCW, "
    sSQL = sSQL & "WFSMPLIDOICW, "
    sSQL = sSQL & "WFSMPLIDB1CW, "
    sSQL = sSQL & "WFSMPLIDB2CW, "
    sSQL = sSQL & "WFSMPLIDB3CW, "
    sSQL = sSQL & "WFSMPLIDL1CW, "
    sSQL = sSQL & "WFSMPLIDL2CW, "
    sSQL = sSQL & "WFSMPLIDL3CW, "
    sSQL = sSQL & "WFSMPLIDL4CW, "
    sSQL = sSQL & "WFSMPLIDDSCW, "
    sSQL = sSQL & "WFSMPLIDDZCW, "
    sSQL = sSQL & "WFSMPLIDSPCW, "
    sSQL = sSQL & "WFSMPLIDDO1CW, "
    sSQL = sSQL & "WFSMPLIDDO2CW, "
    sSQL = sSQL & "WFSMPLIDDO3CW, "
    sSQL = sSQL & "WFSMPLIDAOICW, "
    sSQL = sSQL & "WFSMPLIDGDCW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    sSQL = sSQL & "EPSMPLIDB1CW, "
    sSQL = sSQL & "EPSMPLIDB2CW, "
    sSQL = sSQL & "EPSMPLIDB3CW, "
    sSQL = sSQL & "EPSMPLIDL1CW, "
    sSQL = sSQL & "EPSMPLIDL2CW, "
    sSQL = sSQL & "EPSMPLIDL3CW, "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    sSQL = sSQL & "WFHSGDCW, "
    sSQL = sSQL & "TDAYCW, "
    sSQL = sSQL & "KDAYCW "
    sSQL = sSQL & "from XSDCW "
    sSQL = sSQL & "where LIVKCW = '0' "
'Chg Start 2011/03/07 SMPK Miyata
'    sSQL = sSQL & "and TBKBNCW = 'T' "
    sSQL = sSQL & "and (TBKBNCW = 'T' OR TBKBNCW = 'B')"
'Chg End   2011/03/07 SMPK Miyata
'Add Start 2011/03/07 SMPK Miyata
    sSQL = sSQL & "union all "
    sSQL = sSQL & "select "
    sSQL = sSQL & "SXLIDCW, "
    sSQL = sSQL & "XTALCW, "
    sSQL = sSQL & "INPOSCW, "
    sSQL = sSQL & "TBKBNCW, "
    sSQL = sSQL & "SMPKBNCW, "
    sSQL = sSQL & "REPSMPLIDCW, "
    sSQL = sSQL & "HINBCW, "
    sSQL = sSQL & "REVNUMCW, "
    sSQL = sSQL & "FACTORYCW, "
    sSQL = sSQL & "OPECW, "
    sSQL = sSQL & "KTKBNCW, "
    sSQL = sSQL & "WFINDRSCW, "
    sSQL = sSQL & "WFINDOICW, "
    sSQL = sSQL & "WFINDB1CW, "
    sSQL = sSQL & "WFINDB2CW, "
    sSQL = sSQL & "WFINDB3CW, "
    sSQL = sSQL & "WFINDL1CW, "
    sSQL = sSQL & "WFINDL2CW, "
    sSQL = sSQL & "WFINDL3CW, "
    sSQL = sSQL & "WFINDL4CW, "
    sSQL = sSQL & "WFINDDSCW, "
    sSQL = sSQL & "WFINDDZCW, "
    sSQL = sSQL & "WFINDSPCW, "
    sSQL = sSQL & "WFINDDO1CW, "
    sSQL = sSQL & "WFINDDO2CW, "
    sSQL = sSQL & "WFINDDO3CW, "
    sSQL = sSQL & "WFINDAOICW, "
    sSQL = sSQL & "WFINDGDCW, "
    sSQL = sSQL & "EPINDB1CW, "
    sSQL = sSQL & "EPINDB2CW, "
    sSQL = sSQL & "EPINDB3CW, "
    sSQL = sSQL & "EPINDL1CW, "
    sSQL = sSQL & "EPINDL2CW, "
    sSQL = sSQL & "EPINDL3CW, "
    sSQL = sSQL & "WFRESRS1CW, "
    sSQL = sSQL & "WFRESOICW, "
    sSQL = sSQL & "WFRESB1CW, "
    sSQL = sSQL & "WFRESB2CW, "
    sSQL = sSQL & "WFRESB3CW, "
    sSQL = sSQL & "WFRESL1CW, "
    sSQL = sSQL & "WFRESL2CW, "
    sSQL = sSQL & "WFRESL3CW, "
    sSQL = sSQL & "WFRESL4CW, "
    sSQL = sSQL & "WFRESDSCW, "
    sSQL = sSQL & "WFRESDZCW, "
    sSQL = sSQL & "WFRESSPCW, "
    sSQL = sSQL & "WFRESDO1CW, "
    sSQL = sSQL & "WFRESDO2CW, "
    sSQL = sSQL & "WFRESDO3CW, "
    sSQL = sSQL & "WFRESAOICW, "
    sSQL = sSQL & "WFRESGDCW, "
    sSQL = sSQL & "EPRESB1CW, "
    sSQL = sSQL & "EPRESB2CW, "
    sSQL = sSQL & "EPRESB3CW, "
    sSQL = sSQL & "EPRESL1CW, "
    sSQL = sSQL & "EPRESL2CW, "
    sSQL = sSQL & "EPRESL3CW, "
    sSQL = sSQL & "WFSMPLIDRSCW, "
    sSQL = sSQL & "WFSMPLIDOICW, "
    sSQL = sSQL & "WFSMPLIDB1CW, "
    sSQL = sSQL & "WFSMPLIDB2CW, "
    sSQL = sSQL & "WFSMPLIDB3CW, "
    sSQL = sSQL & "WFSMPLIDL1CW, "
    sSQL = sSQL & "WFSMPLIDL2CW, "
    sSQL = sSQL & "WFSMPLIDL3CW, "
    sSQL = sSQL & "WFSMPLIDL4CW, "
    sSQL = sSQL & "WFSMPLIDDSCW, "
    sSQL = sSQL & "WFSMPLIDDZCW, "
    sSQL = sSQL & "WFSMPLIDSPCW, "
    sSQL = sSQL & "WFSMPLIDDO1CW, "
    sSQL = sSQL & "WFSMPLIDDO2CW, "
    sSQL = sSQL & "WFSMPLIDDO3CW, "
    sSQL = sSQL & "WFSMPLIDAOICW, "
    sSQL = sSQL & "WFSMPLIDGDCW, "
    sSQL = sSQL & "EPSMPLIDB1CW, "
    sSQL = sSQL & "EPSMPLIDB2CW, "
    sSQL = sSQL & "EPSMPLIDB3CW, "
    sSQL = sSQL & "EPSMPLIDL1CW, "
    sSQL = sSQL & "EPSMPLIDL2CW, "
    sSQL = sSQL & "EPSMPLIDL3CW, "
    sSQL = sSQL & "WFHSGDCW, "
    sSQL = sSQL & "TDAYCW, "
    sSQL = sSQL & "KDAYCW "
    sSQL = sSQL & "from XSDCW_1 "
    sSQL = sSQL & "where LIVKCW = '0' "
    sSQL = sSQL & "and TBKBNCW = 'C'"
'Add End   2011/03/07 SMPK Miyata
    sSQL = sSQL & ") "
    sSQL = sSQL & "where SXLID = SXLIDCW(+) "

'Del Start 2011/03/07 SMPK Miyata
'    sSQL = sSQL & "union all "
'
'    sSQL = sSQL & "select "
'    sSQL = sSQL & "CRYNUM, "
'    sSQL = sSQL & "INGOTPOS, "
'    sSQL = sSQL & "LENGTH, "
'    sSQL = sSQL & "SXLID, "
'    sSQL = sSQL & "KRPROCCD, "
'    sSQL = sSQL & "NOWPROC, "
'    sSQL = sSQL & "LPKRPROCCD, "
'    sSQL = sSQL & "LASTPASS, "
'    sSQL = sSQL & "DELCLS, "
'    sSQL = sSQL & "LSTATCLS, "
'    sSQL = sSQL & "HOLDCLS, "
'    sSQL = sSQL & "HINBAN, "
'    sSQL = sSQL & "REVNUM, "
'    sSQL = sSQL & "FACTORY, "
'    sSQL = sSQL & "OPECOND, "
'    sSQL = sSQL & "MAICB, "
'    sSQL = sSQL & "REGDATE, "
'    sSQL = sSQL & "UPDDATE, "
'    sSQL = sSQL & "HOLDBCB, "             'ΰ��ދ敪�@06/02/08 ooba
'    sSQL = sSQL & "WFHOLDFLGCB, "         'WFΰ��ދ敪�@06/02/08 ooba
'    sSQL = sSQL & "KBLKFLGCB, "           '�֘A��ۯ��׸ށ@08/01/31 ooba
'    sSQL = sSQL & "PLANTCAT, "            '���� 07/09/05 SPK Tsutsumi Add
'    sSQL = sSQL & "XTALCW, "
'    sSQL = sSQL & "INPOSCW, "
'    sSQL = sSQL & "nvl(TBKBNCW,'B') as TBKBNCW, "
'    sSQL = sSQL & "SMPKBNCW, "
'    sSQL = sSQL & "REPSMPLIDCW, "
'    sSQL = sSQL & "HINBCW, "
'    sSQL = sSQL & "REVNUMCW, "
'    sSQL = sSQL & "FACTORYCW, "
'    sSQL = sSQL & "OPECW, "
'    sSQL = sSQL & "KTKBNCW, "
'    sSQL = sSQL & "WFINDRSCW, "
'    sSQL = sSQL & "WFINDOICW, "
'    sSQL = sSQL & "WFINDB1CW, "
'    sSQL = sSQL & "WFINDB2CW, "
'    sSQL = sSQL & "WFINDB3CW, "
'    sSQL = sSQL & "WFINDL1CW, "
'    sSQL = sSQL & "WFINDL2CW, "
'    sSQL = sSQL & "WFINDL3CW, "
'    sSQL = sSQL & "WFINDL4CW, "
'    sSQL = sSQL & "WFINDDSCW, "
'    sSQL = sSQL & "WFINDDZCW, "
'    sSQL = sSQL & "WFINDSPCW, "
'    sSQL = sSQL & "WFINDDO1CW, "
'    sSQL = sSQL & "WFINDDO2CW, "
'    sSQL = sSQL & "WFINDDO3CW, "
'    sSQL = sSQL & "WFINDAOICW, "
'    sSQL = sSQL & "WFINDGDCW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
'    sSQL = sSQL & "EPINDB1CW, "
'    sSQL = sSQL & "EPINDB2CW, "
'    sSQL = sSQL & "EPINDB3CW, "
'    sSQL = sSQL & "EPINDL1CW, "
'    sSQL = sSQL & "EPINDL2CW, "
'    sSQL = sSQL & "EPINDL3CW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'    sSQL = sSQL & "WFRESRS1CW, "
'    sSQL = sSQL & "WFRESOICW, "
'    sSQL = sSQL & "WFRESB1CW, "
'    sSQL = sSQL & "WFRESB2CW, "
'    sSQL = sSQL & "WFRESB3CW, "
'    sSQL = sSQL & "WFRESL1CW, "
'    sSQL = sSQL & "WFRESL2CW, "
'    sSQL = sSQL & "WFRESL3CW, "
'    sSQL = sSQL & "WFRESL4CW, "
'    sSQL = sSQL & "WFRESDSCW, "
'    sSQL = sSQL & "WFRESDZCW, "
'    sSQL = sSQL & "WFRESSPCW, "
'    sSQL = sSQL & "WFRESDO1CW, "
'    sSQL = sSQL & "WFRESDO2CW, "
'    sSQL = sSQL & "WFRESDO3CW, "
'    sSQL = sSQL & "WFRESAOICW, "
'    sSQL = sSQL & "WFRESGDCW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
'    sSQL = sSQL & "EPRESB1CW, "
'    sSQL = sSQL & "EPRESB2CW, "
'    sSQL = sSQL & "EPRESB3CW, "
'    sSQL = sSQL & "EPRESL1CW, "
'    sSQL = sSQL & "EPRESL2CW, "
'    sSQL = sSQL & "EPRESL3CW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'    sSQL = sSQL & "WFSMPLIDRSCW, "
'    sSQL = sSQL & "WFSMPLIDOICW, "
'    sSQL = sSQL & "WFSMPLIDB1CW, "
'    sSQL = sSQL & "WFSMPLIDB2CW, "
'    sSQL = sSQL & "WFSMPLIDB3CW, "
'    sSQL = sSQL & "WFSMPLIDL1CW, "
'    sSQL = sSQL & "WFSMPLIDL2CW, "
'    sSQL = sSQL & "WFSMPLIDL3CW, "
'    sSQL = sSQL & "WFSMPLIDL4CW, "
'    sSQL = sSQL & "WFSMPLIDDSCW, "
'    sSQL = sSQL & "WFSMPLIDDZCW, "
'    sSQL = sSQL & "WFSMPLIDSPCW, "
'    sSQL = sSQL & "WFSMPLIDDO1CW, "
'    sSQL = sSQL & "WFSMPLIDDO2CW, "
'    sSQL = sSQL & "WFSMPLIDDO3CW, "
'    sSQL = sSQL & "WFSMPLIDAOICW, "
'    sSQL = sSQL & "WFSMPLIDGDCW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
'    sSQL = sSQL & "EPSMPLIDB1CW, "
'    sSQL = sSQL & "EPSMPLIDB2CW, "
'    sSQL = sSQL & "EPSMPLIDB3CW, "
'    sSQL = sSQL & "EPSMPLIDL1CW, "
'    sSQL = sSQL & "EPSMPLIDL2CW, "
'    sSQL = sSQL & "EPSMPLIDL3CW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'    sSQL = sSQL & "WFHSGDCW, "
'    sSQL = sSQL & "TDAYCW, "
'    sSQL = sSQL & "KDAYCW "
'    sSQL = sSQL & "from "
'
'    sSQL = sSQL & "(select "
'    sSQL = sSQL & "xtalcb as CRYNUM, "
'    sSQL = sSQL & "inposcb as INGOTPOS, "
''    sSql = sSql & "rlencb as LENGTH, "
'    sSQL = sSQL & "LENCB as LENGTH, "         '���_�����������@06/11/09 ooba
'    sSQL = sSQL & "sxlidcb as SXLID, "
'    sSQL = sSQL & "' ' as KRPROCCD, "
'    sSQL = sSQL & "gnwkntcb as NOWPROC, "
'    sSQL = sSQL & "' ' as LPKRPROCCD, "
'    sSQL = sSQL & "newkntcb as LASTPASS, "
'    sSQL = sSQL & "livkcb as DELCLS, "
'    sSQL = sSQL & "lstccb as LSTATCLS, "
'    sSQL = sSQL & "sholdclscb HOLDCLS, "
'    sSQL = sSQL & "hinbcb as HINBAN, "
'    sSQL = sSQL & "revnumcb as REVNUM, "
'    sSQL = sSQL & "factorycb as FACTORY, "
'    sSQL = sSQL & "opecb as OPECOND, "
'    sSQL = sSQL & "MAICB, "
'    sSQL = sSQL & "tdaycb as REGDATE, "
'    sSQL = sSQL & "kdaycb as UPDDATE, "
'    sSQL = sSQL & "HOLDBCB, "
'' 2007/09/04 SPK Tsutsumi Add Start
'    sSQL = sSQL & "WFHOLDFLGCB, "
'    sSQL = sSQL & "KBLKFLGCB, "           '�֘A��ۯ��׸ށ@08/01/31 ooba
'    sSQL = sSQL & "PLANTCATCB as PLANTCAT "
'' 2007/09/04 SPK Tsutsumi Add End
'    sSQL = sSQL & "from XSDCB "
'    sSQL = sSQL & "where GNWKNTCB = '" & inNowPorc & "' "
'    sSQL = sSQL & "and livkcb = '0' "
'
'' 2007/09/04 SPK Tsutsumi Add Start
'    If sCmbMukesaki <> "ALL" Then
'        sSQL = sSQL & "   AND PLANTCATCB      = '" & sCmbMukesaki & "'"
'    End If
'' 2007/09/04 SPK Tsutsumi Add End
'
'    sSQL = sSQL & "), "
'
'    sSQL = sSQL & "(select "
'    sSQL = sSQL & "SXLIDCW, "
'    sSQL = sSQL & "XTALCW, "
'    sSQL = sSQL & "INPOSCW, "
'    sSQL = sSQL & "TBKBNCW, "
'    sSQL = sSQL & "SMPKBNCW, "
'    sSQL = sSQL & "REPSMPLIDCW, "
'    sSQL = sSQL & "HINBCW, "
'    sSQL = sSQL & "REVNUMCW, "
'    sSQL = sSQL & "FACTORYCW, "
'    sSQL = sSQL & "OPECW, "
'    sSQL = sSQL & "KTKBNCW, "
'    sSQL = sSQL & "WFINDRSCW, "
'    sSQL = sSQL & "WFINDOICW, "
'    sSQL = sSQL & "WFINDB1CW, "
'    sSQL = sSQL & "WFINDB2CW, "
'    sSQL = sSQL & "WFINDB3CW, "
'    sSQL = sSQL & "WFINDL1CW, "
'    sSQL = sSQL & "WFINDL2CW, "
'    sSQL = sSQL & "WFINDL3CW, "
'    sSQL = sSQL & "WFINDL4CW, "
'    sSQL = sSQL & "WFINDDSCW, "
'    sSQL = sSQL & "WFINDDZCW, "
'    sSQL = sSQL & "WFINDSPCW, "
'    sSQL = sSQL & "WFINDDO1CW, "
'    sSQL = sSQL & "WFINDDO2CW, "
'    sSQL = sSQL & "WFINDDO3CW, "
'    sSQL = sSQL & "WFINDAOICW, "
'    sSQL = sSQL & "WFINDGDCW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
'    sSQL = sSQL & "EPINDB1CW, "
'    sSQL = sSQL & "EPINDB2CW, "
'    sSQL = sSQL & "EPINDB3CW, "
'    sSQL = sSQL & "EPINDL1CW, "
'    sSQL = sSQL & "EPINDL2CW, "
'    sSQL = sSQL & "EPINDL3CW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'    sSQL = sSQL & "WFRESRS1CW, "
'    sSQL = sSQL & "WFRESOICW, "
'    sSQL = sSQL & "WFRESB1CW, "
'    sSQL = sSQL & "WFRESB2CW, "
'    sSQL = sSQL & "WFRESB3CW, "
'    sSQL = sSQL & "WFRESL1CW, "
'    sSQL = sSQL & "WFRESL2CW, "
'    sSQL = sSQL & "WFRESL3CW, "
'    sSQL = sSQL & "WFRESL4CW, "
'    sSQL = sSQL & "WFRESDSCW, "
'    sSQL = sSQL & "WFRESDZCW, "
'    sSQL = sSQL & "WFRESSPCW, "
'    sSQL = sSQL & "WFRESDO1CW, "
'    sSQL = sSQL & "WFRESDO2CW, "
'    sSQL = sSQL & "WFRESDO3CW, "
'    sSQL = sSQL & "WFRESAOICW, "
'    sSQL = sSQL & "WFRESGDCW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
'    sSQL = sSQL & "EPRESB1CW, "
'    sSQL = sSQL & "EPRESB2CW, "
'    sSQL = sSQL & "EPRESB3CW, "
'    sSQL = sSQL & "EPRESL1CW, "
'    sSQL = sSQL & "EPRESL2CW, "
'    sSQL = sSQL & "EPRESL3CW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'    sSQL = sSQL & "WFSMPLIDRSCW, "
'    sSQL = sSQL & "WFSMPLIDOICW, "
'    sSQL = sSQL & "WFSMPLIDB1CW, "
'    sSQL = sSQL & "WFSMPLIDB2CW, "
'    sSQL = sSQL & "WFSMPLIDB3CW, "
'    sSQL = sSQL & "WFSMPLIDL1CW, "
'    sSQL = sSQL & "WFSMPLIDL2CW, "
'    sSQL = sSQL & "WFSMPLIDL3CW, "
'    sSQL = sSQL & "WFSMPLIDL4CW, "
'    sSQL = sSQL & "WFSMPLIDDSCW, "
'    sSQL = sSQL & "WFSMPLIDDZCW, "
'    sSQL = sSQL & "WFSMPLIDSPCW, "
'    sSQL = sSQL & "WFSMPLIDDO1CW, "
'    sSQL = sSQL & "WFSMPLIDDO2CW, "
'    sSQL = sSQL & "WFSMPLIDDO3CW, "
'    sSQL = sSQL & "WFSMPLIDAOICW, "
'    sSQL = sSQL & "WFSMPLIDGDCW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
'    sSQL = sSQL & "EPSMPLIDB1CW, "
'    sSQL = sSQL & "EPSMPLIDB2CW, "
'    sSQL = sSQL & "EPSMPLIDB3CW, "
'    sSQL = sSQL & "EPSMPLIDL1CW, "
'    sSQL = sSQL & "EPSMPLIDL2CW, "
'    sSQL = sSQL & "EPSMPLIDL3CW, "
''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'    sSQL = sSQL & "WFHSGDCW, "
'    sSQL = sSQL & "TDAYCW, "
'    sSQL = sSQL & "KDAYCW "
'    sSQL = sSQL & "from XSDCW "
'    sSQL = sSQL & "where LIVKCW = '0' "
'    sSQL = sSQL & "and TBKBNCW = 'B' "
'    sSQL = sSQL & ") "
'    sSQL = sSQL & "where SXLID = SXLIDCW(+) "
'Del End   2011/03/07 SMPK Miyata

    sSQL = sSQL & "order by CRYNUM,INGOTPOS,TBKBNCW DESC "

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)

    intSxlCount = rs.RecordCount

    '���R�[�h0��������I��
    If intSxlCount = 0 Then
        rs.Close
        ReDim sxl(0)
        GoTo proc_exit
    End If

    j = 0
    sNullSXLID = ""
'Add Start 2011/03/07 SMPK Miyata
    lngCmpCnt = 0               '�T���v���� ������
    ReDim Preserve sxl(0)       'SXL�Ǘ�    ������
'Add End   2011/03/07 SMPK Miyata
    
'Chg Start 2011/03/07 SMPK Miyata
'    For i = 1 To intSxlCount Step 2
    For i = 1 To intSxlCount
'Chg End   2011/03/07 SMPK Miyata

'Add Start 2011/03/07 SMPK Miyata
        If rs("SXLID") <> sxl(j).SXLIDCA Then
'Add End   2011/03/07 SMPK Miyata

            lngNullCnt = 0
            j = j + 1
            ReDim Preserve sxl(j)
            ReDim Preserve WFJudgExecOkFlag(j)
            WFJudgExecOkFlag(j) = True '�\���F�f�t�H���g�ݒ�(��)
            lngCmpCnt = 0               '�T���v����     'Add 2011/03/07 SMPK Miyata
                    
            With sxl(j)
                If IsNull(rs("CRYNUM")) = False Then .CRYNUMCA = rs("CRYNUM") Else lngNullCnt = lngNullCnt + 1        ' �����ԍ�
                If IsNull(rs("INGOTPOS")) = False Then .INPOSCA = rs("INGOTPOS") Else lngNullCnt = lngNullCnt + 1     ' �������J�n�ʒu
                If IsNull(rs("LENGTH")) = False Then .GNLCA = rs("LENGTH") Else lngNullCnt = lngNullCnt + 1           ' ����
                If IsNull(rs("SXLID")) = False Then .SXLIDCA = rs("SXLID") Else lngNullCnt = lngNullCnt + 1           ' SXLID
                If IsNull(rs("NOWPROC")) = False Then .NOWPROC = rs("NOWPROC") Else lngNullCnt = lngNullCnt + 1       ' ���ݍH��
                If IsNull(rs("LASTPASS")) = False Then .NEWKNTCA = rs("LASTPASS") Else lngNullCnt = lngNullCnt + 1    ' �ŏI�ʉߍH��
                If IsNull(rs("DELCLS")) = False Then .SAKJCA = rs("DELCLS") Else lngNullCnt = lngNullCnt + 1          ' �폜�敪
                If IsNull(rs("LSTATCLS")) = False Then .LSTATBCA = rs("LSTATCLS") Else lngNullCnt = lngNullCnt + 1    ' �ŏI��ԋ敪
                If IsNull(rs("HOLDCLS")) = False Then .HOLDBCA = rs("HOLDCLS") Else lngNullCnt = lngNullCnt + 1       ' �z�[���h�敪
                If IsNull(rs("HINBAN")) = False Then .HINBCA = rs("HINBAN") Else lngNullCnt = lngNullCnt + 1          ' �i��
                If IsNull(rs("REVNUM")) = False Then .REVNUMCA = rs("REVNUM") Else lngNullCnt = lngNullCnt + 1        ' ���i�ԍ������ԍ�
                If IsNull(rs("FACTORY")) = False Then .FACTORYCA = rs("FACTORY") Else lngNullCnt = lngNullCnt + 1     ' �H��
                If IsNull(rs("OPECOND")) = False Then .OPECA = rs("OPECOND") Else lngNullCnt = lngNullCnt + 1         ' ���Ə���
                If IsNull(rs("MAICB")) = False Then .MAICB = rs("MAICB") Else lngNullCnt = lngNullCnt + 1             ' ����
                If IsNull(rs("REGDATE")) = False Then .TDAYCB = rs("REGDATE") Else lngNullCnt = lngNullCnt + 1        ' �o�^���t
                If IsNull(rs("UPDDATE")) = False Then .KDAYCA = rs("UPDDATE") Else lngNullCnt = lngNullCnt + 1        ' �X�V���t
                If IsNull(rs("HOLDBCB")) = False Then .HOLDBCB = rs("HOLDBCB") Else .HOLDBCB = " "                  ' ΰ��ދ敪�@06/02/08 ooba
                If IsNull(rs("WFHOLDFLGCB")) = False Then .WFHOLDFLGCB = rs("WFHOLDFLGCB") Else .WFHOLDFLGCB = " "  ' WFΰ��ދ敪�@06/02/08 ooba
                If IsNull(rs("KBLKFLGCB")) = False Then .KANREN = rs("KBLKFLGCB") Else .KANREN = " "            ' �֘A��ۯ��L���@08/01/31 ooba
    
                ' ���� 07/09/04 SPK Tsutsumi Add Start
                If IsNull(rs("PLANTCAT")) = False Then
                    For k = 0 To UBound(s_MukesakiBase)
                        If s_MukesakiBase(k).sMukeCode = rs("PLANTCAT") Then
                           .PLANTCAT = s_MukesakiBase(k).sMukeName
                           Exit For
                        End If
                    Next k
                Else
                    .PLANTCAT = " "
                End If
                ' ���� 07/09/04 SPK Tsutsumi Add end
            End With
        End If              'Add 2011/03/07 SMPK Miyata

'Add Start 2011/03/07 SMPK Miyata
        If rs("SXLID") = sxl(j).SXLIDCA Then
            lngCmpCnt = lngCmpCnt + 1       '�T���v�����J�E���g     'Add 2011/03/07 SMPK Miyata
'Add End   2011/03/07 SMPK Miyata

'Chg Start 2011/03/07 SMPK Miyata
'Chg End   2011/03/07 SMPK Miyata

'Chg Start 2011/03/07 SMPK Miyata
'        ReDim Preserve SXL(j).WFSMP(2)
'
'        For k = 1 To 2
'            Call Init_SXL_WFSMP(SXL(j).WFSMP(k))
'            With SXL(j).WFSMP(k)
            ReDim Preserve sxl(j).WFSMP(lngCmpCnt)
            With sxl(j).WFSMP(lngCmpCnt)
'Chg End   2011/03/07 SMPK Miyata

                If IsNull(rs("XTALCW")) = False Then .XTALCW = rs("XTALCW") Else lngNullCnt = lngNullCnt + 1                    ' �����ԍ�
                If IsNull(rs("INPOSCW")) = False Then .INPOSCW = rs("INPOSCW") Else lngNullCnt = lngNullCnt + 1                 ' �������ʒu
                If IsNull(rs("SMPKBNCW")) = False Then .SMPKBNCW = rs("SMPKBNCW") Else lngNullCnt = lngNullCnt + 1              ' �T���v���敪
                If IsNull(rs("REPSMPLIDCW")) = False Then .REPSMPLIDCW = rs("REPSMPLIDCW") Else lngNullCnt = lngNullCnt + 1     ' �T���v��ID
                If IsNull(rs("HINBCW")) = False Then .HINBCW = rs("HINBCW") Else lngNullCnt = lngNullCnt + 1                    ' �i��
                If IsNull(rs("REVNUMCW")) = False Then .REVNUMCW = rs("REVNUMCW") Else lngNullCnt = lngNullCnt + 1              ' ���i�ԍ������ԍ�
                If IsNull(rs("FACTORYCW")) = False Then .FACTORYCW = rs("FACTORYCW") Else lngNullCnt = lngNullCnt + 1           ' �H��
                If IsNull(rs("OPECW")) = False Then .OPECW = rs("OPECW") Else lngNullCnt = lngNullCnt + 1                       ' ���Ə���
                If IsNull(rs("KTKBNCW")) = False Then .KTKBNCW = rs("KTKBNCW") Else lngNullCnt = lngNullCnt + 1                 ' �m��敪

                If IsNull(rs("WFINDRSCW")) = False Then .WFINDRSCW = rs("WFINDRSCW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iRS)
                If IsNull(rs("WFINDOICW")) = False Then .WFINDOICW = rs("WFINDOICW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iOi)
                If IsNull(rs("WFINDB1CW")) = False Then .WFINDB1CW = rs("WFINDB1CW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iB1)
                If IsNull(rs("WFINDB2CW")) = False Then .WFINDB2CW = rs("WFINDB2CW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iB2�j
                If IsNull(rs("WFINDB3CW")) = False Then .WFINDB3CW = rs("WFINDB3CW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iB3)
                If IsNull(rs("WFINDL1CW")) = False Then .WFINDL1CW = rs("WFINDL1CW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iL1)
                If IsNull(rs("WFINDL2CW")) = False Then .WFINDL2CW = rs("WFINDL2CW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iL2)
                If IsNull(rs("WFINDL3CW")) = False Then .WFINDL3CW = rs("WFINDL3CW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iL3)
                If IsNull(rs("WFINDL4CW")) = False Then .WFINDL4CW = rs("WFINDL4CW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iL4)
                If IsNull(rs("WFINDDSCW")) = False Then .WFINDDSCW = rs("WFINDDSCW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iDS)
                If IsNull(rs("WFINDDZCW")) = False Then .WFINDDZCW = rs("WFINDDZCW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iDZ)
                If IsNull(rs("WFINDSPCW")) = False Then .WFINDSPCW = rs("WFINDSPCW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w���iSP)
                If IsNull(rs("WFINDDO1CW")) = False Then .WFINDDO1CW = rs("WFINDDO1CW") Else lngNullCnt = lngNullCnt + 1        ' WF�����w���iDO1)
                If IsNull(rs("WFINDDO2CW")) = False Then .WFINDDO2CW = rs("WFINDDO2CW") Else lngNullCnt = lngNullCnt + 1        ' WF�����w���iDO2)
                If IsNull(rs("WFINDDO3CW")) = False Then .WFINDDO3CW = rs("WFINDDO3CW") Else lngNullCnt = lngNullCnt + 1        ' WF�����w���iDO3)
                If IsNull(rs("WFINDAOICW")) = False Then .WFINDAOICW = rs("WFINDAOICW") Else lngNullCnt = lngNullCnt + 1        ' WF�����w�� (AOi)
                If IsNull(rs("WFINDGDCW")) = False Then .WFINDGDCW = rs("WFINDGDCW") Else lngNullCnt = lngNullCnt + 1           ' WF�����w�� (GD)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                If IsNull(rs("EPINDB1CW")) = False Then .EPINDB1CW = rs("EPINDB1CW") Else lngNullCnt = lngNullCnt + 1           ' EP�����w���iB1E)
                If IsNull(rs("EPINDB2CW")) = False Then .EPINDB2CW = rs("EPINDB2CW") Else lngNullCnt = lngNullCnt + 1           ' EP�����w���iB2E�j
                If IsNull(rs("EPINDB3CW")) = False Then .EPINDB3CW = rs("EPINDB3CW") Else lngNullCnt = lngNullCnt + 1           ' EP�����w���iB3E)
                If IsNull(rs("EPINDL1CW")) = False Then .EPINDL1CW = rs("EPINDL1CW") Else lngNullCnt = lngNullCnt + 1           ' EP�����w���iL1E)
                If IsNull(rs("EPINDL2CW")) = False Then .EPINDL2CW = rs("EPINDL2CW") Else lngNullCnt = lngNullCnt + 1           ' EP�����w���iL2E)
                If IsNull(rs("EPINDL3CW")) = False Then .EPINDL3CW = rs("EPINDL3CW") Else lngNullCnt = lngNullCnt + 1           ' EP�����w���iL3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

                If IsNull(rs("WFRESRS1CW")) = False Then .WFRESRS1CW = rs("WFRESRS1CW") Else lngNullCnt = lngNullCnt + 1        ' WF�������сiRS)
                If IsNull(rs("WFRESOICW")) = False Then .WFRESOICW = rs("WFRESOICW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiOi)
                If IsNull(rs("WFRESB1CW")) = False Then .WFRESB1CW = rs("WFRESB1CW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiB1)
                If IsNull(rs("WFRESB2CW")) = False Then .WFRESB2CW = rs("WFRESB2CW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiB2�j
                If IsNull(rs("WFRESB3CW")) = False Then .WFRESB3CW = rs("WFRESB3CW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiB3)
                If IsNull(rs("WFRESL1CW")) = False Then .WFRESL1CW = rs("WFRESL1CW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiL1)
                If IsNull(rs("WFRESL2CW")) = False Then .WFRESL2CW = rs("WFRESL2CW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiL2)
                If IsNull(rs("WFRESL3CW")) = False Then .WFRESL3CW = rs("WFRESL3CW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiL3)
                If IsNull(rs("WFRESL4CW")) = False Then .WFRESL4CW = rs("WFRESL4CW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiL4)
                If IsNull(rs("WFRESDSCW")) = False Then .WFRESDSCW = rs("WFRESDSCW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiDS)
                If IsNull(rs("WFRESDZCW")) = False Then .WFRESDZCW = rs("WFRESDZCW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiDZ)
                If IsNull(rs("WFRESSPCW")) = False Then .WFRESSPCW = rs("WFRESSPCW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiSP)
                If IsNull(rs("WFRESDO1CW")) = False Then .WFRESDO1CW = rs("WFRESDO1CW") Else lngNullCnt = lngNullCnt + 1        ' WF�������сiDO1)
                If IsNull(rs("WFRESDO2CW")) = False Then .WFRESDO2CW = rs("WFRESDO2CW") Else lngNullCnt = lngNullCnt + 1        ' WF�������сiDO2)
                If IsNull(rs("WFRESDO3CW")) = False Then .WFRESDO3CW = rs("WFRESDO3CW") Else lngNullCnt = lngNullCnt + 1        ' WF�������сiDO3)
                If IsNull(rs("WFRESAOICW")) = False Then .WFRESAOICW = rs("WFRESAOICW") Else lngNullCnt = lngNullCnt + 1        ' WF�������сiAOi)
                If IsNull(rs("WFRESGDCW")) = False Then .WFRESGDCW = rs("WFRESGDCW") Else lngNullCnt = lngNullCnt + 1           ' WF�������сiGD)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                If IsNull(rs("EPRESB1CW")) = False Then .EPRESB1CW = rs("EPRESB1CW") Else lngNullCnt = lngNullCnt + 1           ' EP�������сiB1E)
                If IsNull(rs("EPRESB2CW")) = False Then .EPRESB2CW = rs("EPRESB2CW") Else lngNullCnt = lngNullCnt + 1           ' EP�������сiB2E�j
                If IsNull(rs("EPRESB3CW")) = False Then .EPRESB3CW = rs("EPRESB3CW") Else lngNullCnt = lngNullCnt + 1           ' EP�������сiB3E)
                If IsNull(rs("EPRESL1CW")) = False Then .EPRESL1CW = rs("EPRESL1CW") Else lngNullCnt = lngNullCnt + 1           ' EP�������сiL1E)
                If IsNull(rs("EPRESL2CW")) = False Then .EPRESL2CW = rs("EPRESL2CW") Else lngNullCnt = lngNullCnt + 1           ' EP�������сiL2E)
                If IsNull(rs("EPRESL3CW")) = False Then .EPRESL3CW = rs("EPRESL3CW") Else lngNullCnt = lngNullCnt + 1           ' EP�������сiL3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

                If IsNull(rs("WFSMPLIDRSCW")) = False Then .WFSMPLIDRSCW = rs("WFSMPLIDRSCW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iRS)
                If IsNull(rs("WFSMPLIDOICW")) = False Then .WFSMPLIDOICW = rs("WFSMPLIDOICW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iOi)
                If IsNull(rs("WFSMPLIDB1CW")) = False Then .WFSMPLIDB1CW = rs("WFSMPLIDB1CW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iB1)
                If IsNull(rs("WFSMPLIDB2CW")) = False Then .WFSMPLIDB2CW = rs("WFSMPLIDB2CW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iB2�j
                If IsNull(rs("WFSMPLIDB3CW")) = False Then .WFSMPLIDB3CW = rs("WFSMPLIDB3CW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iB3)
                If IsNull(rs("WFSMPLIDL1CW")) = False Then .WFSMPLIDL1CW = rs("WFSMPLIDL1CW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iL1)
                If IsNull(rs("WFSMPLIDL2CW")) = False Then .WFSMPLIDL2CW = rs("WFSMPLIDL2CW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iL2)
                If IsNull(rs("WFSMPLIDL3CW")) = False Then .WFSMPLIDL3CW = rs("WFSMPLIDL3CW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iL3)
                If IsNull(rs("WFSMPLIDL4CW")) = False Then .WFSMPLIDL4CW = rs("WFSMPLIDL4CW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iL4)
                If IsNull(rs("WFSMPLIDDSCW")) = False Then .WFSMPLIDDSCW = rs("WFSMPLIDDSCW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iDS)
                If IsNull(rs("WFSMPLIDDZCW")) = False Then .WFSMPLIDDZCW = rs("WFSMPLIDDZCW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iDZ)
                If IsNull(rs("WFSMPLIDSPCW")) = False Then .WFSMPLIDSPCW = rs("WFSMPLIDSPCW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iSP)
                If IsNull(rs("WFSMPLIDDO1CW")) = False Then .WFSMPLIDDO1CW = rs("WFSMPLIDDO1CW") Else lngNullCnt = lngNullCnt + 1   ' WF�����ID�iDO1)
                If IsNull(rs("WFSMPLIDDO2CW")) = False Then .WFSMPLIDDO2CW = rs("WFSMPLIDDO2CW") Else lngNullCnt = lngNullCnt + 1   ' WF�����ID�iDO2)
                If IsNull(rs("WFSMPLIDDO3CW")) = False Then .WFSMPLIDDO3CW = rs("WFSMPLIDDO3CW") Else lngNullCnt = lngNullCnt + 1   ' WF�����ID�iDO3)
                If IsNull(rs("WFSMPLIDAOICW")) = False Then .WFSMPLIDAOICW = rs("WFSMPLIDAOICW") Else lngNullCnt = lngNullCnt + 1   ' WF�����ID�iAOi)
                If IsNull(rs("WFSMPLIDGDCW")) = False Then .WFSMPLIDGDCW = rs("WFSMPLIDGDCW") Else lngNullCnt = lngNullCnt + 1      ' WF�����ID�iGD)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                If IsNull(rs("EPSMPLIDB1CW")) = False Then .EPSMPLIDB1CW = rs("EPSMPLIDB1CW") Else lngNullCnt = lngNullCnt + 1      ' EP�����ID�iB1E)
                If IsNull(rs("EPSMPLIDB2CW")) = False Then .EPSMPLIDB2CW = rs("EPSMPLIDB2CW") Else lngNullCnt = lngNullCnt + 1      ' EP�����ID�iB2E�j
                If IsNull(rs("EPSMPLIDB3CW")) = False Then .EPSMPLIDB3CW = rs("EPSMPLIDB3CW") Else lngNullCnt = lngNullCnt + 1      ' EP�����ID�iB3E)
                If IsNull(rs("EPSMPLIDL1CW")) = False Then .EPSMPLIDL1CW = rs("EPSMPLIDL1CW") Else lngNullCnt = lngNullCnt + 1      ' EP�����ID�iL1E)
                If IsNull(rs("EPSMPLIDL2CW")) = False Then .EPSMPLIDL2CW = rs("EPSMPLIDL2CW") Else lngNullCnt = lngNullCnt + 1      ' EP�����ID�iL2E)
                If IsNull(rs("EPSMPLIDL3CW")) = False Then .EPSMPLIDL3CW = rs("EPSMPLIDL3CW") Else lngNullCnt = lngNullCnt + 1      ' EP�����ID�iL3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                If IsNull(rs("WFHSGDCW")) = False Then .WFHSGDCW = rs("WFHSGDCW") Else lngNullCnt = lngNullCnt + 1                  ' WF�����ۏ� (GD)
                If IsNull(rs("TDAYCW")) = False Then .TDAYCW = rs("TDAYCW") Else lngNullCnt = lngNullCnt + 1                        ' �o�^���t
                If IsNull(rs("KDAYCW")) = False Then .KDAYCW = rs("KDAYCW") Else lngNullCnt = lngNullCnt + 1                        ' �X�V���t
            End With

'Chg Start 2011/03/07 SMPK Miyata
'            If Not rs.EOF Then
'                rs.MoveNext
'            End If
'        Next k
            rs.MoveNext
        End If
'Chg End   2011/03/07 SMPK Miyata

        If lngNullCnt > 0 And sNullSXLID = "" Then
            sNullSXLID = sxl(j).SXLIDCA
        End If

        If lngNullCnt > 0 Then
            WFJudgExecOkFlag(j) = False
        End If
    Next i
    rs.Close
    '�d�|SXL�擾SQL�ύX�@06/02/07 ooba END =================================================>

'=================================================================================
' 2011/02/16 tkimura MOD START
' --- 6�Ԃ̏������R�����g�A�E�g����ƍ����������������B
Debug.Print "6 " & Now & " ���茋�ʂ̎�M�m�F"

    '����]�����ʎ�M�m�F
    '�w���ɑ΂��鑪��]�����ʂ���M���Ă��邩�ǂ����̃`�F�b�N
    '��M���Ă���΁AWF�T���v���Ǘ����X�V
'Cng Start 2011/06/16 Y.Hitomi MQ��M����SIRD���f�s��Ή�
    If MeasRsltCheck1(sxl()) = FUNCTION_RETURN_FAILURE Then
'    If MeasRsltCheck(SXL()) = FUNCTION_RETURN_FAILURE Then
'Cng End   2011/06/16 Y.Hitomi
        DBDRV_scmzc_fcmlc001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


Debug.Print "7 " & Now & " ���������`�F�b�N"

    '���ׂĂ̌������������Ă��邩�ǂ����̃`�F�b�N
    For i = 1 To UBound(sxl)
        '�������ڃ`�F�b�N
        Select Case UBound(sxl(i).WFSMP)
            Case 1
                ReDim Preserve sxl(i).WFSMP(2) As typ_XSDCW
                    WFJudgExecOkFlag(i) = False
            Case 2
                If Trim(sxl(i).WFSMP(1).XTALCW) = "" And Trim(sxl(i).WFSMP(2).XTALCW) = "" Then
                Else
                    If Not (ChkRslt(sxl(i).WFSMP(1)) And ChkRslt(sxl(i).WFSMP(2))) Then
                        WFJudgExecOkFlag(i) = False
                    End If
                End If
'Add Start 2011/03/07 SMPK Miyata
            Case Else
                For k = 1 To UBound(sxl(i).WFSMP)
                    If Not ChkRslt(sxl(i).WFSMP(k)) Then
                        WFJudgExecOkFlag(i) = False
                        Exit For
                    End If
                Next k
'Add End   2011/03/07 SMPK Miyata
        End Select
    Next
Debug.Print "8 " & Now

    If sNullSXLID <> "" Then
        f_cmbc039_1.lblMsg.Caption = "�f�[�^�s�� SXLID=" & sNullSXLID
        GoTo proc_exit
    Else
        f_cmbc039_1.lblMsg.Caption = ""
    End If

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001b_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : Init_SXL_WFSMP
'*
'*    �����T�v      : 1.�V�T���v���Ǘ��e�[�u���̏�����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^         ,����
'*                    WFSMP         ,O  ,typ_XSDCW  ,�V�T���v���Ǘ��iSXL�j
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Sub Init_SXL_WFSMP(WFSMP As typ_XSDCW)
    With WFSMP
        .FACTORYCW = ""
        .HINBCW = ""
        .INPOSCW = 0
        .KDAYCW = 0
        .KSTAFFCW = ""
        .KTKBNCW = ""
        .LIVKCW = ""
        .OPECW = ""
        .REPSMPLIDCW = ""
        .REVNUMCW = 0
        .SMCRYNUMCW = ""
        .SMPKBNCW = ""
        .SMPLNUMCW = 0
        .SNDDAYCW = 0
        .SNDKCW = ""
        .SXLIDCW = ""
        .TBKBNCW = ""
        .TDAYCW = 0
        .TSTAFFCW = ""
        .WFINDAOICW = ""
        .WFINDB1CW = ""
        .WFINDB1CW = ""
        .WFINDB2CW = ""
        .WFINDB3CW = ""
        .WFINDDO1CW = ""
        .WFINDDO2CW = ""
        .WFINDDO3CW = ""
        .WFINDDSCW = ""
        .WFINDDZCW = ""
        .WFINDL1CW = ""
        .WFINDL2CW = ""
        .WFINDL3CW = ""
        .WFINDL4CW = ""
        .WFINDOICW = ""
        .WFINDOT1CW = ""
        .WFINDOT2CW = ""
        .WFINDRSCW = ""
        .WFINDSPCW = ""
        .WFINDGDCW = ""         '05/02/04 ooba
        .WFRESB1CW = ""
        .WFRESAOICW = ""
        .WFRESB2CW = ""
        .WFRESB3CW = ""
        .WFRESDO1CW = ""
        .WFRESDO2CW = ""
        .WFRESDO3CW = ""
        .WFRESDSCW = ""
        .WFRESDZCW = ""
        .WFRESL1CW = ""
        .WFRESL2CW = ""
        .WFRESL3CW = ""
        .WFRESL4CW = ""
        .WFRESOICW = ""
        .WFRESOT1CW = ""
        .WFRESOT2CW = ""
        .WFRESRS1CW = ""
        .WFRESRS2CW = ""
        .WFRESSPCW = ""
        .WFRESGDCW = ""         '05/02/04 ooba
        .WFSMPLIDAOICW = ""
        .WFSMPLIDB1CW = ""
        .WFSMPLIDB2CW = ""
        .WFSMPLIDB3CW = ""
        .WFSMPLIDDO1CW = ""
        .WFSMPLIDDO2CW = ""
        .WFSMPLIDDO3CW = ""
        .WFSMPLIDDSCW = ""
        .WFSMPLIDDZCW = ""
        .WFSMPLIDL1CW = ""
        .WFSMPLIDL2CW = ""
        .WFSMPLIDL3CW = ""
        .WFSMPLIDL4CW = ""
        .WFSMPLIDOICW = ""
        .WFSMPLIDOT1CW = ""
        .WFSMPLIDOT2CW = ""
        .WFSMPLIDRS1CW = ""
        .WFSMPLIDRS2CW = ""
        .WFSMPLIDRSCW = ""
        .WFSMPLIDSPCW = ""
        .WFSMPLIDGDCW = ""      '05/02/04 ooba
        .WFHSGDCW = ""          '05/02/04 ooba
        .XTALCW = ""
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
        .EPSMPLIDB1CW = ""      '�T���v��ID(BMD1)
        .EPINDB1CW = ""         '���FLG(BMD1)
        .EPRESB1CW = ""         '����FLG(BMD1)
        .EPSMPLIDB2CW = ""      '�T���v��ID(BMD2)
        .EPINDB2CW = ""         '���FLG(BMD2)
        .EPRESB2CW = ""         '����FLG(BMD2)
        .EPSMPLIDB3CW = ""      '�T���v��ID(BMD3)
        .EPINDB3CW = ""         '���FLG(BMD3)
        .EPRESB3CW = ""         '����FLG(BMD3)
        .EPSMPLIDL1CW = ""      '�T���v��ID(OSF1)
        .EPINDL1CW = ""         '���FLG(OSF1)
        .EPRESL1CW = ""         '����FLG(OSF1)
        .EPSMPLIDL2CW = ""      '�T���v��ID(OSF2)
        .EPINDL2CW = ""         '���FLG(OSF2)
        .EPRESL2CW = ""         '����FLG(OSF2)
        .EPSMPLIDL3CW = ""      '�T���v��ID(OSF3)
        .EPINDL3CW = ""         '���FLG(OSF3)
        .EPRESL3CW = ""         '����FLG(OSF3)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    End With
End Sub

'**************************************************************************************
'*    �֐���        : GetSxlidINBlkid
'*
'*    �����T�v      : 1.�������ъ����`�F�b�N
'*                      (�����w������Ă��錟�����I�����Ă��邩�`�F�b�N����)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^           ,����
'*                    typ_WfSmp     ,I  ,typ_XSDCW    ,�V�T���v���Ǘ��iSXL�j���\����
'*
'*    �߂�l        : Boolean
'*
'**************************************************************************************
Public Function ChkRslt(typ_WFSmp As typ_XSDCW) As Boolean
    With typ_WFSmp
        If .WFINDRSCW <> "0" And .WFRESRS1CW = "0" Then               ' ���FLG�iRs)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDOICW <> "0" And .WFRESOICW = "0" Then           ' ���FLG�iOi)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDB1CW <> "0" And .WFRESB1CW = "0" Then           ' ���FLG�iB1)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDB2CW <> "0" And .WFRESB2CW = "0" Then           ' ���FLG�iB2�j
            ChkRslt = False
            Exit Function
        ElseIf .WFINDB3CW <> "0" And .WFRESB3CW = "0" Then           ' ���FLG�iB3)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDL1CW <> "0" And .WFRESL1CW = "0" Then           ' ���FLG�iL1)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDL2CW <> "0" And .WFRESL2CW = "0" Then           ' ���FLG�iL2)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDL3CW <> "0" And .WFRESL3CW = "0" Then           ' ���FLG�iL3)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDL4CW <> "0" And .WFRESL4CW = "0" Then           ' ���FLG�iL4)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDDSCW <> "0" And .WFRESDSCW = "0" Then           ' ���FLG�iDS)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDDZCW <> "0" And .WFRESDZCW = "0" Then           ' ���FLG�iDZ)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDSPCW <> "0" And .WFRESSPCW = "0" Then           ' ���FLG�iSP)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDDO1CW <> "0" And .WFRESDO1CW = "0" Then         ' ���FLG�iDO1)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDDO2CW <> "0" And .WFRESDO2CW = "0" Then         ' ���FLG�iDO2)
            ChkRslt = False
            Exit Function
        ElseIf .WFINDDO3CW <> "0" And .WFRESDO3CW = "0" Then         ' ���FLG�iDO3)
            ChkRslt = False
            Exit Function
        ''�c���_�f�ǉ��@03/12/15 ooba
        ElseIf .WFINDAOICW <> "0" And .WFRESAOICW = "0" Then         ' ���FLG (AOi)
            ChkRslt = False
            Exit Function
        'GD�ǉ��@05/02/17 ooba
        ElseIf .WFINDGDCW <> "0" And .WFRESGDCW = "0" Then           ' ���FLG (GD)
            ChkRslt = False
            Exit Function
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
        ElseIf .EPINDB1CW <> "0" And .EPRESB1CW = "0" Then           ' ���FLG�iB1E)
            ChkRslt = False
            Exit Function
        ElseIf .EPINDB2CW <> "0" And .EPRESB2CW = "0" Then           ' ���FLG�iB2E�j
            ChkRslt = False
            Exit Function
        ElseIf .EPINDB3CW <> "0" And .EPRESB3CW = "0" Then           ' ���FLG�iB3E)
            ChkRslt = False
            Exit Function
        ElseIf .EPINDL1CW <> "0" And .EPRESL1CW = "0" Then           ' ���FLG�iL1E)
            ChkRslt = False
            Exit Function
        ElseIf .EPINDL2CW <> "0" And .EPRESL2CW = "0" Then           ' ���FLG�iL2E)
            ChkRslt = False
            Exit Function
        ElseIf .EPINDL3CW <> "0" And .EPRESL3CW = "0" Then           ' ���FLG�iL3E)
            ChkRslt = False
            Exit Function
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
        End If
    End With
    ChkRslt = True
End Function

'**************************************************************************************************
'*    �֐���        : DBDRV_GetTBCMY013
'*
'*    �����T�v      : 1.�e�[�u���uTBCMY013�v��������ɂ��������R�[�h�𒊏o����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^           ,����
'*                    records()     ,O  ,typ_TBCMY013 ,���o���R�[�h
'*                    sSqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'*                    sSqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'**************************************************************************************************
Public Function DBDRV_GetTBCMY013(records() As typ_TBCMY013, Optional sSqlWhere$ = vbNullString, _
                                   Optional sSqlOrder$ = vbNullString) As FUNCTION_RETURN
    Dim sSQL        As String       'SQL�S��
    Dim sSqlBase    As String       'SQL��{��(WHERE�߂̑O�܂�)
    Dim rs          As OraDynaset   'RecordSet
    Dim lngRecCnt   As Long         '���R�[�h��
    Dim i           As Long

    ''SQL��g�ݗ��Ă�
    sSqlBase = "Select SAMPLEID, OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5," & _
              " MESDATA6, MESDATA7, MESDATA8, MESDATA9, MESDATA10, MESDATA11, MESDATA12, MESDATA13, MESDATA14, MESDATA15," & _
              " TXID, REGDATE, SENDFLAG, SENDDATE "
    sSqlBase = sSqlBase & "From TBCMY013"
    sSQL = sSqlBase
    If (sSqlWhere <> vbNullString) Or (sSqlOrder <> vbNullString) Then
        sSQL = sSQL & " " & sSqlWhere & " " & sSqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMY013 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    lngRecCnt = rs.RecordCount
    ReDim records(lngRecCnt)
    For i = 1 To lngRecCnt
        With records(i)
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
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMY013 = FUNCTION_RETURN_SUCCESS
End Function

'**************************************************************************************************
'*    �֐���        : DBDRV_scmzc_fcmlc001c_Disp
'*
'*    �����T�v      : 1.WF�������� �\���p�c�a�h���C�o
'*
'*    �p�����[�^    : �ϐ���       ,IO  ,�^                                    ,����
'*                    typIn        ,I   ,type_DBDRV_scmzc_fcmlc001c_In         ,���͗p
'*                    Siyou        ,O   ,type_DBDRV_scmzc_fcmlc001c_Siyou      ,WF�d�l�p
'*                    Sokutei      ,O   ,typ_TBCMY013                          ,����]������
'*              �@�@  sErrMsg �@�@ ,O   ,String    �@�@�@�@�@�@�@�@�@�@�@      ,�G���[���b�Z�[�W
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'**************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_Disp(typIn As type_DBDRV_scmzc_fcmlc001c_In039, _
                                           siyou As type_DBDRV_scmzc_fcmlc001c_Siyou039, _
                                           Sokutei() As typ_TBCMY013, _
                                           sErrMsg As String) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Integer
    Dim i           As Long
    Dim sDBName     As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_Disp"

    DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_SUCCESS

    'WF�d�l�擾
    sDBName = "V001"
    sSQL = "select "
    sSQL = sSQL & "E021HWFTYPE, "            ' �i�v�e�^�C�v
    sSQL = sSQL & "E022HWFCDIR, "            ' �i�v�e�����ʕ�
    sSQL = sSQL & "E023HWFCDOP, "            ' �i�v�e�����h�[�v

    sSQL = sSQL & "E021HWFRMIN, "            ' �i�v�e���R����
    sSQL = sSQL & "E021HWFRMAX, "            ' �i�v�e���R���
    sSQL = sSQL & "E021HWFRSPOH, "           ' �i�v�e���R����ʒu�Q��
    sSQL = sSQL & "E021HWFRSPOT, "           ' �i�v�e���R����ʒu�Q�_
    sSQL = sSQL & "E021HWFRSPOI, "           ' �i�v�e���R����ʒu�Q��
    sSQL = sSQL & "E021HWFRHWYT, "           ' �i�v�e���R�ۏؕ��@�Q��
    sSQL = sSQL & "E021HWFRHWYS, "           ' �i�v�e���R�ۏؕ��@�Q��
    sSQL = sSQL & "E021HWFRMCAL, "           ' �i�v�e���R�ʓ��v�Z 2001/11/08 S.Sano
    sSQL = sSQL & "E021HWFRAMIN, "           ' �i�v�e���R���ω���
    sSQL = sSQL & "E021HWFRAMAX, "           ' �i�v�e���R���Ϗ��
    sSQL = sSQL & "E021HWFRMBNP, "           ' �i�v�e���R�ʓ����z

    sSQL = sSQL & "E024HWFMKMIN, "           ' �i�v�e�����בw����
    sSQL = sSQL & "E024HWFMKMAX, "           ' �i�v�e�����בw���
    sSQL = sSQL & "E024HWFMKSPH, "           ' �i�v�e�����בw����ʒu�Q��
    sSQL = sSQL & "E024HWFMKSPT, "           ' �i�v�e�����בw����ʒu�Q�_
    sSQL = sSQL & "E024HWFMKSPR, "           ' �i�v�e�����בw����ʒu�Q��
    sSQL = sSQL & "E024HWFMKHWT, "           ' �i�v�e�����בw�ۏؕ��@�Q��
    sSQL = sSQL & "E024HWFMKHWS, "           ' �i�v�e�����בw�ۏؕ��@�Q��

    sSQL = sSQL & "E025HWFONMIN, "           ' �i�v�e�_�f�Z�x����
    sSQL = sSQL & "E025HWFONMAX, "           ' �i�v�e�_�f�Z�x���
    sSQL = sSQL & "E025HWFONSPH, "           ' �i�v�e�_�f�Z�x����ʒu�Q��
    sSQL = sSQL & "E025HWFONSPT, "           ' �i�v�e�_�f�Z�x����ʒu�Q�_
    sSQL = sSQL & "E025HWFONSPI, "           ' �i�v�e�_�f�Z�x����ʒu�Q��
    sSQL = sSQL & "E025HWFONHWT, "           ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
    sSQL = sSQL & "E025HWFONHWS, "           ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
    sSQL = sSQL & "E025HWFONMCL, "           ' �i�v�e�_�f�Z�x�ʓ��v�Z 2001/11/08 S.Sano
    sSQL = sSQL & "E025HWFONMBP, "           ' �i�v�e�_�f�Z�x�ʓ����z
    sSQL = sSQL & "E025HWFONAMN, "           ' �i�v�e�_�f�Z�x���ω���
    sSQL = sSQL & "E025HWFONAMX, "           ' �i�v�e�_�f�Z�x���Ϗ��

    sSQL = sSQL & "E025HWFOS1MN, "           ' �i�v�e�_�f�͏o�P����
    sSQL = sSQL & "E025HWFOS1MX, "           ' �i�v�e�_�f�͏o�P���
    sSQL = sSQL & "E025HWFOS1SH, "           ' �i�v�e�_�f�͏o�P����ʒu�Q��
    sSQL = sSQL & "E025HWFOS1ST, "           ' �i�v�e�_�f�͏o�P����ʒu�Q�_
    sSQL = sSQL & "E025HWFOS1SI, "           ' �i�v�e�_�f�͏o�P����ʒu�Q��
    sSQL = sSQL & "E025HWFOS1HT, "           ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
    sSQL = sSQL & "E025HWFOS1HS, "           ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
    sSQL = sSQL & "E025HWFOS2SH, "           ' �i�v�e�_�f�͏o�Q����ʒu�Q��
    sSQL = sSQL & "E025HWFOS2ST, "           ' �i�v�e�_�f�͏o�Q����ʒu�Q�_
    sSQL = sSQL & "E025HWFOS2SI, "           ' �i�v�e�_�f�͏o�Q����ʒu�Q��
    sSQL = sSQL & "E025HWFOS2MN, "           ' �i�v�e�_�f�͏o�Q����
    sSQL = sSQL & "E025HWFOS2MX, "           ' �i�v�e�_�f�͏o�Q���
    sSQL = sSQL & "E025HWFOS2HT, "           ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
    sSQL = sSQL & "E025HWFOS2HS, "           ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
    sSQL = sSQL & "E025HWFOS3MN, "           ' �i�v�e�_�f�͏o�R����
    sSQL = sSQL & "E025HWFOS3MX, "           ' �i�v�e�_�f�͏o�R���
    sSQL = sSQL & "E025HWFOS3SH, "           ' �i�v�e�_�f�͏o�R����ʒu�Q��
    sSQL = sSQL & "E025HWFOS3ST, "           ' �i�v�e�_�f�͏o�R����ʒu�Q�_
    sSQL = sSQL & "E025HWFOS3SI, "           ' �i�v�e�_�f�͏o�R����ʒu�Q��
    sSQL = sSQL & "E025HWFOS3HT, "           ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��
    sSQL = sSQL & "E025HWFOS3HS, "           ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��

    sSQL = sSQL & "E026HWFDSOMX, "           ' �i�v�e�c�r�n�c���
    sSQL = sSQL & "E026HWFDSOMN, "           ' �i�v�e�c�r�n�c����
    sSQL = sSQL & "E026HWFDSOAX, "           ' �i�v�e�c�r�n�c�̈���
    sSQL = sSQL & "E026HWFDSOAN, "           ' �i�v�e�c�r�n�c�̈扺��
    sSQL = sSQL & "E026HWFDSOHT, "           ' �i�v�e�c�r�n�c�ۏؕ��@�Q��
    sSQL = sSQL & "E026HWFDSOHS, "           ' �i�v�e�c�r�n�c�ۏؕ��@�Q��

    sSQL = sSQL & "E028HWFSPVMX, "           ' �i�v�e�r�o�u�e�d���
    sSQL = sSQL & "E028HWFSPVSH, "           ' �i�v�e�r�o�u�e�d����ʒu�Q��
    sSQL = sSQL & "E028HWFSPVST, "           ' �i�v�e�r�o�u�e�d����ʒu�Q�_
    sSQL = sSQL & "E028HWFSPVSI, "           ' �i�v�e�r�o�u�e�d����ʒu�Q��
    sSQL = sSQL & "E028HWFSPVHT, "           ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
    sSQL = sSQL & "E028HWFSPVHS, "           ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
    sSQL = sSQL & "E028HWFDLSPH, "           ' �i�v�e�g�U������ʒu�Q��
    sSQL = sSQL & "E028HWFDLSPT, "           ' �i�v�e�g�U������ʒu�Q�_
    sSQL = sSQL & "E028HWFDLSPI, "           ' �i�v�e�g�U������ʒu�Q��
    sSQL = sSQL & "E028HWFDLHWT, "           ' �i�v�e�g�U���ۏؕ��@�Q��
    sSQL = sSQL & "E028HWFDLHWS, "           ' �i�v�e�g�U���ۏؕ��@�Q��
    sSQL = sSQL & "E028HWFDLMIN, "           ' �i�v�e�g�U������
    sSQL = sSQL & "E028HWFDLMAX, "           ' �i�v�e�g�U�����

    sSQL = sSQL & "E029HWFOF1AX, "          ' �i�v�e�n�r�e�P���Ϗ��
    sSQL = sSQL & "E029HWFOF1MX, "          ' �i�v�e�n�r�e�P���
    sSQL = sSQL & "E029HWFOF1SH, "          ' �i�v�e�n�r�e�P����ʒu�Q��
    sSQL = sSQL & "E029HWFOF1ST, "          ' �i�v�e�n�r�e�P����ʒu�Q�_
    sSQL = sSQL & "E029HWFOF1SR, "          ' �i�v�e�n�r�e�P����ʒu�Q��
    sSQL = sSQL & "E029HWFOF1HT, "          ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF1HS, "          ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF2AX, "          ' �i�v�e�n�r�e�Q���Ϗ��
    sSQL = sSQL & "E029HWFOF2MX, "          ' �i�v�e�n�r�e�Q���
    sSQL = sSQL & "E029HWFOF2SH, "          ' �i�v�e�n�r�e�Q����ʒu�Q��
    sSQL = sSQL & "E029HWFOF2ST, "          ' �i�v�e�n�r�e�Q����ʒu�Q�_
    sSQL = sSQL & "E029HWFOF2SR, "          ' �i�v�e�n�r�e�Q����ʒu�Q��
    sSQL = sSQL & "E029HWFOF2HT, "          ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF2HS, "          ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF3AX, "          ' �i�v�e�n�r�e�R���Ϗ��
    sSQL = sSQL & "E029HWFOF3MX, "          ' �i�v�e�n�r�e�R���
    sSQL = sSQL & "E029HWFOF3SH, "          ' �i�v�e�n�r�e�R����ʒu�Q��
    sSQL = sSQL & "E029HWFOF3ST, "          ' �i�v�e�n�r�e�R����ʒu�Q�_
    sSQL = sSQL & "E029HWFOF3SR, "          ' �i�v�e�n�r�e�R����ʒu�Q��
    sSQL = sSQL & "E029HWFOF3HT, "          ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF3HS, "          ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF4AX, "          ' �i�v�e�n�r�e�S���Ϗ��
    sSQL = sSQL & "E029HWFOF4MX, "          ' �i�v�e�n�r�e�S���
    sSQL = sSQL & "E029HWFOF4SH, "          ' �i�v�e�n�r�e�S����ʒu�Q��
    sSQL = sSQL & "E029HWFOF4ST, "          ' �i�v�e�n�r�e�S����ʒu�Q�_
    sSQL = sSQL & "E029HWFOF4SR, "          ' �i�v�e�n�r�e�S����ʒu�Q��
    sSQL = sSQL & "E029HWFOF4HT, "          ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF4HS, "          ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM1AN, "          ' �i�v�e�a�l�c�P���ω���
    sSQL = sSQL & "E029HWFBM1AX, "          ' �i�v�e�a�l�c�P���Ϗ��
    sSQL = sSQL & "E029HWFBM1SH, "          ' �i�v�e�a�l�c�P����ʒu�Q��
    sSQL = sSQL & "E029HWFBM1ST, "          ' �i�v�e�a�l�c�P����ʒu�Q�_
    sSQL = sSQL & "E029HWFBM1SR, "          ' �i�v�e�a�l�c�P����ʒu�Q��
    sSQL = sSQL & "E029HWFBM1HT, "          ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM1HS, "          ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM2AN, "          ' �i�v�e�a�l�c�Q���ω���
    sSQL = sSQL & "E029HWFBM2AX, "          ' �i�v�e�a�l�c�Q���Ϗ��
    sSQL = sSQL & "E029HWFBM2SH, "          ' �i�v�e�a�l�c�Q����ʒu�Q��
    sSQL = sSQL & "E029HWFBM2ST, "          ' �i�v�e�a�l�c�Q����ʒu�Q�_
    sSQL = sSQL & "E029HWFBM2SR, "          ' �i�v�e�a�l�c�Q����ʒu�Q��
    sSQL = sSQL & "E029HWFBM2HT, "          ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM2HS, "          ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM3AN, "          ' �i�v�e�a�l�c�R���ω���
    sSQL = sSQL & "E029HWFBM3AX, "          ' �i�v�e�a�l�c�R���Ϗ��
    sSQL = sSQL & "E029HWFBM3SH, "          ' �i�v�e�a�l�c�R����ʒu�Q��
    sSQL = sSQL & "E029HWFBM3ST, "          ' �i�v�e�a�l�c�R����ʒu�Q�_
    sSQL = sSQL & "E029HWFBM3SR, "          ' �i�v�e�a�l�c�R����ʒu�Q��
    sSQL = sSQL & "E029HWFBM3HT, "          ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM3HS, "          ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOSF1PTK, "        ' �i�v�e�n�r�e�P�p�^���敪�@��2003/05/14 ooba
    sSQL = sSQL & "E029HWFOSF2PTK, "        ' �i�v�e�n�r�e�Q�p�^���敪
    sSQL = sSQL & "E029HWFOSF3PTK, "        ' �i�v�e�n�r�e�R�p�^���敪
    sSQL = sSQL & "E029HWFOSF4PTK, "        ' �i�v�e�n�r�e�S�p�^���敪
    sSQL = sSQL & "E029HWFBM1MBP, "         ' �i�v�e�a�l�c�P�ʓ����z
    sSQL = sSQL & "E029HWFBM2MBP, "         ' �i�v�e�a�l�c�Q�ʓ����z
    sSQL = sSQL & "E029HWFBM3MBP, "         ' �i�v�e�a�l�c�R�ʓ����z
    sSQL = sSQL & "E029HWFBM1MCL, "         ' �i�v�e�a�l�c�P�ʓ��v�Z
    sSQL = sSQL & "E029HWFBM2MCL, "         ' �i�v�e�a�l�c�Q�ʓ��v�Z
    sSQL = sSQL & "E029HWFBM3MCL, "         ' �i�v�e�a�l�c�R�ʓ��v�Z�@��2003/05/14 ooba
    sSQL = sSQL & "E025HWFOS1NS, "          ' �i�v�e�_�f�͏o�P�M�����@
    sSQL = sSQL & "E025HWFOS2NS, "          ' �i�v�e�_�f�͏o�Q�M�����@
    sSQL = sSQL & "E025HWFOS3NS, "          ' �i�v�e�_�f�͏o�R�M�����@

    sSQL = sSQL & "E029HWFOF1NS, "          ' �i�v�e�n�r�e�P�M�����@
    sSQL = sSQL & "E029HWFOF2NS, "          ' �i�v�e�n�r�e�Q�M�����@
    sSQL = sSQL & "E029HWFOF3NS, "          ' �i�v�e�n�r�e�R�M�����@
    sSQL = sSQL & "E029HWFOF4NS, "          ' �i�v�e�n�r�e�S�M�����@

    sSQL = sSQL & "E029HWFBM1NS, "          ' �i�v�e�a�l�c�P�M�����@
    sSQL = sSQL & "E029HWFBM2NS, "          ' �i�v�e�a�l�c�Q�M�����@
    sSQL = sSQL & "E029HWFBM3NS, "          ' �i�v�e�a�l�c�R�M�����@

    sSQL = sSQL & "E025HWFANTIM, "          ' �i�v�e�`�m����
    sSQL = sSQL & "E025HWFANTNP, "          ' �i�v�e�`�m���x

    sSQL = sSQL & "E029HWFOF1ET, "          ' �i�v�e�n�r�e�P�I���d�s��
    sSQL = sSQL & "E029HWFOF2ET, "          ' �i�v�e�n�r�e�Q�I���d�s��
    sSQL = sSQL & "E029HWFOF3ET, "          ' �i�v�e�n�r�e�R�I���d�s��
    sSQL = sSQL & "E029HWFOF4ET, "          ' �i�v�e�n�r�e�S�I���d�s��
    sSQL = sSQL & "E029HWFBM1ET, "          ' �i�v�e�a�l�c�P�I���d�s��
    sSQL = sSQL & "E029HWFBM2ET, "          ' �i�v�e�a�l�c�Q�I���d�s��
    sSQL = sSQL & "E029HWFBM3ET, "          ' �i�v�e�a�l�c�R�I���d�s��

    sSQL = sSQL & "E029HWFOF1SZ, "          ' �i�v�e�n�r�e�P�������
    sSQL = sSQL & "E029HWFOF2SZ, "          ' �i�v�e�n�r�e�Q�������
    sSQL = sSQL & "E029HWFOF3SZ, "          ' �i�v�e�n�r�e�R�������
    sSQL = sSQL & "E029HWFOF4SZ, "          ' �i�v�e�n�r�e�S�������
    sSQL = sSQL & "E029HWFBM1SZ, "          ' �i�v�e�a�l�c�P�������
    sSQL = sSQL & "E029HWFBM2SZ, "          ' �i�v�e�a�l�c�Q�������
    sSQL = sSQL & "E029HWFBM3SZ "           ' �i�v�e�a�l�c�R�������

    sSQL = sSQL & " from VECME001"
    sSQL = sSQL & " where E018HINBAN='" & typIn.HIN.hinban & "' and "
    sSQL = sSQL & " E018MNOREVNO=" & typIn.HIN.mnorevno & " and "
    sSQL = sSQL & " E018FACTORY='" & typIn.HIN.factory & "' and "
    sSQL = sSQL & " E018OPECOND='" & typIn.HIN.opecond & "' "

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    '���R�[�h0���̓G���[�I��
    If rs.RecordCount = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With siyou
        .HWFTYPE = rs("E021HWFTYPE")              ' �i�v�e�^�C�v
        .HWFCDIR = rs("E022HWFCDIR")              ' �i�v�e�����ʕ�
        .HWFCDOP = rs("E023HWFCDOP")              ' �i�v�e�����h�[�v

        .HWFRMIN = rs("E021HWFRMIN")              ' �i�v�e���R����
        .HWFRMAX = rs("E021HWFRMAX")              ' �i�v�e���R���
        .HWFRSPOH = rs("E021HWFRSPOH")            ' �i�v�e���R����ʒu�Q��
        .HWFRSPOT = rs("E021HWFRSPOT")            ' �i�v�e���R����ʒu�Q�_
        .HWFRSPOI = rs("E021HWFRSPOI")            ' �i�v�e���R����ʒu�Q��
        .HWFRHWYT = rs("E021HWFRHWYT")            ' �i�v�e���R�ۏؕ��@�Q��
        .HWFRHWYS = rs("E021HWFRHWYS")            ' �i�v�e���R�ۏؕ��@�Q��
        .HWFRMCAL = rs("E021HWFRMCAL")            ' �i�v�e���R�ʓ��v�Z 2001/11/08 S.Sano
        .HWFRAMIN = rs("E021HWFRAMIN")            ' �i�v�e���R���ω���
        .HWFRAMAX = rs("E021HWFRAMAX")            ' �i�v�e���R���Ϗ��
        .HWFRMBNP = rs("E021HWFRMBNP")            ' �i�v�e���R�ʓ����z

        .HWFMKMIN = rs("E024HWFMKMIN")            ' �i�v�e�����בw����
        .HWFMKMAX = rs("E024HWFMKMAX")            ' �i�v�e�����בw���
        .HWFMKSPH = rs("E024HWFMKSPH")            ' �i�v�e�����בw����ʒu�Q��
        .HWFMKSPT = rs("E024HWFMKSPT")            ' �i�v�e�����בw����ʒu�Q�_
        .HWFMKSPR = rs("E024HWFMKSPR")            ' �i�v�e�����בw����ʒu�Q��
        .HWFMKHWT = rs("E024HWFMKHWT")            ' �i�v�e�����בw�ۏؕ��@�Q��
        .HWFMKHWS = rs("E024HWFMKHWS")            ' �i�v�e�����בw�ۏؕ��@�Q��

        .HWFONMIN = rs("E025HWFONMIN")            ' �i�v�e�_�f�Z�x����
        .HWFONMAX = rs("E025HWFONMAX")            ' �i�v�e�_�f�Z�x���
        .HWFONSPH = rs("E025HWFONSPH")            ' �i�v�e�_�f�Z�x����ʒu�Q��
        .HWFONSPT = rs("E025HWFONSPT")            ' �i�v�e�_�f�Z�x����ʒu�Q�_
        .HWFONSPI = rs("E025HWFONSPI")            ' �i�v�e�_�f�Z�x����ʒu�Q��
        .HWFONHWT = rs("E025HWFONHWT")            ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
        .HWFONHWS = rs("E025HWFONHWS")            ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
        .HWFONMCL = rs("E025HWFONMCL")            ' �i�v�e�_�f�Z�x�ʓ��v�Z 2001/11/08 S.Sano
        .HWFONMBP = rs("E025HWFONMBP")            ' �i�v�e�_�f�Z�x�ʓ����z
        .HWFONAMN = rs("E025HWFONAMN")            ' �i�v�e�_�f�Z�x���ω���
        .HWFONAMX = rs("E025HWFONAMX")            ' �i�v�e�_�f�Z�x���Ϗ��

        .HWFOS1MN = rs("E025HWFOS1MN")            ' �i�v�e�_�f�͏o�P����
        .HWFOS1MX = rs("E025HWFOS1MX")            ' �i�v�e�_�f�͏o�P���
        .HWFOS1SH = rs("E025HWFOS1SH")            ' �i�v�e�_�f�͏o�P����ʒu�Q��
        .HWFOS1ST = rs("E025HWFOS1ST")            ' �i�v�e�_�f�͏o�P����ʒu�Q�_
        .HWFOS1SI = rs("E025HWFOS1SI")            ' �i�v�e�_�f�͏o�P����ʒu�Q��
        .HWFOS1HT = rs("E025HWFOS1HT")            ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        .HWFOS1HS = rs("E025HWFOS1HS")            ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        .HWFOS2SH = rs("E025HWFOS2SH")            ' �i�v�e�_�f�͏o�Q����ʒu�Q��
        .HWFOS2ST = rs("E025HWFOS2ST")            ' �i�v�e�_�f�͏o�Q����ʒu�Q�_
        .HWFOS2SI = rs("E025HWFOS2SI")            ' �i�v�e�_�f�͏o�Q����ʒu�Q��
        .HWFOS2MN = rs("E025HWFOS2MN")            ' �i�v�e�_�f�͏o�Q����
        .HWFOS2MX = rs("E025HWFOS2MX")            ' �i�v�e�_�f�͏o�Q���
        .HWFOS2HT = rs("E025HWFOS2HT")            ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
        .HWFOS2HS = rs("E025HWFOS2HS")            ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
        .HWFOS3MN = rs("E025HWFOS3MN")            ' �i�v�e�_�f�͏o�R����
        .HWFOS3MX = rs("E025HWFOS3MX")            ' �i�v�e�_�f�͏o�R���
        .HWFOS3SH = rs("E025HWFOS3SH")            ' �i�v�e�_�f�͏o�R����ʒu�Q��
        .HWFOS3ST = rs("E025HWFOS3ST")            ' �i�v�e�_�f�͏o�R����ʒu�Q�_
        .HWFOS3SI = rs("E025HWFOS3SI")            ' �i�v�e�_�f�͏o�R����ʒu�Q��
        .HWFOS3HT = rs("E025HWFOS3HT")            ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��
        .HWFOS3HS = rs("E025HWFOS3HS")            ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��

        .HWFDSOMX = rs("E026HWFDSOMX")            ' �i�v�e�c�r�n�c���
        .HWFDSOMN = rs("E026HWFDSOMN")            ' �i�v�e�c�r�n�c����
        .HWFDSOAX = rs("E026HWFDSOAX")            ' �i�v�e�c�r�n�c�̈���
        .HWFDSOAN = rs("E026HWFDSOAN")            ' �i�v�e�c�r�n�c�̈扺��
        .HWFDSOHT = rs("E026HWFDSOHT")            ' �i�v�e�c�r�n�c�ۏؕ��@�Q��
        .HWFDSOHS = rs("E026HWFDSOHS")            ' �i�v�e�c�r�n�c�ۏؕ��@�Q��

        .HWFSPVMX = rs("E028HWFSPVMX")            ' �i�v�e�r�o�u�e�d���
        .HWFSPVSH = rs("E028HWFSPVSH")            ' �i�v�e�r�o�u�e�d����ʒu�Q��
        .HWFSPVST = rs("E028HWFSPVST")            ' �i�v�e�r�o�u�e�d����ʒu�Q�_
        .HWFSPVSI = rs("E028HWFSPVSI")            ' �i�v�e�r�o�u�e�d����ʒu�Q��
        .HWFSPVHT = rs("E028HWFSPVHT")            ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
        .HWFSPVHS = rs("E028HWFSPVHS")            ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
        .HWFDLSPH = rs("E028HWFDLSPH")            ' �i�v�e�g�U������ʒu�Q��
        .HWFDLSPT = rs("E028HWFDLSPT")            ' �i�v�e�g�U������ʒu�Q�_
        .HWFDLSPI = rs("E028HWFDLSPI")            ' �i�v�e�g�U������ʒu�Q��
        .HWFDLHWT = rs("E028HWFDLHWT")            ' �i�v�e�g�U���ۏؕ��@�Q��
        .HWFDLHWS = rs("E028HWFDLHWS")            ' �i�v�e�g�U���ۏؕ��@�Q��
        .HWFDLMIN = rs("E028HWFDLMIN")            ' �i�v�e�g�U������
        .HWFDLMAX = rs("E028HWFDLMAX")            ' �i�v�e�g�U�����

        .HWFOF1AX = rs("E029HWFOF1AX")           ' �i�v�e�n�r�e�P���Ϗ��
        .HWFOF1MX = rs("E029HWFOF1MX")           ' �i�v�e�n�r�e�P���
        .HWFOF1SH = rs("E029HWFOF1SH")           ' �i�v�e�n�r�e�P����ʒu�Q��
        .HWFOF1ST = rs("E029HWFOF1ST")           ' �i�v�e�n�r�e�P����ʒu�Q�_
        .HWFOF1SR = rs("E029HWFOF1SR")           ' �i�v�e�n�r�e�P����ʒu�Q��
        .HWFOF1HT = rs("E029HWFOF1HT")           ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
        .HWFOF1HS = rs("E029HWFOF1HS")           ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
        .HWFOF2AX = rs("E029HWFOF2AX")           ' �i�v�e�n�r�e�Q���Ϗ��
        .HWFOF2MX = rs("E029HWFOF2MX")           ' �i�v�e�n�r�e�Q���
        .HWFOF2SH = rs("E029HWFOF2SH")           ' �i�v�e�n�r�e�Q����ʒu�Q��
        .HWFOF2ST = rs("E029HWFOF2ST")           ' �i�v�e�n�r�e�Q����ʒu�Q�_
        .HWFOF2SR = rs("E029HWFOF2SR")           ' �i�v�e�n�r�e�Q����ʒu�Q��
        .HWFOF2HT = rs("E029HWFOF2HT")           ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
        .HWFOF2HS = rs("E029HWFOF2HS")           ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
        .HWFOF3AX = rs("E029HWFOF3AX")           ' �i�v�e�n�r�e�R���Ϗ��
        .HWFOF3MX = rs("E029HWFOF3MX")           ' �i�v�e�n�r�e�R���
        .HWFOF3SH = rs("E029HWFOF3SH")           ' �i�v�e�n�r�e�R����ʒu�Q��
        .HWFOF3ST = rs("E029HWFOF3ST")           ' �i�v�e�n�r�e�R����ʒu�Q�_
        .HWFOF3SR = rs("E029HWFOF3SR")           ' �i�v�e�n�r�e�R����ʒu�Q��
        .HWFOF3HT = rs("E029HWFOF3HT")           ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
        .HWFOF3HS = rs("E029HWFOF3HS")           ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
        .HWFOF4AX = rs("E029HWFOF4AX")           ' �i�v�e�n�r�e�S���Ϗ��
        .HWFOF4MX = rs("E029HWFOF4MX")           ' �i�v�e�n�r�e�S���
        .HWFOF4SH = rs("E029HWFOF4SH")           ' �i�v�e�n�r�e�S����ʒu�Q��
        .HWFOF4ST = rs("E029HWFOF4ST")           ' �i�v�e�n�r�e�S����ʒu�Q�_
        .HWFOF4SR = rs("E029HWFOF4SR")           ' �i�v�e�n�r�e�S����ʒu�Q��
        .HWFOF4HT = rs("E029HWFOF4HT")           ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
        .HWFOF4HS = rs("E029HWFOF4HS")           ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
        If IsNull(rs("E029HWFOSF1PTK")) = False Then .HWFOSF1PTK = rs("E029HWFOSF1PTK")       ' �i�v�e�n�r�e�P�p�^���敪�@��2003/05/14 ooba
        If IsNull(rs("E029HWFOSF2PTK")) = False Then .HWFOSF2PTK = rs("E029HWFOSF2PTK")       ' �i�v�e�n�r�e�Q�p�^���敪
        If IsNull(rs("E029HWFOSF3PTK")) = False Then .HWFOSF3PTK = rs("E029HWFOSF3PTK")       ' �i�v�e�n�r�e�R�p�^���敪
        If IsNull(rs("E029HWFOSF4PTK")) = False Then .HWFOSF4PTK = rs("E029HWFOSF4PTK")       ' �i�v�e�n�r�e�S�p�^���敪�@��2003/05/14 ooba

        'BMD�ׂ��搔�ύX�Ή��@2003/05/19 osawa
        .HWFBM1AN = rs("E029HWFBM1AN")           ' �i�v�e�a�l�c�P���ω���
        .HWFBM1AX = rs("E029HWFBM1AX")           ' �i�v�e�a�l�c�P���Ϗ��
        .HWFBM1SH = rs("E029HWFBM1SH")           ' �i�v�e�a�l�c�P����ʒu�Q��
        .HWFBM1ST = rs("E029HWFBM1ST")           ' �i�v�e�a�l�c�P����ʒu�Q�_
        .HWFBM1SR = rs("E029HWFBM1SR")           ' �i�v�e�a�l�c�P����ʒu�Q��
        .HWFBM1HT = rs("E029HWFBM1HT")           ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
        .HWFBM1HS = rs("E029HWFBM1HS")           ' �i�v�e�a�l�c�P�ۏؕ��@�Q��

        'BMD�ׂ��搔�ύX�Ή��@2003/05/19 osawa
        .HWFBM2AN = rs("E029HWFBM2AN")           ' �i�v�e�a�l�c�Q���ω���
        .HWFBM2AX = rs("E029HWFBM2AX")           ' �i�v�e�a�l�c�Q���Ϗ��
        .HWFBM2SH = rs("E029HWFBM2SH")           ' �i�v�e�a�l�c�Q����ʒu�Q��
        .HWFBM2ST = rs("E029HWFBM2ST")           ' �i�v�e�a�l�c�Q����ʒu�Q�_
        .HWFBM2SR = rs("E029HWFBM2SR")           ' �i�v�e�a�l�c�Q����ʒu�Q��
        .HWFBM2HT = rs("E029HWFBM2HT")           ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
        .HWFBM2HS = rs("E029HWFBM2HS")           ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��

        'BMD�ׂ��搔�ύX�Ή��@2003/05/19 osawa
        .HWFBM3AN = rs("E029HWFBM3AN")           ' �i�v�e�a�l�c�R���ω���
        .HWFBM3AX = rs("E029HWFBM3AX")           ' �i�v�e�a�l�c�R���Ϗ��
        .HWFBM3SH = rs("E029HWFBM3SH")           ' �i�v�e�a�l�c�R����ʒu�Q��
        .HWFBM3ST = rs("E029HWFBM3ST")           ' �i�v�e�a�l�c�R����ʒu�Q�_
        .HWFBM3SR = rs("E029HWFBM3SR")           ' �i�v�e�a�l�c�R����ʒu�Q��
        .HWFBM3HT = rs("E029HWFBM3HT")           ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
        .HWFBM3HS = rs("E029HWFBM3HS")           ' �i�v�e�a�l�c�R�ۏؕ��@�Q��

        If IsNull(rs("E029HWFBM1MBP")) = True Then    ' �i�v�e�a�l�c�P�ʓ����z�@��2003/05/14 ooba
            .HWFBM1MBP = -1
        Else
            .HWFBM1MBP = rs("E029HWFBM1MBP")
        End If

        If IsNull(rs("E029HWFBM2MBP")) = True Then    ' �i�v�e�a�l�c�Q�ʓ����z
            .HWFBM2MBP = -1
        Else
            .HWFBM2MBP = rs("E029HWFBM2MBP")
        End If

        If IsNull(rs("E029HWFBM3MBP")) = True Then    ' �i�v�e�a�l�c�R�ʓ����z
            .HWFBM3MBP = -1
        Else
            .HWFBM3MBP = rs("E029HWFBM3MBP")
        End If

        If IsNull(rs("E029HWFBM1MCL")) = False Then .HWFBM1MCL = rs("E029HWFBM1MCL")         ' �i�v�e�a�l�c�P�ʓ��v�Z
        If IsNull(rs("E029HWFBM2MCL")) = False Then .HWFBM2MCL = rs("E029HWFBM2MCL")         ' �i�v�e�a�l�c�Q�ʓ��v�Z
        If IsNull(rs("E029HWFBM3MCL")) = False Then .HWFBM3MCL = rs("E029HWFBM3MCL")         ' �i�v�e�a�l�c�R�ʓ��v�Z�@��2003/05/14 ooba

        .HWFOS1NS = rs("E025HWFOS1NS")           ' �i�v�e�_�f�͏o�P�M�����@
        .HWFOS2NS = rs("E025HWFOS2NS")           ' �i�v�e�_�f�͏o�Q�M�����@
        .HWFOS3NS = rs("E025HWFOS3NS")           ' �i�v�e�_�f�͏o�R�M�����@
        .HWFOF1NS = rs("E029HWFOF1NS")           ' �i�v�e�n�r�e�P�M�����@
        .HWFOF2NS = rs("E029HWFOF2NS")           ' �i�v�e�n�r�e�Q�M�����@
        .HWFOF3NS = rs("E029HWFOF3NS")           ' �i�v�e�n�r�e�R�M�����@
        .HWFOF4NS = rs("E029HWFOF4NS")           ' �i�v�e�n�r�e�S�M�����@
        .HWFBM1NS = rs("E029HWFBM1NS")           ' �i�v�e�a�l�c�P�M�����@
        .HWFBM2NS = rs("E029HWFBM2NS")           ' �i�v�e�a�l�c�Q�M�����@
        .HWFBM3NS = rs("E029HWFBM3NS")           ' �i�v�e�a�l�c�R�M�����@

        .HWFANTIM = rs("E025HWFANTIM")           ' �i�v�e�`�m����
        .HWFANTNP = rs("E025HWFANTNP")           ' �i�v�e�`�m���x

        .HWFOF1ET = rs("E029HWFOF1ET")           ' �i�v�e�n�r�e�P�I���d�s��
        .HWFOF2ET = rs("E029HWFOF2ET")           ' �i�v�e�n�r�e�Q�I���d�s��
        .HWFOF3ET = rs("E029HWFOF3ET")           ' �i�v�e�n�r�e�R�I���d�s��
        .HWFOF4ET = rs("E029HWFOF4ET")           ' �i�v�e�n�r�e�S�I���d�s��
        .HWFBM1ET = rs("E029HWFBM1ET")           ' �i�v�e�a�l�c�P�I���d�s��
        .HWFBM2ET = rs("E029HWFBM2ET")           ' �i�v�e�a�l�c�Q�I���d�s��
        .HWFBM3ET = rs("E029HWFBM3ET")           ' �i�v�e�a�l�c�R�I���d�s��

        .HWFOF1SZ = rs("E029HWFOF1SZ")           ' �i�v�e�n�r�e�P�������
        .HWFOF2SZ = rs("E029HWFOF2SZ")           ' �i�v�e�n�r�e�Q�������
        .HWFOF3SZ = rs("E029HWFOF3SZ")           ' �i�v�e�n�r�e�R�������
        .HWFOF4SZ = rs("E029HWFOF4SZ")           ' �i�v�e�n�r�e�S�������
        .HWFBM1SZ = rs("E029HWFBM1SZ")           ' �i�v�e�a�l�c�P�������
        .HWFBM2SZ = rs("E029HWFBM2SZ")           ' �i�v�e�a�l�c�Q�������
        .HWFBM3SZ = rs("E029HWFBM3SZ")           ' �i�v�e�a�l�c�R�������
    End With

    '����]�����ʎ擾
    sDBName = "Y013"
    If DBDRV_GetTBCMY013(Sokutei(), " where SAMPLEID='" & typIn.SAMPLEID & "' ", "order by OSITEM") = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '2001/08/20�@�r���[����u���b�NID���擾����悤�ɏC��
    '�u���b�NID�擾
    sDBName = "E040"

    'XSDCB���g�����ꍇ
    sSQL = "select "
    sSQL = sSQL & " BLOCKID "
    sSQL = sSQL & " from "
    sSQL = sSQL & " VECME013xsb "
    sSQL = sSQL & " where "
    sSQL = sSQL & " CRYNUM = '" & typ_AType.typ_Param.CRYNUMCA & "' "
    sSQL = sSQL & " and INPOSCB = " & typ_AType.typ_Param.INPOSCA & " "
    sSQL = sSQL & " and INGOTPOS >= 0 "

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    '���R�[�h0���̓G���[�I��
    If rs.RecordCount = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    lngRecCnt = rs.RecordCount

    ReDim siyou.BLOCKID(lngRecCnt)

    For i = 1 To lngRecCnt
        ' SXL�Ǘ�
        siyou.BLOCKID(i) = rs("BLOCKID")           ' �u���b�NID
        rs.MoveNext
    Next
    rs.Close


'2002/01/30 S.Sano Start
'    HWFRSPOT As String * 1          ' �i�v�e���R����ʒu�Q�_
'    HWFRSPOI As String * 1          ' �i�v�e���R����ʒu�Q��
'    HWFONSPT As String * 1          ' �i�v�e�_�f�Z�x����ʒu�Q�_
'    HWFONSPI As String * 1          ' �i�v�e�_�f�Z�x����ʒu�Q��
'�̑����
' ���i�d�lSXL�ް��P
'Public Type typ_TBCME018
'    HSXRSPOT As String * 1          ' �i�r�w���R����ʒu�Q�_
'    HSXRSPOI As String * 1          ' �i�r�w���R����ʒu�Q��
' ���i�d�lSXL�ް��Q
'Public Type typ_TBCME019
'    HSXONSPT As String * 1          ' �i�r�w�_�f�Z�x����ʒu�Q�_
'    HSXONSPI As String * 1          ' �i�r�w�_�f�Z�x����ʒu�Q��
'���g�p����B
    sSQL = "select "
    sSQL = sSQL & " HSXRSPOT, "
    sSQL = sSQL & " HSXRSPOI, "
    sSQL = sSQL & " HSXONSPT, "
    sSQL = sSQL & " HSXONSPI "
    sSQL = sSQL & " from "
    sSQL = sSQL & " TBCME018 K01, TBCME019 K12 "
    sSQL = sSQL & " where K01.HINBAN='" & typIn.HIN.hinban & "' and "
    sSQL = sSQL & " K01.MNOREVNO=" & typIn.HIN.mnorevno & " and "
    sSQL = sSQL & " K01.FACTORY='" & typIn.HIN.factory & "' and "
    sSQL = sSQL & " K01.OPECOND='" & typIn.HIN.opecond & "' and "
    sSQL = sSQL & " K12.HINBAN='" & typIn.HIN.hinban & "' and "
    sSQL = sSQL & " K12.MNOREVNO=" & typIn.HIN.mnorevno & " and "
    sSQL = sSQL & " K12.FACTORY='" & typIn.HIN.factory & "' and "
    sSQL = sSQL & " K12.OPECOND='" & typIn.HIN.opecond & "'"

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    '���R�[�h0���̓G���[�I��
    If rs.RecordCount = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With siyou
        .HWFRSPOT = rs("HSXRSPOT") ' �i�r�w���R����ʒu�Q�_
        .HWFRSPOI = rs("HSXRSPOI") ' �i�r�w���R����ʒu�Q��
        .HWFONSPT = rs("HSXONSPT") ' �i�r�w�_�f�Z�x����ʒu�Q�_
        .HWFONSPI = rs("HSXONSPI") ' �i�r�w�_�f�Z�x����ʒu�Q��
    End With

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001c_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************************
'*    �֐���        : DBDRV_scmzc_fcmlc001c_InsWfSougou
'*
'*    �����T�v      : 1.WF�������� WF����������ё}���p�c�a�h���C�o
'*
'*    �p�����[�^    : �ϐ���        ,IO  ,�^               ,����
'*                  :WFSougou       ,I   ,typ_TBCMW005     ,WF����������ё}���p
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_InsWfSougou(WfSougou As typ_TBCMW005) As FUNCTION_RETURN
    Dim sSQL As String

    'WF����������тւ̑}��
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_InsWfSougou"

    sSQL = "insert into TBCMW005 ( "
    sSQL = sSQL & "CRYNUM, "           ' �����ԍ�
    sSQL = sSQL & "INGOTPOS, "         ' �C���S�b�g���ʒu
    sSQL = sSQL & "TRANCNT, "          ' ������
    sSQL = sSQL & "CRYLEN, "           ' ����
    sSQL = sSQL & "KRPROCCD, "         ' �Ǘ��H���R�[�h
    sSQL = sSQL & "PROCCODE, "         ' �H���R�[�h
    sSQL = sSQL & "SXLID, "            ' SXLID
    sSQL = sSQL & "CODE, "             ' �敪�R�[�h
    sSQL = sSQL & "TSTAFFID, "         ' �o�^�Ј�ID
    sSQL = sSQL & "REGDATE, "          ' �o�^���t
    sSQL = sSQL & "KSTAFFID, "         ' �X�V�Ј�ID
    sSQL = sSQL & "UPDDATE, "          ' �X�V���t
    sSQL = sSQL & "SENDFLAG, "         ' ���M�t���O
    sSQL = sSQL & "SENDDATE, "        ' ���M���t
    sSQL = sSQL & "PLANTCAT) "          ' ����

    With WfSougou
        sSQL = sSQL & " select "
        sSQL = sSQL & " '" & .CRYNUM & "', "           ' �����ԍ�
        sSQL = sSQL & " " & .INGOTPOS & ", "           ' �C���S�b�g���ʒu
        sSQL = sSQL & " nvl(max(TRANCNT),0)+1, "       ' ������
        sSQL = sSQL & " " & .CRYLEN & ", "             ' ����
        sSQL = sSQL & " '" & .KRPROCCD & "', "         ' �Ǘ��H���R�[�h
        sSQL = sSQL & " '" & .PROCCODE & "', "         ' �H���R�[�h
        sSQL = sSQL & " '" & .SXLID & "', "            ' SXLID
        sSQL = sSQL & " '" & .CODE & "', "             ' �敪�R�[�h
        sSQL = sSQL & " '" & .TSTAFFID & "', "         ' �o�^�Ј�ID
        sSQL = sSQL & " sysdate, "                     ' �o�^���t
        sSQL = sSQL & " '" & .TSTAFFID & "', "         ' �X�V�Ј�ID
        sSQL = sSQL & " sysdate, "                     ' �X�V���t
        sSQL = sSQL & " '0', "                         ' ���M�t���O
        sSQL = sSQL & " sysdate "                      ' ���M���t
        sSQL = sSQL & " , '" & sCmbMukesaki & "'"      ' ���� 2007/09/04 SPK Tsutsumi Add
        sSQL = sSQL & " from TBCMW005 "
        sSQL = sSQL & " where CRYNUM='" & .CRYNUM & "' "
        sSQL = sSQL & " and INGOTPOS=" & .INGOTPOS
    End With

    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        DBDRV_scmzc_fcmlc001c_InsWfSougou = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmlc001c_InsWfSougou = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001c_InsWfSougou = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'************************************************************************************************************
'*    �֐���        : DBDRV_scmzc_fcmlc001c_UpdGDdata
'*
'*    �����T�v      : 1.WF�������� WF_GD���эX�V�p�c�a�h���C�o
'*
'*    �p�����[�^    : �ϐ���       ,IO  ,�^                  ,����
'*                  : UpdGD        ,I   ,typ_TBCMJ015        ,WF_GD���эX�V�p
'*                  : sStaffID     ,I   ,sStaffID            ,�S���Һ���
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_UpdGDdata(UpdGD As typ_TBCMJ015, sStaffID As String) As FUNCTION_RETURN
    Dim sSQL As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_UpdGDdata"

    '05/10/25 ooba START ============================================================>
    If UpdGD.MSRSDEN = -1 And UpdGD.MSRSLDL = -1 And UpdGD.MSRSDVD2 = -1 Then
        DBDRV_scmzc_fcmlc001c_UpdGDdata = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    '05/10/25 ooba END ==============================================================>

    With UpdGD
        sSQL = "UPDATE TBCMJ015 "
        sSQL = sSQL & "SET "

        If .MSRSDEN <> -1 Then      '05/10/25 ooba
            sSQL = sSQL & "MSRSDEN = " & .MSRSDEN & ", "        ' ���茋�� Den
        End If

        If .MSRSLDL <> -1 Then      '05/10/25 ooba
            sSQL = sSQL & "MSRSLDL = " & .MSRSLDL & ", "        ' ���茋�� L/DL
        End If

        If .MSRSDVD2 <> -1 Then     '05/10/25 ooba
            sSQL = sSQL & "MSRSDVD2 = " & .MSRSDVD2 & ", "      ' ���茋�� DVD2
        End If

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        If .MSZEROMN <> -1 Then
            sSQL = sSQL & "MSZEROMN = " & .MSZEROMN & ", "      ' L/DL0�A�����ŏ��l
        End If
        If .MSZEROMX <> -1 Then
            sSQL = sSQL & "MSZEROMX = " & .MSZEROMX & ", "      ' L/DL0�A�����ő�l
        End If
        sSQL = sSQL & "PTNJUDGRES = '" & .PTNJUDGRES & "', "      ' �p�^�[�����茋��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

        sSQL = sSQL & "KSTAFFID = '" & sStaffID & "', "         ' �X�V�Ј�ID
        sSQL = sSQL & "UPDDATE = SYSDATE "                      ' �X�V���t
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "CRYNUM = '" & .CRYNUM & "' "             ' �����ԍ�
        sSQL = sSQL & "AND POSITION = " & .POSITION & " "       ' �ʒu
        sSQL = sSQL & "AND SMPKBN = '" & .SMPKBN & "' "         ' �T���v���敪
        sSQL = sSQL & "AND TRANCOND = '" & .TRANCOND & "' "     ' ��������
        sSQL = sSQL & "AND TRANCNT = " & .TRANCNT & " "         ' ������
        sSQL = sSQL & "AND HSFLG = '" & .HSFLG & "' "           ' �ۏ؃t���O
        sSQL = sSQL & "AND SMPLNO = '" & .SMPLNO & "' "         ' �T���v���m��
    End With

    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        DBDRV_scmzc_fcmlc001c_UpdGDdata = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmlc001c_UpdGDdata = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001c_UpdGDdata = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************************************
'*    �֐���        : DBDRV_scmzc_fcmlc001c_UpdSXL1
'*
'*    �����T�v      : 1.WF�������� SXL�Ǘ��X�V�p�c�a�h���C�o�i���ݍH���A�ŏI�ʉߍH���X�V�j
'*                      (���݂́ATrue��Ԃ��Ă��邾��)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                                  ,����
'*                    SXL           ,O  ,type_DBDRV_scmzc_fcmlc001c_UpdSXL1  ,SXL�Ǘ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_UpdSXL1(sxl As type_DBDRV_scmzc_fcmlc001c_UpdSXL1) As FUNCTION_RETURN
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP���{
    DBDRV_scmzc_fcmlc001c_UpdSXL1 = FUNCTION_RETURN_SUCCESS
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP���{
End Function

'***********************************************************************************************************
'*    �֐���        : GetSxlidINBlkid
'*
'*    �����T�v      : 1.WF�������� SXL�Ǘ��X�V�p�c�a�h���C�o�i�폜�敪�A�ŏI��ԋ敪�X�V�j
'*
'*    �p�����[�^    : �ϐ���       ,IO  ,�^                 ,����
'*                    WFSoku       ,I   ,typ_TBCMW009       ,WF�Z���^�[�������葪��l�}���p
'*                    WFSougou     ,I   ,typ_TBCMW005       ,WF����������ё}���p
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_UpdSXL2(sxl As type_DBDRV_scmzc_fcmlc001c_UpdSXL2) As FUNCTION_RETURN
    Dim sSQL As String

    'SXL�Ǘ��̍X�V

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_UpdSXL2"

    sSQL = "update TBCME042 set "
    sSQL = sSQL & " DELCLS='" & sxl.DELCLS & "', "
    sSQL = sSQL & " LSTATCLS='" & sxl.LSTATCLS & "', "
    sSQL = sSQL & " UPDDATE=sysdate, "
    sSQL = sSQL & " SENDFLAG='0' "
    sSQL = sSQL & " where "
    sSQL = sSQL & " CRYNUM='" & sxl.CRYNUM & "' "
    sSQL = sSQL & " and INGOTPOS=" & sxl.INGOTPOS

    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        DBDRV_scmzc_fcmlc001c_UpdSXL2 = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmlc001c_UpdSXL2 = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'********************************************************************************************
'*    �֐���        : GetSxlidINBlkid
'*
'*    �����T�v      : 1.WF�������� �U�֔p�����ё}���p�c�a�h���C�o
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^                   ,����
'*                    Hurikae      ,I  ,typ_TBCMW006         ,�U�֔p�����ё}���p
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'********************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_InsHurikae(Hurikae As typ_TBCMW006) As FUNCTION_RETURN
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_InsHurikae"

    '�U�֔p�����тւ̑}��
    If DBDRV_Furikae_Ins(Hurikae) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmlc001c_InsHurikae = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmlc001c_InsHurikae = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'******************************************************************************************
'*    �֐���        : DBDRV_scmzc_fcmlc001c_InsSxlKakutei
'*
'*    �����T�v      : 1.WF�������� SXL�m��w���}���p�c�a�h���C�o
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^                  ,����
'*                    Hurikae      ,I  ,typ_TBCMY007        ,SXL�m��w��
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*                    (�g�p���Ă��Ȃ�)
'******************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_InsSxlKakutei(sxl As typ_TBCMY007) As FUNCTION_RETURN
    Dim sSQL As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_InsSxlKakutei"

    sSQL = "insert into TBCMY007 ("
    sSQL = sSQL & "SXL_ID, "           ' SXL-ID
    sSQL = sSQL & "SAMPLE_FROM, "      ' �T���v��ID (From)
    sSQL = sSQL & "SAMPLE_TO, "        ' �T���v��ID (To)
    sSQL = sSQL & "BLOCKID, "          ' �u���b�N�h�c
    sSQL = sSQL & "HINBAN, "           ' �m��i��
    sSQL = sSQL & "KUBUN, "            ' �敪�R�[�h
    sSQL = sSQL & "TXID, "             ' �g�����U�N�V����ID
    sSQL = sSQL & "REGDATE, "          ' �o�^���t
    sSQL = sSQL & "SUMMITSENDFLAG, "   ' SUMMIT���M�t���O
    sSQL = sSQL & "SENDFLAG, "         ' ���M�t���O
    sSQL = sSQL & "SENDDATE, "                ' ���M���t
    sSQL = sSQL & "PLANTCAT) "                ' ����

    With sxl
        sSQL = sSQL & "values ("
        sSQL = sSQL & " '" & .SXL_ID & "', "           ' SXL-ID
        sSQL = sSQL & " '" & .SAMPLE_FROM & "', "      ' �T���v��ID (From)
        sSQL = sSQL & " '" & .SAMPLE_TO & "', "        ' �T���v��ID (To)
        sSQL = sSQL & " '" & .BLOCKID & "', "          ' �u���b�N�h�c
        sSQL = sSQL & " '" & .hinban & "', "           ' �m��i��
        sSQL = sSQL & " '" & .KUBUN & "', "            ' �敪�R�[�h
        sSQL = sSQL & " '" & .TXID & "', "             ' �g�����U�N�V����ID
        sSQL = sSQL & " sysdate, "                     ' �o�^���t
        sSQL = sSQL & " '0', "                         ' SUMMIT���M�t���O
        sSQL = sSQL & " '0', "                         ' ���M�t���O
        sSQL = sSQL & " sysdate , "                    ' ���M���t
        sSQL = sSQL & " , '" & sCmbMukesaki & "'"      ' ���� 2007/09/04 SPK Tsutsumi Add
    End With

    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        DBDRV_scmzc_fcmlc001c_InsSxlKakutei = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmlc001c_InsSxlKakutei = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001c_InsSxlKakutei = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : GetSxlidINBlkid
'*
'*    �����T�v      : 1.SXL�̑S�u���b�N���Ƀ`�F�b�N
'*
'*    �p�����[�^    : �ϐ���        ,IO  ,�^             ,����
'*                    sCryNum       ,I   ,String         ,�����ԍ�
'*                    pTbcmh004()   ,O   ,typ_TBCMH004   ,���グ�I�����ю擾�p
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function s_cmmc001db_sSql039(ByVal sCryNum As String, _
                pTbcmh004() As typ_TBCMH004) As Double
    Dim sSQL    As String
    Dim intRet  As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function s_cmmc001db_sSql039"

    sSQL = " where CRYNUM = '" & sCryNum & "' "

    If DBDRV_GetTBCMH004(pTbcmh004, sSQL, "order by CRYNUM") = FUNCTION_RETURN_FAILURE Then
        s_cmmc001db_sSql039 = FUNCTION_RETURN_FAILURE
    Else
        s_cmmc001db_sSql039 = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    s_cmmc001db_sSql039 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'**********************************************************************************************************
'*    �֐���        : DBDRV_scmzc_fcmlc001c_UpdWfCrySmp
'*
'*    �����T�v      : 1.�������� WF�T���v���Ǘ��X�V�p�i�m��敪���P�ɍX�V�j
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^                                       ,����
'*                    WfCrySmp     ,O  ,type_DBDRV_scmzc_fcmlc001c_UpdWfCrySmp   ,�����T���v���Ǘ��X�V�p
'*                    sSXLID       ,I  ,String                                   ,SXLID 09/05/25 ooba
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'**********************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001c_UpdWfCrySmp(sSXLID As String) As FUNCTION_RETURN
''Public Function DBDRV_scmzc_fcmlc001c_UpdWfCrySmp(WfCrySmp() As type_DBDRV_scmzc_fcmlc001c_UpdWfCrySmp) _
''                 As FUNCTION_RETURN
    Dim sSQL As String
    Dim i   As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001c_UpdCrySmp"

''    For i = 1 To UBound(WfCrySmp)
        ' WF�T���v���Ǘ��̍X�V
        sSQL = "update XSDCW set "
        sSQL = sSQL & "  KTKBNCW='1' "          '�m��敪
        sSQL = sSQL & ", KDAYCW=sysdate "
        sSQL = sSQL & ", SNDKCW='0' "
        sSQL = sSQL & " where SXLIDCW = '" & sSXLID & "' "      '�����ύX(��߰�މ�) 09/05/25 ooba
''        sSql = sSql & " and XTALCW='" & WfCrySmp(i).CRYNUM & "' "
''        sSql = sSql & " and INPOSCW=" & WfCrySmp(i).INGOTPOS & " "
''        sSql = sSql & " and SMPKBNCW='" & WfCrySmp(i).SMPKBN & "' "

        If 0 >= OraDB.ExecuteSQL(sSQL) Then
            DBDRV_scmzc_fcmlc001c_UpdWfCrySmp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
''     Next

     DBDRV_scmzc_fcmlc001c_UpdWfCrySmp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_scmzc_fcmlc001c_UpdWfCrySmp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'*****************************************************************************************************
'*    �֐���        : DBDRV_GetTBCMH001039
'*
'*    �����T�v      : 1.�e�[�u���uTBCMH001�v��������ɂ��������R�[�h�𒊏o����
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^           ,����
'*                   records()     ,O  ,typ_TBCMH001 ,���o���R�[�h
'*                   sSqlWhere      ,I  ,String      ,���o����(SQL��Where��:�ȗ��\)
'*                   sSqlOrder      ,I  ,String      ,���o����(SQL��Order by��:�ȗ��\)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************************
Public Function DBDRV_GetTBCMH001039(records() As typ_TBCMH001, Optional sSqlWhere$ = vbNullString, _
                                        Optional sSqlOrder$ = vbNullString) As FUNCTION_RETURN
    Dim sSQL        As String       'SQL�S��
    Dim sSqlBase    As String       'SQL��{��(WHERE�߂̑O�܂�)
    Dim rs          As OraDynaset   'RecordSet
    Dim lngRecCnt   As Long         '���R�[�h��
    Dim i           As Long

    ''SQL��g�ݗ��Ă�
    sSqlBase = "Select UPINDNO, KRPROCCD, PROCCODE, MODEL, GOUKI, PGID, CPORGIND, HINBAN, NMNOREVNO, NFACTORY, NOPECOND, NUMNOTE1," & _
              " NUMNOTE2, SEED, SEKIERTB, DPNTCLS, DOPANT, AMRESIST, CRYDOPCL, CRYDOPVL, UPBTCHNM, ADDDOPCL, ADDDOPVL, ADDDOPPT," & _
              " BCNT1COD, BCNT1CMT, BCNT2COD, BCNT2CMT, MTCLS1, MTWGHT1, ESWGHT1, MTCLS2, MTWGHT2, ESWGHT2, MTCLS3, MTWGHT3," & _
              " ESWGHT3, MTCLS4, MTWGHT4, ESWGHT4, MTCLS5, MTWGHT5, ESWGHT5, MTCLS6, MTWGHT6, ESWGHT6, MTCLS7, MTWGHT7, ESWGHT7," & _
              " MTCLS8, MTWGHT8, ESWGHT8, MTCLS9, MTWGHT9, ESWGHT9, MTCLS10, MTWGHT10, ESWGHT10, MTCLS11, MTWGHT11, ESWGHT11," & _
              " MTCLS12, MTWGHT12, ESWGHT12, MTCLS13, MTWGHT13, ESWGHT13, MTCLS14, MTWGHT14, ESWGHT14, MTCLS15, MTWGHT15," & _
              " ESWGHT15, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sSqlBase = sSqlBase & "From TBCMH001"
    sSQL = sSqlBase

    If (sSqlWhere <> vbNullString) Or (sSqlOrder <> vbNullString) Then
        sSQL = sSQL & " " & sSqlWhere & " " & sSqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMH001039 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    lngRecCnt = rs.RecordCount
    ReDim records(lngRecCnt)
    For i = 1 To lngRecCnt
        With records(i)
            .UPINDNO = rs("UPINDNO")         ' ���グ�w��No.
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .MODEL = rs("MODEL")             ' �@��
            .GOUKI = rs("GOUKI")             ' ���@
            .PGID = rs("PGID")               ' PG-ID
            .CPORGIND = rs("CPORGIND")       ' ���ʌ��w��No
            .hinban = rs("HINBAN")           ' �i��
            .NMNOREVNO = rs("NMNOREVNO")     ' ���i�ԍ������ԍ�
            .NFACTORY = rs("NFACTORY")       ' �H��
            .NOPECOND = rs("NOPECOND")       ' ���Ə���
            .NUMNOTE1 = rs("NUMNOTE1")       ' �i�Ԕ��l�P
            .NUMNOTE2 = rs("NUMNOTE2")       ' �i�Ԕ��l�Q
            .SEED = rs("SEED")               ' �V�[�h
            .SEKIERTB = rs("SEKIERTB")       ' �Ήp���c�{
            .DPNTCLS = rs("DPNTCLS")         ' �h�[�p���g���
            .DOPANT = rs("DOPANT")           ' �h�[�p���g��
            .AMRESIST = rs("AMRESIST")       ' �˂炢��R
            .CRYDOPCL = rs("CRYDOPCL")       ' �����h�[�v���
            .CRYDOPVL = rs("CRYDOPVL")       ' �����h�[�v��
            .UPBTCHNM = rs("UPBTCHNM")       ' ���グ�o�b�`��
            .ADDDOPCL = rs("ADDDOPCL")       ' �ǉ��h�[�p���g���
            .ADDDOPVL = rs("ADDDOPVL")       ' �ǉ��h�[�p���g��
            .ADDDOPPT = rs("ADDDOPPT")       ' �ǉ��h�[�p���g�ʒu
            .BCNT1COD = rs("BCNT1COD")       ' �o�b�`���l1�i�R�[�h�j
            .BCNT1CMT = rs("BCNT1CMT")       ' �o�b�`���l1�i���āj
            .BCNT2COD = rs("BCNT2COD")       ' �o�b�`���l2�i�R�[�h�j
            .BCNT2CMT = rs("BCNT2CMT")       ' �o�b�`���l2�i���āj
            .MTCLS1 = rs("MTCLS1")           ' �������1
            .MTWGHT1 = rs("MTWGHT1")         ' �����d��1
            .ESWGHT1 = rs("ESWGHT1")         ' ����c�d��1
            .MTCLS2 = rs("MTCLS2")           ' �������2
            .MTWGHT2 = rs("MTWGHT2")         ' �����d��2
            .ESWGHT2 = rs("ESWGHT2")         ' ����c�d��2
            .MTCLS3 = rs("MTCLS3")           ' �������3
            .MTWGHT3 = rs("MTWGHT3")         ' �����d��3
            .ESWGHT3 = rs("ESWGHT3")         ' ����c�d��3
            .MTCLS4 = rs("MTCLS4")           ' �������4
            .MTWGHT4 = rs("MTWGHT4")         ' �����d��4
            .ESWGHT4 = rs("ESWGHT4")         ' ����c�d��4
            .MTCLS5 = rs("MTCLS5")           ' �������5
            .MTWGHT5 = rs("MTWGHT5")         ' �����d��5
            .ESWGHT5 = rs("ESWGHT5")         ' ����c�d��5
            .MTCLS6 = rs("MTCLS6")           ' �������6
            .MTWGHT6 = rs("MTWGHT6")         ' �����d��6
            .ESWGHT6 = rs("ESWGHT6")         ' ����c�d��6
            .MTCLS7 = rs("MTCLS7")           ' �������7
            .MTWGHT7 = rs("MTWGHT7")         ' �����d��7
            .ESWGHT7 = rs("ESWGHT7")         ' ����c�d��7
            .MTCLS8 = rs("MTCLS8")           ' �������8
            .MTWGHT8 = rs("MTWGHT8")         ' �����d��8
            .ESWGHT8 = rs("ESWGHT8")         ' ����c�d��8
            .MTCLS9 = rs("MTCLS9")           ' �������9
            .MTWGHT9 = rs("MTWGHT9")         ' �����d��9
            .ESWGHT9 = rs("ESWGHT9")         ' ����c�d��9
            .MTCLS10 = rs("MTCLS10")         ' �������10
            .MTWGHT10 = rs("MTWGHT10")       ' �����d��10
            .ESWGHT10 = rs("ESWGHT10")       ' ����c�d��10
            .MTCLS11 = rs("MTCLS11")         ' �������11
            .MTWGHT11 = rs("MTWGHT11")       ' �����d��11
            .ESWGHT11 = rs("ESWGHT11")       ' ����c�d��11
            .MTCLS12 = rs("MTCLS12")         ' �������12
            .MTWGHT12 = rs("MTWGHT12")       ' �����d��12
            .ESWGHT12 = rs("ESWGHT12")       ' ����c�d��12
            .MTCLS13 = rs("MTCLS13")         ' �������13
            .MTWGHT13 = rs("MTWGHT13")       ' �����d��13
            .ESWGHT13 = rs("ESWGHT13")       ' ����c�d��13
            .MTCLS14 = rs("MTCLS14")         ' �������14
            .MTWGHT14 = rs("MTWGHT14")       ' �����d��14
            .ESWGHT14 = rs("ESWGHT14")       ' ����c�d��14
            .MTCLS15 = rs("MTCLS15")         ' �������15
            .MTWGHT15 = rs("MTWGHT15")       ' �����d��15
            .ESWGHT15 = rs("ESWGHT15")       ' ����c�d��15
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMH001039 = FUNCTION_RETURN_SUCCESS
End Function

'*****************************************************************************************************
'*    �֐���        : DBDRV_GetTBCMH004
'*
'*    �����T�v      : 1.�e�[�u���uTBCMH004�v��������ɂ��������R�[�h�𒊏o����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                            ,����
'*                   records()     ,O  ,typ_TBCMH001 ,���o���R�[�h
'*                   sSqlWhere      ,I  ,String      ,���o����(SQL��Where��:�ȗ��\)
'*                   sSqlOrder      ,I  ,String      ,���o����(SQL��Order by��:�ȗ��\)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************************
Public Function DBDRV_GetTBCMH004(records() As typ_TBCMH004, Optional sSqlWhere$ = vbNullString, _
                                    Optional sSqlOrder$ = vbNullString) As FUNCTION_RETURN
    Dim sSQL        As String       'SQL�S��
    Dim sSqlBase    As String       'SQL��{��(WHERE�߂̑O�܂�)
    Dim rs          As OraDynaset   'RecordSet
    Dim lngRecCnt   As Long         '���R�[�h��
    Dim i           As Long

    ''SQL��g�ݗ��Ă�
    sSqlBase = "Select CRYNUM, KRPROCCD, PROCCODE, LENGTOP, LENGTKDO, LENGTAIL, LENGFREE, DM1, DM2, DM3, WGHTTOP, WGHTTKDO," & _
              " WGHTTAIL, WGHTFREE, WGTOPCUT, UPWEIGHT, CHARGE, SEED, STATCLS, JDGECODE, PWTIME, ADDDPPOS, ADDDPCLS, ADDDPVAL," & _
              " ADDDPNAM, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sSqlBase = sSqlBase & "From TBCMH004"
    sSQL = sSqlBase
    If (sSqlWhere <> vbNullString) Or (sSqlOrder <> vbNullString) Then
        sSQL = sSQL & " " & sSqlWhere & " " & sSqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMH004 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    lngRecCnt = rs.RecordCount
    ReDim records(lngRecCnt)
    For i = 1 To lngRecCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .LENGTOP = rs("LENGTOP")         ' �����iTOP�j
            .LENGTKDO = rs("LENGTKDO")       ' �����i�����j
            .LENGTAIL = rs("LENGTAIL")       ' �����iTAIL�j
            .LENGFREE = rs("LENGFREE")       ' �t���[����
            .DM1 = rs("DM1")                 ' �������a�P
            .DM2 = rs("DM2")                 ' �������a�Q
            .DM3 = rs("DM3")                 ' �������a�R
            .WGHTTOP = rs("WGHTTOP")         ' �d�ʁiTOP�j
            .WGHTTKDO = rs("WGHTTKDO")       ' �d�ʁi�����j
            .WGHTTAIL = rs("WGHTTAIL")       ' �d�ʁiTAIL)
            .WGHTFREE = rs("WGHTFREE")       ' �d�ʁi�t���[�����j
            .WGTOPCUT = rs("WGTOPCUT")       ' �g�b�v�J�b�g�d��
            .UPWEIGHT = rs("UPWEIGHT")       ' ���グ�d��
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .SEED = rs("SEED")               ' �V�[�h
            .STATCLS = rs("STATCLS")         ' BOT�󋵋敪
            .JDGECODE = rs("JDGECODE")       ' ����R�[�h
            .PWTIME = rs("PWTIME")           ' �p���[����
            .ADDDPPOS = rs("ADDDPPOS")       ' �ǉ��h�[�v�ʒu
            .ADDDPCLS = rs("ADDDPCLS")       ' �ǉ��h�[�p���g���
            .ADDDPVAL = rs("ADDDPVAL")       ' �ǉ��h�[�v��
            .ADDDPNAM = rs("ADDDPNAM")       ' �ǉ��h�[�v��
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMH004 = FUNCTION_RETURN_SUCCESS
End Function

'***********************************************************************************************
'*    �֐���        : DBDRV_scmzc_fcmlc001d_DispSiyou
'*
'*    �����T�v      : 1.�Ĕ����w�� �\���p�c�a�h���C�o�iWF�d�l�j
'*
'*    �p�����[�^    : �ϐ���       ,IO   ,�^                                 ,����
'*                    typIn        ,I    ,type_DBDRV_scmzc_fcmlc001d_In      ,���͗p
'*                    WfSiyou      ,I    ,type_DBDRV_scmzc_fcmlc001d_WfSiyou ,WF�d�l�擾�p
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function DBDRV_scmzc_fcmlc001d_DispSiyou(typIn() As type_DBDRV_scmzc_fcmlc001d_In, _
                                            WfSiyou() As type_DBDRV_scmzc_fcmlc001d_WfSiyou, _
                                            sErrMsg As String _
                                            ) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim i           As Long
    Dim lngInCnt    As Long
    Dim sDBName     As String
    Dim sOT1        As String
    Dim sOT2        As String
    Dim rtn         As FUNCTION_RETURN
    Dim sMAI1       As String
    Dim sMAI2       As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001d_DispSiyou"

    lngInCnt = UBound(typIn)

    ReDim WfSiyou(lngInCnt)
    sDBName = "(V001)"

    For i = 1 To lngInCnt
        DoEvents
        ' WF�d�l�̎擾
        sSQL = "select "
        sSQL = sSQL & "E021HWFRMIN, "            ' �i�v�e���R����
        sSQL = sSQL & "E021HWFRMAX, "            ' �i�v�e���R���
        sSQL = sSQL & "E021HWFRHWYS, "           ' �i�v�e���R�ۏؕ��@�Q��(Rs)
        sSQL = sSQL & "E025HWFONHWS, "           ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��(Oi)
        sSQL = sSQL & "E029HWFBM1HS, "           ' �i�v�e�a�l�c�P�ۏؕ��@�Q��(B1)
        sSQL = sSQL & "E029HWFBM2HS, "           ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��(B2)
        sSQL = sSQL & "E029HWFBM3HS, "           ' �i�v�e�a�l�c�R�ۏؕ��@�Q��(B3)
        sSQL = sSQL & "E029HWFOF1HS, "           ' �i�v�e�n�r�e�P�ۏؕ��@�Q��(L1)
        sSQL = sSQL & "E029HWFOF2HS, "           ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��(L2)
        sSQL = sSQL & "E029HWFOF3HS, "           ' �i�v�e�n�r�e�R�ۏؕ��@�Q��(L3)
        sSQL = sSQL & "E029HWFOF4HS, "           ' �i�v�e�n�r�e�S�ۏؕ��@�Q��(L4)
        sSQL = sSQL & "E026HWFDSOHS, "           ' �i�v�e�c�r�n�c�ۏؕ��@�Q��(DS)
        sSQL = sSQL & "E024HWFMKHWS, "           ' �i�v�e�����בw�ۏؕ��@�Q��(DZ)
        sSQL = sSQL & "E028HWFSPVHS, "           ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��(SP)
        sSQL = sSQL & "E028HWFDLHWS, "           ' �i�v�e�g�U���ۏؕ��@�Q��(KL)�@06/06/08 ooba
        sSQL = sSQL & "E025HWFOS1HS, "           ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��(D1)
        sSQL = sSQL & "E025HWFOS2HS, "           ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��(D2)
        sSQL = sSQL & "E025HWFOS3HS "            ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��(D3)
        sSQL = sSQL & " from VECME001 "
        sSQL = sSQL & " where E018HINBAN='" & typIn(i).HIN.hinban & "' " & _
                    " and E018MNOREVNO=" & typIn(i).HIN.mnorevno & " " & _
                    " and E018FACTORY='" & typIn(i).HIN.factory & "' " & _
                    " and E018OPECOND='" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        '���R�[�h0�����̓G���[
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & sDBName
            rs.Close
            GoTo proc_exit
        End If

        With WfSiyou(i)
            .HWFRMIN = fncNullCheck(rs("E021HWFRMIN"))            ' �i�v�e���R����
            .HWFRMAX = fncNullCheck(rs("E021HWFRMAX"))            ' �i�v�e���R���
            .HWFRHWYS = rs("E021HWFRHWYS")          ' �i�v�e���R�ۏؕ��@�Q��(Rs)
            .HWFONHWS = rs("E025HWFONHWS")          ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��(Oi)
            .HWFBM1HS = rs("E029HWFBM1HS")          ' �i�v�e�a�l�c�P�ۏؕ��@�Q��(B1)
            .HWFBM2HS = rs("E029HWFBM2HS")          ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��(B2)
            .HWFBM3HS = rs("E029HWFBM3HS")          ' �i�v�e�a�l�c�R�ۏؕ��@�Q��(B3)
            .HWFOF1HS = rs("E029HWFOF1HS")          ' �i�v�e�n�r�e�P�ۏؕ��@�Q��(L1)
            .HWFOF2HS = rs("E029HWFOF2HS")          ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��(L2)
            .HWFOF3HS = rs("E029HWFOF3HS")          ' �i�v�e�n�r�e�R�ۏؕ��@�Q��(L3)
            .HWFOF4HS = rs("E029HWFOF4HS")          ' �i�v�e�n�r�e�S�ۏؕ��@�Q��(L4)
            .HWFDSOHS = rs("E026HWFDSOHS")          ' �i�v�e�c�r�n�c�ۏؕ��@�Q��(DS)
            .HWFMKHWS = rs("E024HWFMKHWS")          ' �i�v�e�����בw�ۏؕ��@�Q��(DZ)
            .HWFSPVHS = rs("E028HWFSPVHS")          ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��(SP)
            .HWFDLHWS = rs("E028HWFDLHWS")          ' �i�v�e�g�U���ۏؕ��@�Q��(KL)�@06/06/08 ooba
            .HWFOS1HS = rs("E025HWFOS1HS")          ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��(D1)
            .HWFOS2HS = rs("E025HWFOS2HS")          ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��(D2)
            .HWFOS3HS = rs("E025HWFOS3HS")          ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��(D3)
            'rtn = scmzc_getE036(typIn(i).HIN, sOT1, sOT2)    '03/05/26
            rtn = scmzc_getE036(typIn(i).HIN, sOT1, sOT2, sMAI1, sMAI2)  '04/07/16
            If rtn = FUNCTION_RETURN_FAILURE Then
                rs.Close
                DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            .HWFOT1 = sOT1 '### 03/05/26
            .HWFOT2 = sOT2
            .HWFMAI1 = sMAI1  '04/07/16
            .HWFMAI2 = sMAI2
        End With

        ''�c���_�f�d�l�擾�ǉ��@03/12/15 ooba START ==============================>
        sSQL = "select HWFZOHWS from TBCME025 "
        sSQL = sSQL & "where HINBAN = '" & typIn(i).HIN.hinban & "' "
        sSQL = sSQL & "and MNOREVNO = " & typIn(i).HIN.mnorevno & " "
        sSQL = sSQL & "and FACTORY = '" & typIn(i).HIN.factory & "' "
        sSQL = sSQL & "and OPECOND = '" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        '���R�[�h0�����̓G���[
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & "(E025)"
            rs.Close
            GoTo proc_exit
        End If

        If IsNull(rs("HWFZOHWS")) = False Then WfSiyou(i).HWFZOHWS = rs("HWFZOHWS") Else WfSiyou(i).HWFZOHWS = " "  '�iWF�c���_�f�ۏؕ��@_��

        ''�c���_�f�d�l�`�F�b�N
        iChkAoi = ChkAoiSiyou(typIn(i).HIN)
        If iChkAoi < 0 Then
            sErrMsg = "�c���_�f(AOi)�d�l�G���["
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        ''�c���_�f�d�l�擾�ǉ��@03/12/15 ooba END ================================>

        '' GD�d�l�擾�@05/02/17 ooba START ==========================================>
        sDBName = "(E026)"
        sSQL = "select "
        sSQL = sSQL & "HWFDENHS, "                    ' �i�v�e�c�����ۏؕ��@�Q��(GD)
        sSQL = sSQL & "HWFDVDHS, "                    ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��(GD)
        sSQL = sSQL & "HWFLDLHS "                     ' �i�v�e�k�^�c�k�ۏؕ��@�Q��(GD)
        sSQL = sSQL & "from TBCME026 "
        sSQL = sSQL & "where HINBAN = '" & typIn(i).HIN.hinban & "' "
        sSQL = sSQL & "and MNOREVNO = " & typIn(i).HIN.mnorevno & " "
        sSQL = sSQL & "and FACTORY = '" & typIn(i).HIN.factory & "' "
        sSQL = sSQL & "and OPECOND = '" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        '���R�[�h0�����̓G���[
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & sDBName
            rs.Close
            GoTo proc_exit
        End If

        WfSiyou(i).HWFDENHS = rs("HWFDENHS")        ' �i�v�e�c�����ۏؕ��@�Q��(GD)
        WfSiyou(i).HWFDVDHS = rs("HWFDVDHS")        ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��(GD)
        WfSiyou(i).HWFLDLHS = rs("HWFLDLHS")        ' �i�v�e�k�^�c�k�ۏؕ��@�Q��(GD)
        '' GD�d�l�擾�@05/02/17 ooba END ============================================>

        '' SPVNr�Z�x�d�l�擾�@06/06/08 ooba START ===========================>
        sDBName = "E048"
        sSQL = "select "
        sSQL = sSQL & "HWFNRHS ,"         '�iWFSPVNR�ۏؕ��@_��
        sSQL = sSQL & "HWFSIRDHS "        '����]�ʕۏؕ��@�Q���@Add 2010/01/06 SIRD�Ή� Y.Hitomi
        sSQL = sSQL & "from TBCME048 "
        sSQL = sSQL & "where HINBAN = '" & typIn(i).HIN.hinban & "' "
        sSQL = sSQL & "and MNOREVNO = " & typIn(i).HIN.mnorevno & " "
        sSQL = sSQL & "and FACTORY = '" & typIn(i).HIN.factory & "' "
        sSQL = sSQL & "and OPECOND = '" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & sDBName
            rs.Close
            GoTo proc_exit
        End If

        If IsNull(rs("HWFNRHS")) = False Then WfSiyou(i).HWFNRHS = rs("HWFNRHS") Else WfSiyou(i).HWFNRHS = " "
        'Add 2010/01/06 SIRD�Ή� Y.Hitomi
        If IsNull(rs("HWFSIRDHS")) = False Then WfSiyou(i).HWFSIRDHS = rs("HWFSIRDHS") Else WfSiyou(i).HWFSIRDHS = " "

        rs.Close
        '' SPVNr�Z�x�d�l�擾�@06/06/08 ooba START ===========================>

        '�v�撷�擾
    '2004.09.08 Y.K �R�t���ύX
        sSQL = "select "
        sSQL = sSQL & " nvl(SUM(LENGTH),0) as Alllen"
        sSQL = sSQL & " from TBCME039 "
        sSQL = sSQL & " where substr(CRYNUM,1,9) = '" & Mid(typIn(i).CRYNUM, 1, 7) & "0" & Mid(typIn(i).CRYNUM, 9, 1) & "' " & _
                    " and HINBAN='" & typIn(i).HIN.hinban & "' " & _
                    " and REVNUM=" & typIn(i).HIN.mnorevno & " " & _
                    " and FACT='" & typIn(i).HIN.factory & "' " & _
                    " and OPCOND='" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            rs.Close
            GoTo proc_exit
        End If

        WfSiyou(i).KEIKAKUL = rs("Alllen")            ' �v�撷

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
        '' �G�s�d�l�擾(OSF�ABND)
        sDBName = "E050"
        sSQL = "select "
        sSQL = sSQL & "HEPOF1HS, "        '�iEPOSF1�ۏؕ��@_��
        sSQL = sSQL & "HEPOF2HS, "        '�iEPOSF2�ۏؕ��@_��
        sSQL = sSQL & "HEPOF3HS, "        '�iEPOSF3�ۏؕ��@_��
        sSQL = sSQL & "HEPBM1HS, "        '�iEPBMD1�ۏؕ��@_��
        sSQL = sSQL & "HEPBM2HS, "        '�iEPBMD2�ۏؕ��@_��
        sSQL = sSQL & "HEPBM3HS "         '�iEPBMD3�ۏؕ��@_��
        sSQL = sSQL & "from TBCME050 "    '���i�d�l�G�s�f�[�^�P
        sSQL = sSQL & "where HINBAN = '" & typIn(i).HIN.hinban & "' "
        sSQL = sSQL & "and MNOREVNO = " & typIn(i).HIN.mnorevno & " "
        sSQL = sSQL & "and FACTORY = '" & typIn(i).HIN.factory & "' "
        sSQL = sSQL & "and OPECOND = '" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & sDBName
            rs.Close
            GoTo proc_exit
        End If

        With WfSiyou(i)
            If IsNull(rs("HEPOF1HS")) = False Then .HEPOF1HS = rs("HEPOF1HS") Else .HEPOF1HS = " "   '�iEPOSF1�ۏؕ��@_��
            If IsNull(rs("HEPOF2HS")) = False Then .HEPOF2HS = rs("HEPOF2HS") Else .HEPOF2HS = " "   '�iEPOSF2�ۏؕ��@_��
            If IsNull(rs("HEPOF3HS")) = False Then .HEPOF3HS = rs("HEPOF3HS") Else .HEPOF3HS = " "   '�iEPOSF3�ۏؕ��@_��
            If IsNull(rs("HEPBM1HS")) = False Then .HEPBM1HS = rs("HEPBM1HS") Else .HEPBM1HS = " "   '�iEPBMD1�ۏؕ��@_��
            If IsNull(rs("HEPBM2HS")) = False Then .HEPBM2HS = rs("HEPBM2HS") Else .HEPBM2HS = " "   '�iEPBMD2�ۏؕ��@_��
            If IsNull(rs("HEPBM3HS")) = False Then .HEPBM3HS = rs("HEPBM3HS") Else .HEPBM3HS = " "   '�iEPBMD3�ۏؕ��@_��
        End With
        rs.Close
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
        '>>>>> ���Ԕ����K�i�Z�b�g�ǉ� 2011/07/15 Marushita
        sDBName = "E036"
        sSQL = "select "
        sSQL = sSQL & "MSMPFLG,     "     '���Ԕ����t���O
        sSQL = sSQL & "MSMPTANIMAI, "     '���Ԕ����P��(��)
        sSQL = sSQL & "MSMPCONSTMAI "     '���Ԕ������e�l
        sSQL = sSQL & "from TBCME036 "
        sSQL = sSQL & "where HINBAN = '" & typIn(i).HIN.hinban & "' "
        sSQL = sSQL & "and MNOREVNO = " & typIn(i).HIN.mnorevno & " "
        sSQL = sSQL & "and FACTORY = '" & typIn(i).HIN.factory & "' "
        sSQL = sSQL & "and OPECOND = '" & typIn(i).HIN.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_FAILURE
            ReDim WfSiyou(0)
            sErrMsg = GetMsgStr("EGET") & sDBName
            rs.Close
            GoTo proc_exit
        End If

        With WfSiyou(i)
            If IsNull(rs("MSMPFLG")) = False Then .CHUFLG = rs("MSMPFLG") Else .CHUFLG = "0"            '���Ԕ����t���O
            If IsNull(rs("MSMPTANIMAI")) = False Then .CHUTAN = rs("MSMPTANIMAI") Else .CHUTAN = "0"    '���Ԕ����P��(��)
            If IsNull(rs("MSMPCONSTMAI")) = False Then .CHUKYO = rs("MSMPCONSTMAI") Else .CHUKYO = "0"  '���Ԕ������e�l
        End With
        rs.Close
        '<<<<< ���Ԕ����K�i�Z�b�g�ǉ� 2011/07/15 Marushita
    Next

    DBDRV_scmzc_fcmlc001d_DispSiyou = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*********************************************************************************************
'*    �֐���        : DBDRV_scmzc_fcmlc001d_DispSmp
'*
'*    �����T�v      : 1.�Ĕ����w�� �\���p�c�a�h���C�o�iWF�T���v���Ǘ��j
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^                                ,����
'*                    inSXLID      ,I  ,String                            ,���͗pSXLID
'*                    WfSmp        ,O  ,type_DBDRV_scmzc_fcmlc001d_WfSmp  ,WF�T���v���Ǘ��p
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function DBDRV_scmzc_fcmlc001d_DispSmp(inSXLID As String, _
                                            WFSMP() As type_DBDRV_scmzc_fcmlc001d_WfSmp, _
                                            sErrMsg As String _
                                            ) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim i           As Long
    Dim sDBName     As String

    ' WF�T���v���Ǘ��擾
    ' �r���[VECME011(SXL�Ǘ����������A���̃u���b�N�ɑ΂���T���v����\������)���g�p

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001d_DispSmp"

    sDBName = "(XSDCW)"

    sSQL = "select "
    sSQL = sSQL & "XTALCW, "           ' �����ԍ�
    sSQL = sSQL & "INPOSCW, "         ' �������ʒu
    sSQL = sSQL & "SMPKBNCW, "           ' �T���v���敪
    sSQL = sSQL & "REPSMPLIDCW, "           ' �T���v��ID
    sSQL = sSQL & "HINBCW, "           ' �i��
    sSQL = sSQL & "REVNUMCW, "           ' ���i�ԍ������ԍ�
    sSQL = sSQL & "FACTORYCW, "          ' �H��
    sSQL = sSQL & "OPECW, "          ' ���Ə���
    sSQL = sSQL & "KTKBNCW, "            ' �m��敪
    sSQL = sSQL & "WFINDRSCW, "          ' ���FLG�iRs)
    sSQL = sSQL & "WFINDOICW, "          ' ���FLG�iOi)
    sSQL = sSQL & "WFINDB1CW, "          ' ���FLG�iB1)
    sSQL = sSQL & "WFINDB2CW, "          ' ���FLG�iB2�j
    sSQL = sSQL & "WFINDB3CW, "          ' ���FLG�iB3)
    sSQL = sSQL & "WFINDL1CW, "          ' ���FLG�iL1)
    sSQL = sSQL & "WFINDL2CW, "          ' ���FLG�iL2)
    sSQL = sSQL & "WFINDL3CW, "          ' ���FLG�iL3)
    sSQL = sSQL & "WFINDL4CW, "          ' ���FLG�iL4)
    sSQL = sSQL & "WFINDDSCW, "          ' ���FLG�iDS)
    sSQL = sSQL & "WFINDDZCW, "          ' ���FLG�iDZ)
    sSQL = sSQL & "WFINDSPCW, "          ' ���FLG�iSP)
    sSQL = sSQL & "WFINDDO1CW, "         ' ���FLG�iDO1)
    sSQL = sSQL & "WFINDDO2CW, "         ' ���FLG�iDO2)
    sSQL = sSQL & "WFINDDO3CW, "         ' ���FLG�iDO3)
    sSQL = sSQL & "WFINDOT1CW, "         ' ���FLG�iOT1)
    sSQL = sSQL & "WFINDOT2CW, "         ' ���FLG�iOT2)
    sSQL = sSQL & "WFINDAOICW, "         ' ���FLG (AOi)      '�c���_�f�ǉ��@03/12/15 ooba
    sSQL = sSQL & "WFINDGDCW, "          ' ���FLG (GD)       'GD�ǉ��@05/02/17 ooba
    sSQL = sSQL & "WFHSGDCW "            ' �ۏ�FLG (GD)       'GD�ǉ��@05/02/17 ooba
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    sSQL = sSQL & ",EPINDB1CW, "          ' ���FLG (OSF1E)
    sSQL = sSQL & "EPINDB2CW, "           ' ���FLG (OSF2E)
    sSQL = sSQL & "EPINDB3CW, "           ' ���FLG (OSF3E)
    sSQL = sSQL & "EPINDL1CW, "           ' ���FLG (BMD1E)
    sSQL = sSQL & "EPINDL2CW, "           ' ���FLG (BMD2E)
    sSQL = sSQL & "EPINDL3CW "            ' ���FLG (BMD3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    sSQL = sSQL & " from XSDCW"
    sSQL = sSQL & " where SXLIDCW='" & inSXLID & "' " & _
                " and LIVKCW='0' " & _
                " order by INPOSCW "


    Debug.Print sSQL
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
         DBDRV_scmzc_fcmlc001d_DispSmp = FUNCTION_RETURN_FAILURE
         ReDim WFSMP(0)
         sErrMsg = GetMsgStr("EGET") & sDBName
         rs.Close
         GoTo proc_exit
    End If

    lngRecCnt = rs.RecordCount
    ReDim WFSMP(lngRecCnt)

    For i = 1 To lngRecCnt
        DoEvents
        With WFSMP(i)
            .INGOTPOS = rs("INPOSCW")        ' �������ʒu
            .SMPLID = rs("REPSMPLIDCW")            ' �T���v��ID
            .hinban = rs("HINBCW")            ' �i��
            .REVNUM = rs("REVNUMCW")            ' ���i�ԍ������ԍ�
            .factory = rs("FACTORYCW")          ' �H��
            .opecond = rs("OPECW")          ' ���Ə���
            .WFINDRS = rs("WFINDRSCW")          ' ���FLG�iRs)
            .WFINDOI = rs("WFINDOICW")          ' ���FLG�iOi)
            .WFINDB1 = rs("WFINDB1CW")          ' ���FLG�iB1)
            .WFINDB2 = rs("WFINDB2CW")          ' ���FLG�iB2�j
            .WFINDB3 = rs("WFINDB3CW")          ' ���FLG�iB3)
            .WFINDL1 = rs("WFINDL1CW")          ' ���FLG�iL1)
            .WFINDL2 = rs("WFINDL2CW")          ' ���FLG�iL2)
            .WFINDL3 = rs("WFINDL3CW")          ' ���FLG�iL3)
            .WFINDL4 = rs("WFINDL4CW")          ' ���FLG�iL4)
            .WFINDDS = rs("WFINDDSCW")          ' ���FLG�iDS)
            .WFINDDZ = rs("WFINDDZCW")          ' ���FLG�iDZ)
            .WFINDSP = rs("WFINDSPCW")          ' ���FLG�iSP)
            .WFINDDO1 = rs("WFINDDO1CW")        ' ���FLG�iDO1)
            .WFINDDO2 = rs("WFINDDO2CW")        ' ���FLG�iDO2)
            .WFINDDO3 = rs("WFINDDO3CW")        ' ���FLG�iDO3)
            .WFINDOTHER1 = rs("WFINDOT1CW")        ' ���FLG�iDO2)
            .WFINDOTHER2 = rs("WFINDOT2CW")        ' ���FLG�iDO3)
            ''�c���_�f�ǉ�
            If IsNull(rs("WFINDAOICW")) = False Then .WFINDAOI = rs("WFINDAOICW")  ' ���FLG (AOi)
            .WFINDGD = rs("WFINDGDCW")          ' ���FLG�iGD)      'GD�ǉ��@05/02/17 ooba
            .WFHSGD = rs("WFHSGDCW")            ' �ۏ�FLG�iGD)      'GD�ǉ��@05/02/17 ooba
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
            .EPINDB1CW = rs("EPINDB1CW")
            .EPINDB2CW = rs("EPINDB2CW")
            .EPINDB3CW = rs("EPINDB3CW")
            .EPINDL1CW = rs("EPINDL1CW")
            .EPINDL2CW = rs("EPINDL2CW")
            .EPINDL3CW = rs("EPINDL3CW")
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_scmzc_fcmlc001d_DispSmp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*********************************************************************************************************
'*    �֐���        : DBDRV_scmzc_fcmlc001d_Exec
'*
'*    �����T�v      : 1.�Ĕ����w�� �X�V�A�}���p�c�a�h���C�o
'*
'*    �p�����[�^    : �ϐ���       ,IO   ,�^               ,����
'*                    WfSampleGr   ,I    ,typ_WfSampleGr   ,WF�T���v���Ǘ��p
'*                    SXL          ,O    ,typ_TBCME042     ,SXL�Ǘ��p
'*                                                          �i�����ԍ���Null��������X�V�A����ȊO�͑}���j
'*                    WfHantei     ,O    ,typ_TBCMW005     ,WF����������їp
'*                    HuriHai      ,O    ,typ_TBCMW006     ,�U�֔p�����їp
'*                    SokuSizi     ,O    ,typ_TBCMY003     ,����]�����@�w���}���p
'*                    SXLKakuSiji  ,O    ,typ_TBCMY007     ,SXL�m��w��
'*      �@            pEpMesInd �@ ,I    ,typ_TBCMY020   �@,EP����]���w��
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************************
Public Function DBDRV_scmzc_fcmlc001d_Exec(WfSampleGr() As typ_WfSampleGr, _
                                           sxl() As typ_TBCME042, _
                                           WfHantei As typ_TBCMW005, _
                                           HuriHai() As typ_TBCMW006, _
                                           SokuSizi() As typ_TBCMY003, _
                                           SXLKakuSiji() As typ_TBCMY007, _
                                           pEpMesInd() As typ_TBCMY020, _
                                           sErrMsg As String) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim i           As Long
    Dim sDBName     As String
    Dim intFromPos  As Integer
    Dim intToPos    As Integer
    Dim vGetFromPos As Variant
    Dim vGetToPos   As Variant
    Dim sTmpSxl()   As String     '�d�|�H���������pSXLID�@06/03/14 ooba

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001d_Exec"

    '' WriteDBLog "Start"

    'WF�T���v���Ǘ��ւ̑}��(���W���[�� s_cmzcDBdriverCOM_SQL.bas �g�p)
    '���V�T���v���Ǘ��ւ̑}���ɕύX�@2003/09/29 iida
    sDBName = "(XSDCW)" '�V�T���v���Ǘ�


    For i = 1 To UBound(WfSampleGr)
        If Trim(WfSampleGr(i).BLOCKID) = "" Then
            If WfSampleGr(i).WFSMP.REPSMPLIDCW <> "" Then  'SXL�̐擪�����L�̏ꍇ�A�T���v��ID�͐ݒ肳��Ă��Ȃ� 2003/04/22 okazaki
                With WfSampleGr(i).WFSMP
                    sSQL = "update XSDCW set "
                     sSQL = sSQL & "SXLIDCW = '" & .SXLIDCW & "',"          'SXLID
                    sSQL = sSQL & "HINBCW = '" & .HINBCW & "', "            ' �i��
                    sSQL = sSQL & "REVNUMCW = " & .REVNUMCW & ", "              ' ���i�ԍ������ԍ�
                    sSQL = sSQL & "FACTORYCW = '" & .FACTORYCW & "', "          ' �H��
                    sSQL = sSQL & "OPECW = '" & .OPECW & "', "          ' ���Ə���

                    ''�S�U�֎��̌���GD���p���Ή��@05/08/04 ooba START =====================>
                    If (i = 1 And bMotoGDcpyFlg(1)) Or _
                        (i = UBound(WfSampleGr) And bMotoGDcpyFlg(2)) Then

                        sSQL = sSQL & "WFSMPLIDGDCW= '" & .WFSMPLIDGDCW & "', "   '�T���v��ID(GD)
                        sSQL = sSQL & "WFINDGDCW= '" & .WFINDGDCW & "', "         '���FLG(GD)
                        sSQL = sSQL & "WFRESGDCW= '" & .WFRESGDCW & "', "         '����FLG(GD)
                        sSQL = sSQL & "WFHSGDCW= '" & .WFHSGDCW & "', "           '�ۏ�FLG(GD)
                    End If
                    ''�S�U�֎��̌���GD���p���Ή��@05/08/04 ooba END =======================>

                    sSQL = sSQL & "NUKISIFLGCW = '1', "                   ' �����w���ʉ߃t���O 09/05/26 ooba
                    sSQL = sSQL & "KDAYCW = sysdate , "                   ' �X�V���t
                    sSQL = sSQL & "SNDKCW = '0', "                       ' ���M�t���O
                    sSQL = sSQL & "SNDDAYCW = sysdate ,  "                   ' ���M���t
                    sSQL = sSQL & "KSTAFFCW = '" & .KSTAFFCW & "' "        '�X�V�Ј�ID

                    sSQL = sSQL & "where XTALCW ='" & .XTALCW & "' and "   ' �����ԍ�
                    sSQL = sSQL & "SXLIDCW = '" & tblSXL.SXLID & "' and "      ' SXLID 2003/11/05 ������ǉ�(�ŏ��ƍŌ�͌���SXLID���X�V)
                    sSQL = sSQL & "INPOSCW = " & .INPOSCW & " and "      ' �������ʒu
                    sSQL = sSQL & "SMPKBNCW = '" & .SMPKBNCW & "'"            ' �T���v���敪
                End With

                '' WriteDBLog sSql, sDBName
                If 1 <> OraDB.ExecuteSQL(sSQL) Then
                    DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
                    sErrMsg = GetMsgStr("EAPLY") & sDBName
                    GoTo proc_exit
                End If
            End If
        Else
            sSQL = "insert into XSDCW ( "
            sSQL = sSQL & "SXLIDCW, "             ' SXLID
            sSQL = sSQL & "SMPKBNCW, "            ' �T���v���敪
            sSQL = sSQL & "TBKBNCW, "             ' T/B�敪
            sSQL = sSQL & "REPSMPLIDCW, "         ' �T���v��ID
            sSQL = sSQL & "XTALCW, "              ' �����ԍ�
            sSQL = sSQL & "INPOSCW, "             ' �������ʒu
            sSQL = sSQL & "HINBCW, "              ' �i��
            sSQL = sSQL & "REVNUMCW, "            ' ���i�ԍ������ԍ�
            sSQL = sSQL & "FACTORYCW, "           ' �H��
            sSQL = sSQL & "OPECW, "               ' ���Ə���
            sSQL = sSQL & "KTKBNCW, "             ' �m��敪
            sSQL = sSQL & "SMCRYNUMCW, "          ' �T���v���u���b�NID
            sSQL = sSQL & "WFSMPLIDRSCW, "        ' �T���v��ID()
            sSQL = sSQL & "WFSMPLIDRS1CW, "       ' ����T���v��ID1
            sSQL = sSQL & "WFSMPLIDRS2CW, "       ' ����T���v��ID2
            sSQL = sSQL & "WFINDRSCW, "           ' ���FLG�iRs)
            sSQL = sSQL & "WFRESRS1CW, "           ' ����FLG1�iRs)
            sSQL = sSQL & "WFRESRS2CW, "           ' ����FLG2�iRs)
            sSQL = sSQL & "WFSMPLIDOICW, "        ' �T���v��ID(Oi)
            sSQL = sSQL & "WFINDOICW, "           ' ���FLG�iOi)
            sSQL = sSQL & "WFRESOICW, "           ' ����FLG�iOi)
            sSQL = sSQL & "WFSMPLIDB1CW, "        ' �T���v��ID(B1)
            sSQL = sSQL & "WFINDB1CW, "           ' ���FLG�iB1)
            sSQL = sSQL & "WFRESB1CW, "           ' ����FLG�iB1)
            sSQL = sSQL & "WFSMPLIDB2CW, "        ' �T���v��ID(B2)
            sSQL = sSQL & "WFINDB2CW, "           ' ���FLG�iB2�j
            sSQL = sSQL & "WFRESB2CW, "           ' ����FLG�iB2�j
            sSQL = sSQL & "WFSMPLIDB3CW, "        ' �T���v��ID(B3)
            sSQL = sSQL & "WFINDB3CW, "           ' ���FLG�iB3)
            sSQL = sSQL & "WFRESB3CW, "           ' ����FLG�iB3)
            sSQL = sSQL & "WFSMPLIDL1CW, "        ' �T���v��ID(L1)
            sSQL = sSQL & "WFINDL1CW, "           ' ���FLG�iL1)
            sSQL = sSQL & "WFRESL1CW, "           ' ����FLG�iL1)
            sSQL = sSQL & "WFSMPLIDL2CW, "        ' �T���v��ID(L2)
            sSQL = sSQL & "WFINDL2CW, "           ' ���FLG�iL2)
            sSQL = sSQL & "WFRESL2CW, "           ' ����FLG�iL2)
            sSQL = sSQL & "WFSMPLIDL3CW, "        ' �T���v��ID(L3)
            sSQL = sSQL & "WFINDL3CW, "           ' ���FLG�iL3)
            sSQL = sSQL & "WFRESL3CW, "           ' ����FLG�iL3)
            sSQL = sSQL & "WFSMPLIDL4CW, "        ' �T���v��ID(L4)
            sSQL = sSQL & "WFINDL4CW, "           ' ���FLG�iL4)
            sSQL = sSQL & "WFRESL4CW, "           ' ����FLG�iL4)
            sSQL = sSQL & "WFSMPLIDDSCW, "        ' �T���v��ID(DS)
            sSQL = sSQL & "WFINDDSCW, "           ' ���FLG�iDS)
            sSQL = sSQL & "WFRESDSCW, "           ' ����FLG�iDS)
            sSQL = sSQL & "WFSMPLIDDZCW, "        ' �T���v��ID(DZ)
            sSQL = sSQL & "WFINDDZCW, "           ' ���FLG�iDZ)
            sSQL = sSQL & "WFRESDZCW, "           ' ����FLG�iDZ)
            sSQL = sSQL & "WFSMPLIDSPCW, "        ' �T���v��ID(SP)
            sSQL = sSQL & "WFINDSPCW, "           ' ���FLG�iSP)
            sSQL = sSQL & "WFRESSPCW, "           ' ����FLG�iSP)
            sSQL = sSQL & "WFSMPLIDDO1CW, "       ' �T���v��ID(DO1)
            sSQL = sSQL & "WFINDDO1CW, "          ' ���FLG�iDO1)
            sSQL = sSQL & "WFRESDO1CW, "          ' ����FLG�iDO1)
            sSQL = sSQL & "WFSMPLIDDO2CW, "       ' �T���v��ID(DO2)
            sSQL = sSQL & "WFINDDO2CW, "          ' ���FLG�iDO2)
            sSQL = sSQL & "WFRESDO2CW, "          ' ����FLG�iDO2)
            sSQL = sSQL & "WFSMPLIDDO3CW, "       ' �T���v��ID(DO3)
            sSQL = sSQL & "WFINDDO3CW, "          ' ���FLG�iDO3)
            sSQL = sSQL & "WFRESDO3CW, "          ' ����FLG�iDO3)
             'add start 2003/05/26 hitec)�㓡 -------------------------
            sSQL = sSQL & "WFSMPLIDOT1CW, "       ' �T���v��ID(OT1)
            sSQL = sSQL & "WFINDOT1CW, "            ' ���FLG�iOT1)
            sSQL = sSQL & "WFRESOT1CW, "          ' ����FLG�iOT1)
            sSQL = sSQL & "WFSMPLIDOT2CW, "       ' �T���v��ID(OT2)
            sSQL = sSQL & "WFINDOT2CW, "            ' ���FLG�iOT2)
            sSQL = sSQL & "WFRESOT2CW, "          ' ����FLG�iOT2)
            'add end   2003/05/26 hitec)�㓡 -------------------------
            sSQL = sSQL & "WFSMPLIDAOICW, "       ' �T���v��ID(AOi)
            sSQL = sSQL & "WFINDAOICW, "          ' ���FLG(AOi)
            sSQL = sSQL & "WFRESAOICW, "          ' ����FLG(AOi)
            '' GD�ǉ��@05/02/21 ooba START =====================================>
            sSQL = sSQL & "WFSMPLIDGDCW, "        ' �T���v��ID (GD)
            sSQL = sSQL & "WFINDGDCW, "           ' ���FLG (GD)
            sSQL = sSQL & "WFRESGDCW, "           ' ����FLG (GD)
            sSQL = sSQL & "WFHSGDCW, "            ' �ۏ�FLG (GD)
            '' GD�ǉ��@05/02/21 ooba END =======================================>
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
            sSQL = sSQL & "EPSMPLIDL1CW, "    ' �T���v��ID (OSF1E)
            sSQL = sSQL & "EPINDL1CW, "       ' ���FLG (OSF1E)
            sSQL = sSQL & "EPRESL1CW, "       ' ����FLG (OSF1E)
            sSQL = sSQL & "EPSMPLIDL2CW, "    ' �T���v��ID (OSF2E)
            sSQL = sSQL & "EPINDL2CW, "       ' ���FLG (OSF2E)
            sSQL = sSQL & "EPRESL2CW, "       ' ����FLG (OSF2E)
            sSQL = sSQL & "EPSMPLIDL3CW, "    ' �T���v��ID (OSF3E)
            sSQL = sSQL & "EPINDL3CW, "       ' ���FLG (OSF3E)
            sSQL = sSQL & "EPRESL3CW, "       ' ����FLG (OSF3E)
            sSQL = sSQL & "EPSMPLIDB1CW, "    ' �T���v��ID (BMD1E)
            sSQL = sSQL & "EPINDB1CW, "       ' ���FLG (BMD1E)
            sSQL = sSQL & "EPRESB1CW, "       ' ����FLG (BMD1E)
            sSQL = sSQL & "EPSMPLIDB2CW, "    ' �T���v��ID (BMD2E)
            sSQL = sSQL & "EPINDB2CW, "       ' ���FLG (BMD2E)
            sSQL = sSQL & "EPRESB2CW, "       ' ����FLG (BMD2E)
            sSQL = sSQL & "EPSMPLIDB3CW, "    ' �T���v��ID (BMD3E)
            sSQL = sSQL & "EPINDB3CW, "       ' ���FLG (BMD3E)
            sSQL = sSQL & "EPRESB3CW, "       ' ����FLG (BMD3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
            sSQL = sSQL & "SMPLNUMCW, "           ' �T���v������
            sSQL = sSQL & "SMPLPATCW, "           ' �T���v���p�^�[��
            sSQL = sSQL & "LIVKCW, "              ' �����敪
            sSQL = sSQL & "NUKISIFLGCW, "         ' �����w���ʉ߃t���O 09/05/26 ooba
            sSQL = sSQL & "TSTAFFCW, "            ' �o�^�Ј�ID
            sSQL = sSQL & "TDAYCW, "              ' �o�^���t
            sSQL = sSQL & "KSTAFFCW, "            ' �X�V�Ј�ID
            sSQL = sSQL & "KDAYCW, "              ' �X�V���t
            sSQL = sSQL & "SNDKCW, "              ' ���M�t���O
            sSQL = sSQL & "SNDDAYCW ) "           ' ���M���t

            With WfSampleGr(i).WFSMP
                sSQL = sSQL & " values ('"
                sSQL = sSQL & .SXLIDCW & "', '"       ' SXLID
                sSQL = sSQL & .SMPKBNCW & "', '"      ' �T���v���敪
                sSQL = sSQL & .TBKBNCW & "', '"       ' T/B�敪
                sSQL = sSQL & .REPSMPLIDCW & "', '"   ' �T���v��ID
                sSQL = sSQL & .XTALCW & "', "         ' �����ԍ�
                sSQL = sSQL & .INPOSCW & ", '"        ' �������ʒu
                sSQL = sSQL & .HINBCW & "', "         ' �i��
                sSQL = sSQL & .REVNUMCW & ", '"       ' ���i�ԍ������ԍ�
                sSQL = sSQL & .FACTORYCW & "', '"     ' �H��
                sSQL = sSQL & .OPECW & "', '"         ' ���Ə���
                sSQL = sSQL & .KTKBNCW & "', '"       ' �m��敪
                sSQL = sSQL & .SMCRYNUMCW & "', '"    ' �T���v���u���b�NID
                sSQL = sSQL & .WFSMPLIDRSCW & "', "  ' �T���v��ID�iRs�j
                sSQL = sSQL & "Null, "               ' ����T���v��ID1�iRs�j
                sSQL = sSQL & "Null, '" ' ����T���v��ID2�iRs�j
                sSQL = sSQL & .WFINDRSCW & "', '"     ' ���FLG�iRs)
                sSQL = sSQL & .WFRESRS1CW & "' , "    ' ����FLG1�iRs)
                sSQL = sSQL & "Null, '"      ' ����FLG2�iRs)
                sSQL = sSQL & .WFSMPLIDOICW & "', '"  ' �T���v��ID�iOi�j
                sSQL = sSQL & .WFINDOICW & "', '"     ' ���FLG�iOi)
                sSQL = sSQL & .WFRESOICW & "', '"     ' ����FLG�iOi)
                sSQL = sSQL & .WFSMPLIDB1CW & "', '"  ' �T���v��ID�iB1�j
                sSQL = sSQL & .WFINDB1CW & "', '"     ' ���FLG�iB1)
                sSQL = sSQL & .WFRESB1CW & "', '"     ' ����FLG�iB1)
                sSQL = sSQL & .WFSMPLIDB2CW & "', '"  ' �T���v��ID�iB2�j
                sSQL = sSQL & .WFINDB2CW & "', '"     ' ���FLG�iB2)
                sSQL = sSQL & .WFRESB2CW & "', '"     ' ����FLG�iB2)
                sSQL = sSQL & .WFSMPLIDB3CW & "', '"  ' �T���v��ID�iB3�j
                sSQL = sSQL & .WFINDB3CW & "', '"     ' ���FLG�iB3)
                sSQL = sSQL & .WFRESB3CW & "', '"     ' ����FLG�iB3)
                sSQL = sSQL & .WFSMPLIDL1CW & "', '"  ' �T���v��ID�iL1�j
                sSQL = sSQL & .WFINDL1CW & "', '"     ' ���FLG�iL1)
                sSQL = sSQL & .WFRESL1CW & "', '"     ' ����FLG�iL1)
                sSQL = sSQL & .WFSMPLIDL2CW & "', '"  ' �T���v��ID�iL2�j
                sSQL = sSQL & .WFINDL2CW & "', '"     ' ���FLG�iL2)
                sSQL = sSQL & .WFRESL2CW & "', '"     ' ����FLG�iL2)
                sSQL = sSQL & .WFSMPLIDL3CW & "', '"  ' �T���v��ID�iL3�j
                sSQL = sSQL & .WFINDL3CW & "', '"     ' ���FLG�iL3)
                sSQL = sSQL & .WFRESL3CW & "', '"     ' ����FLG�iL3)
                sSQL = sSQL & .WFSMPLIDL4CW & "', '"  ' �T���v��ID�iL4�j
                sSQL = sSQL & .WFINDL4CW & "', '"     ' ���FLG�iL4)
                sSQL = sSQL & .WFRESL4CW & "', '"     ' ����FLG�iL4)
                sSQL = sSQL & .WFSMPLIDDSCW & "', '"  ' �T���v��ID�iDS�j
                sSQL = sSQL & .WFINDDSCW & "', '"     ' ���FLG�iDS)
                sSQL = sSQL & .WFRESDSCW & "', '"     ' ����FLG�iDS)
                sSQL = sSQL & .WFSMPLIDDZCW & "', '"  ' �T���v��ID�iDZ�j
                sSQL = sSQL & .WFINDDZCW & "', '"     ' ���FLG�iDZ)
                sSQL = sSQL & .WFRESDZCW & "', '"     ' ����FLG�iDZ)
                sSQL = sSQL & .WFSMPLIDSPCW & "', '"  ' �T���v��ID�iSP�j
                sSQL = sSQL & .WFINDSPCW & "', '"     ' ���FLG�iSP)
                sSQL = sSQL & .WFRESSPCW & "', '"     ' ����FLG�iSP)
                sSQL = sSQL & .WFSMPLIDDO1CW & "', '" ' �T���v��ID�iDO1�j
                sSQL = sSQL & .WFINDDO1CW & "', '"    ' ���FLG�iDO1)
                sSQL = sSQL & .WFRESDO1CW & "', '"    ' ����FLG�iDO1)
                sSQL = sSQL & .WFSMPLIDDO2CW & "', '" ' �T���v��ID�iDO2�j
                sSQL = sSQL & .WFINDDO2CW & "', '"    ' ���FLG�iDO2)
                sSQL = sSQL & .WFRESDO2CW & "', '"    ' ����FLG�iDO2)
                sSQL = sSQL & .WFSMPLIDDO3CW & "', '" ' �T���v��ID�iDO3�j
                sSQL = sSQL & .WFINDDO3CW & "', '"    ' ���FLG�iDO3)
                sSQL = sSQL & .WFRESDO3CW & "', '"    ' ����FLG�iDO3)
'                sSQL = sSQL & .WFSMPLIDOT1CW & "', '" ' �T���v��ID�iOT1�j
                sSQL = sSQL & "                ', '"   ' �T���v��ID�iOT1�j2010/04/08 Y.Hitomi OT1�b��Ή�
                sSQL = sSQL & .WFINDOT1CW & "', '"    ' ���FLG�iOT1)
                sSQL = sSQL & .WFRESOT1CW & "', '"    ' ����FLG�iOT1)
                sSQL = sSQL & .WFSMPLIDOT2CW & "', '" ' �T���v��ID�iOT2�j
                sSQL = sSQL & .WFINDOT2CW & "', '"    ' ���FLG�iOT2)
                sSQL = sSQL & .WFRESOT2CW & "', '"    ' ����FLG�iOT2)
                ''�c���_�f�ǉ��@03/12/15 ooba START ================================>
                sSQL = sSQL & .WFSMPLIDAOICW & "', '" ' �T���v��ID�iAOi�j
                sSQL = sSQL & .WFINDAOICW & "', '"    ' ���FLG�iAOi�j
                sSQL = sSQL & .WFRESAOICW & "', '"    ' ����FLG�iAOi�j
                ''�c���_�f�ǉ��@03/12/15 ooba END ==================================>
                '' GD�ǉ��@05/02/21 ooba START =====================================>
                sSQL = sSQL & .WFSMPLIDGDCW & "', '"  ' �T���v��ID (GD)
                sSQL = sSQL & .WFINDGDCW & "', '"     ' ���FLG (GD)
                sSQL = sSQL & .WFRESGDCW & "', '"     ' ����FLG (GD)
                sSQL = sSQL & .WFHSGDCW & "', "       ' �ۏ�FLG (GD)
                '' GD�ǉ��@05/02/21 ooba END =======================================>
    '--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                sSQL = sSQL & "'" & .EPSMPLIDL1CW & "', '"  ' �T���v��ID (OSF1E)
                sSQL = sSQL & .EPINDL1CW & "', '"         ' ���FLG (OSF1E)
                sSQL = sSQL & .EPRESL1CW & "', '"         ' ����FLG (OSF1E)
                sSQL = sSQL & .EPSMPLIDL2CW & "', '"      ' �T���v��ID (OSF2E)
                sSQL = sSQL & .EPINDL2CW & "', '"         ' ���FLG (OSF2E)
                sSQL = sSQL & .EPRESL2CW & "', '"         ' ����FLG (OSF2E)
                sSQL = sSQL & .EPSMPLIDL3CW & "', '"      ' �T���v��ID (OSF3E)
                sSQL = sSQL & .EPINDL3CW & "', '"         ' ���FLG (OSF3E)
                sSQL = sSQL & .EPRESL3CW & "', '"         ' ����FLG (OSF3E)
                sSQL = sSQL & .EPSMPLIDB1CW & "', '"      ' �T���v��ID (BMD1E)
                sSQL = sSQL & .EPINDB1CW & "', '"         ' ���FLG (BMD1E)
                sSQL = sSQL & .EPRESB1CW & "', '"         ' ����FLG (BMD1E)
                sSQL = sSQL & .EPSMPLIDB2CW & "', '"      ' �T���v��ID (BMD2E)
                sSQL = sSQL & .EPINDB2CW & "', '"         ' ���FLG (BMD2E)
                sSQL = sSQL & .EPRESB2CW & "', '"         ' ����FLG (BMD2E)
                sSQL = sSQL & .EPSMPLIDB3CW & "', '"      ' �T���v��ID (BMD3E)
                sSQL = sSQL & .EPINDB3CW & "', '"         ' ���FLG (BMD3E)
                sSQL = sSQL & .EPRESB3CW & "', "          ' ����FLG (BMD3E)
                sSQL = sSQL & "NULL, "                ' �T���v������
                sSQL = sSQL & "NULL, '"               ' �T���v���p�^�[��
                sSQL = sSQL & .LIVKCW & "', '"        ' �����敪
                sSQL = sSQL & "1', '"                 ' �����w���ʉ߃t���O 09/05/26 ooba
                sSQL = sSQL & .TSTAFFCW & "', "       ' �o�^�Ј�ID
                sSQL = sSQL & "sysdate, '"            ' �o�^���t
                sSQL = sSQL & .KSTAFFCW & "', "       ' �X�V�Ј�ID
                sSQL = sSQL & "sysdate, "             ' �X�V���t
                sSQL = sSQL & "'0', "                 ' ���M�t���O
                sSQL = sSQL & "sysdate)"              ' ���M���t
            End With

            '' WriteDBLog sSql, sDBName
            If 1 <> OraDB.ExecuteSQL(sSQL) Then
                DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
                sErrMsg = GetMsgStr("EAPLY") & sDBName
                GoTo proc_exit
            End If
        End If
    Next

    'Add Start 2011/04/25 SMPK Miyata
    sDBName = "(XSDCW_1)" '�V�T���v���Ǘ�(���Ԕ���)

    For i = 1 To UBound(sxl)
        sSQL = "update XSDCW_1 set "
        sSQL = sSQL & "SXLIDCW = '" & sxl(i).SXLID & "',"           ' SXLID
        sSQL = sSQL & "HINBCW = '" & sxl(i).hinban & "', "          ' �i��
        sSQL = sSQL & "REVNUMCW = " & sxl(i).REVNUM & ", "          ' ���i�ԍ������ԍ�
        sSQL = sSQL & "FACTORYCW = '" & sxl(i).factory & "', "      ' �H��
        sSQL = sSQL & "OPECW = '" & sxl(i).opecond & "', "          ' ���Ə���
        sSQL = sSQL & "NUKISIFLGCW = '1', "                         ' �����w���ʉ߃t���O 09/05/26 ooba
        sSQL = sSQL & "KDAYCW = sysdate , "                         ' �X�V���t
        sSQL = sSQL & "SNDKCW = '0', "                              ' ���M�t���O
        sSQL = sSQL & "SNDDAYCW = sysdate ,  "                      ' ���M���t
        sSQL = sSQL & "KSTAFFCW = '" & WfSampleGr(1).WFSMP.KSTAFFCW & "' "  ' �X�V�Ј�ID

        sSQL = sSQL & "where XTALCW ='" & tblSXL.CRYNUM & "' and "  ' �����ԍ�
        sSQL = sSQL & "SXLIDCW = '" & tblSXL.SXLID & "' and "       ' SXLID(����SXLID)
        sSQL = sSQL & "INPOSCW > " & sxl(i).INGOTPOS & " and "      ' �������ʒu
        sSQL = sSQL & "INPOSCW <= " & sxl(i).INGOTPOS + sxl(i).LENGTH   ' �������ʒu

        Call OraDB.ExecuteSQL(sSQL)

    Next i

    'Add End   2011/04/25 SMPK Miyata


    '�d�|�H���ă`�F�b�N�@�\�ǉ��@06/03/14 ooba
    ReDim sTmpSxl(1)
    sTmpSxl(1) = Trim(f_cmbc039_3.txtKSXLID.text)
    If DBDRV_CheckCodeXSDCB(sTmpSxl, PROCD_WFC_SOUGOUHANTEI, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ' SXL�Ǘ��ւ̍X�Vor�}��
    sDBName = "(XSDCB)"
    For i = 1 To UBound(sxl)
        'CRYNUM��Null�̎���SXLID�ōX�V
        If Trim(sxl(i).CRYNUM) = "" Then    '2003/04/20 okazaki

            '�T���v��ID�̕ύX�������ꍇ��SXLID���A�J�n�ʒu�ς��Ȃ�
            With sxl(i)
                sSQL = "update XSDCB set "
                sSQL = sSQL & "rlencb=" & .LENGTH & ", "          ' ����
                sSQL = sSQL & "gnwkntcb='" & .NOWPROC & "', "     ' ���ݍH��
                sSQL = sSQL & "newkntcb='" & .LASTPASS & "', "    ' �ŏI�ʉߍH��
                sSQL = sSQL & "livkcb='" & .DELCLS & "', "        ' �폜�敪
                sSQL = sSQL & "lstccb='" & .LSTATCLS & "', "      ' �ŏI��ԋ敪
                sSQL = sSQL & "sholdclscb='" & .HOLDCLS & "', "   ' �z�[���h�敪
                sSQL = sSQL & "hinbcb='" & .hinban & "', "        ' �i��
                sSQL = sSQL & "revnumcb=" & .REVNUM & ", "        ' ���i�ԍ������ԍ�
                sSQL = sSQL & "factorycb='" & .factory & "', "    ' �H��
                sSQL = sSQL & "opecb='" & .opecond & "', "        ' ���Ə���
                sSQL = sSQL & "furyccb='" & .BDCAUS & "', "       ' �s�Ǘ��R
                sSQL = sSQL & "maicb=" & .COUNT & ", "            ' ����
                sSQL = sSQL & "kdaycb=sysdate, "                  ' �X�V���t
                sSQL = sSQL & "sndkcb='0' "                       ' ���MFLG
                sSQL = sSQL & " where sxlidcb='" & .SXLID & "' "
            End With

            '' WriteDBLog sSql, sDBName
            If 0 >= OraDB.ExecuteSQL(sSQL) Then
                DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
                sErrMsg = GetMsgStr("EAPLY") & sDBName
                GoTo proc_exit
            End If
        Else
            sSQL = "insert into XSDCB ( "
            sSQL = sSQL & "xtalcb, "          ' �����ԍ�
            sSQL = sSQL & "inposcb, "         ' �������J�n�ʒu
            sSQL = sSQL & "rlencb, "          ' ����
            sSQL = sSQL & "sxlidcb, "         ' SXLID
            sSQL = sSQL & "gnwkntcb, "        ' ���ݍH��
            sSQL = sSQL & "newkntcb, "        ' �ŏI�ʉߍH��
            sSQL = sSQL & "livkcb, "          ' �폜�敪
            sSQL = sSQL & "lstccb, "          ' �ŏI��ԋ敪
            sSQL = sSQL & "sholdclscb, "      ' �z�[���h�敪
            sSQL = sSQL & "hinbcb, "          ' �i��
            sSQL = sSQL & "revnumcb, "        ' ���i�ԍ������ԍ�
            sSQL = sSQL & "factorycb, "       ' �H��
            sSQL = sSQL & "opecb, "           ' ���Ə���
            sSQL = sSQL & "furyccb, "         ' �s�Ǘ��R
            sSQL = sSQL & "maicb, "           ' ����
            sSQL = sSQL & "tdaycb, "          ' �o�^���t
            sSQL = sSQL & "kdaycb, "          ' �X�V���t
            sSQL = sSQL & "sndkcb, "          ' ���M�t���O
            sSQL = sSQL & "WSRMAICB, "        ' WS��㖇��
            sSQL = sSQL & "WSNMAICB, "        ' WS��򌇗�����
            sSQL = sSQL & "WFCMAICB, "        ' �������
            sSQL = sSQL & "SXLRMAICB, "       ' SXL�w��(�Ǖi)
            sSQL = sSQL & "WFCNMAICB, "       ' WFC����������
            sSQL = sSQL & "SXLEMAICB, "       ' SXL�m�薇��
            sSQL = sSQL & "SRMAICB, "         ' �T���v�����w������
            sSQL = sSQL & "SNMAICB, "         ' �T���v�����w���s�ǖ���
            sSQL = sSQL & "STMAICB, "         ' �T���v������
            sSQL = sSQL & "FURIMAICB, "       ' �U�֖���
            sSQL = sSQL & "XTWORKCB, "        ' �����H��
            sSQL = sSQL & "WFWORKCB, "        ' �E�F�[�n����
            sSQL = sSQL & "LUFRCCB, "         ' �i��R�[�h
            sSQL = sSQL & "LUFRBCB, "         ' �i��敪
            sSQL = sSQL & "LDERCCB, "         ' �i���R�[�h
            sSQL = sSQL & "HOLDCCB, "         ' �z�[���h�R�[�h
            sSQL = sSQL & "HOLDBCB, "         ' �z�[���h�敪
            sSQL = sSQL & "EXKUBCB, "         ' ��O�敪
            sSQL = sSQL & "HENPKCB, "         ' �ԕi�敪
            sSQL = sSQL & "KANKCB, "          ' �����敪
            sSQL = sSQL & "NFCB, "            ' ���ɋ敪
            sSQL = sSQL & "SAKJCB, "          ' �폜�敪
            sSQL = sSQL & "SUMITCB "          ' SUMIT���M�t���O
            sSQL = sSQL & " ) "

            With sxl(i)
                sSQL = sSQL & " values ( "
                sSQL = sSQL & " '" & .CRYNUM & "', "           ' �����ԍ�
                sSQL = sSQL & " " & .INGOTPOS & ", "           ' �������J�n�ʒu
                sSQL = sSQL & " " & .LENGTH & ", "             ' ����
                sSQL = sSQL & " '" & .SXLID & "', "            ' SXLID
                sSQL = sSQL & " '" & .NOWPROC & "', "          ' ���ݍH��
                sSQL = sSQL & " '" & .LASTPASS & "', "         ' �ŏI�ʉߍH��
                sSQL = sSQL & " '" & .DELCLS & "', "           ' �폜�敪
                sSQL = sSQL & " '" & .LSTATCLS & "', "         ' �ŏI��ԋ敪
                sSQL = sSQL & " '" & .HOLDCLS & "', "          ' �z�[���h�敪
                sSQL = sSQL & " '" & .hinban & "', "           ' �i��
                sSQL = sSQL & " " & .REVNUM & ", "             ' ���i�ԍ������ԍ�
                sSQL = sSQL & " '" & .factory & "', "          ' �H��
                sSQL = sSQL & " '" & .opecond & "', "          ' ���Ə���
                sSQL = sSQL & " '" & .BDCAUS & "', "           ' �s�Ǘ��R
                sSQL = sSQL & " " & .COUNT & ", "              ' ����
                sSQL = sSQL & " sysdate, "                     ' �o�^���t
                sSQL = sSQL & " sysdate, "                     ' �X�V���t
                sSQL = sSQL & " '0', "                         ' ���M�t���O
                sSQL = sSQL & " '0', "                         ' WS��㖇��
                sSQL = sSQL & " '0', "                         ' WS��򌇗�����
                sSQL = sSQL & " '0', "                         ' �������
                sSQL = sSQL & " '0', "                         ' SXL�w��(�Ǖi)
                sSQL = sSQL & " '0', "                         ' WFC����������
                sSQL = sSQL & " '0', "                         ' SXL�m�薇��
                sSQL = sSQL & " '0', "                         ' �T���v�����w������
                sSQL = sSQL & " '0', "                         ' �T���v�����w���s�ǖ���
                sSQL = sSQL & " '0', "                         ' �T���v������
                sSQL = sSQL & " '0', "                         ' �U�֖���
                sSQL = sSQL & " '42', "                        ' �����H��
                sSQL = sSQL & " '  ', "                        ' �E�F�[�n����
                sSQL = sSQL & " '   ', "                       ' �i��R�[�h
                sSQL = sSQL & " ' ', "                         ' �i��敪
                sSQL = sSQL & " '   ', "                       ' �i���R�[�h
                sSQL = sSQL & " '   ', "                       ' �z�[���h�R�[�h
                sSQL = sSQL & " '0', "                         ' �z�[���h�敪
                sSQL = sSQL & " ' ', "                         ' ��O�敪
                sSQL = sSQL & " ' ', "                         ' �ԕi�敪
                sSQL = sSQL & " '0', "                         ' �����敪
                sSQL = sSQL & " '0', "                         ' ���ɋ敪
                sSQL = sSQL & " '0', "                         ' �폜�敪
                sSQL = sSQL & " '0' "                          ' SUMIT���M�t���O
                sSQL = sSQL & " ) "
            End With

            '' WriteDBLog sSql, sDBName
            If 0 >= OraDB.ExecuteSQL(sSQL) Then
                DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
                sErrMsg = GetMsgStr("EAPLY") & sDBName
                GoTo proc_exit
            End If
        End If
    Next

    sDBName = "(W005)"

    ' WF����������тւ̑}��
    sSQL = "insert into TBCMW005 ( "
    sSQL = sSQL & "CRYNUM, "           ' �����ԍ�
    sSQL = sSQL & "INGOTPOS, "         ' �C���S�b�g�ʒu
    sSQL = sSQL & "TRANCNT, "          ' ������
    sSQL = sSQL & "CRYLEN, "           ' ����
    sSQL = sSQL & "KRPROCCD, "         ' �Ǘ��H���R�[�h
    sSQL = sSQL & "PROCCODE, "         ' �H���R�[�h
    sSQL = sSQL & "SXLID, "            ' SXLID
    sSQL = sSQL & "CODE, "             ' �敪�R�[�h
    sSQL = sSQL & "TSTAFFID, "         ' �o�^�Ј�ID
    sSQL = sSQL & "REGDATE, "          ' �o�^���t
    sSQL = sSQL & "KSTAFFID, "         ' �X�V�Ј�ID
    sSQL = sSQL & "UPDDATE, "          ' �X�V���t
    sSQL = sSQL & "SENDFLAG, "         ' ���M�t���O
    sSQL = sSQL & "SENDDATE, "        ' ���M���t
    sSQL = sSQL & "PLANTCAT) "          ' ����

    With WfHantei
        sSQL = sSQL & " select "
        sSQL = sSQL & " '" & .CRYNUM & "', "           ' �����ԍ�
        sSQL = sSQL & " " & .INGOTPOS & ", "           ' �C���S�b�g�ʒu
        sSQL = sSQL & " nvl(max(TRANCNT),0)+1, "       ' ������
        sSQL = sSQL & " " & .CRYLEN & ", "             ' ����
        sSQL = sSQL & " '" & .KRPROCCD & "', "         ' �Ǘ��H���R�[�h
        sSQL = sSQL & " '" & .PROCCODE & "', "         ' �H���R�[�h
        sSQL = sSQL & " '" & .SXLID & "', "            ' SXLID
        sSQL = sSQL & " '" & .CODE & "', "             ' �敪�R�[�h
        sSQL = sSQL & " '" & .TSTAFFID & "', "         ' �o�^�Ј�ID
        sSQL = sSQL & " sysdate, "                     ' �o�^���t
        sSQL = sSQL & " '" & .KSTAFFID & "', "         ' �X�V�Ј�ID
        sSQL = sSQL & " sysdate, "                     ' �X�V���t
        sSQL = sSQL & " '0', "                         ' ���M�t���O
        sSQL = sSQL & " sysdate "                      ' ���M���t
        sSQL = sSQL & " , '" & sCmbMukesaki & "'"      ' ���� 2007/09/04 SPK Tsutsumi Add
        sSQL = sSQL & " from TBCMW005 "
        sSQL = sSQL & " where CRYNUM='" & .CRYNUM & "' " & _
                    " and INGOTPOS=" & .INGOTPOS
    End With

    '' WriteDBLog sSql, sDBName
    If 0 >= OraDB.ExecuteSQL(sSQL) Then
        DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
        sErrMsg = GetMsgStr("EAPLY") & sDBName
        GoTo proc_exit
    End If

    sDBName = "(W006)"
    For i = 1 To UBound(HuriHai)
        ' �U�֔p�����тւ̑}�� (���W���[�� s_cmzcDBdriverCOM_SQL.bas �g�p)
        If DBDRV_Furikae_Ins(HuriHai(i)) = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
            sErrMsg = GetMsgStr("EAPLY") & sDBName
            GoTo proc_exit
        End If
    Next

    sDBName = "(Y003)"
    ' ����]�����@�w���ւ̑}�� (���W���[�� s_cmzcDBdriverCOM_SQL.bas �g�p)
    If DBDRV_SokuSizi_Ins(SokuSizi()) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
        sErrMsg = GetMsgStr("EAPLY") & sDBName
        GoTo proc_exit
    End If

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    '' �G�s����]���w�����̑}��(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    sDBName = "(Y020)"
    If DBDRV_SokuSizi_EP_Ins(pEpMesInd()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EAPLY") & sDBName
        GoTo proc_exit
    End If
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    sDBName = "(Y007)"
    '2001/08/04�@�ǉ�
    For i = 1 To UBound(SXLKakuSiji)
        ' SXL�m��w���ւ̑}��
        sSQL = "insert into TBCMY007 ("
        sSQL = sSQL & "SXL_ID, "           ' SXL-ID
        sSQL = sSQL & "SAMPLE_FROM, "      ' �T���v��ID (From)
        sSQL = sSQL & "SAMPLE_TO, "        ' �T���v��ID (To)
        sSQL = sSQL & "BLOCKID, "          ' �u���b�N�h�c
        sSQL = sSQL & "HINBAN, "           ' �m��i��
        sSQL = sSQL & "KUBUN, "            ' �敪�R�[�h
        sSQL = sSQL & "TXID, "             ' �g�����U�N�V����ID
        sSQL = sSQL & "REGDATE, "          ' �o�^���t
        sSQL = sSQL & "SUMMITSENDFLAG, "   ' SUMMIT���M�t���O
        sSQL = sSQL & "SENDFLAG, "         ' ���M�t���O
        sSQL = sSQL & "SENDDATE, "         ' ���M���t
        sSQL = sSQL & "PLANTCAT, "         ' ���� 2007/09/04 SPK Tsutsumi Add
        sSQL = sSQL & "MESDATA1TOP, "      ' ����l�P(Top)  center        '04/02/12 ooba START ====>
        sSQL = sSQL & "MESDATA2TOP, "      ' ����l�Q(Top)  R/2
        sSQL = sSQL & "MESDATA3TOP, "      ' ����l�R(Top)  Inside 10mm
        sSQL = sSQL & "MESDATA4TOP, "      ' ����l�S(Top)  Inside   6mm
        sSQL = sSQL & "MESDATA5TOP, "      ' ����l�T(Top)  Inside   3mm
        sSQL = sSQL & "MESDATA1BOT, "      ' ����l�P(Tail)  center
        sSQL = sSQL & "MESDATA2BOT, "      ' ����l�Q(Tail)  R/2
        sSQL = sSQL & "MESDATA3BOT, "      ' ����l�R(Tail)  Inside 10mm
        sSQL = sSQL & "MESDATA4BOT, "      ' ����l�S(Tail)  Inside   6mm
        sSQL = sSQL & "MESDATA5BOT )"      ' ����l�T(Tail)  Inside   3mm '04/02/12 ooba END ======>

        With SXLKakuSiji(i)
            sSQL = sSQL & "values ("
            sSQL = sSQL & " '" & .SXL_ID & "', "           ' SXL-ID
            sSQL = sSQL & " '" & .SAMPLE_FROM & "', "      ' �T���v��ID (From)
            sSQL = sSQL & " '" & .SAMPLE_TO & "', "        ' �T���v��ID (To)
            sSQL = sSQL & " '" & .BLOCKID & "', "          ' �u���b�N�h�c
            sSQL = sSQL & " '" & .hinban & "', "           ' �m��i��
            sSQL = sSQL & " '" & .KUBUN & "', "            ' �敪�R�[�h
            sSQL = sSQL & " '" & .TXID & "', "             ' �g�����U�N�V����ID
            sSQL = sSQL & " sysdate, "                     ' �o�^���t
            sSQL = sSQL & " '0', "                         ' SUMMIT���M�t���O
            sSQL = sSQL & " '3', "                         ' ���M�t���O   'upd 2003/06/05 hitec)matsumoto ���M�t���O��3�ɕύX
            sSQL = sSQL & " sysdate, "                     ' ���M���t
            sSQL = sSQL & " '" & sCmbMukesaki & "', "      ' ���� 2007/09/04 SPK Tsutsumi Add
            sSQL = sSQL & " '" & .MESDATA1TOP & "', "      ' ����l�P(Top)  center        '04/02/12 ooba START ====>
            sSQL = sSQL & " '" & .MESDATA2TOP & "', "      ' ����l�Q(Top)  R/2
            sSQL = sSQL & " '" & .MESDATA3TOP & "', "      ' ����l�R(Top)  Inside 10mm
            sSQL = sSQL & " '" & .MESDATA4TOP & "', "      ' ����l�S(Top)  Inside   6mm
            sSQL = sSQL & " '" & .MESDATA5TOP & "', "      ' ����l�T(Top)  Inside   3mm
            sSQL = sSQL & " '" & .MESDATA1BOT & "', "      ' ����l�P(Tail)  center
            sSQL = sSQL & " '" & .MESDATA2BOT & "', "      ' ����l�Q(Tail)  R/2
            sSQL = sSQL & " '" & .MESDATA3BOT & "', "      ' ����l�R(Tail)  Inside 10mm
            sSQL = sSQL & " '" & .MESDATA4BOT & "', "      ' ����l�S(Tail)  Inside   6mm
            sSQL = sSQL & " '" & .MESDATA5BOT & "' ) "     ' ����l�T(Tail)  Inside   3mm '04/02/12 ooba END ======>
        End With

        '' WriteDBLog sSql, sDBName
        If 0 >= OraDB.ExecuteSQL(sSQL) Then
            DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_FAILURE
            sErrMsg = GetMsgStr("EAPLY") & sDBName
            GoTo proc_exit
        End If

    Next

    '�֘A�u���b�N���o�^��~�@08/01/23 ooba
''    '�֘A��ۯ����o�^�@07/08/06 ooba START =====================================>
''    If UBound(tSXLID) > 0 Then
''        sDbName = "(Y023)"
''        If DBDRV_KanrenBlk(WfHantei.CRYNUM, tSXLID(), _
''                            SIngotP, EIngotP) = FUNCTION_RETURN_FAILURE Then
''
''            sErrMsg = GetMsgStr("EAPLY") & sDbName
''            GoTo proc_exit
''        End If
''    End If
''    '�֘A��ۯ����o�^�@07/08/06 ooba END =======================================>

    DBDRV_scmzc_fcmlc001d_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : GetSxlidINBlkid
'*
'*    �����T�v      : 1.�Ĕ����w�� �\���p�c�a�h���C�o�iWF�T���v���Ǘ��j
'*                      (�������̎擾)
'*
'*    �p�����[�^    : �ϐ���       ,IO   ,�^                ,����
'*                    CRYNUM       ,I    ,String            ,�����ԍ�
'*                    pLackWaf     ,O    ,typ_LackWaf       ,�������
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function DBDRV_scmzc_fcmlc001d_LostInfo(CRYNUM As String, _
                                            pLackWaf() As typ_LackWaf _
                                            ) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim lngRecCnt   As Long
    Dim i           As Long
    Dim rs          As OraDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_scmzc_fcmlc001d_LostInfo"

     '' �����E�F�n�[���̎擾
    sSQL = "select distinct LOTID as BLOCKID, REJWFFROM as WAFERNO, REJFROM as TOP_POS, REJTO as TAIL_POS "
    sSQL = sSQL & "from VECMW005 "
    sSQL = sSQL & "where (REJCAT='A' or ALLSCRAP='Y') "
    sSQL = sSQL & "  and LOTID like '" & left$(CRYNUM, 9) & "%'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    lngRecCnt = rs.RecordCount

    ReDim pLackWaf(lngRecCnt)

    For i = 1 To lngRecCnt
        DoEvents
        With pLackWaf(i)
            .BLOCKID = rs("BLOCKID")    ' �u���b�NID
            .WAFERNO = rs("WAFERNO")    ' �E�F�n�[�A��
            .TOP_POS = rs("TOP_POS")    ' �E�F�n�[�J�n�ʒu
            .TAIL_POS = rs("TAIL_POS")  ' �E�F�n�[�I���ʒu
        End With
        rs.MoveNext
    Next i
    rs.Close

    DBDRV_scmzc_fcmlc001d_LostInfo = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    '' WriteDBLog " ", "End"
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'********************************************************************************************
'*    �֐���        : DBDRV_GetDSODSpec
'*
'*    �����T�v      : 1.�iWFDSOD�������擾
'*                      (�g�p���Ă��Ȃ�)
'*    �p�����[�^    : �ϐ���        ,IO ,�^                      ,����
'*                    HIN           ,O  ,tFullHinban�@           ,�i�ԏ��
'*                    HWFDSOKE      ,O  ,String     �@           ,�iWFDSOD����
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'********************************************************************************************
Public Function DBDRV_GetDSODSpec(HIN As tFullHinban, HWFDSOKE As String) As FUNCTION_RETURN
    Dim sSQL    As String
    Dim rs      As OraDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_GetDSODSpec"
    DBDRV_GetDSODSpec = FUNCTION_RETURN_FAILURE

    sSQL = "select HWFDSOKE "
    sSQL = sSQL & "from TBCME026 "
    sSQL = sSQL & "where HINBAN = '" & HIN.hinban & "' and "
    sSQL = sSQL & "MNOREVNO = " & HIN.mnorevno & " and "
    sSQL = sSQL & "FACTORY = '" & HIN.factory & "' and "
    sSQL = sSQL & "OPECOND = '" & HIN.opecond & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        GoTo proc_exit
    End If

    HWFDSOKE = rs("HWFDSOKE") ' �iWFDSOD����

    rs.Close

    DBDRV_GetDSODSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************************
'*    �֐���        : DBDRV_GetDZSpec
'*
'*    �����T�v      : 1.�iWFDSOD�������擾
'*                      (�g�p���Ă��Ȃ�)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                      ,����
'*                    HIN           ,O  ,tFullHinban�@           ,�i�ԏ��
'*                    HWFMKSZY      ,O  ,String     �@           ,�iWF�����בw�������
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function DBDRV_GetDZSpec(HIN As tFullHinban, HWFMKSZY As String) As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_GetDZSpec"
    DBDRV_GetDZSpec = FUNCTION_RETURN_FAILURE

    sql = "select HWFDSOKE "
    sql = sql & "from TBCME026 "
    sql = sql & "where HINBAN = '" & HIN.hinban & "' and "
    sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
    sql = sql & "FACTORY = '" & HIN.factory & "' and "
    sql = sql & "OPECOND = '" & HIN.opecond & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        GoTo proc_exit
    End If

    HWFMKSZY = rs("HWFMKSZY") ' �iWF�����בw�������

    rs.Close

    DBDRV_GetDZSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'*****************************************************************************************************
'*    �֐���        : DBDRV_GetNoTestHinInfo
'*
'*    �����T�v      : 1.SXL�̑S�u���b�N���Ƀ`�F�b�N
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*                    HIN           ,I  ,tFullHinban    ,�i�ԏ��
'*                    Inf           ,O  ,NoTest_Info    ,
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************************
Public Function DBDRV_GetNoTestHinInfo(HIN() As tFullHinban, Inf() As NoTest_Info) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long
    Dim c0          As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_GetNoTestHinInfo"
    DBDRV_GetNoTestHinInfo = FUNCTION_RETURN_FAILURE

    For c0 = 0 To 1
        sSQL = "select "
        sSQL = sSQL & "HWFRHWYS " '�iWF���R�ۏؕ��@�Q��
        sSQL = sSQL & "from TBCME021 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).Res.HWFRHWYS = rs("HWFRHWYS")  '�iWF���R�ۏؕ��@�Q��
        rs.Close

        sSQL = "select "
        sSQL = sSQL & "HWFMKHWS, " '�iWF�����בw�ۏؕ��@�Q��
        sSQL = sSQL & "HWFMKSZY, " '�iWF�����בw�������
        sSQL = sSQL & "HWFMKSPH, " '�iWF�����בw����ʒu_��
        sSQL = sSQL & "HWFMKSPT, " '�iWF�����בw����ʒu_�_
        sSQL = sSQL & "HWFMKSPR " '�iWF�����בw����ʒu_��
        sSQL = sSQL & "from TBCME024 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).DZ.HWFMKHWS = rs("HWFMKHWS") '�iWF�����בw�ۏؕ��@�Q��
        Inf(c0).DZ.HWFMKSZY = rs("HWFMKSZY") '�iWF�����בw�������
        Inf(c0).DZ.HWFMKSPH = rs("HWFMKSPH") '�iWF�����בw����ʒu_��
        Inf(c0).DZ.HWFMKSPT = rs("HWFMKSPT") '�iWF�����בw����ʒu_�_
        Inf(c0).DZ.HWFMKSPR = rs("HWFMKSPR") '�iWF�����בw����ʒu_��
        rs.Close

        sSQL = "select "
        sSQL = sSQL & "HWFONHWS, " '�iWF�_�f�Z�x�ۏؕ��@�Q��
        sSQL = sSQL & "HWFOS1HS, " '�iWF�_�f�͏o1�ۏؕ��@�Q��
        sSQL = sSQL & "HWFOS1NS, " '�iWF�_�f�͏o1�M�����@
        sSQL = sSQL & "HWFOS1SH, " '�iWF�_�f�͏o1����ʒu_��
        sSQL = sSQL & "HWFOS1ST, " '�iWF�_�f�͏o1����ʒu_�_
        sSQL = sSQL & "HWFOS1SI, " '�iWF�_�f�͏o1����ʒu�Q��
        sSQL = sSQL & "HWFOS2HS, " '�iWF�_�f�͏o2�ۏؕ��@�Q��
        sSQL = sSQL & "HWFOS2NS, " '�iWF�_�f�͏o2�M�����@
        sSQL = sSQL & "HWFOS2SH, " '�iWF�_�f�͏o2����ʒu_��
        sSQL = sSQL & "HWFOS2ST, " '�iWF�_�f�͏o2����ʒu_�_
        sSQL = sSQL & "HWFOS2SI, " '�iWF�_�f�͏o2����ʒu�Q��
        sSQL = sSQL & "HWFOS3HS, " '�iWF�_�f�͏o3�ۏؕ��@�Q��
        sSQL = sSQL & "HWFOS3NS, " '�iWF�_�f�͏o3�M�����@
        sSQL = sSQL & "HWFOS3SH, " '�iWF�_�f�͏o3����ʒu_��
        sSQL = sSQL & "HWFOS3ST, " '�iWF�_�f�͏o3����ʒu_�_
        sSQL = sSQL & "HWFOS3SI " '�iWF�_�f�͏o3����ʒu�Q��
        sSQL = sSQL & "from TBCME025 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).Oi.HWFONHWS = rs("HWFONHWS") '�iWF�_�f�Z�x�ۏؕ��@�Q��
        Inf(c0).Doi(0).HWFOSxHS = rs("HWFOS1HS")  '�iWF�_�f�͏o1�ۏؕ��@�Q��
        Inf(c0).Doi(0).HWFOSxNS = rs("HWFOS1NS")  '�iWF�_�f�͏o1�M�����@
        Inf(c0).Doi(0).HWFOSxSH = rs("HWFOS1SH")  '�iWF�_�f�͏o1����ʒu_��
        Inf(c0).Doi(0).HWFOSxST = rs("HWFOS1ST")  '�iWF�_�f�͏o1����ʒu_�_
        Inf(c0).Doi(0).HWFOSxSI = rs("HWFOS1SI")  '�iWF�_�f�͏o1����ʒu�Q��
        Inf(c0).Doi(1).HWFOSxHS = rs("HWFOS2HS")  '�iWF�_�f�͏o2�ۏؕ��@�Q��
        Inf(c0).Doi(1).HWFOSxNS = rs("HWFOS2NS")  '�iWF�_�f�͏o2�M�����@
        Inf(c0).Doi(1).HWFOSxSH = rs("HWFOS2SH")  '�iWF�_�f�͏o2����ʒu_��
        Inf(c0).Doi(1).HWFOSxST = rs("HWFOS2ST")  '�iWF�_�f�͏o2����ʒu_�_
        Inf(c0).Doi(1).HWFOSxSI = rs("HWFOS2SI")  '�iWF�_�f�͏o2����ʒu�Q��
        Inf(c0).Doi(2).HWFOSxHS = rs("HWFOS3HS")  '�iWF�_�f�͏o3�ۏؕ��@�Q��
        Inf(c0).Doi(2).HWFOSxNS = rs("HWFOS3NS")  '�iWF�_�f�͏o3�M�����@
        Inf(c0).Doi(2).HWFOSxSH = rs("HWFOS3SH")  '�iWF�_�f�͏o3����ʒu_��
        Inf(c0).Doi(2).HWFOSxST = rs("HWFOS3ST")  '�iWF�_�f�͏o3����ʒu_�_
        Inf(c0).Doi(2).HWFOSxSI = rs("HWFOS3SI")  '�iWF�_�f�͏o3����ʒu�Q��
        rs.Close


        sSQL = "select "
        sSQL = sSQL & "HWFDSOHS, " '�iWFDSOD�ۏؕ��@�Q��
        sSQL = sSQL & "HWFDSOKE " '�iWFDSOD����"
        sSQL = sSQL & "from TBCME026 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).Dsod.HWFDSOHS = rs("HWFDSOHS") '�iWFDSOD�ۏؕ��@�Q��
        Inf(c0).Dsod.HWFDSOKE = rs("HWFDSOKE") '�iWFDSOD����"
        rs.Close

        sSQL = "select "
        sSQL = sSQL & "HWFSPVHS, " '�iWFSPVFE�ۏؕ��@�Q��
        sSQL = sSQL & "HWFSPVSH, " '�iWFSPVFE����ʒu_��
        sSQL = sSQL & "HWFSPVST, " '�iWFSPVFE����ʒu_�_
        sSQL = sSQL & "HWFSPVSI, " '�iWFSPVFE����ʒu�Q��
        sSQL = sSQL & "HWFDLHWS, " '�iWF�g�U���ۏؕ��@�Q��
        sSQL = sSQL & "HWFDLSPH, " '�iWF�g�U������ʒu_��
        sSQL = sSQL & "HWFDLSPT, " '�iWF�g�U������ʒu_�_
        sSQL = sSQL & "HWFDLSPI " '�iWF�g�U������ʒu�Q��
        sSQL = sSQL & "from TBCME028 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).SpvFe.HWFSPVHS = rs("HWFSPVHS") '�iWFSPVFE�ۏؕ��@�Q��
        Inf(c0).SpvFe.HWFSPVSH = rs("HWFSPVSH") '�iWFSPVFE����ʒu_��
        Inf(c0).SpvFe.HWFSPVST = rs("HWFSPVST") '�iWFSPVFE����ʒu_�_
        Inf(c0).SpvFe.HWFSPVSI = rs("HWFSPVSI") '�iWFSPVFE����ʒu�Q��
        Inf(c0).Spv.HWFDLHWS = rs("HWFDLHWS") '�iWF�g�U���ۏؕ��@�Q��
        Inf(c0).Spv.HWFDLSPH = rs("HWFDLSPH")  '�iWF�g�U������ʒu_��
        Inf(c0).Spv.HWFDLSPT = rs("HWFDLSPT") '�iWF�g�U������ʒu_�_
        Inf(c0).Spv.HWFDLSPI = rs("HWFDLSPI") '�iWF�g�U������ʒu�Q��
        rs.Close

        sSQL = "select "
        sSQL = sSQL & "HWFBM1HS, " '�iWFBMD1�ۏؕ��@�Q��
        sSQL = sSQL & "HWFBM1ET, " '�iWFBMD1�I��ET��
        sSQL = sSQL & "HWFBM1NS, " '�iWFBMD1�M�����@
        sSQL = sSQL & "HWFBM1SZ, " '�iWFBMD1�������
        sSQL = sSQL & "HWFBM1SH, " '�iWFBMD1����ʒu_��
        sSQL = sSQL & "HWFBM1ST, " '�iWFBMD1����ʒu_�_
        sSQL = sSQL & "HWFBM1SR, " '�iWFBMD1����ʒu_��
        sSQL = sSQL & "HWFBM2HS, " '�iWFBMD2�ۏؕ��@�Q��
        sSQL = sSQL & "HWFBM2ET, " '�iWFBMD2�I��ET��
        sSQL = sSQL & "HWFBM2NS, " '�iWFBMD2�M�����@
        sSQL = sSQL & "HWFBM2SZ, " '�iWFBMD2�������
        sSQL = sSQL & "HWFBM2SH, " '�iWFBMD2����ʒu_��
        sSQL = sSQL & "HWFBM2ST, " '�iWFBMD2����ʒu_�_
        sSQL = sSQL & "HWFBM2SR, " '�iWFBMD2����ʒu_��
        sSQL = sSQL & "HWFBM3HS, " '�iWFBMD3�ۏؕ��@�Q��
        sSQL = sSQL & "HWFBM3ET, " '�iWFBMD3�I��ET��
        sSQL = sSQL & "HWFBM3NS, " '�iWFBMD3�M�����@
        sSQL = sSQL & "HWFBM3SZ, " '�iWFBMD3�������
        sSQL = sSQL & "HWFBM3SH, " '�iWFBMD3����ʒu_��
        sSQL = sSQL & "HWFBM3ST, " '�iWFBMD3����ʒu_�_
        sSQL = sSQL & "HWFBM3SR, " '�iWFBMD3����ʒu_��
        sSQL = sSQL & "HWFOF1HS, " '�iWFOSF1�ۏؕ��@�Q��
        sSQL = sSQL & "HWFOF1ET, " '�iWFOSF1�I��ET��
        sSQL = sSQL & "HWFOF1NS, " '�iWFOSF1�M�����@
        sSQL = sSQL & "HWFOF1SZ, " '�iWFOSF1�������
        sSQL = sSQL & "HWFOF1SH, " '�iWFOSF1����ʒu_��
        sSQL = sSQL & "HWFOF1ST, " '�iWFOSF1����ʒu_�_
        sSQL = sSQL & "HWFOF1SR, " '�iWFOSF1����ʒu_��
        sSQL = sSQL & "HWFOF2HS, " '�iWFOSF2�ۏؕ��@�Q��
        sSQL = sSQL & "HWFOF2ET, " '�iWFOSF2�I��ET��
        sSQL = sSQL & "HWFOF2NS, " '�iWFOSF2�M�����@
        sSQL = sSQL & "HWFOF2SZ, " '�iWFOSF2�������
        sSQL = sSQL & "HWFOF2SH, " '�iWFOSF2����ʒu_��
        sSQL = sSQL & "HWFOF2ST, " '�iWFOSF2����ʒu_�_
        sSQL = sSQL & "HWFOF2SR, " '�iWFOSF2����ʒu_��
        sSQL = sSQL & "HWFOF3HS, " '�iWFOSF3�ۏؕ��@�Q��
        sSQL = sSQL & "HWFOF3ET, " '�iWFOSF3�I��ET��
        sSQL = sSQL & "HWFOF3NS, " '�iWFOSF3�M�����@
        sSQL = sSQL & "HWFOF3SZ, " '�iWFOSF3�������
        sSQL = sSQL & "HWFOF3SH, " '�iWFOSF3����ʒu_��
        sSQL = sSQL & "HWFOF3ST, " '�iWFOSF3����ʒu_�_
        sSQL = sSQL & "HWFOF3SR, " '�iWFOSF3����ʒu_��
        sSQL = sSQL & "HWFOF4HS, " '�iWFOSF4�ۏؕ��@�Q��
        sSQL = sSQL & "HWFOF4ET, " '�iWFOSF4�I��ET��
        sSQL = sSQL & "HWFOF4NS, " '�iWFOSF4�M�����@
        sSQL = sSQL & "HWFOF4SZ, " '�iWFOSF4�������
        sSQL = sSQL & "HWFOF4SH, " '�iWFOSF4����ʒu_��
        sSQL = sSQL & "HWFOF4ST, " '�iWFOSF4����ʒu_�_
        sSQL = sSQL & "HWFOF4SR " '�iWFOSF4����ʒu_��
        sSQL = sSQL & "from TBCME029 "
        sSQL = sSQL & "where "
        sSQL = sSQL & "HINBAN = '" & HIN(c0).hinban & "' and "
        sSQL = sSQL & "MNOREVNO = " & HIN(c0).mnorevno & " and "
        sSQL = sSQL & "FACTORY = '" & HIN(c0).factory & "' and "
        sSQL = sSQL & "OPECOND = '" & HIN(c0).opecond & "'"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        lngRecCnt = rs.RecordCount

        If lngRecCnt = 0 Then
            rs.Close
            GoTo proc_exit
        End If

        Inf(c0).BMD(0).HWFBMxHS = rs("HWFBM1HS") '�iWFBMD1�ۏؕ��@�Q��
        Inf(c0).BMD(0).HWFBMxET = rs("HWFBM1ET") '�iWFBMD1�I��ET��
        Inf(c0).BMD(0).HWFBMxNS = rs("HWFBM1NS") '�iWFBMD1�M�����@
        Inf(c0).BMD(0).HWFBMxSZ = rs("HWFBM1SZ") '�iWFBMD1�������
        Inf(c0).BMD(0).HWFBMxSH = rs("HWFBM1SH") '�iWFBMD1����ʒu_��
        Inf(c0).BMD(0).HWFBMxST = rs("HWFBM1ST") '�iWFBMD1����ʒu_�_
        Inf(c0).BMD(0).HWFBMxSR = rs("HWFBM1SR") '�iWFBMD1����ʒu_��
        Inf(c0).BMD(1).HWFBMxHS = rs("HWFBM2HS") '�iWFBMD2�ۏؕ��@�Q��
        Inf(c0).BMD(1).HWFBMxET = rs("HWFBM2ET") '�iWFBMD2�I��ET��
        Inf(c0).BMD(1).HWFBMxNS = rs("HWFBM2NS") '�iWFBMD2�M�����@
        Inf(c0).BMD(1).HWFBMxSZ = rs("HWFBM2SZ") '�iWFBMD2�������
        Inf(c0).BMD(1).HWFBMxSH = rs("HWFBM2SH") '�iWFBMD2����ʒu_��
        Inf(c0).BMD(1).HWFBMxST = rs("HWFBM2ST") '�iWFBMD2����ʒu_�_
        Inf(c0).BMD(1).HWFBMxSR = rs("HWFBM2SR") '�iWFBMD2����ʒu_��
        Inf(c0).BMD(2).HWFBMxHS = rs("HWFBM3HS") '�iWFBMD3�ۏؕ��@�Q��
        Inf(c0).BMD(2).HWFBMxET = rs("HWFBM3ET") '�iWFBMD3�I��ET��
        Inf(c0).BMD(2).HWFBMxNS = rs("HWFBM3NS") '�iWFBMD3�M�����@
        Inf(c0).BMD(2).HWFBMxSZ = rs("HWFBM3SZ") '�iWFBMD3�������
        Inf(c0).BMD(2).HWFBMxSH = rs("HWFBM3SH") '�iWFBMD3����ʒu_��
        Inf(c0).BMD(2).HWFBMxST = rs("HWFBM3ST") '�iWFBMD3����ʒu_�_
        Inf(c0).BMD(2).HWFBMxSR = rs("HWFBM3SR") '�iWFBMD3����ʒu_��
        Inf(c0).OSF(0).HWFOFxHS = rs("HWFOF1HS") '�iWFOSF1�ۏؕ��@�Q��
        Inf(c0).OSF(0).HWFOFxET = rs("HWFOF1ET") '�iWFOSF1�I��ET��
        Inf(c0).OSF(0).HWFOFxNS = rs("HWFOF1NS") '�iWFOSF1�M�����@
        Inf(c0).OSF(0).HWFOFxSZ = rs("HWFOF1SZ") '�iWFOSF1�������
        Inf(c0).OSF(0).HWFOFxSH = rs("HWFOF1SH") '�iWFOSF1����ʒu_��
        Inf(c0).OSF(0).HWFOFxST = rs("HWFOF1ST") '�iWFOSF1����ʒu_�_
        Inf(c0).OSF(0).HWFOFxSR = rs("HWFOF1SR") '�iWFOSF1����ʒu_��
        Inf(c0).OSF(1).HWFOFxHS = rs("HWFOF2HS") '�iWFOSF2�ۏؕ��@�Q��
        Inf(c0).OSF(1).HWFOFxET = rs("HWFOF2ET") '�iWFOSF2�I��ET��
        Inf(c0).OSF(1).HWFOFxNS = rs("HWFOF2NS") '�iWFOSF2�M�����@
        Inf(c0).OSF(1).HWFOFxSZ = rs("HWFOF2SZ") '�iWFOSF2�������
        Inf(c0).OSF(1).HWFOFxSH = rs("HWFOF2SH") '�iWFOSF2����ʒu_��
        Inf(c0).OSF(1).HWFOFxST = rs("HWFOF2ST") '�iWFOSF2����ʒu_�_
        Inf(c0).OSF(1).HWFOFxSR = rs("HWFOF2SR") '�iWFOSF2����ʒu_��
        Inf(c0).OSF(2).HWFOFxHS = rs("HWFOF3HS") '�iWFOSF3�ۏؕ��@�Q��
        Inf(c0).OSF(2).HWFOFxET = rs("HWFOF3ET") '�iWFOSF3�I��ET��
        Inf(c0).OSF(2).HWFOFxNS = rs("HWFOF3NS") '�iWFOSF3�M�����@
        Inf(c0).OSF(2).HWFOFxSZ = rs("HWFOF3SZ") '�iWFOSF3�������
        Inf(c0).OSF(2).HWFOFxSH = rs("HWFOF3SH") '�iWFOSF3����ʒu_��
        Inf(c0).OSF(2).HWFOFxST = rs("HWFOF3ST") '�iWFOSF3����ʒu_�_
        Inf(c0).OSF(2).HWFOFxSR = rs("HWFOF3SR") '�iWFOSF3����ʒu_��
        Inf(c0).OSF(3).HWFOFxHS = rs("HWFOF4HS") '�iWFOSF4�ۏؕ��@�Q��
        Inf(c0).OSF(3).HWFOFxET = rs("HWFOF4ET") '�iWFOSF4�I��ET��
        Inf(c0).OSF(3).HWFOFxNS = rs("HWFOF4NS") '�iWFOSF4�M�����@
        Inf(c0).OSF(3).HWFOFxSZ = rs("HWFOF4SZ") '�iWFOSF4�������
        Inf(c0).OSF(3).HWFOFxSH = rs("HWFOF4SH") '�iWFOSF4����ʒu_��
        Inf(c0).OSF(3).HWFOFxST = rs("HWFOF4ST") '�iWFOSF4����ʒu_�_
        Inf(c0).OSF(3).HWFOFxSR = rs("HWFOF4SR") '�iWFOSF4����ʒu_��
        rs.Close
    Next

    DBDRV_GetNoTestHinInfo = FUNCTION_RETURN_SUCCESS

proc_exit:
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : scmzc_getWF
'*
'*    �����T�v      : 1.���i�d�lWF�f�[�^�̎擾�h���C�o
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^               ,����
'*�@�@                pSpWFSamp�@�@ ,IO ,typ_SpWFSamp   �@,WF�T���v���d�l
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function scmzc_getWF(pSpWFSamp As typ_SpWFSamp) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sSQL    As String
    Dim sOT1    As String
    Dim sOT2    As String
    Dim sMAI1   As String     '04/07/16
    Dim sMAI2   As String
    Dim rtn     As FUNCTION_RETURN

     '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function scmzc_getWF"

    '' ���i�d�l�̎擾
    'DK���x�ǉ�      08/08/25 Systech
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
'''    sSql = "select " & _
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
'''          ", NVL(U.HSXDKTMP, ' ') as HSXDKTMP" & _
'''          " from  VECME001,TBCME036 U" & _
'''          " where E018HINBAN='" & pSpWFSamp.HIN.hinban & "' and E018MNOREVNO=" & pSpWFSamp.HIN.mnorevno & _
'''          " and E018FACTORY='" & pSpWFSamp.HIN.factory & "' and E018OPECOND='" & pSpWFSamp.HIN.opecond & "'" & _
'''          " and U.HINBAN = E018HINBAN and U.MNOREVNO = E018MNOREVNO and U.FACTORY = E018FACTORY and U.OPECOND = E018OPECOND"

    sSQL = "select " & _
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
          ", NVL(U.HSXDKTMP, ' ') as HSXDKTMP" & _
          " from  VECME001,TBCME036 U" & _
          " where E018HINBAN='" & pSpWFSamp.HIN.hinban & "' and E018MNOREVNO=" & pSpWFSamp.HIN.mnorevno & _
          " and E018FACTORY='" & pSpWFSamp.HIN.factory & "' and E018OPECOND='" & pSpWFSamp.HIN.opecond & "'" & _
          " and U.HINBAN = E018HINBAN and U.MNOREVNO = E018MNOREVNO and U.FACTORY = E018FACTORY and U.OPECOND = E018OPECOND"
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
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
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        rtn = scmzc_getE036(pSpWFSamp.HIN, sOT1, sOT2, sMAI1, sMAI2)
        If rtn = FUNCTION_RETURN_FAILURE Then
            rs.Close
            scmzc_getWF = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        .HWOTHER1 = sOT1 '### 03/05/26
        .HWOTHER2 = sOT2
        .HWOTHER1MAI = sMAI1    '04/07/16
        .HWOTHER2MAI = sMAI2
    End With
    rs.Close

    '�����p�x_���ް��擾�@04/04/13 ooba START =================================================>
    sSQL = "select "
    sSQL = sSQL & "TBCME024.HWFANGZY, "               '�iWF����AN�޽�����@04/07/29 ooba
    sSQL = sSQL & "TBCME021.HWFRKHNN, "
    sSQL = sSQL & "TBCME025.HWFONKHN, "
    sSQL = sSQL & "TBCME029.HWFOF1KN, "
    sSQL = sSQL & "TBCME029.HWFOF2KN, "
    sSQL = sSQL & "TBCME029.HWFOF3KN, "
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
'''    sSql = sSql & "TBCME029.HWFOF4KN, "
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
    sSQL = sSQL & "TBCME029.HWFBM1KN, "
    sSQL = sSQL & "TBCME029.HWFBM2KN, "
    sSQL = sSQL & "TBCME029.HWFBM3KN, "
    sSQL = sSQL & "TBCME025.HWFOS1KN, "
    sSQL = sSQL & "TBCME025.HWFOS2KN, "
    sSQL = sSQL & "TBCME025.HWFOS3KN, "
    sSQL = sSQL & "TBCME026.HWFDSOKN, "
    sSQL = sSQL & "TBCME024.HWFMKKHN, "
    sSQL = sSQL & "TBCME028.HWFSPVKN, "
    sSQL = sSQL & "TBCME028.HWFDLKHN, "
    sSQL = sSQL & "TBCME025.HWFZOKHN, "
    sSQL = sSQL & "TBCME026.HWFGDKHN "                '�����p�x_��(GD)�@05/02/18 ooba
    sSQL = sSQL & "from TBCME021, TBCME024, TBCME025, TBCME026, TBCME028, TBCME029 "
    sSQL = sSQL & "where TBCME021.HINBAN = TBCME024.HINBAN "
    sSQL = sSQL & "and TBCME021.MNOREVNO = TBCME024.MNOREVNO "
    sSQL = sSQL & "and TBCME021.FACTORY = TBCME024.FACTORY "
    sSQL = sSQL & "and TBCME021.OPECOND = TBCME024.OPECOND "
    sSQL = sSQL & "and TBCME021.HINBAN = TBCME025.HINBAN "
    sSQL = sSQL & "and TBCME021.MNOREVNO = TBCME025.MNOREVNO "
    sSQL = sSQL & "and TBCME021.FACTORY = TBCME025.FACTORY "
    sSQL = sSQL & "and TBCME021.OPECOND = TBCME025.OPECOND "
    sSQL = sSQL & "and TBCME021.HINBAN = TBCME026.HINBAN "
    sSQL = sSQL & "and TBCME021.MNOREVNO = TBCME026.MNOREVNO "
    sSQL = sSQL & "and TBCME021.FACTORY = TBCME026.FACTORY "
    sSQL = sSQL & "and TBCME021.OPECOND = TBCME026.OPECOND "
    sSQL = sSQL & "and TBCME021.HINBAN = TBCME028.HINBAN "
    sSQL = sSQL & "and TBCME021.MNOREVNO = TBCME028.MNOREVNO "
    sSQL = sSQL & "and TBCME021.FACTORY = TBCME028.FACTORY "
    sSQL = sSQL & "and TBCME021.OPECOND = TBCME028.OPECOND "
    sSQL = sSQL & "and TBCME021.HINBAN = TBCME029.HINBAN "
    sSQL = sSQL & "and TBCME021.MNOREVNO = TBCME029.MNOREVNO "
    sSQL = sSQL & "and TBCME021.FACTORY = TBCME029.FACTORY "
    sSQL = sSQL & "and TBCME021.OPECOND = TBCME029.OPECOND "
    sSQL = sSQL & "and TBCME021.HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
    sSQL = sSQL & "and TBCME021.MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
    sSQL = sSQL & "and TBCME021.FACTORY = '" & pSpWFSamp.HIN.factory & "' "
    sSQL = sSQL & "and TBCME021.OPECOND = '" & pSpWFSamp.HIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With pSpWFSamp
        If IsNull(rs("HWFANGZY")) = False Then .HWFANGZY = rs("HWFANGZY") Else .HWFANGZY = " "  '�iWF����AN�޽�����@04/07/29 ooba
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
        If IsNull(rs("HWFGDKHN")) = False Then .HWFGDKHN = rs("HWFGDKHN") Else .HWFGDKHN = " "  '�����p�x_��(GD)�@05/02/18 ooba
    End With
    rs.Close
    '�����p�x_���ް��擾�@04/04/13 ooba END ===================================================>

    ''�c���_�f�d�l�擾�ǉ��@03/12/15 ooba START ================================================>
    sSQL = "select HWFZOHWS, HWFZOSPH, HWFZOSPT, HWFZOSPI, HWFZONSW from TBCME025 "
    sSQL = sSQL & "where HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & pSpWFSamp.HIN.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & pSpWFSamp.HIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    If IsNull(rs("HWFZOHWS")) = False Then pSpWFSamp.HWFZOHWS = rs("HWFZOHWS") Else pSpWFSamp.HWFZOHWS = " "  ' �������@(AO)
    If IsNull(rs("HWFZOSPH")) = False Then pSpWFSamp.HWFZOSPH = rs("HWFZOSPH") Else pSpWFSamp.HWFZOSPH = " "  ' ������@(AO)
    If IsNull(rs("HWFZOSPT")) = False Then pSpWFSamp.HWFZOSPT = rs("HWFZOSPT") Else pSpWFSamp.HWFZOSPT = " "  ' ����_��(AO)
    If IsNull(rs("HWFZOSPI")) = False Then pSpWFSamp.HWFZOSPI = rs("HWFZOSPI") Else pSpWFSamp.HWFZOSPI = " "  ' ����ʒu(AO)
    If IsNull(rs("HWFZONSW")) = False Then pSpWFSamp.HWFZONSW = rs("HWFZONSW") Else pSpWFSamp.HWFZONSW = " "  ' �M�����@(AO)

    rs.Close
    ''�c���_�f�d�l�擾�ǉ��@03/12/15 ooba END ==================================================>

    '' GD�d�l�擾�@05/02/18 ooba START ========================================================>
''Upd start (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�
    sSQL = "select "
    sSQL = sSQL & "T1.HWFGDSPH AS HWFGDSPH, "         '������@(GD)�@05/10/25 ooba
    sSQL = sSQL & "T1.HWFGDSPT AS HWFGDSPT, "         '����_��(GD)�@05/10/25 ooba
    sSQL = sSQL & "T1.HWFGDZAR AS HWFGDZAR, "         '���O�̈�(GD)�@05/10/25 ooba
    sSQL = sSQL & "T1.HWFDENHS AS HWFDENHS, "         '�������@(GD/DEN)
    sSQL = sSQL & "T1.HWFLDLHS AS HWFLDLHS, "         '�������@(GD/LDL)
    sSQL = sSQL & "T1.HWFDVDHS AS HWFDVDHS"           '�������@(GD/DVD2)
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga START   ---
    sSQL = sSQL & ",T1.HWFGDSZY AS HWFGDSZY"          '�������(GD)
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga END     ---
    sSQL = sSQL & ",T2.HWFGDLINE AS HWFGDLINE "       'ײݐ�
    sSQL = sSQL & "from TBCME026 T1,TBCME036 T2 "
    sSQL = sSQL & "where T1.HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
    sSQL = sSQL & "and T1.MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
    sSQL = sSQL & "and T1.FACTORY = '" & pSpWFSamp.HIN.factory & "' "
    sSQL = sSQL & "and T1.OPECOND = '" & pSpWFSamp.HIN.opecond & "' "
    sSQL = sSQL & "and T1.HINBAN = T2.HINBAN "
    sSQL = sSQL & "and T1.MNOREVNO = T2.MNOREVNO "
    sSQL = sSQL & "and T1.FACTORY = T2.FACTORY "
    sSQL = sSQL & "and T1.OPECOND = T2.OPECOND "
''Upd end   (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    If IsNull(rs("HWFGDSPH")) = False Then pSpWFSamp.HWFGDSPH = rs("HWFGDSPH") Else pSpWFSamp.HWFGDSPH = " "  '05/10/25 ooba
    If IsNull(rs("HWFGDSPT")) = False Then pSpWFSamp.HWFGDSPT = rs("HWFGDSPT") Else pSpWFSamp.HWFGDSPT = " "  '05/10/25 ooba
    If IsNull(rs("HWFGDZAR")) = False Then pSpWFSamp.HWFGDZAR = rs("HWFGDZAR") Else pSpWFSamp.HWFGDZAR = " "  '05/10/25 ooba
    If IsNull(rs("HWFDENHS")) = False Then pSpWFSamp.HWFDENHS = rs("HWFDENHS") Else pSpWFSamp.HWFDENHS = " "  '�������@(GD/DEN)
    If IsNull(rs("HWFLDLHS")) = False Then pSpWFSamp.HWFLDLHS = rs("HWFLDLHS") Else pSpWFSamp.HWFLDLHS = " "  '�������@(GD/LDL)
    If IsNull(rs("HWFDVDHS")) = False Then pSpWFSamp.HWFDVDHS = rs("HWFDVDHS") Else pSpWFSamp.HWFDVDHS = " "  '�������@(GD/DVD2)
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga START   ---
    If IsNull(rs("HWFGDSZY")) = False Then pSpWFSamp.HWFGDSZY = rs("HWFGDSZY") Else pSpWFSamp.HWFGDSZY = " "
'GDײ������@�\�ǉ� 2007/06/25 M.Kaga END     ---
''Upd Start (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�
    If IsNull(rs("HWFGDLINE")) = False Then pSpWFSamp.HWFGDLINE = CStr(rs("HWFGDLINE"))
''Upd End   (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�

    rs.Close
    '' GD�d�l�擾�@05/02/18 ooba END ==========================================================>

    '' SPV�d�l�擾�@06/06/08 ooba START ===============================================>
    sSQL = "select HWFNRHS, "                    '�iWFSPVNR�ۏؕ��@_��
    sSQL = sSQL & "HWFNRSH, "                     '�iWFSPVNR����ʒu_��
    sSQL = sSQL & "HWFNRST, "                     '�iWFSPVNR����ʒu_�_
    sSQL = sSQL & "HWFNRSI, "                     '�iWFSPVNR����ʒu_��
    sSQL = sSQL & "HWFNRKN, "                     '�iWFSPVNR�����p�x_��
    sSQL = sSQL & "HWFSPVPUG, "                   '�iWFSPVFEPUA��
    sSQL = sSQL & "HWFSPVPUR, "                   '�iWFSPVFEPUA��
    sSQL = sSQL & "HWFSPVSTD, "                   '�iWFSPVFE�W���΍�
    sSQL = sSQL & "HWFDLPUG, "                    '�iWF�g�U��PUA��
    sSQL = sSQL & "HWFDLPUR, "                    '�iWF�g�U��PUA��
    sSQL = sSQL & "HWFNRPUG, "                    '�iWFSPVNRPUA��
    sSQL = sSQL & "HWFNRPUR, "                    '�iWFSPVNRPUA��
    sSQL = sSQL & "HWFNRSTD "                     '�iWFSPVNR�W���΍�
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
    sSQL = sSQL & ",HWFSIRDMX, "                  '����]�ʏ��
    sSQL = sSQL & "HWFSIRDSZ, "                   '����]�ʑ������
    sSQL = sSQL & "HWFSIRDHT, "                   '����]�ʕۏؕ��@�Q��
    sSQL = sSQL & "HWFSIRDHS, "                   '����]�ʕۏؕ��@_��
    sSQL = sSQL & "HWFSIRDKM, "                   '����]�ʌ����p�x�Q��
    sSQL = sSQL & "HWFSIRDKN, "                   '����]�ʌ����p�x_��
    sSQL = sSQL & "HWFSIRDKH, "                   '����]�ʌ����p�x�Q��
    sSQL = sSQL & "HWFSIRDKU, "                   '����]�ʌ����p�x�Q�E
    sSQL = sSQL & "HWFSIRDPS  "                   '����]��TB�ۏ؈ʒu
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
    sSQL = sSQL & "from TBCME048 "
    sSQL = sSQL & "where HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & pSpWFSamp.HIN.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & pSpWFSamp.HIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
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
    sSQL = "select HWFIGKBN from TBCME017" & _
          " where HINBAN='" & pSpWFSamp.HIN.hinban & "' and MNOREVNO=" & pSpWFSamp.HIN.mnorevno & _
          " and FACTORY='" & pSpWFSamp.HIN.factory & "'"
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pSpWFSamp.HWFIGKBN = rs("HWFIGKBN")
    rs.Close

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    '' �G�s�d�l�擾(BMD1E�`BMD3E,OSF1E�`OSF3E)
    sSQL = "select HEPOF1NS, "                   ' �i�M�����@(OSF1E)
    sSQL = sSQL & "HEPOF1SZ, "                    ' �i�������(OSF1E)
    sSQL = sSQL & "HEPOF1ET, "                    ' �i�I��ET��(OSF1E)
    sSQL = sSQL & "HEPOF1HS, "                    ' �i�ۏؕ��@_��(OSF1E)
    sSQL = sSQL & "HEPOF1SH, "                    ' �i����ʒu_��(OSF1E)
    sSQL = sSQL & "HEPOF1ST, "                    ' �i����ʒu_�_(OSF1E)
    sSQL = sSQL & "HEPOF1SR, "                    ' �i����ʒu_��(OSF1E)
    sSQL = sSQL & "HEPOF2NS, "                    ' �i�M�����@(OSF2E)
    sSQL = sSQL & "HEPOF2SZ, "                    ' �i�������(OSF2E)
    sSQL = sSQL & "HEPOF2ET, "                    ' �i�I��ET��(OSF2E)
    sSQL = sSQL & "HEPOF2HS, "                    ' �i�ۏؕ��@_��(OSF2E)
    sSQL = sSQL & "HEPOF2SH, "                    ' �i����ʒu_��(OSF2E)
    sSQL = sSQL & "HEPOF2ST, "                    ' �i����ʒu_�_(OSF2E)
    sSQL = sSQL & "HEPOF2SR, "                    ' �i����ʒu_��(OSF2E)
    sSQL = sSQL & "HEPOF3NS, "                    ' �i�M�����@(OSF3E)
    sSQL = sSQL & "HEPOF3SZ, "                    ' �i�������(OSF3E)
    sSQL = sSQL & "HEPOF3ET, "                    ' �i�I��ET��(OSF3E)
    sSQL = sSQL & "HEPOF3HS, "                    ' �i�ۏؕ��@_��(OSF3E)
    sSQL = sSQL & "HEPOF3SH, "                    ' �i����ʒu_��(OSF3E)
    sSQL = sSQL & "HEPOF3ST, "                    ' �i����ʒu_�_(OSF3E)
    sSQL = sSQL & "HEPOF3SR, "                    ' �i����ʒu_��(OSF3E)
    sSQL = sSQL & "HEPBM1NS, "                    ' �i�M�����@(BMD1E)
    sSQL = sSQL & "HEPBM1SZ, "                    ' �i�������(BMD1E)
    sSQL = sSQL & "HEPBM1ET, "                    ' �i�I��ET��(BMD1E)
    sSQL = sSQL & "HEPBM1HS, "                    ' �i�ۏؕ��@_��(BMD1E)
    sSQL = sSQL & "HEPBM1SH, "                    ' �i����ʒu_��(BMD1E)
    sSQL = sSQL & "HEPBM1ST, "                    ' �i����ʒu_�_(BMD1E)
    sSQL = sSQL & "HEPBM1SR, "                    ' �i����ʒu_��(BMD1E)
    sSQL = sSQL & "HEPBM2NS, "                    ' �i�M�����@(BMD2E)
    sSQL = sSQL & "HEPBM2SZ, "                    ' �i�������(BMD2E)
    sSQL = sSQL & "HEPBM2ET, "                    ' �i�I��ET��(BMD2E)
    sSQL = sSQL & "HEPBM2HS, "                    ' �i�ۏؕ��@_��(BMD1E)
    sSQL = sSQL & "HEPBM2SH, "                    ' �i����ʒu_��(BMD2E)
    sSQL = sSQL & "HEPBM2ST, "                    ' �i����ʒu_�_(BMD2E)
    sSQL = sSQL & "HEPBM2SR, "                    ' �i����ʒu_��(BMD2E)
    sSQL = sSQL & "HEPBM3NS, "                    ' �i�M�����@(BMD3E)
    sSQL = sSQL & "HEPBM3SZ, "                    ' �i�������(BMD3E)
    sSQL = sSQL & "HEPBM3ET, "                    ' �i�I��ET��(BMD3E)
    sSQL = sSQL & "HEPBM3HS, "                    ' �i�ۏؕ��@_��(BMD1E)
    sSQL = sSQL & "HEPBM3SH, "                    ' �i����ʒu_��(BMD3E)
    sSQL = sSQL & "HEPBM3ST, "                    ' �i����ʒu_�_(BMD3E)
    sSQL = sSQL & "HEPBM3SR, "                    ' �i����ʒu_��(BMD3E)
    sSQL = sSQL & "HEPACEN, "                     ' �iE1�����S
    sSQL = sSQL & "HEPANTNP, "                    ' �iEPAN���x
    sSQL = sSQL & "HEPANTIM, "                    ' �iEPAN����
    sSQL = sSQL & "HEPIGKBN, "                    ' �iEPIG�敪
    sSQL = sSQL & "HEPANGZY "                     ' �iEP����AN�K�X����
    sSQL = sSQL & "from TBCME050 "
    sSQL = sSQL & "where HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & pSpWFSamp.HIN.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & pSpWFSamp.HIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    If IsNull(rs("HEPOF1NS")) = False Then pSpWFSamp.HEPOF1NS = rs("HEPOF1NS") Else pSpWFSamp.HEPOF1NS = " "
    If IsNull(rs("HEPOF1SZ")) = False Then pSpWFSamp.HEPOF1SZ = rs("HEPOF1SZ") Else pSpWFSamp.HEPOF1SZ = " "
    pSpWFSamp.HEPOF1ET = fncNullCheck(rs("HEPOF1ET"))
    If IsNull(rs("HEPOF1HS")) = False Then pSpWFSamp.HEPOF1HS = rs("HEPOF1HS") Else pSpWFSamp.HEPOF1HS = " "
    If IsNull(rs("HEPOF1SH")) = False Then pSpWFSamp.HEPOF1SH = rs("HEPOF1SH") Else pSpWFSamp.HEPOF1SH = " "
    If IsNull(rs("HEPOF1ST")) = False Then pSpWFSamp.HEPOF1ST = rs("HEPOF1ST") Else pSpWFSamp.HEPOF1ST = " "
    If IsNull(rs("HEPOF1SR")) = False Then pSpWFSamp.HEPOF1SR = rs("HEPOF1SR") Else pSpWFSamp.HEPOF1SR = " "
    If IsNull(rs("HEPOF2NS")) = False Then pSpWFSamp.HEPOF2NS = rs("HEPOF2NS") Else pSpWFSamp.HEPOF2NS = " "
    If IsNull(rs("HEPOF2SZ")) = False Then pSpWFSamp.HEPOF2SZ = rs("HEPOF2SZ") Else pSpWFSamp.HEPOF2SZ = " "
    pSpWFSamp.HEPOF2ET = fncNullCheck(rs("HEPOF2ET"))
    If IsNull(rs("HEPOF2HS")) = False Then pSpWFSamp.HEPOF2HS = rs("HEPOF2HS") Else pSpWFSamp.HEPOF2HS = " "
    If IsNull(rs("HEPOF2SH")) = False Then pSpWFSamp.HEPOF2SH = rs("HEPOF2SH") Else pSpWFSamp.HEPOF2SH = " "
    If IsNull(rs("HEPOF2ST")) = False Then pSpWFSamp.HEPOF2ST = rs("HEPOF2ST") Else pSpWFSamp.HEPOF2ST = " "
    If IsNull(rs("HEPOF2SR")) = False Then pSpWFSamp.HEPOF2SR = rs("HEPOF2SR") Else pSpWFSamp.HEPOF2SR = " "
    If IsNull(rs("HEPOF3NS")) = False Then pSpWFSamp.HEPOF3NS = rs("HEPOF3NS") Else pSpWFSamp.HEPOF3NS = " "
    If IsNull(rs("HEPOF3SZ")) = False Then pSpWFSamp.HEPOF3SZ = rs("HEPOF3SZ") Else pSpWFSamp.HEPOF3SZ = " "
    pSpWFSamp.HEPOF3ET = fncNullCheck(rs("HEPOF3ET"))
    If IsNull(rs("HEPOF3HS")) = False Then pSpWFSamp.HEPOF3HS = rs("HEPOF3HS") Else pSpWFSamp.HEPOF3HS = " "
    If IsNull(rs("HEPOF3SH")) = False Then pSpWFSamp.HEPOF3SH = rs("HEPOF3SH") Else pSpWFSamp.HEPOF3SH = " "
    If IsNull(rs("HEPOF3ST")) = False Then pSpWFSamp.HEPOF3ST = rs("HEPOF3ST") Else pSpWFSamp.HEPOF3ST = " "
    If IsNull(rs("HEPOF3SR")) = False Then pSpWFSamp.HEPOF3SR = rs("HEPOF3SR") Else pSpWFSamp.HEPOF3SR = " "
    If IsNull(rs("HEPBM1NS")) = False Then pSpWFSamp.HEPBM1NS = rs("HEPBM1NS") Else pSpWFSamp.HEPBM1NS = " "
    If IsNull(rs("HEPBM1SZ")) = False Then pSpWFSamp.HEPBM1SZ = rs("HEPBM1SZ") Else pSpWFSamp.HEPBM1SZ = " "
    pSpWFSamp.HEPBM1ET = fncNullCheck(rs("HEPBM1ET"))
    If IsNull(rs("HEPBM1HS")) = False Then pSpWFSamp.HEPBM1HS = rs("HEPBM1HS") Else pSpWFSamp.HEPBM1HS = " "
    If IsNull(rs("HEPBM1SH")) = False Then pSpWFSamp.HEPBM1SH = rs("HEPBM1SH") Else pSpWFSamp.HEPBM1SH = " "
    If IsNull(rs("HEPBM1ST")) = False Then pSpWFSamp.HEPBM1ST = rs("HEPBM1ST") Else pSpWFSamp.HEPBM1ST = " "
    If IsNull(rs("HEPBM1SR")) = False Then pSpWFSamp.HEPBM1SR = rs("HEPBM1SR") Else pSpWFSamp.HEPBM1SR = " "
    If IsNull(rs("HEPBM2NS")) = False Then pSpWFSamp.HEPBM2NS = rs("HEPBM2NS") Else pSpWFSamp.HEPBM2NS = " "
    If IsNull(rs("HEPBM2SZ")) = False Then pSpWFSamp.HEPBM2SZ = rs("HEPBM2SZ") Else pSpWFSamp.HEPBM2SZ = " "
    pSpWFSamp.HEPBM2ET = fncNullCheck(rs("HEPBM2ET"))
    If IsNull(rs("HEPBM2HS")) = False Then pSpWFSamp.HEPBM2HS = rs("HEPBM2HS") Else pSpWFSamp.HEPBM2HS = " "
    If IsNull(rs("HEPBM2SH")) = False Then pSpWFSamp.HEPBM2SH = rs("HEPBM2SH") Else pSpWFSamp.HEPBM2SH = " "
    If IsNull(rs("HEPBM2ST")) = False Then pSpWFSamp.HEPBM2ST = rs("HEPBM2ST") Else pSpWFSamp.HEPBM2ST = " "
    If IsNull(rs("HEPBM2SR")) = False Then pSpWFSamp.HEPBM2SR = rs("HEPBM2SR") Else pSpWFSamp.HEPBM2SR = " "
    If IsNull(rs("HEPBM3NS")) = False Then pSpWFSamp.HEPBM3NS = rs("HEPBM3NS") Else pSpWFSamp.HEPBM3NS = " "
    If IsNull(rs("HEPBM3SZ")) = False Then pSpWFSamp.HEPBM3SZ = rs("HEPBM3SZ") Else pSpWFSamp.HEPBM3SZ = " "
    pSpWFSamp.HEPBM3ET = fncNullCheck(rs("HEPBM3ET"))
    If IsNull(rs("HEPBM3HS")) = False Then pSpWFSamp.HEPBM3HS = rs("HEPBM3HS") Else pSpWFSamp.HEPBM3HS = " "
    If IsNull(rs("HEPBM3SH")) = False Then pSpWFSamp.HEPBM3SH = rs("HEPBM3SH") Else pSpWFSamp.HEPBM3SH = " "
    If IsNull(rs("HEPBM3ST")) = False Then pSpWFSamp.HEPBM3ST = rs("HEPBM3ST") Else pSpWFSamp.HEPBM3ST = " "
    If IsNull(rs("HEPBM3SR")) = False Then pSpWFSamp.HEPBM3SR = rs("HEPBM3SR") Else pSpWFSamp.HEPBM3SR = " "
    pSpWFSamp.HEPACEN = fncNullCheck(rs("HEPACEN"))
    pSpWFSamp.HEPANTNP = fncNullCheck(rs("HEPANTNP"))
    pSpWFSamp.HEPANTIM = fncNullCheck(rs("HEPANTIM"))
    If IsNull(rs("HEPIGKBN")) = False Then pSpWFSamp.HEPIGKBN = rs("HEPIGKBN") Else pSpWFSamp.HEPIGKBN = " "
    If IsNull(rs("HEPANGZY")) = False Then pSpWFSamp.HEPANGZY = rs("HEPANGZY") Else pSpWFSamp.HEPANGZY = " "
    rs.Close
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    scmzc_getWF = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    scmzc_getWF = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'***********************************************************************************
'*    �֐���        : MakeParameter
'*
'*    �����T�v      : 1.�VDB�����ݏ���
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    intFormID     ,I  ,Integer  ,�i1:WF�Z���^��������@2:�Ĕ���
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************
Public Function MakeParameter(ByVal intFormID As Integer) As FUNCTION_RETURN
    Dim sErrTbl             As String
    Dim vBeforeBlock        As Variant  '���ݍs�̃u���b�NID
    Dim vAfterBlock         As Variant  '���̍s�̃u���b�NID
    Dim intSprCnt           As Integer  '�X�v���b�h���[�v�J�E���g
    Dim sErrMsg             As String
    Dim vIngotpos           As Variant
    Dim lngBeginIngotpos    As Long
    Dim lngEndIngotpos      As Long
    Dim vBeginSeq           As Variant
    Dim lngWfBeginSeq       As Long
    Dim lngWfEndSeq         As Long
    Dim sCryNum             As String
    Dim sSXLID              As String

    If intFormID = 1 Then 'WF�Z���^�[����������s
        '�\���̍쐬
        If cmbc039_2_CreateTable(sErrMsg) = FUNCTION_RETURN_FAILURE Then
            MakeParameter = FUNCTION_RETURN_FAILURE
            f_cmbc039_2.lblMsg.Caption = sErrMsg
            Exit Function
        End If
    ElseIf intFormID = 2 Then '�Ĕ����w�����s
        sSXLID = Trim(f_cmbc039_3.txtKSXLID.text)
        '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX-------start iida 2003/09/06
        With f_cmbc039_3.sprExamine
            lngBeginIngotpos = SIngotP  '2003/04/22 okazaki
            lngEndIngotpos = EIngotP  '2003/05/01 hitec)matsumoto
            .GetText 6, 1, vBeginSeq
            '�����E�V�K�u���b�N�ʒu�̏C�� 2003/04/22
            lngWfBeginSeq = CInt(Trim(vBeginSeq))
            .GetText 6, .MaxRows, vBeginSeq
            lngWfEndSeq = CInt(Trim(vBeginSeq))
        End With

        '�\���̍쐬
        intSprCnt = 0
        '�e�[�u���W�J����
        If cmbc039_3_CreateTable(sSXLID, lngBeginIngotpos, lngEndIngotpos, lngWfBeginSeq, lngWfEndSeq, sErrMsg) = FUNCTION_RETURN_FAILURE Then 'upd 2003/03/29 hitec)matsumoto lngWfBeginSeq,lngWfEndSeq�ǉ�
            MakeParameter = FUNCTION_RETURN_FAILURE
            f_cmbc039_3.lblMsg.Caption = sErrMsg
            Exit Function
        End If
    End If
    MakeParameter = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

End Function

'***********************************************************************************************
'*    �֐���        : cmbc039_2_CreateXSDC2
'*
'*    �����T�v      : 1.���������i�u���b�N�j�O�H�����ю擾���\���̍쐬
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    intBlockCnt   ,I  ,Integer  ,�u���b�N��
'*                    bNoData�@�@   ,I  ,Boolean  ,�f�[�^�L���t���O
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function cmbc039_2_CreateXSDC2(ByVal intBlockCnt As Integer, ByRef bNoData As Boolean) _
                                        As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sSQL        As String
    Dim intProcNo   As Integer
    Dim dblDiameter As Double
    Dim intNum      As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False

    '�u���b�NID�𓾂�
    sSQL = "SELECT * from XSDC2 "
    sSQL = sSQL & " WHERE CRYNUMC2='" & strBlockID(intBlockCnt) & "'"
    sSQL = sSQL & "   AND LIVKC2= '0'"   '�����敪

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc039_2_CreateXSDC2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    rs.MoveFirst
    If rs.EOF = False Then
        With BlkOld
            If IsNull(rs.Fields("CRYNUMC2")) = False Then .CRYNUMC2 = rs.Fields("CRYNUMC2")
            If IsNull(rs.Fields("KCNTC2")) = False Then .KCNTC2 = rs.Fields("KCNTC2")       '�H���A��
            If IsNull(rs.Fields("XTALC2")) = False Then .XTALC2 = rs.Fields("XTALC2")
            If IsNull(rs.Fields("INPOSC2")) = False Then .INPOSC2 = rs.Fields("INPOSC2")
            If IsNull(rs.Fields("NEKKNTC2")) = False Then .NEKKNTC2 = rs.Fields("NEKKNTC2")
            If IsNull(rs.Fields("NEWKNTC2")) = False Then .NEWKNTC2 = rs.Fields("NEWKNTC2")
            If IsNull(rs.Fields("NEWKKBC2")) = False Then .NEWKKBC2 = rs.Fields("NEWKKBC2")
            If IsNull(rs.Fields("NEMACOC2")) = False Then .NEMACOC2 = rs.Fields("NEMACOC2")
            If IsNull(rs.Fields("GNKKNTC2")) = False Then .GNKKNTC2 = rs.Fields("GNKKNTC2")
            If IsNull(rs.Fields("GNWKNTC2")) = False Then .GNWKNTC2 = rs.Fields("GNWKNTC2")
            If IsNull(rs.Fields("GNWKKBC2")) = False Then .GNWKKBC2 = rs.Fields("GNWKKBC2")
            If IsNull(rs.Fields("GNMACOC2")) = False Then .GNMACOC2 = rs.Fields("GNMACOC2")
            If IsNull(rs.Fields("GNDAYC2")) = False Then .GNDAYC2 = rs.Fields("GNDAYC2")
            If IsNull(rs.Fields("GNLC2")) = False Then .GNLC2 = rs.Fields("GNLC2")          '���ݒ����i�O�H�������j
            If IsNull(rs.Fields("GNWC2")) = False Then .GNWC2 = rs.Fields("GNWC2")          '���ݏd��
            If IsNull(rs.Fields("GNMC2")) = False Then .GNMC2 = rs.Fields("GNMC2")          '���ݖ����i�O�H�������j
            If IsNull(rs.Fields("SUMITLC2")) = False Then .SUMITLC2 = rs.Fields("SUMITLC2")
            If IsNull(rs.Fields("SUMITWC2")) = False Then .SUMITWC2 = rs.Fields("SUMITWC2")
            If IsNull(rs.Fields("SUMITMC2")) = False Then .SUMITMC2 = rs.Fields("SUMITMC2")
            If IsNull(rs.Fields("CHGC2")) = False Then .CHGC2 = rs.Fields("CHGC2")
            If IsNull(rs.Fields("KAKOUBC2")) = False Then .KAKOUBC2 = rs.Fields("KAKOUBC2")
            If IsNull(rs.Fields("KEIDAYC2")) = False Then .KEIDAYC2 = rs.Fields("KEIDAYC2")
            If IsNull(rs.Fields("GNTKUBC2")) = False Then .GNTKUBC2 = rs.Fields("GNTKUBC2")
            If IsNull(rs.Fields("GNTNOC2")) = False Then .GNTNOC2 = rs.Fields("GNTNOC2")
            If IsNull(rs.Fields("XTWORKC2")) = False Then .XTWORKC2 = rs.Fields("XTWORKC2")
            If IsNull(rs.Fields("WFWORKC2")) = False Then .WFWORKC2 = rs.Fields("WFWORKC2")
            If IsNull(rs.Fields("LSTATBC2")) = False Then .LSTATBC2 = rs.Fields("LSTATBC2")
            If IsNull(rs.Fields("RSTATBC2")) = False Then .RSTATBC2 = rs.Fields("RSTATBC2")
            If IsNull(rs.Fields("LUFRCC2")) = False Then .LUFRCC2 = rs.Fields("LUFRCC2")
            If IsNull(rs.Fields("LUFRBC2")) = False Then .LUFRBC2 = rs.Fields("LUFRBC2")
            If IsNull(rs.Fields("LDFRCC2")) = False Then .LDFRCC2 = rs.Fields("LDFRCC2")
            If IsNull(rs.Fields("LDFRBC2")) = False Then .LDFRBC2 = rs.Fields("LDFRBC2")
            If IsNull(rs.Fields("HOLDCC2")) = False Then .HOLDCC2 = rs.Fields("HOLDCC2")
            If IsNull(rs.Fields("HOLDBC2")) = False Then .HOLDBC2 = rs.Fields("HOLDBC2")
            If IsNull(rs.Fields("EXKUBC2")) = False Then .EXKUBC2 = rs.Fields("EXKUBC2")
            If IsNull(rs.Fields("HENPKC2")) = False Then .HENPKC2 = rs.Fields("HENPKC2")
            If IsNull(rs.Fields("LIVKC2")) = False Then .LIVKC2 = rs.Fields("LIVKC2")
            If IsNull(rs.Fields("KANKC2")) = False Then .KANKC2 = rs.Fields("KANKC2")
            If IsNull(rs.Fields("NFC2")) = False Then .NFC2 = rs.Fields("NFC2")
            If IsNull(rs.Fields("SAKJC2")) = False Then .SAKJC2 = rs.Fields("SAKJC2")
            If IsNull(rs.Fields("TDAYC2")) = False Then .TDAYC2 = rs.Fields("TDAYC2")
            If IsNull(rs.Fields("KDAYC2")) = False Then .KDAYC2 = rs.Fields("KDAYC2")
            If IsNull(rs.Fields("SUMITBC2")) = False Then .SUMITBC2 = rs.Fields("SUMITBC2")
            If IsNull(rs.Fields("SNDKC2")) = False Then .SNDKC2 = rs.Fields("SNDKC2")
            If IsNull(rs.Fields("SNDDAYC2")) = False Then .SNDDAYC2 = rs.Fields("SNDDAYC2")
            If IsNull(rs.Fields("PLANTCATC2")) = False Then .PLANTCATC2 = rs.Fields("PLANTCATC2")   ' 2007/09/04 SPK Tsutsumi Add
        End With

        '�O�H���̍\���̂����ݍH���̍\���̂ɃR�s�[
        BlkNow = BlkOld

        '���ݍH���̍H���A�Ԃ��C��
        With BlkNow
            .KCNTC2 = CInt(.KCNTC2) + 1     '�H���A��
            'Cng Start  2010/09/02 Y.Hitomi
            '�u���b�N��SXL��1�ł��������Ă����ꍇ�A�H���R�[�h���X�V���Ȃ��悤�ɂ���B
            If (.GNWKNTC2 <> "     " Or _
                .GNWKNTC2 <> "CW800" Or _
                .GNWKNTC2 <> "TX860") Then
            
                .NEWKNTC2 = Kihon.NOWPROC       '�O�H���R�[�h���ŏI�ʉߍH���ɃZ�b�g
                .GNWKNTC2 = Kihon.NEWPROC       '���ݍH���R�[�h�����ݍH���փZ�b�g
            End If
            '  .NEWKNTC2 = Kihon.NOWPROC       '�O�H���R�[�h���ŏI�ʉߍH���ɃZ�b�g
            '  .GNWKNTC2 = Kihon.NEWPROC       '���ݍH���R�[�h�����ݍH���փZ�b�g
            'Cng End    2010/09/02 Y.Hitomi
            
            .SUMITBC2 = "0"
            .SUMITLC2 = "0"
            .SUMITMC2 = "0"
            .SUMITWC2 = "0"
            '���ݏd�ʂ����߂�
            If GetDiameter(strBlockID(intBlockCnt), dblDiameter) = FUNCTION_RETURN_FAILURE Then  '���a�����߂�
                dblDiameter = 0
            End If
            '��{���̒��a���Z�b�g
            Kihon.DIAMETER = dblDiameter
        End With
    End If

    rs.Close
    cmbc039_2_CreateXSDC2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    cmbc039_2_CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'********************************************************************************************
'*    �֐���        : cmbc039_2_CreateXSDCA
'*
'*    �����T�v      : 1.���������i�i�ԁj�O�H�����ю擾���\���̍쐬
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                            ,����
'*                    intBlockCnt   ,I  ,Integer  ,�u���b�N��
'*                    bNoData�@�@   ,I  ,Boolean  ,�f�[�^�L���t���O
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'********************************************************************************************
'Cng Start 2010/10/03 Y.Hitomi
Public Function cmbc039_2_CreateXSDCA(ByVal intBlockCnt As Integer, ByRef bNoData As Boolean, ByVal strSXLID As String) _
                                        As FUNCTION_RETURN
'Public Function cmbc039_2_CreateXSDCA(ByVal intBlockCnt As Integer, ByRef bNoData As Boolean) _
                                        As FUNCTION_RETURN
'Cng End 2010/10/03 Y.Hitomi
    
    Dim intLoopCnt  As Integer
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim intProcNo   As Integer
    Dim dblDiameter As Double
    Dim intNum      As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False

    '�u���b�NID�𓾂�
    sql = "SELECT * from XSDCA"
    sql = sql & " WHERE CRYNUMCA='" & strBlockID(intBlockCnt) & "'"
    sql = sql & "   AND LIVKCA= '0'"   '�����敪

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc039_2_CreateXSDCA = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    rs.MoveFirst
    intLoopCnt = 0

    Do While Not rs.EOF
        ReDim Preserve HinOld(intLoopCnt)
        ReDim Preserve HinNow(intLoopCnt)
        Kihon.CNTHINOLD = intLoopCnt + 1
        Kihon.CNTHINNOW = intLoopCnt + 1
        With HinOld(intLoopCnt)
            If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
            If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
            If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
            If IsNull(rs.Fields("REVNUMCA")) = False Then .REVNUMCA = rs.Fields("REVNUMCA")
            If IsNull(rs.Fields("FACTORYCA")) = False Then .FACTORYCA = rs.Fields("FACTORYCA")
            If IsNull(rs.Fields("OPECA")) = False Then .OPECA = rs.Fields("OPECA")
            If IsNull(rs.Fields("KCKNTCA")) = False Then .KCKNTCA = rs.Fields("KCKNTCA")
            If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLIDCA = rs.Fields("SXLIDCA")
            If IsNull(rs.Fields("XTALCA")) = False Then .XTALCA = rs.Fields("XTALCA")
            If IsNull(rs.Fields("NEKKNTCA")) = False Then .NEKKNTCA = rs.Fields("NEKKNTCA")
            If IsNull(rs.Fields("NEWKNTCA")) = False Then .NEWKNTCA = rs.Fields("NEWKNTCA")
            If IsNull(rs.Fields("NEWKKBCA")) = False Then .NEWKKBCA = rs.Fields("NEWKKBCA")
            If IsNull(rs.Fields("NEMACOCA")) = False Then .NEMACOCA = rs.Fields("NEMACOCA")
            If IsNull(rs.Fields("GNKKNTCA")) = False Then .GNKKNTCA = rs.Fields("GNKKNTCA")
            If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
            If IsNull(rs.Fields("GNWKKBCA")) = False Then .GNWKKBCA = rs.Fields("GNWKKBCA")
            If IsNull(rs.Fields("GNMACOCA")) = False Then .GNMACOCA = rs.Fields("GNMACOCA")
            If IsNull(rs.Fields("GNDAYCA")) = False Then .GNDAYCA = rs.Fields("GNDAYCA")
            If IsNull(rs.Fields("GNLCA")) = False Then .GNLCA = rs.Fields("GNLCA")
            If IsNull(rs.Fields("GNWCA")) = False Then .GNWCA = rs.Fields("GNWCA")
            If IsNull(rs.Fields("GNMCA")) = False Then .GNMCA = rs.Fields("GNMCA")
            If IsNull(rs.Fields("SUMITLCA")) = False Then .SUMITLCA = rs.Fields("SUMITLCA")
            If IsNull(rs.Fields("SUMITWCA")) = False Then .SUMITWCA = rs.Fields("SUMITWCA")
            If IsNull(rs.Fields("SUMITMCA")) = False Then .SUMITMCA = rs.Fields("SUMITMCA")
            If IsNull(rs.Fields("CHGCA")) = False Then .CHGCA = rs.Fields("CHGCA")
            If IsNull(rs.Fields("KAKOUBCA")) = False Then .KAKOUBCA = rs.Fields("KAKOUBCA")
            If IsNull(rs.Fields("KEIDAYCA")) = False Then .KEIDAYCA = rs.Fields("KEIDAYCA")
            If IsNull(rs.Fields("GNTKUBCA")) = False Then .GNTKUBCA = rs.Fields("GNTKUBCA")
            If IsNull(rs.Fields("GNTNOCA")) = False Then .GNTNOCA = rs.Fields("GNTNOCA")
            If IsNull(rs.Fields("XTWORKCA")) = False Then .XTWORKCA = rs.Fields("XTWORKCA")
            If IsNull(rs.Fields("WFWORKCA")) = False Then .WFWORKCA = rs.Fields("WFWORKCA")
            If IsNull(rs.Fields("LSTATBCA")) = False Then .LSTATBCA = rs.Fields("LSTATBCA")
            If IsNull(rs.Fields("RSTATBCA")) = False Then .RSTATBCA = rs.Fields("RSTATBCA")
            If IsNull(rs.Fields("LUFRCCA")) = False Then .LUFRCCA = rs.Fields("LUFRCCA")
            If IsNull(rs.Fields("LUFRBCA")) = False Then .LUFRBCA = rs.Fields("LUFRBCA")
            If IsNull(rs.Fields("LDFRCCA")) = False Then .LDFRCCA = rs.Fields("LDFRCCA")
            If IsNull(rs.Fields("LDFRBCA")) = False Then .LDFRBCA = rs.Fields("LDFRBCA")
            If IsNull(rs.Fields("HOLDCCA")) = False Then .HOLDCCA = rs.Fields("HOLDCCA")
            If IsNull(rs.Fields("HOLDBCA")) = False Then .HOLDBCA = rs.Fields("HOLDBCA")
            If IsNull(rs.Fields("EXKUBCA")) = False Then .EXKUBCA = rs.Fields("EXKUBCA")
            If IsNull(rs.Fields("HENPKCA")) = False Then .HENPKCA = rs.Fields("HENPKCA")
            If IsNull(rs.Fields("LIVKCA")) = False Then .LIVKCA = rs.Fields("LIVKCA")
            If IsNull(rs.Fields("KANKCA")) = False Then .KANKCA = rs.Fields("KANKCA")
            If IsNull(rs.Fields("NFCA")) = False Then .NFCA = rs.Fields("NFCA")
            If IsNull(rs.Fields("SAKJCA")) = False Then .SAKJCA = rs.Fields("SAKJCA")
            If IsNull(rs.Fields("TDAYCA")) = False Then .TDAYCA = rs.Fields("TDAYCA")
            If IsNull(rs.Fields("KDAYCA")) = False Then .KDAYCA = rs.Fields("KDAYCA")
            If IsNull(rs.Fields("SUMITBCA")) = False Then .SUMITBCA = rs.Fields("SUMITBCA")
            If IsNull(rs.Fields("SNDKCA")) = False Then .SNDKCA = rs.Fields("SNDKCA")
            If IsNull(rs.Fields("SNDDAYCA")) = False Then .SNDDAYCA = rs.Fields("SNDDAYCA")
            If IsNull(rs.Fields("PLANTCATCA")) = False Then .PLANTCATCA = rs.Fields("PLANTCATCA")   '2007/09/04 SPK Tsutsumi Add
        End With

        '�O�H���̍\���̂����ݍH���̍\���̂փR�s�[
        HinNow(intLoopCnt) = HinOld(intLoopCnt)

        With HinNow(intLoopCnt)
            .KCKNTCA = CInt(.KCKNTCA) + 1
            'Cng Start 2010/10/03 Y.Hitomi
            '���s�w��SXLID�̂ݍH���R�[�h��ύX���A����ȊO�́A�O�H���������p��
            If strSXLID = .SXLIDCA Then
                .NEWKNTCA = Kihon.NOWPROC             '�O�H���R�[�h���ŏI�ʉߍH���ɃZ�b�g
                .GNWKNTCA = Kihon.NEWPROC             '���ݍH���R�[�h�����ݍH���փZ�b�g
            Else
                .NEWKNTCA = rs.Fields("NEWKNTCA")     '�O�H���R�[�h���ŏI�ʉߍH���ɃZ�b�g
                .GNWKNTCA = rs.Fields("GNWKNTCA")     '���ݍH���R�[�h�����ݍH���փZ�b�g
            End If
            '.NEWKNTCA = Kihon.NOWPROC             '�O�H���R�[�h���ŏI�ʉߍH���ɃZ�b�g
            '.GNWKNTCA = Kihon.NEWPROC             '���ݍH���R�[�h�����ݍH���փZ�b�g
            'Cng End   2010/10/03 Y.Hitomi
            .SUMITBCA = "0"
            .SUMITLCA = HinOld(intLoopCnt).SUMITLCA   ''03/05/13 �㓡
            .SUMITMCA = HinOld(intLoopCnt).SUMITMCA   ''03/05/13 �㓡
            .SUMITWCA = HinOld(intLoopCnt).SUMITWCA   ''03/05/13 �㓡
            '���ݏd�ʂ����߂�
            If GetDiameter(strBlockID(intBlockCnt), dblDiameter) = FUNCTION_RETURN_FAILURE Then  '���a�����߂�
                dblDiameter = 0
            End If
        End With

        intLoopCnt = intLoopCnt + 1
        rs.MoveNext
    Loop

    rs.Close
    cmbc039_2_CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc039_2_CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'*************************************************************************************************************
'*    �֐���        : cmbc039_3_CreateTable
'*
'*    �����T�v      : 1.�\���̍쐬����
'*
'*    �p�����[�^    : �ϐ���           ,IO ,�^      ,����
'*                    strSXLID         ,I  ,String  ,SXL-ID
'*                    lngBeginIngotpos ,I  ,Long    ,�u���b�N�Ǘ��f�[�^�̒���
'*                    strErrMsg        ,O  ,String  ,ErrMsg�i�[
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*************************************************************************************************************
Public Function cmbc039_3_CreateTable(ByVal strSXLID As String, ByVal lngBeginIngotpos As Long, _
                                      ByVal lngEndIngotpos As Long, ByVal lngWfBeginSeq As Long, _
                                      ByVal lngWfEndSeq As Long, ByRef strErrMsg As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sErrTbl     As String
    Dim intSprCnt   As Integer
    Dim intRowCnt   As Integer
    Dim sDBName     As String
    Dim blNoData    As Boolean
    Dim sSQL        As String
    Dim intLoopCnt  As Integer

    blNoData = False

    giInpos = 9000  '�݌Ɍ��A�U�֏��̈ʒu��������

    '�u���b�N�Ǘ�����u���b�N�h�c���擾
    sSQL = "SELECT DISTINCT(CRYNUMCA) "
    sSQL = sSQL & " FROM XSDCA"
    sSQL = sSQL & " WHERE CRYNUMCA like '" & left(strSXLID, 9) & "%'"   '���ޯ�����ڒǉ� 09/05/25 ooba
    sSQL = sSQL & "   AND SXLIDCA = '" & strSXLID & "'"
    sSQL = sSQL & "   AND LIVKCA = '0'"   'add 2003/05/19 hitec)matsumoto

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    '�u���b�NID���擾
    intLoopCnt = 0
    Do While Not rs.EOF
        ReDim Preserve strBlockID(intLoopCnt) As String
        If IsNull(rs("CRYNUMCA")) = True Then
            strBlockID(intLoopCnt) = ""
        Else
            strBlockID(intLoopCnt) = rs("CRYNUMCA")            '�u���b�NID
        End If
        '��{���\����
        With Kihon
            .STAFFID = Trim(f_cmbc039_3.txtStaffID.text)
            .NEWPROC = PROCD_WFC_SOUGOUHANTEI
            .NOWPROC = PROCD_WFC_SAINUKISI
            .DIAMETER = 0
        End With

        '���������i�u���b�N�j����O�H�����ю擾
        sDBName = "XSDC2"
        If cmbc039_3_CreateXSDC2(strBlockID(intLoopCnt), blNoData) = FUNCTION_RETURN_FAILURE Then
            If blNoData = True Then
                cmbc039_3_CreateTable = FUNCTION_RETURN_SUCCESS
                GoTo proc_exit
            Else
                cmbc039_3_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EGET") & sDBName
                GoTo proc_exit
            End If
        End If

        '���������i�i�ԁj����O�H�����ю擾
        sDBName = "XSDCA"
        If cmbc039_3_CreateXSDCA(strBlockID(intLoopCnt), blNoData) = FUNCTION_RETURN_FAILURE Then
            If blNoData = True Then
                cmbc039_3_CreateTable = FUNCTION_RETURN_SUCCESS
                GoTo proc_exit
            Else
                cmbc039_3_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EGET") & sDBName
                GoTo proc_exit
            End If
        End If

        '���ݍH�����э쐬
        sDBName = "XSDC2,XSDCA"
        strErrMsg = GetMsgStr("EAPLY") & sDBName
        
        'Cng Start 2010/10/03 Y.Hitomi
        'If cmbc039_3_CreateNowProc(strBlockID(intLoopCnt), lngBeginIngotpos, lngEndIngotpos, lngWfBeginSeq, lngWfEndSeq, strErrMsg) = FUNCTION_RETURN_FAILURE Then   'upd 2003/03/29 hitec)matsumoto �������ʒu���g�p���Ă������A�}�b�v�ʒu�ɕύX
        If cmbc039_3_CreateNowProc(strBlockID(intLoopCnt), lngBeginIngotpos, lngEndIngotpos, lngWfBeginSeq, lngWfEndSeq, strErrMsg, strSXLID) = FUNCTION_RETURN_FAILURE Then  'upd 2003/03/29 hitec)matsumoto �������ʒu���g�p���Ă������A�}�b�v�ʒu�ɕύX
        'Cng End  2010/10/03 Y.Hitomi
            cmbc039_3_CreateTable = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        strErrMsg = ""

        '��{����
        If KihonProc = FUNCTION_RETURN_FAILURE Then
            cmbc039_3_CreateTable = FUNCTION_RETURN_FAILURE
            strErrMsg = GetMsgStr("EAPLY")
            GoTo proc_exit
        End If

        rs.MoveNext
        intLoopCnt = intLoopCnt + 1
    Loop

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

End Function

'**********************************************************************************
'*    �֐���        : cmbc039_2_CreateTable
'*
'*    �����T�v      : 1.�\���̍쐬����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^      ,����
'*                    strErrMsg     ,O  ,String  ,ErrMsg�i�[
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'**********************************************************************************
Public Function cmbc039_2_CreateTable(ByRef strErrMsg As String) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rsMain      As OraDynaset
    Dim sErrTbl     As String
    Dim intBlockCnt As Integer
    Dim sDBName     As String
    Dim blNoData    As Boolean
    Dim sTmpSxl()   As String       '�d�|�H���������pSXLID�@06/03/14 ooba
    Dim blKouteiChk As Boolean      '�H�������׸ށ@06/03/14 ooba

    On Error GoTo proc_err

    blNoData = False
    '�u���b�NID�擾
    sDBName = "XSDCA"
    '���ޯ������(CRYNUMCA)�ǉ� 09/05/25 ooba
    sSQL = "select DISTINCT(CRYNUMCA) from XSDCA " & _
          "where CRYNUMCA like '" & left(SelectSxlID039, 9) & "%' " & _
          "  and SXLIDCA='" & SelectSxlID039 & "' " & _
          "  and LIVKCA= '0'"
    Set rsMain = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rsMain.RecordCount = 0 Then
        Debug.Print "XSDCA�F�O�H�����тȂ�"
        Debug.Print sSQL
        rsMain.Close
        cmbc039_2_CreateTable = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If

    intBlockCnt = 0
    blNoData = False
    blKouteiChk = False      '06/03/14 ooba

    Do While Not rsMain.EOF
        intBlockCnt = intBlockCnt + 1
        ReDim Preserve strBlockID(intBlockCnt)
        strBlockID(intBlockCnt) = rsMain("CRYNUMCA")

        With Kihon
            .STAFFID = Trim(f_cmbc039_2.txtStaffID.text)
            .NOWPROC = PROCD_WFC_SOUGOUHANTEI
            .NEWPROC = PROCD_SXL_KAKUTEI
            .DIAMETER = 0       '--------------�ۗ�
            .ALLSCRAP = "N" '�S���X�N���b�v
            .FURYOUMU = "N"   '�s�ǖ���
        End With

        '���������i�u���b�N�j����O�H�����ю擾
        sDBName = "XSDC2"
        If cmbc039_2_CreateXSDC2(intBlockCnt, blNoData) = FUNCTION_RETURN_FAILURE Then
            If blNoData = True Then
                rsMain.Close
                Debug.Print "cmbc039_2_CreateXSDC2(" & intBlockCnt & "," & blNoData & "):XSDC2�O�H�����тȂ�"
                cmbc039_2_CreateTable = FUNCTION_RETURN_SUCCESS
                Exit Function
            Else
                rsMain.Close
                Debug.Print "cmbc039_2_CreateXSDC2(" & intBlockCnt & "," & blNoData & "):XSDC2�O�H�����ѓǍ��݃G���["
                cmbc039_2_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EGET") & sDBName
                Exit Function
            End If
        End If

        sDBName = "XSDCA"
        '���������i�i�ԁj����O�H�����ю擾
'2010/10/03 Cng Start Y.Hitomi
'        If cmbc039_2_CreateXSDCA(intBlockCnt, blNoData) = FUNCTION_RETURN_FAILURE Then
        If cmbc039_2_CreateXSDCA(intBlockCnt, blNoData, SelectSxlID039) = FUNCTION_RETURN_FAILURE Then
'2010/10/03 Cng End Y.Hitomi
            If blNoData = True Then
                rsMain.Close
                Debug.Print "cmbc039_2_CreateXSDCA(" & intBlockCnt & "," & blNoData & "):XSDCA�O�H�����тȂ�"
                cmbc039_2_CreateTable = FUNCTION_RETURN_SUCCESS
                Exit Function
            Else
                rsMain.Close
                Debug.Print "cmbc039_2_CreateXSDCA(" & intBlockCnt & "," & blNoData & "):XSDCA�O�H�����ѓǍ��݃G���["
                cmbc039_2_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EGET") & sDBName
                Exit Function
            End If
        End If

        '�d�|�H���ă`�F�b�N�@�\�ǉ��@06/03/14 ooba
        ReDim sTmpSxl(1)
        sTmpSxl(1) = SelectSxlID039
        If Not blKouteiChk Then
            If DBDRV_CheckCodeXSDCB(sTmpSxl, Kihon.NOWPROC, strErrMsg) = FUNCTION_RETURN_FAILURE Then
                cmbc039_2_CreateTable = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            blKouteiChk = True
        End If

        '��{����
        If KihonProc = FUNCTION_RETURN_FAILURE Then
            rsMain.Close
            Debug.Print "KihonProc()�F��{�����ُ�I��"
            cmbc039_2_CreateTable = FUNCTION_RETURN_FAILURE
            strErrMsg = GetMsgStr("EAPLY")
            Exit Function
        End If

        rsMain.MoveNext
    Loop
    rsMain.Close
    cmbc039_2_CreateTable = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    cmbc039_2_CreateTable = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'*******************************************************************************************
'*    �֐���        : cmbc039_3_CreateXSDCA
'*
'*    �����T�v      : 1.���������i�i�ԁj�O�H�����ю擾���\���̍쐬
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^      ,����
'*                    strBlockID    ,O  ,String  ,�u���b�NID
'*                    bNoData�@�@   ,I  ,Boolean ,�f�[�^�L���t���O
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function cmbc039_3_CreateXSDCA(ByVal strBlockID As String, ByRef bNoData As Boolean) _
                                        As FUNCTION_RETURN
    Dim intLoopCnt  As Integer
    Dim rs          As OraDynaset
    Dim sSQL        As String
    Dim intProcNo   As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0

    '�u���b�NID�𓾂�
    sSQL = "SELECT * from XSDCA"
    sSQL = sSQL & " WHERE CRYNUMCA='" & strBlockID & "'"
    sSQL = sSQL & "   AND LIVKCA= '0'"   '�����敪

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        cmbc039_3_CreateXSDCA = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If

    rs.MoveFirst
    intLoopCnt = 0
    BlkOld.GNLC2 = 0
    BlkOld.GNWC2 = 0
    BlkOld.GNMC2 = 0

    Do While Not rs.EOF
        ReDim Preserve HinOld(intLoopCnt)
        ReDim Preserve HinNow(intLoopCnt)
        With HinOld(intLoopCnt)
            If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
            If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
            If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
            If IsNull(rs.Fields("REVNUMCA")) = False Then .REVNUMCA = rs.Fields("REVNUMCA")
            If IsNull(rs.Fields("FACTORYCA")) = False Then .FACTORYCA = rs.Fields("FACTORYCA")
            If IsNull(rs.Fields("OPECA")) = False Then .OPECA = rs.Fields("OPECA")
            If IsNull(rs.Fields("KCKNTCA")) = False Then .KCKNTCA = rs.Fields("KCKNTCA")
            If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLIDCA = rs.Fields("SXLIDCA")
            If IsNull(rs.Fields("XTALCA")) = False Then .XTALCA = rs.Fields("XTALCA")
            If IsNull(rs.Fields("NEKKNTCA")) = False Then .NEKKNTCA = rs.Fields("NEKKNTCA")
            If IsNull(rs.Fields("NEWKNTCA")) = False Then .NEWKNTCA = rs.Fields("NEWKNTCA")
            If IsNull(rs.Fields("NEWKKBCA")) = False Then .NEWKKBCA = rs.Fields("NEWKKBCA")
            If IsNull(rs.Fields("NEMACOCA")) = False Then .NEMACOCA = rs.Fields("NEMACOCA")
            If IsNull(rs.Fields("GNKKNTCA")) = False Then .GNKKNTCA = rs.Fields("GNKKNTCA")
            If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
            If IsNull(rs.Fields("GNWKKBCA")) = False Then .GNWKKBCA = rs.Fields("GNWKKBCA")
            If IsNull(rs.Fields("GNMACOCA")) = False Then .GNMACOCA = rs.Fields("GNMACOCA")
            If IsNull(rs.Fields("GNDAYCA")) = False Then .GNDAYCA = rs.Fields("GNDAYCA")
            If IsNull(rs.Fields("GNLCA")) = False Then .GNLCA = rs.Fields("GNLCA")
            If IsNull(rs.Fields("GNWCA")) = False Then .GNWCA = rs.Fields("GNWCA")
            If IsNull(rs.Fields("GNMCA")) = False Then .GNMCA = rs.Fields("GNMCA")
            BlkOld.GNLC2 = CLng(BlkOld.GNLC2) + CLng(.GNLCA)
            BlkOld.GNWC2 = CLng(BlkOld.GNWC2) + CLng(.GNWCA)
            BlkOld.GNMC2 = CLng(BlkOld.GNMC2) + CLng(.GNMCA)
            If IsNull(rs.Fields("SUMITLCA")) = False Then .SUMITLCA = rs.Fields("SUMITLCA")
            If IsNull(rs.Fields("SUMITWCA")) = False Then .SUMITWCA = rs.Fields("SUMITWCA")
            If IsNull(rs.Fields("SUMITMCA")) = False Then .SUMITMCA = rs.Fields("SUMITMCA")
            If IsNull(rs.Fields("CHGCA")) = False Then .CHGCA = rs.Fields("CHGCA")
            If IsNull(rs.Fields("KAKOUBCA")) = False Then .KAKOUBCA = rs.Fields("KAKOUBCA")
            If IsNull(rs.Fields("KEIDAYCA")) = False Then .KEIDAYCA = rs.Fields("KEIDAYCA")
            If IsNull(rs.Fields("GNTKUBCA")) = False Then .GNTKUBCA = rs.Fields("GNTKUBCA")
            If IsNull(rs.Fields("GNTNOCA")) = False Then .GNTNOCA = rs.Fields("GNTNOCA")
            If IsNull(rs.Fields("XTWORKCA")) = False Then .XTWORKCA = rs.Fields("XTWORKCA")
            If IsNull(rs.Fields("WFWORKCA")) = False Then .WFWORKCA = rs.Fields("WFWORKCA")
            If IsNull(rs.Fields("LSTATBCA")) = False Then .LSTATBCA = rs.Fields("LSTATBCA")
            If IsNull(rs.Fields("RSTATBCA")) = False Then .RSTATBCA = rs.Fields("RSTATBCA")
            If IsNull(rs.Fields("LUFRCCA")) = False Then .LUFRCCA = rs.Fields("LUFRCCA")
            If IsNull(rs.Fields("LUFRBCA")) = False Then .LUFRBCA = rs.Fields("LUFRBCA")
            If IsNull(rs.Fields("LDFRCCA")) = False Then .LDFRCCA = rs.Fields("LDFRCCA")
            If IsNull(rs.Fields("LDFRBCA")) = False Then .LDFRBCA = rs.Fields("LDFRBCA")
            If IsNull(rs.Fields("HOLDCCA")) = False Then .HOLDCCA = rs.Fields("HOLDCCA")
            If IsNull(rs.Fields("HOLDBCA")) = False Then .HOLDBCA = rs.Fields("HOLDBCA")
            If IsNull(rs.Fields("EXKUBCA")) = False Then .EXKUBCA = rs.Fields("EXKUBCA")
            If IsNull(rs.Fields("HENPKCA")) = False Then .HENPKCA = rs.Fields("HENPKCA")
            If IsNull(rs.Fields("LIVKCA")) = False Then .LIVKCA = rs.Fields("LIVKCA")
            If IsNull(rs.Fields("KANKCA")) = False Then .KANKCA = rs.Fields("KANKCA")
            If IsNull(rs.Fields("NFCA")) = False Then .NFCA = rs.Fields("NFCA")
            If IsNull(rs.Fields("SAKJCA")) = False Then .SAKJCA = rs.Fields("SAKJCA")
            If IsNull(rs.Fields("TDAYCA")) = False Then .TDAYCA = rs.Fields("TDAYCA")
            If IsNull(rs.Fields("KDAYCA")) = False Then .KDAYCA = rs.Fields("KDAYCA")
            If IsNull(rs.Fields("SUMITBCA")) = False Then .SUMITBCA = rs.Fields("SUMITBCA")
            If IsNull(rs.Fields("SNDKCA")) = False Then .SNDKCA = rs.Fields("SNDKCA")
            If IsNull(rs.Fields("SNDDAYCA")) = False Then .SNDDAYCA = rs.Fields("SNDDAYCA")
            If IsNull(rs.Fields("PLANTCATCA")) = False Then .PLANTCATCA = rs.Fields("PLANTCATCA")   '2007/09/04 SPK Tsutsumi Add
        End With
        '��{���Ƀf�[�^�������Z�b�g
        With Kihon
            .CNTHINOLD = intLoopCnt + 1
        End With
        intLoopCnt = intLoopCnt + 1
        rs.MoveNext
    Loop

    rs.Close
    cmbc039_3_CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    cmbc039_3_CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'*******************************************************************************************
'*    �֐���        : cmbc039_3_CreateXSDC2
'*
'*    �����T�v      : 1.���������i�u���b�N�j�O�H�����ю擾���\���̍쐬
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^      ,����
'*                    strBlockID    ,O  ,String  ,�u���b�NID
'*                    bNoData�@�@   ,I  ,Boolean ,�f�[�^�L���t���O
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function cmbc039_3_CreateXSDC2(ByVal strBlockID As String, ByRef bNoData As Boolean) _
                                        As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sSQL        As String
    Dim intProcNo   As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False

    '�u���b�NID�𓾂�
    sSQL = "SELECT * from XSDC2 "
    sSQL = sSQL & " WHERE CRYNUMC2='" & strBlockID & "'"
    sSQL = sSQL & "   AND LIVKC2= '0'"   '�����敪

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        bNoData = True
        cmbc039_3_CreateXSDC2 = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If

    rs.MoveFirst
    If rs.EOF = False Then
        With BlkOld
            If IsNull(rs.Fields("CRYNUMC2")) = False Then .CRYNUMC2 = rs.Fields("CRYNUMC2")
            If IsNull(rs.Fields("KCNTC2")) = False Then .KCNTC2 = rs.Fields("KCNTC2")       '�H���A��
            If IsNull(rs.Fields("XTALC2")) = False Then .XTALC2 = rs.Fields("XTALC2")
            If IsNull(rs.Fields("INPOSC2")) = False Then .INPOSC2 = rs.Fields("INPOSC2")
            If IsNull(rs.Fields("NEKKNTC2")) = False Then .NEKKNTC2 = rs.Fields("NEKKNTC2")
            If IsNull(rs.Fields("NEWKNTC2")) = False Then .NEWKNTC2 = rs.Fields("NEWKNTC2")
            If IsNull(rs.Fields("NEWKKBC2")) = False Then .NEWKKBC2 = rs.Fields("NEWKKBC2")
            If IsNull(rs.Fields("NEMACOC2")) = False Then .NEMACOC2 = rs.Fields("NEMACOC2")
            If IsNull(rs.Fields("GNKKNTC2")) = False Then .GNKKNTC2 = rs.Fields("GNKKNTC2")
            If IsNull(rs.Fields("GNWKNTC2")) = False Then .GNWKNTC2 = rs.Fields("GNWKNTC2")
            If IsNull(rs.Fields("GNWKKBC2")) = False Then .GNWKKBC2 = rs.Fields("GNWKKBC2")
            If IsNull(rs.Fields("GNMACOC2")) = False Then .GNMACOC2 = rs.Fields("GNMACOC2")
            If IsNull(rs.Fields("GNDAYC2")) = False Then .GNDAYC2 = rs.Fields("GNDAYC2")
            If IsNull(rs.Fields("GNLC2")) = False Then .GNLC2 = rs.Fields("GNLC2")          '���ݒ���
            If IsNull(rs.Fields("GNWC2")) = False Then .GNWC2 = rs.Fields("GNWC2")          '���ݏd��
            If IsNull(rs.Fields("GNMC2")) = False Then .GNMC2 = rs.Fields("GNMC2")          '���ݖ���
            If IsNull(rs.Fields("SUMITLC2")) = False Then .SUMITLC2 = rs.Fields("SUMITLC2")
            If IsNull(rs.Fields("SUMITWC2")) = False Then .SUMITWC2 = rs.Fields("SUMITWC2")
            If IsNull(rs.Fields("SUMITMC2")) = False Then .SUMITMC2 = rs.Fields("SUMITMC2")
            If IsNull(rs.Fields("CHGC2")) = False Then .CHGC2 = rs.Fields("CHGC2")
            If IsNull(rs.Fields("KAKOUBC2")) = False Then .KAKOUBC2 = rs.Fields("KAKOUBC2")
            If IsNull(rs.Fields("KEIDAYC2")) = False Then .KEIDAYC2 = rs.Fields("KEIDAYC2")
            If IsNull(rs.Fields("GNTKUBC2")) = False Then .GNTKUBC2 = rs.Fields("GNTKUBC2")
            If IsNull(rs.Fields("GNTNOC2")) = False Then .GNTNOC2 = rs.Fields("GNTNOC2")
            If IsNull(rs.Fields("XTWORKC2")) = False Then .XTWORKC2 = rs.Fields("XTWORKC2")
            If IsNull(rs.Fields("WFWORKC2")) = False Then .WFWORKC2 = rs.Fields("WFWORKC2")
            If IsNull(rs.Fields("LSTATBC2")) = False Then .LSTATBC2 = rs.Fields("LSTATBC2")
            If IsNull(rs.Fields("RSTATBC2")) = False Then .RSTATBC2 = rs.Fields("RSTATBC2")
            If IsNull(rs.Fields("LUFRCC2")) = False Then .LUFRCC2 = rs.Fields("LUFRCC2")
            If IsNull(rs.Fields("LUFRBC2")) = False Then .LUFRBC2 = rs.Fields("LUFRBC2")
            If IsNull(rs.Fields("LDFRCC2")) = False Then .LDFRCC2 = rs.Fields("LDFRCC2")
            If IsNull(rs.Fields("LDFRBC2")) = False Then .LDFRBC2 = rs.Fields("LDFRBC2")
            If IsNull(rs.Fields("HOLDCC2")) = False Then .HOLDCC2 = rs.Fields("HOLDCC2")
            If IsNull(rs.Fields("HOLDBC2")) = False Then .HOLDBC2 = rs.Fields("HOLDBC2")
            If IsNull(rs.Fields("EXKUBC2")) = False Then .EXKUBC2 = rs.Fields("EXKUBC2")
            If IsNull(rs.Fields("HENPKC2")) = False Then .HENPKC2 = rs.Fields("HENPKC2")
            If IsNull(rs.Fields("LIVKC2")) = False Then .LIVKC2 = rs.Fields("LIVKC2")
            If IsNull(rs.Fields("KANKC2")) = False Then .KANKC2 = rs.Fields("KANKC2")
            If IsNull(rs.Fields("NFC2")) = False Then .NFC2 = rs.Fields("NFC2")
            If IsNull(rs.Fields("SAKJC2")) = False Then .SAKJC2 = rs.Fields("SAKJC2")
            If IsNull(rs.Fields("TDAYC2")) = False Then .TDAYC2 = rs.Fields("TDAYC2")
            If IsNull(rs.Fields("KDAYC2")) = False Then .KDAYC2 = rs.Fields("KDAYC2")
            If IsNull(rs.Fields("SUMITBC2")) = False Then .SUMITBC2 = rs.Fields("SUMITBC2")
            If IsNull(rs.Fields("SNDKC2")) = False Then .SNDKC2 = rs.Fields("SNDKC2")
            If IsNull(rs.Fields("SNDDAYC2")) = False Then .SNDDAYC2 = rs.Fields("SNDDAYC2")
            If IsNull(rs.Fields("PLANTCATC2")) = False Then .PLANTCATC2 = rs.Fields("PLANTCATC2")   '2007/09/04 SPK Tsutsumi Add
        End With
    End If

    rs.Close
    cmbc039_3_CreateXSDC2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    cmbc039_3_CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'***************************************************************************************************
'*    �֐���        : cmbc039_3_CreateNowProc
'*
'*    �����T�v      : 1.���ݍH���\���̍쐬
'*
'*    �p�����[�^    : �ϐ���           ,IO ,�^      ,����
'*                    strSXLID         ,I  ,String  ,SXL-ID
'*                    lngBeginIngotpos ,I  ,Long    ,�u���b�N�Ǘ��f�[�^�̒���
'*                    strErrMsg        ,O  ,String  ,ErrMsg�i�[
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************************
'Cng Start 2010/10/03 Y.Hitomi
Public Function cmbc039_3_CreateNowProc(ByVal strBlockID As String, ByVal lngBeginIngotpos As Long, _
                                        ByVal lngEndIngotpos As Long, ByVal lngWfBeginSeq As Long, _
                                        ByVal lngWfEndSeq As Long, _
                                        ByRef strErrMsg As String, ByVal strSXLID As String) As FUNCTION_RETURN
'Public Function cmbc039_3_CreateNowProc(ByVal strBlockID As String, ByVal lngBeginIngotpos As Long, _
                                        ByVal lngEndIngotpos As Long, ByVal lngWfBeginSeq As Long, _
                                        ByVal lngWfEndSeq As Long, _
                                        ByRef strErrMsg As String) As FUNCTION_RETURN
'Cng End 2010/10/03 Y.Hitomi

    Dim rs              As OraDynaset
    Dim rs2             As OraDynaset
    Dim sSQL            As String
    Dim intProcNo       As Integer
    Dim intHinOldCnt    As Integer
    Dim intLengthCnt    As Integer
    Dim intLoopCnt      As Integer
    Dim dblDiameter     As Double
    Dim intNum          As Integer
    Dim sCryNum         As String
    Dim intBlkLength    As Integer  '�u���b�N�Ǘ��f�[�^�̒���
    Dim intBlkIngotPos  As Integer  '�u���b�N�Ǘ��f�[�^�̈ʒu
    Dim intSxlLength    As Integer  '�V���O���Ǘ��f�[�^�̒���
    Dim intSxlIngotPos  As Integer  '�V���O���Ǘ��f�[�^�̈ʒu
    Dim blFlg           As Boolean
    Dim intSP           As Integer  '��������p
    Dim intEP           As Integer  '��������p
    Dim intSBP          As Integer  '��������p
    Dim intEBP          As Integer  '��������p
    Dim intLength       As Integer  '����
    Dim intIngotpos     As Integer  '�ʒu
    Dim blRtn           As Boolean  '�߂�l
    Dim intWFcnt        As Integer  'WF�}�b�v���� add 2003/03/29 hitec)matsumoto
    Dim sMotoHinban     As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0

    '�u���b�N�Ǘ�������擾
    sSQL = "SELECT * from TBCME040 "
    sSQL = sSQL & " WHERE BLOCKID='" & strBlockID & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    intLoopCnt = 0
    If rs.EOF = False Then
        If IsNull(rs("CRYNUM")) = False Then sCryNum = rs("CRYNUM")           '�����ԍ�
        If IsNull(rs("LENGTH")) = False Then intBlkLength = rs("LENGTH")        '����
        If IsNull(rs("INGOTPOS")) = False Then intBlkIngotPos = rs("ingotpos")  '�ʒu
    End If

    rs.Close

    'upd start 2003/03/29 hitec)matsumoto �S���p�����́AWF�}�b�v������----------

    '�u���b�NID�𓾂�       'del 2003/03/29 hitec)matsumoto ���ֈړ�
    sSQL = "SELECT LOTID from TBCMY011 "
    sSQL = sSQL & " WHERE LOTID='" & strBlockID & "'"     '2003/04/03 hitec)matsumoto �S���X�N���b�v="Y"�̓u���b�N�P�ʂȂ̂ŁA�V���O���͈͂Ŏ��Ȃ�
    sSQL = sSQL & "   AND TO_NUMBER(WFSTA) <= 1"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        '�O�H�����т��A���ݍH�����тɃR�s�[
        BlkNow = BlkOld
        BlkNow.GNLC2 = "0"
        BlkNow.GNWC2 = "0"
        BlkNow.GNMC2 = "0"
        BlkNow.GNWKNTC2 = Kihon.NEWPROC
        BlkNow.NEWKNTC2 = Kihon.NOWPROC
        For intHinOldCnt = 0 To Kihon.CNTHINOLD - 1
            ReDim Preserve HinNow(intHinOldCnt) As typ_XSDCA_Update
            HinNow(intHinOldCnt) = HinOld(intHinOldCnt)
            HinNow(intHinOldCnt).GNLCA = "0"    '�S���X�N���b�v=������0
            HinNow(intHinOldCnt).GNWCA = "0"    '�d�� = 0
            HinNow(intHinOldCnt).GNMCA = "0"    '���� = 0
            HinNow(intHinOldCnt).GNWKNTCA = Kihon.NEWPROC
            HinNow(intHinOldCnt).NEWKNTCA = Kihon.NOWPROC
        Next
        Kihon.CNTHINNOW = 1
        Kihon.ALLSCRAP = "Y"

    '�O�H���̒����ƌ��ݍH���̒���������ׁA�s�ǂ����݂��邩���� 'upd 2003/03/29 hitec)matsumoto �����ł͂Ȃ������Ŕ�ׂ�
        If CInt(BlkNow.GNMC2) = CInt(BlkOld.GNMC2) Then '�s�ǂȂ�
            '��{���\����
            With Kihon
                .FURYOUMU = "N"
            End With
        Else
            rs.Close
            strErrMsg = GetMsgStr("EWFM5", "�O�H��=" & BlkOld.GNMC2 & "�F���ݍH��=" & BlkNow.GNMC2) '03/06/06 �㓡
            cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        rs.Close
        cmbc039_3_CreateNowProc = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    rs.Close

    '�O�H���̍\���̂����ݍH���̍\���̂փR�s�[
    BlkNow = BlkOld
    '�H���A�ԂɁ{�P����
    With BlkNow
        If BlkNow.KCNTC2 = "" Then BlkNow.KCNTC2 = "0"
        .KCNTC2 = CInt(BlkNow.KCNTC2) + 1   '�H���A��
        'Cng Start 2010/09/02 Y.Hitomi
        ''�u���b�N��SXL��1�ł��������Ă����ꍇ�A�H���R�[�h���X�V���Ȃ��悤�ɂ���
        If (.GNWKNTC2 <> "     " Or _
            .GNWKNTC2 <> "CW800" Or _
            .GNWKNTC2 <> "TX860") Then
            
            .NEWKNTC2 = Kihon.NOWPROC           '�O�H��
            .GNWKNTC2 = Kihon.NEWPROC           '���ݍH��
        End If
'            .NEWKNTC2 = Kihon.NOWPROC           '�O�H��
'            .GNWKNTC2 = Kihon.NEWPROC           '���ݍH��
        'Cng End 2010/09/02 Y.Hitomi
        
        .SUMITBC2 = "0"
        .SUMITLC2 = "0"
        .SUMITMC2 = "0"
        .SUMITWC2 = "0"
    End With

    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "  xtalcb as CRYNUM"        ' �����ԍ�
    sSQL = sSQL & " ,inposcb as INGOTPOS"     ' �������J�n�ʒu
    sSQL = sSQL & " ,rlencb as LENGTH"        ' ����
    sSQL = sSQL & " ,sxlidcb as SXLID"        ' SXLID
    sSQL = sSQL & " ,' ' as KRPROCCD"         ' �Ǘ��H��
    sSQL = sSQL & " ,gnwkntcb as NOWPROC"     ' ���ݍH��
    sSQL = sSQL & " ,' ' as LPKRPROCCD"       ' �ŏI�ʉߊǗ��H��
    sSQL = sSQL & " ,newkntcb as LASTPASS"    ' �ŏI�ʉߍH��
    sSQL = sSQL & " ,livkcb as DELCLS"        ' �폜�敪
    sSQL = sSQL & " ,lstccb as LSTATCLS"      ' �ŏI��ԋ敪
    sSQL = sSQL & " ,sholdclscb HOLDCLS"      ' �z�[���h�敪
    sSQL = sSQL & " ,hinbcb as HINBAN"        ' �i��
    sSQL = sSQL & " ,revnumcb as REVNUM"      ' ���i�ԍ������ԍ�
    sSQL = sSQL & " ,factorycb as FACTORY"    ' �H��
    sSQL = sSQL & " ,opecb as OPECOND"        ' ���Ə���
    sSQL = sSQL & " ,maicb"                   ' ����
    sSQL = sSQL & " ,tdaycb as REGDATE"       ' �o�^���t
    sSQL = sSQL & " ,kdaycb as UPDDATE"       ' �X�V���t
    sSQL = sSQL & " ,' ' as SUMMITSENDFLAG"   ' SUMMIT���M�t���O
    sSQL = sSQL & " ,sndkcb as SENDFLAG"      ' ���M�t���O
    sSQL = sSQL & " ,sndaycb as SENDDATE"     ' ���M���t
    sSQL = sSQL & " ,plantcatcb as PLANTCAT"  ' ���� 2007/09/04 SPK Tsutsumi Add
    sSQL = sSQL & " FROM XSDCB"
    sSQL = sSQL & " WHERE sxlidcb like '" & left(sCryNum, 9) & "%'"     '���ޯ�����ڒǉ� 09/05/25 ooba
    sSQL = sSQL & "   AND xtalcb = '" & sCryNum & "'"
    sSQL = sSQL & "   AND ((inposcb >=" & lngBeginIngotpos & ") And (inposcb + rlencb <= " & lngEndIngotpos & "))"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    intLoopCnt = 0
    BlkNow.GNMC2 = 0    '���ݍH���i�u���b�N�j�̖������N���A���Ă���
    Do While Not rs.EOF
        ReDim Preserve HinNow(intLoopCnt) As typ_XSDCA_Update

        If IsNull(rs("CRYNUM")) = False Then sCryNum = rs("CRYNUM")
        If IsNull(rs("LENGTH")) = False Then intSxlLength = rs("LENGTH")
        If IsNull(rs("INGOTPOS")) = False Then intSxlIngotPos = rs("INGOTPOS")

        '-- �u���b�N�ƃV���O���̈ʒu�֌W�𔻒肵�A�������Z�o --------
        intSP = intSxlIngotPos         '�V���O���J�n�ʒu
        intEP = intSP + intSxlLength      '�V���O���I�[�ʒu
        intSBP = intBlkIngotPos        '�u���b�N�J�n�ʒu
        intEBP = intSBP + intBlkLength    '�u���b�N�I�[�ʒu

        '' �u���b�N��SXL�̒��Ɋ��S�Ɋ܂܂�Ă���ꍇ ---------
        If intSP <= intSBP And intEP >= intEBP Then

            intLength = intBlkLength                    '�u���b�N�Ǘ��̒������g�p
            intIngotpos = intBlkIngotPos

        '' �u���b�N��SXL�̊J�n�ʒu����ɂ���A���I�[�ʒu���������ꍇ ---------
        ElseIf intSP >= intSBP And intEP <= intEBP Then

            intLength = intSxlLength                  '�V���O���Ǘ��̒������g�p
            intIngotpos = intSxlIngotPos

        '' �u���b�N���ꕔSXL�ɂ������Ă���ꍇ
        '' (�u���b�N���㑤�B�������u���b�N�̏I�[��SXL�̊J�n�ʒu����v���Ȃ�����) ------------
        ElseIf intSP > intSBP And intSP < intEBP And intSP <> intEBP Then

            intLength = intEBP - intSP                        '�u���b�N�̏I�[�ʒu - �V���O���̊J�n�ʒu
            intIngotpos = intSxlIngotPos

        '' �u���b�N���ꕔSXL�ɂ������Ă���ꍇ
        '' (�u���b�N�������B������SXL�̏I�[�ƃu���b�N�̊J�n�ʒu����v���Ȃ�����) ----------
        ElseIf intSP < intSBP And intEP > intSBP And intEP <> intSBP Then

            intLength = intEP - intSBP                        '�V���O���̏I�[�ʒu - �u���b�N�̊J�n�ʒu
            intIngotpos = intBlkIngotPos

        Else
            GoTo LoopNext
        End If
        '----------------------------------------------------

        '���ݍH���ҏW
        With HinNow(intLoopCnt)
            .CRYNUMCA = strBlockID       '�u���b�NID
            If IsNull(rs("HINBAN")) = False Then .HINBCA = rs("HINBAN")         '�i��
            If IsNull(rs("SXLID")) = False Then .SXLIDCA = rs("SXLID")          '�V���O��ID
            If IsNull(rs("REVNUM")) = False Then .REVNUMCA = rs("REVNUM")       '���i�ԍ������ԍ�
            If IsNull(rs("FACTORY")) = False Then .FACTORYCA = rs("FACTORY")    '�H��
            If IsNull(rs("OPECOND")) = False Then .OPECA = rs("OPECOND")        '���Ə���
            If IsNull(rs("CRYNUM")) = False Then .XTALCA = rs("CRYNUM")         '�����ԍ�
            'add start 2003/04/27 hitec)matsumoto Z�i�Ԃ̂��̂́A�Ǖi�Ƃ����邽�߁A���i�Ԃ��擾����------------------
            If Trim(.HINBCA) = "Z" Then
'Cng Start 2012/4/26 Y.Hitomi
'                If GetZMotoHinban(.SXLIDCA, sMotoHinban) = FUNCTION_RETURN_FAILURE Then
'                    cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
'                    rs.Close
'                    GoTo proc_exit
'                End If
'                If sMotoHinban = vbNullString Then
'                    cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
'                    rs.Close
'                    GoTo proc_exit
'                End If
'                .HINBCA = sMotoHinban
                .HINBCA = tblSXL.hinban '���i�ԏ��
'Cng End 2012/4/26 Y.Hitomi
                .FACTORYCA = tblSXL.factory   '�H��
                .OPECA = tblSXL.opecond       '���Ə���
                .REVNUMCA = tblSXL.REVNUM     '�����ԍ�
                'Add Start 2010/10/14 Y.Hitomi
                .GNWKNTCA = PROCD_SXL_MAP     '���ݍH��
                'Add End   2010/10/14 Y.Hitomi
            Else
                If IsNull(rs("FACTORY")) = False Then .FACTORYCA = rs("FACTORY")    '�H��
                If IsNull(rs("OPECOND")) = False Then .OPECA = rs("OPECOND")        '���Ə���
                If IsNull(rs("REVNUM")) = False Then .REVNUMCA = rs("REVNUM")       '���i�ԍ������ԍ�
                If IsNull(rs("PLANTCAT")) = False Then .PLANTCATCA = rs("PLANTCAT")     '���� 2007/09/04 SPK Tsutsumi Add
                'Add Start  2010/10/14 Y.Hitomi
                .GNWKNTCA = Kihon.NEWPROC     '���ݍH��
                'Add End    2010/10/14 Y.Hitomi
            End If
            'add end   2003/04/27 hitec)matsumoto ------------------
        
            .NEWKNTCA = Kihon.NOWPROC   '�O�H��
            'Cng Start 2010/10/14 Y.Hitomi
            '.GNWKNTCA = Kihon.NEWPROC   '���ݍH��
            'Cng End   2010/10/14 Y.Hitomi
            
            .KCKNTCA = BlkNow.KCNTC2    '�H���A��
            .NEMACOCA = BlkNow.NEMACOC2 '�ŏI�ʉߏ�����
            .GNMACOCA = BlkNow.GNMACOC2 '���ݏ�����
            .SUMITBCA = "0"
            .SUMITLCA = "0"
            .SUMITMCA = "0"
            .SUMITWCA = "0"
            .INPOSCA = intIngotpos      '�������J�n�ʒu
            .GNLCA = intLength          '����

            '���ݏd�ʂ����߂�
            If GetDiameter(strBlockID, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '���a�����߂�
                dblDiameter = 0
            End If
            '�擾�������a�����ɏd�ʂ����߂�
            HinNow(intLoopCnt).GNWCA = CStr(WeightOfCylinder(dblDiameter, CDbl(.GNLCA)))

            sSQL = "SELECT LOTID from TBCMY011 "         'upd 2003/04/29 hitec)matsumoto �V���O���������Ɏ擾����̂ł͂Ȃ��A�u���b�N�P�ʂŖ������擾�ł���悤�A������ύX
            sSQL = sSQL & " WHERE LOTID = '" & .CRYNUMCA & "'"
            sSQL = sSQL & " AND MSXLID ='" & .SXLIDCA & "'"    ''' 03/04/27 �C�� �㓡
            sSQL = sSQL & "   AND TO_NUMBER(WFSTA) <= 1"

            Set rs2 = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
            intWFcnt = 0
            Do While Not rs2.EOF
                intWFcnt = intWFcnt + 1
                rs2.MoveNext
            Loop
            rs2.Close

            .GNMCA = intWFcnt   'add 2003/03/29 hitec)matsumoto ���WF�}�b�v�e�[�u�����疇���J�E���g�擾���Ă���̂ŁA�����Ǖi�����Ƃ���
            .SUMITLCA = .GNLCA ''' 03/05/13 �㓡
            .SUMITMCA = .GNMCA
            .SUMITWCA = .GNWCA
        End With

        With BlkNow
            '���ݏd�ʂ����߂�
            If GetDiameter(strBlockID, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '���a�����߂�
                dblDiameter = 0
            End If
            '��{���̒��a�Z�b�g
            Kihon.DIAMETER = dblDiameter
            '�擾�������a�����ɏd�ʂ����߂�
            .GNMC2 = CStr(CLng(BlkNow.GNMC2) + CLng(HinNow(intLoopCnt).GNMCA))  '���� 'upd 2003/03/29 hitec)matsumoto �����Čv�Z�͂��Ȃ�
        End With
        intLoopCnt = intLoopCnt + 1
        '�Ǖi�����Z�b�g
        With Kihon
            .CNTHINNOW = intLoopCnt
        End With

LoopNext:

        rs.MoveNext
    Loop

    rs.Close

    '���ݍH���ō��ꂽ���̈ȊO�̃f�[�^���O�H���ɑ��݂����ꍇ�擾
    If GetOtherData(intBlkLength, intBlkIngotPos) = FUNCTION_RETURN_FAILURE Then
        cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '�O�H���̒����ƌ��ݍH���̒���������ׁA�s�ǂ����݂��邩����
    If BlkNow.GNLC2 = "" Then BlkNow.GNLC2 = "0"
    If BlkOld.GNLC2 = "" Then BlkOld.GNLC2 = "0"
    If BlkNow.GNMC2 = "" Then BlkNow.GNMC2 = "0"
    If BlkOld.GNMC2 = "" Then BlkOld.GNMC2 = "0"
    If CInt(BlkNow.GNMC2) = CInt(BlkOld.GNMC2) Then '�s�ǂȂ�
        '��{���\����
        With Kihon
            .FURYOUMU = "N"
        End With
    Else                                            '�s�ǂ���
        rs.Close
        strErrMsg = GetMsgStr("EWFM5", "�O�H��=" & BlkOld.GNMC2 & "�F���ݍH��=" & BlkNow.GNMC2) '03/06/06 �㓡
        cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    cmbc039_3_CreateNowProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    cmbc039_3_CreateNowProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    �֐���        : GetOtherData
'*
'*    �����T�v      : 1.SXL�̑S�u���b�N���Ƀ`�F�b�N
'*
'*    �p�����[�^    : �ϐ���           ,IO ,�^      ,����
'*                    intBlkLength     ,I  ,Integer ,�u���b�N�Ǘ��f�[�^�̒���
'*                    lngBeginIngotpos ,I  ,Integer ,�u���b�N�Ǘ��f�[�^�̒���
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Private Function GetOtherData(ByVal intBlkLength As Integer, ByVal intBlkIngotPos As Integer) _
                                As FUNCTION_RETURN
    Dim rs              As OraDynaset
    Dim sSQL            As String
    Dim intHinOldCnt    As Integer
    Dim intHinNowCnt    As Integer
    Dim intHinCnt       As Integer
    Dim blUpdFlg        As Boolean
    Dim intFuryouCnt    As Integer
    Dim intWFcnt        As Integer

    intHinCnt = Kihon.CNTHINNOW

    For intHinOldCnt = 0 To Kihon.CNTHINOLD - 1
        blUpdFlg = False
        For intHinNowCnt = 0 To intHinCnt - 1
            If (HinOld(intHinOldCnt).XTALCA = HinNow(intHinNowCnt).XTALCA) _
                And (HinOld(intHinOldCnt).INPOSCA = HinNow(intHinNowCnt).INPOSCA) Then

                blUpdFlg = True
            End If
        Next

        If blUpdFlg = False Then
            '�O�H���i�Ԃɂ����āA���ݍH���i�ԂɂȂ����̂́A�O�H���i�Ԃ����ݍH���i�ԂɃR�s�[
            ReDim Preserve HinNow(Kihon.CNTHINNOW) As typ_XSDCA_Update
            '�O�H���i�Ԃ��R�s�[
            HinNow(Kihon.CNTHINNOW) = HinOld(intHinOldCnt)
            With HinNow(Kihon.CNTHINNOW)
                .KCKNTCA = BlkNow.KCNTC2
                'Cng Start 2010/10/14 Y.Hitomi
                .NEWKNTCA = HinOld(intHinOldCnt).NEWKNTCA
                .GNWKNTCA = HinOld(intHinOldCnt).GNWKNTCA
'                .NEWKNTCA = Kihon.NOWPROC
'                .GNWKNTCA = Kihon.NEWPROC
                'Cng End 2010/10/14 Y.Hitomi
                .SUMITBCA = "0"

                sSQL = "SELECT LOTID from TBCMY011 "
                sSQL = sSQL & " WHERE LOTID = '" & .CRYNUMCA & "'"
                sSQL = sSQL & "   AND MSXLID ='" & .SXLIDCA & "'"
                sSQL = sSQL & "   AND TO_NUMBER(WFSTA) <= 1"

                Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
                intWFcnt = 0
                Do While Not rs.EOF
                    intWFcnt = intWFcnt + 1
                    rs.MoveNext
                Loop
                rs.Close
                .GNMCA = intWFcnt

                BlkNow.GNLC2 = CLng(BlkNow.GNLC2) + val(.GNLCA)
                BlkNow.GNMC2 = CLng(BlkNow.GNMC2) + val(.GNMCA)
                BlkNow.GNWC2 = CLng(BlkNow.GNWC2) + val(.GNWCA)
            End With
            Kihon.CNTHINNOW = Kihon.CNTHINNOW + 1
        End If
    Next
    GetOtherData = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function
End Function

'*******************************************************************************
'*    �֐���        : Shikibetsu
'*
'*    �����T�v      : 1.�T���v���f�[�^�쐬����
'*                    (��R�ۏ؃t���O(HSXRHWYS)�A�_�f�ۏ؃t���O(HSXONHWS))
'*�@�@�@�@�@�@�@�@�@�@(�ǂ��炩�ł��u�g�v��������True)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^      ,����
'*                    hinb          ,I  ,String  ,�i��
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function Shikibetsu(ByVal hinb As String) As Boolean
    Dim blDbIsMine      As Boolean
    Dim rs              As OraDynaset
    Dim sSQL            As String
    Dim i               As Integer
    Dim intRecCnt       As Integer
    Dim sWork1, sWork2  As String

    Shikibetsu = False
    If OraDB Is Nothing Then
        blDbIsMine = True
        OraDBOpen
    End If

    ''�ėp�R�[�h�}�X�^����A�R�[�hNO�ɑΉ�����R�[�h�̈ꗗ�𓾂�
    sSQL = "select E18.HSXRHWYS, E19.HSXONHWS"
    sSQL = sSQL & " from TBCME018 E18 ,TBCME019 E19 "
    sSQL = sSQL & " where rtrim(E18.HINBAN) = '" & Trim$(hinb) & "' "
    sSQL = sSQL & " and E18.HINBAN = E19.HINBAN "
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs.EOF <> True Then
        sWork1 = rs("HSXRHWYS")
        sWork2 = rs("HSXONHWS")
        If Trim(sWork1) = "H" Or Trim(sWork2) = "H" Then
            Shikibetsu = True
        End If
    End If
    rs.Close

    If blDbIsMine Then
        OraDBClose
    End If
End Function

'****************************************************************************************
'*    �֐���        : DBDRV_MIN_MAX_SEQGET
'*
'*    �����T�v      : 1.�����w�� MIN,MAX�l���擾
'*                      (SXLID,BLOCKID���ő�A�ŏ��i�u���b�N�o�Ŕ���j�̃f�[�^���擾����)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    iWfNum        ,O  ,Integer  ,SXL�Ǘ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************
Public Function DBDRV_MIN_MAX_SEQGET(ByRef iWfNum As Integer) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim lngCnt      As Long
    Dim sDBName     As String
    Dim intUCount   As Integer
    Dim dblWFLen    As Double  '2003/04/25 hitec)okazaki
    Dim iRtn        As FUNCTION_RETURN
    Dim dblEPS      As Double

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_MIN_MAX_SEQGET"

    dblEPS = 0.000001        '�Â̐ݒ�
    iWfNum = 0
    sDBName = "(V001)"
    i = 0

    ' SXLID�̎擾
    For i = 0 To UBound(tSXLID())
        sSQL = "select "
        sSQL = sSQL & "LOTID,"                ' �u���b�NID"
        sSQL = sSQL & "MSXLID,"                ' SXLID"
        sSQL = sSQL & "blockseq,"             ' �u���b�N���A��"
        sSQL = sSQL & "WFSTA,"                ' WF���"
        sSQL = sSQL & "MHINBAN,"               ' �i��"
        sSQL = sSQL & "RTOP_POS,"             ' �_���u���b�N���ʒu"
        sSQL = sSQL & "RITOP_POS,"            ' �_���������ʒu"
        sSQL = sSQL & "MSMPLEID,"              ' �����ʒu"
        sSQL = sSQL & "SHAFLAG,"               ' �T���v���t���O"
        sSQL = sSQL & "INDTM,"
        sSQL = sSQL & "BASKETID,"
        sSQL = sSQL & "SLOTNO,"
        sSQL = sSQL & "CURRWPCS,"
        sSQL = sSQL & "EXISTFLG,"
        sSQL = sSQL & "TOP_POS,"
        sSQL = sSQL & "REJCAT,"
        sSQL = sSQL & "TXID,"
        sSQL = sSQL & "REGDATE,"
        sSQL = sSQL & "SUMMITSENDFLAG,"
        sSQL = sSQL & "SENDFLAG,"
        sSQL = sSQL & "SENDDATE,"
        sSQL = sSQL & "HREJCODE,"
        sSQL = sSQL & "UPDPROC,"
        sSQL = sSQL & "UPDDATE,"
        sSQL = sSQL & "MREVNUM,"
        sSQL = sSQL & "MFACTORY,"
        sSQL = sSQL & "MOPECOND,"
        sSQL = sSQL & "kankbn,"
        sSQL = sSQL & "NREJCODE"
        sSQL = sSQL & " from TBCMY011 "
        sSQL = sSQL & " where LOTID ='" & tSXLID(i).LOTID & "'"
        sSQL = sSQL & "   and MSXLID ='" & tSXLID(i).SXLID & "'"
        sSQL = sSQL & " ORDER BY blockseq ASC"

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        iWfNum = 0
        Do While Not rs.EOF
            If CInt(rs.Fields("WFSTA")) <= 1 Then
                iWfNum = iWfNum + 1
            End If
            rs.MoveNext
        Loop
        If rs.RecordCount = 0 Then
            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
            f_cmbc039_3.lblMsg.Caption = GetMsgStr("EWFM6", "(Y011)")
            Exit Function
        End If
        rs.MoveFirst    '�擪ں��ނɈړ�

        Do While Not rs.EOF
            '�擪ں���
            ReDim Preserve tExamine(intUCount)   '�z��̍Ē�`
            With tExamine(intUCount)
                If IsNull(rs!LOTID) = True Then
                    .LOTID = vbNullString           ' �u���b�NID
                 Else
                    .LOTID = rs!LOTID
                End If
                If IsNull(rs!MSXLID) = True Then
                    .SXLID = vbNullString
                Else
                    .SXLID = rs!MSXLID               ' SXLID
                End If
                .MinMax = 0                         ' 0:MIN 1:MAX
                If IsNull(rs!BLOCKSEQ) = True Then
                    .BLOCKSEQ = vbNullString
                Else
                    .BLOCKSEQ = rs!BLOCKSEQ         ' �u���b�N���A��
                End If
                If IsNull(rs!WFSTA) = True Then
                    .WFSTA = vbNullString
                Else
                    .WFSTA = rs!WFSTA               ' WF���
                End If
                If IsNull(rs!mhinban) = True Then
                    .hinban = vbNullString
                Else
                    .hinban = rs!mhinban             ' �i��
                End If
                If IsNull(rs!MSMPLEID) = True Then
                    .SMPLEID = vbNullString
                Else
                    .SMPLEID = rs!MSMPLEID          ' �����ʒu
                End If
                If IsNull(rs!RTOP_POS) = True Then
                    .RTOP_POS = 0
                Else
                    'WF�ꖇ�̒����擾                                   '2003/04/25 hitec)okazaki
                    iRtn = DBDRV_WFLENGET(tSXLID(i).LOTID, dblWFLen)
                    '�u���b�N�擪�̕\���ʒu��WF�ꖇ�̒���������������   '2003/04/25 hitec)okazaki
                    If Right(.SMPLEID, 1) <> "D" Then
                        .RTOP_POS = Fix(CDbl(rs.Fields("RTOP_POS")) - dblWFLen + 0.99999)       ' �_���u���b�N���ʒu  'upd 2003/08/06 hitec)matsumoto
                    Else
                        .RTOP_POS = Int(CDbl(rs.Fields("RTOP_POS")) + dblEPS)      'D�̏ꍇ�؂�̂�(WF�����ɐ����ʒu��2�ȏ゠��ꍇ�̑Ή�)�@06/10/27 ooba
                    End If
                End If
                If IsNull(rs!RITOP_POS) = True Then
                    .RITOP_POS = 0
                Else
                    '�u���b�N�擪�̕\���ʒu��WF�ꖇ�̒���������������   '2003/04/25 hitec)okazaki
                    If Right(.SMPLEID, 1) <> "D" Then
                        .RITOP_POS = Fix(CDbl(rs.Fields("RITOP_POS")) - dblWFLen + 0.99999)        ' �_���������ʒu   'upd 2003/08/06 hitec)matsumoto
                    Else
                        .RITOP_POS = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS)    'D�̏ꍇ�؂�̂�(WF�����ɐ����ʒu��2�ȏ゠��ꍇ�̑Ή�)�@06/10/27 ooba
                    End If
                End If
                If IsNull(rs!SHAFLAG) = True Then
                    .SHAFLAG = vbNullString
                Else
                    .SHAFLAG = rs!SHAFLAG           ' �T���v���t���O
                    If Trim(.SHAFLAG) = "1" Then
                        If Trim(.SMPLEID) = vbNullString Then   'add 2003/06/24 hitec)matsumoto �T���v���t���O��
                            f_cmbc039_3.cmdF(6).Enabled = False
                            f_cmbc039_3.cmdF(7).Enabled = False
                            f_cmbc039_3.cmdF(8).Enabled = False
                            f_cmbc039_3.cmdF(9).Enabled = False
                            f_cmbc039_3.cmdF(10).Enabled = False
                            f_cmbc039_3.cmdF(12).Enabled = False
                            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
                            f_cmbc039_3.lblMsg.Caption = GetMsgStr("ENSP4", "Y011")
                            rs.Close
                            Exit Function
                        End If
                    End If
                End If
                If IsNull(rs!INDTM) = True Then
                Else
                    .INDTM = rs!INDTM               ' �E�F�n�[�Z���^�[���ɓ���
                End If
                If IsNull(rs!BASKETID) = True Then
                    .BASKETID = vbNullString
                Else
                    .BASKETID = rs!BASKETID         ' �o�X�P�b�gID
                End If
                If IsNull(rs!SLOTNO) = True Then
                    .SLOTNO = vbNullString
                Else
                    .SLOTNO = rs!SLOTNO             ' �X���b�gNO
                End If
                If IsNull(rs!CURRWPCS) = True Then
                    .CURRWPCS = vbNullString
                Else
                    .CURRWPCS = iWfNum              ' �E�F�n�[����
                End If
                If IsNull(rs!EXISTFLG) = True Then
                    .EXISTFLG = vbNullString
                Else
                    .EXISTFLG = rs!EXISTFLG         ' ���݃t���O
                End If
                If IsNull(rs!TOP_POS) = True Then
                    .TOP_POS = 0
                Else
                    .TOP_POS = Int(CDbl(rs!TOP_POS) / 10 + dblEPS)         ' �u���b�N��TOP����̈ʒu   'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                End If
                If IsNull(rs!REJCAT) = True Then
                    .REJCAT = vbNullString
                Else
                    .REJCAT = rs!REJCAT             ' �������R
                End If
                If IsNull(rs!TXID) = True Then
                    .TXID = vbNullString
                Else
                    .TXID = rs!TXID                 ' �g�����U�N�V����ID
                End If
                If IsNull(rs!REGDATE) = True Then
                Else
                    .REGDATE = rs!REGDATE           ' �o�^���t
                End If
                If IsNull(rs!SUMMITSENDFLAG) = True Then
                    .SUMMITSENDFLAG = vbNullString
                Else
                    .SUMMITSENDFLAG = rs!SUMMITSENDFLAG ' SUMIT���M�t���O
                End If
                If IsNull(rs!SENDFLAG) = True Then
                    .SENDFLAG = vbNullString
                Else
                    .SENDFLAG = rs!SENDFLAG         ' ���M�t���O
                End If
                If IsNull(rs!SENDDATE) = True Then
                Else
                    .SENDDATE = rs!SENDDATE         ' ���M���t
                End If
                If IsNull(rs!HREJCODE) = True Then
                    .HREJCODE = vbNullString
                Else
                    .HREJCODE = rs!HREJCODE         ' �s�Ǘ��R�R�[�h
                End If
                If IsNull(rs!UPDPROC) = True Then
                    .UPDPROC = vbNullString
                Else
                    .UPDPROC = rs!UPDPROC           ' �X�V�H��
                End If
                If IsNull(rs!UPDDATE) = True Then
                Else
                    .UPDDATE = rs!UPDDATE           ' �X�V���t
                End If
                If IsNull(rs!MREVNUM) = True Then
                    .REVNUM = 0
                Else
                    .REVNUM = rs!MREVNUM             ' ���i�ԍ������ԍ�
                End If
                If IsNull(rs!Mfactory) = True Then
                    .factory = vbNullString
                Else
                    .factory = rs!Mfactory           ' �H��
                End If
                If IsNull(rs!Mopecond) = True Then
                    .opecond = vbNullString
                Else
                    .opecond = rs!Mopecond           ' ���Ə���
                End If
                If IsNull(rs!KANKBN) = True Then
                    .KANKBN = vbNullString
                Else
                    .KANKBN = rs!KANKBN             ' �����敪
                End If
                If IsNull(rs!NREJCODE) = True Then
                    .NREJCODE = vbNullString
                Else
                    .NREJCODE = rs!NREJCODE         ' �����ԓ����R�R�[�h
                End If
            End With

            '�ŏIں���
            rs.MoveLast                             '�ŏIں��ނɈړ�
            intUCount = intUCount + 1
            ReDim Preserve tExamine(intUCount)    '�z��̍Ē�`
            With tExamine(intUCount)
                If IsNull(rs!LOTID) = True Then
                    .LOTID = vbNullString           ' �u���b�NID
                Else
                    .LOTID = rs!LOTID
                End If
                If IsNull(rs!MSXLID) = True Then
                    .SXLID = vbNullString
                Else
                    .SXLID = rs!MSXLID               ' SXLID
                End If
                .MinMax = 1                          ' 0:MIN 1:MAX
                If IsNull(rs!BLOCKSEQ) = True Then
                    .BLOCKSEQ = vbNullString
                Else
                    .BLOCKSEQ = rs!BLOCKSEQ         ' �u���b�N���A��
                End If
                If IsNull(rs!WFSTA) = True Then
                    .WFSTA = vbNullString
                Else
                    .WFSTA = rs!WFSTA               ' WF���
                End If
                If IsNull(rs!mhinban) = True Then
                    .hinban = vbNullString
                Else
                    .hinban = rs!mhinban             ' �i��
                End If
                If IsNull(rs!RTOP_POS) = True Then
                    .RTOP_POS = 0
                Else
                    .RTOP_POS = Int(CDbl(rs.Fields("RTOP_POS")) + dblEPS)           ' �_���u���b�N���ʒu   'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                End If
                If IsNull(rs!RITOP_POS) = True Then
                    .RITOP_POS = 0
                Else
                    .RITOP_POS = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS)        ' �_���������ʒu    'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                End If
                If IsNull(rs!MSMPLEID) = True Then
                    .SMPLEID = vbNullString
                Else
                    .SMPLEID = rs!MSMPLEID           ' �����ʒu
                    .SMPLEID = tblsmp(2).SMPLID      ' �����ʒu    2003/10/26 SystemBrain
                End If
                If IsNull(rs!SHAFLAG) = True Then
                    .SHAFLAG = vbNullString
                Else
                    .SHAFLAG = rs!SHAFLAG            ' �T���v���t���O
                    If Trim(.SHAFLAG) = "1" Then
                        If Trim(.SMPLEID) = vbNullString Then   'add 2003/06/24 hitec)matsumoto �T���v���t���O��
                            f_cmbc039_3.cmdF(6).Enabled = False
                            f_cmbc039_3.cmdF(7).Enabled = False
                            f_cmbc039_3.cmdF(8).Enabled = False
                            f_cmbc039_3.cmdF(9).Enabled = False
                            f_cmbc039_3.cmdF(10).Enabled = False
                            f_cmbc039_3.cmdF(12).Enabled = False
                            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
                            f_cmbc039_3.lblMsg.Caption = GetMsgStr("ENSP4", "Y011")
                            rs.Close
                            Exit Function
                        End If
                    End If
                End If
                If IsNull(rs!INDTM) = True Then
                Else
                    .INDTM = rs!INDTM               ' �E�F�n�[�Z���^�[���ɓ���
                End If
                If IsNull(rs!BASKETID) = True Then
                    .BASKETID = vbNullString
                Else
                    .BASKETID = rs!BASKETID         ' �o�X�P�b�gID
                End If
                If IsNull(rs!SLOTNO) = True Then
                    .SLOTNO = vbNullString
                Else
                    .SLOTNO = rs!SLOTNO             ' �X���b�gNO
                End If
                If IsNull(rs!CURRWPCS) = True Then
                    .CURRWPCS = vbNullString
                Else
                    .CURRWPCS = rs!CURRWPCS         ' �E�F�n�[����
                End If
                If IsNull(rs!EXISTFLG) = True Then
                    .EXISTFLG = vbNullString
                Else
                    .EXISTFLG = rs!EXISTFLG         ' ���݃t���O
                End If
                If IsNull(rs!TOP_POS) = True Then
                    .TOP_POS = vbNullString
                Else
                    .TOP_POS = Int(CDbl(rs!TOP_POS) / 10 + 0.9 + dblEPS)        ' �u���b�N��TOP����̈ʒu  'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                End If
                If IsNull(rs!REJCAT) = True Then
                    .REJCAT = vbNullString
                Else
                    .REJCAT = rs!REJCAT             ' �������R
                End If
                If IsNull(rs!TXID) = True Then
                    .TXID = vbNullString
                Else
                    .TXID = rs!TXID                 ' �g�����U�N�V����ID
                End If
                If IsNull(rs!REGDATE) = True Then
                Else
                    .REGDATE = rs!REGDATE           ' �o�^���t
                End If
                If IsNull(rs!SUMMITSENDFLAG) = True Then
                    .SUMMITSENDFLAG = vbNullString
                Else
                    .SUMMITSENDFLAG = rs!SUMMITSENDFLAG ' SUMIT���M�t���O
                End If
                If IsNull(rs!SENDFLAG) = True Then
                    .SENDFLAG = vbNullString
                Else
                    .SENDFLAG = rs!SENDFLAG         ' ���M�t���O
                End If
                If IsNull(rs!SENDDATE) = True Then
                Else
                    .SENDDATE = rs!SENDDATE         ' ���M���t
                End If
                If IsNull(rs!HREJCODE) = True Then
                    .HREJCODE = vbNullString
                Else
                    .HREJCODE = rs!HREJCODE         ' �s�Ǘ��R�R�[�h
                End If
                If IsNull(rs!UPDPROC) = True Then
                    .UPDPROC = vbNullString
                Else
                    .UPDPROC = rs!UPDPROC           ' �X�V�H��
                End If
                If IsNull(rs!UPDDATE) = True Then
                Else
                    .UPDDATE = rs!UPDDATE           ' �X�V���t
                End If
                If IsNull(rs!MREVNUM) = True Then
                    .REVNUM = 0
                Else
                    .REVNUM = rs!MREVNUM             ' ���i�ԍ������ԍ�
                End If
                If IsNull(rs!Mfactory) = True Then
                    .factory = vbNullString
                Else
                    .factory = rs!Mfactory           ' �H��
                End If
                If IsNull(rs!Mopecond) = True Then
                    .opecond = vbNullString
                Else
                    .opecond = rs!Mopecond           ' ���Ə���
                End If
                If IsNull(rs!KANKBN) = True Then
                    .KANKBN = vbNullString
                Else
                    .KANKBN = rs!KANKBN             ' �����敪
                End If
                If IsNull(rs!NREJCODE) = True Then
                    .NREJCODE = vbNullString
                Else
                    .NREJCODE = rs!NREJCODE         ' �����ԓ����R�R�[�h
                End If
            End With
            intUCount = intUCount + 1
            rs.MoveNext
        Loop
    Next
    '�f���[�v�I��

    DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : DBDRV_BLOCKIDGET
'*
'*    �����T�v      : 1.�����w�� �u���b�N�h�c(SXL�����̃u���b�N�Ɍׂ�ꍇ�j���擾
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                            ,����
'*                    �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function DBDRV_BLOCKIDGET() As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim lngCnt      As Long
    Dim sDBName     As String
    Dim intUCount   As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_BLOCKIDGET"

    sDBName = "(V001)"

    ' SXLID�̎擾
    sSQL = "select"
    sSQL = sSQL & " CRYNUMCA,SXLIDCA"
    sSQL = sSQL & " from XSDCA "
    sSQL = sSQL & " where SXLIDCA ='" & tSXLID(0).SXLID & "'"
    sSQL = sSQL & "   and LIVKCA = '0'"
    sSQL = sSQL & "   ORDER BY CRYNUMCA,SXLIDCA"  'add 2003/04/30 hitec)matsumoto

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    '''���o���R�[�h�����݂Ȃ�ΊY��
    If rs.EOF Then
        DBDRV_BLOCKIDGET = FUNCTION_RETURN_FAILURE
        Exit Function
    Else
        '�f���o���R�[�h�����ׂĎ擾�i���[�v�j
        rs.MoveFirst  '�擪ں��ނɈړ�
        intUCount = 0
        Do While Not rs.EOF
            '�f�z��ɂ��̑g�ݍ��킹��ǉ�����
            If intUCount = 0 Then  '0�i�܂����b�g�����Ă��Ȃ���ԁj
                With tSXLID(intUCount)
                    .SXLID = rs.Fields("SXLIDCA")         'SXLIDCA
                    .LOTID = rs.Fields("CRYNUMCA")
                End With
            Else    '�Ώۃ��b�g��������������
                ReDim Preserve tSXLID(intUCount)  '�z��̍Ē�`
                With tSXLID(intUCount)
                    .SXLID = rs.Fields("SXLIDCA")         'SXLIDCA
                    .LOTID = rs.Fields("CRYNUMCA")
                End With
            End If
            intUCount = intUCount + 1
            rs.MoveNext
        Loop
        rs.Close
    End If

    DBDRV_BLOCKIDGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : DVDRV_KENSA_KOUMOKU
'*
'*    �����T�v      : 1.�����w���@�������ڂ��擾
'*                      (SXLID,BLOCKID���ő�A�ŏ��i�u���b�N�o�Ŕ���j�̃f�[�^���擾����)
'*    �p�����[�^    : �ϐ���        ,IO ,�^               ,����
'*                    tKensa        ,I  ,typ_XSDCW        ,�V�T���v���Ǘ��iSXL�j
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function DVDRV_KENSA_KOUMOKU(tKensa() As typ_XSDCW) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim sDBName     As String
    Dim intUCount   As Integer
    Dim tHIN        As tFullHinban
    Dim sOT1        As String
    Dim sOT2        As String
    Dim sMAI1       As String      '04/07/16
    Dim sMAI2       As String
    Dim rtn         As FUNCTION_RETURN
    Dim intIdx      As Integer
    Dim intCnt      As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DVDRV_KENSA_KOUMOKU"

    sDBName = "(V001)"
    intUCount = UBound(tSXLID)
    ReDim tKensa((intUCount * 2) + 1)            '�̈�Ē�`
    intIdx = 0

    '�f���[�v�J�n
    For i = 0 To intUCount
        ' SXLID�̎擾
        If Trim(tSXLID(i).SXLID) <> "" Then
            sSQL = "select "
            sSQL = sSQL & "SXLIDCW,"
            sSQL = sSQL & "SMPKBNCW,"
            sSQL = sSQL & "TBKBNCW,"
            sSQL = sSQL & "REPSMPLIDCW,"
            sSQL = sSQL & "XTALCW,"
            sSQL = sSQL & "INPOSCW,"
            sSQL = sSQL & "HINBCW,"
            sSQL = sSQL & "REVNUMCW,"
            sSQL = sSQL & "FACTORYCW,"
            sSQL = sSQL & "OPECW,"
            sSQL = sSQL & "KTKBNCW,"
            sSQL = sSQL & "SMCRYNUMCW,"
            sSQL = sSQL & "WFSMPLIDRSCW,"
            sSQL = sSQL & "WFSMPLIDRS1CW,"
            sSQL = sSQL & "WFSMPLIDRS2CW,"
            sSQL = sSQL & "WFINDRSCW,"
            sSQL = sSQL & "WFRESRS1CW,"
            sSQL = sSQL & "WFSMPLIDOICW,"
            sSQL = sSQL & "WFINDOICW,"
            sSQL = sSQL & "WFRESOICW,"
            sSQL = sSQL & "WFSMPLIDB1CW,"
            sSQL = sSQL & "WFINDB1CW,"
            sSQL = sSQL & "WFRESB1CW,"
            sSQL = sSQL & "WFSMPLIDB2CW,"
            sSQL = sSQL & "WFINDB2CW,"
            sSQL = sSQL & "WFRESB2CW,"
            sSQL = sSQL & "WFSMPLIDB3CW,"
            sSQL = sSQL & "WFINDB3CW,"
            sSQL = sSQL & "WFRESB3CW,"
            sSQL = sSQL & "WFSMPLIDL1CW,"
            sSQL = sSQL & "WFINDL1CW,"
            sSQL = sSQL & "WFRESL1CW,"
            sSQL = sSQL & "WFSMPLIDL2CW,"
            sSQL = sSQL & "WFINDL2CW,"
            sSQL = sSQL & "WFRESL2CW,"
            sSQL = sSQL & "WFSMPLIDL3CW,"
            sSQL = sSQL & "WFINDL3CW,"
            sSQL = sSQL & "WFRESL3CW,"
            sSQL = sSQL & "WFSMPLIDL4CW,"
            sSQL = sSQL & "WFINDL4CW,"
            sSQL = sSQL & "WFRESL4CW,"
            sSQL = sSQL & "WFSMPLIDDSCW,"
            sSQL = sSQL & "WFINDDSCW,"
            sSQL = sSQL & "WFRESDSCW,"
            sSQL = sSQL & "WFSMPLIDDZCW,"
            sSQL = sSQL & "WFINDDZCW,"
            sSQL = sSQL & "WFRESDZCW,"
            sSQL = sSQL & "WFSMPLIDSPCW,"
            sSQL = sSQL & "WFINDSPCW,"
            sSQL = sSQL & "WFRESSPCW,"
            sSQL = sSQL & "WFSMPLIDDO1CW,"
            sSQL = sSQL & "WFINDDO1CW,"
            sSQL = sSQL & "WFRESDO1CW,"
            sSQL = sSQL & "WFSMPLIDDO2CW,"
            sSQL = sSQL & "WFINDDO2CW,"
            sSQL = sSQL & "WFRESDO2CW,"
            sSQL = sSQL & "WFSMPLIDDO3CW,"
            sSQL = sSQL & "WFINDDO3CW,"
            sSQL = sSQL & "WFRESDO3CW,"
            sSQL = sSQL & "WFSMPLIDOT1CW,"
            sSQL = sSQL & "WFSMPLIDOT2CW,"
            sSQL = sSQL & "NVL(WFINDOT1CW,'0') as DOT1,"            ' ���FLG�iOT1)
            sSQL = sSQL & "NVL(WFRESOT1CW,'0') as SOT1,"            ' ����FLG�iOT1)
            sSQL = sSQL & "NVL(WFINDOT2CW,'0') as DOT2,"            ' ���FLG�iOT2)
            sSQL = sSQL & "NVL(WFRESOT2CW,'0') as SOT2,"            ' ����FLG�iOT2)
            sSQL = sSQL & "WFSMPLIDAOICW,"
            sSQL = sSQL & "WFINDAOICW,"
            sSQL = sSQL & "WFRESAOICW,"
            sSQL = sSQL & "SMPLNUMCW,"
            sSQL = sSQL & "SMPLPATCW,"
            sSQL = sSQL & "TSTAFFCW,"
            sSQL = sSQL & "TDAYCW,"
            sSQL = sSQL & "KSTAFFCW,"
            sSQL = sSQL & "KDAYCW,"
            sSQL = sSQL & "SNDKCW,"
            sSQL = sSQL & "SNDDAYCW,"
            sSQL = sSQL & "WFSMPLIDGDCW,"     '' GD�ǉ��@05/02/17 ooba START ===========>
            sSQL = sSQL & "WFINDGDCW,"
            sSQL = sSQL & "WFRESGDCW,"
            sSQL = sSQL & "WFHSGDCW"          '' GD�ǉ��@05/02/17 ooba END =============>
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
            sSQL = sSQL & ",EPSMPLIDB1CW, "
            sSQL = sSQL & "EPINDB1CW, "
            sSQL = sSQL & "EPRESB1CW, "
            sSQL = sSQL & "EPSMPLIDB2CW, "
            sSQL = sSQL & "EPINDB2CW, "
            sSQL = sSQL & "EPRESB2CW, "
            sSQL = sSQL & "EPSMPLIDB3CW, "
            sSQL = sSQL & "EPINDB3CW, "
            sSQL = sSQL & "EPRESB3CW, "
            sSQL = sSQL & "EPSMPLIDL1CW, "
            sSQL = sSQL & "EPINDL1CW, "
            sSQL = sSQL & "EPRESL1CW, "
            sSQL = sSQL & "EPSMPLIDL2CW, "
            sSQL = sSQL & "EPINDL2CW, "
            sSQL = sSQL & "EPRESL2CW, "
            sSQL = sSQL & "EPSMPLIDL3CW, "
            sSQL = sSQL & "EPINDL3CW, "
            sSQL = sSQL & "EPRESL3CW "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
            sSQL = sSQL & " from XSDCW "
            sSQL = sSQL & " where SXLIDCW ='" & tSXLID(i).SXLID & "'"
            sSQL = sSQL & "   and LIVKCW  ='0'"                           ' �����敪�͕K���m�F���鎖
            sSQL = sSQL & " order by INPOSCW"

            Debug.Print sSQL
            Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

            '''���o���R�[�h�����݂Ȃ�ΊY��
            If Not rs.EOF Then
                intCnt = 0
                Do While Not rs.EOF
                    intCnt = intCnt + 1
                    ' �R���ڈȍ~�����݂���ꍇ�G���[
                    If intCnt > 2 Then
                        Exit Do
                    End If

                    With tKensa(intIdx)
                        .SXLIDCW = rs("SXLIDCW")
                        .SMPKBNCW = rs!SMPKBNCW
                        .TBKBNCW = rs!TBKBNCW
                        .REPSMPLIDCW = rs!REPSMPLIDCW
                        .XTALCW = rs!XTALCW
                        .INPOSCW = rs!INPOSCW
                        .HINBCW = rs!HINBCW
                        .REVNUMCW = rs!REVNUMCW
                        .FACTORYCW = rs!FACTORYCW
                        .OPECW = rs!OPECW
                        .KTKBNCW = rs!KTKBNCW
                        .SMCRYNUMCW = rs!SMCRYNUMCW
                        If IsNull(rs!WFSMPLIDRSCW) Then
                            .WFSMPLIDRSCW = ""
                        Else
                            .WFSMPLIDRSCW = rs!WFSMPLIDRSCW
                        End If
                        If IsNull(rs("WFSMPLIDRS1CW")) Then
                            .WFSMPLIDRS1CW = ""
                        Else
                            .WFSMPLIDRS1CW = rs("WFSMPLIDRS1CW")
                        End If
                        If IsNull(rs!WFSMPLIDRS2CW) Then
                            .WFSMPLIDRS2CW = ""
                        Else
                            .WFSMPLIDRS2CW = rs!WFSMPLIDRS2CW
                        End If
                        If IsNull(rs!WFINDRSCW) Then
                            .WFINDRSCW = ""
                        Else
                            .WFINDRSCW = rs!WFINDRSCW
                        End If
                        If IsNull(rs!WFRESRS1CW) Then
                            .WFRESRS1CW = ""
                        Else
                            .WFRESRS1CW = rs!WFRESRS1CW
                        End If
                        If IsNull(rs!WFSMPLIDOICW) Then
                            .WFSMPLIDOICW = ""
                        Else
                            .WFSMPLIDOICW = rs!WFSMPLIDOICW
                        End If
                        If IsNull(rs!WFINDOICW) Then
                            .WFINDOICW = ""
                        Else
                            .WFINDOICW = rs!WFINDOICW
                        End If
                        If IsNull(rs!WFRESOICW) Then
                            .WFRESOICW = ""
                        Else
                            .WFRESOICW = rs!WFRESOICW
                        End If
                        If IsNull(rs!WFSMPLIDB1CW) Then
                            .WFSMPLIDB1CW = ""
                        Else
                            .WFSMPLIDB1CW = rs!WFSMPLIDB1CW
                        End If
                        If IsNull(rs!WFINDB1CW) Then
                            .WFINDB1CW = ""
                        Else
                            .WFINDB1CW = rs!WFINDB1CW
                        End If
                        If IsNull(rs!WFRESB1CW) Then
                            .WFRESB1CW = ""
                        Else
                            .WFRESB1CW = rs!WFRESB1CW
                        End If
                        If IsNull(rs!WFSMPLIDB2CW) Then
                            .WFSMPLIDB2CW = ""
                        Else
                            .WFSMPLIDB2CW = rs!WFSMPLIDB2CW
                        End If
                        If IsNull(rs!WFINDB2CW) Then
                            .WFINDB2CW = ""
                        Else
                            .WFINDB2CW = rs!WFINDB2CW
                        End If
                        If IsNull(rs!WFRESB2CW) Then
                            .WFRESB2CW = ""
                        Else
                            .WFRESB2CW = rs!WFRESB2CW
                        End If
                        If IsNull(rs!WFSMPLIDB3CW) Then
                            .WFSMPLIDB3CW = ""
                        Else
                            .WFSMPLIDB3CW = rs!WFSMPLIDB3CW
                        End If
                        If IsNull(rs!WFINDB3CW) Then
                            .WFINDB3CW = ""
                        Else
                            .WFINDB3CW = rs!WFINDB3CW
                        End If
                        If IsNull(rs!WFRESB3CW) Then
                            .WFRESB3CW = ""
                        Else
                            .WFRESB3CW = rs!WFRESB3CW
                        End If
                        If IsNull(rs!WFSMPLIDL1CW) Then
                            .WFSMPLIDL1CW = ""
                        Else
                            .WFSMPLIDL1CW = rs!WFSMPLIDL1CW
                        End If
                        If IsNull(rs!WFINDL1CW) Then
                            .WFINDL1CW = ""
                        Else
                            .WFINDL1CW = rs!WFINDL1CW
                        End If
                        If IsNull(rs!WFRESL1CW) Then
                            .WFRESL1CW = ""
                        Else
                            .WFRESL1CW = rs!WFRESL1CW
                        End If
                        If IsNull(rs!WFSMPLIDL2CW) Then
                            .WFSMPLIDL2CW = ""
                        Else
                            .WFSMPLIDL2CW = rs!WFSMPLIDL2CW
                        End If
                        If IsNull(rs!WFINDL2CW) Then
                            .WFINDL2CW = ""
                        Else
                            .WFINDL2CW = rs!WFINDL2CW
                        End If
                        If IsNull(rs!WFRESL2CW) Then
                            .WFRESL2CW = ""
                        Else
                            .WFRESL2CW = rs!WFRESL2CW
                        End If
                        If IsNull(rs!WFSMPLIDL3CW) Then
                            .WFSMPLIDL3CW = ""
                        Else
                            .WFSMPLIDL3CW = rs!WFSMPLIDL3CW
                        End If
                        If IsNull(rs!WFINDL3CW) Then
                            .WFINDL3CW = ""
                        Else
                            .WFINDL3CW = rs!WFINDL3CW
                        End If
                        If IsNull(rs!WFRESL3CW) Then
                            .WFRESL3CW = ""
                        Else
                            .WFRESL3CW = rs!WFRESL3CW
                        End If
                        If IsNull(rs!WFSMPLIDL4CW) Then
                            .WFSMPLIDL4CW = ""
                        Else
                            .WFSMPLIDL4CW = rs!WFSMPLIDL4CW
                        End If
                        If IsNull(rs!WFINDL4CW) Then
                            .WFINDL4CW = ""
                        Else
                            .WFINDL4CW = rs!WFINDL4CW
                        End If
                        If IsNull(rs!WFRESL4CW) Then
                            .WFRESL4CW = ""
                        Else
                            .WFRESL4CW = rs!WFRESL4CW
                        End If
                        If IsNull(rs!WFSMPLIDDSCW) Then
                            .WFSMPLIDDSCW = ""
                        Else
                            .WFSMPLIDDSCW = rs!WFSMPLIDDSCW
                        End If
                        If IsNull(rs!WFINDDSCW) Then
                            .WFINDDSCW = ""
                        Else
                            .WFINDDSCW = rs!WFINDDSCW
                        End If
                        If IsNull(rs!WFRESDSCW) Then
                            .WFRESDSCW = ""
                        Else
                            .WFRESDSCW = rs!WFRESDSCW
                        End If
                        If IsNull(rs!WFSMPLIDDZCW) Then
                            .WFSMPLIDDZCW = ""
                        Else
                            .WFSMPLIDDZCW = rs!WFSMPLIDDZCW
                        End If
                        If IsNull(rs!WFINDDZCW) Then
                            .WFINDDZCW = ""
                        Else
                            .WFINDDZCW = rs!WFINDDZCW
                        End If
                        If IsNull(rs!WFRESDZCW) Then
                            .WFRESDZCW = ""
                        Else
                            .WFRESDZCW = rs!WFRESDZCW
                        End If
                        If IsNull(rs!WFSMPLIDSPCW) Then
                            .WFSMPLIDSPCW = ""
                        Else
                            .WFSMPLIDSPCW = rs!WFSMPLIDSPCW
                        End If
                        If IsNull(rs!WFINDSPCW) Then
                            .WFINDSPCW = ""
                        Else
                            .WFINDSPCW = rs!WFINDSPCW
                        End If
                        If IsNull(rs!WFRESSPCW) Then
                            .WFRESSPCW = ""
                        Else
                            .WFRESSPCW = rs!WFRESSPCW
                        End If
                        If IsNull(rs!WFSMPLIDDO1CW) Then
                            .WFSMPLIDDO1CW = ""
                        Else
                            .WFSMPLIDDO1CW = rs!WFSMPLIDDO1CW
                        End If
                        If IsNull(rs!WFINDDO1CW) Then
                            .WFINDDO1CW = ""
                        Else
                            .WFINDDO1CW = rs!WFINDDO1CW
                        End If
                        If IsNull(rs!WFRESDO1CW) Then
                            .WFRESDO1CW = ""
                        Else
                            .WFRESDO1CW = rs!WFRESDO1CW
                        End If
                        If IsNull(rs!WFSMPLIDDO2CW) Then
                            .WFSMPLIDDO2CW = ""
                        Else
                            .WFSMPLIDDO2CW = rs!WFSMPLIDDO2CW
                        End If
                        If IsNull(rs!WFINDDO2CW) Then
                            .WFINDDO2CW = ""
                        Else
                            .WFINDDO2CW = rs!WFINDDO2CW
                        End If
                        If IsNull(rs!WFRESDO2CW) Then
                            .WFRESDO2CW = ""
                        Else
                            .WFRESDO2CW = rs!WFRESDO2CW
                        End If
                        If IsNull(rs!WFSMPLIDDO3CW) Then
                            .WFSMPLIDDO3CW = ""
                        Else
                            .WFSMPLIDDO3CW = rs!WFSMPLIDDO3CW
                        End If
                        If IsNull(rs!WFINDDO3CW) Then
                            .WFINDDO3CW = ""
                        Else
                            .WFINDDO3CW = rs!WFINDDO3CW
                        End If
                        If IsNull(rs!WFRESDO3CW) Then
                            .WFRESDO3CW = ""
                        Else
                            .WFRESDO3CW = rs!WFRESDO3CW
                        End If
                        If IsNull(rs!WFSMPLIDOT1CW) Then
                            .WFSMPLIDOT1CW = ""
                        Else
                            .WFSMPLIDOT1CW = rs!WFSMPLIDOT1CW
                        End If
                        If IsNull(rs!sOT1) Then
                            .WFRESOT1CW = ""
                        Else
                            .WFRESOT1CW = rs!sOT1
                        End If
                        If IsNull(rs!WFSMPLIDOT2CW) Then
                            .WFSMPLIDOT2CW = ""
                        Else
                            .WFSMPLIDOT2CW = rs!WFSMPLIDOT2CW
                        End If
                        If IsNull(rs!sOT2) Then
                            .WFRESOT2CW = ""
                        Else
                            .WFRESOT2CW = rs!sOT2
                        End If

                        tHIN.hinban = .HINBCW
                        tHIN.factory = .FACTORYCW
                        tHIN.mnorevno = .REVNUMCW
                        tHIN.opecond = .OPECW

                        rtn = scmzc_getE036(tHIN, sOT1, sOT2, sMAI1, sMAI2)
                        If rtn = FUNCTION_RETURN_FAILURE Then
                            rs.Close
                            DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
                            GoTo proc_exit
                        End If
                        If sOT1 = "1" Then
                            .WFINDOT1CW = rs!DOT1
                        Else
                            .WFINDOT1CW = 0
                        End If
                        If sOT2 = "1" Then
                            .WFINDOT2CW = rs!DOT2
                        Else
                            .WFINDOT2CW = 0
                        End If

                        If IsNull(rs!WFSMPLIDAOICW) Then
                            .WFSMPLIDAOICW = ""
                        Else
                            .WFSMPLIDAOICW = rs!WFSMPLIDAOICW
                        End If
                        If IsNull(rs!WFINDAOICW) Then
                            .WFINDAOICW = ""
                        Else
                            .WFINDAOICW = rs!WFINDAOICW
                        End If
                        If IsNull(rs!WFRESAOICW) Then
                            .WFRESAOICW = ""
                        Else
                            .WFRESAOICW = rs!WFRESAOICW
                        End If
                        If IsNull(rs!SMPLNUMCW) Then
                            .SMPLNUMCW = 0
                        Else
                            .SMPLNUMCW = rs!SMPLNUMCW
                        End If
                        If IsNull(rs!SMPLPATCW) Then
                            .SMPLPATCW = ""
                        Else
                            .SMPLPATCW = rs!SMPLPATCW
                        End If
                        If IsNull(rs!TSTAFFCW) Then
                            .TSTAFFCW = ""
                        Else
                            .TSTAFFCW = rs!TSTAFFCW
                        End If
                        If IsNull(rs!TDAYCW) Then
                            .TDAYCW = "2003/10/3"
                        Else
                            .TDAYCW = rs!TDAYCW
                        End If
                        If IsNull(rs!KSTAFFCW) Then
                            .KSTAFFCW = ""
                        Else
                            .KSTAFFCW = rs!KSTAFFCW
                        End If
                        If IsNull(rs!KDAYCW) Then
                            .KDAYCW = "2003/10/3"
                        Else
                            .KDAYCW = rs!KDAYCW
                        End If
                        If IsNull(rs!SNDKCW) Then
                            .SNDKCW = ""
                        Else
                            .SNDKCW = rs!SNDKCW
                        End If
                        If IsNull(rs!SNDDAYCW) Then
                            .SNDDAYCW = "2003/10/3"
                        Else
                            .SNDDAYCW = rs!SNDDAYCW
                        End If
                        '' GD�ǉ��@05/02/17 ooba START ==================>
                        If IsNull(rs!WFSMPLIDGDCW) Then
                            .WFSMPLIDGDCW = ""
                        Else
                            .WFSMPLIDGDCW = rs!WFSMPLIDGDCW
                        End If
                        If IsNull(rs!WFINDGDCW) Then
                            .WFINDGDCW = ""
                        Else
                            .WFINDGDCW = rs!WFINDGDCW
                        End If
                        If IsNull(rs!WFRESGDCW) Then
                            .WFRESGDCW = ""
                        Else
                            .WFRESGDCW = rs!WFRESGDCW
                        End If
                        If IsNull(rs!WFHSGDCW) Then
                            .WFHSGDCW = ""
                        Else
                            .WFHSGDCW = rs!WFHSGDCW
                        End If
                        '' GD�ǉ��@05/02/17 ooba END ====================>

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                        ' BMD1E
                        If IsNull(rs!EPSMPLIDB1CW) Then
                            .EPSMPLIDB1CW = ""
                        Else
                            .EPSMPLIDB1CW = rs!EPSMPLIDB1CW
                        End If
                        If IsNull(rs!EPINDB1CW) Then
                            .EPINDB1CW = ""
                        Else
                            .EPINDB1CW = rs!EPINDB1CW
                        End If
                        If IsNull(rs!EPRESB1CW) Then
                            .EPRESB1CW = ""
                        Else
                            .EPRESB1CW = rs!EPRESB1CW
                        End If
                        ' BMD2E
                        If IsNull(rs!EPSMPLIDB2CW) Then
                            .EPSMPLIDB2CW = ""
                        Else
                            .EPSMPLIDB2CW = rs!EPSMPLIDB2CW
                        End If
                        If IsNull(rs!EPINDB2CW) Then
                            .EPINDB2CW = ""
                        Else
                            .EPINDB2CW = rs!EPINDB2CW
                        End If
                        If IsNull(rs!EPRESB2CW) Then
                            .EPRESB2CW = ""
                        Else
                            .EPRESB2CW = rs!EPRESB2CW
                        End If
                        ' BMD3E
                        If IsNull(rs!EPSMPLIDB3CW) Then
                            .EPSMPLIDB3CW = ""
                        Else
                            .EPSMPLIDB3CW = rs!EPSMPLIDB3CW
                        End If
                        If IsNull(rs!EPINDB3CW) Then
                            .EPINDB3CW = ""
                        Else
                            .EPINDB3CW = rs!EPINDB3CW
                        End If
                        If IsNull(rs!EPRESB3CW) Then
                            .EPRESB3CW = ""
                        Else
                            .EPRESB3CW = rs!EPRESB3CW
                        End If
                        ' OSF1E
                        If IsNull(rs!EPSMPLIDL1CW) Then
                            .EPSMPLIDL1CW = ""
                        Else
                            .EPSMPLIDL1CW = rs!EPSMPLIDL1CW
                        End If
                        If IsNull(rs!EPINDL1CW) Then
                            .EPINDL1CW = ""
                        Else
                            .EPINDL1CW = rs!EPINDL1CW
                        End If
                        If IsNull(rs!EPRESL1CW) Then
                            .EPRESL1CW = ""
                        Else
                            .EPRESL1CW = rs!EPRESL1CW
                        End If
                        ' OSF2E
                        If IsNull(rs!EPSMPLIDL2CW) Then
                            .EPSMPLIDL2CW = ""
                        Else
                            .EPSMPLIDL2CW = rs!EPSMPLIDL2CW
                        End If
                        If IsNull(rs!EPINDL2CW) Then
                            .EPINDL2CW = ""
                        Else
                            .EPINDL2CW = rs!EPINDL2CW
                        End If
                        If IsNull(rs!EPRESL2CW) Then
                            .EPRESL2CW = ""
                        Else
                            .EPRESL2CW = rs!EPRESL2CW
                        End If
                        ' OSF3E
                        If IsNull(rs!EPSMPLIDL3CW) Then
                            .EPSMPLIDL3CW = ""
                        Else
                            .EPSMPLIDL3CW = rs!EPSMPLIDL3CW
                        End If
                        If IsNull(rs!EPINDL3CW) Then
                            .EPINDL3CW = ""
                        Else
                            .EPINDL3CW = rs!EPINDL3CW
                        End If
                        If IsNull(rs!EPRESL3CW) Then
                            .EPRESL3CW = ""
                        Else
                            .EPRESL3CW = rs!EPRESL3CW
                        End If
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                    End With

                    intIdx = intIdx + 1
                    rs.MoveNext
                Loop
                rs.Close

                ' �擾�������Q���łȂ��ꍇ�G���[
                If intCnt <> 2 Then
                    f_cmbc039_3.lblMsg.Caption = GetMsgStr("ENSP2")
                    DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
            Else
                 f_cmbc039_3.lblMsg.Caption = GetMsgStr("ENSP2")
                DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        End If
    Next i
    '�f���[�v�I��

    DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    GoTo proc_exit
End Function

'*****************************************************************************************************************
'*    �֐���        : DBDRV_GET_WFMAP
'*
'*    �����T�v      : 1.�����w�� ���͂����u���b�N�o����A�Y���v�e������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^      ,����
'*                    sBlkId        ,I  ,String  ,�u���b�NID
'*                    sSXLID        ,I  ,String  ,SXL-ID
'*                    iBlkP         ,I  ,Integer ,�u���b�NP
'*                    sBlkP         ,I  ,Variant ,Spread�ɋL�ڂ���Ă���u���b�NP
'*                    sKessyoP      ,O  ,Variant ,�_���������ʒu
'*                    sNextIngotP   ,O  ,String  ,�������ʒu
'*                    sBlkSeq       ,O  ,Variant ,�u���b�N���A��
'*                    sBlkSeq2      ,O  ,Variant ,�u���b�N���A��
'*                    sSmpId1       ,O  ,Variant ,�T���v��ID
'*                    sSmpId2       ,O  ,Variant ,�T���v��ID
'*                    iNextBlkP     ,O  ,Integer ,���u���b�NP
'*                    vWfNum        ,O  ,Variant ,Wafer��
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************************************
Public Function DBDRV_GET_WFMAP(ByVal sBlkId As String, ByVal sSXLID As String, ByVal iBlkP As Integer, _
                                ByRef sBlkP As Variant, ByRef sKessyoP As Variant, ByRef sNextIngotP As String, _
                                ByRef sBlkSeq As Variant, ByRef sBlkSeq2 As Variant, ByRef sSmpId1 As Variant, _
                                ByRef sSmpId2 As Variant, ByRef iNextBlkP As Integer, ByRef vWfNum As Variant) _
                                    As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim lngCnt      As Long
    Dim sDBName     As String
    Dim intLoopCnt  As Integer
    Dim dblChkBlkP  As Double
    Dim intChkBlkP  As Integer
    Dim intTopPos   As Integer
    Dim sAddSmpId1  As String
    Dim sAddSmpId2  As String
    Dim vBlkId      As Variant
    Dim intBlkflg   As Integer
    Dim dblWFLen    As Double
    Dim iRtn        As FUNCTION_RETURN
    Dim intSearchWf As Integer
    Dim dblEPS      As Double

    dblEPS = 0.000001        '�Â̐ݒ� 'add 2003/06/13 hitec)matsumoto

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_GET_WFMAP"

    sDBName = "(Y011)"
    i = 0

    sSQL = "select "
    sSQL = sSQL & "LOTID,"                ' �u���b�NID"
    sSQL = sSQL & "MSXLID,"                ' SXLID"
    sSQL = sSQL & "blockseq,"             ' �u���b�N���A��"
    sSQL = sSQL & "WFSTA,"                ' WF���"
    sSQL = sSQL & "RTOP_POS,"             ' �_���u���b�N���ʒu"
    sSQL = sSQL & "RITOP_POS,"            ' �_���������ʒu"
    sSQL = sSQL & "MSMPLEID,"              ' �����ʒu"
    sSQL = sSQL & "SHAFLAG,"              ' �T���v���t���O"
    sSQL = sSQL & "TOP_POS"               ' �u���b�N���ʒu
    sSQL = sSQL & " from TBCMY011 "
    sSQL = sSQL & " where MSXLID ='" & sSXLID & "'"
    sSQL = sSQL & "   AND LOTID ='" & sBlkId & "'"
    sSQL = sSQL & "   AND TO_NUMBER(WFSTA) <= 1"  'del 2003/04/28 hitec)matsumoto ��Ԃ����Ȃ�
    sSQL = sSQL & " ORDER BY BLOCKSEQ ASC"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    intLoopCnt = 0
    vWfNum = 0
    Do While Not rs.EOF
        If IsNull(rs.Fields("RTOP_POS")) = True Then
            dblChkBlkP = 0
        Else
            dblChkBlkP = CDbl(rs.Fields("RTOP_POS"))
        End If
        If (iBlkP < dblChkBlkP) And (dblChkBlkP <= iNextBlkP) Then
            vWfNum = CInt(vWfNum) + 1
        End If
        rs.MoveNext
    Loop
    rs.Close

    sSQL = "select "
    sSQL = sSQL & "LOTID,"                ' �u���b�NID"
    sSQL = sSQL & "MSXLID,"                ' SXLID"
    sSQL = sSQL & "blockseq,"             ' �u���b�N���A��"
    sSQL = sSQL & "WFSTA,"                ' WF���"
    sSQL = sSQL & "RTOP_POS,"             ' �_���u���b�N���ʒu"
    sSQL = sSQL & "RITOP_POS,"            ' �_���������ʒu"
    sSQL = sSQL & "MSMPLEID,"              ' �����ʒu"
    sSQL = sSQL & "SHAFLAG,"              ' �T���v���t���O"
    sSQL = sSQL & "TOP_POS"               ' �u���b�N���ʒu
    sSQL = sSQL & " from TBCMY011 "
    sSQL = sSQL & " where MSXLID ='" & sSXLID & "'"
    sSQL = sSQL & "   AND LOTID ='" & sBlkId & "'"
    sSQL = sSQL & " ORDER BY BLOCKSEQ ASC"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    intLoopCnt = 0
    rs.MoveFirst
    Do While Not rs.EOF
        Select Case Right(sSmpId1, 1)
            Case "T"
                rs.MoveFirst
                'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)   'add 2003/05/05 hitec)matsumoto
                If CStr(rs.Fields("WFSTA")) = "4" Then
                    rs.Close
                    DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + dblEPS)  'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                End If
                If IsNull(rs.Fields("RITOP_POS")) = False Then
                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS)  'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                End If
                sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + dblEPS)  '�؂�̂� 'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "T"
                Exit Do
            Case "U"
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    dblChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                End If
                If dblChkBlkP > iBlkP Or dblChkBlkP = iBlkP Then
                    If dblChkBlkP > iBlkP Then
                        rs.MovePrevious
                        If IsNull(rs.Fields("BLOCKSEQ")) = True Then    'add 2003/04/28 hitec)matsumoto  NULL�̏ꍇ�i�Y��WF�����j�́A���Ɍ�������
                            Do
                                rs.MoveNext
                                If IsNull(rs.Fields("RTOP_POS")) = False Then
                                    Exit Do
                                End If
                            Loop
                        End If
                        'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)   'add 2003/05/05 hitec)matsumoto
                        If CStr(rs.Fields("WFSTA")) = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                        intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + dblEPS) '�؂�グ    'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "U"
                        rs.MoveNext
                    ElseIf dblChkBlkP = iBlkP Then
                        'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)   'add 2003/05/05 hitec)matsumoto
                        If CStr(rs.Fields("WFSTA")) = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                        intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + dblEPS) '�؂�グ    'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "U"
                        rs.MoveNext
                    End If

                    If Not rs.EOF Then
                        If sSmpId2 <> vbNullString Then 'D�̃T���v�����쐬
                            '0�ȊO��0.1mm�����Đ؎̂�(WF����:D�͊Y���ʒu���܂܂��ɉ����������) 08/11/06 ooba
                            If rs.Fields("TOP_POS") > 0 Then
                                intTopPos = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + dblEPS)   '0.1mm�����Đ؎̂�
                            Else
                                intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + dblEPS)  '�؂�̂� 'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                            End If
                            sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "D"
                        End If
                        sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                        sNextIngotP = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS)   'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                        Exit Do
                    Else
                        '���݂̃u���b�NID�̎��̃u���b�NID���擾
                        With f_cmbc039_3.sprExamine
                            intBlkflg = 0
                            For i = 1 To .MaxRows
                                .GetText 1, i, vBlkId
                                If vBlkId <> "" And Len(vBlkId) <> 1 Then    '03/05/31
                                    If intBlkflg = 1 Then
                                        sBlkId = left(sBlkId, 9) & CStr(vBlkId) '����BLID�擾
                                        Exit For
                                    ElseIf Right(sBlkId, 3) = vBlkId Then
                                        intBlkflg = 1
                                    End If
                                End If
                            Next i
                        End With
                        rs.Close

                        sSQL = "select "
                        sSQL = sSQL & "LOTID,"                ' �u���b�NID"
                        sSQL = sSQL & "MSXLID,"                ' SXLID"
                        sSQL = sSQL & "blockseq,"             ' �u���b�N���A��"
                        sSQL = sSQL & "WFSTA,"                ' WF���"
                        sSQL = sSQL & "RTOP_POS,"             ' �_���u���b�N���ʒu"
                        sSQL = sSQL & "RITOP_POS,"            ' �_���������ʒu"
                        sSQL = sSQL & "MSMPLEID,"              ' �����ʒu"
                        sSQL = sSQL & "SHAFLAG,"              ' �T���v���t���O"
                        sSQL = sSQL & "TOP_POS"               ' �u���b�N���ʒu
                        sSQL = sSQL & " from TBCMY011 "
                        sSQL = sSQL & " where MSXLID ='" & sSXLID & "'"
                        sSQL = sSQL & "   AND LOTID ='" & sBlkId & "'"
                        sSQL = sSQL & " ORDER BY BLOCKSEQ ASC"

                        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

                        intLoopCnt = 0
                        rs.MoveFirst
                        'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)   'add 2003/05/05 hitec)matsumoto
                        If CStr(rs.Fields("WFSTA")) = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        If IsNull(rs.Fields("RTOP_POS")) = False Then
                            'WF�ꖇ�̒����擾                                   '2003/04/25 hitec)okazaki
                            iRtn = DBDRV_WFLENGET(sBlkId, dblWFLen)
                            '�u���b�N�擪�̕\���ʒu��WF�ꖇ�̒���������������   '2003/04/25 hitec)okazaki
                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) - dblWFLen + dblEPS)   'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                        End If
                        If IsNull(rs.Fields("RITOP_POS")) = False Then
                            sNextIngotP = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS) 'upd 2003/06/13 hitec)matsumoto
                        End If
                        If sSmpId2 <> vbNullString Then 'D�̃T���v�����쐬
                            '0�ȊO��0.1mm�����Đ؎̂�(WF����:D�͊Y���ʒu���܂܂��ɉ����������) 08/11/06 ooba
                            If rs.Fields("TOP_POS") > 0 Then
                                intTopPos = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + dblEPS)   '0.1mm�����Đ؎̂�
                            Else
                                intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + dblEPS)  '�؂�̂� 'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                            End If
                            sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "D"
                        End If
                        sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                        Exit Do
                    End If
                End If
            Case "D"
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    dblChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                End If
                If dblChkBlkP > iBlkP Then
                    sNextIngotP = Int(CDbl(rs.Fields("RITOP_POS")) + dblEPS)   'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                    sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                    '0�ȊO��0.1mm�����Đ؎̂�(WF����:D�͊Y���ʒu���܂܂��ɉ����������) 08/11/06 ooba
                    If rs.Fields("TOP_POS") > 0 Then
                        intTopPos = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + dblEPS)   '0.1mm�����Đ؎̂�
                    Else
                        intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + dblEPS)  '�؂�̂� 'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                    End If
                    sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "D"   'D�Ȃ̂�sAddSmpId2�ɓ����
                    rs.MovePrevious
                    'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)   'add 2003/05/05 hitec)matsumoto
                    If CStr(rs.Fields("WFSTA")) = "4" Then
                        rs.Close
                        DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                    If sSmpId2 <> vbNullString Then 'U�̃T���v�����쐬
                        intTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + dblEPS)   '�؂�グ  'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "U"   '"U"�Ȃ̂�sAddSmpId1�ɓ����
                    End If
                    sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                    sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                    Exit Do
                End If
            Case "B"
                rs.MoveLast
                'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)   'add 2003/05/05 hitec)matsumoto
                If CStr(rs.Fields("WFSTA")) = "4" Then
                    rs.Close
                    DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                End If
                If IsNull(rs.Fields("RITOP_POS")) = False Then
                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + dblEPS)    'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                End If
                sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                intTopPos = Int(CDbl(rs.Fields("TOP_POS")) + 0.9 + dblEPS) '�؂�グ 'add 2003/06/13 hitec)matsumoto [+ dblEPS]�ǉ�
                sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(intTopPos), "000") & "B"
                Exit Do
        End Select
        rs.MoveNext
    Loop
    sSmpId1 = sAddSmpId1
    sSmpId2 = sAddSmpId2
    rs.Close

    DBDRV_GET_WFMAP = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
'    gErr.HandleError
    DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : DBDRV_UPD_WFMap
'*
'*    �����T�v      : 1.WF�}�b�v�e�[�u���X�V
'*                    (WF�}�b�v�e�[�u��(TBCMY011)���X�V����)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                            ,����
'*                    SXL           ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL�Ǘ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function DBDRV_UPD_WFMap() As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim i               As Long
    Dim lngLoopCnt      As Long
    Dim sDBName         As String
    Dim intUCount       As Integer
    Dim dtmNowtime      As Date
    Dim vGetMaxSeq      As Variant
    Dim sGetSXLid       As String
    Dim intNowIngotPos  As Integer
    Dim intGetSmplLoop  As Integer
    Dim intFromBlkSeq   As Integer
    Dim intToBlkSeq     As Integer
    Dim intNextLoopCnt  As Integer
    Dim vGetSample      As Variant
    Dim m               As Integer
    Dim intGetNextSeq   As Integer
    Dim vGetHinban      As Variant
    Dim intAllScrapCnt  As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_UPD_WFMap"

    sDBName = "(Y011)"
    With f_cmbc039_3.sprExamine
        m = .MaxRows
        For lngLoopCnt = 1 To m Step 2
            intFromBlkSeq = gtSprWfMap(lngLoopCnt).BLOCKSEQ '�u���b�NSEQ���擾
            intToBlkSeq = gtSprWfMap(lngLoopCnt + 1).BLOCKSEQ '�u���b�NSEQ���擾
            .row = lngLoopCnt
            .col = 10
            If (Len(Trim(.text)) > 0) Or (gtSprWfMap(lngLoopCnt).hinban = "Z") Then    '�T���v���s�̏ꍇ
                If (gtSprWfMap(lngLoopCnt).hinban = "Z") And (gtSprWfMap(lngLoopCnt - 1).hinban = "Z") Then 'Z���A�����Ă������ꍇ�́A�㑤��SXLID������
                    'SXLID�͂��Ȃ�
                    If CheckGetSampleID(lngLoopCnt - 1) = True Then
                        .GetText 5, lngLoopCnt, gtSprWfMap(lngLoopCnt).KESSYOUP                 '2003/06/01 add (��SXL�̏I�������ʒu�Ɖ�SXL�̊J�n�ʒu���قȂ�P�[�X�����݂��邽�߁j
                        sGetSXLid = Mid(gtSprWfMap(lngLoopCnt).LOTID, 1, 10) & GetWafPos(CInt(gtSprWfMap(lngLoopCnt).KESSYOUP))
                    End If
                Else
                    If lngLoopCnt = 1 Then
                        sGetSXLid = tblSXL.SXLID 'upd 2003/05/19 hitec)matsumoto �ʒu����SXLID�����Ȃ�
                    Else
                        .GetText 5, lngLoopCnt, gtSprWfMap(lngLoopCnt).KESSYOUP                 '2003/06/01 add (��SXL�̏I�������ʒu�Ɖ�SXL�̊J�n�ʒu���قȂ�P�[�X�����݂��邽�߁j
                        sGetSXLid = Mid(gtSprWfMap(lngLoopCnt).LOTID, 1, 10) & GetWafPos(CInt(gtSprWfMap(lngLoopCnt).KESSYOUP))
                    End If
                End If
            End If
            If gtSprWfMap(lngLoopCnt).hinban = "Z" Then
                sSQL = "UPDATE TBCMY011 SET"
                dtmNowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")

                '�擪�T���v��ID�͕ς��Ȃ�����SXLID�������̕����ō��ꂽ�܂܂Ƃ���@2003/04/22
                If lngLoopCnt = 1 Then
                    intNowIngotPos = SIngotP
                Else
                    intNowIngotPos = gtSprWfMap(lngLoopCnt).KESSYOUP
                End If

                sSQL = sSQL & " MSXLID = '" & sGetSXLid & "'"

                sSQL = sSQL & ",UPDPROC = 'CW760'"             ' �X�V�H��
                sSQL = sSQL & ",UPDDATE = sysdate"    'upd 2003/05/03 hitec)matsumoto
                sSQL = sSQL & " WHERE LOTID ='" & gtSprWfMap(lngLoopCnt).LOTID & "'" ' �u���b�NID"
                If intFromBlkSeq <= intToBlkSeq Then
                    sSQL = sSQL & "   AND ((BLOCKSEQ >= " & intFromBlkSeq & ")"    ' �u���b�N���A��"
                    sSQL = sSQL & "       AND (BLOCKSEQ <= " & intToBlkSeq & "))"
                Else
                    sSQL = sSQL & "   AND (BLOCKSEQ >= " & intFromBlkSeq & ")"    ' �u���b�N���A��"
                End If
                '' WriteDBLog sSql
                If 0 >= OraDB.ExecuteSQL(sSQL) Then
                    DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            Else
                sSQL = "UPDATE TBCMY011 SET"
                sSQL = sSQL & " mhinban = '" & gtSprWfMap(lngLoopCnt).hinban & "'"    ' �i��"
                dtmNowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
                '�擪�T���v��ID�͕ς��Ȃ�����SXLID�������̕����ō��ꂽ�܂܂Ƃ���@2003/04/22
                If lngLoopCnt = 1 Then
                    intNowIngotPos = SIngotP
                Else
                    intNowIngotPos = gtSprWfMap(lngLoopCnt).KESSYOUP
                End If
                sSQL = sSQL & ",MSXLID = '" & sGetSXLid & "'"

                sSQL = sSQL & ",UPDPROC = 'CW760'"             ' �X�V�H��
                sSQL = sSQL & ",UPDDATE = sysdate"    'upd 2003/05/03 hitec)matsumoto
                sSQL = sSQL & ",MREVNUM = " & gtSprWfMap(lngLoopCnt).REVNUM          ' ���i�ԍ������ԍ�
                sSQL = sSQL & ",MFACTORY = '" & gtSprWfMap(lngLoopCnt).factory & "'" ' �H��
                sSQL = sSQL & ",MOPECOND = '" & gtSprWfMap(lngLoopCnt).opecond & "'" ' ���Ə���
                sSQL = sSQL & " WHERE LOTID ='" & gtSprWfMap(lngLoopCnt).LOTID & "'"                   ' �u���b�NID"
                If (intFromBlkSeq <= intToBlkSeq) Then

                    sSQL = sSQL & "   AND ((BLOCKSEQ >= " & intFromBlkSeq & ")"    ' �u���b�N���A��"
                    sSQL = sSQL & "       AND (BLOCKSEQ <= " & intToBlkSeq & "))"
                Else
                    sSQL = sSQL & "   AND (BLOCKSEQ >= " & intFromBlkSeq & ")"    ' �u���b�N���A��"

                End If
                '' WriteDBLog sSql

                If 0 >= OraDB.ExecuteSQL(sSQL) Then
                    DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If

                sSQL = "UPDATE TBCMY011 SET"
                    sSQL = sSQL & " SHAFLAG = '0'"             ' �T���v���t���O"
                    sSQL = sSQL & ",WFSTA = '0'"               ' WF���
                sSQL = sSQL & " WHERE LOTID ='" & gtSprWfMap(lngLoopCnt).LOTID & "'"                   ' �u���b�NID"
                If (intFromBlkSeq <= intToBlkSeq) Then

                    sSQL = sSQL & "   AND ((BLOCKSEQ >= " & intFromBlkSeq & ")"    ' �u���b�N���A��"
                    sSQL = sSQL & "       AND (BLOCKSEQ <= " & intToBlkSeq & "))"

                Else
                    sSQL = sSQL & "   AND (BLOCKSEQ >= " & intFromBlkSeq & ")"    ' �u���b�N���A��"

                End If
                sSQL = sSQL & "  AND ( WFSTA <> '0'"
                sSQL = sSQL & "  AND  WFSTA <> '4')"
                '' WriteDBLog sSql

                If 0 >= OraDB.ExecuteSQL(sSQL) Then
                End If
            End If
        Next

        For lngLoopCnt = 1 To UBound(gtSprWfMap())
            If 0 = lngLoopCnt Mod 2 Then
                .GetText 2, lngLoopCnt - 1, vGetHinban
            Else
                .GetText 2, lngLoopCnt, vGetHinban
            End If
            If Trim(vGetHinban) <> "Z" Then
                .GetText 10, lngLoopCnt, vGetSample
                If (vGetSample <> vbNullString) Then
                    sSQL = "UPDATE TBCMY011 SET"
                    .GetText 10, lngLoopCnt, vGetSample
                    If vGetSample = gsWF_SMPL_JOINT Then    '���L
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
                        .GetText 38, lngLoopCnt, vGetSample
                        '############################# 2003/05/23 end

                        Call Cnv_GetSample(vGetSample)      '2004/01/29 ooba

                        sSQL = sSQL & " MSMPLEID = '" & vGetSample & "'" ' �����ʒu"
                        sSQL = sSQL & ",SHAFLAG = '1'"             ' �T���v���t���O"
                        sSQL = sSQL & ",WFSTA = '1'"    ' WF��ԃT���v��
                    Else
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
                        .GetText 38, lngLoopCnt, vGetSample

                        Call Cnv_GetSample(vGetSample)      '2004/01/29 ooba

                        sSQL = sSQL & " MSMPLEID = '" & vGetSample & "'" ' �����ʒu"
                        If lngLoopCnt <> 1 And lngLoopCnt <> UBound(gtSprWfMap()) Then  'upd hitec)matsumoto �����\���T���v���̃t���O�͍X�V���Ȃ�
                            sSQL = sSQL & ",WFSTA = '0'"    ' WF��ԃT���v��
                            sSQL = sSQL & ",SHAFLAG = '1'"             ' �T���v���t���O"
                        End If
                    End If

                    dtmNowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")

                    sSQL = sSQL & ",UPDPROC = 'CW760'"             ' �X�V�H��
                    sSQL = sSQL & ",UPDDATE = sysdate"    'upd 2003/05/03 hitec)matsumoto

                    sSQL = sSQL & " WHERE LOTID ='" & gtSprWfMap(lngLoopCnt).LOTID & "'"                   ' �u���b�NID"
                    sSQL = sSQL & " AND BLOCKSEQ = " & gtSprWfMap(lngLoopCnt).BLOCKSEQ              ' �u���b�N���A��"
                    '' WriteDBLog sSql
                    If 0 >= OraDB.ExecuteSQL(sSQL) Then
                        DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
            End If
        Next
    End With

    DBDRV_UPD_WFMap = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : DBDRV_WFLENGET
'*
'*    �����T�v      : 1.�Y���u���b�N��WF�P���̒����i�v�Z���j���擾
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                    ,����
'*                    BLOCKID       ,I  ,STRING                ,�u���b�N�h�c
'*                    dblWFLen      ,O  ,DOUBLE        �@�@    ,WF1���̌v�Z����
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function DBDRV_WFLENGET(ByVal strBlockID As String, _
                                ByRef dblWFLen As Double) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim intRealLen  As Integer
    Dim intWFcnt    As Integer
    Dim rs          As OraDynaset
    Dim intKetuFrom As Integer
    Dim intKetuTo   As Integer
    Dim intKetuLen  As Integer
    Dim sDBName     As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_WFLENGET"

    '�������AWF�����擾
    sDBName = "(Y011)"

    sSQL = "select e40.blockid,e40.reallen,y11.cnt"
    sSQL = sSQL & " from tbcme040 e40,"
    sSQL = sSQL & " xsdca xa,"
    sSQL = sSQL & " (select lotid,count(lotid) cnt"
    sSQL = sSQL & " from tbcmy011"
    sSQL = sSQL & " where lotid ='" & strBlockID & "'"
    sSQL = sSQL & " group by lotid  ) y11"
    sSQL = sSQL & " where e40.blockid = xa.CRYNUMCA"
    sSQL = sSQL & " and   y11.lotid   = xa.CRYNUMCA"
    sSQL = sSQL & " and   y11.lotid  = '" & strBlockID & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If Not rs.EOF Then
           intRealLen = CInt(rs!REALLEN)
           intWFcnt = CInt(rs!cnt)
    Else
        rs.Close
        DBDRV_WFLENGET = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    rs.Close

    '���������擾
    sDBName = "(Y012)"
    sSQL = "SELECT DISTINCT LENFROM,LENTO FROM TBCMY012"
    sSQL = sSQL & " Where "
    sSQL = sSQL & " LOTID   = '" & strBlockID & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    intKetuLen = 0
    Do While Not rs.EOF
        If (IsNull(rs.Fields("LENFROM")) = True) Or rs.Fields("LENFROM") = -1 Or _
            (IsNull(rs.Fields("LENTO")) = True) Or rs.Fields("LENTO") = -1 Then
        Else
            intKetuFrom = CInt(rs.Fields("LENFROM"))
            intKetuTo = CInt(rs.Fields("LENTO"))
            intKetuLen = intKetuLen + intKetuTo - intKetuFrom
        End If
        rs.MoveNext
    Loop
    rs.Close

    'WF�����v�Z
    dblWFLen = (intRealLen - intKetuLen) / intWFcnt

    DBDRV_WFLENGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    DBDRV_WFLENGET = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'**********************************************************************************************************
'*    �֐���        : GetZMotoHinban
'*
'*    �����T�v      : 1.TBCMY007����Z�i�Ԃ̌��i�Ԃ��擾����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^      ,����
'*                    strSXLID      ,I  ,String  ,SXL-ID
'*                    strMotoHinban ,O  ,String  ,���i��
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'**********************************************************************************************************
Public Function GetZMotoHinban(ByVal strSXLID As String, ByRef strMotoHinban As String) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rsMain      As OraDynaset
    Dim sErrTbl     As String
    Dim strDBName   As String

    On Error GoTo proc_err

    '�u���b�NID�擾
    strDBName = "Y007"
    sSQL = " select HINBAN from TBCMY007"
    sSQL = sSQL & "  where SXL_ID='" & strSXLID & "'"

    Set rsMain = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rsMain.RecordCount = 0 Then
        rsMain.Close
        GetZMotoHinban = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    Do While Not rsMain.EOF
        If IsNull(rsMain.Fields("HINBAN")) = True Then
            strMotoHinban = vbNullString
        Else
            strMotoHinban = Mid(Trim(rsMain.Fields("HINBAN")), 1, 8)
        End If
        rsMain.MoveNext
    Loop
    rsMain.Close
    GetZMotoHinban = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    GetZMotoHinban = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'*******************************************************************************************************
'*    �֐���        : scmzc_getE036
'*
'*    �����T�v      : 1.���i�d�lWF�f�[�^�iOT�P�AOT2)�̎擾�h���C�o
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^           ,����
'*                    pHIN          ,I  ,tFullHinban  ,�i�ԏ��
'*                    strOT1        ,O  ,String       ,���̑��T���v��1
'*                    strOT2        ,O  ,String       ,���̑��T���v��2
'*                    strMAI1       ,O  ,String       ,���̑��T���v��1����
'*                    strMAI2       ,O  ,String       ,���̑��T���v��2����
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************************
Public Function scmzc_getE036(pHIN As tFullHinban, strOT1 As String, strOT2 As String, strMAI1, strMAI2) _
                                As FUNCTION_RETURN
    Dim sSQL As String
    Dim rs  As OraDynaset

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function scmzc_getE036"

    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "   a.ot1 AS other1"
    sSQL = sSQL & "  ,a.ot1m AS other1mai"
    sSQL = sSQL & "  ,b.ot2 AS other2"
    sSQL = sSQL & "  ,b.ot2m AS other2mai"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "   ("
    sSQL = sSQL & "    SELECT"
    sSQL = sSQL & "      COUNT(other1)"
    sSQL = sSQL & "     ,MAX(other1) AS ot1"
    sSQL = sSQL & "     ,MAX(other1mai) AS ot1m"
    sSQL = sSQL & "    FROM"
    sSQL = sSQL & "      tbcme036"
    sSQL = sSQL & "    WHERE hinban   = '" & pHIN.hinban & "'"
    sSQL = sSQL & "      AND mnorevno = " & pHIN.mnorevno
    sSQL = sSQL & "      AND factory  = '" & pHIN.factory & "'"
    sSQL = sSQL & "      AND opecond  = '" & pHIN.opecond & "'"
    sSQL = sSQL & "      AND othertime > SYSDATE"
    sSQL = sSQL & "   ) a"
    sSQL = sSQL & "  ,("
    sSQL = sSQL & "    SELECT"
    sSQL = sSQL & "      COUNT(other2)"
    sSQL = sSQL & "     ,MAX(other2) AS ot2"
    sSQL = sSQL & "     ,MAX(other2mai) AS ot2m"
    sSQL = sSQL & "    FROM"
    sSQL = sSQL & "      tbcme036"
    sSQL = sSQL & "    WHERE hinban   = '" & pHIN.hinban & "'"
    sSQL = sSQL & "      AND mnorevno = " & pHIN.mnorevno
    sSQL = sSQL & "      AND factory  = '" & pHIN.factory & "'"
    sSQL = sSQL & "      AND opecond  = '" & pHIN.opecond & "'"
    sSQL = sSQL & "      AND othertime2 > SYSDATE"
    sSQL = sSQL & "   ) b"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        strOT1 = "0"
        strOT2 = "0"
        strMAI1 = "0"
        strMAI2 = "0"
        GoTo proc_exit
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
    scmzc_getE036 = FUNCTION_RETURN_SUCCESS
    rs.Close

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    scmzc_getE036 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : Pic_Disp
'*
'*    �����T�v      : 1.SXL�`�F�b�N�{�b�N�X�ڍׂ̕\��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    iIndex        ,I  ,Integer  ,�P�F�\�� / �O�F��\��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Public Sub Pic_Disp(iIndex As Integer)
    Dim intCnt    As Integer

    With f_cmbc039_3
        If iIndex = 0 Then
            For intCnt = 0 To 2
            .lbl_check(intCnt).Visible = False
            Next
            .pic_check(0).Visible = False
            .pic_check(1).Visible = False
        ElseIf iIndex = 1 Then
            For intCnt = 0 To 2
            .lbl_check(intCnt).Visible = True
            Next
            .pic_check(0).Visible = True
            .pic_check(1).Visible = True
        End If
    End With
End Sub

'*******************************************************************************
'*    �֐���        : CheckGetSampleID
'*
'*    �����T�v      : 1.�T���v���h�c�̎擾����
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^       ,����
'*                    iWafPos      ,I  ,Integer�@,�����w���e�[�u���ʒu
'*
'*    �߂�l        : Boolean(�I���̗L��)
'*
'*******************************************************************************
Public Function CheckGetSampleID(iWafPos As Integer) As Boolean
    Dim vNowhinban As Variant
    Dim vUDhinban  As Variant
    Dim vFlg       As Variant
    Dim sSampID    As Variant
    Dim vSampleID  As Variant
    Dim intPointer As Integer
    Dim vOldHinban As Variant
    Dim blCheckbox As Boolean       '�`�F�b�N�{�b�N�X�t���O
    Dim lngRow     As Long

    CheckGetSampleID = False

    With f_cmbc039_3.sprExamine
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
        .GetText 37, iWafPos, vFlg
        If vFlg = "1" Then  '�擪�ƍŏI�s�i�擪�͂��̊֐��ɂ͗��Ȃ��͂��j
            Exit Function

        ElseIf vFlg = "3" And iWafPos Mod 2 = 0 Then    '�����\���ŃT���v���̖����s�i�u���b�N�̋��j
            .GetText 2, iWafPos - 1, vNowhinban         '���݂̕i��
            .GetText 2, iWafPos + 1, vUDhinban  '���i��
            If vUDhinban <> vNowhinban Then
                '�`�F�b�N�{�b�N�X����
                .col = 1
                .row = iWafPos
                .CellType = CellTypeEdit
                .text = ""
                Call Pic_Disp(0) '03/05/31
                CheckGetSampleID = True
            Else
                If ADD_CHECKBOX(iWafPos, blCheckbox) = FUNCTION_RETURN_FAILURE Then
                    CheckGetSampleID = True
                End If
            End If
        ElseIf iWafPos Mod 2 = 0 Then   '�����s
            .GetText 2, iWafPos - 1, vNowhinban
            .GetText 2, iWafPos + 1, vUDhinban  '���i��
            If vNowhinban = vUDhinban Then
                If vFlg = "2" Or vFlg = "" Or vFlg = "0" Then
                    CheckGetSampleID = True
                End If
            Else
                CheckGetSampleID = True
            End If
        End If
    End With
End Function

'*****************************************************************************************************************
'*    �֐���        : GetSxlidINBlkid
'*
'*    �����T�v      : 1.����SXL����i�Ԃ̃u���b�N���E�ɔ����L���̃`�F�b�N�{�b�N�X�\���A�`�F�b�N�{�b�N�X�̓��e����
'*                      CW740�̏ꍇ�AA�������΂��ău���b�N�ŏI�s�Ƀ`�F�b�N�{�b�N�X��z�u����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^         ,����
'*�@�@                iWafPos       ,I  ,Integer�@  ,Spread�s
'*                    bCheckbox     ,IO ,Boolean    ,�`�F�b�N�{�b�N�X�\���t���O
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*****************************************************************************************************************
Private Function ADD_CHECKBOX(ByVal iWafPos As Integer, bCheckbox As Boolean) As FUNCTION_RETURN
    Dim j           As Integer
    Dim intRowCnt   As Integer
    Dim vGetLot     As Variant   'CW740�p
    Dim vGetLot2    As Variant   'CW740�p
    Dim intCnt      As Integer
    Dim vBackColor  As Variant

    ADD_CHECKBOX = FUNCTION_RETURN_SUCCESS

    intRowCnt = iWafPos

    With f_cmbc039_3.sprExamine
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
        .GetText 39, iWafPos, vGetLot
        .GetText 39, iWafPos + 1, vGetLot2
        If vGetLot = vGetLot2 Then
            Exit Function
        End If

        .col = 1
        .row = intRowCnt
        If .CellType <> CellTypeCheckBox Then
            bCheckbox = True
        ElseIf .text = "1" Then
            ADD_CHECKBOX = FUNCTION_RETURN_FAILURE
            Exit Function
        Else
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
            For j = 10 To 35
                If j <> 27 And j <> 35 Then
                    .SetText j, intRowCnt, vbNullString
                    .SetText j, intRowCnt + 1, vbNullString
                End If
            Next j

            .col = 2
            .row = intRowCnt - 1
            If .text <> "Z" Then
                .col = 11
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
                .col2 = 35
                .row = intRowCnt
                .row2 = intRowCnt
                .Lock = True
                .BlockMode = True
                .backColor = vbWhite
                .BlockMode = False
            Else
                .col = 11
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
                .col2 = 35
                .row = intRowCnt
                .row2 = intRowCnt
                .Lock = True
                .BlockMode = True
                .backColor = &H8080FF
                .BlockMode = False
            End If
            .col = 2
            .row = intRowCnt + 1
            If .text <> "Z" Then
                .col = 11
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
                .col2 = 35
                .row = intRowCnt + 1
                .row2 = intRowCnt + 1
                .BlockMode = True
                .backColor = vbWhite
                .BlockMode = False
                .Lock = True
            Else
                .col = 11
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
                .col2 = 35
                .row = intRowCnt + 1
                .row2 = intRowCnt + 1
                .BlockMode = True
                .backColor = &H8080FF
                .BlockMode = False
                .Lock = True
            End If
        End If
        If bCheckbox = True Then
            .col = 1
            .row = intRowCnt
            .Lock = False
            .CellType = CellTypeCheckBox
            Call Pic_Disp(1) '03/05/31
            .TypeCheckTextAlign = TypeCheckTextAlignLeft
            .TypeCheckType = TypeCheckTypeNormal
            .TypeCheckCenter = False
        End If
    End With
End Function

'*******************************************************************************
'*    �֐���        : Cnv_GetSample
'*
'*    �����T�v      : 1.�T���v��ID�̕ϊ�����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    vGetSample    ,I  ,Variant  ,SXL�Ǘ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Public Sub Cnv_GetSample(ByRef vGetSample As Variant)
    Dim i       As Integer
    Dim sKbn    As String

    For i = 1 To UBound(CngSmpID_UD)
        If CngSmpID_UD(i) = vGetSample Then
           sKbn = Cnv_Smp_KB(Right(vGetSample, 1))
           vGetSample = left(vGetSample, Len(vGetSample) - 1) + sKbn
           Exit Sub
        End If
    Next
End Sub

'*******************************************************************************
'*    �֐���        : Cnv_Smp_KB
'*
'*    �����T�v      : 1.�T���v���敪�̕ϊ�����
'*�@�@�@�@�@�@�@�@�@�@(�t�˂a�@�c�˂s�ɕϊ�)
'*    �p�����[�^    : �ϐ���        ,IO ,�^      ,����
'*                    SmpKb         ,I  ,String  ,�T���v���敪
'*
'*    �߂�l        : String�i�T���v���敪�j
'*
'*
'*******************************************************************************
Public Function Cnv_Smp_KB(SmpKb As String) As String
    If SmpKb = "U" Then
        Cnv_Smp_KB = "B"
        Exit Function
    End If

    If SmpKb = "D" Then
        Cnv_Smp_KB = "T"
        Exit Function
    End If
End Function

'********************************************************************************************************
'*    �֐���        : chkComSAMPL
'*
'*    �����T�v      : 1.���L�T���v���`�F�b�N����
'*                    (�w�肳�ꂽ�����ID���S���L���ǂ������������A�S���L�̏ꍇ�A���L�����ID���擾���Ԃ�)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^         ,����
'*                    inSXLID       ,I  ,String     , SXL-ID
'*                    inSMPLID      ,I  ,String     , �����ID
'*                    outSMPLID     ,O  ,String     , ���L�����ID(���L�łȂ��ꍇ�AinSMPLID��Ԃ�)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************************
Public Function chkComSAMPL(inSXLID As String, inSMPLID As String, outSMPLID As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sSQL        As String
    Dim sXTALCW     As String
    Dim sINPOSCW    As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function chkComSAMPL"

    '-------------------- �����ر ----------------------------------------
    chkComSAMPL = FUNCTION_RETURN_SUCCESS
    outSMPLID = inSMPLID

    '-------------------- �S���L�m�F(XSDCW) ----------------------------------------
    sSQL = "select XTALCW, INPOSCW from XSDCW "
    sSQL = sSQL & "where SXLIDCW = '" & inSXLID & "' and "
    sSQL = sSQL & "      REPSMPLIDCW = '" & inSMPLID & "' and "
    sSQL = sSQL & "      (WFINDRSCW = '2' or WFINDRSCW = '0' or WFINDRSCW = ' ' or WFINDRSCW is null) and "
    sSQL = sSQL & "      (WFINDOICW = '2' or WFINDOICW = '0' or WFINDOICW = ' ' or WFINDOICW is null) and "
    sSQL = sSQL & "      (WFINDB1CW = '2' or WFINDB1CW = '0' or WFINDB1CW = ' ' or WFINDB1CW is null) and "
    sSQL = sSQL & "      (WFINDB2CW = '2' or WFINDB2CW = '0' or WFINDB2CW = ' ' or WFINDB2CW is null) and "
    sSQL = sSQL & "      (WFINDB2CW = '2' or WFINDB3CW = '0' or WFINDB3CW = ' ' or WFINDB3CW is null) and "
    sSQL = sSQL & "      (WFINDL1CW = '2' or WFINDL1CW = '0' or WFINDL1CW = ' ' or WFINDL1CW is null) and "
    sSQL = sSQL & "      (WFINDL2CW = '2' or WFINDL2CW = '0' or WFINDL2CW = ' ' or WFINDL2CW is null) and "
    sSQL = sSQL & "      (WFINDL3CW = '2' or WFINDL3CW = '0' or WFINDL3CW = ' ' or WFINDL3CW is null) and "
    sSQL = sSQL & "      (WFINDL4CW = '2' or WFINDL4CW = '0' or WFINDL4CW = ' ' or WFINDL4CW is null) and "
    sSQL = sSQL & "      (WFINDDSCW = '2' or WFINDDSCW = '0' or WFINDDSCW = ' ' or WFINDDSCW is null) and "
    sSQL = sSQL & "      (WFINDDZCW = '2' or WFINDDZCW = '0' or WFINDDZCW = ' ' or WFINDDZCW is null) and "
    sSQL = sSQL & "      (WFINDSPCW = '2' or WFINDSPCW = '0' or WFINDSPCW = ' ' or WFINDSPCW is null) and "
    sSQL = sSQL & "      (WFINDDO1CW = '2' or WFINDDO1CW = '0' or WFINDDO1CW = ' ' or WFINDDO1CW is null) and "
    sSQL = sSQL & "      (WFINDDO2CW = '2' or WFINDDO2CW = '0' or WFINDDO2CW = ' ' or WFINDDO2CW is null) and "
    sSQL = sSQL & "      (WFINDDO3CW = '2' or WFINDDO3CW = '0' or WFINDDO3CW = ' ' or WFINDDO3CW is null) and "
    sSQL = sSQL & "      (WFINDOT1CW = '2' or WFINDOT1CW = '0' or WFINDOT1CW = ' ' or WFINDOT1CW is null) and "
    sSQL = sSQL & "      (WFINDOT2CW = '2' or WFINDOT2CW = '0' or WFINDOT2CW = ' ' or WFINDOT2CW is null) and "
    sSQL = sSQL & "      (WFINDAOICW = '2' or WFINDAOICW = '0' or WFINDAOICW = ' ' or WFINDAOICW is null) and "   '�c���_�f�ǉ��@03/12/19 ooba
    sSQL = sSQL & "      (WFINDGDCW = '2' or WFINDGDCW = '0' or WFINDGDCW = ' ' or WFINDGDCW is null or (WFINDGDCW = '1' and WFHSGDCW = '1')) "   'GD�ǉ��@05/02/24 ooba
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    sSQL = sSQL & "  and (EPINDB1CW = '2' or EPINDB1CW = '0' or EPINDB1CW = ' ' or EPINDB1CW is null) and "
    sSQL = sSQL & "      (EPINDB2CW = '2' or EPINDB2CW = '0' or EPINDB2CW = ' ' or EPINDB2CW is null) and "
'--- 2009/07/30 Change Y.Hitomi
'    sSql = sSql & "      (EPINDB2CW = '2' or EPINDB3CW = '0' or EPINDB3CW = ' ' or EPINDB3CW is null) and "
    sSQL = sSQL & "      (EPINDB3CW = '2' or EPINDB3CW = '0' or EPINDB3CW = ' ' or EPINDB3CW is null) and "
    sSQL = sSQL & "      (EPINDL1CW = '2' or EPINDL1CW = '0' or EPINDL1CW = ' ' or EPINDL1CW is null) and "
    sSQL = sSQL & "      (EPINDL2CW = '2' or EPINDL2CW = '0' or EPINDL2CW = ' ' or EPINDL2CW is null) and "
    sSQL = sSQL & "      (EPINDL3CW = '2' or EPINDL3CW = '0' or EPINDL3CW = ' ' or EPINDL3CW is null) "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    sXTALCW = rs("XTALCW")      '�����ԍ�
    sINPOSCW = rs("INPOSCW")    '�������ʒu
    Set rs = Nothing

    '-------------------- ���L�����ID�̎擾(XSDCW) ----------------------------------------
    sSQL = "select REPSMPLIDCW from XSDCW "
    sSQL = sSQL & "where SXLIDCW like '" & left(sXTALCW, 9) & "%' and "     '���ޯ�����ڒǉ� 09/05/25 ooba
    sSQL = sSQL & "      XTALCW = '" & sXTALCW & "' and "
    sSQL = sSQL & "      INPOSCW = '" & sINPOSCW & "' and "
    sSQL = sSQL & "      SXLIDCW != '" & inSXLID & "' and "
    sSQL = sSQL & "      REPSMPLIDCW != '" & inSMPLID & "' "
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    outSMPLID = rs("REPSMPLIDCW")       '��\�����ID(���L)
    Set rs = Nothing

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    chkComSAMPL = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************************
'*    �֐���        : cmbc039_GetSxlRsData
'*
'*    �����T�v      : 1.SXL�m��w��(TBCMY007)ð��قɾ�Ă���SXL�̔��R�ް����擾����B
'*
'*    �p�����[�^    : �ϐ���        ,IO  ,�^                ,����
'*                    oldSXLID      ,I   ,String            ��SXLID
'*                    newSXLID      ,I   ,String            �VSXLID
'*                    iRow          ,I   ,String            ��ʽ��گ�ލs��
'*                    sDataPattern  ,I   ,String            ���R�ް��擾�����
'*                                                            �������A : WF�����ް��擾
'*                                                            �������B : ���������ް��擾
'*                                                            �������C : �擾�ް��Ȃ�
'*                    iSxlPattern   ,I   ,String            �o�^SXL�����
'*                                                            �������1 : �S�p��SXL
'*                                                            �������2 : ��Ǎ���SXL
'*                                                            �������3 : ���Ǎ���SXL
'*                                                            �������4 : SXL�̊Ԃ�Z
'*                                                            �������0 : �擾�ް��Ȃ�
'*                    mesdata()     ,O   ,String            ���R�ް�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function cmbc039_GetSxlRsData(oldSXLID As String, newSXLID As String, IRow As Integer, _
                                        sDataPattern As String, iSxlPattern As Integer, _
                                        mesdata() As String) As FUNCTION_RETURN
    Dim sTBkbn      As String        'T/B�敪
    Dim sBlkId      As String        '�������ۯ�ID
    Dim sSmpId      As String        '�����ID(Rs)
    Dim i           As Integer
    Dim j           As Integer
    Dim intChkRow   As Integer
    Dim sSQL        As String
    Dim rs          As OraDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function cmbc039_GetSxlRsData"
    cmbc039_GetSxlRsData = FUNCTION_RETURN_FAILURE

    '���R�ް�������
    For i = 1 To 10
        mesdata(i) = ""
    Next

    '���R�ް��擾����݂��wA�x�̏ꍇ�AWF�����ް�(TBCMY013)���擾����B
    If sDataPattern = "A" Then
        For i = 1 To 2
            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"

            '�o�^SXL���w�����1�x�A�w�����2�xTOP���A�w�����3�xBOT����ð��ق�������ID(Rs)���擾����B
            If iSxlPattern = 1 Or (iSxlPattern = 2 And sTBkbn = "T") Or _
                (iSxlPattern = 3 And sTBkbn = "B") Then

                '�Y��SXL���A�V����يǗ�-WF<XSDCW>�̻����ID_Rs���擾�B
                sSQL = "select WFSMPLIDRSCW "
                sSQL = sSQL & "from XSDCW "
                sSQL = sSQL & "where TBKBNCW = '" & sTBkbn & "' "
                sSQL = sSQL & "and SXLIDCW = '" & oldSXLID & "' "

                Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

                If rs.RecordCount = 1 Then
                    sSmpId = rs("WFSMPLIDRSCW")
                End If
                Set rs = Nothing

            '�o�^SXL���w�����2�xBOT���A�w�����3�xTOP���A�w�����4�x�͓����ް���������ID(Rs)���擾����B
            ElseIf (iSxlPattern = 2 And sTBkbn = "B") Or (iSxlPattern = 3 And sTBkbn = "T") Or _
                    iSxlPattern = 4 Then

                If f_cmbc039_3.sprExamine.MaxRows = UBound(tblWfSample) Then
                    If sTBkbn = "T" Then
                        sSmpId = tblWfSample(IRow).WFSMP.WFSMPLIDRSCW
                    ElseIf sTBkbn = "B" Then
                        sSmpId = tblWfSample(IRow + 1).WFSMP.WFSMPLIDRSCW
                    End If

                '1SXL������ۯ���SXL�������Ȃ��ꍇ
                Else
                    If IRow > UBound(tblWfSample) Then
                        intChkRow = UBound(tblWfSample)
                    Else
                        intChkRow = IRow + 1
                    End If
                    For j = intChkRow To 1 Step -1
                        If tblWfSample(j).WFSMP.SXLIDCW = newSXLID Then
                            'TOP���͊�s�ABOT���͋����s
                            If (sTBkbn = "T" And j Mod 2 = 1) Or (sTBkbn = "B" And j Mod 2 = 0) Then
                                sSmpId = tblWfSample(j).WFSMP.WFSMPLIDRSCW
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If

            If Trim(sSmpId) <> "" Then
                '�����ID_Rs����A����]������<TBCMY013>�̔��R�����ް�(TOP��/BOT��)���擾����B
                sSQL = "select MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5 "
                sSQL = sSQL & "from TBCMY013 "
                sSQL = sSQL & "where OSITEM = 'RES' "
                sSQL = sSQL & "and SAMPLEID = '" & sSmpId & "' "

                Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

                If rs.RecordCount = 1 Then
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
                End If
                Set rs = Nothing
            End If
        Next
    '���R�ް��擾����݂��wB�x�̏ꍇ�A���������ް�(TBCMJ002)���擾����B
    ElseIf sDataPattern = "B" Then
        For i = 1 To 2
            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"

            '�o�^SXL���w�����1�x�A�w�����2�xTOP���A�w�����3�xBOT����ð��ق���������ۯ�ID���擾����B
            If iSxlPattern = 1 Or (iSxlPattern = 2 And sTBkbn = "T") Or _
                (iSxlPattern = 3 And sTBkbn = "B") Then

                '�Y��SXL���A�V����يǗ�-WF<XSDCW>�̻������ۯ�ID���擾
                sSQL = "select SMCRYNUMCW "
                sSQL = sSQL & "from XSDCW "
                sSQL = sSQL & "where TBKBNCW = '" & sTBkbn & "' "
                sSQL = sSQL & "and SXLIDCW = '" & oldSXLID & "' "

                Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

                If rs.RecordCount = 1 Then
                    sBlkId = rs("SMCRYNUMCW")
                End If
                Set rs = Nothing
            '�o�^SXL���w�����2�xBOT���A�w�����3�xTOP���A�w�����4�x�͓����ް�����������ۯ�ID���擾����B
            ElseIf (iSxlPattern = 2 And sTBkbn = "B") Or (iSxlPattern = 3 And sTBkbn = "T") Or _
                    iSxlPattern = 4 Then

                If f_cmbc039_3.sprExamine.MaxRows = UBound(tblWfSample) Then
                    If sTBkbn = "T" Then
                        sBlkId = tblWfSample(IRow).WFSMP.SMCRYNUMCW
                    ElseIf sTBkbn = "B" Then
                        sBlkId = tblWfSample(IRow + 1).WFSMP.SMCRYNUMCW
                    End If
                '1SXL������ۯ���SXL�������Ȃ��ꍇ
                Else
                    If IRow > UBound(tblWfSample) Then
                        intChkRow = UBound(tblWfSample)
                    Else
                        intChkRow = IRow + 1
                    End If
                    For j = intChkRow To 1 Step -1
                        If tblWfSample(j).WFSMP.SXLIDCW = newSXLID Then
                            'TOP���͊�s�ABOT���͋����s
                            If (sTBkbn = "T" And j Mod 2 = 1) Or (sTBkbn = "B" And j Mod 2 = 0) Then
                                sBlkId = tblWfSample(j).WFSMP.SMCRYNUMCW
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If

            If Trim(sBlkId) <> "" Then
                'T/B�敪�A�������ۯ�ID����A�V����يǗ�-��ۯ�<XSDCS>�̌����ԍ��A�����ID_Rs���擾�B
                '�����ԍ��A�����ID_Rs����A������R����<TBCMJ002>�̔��R�����ް�(TOP��/BOT��)���擾����B
                sSQL = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 "
                sSQL = sSQL & "from TBCMJ002 "
                sSQL = sSQL & "where (CRYNUM, SMPLNO) in ( "
                sSQL = sSQL & "         select XTALCS, CRYSMPLIDRSCS "
                sSQL = sSQL & "         from XSDCS "
                sSQL = sSQL & "         where TBKBNCS = '" & sTBkbn & "' "
                sSQL = sSQL & "         and CRYNUMCS = '" & sBlkId & "') "
                sSQL = sSQL & "and TRANCNT = ( "
                sSQL = sSQL & "         select max(TRANCNT) "
                sSQL = sSQL & "         from TBCMJ002 "
                sSQL = sSQL & "         where (CRYNUM, SMPLNO) in ( "
                sSQL = sSQL & "                  select XTALCS, CRYSMPLIDRSCS "
                sSQL = sSQL & "                  from XSDCS "
                sSQL = sSQL & "                  where TBKBNCS = '" & sTBkbn & "' "
                sSQL = sSQL & "                  and CRYNUMCS = '" & sBlkId & "')) "

                Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

                If rs.RecordCount = 1 Then
                    'TOP�������ް�
                    If sTBkbn = "T" Then
                        mesdata(1) = CStr(rs("MEAS1"))
                        mesdata(2) = CStr(rs("MEAS2"))
                        mesdata(3) = CStr(rs("MEAS3"))
                        mesdata(4) = CStr(rs("MEAS4"))
                        mesdata(5) = CStr(rs("MEAS5"))
                    'BOT�������ް�
                    ElseIf sTBkbn = "B" Then
                        mesdata(6) = CStr(rs("MEAS1"))
                        mesdata(7) = CStr(rs("MEAS2"))
                        mesdata(8) = CStr(rs("MEAS3"))
                        mesdata(9) = CStr(rs("MEAS4"))
                        mesdata(10) = CStr(rs("MEAS5"))
                    End If
                End If
                Set rs = Nothing
            End If
        Next
    '���R�ް��擾����݂��wC�x�̏ꍇ�A�擾�����ް��Ȃ��B
    ElseIf sDataPattern = "C" Then
    End If

    '�擾�ް�����/-1/NULL�̎��ͽ�߰���Ă���B
    For i = 1 To 10
        If mesdata(i) = "" Or mesdata(i) = "-1" Or mesdata(i) = vbNullString Then
            mesdata(i) = " "
        End If
    Next

    cmbc039_GetSxlRsData = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    cmbc039_GetSxlRsData = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************************
'*    �֐���        : DBDRV_GetTBCMJ015Cnt
'*
'*    �����T�v      : 1.GD���ё�������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^              ,����
'*      �@�@          sSampleid�@�@ ,I  ,String        �@,�����ID
'*      �@�@          iRecCnt�@�@   ,O  ,Integer         ,ں��ސ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function DBDRV_GetTBCMJ015Cnt(sSampleid As String, iRecCnt As Integer) As FUNCTION_RETURN
    Dim sSQL    As String
    Dim rs      As OraDynaset

    '�����ID�ƕۏ��׸ނ�����GD���т��擾
    sSQL = "SELECT CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, HSFLG, SMPLNO "
    sSQL = sSQL & "FROM TBCMJ015 "
    sSQL = sSQL & "WHERE SMPLNO = '" & sSampleid & "' "
    sSQL = sSQL & "AND HSFLG = '1' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        iRecCnt = 0
        DBDRV_GetTBCMJ015Cnt = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '���o���ʂ�ں��ސ���o�^
    iRecCnt = rs.RecordCount
    rs.Close

    DBDRV_GetTBCMJ015Cnt = FUNCTION_RETURN_SUCCESS
End Function

'*****************************************************************************************
'*    �֐���        : ChkAoiSiyou
'*
'*    �����T�v      : 1.�_�f�͏o�Ǝc���_�f�̎d�l�`�F�b�N
'*                    (�_�f�͏o(��oi)�Ǝc���_�f�̗����Ɏd�l�������Ă����ꍇ�G���[��Ԃ�)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^              ,����
'*                �@�@pHin�@�@    �@,I  ,tFullHinban   �@,�i��
'*
'*    �߂�l        : �d�l�`�F�b�N����(-1:�װ�C0:AOi�d�l���C1:AOi�d�l�L)
'*
'*****************************************************************************************
Public Function ChkAoiSiyou(pHIN As tFullHinban) As Integer
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim sDoiSiyou(2)    As String       '�����L��(DOi1�`3)
    Dim sAoiSiyou       As String       '�����L��(AOi)
    Dim intCnt          As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function ChkAoiSiyou"

    sSQL = "select HWFOS1HS, HWFOS2HS, HWFOS3HS, HWFZOHWS from TBCME025 "
    sSQL = sSQL & "where HINBAN = '" & pHIN.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & pHIN.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & pHIN.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & pHIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        ChkAoiSiyou = -1
        GoTo proc_exit
    End If

    If IsNull(rs("HWFOS1HS")) = False Then sDoiSiyou(0) = rs("HWFOS1HS") '�iWF�_�f�͏o1�ۏؕ��@_��
    If IsNull(rs("HWFOS2HS")) = False Then sDoiSiyou(1) = rs("HWFOS2HS") '�iWF�_�f�͏o2�ۏؕ��@_��
    If IsNull(rs("HWFOS3HS")) = False Then sDoiSiyou(2) = rs("HWFOS3HS") '�iWF�_�f�͏o3�ۏؕ��@_��
    If IsNull(rs("HWFZOHWS")) = False Then sAoiSiyou = rs("HWFZOHWS")    '�iWF�c���_�f�ۏؕ��@_��

    '�_�f�͏o�Ǝc���_�f�̎d�l�`�F�b�N
    ChkAoiSiyou = 0
    For intCnt = 0 To 2
        If sDoiSiyou(intCnt) = "H" Or sDoiSiyou(intCnt) = "S" Then
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

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    ChkAoiSiyou = -1
    Resume proc_exit
End Function

'***********************************************************************************************
'*    �֐���        : DBDRV_GetTBCMJ016Cnt
'*
'*    �����T�v      : 1.SPV���ё�������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^              ,����
'*      �@�@          sSampleid�@�@ ,I  ,String        �@,�����ID
'*      �@�@          iRecCnt�@�@   ,O  ,Integer         ,ں��ސ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************************
Public Function DBDRV_GetTBCMJ016Cnt(sSampleid As String, iRecCnt As Integer) As FUNCTION_RETURN
    Dim sSQL    As String
    Dim rs      As OraDynaset

    '�����ID�ƕۏ��׸ނ�����SPV���т��擾
    sSQL = "SELECT CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, HSFLG, SMPLNO "
    sSQL = sSQL & "FROM TBCMJ016 "
    sSQL = sSQL & "WHERE SMPLNO = '" & sSampleid & "' "
    sSQL = sSQL & "AND HSFLG = '1' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        iRecCnt = 0
        DBDRV_GetTBCMJ016Cnt = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '���o���ʂ�ں��ސ���o�^
    iRecCnt = rs.RecordCount
    rs.Close

    DBDRV_GetTBCMJ016Cnt = FUNCTION_RETURN_SUCCESS

End Function

'***************************************************************************************
'*    �֐���        : DBDRV_WARPMAPGET
'*
'*    �����T�v      : 1.WFϯ���ް��擾(Warp����p)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                  ,����
'*          �@�@      tWarpMapTmp() ,I  ,type_DBDRV_Nukisi   ,WFϯ���ް�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************
Public Function DBDRV_WARPMAPGET(tWarpMapTmp() As type_DBDRV_Nukisi) As FUNCTION_RETURN
    Dim i, j, k, m, n   As Integer
    Dim sSQL            As String
    Dim rs              As OraDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_WARPMAPGET"

    m = 0
    ReDim sWrpLOTID(0)
    ReDim iWrpBLOCKSEQ(0)
    ReDim tWarpMapTmp(0)

    For i = 0 To UBound(tSXLID)
        sSQL = "select "
        sSQL = sSQL & "LOTID, "
        sSQL = sSQL & "BLOCKSEQ, "
        sSQL = sSQL & "MSXLID, "
        sSQL = sSQL & "MHINBAN, "
        sSQL = sSQL & "MREVNUM, "
        sSQL = sSQL & "MFACTORY, "
        sSQL = sSQL & "MOPECOND, "
        sSQL = sSQL & "SHAFLAG, "
        sSQL = sSQL & "MSMPLEID "
        sSQL = sSQL & "from TBCMY011 "
        sSQL = sSQL & "where LOTID = '" & tSXLID(i).LOTID & "' "
        sSQL = sSQL & "and MSXLID = '" & tSXLID(i).SXLID & "' "
        sSQL = sSQL & "order by BLOCKSEQ "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_WARPMAPGET = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        k = UBound(sWrpLOTID)
        n = rs.RecordCount
        j = 0
        ReDim Preserve sWrpLOTID(k + n)         '��ۯ�ID
        ReDim Preserve iWrpBLOCKSEQ(k + n)      '��ۯ����A��

        Do While Not rs.EOF
            j = j + 1
            '��ۯ�ID
            If IsNull(rs("LOTID")) Then
                sWrpLOTID(k + j) = ""
            Else
                sWrpLOTID(k + j) = rs("LOTID")
            End If
            '��ۯ����A��
            If IsNull(rs("BLOCKSEQ")) Then
                iWrpBLOCKSEQ(k + j) = 0
            Else
                iWrpBLOCKSEQ(k + j) = rs("BLOCKSEQ")
            End If
            rs.MoveNext
        Loop

        'SXL��TOP��
        rs.MoveFirst
        m = m + 1
        ReDim Preserve tWarpMapTmp(m)
        With tWarpMapTmp(m)
            '��ۯ�ID
            If IsNull(rs("LOTID")) = False Then .LOTID = rs("LOTID") Else .LOTID = vbNullString
            '��ۯ����A��
            If IsNull(rs("BLOCKSEQ")) = False Then .BLOCKSEQ = rs("BLOCKSEQ") Else .BLOCKSEQ = "0"
            'SXLID
            If IsNull(rs("MSXLID")) = False Then .SXLID = rs("MSXLID") Else .SXLID = vbNullString
            '�i��
            If IsNull(rs("MHINBAN")) = False Then .hinban = rs("MHINBAN") Else .hinban = vbNullString
            '���i�ԍ������ԍ�
            If IsNull(rs("MREVNUM")) = False Then .REVNUM = rs("MREVNUM") Else .REVNUM = 0
            '�H��
            If IsNull(rs("MFACTORY")) = False Then .factory = rs("MFACTORY") Else .factory = vbNullString
            '���Ə���
            If IsNull(rs("MOPECOND")) = False Then .opecond = rs("MOPECOND") Else .opecond = vbNullString
            '�����ʒu(�����ID)
            If IsNull(rs("MSMPLEID")) = False Then .SMPLEID = rs("MSMPLEID") Else .SMPLEID = vbNullString
            '������׸�
            If IsNull(rs("SHAFLAG")) = False Then .SHAFLAG = rs("SHAFLAG") Else .SHAFLAG = vbNullString
            If Trim(.SHAFLAG) = "1" Then
                If Trim(.SMPLEID) = vbNullString Then
                    DBDRV_WARPMAPGET = FUNCTION_RETURN_FAILURE
                    rs.Close
                    GoTo proc_exit
                End If
            End If
        End With

        'SXL��BOT��
        rs.MoveLast
        m = m + 1
        ReDim Preserve tWarpMapTmp(m)
        With tWarpMapTmp(m)
            '��ۯ�ID
            If IsNull(rs("LOTID")) = False Then .LOTID = rs("LOTID") Else .LOTID = vbNullString
            '��ۯ����A��
            If IsNull(rs("BLOCKSEQ")) = False Then .BLOCKSEQ = rs("BLOCKSEQ") Else .BLOCKSEQ = "0"
            'SXLID
            If IsNull(rs("MSXLID")) = False Then .SXLID = rs("MSXLID") Else .SXLID = vbNullString
            '�i��
            If IsNull(rs("MHINBAN")) = False Then .hinban = rs("MHINBAN") Else .hinban = vbNullString
            '���i�ԍ������ԍ�
            If IsNull(rs("MREVNUM")) = False Then .REVNUM = rs("MREVNUM") Else .REVNUM = 0
            '�H��
            If IsNull(rs("MFACTORY")) = False Then .factory = rs("MFACTORY") Else .factory = vbNullString
            '���Ə���
            If IsNull(rs("MOPECOND")) = False Then .opecond = rs("MOPECOND") Else .opecond = vbNullString
            '�����ʒu(�����ID)
            If IsNull(rs("MSMPLEID")) = False Then .SMPLEID = rs("MSMPLEID") Else .SMPLEID = vbNullString
            '������׸�
            If IsNull(rs("SHAFLAG")) = False Then .SHAFLAG = rs("SHAFLAG") Else .SHAFLAG = vbNullString
            If Trim(.SHAFLAG) = "1" Then
                If Trim(.SMPLEID) = vbNullString Then
                    DBDRV_WARPMAPGET = FUNCTION_RETURN_FAILURE
                    rs.Close
                    GoTo proc_exit
                End If
            End If
        End With
        rs.Close
    Next i

    DBDRV_WARPMAPGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    DBDRV_WARPMAPGET = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'***************************************************************************************
'*    �֐���        : DBDRV_KanrenBlk
'*
'*    �����T�v      : 1.�֘A��ۯ��R�t�R��(TBCMY023)�o�^
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^                 ,����
'*      �@�@      �@�@sCrynum     ,I  ,String         �@  ,�����ԍ�
'*      �@�@      �@�@sKblockid() ,I  ,type_DBDRV_LOTSXL  ,�֘A��ۯ�
'*      �@�@      �@�@iSpos       ,I  ,Integer        �@  ,�������J�n�ʒu
'*                �@�@iEpos       ,I  ,Integer        �@  ,�������I���ʒu
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************
Public Function DBDRV_KanrenBlk(sCryNum As String, sKblockid() As type_DBDRV_LOTSXL, _
                                iSpos As Integer, iEpos As Integer) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim i, j            As Long
    Dim rs              As OraDynaset
    Dim lngRecCnt       As Long             'ں��ސ�
    Dim sLotid          As String           '��ۯ�ID(WFϯ��)
    Dim sSXLID          As String           'SXLID(WFϯ��)
    Dim udtKanrenData() As typ_TBCMY023     '�֘A��ۯ��R�t�R���ް�
    Dim blCutFlg        As Boolean          '�֘A��ۯ��R�؂��׸�
    Dim intTrnCnt       As Integer          '������

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_KanrenBlk"

    DBDRV_KanrenBlk = FUNCTION_RETURN_FAILURE

    '�����񐔎擾
    sSQL = "SELECT NVL(MAX(TRANCNT),0) MAXCNT FROM TBCMY023"
    sSQL = sSQL & " WHERE CRYNUM = '" & sCryNum & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        intTrnCnt = 1
    Else
        intTrnCnt = rs("MAXCNT") + 1          '������(�ő�) + 1
    End If
    rs.Close

    lngRecCnt = 0             '�o�^ں��ސ�
    blCutFlg = False         '�֘A��ۯ��R�؂��׸�(False:�R�؂薳)

    '�֘A��ۯ��R���ް����
    For i = 0 To UBound(sKblockid)
        lngRecCnt = lngRecCnt + 1
        ReDim Preserve udtKanrenData(lngRecCnt)
        With udtKanrenData(lngRecCnt)
            .CRYNUM = sCryNum               '�����ԍ�
            .TRANCNT = intTrnCnt              '������
            .BLOCKID = sKblockid(i).LOTID   '��ۯ�ID
            .PROCCAT = "D"                  '�����敪(D:�R��)
            .TXID = "TX879I"                '��ݻ޸���ID
        End With
    Next i

    'WFϯ�߂����ۯ�ID,SXLID���擾
    sSQL = "SELECT LOTID, MSXLID FROM TBCMY011"
    sSQL = sSQL & " WHERE LOTID LIKE '" & left(sCryNum, 9) & "%'"
    sSQL = sSQL & " AND (WFSTA = '0' OR WFSTA = '1')"
    sSQL = sSQL & " AND RITOP_POS > " & iSpos
    sSQL = sSQL & " AND RITOP_POS <= " & iEpos
    sSQL = sSQL & " AND MSXLID IS NOT NULL"
    sSQL = sSQL & " GROUP BY LOTID, MSXLID"
    sSQL = sSQL & " ORDER BY LOTID, MAX(BLOCKSEQ)"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    '�ް��Ȃ�
    If rs.RecordCount = 0 Then
        rs.Close
        GoTo proc_exit
    End If

    '�֘A��ۯ��R�t�ް����
    For i = 1 To rs.RecordCount
        If i > 1 Then
            '����ۯ��œ���SXL(�֘A��ۯ���)
            If sLotid <> rs("LOTID") And sSXLID = rs("MSXLID") Then
                '�֘A��ۯ�(��)
                If udtKanrenData(lngRecCnt).BLOCKID <> sLotid Then
                    intTrnCnt = intTrnCnt + 1       '������
                    lngRecCnt = lngRecCnt + 1
                    ReDim Preserve udtKanrenData(lngRecCnt)
                    With udtKanrenData(lngRecCnt)
                        .CRYNUM = sCryNum               '�����ԍ�
                        .TRANCNT = intTrnCnt              '������
                        .BLOCKID = sLotid               '��ۯ�ID
                        .PROCCAT = "C"                  '�����敪(C:�t�ւ�)
                        .TXID = "TX879I"                '��ݻ޸���ID
                    End With
                End If
                '�֘A��ۯ�(��)
                lngRecCnt = lngRecCnt + 1
                ReDim Preserve udtKanrenData(lngRecCnt)
                With udtKanrenData(lngRecCnt)
                    .CRYNUM = sCryNum                   '�����ԍ�
                    .TRANCNT = intTrnCnt                  '������
                    .BLOCKID = rs("LOTID")              '��ۯ�ID
                    .PROCCAT = "C"                      '�����敪(C:�t�ւ�)
                    .TXID = "TX879I"                    '��ݻ޸���ID
                End With

            '����ۯ��ŕ�SXL(�֘A��ۯ��~)
            ElseIf sLotid <> rs("LOTID") And sSXLID <> rs("MSXLID") Then
                blCutFlg = True          '�֘A��ۯ��R�؂��׸�(True:�R�؂�L)
            End If
        End If
        sLotid = rs("LOTID")        '��ۯ�ID
        sSXLID = rs("MSXLID")       'SXLID
        rs.MoveNext
    Next i
    rs.Close

    '�֘A��ۯ��R�؂肪���������ꍇ�A�֘A��ۯ��R�t�R��(TBCMY023)�ɓo�^
    If blCutFlg Then
        For i = 1 To UBound(udtKanrenData)
            With udtKanrenData(i)
                sSQL = "INSERT INTO TBCMY023"
                sSQL = sSQL & " (CRYNUM,"
                sSQL = sSQL & " TRANCNT,"
                sSQL = sSQL & " BLOCKID,"
                sSQL = sSQL & " PROCCAT,"
                sSQL = sSQL & " TXID,"
                sSQL = sSQL & " REGDATE,"
                sSQL = sSQL & " SUMITFLAG,"               '07/12/21 ooba
                sSQL = sSQL & " SUMITSND,"                '07/12/21 ooba
                sSQL = sSQL & " SSENDNO,"                 '07/12/21 ooba
                sSQL = sSQL & " SENDFLAG,"
                sSQL = sSQL & " SENDDATE, "
                sSQL = sSQL & " PLANTCAT) "
                sSQL = sSQL & " VALUES"
                sSQL = sSQL & " ('" & .CRYNUM & "',"      '�����ԍ�
                sSQL = sSQL & .TRANCNT & ","              '������
                sSQL = sSQL & " '" & .BLOCKID & "',"      '��ۯ�ID
                sSQL = sSQL & " '" & .PROCCAT & "',"      '�����敪
                sSQL = sSQL & " '" & .TXID & "',"         '��ݻ޸���ID
                sSQL = sSQL & " SYSDATE,"                 '�o�^���t
                sSQL = sSQL & " '0',"                     'SUMIT���M�׸�  07/12/21 ooba
                sSQL = sSQL & " NULL,"                    'SUMIT���M���t  07/12/21 ooba
                sSQL = sSQL & " NULL,"                    '���M���A��  07/12/21 ooba
                sSQL = sSQL & " '0',"                     '���M�׸�
                sSQL = sSQL & " NULL, "                    '���M���t
                sSQL = sSQL & "  '" & sCmbMukesaki & "') "  '����
            End With

            '�o�^���s
            If OraDB.ExecuteSQL(sSQL) <= 0 Then
                GoTo proc_exit
            End If
        Next i
    End If

    DBDRV_KanrenBlk = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    gErr.HandleError
    Resume proc_exit
End Function
' add SETkimizuka Start 09/03/17
'***************************************************************************************
'*    �֐���        : DBDRV_XODY4GET
'*
'*    �����T�v      : ������~���ڎ擾
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^                 ,����
'*      �@�@      �@�@sCrynum     ,I  ,String         �@  ,�����ԍ�
'*      �@�@      �@�@sKblockid() ,I  ,type_DBDRV_LOTSXL  ,�֘A��ۯ�
'*      �@�@      �@�@iSpos       ,I  ,Integer        �@  ,�������J�n�ʒu
'*                �@�@iEpos       ,I  ,Integer        �@  ,�������I���ʒu
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'***************************************************************************************
Public Function DBDRV_XODY4GET(udt_ww() As DBDRV_scmzc_fcmlc001b_SXL039) As FUNCTION_RETURN
    Dim sql             As String
    Dim rs              As OraDynaset
    Dim sOldID          As String
    Dim iCnt            As Integer
    Dim sSxl            As String
    
    sSxl = "("
    For iCnt = 1 To UBound(udt_ww)
        sSxl = sSxl & "'" & udt_ww(iCnt).SXLIDCA & "'"
        If iCnt < UBound(udt_ww) Then
            sSxl = sSxl & ","
        End If
    Next
    sSxl = sSxl & ")"
    
    
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/30
    sql = "SELECT "
    sql = sql & "   NVL(SXLIDY3,' ') as SXLIDY4"          '
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUSY4 "
    sql = sql & " , DECODE(CAUSEY4,NULL,' ',TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSEY4"        '
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNOY4"
    sql = sql & " , NVL(STOPY4,'0') as STOP "
    sql = sql & " , NVL(WKKTY4,' ') as WKKTY4 "
    sql = sql & " FROM XODY3 Y3,XODY4 Y4,KODA9 A9 "
    sql = sql & " WHERE  "
    sql = sql & "  SXLIDY3 IN " & sSxl
    sql = sql & " AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & " AND Y3.RCNTY3 = Y4.RCNTY4(+) "
    sql = sql & " AND Y3.LIVKY3(+) = '0' "
    sql = sql & " AND Y4.LIVKY4(+) = '0' "
    sql = sql & " AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
'    sql = sql & " GROUP BY SXLIDY3,STOPY4,CAUSEY4,Y4.PRINTNOY4,Y4.PRINTKINDY4,NAMEJA9,AGRSTATUSY4,WKKTY4 "
    
    sql = sql & " UNION SELECT "
    sql = sql & "   NVL(SXLIDY3,' ') as SXLIDY4"          '
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUSY4 "
    sql = sql & " , DECODE(CAUSEY4,NULL,' ',TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSEY4"        '
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNOY4"
    sql = sql & " , NVL(STOPY4,'0') as STOP "
    sql = sql & " , NVL(WKKTY4,' ') as WKKTY4 "
    sql = sql & " FROM XODY3 Y3,XODY4 Y4,KODA9 A9 "
    sql = sql & " WHERE  "
    sql = sql & "  SXLIDY3 IN " & sSxl
    sql = sql & " AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & " AND Y4.WKKTY4(+) = 'CW000'"
    sql = sql & " AND Y3.LIVKY3(+) = '0' "
    sql = sql & " AND Y4.LIVKY4(+) = '0' "
    sql = sql & " AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
'    sql = sql & " GROUP BY SXLIDY3,STOPY4,CAUSEY4,Y4.PRINTNOY4,Y4.PRINTKINDY4,NAMEJA9,AGRSTATUSY4,WKKTY4 "
    
'    sql = "SELECT "
'    sql = sql & "   NVL(SXLIDY3,' ') as SXLIDY4"          '
'    sql = sql & " , NVL(TO_CHAR(MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4))),' ') as AGRSTATUSY4 "
'    sql = sql & " , DECODE(CAUSEY4,NULL,' ',TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSEY4"        '
'    sql = sql & " , NVL(Y5.PRINTKIND || Y5.PRINTNO,' ') as PRINTNOY4"
'    sql = sql & " , NVL(STOPY4,'0') as STOP "
'   sql = sql & "      FROM XODY3  "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XODY4.AGRSTATUSY4,XODY4.CAUSEY4,XODY4.STOPY4,XODY4.XTALNOY4 FROM XODY3,XODY4 "
'    sql = sql & "           INNER JOIN (SELECT MIN(DECODE(WKKTY4,'CW750',3,'CW760',2,'CW000',1,9)) as WKKTY4 ,XTALNOY4 FROM XODY4  "
'    sql = sql & "              WHERE STOPY4 <> '2' AND LIVKY4 = '0' AND WKKTY4 ='CW000'"
'    sql = sql & "              GROUP BY XTALNOY4) Y4_WKKT ON (XODY4.XTALNOY4 = Y4_WKKT.XTALNOY4 AND XODY4.WKKTY4    = DECODE(Y4_WKKT.WKKTY4,3,'CW750',2,'CW760',1,'CW000',' ') ) "
'    sql = sql & "           WHERE XTALNOY3 = XODY4.XTALNOY4 AND LIVKY3 = '0' AND XODY4.STOPY4 <> '2' AND XODY4.LIVKY4 = '0' AND XODY4.WKKTY4 ='CW000'"
'    sql = sql & "           GROUP BY XODY4.AGRSTATUSY4,XODY4.CAUSEY4,XODY4.STOPY4,XODY4.XTALNOY4 ) XODY4  on ( XTALNOY3 = XTALNOY4" & ")"
'    sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
'    sql = sql & "                FROM XODY3,XODY4,XODY5 "
'    sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
'    sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
'    sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
'    sql = sql & "      WHERE  "
'    sql = sql & "       ( AGRSTATUSY4 IS NOT NULL OR Y5.PRINTKIND IS NOT NULL) AND LIVKY3    = '0' AND SXLIDY3 IN " & sSxl
'    sql = sql & " GROUP BY SXLIDY3,STOPY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9,AGRSTATUSY4 "
'
'    sql = sql & " UNION SELECT "
'    sql = sql & "   NVL(SXLIDY3,' ') as SXLIDY4"          '
'    sql = sql & " , NVL(TO_CHAR(MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4))),' ') as AGRSTATUSY4 "
'    sql = sql & " , DECODE(CAUSEY4,NULL,' ',TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSEY4"        '
'    sql = sql & " , NVL(Y5.PRINTKIND || Y5.PRINTNO,' ') as PRINTNOY4"
'    sql = sql & " , NVL(STOPY4,'0') as STOP "
'    sql = sql & "      FROM XODY3  "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XODY4.AGRSTATUSY4,XODY4.CAUSEY4,XODY4.STOPY4,XODY4.SXLIDY4 FROM XODY3,XODY4 "
'    sql = sql & "           INNER JOIN (SELECT MIN(DECODE(WKKTY4,'CW750',3,'CW760',2,'CW000',1,9)) as WKKTY4 ,XTALNOY4,SXLIDY4 FROM XODY4  "
'    sql = sql & "              WHERE STOPY4 <> '2' AND LIVKY4 = '0' AND WKKTY4 in ('CW750','CW760')"
'    sql = sql & "              GROUP BY XTALNOY4,SXLIDY4) Y4_WKKT ON (XODY4.XTALNOY4 = Y4_WKKT.XTALNOY4 AND XODY4.SXLIDY4 = Y4_WKKT.SXLIDY4 AND XODY4.WKKTY4    = DECODE(Y4_WKKT.WKKTY4,3,'CW750',2,'CW760',1,'CW000',' ') ) "
'    sql = sql & "           WHERE XTALNOY3 = XODY4.XTALNOY4 AND RCNTY3 = XODY4.RCNTY4 AND LIVKY3 = '0' AND XODY4.STOPY4 <> '2' AND XODY4.LIVKY4 = '0' AND XODY4.WKKTY4 IN ('CW750','CW760')"
'    sql = sql & "           GROUP BY XODY4.AGRSTATUSY4,XODY4.CAUSEY4,XODY4.STOPY4,XODY4.SXLIDY4 ) XODY4  on ( SXLIDY3 = SXLIDY4" & ")"
'    sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
'    sql = sql & "                FROM XODY3,XODY4,XODY5 "
'    sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
'    sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
'    sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
'    sql = sql & "      WHERE  "
'    sql = sql & "       LIVKY3    = '0' AND SXLIDY3 IN " & sSxl
'    sql = sql & " GROUP BY SXLIDY3,STOPY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9,AGRSTATUSY4 "
'    sql = sql & " ORDER BY SXLIDY3"
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/30
'Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    Do While Not rs.EOF
        
        For iCnt = 1 To UBound(udt_ww)
            If udt_ww(iCnt).SXLIDCA = rs("SXLIDY4") Then
                If rs("STOP") <> "2" And (rs("WKKTY4") = "CW750" Or rs("WKKTY4") = "CW760" Or rs("WKKTY4") = "CW000") Then
                    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/30
                    'udt_ww(iCnt).STOP = rs("STOP")
                    'udt_ww(iCnt).AGRSTATUS = rs("AGRSTATUSY4")
                    If Trim(udt_ww(iCnt).AGRSTATUS) = "" Or rs("AGRSTATUSY4") < udt_ww(iCnt).AGRSTATUS Then
                        udt_ww(iCnt).STOP = rs("STOP")
                        udt_ww(iCnt).AGRSTATUS = rs("AGRSTATUSY4")
                    End If
                    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/30
                    If Trim(rs("CAUSEY4")) <> "" And InStr(udt_ww(iCnt).CAUSE, rs("CAUSEY4")) = 0 Then
                        udt_ww(iCnt).CAUSE = udt_ww(iCnt).CAUSE & rs("CAUSEY4") & vbTab
                    End If
                End If
                If Trim(rs("PRINTNOY4")) <> "" And InStr(udt_ww(iCnt).PRINTNO, rs("PRINTNOY4")) = 0 Then
                    udt_ww(iCnt).PRINTNO = udt_ww(iCnt).PRINTNO & rs("PRINTNOY4") & vbTab
                End If
                Exit For
            End If
        Next
        rs.MoveNext
    Loop

    rs.Close
    
proc_exit:
    '' �I��
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
' add SETkimizuka End 09/03/17

' add SPK_Hitomi Start 09/10/20
'********************************************************************************************************
'*    �֐���        : ChkHosho
'*
'*    �����T�v      : 1.�ۏؕ��@����
'*                    (XSDCW�̊m��敪���A��ۯ�,WF�ۏ؂𔻒肷��)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^         ,����
'*                    inSXLID       ,I  ,String     , SXL-ID
'*                    outHosho  �@  ,O  ,String     , �ۏؕ��@(1:��ۯ��ۏ�,2:WF�ۏ�)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************************
Public Function ChkHosho(inSXLID As String, outHosho As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sSQL        As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    '-------------------- �����ر ----------------------------------------
    ChkHosho = FUNCTION_RETURN_SUCCESS

    sSQL = "select REPSMPLIDCW,WFSMPLIDGDCW from XSDCW where SXLIDCW = '" & inSXLID & "' and KTKBNCW = '9' and LIVKCW = '0'"
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount > 0 Then
        outHosho = 1 '��ۯ��ۏ�
    ElseIf rs.RecordCount = 0 Then
        outHosho = 2 'WF�ۏ�
    End If
    
    Set rs = Nothing

proc_exit:

    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    ChkHosho = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START
'---------------------------------------------------------------------------
'�T�v      :�����ԍ���X�����TBCMJ022���������ASIRD��������Ԃ�
'---------------------------------------------------------------------------
'���Ұ�    :�ϐ���      ,IO     ,�^                     ,����
'          :pCRYNUM     ,I  �@�@,String                 ,�����ԍ�
'          :pflgSird    ,O  �@�@,Boolean                ,SIRD�������̗L��(True:�L�AFalse�F��)
'          :pSMPLID     ,O  �@�@,String                 ,SIRD�������̑�\�����ID
'          :�߂�l      ,O      ,Boolean                ,[True:OK�^False:NG]
'---------------------------------------------------------------------------
Public Function fncGetSirdSample(ByVal pCRYNUM As String, ByRef pflgSird As Boolean, ByRef pSMPLID As String) As Boolean

    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '--�װ����ׂ̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function fncGetSirdSample"
    
    '--������
    fncGetSirdSample = False: pflgSird = False: pSMPLID = ""
    Set rs = Nothing      'Oracle RecordSet Free

    '--SQL������
    sql = "select SMPLNO  "
    sql = sql & "from TBCMJ022 " & vbCrLf
    sql = sql & "where" & vbCrLf
'Cng Start 2011/10/06 Y.Hitomi
    sql = sql & "     substr(CRYNUM,1,9) = '" & left(pCRYNUM, 9) & "'" & vbCrLf     '�����ԍ�(��9��)
'    sql = sql & "     substr(CRYNUM,1,7) = '" & left(pCRYNUM, 7) & "'" & vbCrLf     '�����ԍ�(��7��)
'Cng Start 2011/10/06 Y.Hitomi
    sql = sql & " and TRANCNT   = 0" & vbCrLf                                       '������
    

    '--�ް��𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY Or ORADYN_NOCACHE)
    If rs Is Nothing Then
        GoTo proc_exit
    End If
    
    '--���o���ʎQ��
    If Not (rs.EOF) Then
        '<< �ް��L�� >>
        rs.MoveFirst
        pflgSird = True                 '[SIRD�������L��]
        pSMPLID = rs("SMPLNO")          '[��\�����ID]
    End If
    
    fncGetSirdSample = True

proc_exit:
    '<< �I�� >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing
    
    gErr.Pop
    Exit Function

proc_err:
    '<< �װ����� >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing

    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    
    gErr.HandleError
    Resume proc_exit

End Function
'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD END

'�T�v      :���Ԕ����P��(mm)�̎擾
'���Ұ��@�@:�ϐ���      ,IO   ,�^                ,����
'�@�@      :HIN  �@�@ �@,I    ,tFullHinban �@    ,12���i��
'�@�@      :iMSMPTANI�@ ,O    ,Integer �@        ,���Ԕ����P��(mm)
'      �@�@:�߂�l      ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :2011/06/30 Marushita
Public Function getMSMPTANI(HIN As tFullHinban, iMSMPTANI As Integer) As FUNCTION_RETURN

    Dim sSQL As String
    Dim rs As OraDynaset
    
    getMSMPTANI = FUNCTION_RETURN_FAILURE
        
    iMSMPTANI = 0
    If Trim(HIN.hinban) = "Z" Or Trim(HIN.hinban) = "G" Or Trim(HIN.hinban) = "" Then
        Exit Function
    End If
    
    sSQL = "SELECT MSMPTANI"
    sSQL = sSQL & " FROM TBCME036"
    sSQL = sSQL & " WHERE"
    sSQL = sSQL & " HINBAN = '" & HIN.hinban & "'"
    sSQL = sSQL & " AND MNOREVNO = " & HIN.mnorevno
    sSQL = sSQL & " AND FACTORY = '" & HIN.factory & "'"
    sSQL = sSQL & " AND OPECOND = '" & HIN.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields("MSMPTANI")) = False Then iMSMPTANI = rs.Fields("MSMPTANI") Else iMSMPTANI = 0
        getMSMPTANI = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
End Function
'�T�v      :�V���O���m��ۂ̃`�F�b�N
'���Ұ��@�@:�ϐ���      ,IO   ,�^                ,����
'�@�@      :HIN  �@�@ �@,I    ,tFullHinban �@    ,12���i��
'      �@�@:sSXLIDFLG   ,O    ,String   �@�@�@�@ ,SXLID�m��ۃt���O
'      �@�@:�߂�l      ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :2011/09/29 Y.Hitomi
Public Function getSXLIDFLG(HIN As tFullHinban, sSXLIDFLG) As FUNCTION_RETURN

    Dim sSQL As String
    Dim rs As OraDynaset
    
    getSXLIDFLG = FUNCTION_RETURN_FAILURE
        
    If Trim(HIN.hinban) = "Z" Or Trim(HIN.hinban) = "G" Or Trim(HIN.hinban) = "" Then
        Exit Function
    End If
    
    sSQL = "SELECT NVL(SXLIDFLG,'0') as SXLIDFLG "
    sSQL = sSQL & " FROM TBCME036"
    sSQL = sSQL & " WHERE"
    sSQL = sSQL & " HINBAN = '" & HIN.hinban & "'"
    sSQL = sSQL & " AND MNOREVNO = " & HIN.mnorevno
    sSQL = sSQL & " AND FACTORY = '" & HIN.factory & "'"
    sSQL = sSQL & " AND OPECOND = '" & HIN.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
'Add Start 2011/10/03 Y.Hitomi
    If rs.RecordCount > 0 And IsNull(rs.Fields("SXLIDFLG")) = False Then
'    If rs.RecordCount > 0 Then
'Add End 2011/10/03 Y.Hitomi
        sSXLIDFLG = rs.Fields("SXLIDFLG")
        getSXLIDFLG = FUNCTION_RETURN_SUCCESS
    Else
        sSXLIDFLG = "0"
    End If
           
    rs.Close
    
End Function

