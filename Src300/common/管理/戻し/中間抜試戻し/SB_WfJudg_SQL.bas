Attribute VB_Name = "SB_WfJudg_SQL"
Option Explicit

'' WF�Z���^�[��������҂��ꗗ

' SXL�Ǘ�
Public Type DBDRV_scmzc_fcmlc001b_SXL
    CRYNUM      As String * 12      ' �����ԍ�
    INGOTPOS    As Integer          ' �������J�n�ʒu
    Length      As Integer          ' ����
    SXLID       As String * 13      ' SXLID
    KRPROCCD    As String * 5       ' �Ǘ��H��
    NOWPROC     As String * 5       ' ���ݍH��
    LPKRPROCCD  As String * 5       ' �ŏI�ʉߊǗ��H��
    LASTPASS    As String * 5       ' �ŏI�ʉߍH��
    DELCLS      As String * 1       ' �폜�敪
    LSTATCLS    As String * 1       ' �ŏI��ԋ敪
    HOLDCLS     As String * 1       ' �z�[���h�敪
    hinban      As String * 8       ' �i��
    REVNUM      As Integer          ' ���i�ԍ������ԍ�
    factory     As String * 1       ' �H��
    opecond     As String * 1       ' ���Ə���
    COUNT       As Integer          ' ����
    REGDATE     As Date             ' �o�^���t
    UPDDATE     As Date             ' �X�V���t
    KETURAKU    As Boolean          ' �������L���t���O
    WFSMP()     As typ_XSDCW        ' �T���v���Ǘ��iTOP�ATAIL�� �Q���R�[�h�j
End Type


' WF�Z���^�[��������
' ���͗p
Public Type type_DBDRV_scmzc_fcmlc001c_In
    HIN         As tFullHinban      ' �i��(full)
    SAMPLEID    As String * 16      ' �T���v��ID
    SXLID       As String * 13      ' SXLID
    WFSMP       As typ_XSDCW        ' ����يǗ�
End Type

' WF���i�d�l�擾�p
Public Type type_DBDRV_scmzc_fcmlc001c_Siyou
    HWFTYPE As String * 1           ' �i�v�e�^�C�v
    HWFCDIR As String * 1           ' �i�v�e�����ʕ�
    HWFCDOP As String * 1           ' �i�v�e�����h�[�v
    HWFRMIN As Double               ' �i�v�e���R����
    HWFRMAX As Double               ' �i�v�e���R���
    HWFRSPOH As String * 1          ' �i�v�e���R����ʒu�Q��
    HWFRSPOT As String * 1          ' �i�v�e���R����ʒu�Q�_
    HWFRSPOI As String * 1          ' �i�v�e���R����ʒu�Q��
    HWFRHWYT As String * 1          ' �i�v�e���R�ۏؕ��@�Q��
    HWFRHWYS As String * 1          ' �i�v�e���R�ۏؕ��@�Q��
    HWFRMCAL As String * 1          ' �i�v�e���R�ʓ��v�Z
    HWFRAMIN As Double              ' �i�v�e���R���ω���
    HWFRAMAX As Double              ' �i�v�e���R���Ϗ��
    HWFRMBNP As Double              ' �i�v�e���R�ʓ����z
    
    HWFMKMIN As Double              ' �i�v�e�����בw����
    HWFMKMAX As Double              ' �i�v�e�����בw���
    HWFMKSPH As String * 1          ' �i�v�e�����בw����ʒu�Q��
    HWFMKSPT As String * 1          ' �i�v�e�����בw����ʒu�Q�_
    HWFMKSPR As String * 1          ' �i�v�e�����בw����ʒu�Q��
    HWFMKHWT As String * 1          ' �i�v�e�����בw�ۏؕ��@�Q��
    HWFMKHWS As String * 1          ' �i�v�e�����בw�ۏؕ��@�Q��

    HWFONMIN As Double              ' �i�v�e�_�f�Z�x����
    HWFONMAX As Double              ' �i�v�e�_�f�Z�x���
    HWFONSPH As String * 1          ' �i�v�e�_�f�Z�x����ʒu�Q��
    HWFONSPT As String * 1          ' �i�v�e�_�f�Z�x����ʒu�Q�_
    HWFONSPI As String * 1          ' �i�v�e�_�f�Z�x����ʒu�Q��
    HWFONHWT As String * 1          ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
    HWFONHWS As String * 1          ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
    HWFONMCL As String * 1          ' �i�v�e�_�f�Z�x�ʓ��v�Z
    HWFONMBP As Double              ' �i�v�e�_�f�Z�x�ʓ����z
    HWFONAMN As Double              ' �i�v�e�_�f�Z�x���ω���
    HWFONAMX As Double              ' �i�v�e�_�f�Z�x���Ϗ��

    HWFOS1MN As Double              ' �i�v�e�_�f�͏o�P����
    HWFOS1MX As Double              ' �i�v�e�_�f�͏o�P���
    HWFOS1SH As String * 1          ' �i�v�e�_�f�͏o�P����ʒu�Q��
    HWFOS1ST As String * 1          ' �i�v�e�_�f�͏o�P����ʒu�Q�_
    HWFOS1SI As String * 1          ' �i�v�e�_�f�͏o�P����ʒu�Q��
    HWFOS1HT As String * 1          ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
    HWFOS1HS As String * 1          ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
    HWFOS2SH As String * 1          ' �i�v�e�_�f�͏o�Q����ʒu�Q��
    HWFOS2ST As String * 1          ' �i�v�e�_�f�͏o�Q����ʒu�Q�_
    HWFOS2SI As String * 1          ' �i�v�e�_�f�͏o�Q����ʒu�Q��
    HWFOS2MN As Double              ' �i�v�e�_�f�͏o�Q����
    HWFOS2MX As Double              ' �i�v�e�_�f�͏o�Q���
    HWFOS2HT As String * 1          ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
    HWFOS2HS As String * 1          ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
    HWFOS3MN As Double              ' �i�v�e�_�f�͏o�R����
    HWFOS3MX As Double              ' �i�v�e�_�f�͏o�R���
    HWFOS3SH As String * 1          ' �i�v�e�_�f�͏o�R����ʒu�Q��
    HWFOS3ST As String * 1          ' �i�v�e�_�f�͏o�R����ʒu�Q�_
    HWFOS3SI As String * 1          ' �i�v�e�_�f�͏o�R����ʒu�Q��
    HWFOS3HT As String * 1          ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��
    HWFOS3HS As String * 1          ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��

    HWFZOMIN As Double              ' �i�v�e�c���_�f����
    HWFZOMAX As Double              ' �i�v�e�c���_�f���
    HWFZOSPH As String * 1          ' �i�v�e�c���_�f����ʒu�Q��
    HWFZOSPT As String * 1          ' �i�v�e�c���_�f����ʒu�Q�_
    HWFZOSPI As String * 1          ' �i�v�e�c���_�f����ʒu�Q��
    HWFZOHWT As String * 1          ' �i�v�e�c���_�f�ۏؕ��@�Q��
    HWFZOHWS As String * 1          ' �i�v�e�c���_�f�ۏؕ��@�Q��
    
    HWFDSOMX As Double              ' �i�v�e�c�r�n�c���
    HWFDSOMN As Double              ' �i�v�e�c�r�n�c����
    HWFDSOAX As Integer             ' �i�v�e�c�r�n�c�̈���
    HWFDSOAN As Integer             ' �i�v�e�c�r�n�c�̈扺��
    HWFDSOHT As String * 1          ' �i�v�e�c�r�n�c�ۏؕ��@�Q��
    HWFDSOHS As String * 1          ' �i�v�e�c�r�n�c�ۏؕ��@�Q��
    HWFDSOPTK As String * 1         ' �i�v�e�c�r�n�c�p�^���敪
    
    HWFSPVMX As Double              ' �i�v�e�r�o�u�e�d���
    HWFSPVAM As Double              ' �i�v�e�r�o�u�e�d���Ϗ��
    HWFSPVSH As String * 1          ' �i�v�e�r�o�u�e�d����ʒu�Q��
    HWFSPVST As String * 1          ' �i�v�e�r�o�u�e�d����ʒu�Q�_
    HWFSPVSI As String * 1          ' �i�v�e�r�o�u�e�d����ʒu�Q��
    HWFSPVHT As String * 1          ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
    HWFSPVHS As String * 1          ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
    HWFDLSPH As String * 1          ' �i�v�e�g�U������ʒu�Q��
    HWFDLSPT As String * 1          ' �i�v�e�g�U������ʒu�Q�_
    HWFDLSPI As String * 1          ' �i�v�e�g�U������ʒu�Q��
    HWFDLHWT As String * 1          ' �i�v�e�g�U���ۏؕ��@�Q��
    HWFDLHWS As String * 1          ' �i�v�e�g�U���ۏؕ��@�Q��
    HWFDLMIN As Integer             ' �i�v�e�g�U������
    HWFDLMAX As Integer             ' �i�v�e�g�U�����
    HWFNRHS As String * 1           ' �i�v�e�r�o�u�m�q�ۏؕ��@�Q��
    HWFNRKN As String * 1           ' �i�v�e�r�o�u�m�q�ۏؕ��@�Q��
    
    HWFOF1AX As Double              ' �i�v�e�n�r�e�P���Ϗ��
    HWFOF1MX As Double              ' �i�v�e�n�r�e�P���
    HWFOF1SH As String * 1          ' �i�v�e�n�r�e�P����ʒu�Q��
    HWFOF1ST As String * 1          ' �i�v�e�n�r�e�P����ʒu�Q�_
    HWFOF1SR As String * 1          ' �i�v�e�n�r�e�P����ʒu�Q��
    HWFOF1HT As String * 1          ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    HWFOF1HS As String * 1          ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    HWFOF2AX As Double              ' �i�v�e�n�r�e�Q���Ϗ��
    HWFOF2MX As Double              ' �i�v�e�n�r�e�Q���
    HWFOF2SH As String * 1          ' �i�v�e�n�r�e�Q����ʒu�Q��
    HWFOF2ST As String * 1          ' �i�v�e�n�r�e�Q����ʒu�Q�_
    HWFOF2SR As String * 1          ' �i�v�e�n�r�e�Q����ʒu�Q��
    HWFOF2HT As String * 1          ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    HWFOF2HS As String * 1          ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    HWFOF3AX As Double              ' �i�v�e�n�r�e�R���Ϗ��
    HWFOF3MX As Double              ' �i�v�e�n�r�e�R���
    HWFOF3SH As String * 1          ' �i�v�e�n�r�e�R����ʒu�Q��
    HWFOF3ST As String * 1          ' �i�v�e�n�r�e�R����ʒu�Q�_
    HWFOF3SR As String * 1          ' �i�v�e�n�r�e�R����ʒu�Q��
    HWFOF3HT As String * 1          ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    HWFOF3HS As String * 1          ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    HWFOF4AX As Double              ' �i�v�e�n�r�e�S���Ϗ��
    HWFOF4MX As Double              ' �i�v�e�n�r�e�S���
    HWFOF4SH As String * 1          ' �i�v�e�n�r�e�S����ʒu�Q��
    HWFOF4ST As String * 1          ' �i�v�e�n�r�e�S����ʒu�Q�_
    HWFOF4SR As String * 1          ' �i�v�e�n�r�e�S����ʒu�Q��
    HWFOF4HT As String * 1          ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    HWFOF4HS As String * 1          ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    HWFOSF1PTK As String * 1        ' �i�v�e�n�r�e�P�p�^���敪
    HWFOSF2PTK As String * 1        ' �i�v�e�n�r�e�Q�p�^���敪
    HWFOSF3PTK As String * 1        ' �i�v�e�n�r�e�R�p�^���敪
    HWFOSF4PTK As String * 1        ' �i�v�e�n�r�e�S�p�^���敪
    
    HWFBM1AN As Double              ' �i�v�e�a�l�c�P���ω���
    HWFBM1AX As Double              ' �i�v�e�a�l�c�P���Ϗ��
    HWFBM1SH As String * 1          ' �i�v�e�a�l�c�P����ʒu�Q��
    HWFBM1ST As String * 1          ' �i�v�e�a�l�c�P����ʒu�Q�_
    HWFBM1SR As String * 1          ' �i�v�e�a�l�c�P����ʒu�Q��
    HWFBM1HT As String * 1          ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    HWFBM1HS As String * 1          ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    HWFBM2AN As Double              ' �i�v�e�a�l�c�Q���ω���
    HWFBM2AX As Double              ' �i�v�e�a�l�c�Q���Ϗ��
    HWFBM2SH As String * 1          ' �i�v�e�a�l�c�Q����ʒu�Q��
    HWFBM2ST As String * 1          ' �i�v�e�a�l�c�Q����ʒu�Q�_
    HWFBM2SR As String * 1          ' �i�v�e�a�l�c�Q����ʒu�Q��
    HWFBM2HT As String * 1          ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    HWFBM2HS As String * 1          ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    HWFBM3AN As Double              ' �i�v�e�a�l�c�R���ω���
    HWFBM3AX As Double              ' �i�v�e�a�l�c�R���Ϗ��
    HWFBM3SH As String * 1          ' �i�v�e�a�l�c�R����ʒu�Q��
    HWFBM3ST As String * 1          ' �i�v�e�a�l�c�R����ʒu�Q�_
    HWFBM3SR As String * 1          ' �i�v�e�a�l�c�R����ʒu�Q��
    HWFBM3HT As String * 1          ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    HWFBM3HS As String * 1          ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    HWFBM1MBP As Single             ' �i�v�e�a�l�c�P�ʓ����z
    HWFBM2MBP As Single             ' �i�v�e�a�l�c�Q�ʓ����z
    HWFBM3MBP As Single             ' �i�v�e�a�l�c�R�ʓ����z
    HWFBM1MCL As String * 2         ' �i�v�e�a�l�c�P�ʓ��v�Z
    HWFBM2MCL As String * 2         ' �i�v�e�a�l�c�Q�ʓ��v�Z
    HWFBM3MCL As String * 2         ' �i�v�e�a�l�c�R�ʓ��v�Z
    
    HWFOS1NS As String * 2          ' �i�v�e�_�f�͏o�P�M�����@
    HWFOS2NS As String * 2          ' �i�v�e�_�f�͏o�Q�M�����@
    HWFOS3NS As String * 2          ' �i�v�e�_�f�͏o�R�M�����@
    HWFZONSW As String * 2          ' �i�v�e�c���_�f�M�����@
    HWFOF1NS As String * 2          ' �i�v�e�n�r�e�P�M�����@
    HWFOF2NS As String * 2          ' �i�v�e�n�r�e�Q�M�����@
    HWFOF3NS As String * 2          ' �i�v�e�n�r�e�R�M�����@
    HWFOF4NS As String * 2          ' �i�v�e�n�r�e�S�M�����@
    HWFBM1NS As String * 2          ' �i�v�e�a�l�c�P�M�����@
    HWFBM2NS As String * 2          ' �i�v�e�a�l�c�Q�M�����@
    HWFBM3NS As String * 2          ' �i�v�e�a�l�c�R�M�����@
    
    HWFANTIM As Integer             ' �i�v�e�`�m����
    HWFANTNP As Integer             ' �i�v�e�`�m���x

    HWFOF1ET As Integer             ' �i�v�e�n�r�e�P�I���d�s��
    HWFOF2ET As Integer             ' �i�v�e�n�r�e�Q�I���d�s��
    HWFOF3ET As Integer             ' �i�v�e�n�r�e�R�I���d�s��
    HWFOF4ET As Integer             ' �i�v�e�n�r�e�S�I���d�s��
    HWFBM1ET As Integer             ' �i�v�e�a�l�c�P�I���d�s��
    HWFBM2ET As Integer             ' �i�v�e�a�l�c�Q�I���d�s��
    HWFBM3ET As Integer             ' �i�v�e�a�l�c�R�I���d�s��

    HWFOF1SZ As String * 1          ' �i�v�e�n�r�e�P�������
    HWFOF2SZ As String * 1          ' �i�v�e�n�r�e�Q�������
    HWFOF3SZ As String * 1          ' �i�v�e�n�r�e�R�������
    HWFOF4SZ As String * 1          ' �i�v�e�n�r�e�S�������
    HWFBM1SZ As String * 1          ' �i�v�e�a�l�c�P�������
    HWFBM2SZ As String * 1          ' �i�v�e�a�l�c�Q�������
    HWFBM3SZ As String * 1          ' �i�v�e�a�l�c�R�������
    
    HWFDENKU As String * 1          ' �i�v�e�c���������L��
    HWFDENMX As Integer             ' �i�v�e�c�������
    HWFDENMN As Integer             ' �i�v�e�c��������
    HWFDENHT As String * 1          ' �i�v�e�c�����ۏؕ��@�Q��
    HWFDENHS As String * 1          ' �i�v�e�c�����ۏؕ��@�Q��
    HWFDVDKU As String * 1          ' �i�v�e�c�u�c�Q�����L��
    HWFDVDMXN As Integer            ' �i�v�e�c�u�c�Q���
    HWFDVDMNN As Integer            ' �i�v�e�c�u�c�Q����
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
    HWFGDLINE As Single             '�i�v�e�f�c���C����
    HWFRKHNN As String * 1          ' �i�v�e���R�����p�x�Q��
    HWFONKHN As String * 1          ' �i�v�e�_�f�Z�x�����p�x�Q��
    HWFOF1KN As String * 1          ' �i�v�e�n�r�e�P�����p�x�Q��
    HWFOF2KN As String * 1          ' �i�v�e�n�r�e�Q�����p�x�Q��
    HWFOF3KN As String * 1          ' �i�v�e�n�r�e�R�����p�x�Q��
    HWFOF4KN As String * 1          ' �i�v�e�n�r�e�S�����p�x�Q��
    HWFBM1KN As String * 1          ' �i�v�e�a�l�c�P�����p�x�Q��
    HWFBM2KN As String * 1          ' �i�v�e�a�l�c�Q�����p�x�Q��
    HWFBM3KN As String * 1          ' �i�v�e�a�l�c�R�����p�x�Q��
    HWFOS1KN As String * 1          ' �i�v�e�_�f�͏o�P�����p�x�Q��
    HWFOS2KN As String * 1          ' �i�v�e�_�f�͏o�Q�����p�x�Q��
    HWFOS3KN As String * 1          ' �i�v�e�_�f�͏o�R�����p�x�Q��
    HWFDSOKN As String * 1          ' �i�v�e�c�r�n�c�����p�x�Q��
    HWFMKKHN As String * 1          ' �i�v�e�����בw�����p�x�Q��
    HWFSPVKN As String * 1          ' �i�v�e�r�o�u�e�d�����p�x�Q��
    HWFDLKHN As String * 1          ' �i�v�e�g�U�������p�x�Q��
    HWFZOKHN As String * 1          ' �i�v�e�c���_�f�����p�x�Q��
    HWFGDKHN As String * 1          ' �i�v�e�f�c�����p�x�Q��
    BLOCKID() As String * 12        ' �u���b�NID

''SPV���菈���ǉ�
''�����̍\���̂ɍ��ڒǉ������VB�̐����Ɉ���������̂ŁA�ʂŊǗ�����B
''WF���i�d�l�擾�p(CMBC039�p)
    HWFSPVPUG As Double             ' �i�v�e�r�o�u�e�d�o�t�`��
    HWFSPVPUR As Double             ' �i�v�e�r�o�u�e�d�o�t�`��
    HWFSPVSTD As Double             ' �i�v�e�r�o�u�e�d�W���΍�
    HWFNRMX   As Double             ' �i�v�e�r�o�u�m�q���
    HWFNRPUG  As Double             ' �i�v�e�r�o�u�m�q�o�t�`��
    HWFNRPUR  As Double             ' �i�v�e�r�o�u�m�q�o�t�`��
    HWFNRSTD  As Double             ' �i�v�e�r�o�u�m�q�W���΍�
    HWFDLPUG  As Double             ' �i�v�e�g�U���o�t�`��
    HWFDLPUR  As Double             ' �i�v�e�g�U���o�t�`��
    HWFNRAM   As Double             ' �i�v�e�r�o�u�m�q����
    HWFNRSH   As String * 1         ' �i�v�e�r�o�u�m�q����ʒu_��
    HWFNRST   As String * 1         ' �i�v�e�r�o�u�m�q����ʒu_�_
    HWFNRHT   As String * 1         ' �i�v�e�r�o�u�m�q�ۏؕ��@_��
    HWFNRSI   As String * 1         ' �i�v�e�r�o�u�m�q����ʒu_��

' �G�s��s�]���ǉ��Ή�
    HEPHS      As Boolean           ' �G�s�d�l�t���O(�L:1,��:0)
    HEPANTNP   As Integer           ' �iEPAN���x
    HEPOF1AX   As Double            ' �iEPOSF1���Ϗ��
    HEPOF1MX   As Double            ' �iEPOSF1���
    HEPOF1ET   As Double            ' �iEPOSF1�I��ET��
    HEPOF1NS   As String * 2        ' �iEPOSF1�M�����@
    HEPOF1SZ   As String * 1        ' �iEPOSF1�������
    HEPOF1SH   As String * 1        ' �iEPOSF1����ʒu_��
    HEPOF1ST   As String * 1        ' �iEPOSF1����ʒu_�_
    HEPOF1SR   As String * 1        ' �iEPOSF1����ʒu_��
    HEPOF1HT   As String * 1        ' �iEPOSF1�ۏؕ��@_��
    HEPOF1HS   As String * 1        ' �iEPOSF1�ۏؕ��@_��
    HEPOF1KM   As String * 1        ' �iEPOSF1�����p�x_��
    HEPOF1KN   As String * 1        ' �iEPOSF1�����p�x_��
    HEPOF1KH   As String * 1        ' �iEPOSF1�����p�x_��
    HEPOF1KU   As String * 1        ' �iEPOSF1�����p�x_�
    HEPOSF1PTK As String * 1        ' �iEPOSF1���݋敪
    HEPOF2AX   As Double            ' �iEPOSF2���Ϗ��
    HEPOF2MX   As Double            ' �iEPOSF2���
    HEPOF2ET   As Double            ' �iEPOSF2�I��ET��
    HEPOF2NS   As String * 2        ' �iEPOSF2�M�����@
    HEPOF2SZ   As String * 1        ' �iEPOSF2�������
    HEPOF2SH   As String * 1        ' �iEPOSF2����ʒu_��
    HEPOF2ST   As String * 1        ' �iEPOSF2����ʒu_�_
    HEPOF2SR   As String * 1        ' �iEPOSF2����ʒu_��
    HEPOF2HT   As String * 1        ' �iEPOSF2�ۏؕ��@_��
    HEPOF2HS   As String * 1        ' �iEPOSF2�ۏؕ��@_��
    HEPOF2KM   As String * 1        ' �iEPOSF2�����p�x_��
    HEPOF2KN   As String * 1        ' �iEPOSF2�����p�x_��
    HEPOF2KH   As String * 1        ' �iEPOSF2�����p�x_��
    HEPOF2KU   As String * 1        ' �iEPOSF2�����p�x_�
    HEPOSF2PTK As String * 1        ' �iEPOSF2���݋敪
    HEPOF3AX   As Double            ' �iEPOSF3���Ϗ��
    HEPOF3MX   As Double            ' �iEPOSF3���
    HEPOF3ET   As Double            ' �iEPOSF3�I��ET��
    HEPOF3NS   As String * 2        ' �iEPOSF3�M�����@
    HEPOF3SZ   As String * 1        ' �iEPOSF3�������
    HEPOF3SH   As String * 1        ' �iEPOSF3����ʒu_��
    HEPOF3ST   As String * 1        ' �iEPOSF3����ʒu_�_
    HEPOF3SR   As String * 1        ' �iEPOSF3����ʒu_��
    HEPOF3HT   As String * 1        ' �iEPOSF3�ۏؕ��@_��
    HEPOF3HS   As String * 1        ' �iEPOSF3�ۏؕ��@_��
    HEPOF3KM   As String * 1        ' �iEPOSF3�����p�x_��
    HEPOF3KN   As String * 1        ' �iEPOSF3�����p�x_��
    HEPOF3KH   As String * 1        ' �iEPOSF3�����p�x_��
    HEPOF3KU   As String * 1        ' �iEPOSF3�����p�x_�
    HEPOSF3PTK As String * 1        ' �iEPOSF3���݋敪
    HEPBM1AN   As Double            ' �iEPBMD1���ω���
    HEPBM1AX   As Double            ' �iEPBMD1���Ϗ��
    HEPBM1ET   As Double            ' �iEPBMD1�I��ET��
    HEPBM1NS   As String * 2        ' �iEPBMD1�M�����@
    HEPBM1SZ   As String * 1        ' �iEPBMD1�������
    HEPBM1SH   As String * 1        ' �iEPBMD1����ʒu_��
    HEPBM1ST   As String * 1        ' �iEPBMD1����ʒu_�_
    HEPBM1SR   As String * 1        ' �iEPBMD1����ʒu_��
    HEPBM1HT   As String * 1        ' �iEPBMD1�ۏؕ��@_��
    HEPBM1HS   As String * 1        ' �iEPBMD1�ۏؕ��@_��
    HEPBM1KM   As String * 1        ' �iEPBMD1�����p�x_��
    HEPBM1KN   As String * 1        ' �iEPBMD1�����p�x_��
    HEPBM1KH   As String * 1        ' �iEPBMD1�����p�x_��
    HEPBM1KU   As String * 1        ' �iEPBMD1�����p�x_�
    HEPBM1MBP  As Double            ' �iEPBMD1�ʓ����z
    HEPBM1MCL  As String * 2        ' �iEPBMD1�ʓ��v�Z
    HEPBM2AN   As Double            ' �iEPBMD2���ω���
    HEPBM2AX   As Double            ' �iEPBMD2���Ϗ��
    HEPBM2ET   As Double            ' �iEPBMD2�I��ET��
    HEPBM2NS   As String * 2        ' �iEPBMD2�M�����@
    HEPBM2SZ   As String * 1        ' �iEPBMD2�������
    HEPBM2SH   As String * 1        ' �iEPBMD2����ʒu_��
    HEPBM2ST   As String * 1        ' �iEPBMD2����ʒu_�_
    HEPBM2SR   As String * 1        ' �iEPBMD2����ʒu_��
    HEPBM2HT   As String * 1        ' �iEPBMD2�ۏؕ��@_��
    HEPBM2HS   As String * 1        ' �iEPBMD2�ۏؕ��@_��
    HEPBM2KM   As String * 1        ' �iEPBMD2�����p�x_��
    HEPBM2KN   As String * 1        ' �iEPBMD2�����p�x_��
    HEPBM2KH   As String * 1        ' �iEPBMD2�����p�x_��
    HEPBM2KU   As String * 1        ' �iEPBMD2�����p�x_�
    HEPBM2MBP  As Double            ' �iEPBMD2�ʓ����z
    HEPBM2MCL  As String * 2        ' �iEPBMD2�ʓ��v�Z
    HEPBM3AN   As Double            ' �iEPBMD3���ω���
    HEPBM3AX   As Double            ' �iEPBMD3���Ϗ��
    HEPBM3GSAN As Double            ' �iEPBMD3���ω���(�O��)�@09/05/07 ooba
    HEPBM3GSAX As Double            ' �iEPBMD3���Ϗ��(�O��)�@09/05/07 ooba
    HEPBM3ET   As Double            ' �iEPBMD3�I��ET��
    HEPBM3NS   As String * 2        ' �iEPBMD3�M�����@
    HEPBM3SZ   As String * 1        ' �iEPBMD3�������
    HEPBM3SH   As String * 1        ' �iEPBMD3����ʒu_��
    HEPBM3ST   As String * 1        ' �iEPBMD3����ʒu_�_
    HEPBM3SR   As String * 1        ' �iEPBMD3����ʒu_��
    HEPBM3HT   As String * 1        ' �iEPBMD3�ۏؕ��@_��
    HEPBM3HS   As String * 1        ' �iEPBMD3�ۏؕ��@_��
    HEPBM3KM   As String * 1        ' �iEPBMD3�����p�x_��
    HEPBM3KN   As String * 1        ' �iEPBMD3�����p�x_��
    HEPBM3KH   As String * 1        ' �iEPBMD3�����p�x_��
    HEPBM3KU   As String * 1        ' �iEPBMD3�����p�x_�
    HEPBM3MBP  As Double            ' �iEPBMD3�ʓ����z
    HEPBM3MCL  As String * 2        ' �iEPBMD3�ʓ��v�Z

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP    As String * 1       ' DK���x�i�d�l�j
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    HSXLDLRMN   As Integer          ' �iSXL/DL�A��0����
    HSXLDLRMX   As Integer          ' �iSXL/DL�A��0���
    HWFLDLRMN   As Integer          ' �iWFL/DL�A��0����
    HWFLDLRMX   As Integer          ' �iWFL/DL�A��0���
    HWFGDPTK    As String * 1       ' �i�v�e�f�c�p�^���敪
    HSXGDPTK    As String * 1       ' �i�r�w�f�c�p�^���敪
    WFHSGDCW    As String * 1       ' �ۏ�FLG�iGD)
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    
'' ��Add 2008/10/01 SIRD�Ή� Y.Hitomi
    HWFSIRDMX   As Integer          ' ����]�ʏ��
    HWFSIRDSZ   As String * 1       ' ����]�ʑ������
    HWFSIRDHT   As String * 1       ' ����]�ʕۏؕ��@�Q��
    HWFSIRDHS   As String * 1       ' ����]�ʕۏؕ��@�Q��
    HWFSIRDKM   As String * 1       ' ����]�ʌ����p�x�Q��
    HWFSIRDKN   As String * 1       ' ����]�ʌ����p�x�Q��
    HWFSIRDKH   As String * 1       ' ����]�ʌ����p�x�Q��
    HWFSIRDKU   As String * 1       ' ����]�ʌ����p�x�Q�E
'' ��Add 2008/10/01 SIRD�Ή� Y.Hitomi
End Type

'*******************************************************************************************
'*    �֐���        : SetInitData
'*
'*    �����T�v      : 1.�����ݒ菈��
'*
'*    �p�����[�^    : �ϐ���      ,IO  ,�^                           ,����
'*                   intSXLID     ,I   ,String                       ,SXL-ID
'*                   udtNew_Hinban,I   ,tFullHinban                  ,�Y���i��(�\����)
'*                   udtSXL       ,O   ,DBDRV_scmzc_fcmlc001b_SXL    ,SXL�Ǘ��p
'*                   intSmpGetFlg ,I   ,Integer                      ,����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
'*                   sSamplID1    ,I   ,String                       ,TOP�����ID(�ȗ���)
'*                   sSamplID2    ,I   ,String                       ,BOT�����ID(�ȗ���)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function SetInitData(intSXLID As String, udtNew_Hinban As tFullHinban, udtSXL As DBDRV_scmzc_fcmlc001b_SXL, _
                            intSmpGetFlg As Integer, sSamplID1 As String, sSamplID2 As String) As FUNCTION_RETURN

    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim intRecCnt   As Integer
    Dim i           As Integer
    Dim intIngotpos As Integer              ' �������ʒu
    Dim intLength   As Integer              ' ����
    Dim sCryNum     As String               ' �����ԍ�
    
    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function SetInitData"

    SetInitData = FUNCTION_RETURN_SUCCESS

Debug.Print "1-1 " & Now & " SXL�Ǘ����擾 SQL���s"
    ' SXL�Ǘ����擾
    sSQL = "select "
    sSQL = sSQL & "xtalcb as CRYNUM, "      ' �����ԍ�
    sSQL = sSQL & "inposcb as INGOTPOS, "   ' �������J�n�ʒu
    sSQL = sSQL & "rlencb as LENGTH, "      ' ����
    sSQL = sSQL & "sxlidcb as SXLID, "      ' SXLID
    sSQL = sSQL & "' ' as KRPROCCD, "       ' �Ǘ��H��
    sSQL = sSQL & "gnwkntcb as NOWPROC, "   ' ���ݍH��
    sSQL = sSQL & "' ' as LPKRPROCCD, "     ' �ŏI�ʉߊǗ��H��
    sSQL = sSQL & "newkntcb as LASTPASS, "  ' �ŏI�ʉߍH��
    sSQL = sSQL & "livkcb as DELCLS, "      ' �폜�敪
    sSQL = sSQL & "lstccb as LSTATCLS, "    ' �ŏI��ԋ敪
    sSQL = sSQL & "sholdclscb HOLDCLS, "    ' �z�[���h�敪
    sSQL = sSQL & "hinbcb as HINBAN, "      ' �i��
    sSQL = sSQL & "revnumcb as REVNUM, "    ' ���i�ԍ������ԍ�
    sSQL = sSQL & "factorycb as FACTORY, "  ' �H��
    sSQL = sSQL & "opecb as OPECOND, "      ' ���Ə���
    sSQL = sSQL & "tdaycb as REGDATE, "     ' �o�^���t
    sSQL = sSQL & "kdaycb as UPDDATE "      ' �X�V���t
    sSQL = sSQL & " from XSDCB "
    sSQL = sSQL & " where sxlidcb = '" & intSXLID & "'"
    
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)

    ' ���R�[�h0��������I��
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        SetInitData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
Debug.Print "1-2 " & Now & " �z��ɃZ�b�g"
    With udtSXL
        .CRYNUM = rs("CRYNUM")                                          ' �����ԍ�
        .INGOTPOS = rs("INGOTPOS")                                      ' �������J�n�ʒu
        If IsNull(rs("LENGTH")) = False Then .Length = rs("LENGTH")     ' ����
        .SXLID = rs("SXLID")                                            ' SXLID
        .KRPROCCD = rs("KRPROCCD")                                      ' �Ǘ��H��
        .NOWPROC = rs("NOWPROC")                                        ' ���ݍH��
        .LPKRPROCCD = rs("LPKRPROCCD")                                  ' �ŏI�ʉߊǗ��H��
        .LASTPASS = rs("LASTPASS")                                      ' �ŏI�ʉߍH��
        .DELCLS = rs("DELCLS")                                          ' �폜�敪
        .LSTATCLS = rs("LSTATCLS")                                      ' �ŏI��ԋ敪
        If IsNull(rs("HOLDCLS")) = False Then .HOLDCLS = rs("HOLDCLS")  ' �z�[���h�敪
        .hinban = rs("HINBAN")                                          ' �i��
        .REVNUM = rs("REVNUM")                                          ' ���i�ԍ������ԍ�
        .factory = rs("FACTORY")                                        ' �H��
        .opecond = rs("OPECOND")                                        ' ���Ə���
        .REGDATE = rs("REGDATE")                                        ' �o�^���t
        .UPDDATE = rs("UPDDATE")                                        ' �X�V���t
    End With
    
    Set rs = Nothing
    
    ' �H�������ް��擾�֐������ް����擾���ݒ肷��
    If intSmpGetFlg <> 0 Then
        If GET_hurikaeC3(intSXLID, wiKcnt, intIngotpos, intLength, sCryNum) = FUNCTION_RETURN_FAILURE Then
            SetInitData = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
            
        With udtSXL
            .CRYNUM = sCryNum                                           ' �����ԍ�
            .INGOTPOS = intIngotpos                                     ' �������J�n�ʒu
            .Length = intLength                                         ' ����
        End With
    End If
        
Debug.Print "2-1 " & Now & " �V�T���v���Ǘ�(SXL)���擾 SQL���s"
    ' �V�T���v���Ǘ�(SXL)���擾
    ' �G�s��s�]���ǉ��Ή�
    sSQL = "select SXLIDCW, SMPKBNCW, TBKBNCW, REVNUMCW, XTALCW, INPOSCW, REPSMPLIDCW, HINBCW, FACTORYCW, OPECW, KTKBNCW, SMCRYNUMCW, "
    sSQL = sSQL & "WFSMPLIDRSCW, WFSMPLIDRS1CW, WFSMPLIDRS2CW, WFINDRSCW, WFRESRS1CW, WFRESRS2CW, WFSMPLIDOICW, WFINDOICW, WFRESOICW, "
    sSQL = sSQL & "WFSMPLIDB1CW, WFINDB1CW, WFRESB1CW, WFSMPLIDB2CW, WFINDB2CW, WFRESB2CW, WFSMPLIDB3CW, WFINDB3CW, WFRESB3CW, "
    sSQL = sSQL & "WFSMPLIDL1CW, WFINDL1CW, WFRESL1CW, WFSMPLIDL2CW, WFINDL2CW, WFRESL2CW, WFSMPLIDL3CW, WFINDL3CW, WFRESL3CW, "
    sSQL = sSQL & "WFSMPLIDL4CW, WFINDL4CW, WFRESL4CW, WFSMPLIDDSCW, WFINDDSCW, WFRESDSCW, WFSMPLIDDZCW, WFINDDZCW, WFRESDZCW, "
    sSQL = sSQL & "WFSMPLIDSPCW, WFINDSPCW, WFRESSPCW, WFSMPLIDDO1CW, WFINDDO1CW, WFRESDO1CW, WFSMPLIDDO2CW, WFINDDO2CW, WFRESDO2CW, "
    sSQL = sSQL & "WFSMPLIDDO3CW, WFINDDO3CW, WFRESDO3CW, WFSMPLIDOT1CW, WFINDOT1CW, WFRESOT1CW, WFSMPLIDOT2CW, WFINDOT2CW, WFRESOT2CW, "
    sSQL = sSQL & "WFSMPLIDAOICW , WFINDAOICW, WFRESAOICW, SMPLNUMCW, SMPLPATCW, TSTAFFCW, TDAYCW, KSTAFFCW, KDAYCW, SNDKCW, SNDDAYCW, "
    sSQL = sSQL & "WFSMPLIDGDCW, WFINDGDCW, WFRESGDCW, WFHSGDCW "
    sSQL = sSQL & ",EPSMPLIDB1CW, EPINDB1CW, EPRESB1CW, EPSMPLIDB2CW, EPINDB2CW, EPRESB2CW, EPSMPLIDB3CW, EPINDB3CW, EPRESB3CW, "
    sSQL = sSQL & "EPSMPLIDL1CW, EPINDL1CW, EPRESL1CW, EPSMPLIDL2CW, EPINDL2CW, EPRESL2CW, EPSMPLIDL3CW, EPINDL3CW, EPRESL3CW "
    sSQL = sSQL & "from XSDCW "
    
    If intSmpGetFlg = 0 Then        ' SXL-ID�Ō���(�����敪=��ۯ�)
        sSQL = sSQL & "where SXLIDCW = '" & intSXLID & "' and "
        sSQL = sSQL & "      LIVKCW = '0' "
    Else                            ' �����ԍ��ƻ����ID�Ō���
        sSQL = sSQL & "where XTALCW = substr('" & intSXLID & "', 1, 9) || '000' and "
        sSQL = sSQL & "      REPSMPLIDCW in ('" & sSamplID1 & "', '" & sSamplID2 & "') "
    End If
    sSQL = sSQL & "order by INPOSCW"
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)

    ' ���R�[�h0��������I��
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        SetInitData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

Debug.Print "2-2 " & Now & " �z��ɃZ�b�g"
    intRecCnt = rs.RecordCount
    ReDim udtSXL.WFSMP(intRecCnt)
    For i = 1 To intRecCnt
        With udtSXL.WFSMP(i)
            .SXLIDCW = rs("SXLIDCW")                                                            ' SXL-ID
            .SMPKBNCW = rs("SMPKBNCW")                                                          ' �T���v���敪
            .TBKBNCW = rs("TBKBNCW")                                                            ' T/B�敪
            
            If IsNull(rs("REPSMPLIDCW")) = False Then .REPSMPLIDCW = rs("REPSMPLIDCW")          ' ��\�T���v��ID
            If IsNull(rs("XTALCW")) = False Then .XTALCW = rs("XTALCW")                         ' �����ԍ�
            If IsNull(rs("INPOSCW")) = False Then .INPOSCW = rs("INPOSCW")                      ' �������ʒu
            If IsNull(rs("HINBCW")) = False Then .HINBCW = rs("HINBCW")                         ' �i��
            If IsNull(rs("REVNUMCW")) = False Then .REVNUMCW = rs("REVNUMCW")                   ' ���i�ԍ������ԍ�
            If IsNull(rs("FACTORYCW")) = False Then .FACTORYCW = rs("FACTORYCW")                ' �H��
            If IsNull(rs("OPECW")) = False Then .OPECW = rs("OPECW")                            ' ���Ə���
            If IsNull(rs("KTKBNCW")) = False Then .KTKBNCW = rs("KTKBNCW")                      ' �m��敪
            If IsNull(rs("SMCRYNUMCW")) = False Then .SMCRYNUMCW = rs("SMCRYNUMCW")             ' �������ۯ�ID
            If IsNull(rs("WFSMPLIDRSCW")) = False Then .WFSMPLIDRSCW = rs("WFSMPLIDRSCW")       ' �T���v��ID(Rs)
            If IsNull(rs("WFSMPLIDRS1CW")) = False Then .WFSMPLIDRS1CW = rs("WFSMPLIDRS1CW")    ' ����T���v��ID1(Rs)
            If IsNull(rs("WFSMPLIDRS2CW")) = False Then .WFSMPLIDRS2CW = rs("WFSMPLIDRS2CW")    ' ����T���v��ID2(Rs)
            If IsNull(rs("WFINDRSCW")) = False Then .WFINDRSCW = rs("WFINDRSCW")                ' ���FLG(Rs)
            If IsNull(rs("WFRESRS1CW")) = False Then .WFRESRS1CW = rs("WFRESRS1CW")             ' ����FLG1(Rs)
            If IsNull(rs("WFRESRS2CW")) = False Then .WFRESRS2CW = rs("WFRESRS2CW")             ' ����FLG2(Rs)
            If IsNull(rs("WFSMPLIDOICW")) = False Then .WFSMPLIDOICW = rs("WFSMPLIDOICW")       ' �T���v��ID(Oi)
            If IsNull(rs("WFINDOICW")) = False Then .WFINDOICW = rs("WFINDOICW")                ' ���FLG(Oi)
            If IsNull(rs("WFRESOICW")) = False Then .WFRESOICW = rs("WFRESOICW")                ' ����FLG(Oi)
            If IsNull(rs("WFSMPLIDB1CW")) = False Then .WFSMPLIDB1CW = rs("WFSMPLIDB1CW")       ' �T���v��ID(B1)
            If IsNull(rs("WFINDB1CW")) = False Then .WFINDB1CW = rs("WFINDB1CW")                ' ���FLG(B1)
            If IsNull(rs("WFRESB1CW")) = False Then .WFRESB1CW = rs("WFRESB1CW")                ' ����FLG(B1)
            If IsNull(rs("WFSMPLIDB2CW")) = False Then .WFSMPLIDB2CW = rs("WFSMPLIDB2CW")       ' �T���v��ID(B2)
            If IsNull(rs("WFINDB2CW")) = False Then .WFINDB2CW = rs("WFINDB2CW")                ' ���FLG(B2)
            If IsNull(rs("WFRESB2CW")) = False Then .WFRESB2CW = rs("WFRESB2CW")                ' ����FLG(B2)
            If IsNull(rs("WFSMPLIDB3CW")) = False Then .WFSMPLIDB3CW = rs("WFSMPLIDB3CW")       ' �T���v��ID(B3)
            If IsNull(rs("WFINDB3CW")) = False Then .WFINDB3CW = rs("WFINDB3CW")                ' ���FLG(B3)
            If IsNull(rs("WFRESB3CW")) = False Then .WFRESB3CW = rs("WFRESB3CW")                ' ����FLG(B3)
            If IsNull(rs("WFSMPLIDL1CW")) = False Then .WFSMPLIDL1CW = rs("WFSMPLIDL1CW")       ' �T���v��ID(L1)
            If IsNull(rs("WFINDL1CW")) = False Then .WFINDL1CW = rs("WFINDL1CW")                ' ���FLG(L1)
            If IsNull(rs("WFRESL1CW")) = False Then .WFRESL1CW = rs("WFRESL1CW")                ' ����FLG(L1)
            If IsNull(rs("WFSMPLIDL2CW")) = False Then .WFSMPLIDL2CW = rs("WFSMPLIDL2CW")       ' �T���v��ID(L2)
            If IsNull(rs("WFINDL2CW")) = False Then .WFINDL2CW = rs("WFINDL2CW")                ' ���FLG(L2)
            If IsNull(rs("WFRESL2CW")) = False Then .WFRESL2CW = rs("WFRESL2CW")                ' ����FLG(L2)
            If IsNull(rs("WFSMPLIDL3CW")) = False Then .WFSMPLIDL3CW = rs("WFSMPLIDL3CW")       ' �T���v��ID(L3)
            If IsNull(rs("WFINDL3CW")) = False Then .WFINDL3CW = rs("WFINDL3CW")                ' ���FLG(L3)
            If IsNull(rs("WFRESL3CW")) = False Then .WFRESL3CW = rs("WFRESL3CW")                ' ����FLG(L3)
            If IsNull(rs("WFSMPLIDL4CW")) = False Then .WFSMPLIDL4CW = rs("WFSMPLIDL4CW")       ' �T���v��ID(L4)
            If IsNull(rs("WFINDL4CW")) = False Then .WFINDL4CW = rs("WFINDL4CW")                ' ���FLG(L4)
            If IsNull(rs("WFRESL4CW")) = False Then .WFRESL4CW = rs("WFRESL4CW")                ' ����FLG(L4)
            If IsNull(rs("WFSMPLIDDSCW")) = False Then .WFSMPLIDDSCW = rs("WFSMPLIDDSCW")       ' �T���v��ID(DS)
            If IsNull(rs("WFINDDSCW")) = False Then .WFINDDSCW = rs("WFINDDSCW")                ' ���FLG(DS)
            If IsNull(rs("WFRESDSCW")) = False Then .WFRESDSCW = rs("WFRESDSCW")                ' ����FLG(DS)
            If IsNull(rs("WFSMPLIDDZCW")) = False Then .WFSMPLIDDZCW = rs("WFSMPLIDDZCW")       ' �T���v��ID(DZ)
            If IsNull(rs("WFINDDZCW")) = False Then .WFINDDZCW = rs("WFINDDZCW")                ' ���FLG(DZ)
            If IsNull(rs("WFRESDZCW")) = False Then .WFRESDZCW = rs("WFRESDZCW")                ' ����FLG(DZ)
            If IsNull(rs("WFSMPLIDSPCW")) = False Then .WFSMPLIDSPCW = rs("WFSMPLIDSPCW")       ' �T���v��ID(SP)
            If IsNull(rs("WFINDSPCW")) = False Then .WFINDSPCW = rs("WFINDSPCW")                ' ���FLG(SP)
            If IsNull(rs("WFRESSPCW")) = False Then .WFRESSPCW = rs("WFRESSPCW")                ' ����FLG(SP)
            If IsNull(rs("WFSMPLIDDO1CW")) = False Then .WFSMPLIDDO1CW = rs("WFSMPLIDDO1CW")    ' �T���v��ID(DO1)
            If IsNull(rs("WFINDDO1CW")) = False Then .WFINDDO1CW = rs("WFINDDO1CW")             ' ���FLG(DO1)
            If IsNull(rs("WFRESDO1CW")) = False Then .WFRESDO1CW = rs("WFRESDO1CW")             ' ����FLG(DO1)
            If IsNull(rs("WFSMPLIDDO2CW")) = False Then .WFSMPLIDDO2CW = rs("WFSMPLIDDO2CW")    ' �T���v��ID(DO2)
            If IsNull(rs("WFINDDO2CW")) = False Then .WFINDDO2CW = rs("WFINDDO2CW")             ' ���FLG(DO2)
            If IsNull(rs("WFRESDO2CW")) = False Then .WFRESDO2CW = rs("WFRESDO2CW")             ' ����FLG(DO2)
            If IsNull(rs("WFSMPLIDDO3CW")) = False Then .WFSMPLIDDO3CW = rs("WFSMPLIDDO3CW")    ' �T���v��ID(DO3)
            If IsNull(rs("WFINDDO3CW")) = False Then .WFINDDO3CW = rs("WFINDDO3CW")             ' ���FLG(DO3)
            If IsNull(rs("WFRESDO3CW")) = False Then .WFRESDO3CW = rs("WFRESDO3CW")             ' ����FLG(DO3)
            If IsNull(rs("WFSMPLIDOT1CW")) = False Then .WFSMPLIDOT1CW = rs("WFSMPLIDOT1CW")    ' �T���v��ID(OT1)
            If IsNull(rs("WFINDOT1CW")) = False Then .WFINDOT1CW = rs("WFINDOT1CW")             ' ���FLG(OT1)
            If IsNull(rs("WFRESOT1CW")) = False Then .WFRESOT1CW = rs("WFRESOT1CW")             ' ����FLG(OT1)
            If IsNull(rs("WFSMPLIDOT2CW")) = False Then .WFSMPLIDOT2CW = rs("WFSMPLIDOT2CW")    ' �T���v��ID(OT2)
            If IsNull(rs("WFINDOT2CW")) = False Then .WFINDOT2CW = rs("WFINDOT2CW")             ' ���FLG(OT2)
            If IsNull(rs("WFRESOT2CW")) = False Then .WFRESOT2CW = rs("WFRESOT2CW")             ' ����FLG(OT2)
            If IsNull(rs("WFSMPLIDAOICW")) = False Then .WFSMPLIDAOICW = rs("WFSMPLIDAOICW")    ' �T���v��ID(AOI)
            If IsNull(rs("WFINDAOICW")) = False Then .WFINDAOICW = rs("WFINDAOICW")             ' ���FLG(AOI)
            If IsNull(rs("WFRESAOICW")) = False Then .WFRESAOICW = rs("WFRESAOICW")             ' ����FLG(AOI)
            If IsNull(rs("SMPLNUMCW")) = False Then .SMPLNUMCW = rs("SMPLNUMCW")                ' ����ٖ���
            If IsNull(rs("SMPLPATCW")) = False Then .SMPLPATCW = rs("SMPLPATCW")                ' ����������
            If IsNull(rs("TSTAFFCW")) = False Then .TSTAFFCW = rs("TSTAFFCW")                   ' �o�^�Ј�ID
            If IsNull(rs("TDAYCW")) = False Then .TDAYCW = rs("TDAYCW")                         ' �o�^���t
            If IsNull(rs("KSTAFFCW")) = False Then .KSTAFFCW = rs("KSTAFFCW")                   ' �X�V�Ј�ID
            If IsNull(rs("KDAYCW")) = False Then .KDAYCW = rs("KDAYCW")                         ' �X�V���t
            If IsNull(rs("SNDKCW")) = False Then .SNDKCW = rs("SNDKCW")                         ' ���M�׸�
            If IsNull(rs("SNDDAYCW")) = False Then .SNDDAYCW = rs("SNDDAYCW")                   ' ���M���t
            If IsNull(rs("WFSMPLIDGDCW")) = False Then .WFSMPLIDGDCW = rs("WFSMPLIDGDCW")       ' �T���v��ID(GD)
            If IsNull(rs("WFINDGDCW")) = False Then .WFINDGDCW = rs("WFINDGDCW")                ' ���FLG(GD)
            If IsNull(rs("WFRESGDCW")) = False Then .WFRESGDCW = rs("WFRESGDCW")                ' ����FLG(GD)
            If IsNull(rs("WFHSGDCW")) = False Then .WFHSGDCW = rs("WFHSGDCW")                   ' �ۏ�FLG(GD)
            
            ' �G�s��s�]���ǉ�
            If IsNull(rs("EPSMPLIDB1CW")) = False Then .EPSMPLIDB1CW = rs("EPSMPLIDB1CW")       ' �T���v��ID(B1E)
            If IsNull(rs("EPINDB1CW")) = False Then .EPINDB1CW = rs("EPINDB1CW")                ' ���FLG(B1E)
            If IsNull(rs("EPRESB1CW")) = False Then .EPRESB1CW = rs("EPRESB1CW")                ' ����FLG(B1E)
            If IsNull(rs("EPSMPLIDB2CW")) = False Then .EPSMPLIDB2CW = rs("EPSMPLIDB2CW")       ' �T���v��ID(B2E)
            If IsNull(rs("EPINDB2CW")) = False Then .EPINDB2CW = rs("EPINDB2CW")                ' ���FLG(B2E)
            If IsNull(rs("EPRESB2CW")) = False Then .EPRESB2CW = rs("EPRESB2CW")                ' ����FLG(B2E)
            If IsNull(rs("EPSMPLIDB3CW")) = False Then .EPSMPLIDB3CW = rs("EPSMPLIDB3CW")       ' �T���v��ID(B3E)
            If IsNull(rs("EPINDB3CW")) = False Then .EPINDB3CW = rs("EPINDB3CW")                ' ���FLG(B3E)
            If IsNull(rs("EPRESB3CW")) = False Then .EPRESB3CW = rs("EPRESB3CW")                ' ����FLG(B3E)
            If IsNull(rs("EPSMPLIDL1CW")) = False Then .EPSMPLIDL1CW = rs("EPSMPLIDL1CW")       ' �T���v��ID(L1E)
            If IsNull(rs("EPINDL1CW")) = False Then .EPINDL1CW = rs("EPINDL1CW")                ' ���FLG(L1E)
            If IsNull(rs("EPRESL1CW")) = False Then .EPRESL1CW = rs("EPRESL1CW")                ' ����FLG(L1E)
            If IsNull(rs("EPSMPLIDL2CW")) = False Then .EPSMPLIDL2CW = rs("EPSMPLIDL2CW")       ' �T���v��ID(L2E)
            If IsNull(rs("EPINDL2CW")) = False Then .EPINDL2CW = rs("EPINDL2CW")                ' ���FLG(L2E)
            If IsNull(rs("EPRESL2CW")) = False Then .EPRESL2CW = rs("EPRESL2CW")                ' ����FLG(L2E)
            If IsNull(rs("EPSMPLIDL3CW")) = False Then .EPSMPLIDL3CW = rs("EPSMPLIDL3CW")       ' �T���v��ID(L3E)
            If IsNull(rs("EPINDL3CW")) = False Then .EPINDL3CW = rs("EPINDL3CW")                ' ���FLG(L3E)
            If IsNull(rs("EPRESL3CW")) = False Then .EPRESL3CW = rs("EPRESL3CW")                ' ����FLG(L3E)
        End With
        rs.MoveNext
    Next i
    
    Set rs = Nothing

Debug.Print "3 " & Now & " �����L�����擾"
    ' �������擾
    If KeturakuInfo(udtSXL) = FUNCTION_RETURN_FAILURE Then
        SetInitData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

Debug.Print "4 " & Now & " �������擾"
    ' �����擾
    If GetMaisu(udtSXL) = FUNCTION_RETURN_FAILURE Then
        SetInitData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
Debug.Print "8 " & Now

    SetInitData = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
'    gErr.Pop
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    SetInitData = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        :
'*
'*    �����T�v      : 1.�����L���擾
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                        ,����
'*                    udtSXL        ,IO ,DBDRV_scmzc_fcmlc001b_SXL ,SXL�Ǘ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function KeturakuInfo(udtSXL As DBDRV_scmzc_fcmlc001b_SXL) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim lngRecCnt   As Long
    Dim sSXLID      As String

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function KeturakuInfo"

    KeturakuInfo = FUNCTION_RETURN_SUCCESS

    sSQL = "select distinct SXL.SXLIDCB "
    sSQL = sSQL & "from XSDCB SXL, TBCME040 BLK, TBCMY012 REJ "
    sSQL = sSQL & "where"
    sSQL = sSQL & "  REJ.LOTID=BLK.BLOCKID"
    sSQL = sSQL & "  and SXL.SXLIDCB = '" & udtSXL.SXLID & "'"
    sSQL = sSQL & "  and SXL.XTALCB=BLK.CRYNUM"
    sSQL = sSQL & "  and SXL.LIVKCB<>'1'"
    sSQL = sSQL & "  and ("
    sSQL = sSQL & "    ("
    sSQL = sSQL & "      REJ.ALLSCRAP='Y'"
    sSQL = sSQL & "      and SXL.INPOSCB<BLK.INGOTPOS+BLK.LENGTH"
    sSQL = sSQL & "      and SXL.INPOSCB+SXL.RLENCB>BLK.INGOTPOS"
    sSQL = sSQL & "    ) or ("
    sSQL = sSQL & "      REJ.ALLSCRAP='N'"
    sSQL = sSQL & "      and REJ.REJCAT='A'"
    sSQL = sSQL & "      and (SXL.INPOSCB < BLK.INGOTPOS + REJ.LENTO)"
    sSQL = sSQL & "      and (SXL.INPOSCB + SXL.RLENCB > BLK.INGOTPOS + REJ.LENFROM)"
    sSQL = sSQL & "    ) or ("
    sSQL = sSQL & "      REJ.REJCAT='B'"
    sSQL = sSQL & "      and BLK.INGOTPOS + REJ.TOP_POS/10.0 between SXL.INPOSCB and SXL.INPOSCB + SXL.RLENCB"
    sSQL = sSQL & "    )"
    sSQL = sSQL & "  )"
    
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    ' SQL���ʂ�SXLID�����������SXLID
    If rs.RecordCount = 0 Then
        udtSXL.KETURAKU = False
    Else
        udtSXL.KETURAKU = True
    End If
    Set rs = Nothing

proc_exit:
    ' �I��
'    gErr.Pop
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    KeturakuInfo = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : GetMaisu
'*
'*    �����T�v      : 1.WF�����擾
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^                        ,����
'*                    udtSXL        ,IO ,DBDRV_scmzc_fcmlc001b_SXL ,SXL�Ǘ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function GetMaisu(udtSXL As DBDRV_scmzc_fcmlc001b_SXL) As FUNCTION_RETURN
    Dim sSQL    As String
    Dim rs      As OraDynaset

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function getMaisu"

    GetMaisu = FUNCTION_RETURN_SUCCESS

    sSQL = sSQL & "select MAICB from XSDCB where SXLIDCB = '" & udtSXL.SXLID & "'"
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        udtSXL.COUNT = 0
    Else
        udtSXL.COUNT = rs("MAICB")
    End If
    Set rs = Nothing

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    GetMaisu = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : funWfcGetDataEtc
'*
'*    �����T�v      : 1.WF�������� �e��f�[�^�擾
'*
'*    �p�����[�^    : �ϐ���      ,IO  ,�^                                 ,����
'*               �@�@udtTypIn     ,I   ,type_DBDRV_scmzc_fcmlc001c_In      ,���͗p
'*               �@�@Siyou        ,O   ,type_DBDRV_scmzc_fcmlc001c_Siyou   ,WF�d�l�p
'*               �@�@udtSokutei   ,O   ,typ_TBCMY013                       ,����]������
'*               �@�@sErrMsg �@�@ ,O   ,String    �@�@�@�@�@�@�@�@�@�@�@   ,�G���[���b�Z�[�W
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function funWfcGetDataEtc(udtTypIn As type_DBDRV_scmzc_fcmlc001c_In, udtNew_Hinban As tFullHinban, intSmpGetFlg As Integer, _
                                 udtSiyou As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                                 udtSokutei() As typ_TBCMY013, _
                                 sErrMsg As String) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim intRecCnt   As Integer
    Dim i           As Long
    Dim sDBName     As String
    Dim intPos      As Integer      ' ����وʒu
    
    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funWfcGetDataEtc"

    funWfcGetDataEtc = FUNCTION_RETURN_SUCCESS
    
    ' WF�d�l�擾
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
    sSQL = sSQL & "E021HWFRMCAL, "           ' �i�v�e���R�ʓ��v�Z
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
    sSQL = sSQL & "E025HWFONMCL, "           ' �i�v�e�_�f�Z�x�ʓ��v�Z
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
    
    sSQL = sSQL & "E029HWFOF1AX, "           ' �i�v�e�n�r�e�P���Ϗ��
    sSQL = sSQL & "E029HWFOF1MX, "           ' �i�v�e�n�r�e�P���
    sSQL = sSQL & "E029HWFOF1SH, "           ' �i�v�e�n�r�e�P����ʒu�Q��
    sSQL = sSQL & "E029HWFOF1ST, "           ' �i�v�e�n�r�e�P����ʒu�Q�_
    sSQL = sSQL & "E029HWFOF1SR, "           ' �i�v�e�n�r�e�P����ʒu�Q��
    sSQL = sSQL & "E029HWFOF1HT, "           ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF1HS, "           ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF2AX, "           ' �i�v�e�n�r�e�Q���Ϗ��
    sSQL = sSQL & "E029HWFOF2MX, "           ' �i�v�e�n�r�e�Q���
    sSQL = sSQL & "E029HWFOF2SH, "           ' �i�v�e�n�r�e�Q����ʒu�Q��
    sSQL = sSQL & "E029HWFOF2ST, "           ' �i�v�e�n�r�e�Q����ʒu�Q�_
    sSQL = sSQL & "E029HWFOF2SR, "           ' �i�v�e�n�r�e�Q����ʒu�Q��
    sSQL = sSQL & "E029HWFOF2HT, "           ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF2HS, "           ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF3AX, "           ' �i�v�e�n�r�e�R���Ϗ��
    sSQL = sSQL & "E029HWFOF3MX, "           ' �i�v�e�n�r�e�R���
    sSQL = sSQL & "E029HWFOF3SH, "           ' �i�v�e�n�r�e�R����ʒu�Q��
    sSQL = sSQL & "E029HWFOF3ST, "           ' �i�v�e�n�r�e�R����ʒu�Q�_
    sSQL = sSQL & "E029HWFOF3SR, "           ' �i�v�e�n�r�e�R����ʒu�Q��
    sSQL = sSQL & "E029HWFOF3HT, "           ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF3HS, "           ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF4AX, "           ' �i�v�e�n�r�e�S���Ϗ��
    sSQL = sSQL & "E029HWFOF4MX, "           ' �i�v�e�n�r�e�S���
    sSQL = sSQL & "E029HWFOF4SH, "           ' �i�v�e�n�r�e�S����ʒu�Q��
    sSQL = sSQL & "E029HWFOF4ST, "           ' �i�v�e�n�r�e�S����ʒu�Q�_
    sSQL = sSQL & "E029HWFOF4SR, "           ' �i�v�e�n�r�e�S����ʒu�Q��
    sSQL = sSQL & "E029HWFOF4HT, "           ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOF4HS, "           ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM1AN, "           ' �i�v�e�a�l�c�P���ω���
    sSQL = sSQL & "E029HWFBM1AX, "           ' �i�v�e�a�l�c�P���Ϗ��
    sSQL = sSQL & "E029HWFBM1SH, "           ' �i�v�e�a�l�c�P����ʒu�Q��
    sSQL = sSQL & "E029HWFBM1ST, "           ' �i�v�e�a�l�c�P����ʒu�Q�_
    sSQL = sSQL & "E029HWFBM1SR, "           ' �i�v�e�a�l�c�P����ʒu�Q��
    sSQL = sSQL & "E029HWFBM1HT, "           ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM1HS, "           ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM2AN, "           ' �i�v�e�a�l�c�Q���ω���
    sSQL = sSQL & "E029HWFBM2AX, "           ' �i�v�e�a�l�c�Q���Ϗ��
    sSQL = sSQL & "E029HWFBM2SH, "           ' �i�v�e�a�l�c�Q����ʒu�Q��
    sSQL = sSQL & "E029HWFBM2ST, "           ' �i�v�e�a�l�c�Q����ʒu�Q�_
    sSQL = sSQL & "E029HWFBM2SR, "           ' �i�v�e�a�l�c�Q����ʒu�Q��
    sSQL = sSQL & "E029HWFBM2HT, "           ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM2HS, "           ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM3AN, "           ' �i�v�e�a�l�c�R���ω���
    sSQL = sSQL & "E029HWFBM3AX, "           ' �i�v�e�a�l�c�R���Ϗ��
    sSQL = sSQL & "E029HWFBM3SH, "           ' �i�v�e�a�l�c�R����ʒu�Q��
    sSQL = sSQL & "E029HWFBM3ST, "           ' �i�v�e�a�l�c�R����ʒu�Q�_
    sSQL = sSQL & "E029HWFBM3SR, "           ' �i�v�e�a�l�c�R����ʒu�Q��
    sSQL = sSQL & "E029HWFBM3HT, "           ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFBM3HS, "           ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    sSQL = sSQL & "E029HWFOSF1PTK, "         ' �i�v�e�n�r�e�P�p�^���敪
    sSQL = sSQL & "E029HWFOSF2PTK, "         ' �i�v�e�n�r�e�Q�p�^���敪
    sSQL = sSQL & "E029HWFOSF3PTK, "         ' �i�v�e�n�r�e�R�p�^���敪
    sSQL = sSQL & "E029HWFOSF4PTK, "         ' �i�v�e�n�r�e�S�p�^���敪
    sSQL = sSQL & "E029HWFBM1MBP, "          ' �i�v�e�a�l�c�P�ʓ����z
    sSQL = sSQL & "E029HWFBM2MBP, "          ' �i�v�e�a�l�c�Q�ʓ����z
    sSQL = sSQL & "E029HWFBM3MBP, "          ' �i�v�e�a�l�c�R�ʓ����z
    sSQL = sSQL & "E029HWFBM1MCL, "          ' �i�v�e�a�l�c�P�ʓ��v�Z
    sSQL = sSQL & "E029HWFBM2MCL, "          ' �i�v�e�a�l�c�Q�ʓ��v�Z
    sSQL = sSQL & "E029HWFBM3MCL, "          ' �i�v�e�a�l�c�R�ʓ��v�Z
    sSQL = sSQL & "E025HWFOS1NS, "           ' �i�v�e�_�f�͏o�P�M�����@
    sSQL = sSQL & "E025HWFOS2NS, "           ' �i�v�e�_�f�͏o�Q�M�����@
    sSQL = sSQL & "E025HWFOS3NS, "           ' �i�v�e�_�f�͏o�R�M�����@
    
    sSQL = sSQL & "E029HWFOF1NS, "           ' �i�v�e�n�r�e�P�M�����@
    sSQL = sSQL & "E029HWFOF2NS, "           ' �i�v�e�n�r�e�Q�M�����@
    sSQL = sSQL & "E029HWFOF3NS, "           ' �i�v�e�n�r�e�R�M�����@
    sSQL = sSQL & "E029HWFOF4NS, "           ' �i�v�e�n�r�e�S�M�����@
    
    sSQL = sSQL & "E029HWFBM1NS, "           ' �i�v�e�a�l�c�P�M�����@
    sSQL = sSQL & "E029HWFBM2NS, "           ' �i�v�e�a�l�c�Q�M�����@
    sSQL = sSQL & "E029HWFBM3NS, "           ' �i�v�e�a�l�c�R�M�����@

    sSQL = sSQL & "E025HWFANTIM, "           ' �i�v�e�`�m����
    sSQL = sSQL & "E025HWFANTNP, "           ' �i�v�e�`�m���x

    sSQL = sSQL & "E029HWFOF1ET, "           ' �i�v�e�n�r�e�P�I���d�s��
    sSQL = sSQL & "E029HWFOF2ET, "           ' �i�v�e�n�r�e�Q�I���d�s��
    sSQL = sSQL & "E029HWFOF3ET, "           ' �i�v�e�n�r�e�R�I���d�s��
    sSQL = sSQL & "E029HWFOF4ET, "           ' �i�v�e�n�r�e�S�I���d�s��
    sSQL = sSQL & "E029HWFBM1ET, "           ' �i�v�e�a�l�c�P�I���d�s��
    sSQL = sSQL & "E029HWFBM2ET, "           ' �i�v�e�a�l�c�Q�I���d�s��
    sSQL = sSQL & "E029HWFBM3ET, "           ' �i�v�e�a�l�c�R�I���d�s��

    sSQL = sSQL & "E029HWFOF1SZ, "           ' �i�v�e�n�r�e�P�������
    sSQL = sSQL & "E029HWFOF2SZ, "           ' �i�v�e�n�r�e�Q�������
    sSQL = sSQL & "E029HWFOF3SZ, "           ' �i�v�e�n�r�e�R�������
    sSQL = sSQL & "E029HWFOF4SZ, "           ' �i�v�e�n�r�e�S�������
    sSQL = sSQL & "E029HWFBM1SZ, "           ' �i�v�e�a�l�c�P�������
    sSQL = sSQL & "E029HWFBM2SZ, "           ' �i�v�e�a�l�c�Q�������
    sSQL = sSQL & "E029HWFBM3SZ, "           ' �i�v�e�a�l�c�R�������
    sSQL = sSQL & "E028HWFSPVAM "            ' �i�v�e�r�o�u���Ϗ��
    
    '' SPV9�_�Ή�  ������2����3���ύX�Ή�
    sSQL = sSQL & ",E028HWFSPVMXN"
    sSQL = sSQL & ",E028HWFSPVAMN"
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sSQL = sSQL & ",NVL(E36.HSXDKTMP, ' ') HSXDKTMP"
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    sSQL = sSQL & ",HSXLDLRMN"              ' �iSXL/DL�A��0����
    sSQL = sSQL & ",HSXLDLRMX"              ' �iSXL/DL�A��0���
    sSQL = sSQL & ",HWFLDLRMN"              ' �iWFL/DL�A��0����
    sSQL = sSQL & ",HWFLDLRMX"              ' �iWFL/DL�A��0���
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    

'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'    sSql = sSql & " from VECME001"
    sSQL = sSQL & " from VECME001, TBCME036 E36"
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
    
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where E018HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " E018MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " E018FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " E018OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where E018HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " E018MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " E018FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " E018OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sSQL = sSQL & " AND E36.HINBAN = E018HINBAN"
    sSQL = sSQL & " AND E36.MNOREVNO = E018MNOREVNO"
    sSQL = sSQL & " AND E36.FACTORY = E018FACTORY"
    sSQL = sSQL & " AND E36.OPECOND = E018OPECOND"
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
    Debug.Print sSQL
    
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' ���R�[�h0���̓G���[�I��
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With udtSiyou
        .HWFTYPE = rs("E021HWFTYPE")                    ' �i�v�e�^�C�v
        .HWFCDIR = rs("E022HWFCDIR")                    ' �i�v�e�����ʕ�
        .HWFCDOP = rs("E023HWFCDOP")                    ' �i�v�e�����h�[�v

        .HWFRMIN = fncNullCheck(rs("E021HWFRMIN"))      ' �i�v�e���R����
        .HWFRMAX = fncNullCheck(rs("E021HWFRMAX"))      ' �i�v�e���R���
        .HWFRSPOH = rs("E021HWFRSPOH")                  ' �i�v�e���R����ʒu�Q��
        .HWFRSPOT = rs("E021HWFRSPOT")                  ' �i�v�e���R����ʒu�Q�_
        .HWFRSPOI = rs("E021HWFRSPOI")                  ' �i�v�e���R����ʒu�Q��
        .HWFRHWYT = rs("E021HWFRHWYT")                  ' �i�v�e���R�ۏؕ��@�Q��
        .HWFRHWYS = rs("E021HWFRHWYS")                  ' �i�v�e���R�ۏؕ��@�Q��
        .HWFRMCAL = rs("E021HWFRMCAL")                  ' �i�v�e���R�ʓ��v�Z
        .HWFRAMIN = fncNullCheck(rs("E021HWFRAMIN"))    ' �i�v�e���R���ω���
        .HWFRAMAX = fncNullCheck(rs("E021HWFRAMAX"))    ' �i�v�e���R���Ϗ��
        .HWFRMBNP = fncNullCheck(rs("E021HWFRMBNP"))    ' �i�v�e���R�ʓ����z

        .HWFMKMIN = fncNullCheck(rs("E024HWFMKMIN"))    ' �i�v�e�����בw����
        .HWFMKMAX = fncNullCheck(rs("E024HWFMKMAX"))    ' �i�v�e�����בw���
        .HWFMKSPH = rs("E024HWFMKSPH")                  ' �i�v�e�����בw����ʒu�Q��
        .HWFMKSPT = rs("E024HWFMKSPT")                  ' �i�v�e�����בw����ʒu�Q�_
        .HWFMKSPR = rs("E024HWFMKSPR")                  ' �i�v�e�����בw����ʒu�Q��
        .HWFMKHWT = rs("E024HWFMKHWT")                  ' �i�v�e�����בw�ۏؕ��@�Q��
        .HWFMKHWS = rs("E024HWFMKHWS")                  ' �i�v�e�����בw�ۏؕ��@�Q��

        .HWFONMIN = fncNullCheck(rs("E025HWFONMIN"))    ' �i�v�e�_�f�Z�x����
        .HWFONMAX = fncNullCheck(rs("E025HWFONMAX"))    ' �i�v�e�_�f�Z�x���
        .HWFONSPH = rs("E025HWFONSPH")                  ' �i�v�e�_�f�Z�x����ʒu�Q��
        .HWFONSPT = rs("E025HWFONSPT")                  ' �i�v�e�_�f�Z�x����ʒu�Q�_
        .HWFONSPI = rs("E025HWFONSPI")                  ' �i�v�e�_�f�Z�x����ʒu�Q��
        .HWFONHWT = rs("E025HWFONHWT")                  ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
        .HWFONHWS = rs("E025HWFONHWS")                  ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
        .HWFONMCL = rs("E025HWFONMCL")                  ' �i�v�e�_�f�Z�x�ʓ��v�Z
        .HWFONMBP = fncNullCheck(rs("E025HWFONMBP"))    ' �i�v�e�_�f�Z�x�ʓ����z
        .HWFONAMN = fncNullCheck(rs("E025HWFONAMN"))    ' �i�v�e�_�f�Z�x���ω���
        .HWFONAMX = fncNullCheck(rs("E025HWFONAMX"))    ' �i�v�e�_�f�Z�x���Ϗ��

        .HWFOS1MN = fncNullCheck(rs("E025HWFOS1MN"))    ' �i�v�e�_�f�͏o�P����
        .HWFOS1MX = fncNullCheck(rs("E025HWFOS1MX"))    ' �i�v�e�_�f�͏o�P���
        .HWFOS1SH = rs("E025HWFOS1SH")                  ' �i�v�e�_�f�͏o�P����ʒu�Q��
        .HWFOS1ST = rs("E025HWFOS1ST")                  ' �i�v�e�_�f�͏o�P����ʒu�Q�_
        .HWFOS1SI = rs("E025HWFOS1SI")                  ' �i�v�e�_�f�͏o�P����ʒu�Q��
        .HWFOS1HT = rs("E025HWFOS1HT")                  ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        .HWFOS1HS = rs("E025HWFOS1HS")                  ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        .HWFOS2SH = rs("E025HWFOS2SH")                  ' �i�v�e�_�f�͏o�Q����ʒu�Q��
        .HWFOS2ST = rs("E025HWFOS2ST")                  ' �i�v�e�_�f�͏o�Q����ʒu�Q�_
        .HWFOS2SI = rs("E025HWFOS2SI")                  ' �i�v�e�_�f�͏o�Q����ʒu�Q��
        .HWFOS2MN = fncNullCheck(rs("E025HWFOS2MN"))    ' �i�v�e�_�f�͏o�Q����
        .HWFOS2MX = fncNullCheck(rs("E025HWFOS2MX"))    ' �i�v�e�_�f�͏o�Q���
        .HWFOS2HT = rs("E025HWFOS2HT")                  ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
        .HWFOS2HS = rs("E025HWFOS2HS")                  ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
        .HWFOS3MN = fncNullCheck(rs("E025HWFOS3MN"))    ' �i�v�e�_�f�͏o�R����
        .HWFOS3MX = fncNullCheck(rs("E025HWFOS3MX"))    ' �i�v�e�_�f�͏o�R���
        .HWFOS3SH = rs("E025HWFOS3SH")                  ' �i�v�e�_�f�͏o�R����ʒu�Q��
        .HWFOS3ST = rs("E025HWFOS3ST")                  ' �i�v�e�_�f�͏o�R����ʒu�Q�_
        .HWFOS3SI = rs("E025HWFOS3SI")                  ' �i�v�e�_�f�͏o�R����ʒu�Q��
        .HWFOS3HT = rs("E025HWFOS3HT")                  ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��
        .HWFOS3HS = rs("E025HWFOS3HS")                  ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��

        .HWFDSOMX = fncNullCheck(rs("E026HWFDSOMX"))    ' �i�v�e�c�r�n�c���
        .HWFDSOMN = fncNullCheck(rs("E026HWFDSOMN"))    ' �i�v�e�c�r�n�c����
        .HWFDSOAX = fncNullCheck(rs("E026HWFDSOAX"))    ' �i�v�e�c�r�n�c�̈���
        .HWFDSOAN = fncNullCheck(rs("E026HWFDSOAN"))    ' �i�v�e�c�r�n�c�̈扺��
        .HWFDSOHT = rs("E026HWFDSOHT")                  ' �i�v�e�c�r�n�c�ۏؕ��@�Q��
        .HWFDSOHS = rs("E026HWFDSOHS")                  ' �i�v�e�c�r�n�c�ۏؕ��@�Q��
        
        ' SPV9�_�Ή�  ������2����3���ύX�Ή�
        .HWFSPVMX = fncNullCheck(rs("E028HWFSPVMXN"))   ' �i�v�e�r�o�u�e�d���
        
        .HWFSPVSH = rs("E028HWFSPVSH")                  ' �i�v�e�r�o�u�e�d����ʒu�Q��
        .HWFSPVST = rs("E028HWFSPVST")                  ' �i�v�e�r�o�u�e�d����ʒu�Q�_
        .HWFSPVSI = rs("E028HWFSPVSI")                  ' �i�v�e�r�o�u�e�d����ʒu�Q��
        .HWFSPVHT = rs("E028HWFSPVHT")                  ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
        .HWFSPVHS = rs("E028HWFSPVHS")                  ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
        .HWFDLSPH = rs("E028HWFDLSPH")                  ' �i�v�e�g�U������ʒu�Q��
        .HWFDLSPT = rs("E028HWFDLSPT")                  ' �i�v�e�g�U������ʒu�Q�_
        .HWFDLSPI = rs("E028HWFDLSPI")                  ' �i�v�e�g�U������ʒu�Q��
        .HWFDLHWT = rs("E028HWFDLHWT")                  ' �i�v�e�g�U���ۏؕ��@�Q��
        .HWFDLHWS = rs("E028HWFDLHWS")                  ' �i�v�e�g�U���ۏؕ��@�Q��
        .HWFDLMIN = fncNullCheck(rs("E028HWFDLMIN"))    ' �i�v�e�g�U������
        .HWFDLMAX = fncNullCheck(rs("E028HWFDLMAX"))    ' �i�v�e�g�U�����
                    
        .HWFOF1AX = fncNullCheck(rs("E029HWFOF1AX"))    ' �i�v�e�n�r�e�P���Ϗ��
        .HWFOF1MX = fncNullCheck(rs("E029HWFOF1MX"))    ' �i�v�e�n�r�e�P���
        .HWFOF1SH = rs("E029HWFOF1SH")                  ' �i�v�e�n�r�e�P����ʒu�Q��
        .HWFOF1ST = rs("E029HWFOF1ST")                  ' �i�v�e�n�r�e�P����ʒu�Q�_
        .HWFOF1SR = rs("E029HWFOF1SR")                  ' �i�v�e�n�r�e�P����ʒu�Q��
        .HWFOF1HT = rs("E029HWFOF1HT")                  ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
        .HWFOF1HS = rs("E029HWFOF1HS")                  ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
        .HWFOF2AX = fncNullCheck(rs("E029HWFOF2AX"))    ' �i�v�e�n�r�e�Q���Ϗ��
        .HWFOF2MX = fncNullCheck(rs("E029HWFOF2MX"))    ' �i�v�e�n�r�e�Q���
        .HWFOF2SH = rs("E029HWFOF2SH")                  ' �i�v�e�n�r�e�Q����ʒu�Q��
        .HWFOF2ST = rs("E029HWFOF2ST")                  ' �i�v�e�n�r�e�Q����ʒu�Q�_
        .HWFOF2SR = rs("E029HWFOF2SR")                  ' �i�v�e�n�r�e�Q����ʒu�Q��
        .HWFOF2HT = rs("E029HWFOF2HT")                  ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
        .HWFOF2HS = rs("E029HWFOF2HS")                  ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
        .HWFOF3AX = fncNullCheck(rs("E029HWFOF3AX"))    ' �i�v�e�n�r�e�R���Ϗ��
        .HWFOF3MX = fncNullCheck(rs("E029HWFOF3MX"))    ' �i�v�e�n�r�e�R���
        .HWFOF3SH = rs("E029HWFOF3SH")                  ' �i�v�e�n�r�e�R����ʒu�Q��
        .HWFOF3ST = rs("E029HWFOF3ST")                  ' �i�v�e�n�r�e�R����ʒu�Q�_
        .HWFOF3SR = rs("E029HWFOF3SR")                  ' �i�v�e�n�r�e�R����ʒu�Q��
        .HWFOF3HT = rs("E029HWFOF3HT")                  ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
        .HWFOF3HS = rs("E029HWFOF3HS")                  ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
        .HWFOF4AX = fncNullCheck(rs("E029HWFOF4AX"))    ' �i�v�e�n�r�e�S���Ϗ��
        .HWFOF4MX = fncNullCheck(rs("E029HWFOF4MX"))    ' �i�v�e�n�r�e�S���
        .HWFOF4SH = rs("E029HWFOF4SH")                  ' �i�v�e�n�r�e�S����ʒu�Q��
        .HWFOF4ST = rs("E029HWFOF4ST")                  ' �i�v�e�n�r�e�S����ʒu�Q�_
        .HWFOF4SR = rs("E029HWFOF4SR")                  ' �i�v�e�n�r�e�S����ʒu�Q��
        .HWFOF4HT = rs("E029HWFOF4HT")                  ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
        .HWFOF4HS = rs("E029HWFOF4HS")                  ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
        If IsNull(rs("E029HWFOSF1PTK")) = False Then .HWFOSF1PTK = rs("E029HWFOSF1PTK")       ' �i�v�e�n�r�e�P�p�^���敪�@��2003/05/14 ooba
        If IsNull(rs("E029HWFOSF2PTK")) = False Then .HWFOSF2PTK = rs("E029HWFOSF2PTK")       ' �i�v�e�n�r�e�Q�p�^���敪
        If IsNull(rs("E029HWFOSF3PTK")) = False Then .HWFOSF3PTK = rs("E029HWFOSF3PTK")       ' �i�v�e�n�r�e�R�p�^���敪
        If IsNull(rs("E029HWFOSF4PTK")) = False Then .HWFOSF4PTK = rs("E029HWFOSF4PTK")       ' �i�v�e�n�r�e�S�p�^���敪�@��2003/05/14 ooba

        ' BMD�ׂ��搔�ύX�Ή�
        .HWFBM1AN = fncNullCheck(rs("E029HWFBM1AN"))    ' �i�v�e�a�l�c�P���ω���
        .HWFBM1AX = fncNullCheck(rs("E029HWFBM1AX"))    ' �i�v�e�a�l�c�P���Ϗ��
        .HWFBM1SH = rs("E029HWFBM1SH")                  ' �i�v�e�a�l�c�P����ʒu�Q��
        .HWFBM1ST = rs("E029HWFBM1ST")                  ' �i�v�e�a�l�c�P����ʒu�Q�_
        .HWFBM1SR = rs("E029HWFBM1SR")                  ' �i�v�e�a�l�c�P����ʒu�Q��
        .HWFBM1HT = rs("E029HWFBM1HT")                  ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
        .HWFBM1HS = rs("E029HWFBM1HS")                  ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
        
        'BMD�ׂ��搔�ύX�Ή�
        .HWFBM2AN = fncNullCheck(rs("E029HWFBM2AN"))    ' �i�v�e�a�l�c�Q���ω���
        .HWFBM2AX = fncNullCheck(rs("E029HWFBM2AX"))    ' �i�v�e�a�l�c�Q���Ϗ��
        .HWFBM2SH = rs("E029HWFBM2SH")                  ' �i�v�e�a�l�c�Q����ʒu�Q��
        .HWFBM2ST = rs("E029HWFBM2ST")                  ' �i�v�e�a�l�c�Q����ʒu�Q�_
        .HWFBM2SR = rs("E029HWFBM2SR")                  ' �i�v�e�a�l�c�Q����ʒu�Q��
        .HWFBM2HT = rs("E029HWFBM2HT")                  ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
        .HWFBM2HS = rs("E029HWFBM2HS")                  ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��

        ' BMD�ׂ��搔�ύX�Ή�
        .HWFBM3AN = fncNullCheck(rs("E029HWFBM3AN"))    ' �i�v�e�a�l�c�R���ω���
        .HWFBM3AX = fncNullCheck(rs("E029HWFBM3AX"))    ' �i�v�e�a�l�c�R���Ϗ��
        .HWFBM3SH = rs("E029HWFBM3SH")                  ' �i�v�e�a�l�c�R����ʒu�Q��
        .HWFBM3ST = rs("E029HWFBM3ST")                  ' �i�v�e�a�l�c�R����ʒu�Q�_
        .HWFBM3SR = rs("E029HWFBM3SR")                  ' �i�v�e�a�l�c�R����ʒu�Q��
        .HWFBM3HT = rs("E029HWFBM3HT")                  ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
        .HWFBM3HS = rs("E029HWFBM3HS")                  ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
        
        .HWFBM1MBP = fncNullCheck(rs("E029HWFBM1MBP"))  ' �i�v�e�a�l�c�P�ʓ����z
        .HWFBM2MBP = fncNullCheck(rs("E029HWFBM2MBP"))  ' �i�v�e�a�l�c�Q�ʓ����z
        .HWFBM3MBP = fncNullCheck(rs("E029HWFBM3MBP"))  ' �i�v�e�a�l�c�R�ʓ����z
        If IsNull(rs("E029HWFBM1MCL")) = False Then .HWFBM1MCL = rs("E029HWFBM1MCL")         ' �i�v�e�a�l�c�P�ʓ��v�Z
        If IsNull(rs("E029HWFBM2MCL")) = False Then .HWFBM2MCL = rs("E029HWFBM2MCL")         ' �i�v�e�a�l�c�Q�ʓ��v�Z
        If IsNull(rs("E029HWFBM3MCL")) = False Then .HWFBM3MCL = rs("E029HWFBM3MCL")         ' �i�v�e�a�l�c�R�ʓ��v�Z�@��2003/05/14 ooba

        .HWFOS1NS = rs("E025HWFOS1NS")                  ' �i�v�e�_�f�͏o�P�M�����@
        .HWFOS2NS = rs("E025HWFOS2NS")                  ' �i�v�e�_�f�͏o�Q�M�����@
        .HWFOS3NS = rs("E025HWFOS3NS")                  ' �i�v�e�_�f�͏o�R�M�����@
        .HWFOF1NS = rs("E029HWFOF1NS")                  ' �i�v�e�n�r�e�P�M�����@
        .HWFOF2NS = rs("E029HWFOF2NS")                  ' �i�v�e�n�r�e�Q�M�����@
        .HWFOF3NS = rs("E029HWFOF3NS")                  ' �i�v�e�n�r�e�R�M�����@
        .HWFOF4NS = rs("E029HWFOF4NS")                  ' �i�v�e�n�r�e�S�M�����@
        .HWFBM1NS = rs("E029HWFBM1NS")                  ' �i�v�e�a�l�c�P�M�����@
        .HWFBM2NS = rs("E029HWFBM2NS")                  ' �i�v�e�a�l�c�Q�M�����@
        .HWFBM3NS = rs("E029HWFBM3NS")                  ' �i�v�e�a�l�c�R�M�����@

        .HWFANTIM = fncNullCheck(rs("E025HWFANTIM"))    ' �i�v�e�`�m����
        .HWFANTNP = fncNullCheck(rs("E025HWFANTNP"))    ' �i�v�e�`�m���x

        .HWFOF1ET = fncNullCheck(rs("E029HWFOF1ET"))    ' �i�v�e�n�r�e�P�I���d�s��
        .HWFOF2ET = fncNullCheck(rs("E029HWFOF2ET"))    ' �i�v�e�n�r�e�Q�I���d�s��
        .HWFOF3ET = fncNullCheck(rs("E029HWFOF3ET"))    ' �i�v�e�n�r�e�R�I���d�s��
        .HWFOF4ET = fncNullCheck(rs("E029HWFOF4ET"))    ' �i�v�e�n�r�e�S�I���d�s��
        .HWFBM1ET = fncNullCheck(rs("E029HWFBM1ET"))    ' �i�v�e�a�l�c�P�I���d�s��
        .HWFBM2ET = fncNullCheck(rs("E029HWFBM2ET"))    ' �i�v�e�a�l�c�Q�I���d�s��
        .HWFBM3ET = fncNullCheck(rs("E029HWFBM3ET"))    ' �i�v�e�a�l�c�R�I���d�s��

        .HWFOF1SZ = rs("E029HWFOF1SZ")                  ' �i�v�e�n�r�e�P�������
        .HWFOF2SZ = rs("E029HWFOF2SZ")                  ' �i�v�e�n�r�e�Q�������
        .HWFOF3SZ = rs("E029HWFOF3SZ")                  ' �i�v�e�n�r�e�R�������
        .HWFOF4SZ = rs("E029HWFOF4SZ")                  ' �i�v�e�n�r�e�S�������
        .HWFBM1SZ = rs("E029HWFBM1SZ")                  ' �i�v�e�a�l�c�P�������
        .HWFBM2SZ = rs("E029HWFBM2SZ")                  ' �i�v�e�a�l�c�Q�������
        .HWFBM3SZ = rs("E029HWFBM3SZ")                  ' �i�v�e�a�l�c�R�������
    
        ' SPV9�_�Ή�  ������2����3���ύX�Ή�
        .HWFSPVAM = fncNullCheck(rs("E028HWFSPVAMN"))   ' �i�v�e�r�o�u�e�d���Ϗ��

'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")                      ' DK���x�i�d�l�j
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        .HSXLDLRMN = fncNullCheck(rs("HSXLDLRMN"))      ' �iSXL/DL�A��0����
        .HSXLDLRMX = fncNullCheck(rs("HSXLDLRMX"))      ' �iSXL/DL�A��0���
        .HWFLDLRMN = fncNullCheck(rs("HWFLDLRMN"))      ' �iWFL/DL�A��0����
        .HWFLDLRMX = fncNullCheck(rs("HWFLDLRMX"))      ' �iWFL/DL�A��0���
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
        
        End With
    Set rs = Nothing
    
    ' �����p�x_���ް��擾
    sSQL = "select "
    sSQL = sSQL & "TBCME026.HWFDSOPTK, "                ' �iWFDSOD�p�^���敪
    sSQL = sSQL & "TBCME021.HWFRKHNN, "                 ' �i�v�e���R�����p�x�Q��
    sSQL = sSQL & "TBCME025.HWFONKHN, "                 ' �i�v�e�_�f�Z�x�����p�x�Q��
    sSQL = sSQL & "TBCME029.HWFOF1KN, "                 ' �i�v�e�n�r�e�P�����p�x�Q��
    sSQL = sSQL & "TBCME029.HWFOF2KN, "                 ' �i�v�e�n�r�e�Q�����p�x�Q��
    sSQL = sSQL & "TBCME029.HWFOF3KN, "                 ' �i�v�e�n�r�e�R�����p�x�Q��
    sSQL = sSQL & "TBCME029.HWFOF4KN, "                 ' �i�v�e�n�r�e�S�����p�x�Q��
    sSQL = sSQL & "TBCME029.HWFBM1KN, "                 ' �i�v�e�a�l�c�P�����p�x�Q��
    sSQL = sSQL & "TBCME029.HWFBM2KN, "                 ' �i�v�e�a�l�c�Q�����p�x�Q��
    sSQL = sSQL & "TBCME029.HWFBM3KN, "                 ' �i�v�e�a�l�c�R�����p�x�Q��
    sSQL = sSQL & "TBCME025.HWFOS1KN, "                 ' �i�v�e�_�f�͏o�P�����p�x�Q��
    sSQL = sSQL & "TBCME025.HWFOS2KN, "                 ' �i�v�e�_�f�͏o�Q�����p�x�Q��
    sSQL = sSQL & "TBCME025.HWFOS3KN, "                 ' �i�v�e�_�f�͏o�R�����p�x�Q��
    sSQL = sSQL & "TBCME026.HWFDSOKN, "                 ' �i�v�e�c�r�n�c�����p�x�Q��
    sSQL = sSQL & "TBCME024.HWFMKKHN, "                 ' �i�v�e�����בw�����p�x�Q��
    sSQL = sSQL & "TBCME028.HWFSPVKN, "                 ' �i�v�e�r�o�u�e�d�����p�x�Q��
    sSQL = sSQL & "TBCME028.HWFDLKHN, "                 ' �i�v�e�g�U�������p�x�Q��
    sSQL = sSQL & "TBCME025.HWFZOKHN, "                 ' �i�v�e�c���_�f�����p�x�Q��
    sSQL = sSQL & "TBCME026.HWFGDKHN "                  ' �i�v�e�f�c�����p�x�Q��
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
    
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & "and TBCME021.HINBAN = '" & udtTypIn.HIN.hinban & "' "
        sSQL = sSQL & "and TBCME021.MNOREVNO = " & udtTypIn.HIN.mnorevno & " "
        sSQL = sSQL & "and TBCME021.FACTORY = '" & udtTypIn.HIN.factory & "' "
        sSQL = sSQL & "and TBCME021.OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & "and TBCME021.HINBAN = '" & udtNew_Hinban.hinban & "' "
        sSQL = sSQL & "and TBCME021.MNOREVNO = " & udtNew_Hinban.mnorevno & " "
        sSQL = sSQL & "and TBCME021.FACTORY = '" & udtNew_Hinban.factory & "' "
        sSQL = sSQL & "and TBCME021.OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If

    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' ���R�[�h0���̓G���[�I��
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    With udtSiyou
        If IsNull(rs("HWFDSOPTK")) = False Then .HWFDSOPTK = rs("HWFDSOPTK") Else .HWFDSOPTK = " "  ' �iWFDSOD�p�^���敪
        If IsNull(rs("HWFRKHNN")) = False Then .HWFRKHNN = rs("HWFRKHNN") Else .HWFRKHNN = " "      ' �i�v�e���R�����p�x�Q��
        If IsNull(rs("HWFONKHN")) = False Then .HWFONKHN = rs("HWFONKHN") Else .HWFONKHN = " "      ' �i�v�e�_�f�Z�x�����p�x�Q��
        If IsNull(rs("HWFOF1KN")) = False Then .HWFOF1KN = rs("HWFOF1KN") Else .HWFOF1KN = " "      ' �i�v�e�n�r�e�P�����p�x�Q��
        If IsNull(rs("HWFOF2KN")) = False Then .HWFOF2KN = rs("HWFOF2KN") Else .HWFOF2KN = " "      ' �i�v�e�n�r�e�Q�����p�x�Q��
        If IsNull(rs("HWFOF3KN")) = False Then .HWFOF3KN = rs("HWFOF3KN") Else .HWFOF3KN = " "      ' �i�v�e�n�r�e�R�����p�x�Q��
        If IsNull(rs("HWFOF4KN")) = False Then .HWFOF4KN = rs("HWFOF4KN") Else .HWFOF4KN = " "      ' �i�v�e�n�r�e�S�����p�x�Q��
        If IsNull(rs("HWFBM1KN")) = False Then .HWFBM1KN = rs("HWFBM1KN") Else .HWFBM1KN = " "      ' �i�v�e�a�l�c�P�����p�x�Q��
        If IsNull(rs("HWFBM2KN")) = False Then .HWFBM2KN = rs("HWFBM2KN") Else .HWFBM2KN = " "      ' �i�v�e�a�l�c�Q�����p�x�Q��
        If IsNull(rs("HWFBM3KN")) = False Then .HWFBM3KN = rs("HWFBM3KN") Else .HWFBM3KN = " "      ' �i�v�e�a�l�c�R�����p�x�Q��
        If IsNull(rs("HWFOS1KN")) = False Then .HWFOS1KN = rs("HWFOS1KN") Else .HWFOS1KN = " "      ' �i�v�e�_�f�͏o�P�����p�x�Q��
        If IsNull(rs("HWFOS2KN")) = False Then .HWFOS2KN = rs("HWFOS2KN") Else .HWFOS2KN = " "      ' �i�v�e�_�f�͏o�Q�����p�x�Q��
        If IsNull(rs("HWFOS3KN")) = False Then .HWFOS3KN = rs("HWFOS3KN") Else .HWFOS3KN = " "      ' �i�v�e�_�f�͏o�R�����p�x�Q��
        If IsNull(rs("HWFDSOKN")) = False Then .HWFDSOKN = rs("HWFDSOKN") Else .HWFDSOKN = " "      ' �i�v�e�c�r�n�c�����p�x�Q��
        If IsNull(rs("HWFMKKHN")) = False Then .HWFMKKHN = rs("HWFMKKHN") Else .HWFMKKHN = " "      ' �i�v�e�����בw�����p�x�Q��
        If IsNull(rs("HWFSPVKN")) = False Then .HWFSPVKN = rs("HWFSPVKN") Else .HWFSPVKN = " "      ' �i�v�e�r�o�u�e�d�����p�x�Q��
        If IsNull(rs("HWFDLKHN")) = False Then .HWFDLKHN = rs("HWFDLKHN") Else .HWFDLKHN = " "      ' �i�v�e�g�U�������p�x�Q��
        If IsNull(rs("HWFZOKHN")) = False Then .HWFZOKHN = rs("HWFZOKHN") Else .HWFZOKHN = " "      ' �i�v�e�c���_�f�����p�x�Q��
        If IsNull(rs("HWFGDKHN")) = False Then .HWFGDKHN = rs("HWFGDKHN") Else .HWFGDKHN = " "      ' �i�v�e�f�c�����p�x�Q��
    End With
    Set rs = Nothing
    
    ''�c���_�f�d�l�擾
    sDBName = "E025"
    sSQL = "select HWFZOMIN, HWFZOMAX, HWFZOSPH, HWFZOSPT, HWFZOSPI, HWFZOHWT, "
    sSQL = sSQL & "HWFZOHWS, HWFZONSW from TBCME025 "
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If
    
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' ���R�[�h0���̓G���[�I��
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    udtSiyou.HWFZOMIN = fncNullCheck(rs("HWFZOMIN"))                                                        ' �i�v�e�c���_�f����
    udtSiyou.HWFZOMAX = fncNullCheck(rs("HWFZOMAX"))                                                        ' �i�v�e�c���_�f���
    If IsNull(rs("HWFZOSPH")) = False Then udtSiyou.HWFZOSPH = rs("HWFZOSPH") Else udtSiyou.HWFZOSPH = " "  ' �i�v�e�c���_�f����ʒu�Q��
    If IsNull(rs("HWFZOSPT")) = False Then udtSiyou.HWFZOSPT = rs("HWFZOSPT") Else udtSiyou.HWFZOSPT = " "  ' �i�v�e�c���_�f����ʒu�Q�_
    If IsNull(rs("HWFZOSPI")) = False Then udtSiyou.HWFZOSPI = rs("HWFZOSPI") Else udtSiyou.HWFZOSPI = " "  ' �i�v�e�c���_�f����ʒu�Q��
    If IsNull(rs("HWFZOHWT")) = False Then udtSiyou.HWFZOHWT = rs("HWFZOHWT") Else udtSiyou.HWFZOHWT = " "  ' �i�v�e�c���_�f�ۏؕ��@�Q��
    If IsNull(rs("HWFZOHWS")) = False Then udtSiyou.HWFZOHWS = rs("HWFZOHWS") Else udtSiyou.HWFZOHWS = " "  ' �i�v�e�c���_�f�ۏؕ��@�Q��
    If IsNull(rs("HWFZONSW")) = False Then udtSiyou.HWFZONSW = rs("HWFZONSW") Else udtSiyou.HWFZONSW = " "  ' �i�v�e�c���_�f�M�����@
        
    Set rs = Nothing
    
    ' GD�d�l�擾
    sDBName = "E026"
    sSQL = "select "
    sSQL = sSQL & "HWFDENKU, "        ' �i�v�e�c���������L��
    sSQL = sSQL & "HWFDENMX, "        ' �i�v�e�c�������
    sSQL = sSQL & "HWFDENMN, "        ' �i�v�e�c��������
    sSQL = sSQL & "HWFDENHT, "        ' �i�v�e�c�����ۏؕ��@�Q��
    sSQL = sSQL & "HWFDENHS, "        ' �i�v�e�c�����ۏؕ��@�Q��
    sSQL = sSQL & "HWFDVDKU, "        ' �i�v�e�c�u�c�Q�����L��
    sSQL = sSQL & "HWFDVDMXN, "       ' �i�v�e�c�u�c�Q���
    sSQL = sSQL & "HWFDVDMNN, "       ' �i�v�e�c�u�c�Q����
    sSQL = sSQL & "HWFDVDHT, "        ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��
    sSQL = sSQL & "HWFDVDHS, "        ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��
    sSQL = sSQL & "HWFLDLKU, "        ' �i�v�e�k�^�c�k�����L��
    sSQL = sSQL & "HWFLDLMX, "        ' �i�v�e�k�^�c�k���
    sSQL = sSQL & "HWFLDLMN, "        ' �i�v�e�k�^�c�k����
    sSQL = sSQL & "HWFLDLHT, "        ' �i�v�e�k�^�c�k�ۏؕ��@�Q��
    sSQL = sSQL & "HWFLDLHS, "        ' �i�v�e�k�^�c�k�ۏؕ��@�Q��
    sSQL = sSQL & "HWFGDSPH, "        ' �i�v�e�f�c����ʒu�Q��
    sSQL = sSQL & "HWFGDSPT, "        ' �i�v�e�f�c����ʒu�Q�_
    sSQL = sSQL & "HWFGDSPR "         ' �i�v�e�f�c����ʒu�Q��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    sSQL = sSQL & ",HWFGDPTK "        ' �i�v�e�f�c�p�^���敪
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    sSQL = sSQL & "from TBCME026 "
    
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' ���R�[�h0���̓G���[�I��
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
        
    With udtSiyou
        .HWFDENKU = rs("HWFDENKU")                      ' �i�v�e�c���������L��
        .HWFDENMX = fncNullCheck(rs("HWFDENMX"))        ' �i�v�e�c�������
        .HWFDENMN = fncNullCheck(rs("HWFDENMN"))        ' �i�v�e�c��������
        .HWFDENHT = rs("HWFDENHT")                      ' �i�v�e�c�����ۏؕ��@�Q��
        .HWFDENHS = rs("HWFDENHS")                      ' �i�v�e�c�����ۏؕ��@�Q��
        .HWFDVDKU = rs("HWFDVDKU")                      ' �i�v�e�c�u�c�Q�����L��
        .HWFDVDMXN = fncNullCheck(rs("HWFDVDMXN"))      ' �i�v�e�c�u�c�Q���
        .HWFDVDMNN = fncNullCheck(rs("HWFDVDMNN"))      ' �i�v�e�c�u�c�Q����
        .HWFDVDHT = rs("HWFDVDHT")                      ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��
        .HWFDVDHS = rs("HWFDVDHS")                      ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��
        .HWFLDLKU = rs("HWFLDLKU")                      ' �i�v�e�k�^�c�k�����L��
        .HWFLDLMX = fncNullCheck(rs("HWFLDLMX"))        ' �i�v�e�k�^�c�k���
        .HWFLDLMN = fncNullCheck(rs("HWFLDLMN"))        ' �i�v�e�k�^�c�k����
        .HWFLDLHT = rs("HWFLDLHT")                      ' �i�v�e�k�^�c�k�ۏؕ��@�Q��
        .HWFLDLHS = rs("HWFLDLHS")                      ' �i�v�e�k�^�c�k�ۏؕ��@�Q��
        .HWFGDSPH = rs("HWFGDSPH")                      ' �i�v�e�f�c����ʒu�Q��
        .HWFGDSPT = rs("HWFGDSPT")                      ' �i�v�e�f�c����ʒu�Q�_
        .HWFGDSPR = rs("HWFGDSPR")                      ' �i�v�e�f�c����ʒu�Q��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        If IsNull(rs("HWFGDPTK")) = False Then .HWFGDPTK = rs("HWFGDPTK") Else .HWFGDPTK = " "      ' �i�v�e�f�c�p�^���敪
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    End With
    Set rs = Nothing

    ' �i�v�e�f�c���C�����̎擾
    sDBName = "E036"
    
    sSQL = "select "
    sSQL = sSQL & "HWFGDLINE "        ' �i�v�e�f�c���C�����̎擾
    sSQL = sSQL & "from TBCME036 "
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' ���R�[�h0���̓G���[�I��
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
        
    With udtSiyou
        .HWFGDLINE = fncNullCheck(rs("HWFGDLINE"))  ' �i�v�e�f�c���C�����̎擾
    End With
    
    Set rs = Nothing
    
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    ' �i�r�w�f�c�p�^���敪�̎擾
    sDBName = "E020"
    
    sSQL = "select "
    sSQL = sSQL & "HSXGDPTK "        ' �i�r�w�f�c�p�^���敪
    sSQL = sSQL & "from TBCME020 "
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' ���R�[�h0���̓G���[�I��
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
        
    With udtSiyou
        If IsNull(rs("HSXGDPTK")) = False Then .HSXGDPTK = rs("HSXGDPTK") Else .HSXGDPTK = " "
    End With
    
    Set rs = Nothing
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    
    ' SPVNr�Z�x�d�l�擾
    ' Add 2010/01/06 SIRD�Ή� Y.Hitomi
    sDBName = "E048"
    sSQL = "select "
    sSQL = sSQL & "HWFNRHS, "                       ' �iWFSPVNR�ۏؕ��@_��
    sSQL = sSQL & "HWFNRKN, "                       ' �iWFSPVNR�ۏؕ��@_��
    ' ��Add 2010/01/06 SIRD�Ή� Y.Hitomi
    sSQL = sSQL & "HWFSIRDMX, "                     ' ����]�ʏ��
    sSQL = sSQL & "HWFSIRDSZ, "                     ' ����]�ʑ������
    sSQL = sSQL & "HWFSIRDHT, "                     ' ����]�ʕۏؕ��@�Q��
    sSQL = sSQL & "HWFSIRDHS, "                     ' ����]�ʕۏؕ��@�Q��
    sSQL = sSQL & "HWFSIRDKM, "                     ' ����]�ʌ����p�x�Q��
    sSQL = sSQL & "HWFSIRDKN, "                     ' ����]�ʌ����p�x�Q��
    sSQL = sSQL & "HWFSIRDKH, "                     ' ����]�ʌ����p�x�Q��
    sSQL = sSQL & "HWFSIRDKU  "                     ' ����]�ʌ����p�x�Q�E
    ' ��Add 2010/01/06 SIRD�Ή� Y.Hitomi
    sSQL = sSQL & "from TBCME048 "
    sSQL = sSQL & "where HINBAN = '" & udtNew_Hinban.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & udtNew_Hinban.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & udtNew_Hinban.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & udtNew_Hinban.opecond & "' "
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    If IsNull(rs("HWFNRHS")) = False Then udtSiyou.HWFNRHS = rs("HWFNRHS") Else udtSiyou.HWFNRHS = " "
    If IsNull(rs("HWFNRKN")) = False Then udtSiyou.HWFNRKN = rs("HWFNRKN") Else udtSiyou.HWFNRKN = " "
    ' ��Add 2010/01/06 SIRD�Ή� Y.Hitomi
    If IsNull(rs("HWFSIRDMX")) = False Then udtSiyou.HWFSIRDMX = rs("HWFSIRDMX") Else udtSiyou.HWFSIRDMX = fncNullCheck(rs("HWFSIRDMX"))
    If IsNull(rs("HWFSIRDSZ")) = False Then udtSiyou.HWFSIRDSZ = rs("HWFSIRDSZ") Else udtSiyou.HWFSIRDSZ = " "
    If IsNull(rs("HWFSIRDHT")) = False Then udtSiyou.HWFSIRDHT = rs("HWFSIRDHT") Else udtSiyou.HWFSIRDHT = " "
    If IsNull(rs("HWFSIRDHS")) = False Then udtSiyou.HWFSIRDHS = rs("HWFSIRDHS") Else udtSiyou.HWFSIRDHS = " "
    If IsNull(rs("HWFSIRDKM")) = False Then udtSiyou.HWFSIRDKM = rs("HWFSIRDKM") Else udtSiyou.HWFSIRDKM = " "
    If IsNull(rs("HWFSIRDKN")) = False Then udtSiyou.HWFSIRDKN = rs("HWFSIRDKN") Else udtSiyou.HWFSIRDKN = " "
    If IsNull(rs("HWFSIRDKH")) = False Then udtSiyou.HWFSIRDKH = rs("HWFSIRDKH") Else udtSiyou.HWFSIRDKH = " "
    If IsNull(rs("HWFSIRDKU")) = False Then udtSiyou.HWFSIRDKU = rs("HWFSIRDKU") Else udtSiyou.HWFSIRDKU = " "
    ' ��Add 2010/01/06 SIRD�Ή� Y.Hitomi
    
    rs.Close
    
    ' ����]�����ʎ擾
    sDBName = "Y013"
    If funGetTBCMY013(udtTypIn, udtSokutei()) = FUNCTION_RETURN_FAILURE Then
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
        
    ' GD���ю擾
    If udtTypIn.WFSMP.WFINDGDCW <> "0" Then
        ' ����وʒu���
        If udtTypIn.WFSMP.TBKBNCW = "T" Then intPos = SxlTop Else intPos = SxlTail
        
        ' ����GD���ю擾
        If udtTypIn.WFSMP.WFHSGDCW = "1" Then
            sDBName = "J006"
            If funGetGDJisseki_J006(udtTypIn.WFSMP.XTALCW, udtTypIn.WFSMP.WFSMPLIDGDCW, _
                                        typ_J015_WFGDJudg(intPos)) = FUNCTION_RETURN_FAILURE Then
                funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        ' WF_GD���ю擾
        Else
            sDBName = "J015"
            If funGetGDJisseki_J015(udtTypIn.WFSMP.XTALCW, udtTypIn.WFSMP.WFSMPLIDGDCW, _
                                        udtTypIn.WFSMP.WFHSGDCW, typ_J015_WFGDJudg(intPos)) _
                                                                = FUNCTION_RETURN_FAILURE Then
                funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End If
    End If

    '��SIRD���ю擾 Add 2010/01/07 SIRD�Ή� Y.Hitomi
        ' ����وʒu���
    If udtTypIn.WFSMP.TBKBNCW = "T" Then
        intPos = SxlTop
    Else
        intPos = SxlTail
    End If

    ' WF_SIRD���ю擾
    If udtTypIn.WFSMP.WFINDL4CW <> "0" Then
        sDBName = "J022"
        If funGetSDJisseki_J022(udtTypIn.WFSMP.XTALCW, udtTypIn.WFSMP.WFSMPLIDL4CW, _
                                   typ_J022_WFSDJudg(intPos)) = FUNCTION_RETURN_FAILURE Then
            funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    '��SIRD���ю擾 Add 2010/01/07 SIRD�Ή� Y.Hitomi

    ' SPV9�_�Ή�
    ' ����وʒu���
    If udtTypIn.WFSMP.TBKBNCW = "T" Then intPos = SxlTop Else intPos = SxlTail
    
    If udtTypIn.WFSMP.WFINDSPCW <> "0" Then
        sDBName = "J016"
        If funGetSPVJisseki_J016(udtTypIn.WFSMP.XTALCW, udtTypIn.WFSMP.WFSMPLIDSPCW, _
                                    typ_J016_WFSPVJudg(intPos), udtSiyou) = FUNCTION_RETURN_FAILURE Then
            funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

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
    
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where K01.HINBAN='" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " K01.MNOREVNO=" & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " K01.FACTORY='" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " K01.OPECOND='" & udtTypIn.HIN.opecond & "' and "
        sSQL = sSQL & " K12.HINBAN='" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " K12.MNOREVNO=" & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " K12.FACTORY='" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " K12.OPECOND='" & udtTypIn.HIN.opecond & "'"
    Else
        sSQL = sSQL & " where K01.HINBAN='" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " K01.MNOREVNO=" & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " K01.FACTORY='" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " K01.OPECOND='" & udtNew_Hinban.opecond & "' and "
        sSQL = sSQL & " K12.HINBAN='" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " K12.MNOREVNO=" & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " K12.FACTORY='" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " K12.OPECOND='" & udtNew_Hinban.opecond & "'"
    End If
    
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    ' ���R�[�h0���̓G���[�I��
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With udtSiyou
        .HWFRSPOT = rs("HSXRSPOT")      ' �i�r�w���R����ʒu�Q�_
        .HWFRSPOI = rs("HSXRSPOI")      ' �i�r�w���R����ʒu�Q��
        .HWFONSPT = rs("HSXONSPT")      ' �i�r�w�_�f�Z�x����ʒu�Q�_
        .HWFONSPI = rs("HSXONSPI")      ' �i�r�w�_�f�Z�x����ʒu�Q��
    End With
    
    Set rs = Nothing
    
    '' �G�s�d�l�擾(BMD,OSF)
    sSQL = "select "
    sSQL = sSQL & "HEPANTNP, "          ' �iEPAN���x
    sSQL = sSQL & "HEPBM1HS, "          ' �iEPBMD1�ۏؕ��@�Q��
    sSQL = sSQL & "HEPBM1AN, "          ' �iEPBMD1���ω���
    sSQL = sSQL & "HEPBM1AX, "          ' �iEPBMD1���Ϗ��
    sSQL = sSQL & "HEPBM2HS, "          ' �iEPBMD1�ۏؕ��@�Q��
    sSQL = sSQL & "HEPBM2AN, "          ' �iEPBMD2���ω���
    sSQL = sSQL & "HEPBM2AX, "          ' �iEPBMD2���Ϗ��
    sSQL = sSQL & "HEPBM3HS, "          ' �iEPBMD1�ۏؕ��@�Q��
    sSQL = sSQL & "HEPBM3AN, "          ' �iEPBMD3���ω���
    sSQL = sSQL & "HEPBM3AX, "          ' �iEPBMD3���Ϗ��
    sSQL = sSQL & "HEPBM3GSAN, "        ' �iEPBMD3���ω���(�O��)�@09/05/07 ooba
    sSQL = sSQL & "HEPBM3GSAX, "        ' �iEPBMD3���Ϗ��(�O��)�@09/05/07 ooba
    sSQL = sSQL & "HEPOF1HS, "          ' �iEPOSF1�ۏؕ��@�Q��
    sSQL = sSQL & "HEPOF1AX, "          ' �iEPOSF1���Ϗ��
    sSQL = sSQL & "HEPOF1MX, "          ' �iEPOSF1���
    sSQL = sSQL & "HEPOF2HS, "          ' �iEPOSF2�ۏؕ��@�Q��
    sSQL = sSQL & "HEPOF2AX, "          ' �iEPOSF2���Ϗ��
    sSQL = sSQL & "HEPOF2MX, "          ' �iEPOSF2���
    sSQL = sSQL & "HEPOF3HS, "          ' �iEPOSF3�ۏؕ��@�Q��
    sSQL = sSQL & "HEPOF3AX, "          ' �iEPOSF3���Ϗ��
    sSQL = sSQL & "HEPOF3MX  "          ' �iEPOSF3���
    sSQL = sSQL & "from TBCME050 "      ' ���i�d�l�G�s�f�[�^�P
    
    If intSmpGetFlg = 0 Then
        sSQL = sSQL & " where HINBAN = '" & udtTypIn.HIN.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtTypIn.HIN.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtTypIn.HIN.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtTypIn.HIN.opecond & "' "
    Else
        sSQL = sSQL & " where HINBAN = '" & udtNew_Hinban.hinban & "' and "
        sSQL = sSQL & " MNOREVNO = " & udtNew_Hinban.mnorevno & " and "
        sSQL = sSQL & " FACTORY = '" & udtNew_Hinban.factory & "' and "
        sSQL = sSQL & " OPECOND = '" & udtNew_Hinban.opecond & "' "
    End If
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' �f�[�^�����̏ꍇ�͏I��
    If rs.RecordCount = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    ' �iEPBMD1�ۏؕ��@�Q����"H","S"�̂��̂���ł�����΁A�G�s�d�l����
    If rs("HEPBM1HS") = "H" Or rs("HEPBM2HS") = "H" Or rs("HEPBM3HS") = "H" Or _
       rs("HEPOF1HS") = "H" Or rs("HEPOF2HS") = "H" Or rs("HEPOF3HS") = "H" Or _
       rs("HEPBM1HS") = "S" Or rs("HEPBM2HS") = "S" Or rs("HEPBM3HS") = "S" Or _
       rs("HEPOF1HS") = "S" Or rs("HEPOF2HS") = "S" Or rs("HEPOF3HS") = "S" Then
        udtSiyou.HEPHS = True
    Else
        udtSiyou.HEPHS = False
    End If
    
    If udtSiyou.HEPHS = True Then
        With udtSiyou
            .HEPBM1AN = fncNullCheck(rs("HEPBM1AN"))   ' �iEPBMD1���ω���
            .HEPBM1AX = fncNullCheck(rs("HEPBM1AX"))   ' �iEPBMD1���Ϗ��
            .HEPBM2AN = fncNullCheck(rs("HEPBM2AN"))   ' �iEPBMD2���ω���
            .HEPBM2AX = fncNullCheck(rs("HEPBM2AX"))   ' �iEPBMD2���Ϗ��
            .HEPBM3AN = fncNullCheck(rs("HEPBM3AN"))   ' �iEPBMD3���ω���
            .HEPBM3AX = fncNullCheck(rs("HEPBM3AX"))   ' �iEPBMD3���Ϗ��
            .HEPBM3GSAN = fncNullCheck(rs("HEPBM3GSAN"))    ' �iEPBMD3���ω���(�O��)�@09/05/07 ooba
            .HEPBM3GSAX = fncNullCheck(rs("HEPBM3GSAX"))    ' �iEPBMD3���Ϗ��(�O��)�@09/05/07 ooba
            .HEPOF1AX = fncNullCheck(rs("HEPOF1AX"))   ' �iEPOSF1���ω���
            .HEPOF1MX = fncNullCheck(rs("HEPOF1MX"))   ' �iEPOSF1���
            .HEPOF2AX = fncNullCheck(rs("HEPOF2AX"))   ' �iEPOSF2���ω���
            .HEPOF2MX = fncNullCheck(rs("HEPOF2MX"))   ' �iEPOSF2���
            .HEPOF3AX = fncNullCheck(rs("HEPOF3AX"))   ' �iEPOSF3���ω���
            .HEPOF3MX = fncNullCheck(rs("HEPOF3MX"))   ' �iEPOSF3���
            .HEPANTNP = fncNullCheck(rs("HEPANTNP"))   ' �iEPAN���x
        End With
    End If
    
    rs.Close

proc_exit:
    ' �I��
'    gErr.Pop
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funWfcGetDataEtc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'************************************************************************************
'*    �֐���        : funGetTBCMY013
'*
'*    �����T�v      : 1.�e�[�u���uTBCMY013�v��������ɂ��������R�[�h�𒊏o����
'*                      (����]�����ʎ擾)
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^                             ,����
'*                   udtTypIn      ,I  ,type_DBDRV_scmzc_fcmlc001c_In  ,���͗p
'*                   udtRecords()  ,O  ,typ_TBCMY013                   ,���o���R�[�h
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Private Function funGetTBCMY013(udtTypIn As type_DBDRV_scmzc_fcmlc001c_In, udtRecords() As typ_TBCMY013) As FUNCTION_RETURN
    Dim sSQL        As String       ' SQL�S��
    Dim rs          As OraDynaset   ' RecordSet
    Dim lngRecCnt   As Long         ' ���R�[�h��
    Dim i           As Long

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetTBCMY013"

    ' SQL��g�ݗ��Ă�
    sSQL = "select SAMPLEID, OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5, "
    sSQL = sSQL & "MESDATA6, MESDATA7, MESDATA8, MESDATA9, MESDATA10, MESDATA11, MESDATA12, MESDATA13, MESDATA14, MESDATA15, "
    sSQL = sSQL & "TXID, REGDATE, SENDFLAG, SENDDATE "
    sSQL = sSQL & "from TBCMY013 "
    sSQL = sSQL & "where ('" & udtTypIn.WFSMP.WFINDRSCW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDRSCW & "' and SPEC = '" & OSWFRES & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDOICW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDOICW & "' and SPEC = '" & OSWFOI & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDB1CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDB1CW & "' and SPEC = '" & OSWFBMD1 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDB2CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDB2CW & "' and SPEC = '" & OSWFBMD2 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDB3CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDB3CW & "' and SPEC = '" & OSWFBMD3 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDL1CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDL1CW & "' and SPEC = '" & OSWFOSF1 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDL2CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDL2CW & "' and SPEC = '" & OSWFOSF2 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDL3CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDL3CW & "' and SPEC = '" & OSWFOSF3 & "') or "
'    sSql = sSql & "      ('" & udtTypIn.WFSMP.WFINDL4CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDL4CW & "' and SPEC = '" & OSWFOSF4 & "') or "
'Upd 2010/01/07 SIRD�Ή� Y.Hitomi
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDL4CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDL4CW & "' and SPEC = '" & OSWFSIRD & "') or "
    
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDDSCW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDDSCW & "' and SPEC = '" & OSWFDS & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDDZCW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDDZCW & "' and SPEC = '" & OSWFDZ & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDDO1CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDDO1CW & "' and SPEC = '" & OSWFDOI1 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDDO2CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDDO2CW & "' and SPEC = '" & OSWFDOI2 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDDO3CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDDO3CW & "' and SPEC = '" & OSWFDOI3 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDOT1CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDOT1CW & "' and SPEC = '" & OSWFOT1 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDOT2CW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDOT2CW & "' and SPEC = '" & OSWFOT2 & "') or "
    sSQL = sSQL & "      ('" & udtTypIn.WFSMP.WFINDAOICW & "' > '0' and SAMPLEID = '" & udtTypIn.WFSMP.WFSMPLIDAOICW & "' and SPEC = '" & OSWFAOI & "')"
    
    Debug.Print sSQL
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        Set rs = Nothing
        ReDim udtRecords(0)
        funGetTBCMY013 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ' ���o���ʂ��i�[����
    lngRecCnt = rs.RecordCount
    ReDim udtRecords(lngRecCnt)
    For i = 1 To lngRecCnt
        With udtRecords(i)
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
    Set rs = Nothing

    funGetTBCMY013 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funGetTBCMY013 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'************************************************************************************
'*    �֐���        : funGetGDJisseki_J006
'*
'*    �����T�v      : 1.����GD����(TBCMJ006)�̎擾����
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^            ,����
'*                   sCryNum       ,I  ,String        ,���͗p
'*                   sSmplID       ,I  ,String        ,���o���R�[�h
'*                   udtGDjisseki  ,O  ,typ_TBCMJ015  ,����GD����(�\����)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Public Function funGetGDJisseki_J006(sCryNum As String, sSmplID As String, _
                                                    udtGDjisseki As typ_TBCMJ015) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim lngSmplID   As Long         ' �ް��^�ύX
    
    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetGDJisseki_J006"
    
    ' �����ID�����l�łȂ��ꍇ
    If IsNumeric(sSmplID) = False Then
        funGetGDJisseki_J006 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    lngSmplID = CLng(sSmplID)         '�ް��^�ύX
    
    ' �����ԍ��A�����ID����TBCMJ006�̌���GD���ђl����������B
    sSQL = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MSRSDEN, MSRSLDL, MSRSDVD2, "
    sSQL = sSQL & "MS01LDL1, MS01LDL2, MS01LDL3, MS01LDL4, MS01LDL5, MS01DEN1, MS01DEN2, MS01DEN3, MS01DEN4, MS01DEN5, "
    sSQL = sSQL & "MS02LDL1, MS02LDL2, MS02LDL3, MS02LDL4, MS02LDL5, MS02DEN1, MS02DEN2, MS02DEN3, MS02DEN4, MS02DEN5, "
    sSQL = sSQL & "MS03LDL1, MS03LDL2, MS03LDL3, MS03LDL4, MS03LDL5, MS03DEN1, MS03DEN2, MS03DEN3, MS03DEN4, MS03DEN5, "
    sSQL = sSQL & "MS04LDL1, MS04LDL2, MS04LDL3, MS04LDL4, MS04LDL5, MS04DEN1, MS04DEN2, MS04DEN3, MS04DEN4, MS04DEN5, "
    sSQL = sSQL & "MS05LDL1, MS05LDL2, MS05LDL3, MS05LDL4, MS05LDL5, MS05DEN1, MS05DEN2, MS05DEN3, MS05DEN4, MS05DEN5, "
    sSQL = sSQL & "MS06LDL1, MS06LDL2, MS06LDL3, MS06LDL4, MS06LDL5, MS06DEN1, MS06DEN2, MS06DEN3, MS06DEN4, MS06DEN5, "
    sSQL = sSQL & "MS07LDL1, MS07LDL2, MS07LDL3, MS07LDL4, MS07LDL5, MS07DEN1, MS07DEN2, MS07DEN3, MS07DEN4, MS07DEN5, "
    sSQL = sSQL & "MS08LDL1, MS08LDL2, MS08LDL3, MS08LDL4, MS08LDL5, MS08DEN1, MS08DEN2, MS08DEN3, MS08DEN4, MS08DEN5, "
    sSQL = sSQL & "MS09LDL1, MS09LDL2, MS09LDL3, MS09LDL4, MS09LDL5, MS09DEN1, MS09DEN2, MS09DEN3, MS09DEN4, MS09DEN5, "
    sSQL = sSQL & "MS10LDL1, MS10LDL2, MS10LDL3, MS10LDL4, MS10LDL5, MS10DEN1, MS10DEN2, MS10DEN3, MS10DEN4, MS10DEN5, "
    sSQL = sSQL & "MS11LDL1, MS11LDL2, MS11LDL3, MS11LDL4, MS11LDL5, MS11DEN1, MS11DEN2, MS11DEN3, MS11DEN4, MS11DEN5, "
    sSQL = sSQL & "MS12LDL1, MS12LDL2, MS12LDL3, MS12LDL4, MS12LDL5, MS12DEN1, MS12DEN2, MS12DEN3, MS12DEN4, MS12DEN5, "
    sSQL = sSQL & "MS13LDL1, MS13LDL2, MS13LDL3, MS13LDL4, MS13LDL5, MS13DEN1, MS13DEN2, MS13DEN3, MS13DEN4, MS13DEN5, "
    sSQL = sSQL & "MS14LDL1, MS14LDL2, MS14LDL3, MS14LDL4, MS14LDL5, MS14DEN1, MS14DEN2, MS14DEN3, MS14DEN4, MS14DEN5, "
    sSQL = sSQL & "MS15LDL1, MS15LDL2, MS15LDL3, MS15LDL4, MS15LDL5, MS15DEN1, MS15DEN2, MS15DEN3, MS15DEN4, MS15DEN5, "
    sSQL = sSQL & "MS01DVD2, MS02DVD2, MS03DVD2, MS04DVD2, MS05DVD2, REGDATE "
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    sSQL = sSQL & ", MSZEROMN, MSZEROMX "
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    sSQL = sSQL & "from TBCMJ006 "
    sSQL = sSQL & "where CRYNUM = '" & sCryNum & "' and "
    sSQL = sSQL & "      SMPLNO = " & lngSmplID & " and "
    sSQL = sSQL & "      TRANCNT = (select max(TRANCNT) from TBCMJ006 "
    sSQL = sSQL & "                 where CRYNUM = '" & sCryNum & "' and "
    sSQL = sSQL & "                       SMPLNO = " & lngSmplID & ")"
    
    ' SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' �Y���ް��Ȃ�
    If rs.EOF Then
        udtGDjisseki.SMPLNO = ""
        funGetGDJisseki_J006 = FUNCTION_RETURN_SUCCESS
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtGDjisseki
        .CRYNUM = rs("CRYNUM")          ' �����ԍ�
        .POSITION = rs("POSITION")      ' �ʒu
        .SMPKBN = rs("SMPKBN")          ' �T���v���敪
        .TRANCOND = rs("TRANCOND")      ' ��������
        .TRANCNT = rs("TRANCNT")        ' ������
        .HSFLG = "0"                    ' �ۏ؃t���O
        .SMPLNO = CStr(rs("SMPLNO"))    ' �T���v���m��
        .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
        .MSRSDEN = rs("MSRSDEN")        ' ���茋�� Den
        .MSRSLDL = rs("MSRSLDL")        ' ���茋�� L/DL
        .MSRSDVD2 = rs("MSRSDVD2")      ' ���茋�� DVD2
        .MS01LDL1 = rs("MS01LDL1")      ' ����l01 L/DL1
        .MS01LDL2 = rs("MS01LDL2")      ' ����l01 L/DL2
        .MS01LDL3 = rs("MS01LDL3")      ' ����l01 L/DL3
        .MS01LDL4 = rs("MS01LDL4")      ' ����l01 L/DL4
        .MS01LDL5 = rs("MS01LDL5")      ' ����l01 L/DL5
        .MS01DEN1 = rs("MS01DEN1")      ' ����l01 Den1
        .MS01DEN2 = rs("MS01DEN2")      ' ����l01 Den2
        .MS01DEN3 = rs("MS01DEN3")      ' ����l01 Den3
        .MS01DEN4 = rs("MS01DEN4")      ' ����l01 Den4
        .MS01DEN5 = rs("MS01DEN5")      ' ����l01 Den5
        .MS02LDL1 = rs("MS02LDL1")      ' ����l02 L/DL1
        .MS02LDL2 = rs("MS02LDL2")      ' ����l02 L/DL2
        .MS02LDL3 = rs("MS02LDL3")      ' ����l02 L/DL3
        .MS02LDL4 = rs("MS02LDL4")      ' ����l02 L/DL4
        .MS02LDL5 = rs("MS02LDL5")      ' ����l02 L/DL5
        .MS02DEN1 = rs("MS02DEN1")      ' ����l02 Den1
        .MS02DEN2 = rs("MS02DEN2")      ' ����l02 Den2
        .MS02DEN3 = rs("MS02DEN3")      ' ����l02 Den3
        .MS02DEN4 = rs("MS02DEN4")      ' ����l02 Den4
        .MS02DEN5 = rs("MS02DEN5")      ' ����l02 Den5
        .MS03LDL1 = rs("MS03LDL1")      ' ����l03 L/DL1
        .MS03LDL2 = rs("MS03LDL2")      ' ����l03 L/DL2
        .MS03LDL3 = rs("MS03LDL3")      ' ����l03 L/DL3
        .MS03LDL4 = rs("MS03LDL4")      ' ����l03 L/DL4
        .MS03LDL5 = rs("MS03LDL5")      ' ����l03 L/DL5
        .MS03DEN1 = rs("MS03DEN1")      ' ����l03 Den1
        .MS03DEN2 = rs("MS03DEN2")      ' ����l03 Den2
        .MS03DEN3 = rs("MS03DEN3")      ' ����l03 Den3
        .MS03DEN4 = rs("MS03DEN4")      ' ����l03 Den4
        .MS03DEN5 = rs("MS03DEN5")      ' ����l03 Den5
        .MS04LDL1 = rs("MS04LDL1")      ' ����l04 L/DL1
        .MS04LDL2 = rs("MS04LDL2")      ' ����l04 L/DL2
        .MS04LDL3 = rs("MS04LDL3")      ' ����l04 L/DL3
        .MS04LDL4 = rs("MS04LDL4")      ' ����l04 L/DL4
        .MS04LDL5 = rs("MS04LDL5")      ' ����l04 L/DL5
        .MS04DEN1 = rs("MS04DEN1")      ' ����l04 Den1
        .MS04DEN2 = rs("MS04DEN2")      ' ����l04 Den2
        .MS04DEN3 = rs("MS04DEN3")      ' ����l04 Den3
        .MS04DEN4 = rs("MS04DEN4")      ' ����l04 Den4
        .MS04DEN5 = rs("MS04DEN5")      ' ����l04 Den5
        .MS05LDL1 = rs("MS05LDL1")      ' ����l05 L/DL1
        .MS05LDL2 = rs("MS05LDL2")      ' ����l05 L/DL2
        .MS05LDL3 = rs("MS05LDL3")      ' ����l05 L/DL3
        .MS05LDL4 = rs("MS05LDL4")      ' ����l05 L/DL4
        .MS05LDL5 = rs("MS05LDL5")      ' ����l05 L/DL5
        .MS05DEN1 = rs("MS05DEN1")      ' ����l05 Den1
        .MS05DEN2 = rs("MS05DEN2")      ' ����l05 Den2
        .MS05DEN3 = rs("MS05DEN3")      ' ����l05 Den3
        .MS05DEN4 = rs("MS05DEN4")      ' ����l05 Den4
        .MS05DEN5 = rs("MS05DEN5")      ' ����l05 Den5
        .MS06LDL1 = rs("MS06LDL1")      ' ����l06 L/DL1
        .MS06LDL2 = rs("MS06LDL2")      ' ����l06 L/DL2
        .MS06LDL3 = rs("MS06LDL3")      ' ����l06 L/DL3
        .MS06LDL4 = rs("MS06LDL4")      ' ����l06 L/DL4
        .MS06LDL5 = rs("MS06LDL5")      ' ����l06 L/DL5
        .MS06DEN1 = rs("MS06DEN1")      ' ����l06 Den1
        .MS06DEN2 = rs("MS06DEN2")      ' ����l06 Den2
        .MS06DEN3 = rs("MS06DEN3")      ' ����l06 Den3
        .MS06DEN4 = rs("MS06DEN4")      ' ����l06 Den4
        .MS06DEN5 = rs("MS06DEN5")      ' ����l06 Den5
        .MS07LDL1 = rs("MS07LDL1")      ' ����l07 L/DL1
        .MS07LDL2 = rs("MS07LDL2")      ' ����l07 L/DL2
        .MS07LDL3 = rs("MS07LDL3")      ' ����l07 L/DL3
        .MS07LDL4 = rs("MS07LDL4")      ' ����l07 L/DL4
        .MS07LDL5 = rs("MS07LDL5")      ' ����l07 L/DL5
        .MS07DEN1 = rs("MS07DEN1")      ' ����l07 Den1
        .MS07DEN2 = rs("MS07DEN2")      ' ����l07 Den2
        .MS07DEN3 = rs("MS07DEN3")      ' ����l07 Den3
        .MS07DEN4 = rs("MS07DEN4")      ' ����l07 Den4
        .MS07DEN5 = rs("MS07DEN5")      ' ����l07 Den5
        .MS08LDL1 = rs("MS08LDL1")      ' ����l08 L/DL1
        .MS08LDL2 = rs("MS08LDL2")      ' ����l08 L/DL2
        .MS08LDL3 = rs("MS08LDL3")      ' ����l08 L/DL3
        .MS08LDL4 = rs("MS08LDL4")      ' ����l08 L/DL4
        .MS08LDL5 = rs("MS08LDL5")      ' ����l08 L/DL5
        .MS08DEN1 = rs("MS08DEN1")      ' ����l08 Den1
        .MS08DEN2 = rs("MS08DEN2")      ' ����l08 Den2
        .MS08DEN3 = rs("MS08DEN3")      ' ����l08 Den3
        .MS08DEN4 = rs("MS08DEN4")      ' ����l08 Den4
        .MS08DEN5 = rs("MS08DEN5")      ' ����l08 Den5
        .MS09LDL1 = rs("MS09LDL1")      ' ����l09 L/DL1
        .MS09LDL2 = rs("MS09LDL2")      ' ����l09 L/DL2
        .MS09LDL3 = rs("MS09LDL3")      ' ����l09 L/DL3
        .MS09LDL4 = rs("MS09LDL4")      ' ����l09 L/DL4
        .MS09LDL5 = rs("MS09LDL5")      ' ����l09 L/DL5
        .MS09DEN1 = rs("MS09DEN1")      ' ����l09 Den1
        .MS09DEN2 = rs("MS09DEN2")      ' ����l09 Den2
        .MS09DEN3 = rs("MS09DEN3")      ' ����l09 Den3
        .MS09DEN4 = rs("MS09DEN4")      ' ����l09 Den4
        .MS09DEN5 = rs("MS09DEN5")      ' ����l09 Den5
        .MS10LDL1 = rs("MS10LDL1")      ' ����l10 L/DL1
        .MS10LDL2 = rs("MS10LDL2")      ' ����l10 L/DL2
        .MS10LDL3 = rs("MS10LDL3")      ' ����l10 L/DL3
        .MS10LDL4 = rs("MS10LDL4")      ' ����l10 L/DL4
        .MS10LDL5 = rs("MS10LDL5")      ' ����l10 L/DL5
        .MS10DEN1 = rs("MS10DEN1")      ' ����l10 Den1
        .MS10DEN2 = rs("MS10DEN2")      ' ����l10 Den2
        .MS10DEN3 = rs("MS10DEN3")      ' ����l10 Den3
        .MS10DEN4 = rs("MS10DEN4")      ' ����l10 Den4
        .MS10DEN5 = rs("MS10DEN5")      ' ����l10 Den5
        .MS11LDL1 = rs("MS11LDL1")      ' ����l11 L/DL1
        .MS11LDL2 = rs("MS11LDL2")      ' ����l11 L/DL2
        .MS11LDL3 = rs("MS11LDL3")      ' ����l11 L/DL3
        .MS11LDL4 = rs("MS11LDL4")      ' ����l11 L/DL4
        .MS11LDL5 = rs("MS11LDL5")      ' ����l11 L/DL5
        .MS11DEN1 = rs("MS11DEN1")      ' ����l11 Den1
        .MS11DEN2 = rs("MS11DEN2")      ' ����l11 Den2
        .MS11DEN3 = rs("MS11DEN3")      ' ����l11 Den3
        .MS11DEN4 = rs("MS11DEN4")      ' ����l11 Den4
        .MS11DEN5 = rs("MS11DEN5")      ' ����l11 Den5
        .MS12LDL1 = rs("MS12LDL1")      ' ����l12 L/DL1
        .MS12LDL2 = rs("MS12LDL2")      ' ����l12 L/DL2
        .MS12LDL3 = rs("MS12LDL3")      ' ����l12 L/DL3
        .MS12LDL4 = rs("MS12LDL4")      ' ����l12 L/DL4
        .MS12LDL5 = rs("MS12LDL5")      ' ����l12 L/DL5
        .MS12DEN1 = rs("MS12DEN1")      ' ����l12 Den1
        .MS12DEN2 = rs("MS12DEN2")      ' ����l12 Den2
        .MS12DEN3 = rs("MS12DEN3")      ' ����l12 Den3
        .MS12DEN4 = rs("MS12DEN4")      ' ����l12 Den4
        .MS12DEN5 = rs("MS12DEN5")      ' ����l12 Den5
        .MS13LDL1 = rs("MS13LDL1")      ' ����l13 L/DL1
        .MS13LDL2 = rs("MS13LDL2")      ' ����l13 L/DL2
        .MS13LDL3 = rs("MS13LDL3")      ' ����l13 L/DL3
        .MS13LDL4 = rs("MS13LDL4")      ' ����l13 L/DL4
        .MS13LDL5 = rs("MS13LDL5")      ' ����l13 L/DL5
        .MS13DEN1 = rs("MS13DEN1")      ' ����l13 Den1
        .MS13DEN2 = rs("MS13DEN2")      ' ����l13 Den2
        .MS13DEN3 = rs("MS13DEN3")      ' ����l13 Den3
        .MS13DEN4 = rs("MS13DEN4")      ' ����l13 Den4
        .MS13DEN5 = rs("MS13DEN5")      ' ����l13 Den5
        .MS14LDL1 = rs("MS14LDL1")      ' ����l14 L/DL1
        .MS14LDL2 = rs("MS14LDL2")      ' ����l14 L/DL2
        .MS14LDL3 = rs("MS14LDL3")      ' ����l14 L/DL3
        .MS14LDL4 = rs("MS14LDL4")      ' ����l14 L/DL4
        .MS14LDL5 = rs("MS14LDL5")      ' ����l14 L/DL5
        .MS14DEN1 = rs("MS14DEN1")      ' ����l14 Den1
        .MS14DEN2 = rs("MS14DEN2")      ' ����l14 Den2
        .MS14DEN3 = rs("MS14DEN3")      ' ����l14 Den3
        .MS14DEN4 = rs("MS14DEN4")      ' ����l14 Den4
        .MS14DEN5 = rs("MS14DEN5")      ' ����l14 Den5
        .MS15LDL1 = rs("MS15LDL1")      ' ����l15 L/DL1
        .MS15LDL2 = rs("MS15LDL2")      ' ����l15 L/DL2
        .MS15LDL3 = rs("MS15LDL3")      ' ����l15 L/DL3
        .MS15LDL4 = rs("MS15LDL4")      ' ����l15 L/DL4
        .MS15LDL5 = rs("MS15LDL5")      ' ����l15 L/DL5
        .MS15DEN1 = rs("MS15DEN1")      ' ����l15 Den1
        .MS15DEN2 = rs("MS15DEN2")      ' ����l15 Den2
        .MS15DEN3 = rs("MS15DEN3")      ' ����l15 Den3
        .MS15DEN4 = rs("MS15DEN4")      ' ����l15 Den4
        .MS15DEN5 = rs("MS15DEN5")      ' ����l15 Den5
        If IsNull(rs("MS01DVD2")) = False Then .MS01DVD2 = rs("MS01DVD2") Else .MS01DVD2 = -1   ' ����l01 DVD2
        If IsNull(rs("MS02DVD2")) = False Then .MS02DVD2 = rs("MS02DVD2") Else .MS02DVD2 = -1   ' ����l02 DVD2
        If IsNull(rs("MS03DVD2")) = False Then .MS03DVD2 = rs("MS03DVD2") Else .MS03DVD2 = -1   ' ����l03 DVD2
        If IsNull(rs("MS04DVD2")) = False Then .MS04DVD2 = rs("MS04DVD2") Else .MS04DVD2 = -1   ' ����l04 DVD2
        If IsNull(rs("MS05DVD2")) = False Then .MS05DVD2 = rs("MS05DVD2") Else .MS05DVD2 = -1   ' ����l05 DVD2
        .REGDATE = rs("REGDATE")        ' �o�^���t
        
        '���ǉ� �M�������f�����ǉ�
        '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
        '�����̃f�[�^��AN���x�������Ă��Ȃ��̂ŁA�\�����Ȃ��悤�ɂ���
        'DBData2DispDate�Ńf�[�^���`���Ă���̂ŁA����ɂ��킹��-1�������
        .DKAN = "  -1"
        
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        If IsNull(rs("MSZEROMN")) = False Then .MSZEROMN = rs("MSZEROMN") Else .MSZEROMN = -1   ' L/DL0�A�����ŏ��l
        If IsNull(rs("MSZEROMX")) = False Then .MSZEROMX = rs("MSZEROMX") Else .MSZEROMX = -1   ' L/DL0�A�����ő�l
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
        
    End With
    
    Set rs = Nothing

    funGetGDJisseki_J006 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetGDJisseki_J006 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'************************************************************************************
'*    �֐���        : funGetGDJisseki_J015
'*
'*    �����T�v      : 1.GD����(TBCMJ015)�̎擾����
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^            ,����
'*                   sCryNum       ,I  ,String        ,���͗p
'*                   sSmplID       ,I  ,String        ,���o���R�[�h
'*                   sHsFlg_XSDCW  ,I  ,String        ,�ۏ�FLG(0:WF���сA1:��������)
'*                   udtGDjisseki  ,O  ,typ_TBCMJ015  ,����GD����(�\����)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Public Function funGetGDJisseki_J015(sCryNum As String, sSmplID As String, _
                                        sHsFlg_XSDCW As String, udtGDjisseki As typ_TBCMJ015) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim intSmplID       As Integer
    Dim sHsFlg_J015     As String       ' �ۏ�FLG(1:WF����)
    
    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetGDJisseki_J015"
    
    If sHsFlg_XSDCW = "0" Then
        ' WF����
        sHsFlg_J015 = "1"
    Else
        ' ��������
        sHsFlg_J015 = "0"
    End If
    
    '�����ԍ��A�����ID�A�ۏ�FLG����TBCMJ015��GD���ђl����������B
    sSQL = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, HSFLG, SMPLNO, SMPLUMU, "
    sSQL = sSQL & "HINBAN, REVNUM, FACTORY, OPECOND, SXLID, KRPROCCD, PROCCODE, GOUKI, "
    sSQL = sSQL & "OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, ETMAE_RYO01, ETATO_RYO01, MSRSDEN, MSRSLDL, MSRSDVD2, "
    sSQL = sSQL & "MS01LDL1, MS01LDL2, MS01LDL3, MS01LDL4, MS01LDL5, MS01DEN1, MS01DEN2, MS01DEN3, MS01DEN4, MS01DEN5, "
    sSQL = sSQL & "MS02LDL1, MS02LDL2, MS02LDL3, MS02LDL4, MS02LDL5, MS02DEN1, MS02DEN2, MS02DEN3, MS02DEN4, MS02DEN5, "
    sSQL = sSQL & "MS03LDL1, MS03LDL2, MS03LDL3, MS03LDL4, MS03LDL5, MS03DEN1, MS03DEN2, MS03DEN3, MS03DEN4, MS03DEN5, "
    sSQL = sSQL & "MS04LDL1, MS04LDL2, MS04LDL3, MS04LDL4, MS04LDL5, MS04DEN1, MS04DEN2, MS04DEN3, MS04DEN4, MS04DEN5, "
    sSQL = sSQL & "MS05LDL1, MS05LDL2, MS05LDL3, MS05LDL4, MS05LDL5, MS05DEN1, MS05DEN2, MS05DEN3, MS05DEN4, MS05DEN5, "
    sSQL = sSQL & "MS06LDL1, MS06LDL2, MS06LDL3, MS06LDL4, MS06LDL5, MS06DEN1, MS06DEN2, MS06DEN3, MS06DEN4, MS06DEN5, "
    sSQL = sSQL & "MS07LDL1, MS07LDL2, MS07LDL3, MS07LDL4, MS07LDL5, MS07DEN1, MS07DEN2, MS07DEN3, MS07DEN4, MS07DEN5, "
    sSQL = sSQL & "MS08LDL1, MS08LDL2, MS08LDL3, MS08LDL4, MS08LDL5, MS08DEN1, MS08DEN2, MS08DEN3, MS08DEN4, MS08DEN5, "
    sSQL = sSQL & "MS09LDL1, MS09LDL2, MS09LDL3, MS09LDL4, MS09LDL5, MS09DEN1, MS09DEN2, MS09DEN3, MS09DEN4, MS09DEN5, "
    sSQL = sSQL & "MS10LDL1, MS10LDL2, MS10LDL3, MS10LDL4, MS10LDL5, MS10DEN1, MS10DEN2, MS10DEN3, MS10DEN4, MS10DEN5, "
    sSQL = sSQL & "MS11LDL1, MS11LDL2, MS11LDL3, MS11LDL4, MS11LDL5, MS11DEN1, MS11DEN2, MS11DEN3, MS11DEN4, MS11DEN5, "
    sSQL = sSQL & "MS12LDL1, MS12LDL2, MS12LDL3, MS12LDL4, MS12LDL5, MS12DEN1, MS12DEN2, MS12DEN3, MS12DEN4, MS12DEN5, "
    sSQL = sSQL & "MS13LDL1, MS13LDL2, MS13LDL3, MS13LDL4, MS13LDL5, MS13DEN1, MS13DEN2, MS13DEN3, MS13DEN4, MS13DEN5, "
    sSQL = sSQL & "MS14LDL1, MS14LDL2, MS14LDL3, MS14LDL4, MS14LDL5, MS14DEN1, MS14DEN2, MS14DEN3, MS14DEN4, MS14DEN5, "
    sSQL = sSQL & "MS15LDL1, MS15LDL2, MS15LDL3, MS15LDL4, MS15LDL5, MS15DEN1, MS15DEN2, MS15DEN3, MS15DEN4, MS15DEN5, "
    sSQL = sSQL & "MS01DVD2, MS02DVD2, MS03DVD2, MS04DVD2, MS05DVD2, "
    sSQL = sSQL & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    sSQL = sSQL & ", MSZEROMN , MSZEROMX "
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    sSQL = sSQL & "from TBCMJ015 "
    sSQL = sSQL & "where CRYNUM = '" & sCryNum & "' and "
    sSQL = sSQL & "      SMPLNO = '" & sSmplID & "' and "
    sSQL = sSQL & "      HSFLG = '" & sHsFlg_J015 & "' and "
    sSQL = sSQL & "      TRANCNT = (select max(TRANCNT) from TBCMJ015 "
    sSQL = sSQL & "                 where CRYNUM = '" & sCryNum & "' and "
    sSQL = sSQL & "                       SMPLNO = '" & sSmplID & "' and "
    sSQL = sSQL & "                       HSFLG = '" & sHsFlg_J015 & "')"
    
    ' SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' �Y���ް��Ȃ�
    If rs.EOF Then
        udtGDjisseki.SMPLNO = ""
        funGetGDJisseki_J015 = FUNCTION_RETURN_SUCCESS
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtGDjisseki
        If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")                             ' �����ԍ�
        If IsNull(rs("POSITION")) = False Then .POSITION = rs("POSITION")                       ' �ʒu
        If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")                             ' �T���v���敪
        If IsNull(rs("TRANCOND")) = False Then .TRANCOND = rs("TRANCOND")                       ' ��������
        If IsNull(rs("TRANCNT")) = False Then .TRANCNT = rs("TRANCNT")                          ' ������
        If IsNull(rs("HSFLG")) = False Then .HSFLG = rs("HSFLG")                                ' �ۏ؃t���O
        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                             ' �T���v���m��
        If IsNull(rs("SMPLUMU")) = False Then .SMPLUMU = rs("SMPLUMU")                          ' �T���v���L��
        If IsNull(rs("HINBAN")) = False Then .hinban = rs("HINBAN")                             ' �i��
        If IsNull(rs("REVNUM")) = False Then .REVNUM = rs("REVNUM")                             ' ���i�ԍ������ԍ�
        If IsNull(rs("FACTORY")) = False Then .factory = rs("FACTORY")                         ' �H��
        If IsNull(rs("OPECOND")) = False Then .opecond = rs("OPECOND")                          ' ���Ə���
        If IsNull(rs("SXLID")) = False Then .SXLID = rs("SXLID")                                ' SXLID
        If IsNull(rs("KRPROCCD")) = False Then .KRPROCCD = rs("KRPROCCD")                       ' �Ǘ��H���R�[�h
        If IsNull(rs("PROCCODE")) = False Then .PROCCODE = rs("PROCCODE")                       ' �H���R�[�h
        If IsNull(rs("GOUKI")) = False Then .GOUKI = rs("GOUKI")                                ' ���@
        If IsNull(rs("OSITEM")) = False Then .OSITEM = rs("OSITEM")                             ' �]������
        If IsNull(rs("MAISU")) = False Then .MAISU = rs("MAISU")                                ' �]������
        If IsNull(rs("SPEC")) = False Then .Spec = rs("SPEC")                                   ' �K�i�l
        If IsNull(rs("NETSU")) = False Then .NETSU = rs("NETSU")                                ' �M��������
        If IsNull(rs("ET")) = False Then .ET = rs("ET")                                         ' �G�b�`���O����
        If IsNull(rs("MES")) = False Then .MES = rs("MES")                                      ' �v�����@
        If IsNull(rs("DKAN")) = False Then .DKAN = rs("DKAN")                                   ' �c�j�A�j�[������
        If IsNull(rs("ETMAE_RYO01")) = False Then .ETMAE_RYO01 = rs("ETMAE_RYO01")              ' ET�O�d��01
        If IsNull(rs("ETATO_RYO01")) = False Then .ETATO_RYO01 = rs("ETATO_RYO01")              ' ET��d��01
        
        If IsNull(rs("MSRSDEN")) = False Then .MSRSDEN = rs("MSRSDEN") Else .MSRSDEN = -1       ' ���茋�� Den
        If IsNull(rs("MSRSLDL")) = False Then .MSRSLDL = rs("MSRSLDL") Else .MSRSLDL = -1       ' ���茋�� L/DL
        If IsNull(rs("MSRSDVD2")) = False Then .MSRSDVD2 = rs("MSRSDVD2") Else .MSRSDVD2 = -1   ' ���茋�� DVD2
        If IsNull(rs("MS01LDL1")) = False Then .MS01LDL1 = rs("MS01LDL1") Else .MS01LDL1 = -1   ' ����l01 L/DL1
        If IsNull(rs("MS01LDL2")) = False Then .MS01LDL2 = rs("MS01LDL2") Else .MS01LDL2 = -1   ' ����l01 L/DL2
        If IsNull(rs("MS01LDL3")) = False Then .MS01LDL3 = rs("MS01LDL3") Else .MS01LDL3 = -1   ' ����l01 L/DL3
        If IsNull(rs("MS01LDL4")) = False Then .MS01LDL4 = rs("MS01LDL4") Else .MS01LDL4 = -1   ' ����l01 L/DL4
        If IsNull(rs("MS01LDL5")) = False Then .MS01LDL5 = rs("MS01LDL5") Else .MS01LDL5 = -1   ' ����l01 L/DL5
        If IsNull(rs("MS01DEN1")) = False Then .MS01DEN1 = rs("MS01DEN1") Else .MS01DEN1 = -1   ' ����l01 Den1
        If IsNull(rs("MS01DEN2")) = False Then .MS01DEN2 = rs("MS01DEN2") Else .MS01DEN2 = -1   ' ����l01 Den2
        If IsNull(rs("MS01DEN3")) = False Then .MS01DEN3 = rs("MS01DEN3") Else .MS01DEN3 = -1   ' ����l01 Den3
        If IsNull(rs("MS01DEN4")) = False Then .MS01DEN4 = rs("MS01DEN4") Else .MS01DEN4 = -1   ' ����l01 Den4
        If IsNull(rs("MS01DEN5")) = False Then .MS01DEN5 = rs("MS01DEN5") Else .MS01DEN5 = -1   ' ����l01 Den5
        If IsNull(rs("MS02LDL1")) = False Then .MS02LDL1 = rs("MS02LDL1") Else .MS02LDL1 = -1   ' ����l02 L/DL1
        If IsNull(rs("MS02LDL2")) = False Then .MS02LDL2 = rs("MS02LDL2") Else .MS02LDL2 = -1   ' ����l02 L/DL2
        If IsNull(rs("MS02LDL3")) = False Then .MS02LDL3 = rs("MS02LDL3") Else .MS02LDL3 = -1   ' ����l02 L/DL3
        If IsNull(rs("MS02LDL4")) = False Then .MS02LDL4 = rs("MS02LDL4") Else .MS02LDL4 = -1   ' ����l02 L/DL4
        If IsNull(rs("MS02LDL5")) = False Then .MS02LDL5 = rs("MS02LDL5") Else .MS02LDL5 = -1   ' ����l02 L/DL5
        If IsNull(rs("MS02DEN1")) = False Then .MS02DEN1 = rs("MS02DEN1") Else .MS02DEN1 = -1   ' ����l02 Den1
        If IsNull(rs("MS02DEN2")) = False Then .MS02DEN2 = rs("MS02DEN2") Else .MS02DEN2 = -1   ' ����l02 Den2
        If IsNull(rs("MS02DEN3")) = False Then .MS02DEN3 = rs("MS02DEN3") Else .MS02DEN3 = -1   ' ����l02 Den3
        If IsNull(rs("MS02DEN4")) = False Then .MS02DEN4 = rs("MS02DEN4") Else .MS02DEN4 = -1   ' ����l02 Den4
        If IsNull(rs("MS02DEN5")) = False Then .MS02DEN5 = rs("MS02DEN5") Else .MS02DEN5 = -1   ' ����l02 Den5
        If IsNull(rs("MS03LDL1")) = False Then .MS03LDL1 = rs("MS03LDL1") Else .MS03LDL1 = -1   ' ����l03 L/DL1
        If IsNull(rs("MS03LDL2")) = False Then .MS03LDL2 = rs("MS03LDL2") Else .MS03LDL2 = -1   ' ����l03 L/DL2
        If IsNull(rs("MS03LDL3")) = False Then .MS03LDL3 = rs("MS03LDL3") Else .MS03LDL3 = -1   ' ����l03 L/DL3
        If IsNull(rs("MS03LDL4")) = False Then .MS03LDL4 = rs("MS03LDL4") Else .MS03LDL4 = -1   ' ����l03 L/DL4
        If IsNull(rs("MS03LDL5")) = False Then .MS03LDL5 = rs("MS03LDL5") Else .MS03LDL5 = -1   ' ����l03 L/DL5
        If IsNull(rs("MS03DEN1")) = False Then .MS03DEN1 = rs("MS03DEN1") Else .MS03DEN1 = -1   ' ����l03 Den1
        If IsNull(rs("MS03DEN2")) = False Then .MS03DEN2 = rs("MS03DEN2") Else .MS03DEN2 = -1   ' ����l03 Den2
        If IsNull(rs("MS03DEN3")) = False Then .MS03DEN3 = rs("MS03DEN3") Else .MS03DEN3 = -1   ' ����l03 Den3
        If IsNull(rs("MS03DEN4")) = False Then .MS03DEN4 = rs("MS03DEN4") Else .MS03DEN4 = -1   ' ����l03 Den4
        If IsNull(rs("MS03DEN5")) = False Then .MS03DEN5 = rs("MS03DEN5") Else .MS03DEN5 = -1   ' ����l03 Den5
        If IsNull(rs("MS04LDL1")) = False Then .MS04LDL1 = rs("MS04LDL1") Else .MS04LDL1 = -1   ' ����l04 L/DL1
        If IsNull(rs("MS04LDL2")) = False Then .MS04LDL2 = rs("MS04LDL2") Else .MS04LDL2 = -1   ' ����l04 L/DL2
        If IsNull(rs("MS04LDL3")) = False Then .MS04LDL3 = rs("MS04LDL3") Else .MS04LDL3 = -1   ' ����l04 L/DL3
        If IsNull(rs("MS04LDL4")) = False Then .MS04LDL4 = rs("MS04LDL4") Else .MS04LDL4 = -1   ' ����l04 L/DL4
        If IsNull(rs("MS04LDL5")) = False Then .MS04LDL5 = rs("MS04LDL5") Else .MS04LDL5 = -1   ' ����l04 L/DL5
        If IsNull(rs("MS04DEN1")) = False Then .MS04DEN1 = rs("MS04DEN1") Else .MS04DEN1 = -1   ' ����l04 Den1
        If IsNull(rs("MS04DEN2")) = False Then .MS04DEN2 = rs("MS04DEN2") Else .MS04DEN2 = -1   ' ����l04 Den2
        If IsNull(rs("MS04DEN3")) = False Then .MS04DEN3 = rs("MS04DEN3") Else .MS04DEN3 = -1   ' ����l04 Den3
        If IsNull(rs("MS04DEN4")) = False Then .MS04DEN4 = rs("MS04DEN4") Else .MS04DEN4 = -1   ' ����l04 Den4
        If IsNull(rs("MS04DEN5")) = False Then .MS04DEN5 = rs("MS04DEN5") Else .MS04DEN5 = -1   ' ����l04 Den5
        If IsNull(rs("MS05LDL1")) = False Then .MS05LDL1 = rs("MS05LDL1") Else .MS05LDL1 = -1   ' ����l05 L/DL1
        If IsNull(rs("MS05LDL2")) = False Then .MS05LDL2 = rs("MS05LDL2") Else .MS05LDL2 = -1   ' ����l05 L/DL2
        If IsNull(rs("MS05LDL3")) = False Then .MS05LDL3 = rs("MS05LDL3") Else .MS05LDL3 = -1   ' ����l05 L/DL3
        If IsNull(rs("MS05LDL4")) = False Then .MS05LDL4 = rs("MS05LDL4") Else .MS05LDL4 = -1   ' ����l05 L/DL4
        If IsNull(rs("MS05LDL5")) = False Then .MS05LDL5 = rs("MS05LDL5") Else .MS05LDL5 = -1   ' ����l05 L/DL5
        If IsNull(rs("MS05DEN1")) = False Then .MS05DEN1 = rs("MS05DEN1") Else .MS05DEN1 = -1   ' ����l05 Den1
        If IsNull(rs("MS05DEN2")) = False Then .MS05DEN2 = rs("MS05DEN2") Else .MS05DEN2 = -1   ' ����l05 Den2
        If IsNull(rs("MS05DEN3")) = False Then .MS05DEN3 = rs("MS05DEN3") Else .MS05DEN3 = -1   ' ����l05 Den3
        If IsNull(rs("MS05DEN4")) = False Then .MS05DEN4 = rs("MS05DEN4") Else .MS05DEN4 = -1   ' ����l05 Den4
        If IsNull(rs("MS05DEN5")) = False Then .MS05DEN5 = rs("MS05DEN5") Else .MS05DEN5 = -1   ' ����l05 Den5
        If IsNull(rs("MS06LDL1")) = False Then .MS06LDL1 = rs("MS06LDL1") Else .MS06LDL1 = -1   ' ����l06 L/DL1
        If IsNull(rs("MS06LDL2")) = False Then .MS06LDL2 = rs("MS06LDL2") Else .MS06LDL2 = -1   ' ����l06 L/DL2
        If IsNull(rs("MS06LDL3")) = False Then .MS06LDL3 = rs("MS06LDL3") Else .MS06LDL3 = -1   ' ����l06 L/DL3
        If IsNull(rs("MS06LDL4")) = False Then .MS06LDL4 = rs("MS06LDL4") Else .MS06LDL4 = -1   ' ����l06 L/DL4
        If IsNull(rs("MS06LDL5")) = False Then .MS06LDL5 = rs("MS06LDL5") Else .MS06LDL5 = -1   ' ����l06 L/DL5
        If IsNull(rs("MS06DEN1")) = False Then .MS06DEN1 = rs("MS06DEN1") Else .MS06DEN1 = -1   ' ����l06 Den1
        If IsNull(rs("MS06DEN2")) = False Then .MS06DEN2 = rs("MS06DEN2") Else .MS06DEN2 = -1   ' ����l06 Den2
        If IsNull(rs("MS06DEN3")) = False Then .MS06DEN3 = rs("MS06DEN3") Else .MS06DEN3 = -1   ' ����l06 Den3
        If IsNull(rs("MS06DEN4")) = False Then .MS06DEN4 = rs("MS06DEN4") Else .MS06DEN4 = -1   ' ����l06 Den4
        If IsNull(rs("MS06DEN5")) = False Then .MS06DEN5 = rs("MS06DEN5") Else .MS06DEN5 = -1   ' ����l06 Den5
        If IsNull(rs("MS07LDL1")) = False Then .MS07LDL1 = rs("MS07LDL1") Else .MS07LDL1 = -1   ' ����l07 L/DL1
        If IsNull(rs("MS07LDL2")) = False Then .MS07LDL2 = rs("MS07LDL2") Else .MS07LDL2 = -1   ' ����l07 L/DL2
        If IsNull(rs("MS07LDL3")) = False Then .MS07LDL3 = rs("MS07LDL3") Else .MS07LDL3 = -1   ' ����l07 L/DL3
        If IsNull(rs("MS07LDL4")) = False Then .MS07LDL4 = rs("MS07LDL4") Else .MS07LDL4 = -1   ' ����l07 L/DL4
        If IsNull(rs("MS07LDL5")) = False Then .MS07LDL5 = rs("MS07LDL5") Else .MS07LDL5 = -1   ' ����l07 L/DL5
        If IsNull(rs("MS07DEN1")) = False Then .MS07DEN1 = rs("MS07DEN1") Else .MS07DEN1 = -1   ' ����l07 Den1
        If IsNull(rs("MS07DEN2")) = False Then .MS07DEN2 = rs("MS07DEN2") Else .MS07DEN2 = -1   ' ����l07 Den2
        If IsNull(rs("MS07DEN3")) = False Then .MS07DEN3 = rs("MS07DEN3") Else .MS07DEN3 = -1   ' ����l07 Den3
        If IsNull(rs("MS07DEN4")) = False Then .MS07DEN4 = rs("MS07DEN4") Else .MS07DEN4 = -1   ' ����l07 Den4
        If IsNull(rs("MS07DEN5")) = False Then .MS07DEN5 = rs("MS07DEN5") Else .MS07DEN5 = -1   ' ����l07 Den5
        If IsNull(rs("MS08LDL1")) = False Then .MS08LDL1 = rs("MS08LDL1") Else .MS08LDL1 = -1   ' ����l08 L/DL1
        If IsNull(rs("MS08LDL2")) = False Then .MS08LDL2 = rs("MS08LDL2") Else .MS08LDL2 = -1   ' ����l08 L/DL2
        If IsNull(rs("MS08LDL3")) = False Then .MS08LDL3 = rs("MS08LDL3") Else .MS08LDL3 = -1   ' ����l08 L/DL3
        If IsNull(rs("MS08LDL4")) = False Then .MS08LDL4 = rs("MS08LDL4") Else .MS08LDL4 = -1   ' ����l08 L/DL4
        If IsNull(rs("MS08LDL5")) = False Then .MS08LDL5 = rs("MS08LDL5") Else .MS08LDL5 = -1   ' ����l08 L/DL5
        If IsNull(rs("MS08DEN1")) = False Then .MS08DEN1 = rs("MS08DEN1") Else .MS08DEN1 = -1   ' ����l08 Den1
        If IsNull(rs("MS08DEN2")) = False Then .MS08DEN2 = rs("MS08DEN2") Else .MS08DEN2 = -1   ' ����l08 Den2
        If IsNull(rs("MS08DEN3")) = False Then .MS08DEN3 = rs("MS08DEN3") Else .MS08DEN3 = -1   ' ����l08 Den3
        If IsNull(rs("MS08DEN4")) = False Then .MS08DEN4 = rs("MS08DEN4") Else .MS08DEN4 = -1   ' ����l08 Den4
        If IsNull(rs("MS08DEN5")) = False Then .MS08DEN5 = rs("MS08DEN5") Else .MS08DEN5 = -1   ' ����l08 Den5
        If IsNull(rs("MS09LDL1")) = False Then .MS09LDL1 = rs("MS09LDL1") Else .MS09LDL1 = -1   ' ����l09 L/DL1
        If IsNull(rs("MS09LDL2")) = False Then .MS09LDL2 = rs("MS09LDL2") Else .MS09LDL2 = -1   ' ����l09 L/DL2
        If IsNull(rs("MS09LDL3")) = False Then .MS09LDL3 = rs("MS09LDL3") Else .MS09LDL3 = -1   ' ����l09 L/DL3
        If IsNull(rs("MS09LDL4")) = False Then .MS09LDL4 = rs("MS09LDL4") Else .MS09LDL4 = -1   ' ����l09 L/DL4
        If IsNull(rs("MS09LDL5")) = False Then .MS09LDL5 = rs("MS09LDL5") Else .MS09LDL5 = -1   ' ����l09 L/DL5
        If IsNull(rs("MS09DEN1")) = False Then .MS09DEN1 = rs("MS09DEN1") Else .MS09DEN1 = -1   ' ����l09 Den1
        If IsNull(rs("MS09DEN2")) = False Then .MS09DEN2 = rs("MS09DEN2") Else .MS09DEN2 = -1   ' ����l09 Den2
        If IsNull(rs("MS09DEN3")) = False Then .MS09DEN3 = rs("MS09DEN3") Else .MS09DEN3 = -1   ' ����l09 Den3
        If IsNull(rs("MS09DEN4")) = False Then .MS09DEN4 = rs("MS09DEN4") Else .MS09DEN4 = -1   ' ����l09 Den4
        If IsNull(rs("MS09DEN5")) = False Then .MS09DEN5 = rs("MS09DEN5") Else .MS09DEN5 = -1   ' ����l09 Den5
        If IsNull(rs("MS10LDL1")) = False Then .MS10LDL1 = rs("MS10LDL1") Else .MS10LDL1 = -1   ' ����l10 L/DL1
        If IsNull(rs("MS10LDL2")) = False Then .MS10LDL2 = rs("MS10LDL2") Else .MS10LDL2 = -1   ' ����l10 L/DL2
        If IsNull(rs("MS10LDL3")) = False Then .MS10LDL3 = rs("MS10LDL3") Else .MS10LDL3 = -1   ' ����l10 L/DL3
        If IsNull(rs("MS10LDL4")) = False Then .MS10LDL4 = rs("MS10LDL4") Else .MS10LDL4 = -1   ' ����l10 L/DL4
        If IsNull(rs("MS10LDL5")) = False Then .MS10LDL5 = rs("MS10LDL5") Else .MS10LDL5 = -1   ' ����l10 L/DL5
        If IsNull(rs("MS10DEN1")) = False Then .MS10DEN1 = rs("MS10DEN1") Else .MS10DEN1 = -1   ' ����l10 Den1
        If IsNull(rs("MS10DEN2")) = False Then .MS10DEN2 = rs("MS10DEN2") Else .MS10DEN2 = -1   ' ����l10 Den2
        If IsNull(rs("MS10DEN3")) = False Then .MS10DEN3 = rs("MS10DEN3") Else .MS10DEN3 = -1   ' ����l10 Den3
        If IsNull(rs("MS10DEN4")) = False Then .MS10DEN4 = rs("MS10DEN4") Else .MS10DEN4 = -1   ' ����l10 Den4
        If IsNull(rs("MS10DEN5")) = False Then .MS10DEN5 = rs("MS10DEN5") Else .MS10DEN5 = -1   ' ����l10 Den5
        If IsNull(rs("MS11LDL1")) = False Then .MS11LDL1 = rs("MS11LDL1") Else .MS11LDL1 = -1   ' ����l11 L/DL1
        If IsNull(rs("MS11LDL2")) = False Then .MS11LDL2 = rs("MS11LDL2") Else .MS11LDL2 = -1   ' ����l11 L/DL2
        If IsNull(rs("MS11LDL3")) = False Then .MS11LDL3 = rs("MS11LDL3") Else .MS11LDL3 = -1   ' ����l11 L/DL3
        If IsNull(rs("MS11LDL4")) = False Then .MS11LDL4 = rs("MS11LDL4") Else .MS11LDL4 = -1   ' ����l11 L/DL4
        If IsNull(rs("MS11LDL5")) = False Then .MS11LDL5 = rs("MS11LDL5") Else .MS11LDL5 = -1   ' ����l11 L/DL5
        If IsNull(rs("MS11DEN1")) = False Then .MS11DEN1 = rs("MS11DEN1") Else .MS11DEN1 = -1   ' ����l11 Den1
        If IsNull(rs("MS11DEN2")) = False Then .MS11DEN2 = rs("MS11DEN2") Else .MS11DEN2 = -1   ' ����l11 Den2
        If IsNull(rs("MS11DEN3")) = False Then .MS11DEN3 = rs("MS11DEN3") Else .MS11DEN3 = -1   ' ����l11 Den3
        If IsNull(rs("MS11DEN4")) = False Then .MS11DEN4 = rs("MS11DEN4") Else .MS11DEN4 = -1   ' ����l11 Den4
        If IsNull(rs("MS11DEN5")) = False Then .MS11DEN5 = rs("MS11DEN5") Else .MS11DEN5 = -1   ' ����l11 Den5
        If IsNull(rs("MS12LDL1")) = False Then .MS12LDL1 = rs("MS12LDL1") Else .MS12LDL1 = -1   ' ����l12 L/DL1
        If IsNull(rs("MS12LDL2")) = False Then .MS12LDL2 = rs("MS12LDL2") Else .MS12LDL2 = -1   ' ����l12 L/DL2
        If IsNull(rs("MS12LDL3")) = False Then .MS12LDL3 = rs("MS12LDL3") Else .MS12LDL3 = -1   ' ����l12 L/DL3
        If IsNull(rs("MS12LDL4")) = False Then .MS12LDL4 = rs("MS12LDL4") Else .MS12LDL4 = -1   ' ����l12 L/DL4
        If IsNull(rs("MS12LDL5")) = False Then .MS12LDL5 = rs("MS12LDL5") Else .MS12LDL5 = -1   ' ����l12 L/DL5
        If IsNull(rs("MS12DEN1")) = False Then .MS12DEN1 = rs("MS12DEN1") Else .MS12DEN1 = -1   ' ����l12 Den1
        If IsNull(rs("MS12DEN2")) = False Then .MS12DEN2 = rs("MS12DEN2") Else .MS12DEN2 = -1   ' ����l12 Den2
        If IsNull(rs("MS12DEN3")) = False Then .MS12DEN3 = rs("MS12DEN3") Else .MS12DEN3 = -1   ' ����l12 Den3
        If IsNull(rs("MS12DEN4")) = False Then .MS12DEN4 = rs("MS12DEN4") Else .MS12DEN4 = -1   ' ����l12 Den4
        If IsNull(rs("MS12DEN5")) = False Then .MS12DEN5 = rs("MS12DEN5") Else .MS12DEN5 = -1   ' ����l12 Den5
        If IsNull(rs("MS13LDL1")) = False Then .MS13LDL1 = rs("MS13LDL1") Else .MS13LDL1 = -1   ' ����l13 L/DL1
        If IsNull(rs("MS13LDL2")) = False Then .MS13LDL2 = rs("MS13LDL2") Else .MS13LDL2 = -1   ' ����l13 L/DL2
        If IsNull(rs("MS13LDL3")) = False Then .MS13LDL3 = rs("MS13LDL3") Else .MS13LDL3 = -1   ' ����l13 L/DL3
        If IsNull(rs("MS13LDL4")) = False Then .MS13LDL4 = rs("MS13LDL4") Else .MS13LDL4 = -1   ' ����l13 L/DL4
        If IsNull(rs("MS13LDL5")) = False Then .MS13LDL5 = rs("MS13LDL5") Else .MS13LDL5 = -1   ' ����l13 L/DL5
        If IsNull(rs("MS13DEN1")) = False Then .MS13DEN1 = rs("MS13DEN1") Else .MS13DEN1 = -1   ' ����l13 Den1
        If IsNull(rs("MS13DEN2")) = False Then .MS13DEN2 = rs("MS13DEN2") Else .MS13DEN2 = -1   ' ����l13 Den2
        If IsNull(rs("MS13DEN3")) = False Then .MS13DEN3 = rs("MS13DEN3") Else .MS13DEN3 = -1   ' ����l13 Den3
        If IsNull(rs("MS13DEN4")) = False Then .MS13DEN4 = rs("MS13DEN4") Else .MS13DEN4 = -1   ' ����l13 Den4
        If IsNull(rs("MS13DEN5")) = False Then .MS13DEN5 = rs("MS13DEN5") Else .MS13DEN5 = -1   ' ����l13 Den5
        If IsNull(rs("MS14LDL1")) = False Then .MS14LDL1 = rs("MS14LDL1") Else .MS14LDL1 = -1   ' ����l14 L/DL1
        If IsNull(rs("MS14LDL2")) = False Then .MS14LDL2 = rs("MS14LDL2") Else .MS14LDL2 = -1   ' ����l14 L/DL2
        If IsNull(rs("MS14LDL3")) = False Then .MS14LDL3 = rs("MS14LDL3") Else .MS14LDL3 = -1   ' ����l14 L/DL3
        If IsNull(rs("MS14LDL4")) = False Then .MS14LDL4 = rs("MS14LDL4") Else .MS14LDL4 = -1   ' ����l14 L/DL4
        If IsNull(rs("MS14LDL5")) = False Then .MS14LDL5 = rs("MS14LDL5") Else .MS14LDL5 = -1   ' ����l14 L/DL5
        If IsNull(rs("MS14DEN1")) = False Then .MS14DEN1 = rs("MS14DEN1") Else .MS14DEN1 = -1   ' ����l14 Den1
        If IsNull(rs("MS14DEN2")) = False Then .MS14DEN2 = rs("MS14DEN2") Else .MS14DEN2 = -1   ' ����l14 Den2
        If IsNull(rs("MS14DEN3")) = False Then .MS14DEN3 = rs("MS14DEN3") Else .MS14DEN3 = -1   ' ����l14 Den3
        If IsNull(rs("MS14DEN4")) = False Then .MS14DEN4 = rs("MS14DEN4") Else .MS14DEN4 = -1   ' ����l14 Den4
        If IsNull(rs("MS14DEN5")) = False Then .MS14DEN5 = rs("MS14DEN5") Else .MS14DEN5 = -1   ' ����l14 Den5
        If IsNull(rs("MS15LDL1")) = False Then .MS15LDL1 = rs("MS15LDL1") Else .MS15LDL1 = -1   ' ����l15 L/DL1
        If IsNull(rs("MS15LDL2")) = False Then .MS15LDL2 = rs("MS15LDL2") Else .MS15LDL2 = -1   ' ����l15 L/DL2
        If IsNull(rs("MS15LDL3")) = False Then .MS15LDL3 = rs("MS15LDL3") Else .MS15LDL3 = -1   ' ����l15 L/DL3
        If IsNull(rs("MS15LDL4")) = False Then .MS15LDL4 = rs("MS15LDL4") Else .MS15LDL4 = -1   ' ����l15 L/DL4
        If IsNull(rs("MS15LDL5")) = False Then .MS15LDL5 = rs("MS15LDL5") Else .MS15LDL5 = -1   ' ����l15 L/DL5
        If IsNull(rs("MS15DEN1")) = False Then .MS15DEN1 = rs("MS15DEN1") Else .MS15DEN1 = -1   ' ����l15 Den1
        If IsNull(rs("MS15DEN2")) = False Then .MS15DEN2 = rs("MS15DEN2") Else .MS15DEN2 = -1   ' ����l15 Den2
        If IsNull(rs("MS15DEN3")) = False Then .MS15DEN3 = rs("MS15DEN3") Else .MS15DEN3 = -1   ' ����l15 Den3
        If IsNull(rs("MS15DEN4")) = False Then .MS15DEN4 = rs("MS15DEN4") Else .MS15DEN4 = -1   ' ����l15 Den4
        If IsNull(rs("MS15DEN5")) = False Then .MS15DEN5 = rs("MS15DEN5") Else .MS15DEN5 = -1   ' ����l15 Den5
        If IsNull(rs("MS01DVD2")) = False Then .MS01DVD2 = rs("MS01DVD2") Else .MS01DVD2 = -1   ' ����l01 DVD2
        If IsNull(rs("MS02DVD2")) = False Then .MS02DVD2 = rs("MS02DVD2") Else .MS02DVD2 = -1   ' ����l02 DVD2
        If IsNull(rs("MS03DVD2")) = False Then .MS03DVD2 = rs("MS03DVD2") Else .MS03DVD2 = -1   ' ����l03 DVD2
        If IsNull(rs("MS04DVD2")) = False Then .MS04DVD2 = rs("MS04DVD2") Else .MS04DVD2 = -1   ' ����l04 DVD2
        If IsNull(rs("MS05DVD2")) = False Then .MS05DVD2 = rs("MS05DVD2") Else .MS05DVD2 = -1   ' ����l05 DVD2
        
        If IsNull(rs("TSTAFFID")) = False Then .TSTAFFID = rs("TSTAFFID")                       ' �o�^�Ј�ID
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")                          ' �o�^���t
        If IsNull(rs("KSTAFFID")) = False Then .KSTAFFID = rs("KSTAFFID")                       ' �X�V�Ј�ID
        If IsNull(rs("UPDDATE")) = False Then .UPDDATE = rs("UPDDATE")                          ' �X�V���t
        If IsNull(rs("SENDFLAG")) = False Then .SENDFLAG = rs("SENDFLAG")                       ' ���M�t���O
        If IsNull(rs("SENDDATE")) = False Then .SENDDATE = rs("SENDDATE")                       ' ���M���t
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        If IsNull(rs("MSZEROMN")) = False Then .MSZEROMN = rs("MSZEROMN") Else .MSZEROMN = -1   ' L/DL0�A�����ŏ��l
        If IsNull(rs("MSZEROMX")) = False Then .MSZEROMX = rs("MSZEROMX") Else .MSZEROMX = -1   ' L/DL0�A�����ő�l
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    
    End With
    
    Set rs = Nothing

    funGetGDJisseki_J015 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetGDJisseki_J015 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function
'************************************************************************************
'*    �֐���        : funGetSDJisseki_J022
'*
'*    �����T�v      : 1.SIRD����(TBCMJ022)�̎擾����
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^            ,����
'*                   sCryNum       ,I  ,String        ,���͗p
'*                   sSmplID       ,I  ,String        ,���o���R�[�h
'*                   sHsFlg_XSDCW  ,I  ,String        ,�ۏ�FLG(0:WF���сA1:��������)
'*                   udtGDjisseki  ,O  ,typ_TBCMJ022  ,SIRD����(�\����)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Public Function funGetSDJisseki_J022(sCryNum As String, sSmplID As String, _
                                         udtSDjisseki As typ_TBCMJ022) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim intSmplID       As Integer
    Dim sHsFlg_J022     As String       ' �ۏ�FLG(1:WF����)

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

'    If sHsFlg_XSDCW = "0" Then
'        ' WF����
'        sHsFlg_J022 = "1"
'    Else
'        ' ��������
'        sHsFlg_J015 = "0"
'    End If

    '�����ԍ��A�����ID����TBCMJ022��SIRD���ђl����������B
    sSQL = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, HSFLG, SMPLNO, SMPLUMU, "
    sSQL = sSQL & "HINBAN, REVNUM, FACTORY, OPECOND, SXLID, KRPROCCD, PROCCODE, GOUKI, "
    sSQL = sSQL & "OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, "
    sSQL = sSQL & "SIRDCNT,"
    sSQL = sSQL & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sSQL = sSQL & "from TBCMJ022 "
    sSQL = sSQL & "where CRYNUM = '" & sCryNum & "' and "
    sSQL = sSQL & "      SMPLNO = '" & sSmplID & "' and "
    sSQL = sSQL & "      TRANCNT = 0 "

    ' SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    ' �Y���ް��Ȃ�
    If rs.EOF Then
        udtSDjisseki.SMPLNO = ""
        funGetSDJisseki_J022 = FUNCTION_RETURN_SUCCESS
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtSDjisseki
        If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")                             ' �����ԍ�
        If IsNull(rs("POSITION")) = False Then .POSITION = rs("POSITION")                       ' �ʒu
        If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")                             ' �T���v���敪
        If IsNull(rs("TRANCOND")) = False Then .TRANCOND = rs("TRANCOND")                       ' ��������
        If IsNull(rs("TRANCNT")) = False Then .TRANCNT = rs("TRANCNT")                          ' ������
        If IsNull(rs("HSFLG")) = False Then .HSFLG = rs("HSFLG")                                ' �ۏ؃t���O
        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                             ' �T���v���m��
        If IsNull(rs("SMPLUMU")) = False Then .SMPLUMU = rs("SMPLUMU")                          ' �T���v���L��
        If IsNull(rs("HINBAN")) = False Then .hinban = rs("HINBAN")                             ' �i��
        If IsNull(rs("REVNUM")) = False Then .REVNUM = rs("REVNUM")                             ' ���i�ԍ������ԍ�
        If IsNull(rs("FACTORY")) = False Then .factory = rs("FACTORY")                         ' �H��
        If IsNull(rs("OPECOND")) = False Then .opecond = rs("OPECOND")                          ' ���Ə���
        If IsNull(rs("SXLID")) = False Then .SXLID = rs("SXLID")                                ' SXLID
        If IsNull(rs("KRPROCCD")) = False Then .KRPROCCD = rs("KRPROCCD")                       ' �Ǘ��H���R�[�h
        If IsNull(rs("PROCCODE")) = False Then .PROCCODE = rs("PROCCODE")                       ' �H���R�[�h
        If IsNull(rs("GOUKI")) = False Then .GOUKI = rs("GOUKI")                                ' ���@
        If IsNull(rs("OSITEM")) = False Then .OSITEM = rs("OSITEM")                             ' �]������
        If IsNull(rs("MAISU")) = False Then .MAISU = rs("MAISU")                                ' �]������
        If IsNull(rs("SPEC")) = False Then .Spec = rs("SPEC")                                   ' �K�i�l
        If IsNull(rs("NETSU")) = False Then .NETSU = rs("NETSU")                                ' �M��������
        If IsNull(rs("ET")) = False Then .ET = rs("ET")                                         ' �G�b�`���O����
        If IsNull(rs("MES")) = False Then .MES = rs("MES")                                      ' �v�����@
        If IsNull(rs("DKAN")) = False Then .DKAN = rs("DKAN")                                   ' �c�j�A�j�[������
        If IsNull(rs("SIRDCNT")) = False Then .SIRDCNT = rs("SIRDCNT")                          ' �ʓ����iSIRD��)
    End With

    Set rs = Nothing

    funGetSDJisseki_J022 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetSDJisseki_J022 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'************************************************************************************
'*    �֐���        : s_cmmc001db_sSql
'*
'*    �����T�v      : 1.���グ�I�����ю擾�֐�
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^            ,����
'*                   sCryNum       ,I  ,String        ,���͗p
'*                   udtTbcmh004   ,O  ,typ_TBCMH004  ,���グ�I�����ю擾�p
'*
'*    �߂�l        : (Double)����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************
Public Function s_cmmc001db_Sql(ByVal sCryNum As String, _
                udtTbcmh004() As typ_TBCMH004) As Double
    Dim sSQL    As String
    Dim intRET  As Integer
    
    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function s_cmmc001db_sSql"

    sSQL = " where CRYNUM = '" & sCryNum & "' "

    If DBDRV_GetTBCMH004(udtTbcmh004, sSQL, "order by CRYNUM") = FUNCTION_RETURN_FAILURE Then
        s_cmmc001db_Sql = FUNCTION_RETURN_FAILURE
    Else
        s_cmmc001db_Sql = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    s_cmmc001db_Sql = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'****************************************************************************************
'*    �֐���        : DBDRV_GetTBCMH001
'*
'*    �����T�v      : 1.�e�[�u���uTBCMH001�v��������ɂ��������R�[�h�𒊏o����
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^           ,����
'*                   udtRecords()  ,O  ,typ_TBCMH001 ,���o���R�[�h
'*                   sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'*                   sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'****************************************************************************************
Public Function DBDRV_GetTBCMH001(udtRecords() As typ_TBCMH001, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    Dim sSQL        As String       ' SQL�S��
    Dim sSqlBase    As String       ' SQL��{��(WHERE�߂̑O�܂�)
    Dim rs          As OraDynaset   ' RecordSet
    Dim lngRecCnt   As Long         ' ���R�[�h��
    Dim i           As Long

    ' SQL��g�ݗ��Ă�
    sSqlBase = "Select UPINDNO, KRPROCCD, PROCCODE, MODEL, GOUKI, PGID, CPORGIND, HINBAN, NMNOREVNO, NFACTORY, NOPECOND, NUMNOTE1," & _
              " NUMNOTE2, SEED, SEKIERTB, DPNTCLS, DOPANT, AMRESIST, CRYDOPCL, CRYDOPVL, UPBTCHNM, ADDDOPCL, ADDDOPVL, ADDDOPPT," & _
              " BCNT1COD, BCNT1CMT, BCNT2COD, BCNT2CMT, MTCLS1, MTWGHT1, ESWGHT1, MTCLS2, MTWGHT2, ESWGHT2, MTCLS3, MTWGHT3," & _
              " ESWGHT3, MTCLS4, MTWGHT4, ESWGHT4, MTCLS5, MTWGHT5, ESWGHT5, MTCLS6, MTWGHT6, ESWGHT6, MTCLS7, MTWGHT7, ESWGHT7," & _
              " MTCLS8, MTWGHT8, ESWGHT8, MTCLS9, MTWGHT9, ESWGHT9, MTCLS10, MTWGHT10, ESWGHT10, MTCLS11, MTWGHT11, ESWGHT11," & _
              " MTCLS12, MTWGHT12, ESWGHT12, MTCLS13, MTWGHT13, ESWGHT13, MTCLS14, MTWGHT14, ESWGHT14, MTCLS15, MTWGHT15," & _
              " ESWGHT15, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sSqlBase = sSqlBase & "From TBCMH001"
    sSQL = sSqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sSQL = sSQL & " " & sqlWhere & " " & sqlOrder
    End If

    ' �f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim udtRecords(0)
        DBDRV_GetTBCMH001 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ' ���o���ʂ��i�[����
    lngRecCnt = rs.RecordCount
    ReDim udtRecords(lngRecCnt)
    For i = 1 To lngRecCnt
        With udtRecords(i)
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

    DBDRV_GetTBCMH001 = FUNCTION_RETURN_SUCCESS
End Function

'****************************************************************************************
'*    �֐���        : DBDRV_GetTBCMH004
'*
'*    �����T�v      : 1.�e�[�u���uTBCMH004�v��������ɂ��������R�[�h�𒊏o����
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^           ,����
'*                   udtRecords()  ,O  ,typ_TBCMH004 ,���o���R�[�h
'*                   sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'*                   sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'****************************************************************************************
Public Function DBDRV_GetTBCMH004(udtRecords() As typ_TBCMH004, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    Dim sSQL        As String       ' SQL�S��
    Dim sSqlBase    As String       ' SQL��{��(WHERE�߂̑O�܂�)
    Dim rs          As OraDynaset   ' RecordSet
    Dim lngRecCnt   As Long         ' ���R�[�h��
    Dim i           As Long

    ' SQL��g�ݗ��Ă�
    sSqlBase = "Select CRYNUM, KRPROCCD, PROCCODE, LENGTOP, LENGTKDO, LENGTAIL, LENGFREE, DM1, DM2, DM3, WGHTTOP, WGHTTKDO," & _
              " WGHTTAIL, WGHTFREE, WGTOPCUT, UPWEIGHT, CHARGE, SEED, STATCLS, JDGECODE, PWTIME, ADDDPPOS, ADDDPCLS, ADDDPVAL," & _
              " ADDDPNAM, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sSqlBase = sSqlBase & "From TBCMH004"
    sSQL = sSqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sSQL = sSQL & " " & sqlWhere & " " & sqlOrder
    End If

    ' �f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim udtRecords(0)
        DBDRV_GetTBCMH004 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ' ���o���ʂ��i�[����
    lngRecCnt = rs.RecordCount
    ReDim udtRecords(lngRecCnt)
    For i = 1 To lngRecCnt
        With udtRecords(i)
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

'*********************************************************************************************
'*    �֐���        : funGetSPVJisseki_J016
'*
'*    �����T�v      : 1.SPV����(TBCMJ016)�̎擾����
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^                               ,����
'*                   sCryNum       ,I  ,String                           ,�����ԍ�
'*                   sSmplID       ,I  ,String                           ,�����ID
'*                   tSPVjisseki   ,O  ,typ_TBCMJ016                     ,����SPV����(�\����)
'*                   Siyou         ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou ,WF�d�l�p
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGetSPVJisseki_J016(sCryNum As String, sSmplID As String, udtSPVJisseki As typ_TBCMJ016 _
                                    , udtSiyou As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    Dim sSokutei        As String       ''������@(����ʒu�Q�� + ����ʒu�Q�_ + ����ʒu�Q��)
    
    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetSPVJisseki_J016"
    
    With udtSPVJisseki
        .MAX_FE = -2
        .MIN_FE = -2
        .AVE_FE = -2
        .CENTER_FE = -2
        .MAX_DIFF = -2
        .MIN_DIFF = -2
        .AVE_DIFF = -2
        .CENTER_DIFF = -2
    End With
            
    ' �����ԍ��A�����ID����TBCMJ016�̌���SPV���ђl����������B
    sSQL = ""
    sSQL = sSQL & " select CRYNUM,POSITION,SMPKBN,TRANCOND,TRANCNT,HSFLG,SMPLNO,SMPLUMU" & vbLf
    sSQL = sSQL & "       ,HINBAN,REVNUM,FACTORY,OPECOND,SXLID,KRPROCCD,PROCCODE,GOUKI" & vbLf
    sSQL = sSQL & "       ,OSITEM,MAISU,SPEC,NETSU,ET,MES,DKAN" & vbLf
    sSQL = sSQL & "       ,SPV_Fe_MAX,SPV_Fe_AVE,SPV_Fe_MIN" & vbLf
    sSQL = sSQL & "       ,ms01_SPV_Fe,ms02_SPV_Fe,ms03_SPV_Fe,ms04_SPV_Fe,ms05_SPV_Fe" & vbLf
    sSQL = sSQL & "       ,ms06_SPV_Fe,ms07_SPV_Fe,ms08_SPV_Fe,ms09_SPV_Fe" & vbLf
    sSQL = sSQL & "       ,SPV_Diff_MAX,SPV_Diff_AVE,SPV_Diff_MIN" & vbLf
    sSQL = sSQL & "       ,ms01_SPV_Diff,ms02_SPV_Diff,ms03_SPV_Diff,ms04_SPV_Diff,ms05_SPV_Diff" & vbLf
    sSQL = sSQL & "       ,ms06_SPV_Diff,ms07_SPV_Diff,ms08_SPV_Diff,ms09_SPV_Diff" & vbLf
    sSQL = sSQL & "       ,TSTAFFID,REGDATE,KSTAFFID,UPDDATE,SENDFLAG,SENDDATE" & vbLf

    ' SPV���菈���ǉ�
    '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
    sSQL = sSQL & "       ,SPV_Fe_PUA,SPV_Fe_PUAP,SPV_Fe_STD,SPV_Diff_PUA,SPV_Diff_PUAP" & vbLf
    sSQL = sSQL & "       ,SPV_Nr_MAX,SPV_Nr_AVE,SPV_Nr_STD,SPV_Nr_PUA,SPV_Nr_PUAP" & vbLf
    sSQL = sSQL & " from   TBCMJ016 " & vbLf
    sSQL = sSQL & " where  CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & " and    SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & " and    HSFLG = '1'" & vbLf
    sSQL = sSQL & " and    TRANCNT = ( select   max(TRANCNT) from TBCMJ016 " & vbLf
    sSQL = sSQL & "                    where    CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "                    and      SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "                    and      HSFLG = '1')" & vbLf
    
    ' SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' �Y���ް��Ȃ�
    If rs.EOF Then
        
        udtSPVJisseki.SMPLNO = "0"
        funGetSPVJisseki_J016 = FUNCTION_RETURN_SUCCESS
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtSPVJisseki
        If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")                                                 ' �����ԍ�
        If IsNull(rs("POSITION")) = False Then .POSITION = rs("POSITION")                                           ' �ʒu
        If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")                                                 ' �T���v���敪
        If IsNull(rs("TRANCOND")) = False Then .TRANCOND = rs("TRANCOND")                                           ' ��������
        If IsNull(rs("TRANCNT")) = False Then .TRANCNT = rs("TRANCNT")                                              ' ������
        If IsNull(rs("HSFLG")) = False Then .HSFLG = rs("HSFLG")                                                    ' �ۏ؃t���O
        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                                                 ' �T���v���m��
        If IsNull(rs("SMPLUMU")) = False Then .SMPLUMU = rs("SMPLUMU")                                              ' �T���v���L��
        If IsNull(rs("HINBAN")) = False Then .hinban = rs("HINBAN")                                                 ' �i��
        If IsNull(rs("REVNUM")) = False Then .REVNUM = rs("REVNUM")                                                 ' ���i�ԍ������ԍ�
        If IsNull(rs("FACTORY")) = False Then .factory = rs("FACTORY")                                              ' �H��
        If IsNull(rs("OPECOND")) = False Then .opecond = rs("OPECOND")                                              ' ���Ə���
        If IsNull(rs("SXLID")) = False Then .SXLID = rs("SXLID")                                                    ' SXLID
        If IsNull(rs("KRPROCCD")) = False Then .KRPROCCD = rs("KRPROCCD")                                           ' �Ǘ��H���R�[�h
        If IsNull(rs("PROCCODE")) = False Then .PROCCODE = rs("PROCCODE")                                           ' �H���R�[�h
        If IsNull(rs("GOUKI")) = False Then .GOUKI = rs("GOUKI")                                                    ' ���@
        If IsNull(rs("OSITEM")) = False Then .OSITEM = rs("OSITEM")                                                 ' �]������
        If IsNull(rs("MAISU")) = False Then .MAISU = rs("MAISU")                                                    ' �]������
        If IsNull(rs("SPEC")) = False Then .Spec = rs("SPEC")                                                       ' �K�i�l
        If IsNull(rs("NETSU")) = False Then .NETSU = rs("NETSU")                                                    ' �M��������
        If IsNull(rs("ET")) = False Then .ET = rs("ET")                                                             ' �G�b�`���O����
        If IsNull(rs("MES")) = False Then .MES = rs("MES")                                                          ' �v�����@
        If IsNull(rs("DKAN")) = False Then .DKAN = rs("DKAN")                                                       ' �c�j�A�j�[������
    
        If IsNull(rs("SPV_Fe_MAX")) = False Then .SPV_Fe_MAX = rs("SPV_Fe_MAX") Else .SPV_Fe_MAX = -1               ' SPV_Fe_MAX
        If IsNull(rs("SPV_Fe_AVE")) = False Then .SPV_Fe_AVE = rs("SPV_Fe_AVE") Else .SPV_Fe_AVE = -1               ' SPV_Fe_AVE
        If IsNull(rs("SPV_Fe_MIN")) = False Then .SPV_Fe_MIN = rs("SPV_Fe_MIN") Else .SPV_Fe_MIN = -1               ' SPV_Fe_MIN
        If IsNull(rs("ms01_SPV_Fe")) = False Then .ms01_SPV_Fe = rs("ms01_SPV_Fe") Else .ms01_SPV_Fe = -1           ' ����l01 SPV_Fe
        If IsNull(rs("ms02_SPV_Fe")) = False Then .ms02_SPV_Fe = rs("ms02_SPV_Fe") Else .ms02_SPV_Fe = -1           ' ����l02 SPV_Fe
        If IsNull(rs("ms03_SPV_Fe")) = False Then .ms03_SPV_Fe = rs("ms03_SPV_Fe") Else .ms03_SPV_Fe = -1           ' ����l03 SPV_Fe
        If IsNull(rs("ms04_SPV_Fe")) = False Then .ms04_SPV_Fe = rs("ms04_SPV_Fe") Else .ms04_SPV_Fe = -1           ' ����l04 SPV_Fe
        If IsNull(rs("ms05_SPV_Fe")) = False Then .ms05_SPV_Fe = rs("ms05_SPV_Fe") Else .ms05_SPV_Fe = -1           ' ����l05 SPV_Fe
        If IsNull(rs("ms06_SPV_Fe")) = False Then .ms06_SPV_Fe = rs("ms06_SPV_Fe") Else .ms06_SPV_Fe = -1           ' ����l06 SPV_Fe
        If IsNull(rs("ms07_SPV_Fe")) = False Then .ms07_SPV_Fe = rs("ms07_SPV_Fe") Else .ms07_SPV_Fe = -1           ' ����l07 SPV_Fe
        If IsNull(rs("ms08_SPV_Fe")) = False Then .ms08_SPV_Fe = rs("ms08_SPV_Fe") Else .ms08_SPV_Fe = -1           ' ����l08 SPV_Fe
        If IsNull(rs("ms09_SPV_Fe")) = False Then .ms09_SPV_Fe = rs("ms09_SPV_Fe") Else .ms09_SPV_Fe = -1           ' ����l09 SPV_Fe
        If IsNull(rs("SPV_Diff_MAX")) = False Then .SPV_Diff_MAX = rs("SPV_Diff_MAX") Else .SPV_Diff_MAX = -1       ' SPV_�g�U��_MAX
        If IsNull(rs("SPV_Diff_AVE")) = False Then .SPV_Diff_AVE = rs("SPV_Diff_AVE") Else .SPV_Diff_AVE = -1       ' SPV_�g�U��_AVE
        If IsNull(rs("SPV_Diff_MIN")) = False Then .SPV_Diff_MIN = rs("SPV_Diff_MIN") Else .SPV_Diff_MIN = -1       ' SPV_�g�U��_MIN
        If IsNull(rs("ms01_SPV_Diff")) = False Then .ms01_SPV_Diff = rs("ms01_SPV_Diff") Else .ms01_SPV_Diff = -1   ' ����l01 SPV_�g�U��
        If IsNull(rs("ms02_SPV_Diff")) = False Then .ms02_SPV_Diff = rs("ms02_SPV_Diff") Else .ms02_SPV_Diff = -1   ' ����l02 SPV_�g�U��
        If IsNull(rs("ms03_SPV_Diff")) = False Then .ms03_SPV_Diff = rs("ms03_SPV_Diff") Else .ms03_SPV_Diff = -1   ' ����l03 SPV_�g�U��
        If IsNull(rs("ms04_SPV_Diff")) = False Then .ms04_SPV_Diff = rs("ms04_SPV_Diff") Else .ms04_SPV_Diff = -1   ' ����l04 SPV_�g�U��
        If IsNull(rs("ms05_SPV_Diff")) = False Then .ms05_SPV_Diff = rs("ms05_SPV_Diff") Else .ms05_SPV_Diff = -1   ' ����l05 SPV_�g�U��
        If IsNull(rs("ms06_SPV_Diff")) = False Then .ms06_SPV_Diff = rs("ms06_SPV_Diff") Else .ms06_SPV_Diff = -1   ' ����l06 SPV_�g�U��
        If IsNull(rs("ms07_SPV_Diff")) = False Then .ms07_SPV_Diff = rs("ms07_SPV_Diff") Else .ms07_SPV_Diff = -1   ' ����l07 SPV_�g�U��
        If IsNull(rs("ms08_SPV_Diff")) = False Then .ms08_SPV_Diff = rs("ms08_SPV_Diff") Else .ms08_SPV_Diff = -1   ' ����l08 SPV_�g�U��
        If IsNull(rs("ms09_SPV_Diff")) = False Then .ms09_SPV_Diff = rs("ms09_SPV_Diff") Else .ms09_SPV_Diff = -1   ' ����l09 SPV_�g�U��
        
        If IsNull(rs("TSTAFFID")) = False Then .TSTAFFID = rs("TSTAFFID")                                           ' �o�^�Ј�ID
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")                                              ' �o�^���t
        If IsNull(rs("KSTAFFID")) = False Then .KSTAFFID = rs("KSTAFFID")                                           ' �X�V�Ј�ID
        If IsNull(rs("UPDDATE")) = False Then .UPDDATE = rs("UPDDATE")                                              ' �X�V���t
        If IsNull(rs("SENDFLAG")) = False Then .SENDFLAG = rs("SENDFLAG")                                           ' ���M�t���O
        If IsNull(rs("SENDDATE")) = False Then .SENDDATE = rs("SENDDATE")                                           ' ���M���t

        ' SPV���菈���ǉ�
        ' ����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
        If IsNull(rs("SPV_Fe_PUA")) = False Then .SPV_Fe_PUA = rs("SPV_Fe_PUA") Else .SPV_Fe_PUA = -1               ' SPV_Fe PUA�l
        If IsNull(rs("SPV_Fe_PUAP")) = False Then .SPV_Fe_PUAP = rs("SPV_Fe_PUAP") Else .SPV_Fe_PUAP = -1           ' SPV_Fe PUA���l
        If IsNull(rs("SPV_Fe_STD")) = False Then .SPV_Fe_STD = rs("SPV_Fe_STD") Else .SPV_Fe_STD = -1               ' SPV_Fe STD
        If IsNull(rs("SPV_Diff_PUA")) = False Then .SPV_Diff_PUA = rs("SPV_Diff_PUA") Else .SPV_Diff_PUA = -1       ' SPV_�g�U�� PUA�l
        If IsNull(rs("SPV_Diff_PUAP")) = False Then .SPV_Diff_PUAP = rs("SPV_Diff_PUAP") Else .SPV_Diff_PUAP = -1   ' SPV_�g�U�� PUA���l
        If IsNull(rs("SPV_Nr_MAX")) = False Then .SPV_Nr_MAX = rs("SPV_Nr_MAX") Else .SPV_Nr_MAX = -1               ' SPV_OtherRecords_MAX
        If IsNull(rs("SPV_Nr_AVE")) = False Then .SPV_Nr_AVE = rs("SPV_Nr_AVE") Else .SPV_Nr_AVE = -1               ' SPV_OtherRecords_AVE
        If IsNull(rs("SPV_Nr_STD")) = False Then .SPV_Nr_STD = rs("SPV_Nr_STD") Else .SPV_Nr_STD = -1               ' SPV_OtherRecords_STD
        If IsNull(rs("SPV_Nr_PUA")) = False Then .SPV_Nr_PUA = rs("SPV_Nr_PUA") Else .SPV_Nr_PUA = -1               ' SPV_OtherRecords_PUA�l
        If IsNull(rs("SPV_Nr_PUAP")) = False Then .SPV_Nr_PUAP = rs("SPV_Nr_PUAP") Else .SPV_Nr_PUAP = -1           ' SPV_OtherRecords_PUA���l

        ' Fe�Z�x������@
        sSokutei = Trim(udtSiyou.HWFSPVSH) & Trim(udtSiyou.HWFSPVST) & Trim(udtSiyou.HWFSPVSI)
    
        ' MAP����̏ꍇ
        If sSokutei = "AMX" Then
            .MAX_FE = .SPV_Fe_MAX
            ' SPV���菈���ǉ�
            ' Map����(AMX)�̏ꍇ�́A�\���f�[�^2(MIN)��\�����Ȃ��悤�ɏC��
            .MIN_FE = -1
            .AVE_FE = .SPV_Fe_AVE
            .CENTER_FE = -1

            ' SPV���菈���ǉ�
            .PUA_FE = .SPV_Fe_PUA
            .PUAP_FE = .SPV_Fe_PUAP
            .STD_FE = .SPV_Fe_STD
        ' 9�_����̏ꍇ
        ElseIf sSokutei = "V9T" Then
            ' Fe�Z�x��MAX,MIN,AVE���擾
            If funGetSPVJisseki_J016_Fe(.CRYNUM, .SMPLNO, .TRANCNT, _
                                    udtSPVJisseki) = FUNCTION_RETURN_FAILURE Then
                funGetSPVJisseki_J016 = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            
            .CENTER_FE = .ms01_SPV_Fe
            
            ' SPV���菈���ǉ�
            .PUA_FE = -1
            .PUAP_FE = -1
            .STD_FE = -1
        Else
            .MAX_FE = -1
            .MIN_FE = -1
            .AVE_FE = -1
            .CENTER_FE = -1
            
            'SPV���菈���ǉ�
            .PUA_FE = -1
            .PUAP_FE = -1
            .STD_FE = -1
        End If
    
        ' �g�U��������@
        sSokutei = Trim(udtSiyou.HWFDLSPH) & Trim(udtSiyou.HWFDLSPT) & Trim(udtSiyou.HWFDLSPI)
    
        ' MAP����̏ꍇ
        If sSokutei = "AMX" Then
            .MAX_DIFF = .SPV_Diff_MAX
            .MIN_DIFF = .SPV_Diff_MIN
            .AVE_DIFF = .SPV_Diff_AVE
            .CENTER_DIFF = -1
            
            ' SPV���菈���ǉ�
            .PUA_DIFF = .SPV_Diff_PUA
            .PUAP_DIFF = .SPV_Diff_PUAP
        
        ' 9�_����̏ꍇ
        ElseIf sSokutei = "V9T" Then
            ' �g�U����MAX,MIN,AVE���擾
            If funGetSPVJisseki_J016_Diff(.CRYNUM, .SMPLNO, .TRANCNT, _
                                    udtSPVJisseki) = FUNCTION_RETURN_FAILURE Then
                funGetSPVJisseki_J016 = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            
            .CENTER_DIFF = .ms01_SPV_Diff
            
            ' SPV���菈���ǉ�
            .PUA_DIFF = -1
            .PUAP_DIFF = -1
        Else
            .MAX_DIFF = -1
            .MIN_DIFF = -1
            .AVE_DIFF = -1
            .CENTER_DIFF = -1
            
            ' SPV���菈���ǉ�
            .PUA_DIFF = -1
            .PUAP_DIFF = -1
        End If
        
        ' SPV���菈���ǉ�
        ' Nr�Z�x������@
        sSokutei = Trim(udtSiyou.HWFNRSH) & Trim(udtSiyou.HWFNRST) & Trim(udtSiyou.HWFNRSI)
        
        ' MAP����̏ꍇ
        If sSokutei = "AMX" Then
            .MAX_NR = .SPV_Nr_MAX
            .MIN_NR = -1
            .AVE_NR = .SPV_Nr_AVE
            .CENTER_NR = -1
            .PUA_NR = .SPV_Nr_PUA
            .PUAP_NR = .SPV_Nr_PUAP
            .STD_NR = .SPV_Nr_STD
        
        ' 9�_����̏ꍇ
        ElseIf sSokutei = "V9T" Then
            .MAX_NR = -1
            .MIN_NR = -1
            .AVE_NR = -1
            .CENTER_NR = -1
            .PUA_NR = -1
            .PUAP_NR = -1
            .STD_NR = -1
        Else
            .MAX_NR = -1
            .MIN_NR = -1
            .AVE_NR = -1
            .CENTER_NR = -1
            .PUA_NR = -1
            .PUAP_NR = -1
            .STD_NR = -1
        End If
    End With
    
    Set rs = Nothing

    funGetSPVJisseki_J016 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetSPVJisseki_J016 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    �֐���        : funGetSPVJisseki_J016_Fe
'*
'*    �����T�v      : 1.SPV����(TBCMJ016)��Fe�Z�x9�_����l��MAX�EMIN�EAVE���擾����
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^                               ,����
'*                   sCryNum       ,I  ,String                           ,�����ԍ�
'*                   sSmplID       ,I  ,String                           ,�����ID
'*                   intTrancnt    ,I  ,Integer                          ,������
'*                   udtSPVJisseki ,O  ,typ_TBCMJ016                     ,����SPV����(�\����)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGetSPVJisseki_J016_Fe(sCryNum As String, sSmplID As String, intTrancnt As Integer, _
                                        udtSPVJisseki As typ_TBCMJ016) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    
    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetSPVJisseki_J016_Fe"
            
    sSQL = ""
    sSQL = sSQL & " SELECT  MAX(SPV_FE) AS MAX_FE,MIN(SPV_FE) AS MIN_FE,AVG(SPV_FE) AS AVE_FE" & vbLf
    sSQL = sSQL & " FROM   (SELECT  CRYNUM,SMPLNO,TRANCNT,ms01_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms02_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms03_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms04_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms05_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms06_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms07_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms08_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms09_SPV_FE AS SPV_FE" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "        )" & vbLf
    
    ' SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' �Y���ް��Ȃ�
    If rs.EOF Then
        funGetSPVJisseki_J016_Fe = FUNCTION_RETURN_SUCCESS
    
        With udtSPVJisseki
            .MAX_FE = -1
            .MIN_FE = -1
            .AVE_FE = -1
        End With
        
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtSPVJisseki
        If IsNull(rs("MAX_FE")) = False Then .MAX_FE = rs("MAX_FE") Else .MAX_FE = -1
        If IsNull(rs("MIN_FE")) = False Then .MIN_FE = rs("MIN_FE") Else .MIN_FE = -1
        If IsNull(rs("AVE_FE")) = False Then .AVE_FE = rs("AVE_FE") Else .AVE_FE = -1
    End With
    
    Set rs = Nothing
    
    funGetSPVJisseki_J016_Fe = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetSPVJisseki_J016_Fe = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    �֐���        : funGetSPVJisseki_J016_Diff
'*
'*    �����T�v      : 1.SPV����(TBCMJ016)�̊g�U��9�_����l��MAX�EMIN�EAVE���擾����
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^                               ,����
'*                   sCryNum       ,I  ,String                           ,�����ԍ�
'*                   sSmplID       ,I  ,String                           ,�����ID
'*                   intTrancnt    ,I  ,Integer                          ,������
'*                   udtSPVJisseki ,O  ,typ_TBCMJ016                     ,����SPV����(�\����)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGetSPVJisseki_J016_Diff(sCryNum As String, sSmplID As String, intTrancnt As Integer, _
                                        udtSPVJisseki As typ_TBCMJ016) As FUNCTION_RETURN
    Dim sSQL            As String
    Dim rs              As OraDynaset
    
    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetSPVJisseki_J016_Diff"
            
    sSQL = ""
    sSQL = sSQL & " SELECT  MAX(SPV_DIFF) AS MAX_DIFF,MIN(SPV_DIFF) AS MIN_DIFF,AVG(SPV_DIFF) AS AVE_DIFF" & vbLf
    sSQL = sSQL & " FROM   (SELECT  CRYNUM,SMPLNO,TRANCNT,ms01_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms02_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms03_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms04_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms05_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms06_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms07_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms08_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "             UNION ALL" & vbLf
    sSQL = sSQL & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms09_SPV_DIFF AS SPV_DIFF" & vbLf
    sSQL = sSQL & "         FROM    TBCMJ016" & vbLf
    sSQL = sSQL & "         WHERE   CRYNUM = '" & sCryNum & "'" & vbLf
    sSQL = sSQL & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
    sSQL = sSQL & "         AND     TRANCNT = " & intTrancnt & vbLf
    sSQL = sSQL & "         AND     HSFLG = '1'" & vbLf
    sSQL = sSQL & "        )" & vbLf
    
    ' SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    ' �Y���ް��Ȃ�
    If rs.EOF Then
        funGetSPVJisseki_J016_Diff = FUNCTION_RETURN_SUCCESS
    
        With udtSPVJisseki
            .MAX_DIFF = -1
            .MIN_DIFF = -1
            .AVE_DIFF = -1
        End With
        
        Set rs = Nothing
        GoTo proc_exit
    End If

    With udtSPVJisseki
        If IsNull(rs("MAX_DIFF")) = False Then .MAX_DIFF = rs("MAX_DIFF") Else .MAX_DIFF = -1
        If IsNull(rs("MIN_DIFF")) = False Then .MIN_DIFF = rs("MIN_DIFF") Else .MIN_DIFF = -1
        If IsNull(rs("AVE_DIFF")) = False Then .AVE_DIFF = rs("AVE_DIFF") Else .AVE_DIFF = -1
    End With
    
    Set rs = Nothing
    
    funGetSPVJisseki_J016_Diff = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Set rs = Nothing
    funGetSPVJisseki_J016_Diff = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    �֐���        : funGetSiyou_Warp
'*
'*    �����T�v      : 1.Warp�d�l�l�̎擾����
'*
'*    �p�����[�^    : �ϐ���    ,IO ,�^           ,����
'*                   udtHIN     ,I  ,tFullHinban  ,�i��
'*                   dblWarpMax ,I  ,Double       ,Warp���
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGetSiyou_Warp(udtHin As tFullHinban, dblWarpMax As Double) As FUNCTION_RETURN

    Dim sSQL    As String           ' SQL�S��
    Dim rs      As OraDynaset       ' RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetSiyou_Warp"

    sSQL = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sSQL = sSQL & "HWFWARMX "
    sSQL = sSQL & "from TBCME027 "
    sSQL = sSQL & "Where HINBAN = '" & udtHin.hinban & "' and "
    sSQL = sSQL & "      MNOREVNO = " & udtHin.mnorevno & " and "
    sSQL = sSQL & "      FACTORY = '" & udtHin.factory & "' and "
    sSQL = sSQL & "      OPECOND = '" & udtHin.opecond & "'"
    
    ' �f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGetSiyou_Warp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ' ���o���ʂ��i�[����
    dblWarpMax = fncNullCheck(rs("HWFWARMX"))         ' �iWFWARP���
        
    Set rs = Nothing

    funGetSiyou_Warp = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funGetSiyou_Warp = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    �֐���        : funGetSiyou_Kaku
'*
'*    �����T�v      : 1.�����p�x�d�l�l�̎擾����
'*
'*    �p�����[�^    : �ϐ���    ,IO ,�^           ,����
'*                   udtHIN     ,I  ,tFullHinban  ,�i��
'*                   dblKakuMin ,I  ,Double       ,�����ʌX����
'*                   dblWarpMax ,I  ,Double       ,�����ʌX���
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGetSiyou_Kaku(udtHin As tFullHinban, dblKakuMin As Double, dblKakuMax As Double) As FUNCTION_RETURN
    Dim sSQL    As String           ' SQL�S��
    Dim rs      As OraDynaset       ' RecordSet

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetSiyou_Kaku"

    sSQL = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sSQL = sSQL & "HWFCSMIN, HWFCSMAX "
    sSQL = sSQL & "from TBCME022 "
    sSQL = sSQL & "Where HINBAN = '" & udtHin.hinban & "' and "
    sSQL = sSQL & "      MNOREVNO = " & udtHin.mnorevno & " and "
    sSQL = sSQL & "      FACTORY = '" & udtHin.factory & "' and "
    sSQL = sSQL & "      OPECOND = '" & udtHin.opecond & "'"
    
    ' �f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGetSiyou_Kaku = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ' ���o���ʂ��i�[����
    dblKakuMin = fncNullCheck(rs("HWFCSMIN"))         ' �iWF�����ʌX����
    dblKakuMax = fncNullCheck(rs("HWFCSMAX"))         ' �iWF�����ʌX���
    
    Set rs = Nothing

    funGetSiyou_Kaku = FUNCTION_RETURN_SUCCESS
  
proc_exit:
    ' �I��
'    gErr.Pop
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funGetSiyou_Kaku = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    �֐���        : funGet_TBCMY018
'*
'*    �����T�v      : 1.�W�������ް�(TBCMY018)�̎擾����
'*
'*    �p�����[�^    : �ϐ���    ,IO ,�^                  ,����
'*                   sBlockID   ,I  ,String              ,��ۯ�ID
'*                   sMeasItem  ,I  ,String              ,���荀�ږ�
'*                   udtMEAS()  ,O  ,typ_WarpKakuData    ,�����ް�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funGet_TBCMY018(sBlockId As String, sMeasItem As String, udtMEAS() As typ_WarpKakuData) As FUNCTION_RETURN
    Dim sSQL        As String           ' SQL�S��
    Dim rs          As OraDynaset       ' RecordSet
    Dim lngRecCnt   As Long             ' ں��ސ�
    Dim i           As Integer

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funGet_TBCMY018"

    sSQL = "select SUBLOTID, MEASITEM, WAFID, MEASDATA "
    sSQL = sSQL & "from TBCMY018 Y018 "
    sSQL = sSQL & "Where SUBLOTID = '" & sBlockId & "' and "
    sSQL = sSQL & "      MEASITEM like '%" & sMeasItem & "%' and "
    sSQL = sSQL & "      TRANCNT = (select MAX(TRANCNT) from TBCMY018 "
    sSQL = sSQL & "                 where SUBLOTID = Y018.SUBLOTID "
    sSQL = sSQL & "                 and WAFID = Y018.WAFID) "
    sSQL = sSQL & "order by WAFID"
    
    ' �f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    If rs Is Nothing Then
        ReDim udtMEAS(0)
        rs.Close
        funGet_TBCMY018 = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    lngRecCnt = rs.RecordCount
    ReDim udtMEAS(lngRecCnt)
    
    ' ���o���ʂ��i�[����
    For i = 1 To lngRecCnt
        ' ��ۯ�ID
        udtMEAS(i).BLOCKID = sBlockId
        
        ' ��ʰID
        If IsNull(rs("WAFID")) Then
            udtMEAS(i).WAFID = -1
        ElseIf Not IsNumeric(rs("WAFID")) Then
            udtMEAS(i).WAFID = -1
        Else
            udtMEAS(i).WAFID = CDbl(rs("WAFID"))
        End If
        
        ' ����l
        If IsNull(rs("MEASDATA")) Then
            udtMEAS(i).MEASDATA = -1
        ElseIf Not IsNumeric(rs("MEASDATA")) Then
            udtMEAS(i).MEASDATA = -1
        Else
            udtMEAS(i).MEASDATA = CDbl(rs("MEASDATA"))
        End If
        
        rs.MoveNext
    Next i
    rs.Close

    funGet_TBCMY018 = FUNCTION_RETURN_SUCCESS
  
proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funGet_TBCMY018 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*********************************************************************************************
'*    �֐���        : funWfcGetDataEtc_SPV
'*
'*    �����T�v      : 1.WF�������� �e��f�[�^�擾(SPV�p)
'*
'*    �p�����[�^    : �ϐ���      ,IO  ,�^                                    ,����
'*                   udtNew_Hinban,I   ,tFullHinban                           ,�i�ԏ��
'*                   Siyou        ,O   ,type_DBDRV_scmzc_fcmlc001c_Siyou_SPV  ,WF�d�l�p
'*                   sErrMsg �@�@ ,O   ,String    �@�@�@�@�@�@�@�@�@�@�@    �@,�G���[���b�Z�[�W
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*********************************************************************************************
Public Function funWfcGetDataEtc_SPV(udtNew_Hinban As tFullHinban, _
                                 udtSiyou As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                                 Optional sErrMsg As String = vbNullString) As FUNCTION_RETURN
    Dim sSQL    As String
    Dim rs      As OraDynaset
    Dim sDBName As String

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "SB_WfJudg_SQL.bas -- Function funWfcGetDataEtc_SPV"

    funWfcGetDataEtc_SPV = FUNCTION_RETURN_SUCCESS

    ' WF�d�l�擾

    sDBName = "E048"
    sSQL = "select "
    sSQL = sSQL & "HWFSPVPUG,"        ' �i�v�e�r�o�u�e�d�o�t�`��
    sSQL = sSQL & "HWFSPVPUR,"        ' �i�v�e�r�o�u�e�d�o�t�`��
    sSQL = sSQL & "HWFSPVSTD,"        ' �i�v�e�r�o�u�e�d�W���΍�
    sSQL = sSQL & "HWFDLPUG,"         ' �i�v�e�g�U���o�t�`��
    sSQL = sSQL & "HWFDLPUR,"         ' �i�v�e�g�U���o�t�`��
    sSQL = sSQL & "HWFNRMX,"          ' �i�v�e�r�o�u�m�q���
    sSQL = sSQL & "HWFNRAM,"          ' �i�v�e�r�o�u�m�q����
    sSQL = sSQL & "HWFNRPUG,"         ' �i�v�e�r�o�u�m�q�o�t�`��
    sSQL = sSQL & "HWFNRPUR,"         ' �i�v�e�r�o�u�m�q�o�t�`��
    sSQL = sSQL & "HWFNRSTD,"         ' �i�v�e�r�o�u�m�q�W���΍�
    sSQL = sSQL & "HWFNRKN,"          ' �i�v�e�r�o�u�m�q�����p�x�Q��
    sSQL = sSQL & "HWFNRHS,"          ' �i�v�e�r�o�u�m�q�ۏؕ��@�Q��
    sSQL = sSQL & "HWFNRSH,"          ' �i�v�e�r�o�u�m�q����ʒu�Q��
    sSQL = sSQL & "HWFNRST,"          ' �i�v�e�r�o�u�m�q����ʒu�Q�_
    sSQL = sSQL & "HWFNRHT,"          ' �i�v�e�r�o�u�m�q�ۏؕ��@�Q��
    sSQL = sSQL & "HWFNRSI "          ' �i�v�e�r�o�u�m�q����ʒu�Q��
    sSQL = sSQL & "from TBCME048 "
    sSQL = sSQL & "where HINBAN = '" & udtNew_Hinban.hinban & "' "
    sSQL = sSQL & "and MNOREVNO = " & udtNew_Hinban.mnorevno & " "
    sSQL = sSQL & "and FACTORY = '" & udtNew_Hinban.factory & "' "
    sSQL = sSQL & "and OPECOND = '" & udtNew_Hinban.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funWfcGetDataEtc_SPV = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With udtSiyou
        ' �i�v�e�r�o�u�e�d�o�t�`��
        If IsNull(rs("HWFSPVPUG")) = False Then .HWFSPVPUG = rs("HWFSPVPUG") Else .HWFSPVPUG = -1
        
        ' �i�v�e�r�o�u�e�d�o�t�`��
        If IsNull(rs("HWFSPVPUR")) = False Then .HWFSPVPUR = rs("HWFSPVPUR") Else .HWFSPVPUR = -1
        
        ' �i�v�e�r�o�u�e�d�W���΍�
        If IsNull(rs("HWFSPVSTD")) = False Then .HWFSPVSTD = rs("HWFSPVSTD") Else .HWFSPVSTD = -1
        
        ' �i�v�e�r�o�u�m�q���
        .HWFNRMX = fncNullCheck(rs("HWFNRMX"))
        
        ' �i�v�e�r�o�u�m�q�o�t�`��
        If IsNull(rs("HWFNRPUG")) = False Then .HWFNRPUG = rs("HWFNRPUG") Else .HWFNRPUG = -1
        
        ' �i�v�e�r�o�u�m�q�o�t�`��
        If IsNull(rs("HWFNRPUR")) = False Then .HWFNRPUR = rs("HWFNRPUR") Else .HWFNRPUR = -1
        
        ' �i�v�e�r�o�u�m�q�W���΍�
        If IsNull(rs("HWFNRSTD")) = False Then .HWFNRSTD = rs("HWFNRSTD") Else .HWFNRSTD = -1
        
        ' �i�v�e�g�U���o�t�`��
        If IsNull(rs("HWFDLPUG")) = False Then .HWFDLPUG = rs("HWFDLPUG") Else .HWFDLPUG = -1
        
        ' �i�v�e�g�U���o�t�`��
        If IsNull(rs("HWFDLPUR")) = False Then .HWFDLPUR = rs("HWFDLPUR") Else .HWFDLPUR = -1
        
        ' �i�v�e�r�o�u�m�q����
        .HWFNRAM = fncNullCheck(rs("HWFNRAM"))
        
        ' �i�v�e�r�o�u�m�q�����p�x�Q��
        If IsNull(rs("HWFNRKN")) = False Then .HWFNRKN = rs("HWFNRKN") Else .HWFNRKN = vbNullString
        
        ' �i�v�e�r�o�u�m�q�ۏؕ��@�Q��
        If IsNull(rs("HWFNRHS")) = False Then .HWFNRHS = rs("HWFNRHS") Else .HWFNRHS = vbNullString
        
        ' �i�v�e�r�o�u�m�q����ʒu_��
        If IsNull(rs("HWFNRSH")) = False Then .HWFNRSH = rs("HWFNRSH") Else .HWFNRSH = vbNullString
        
        ' �i�v�e�r�o�u�m�q����ʒu_�_
        If IsNull(rs("HWFNRST")) = False Then .HWFNRST = rs("HWFNRST") Else .HWFNRST = vbNullString
        
        ' �i�v�e�r�o�u�m�q����ʒu_��
        If IsNull(rs("HWFNRHT")) = False Then .HWFNRHT = rs("HWFNRHT") Else .HWFNRHT = vbNullString
        
        ' �i�v�e�r�o�u�m�q����ʒu_��
        If IsNull(rs("HWFNRSI")) = False Then .HWFNRSI = rs("HWFNRSI") Else .HWFNRSI = vbNullString
    End With
    rs.Close

    Set rs = Nothing

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    funWfcGetDataEtc_SPV = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function
