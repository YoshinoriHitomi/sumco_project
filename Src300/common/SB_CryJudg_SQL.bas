Attribute VB_Name = "SB_CryJudg_SQL"
Option Explicit

'�i�ԁA�d�l�A���������擾�p(TOP,TAIL���łQ���R�[�h�擾)
Public Type type_DBDRV_scmzc_fcmkc001c_Siyou
    
    '�u���b�N�Ǘ�
    CRYNUM      As String * 12        ' �����ԍ�
    INGOTPOS    As Integer            ' �������J�n�ʒu
    Length      As Integer            ' ����
    
    '�i�ԊǗ�
    HIN As tFullHinban                ' �i��(full)
        
    '�������
    PRODCOND    As String * 4         ' �������
    PGID        As String * 8         ' �o�f�|�h�c
    UPLENGTH    As Integer            ' ���グ����
    FREELENG    As Integer            ' �t���[��
    DIAMETER    As Integer            ' ���a 2002/05/01 S.Sano
    CHARGE      As Double             ' �`���[�W��
    SEED        As String * 4         ' �V�[�h
    ADDDPPOS    As Integer            ' �ǉ��h�[�v�ʒu

    '���i�d�l
    HSXTYPE  As String * 1            ' �i�r�w�^�C�v
    HSXD1CEN As Double                ' �i�r�w���a�P���S
    HSXCDIR  As String * 1            ' �i�r�w�����ʕ���
    HSXRMIN  As Double                ' �i�r�w���R����
    HSXRMAX  As Double                ' �i�r�w���R���
    HSXRAMIN As Double                ' �i�r�w���R���ω���
    HSXRAMAX As Double                ' �i�r�w���R���Ϗ��
    HSXRMCAL As String * 1            ' �i�r�w���R�ʓ��v�Z�@�@�@�@'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
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
    HSXONMCL As String * 1            ' �i�r�w�_�f�Z�x�ʓ��v�Z�@�@�@�@'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
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

    HSXOF1AX As Double                ' �i�r�w�n�r�e�P���Ϗ��
    HSXOF1MX As Double                ' �i�r�w�n�r�e�P���
    HSXOF2AX As Double                ' �i�r�w�n�r�e�Q���Ϗ��
    HSXOF2MX As Double                ' �i�r�w�n�r�e�Q���
    HSXOF3AX As Double                ' �i�r�w�n�r�e�R���Ϗ��
    HSXOF3MX As Double                ' �i�r�w�n�r�e�R���
    HSXOF4AX As Double                ' �i�r�w�n�r�e�S���Ϗ��
    HSXOF4MX As Double                ' �i�r�w�n�r�e�S���
    HSXOF1SH As String * 1            ' �i�r�w�n�r�e�P����ʒu�Q��
    HSXOF1ST As String * 1            ' �i�r�w�n�r�e�P����ʒu�Q�_
    HSXOF1SR As String * 1            ' �i�r�w�n�r�e�P����ʒu�Q��
    HSXOF1HT As String * 1            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    HSXOF1HS As String * 1            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    HSXOF2SH As String * 1            ' �i�r�w�n�r�e�Q����ʒu�Q��
    HSXOF2ST As String * 1            ' �i�r�w�n�r�e�Q����ʒu�Q�_
    HSXOF2SR As String * 1            ' �i�r�w�n�r�e�Q����ʒu�Q��
    HSXOF2HT As String * 1            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    HSXOF2HS As String * 1            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    HSXOF3SH As String * 1            ' �i�r�w�n�r�e�R����ʒu�Q��
    HSXOF3ST As String * 1            ' �i�r�w�n�r�e�R����ʒu�Q�_
    HSXOF3SR As String * 1            ' �i�r�w�n�r�e�R����ʒu�Q��
    HSXOF3HT As String * 1            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    HSXOF3HS As String * 1            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    HSXOF4SH As String * 1            ' �i�r�w�n�r�e�S����ʒu�Q��
    HSXOF4ST As String * 1            ' �i�r�w�n�r�e�S����ʒu�Q�_
    HSXOF4SR As String * 1            ' �i�r�w�n�r�e�S����ʒu�Q��
    HSXOF4HT As String * 1            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    HSXOF4HS As String * 1            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    HSXOF1NS As String * 2            ' �i�r�w�n�r�e�P�M�����@
    HSXOF2NS As String * 2            ' �i�r�w�n�r�e�Q�M�����@
    HSXOF3NS As String * 2            ' �i�r�w�n�r�e�R�M�����@
    HSXOF4NS As String * 2            ' �i�r�w�n�r�e�S�M�����@
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
    HSXCNKHI As String * 1            ' �i�r�w�Y�f�Z�x�����p�x�Q�� 09/01/08 ooba

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
''Add Start 2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
    HSXLT10MIN As Integer             ' �i�r�w�k�^�C��10�����Z�����l
''Add End   2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
    HSXLTSPH As String * 1            ' �i�r�w�k�^�C������ʒu�Q��
    HSXLTSPT As String * 1            ' �i�r�w�k�^�C������ʒu�Q�_
    HSXLTSPI As String * 1            ' �i�r�w�k�^�C������ʒu�Q��
    HSXLTHWT As String * 1            ' �i�r�w�k�^�C���ۏؕ��@�Q��
    HSXLTHWS As String * 1            ' �i�r�w�k�^�C���ۏؕ��@�Q��
    '���������Ǘ�
    EPDUP As Integer                  ' EPD�@���
    
    'WF�d�l(��������p)�@08/4/15 ooba START ==========================>
    HWFRHWYS As String * 1          ' �i�v�e���R�ۏؕ��@�Q��
    HWFONHWS As String * 1          ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
    HWFOF1HS As String * 1          ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    HWFOF2HS As String * 1          ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    HWFOF3HS As String * 1          ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    HWFOF4HS As String * 1          ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    HWFBM1HS As String * 1          ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    HWFBM2HS As String * 1          ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    HWFBM3HS As String * 1          ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    HWFDENHS As String * 1          ' �i�v�e�c�����ۏؕ��@�Q��
    HWFDVDHS As String * 1          ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��
    HWFLDLHS As String * 1          ' �i�v�e�k�^�c�k�ۏؕ��@�Q��
    HWFRKHNN As String * 1          ' �i�v�e���R�����p�x�Q��
    HWFONKHN As String * 1          ' �i�v�e�_�f�Z�x�����p�x�Q��
    HWFOF1KN As String * 1          ' �i�v�e�n�r�e�P�����p�x�Q��
    HWFOF2KN As String * 1          ' �i�v�e�n�r�e�Q�����p�x�Q��
    HWFOF3KN As String * 1          ' �i�v�e�n�r�e�R�����p�x�Q��
    HWFOF4KN As String * 1          ' �i�v�e�n�r�e�S�����p�x�Q��
    HWFBM1KN As String * 1          ' �i�v�e�a�l�c�P�����p�x�Q��
    HWFBM2KN As String * 1          ' �i�v�e�a�l�c�Q�����p�x�Q��
    HWFBM3KN As String * 1          ' �i�v�e�a�l�c�R�����p�x�Q��
    HWFGDKHN As String * 1          ' �i�v�e�f�c�����p�x�Q��
    'WF�d�l(��������p)�@08/4/15 ooba END ============================>
    
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 start
    TOPREG  As Integer                ' TOP�K��
    TAILREG As Double                 ' TAIL�K��
    BTMSPRT As Integer                ' �{�g���͏o�K��
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 end

' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
    HSXOSF1PTK As String * 1          ' �i�r�w�n�r�e�P�p�^���敪
    HSXOSF2PTK As String * 1          ' �i�r�w�n�r�e�Q�p�^���敪
    HSXOSF3PTK As String * 1          ' �i�r�w�n�r�e�R�p�^���敪
    HSXOSF4PTK As String * 1          ' �i�r�w�n�r�e�S�p�^���敪
    HSXBMD1MBP As Double              ' �i�r�w�a�l�c�P�ʓ����z
    HSXBMD2MBP As Double              ' �i�r�w�a�l�c�Q�ʓ����z
    HSXBMD3MBP As Double              ' �i�r�w�a�l�c�R�ʓ����z
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
    BLOCKHFLAG As String * 1
''Upd Start (TCS)T.Terauchi 2005/10/12  GDײݐ��\���Ή�
    HSXGDLINE   As String * 3         ' GDײݐ�
''Upd End   (TCS)T.Terauchi 2005/10/12  GDײݐ��\���Ή�

'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
    COSF3FLAG As String * 1
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP    As String * 1              ' DK���x�i�d�l�j
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
    HSXGDPTK As String * 1          ' �i�r�w�f�c�p�^���敪
    HWFGDPTK    As String * 1       ' �i�v�e�f�c�p�^���敪
    WFHSGDCW    As String * 1       ' �ۏ�FLG�iGD)
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
''2009/07/13 add Kameda ���f -----------------------
    HSXCDOPMN As Double
    HSXCDOPMX As Double
    HSXCDPNI As String
    HSXCDOPN As Double
''---------------------------------------------------
''2009/08/12 add Kameda �����ʌX
    HSXCSCEN As Double
    HSXCSMIN As Double
    HSXCSMAX As Double
''2009/09/01 add Kameda �����ʌX
    HSXCYCEN As Double
    HSXCYMIN As Double
    HSXCYMAX As Double
    HSXCTCEN As Double
    HSXCTMIN As Double
    HSXCTMAX As Double
''2010/02/04 add Kameda SIRD
    HWFSIRDMX As Double
    
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎d�l���ڒǉ�
    HSXCPK      As String * 1       ' �i�r�w�b�p�^�[���敪
    HSXCSZ      As String * 1       ' �i�r�w�b�������
    HSXCHT      As String * 1       ' �i�r�w�b�ۏؕ��@�Q��
    HSXCHS      As String * 1       ' �i�r�w�b�ۏؕ��@�Q��
    HSXCJPK     As String * 1       ' �i�r�w�b�i�p�^�[���敪
    HSXCJNS     As String * 2       ' �i�r�w�b�i�M�����@
    HSXCJHT     As String * 1       ' �i�r�w�b�i�ۏؕ��@�Q��
    HSXCJHS     As String * 1       ' �i�r�w�b�i�ۏؕ��@�Q��
    HSXCJLTPK   As String * 1       ' �i�r�w�b�i�k�s�p�^�[���敪
    HSXCJLTNS   As String * 2       ' �i�r�w�b�i�k�s�M�����@
    HSXCJLTHT   As String * 1       ' �i�r�w�b�i�k�s�ۏؕ��@�Q��
    HSXCJLTHS   As String * 1       ' �i�r�w�b�i�k�s�ۏؕ��@�Q��
    HSXCJ2PK    As String * 1       ' �i�r�w�b�i�Q�p�^�[���敪
    HSXCJ2NS    As String * 2       ' �i�r�w�b�i�Q�M�����@
    HSXCJ2HT    As String * 1       ' �i�r�w�b�i�Q�ۏؕ��@�Q��
    HSXCJ2HS    As String * 1       ' �i�r�w�b�i�Q�ۏؕ��@�Q��
    HSXCJLTBND  As Integer          ' �iSXL/CJLT�o���h�� Number(3,0)
  'Add End   2011/01/17 SMPK A.Nagamine

'Add Start 2011/02/28 SMPK H.Ohkubo
    HSXONKHI As String * 1          ' �i�r�w�_�f�Z�x�����p�x�Q��
    FRSFLG   As String * 1          ' FRS����L��
'Add End 2011/02/28 SMPK H.Ohkubo
'Add Start 2012/06/01 SMPK H.Ohkubo
    HSXCOSF3PK   As String * 1      ' �i�r�w�b�n�r�e�R�p�^�[���敪
'Add Start 2012/06/01 SMPK H.Ohkubo
End Type

' �V�T���v���Ǘ�(��ۯ�)�擾�p(TOP,TAIL���łQ���R�[�h�擾)
Public Type type_DBDRV_scmzc_fcmkc001c_CrySmp
    CRYNUMCS        As String * 12      '�u���b�NID
    Length          As Integer          ' ����
    SMPKBNCS        As String * 1       '�T���v���敪
    TBKBNCS         As String * 1       'T/B�敪
    REPSMPLIDCS     As Long             '��\�T���v��ID         Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    XTALCS          As String * 12      '�����ԍ�
    INPOSCS         As Integer          '�������ʒu
    HINBCS          As String * 8       '�i��
    REVNUMCS        As Integer          '���i�ԍ������ԍ�
    FACTORYCS       As String * 1       '�H��
    OPECS           As String * 1       '���Ə���
    KTKBNCS         As String * 1       '�m��敪
    BLKKTFLAGCS     As String * 1       '�u���b�N�m��t���O
    CRYSMPLIDRSCS   As Long             '�T���v��ID(Rs)         Integer��Long �T���v����6���Ή�
    CRYSMPLIDRS1CS  As Long             '����T���v��ID1(Rs)    Integer��Long �T���v����6���Ή�
    CRYSMPLIDRS2CS  As Long             '����T���v��ID2(Rs)    Integer��Long �T���v����6���Ή�
    CRYINDRSCS      As String * 1       '���FLG(Rs)
    CRYRESRS1CS     As String * 1       '����FLG1(Rs)
    CRYRESRS2CS     As String * 1       '����FLG2(Rs)
    CRYSMPLIDOICS   As Long             '�T���v��ID(Oi)         Integer��Long �T���v����6���Ή�
    CRYINDOICS      As String * 1       '���FLG(Oi)
    CRYRESOICS      As String * 1       '����FLG(Oi)
    CRYSMPLIDB1CS   As Long             '�T���v��ID(B1)         Integer��Long �T���v����6���Ή�
    CRYINDB1CS      As String * 1       '���FLG(B1)
    CRYRESB1CS      As String * 1       '����FLG(B1)
    CRYSMPLIDB2CS   As Long             '�T���v��ID(B2)         Integer��Long �T���v����6���Ή�
    CRYINDB2CS      As String * 1       '���FLG(B2)
    CRYRESB2CS      As String * 1       '����FLG(B2)
    CRYSMPLIDB3CS   As Long             '�T���v��ID(B3)         Integer��Long �T���v����6���Ή�
    CRYINDB3CS      As String * 1       '���FLG(B3)
    CRYRESB3CS      As String * 1       '����FLG(B3)
    CRYSMPLIDL1CS   As Long             '�T���v��ID(L1)         Integer��Long �T���v����6���Ή�
    CRYINDL1CS      As String * 1       '���FLG(L1)
    CRYRESL1CS      As String * 1       '����FLG(L1)
    CRYSMPLIDL2CS   As Long             '�T���v��ID(L2)         Integer��Long �T���v����6���Ή�
    CRYINDL2CS      As String * 1       '���FLG(L2)
    CRYRESL2CS      As String * 1       '����FLG(L2)
    CRYSMPLIDL3CS   As Long             '�T���v��ID(L3)         Integer��Long �T���v����6���Ή�
    CRYINDL3CS      As String * 1       '���FLG(L3)
    CRYRESL3CS      As String * 1       '����FLG(L3)
    CRYSMPLIDL4CS   As Long             '�T���v��ID(L4)         Integer��Long �T���v����6���Ή�
    CRYINDL4CS      As String * 1       '���FLG(L4)
    CRYRESL4CS      As String * 1       '����FLG(L4)
    CRYSMPLIDCSCS   As Long             '�T���v��ID(Cs)         Integer��Long �T���v����6���Ή�
    CRYINDCSCS      As String * 1       '���FLG(Cs)
    CRYRESCSCS      As String * 1       '����FLG(Cs)
    CRYSMPLIDGDCS   As Long             '�T���v��ID(GD)         Integer��Long �T���v����6���Ή�
    CRYINDGDCS      As String * 1       '���FLG(GD)
    CRYRESGDCS      As String * 1       '����FLG(GD)
    CRYSMPLIDTCS    As Long             '�T���v��ID(T)          Integer��Long �T���v����6���Ή�
    CRYINDTCS       As String * 1       '���FLG(T)
    CRYRESTCS       As String * 1       '����FLG(T)
''Add Start 2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
    CRYREST10CS     As String * 1       '����FLG(T10)
''Add End   2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
    CRYSMPLIDEPCS   As Long             '�T���v��ID(EPD)        Integer��Long �T���v����6���Ή�
    CRYINDEPCS      As String * 1       '���FLG(EPD)
    CRYRESEPCS      As String * 1       '����FLG(EPD)
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP        As String * 1       'DK���x(����)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    CRYINDXC1       As String * 1       '���FLG(X)     2009/08/12 Kameda
    CRYRESXC1       As String * 1       '����FLG(X)     2009/08/12 Kameda
    SIRDKBNY3       As String * 1       '���FLG(SIRD)  2010/02/04 Kameda
    
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎d�l���ڒǉ�
    CRYSMPLIDCCS    As Long             ' �T���v��ID(C)
    CRYINDCCS       As String * 1       ' ���FLG(C)
    CRYRESCCS       As String * 1       ' ����FLG(C)
    CRYSMPLIDCJCS   As Long             ' �T���v��ID(CJ)
    CRYINDCJCS      As String * 1       ' ���FLG(CJ)
    CRYRESCJCS      As String * 1       ' ����FLG(CJ)
    CRYSMPLIDCJLTCS As Long             ' �T���v��ID(CJ[LT])
    CRYINDCJLTCS    As String * 1       ' ���FLG(CJ[LT])
    CRYRESCJLTCS    As String * 1       ' ����FLG(CJ[LT])
    CRYSMPLIDCJ2CS  As Long             ' �T���v��ID(CJ2)
    CRYINDCJ2CS     As String * 1       ' ���FLG(CJ2)
    CRYRESCJ2CS     As String * 1       ' ����FLG(CJ2)
  'Add End   2011/01/17 SMPK A.Nagamine
End Type

'���т��܂Ƃ߂��\����
Public Type type_DBDRV_scmzc_fcmkc001c_Zisseki
    CRYRZ() As type_DBDRV_scmzc_fcmkc001c_CryR
    OIZ()   As type_DBDRV_scmzc_fcmkc001c_Oi
    BMD1Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD2Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD3Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    OSF1Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF2Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF3Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF4Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    CSZ()   As type_DBDRV_scmzc_fcmkc001c_CS
    GDZ()   As type_DBDRV_scmzc_fcmkc001c_GD
    LTZ()   As type_DBDRV_scmzc_fcmkc001c_LT
    EPDZ()  As type_DBDRV_scmzc_fcmkc001c_EPD
    SURSZ() As type_DBDRV_scmzc_fcmkc001c_CryR
    XZ As type_DBDRV_scmzc_fcmkc001c_X
    SIRD As type_DBDRV_scmzc_fcmkc001c_SIRD
    
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ� SB_CryHanSui.bas
    CuC()       As type_DBDRV_scmzc_fcmkc001c_C     'C     ����
    CuCJ()      As type_DBDRV_scmzc_fcmkc001c_CJ    'CJ    ����
    CuCJLT()    As type_DBDRV_scmzc_fcmkc001c_CJLT  'CJ(LT)����
    CuCJ2()     As type_DBDRV_scmzc_fcmkc001c_CJ2   'CJ2   ����
  'Add End   2011/01/17 SMPK A.Nagamine
End Type

'���茋�ʂ�J014�����v�ۍ\����
Public Type Judg_Spec_Cry
    Enable  As Boolean          '�L���ȕi�Ԃł���
    rs      As Boolean          'Rs�͗v����
    Oi      As Boolean          'Oi�͗v����
    B1      As Boolean          'BMD1�͗v����
    B2      As Boolean          'BMD2�͗v����
    B3      As Boolean          'BMD3�͗v����
    L1      As Boolean          'OSF1�͗v����
    L2      As Boolean          'OSF2�͗v����
    L3      As Boolean          'OSF3�͗v����
    L4      As Boolean          'OSF4�͗v����
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
    COSF3   As Boolean          'C-OSF3�׸�
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---
    Cs      As Boolean          'Cs�͗v����
    GD      As Boolean          'GD�͗v����
    Lt      As Boolean          'LT�͗v����
    EPD     As Boolean          'EPD�͗v����
    
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ�
    CuC     As Boolean          'C�͗v����
    CuCJ    As Boolean          'CJ�͗v����
    CuCJLT  As Boolean          'CJ(LT)�͗v����
    CuCJ2   As Boolean          'CJ2�͗v����
  'Add End   2011/01/17 SMPK A.Nagamine
End Type

' �d�l�̎w���������Ă��锻�f�p
Public Const SIJI = "H"
Public Const SANKOU = "S"

'�T�v      :�������� �e��f�[�^�擾
'���Ұ�    :�ϐ���        ,IO ,�^                                 ,����
'          :inBlockID     ,I  ,String                             ,�Ώۃu���b�NID
'          :tNew_Hinban   ,I  ,tFullHinban                        ,�Ώەi��(�\����)
'          :Siyou()       ,O  ,type_DBDRV_scmzc_fcmkc001c_Siyou   ,�i�ԁA�d�l�A���������擾�p
'          :CrySmp()      ,O  ,type_DBDRV_scmzc_fcmkc001c_CrySmp  ,�����T���v���Ǘ��擾�p
'          :Zisseki       ,O  ,type_DBDRV_scmzc_fcmkc001c_Zisseki ,���їp
'          :sErrMsg       ,O  ,String                             ,
'          :iSmpGetFlg    ,I  ,Integer                            :����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
'          :iSamplID1     ,I  ,Long                               :TOP�����ID(�ȗ���)   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iSamplID2     ,I  ,Long                               :BOT�����ID(�ȗ���)   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O  ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'����      :
'����      :2001/06/26 ���{ �쐬
Public Function funCryGetDataEtc(inBlockID As String, tNew_Hinban As tFullHinban, _
                                 siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                                 CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                                 Zisseki As type_DBDRV_scmzc_fcmkc001c_Zisseki, _
                                 sErrMsg As String, _
                                 iSmpGetFlg As Integer, iSamplID1 As Long, iSamplID2 As Long) As FUNCTION_RETURN

    Dim chk_cnt As Integer
    Dim i       As Integer
    Dim recCnt  As Integer
    Dim sDbName As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function funCryGetDataEtc"

    funCryGetDataEtc = FUNCTION_RETURN_FAILURE

    sDbName = "V011"
    '�i�ԁASXL�d�l����f�[�^�̎擾�i���R�[�h0���̏ꍇ���G���[�j
    If getHinSiyou(inBlockID, tNew_Hinban, siyou()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        funCryGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '����i�Ԃ��R�s�[
    chk_cnt = UBound(siyou)
    If chk_cnt = 1 Then
        ReDim Preserve siyou(chk_cnt + 1)
        siyou(chk_cnt + 1) = siyou(chk_cnt)
    End If
    
    sDbName = "V010"
    '�����T���v���̎擾(���R�[�h0���̏ꍇ���G���[)
    If getCrySmp(inBlockID, CrySmp(), iSmpGetFlg, iSamplID1, iSamplID2) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        funCryGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) :
    recCnt = UBound(CrySmp)
  'Add End   2011/01/17 SMPK A.Nagamine

    With Zisseki
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) :
        'ReDim .CRYRZ(2)
        'ReDim .OIZ(2)
        'ReDim .BMD1Z(2)
        'ReDim .BMD2Z(2)
        'ReDim .BMD3Z(2)
        'ReDim .OSF1Z(2)
        'ReDim .OSF2Z(2)
        'ReDim .OSF3Z(2)
        'ReDim .OSF4Z(2)
        'ReDim .CSZ(2)
        'ReDim .GDZ(2)
        'ReDim .LTZ(2)
        'ReDim .EPDZ(2)
        'ReDim .SURSZ(2)
        
        ReDim .CRYRZ(recCnt)
        ReDim .OIZ(recCnt)
        ReDim .BMD1Z(recCnt)
        ReDim .BMD2Z(recCnt)
        ReDim .BMD3Z(recCnt)
        ReDim .OSF1Z(recCnt)
        ReDim .OSF2Z(recCnt)
        ReDim .OSF3Z(recCnt)
        ReDim .OSF4Z(recCnt)
        ReDim .CSZ(recCnt)
        ReDim .GDZ(recCnt)
        ReDim .LTZ(recCnt)
        ReDim .EPDZ(recCnt)
        ReDim .SURSZ(recCnt)
  'Add End   2011/01/17 SMPK A.Nagamine
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ�
        ReDim .CuC(recCnt)
        ReDim .CuCJ(recCnt)
        ReDim .CuCJLT(recCnt)
        ReDim .CuCJ2(recCnt)
  'Add End   2011/01/17 SMPK A.Nagamine
    End With
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) :
'    recCnt = UBound(CrySmp)
  'Add End   2011/01/17 SMPK A.Nagamine

    '�����T���v���̎w�������Ď��т����
    For i = 1 To recCnt
        
        sDbName = "J002"
        If CryR_Zisseki(siyou(i), CrySmp(i), Zisseki.CRYRZ(i), Zisseki.SURSZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J003"
        If Oi_Zisseki(siyou(i), CrySmp(i), Zisseki.OIZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "1", Zisseki.BMD1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "2", Zisseki.BMD2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "3", Zisseki.BMD3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "1", Zisseki.OSF1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "2", Zisseki.OSF2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "3", Zisseki.OSF3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "4", Zisseki.OSF4Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J004"
        If CS_Zisseki(siyou(i), CrySmp(i), Zisseki.CSZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J006"
        If GD_Zisseki(siyou(i), CrySmp(i), Zisseki.GDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J007"
        If LT_Zisseki(siyou(i), CrySmp(i), Zisseki.LTZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J001"
        If EPD_Zisseki(siyou(i), CrySmp(i), Zisseki.EPDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        
      'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ�
        sDbName = "J023"
        If CuDeco_C_Zisseki(siyou(i), CrySmp(i), Zisseki.CuC(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJ_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJLT_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJLT(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJ2_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJ2(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
      'Add End   2011/01/17 SMPK A.Nagamine
        
    Next
    '2009/08/12 Kameda
    'X����������t���O�̎擾
    If GetXSDC1_XRAY(CrySmp(recCnt)) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", "XSDC1_XRAY")
        funCryGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    sDbName = "J021"
    If X_Zisseki(CrySmp(recCnt).XTALCS, Zisseki.XZ) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
    
    '2010/02/04 Kameda
    'SIRD�]���敪�擾
    If GetXODY3_SIRD(CrySmp(1)) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", "XODY3_SIRD")
        funCryGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    sDbName = "J022"
    If SIRD_Zisseki(CrySmp(1), Zisseki.SIRD) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
    
    
    
    sDbName = ""
    funCryGetDataEtc = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    If Trim$(sDbName) <> "" Then sErrMsg = GetMsgStr("EGET2", sDbName)
    If recCnt > 2 Then
        sErrMsg = "0"
    End If
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    funCryGetDataEtc = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�������� �e��f�[�^�擾(������������F���f�f�[�^�̍��۔�����s��Ȃ��p)
'���Ұ�    :�ϐ���        ,IO ,�^                                 ,����
'          :inBlockID     ,I  ,String                             ,�Ώۃu���b�NID
'          :Top_Hinban      ,I  ,tFullHinban                      ,TOP�i��
'          :Tail_Hinban     ,I  ,tFullHinban                      ,TAIL�i��
'          :Siyou()       ,O  ,type_DBDRV_scmzc_fcmkc001c_Siyou   ,�i�ԁA�d�l�A���������擾�p
'          :CrySmp()      ,O  ,type_DBDRV_scmzc_fcmkc001c_CrySmp  ,�����T���v���Ǘ��擾�p
'          :Zisseki       ,O  ,type_DBDRV_scmzc_fcmkc001c_Zisseki ,���їp
'          :sErrMsg       ,O  ,String                             ,�װү���޺���
'          :iSmpGetFlg    ,I  ,Integer                            ,����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
'          :iSamplID1     ,I  ,Long                               ,TOP�����ID(�ȗ���)   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iSamplID2     ,I  ,Long                               ,BOT�����ID(�ȗ���)   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O  ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'����      :
'����      :2005/02/08 �쐬  ffc)tanabe
Public Function funCryGetDataEtc2(inBlockID As String, Top_Hinban As tFullHinban, Tail_Hinban As tFullHinban, _
                                 siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                                 CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                                 Zisseki As type_DBDRV_scmzc_fcmkc001c_Zisseki, _
                                 sErrMsg As String, _
                                 iSmpGetFlg As Integer, iSamplID1 As Long, iSamplID2 As Long) As FUNCTION_RETURN

    Dim i       As Integer                              'for���p�ϐ�
    Dim recCnt  As Integer                              '�����T���v���w������
    Dim t_Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou   '�d�l�\����
    Dim sDbName As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function funCryGetDataEtc2"

    funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE

    '�d�l�z��̏�����
    ReDim siyou(2)

    sDbName = "V011"
    'TOP��
    '�i�ԁASXL�d�l����f�[�^�̎擾�i���R�[�h0���̏ꍇ���G���[�j
    If getHinSiyou(inBlockID, Top_Hinban, t_Siyou()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    'TOP���̎d�l�f�[�^���i�[����B
    siyou(1) = t_Siyou(1)

    'TAIL��
    '�i�ԁASXL�d�l����f�[�^�̎擾�i���R�[�h0���̏ꍇ���G���[�j
    If getHinSiyou(inBlockID, Tail_Hinban, t_Siyou()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    'TAIL���̎d�l�f�[�^���i�[����B
    siyou(2) = t_Siyou(1)
    
    sDbName = "V010"
    '�����T���v���̎擾(���R�[�h0���̏ꍇ���G���[)
    If getCrySmp(inBlockID, CrySmp(), iSmpGetFlg, iSamplID1, iSamplID2) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) :
    recCnt = UBound(CrySmp)
  'Add End   2011/01/17 SMPK A.Nagamine

    With Zisseki
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) :
        'ReDim .CRYRZ(2)
        'ReDim .OIZ(2)
        'ReDim .BMD1Z(2)
        'ReDim .BMD2Z(2)
        'ReDim .BMD3Z(2)
        'ReDim .OSF1Z(2)
        'ReDim .OSF2Z(2)
        'ReDim .OSF3Z(2)
        'ReDim .OSF4Z(2)
        'ReDim .CSZ(2)
        'ReDim .GDZ(2)
        'ReDim .LTZ(2)
        'ReDim .EPDZ(2)
        'ReDim .SURSZ(2)
  
        ReDim .CRYRZ(recCnt)
        ReDim .OIZ(recCnt)
        ReDim .BMD1Z(recCnt)
        ReDim .BMD2Z(recCnt)
        ReDim .BMD3Z(recCnt)
        ReDim .OSF1Z(recCnt)
        ReDim .OSF2Z(recCnt)
        ReDim .OSF3Z(recCnt)
        ReDim .OSF4Z(recCnt)
        ReDim .CSZ(recCnt)
        ReDim .GDZ(recCnt)
        ReDim .LTZ(recCnt)
        ReDim .EPDZ(recCnt)
        ReDim .SURSZ(recCnt)
  'Add End   2011/01/17 SMPK A.Nagamine
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ�
        ReDim .CuC(recCnt)
        ReDim .CuCJ(recCnt)
        ReDim .CuCJLT(recCnt)
        ReDim .CuCJ2(recCnt)
  'Add End   2011/01/17 SMPK A.Nagamine
    End With
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) :
'    recCnt = UBound(CrySmp)
  'Add End   2011/01/17 SMPK A.Nagamine

    '�����T���v���̎w�������Ď��т����
    For i = 1 To recCnt
        
        sDbName = "J002"
        If CryR_Zisseki(siyou(i), CrySmp(i), Zisseki.CRYRZ(i), Zisseki.SURSZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J003"
        If Oi_Zisseki(siyou(i), CrySmp(i), Zisseki.OIZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "1", Zisseki.BMD1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "2", Zisseki.BMD2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "3", Zisseki.BMD3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "1", Zisseki.OSF1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "2", Zisseki.OSF2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "3", Zisseki.OSF3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "4", Zisseki.OSF4Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J004"
        If CS_Zisseki(siyou(i), CrySmp(i), Zisseki.CSZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J006"
        If GD_Zisseki(siyou(i), CrySmp(i), Zisseki.GDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J007"
        If LT_Zisseki(siyou(i), CrySmp(i), Zisseki.LTZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J001"
        If EPD_Zisseki(siyou(i), CrySmp(i), Zisseki.EPDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        
      'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̍��ڒǉ�
        sDbName = "J023"
        If CuDeco_C_Zisseki(siyou(i), CrySmp(i), Zisseki.CuC(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJ_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJLT_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJLT(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJ2_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJ2(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
      'Add End   2011/01/17 SMPK A.Nagamine
        
    Next
    '2009/08/12 Kameda
    'X����������t���O�̎擾
    If GetXSDC1_XRAY(CrySmp(recCnt)) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", "XSDC1_XRAY")
        funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    sDbName = "J021"
    If X_Zisseki(CrySmp(recCnt).XTALCS, Zisseki.XZ) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
    
    '2010/02/04 Kameda
    'SIRD�]���敪�擾
    If GetXODY3_SIRD(CrySmp(1)) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", "XODY3_SIRD")
        funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    sDbName = "J022"
    If SIRD_Zisseki(CrySmp(1), Zisseki.SIRD) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
    
    sDbName = ""
    funCryGetDataEtc2 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    If Trim$(sDbName) <> "" Then sErrMsg = GetMsgStr("EGET2", sDbName)
    If recCnt > 2 Then
        sErrMsg = "0"
    End If
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�����֐� �i�ԁA�d�l���擾����
'���Ұ�    :�ϐ���        ,IO ,�^                                 ,����
'          :inBlockID     ,I  ,String                             ,�Ώۃu���b�NID
'          :tNew_Hinban   ,I  ,tFullHinban                        ,�Ώەi��(�\����)
'          :Siyou()       ,O  ,type_DBDRV_scmzc_fcmkc001c_Siyou   ,�i�ԁA�d�l�A���������擾�p
'          :�߂�l        ,O  ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'����      :
'����      :
Public Function getHinSiyou(inBlockID As String, tNew_Hinban As tFullHinban, _
                            siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim recCnt      As Integer
    Dim i           As Long
    Dim Jiltuseki   As Judg_Kakou
    Dim iIngotPos   As Integer          '�������ʒu
    Dim iLength     As Integer          '����
    Dim sCryNum     As String           '�����ԍ�
    
    '�i�ԁASXL�d�l����f�[�^�̎擾
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 start
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function getHinSiyou"

    getHinSiyou = FUNCTION_RETURN_SUCCESS

    If ciSmpGetFlg = 0 Then
        sql = "select "
        '��ۯ��Ǘ�(TBCME040)�Q�ƒ�~�@05/10/04 ooba START ===================================>
        sql = sql & "CSTOP.XTALCS as CRYNUM, "                      ' �����ԍ�
        sql = sql & "CSTOP.INPOSCS as INGOTPOS, "                   ' �������J�n�ʒu
        sql = sql & "CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, "     ' ����
        '��ۯ��Ǘ�(TBCME040)�Q�ƒ�~�@05/10/04 ooba END =====================================>
    Else
        '�H�������ް��擾�֐������ް����擾���ݒ肷��
        If GET_hurikaeC3(inBlockID, ciKcnt, iIngotPos, iLength, sCryNum) = FUNCTION_RETURN_FAILURE Then
            getHinSiyou = FUNCTION_RETURN_FAILURE
            ReDim siyou(0)
            GoTo proc_exit
        End If
            
        sql = "select "
        sql = sql & sCryNum & " as CRYNUM, "        ' �����ԍ�
        sql = sql & iIngotPos & " as INGOTPOS, "    ' �������J�n�ʒu
        sql = sql & iLength & " as LENGTH, "        ' ����
    End If
    
    sql = sql & "E037.PRODCOND, "           ' �������
    sql = sql & "E037.PGID, "               ' �o�f�|�h�c
    sql = sql & "E037.UPLENGTH, "           ' ���グ����
    sql = sql & "E037.FREELENG, "           ' �t���[��
    sql = sql & "E037.DIAMETER, "           ' ���a
    sql = sql & "E037.CHARGE, "             ' �`���[�W��
    sql = sql & "E037.SEED, "               ' �V�[�h
    sql = sql & "E037.ADDDPPOS, "           ' �ǉ��h�[�v�ʒu
    
    sql = sql & "E018.HSXTYPE, "             ' �i�r�w�^�C�v
    sql = sql & "E018.HSXD1CEN, "            ' �i�r�w���a�P���S
    sql = sql & "E018.HSXCDIR, "             ' �i�r�w�����ʕ���
    
    sql = sql & "E018.HSXRMIN, "             ' �i�r�w���R����
    sql = sql & "E018.HSXRMAX, "             ' �i�r�w���R���
    sql = sql & "E018.HSXRAMIN, "            ' �i�r�w���R���ω���
    sql = sql & "E018.HSXRAMAX, "            ' �i�r�w���R���Ϗ��
    sql = sql & "E018.HSXRMCAL, "            ' �i�r�w���R�ʓ��v�Z�@�@'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
    sql = sql & "E018.HSXRMBNP, "            ' �i�r�w���R�ʓ����z
    sql = sql & "E018.HSXRSPOH, "            ' �i�r�w���R����ʒu�Q��
    sql = sql & "E018.HSXRSPOT, "            ' �i�r�w���R����ʒu�Q�_
    sql = sql & "E018.HSXRSPOI, "            ' �i�r�w���R����ʒu�Q��
    sql = sql & "E018.HSXRHWYT, "            ' �i�r�w���R�ۏؕ��@�Q��
    sql = sql & "E018.HSXRHWYS, "            ' �i�r�w���R�ۏؕ��@�Q��

    sql = sql & "E019.HSXONMIN, "            ' �i�r�w�_�f�Z�x����
    sql = sql & "E019.HSXONMAX, "            ' �i�r�w�_�f�Z�x���
    sql = sql & "E019.HSXONAMN, "            ' �i�r�w�_�f�Z�x���ω���
    sql = sql & "E019.HSXONAMX, "            ' �i�r�w�_�f�Z�x���Ϗ��
    sql = sql & "E019.HSXONMCL, "            ' �i�r�w�_�f�Z�x�ʓ��v�Z�@�@'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
    sql = sql & "E019.HSXONMBP, "            ' �i�r�w�_�f�Z�x�ʓ����z
    sql = sql & "E019.HSXONSPH, "            ' �i�r�w�_�f�Z�x����ʒu�Q��
    sql = sql & "E019.HSXONSPT, "            ' �i�r�w�_�f�Z�x����ʒu�Q�_
    sql = sql & "E019.HSXONSPI, "            ' �i�r�w�_�f�Z�x����ʒu�Q��
    sql = sql & "E019.HSXONHWT, "            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    sql = sql & "E019.HSXONHWS, "            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��

    sql = sql & "E020.HSXBM1AN, "            ' �i�r�w�a�l�c�P���ω���
    sql = sql & "E020.HSXBM1AX, "            ' �i�r�w�a�l�c�P���Ϗ��
    sql = sql & "E020.HSXBM2AN, "            ' �i�r�w�a�l�c�Q���ω���
    sql = sql & "E020.HSXBM2AX, "            ' �i�r�w�a�l�c�Q���Ϗ��
    sql = sql & "E020.HSXBM3AN, "            ' �i�r�w�a�l�c�R���ω���
    sql = sql & "E020.HSXBM3AX, "            ' �i�r�w�a�l�c�R���Ϗ��
    sql = sql & "E020.HSXBM1SH, "            ' �i�r�w�a�l�c�P����ʒu�Q��
    sql = sql & "E020.HSXBM1ST, "            ' �i�r�w�a�l�c�P����ʒu�Q�_
    sql = sql & "E020.HSXBM1SR, "            ' �i�r�w�a�l�c�P����ʒu�Q��
    sql = sql & "E020.HSXBM1HT, "            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    sql = sql & "E020.HSXBM1HS, "            ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    sql = sql & "E020.HSXBM2SH, "            ' �i�r�w�a�l�c�Q����ʒu�Q��
    sql = sql & "E020.HSXBM2ST, "            ' �i�r�w�a�l�c�Q����ʒu�Q�_
    sql = sql & "E020.HSXBM2SR, "            ' �i�r�w�a�l�c�Q����ʒu�Q��
    sql = sql & "E020.HSXBM2HT, "            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    sql = sql & "E020.HSXBM2HS, "            ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    sql = sql & "E020.HSXBM3SH, "            ' �i�r�w�a�l�c�R����ʒu�Q��
    sql = sql & "E020.HSXBM3ST, "            ' �i�r�w�a�l�c�R����ʒu�Q�_
    sql = sql & "E020.HSXBM3SR, "            ' �i�r�w�a�l�c�R����ʒu�Q��
    sql = sql & "E020.HSXBM3HT, "            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
    sql = sql & "E020.HSXBM3HS, "            ' �i�r�w�a�l�c�R�ۏؕ��@�Q��

    sql = sql & "E020.HSXOF1AX, "            ' �i�r�w�n�r�e�P���Ϗ��
    sql = sql & "E020.HSXOF1MX, "            ' �i�r�w�n�r�e�P���
    sql = sql & "E020.HSXOF2AX, "            ' �i�r�w�n�r�e�Q���Ϗ��
    sql = sql & "E020.HSXOF2MX, "            ' �i�r�w�n�r�e�Q���
    sql = sql & "E020.HSXOF3AX, "            ' �i�r�w�n�r�e�R���Ϗ��
    sql = sql & "E020.HSXOF3MX, "            ' �i�r�w�n�r�e�R���
    sql = sql & "E020.HSXOF4AX, "            ' �i�r�w�n�r�e�S���Ϗ��
    sql = sql & "E020.HSXOF4MX, "            ' �i�r�w�n�r�e�S���
    sql = sql & "E020.HSXOF1SH, "            ' �i�r�w�n�r�e�P����ʒu�Q��
    sql = sql & "E020.HSXOF1ST, "            ' �i�r�w�n�r�e�P����ʒu�Q�_
    sql = sql & "E020.HSXOF1SR, "            ' �i�r�w�n�r�e�P����ʒu�Q��
    sql = sql & "E020.HSXOF1HT, "            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    sql = sql & "E020.HSXOF1HS, "            ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
    sql = sql & "E020.HSXOF2SH, "            ' �i�r�w�n�r�e�Q����ʒu�Q��
    sql = sql & "E020.HSXOF2ST, "            ' �i�r�w�n�r�e�Q����ʒu�Q�_
    sql = sql & "E020.HSXOF2SR, "            ' �i�r�w�n�r�e�Q����ʒu�Q��
    sql = sql & "E020.HSXOF2HT, "            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    sql = sql & "E020.HSXOF2HS, "            ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
    sql = sql & "E020.HSXOF3SH, "            ' �i�r�w�n�r�e�R����ʒu�Q��
    sql = sql & "E020.HSXOF3ST, "            ' �i�r�w�n�r�e�R����ʒu�Q�_
    sql = sql & "E020.HSXOF3SR, "            ' �i�r�w�n�r�e�R����ʒu�Q��
    sql = sql & "E020.HSXOF3HT, "            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    sql = sql & "E020.HSXOF3HS, "            ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
    sql = sql & "E020.HSXOF4SH, "            ' �i�r�w�n�r�e�S����ʒu�Q��
    sql = sql & "E020.HSXOF4ST, "            ' �i�r�w�n�r�e�S����ʒu�Q�_
    sql = sql & "E020.HSXOF4SR, "            ' �i�r�w�n�r�e�S����ʒu�Q��
    sql = sql & "E020.HSXOF4HT, "            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    sql = sql & "E020.HSXOF4HS, "            ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
    sql = sql & "E020.HSXOF1NS, "            ' �i�r�w�n�r�e�P�M�����@
    sql = sql & "E020.HSXOF2NS, "            ' �i�r�w�n�r�e�Q�M�����@
    sql = sql & "E020.HSXOF3NS, "            ' �i�r�w�n�r�e�R�M�����@
    sql = sql & "E020.HSXOF4NS, "            ' �i�r�w�n�r�e�S�M�����@
    sql = sql & "E020.HSXBM1NS, "            ' �i�r�w�a�l�c�P�M�����@
    sql = sql & "E020.HSXBM2NS, "            ' �i�r�w�a�l�c�Q�M�����@
    sql = sql & "E020.HSXBM3NS, "            ' �i�r�w�a�l�c�R�M�����@

    sql = sql & "E019.HSXCNMIN, "            ' �i�r�w�Y�f�Z�x����
    sql = sql & "E019.HSXCNMAX, "            ' �i�r�w�Y�f�Z�x���
    sql = sql & "E019.HSXCNSPH, "            ' �i�r�w�Y�f�Z�x����ʒu�Q��
    sql = sql & "E019.HSXCNSPT, "            ' �i�r�w�Y�f�Z�x����ʒu�Q�_
    sql = sql & "E019.HSXCNSPI, "            ' �i�r�w�Y�f�Z�x����ʒu�Q��
    sql = sql & "E019.HSXCNHWT, "            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    sql = sql & "E019.HSXCNHWS, "            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    sql = sql & "E019.HSXCNKHI, "            ' �i�r�w�Y�f�Z�x�����p�x�Q�� 09/01/08 ooba

    sql = sql & "E020.HSXDENMX, "            ' �i�r�w�c�������
    sql = sql & "E020.HSXDENMN, "            ' �i�r�w�c��������
    sql = sql & "E020.HSXLDLMX, "            ' �i�r�w�k�^�c�k���
    sql = sql & "E020.HSXLDLMN, "            ' �i�r�w�k�^�c�k����
    sql = sql & "E020.HSXDVDMXN, "           ' �i�r�w�c�u�c�Q���   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sql = sql & "E020.HSXDVDMNN, "           ' �i�r�w�c�u�c�Q����   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sql = sql & "E020.HSXDENHT, "            ' �i�r�w�c�����ۏؕ��@�Q��
    sql = sql & "E020.HSXDENHS, "            ' �i�r�w�c�����ۏؕ��@�Q��
    sql = sql & "E020.HSXLDLHT, "            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    sql = sql & "E020.HSXLDLHS, "            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    sql = sql & "E020.HSXDVDHT, "            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    sql = sql & "E020.HSXDVDHS, "            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    sql = sql & "E020.HSXDENKU, "            ' �i�r�w�c���������L��
    sql = sql & "E020.HSXDVDKU, "            ' �i�r�w�c�u�c�Q�����L��
    sql = sql & "E020.HSXLDLKU, "            ' �i�r�w�k�^�c�k�����L��

    sql = sql & "E019.HSXLTMIN, "            ' �i�r�w�k�^�C������
    sql = sql & "E019.HSXLTMAX, "            ' �i�r�w�k�^�C�����
''Add Start 2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
    sql = sql & "E036.LTCONVAL, "            ' �i�r�w�kLT10����
''Add End   2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
    sql = sql & "E019.HSXLTSPH, "            ' �i�r�w�k�^�C������ʒu�Q��
    sql = sql & "E019.HSXLTSPT, "            ' �i�r�w�k�^�C������ʒu�Q�_
    sql = sql & "E019.HSXLTSPI, "            ' �i�r�w�k�^�C������ʒu�Q��
    sql = sql & "E019.HSXLTHWT, "            ' �i�r�w�k�^�C���ۏؕ��@�Q��
    sql = sql & "E019.HSXLTHWS, "            ' �i�r�w�k�^�C���ۏؕ��@�Q��
    sql = sql & "E036.EPDUP, "               ' EPD ���
    sql = sql & "E036.TOPREG, "              ' TOP�K��
    sql = sql & "E036.TAILREG, "             ' TAIL�K��
    sql = sql & "E036.BTMSPRT, "             ' �{�g���͏o�K��
'*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݐ��ǉ�
    sql = sql & "E036.HSXGDLINE, "           ' �i�r�w�k�f�c���C����
'*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݐ��ǉ�

'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
    sql = sql & "E036.COSF3FLAG, "           ' C-OSF3�׸�
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sql = sql & "NVL(E036.HSXDKTMP,' ') HSXDKTMP, "
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    sql = sql & "E036.HSXLDLRMN HSXLDLRMN, "
    sql = sql & "E036.HSXLDLRMX HSXLDLRMX, "
    sql = sql & "E036.HWFLDLRMN HWFLDLRMN, "
    sql = sql & "E036.HWFLDLRMX HWFLDLRMX, "
    sql = sql & "E036.HSXOF1ARPTK HSXOF1ARPTK, "
    sql = sql & "E036.HSXOFARMIN HSXOFARMIN, "
    sql = sql & "E036.HSXOFARMAX HSXOFARMAX, "
    sql = sql & "E036.HSXOFARMHMX HSXOFARMHMX, "
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
    sql = sql & "E020.HSXOSF1PTK, "          ' �i�r�w�n�r�e�P�p�^���敪
    sql = sql & "E020.HSXOSF2PTK, "          ' �i�r�w�n�r�e�Q�p�^���敪
    sql = sql & "E020.HSXOSF3PTK, "          ' �i�r�w�n�r�e�R�p�^���敪
    sql = sql & "E020.HSXOSF4PTK, "          ' �i�r�w�n�r�e�S�p�^���敪
    sql = sql & "E020.HSXBMD1MBP, "          ' �i�r�w�a�l�c�P�ʓ����z
    sql = sql & "E020.HSXBMD2MBP, "          ' �i�r�w�a�l�c�Q�ʓ����z
    sql = sql & "E020.HSXBMD3MBP, "          ' �i�r�w�a�l�c�R�ʓ����z
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
    
    'WF�d�l�擾�@08/04/15 ooba START ===========================================>
    sql = sql & "E021.HWFRHWYS, "            ' �i�v�e���R�ۏؕ��@�Q��
    sql = sql & "E025.HWFONHWS, "            ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
    sql = sql & "E029.HWFOF1HS, "            ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
    sql = sql & "E029.HWFOF2HS, "            ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
    sql = sql & "E029.HWFOF3HS, "            ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
    sql = sql & "E029.HWFOF4HS, "            ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
    sql = sql & "E029.HWFBM1HS, "            ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
    sql = sql & "E029.HWFBM2HS, "            ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
    sql = sql & "E029.HWFBM3HS, "            ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
    sql = sql & "E026.HWFDENHS, "            ' �i�v�e�c�����ۏؕ��@�Q��
    sql = sql & "E026.HWFDVDHS, "            ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��
    sql = sql & "E026.HWFLDLHS, "            ' �i�v�e�k�^�c�k�ۏؕ��@�Q��
    sql = sql & "E021.HWFRKHNN, "            ' �i�v�e���R�����p�x�Q��
    sql = sql & "E025.HWFONKHN, "            ' �i�v�e�_�f�Z�x�����p�x�Q��
    sql = sql & "E029.HWFOF1KN, "            ' �i�v�e�n�r�e�P�����p�x�Q��
    sql = sql & "E029.HWFOF2KN, "            ' �i�v�e�n�r�e�Q�����p�x�Q��
    sql = sql & "E029.HWFOF3KN, "            ' �i�v�e�n�r�e�R�����p�x�Q��
    sql = sql & "E029.HWFOF4KN, "            ' �i�v�e�n�r�e�S�����p�x�Q��
    sql = sql & "E029.HWFBM1KN, "            ' �i�v�e�a�l�c�P�����p�x�Q��
    sql = sql & "E029.HWFBM2KN, "            ' �i�v�e�a�l�c�Q�����p�x�Q��
    sql = sql & "E029.HWFBM3KN, "            ' �i�v�e�a�l�c�R�����p�x�Q��
    sql = sql & "E026.HWFGDKHN "             ' �i�v�e�f�c�����p�x�Q��
    'WF�d�l�擾�@08/04/15 ooba END =============================================>

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    sql = sql & ",E020.HSXGDPTK "            ' �i�r�w�f�c�p�^���敪
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    sql = sql & ",E018.HSXCSCEN "            ' �i�r�w�ʌX�����S  2009/08/12 Kameda
    sql = sql & ",E018.HSXCSMIN "            ' �i�r�w�ʌX������  2009/08/12 Kameda
    sql = sql & ",E018.HSXCSMAX "            ' �i�r�w�ʌX�����  2009/08/12 Kameda
    sql = sql & ",E018.HSXCYCEN "            ' �i�r�w�ʌX�����S  2009/09/01 Kameda
    sql = sql & ",E018.HSXCYMIN "            ' �i�r�w�ʌX������  2009/09/01 Kameda
    sql = sql & ",E018.HSXCYMAX "            ' �i�r�w�ʌX�����  2009/09/01 Kameda
    sql = sql & ",E018.HSXCTCEN "            ' �i�r�w�ʌX�����S  2009/09/01 Kameda
    sql = sql & ",E018.HSXCTMIN "            ' �i�r�w�ʌX������  2009/09/01 Kameda
    sql = sql & ",E018.HSXCTMAX "            ' �i�r�w�ʌX�����  2009/09/01 Kameda
    sql = sql & ",E048.HWFSIRDMX "           ' �iWF�ʓ������  2010/02/04 Kameda
    
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎d�l���ڒǉ�
    sql = sql & ",E020.HSXCPK,    E020.HSXCSZ,    E020.HSXCHT,    E020.HSXCHS,    E020.HSXCJPK   "
    sql = sql & ",E020.HSXCJNS,   E020.HSXCJHT,   E020.HSXCJHS,   E020.HSXCJLTPK, E020.HSXCJLTNS "
    sql = sql & ",E020.HSXCJLTHT, E020.HSXCJLTHS, E020.HSXCJ2PK,  E020.HSXCJ2NS,  E020.HSXCJ2HT  "
    sql = sql & ",E020.HSXCJ2HS,  E036.HSXCJLTBND "
  'Add End   2011/01/17 SMPK A.Nagamine
  'Add Start 2012/06/01 SMPK H.Ohkubo
    sql = sql & ",NVL(E020.HSXCOSF3PK,'4') as HSXCOSF3PK"    '�i�r�w�b�n�r�e�R�p�^�[���敪"
  'Add End 2012/06/01 SMPK H.Ohkubo
    If ciSmpGetFlg = 0 Then
        '��ۯ��Ǘ�(TBCME040)�Q�ƒ�~�@05/10/04 ooba START ===================================>
        sql = sql & " from TBCME037 E037, TBCME018 E018, TBCME019 E019, TBCME020 E020, TBCME036 E036, "
        sql = sql & "      TBCME021 E021, TBCME025 E025, TBCME026 E026, TBCME029 E029, TBCME048 E048, "  '08/04/15 ooba, 2010/02/04 Kameda addE048
        sql = sql & " (select CRYNUMCS, XTALCS, INPOSCS from XSDCS "
        sql = sql & " where TBKBNCS = 'T' and CRYNUMCS = '" & inBlockID & "' "
        sql = sql & " ) CSTOP, "
        sql = sql & " (select CRYNUMCS, XTALCS, INPOSCS from XSDCS "
        sql = sql & " where TBKBNCS = 'B' and CRYNUMCS = '" & inBlockID & "' "
        sql = sql & " ) CSBOT "
        sql = sql & " where CSTOP.CRYNUMCS = CSBOT.CRYNUMCS and "
        sql = sql & "       E037.CRYNUM = '" & left(inBlockID, 9) & "000' and "
        '��ۯ��Ǘ�(TBCME040)�Q�ƒ�~�@05/10/04 ooba END =====================================>
    Else
        sql = sql & " from TBCME037 E037, TBCME018 E018, TBCME019 E019, TBCME020 E020, TBCME036 E036, "
        sql = sql & "      TBCME021 E021, TBCME025 E025, TBCME026 E026, TBCME029 E029, TBCME048 E048 "   '08/04/15 ooba, 2010/02/04 Kameda addE048
        sql = sql & " where E037.CRYNUM = '" & left(inBlockID, 9) & "000' and "
    End If
    sql = sql & "       E018.HINBAN = '" & tNew_Hinban.hinban & "' and "
    sql = sql & "       E018.MNOREVNO = " & tNew_Hinban.mnorevno & " and "
    sql = sql & "       E018.FACTORY = '" & tNew_Hinban.factory & "' and "
    sql = sql & "       E018.OPECOND = '" & tNew_Hinban.opecond & "' and "
    sql = sql & "       E019.HINBAN = E018.HINBAN and E019.MNOREVNO = E018.MNOREVNO and E019.FACTORY = E018.FACTORY and E019.OPECOND = E018.OPECOND and "
    sql = sql & "       E020.HINBAN = E018.HINBAN and E020.MNOREVNO = E018.MNOREVNO and E020.FACTORY = E018.FACTORY and E020.OPECOND = E018.OPECOND and "
    sql = sql & "       E021.HINBAN = E018.HINBAN and E021.MNOREVNO = E018.MNOREVNO and E021.FACTORY = E018.FACTORY and E021.OPECOND = E018.OPECOND and "   '08/04/15 ooba
    sql = sql & "       E025.HINBAN = E018.HINBAN and E025.MNOREVNO = E018.MNOREVNO and E025.FACTORY = E018.FACTORY and E025.OPECOND = E018.OPECOND and "   '08/04/15 ooba
    sql = sql & "       E026.HINBAN = E018.HINBAN and E026.MNOREVNO = E018.MNOREVNO and E026.FACTORY = E018.FACTORY and E026.OPECOND = E018.OPECOND and "   '08/04/15 ooba
    sql = sql & "       E029.HINBAN = E018.HINBAN and E029.MNOREVNO = E018.MNOREVNO and E029.FACTORY = E018.FACTORY and E029.OPECOND = E018.OPECOND and "   '08/04/15 ooba
    sql = sql & "       E036.HINBAN = E018.HINBAN and E036.MNOREVNO = E018.MNOREVNO and E036.FACTORY = E018.FACTORY and E036.OPECOND = E018.OPECOND and "
    sql = sql & "       E048.HINBAN = E018.HINBAN and E048.MNOREVNO = E018.MNOREVNO and E048.FACTORY = E018.FACTORY and E048.OPECOND = E018.OPECOND "       '2010/02/04 Kameda
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
    End If

    recCnt = rs.RecordCount
    ReDim siyou(recCnt)
    For i = 1 To recCnt
    
        With siyou(i)
            .CRYNUM = rs("CRYNUM")                  ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")              ' �������J�n�ʒu
            .Length = rs("LENGTH")                  ' ����
            .HIN.hinban = tNew_Hinban.hinban        ' �i��
            .HIN.mnorevno = tNew_Hinban.mnorevno    ' ���i�ԍ������ԍ�
            .HIN.factory = tNew_Hinban.factory      ' �H��
            .HIN.opecond = tNew_Hinban.opecond      ' ���Ə���
            
            .PRODCOND = rs("PRODCOND")              ' �������
            .PGID = rs("PGID")                      ' �o�f�|�h�c
            .UPLENGTH = rs("UPLENGTH")              ' ���グ����
            .FREELENG = rs("FREELENG")              ' �t���[��
            .DIAMETER = rs("DIAMETER")              ' ���a
            .CHARGE = rs("CHARGE")                  ' �`���[�W��
            .SEED = rs("SEED")                      ' �V�[�h
            .ADDDPPOS = rs("ADDDPPOS")              ' �ǉ��h�[�v�ʒu
    
            .HSXTYPE = rs("HSXTYPE")                        ' �i�r�w�^�C�v"
            .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))        ' �i�r�w���a�P���S"         2003/12/10 SystemBrain Null�Ή�
            .HSXCDIR = rs("HSXCDIR")                        ' �i�r�w�����ʕ���"

            .HSXRMIN = fncNullCheck(rs("HSXRMIN"))          ' �i�r�w���R����          2003/12/10 SystemBrain Null�Ή�
            .HSXRMAX = fncNullCheck(rs("HSXRMAX"))          ' �i�r�w���R���          2003/12/10 SystemBrain Null�Ή�
            .HSXRAMIN = fncNullCheck(rs("HSXRAMIN"))        ' �i�r�w���R���ω���      2003/12/10 SystemBrain Null�Ή�
            .HSXRAMAX = fncNullCheck(rs("HSXRAMAX"))        ' �i�r�w���R���Ϗ��      2003/12/10 SystemBrain Null�Ή�
            .HSXRMCAL = rs("HSXRMCAL")                      ' �i�r�w���R�ʓ��v�Z     '' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
            .HSXRMBNP = fncNullCheck(rs("HSXRMBNP"))        ' �i�r�w���R�ʓ����z      2003/12/10 SystemBrain Null�Ή�
            .HSXRSPOH = rs("HSXRSPOH")                      ' �i�r�w���R����ʒu�Q��
            .HSXRSPOT = rs("HSXRSPOT")                      ' �i�r�w���R����ʒu�Q�_
            .HSXRSPOI = rs("HSXRSPOI")                      ' �i�r�w���R����ʒu�Q��
            .HSXRHWYT = rs("HSXRHWYT")                      ' �i�r�w���R�ۏؕ��@�Q��
            .HSXRHWYS = rs("HSXRHWYS")                      ' �i�r�w���R�ۏؕ��@�Q��

            .HSXONMIN = fncNullCheck(rs("HSXONMIN"))        ' �i�r�w�_�f�Z�x����        2003/12/10 SystemBrain Null�Ή�
            .HSXONMAX = fncNullCheck(rs("HSXONMAX"))        ' �i�r�w�_�f�Z�x���        2003/12/10 SystemBrain Null�Ή�
            .HSXONAMN = fncNullCheck(rs("HSXONAMN"))        ' �i�r�w�_�f�Z�x���ω���    2003/12/10 SystemBrain Null�Ή�
            .HSXONAMX = fncNullCheck(rs("HSXONAMX"))        ' �i�r�w�_�f�Z�x���Ϗ��    2003/12/10 SystemBrain Null�Ή�
            .HSXONMCL = rs("HSXONMCL")                      ' �i�r�w�_�f�Z�x�ʓ��v�Z   '' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
            .HSXONMBP = fncNullCheck(rs("HSXONMBP"))        ' �i�r�w�_�f�Z�x�ʓ����z    2003/12/10 SystemBrain Null�Ή�
            .HSXONSPH = rs("HSXONSPH")                      ' �i�r�w�_�f�Z�x����ʒu�Q��
            .HSXONSPT = rs("HSXONSPT")                      ' �i�r�w�_�f�Z�x����ʒu�Q�_
            .HSXONSPI = rs("HSXONSPI")                      ' �i�r�w�_�f�Z�x����ʒu�Q��
            .HSXONHWT = rs("HSXONHWT")                      ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
            .HSXONHWS = rs("HSXONHWS")                      ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��

            .HSXBM1AN = fncNullCheck(rs("HSXBM1AN"))        ' �i�r�w�a�l�c�P���ω���    2003/12/10 SystemBrain Null�Ή�
            .HSXBM1AX = fncNullCheck(rs("HSXBM1AX"))        ' �i�r�w�a�l�c�P���Ϗ��    2003/12/10 SystemBrain Null�Ή�
            .HSXBM1SH = rs("HSXBM1SH")                      ' �i�r�w�a�l�c�P����ʒu�Q��
            .HSXBM1ST = rs("HSXBM1ST")                      ' �i�r�w�a�l�c�P����ʒu�Q�_
            .HSXBM1SR = rs("HSXBM1SR")                      ' �i�r�w�a�l�c�P����ʒu�Q��
            .HSXBM1HT = rs("HSXBM1HT")                      ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
            .HSXBM1HS = rs("HSXBM1HS")                      ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
            .HSXBM1NS = rs("HSXBM1NS")                      ' �i�r�w�a�l�c�P�M�����@
            .HSXBM2AN = fncNullCheck(rs("HSXBM2AN"))        ' �i�r�w�a�l�c�Q���ω���    2003/12/10 SystemBrain Null�Ή�
            .HSXBM2AX = fncNullCheck(rs("HSXBM2AX"))        ' �i�r�w�a�l�c�Q���Ϗ��    2003/12/10 SystemBrain Null�Ή�
            .HSXBM2SH = rs("HSXBM2SH")                      ' �i�r�w�a�l�c�Q����ʒu�Q��
            .HSXBM2ST = rs("HSXBM2ST")                      ' �i�r�w�a�l�c�Q����ʒu�Q�_
            .HSXBM2SR = rs("HSXBM2SR")                      ' �i�r�w�a�l�c�Q����ʒu�Q��
            .HSXBM2HT = rs("HSXBM2HT")                      ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
            .HSXBM2HS = rs("HSXBM2HS")                      ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
            .HSXBM2NS = rs("HSXBM2NS")                      ' �i�r�w�a�l�c�Q�M�����@
            .HSXBM3AN = fncNullCheck(rs("HSXBM3AN"))        ' �i�r�w�a�l�c�R���ω���    2003/12/10 SystemBrain Null�Ή�
            .HSXBM3AX = fncNullCheck(rs("HSXBM3AX"))        ' �i�r�w�a�l�c�R���Ϗ��    2003/12/10 SystemBrain Null�Ή�
            .HSXBM3SH = rs("HSXBM3SH")                      ' �i�r�w�a�l�c�R����ʒu�Q��
            .HSXBM3ST = rs("HSXBM3ST")                      ' �i�r�w�a�l�c�R����ʒu�Q�_
            .HSXBM3SR = rs("HSXBM3SR")                      ' �i�r�w�a�l�c�R����ʒu�Q��
            .HSXBM3HT = rs("HSXBM3HT")                      ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
            .HSXBM3HS = rs("HSXBM3HS")                      ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
            .HSXBM3NS = rs("HSXBM3NS")                      ' �i�r�w�a�l�c�R�M�����@
            
            .HSXOF1AX = fncNullCheck(rs("HSXOF1AX"))        ' �i�r�w�n�r�e�P���Ϗ��    2003/12/10 SystemBrain Null�Ή�
            .HSXOF1MX = fncNullCheck(rs("HSXOF1MX"))        ' �i�r�w�n�r�e�P���        2003/12/10 SystemBrain Null�Ή�
            .HSXOF1SH = rs("HSXOF1SH")                      ' �i�r�w�n�r�e�P����ʒu�Q��
            .HSXOF1ST = rs("HSXOF1ST")                      ' �i�r�w�n�r�e�P����ʒu�Q�_
            .HSXOF1SR = rs("HSXOF1SR")                      ' �i�r�w�n�r�e�P����ʒu�Q��
            .HSXOF1HT = rs("HSXOF1HT")                      ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
            .HSXOF1HS = rs("HSXOF1HS")                      ' �i�r�w�n�r�e�P�ۏؕ��@�Q��
            .HSXOF1NS = rs("HSXOF1NS")                      ' �i�r�w�n�r�e�P�M�����@
            .HSXOF2AX = fncNullCheck(rs("HSXOF2AX"))        ' �i�r�w�n�r�e�Q���Ϗ��    2003/12/10 SystemBrain Null�Ή�
            .HSXOF2MX = fncNullCheck(rs("HSXOF2MX"))        ' �i�r�w�n�r�e�Q���        2003/12/10 SystemBrain Null�Ή�
            .HSXOF2SH = rs("HSXOF2SH")                      ' �i�r�w�n�r�e�Q����ʒu�Q��
            .HSXOF2ST = rs("HSXOF2ST")                      ' �i�r�w�n�r�e�Q����ʒu�Q�_
            .HSXOF2SR = rs("HSXOF2SR")                      ' �i�r�w�n�r�e�Q����ʒu�Q��
            .HSXOF2HT = rs("HSXOF2HT")                      ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
            .HSXOF2HS = rs("HSXOF2HS")                      ' �i�r�w�n�r�e�Q�ۏؕ��@�Q��
            .HSXOF2NS = rs("HSXOF2NS")                      ' �i�r�w�n�r�e�Q�M�����@
            .HSXOF3AX = fncNullCheck(rs("HSXOF3AX"))        ' �i�r�w�n�r�e�R���Ϗ��    2003/12/10 SystemBrain Null�Ή�
            .HSXOF3MX = fncNullCheck(rs("HSXOF3MX"))        ' �i�r�w�n�r�e�R���        2003/12/10 SystemBrain Null�Ή�
            .HSXOF3SH = rs("HSXOF3SH")                      ' �i�r�w�n�r�e�R����ʒu�Q��
            .HSXOF3ST = rs("HSXOF3ST")                      ' �i�r�w�n�r�e�R����ʒu�Q�_
            .HSXOF3SR = rs("HSXOF3SR")                      ' �i�r�w�n�r�e�R����ʒu�Q��
            .HSXOF3HT = rs("HSXOF3HT")                      ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
            .HSXOF3HS = rs("HSXOF3HS")                      ' �i�r�w�n�r�e�R�ۏؕ��@�Q��
            .HSXOF3NS = rs("HSXOF3NS")                      ' �i�r�w�n�r�e�R�M�����@
            .HSXOF4AX = fncNullCheck(rs("HSXOF4AX"))        ' �i�r�w�n�r�e�S���Ϗ��    2003/12/10 SystemBrain Null�Ή�
            .HSXOF4MX = fncNullCheck(rs("HSXOF4MX"))        ' �i�r�w�n�r�e�S���        2003/12/10 SystemBrain Null�Ή�
            .HSXOF4SH = rs("HSXOF4SH")                      ' �i�r�w�n�r�e�S����ʒu�Q��
            .HSXOF4ST = rs("HSXOF4ST")                      ' �i�r�w�n�r�e�S����ʒu�Q�_
            .HSXOF4SR = rs("HSXOF4SR")                      ' �i�r�w�n�r�e�S����ʒu�Q��
            .HSXOF4HT = rs("HSXOF4HT")                      ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
            .HSXOF4HS = rs("HSXOF4HS")                      ' �i�r�w�n�r�e�S�ۏؕ��@�Q��
            .HSXOF4NS = rs("HSXOF4NS")                      ' �i�r�w�n�r�e�S�M�����@
            
            .HSXCNMIN = fncNullCheck(rs("HSXCNMIN"))        ' �i�r�w�Y�f�Z�x����        2003/12/10 SystemBrain Null�Ή�
            .HSXCNMAX = fncNullCheck(rs("HSXCNMAX"))        ' �i�r�w�Y�f�Z�x���        2003/12/10 SystemBrain Null�Ή�
            .HSXCNSPH = rs("HSXCNSPH")                      ' �i�r�w�Y�f�Z�x����ʒu�Q��
            .HSXCNSPT = rs("HSXCNSPT")                      ' �i�r�w�Y�f�Z�x����ʒu�Q�_
            .HSXCNSPI = rs("HSXCNSPI")                      ' �i�r�w�Y�f�Z�x����ʒu�Q��
            .HSXCNHWT = rs("HSXCNHWT")                      ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
            .HSXCNHWS = rs("HSXCNHWS")                      ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
            .HSXCNKHI = rs("HSXCNKHI")                      ' �i�r�w�Y�f�Z�x�����p�x�Q�� 09/01/08 ooba

            .HSXDENMX = fncNullCheck(rs("HSXDENMX"))        ' �i�r�w�c�������          2003/12/10 SystemBrain Null�Ή�
            .HSXDENMN = fncNullCheck(rs("HSXDENMN"))        ' �i�r�w�c��������          2003/12/10 SystemBrain Null�Ή�
            .HSXLDLMX = fncNullCheck(rs("HSXLDLMX"))        ' �i�r�w�k�^�c�k���        2003/12/10 SystemBrain Null�Ή�
            .HSXLDLMN = fncNullCheck(rs("HSXLDLMN"))        ' �i�r�w�k�^�c�k����        2003/12/10 SystemBrain Null�Ή�
            .HSXDVDMX = fncNullCheck(rs("HSXDVDMXN"))       ' �i�r�w�c�u�c�Q���   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura   2003/12/10 SystemBrain Null�Ή�
            .HSXDVDMN = fncNullCheck(rs("HSXDVDMNN"))       ' �i�r�w�c�u�c�Q����   ���ڒǉ��C�C���Ή� 2003.05.20 yakimura   2003/12/10 SystemBrain Null�Ή�
            .HSXDENHT = rs("HSXDENHT")                      ' �i�r�w�c�����ۏؕ��@�Q��
            .HSXDENHS = rs("HSXDENHS")                      ' �i�r�w�c�����ۏؕ��@�Q��
            .HSXLDLHT = rs("HSXLDLHT")                      ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
            .HSXLDLHS = rs("HSXLDLHS")                      ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
            .HSXDVDHT = rs("HSXDVDHT")                      ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
            .HSXDVDHS = rs("HSXDVDHS")                      ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
            .HSXDENKU = rs("HSXDENKU")                      ' �i�r�w�c���������L��
            .HSXDVDKU = rs("HSXDVDKU")                      ' �i�r�w�c�u�c�Q�����L��
            .HSXLDLKU = rs("HSXLDLKU")                      ' �i�r�w�k�^�c�k�����L��
        '*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݐ��ǉ�
            .HSXGDLINE = fncNullCheck(rs("HSXGDLINE"))      ' �i�r�w�k�f�c���C����
        '*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݐ��ǉ�
            .HSXLTMIN = fncNullCheck(rs("HSXLTMIN"))        ' �i�r�w�k�^�C������        2003/12/10 SystemBrain Null�Ή�
            .HSXLTMAX = fncNullCheck(rs("HSXLTMAX"))        ' �i�r�w�k�^�C�����        2003/12/10 SystemBrain Null�Ή�
''Add Start 2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
            .HSXLT10MIN = fncNullCheck(rs("LTCONVAL"))      ' �i�r�w�kLT10����
''Add End   2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
            .HSXLTSPH = rs("HSXLTSPH")                      ' �i�r�w�k�^�C������ʒu�Q��
            .HSXLTSPT = rs("HSXLTSPT")                      ' �i�r�w�k�^�C������ʒu�Q�_
            .HSXLTSPI = rs("HSXLTSPI")                      ' �i�r�w�k�^�C������ʒu�Q��
            .HSXLTHWT = rs("HSXLTHWT")                      ' �i�r�w�k�^�C���ۏؕ��@�Q��
            .HSXLTHWS = rs("HSXLTHWS")                      ' �i�r�w�k�^�C���ۏؕ��@�Q��
            
            'Null�Ή� 2003/10/22 SystemBrain ��
            .EPDUP = fncNullCheck(rs("EPDUP"))              ' EPD���                   2003/12/10 SystemBrain Null�Ή�
            .TOPREG = fncNullCheck(rs("TOPREG"))            ' TOP�K��                   2003/12/10 SystemBrain Null�Ή�
            .TAILREG = fncNullCheck(rs("TAILREG"))          ' TAIL�K��                  2003/12/10 SystemBrain Null�Ή�
            .BTMSPRT = fncNullCheck(rs("BTMSPRT"))          ' �{�g���͏o�K��            2003/12/10 SystemBrain Null�Ή�
            'Null�Ή� 2003/10/22 SystemBrain ��

'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
            If IsNull(rs("COSF3FLAG")) = False Then .COSF3FLAG = rs("COSF3FLAG") Else .COSF3FLAG = " "            'C-OSF3�׸�
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---

'--------------- 2008/08/25 INSERT START  By Systech ---------------
            .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
            .HSXLDLRMN = fncNullCheck(rs("HSXLDLRMN"))      ' �iSXL/DL�A��0����
            .HSXLDLRMX = fncNullCheck(rs("HSXLDLRMX"))      ' �iSXL/DL�A��0���
            .HWFLDLRMN = fncNullCheck(rs("HWFLDLRMN"))      ' �iWFL/DL�A��0����
            .HWFLDLRMX = fncNullCheck(rs("HWFLDLRMX"))      ' �iWFL/DL�A��0���
            If IsNull(rs("HSXOF1ARPTK")) = False Then .HSXOF1ARPTK = rs("HSXOF1ARPTK") Else .HSXOF1ARPTK = " "  ' �iSXOSF1(ArAN)�p�^���敪
            .HSXOFARMIN = fncNullCheck(rs("HSXOFARMIN"))    ' �iSXOSF(ArAN)����
            .HSXOFARMAX = fncNullCheck(rs("HSXOFARMAX"))    ' �iSXOSF(ArAN)���
            .HSXOFARMHMX = fncNullCheck(rs("HSXOFARMHMX"))  ' �iSXOSF(ArAN)�ʓ�����
            If IsNull(rs("HSXGDPTK")) = False Then .HSXGDPTK = rs("HSXGDPTK") Else .HSXGDPTK = " "  ' �i�r�w�f�c�p�^���敪
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
            If IsNull(rs("HSXOSF1PTK")) = False Then .HSXOSF1PTK = rs("HSXOSF1PTK")   ' �i�r�w�n�r�e�P�p�^���敪
            If IsNull(rs("HSXOSF2PTK")) = False Then .HSXOSF2PTK = rs("HSXOSF2PTK")   ' �i�r�w�n�r�e�Q�p�^���敪
            If IsNull(rs("HSXOSF3PTK")) = False Then .HSXOSF3PTK = rs("HSXOSF3PTK")   ' �i�r�w�n�r�e�R�p�^���敪
            If IsNull(rs("HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("HSXOSF4PTK")   ' �i�r�w�n�r�e�S�p�^���敪
            
            .HSXBMD1MBP = fncNullCheck(rs("HSXBMD1MBP"))    ' �i�r�w�a�l�c�P�ʓ����z    2003/12/10 SystemBrain Null�Ή�
            .HSXBMD2MBP = fncNullCheck(rs("HSXBMD2MBP"))    ' �i�r�w�a�l�c�Q�ʓ����z    2003/12/10 SystemBrain Null�Ή�
            .HSXBMD3MBP = fncNullCheck(rs("HSXBMD3MBP"))    ' �i�r�w�a�l�c�R�ʓ����z    2003/12/10 SystemBrain Null�Ή�
' OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
            
            'WF�d�l�擾�@08/04/15 ooba START ============================================>
            .HWFRHWYS = rs("HWFRHWYS")                      ' �i�v�e���R�ۏؕ��@�Q��
            .HWFONHWS = rs("HWFONHWS")                      ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
            .HWFOF1HS = rs("HWFOF1HS")                      ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
            .HWFOF2HS = rs("HWFOF2HS")                      ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
            .HWFOF3HS = rs("HWFOF3HS")                      ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
            .HWFOF4HS = rs("HWFOF4HS")                      ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
            .HWFBM1HS = rs("HWFBM1HS")                      ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
            .HWFBM2HS = rs("HWFBM2HS")                      ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
            .HWFBM3HS = rs("HWFBM3HS")                      ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
            .HWFDENHS = rs("HWFDENHS")                      ' �i�v�e�c�����ۏؕ��@�Q��
            .HWFDVDHS = rs("HWFDVDHS")                      ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��
            .HWFLDLHS = rs("HWFLDLHS")                      ' �i�v�e�k�^�c�k�ۏؕ��@�Q��
            .HWFRKHNN = rs("HWFRKHNN")                      ' �i�v�e���R�����p�x�Q��
            .HWFONKHN = rs("HWFONKHN")                      ' �i�v�e�_�f�Z�x�����p�x�Q��
            .HWFOF1KN = rs("HWFOF1KN")                      ' �i�v�e�n�r�e�P�����p�x�Q��
            .HWFOF2KN = rs("HWFOF2KN")                      ' �i�v�e�n�r�e�Q�����p�x�Q��
            .HWFOF3KN = rs("HWFOF3KN")                      ' �i�v�e�n�r�e�R�����p�x�Q��
            .HWFOF4KN = rs("HWFOF4KN")                      ' �i�v�e�n�r�e�S�����p�x�Q��
            .HWFBM1KN = rs("HWFBM1KN")                      ' �i�v�e�a�l�c�P�����p�x�Q��
            .HWFBM2KN = rs("HWFBM2KN")                      ' �i�v�e�a�l�c�Q�����p�x�Q��
            .HWFBM3KN = rs("HWFBM3KN")                      ' �i�v�e�a�l�c�R�����p�x�Q��
            .HWFGDKHN = rs("HWFGDKHN")                      ' �i�v�e�f�c�����p�x�Q��
            'WF�d�l�擾�@08/04/15 ooba END ==============================================>
            .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))        ' �i�r�w�ʌX�����S  2009/08/12 Kameda
            .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))        ' �i�r�w�ʌX������  2009/08/12 Kameda
            .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))        ' �i�r�w�ʌX�����  2009/08/12 Kameda
            .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))        ' �i�r�w�ʌX�����  2009/09/01 Kameda
            .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))        ' �i�r�w�ʌX������  2009/09/01 Kameda
            .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))        ' �i�r�w�ʌX�����  2009/09/01 Kameda
            .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))        ' �i�r�w�ʌX�����  2009/09/01 Kameda
            .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))        ' �i�r�w�ʌX������  2009/09/01 Kameda
            .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))        ' �i�r�w�ʌX�����  2009/09/01 Kameda
            .HWFSIRDMX = fncNullCheck(rs("HWFSIRDMX"))      ' �i�ʓ������    2010/02/04 Kameda
            
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎d�l���ڒǉ�
            If IsNull(rs("HSXCPK")) = False Then .HSXCPK = rs("HSXCPK") Else .HSXCPK = " "              ' �i�r�w�b�p�^�[���敪
            If IsNull(rs("HSXCSZ")) = False Then .HSXCSZ = rs("HSXCSZ") Else .HSXCSZ = " "              ' �i�r�w�b�������
            If IsNull(rs("HSXCHT")) = False Then .HSXCHT = rs("HSXCHT") Else .HSXCHT = " "              ' �i�r�w�b�ۏؕ��@�Q��
            If IsNull(rs("HSXCHS")) = False Then .HSXCHS = rs("HSXCHS") Else .HSXCHS = " "              ' �i�r�w�b�ۏؕ��@�Q��
            If IsNull(rs("HSXCJPK")) = False Then .HSXCJPK = rs("HSXCJPK") Else .HSXCJPK = " "          ' �i�r�w�b�i�p�^�[���敪
            If IsNull(rs("HSXCJNS")) = False Then .HSXCJNS = rs("HSXCJNS") Else .HSXCJNS = " "          ' �i�r�w�b�i�M�����@
            If IsNull(rs("HSXCJHT")) = False Then .HSXCJHT = rs("HSXCJHT") Else .HSXCJHT = " "          ' �i�r�w�b�i�ۏؕ��@�Q��
            If IsNull(rs("HSXCJHS")) = False Then .HSXCJHS = rs("HSXCJHS") Else .HSXCJHS = " "          ' �i�r�w�b�i�ۏؕ��@�Q��
            If IsNull(rs("HSXCJLTPK")) = False Then .HSXCJLTPK = rs("HSXCJLTPK") Else .HSXCJLTPK = " "  ' �i�r�w�b�i�k�s�p�^�[���敪
            If IsNull(rs("HSXCJLTNS")) = False Then .HSXCJLTNS = rs("HSXCJLTNS") Else .HSXCJLTNS = " "  ' �i�r�w�b�i�k�s�M�����@
            If IsNull(rs("HSXCJLTHT")) = False Then .HSXCJLTHT = rs("HSXCJLTHT") Else .HSXCJLTHT = " "  ' �i�r�w�b�i�k�s�ۏؕ��@�Q��
            If IsNull(rs("HSXCJLTHS")) = False Then .HSXCJLTHS = rs("HSXCJLTHS") Else .HSXCJLTHS = " "  ' �i�r�w�b�i�k�s�ۏؕ��@�Q��
            If IsNull(rs("HSXCJ2PK")) = False Then .HSXCJ2PK = rs("HSXCJ2PK") Else .HSXCJ2PK = " "      ' �i�r�w�b�i�Q�p�^�[���敪
            If IsNull(rs("HSXCJ2NS")) = False Then .HSXCJ2NS = rs("HSXCJ2NS") Else .HSXCJ2NS = " "      ' �i�r�w�b�i�Q�M�����@
            If IsNull(rs("HSXCJ2HT")) = False Then .HSXCJ2HT = rs("HSXCJ2HT") Else .HSXCJ2HT = " "      ' �i�r�w�b�i�Q�ۏؕ��@�Q��
            If IsNull(rs("HSXCJ2HS")) = False Then .HSXCJ2HS = rs("HSXCJ2HS") Else .HSXCJ2HS = " "      ' �i�r�w�b�i�Q�ۏؕ��@�Q��
            .HSXCJLTBND = fncNullCheck(rs("HSXCJLTBND"))                                                ' �iSXL/CJLT�o���h�� Number(3,0)
    
  'Add End   2011/01/17 SMPK A.Nagamine
  
  'Add Start 2012/06/01 SMPK H.Ohkubo
            If IsNull(rs("HSXCOSF3PK")) = False Then .HSXCOSF3PK = rs("HSXCOSF3PK") Else .HSXCOSF3PK = " "  '�i�r�w�b�n�r�e�R�p�^�[���敪"
  'Add End 2012/06/01 SMPK H.Ohkubo
  
        End With
        rs.MoveNext
    Next

    If scmzc_getKakouJiltuseki(inBlockID, Jiltuseki) = FUNCTION_RETURN_FAILURE Then
        getHinSiyou = FUNCTION_RETURN_FAILURE
        ReDim siyou(0)
        GoTo proc_exit
    End If
    For i = 1 To recCnt
        siyou(i).DIAMETER = (Jiltuseki.top(1) + Jiltuseki.top(2) + Jiltuseki.TAIL(1) + Jiltuseki.TAIL(2)) / 4 ' ���a
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

'�T�v      :�����֐� �T���v���ԍ����擾����
'�T�v      :�������� �e��f�[�^�擾
'���Ұ�    :�ϐ���        ,IO ,�^                                 ,����
'          :inBlockID     ,I  ,String                             ,�Ώۃu���b�NID
'          :CrySmp()      ,O  ,type_DBDRV_scmzc_fcmkc001c_CrySmp  ,�����T���v���Ǘ��擾�p
'          :iSmpGetFlg    ,I  ,Integer                            :����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
'          :iSamplID1     ,I  ,Long                               :TOP�����ID(�ȗ���)   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iSamplID2     ,I  ,Long                               :BOT�����ID(�ȗ���)   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O  ,FUNCTION_RETURN                    ,�ǂݍ��ݐ���
'����      :
'����      :2001/06/26 ���{ �쐬
Private Function getCrySmp(inBlockID As String, _
                           CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                           iSmpGetFlg As Integer, iSamplID1 As Long, iSamplID2 As Long) As FUNCTION_RETURN
    
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim recCnt      As Integer
    Dim i           As Long
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Dim wkXsdcs     As typ_XSDCS
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function getCrySmp"

    If iSmpGetFlg = 0 Then          '��ۯ�ID�Ō���(�����敪=��ۯ�)
        '��ۯ��Ǘ�(TBCME040)�Q�ƒ�~�@05/10/03 ooba
        sql = "select CS.CRYNUMCS, CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, CS.SMPKBNCS, CS.TBKBNCS, CS.REPSMPLIDCS, CS.XTALCS, CS.INPOSCS, "
        sql = sql & "CS.HINBCS, CS.REVNUMCS, CS.FACTORYCS, CS.OPECS, CS.KTKBNCS, CS.BLKKTFLAGCS, "
        sql = sql & "CS.CRYSMPLIDRSCS, CS.CRYSMPLIDRS1CS, CS.CRYSMPLIDRS2CS, CS.CRYINDRSCS, CS.CRYRESRS1CS, CS.CRYRESRS2CS, "
        sql = sql & "CS.CRYSMPLIDOICS, CS.CRYINDOICS, CS.CRYRESOICS, CS.CRYSMPLIDB1CS, CS.CRYINDB1CS, CS.CRYRESB1CS, "
        sql = sql & "CS.CRYSMPLIDB2CS, CS.CRYINDB2CS, CS.CRYRESB2CS, CS.CRYSMPLIDB3CS, CS.CRYINDB3CS, CS.CRYRESB3CS, "
        sql = sql & "CS.CRYSMPLIDL1CS, CS.CRYINDL1CS, CS.CRYRESL1CS, CS.CRYSMPLIDL2CS, CS.CRYINDL2CS, CS.CRYRESL2CS, "
        sql = sql & "CS.CRYSMPLIDL3CS, CS.CRYINDL3CS, CS.CRYRESL3CS, CS.CRYSMPLIDL4CS, CS.CRYINDL4CS, CS.CRYRESL4CS, "
        sql = sql & "CS.CRYSMPLIDCSCS, CS.CRYINDCSCS, CS.CRYRESCSCS, CS.CRYSMPLIDGDCS, CS.CRYINDGDCS, CS.CRYRESGDCS, "
        sql = sql & "CS.CRYSMPLIDTCS, CS.CRYINDTCS, CS.CRYRESTCS, CS.CRYREST10CS, CS.CRYSMPLIDEPCS, CS.CRYINDEPCS, CS.CRYRESEPCS "
      'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎d�l���ڒǉ�
        sql = sql & ", CS.CRYSMPLIDCCS, CS.CRYINDCCS, CS.CRYRESCCS, CS.CRYSMPLIDCJCS, CS.CRYINDCJCS"
        sql = sql & ", CS.CRYRESCJCS, CS.CRYSMPLIDCJLTCS, CS.CRYINDCJLTCS, CS.CRYRESCJLTCS, CS.CRYSMPLIDCJ2CS"
        sql = sql & ", CS.CRYINDCJ2CS, CS.CRYRESCJ2CS "
      'Add End   2011/01/17 SMPK A.Nagamine
        sql = sql & "from XSDCS CS, "
        sql = sql & "(select CRYNUMCS, XTALCS, INPOSCS from XSDCS "
        sql = sql & "where TBKBNCS = 'T' and CRYNUMCS = '" & inBlockID & "' "
        sql = sql & ") CSTOP, "
        sql = sql & "(select CRYNUMCS, XTALCS, INPOSCS from XSDCS "
        sql = sql & "where TBKBNCS = 'B' and CRYNUMCS = '" & inBlockID & "' "
        sql = sql & ") CSBOT "
        sql = sql & "where CSTOP.CRYNUMCS = CSBOT.CRYNUMCS and "
        
        sql = sql & "CS.CRYNUMCS = '" & inBlockID & "' and "
        sql = sql & "CS.LIVKCS = '0'"
    
    Else                            '�����ԍ��ƻ����ID�Ō���
        sql = "select CS.CRYNUMCS, 0 as LENGTH, CS.SMPKBNCS, CS.TBKBNCS, CS.REPSMPLIDCS, CS.XTALCS, CS.INPOSCS, "
        sql = sql & "CS.HINBCS, CS.REVNUMCS, CS.FACTORYCS, CS.OPECS, CS.KTKBNCS, CS.BLKKTFLAGCS, "
        sql = sql & "CS.CRYSMPLIDRSCS, CS.CRYSMPLIDRS1CS, CS.CRYSMPLIDRS2CS, CS.CRYINDRSCS, CS.CRYRESRS1CS, CS.CRYRESRS2CS, "
        sql = sql & "CS.CRYSMPLIDOICS, CS.CRYINDOICS, CS.CRYRESOICS, CS.CRYSMPLIDB1CS, CS.CRYINDB1CS, CS.CRYRESB1CS, "
        sql = sql & "CS.CRYSMPLIDB2CS, CS.CRYINDB2CS, CS.CRYRESB2CS, CS.CRYSMPLIDB3CS, CS.CRYINDB3CS, CS.CRYRESB3CS, "
        sql = sql & "CS.CRYSMPLIDL1CS, CS.CRYINDL1CS, CS.CRYRESL1CS, CS.CRYSMPLIDL2CS, CS.CRYINDL2CS, CS.CRYRESL2CS, "
        sql = sql & "CS.CRYSMPLIDL3CS, CS.CRYINDL3CS, CS.CRYRESL3CS, CS.CRYSMPLIDL4CS, CS.CRYINDL4CS, CS.CRYRESL4CS, "
        sql = sql & "CS.CRYSMPLIDCSCS, CS.CRYINDCSCS, CS.CRYRESCSCS, CS.CRYSMPLIDGDCS, CS.CRYINDGDCS, CS.CRYRESGDCS, "
        sql = sql & "CS.CRYSMPLIDTCS, CS.CRYINDTCS, CS.CRYRESTCS, CS.CRYREST10CS, CS.CRYSMPLIDEPCS, CS.CRYINDEPCS, CS.CRYRESEPCS "
      'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎d�l���ڒǉ�
        sql = sql & ", CS.CRYSMPLIDCCS, CS.CRYINDCCS, CS.CRYRESCCS, CS.CRYSMPLIDCJCS, CS.CRYINDCJCS"
        sql = sql & ", CS.CRYRESCJCS, CS.CRYSMPLIDCJLTCS, CS.CRYINDCJLTCS, CS.CRYRESCJLTCS, CS.CRYSMPLIDCJ2CS"
        sql = sql & ", CS.CRYINDCJ2CS, CS.CRYRESCJ2CS "
      'Add End   2011/01/17 SMPK A.Nagamine
        sql = sql & "from XSDCS CS "
        sql = sql & "where substr(CS.CRYNUMCS, 1, 10) = substr('" & inBlockID & "', 1, 10) and "
        sql = sql & "CS.REPSMPLIDCS in (" & iSamplID1 & ", " & iSamplID2 & ")"
    End If
    
    sql = sql & "order by CS.INPOSCS "  ' TOP TAIL��
    ' SQL���s
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        getCrySmp = FUNCTION_RETURN_FAILURE
        ReDim CrySmp(0)
        GoTo proc_exit
    End If
    
    recCnt = rs.RecordCount
    ReDim CrySmp(recCnt)
    For i = 1 To recCnt
        With CrySmp(i)
            .CRYNUMCS = rs("CRYNUMCS")          '�u���b�NID
            .Length = rs("LENGTH")              ' ����
            If IsNull(rs("SMPKBNCS")) = False Then .SMPKBNCS = rs("SMPKBNCS")                   ' �T���v���敪
            .TBKBNCS = rs("TBKBNCS")            'T/B�敪
            .REPSMPLIDCS = rs("REPSMPLIDCS")    ' ��\�T���v��ID
            
            If IsNull(rs("XTALCS")) = False Then .XTALCS = rs("XTALCS")                         ' �����ԍ�
            If IsNull(rs("INPOSCS")) = False Then .INPOSCS = rs("INPOSCS")                      ' �������ʒu
            If IsNull(rs("HINBCS")) = False Then .HINBCS = rs("HINBCS")                         ' �i��
            If IsNull(rs("REVNUMCS")) = False Then .REVNUMCS = rs("REVNUMCS")                   ' ���i�ԍ������ԍ�
            If IsNull(rs("FACTORYCS")) = False Then .FACTORYCS = rs("FACTORYCS")                ' �H��
            If IsNull(rs("OPECS")) = False Then .OPECS = rs("OPECS")                            ' ���Ə���
            If IsNull(rs("KTKBNCS")) = False Then .KTKBNCS = rs("KTKBNCS")                      ' �m��敪
            If IsNull(rs("BLKKTFLAGCS")) = False Then .BLKKTFLAGCS = rs("BLKKTFLAGCS")          ' �u���b�N�m��t���O
            If IsNull(rs("CRYSMPLIDRSCS")) = False Then .CRYSMPLIDRSCS = rs("CRYSMPLIDRSCS")    ' �T���v��ID(Rs)
            If IsNull(rs("CRYSMPLIDRS1CS")) = False Then .CRYSMPLIDRS1CS = rs("CRYSMPLIDRS1CS") ' ����T���v��ID1(Rs)
            If IsNull(rs("CRYSMPLIDRS2CS")) = False Then .CRYSMPLIDRS2CS = rs("CRYSMPLIDRS2CS") ' ����T���v��ID2(Rs)
            If IsNull(rs("CRYINDRSCS")) = False Then .CRYINDRSCS = rs("CRYINDRSCS")             ' ���FLG(Rs)
            If IsNull(rs("CRYRESRS1CS")) = False Then .CRYRESRS1CS = rs("CRYRESRS1CS")          ' ����FLG1(Rs)
            If IsNull(rs("CRYRESRS2CS")) = False Then .CRYRESRS2CS = rs("CRYRESRS2CS")          ' ����FLG2(Rs)
            If IsNull(rs("CRYSMPLIDOICS")) = False Then .CRYSMPLIDOICS = rs("CRYSMPLIDOICS")    ' �T���v��ID(Oi)
            If IsNull(rs("CRYINDOICS")) = False Then .CRYINDOICS = rs("CRYINDOICS")             ' ���FLG(Oi)
            If IsNull(rs("CRYRESOICS")) = False Then .CRYRESOICS = rs("CRYRESOICS")             ' ����FLG(Oi)
            If IsNull(rs("CRYSMPLIDB1CS")) = False Then .CRYSMPLIDB1CS = rs("CRYSMPLIDB1CS")    ' �T���v��ID(B1)
            If IsNull(rs("CRYINDB1CS")) = False Then .CRYINDB1CS = rs("CRYINDB1CS")             ' ���FLG(B1)
            If IsNull(rs("CRYRESB1CS")) = False Then .CRYRESB1CS = rs("CRYRESB1CS")             ' ����FLG(B1)
            If IsNull(rs("CRYSMPLIDB2CS")) = False Then .CRYSMPLIDB2CS = rs("CRYSMPLIDB2CS")    ' �T���v��ID(B2)
            If IsNull(rs("CRYINDB2CS")) = False Then .CRYINDB2CS = rs("CRYINDB2CS")             ' ���FLG(B2)
            If IsNull(rs("CRYRESB2CS")) = False Then .CRYRESB2CS = rs("CRYRESB2CS")             ' ����FLG(B2)
            If IsNull(rs("CRYSMPLIDB3CS")) = False Then .CRYSMPLIDB3CS = rs("CRYSMPLIDB3CS")    ' �T���v��ID(B3)
            If IsNull(rs("CRYINDB3CS")) = False Then .CRYINDB3CS = rs("CRYINDB3CS")             ' ���FLG(B3)
            If IsNull(rs("CRYRESB3CS")) = False Then .CRYRESB3CS = rs("CRYRESB3CS")             ' ����FLG(B3)
            If IsNull(rs("CRYSMPLIDL1CS")) = False Then .CRYSMPLIDL1CS = rs("CRYSMPLIDL1CS")    ' �T���v��ID(L1)
            If IsNull(rs("CRYINDL1CS")) = False Then .CRYINDL1CS = rs("CRYINDL1CS")             ' ���FLG(L1)
            If IsNull(rs("CRYRESL1CS")) = False Then .CRYRESL1CS = rs("CRYRESL1CS")             ' ����FLG(L1)
            If IsNull(rs("CRYSMPLIDL2CS")) = False Then .CRYSMPLIDL2CS = rs("CRYSMPLIDL2CS")    ' �T���v��ID(L2)
            If IsNull(rs("CRYINDL2CS")) = False Then .CRYINDL2CS = rs("CRYINDL2CS")             ' ���FLG(L2)
            If IsNull(rs("CRYRESL2CS")) = False Then .CRYRESL2CS = rs("CRYRESL2CS")             ' ����FLG(L2)
            If IsNull(rs("CRYSMPLIDL3CS")) = False Then .CRYSMPLIDL3CS = rs("CRYSMPLIDL3CS")    ' �T���v��ID(L3)
            If IsNull(rs("CRYINDL3CS")) = False Then .CRYINDL3CS = rs("CRYINDL3CS")             ' ���FLG(L3)
            If IsNull(rs("CRYRESL3CS")) = False Then .CRYRESL3CS = rs("CRYRESL3CS")             ' ����FLG(L3)
            If IsNull(rs("CRYSMPLIDL4CS")) = False Then .CRYSMPLIDL4CS = rs("CRYSMPLIDL4CS")    ' �T���v��ID(L4)
            If IsNull(rs("CRYINDL4CS")) = False Then .CRYINDL4CS = rs("CRYINDL4CS")             ' ���FLG(L4)
            If IsNull(rs("CRYRESL4CS")) = False Then .CRYRESL4CS = rs("CRYRESL4CS")             ' ����FLG(L4)
            If IsNull(rs("CRYSMPLIDCSCS")) = False Then .CRYSMPLIDCSCS = rs("CRYSMPLIDCSCS")    ' �T���v��ID(Cs)
            If IsNull(rs("CRYINDCSCS")) = False Then .CRYINDCSCS = rs("CRYINDCSCS")             ' ���FLG(Cs)
            If IsNull(rs("CRYRESCSCS")) = False Then .CRYRESCSCS = rs("CRYRESCSCS")             ' ����FLG(Cs)
            If IsNull(rs("CRYSMPLIDGDCS")) = False Then .CRYSMPLIDGDCS = rs("CRYSMPLIDGDCS")    ' �T���v��ID(GD)
            If IsNull(rs("CRYINDGDCS")) = False Then .CRYINDGDCS = rs("CRYINDGDCS")             ' ���FLG(GD)
            If IsNull(rs("CRYRESGDCS")) = False Then .CRYRESGDCS = rs("CRYRESGDCS")             ' ����FLG(GD)
            If IsNull(rs("CRYSMPLIDTCS")) = False Then .CRYSMPLIDTCS = rs("CRYSMPLIDTCS")       ' �T���v��ID(T)
            If IsNull(rs("CRYINDTCS")) = False Then .CRYINDTCS = rs("CRYINDTCS")                ' ���FLG(T)
            If IsNull(rs("CRYRESTCS")) = False Then .CRYRESTCS = rs("CRYRESTCS")                ' ����FLG(T)
''Add Start 2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
            If IsNull(rs("CRYREST10CS")) = False Then .CRYREST10CS = rs("CRYREST10CS")                ' ����FLG(T)
''Add End   2011/07/25 LT10������ǉ��Ή� T.Koi(SETsw)
            If IsNull(rs("CRYSMPLIDEPCS")) = False Then .CRYSMPLIDEPCS = rs("CRYSMPLIDEPCS")    ' �T���v��ID(EPD)
            If IsNull(rs("CRYINDEPCS")) = False Then .CRYINDEPCS = rs("CRYINDEPCS")             ' ���FLG(EPD)
            If IsNull(rs("CRYRESEPCS")) = False Then .CRYRESEPCS = rs("CRYRESEPCS")             ' ����FLG(EPD)
'--------------- 2008/08/25 INSERT START  By Systech ---------------
            ' DK���x�i���сj
            wkXsdcs.HINBCS = .HINBCS
            wkXsdcs.REVNUMCS = .REVNUMCS
            wkXsdcs.FACTORYCS = .FACTORYCS
            wkXsdcs.OPECS = .OPECS
            wkXsdcs.XTALCS = .XTALCS
            wkXsdcs.CRYSMPLIDRSCS = .CRYSMPLIDRSCS
            wkXsdcs.CRYINDRSCS = .CRYINDRSCS
            .HSXDKTMP = GetDKTmpCode(False, wkXsdcs)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
            
          'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎d�l���ڒǉ�
            If IsNull(rs("CRYSMPLIDCCS")) = False Then .CRYSMPLIDCCS = rs("CRYSMPLIDCCS")           ' �T���v��ID(C)
            If IsNull(rs("CRYINDCCS")) = False Then .CRYINDCCS = rs("CRYINDCCS")                    ' ���FLG(C)
            If IsNull(rs("CRYRESCCS")) = False Then .CRYRESCCS = rs("CRYRESCCS")                    ' ����FLG(C)
            If IsNull(rs("CRYSMPLIDCJCS")) = False Then .CRYSMPLIDCJCS = rs("CRYSMPLIDCJCS")        ' �T���v��ID(CJ)
            If IsNull(rs("CRYINDCJCS")) = False Then .CRYINDCJCS = rs("CRYINDCJCS")                 ' ���FLG(CJ)
            If IsNull(rs("CRYRESCJCS")) = False Then .CRYRESCJCS = rs("CRYRESCJCS")                 ' ����FLG(CJ)
            If IsNull(rs("CRYSMPLIDCJLTCS")) = False Then .CRYSMPLIDCJLTCS = rs("CRYSMPLIDCJLTCS")  ' �T���v��ID(CJ[LT])
            If IsNull(rs("CRYINDCJLTCS")) = False Then .CRYINDCJLTCS = rs("CRYINDCJLTCS")           ' ���FLG(CJ[LT])
            If IsNull(rs("CRYRESCJLTCS")) = False Then .CRYRESCJLTCS = rs("CRYRESCJLTCS")           ' ����FLG(CJ[LT])
            If IsNull(rs("CRYSMPLIDCJ2CS")) = False Then .CRYSMPLIDCJ2CS = rs("CRYSMPLIDCJ2CS")     ' �T���v��ID(CJ2)
            If IsNull(rs("CRYINDCJ2CS")) = False Then .CRYINDCJ2CS = rs("CRYINDCJ2CS")              ' ���FLG(CJ2)
            If IsNull(rs("CRYRESCJ2CS")) = False Then .CRYRESCJ2CS = rs("CRYRESCJ2CS")              ' ����FLG(CJ2)
          'Add End   2011/01/17 SMPK A.Nagamine
            
        End With
        rs.MoveNext
    Next
    rs.Close
    
    getCrySmp = FUNCTION_RETURN_SUCCESS
    

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getCrySmp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�����֐� ������R���ю擾�p
Private Function CryR_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                              CryR As type_DBDRV_scmzc_fcmkc001c_CryR, _
                              SuCryR As type_DBDRV_scmzc_fcmkc001c_CryR, _
                              TorB As Integer, _
                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Dim wkXsdcs     As typ_XSDCS
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

    NothingFlag = False

    ' ������R���уe�[�u������l���擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CryR_Zisseki"

    CryR_Zisseki = FUNCTION_RETURN_SUCCESS

    Set rs = Nothing

    ' ����f�[�^�̊m�F�Ɛ���f�[�^�쐬
    If (Samp.CRYINDRSCS = "3") And (Samp.KTKBNCS = "0") And (ciSmpGetFlg = 0) Then
        If (Samp.CRYRESRS1CS <> "0") And (Samp.CRYRESRS2CS <> "0") Then     ' ���茳���т���������
    
            ' ����f�[�^�쐬
            If funComputeSuitei(siyou, Samp, CryR) <> 0 Then
                NothingFlag = True
                CryR_Zisseki = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
    
        Else                                                                ' ���茳���т�����
            NothingFlag = True
            CryR_Zisseki = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    
    ' �w��(�d�l)�Ǝ���FLG���m�F
    ElseIf (Samp.CRYINDRSCS <> "0") And (Samp.CRYRESRS1CS <> "0") And (Samp.KTKBNCS <> "9") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
        '----TEST2004/10
        sql = sql & "MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, REGDATE, KSTAFFID "
        sql = sql & "from TBCMJ002 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDRSCS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ002 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDRSCS & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With CryR
                .CRYNUM = rs("CRYNUM")          ' �����ԍ�
                .POSITION = rs("POSITION")      ' �ʒu
                .SMPKBN = rs("SMPKBN")          ' �T���v���敪
                .TRANCOND = rs("TRANCOND")      ' ��������
                .TRANCNT = rs("TRANCNT")        ' ������
                .SMPLNO = rs("SMPLNO")          ' �T���v���m��
                .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
                .MEAS1 = rs("MEAS1")            ' ����l�P
                .MEAS2 = rs("MEAS2")            ' ����l�Q
                .MEAS3 = rs("MEAS3")            ' ����l�R
                .MEAS4 = rs("MEAS4")            ' ����l�S
                .MEAS5 = rs("MEAS5")            ' ����l�T
                .EFEHS = rs("EFEHS")            ' �����ΐ�
                .RRG = rs("RRG")                ' RRG
                .REGDATE = rs("REGDATE")        ' �o�^���t
                '---TEST2004/10
                .KSTAFFID = rs("KSTAFFID")
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    ' DK���x�i���сj
    wkXsdcs.XTALCS = Samp.XTALCS
    wkXsdcs.CRYSMPLIDRSCS = Samp.CRYSMPLIDRSCS
    wkXsdcs.CRYINDRSCS = "0"
    CryR.HSXDKTMP = GetDKTmpCode(False, wkXsdcs)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    CryR_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' ����f�[�^�쐬
'------------------------------------------------
'�T�v      :�w�肳�ꂽ��񂩂�A����v�Z���s�Ȃ��A������ђl���쐬����B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :Siyou         ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :�d�l���\����
'          :Samp          ,I  ,type_DBDRV_scmzc_fcmkc001c_CrySmp    :�V����يǗ�(��ۯ�)�\����
'          :CryR          ,O  ,type_DBDRV_scmzc_fcmkc001c_CryR      :RS���э\����
'          :�߂�l        ,O  ,Integer                              :����(0:����, 1:�ُ�)
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Private Function funComputeSuitei(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                                  Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                                  CryR As type_DBDRV_scmzc_fcmkc001c_CryR) As Integer
    
    Dim tSuiHin         As tFullHinban
    Dim tCryRs(2)       As type_DBDRV_scmzc_fcmkc001c_CryR          '(0)�����茳Top, (1)�����茳Bot, (2)�������
    Dim getPtrn1        As String                                   'TOP�ʒu����ݺ���
    Dim getPtrn2        As String                                   'BOT�ʒu����ݺ���

    Dim retCode         As Integer
    Dim wGetSPtrn1      As String
    Dim wGetSPtrn2      As String
    Dim wcnt            As Integer
    Dim wMeasTop(4)     As Double                   'Top����l
    Dim wMeasBot(4)     As Double                   'Bot����l
    Dim wMeasSui()      As Double                   '�Z�o����l
    Dim retJudg         As Boolean
    
    '�V����يǗ�(��ۯ�)�̕i�Ԑݒ�
    tSuiHin.hinban = Samp.HINBCS
    tSuiHin.mnorevno = Samp.REVNUMCS
    tSuiHin.factory = Samp.FACTORYCS
    tSuiHin.opecond = Samp.OPECS
    
    '�V����يǗ�(XSDCS)�̐��茳�����ID1����A���茳RS���ђl���擾����B
    If funGetCryRsJisseki(Samp.XTALCS, Samp.CRYSMPLIDRS1CS, tCryRs(0)) <> 0 Then GoTo ComputeSuiteiNG

    '�V����يǗ�(XSDCS)�̐��茳�����ID2����A���茳RS���ђl���擾����B
    If funGetCryRsJisseki(Samp.XTALCS, Samp.CRYSMPLIDRS2CS, tCryRs(1)) <> 0 Then GoTo ComputeSuiteiNG

    '������R���т̏����񐔎擾
    retCode = funGetTrancntRS(Samp)
    If retCode < 0 Then GoTo ComputeSuiteiSonotaErr

    '�����̎��уf�[�^�ҏW
    With tCryRs(2)
        .CRYNUM = Samp.XTALCS               '�����ԍ�
        .POSITION = Samp.INPOSCS            '�ʒu
        .SMPKBN = Samp.TBKBNCS              '����ً敪
        .TRANCOND = "0"                     '��������
        .TRANCNT = retCode                  '������
        .SMPLNO = Samp.CRYSMPLIDRSCS        '�����No
        .SMPLUMU = "0"                      '����ٗL��
    End With
    
    'Top/Bot����l�𐄒�l�Z�o�p�ɃZ�b�g
        wMeasTop(0) = tCryRs(0).MEAS1
        wMeasTop(1) = tCryRs(0).MEAS2
        wMeasTop(2) = tCryRs(0).MEAS3
        wMeasTop(3) = tCryRs(0).MEAS4
        wMeasTop(4) = tCryRs(0).MEAS5
    
        wMeasBot(0) = tCryRs(1).MEAS1
        wMeasBot(1) = tCryRs(1).MEAS2
        wMeasBot(2) = tCryRs(1).MEAS3
        wMeasBot(3) = tCryRs(1).MEAS4
        wMeasBot(4) = tCryRs(1).MEAS5
    
    '�����̑���_�����A����l���Z�o����
    ReDim wMeasSui(4)
    For wcnt = 0 To 4
        
        '����l�̎Z�o
        retCode = new_ResSuitei(Samp.XTALCS, wMeasTop(wcnt), tCryRs(0).POSITION, wMeasBot(wcnt), tCryRs(1).POSITION, Samp.INPOSCS, wMeasSui(wcnt))
        If retCode = FUNCTION_RETURN_FAILURE Then GoTo ComputeSuiteiNG
    
    Next wcnt
    
    '����l�̐ݒ�
    tCryRs(2).MEAS1 = wMeasSui(0)
    tCryRs(2).MEAS2 = wMeasSui(1)
    tCryRs(2).MEAS3 = wMeasSui(2)
    tCryRs(2).MEAS4 = wMeasSui(3)
    tCryRs(2).MEAS5 = wMeasSui(4)
    
    CryR = tCryRs(2)
    funComputeSuitei = 0
    Exit Function

ComputeSuiteiNG:
    funComputeSuitei = 0
    Exit Function

ComputeSuiteiSonotaErr:
    funComputeSuitei = -2
End Function

'------------------------------------------------
' ���R����p�^�[���R�[�h�擾
'------------------------------------------------
'�T�v      :�����ԍ��Ɛ��茳�����ID1�Ɛ��茳�����ID2����A�V����يǗ�(��ۯ�)(XSDCS)���������A���ꂼ��̕i�Ԃ��擾����B
'           ���茳1,���茳2,�����̕i�Ԃ�����R�d�l�l���擾���A���R��������ݺ��ނ��擾����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :sCryNum       ,I  ,String                               :�����ԍ�
'          :tSuiHin       ,I  ,tFullHinban                          :�����i��(�\����)
'          :iSmplID1      ,I  ,Integer                              :���茳�T���v���h�c�P
'          :iSmplID2      ,I  ,Integer                              :���茳�T���v���h�c�Q
'          :sHSXRSPOT     ,I  ,String                               :�����RS����_��
'          :tCryRs()      ,I  ,type_DBDRV_scmzc_fcmkc001c_CryR      :RS���� (0)�����茳Top, (1)�����茳Bot, (2)�������
'          :iGetPCode1    ,O  ,String                               :���茳�p�^�[���P('A' or 'B')
'          :iGetPCode2    ,O  ,String                               :���茳�p�^�[���Q('A' or 'B')
'          :�߂�l        ,O  ,Integer                              :�擾���� = 0 : ����I��
'                                                                               1 : ����I��(�Y���T���v���Ȃ�)
'                                                                              -1 : �ُ�I��
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Private Function funGetPcodeRS(sCryNum As String, tSuiHin As tFullHinban, iSmplID1 As Integer, iSmplID2 As Integer, _
                                                    sHSXRSPOT As String, tCryRs() As type_DBDRV_scmzc_fcmkc001c_CryR, _
                                                    iGetPCode1 As String, iGetPCode2 As String) As Integer
    
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim getNewSpec  As String       '�V����وʒu���R�d�l�l
    Dim wcnt        As Integer
    Dim getTopHin   As tFullHinban  'TOP�ʒu�i��
    Dim getTopSpec  As String       'TOP�ʒu���R�d�l�l
    Dim getTopPtrn  As String       'TOP�ʒu����ݺ���
    Dim getBotHin   As tFullHinban  'BOT�ʒu�i��
    Dim getBotSpec  As String       'BOT�ʒu���R�d�l�l
    Dim getBotPtrn  As String       'BOT�ʒu����ݺ���
    
    '-------------------- ����� --------------------
    '�e�i�Ԃ̔��R�d�l�l�擾
    '��w�肳�ꂽ�V�T���v���ʒu��
    getNewSpec = funGetSuiSpecRS(tSuiHin)
    If getNewSpec = " " Then GoTo GetPcodeRSEmpty
    
    '-------------------- ���茳�P --------------------
    '�w�肳�ꂽ�������ɁA�V����يǗ�(��ۯ�)(XSDCS)����������B
    '�ᐄ�茳�T���v���h�c�P(TOP�ʒu)�̎擾��
    sql = "select HINBCS, REVNUMCS, FACTORYCS, OPECS from XSDCS "
'' 09/03/02 FAE)akiyama start
'    sql = sql & "where XTALCS = '" & sCryNum & "' and "
    sql = sql & "where CRYNUMCS LIKE '" & left(sCryNum, 9) & "%' and "
'' 09/03/02 FAE)akiyama end
    sql = sql & "      REPSMPLIDCS = " & iSmplID1
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        Set rs = Nothing
        GoTo GetPcodeRSEmpty
    End If
    
    'TOP�ʒu�f�[�^�̐ݒ�
    getTopHin.hinban = rs("HINBCS")         'TOP�ʒu�i��
    getTopHin.mnorevno = rs("REVNUMCS")     'TOP�ʒu���i�ԍ������ԍ�
    getTopHin.factory = rs("FACTORYCS")     'TOP�ʒu�H��
    getTopHin.opecond = rs("OPECS")         'TOP�ʒu���Ə���
    Set rs = Nothing
    
    '�ᐄ�茳�T���v���h�c�P(TOP�ʒu)��
    getTopSpec = funGetSuiSpecRS(getTopHin)
    If getTopSpec <> " " Then
        '�R�[�hDB�擾�֐����Ăяo����R�[�h�e�[�u��������R����p�^�[���R�[�h���擾����
        getTopPtrn = "A"
    Else
        '�����ް�����A�������Z�o����
        wcnt = funGetRsCnt(tCryRs(0))
        If wcnt < 1 Then GoTo GetPcodeRSEmpty

        If wcnt = sHSXRSPOT Then
            getTopPtrn = "A"
        ElseIf wcnt > sHSXRSPOT Then
            getTopPtrn = "B"
        Else
            GoTo GetPcodeRSEmpty
        End If
    End If
    
    '-------------------- ���茳�Q --------------------
    '�w�肳�ꂽ�������ɁA�V����يǗ�(��ۯ�)(XSDCS)����������B
    '�ᐄ�茳�T���v���h�c�Q(BOT�ʒu)�̎擾��
    sql = "select HINBCS, REVNUMCS, FACTORYCS, OPECS from XSDCS "
'' 09/03/02 FAE)akiyama start
'    sql = sql & "where XTALCS = '" & sCryNum & "' and "
    sql = sql & "where CRYNUMCS LIKE '" & left(sCryNum, 9) & "%' and "
'' 09/03/02 FAE)akiyama end
    sql = sql & "      REPSMPLIDCS = " & iSmplID2
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        Set rs = Nothing
        GoTo GetPcodeRSEmpty
    End If
    
    'BOT�ʒu�f�[�^�̐ݒ�
    getBotHin.hinban = rs("HINBCS")         'BOT�ʒu�i��
    getBotHin.mnorevno = rs("REVNUMCS")     'BOT�ʒu���i�ԍ������ԍ�
    getBotHin.factory = rs("FACTORYCS")     'BOT�ʒu�H��
    getBotHin.opecond = rs("OPECS")         'BOT�ʒu���Ə���
    Set rs = Nothing
    
    '�ᐄ�茳�T���v���h�c�Q(BOT�ʒu)��
    getBotSpec = funGetSuiSpecRS(getBotHin)
    If getBotSpec <> " " Then
        '�R�[�hDB�擾�֐����Ăяo����R�[�h�e�[�u��������R����p�^�[���R�[�h���擾����
        getBotPtrn = "A"
    Else
        '�����ް�����A�������Z�o����
        wcnt = funGetRsCnt(tCryRs(1))
        If wcnt < 1 Then GoTo GetPcodeRSEmpty

        If wcnt = sHSXRSPOT Then
            getBotPtrn = "A"
        ElseIf wcnt > sHSXRSPOT Then
            getBotPtrn = "B"
        Else
            GoTo GetPcodeRSEmpty
        End If
    End If
    
    '�Ăяo�����ւ̌��ʒʒm
    iGetPCode1 = getTopPtrn         '���茳�p�^�[���P('A' or 'B')
    iGetPCode2 = getBotPtrn         '���茳�p�^�[���Q('A' or 'B')
    
    funGetPcodeRS = 0
    Exit Function

GetPcodeRSEmpty:
    funGetPcodeRS = 1
    Exit Function

GetPcodeRSParameterErr:
    funGetPcodeRS = -1
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
' ������R���т̏����񐔎擾
'------------------------------------------------
'�T�v      :������R����(TBCMJ002)����Y������f�[�^�̏����񐔂��擾����B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :Samp          ,I  ,type_DBDRV_scmzc_fcmkc001c_CrySmp    :�V����يǗ�(��ۯ�)�\����
'          :�߂�l        ,O  ,Integer                              :������(�ő�l�{�P)
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Private Function funGetTrancntRS(Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp) As Integer
    
    Dim sql         As String
    Dim rs          As OraDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function funGetTrancntRS"

    Set rs = Nothing

    ' ������R���уe�[�u������l���擾
    sql = "select TRANCNT+1 MAXCNT from TBCMJ002 "
    sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
    sql = sql & "      SMPLNO = " & Samp.REPSMPLIDCS & " and "
    sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ002 "
    sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
    sql = sql & "                 SMPLNO = " & Samp.REPSMPLIDCS & ")"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs.EOF Or rs.RecordCount = 0 Then
        funGetTrancntRS = 1
    Else
        funGetTrancntRS = rs("MAXCNT")
    End If
    Set rs = Nothing

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    funGetTrancntRS = -1
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�����֐� Oi���ю擾�p
Private Function Oi_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                            Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                            Oi As type_DBDRV_scmzc_fcmkc001c_Oi, _
                            TorB As Integer, _
                            Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean

    NothingFlag = False

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function Oi_Zisseki"

    Oi_Zisseki = FUNCTION_RETURN_SUCCESS

    ' �w��(�d�l)�Ǝ���FLG���m�F
    If (Samp.CRYINDOICS <> "0") And (Samp.CRYRESOICS <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
        sql = sql & "OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, AVE, FTIRCONV, INSPECTWAY, REGDATE "
        sql = sql & "from TBCMJ003 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDOICS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ003 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDOICS & ")"
        sql = sql & "  and TRANCOND = 0 "       'GFA��FTIR���Z�l�\���ُ�Ή� 2011/01/20�ǉ� SETsw kubota
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Oi
                .CRYNUM = rs("CRYNUM")          ' �����ԍ�
                .POSITION = rs("POSITION")      ' �ʒu
                .SMPKBN = rs("SMPKBN")          ' �T���v���敪
                .TRANCOND = rs("TRANCOND")      ' ��������
                .TRANCNT = rs("TRANCNT")        ' ������
                .SMPLNO = rs("SMPLNO")          ' �T���v���m��
                .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
                If IsNull(rs("OIMEAS1")) = False Then .OIMEAS1 = rs("OIMEAS1") Else .OIMEAS1 = -1  '�n������l1
                If IsNull(rs("OIMEAS2")) = False Then .OIMEAS2 = rs("OIMEAS2") Else .OIMEAS2 = -1  '�n������l2
                If IsNull(rs("OIMEAS3")) = False Then .OIMEAS3 = rs("OIMEAS3") Else .OIMEAS3 = -1  '�n������l3
                If IsNull(rs("OIMEAS4")) = False Then .OIMEAS4 = rs("OIMEAS4") Else .OIMEAS4 = -1  '�n������l4
                If IsNull(rs("OIMEAS5")) = False Then .OIMEAS5 = rs("OIMEAS5") Else .OIMEAS5 = -1  '�n������l5
                If IsNull(rs("ORGRES")) = False Then .ORGRES = rs("ORGRES") Else .ORGRES = -1    ' �n�q�f����
'OI_NULL�Ή��@2005/03/08 TUKU END   --------------------------------------------------
                .AVE = rs("AVE")                ' �`�u�d
                .FTIRCONV = rs("FTIRCONV")      ' �e�s�h�q���Z
                .INSPECTWAY = rs("INSPECTWAY")  ' �������@
                .REGDATE = rs("REGDATE")        ' �o�^���t
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Oi_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�����֐� BMD���ю擾�p
Private Function BMD_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             inTRANCOND As Integer, _
                             BMD As type_DBDRV_scmzc_fcmkc001c_BMD, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wHSX_HS     As String
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long         'Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    
    NothingFlag = False

    ' BMD���уe�[�u������l���擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function BMD_Zisseki"

    BMD_Zisseki = FUNCTION_RETURN_SUCCESS

    If inTRANCOND = 1 Then
        wHSX_HS = siyou.HSXBM1HS
        wCryIND = Samp.CRYINDB1CS
        wCryRES = Samp.CRYRESB1CS
        wCrySMPL = Samp.CRYSMPLIDB1CS
    ElseIf inTRANCOND = 2 Then
        wHSX_HS = siyou.HSXBM2HS
        wCryIND = Samp.CRYINDB2CS
        wCryRES = Samp.CRYRESB2CS
        wCrySMPL = Samp.CRYSMPLIDB2CS
    ElseIf inTRANCOND = 3 Then
        wHSX_HS = siyou.HSXBM3HS
        wCryIND = Samp.CRYINDB3CS
        wCryRES = Samp.CRYRESB3CS
        wCrySMPL = Samp.CRYSMPLIDB3CS
    End If

    ' �w��(�d�l)�Ǝ���FLG���m�F
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HTPRC, KKSP, KKSET, "
        sql = sql & "MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN, MEASMAX, MEASAVE, BMDMNBUNP, REGDATE "
        sql = sql & "from TBCMJ008 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & wCrySMPL & " and "
        sql = sql & "      TRANCOND = '" & inTRANCOND & "' and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ008 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & wCrySMPL & " and "
        sql = sql & "                       TRANCOND = '" & inTRANCOND & "')"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With BMD
                .CRYNUM = rs("CRYNUM")          ' �����ԍ�
                .POSITION = rs("POSITION")      ' �ʒu
                .SMPKBN = rs("SMPKBN")          ' �T���v���敪
                .TRANCOND = rs("TRANCOND")      ' ��������
                .TRANCNT = rs("TRANCNT")        ' ������
                .SMPLNO = rs("SMPLNO")          ' �T���v���m��
                .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
                .HTPRC = rs("HTPRC")            ' �M�������@
                .KKSP = rs("KKSP")              ' �������ב���ʒu
                .KKSET = rs("KKSET")            ' �������ב�������{�I��ET��
                .MEAS1 = rs("MEAS1")            ' ����l�P
                .MEAS2 = rs("MEAS2")            ' ����l�Q
                .MEAS3 = rs("MEAS3")            ' ����l�R
                .MEAS4 = rs("MEAS4")            ' ����l�S
                .MEAS5 = rs("MEAS5")            ' ����l�T
                .MEASMIN = rs("MEASMIN")        ' MIN
                .MEASMAX = rs("MEASMAX")        ' MAX
                .MEASAVE = rs("MEASAVE")        ' AVE
                 If IsNull(rs("BMDMNBUNP")) = False Then .BMDMNBUNP = rs("BMDMNBUNP")       ' BMD�ʓ����z
                .REGDATE = rs("REGDATE")        ' �o�^���t
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    BMD_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�����֐� GD���ю擾�p
Private Function OSF_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             inTRANCOND As Integer, _
                             OSF As type_DBDRV_scmzc_fcmkc001c_OSF, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wHSX_HS     As String
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long     'Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota

    NothingFlag = False

    ' OSF���уe�[�u������l���擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function OSF_Zisseki"

    OSF_Zisseki = FUNCTION_RETURN_SUCCESS

    If inTRANCOND = 1 Then
        wHSX_HS = siyou.HSXOF1HS
        wCryIND = Samp.CRYINDL1CS
        wCryRES = Samp.CRYRESL1CS
        wCrySMPL = Samp.CRYSMPLIDL1CS
    ElseIf inTRANCOND = 2 Then
        wHSX_HS = siyou.HSXOF2HS
        wCryIND = Samp.CRYINDL2CS
        wCryRES = Samp.CRYRESL2CS
        wCrySMPL = Samp.CRYSMPLIDL2CS
    ElseIf inTRANCOND = 3 Then
        wHSX_HS = siyou.HSXOF3HS
        wCryIND = Samp.CRYINDL3CS
        wCryRES = Samp.CRYRESL3CS
        wCrySMPL = Samp.CRYSMPLIDL3CS
    ElseIf inTRANCOND = 4 Then
        wHSX_HS = siyou.HSXOF4HS
        wCryIND = Samp.CRYINDL4CS
        wCryRES = Samp.CRYRESL4CS
        wCrySMPL = Samp.CRYSMPLIDL4CS
    End If

    ' �w��(�d�l)�Ǝ���FLG���m�F
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HTPRC, KKSP, KKSET, CALCMAX, CALCAVE, "
        sql = sql & "MEAS1, MEAS2,  MEAS3,  MEAS4,  MEAS5,  MEAS6,  MEAS7,  MEAS8,  MEAS9,  MEAS10, "
        sql = sql & "MEAS11,MEAS12, MEAS13, MEAS14, MEAS15, MEAS16, MEAS17, MEAS18, MEAS19, MEAS20, "
        sql = sql & "OSFPOS1, OSFWID1, OSFRD1, OSFPOS2, OSFWID2, OSFRD2, OSFPOS3, OSFWID3, OSFRD3, REGDATE "
        
        sql = sql & ",CALCMH "  '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
        'Add Start 2012/06/01 SMPK H.Ohkubo
        sql = sql & ",COSF3PTNJSK "
        'Add End 2012/06/01 SMPK H.Ohkubo
        sql = sql & "from TBCMJ005 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & wCrySMPL & " and "
        sql = sql & "      TRANCOND = '" & inTRANCOND & "' and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ005 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & wCrySMPL & " and "
        sql = sql & "                       TRANCOND = '" & inTRANCOND & "')"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With OSF
                .CRYNUM = rs("CRYNUM")          ' �����ԍ�
                .POSITION = rs("POSITION")      ' �ʒu
                .SMPKBN = rs("SMPKBN")          ' �T���v���敪
                .TRANCOND = rs("TRANCOND")      ' ��������
                .TRANCNT = rs("TRANCNT")        ' ������
                .SMPLNO = rs("SMPLNO")          ' �T���v���m��
                .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
                .HTPRC = rs("HTPRC")            ' �M�������@
                .KKSP = rs("KKSP")              ' �������ב���ʒu
                .KKSET = rs("KKSET")            ' �������ב�������{�I��ET��
                .CALCMAX = rs("CALCMAX")       ' �v�Z���� Max
                .CALCAVE = rs("CALCAVE")       ' �v�Z���� Ave
                .MEAS1 = rs("MEAS1")           ' ����l�P
                .MEAS2 = rs("MEAS2")           ' ����l�Q
                .MEAS3 = rs("MEAS3")           ' ����l�R
                .MEAS4 = rs("MEAS4")           ' ����l�S
                .MEAS5 = rs("MEAS5")           ' ����l�T
                .MEAS6 = rs("MEAS6")           ' ����l�U
                .MEAS7 = rs("MEAS7")           ' ����l�V
                .MEAS8 = rs("MEAS8")           ' ����l�W
                .MEAS9 = rs("MEAS9")           ' ����l�X
                .MEAS10 = rs("MEAS10")         ' ����l�P�O
                .MEAS11 = rs("MEAS11")         ' ����l�P�P
                .MEAS12 = rs("MEAS12")         ' ����l�P�Q
                .MEAS13 = rs("MEAS13")         ' ����l�P�R
                .MEAS14 = rs("MEAS14")         ' ����l�P�S
                .MEAS15 = rs("MEAS15")         ' ����l�P�T
                .MEAS16 = rs("MEAS16")         ' ����l�P�U
                .MEAS17 = rs("MEAS17")         ' ����l�P�V
                .MEAS18 = rs("MEAS18")         ' ����l�P�W
                .MEAS19 = rs("MEAS19")         ' ����l�P�X
                .MEAS20 = rs("MEAS20")         ' ����l�Q�O
                 If IsNull(rs("OSFPOS1")) = False Then .OSFPOS1 = rs("OSFPOS1")   '����݋敪�P�ʒu
                 If IsNull(rs("OSFWID1")) = False Then .OSFWID1 = rs("OSFWID1")   '����݋敪�P��
                 If IsNull(rs("OSFRD1")) = False Then .OSFRD1 = rs("OSFRD1")      '����݋敪�PR/D
                 If IsNull(rs("OSFPOS2")) = False Then .OSFPOS2 = rs("OSFPOS2")   '����݋敪�Q�ʒu
                 If IsNull(rs("OSFWID2")) = False Then .OSFWID2 = rs("OSFWID2")   '����݋敪�Q��
                 If IsNull(rs("OSFRD2")) = False Then .OSFRD2 = rs("OSFRD2")      '����݋敪�QR/D
                 If IsNull(rs("OSFPOS3")) = False Then .OSFPOS3 = rs("OSFPOS3")   '����݋敪�R�ʒu
                 If IsNull(rs("OSFWID3")) = False Then .OSFWID3 = rs("OSFWID3")   '����݋敪�R��
                 If IsNull(rs("OSFRD3")) = False Then .OSFRD3 = rs("OSFRD3")      '����݋敪�RR/D
                 If IsNull(rs("CALCMH")) = False Then .CALCMH = rs("CALCMH")      '�ʓ���(MAX/MIN)  '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
                .REGDATE = rs("REGDATE")       ' �o�^���t
                
                'Add Start 2012/06/01 SMPK H.Ohkubo
                If Not IsNull(rs("COSF3PTNJSK")) Then
                    .COSF3PTNJSK = rs("COSF3PTNJSK")   ' �p�^�[���敪����
                Else
                    '�p�^�[�����тȂ�
                    .COSF3PTNJSK = "0"
                End If
                'Add End 2012/06/01 SMPK H.Ohkubo
                
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    OSF_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�����֐� Cs���ю擾�p
Private Function CS_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                            Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                            Cs As type_DBDRV_scmzc_fcmkc001c_CS, _
                            TorB As Integer, _
                            Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean

    NothingFlag = False

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CS_Zisseki"
    
    CS_Zisseki = FUNCTION_RETURN_SUCCESS

    ' �w��(�d�l)�Ǝ���FLG���m�F
    If (Samp.CRYINDCSCS <> "0") And (Samp.CRYRESCSCS <> "0") Then

        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
        sql = sql & "CSMEAS, PRE70P, INSPECTWAY, REGDATE "
        sql = sql & "from TBCMJ004 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDCSCS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ004 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDCSCS & ")"

        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Cs
                .CRYNUM = rs("CRYNUM")          ' �����ԍ�
                .POSITION = rs("POSITION")      ' �ʒu
                .SMPKBN = rs("SMPKBN")          ' �T���v���敪
                .TRANCOND = rs("TRANCOND")      ' ��������
                .TRANCNT = rs("TRANCNT")        ' ������
                .SMPLNO = rs("SMPLNO")          ' �T���v���m��
                .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
                If IsNull(rs("CSMEAS")) = False Then .CSMEAS = rs("CSMEAS") Else .CSMEAS = -1  ' Cs�����l
                If IsNull(rs("PRE70P")) = False Then .PRE70P = rs("PRE70P") Else .PRE70P = -1  ' �V�O������l
'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
                .INSPECTWAY = rs("INSPECTWAY")  ' �������@
                .REGDATE = rs("REGDATE")        ' �o�^���t
            End With
        Else
            NothingFlag = True
        End If

        Set rs = Nothing
    End If
    
    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    gErr.HandleError
    Resume proc_exit
End Function

'�����֐� GD���ю擾�p
Private Function GD_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                            Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                            GD As type_DBDRV_scmzc_fcmkc001c_GD, _
                            TorB As Integer, _
                            Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    
    NothingFlag = False

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function GD_Zisseki"

    GD_Zisseki = FUNCTION_RETURN_SUCCESS

    ' �w��(�d�l)�Ǝ���FLG���m�F
    If (Samp.CRYINDGDCS <> "0") And (Samp.CRYRESGDCS <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MSRSDEN, MSRSLDL, MSRSDVD2, "
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
        
        sql = sql & ",MSZEROMN, MSZEROMX "  '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
        
        sql = sql & "from TBCMJ006 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDGDCS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ006 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDGDCS & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With GD
                .CRYNUM = rs("CRYNUM")          ' �����ԍ�
                .POSITION = rs("POSITION")      ' �ʒu
                .SMPKBN = rs("SMPKBN")          ' �T���v���敪
                .TRANCOND = rs("TRANCOND")      ' ��������
                .TRANCNT = rs("TRANCNT")        ' ������
                .SMPLNO = rs("SMPLNO")          ' �T���v���m��
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
                If IsNull(rs("MS01DVD2")) = False Then .MS01DVD2 = rs("MS01DVD2")   '����l01 DVD2
                If IsNull(rs("MS02DVD2")) = False Then .MS02DVD2 = rs("MS02DVD2")   '����l02 DVD2
                If IsNull(rs("MS03DVD2")) = False Then .MS03DVD2 = rs("MS03DVD2")   '����l03 DVD2
                If IsNull(rs("MS04DVD2")) = False Then .MS04DVD2 = rs("MS04DVD2")   '����l04 DVD2
                If IsNull(rs("MS05DVD2")) = False Then .MS05DVD2 = rs("MS05DVD2")   '����l05 DVD2
                
                If IsNull(rs("MSZEROMN")) = False Then .MSZEROMN = rs("MSZEROMN")   'L/DL0�A�����ŏ��l  '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
                If IsNull(rs("MSZEROMX")) = False Then .MSZEROMX = rs("MSZEROMX")   'L/DL0�A�����ő�l  '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
                
                .REGDATE = rs("REGDATE")        ' �o�^���t
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    GD_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�����֐� ���C�t�^�C�����ю擾�p
Private Function LT_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                            Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                            Lt As type_DBDRV_scmzc_fcmkc001c_LT, _
                            TorB As Integer, _
                            Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    
    NothingFlag = False

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function LT_Zisseki"

    ' ���C�t�^�C�����уe�[�u������l���擾
    LT_Zisseki = FUNCTION_RETURN_SUCCESS

    ' �w��(�d�l)�Ǝ���FLG���m�F
    If (Samp.CRYINDTCS <> "0") And (Samp.CRYRESTCS <> "0") Then
        
        '2005/12/02 mod SET���� ����l�P�`�T�J����NULL���ɂ�NVL�g�p ->
        '                    ����l�U�`�P�O�J�����ǉ�
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MEASPEAK, CALCMEAS, REGDATE, "
        sql = sql & "NVL(MEAS1, -1) MEAS1, "
        sql = sql & "NVL(MEAS2, -1) MEAS2, "
        sql = sql & "NVL(MEAS3, -1) MEAS3, "
        sql = sql & "NVL(MEAS4, -1) MEAS4, "
        sql = sql & "NVL(MEAS5, -1) MEAS5, "
        sql = sql & " NVL(MEAS6,-1) MEAS6, "
        sql = sql & " NVL(MEAS7,-1) MEAS7, "
        sql = sql & " NVL(MEAS8,-1) MEAS8, "
        sql = sql & " NVL(MEAS9,-1) MEAS9, "
        sql = sql & " NVL(MEAS10,-1) MEAS10, "
        sql = sql & " LTSPIFLG "
        sql = sql & ",NVL(CONVAL,-1) CONVAL "
        sql = sql & "from TBCMJ007 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDTCS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ007 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDTCS & ")"
        
        '2005/12/02 mod SET���� ����l�P�`�T�J����NULL���ɂ�NVL�g�p
        '                    ����l�U�`�P�O�J�����ǉ�               <-
        Set rs = OraDB.CreateDynaset(sql, ORADYN_READONLY)
        If rs.RecordCount > 0 Then
            With Lt
                .CRYNUM = rs("CRYNUM")          ' �����ԍ�
                .POSITION = rs("POSITION")      ' �ʒu
                .SMPKBN = rs("SMPKBN")          ' �T���v���敪
                .TRANCOND = rs("TRANCOND")      ' ��������
                .TRANCNT = rs("TRANCNT")        ' ������
                .SMPLNO = rs("SMPLNO")          ' �T���v���m��
                .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
                .MEAS1 = rs("MEAS1")            ' ����l�P
                .MEAS2 = rs("MEAS2")            ' ����l�Q
                .MEAS3 = rs("MEAS3")            ' ����l�R
                .MEAS4 = rs("MEAS4")            ' ����l�S
                .MEAS5 = rs("MEAS5")            ' ����l�T
                .MEASPEAK = rs("MEASPEAK")      ' ����l �s�[�N�l
                .CALCMEAS = rs("CALCMEAS")      ' �v�Z����
                .REGDATE = rs("REGDATE")        ' �o�^���t
''Add Start 2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
                .CONVAL = rs("CONVAL")          ' 10�����Z�l
''Add End   2011/07/22 LT10������ǉ��Ή� T.Koi(SETsw)
                '2005/12/02 add SET���� ����l�U�`�P�O�J�����ǉ��̂��ߒǉ� ->
                .MEAS6 = rs("MEAS6")            ' ����l�U
                .MEAS7 = rs("MEAS7")            ' ����l�V
                .MEAS8 = rs("MEAS8")            ' ����l�W
                .MEAS9 = rs("MEAS9")            ' ����l�X
                .MEAS10 = rs("MEAS10")          ' ����l�P�O
                .LTSPIFLG = Trim(CStr(NulltoStr(rs.Fields("LTSPIFLG").Value)))  '����ʒu����t���O
                '2005/12/02 add SET���� ����l�U�`�P�O�J�����ǉ��̂��ߒǉ� <-
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    LT_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�����֐� EPD���ю擾�p
Private Function EPD_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             EPD As type_DBDRV_scmzc_fcmkc001c_EPD, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean

    NothingFlag = False

    ' EPD���уe�[�u������l���擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function EPD_Zisseki"

    EPD_Zisseki = FUNCTION_RETURN_SUCCESS

    ' �w��(�d�l)�Ǝ���FLG���m�F
    If (Samp.CRYINDEPCS <> "0") And (Samp.CRYRESEPCS <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MEASURE, REGDATE "
        sql = sql & "from TBCMJ001 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDEPCS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ001 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDEPCS & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With EPD
                .CRYNUM = rs("CRYNUM")          ' �����ԍ�
                .POSITION = rs("POSITION")      ' �ʒu
                .SMPKBN = rs("SMPKBN")          ' �T���v���敪
                .TRANCOND = rs("TRANCOND")      ' ��������
                .TRANCNT = rs("TRANCNT")        ' ������
                .SMPLNO = rs("SMPLNO")          ' �T���v���m��
                .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
                .MEASURE = rs("MEASURE")        ' ����l
                .REGDATE = rs("REGDATE")        ' �o�^���t
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    EPD_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'�T�v      :�����֐� X�����ю擾�p    2009/08/12 Kameda
Private Function X_Zisseki(XTALCS As String, x As type_DBDRV_scmzc_fcmkc001c_X, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean

    NothingFlag = False

    ' EPD���уe�[�u������l���擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function X_Zisseki"

    X_Zisseki = FUNCTION_RETURN_SUCCESS

        
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, XRAYX,XRAYY,XRAYXY, REGDATE "
    sql = sql & "from TBCMJ021 "
    sql = sql & "where CRYNUM = '" & XTALCS & "' and "
    'sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDEPCS & " and "
    sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ021 "
    'sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
    'sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDEPCS & ")"
    sql = sql & "                 where CRYNUM = '" & XTALCS & "' )"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount <> 0 Then
        With x
            .CRYNUM = rs("CRYNUM")          ' �����ԍ�
            .POSITION = rs("POSITION")      ' �ʒu
            .SMPKBN = rs("SMPKBN")          ' �T���v���敪
            .TRANCOND = rs("TRANCOND")      ' ��������
            .TRANCNT = rs("TRANCNT")        ' ������
            .SMPLNO = rs("SMPLNO")          ' �T���v���m��
            .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
            .XX = rs("XRAYX")               ' ����lX
            .XY = rs("XRAYY")               ' ����lY
            .XXY = rs("XRAYXY")             ' ����lXY
            .REGDATE = rs("REGDATE")        ' �o�^���t
        End With
    Else
        NothingFlag = True
    End If
    
    Set rs = Nothing
    
    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    X_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'�T�v      :�����֐� SIRD���ю擾�p    2010/02/04 Kameda
Private Function SIRD_Zisseki(Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, SIRD As type_DBDRV_scmzc_fcmkc001c_SIRD, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean

    NothingFlag = False
    SIRD.NothingFlg = ""      '2010/02/18 Kameda
    ' SIRD���уe�[�u������l���擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function SIRD_Zisseki"

    SIRD_Zisseki = FUNCTION_RETURN_SUCCESS

    If Samp.SIRDKBNY3 = "1" Then
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, SIRDCNT, REGDATE "
        sql = sql & "from TBCMJ022 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      TRANCNT = '0'"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
        If rs.RecordCount <> 0 Then
            With SIRD
                .CRYNUM = rs("CRYNUM")          ' �����ԍ�
                .POSITION = rs("POSITION")      ' �ʒu
                .SMPKBN = rs("SMPKBN")          ' �T���v���敪
                .TRANCOND = rs("TRANCOND")      ' ��������
                .TRANCNT = rs("TRANCNT")        ' ������
                .SMPLNO = rs("SMPLNO")          ' �T���v���m��
                .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
                .SIRDCNT = rs("SIRDCNT")        ' ����l
                .REGDATE = rs("REGDATE")        ' �o�^���t
            End With
        Else
            NothingFlag = True
            SIRD.NothingFlg = "1"    '2010/02/18 Kameda
        End If
        
        Set rs = Nothing
    End If
    
    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    SIRD_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :���H���є���ɍ\���̂ɒl���Z�b�g����
'���Ұ�    :�ϐ���        ,IO ,�^             ,����
'          :BLOCKID       ,   ,String         ,�u���b�NID
'          :Kakou         ,   ,type_KakouJudg ,���H���є���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :�u���b�N���S�i�Ԃ̎d�l�Ǝ��т����߂�
'����      :2002/4/16 ���� �쐬
Public Function DBDRV_scmzc_fcmkc001c_Kakou(BLOCKID As String, Kakou As type_KakouJudg) As FUNCTION_RETURN
    Dim sql     As String
    Dim sql1    As String
    Dim rs      As OraDynaset
    Dim recCnt  As Integer
    Dim c0      As Integer
    Dim tHIN()  As tFullHinban

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_Kakou"

    DBDRV_scmzc_fcmkc001c_Kakou = FUNCTION_RETURN_FAILURE

    '�u���b�N���̑S�i�Ԃ����߂�
    '��ۯ��Ǘ�(TBCME040)�Q�ƒ�~�@05/10/03 ooba START ======================================>
    sql = "select HINBAN, REVNUM, FACTORY, OPECOND from XSDC2 C2, TBCME041 E41 "
    sql = sql & "Where E41.CRYNUM = C2.XTALC2 and "
    sql = sql & "C2.CRYNUMC2 = '" & BLOCKID & "' and "
    sql = sql & "C2.INPOSC2 < E41.INGOTPOS+E41.LENGTH and "
    sql = sql & "C2.INPOSC2+C2.GNLC2 > E41.INGOTPOS"
    '��ۯ��Ǘ�(TBCME040)�Q�ƒ�~�@05/10/03 ooba END ========================================>

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim tHIN(recCnt)
    If recCnt = 0 Then
        rs.Close
        GoTo proc_exit
    End If
    For c0 = 1 To recCnt
        tHIN(c0).hinban = rs("HINBAN")
        tHIN(c0).mnorevno = rs("REVNUM")
        tHIN(c0).factory = rs("FACTORY")
        tHIN(c0).opecond = rs("OPECOND")
        rs.MoveNext
    Next
    rs.Close
    
    '���߂��S�i�Ԃ̉��H�d�l�����߂�
    If scmzc_getKakouSpec(tHIN(), Kakou.Spec()) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    
    '�Ώۃu���b�N�̉��H���т����߂�
    If scmzc_getKakouJiltuseki(BLOCKID, Kakou.Jiltuseki) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    
    DBDRV_scmzc_fcmkc001c_Kakou = FUNCTION_RETURN_SUCCESS

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
'�T�v      :�ʌX������X��������ԁA���уt���O�擾
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :CrySmp        ,IO  ,Double       ,
'����      :2009/08/12
Private Function GetXSDC1_XRAY(CrySmp As type_DBDRV_scmzc_fcmkc001c_CrySmp) As FUNCTION_RETURN
    Dim sql             As String
    Dim rs              As OraDynaset
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function GetXSDC1_XRAY"

    GetXSDC1_XRAY = FUNCTION_RETURN_FAILURE
    
    sql = "select "
    sql = sql & "NVL(CRYINDXC1,'0') as CRYINDXC1 "         ' ���FLG(X��)
    sql = sql & ",NVL(CRYRESXC1,'0') as CRYRESXC1 "        ' ����FLG(X��)
    sql = sql & " from XSDC1"
    sql = sql & " where XTALC1 = '" & CrySmp.XTALCS & "'"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount <> 0 Then
        CrySmp.CRYINDXC1 = rs("CRYINDXC1")
        CrySmp.CRYRESXC1 = rs("CRYRESXC1")
    End If
    
    rs.Close

    GetXSDC1_XRAY = FUNCTION_RETURN_SUCCESS
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
'�T�v      :SIRD�]���敪�擾
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :CrySmp        ,IO  ,Double       ,
'����      :2010/02/04
Private Function GetXODY3_SIRD(CrySmp As type_DBDRV_scmzc_fcmkc001c_CrySmp) As FUNCTION_RETURN
    Dim sql             As String
    Dim rs              As OraDynaset
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function GetXODY3_SIRD"

    GetXODY3_SIRD = FUNCTION_RETURN_FAILURE
    
    sql = "select "
    sql = sql & "NVL(SIRDKBNY3,'0') as SIRDKBNY3 "         '
    sql = sql & " from XODY3"
    sql = sql & " where XTALNOY3 = '" & CrySmp.CRYNUMCS & "'"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount <> 0 Then
        CrySmp.SIRDKBNY3 = rs("SIRDKBNY3")
    End If
    
    rs.Close

    GetXODY3_SIRD = FUNCTION_RETURN_SUCCESS
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

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎��я��擾
'�T�v      :�����֐� Cu-Deco C ���ю擾�p
Private Function CuDeco_C_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             Cu_deco_C As type_DBDRV_scmzc_fcmkc001c_C, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long

    NothingFlag = False

    ' Cu_deco���уe�[�u��(TBCMJ023)����l���擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CuDeco_C_Zisseki"

    CuDeco_C_Zisseki = FUNCTION_RETURN_SUCCESS

    wCryIND = Samp.CRYINDCCS        ' ��ԃt���O C
    wCryRES = Samp.CRYRESCCS        ' ���уt���O C
    wCrySMPL = Samp.CRYSMPLIDCCS    ' �T���v��ID C

    ' �w��(�d�l)�Ǝ���FLG���m�F
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO"
        sql = sql & ", SMPLUMUC, REGDATEC"
        sql = sql & ", CPTNJSK, CDISKJSK, CRINGNKJSK, CRINGGKJSK, CHANTEI"
        
        sql = sql & " from TBCMJ023"
        sql = sql & " where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "       SMPLNO = " & wCrySMPL & " and"
        sql = sql & "       TRANCNT = (select max(TRANCNT) from TBCMJ023"
        sql = sql & "                  where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "                        SMPLNO = " & wCrySMPL & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Cu_deco_C
                
                If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")             ' �����ԍ�
                .POSITION = CInt(fncNullCheck(rs("POSITION")))                          ' �ʒu
                If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")             ' �T���v���敪
                .TRANCNT = CInt(fncNullCheck(rs("TRANCNT")))                            ' ������
                .SMPLNO = CLng(fncNullCheck(rs("SMPLNO")))                              ' �T���v���m��
                If IsNull(rs("SMPLUMUC")) = False Then .SMPLUMUC = rs("SMPLUMUC")       ' �T���v���L�� C
                
                If IsNull(rs("CPTNJSK")) = False Then .CPTNJSK = rs("CPTNJSK")          ' C �p�^�[������
                
                .CDISKJSK = CInt(fncNullCheck(rs("CDISKJSK")))                          ' C Disk���a����
                .CRINGNKJSK = CInt(fncNullCheck(rs("CRINGNKJSK")))                      ' C Ring���a����
                .CRINGGKJSK = CInt(fncNullCheck(rs("CRINGGKJSK")))                      ' C Ring�O�a����
                
                If IsNull(rs("CHANTEI")) = False Then .CHANTEI = rs("CHANTEI")          ' C ���茋��
                
                If IsNull(rs("REGDATEC")) = False Then .REGDATE = rs("REGDATEC")        ' �o�^���t
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    CuDeco_C_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add End   2011/01/17 SMPK A.Nagamine


'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎��я��擾
'�T�v      :�����֐� Cu-Deco CJ ���ю擾�p
Private Function CuDeco_CJ_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             Cu_deco_CJ As type_DBDRV_scmzc_fcmkc001c_CJ, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long

    NothingFlag = False

    ' Cu_deco���уe�[�u��(TBCMJ023)����l���擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CuDeco_CJ_Zisseki"

    CuDeco_CJ_Zisseki = FUNCTION_RETURN_SUCCESS

    wCryIND = Samp.CRYINDCJCS           ' ��ԃt���O CJ
    wCryRES = Samp.CRYRESCJCS           ' ���уt���O CJ
    wCrySMPL = Samp.CRYSMPLIDCJCS       ' �T���v��ID CJ

    ' �w��(�d�l)�Ǝ���FLG���m�F
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO"
        sql = sql & ", SMPLUMUCJ, REGDATECJ"
        sql = sql & ", CJPTNJSK, CJDISKJSK, CJRINGNKJSK, CJRINGGKJSK, CJBANDNKJSK"
        sql = sql & ", CJBANDGKJSK, CJRINGCALC, CJPICALC, CJHANTEI, CJDMAXPIC5"
'Chg Start 2012/05/22 SMPK H.Ohkubo CLESTA�]�����x������Ή�
'        sql = sql & ", CJRMAXPIC5, CJDRMAXPIC5, CJALLMAXDIC5, CJALLMINRINC5, CJALLMAXRIGC5"
        sql = sql & ", CJRMAXPIC5, CJDRMAXPIC5, CJALLMINDIC5, CJALLMAXDIC5, CJALLMINRINC5, CJALLMAXRIGC5"
'Chg End 2012/05/22 SMPK H.Ohkubo CLESTA�]�����x������Ή�
        
        sql = sql & " from TBCMJ023"
        sql = sql & " where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "       SMPLNO = " & wCrySMPL & " and"
        sql = sql & "       TRANCNT = (select max(TRANCNT) from TBCMJ023"
        sql = sql & "                  where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "                        SMPLNO = " & wCrySMPL & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Cu_deco_CJ
                
                If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")             ' �����ԍ�
                .POSITION = CInt(fncNullCheck(rs("POSITION")))                          ' �ʒu
                If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")             ' �T���v���敪
                .TRANCNT = CInt(fncNullCheck(rs("TRANCNT")))                            ' ������
                .SMPLNO = CLng(fncNullCheck(rs("SMPLNO")))                              ' �T���v���m��
                If IsNull(rs("SMPLUMUCJ")) = False Then .SMPLUMUCJ = rs("SMPLUMUCJ")    ' �T���v���L�� CJ
                
                If IsNull(rs("CJPTNJSK")) = False Then .CJPTNJSK = rs("CJPTNJSK")                   ' CJ �p�^�[������
                
                .CJDISKJSK = CInt(fncNullCheck(rs("CJDISKJSK")))                                    ' CJ Disk���a����
                .CJRINGNKJSK = CInt(fncNullCheck(rs("CJRINGNKJSK")))                                ' CJ Ring���a����
                .CJRINGGKJSK = CInt(fncNullCheck(rs("CJRINGGKJSK")))                                ' CJ Ring�O�a����
                .CJBANDNKJSK = CInt(fncNullCheck(rs("CJBANDNKJSK")))                                ' CJ Band���a����
                .CJBANDGKJSK = CInt(fncNullCheck(rs("CJBANDGKJSK")))                                ' CJ Band�O�a����
                .CJRINGCALC = CInt(fncNullCheck(rs("CJRINGCALC")))                                  ' CJ Ring���v�Z
                .CJPICALC = CInt(fncNullCheck(rs("CJPICALC")))                                      ' CJ Pi���v�Z
                
                If IsNull(rs("CJHANTEI")) = False Then .CJHANTEI = rs("CJHANTEI")                   ' CJ ���茋��
                
                .CJDMAXPIC5 = CInt(fncNullCheck(rs("CJDMAXPIC5")))                                  ' CJ Disk�̂݃p�^�[�� Pi������l
                .CJRMAXPIC5 = CInt(fncNullCheck(rs("CJRMAXPIC5")))                                  ' CJ Ring�̂݃p�^�[�� Pi������l
                .CJDRMAXPIC5 = CInt(fncNullCheck(rs("CJDRMAXPIC5")))                                ' CJ DiskRing�p�^�[�� Pi������l
'Add Start 2012/05/22 SMPK H.Ohkubo CLESTA�]�����x������Ή�
                .CJALLMINDIC5 = CInt(fncNullCheck(rs("CJALLMINDIC5")))                              ' CJ ����Disk���a�����l
'Add End 2012/05/22 SMPK H.Ohkubo CLESTA�]�����x������Ή�
                .CJALLMAXDIC5 = CInt(fncNullCheck(rs("CJALLMAXDIC5")))                              ' CJ ����Disk���a����l
                .CJALLMINRINC5 = CInt(fncNullCheck(rs("CJALLMINRINC5")))                            ' CJ ����Ring���a�����l
                .CJALLMAXRIGC5 = CInt(fncNullCheck(rs("CJALLMAXRIGC5")))                            ' CJ ����Ring�O�a����l
                
                If IsNull(rs("REGDATECJ")) = False Then .REGDATE = rs("REGDATECJ")       ' �o�^���t
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    CuDeco_CJ_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add End   2011/01/17 SMPK A.Nagamine


'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎��я��擾
'�T�v      :�����֐� Cu-Deco CJ(LT) ���ю擾�p
Private Function CuDeco_CJLT_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             Cu_deco_CJLT As type_DBDRV_scmzc_fcmkc001c_CJLT, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long

    NothingFlag = False

    ' Cu_deco���уe�[�u��(TBCMJ023)����l���擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CuDeco_CJLT_Zisseki"

    CuDeco_CJLT_Zisseki = FUNCTION_RETURN_SUCCESS

    wCryIND = Samp.CRYINDCJLTCS         ' ��ԃt���O CJ(LT)
    wCryRES = Samp.CRYRESCJLTCS         ' ���уt���O CJ(LT)
    wCrySMPL = Samp.CRYSMPLIDCJLTCS     ' �T���v��ID CJ(LT)

    ' �w��(�d�l)�Ǝ���FLG���m�F
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO"
        sql = sql & ", SMPLUMUCJLT, REGDATECJLT"
        sql = sql & ", CJLTPTNJSK, CJLTDISKJSK, CJLTRINGNKJSK, CJLTRINGGKJSK, CJLTBANDNKJSK"
        sql = sql & ", CJLTBANDGKJSK, CJLTRINGCALC, CJLTPICALC, CJLTHANTEI"
        
        sql = sql & " from TBCMJ023"
        sql = sql & " where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "       SMPLNO = " & wCrySMPL & " and"
        sql = sql & "       TRANCNT = (select max(TRANCNT) from TBCMJ023"
        sql = sql & "                  where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "                        SMPLNO = " & wCrySMPL & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Cu_deco_CJLT
                
                If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")             ' �����ԍ�
                .POSITION = CInt(fncNullCheck(rs("POSITION")))                          ' �ʒu
                If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")             ' �T���v���敪
                .TRANCNT = CInt(fncNullCheck(rs("TRANCNT")))                            ' ������
                .SMPLNO = CLng(fncNullCheck(rs("SMPLNO")))                              ' �T���v���m��
                If IsNull(rs("SMPLUMUCJLT")) = False Then .SMPLUMUCJLT = rs("SMPLUMUCJLT")          ' �T���v���L�� CJ(LT)
                
                If IsNull(rs("CJLTPTNJSK")) = False Then .CJLTPTNJSK = rs("CJLTPTNJSK")             ' CJ(LT) �p�^�[������
                
                .CJLTDISKJSK = CInt(fncNullCheck(rs("CJLTDISKJSK")))                                ' CJ(LT) Disk���a����
                .CJLTRINGNKJSK = CInt(fncNullCheck(rs("CJLTRINGNKJSK")))                            ' CJ(LT) Ring���a����
                .CJLTRINGGKJSK = CInt(fncNullCheck(rs("CJLTRINGGKJSK")))                            ' CJ(LT) Ring�O�a����
                .CJLTBANDNKJSK = CInt(fncNullCheck(rs("CJLTBANDNKJSK")))                            ' CJ(LT) Band���a����
                .CJLTBANDGKJSK = CInt(fncNullCheck(rs("CJLTBANDGKJSK")))                            ' CJ(LT) Band�O�a����
                .CJLTRINGCALC = CInt(fncNullCheck(rs("CJLTRINGCALC")))                              ' CJ(LT) Ring���v�Z
                .CJLTPICALC = CInt(fncNullCheck(rs("CJLTPICALC")))                                  ' CJ(LT) Pi���v�Z
                
                If IsNull(rs("CJLTHANTEI")) = False Then .CJLTHANTEI = rs("CJLTHANTEI")             ' CJ(LT) ���茋��
                
                If IsNull(rs("REGDATECJLT")) = False Then .REGDATE = rs("REGDATECJLT")       ' �o�^���t
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    CuDeco_CJLT_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add End   2011/01/17 SMPK A.Nagamine


'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)�̎��я��擾
'�T�v      :�����֐� Cu-Deco CJ2 ���ю擾�p
Private Function CuDeco_CJ2_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             Cu_deco_CJ2 As type_DBDRV_scmzc_fcmkc001c_CJ2, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long

    NothingFlag = False

    ' Cu_deco���уe�[�u��(TBCMJ023)����l���擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CuDeco_CJ2_Zisseki"

    CuDeco_CJ2_Zisseki = FUNCTION_RETURN_SUCCESS

    wCryIND = Samp.CRYINDCJ2CS          ' ��ԃt���O CJ2
    wCryRES = Samp.CRYRESCJ2CS          ' ���уt���O CJ2
    wCrySMPL = Samp.CRYSMPLIDCJ2CS      ' �T���v��ID CJ2

    ' �w��(�d�l)�Ǝ���FLG���m�F
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO"
        sql = sql & ", SMPLUMUCJ2, REGDATECJ2"
        sql = sql & ", CJ2PTNJSK, CJ2DISKJSK, CJ2RINGNKJSK, CJ2RINGGKJSK, CJ2PICALC"
        sql = sql & ", CJ2HANTEI, CJ2DMAXPIC5, CJ2RMAXPIC5, CJ2RMINRINC5, CJ2RMAXRIGC5"
        sql = sql & ", CJ2DRMAXPIC5, CJ2DRMINRINC5, CJ2DRMAXRIGC5"
        
        sql = sql & " from TBCMJ023"
        sql = sql & " where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "       SMPLNO = " & wCrySMPL & " and"
        sql = sql & "       TRANCNT = (select max(TRANCNT) from TBCMJ023"
        sql = sql & "                  where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "                        SMPLNO = " & wCrySMPL & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Cu_deco_CJ2
                
                If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")             ' �����ԍ�
                .POSITION = CInt(fncNullCheck(rs("POSITION")))                          ' �ʒu
                If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")             ' �T���v���敪
                .TRANCNT = CInt(fncNullCheck(rs("TRANCNT")))                            ' ������
                .SMPLNO = CLng(fncNullCheck(rs("SMPLNO")))                              ' �T���v���m��
                If IsNull(rs("SMPLUMUCJ2")) = False Then .SMPLUMUCJ2 = rs("SMPLUMUCJ2")             ' �T���v���L��CJ2
                
                If IsNull(rs("CJ2PTNJSK")) = False Then .CJ2PTNJSK = rs("CJ2PTNJSK")                ' CJ2 �p�^�[������
                
                .CJ2DISKJSK = CInt(fncNullCheck(rs("CJ2DISKJSK")))                                  ' CJ2 Disk���a����
                .CJ2RINGNKJSK = CInt(fncNullCheck(rs("CJ2RINGNKJSK")))                              ' CJ2 Ring���a����
                .CJ2RINGGKJSK = CInt(fncNullCheck(rs("CJ2RINGGKJSK")))                              ' CJ2 Ring�O�a����
                .CJ2PICALC = CInt(fncNullCheck(rs("CJ2PICALC")))                                    ' CJ2 Pi���v�Z
                
                If IsNull(rs("CJ2HANTEI")) = False Then .CJ2HANTEI = rs("CJ2HANTEI")                ' CJ2 ���茋��
                
                .CJ2DMAXPIC5 = CInt(fncNullCheck(rs("CJ2DMAXPIC5")))                                ' CJ2 Disk�̂݃p�^�[�� Pi�������l
                .CJ2RMAXPIC5 = CInt(fncNullCheck(rs("CJ2RMAXPIC5")))                                ' CJ2 Ring�̂݃p�^�[�� Pi�������l
                .CJ2RMINRINC5 = CInt(fncNullCheck(rs("CJ2RMINRINC5")))                              ' CJ2 Ring�̂݃p�^�[�� Ring���a�����l
                .CJ2RMAXRIGC5 = CInt(fncNullCheck(rs("CJ2RMAXRIGC5")))                              ' CJ2 Ring�̂݃p�^�[�� Ring�O�a����l
                .CJ2DRMAXPIC5 = CInt(fncNullCheck(rs("CJ2DRMAXPIC5")))                              ' CJ2 DiskRing�p�^�[�� Pi�������l
                .CJ2DRMINRINC5 = CInt(fncNullCheck(rs("CJ2DRMINRINC5")))                            ' CJ2 DiskRing�p�^�[�� Ring���a�����l
                .CJ2DRMAXRIGC5 = CInt(fncNullCheck(rs("CJ2DRMAXRIGC5")))                            ' CJ2 DiskRing�p�^�[�� Ring�O�a����l
                
                If IsNull(rs("REGDATECJ2")) = False Then .REGDATE = rs("REGDATECJ2")       ' �o�^���t
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    CuDeco_CJ2_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add End   2011/01/17 SMPK A.Nagamine

