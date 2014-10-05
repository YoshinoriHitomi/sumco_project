Attribute VB_Name = "s_cmzccj"
Option Explicit
''Oi����\����
Public Type C_Oi
    GuaranteeOi         As Guarantee    ''�i���ۏ؏��\����
    SpecOiMin           As Double       ''�iSX�_�f�Z�x����
    SpecOiMax           As Double       ''�iSX�_�f�Z�x���
    SpecORG             As Double       ''�iSX�_�f�Z�x�ʓ����z
    SpecOiAveMin        As Double       ''�iSX�_�f�Z�x���ω���
    SpecOiAveMax        As Double       ''�iSX�_�f�Z�x���Ϗ��
    Oi()                As Double       ''Oi����l
    ORG                 As Double       ''ORG�v�Z�l
    JudgData            As Double       ''Oi����Ώۃf�[�^
    JudgOi              As Boolean      ''Oi���茋��
    JudgOrg             As Boolean      ''ORG���茋��
End Type

''Cs����\����
Public Type C_Cs
    GuaranteeCs         As Guarantee    ''�i���ۏ؏��\����
    SpecCsMin           As Double       ''�iSX�Y�f�Z�x����
    SpecCsMax           As Double       ''�iSX�Y�f�Z�x���
    SpecCsKHI           As String * 1   ''�iSX�����p�x_�� 09/01/08 ooba
    Cs                  As Double       ''Cs����l
    JudgCs              As Boolean      ''Cs���茋��
End Type

''����FTIR����\����
Public Type C_FTIR
    GuaranteeOi         As Guarantee    ''�i���ۏ؏��\����
    GuaranteeCs         As Guarantee    ''�i���ۏ؏��\����
    SpecOiMin           As Double       ''�iSX�_�f�Z�x����
    SpecOiMax           As Double       ''�iSX�_�f�Z�x���
    SpecORG             As Double       ''�iSX�_�f�Z�x�ʓ����z
    SpecOiAveMin        As Double       ''�iSX�_�f�Z�x���ω���
    SpecOiAveMax        As Double       ''�iSX�_�f�Z�x���Ϗ��
    SpecCsMin           As Double       ''�iSX�Y�f�Z�x����
    SpecCsMax           As Double       ''�iSX�Y�f�Z�x���
    Oi(4)               As Double       ''Oi����l
    Cs                  As Double       ''Cs����l
    ORG                 As Double       ''ORG�v�Z�l
    JudgData            As Double       ''Oi����Ώۃf�[�^
    JudgOi              As Boolean      ''Oi���茋��
    JudgOrg             As Boolean      ''ORG���茋��
    JudgCs              As Boolean      ''Cs���茋��
End Type

''����GFA����\����
Public Type C_GFA
    GuaranteeOi         As Guarantee    ''�i���ۏ؏��\����
    SpecOiMin           As Double       ''�iSX�_�f�Z�x����
    SpecOiMax           As Double       ''�iSX�_�f�Z�x���
    SpecORG             As Double       ''�iSX�_�f�Z�x�ʓ����z
    SpecOiAveMin        As Double       ''�iSX�_�f�Z�x���ω���
    SpecOiAveMax        As Double       ''�iSX�_�f�Z�x���Ϗ��
    Ftir(19)            As Double       ''FTIR���Z�l
    ORG                 As Double       ''ORG�v�Z�l
    JudgData            As Double       ''Oi����Ώۃf�[�^
    JudgFtir            As Boolean      ''FTIR���茋��
    JudgOrg             As Boolean      ''ORG���茋��
End Type

''�������R����\����
Public Type C_RES
    GuaranteeRes        As Guarantee    ''�i���ۏ؏��\����
    SpecResMin          As Double       ''�iSX���R����
    SpecResMax          As Double       ''�iSX���R���
    SpecResAveMin       As Double       ''�iSX���R���ω���
    SpecResAveMax       As Double       ''�iSX���R���Ϗ��
    SpecRrg             As Double       ''�iSX���R�ʓ����z
    Res(4)              As Double       ''���R����l
    RRG                 As Double       ''RRG�v�Z�l
    JudgData            As Double       ''���R����Ώۃf�[�^
    JudgRes             As Boolean      ''���R����l
    JudgRes1            As Boolean      ''���R����l
    JudgRrg             As Boolean      ''RRG����l
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    DkTmpSiyo           As String       ''DK���x�i�d�l�j
    DkTmpJsk            As String       ''DK���x�i���сj
    JudgDkTmp           As Boolean      ''DK���x����l
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
End Type

''����BMD����\����
Public Type C_BMD
    GuaranteeBmd        As Guarantee    ''�i���ۏ؏��\����
    SpecBmdAveMin       As Double       ''�iSXBMD���ω���
    SpecBmdAveMax       As Double       ''�iSXBMD���Ϗ��
    BMD(4)              As Double       ''BMD����l
    Min                 As Double       ''�ŏ��l
    max                 As Double       ''�ő�l
    AVE                 As Double       ''���ϒl
    JudgBmd             As Boolean      ''BMD���茋��
    Bunpu               As Double       ''�ʓ����z
End Type

''����OSF����\����
Public Type C_OSF
    GuaranteeOsf        As Guarantee    ''�i���ۏ؏��\����
    SpecOsfAveMax       As Double       ''�iSXOSF���Ϗ��
    SpecOsfMax          As Double       ''�iSX���
    OSF(19)             As Double       ''OSF����l
    max                 As Double       ''�ő�l
    AVE                 As Double       ''���ϒl
    JudgOsf             As Boolean      ''OSF���茋��
    RD1                 As String * 1   ''RD1
    RD2                 As String * 1   ''RD2
    RD3                 As String * 1   ''RD3
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    ARPTK               As String * 1   '�iSXOSF1(ArAN)�p�^���敪
    ARMIN               As Double       '�iSXOSF(ArAN)����
    ARMAX               As Double       '�iSXOSF(ArAN)���
    ARMHMX              As Double       '�iSXOSF(ArAN)�ʓ�����
    CALCMH              As Double       '�ʓ���(MAX/MIN)
    ArAveMin            As Double       '
    ArAveMax            As Double       '
    JudgOsfPtn          As Boolean      ''OSF�p�^�[�����茋��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
End Type

''����GD����\����
Public Type C_GD
    GuaranteeDen        As Guarantee    ''�i���ۏ؏��\����
    GuaranteeLdl        As Guarantee    ''�i���ۏ؏��\����
    GuaranteeDvd2       As Guarantee    ''�i���ۏ؏��\����
    JudgFlagDen         As String * 1   ''�iSXDen�����L��
    JudgFlagLdl         As String * 1   ''�iSXL/DL�����L��
    JudgFlagDvd2        As String * 1   ''�iSXDVD2�����L��
    SpecDenMin          As Double       ''�iSXDen����
    SpecDenMax          As Double       ''�iSXDen���
    SpecLdlMin          As Double       ''�iSXLdl����
    SpecLdlMax          As Double       ''�iSXLdl���
    SpecDvd2Min         As Double       ''�iSXDvd2����
    SpecDvd2Max         As Double       ''�iSXDvd2���
'*** UPDATE �� Y.SIMIZU 2005/10/13 �iWFGDײݐ�
    SpecGdLine          As Single       ''�iSXGDײݐ�
'*** UPDATE �� Y.SIMIZU 2005/10/13 �iWFGDײݐ�
    Den                 As Double       ''Den�v�Z�l
    Ldl                 As Double       ''L/DL�v�Z�l
    Dvd2                As Double       ''Dvd2�v�Z�l
    JudgDen             As Boolean      ''Den���茋��
    JudgLdl             As Boolean      ''L/DL���茋��
    JudgDvd2            As Boolean      ''Dvd2���茋��
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    GDPTK               As String * 1   ''�i�r�w�f�c�p�^���敪
    LdlMin              As Integer      ''L/DL�A��0MIN
    LdlMax              As Integer      ''L/DL�A��0MAX
    ZeroLdlMin          As Integer      ''�iSXLdl�A��0����
    ZeroLdlMax          As Integer      ''�iSXLdl�A��0���
    JudgLdlPtn          As Boolean      ''L/DL�p�^�[�����茋��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
End Type

''�������C�t�^�C������\����
Type C_LT
    GuaranteeLt         As Guarantee    ''�i���ۏ؏��\����
    SpecLtMin           As Double       ''�iSXL�^�C������
    SpecLtMax           As Double       ''�iSXL�^�C�����
    SpecLt10Min         As Double       ''�iSXL�^�C������(10�����Z�l)
    Lt                  As Double       ''���C�t�^�C���v�Z�l
    Lt10                As Double       ''�v�Z�l(10�����Z�l) Add 2011/07/21 T.Koi(SETsw)
    JudgLt              As Boolean      ''���C�t�^�C�����茋��
    JudgLt10            As Boolean      ''���茋��(10�����Z�l) Add 2011/07/21 T.Koi(SETsw)
    resLt10             As String       ''0:�҂� 1:OK 2:NG
End Type

''����EPD����\����
Type C_EPD
    SpecEpdMax          As Double       ''���������Ǘ��EPD���
    EPD                 As Double       ''EPD����l
    JudgEpd             As Boolean      ''EPD���茋��
End Type

'2009/08 SUMCO Akizuki
''�w��������� ����\����
Type C_XY
    Spec_X              As Double        ''����l�@������<X>
    SpecX_Max           As Double        ''        ������<X>���
    SpecX_Min           As Double        ''        ������<X>����
      
    Spec_Y              As Double        ''����l  �c����<Y>
    SpecY_Max           As Double        ''        �c����<Y>���
    SpecY_Min           As Double        ''        �c����<Y>����
    
    Spec_XY             As Double        ''�����p<����>
    SpecXY_Max          As Double        ''�w�������p <����>���
    SpecXY_Min          As Double        ''�w�������p <����>����
    
    JudgResult_X       As Boolean        ''X�����@���茋��
    JudgResult_Y       As Boolean        ''Y�����@���茋��
    JudgResult_XY       As Boolean       ''XY�@���茋��
End Type

'Add Start 2011/01/07 SMPK Miyata
''Cu-deco������� ����\����
Public Type C_CUDECO
    GuaranteeC          As Guarantee    ''�i���ۏ؏��\����
    GuaranteeCJ         As Guarantee    ''�i���ۏ؏��\����
    GuaranteeCJLT       As Guarantee    ''�i���ۏ؏��\����
    GuaranteeCJ2        As Guarantee    ''�i���ۏ؏��\����
    ''----- C���� ----------------------
    HSXCPK              As String * 1   ''�i�r�w�b�p�^�[���敪
    HSXCSZ              As String * 1   ''�i�r�w�b�������

    CPTNJSK             As String * 1   ''C �p�^�[������
    CDISKJSK            As Integer      ''C Disk���a����
    CRINGNKJSK          As Integer      ''C Ring���a����
    CRINGGKJSK          As Integer      ''C Ring�O�a����
    ''----- CJ���� ---------------------
    HSXCJPK             As String * 1   ''�i�r�w�b�i�p�^�[���敪
    HSXCJNS             As String * 2   ''�i�r�w�b�i�M�����@

    CJPTNJSK            As String       ''CJ �p�^�[������
    CJDISKJSK           As Integer      ''CJ Disk���a����
    CJRINGNKJSK         As Integer      ''CJ Ring���a����
    CJRINGGKJSK         As Integer      ''CJ Ring�O�a����
    CJBANDNKJSK         As Integer      ''CJ Band���a����
    CJBANDGKJSK         As Integer      ''CJ Band�O�a����
    CJRINGCALC          As Integer      ''CJ Ring���v�Z
    CJPICALC            As Integer      ''CJ Pi���v�Z
    CJHANTEI            As String       ''CJ ���茋��
    CJBUIUMU            As String       ''CJ ���ʕʔ���L��
    CJDMAXPIC5          As Integer      ''CJ Disk�̂݃p�^�[�� Pi������l
    CJRMAXPIC5          As Integer      ''CJ Ring�̂݃p�^�[�� Pi������l
    CJDRMAXPIC5         As Integer      ''CJ DiskRing�p�^�[�� Pi������l
    CJALLMAXDIC5        As Integer      ''CJ ����Disk���a����l
    CJALLMINRINC5       As Integer      ''CJ ����Ring���a�����l
    CJALLMAXRIGC5       As Integer      ''CJ ����Ring�O�a����l
    ''----- CJLT���� -------------------
    HSXCJLTPK           As String * 1   ''�i�r�w�b�i�k�s�p�^�[���敪
    HSXCJLTNS           As String * 2   ''�i�r�w�b�i�k�s�M�����@
    
    CJLTPTNJSK          As String       ''CJ(LT) �p�^�[������
    CJLTDISKJSK         As Integer      ''CJ(LT) Disk���a����
    CJLTRINGNKJSK       As Integer      ''CJ(LT) Ring���a����
    CJLTRINGGKJSK       As Integer      ''CJ(LT) Ring�O�a����
    CJLTBANDNKJSK       As Integer      ''CJ(LT) Band���a����
    CJLTBANDGKJSK       As Integer      ''CJ(LT) Band�O�a����
    CJLTRINGCALC        As Integer      ''CJ(LT) Ring���v�Z
    CJLTPICALC          As Integer      ''CJ(LT) Pi���v�Z
    CJLTBANDCALC        As Integer      ''CJ(LT) Band���v�Z
    HSXCJLTBND          As Integer      ''CJ(LT) Band������l
    ''----- CJ2���� --------------------
    HSXCJ2PK            As String * 1   ''�i�r�w�b�i�Q�p�^�[���敪
    HSXCJ2NS            As String * 2   ''�i�r�w�b�i�Q�M�����@

    CJ2PTNJSK           As String       ''CJ2 �p�^�[������
    CJ2DISKJSK          As Integer      ''CJ2 Disk���a����
    CJ2RINGNKJSK        As Integer      ''CJ2 Ring���a����
    CJ2RINGGKJSK        As Integer      ''CJ2 Ring�O�a����
    CJ2PICALC           As Integer      ''CJ2 Pi���v�Z
    CJ2HANTEI           As String       ''CJ2 ���茋��
    CJ2BUIUMU           As String       ''CJ2 ���ʕʔ���L��
    CJ2DMAXPIC5         As Integer      ''CJ2 Disk�̂݃p�^�[�� Pi������l
    CJ2RMAXPIC5         As Integer      ''CJ2 Ring�̂݃p�^�[�� Pi������l
    CJ2RMINRINC5        As Integer      ''CJ2 Ring�̂݃p�^�[�� Ring���a�����l
    CJ2RMAXRIGC5        As Integer      ''CJ2 Ring�̂݃p�^�[�� Ring�O�a����l
    CJ2DRMAXPIC5        As Integer      ''CJ2 DiskRing�p�^�[�� Pi������l
    CJ2DRMINRINC5       As Integer      ''CJ2 DiskRing�p�^�[�� Ring���a�����l
    CJ2DRMAXRIGC5       As Integer      ''CJ2 DiskRing�p�^�[�� Ring�O�a����l

    JudgC               As Boolean      ''���茋��C
    JudgCJ              As Boolean      ''���茋��CJ
    JudgCJLT            As Boolean      ''���茋��CJLT
    JudgCJ2             As Boolean      ''���茋��CJ2

End Type

''C���� ����\����
Public Type C_C
    GuaranteeC          As Guarantee    ''�i���ۏ؏��\����
    HSXCPK              As String * 1   ''�i�r�w�b�p�^�[���敪
    HSXCSZ              As String * 1   ''�i�r�w�b�������

    CPTNJSK             As String * 1   ''C �p�^�[������
    CDISKJSK            As Integer      ''C Disk���a����
    CRINGNKJSK          As Integer      ''C Ring���a����
    CRINGGKJSK          As Integer      ''C Ring�O�a����

    JudgC               As Boolean      ''���茋��C
End Type

''CJ���� ����\����
Public Type C_CJ
    GuaranteeCJ         As Guarantee    ''�i���ۏ؏��\����
    HSXCJPK             As String * 1   ''�i�r�w�b�i�p�^�[���敪
    HSXCJNS             As String * 2   ''�i�r�w�b�i�M�����@

    CJPTNJSK            As String       ''CJ �p�^�[������
    CJDISKJSK           As Integer      ''CJ Disk���a����
    CJRINGNKJSK         As Integer      ''CJ Ring���a����
    CJRINGGKJSK         As Integer      ''CJ Ring�O�a����
    CJBANDNKJSK         As Integer      ''CJ Band���a����
    CJBANDGKJSK         As Integer      ''CJ Band�O�a����
    CJRINGCALC          As Integer      ''CJ Ring���v�Z
    CJPICALC            As Integer      ''CJ Pi���v�Z
    CJHANTEI            As String       ''CJ ���茋��
    CJBUIUMU            As String       ''CJ ���ʕʔ���L��
    CJDMAXPIC5          As Integer      ''CJ Disk�̂݃p�^�[�� Pi������l
    CJRMAXPIC5          As Integer      ''CJ Ring�̂݃p�^�[�� Pi������l
    CJDRMAXPIC5         As Integer      ''CJ DiskRing�p�^�[�� Pi������l
    CJALLMAXDIC5        As Integer      ''CJ ����Disk���a����l
    CJALLMINRINC5       As Integer      ''CJ ����Ring���a�����l
    CJALLMAXRIGC5       As Integer      ''CJ ����Ring�O�a����l
    
    JudgCJ              As Boolean      ''���茋��CJ
End Type

''CJLT���� ����\����
Public Type C_CJLT
    GuaranteeCJLT       As Guarantee    ''�i���ۏ؏��\����
    HSXCJLTPK           As String * 1   ''�i�r�w�b�i�k�s�p�^�[���敪
    HSXCJLTNS           As String * 2   ''�i�r�w�b�i�k�s�M�����@
    
    CJLTPTNJSK          As String       ''CJ(LT) �p�^�[������
    CJLTDISKJSK         As Integer      ''CJ(LT) Disk���a����
    CJLTRINGNKJSK       As Integer      ''CJ(LT) Ring���a����
    CJLTRINGGKJSK       As Integer      ''CJ(LT) Ring�O�a����
    CJLTBANDNKJSK       As Integer      ''CJ(LT) Band���a����
    CJLTBANDGKJSK       As Integer      ''CJ(LT) Band�O�a����
    CJLTRINGCALC        As Integer      ''CJ(LT) Ring���v�Z
    CJLTPICALC          As Integer      ''CJ(LT) Pi���v�Z
    CJLTBANDCALC        As Integer      ''CJ(LT) Band���v�Z
    HSXCJLTBND          As Integer      ''CJ(LT) Band������l

    JudgCJLT            As Boolean      ''���茋��CJLT
End Type

''CJ2���� ����\����
Public Type C_CJ2
    GuaranteeCJ2        As Guarantee    ''�i���ۏ؏��\����
    HSXCJ2PK            As String * 1   ''�i�r�w�b�i�Q�p�^�[���敪
    HSXCJ2NS            As String * 2   ''�i�r�w�b�i�Q�M�����@

    CJ2PTNJSK           As String       ''CJ2 �p�^�[������
    CJ2DISKJSK          As Integer      ''CJ2 Disk���a����
    CJ2RINGNKJSK        As Integer      ''CJ2 Ring���a����
    CJ2RINGGKJSK        As Integer      ''CJ2 Ring�O�a����
    CJ2PICALC           As Integer      ''CJ2 Pi���v�Z
    CJ2HANTEI           As String       ''CJ2 ���茋��
    CJ2BUIUMU           As String       ''CJ2 ���ʕʔ���L��
    CJ2DMAXPIC5         As Integer      ''CJ2 Disk�̂݃p�^�[�� Pi������l
    CJ2RMAXPIC5         As Integer      ''CJ2 Ring�̂݃p�^�[�� Pi������l
    CJ2RMINRINC5        As Integer      ''CJ2 Ring�̂݃p�^�[�� Ring���a�����l
    CJ2RMAXRIGC5        As Integer      ''CJ2 Ring�̂݃p�^�[�� Ring�O�a����l
    CJ2DRMAXPIC5        As Integer      ''CJ2 DiskRing�p�^�[�� Pi������l
    CJ2DRMINRINC5       As Integer      ''CJ2 DiskRing�p�^�[�� Ring���a�����l
    CJ2DRMAXRIGC5       As Integer      ''CJ2 DiskRing�p�^�[�� Ring�O�a����l

    JudgCJ2             As Boolean      ''���茋��CJ2
End Type
'Add End   2011/01/07 SMPK Miyata

''�����u���b�N�ΐ͔���\����  2005/1/11�ǉ�
Type C_COEF
    NP                   As String       ''
    COEF                 As Double       ''
    JudgCOEF             As Boolean      ''
End Type
''�ΐ͔͈͒l
Public Const PminusMin As Double = 0.7
Public Const PminusMax As Double = 0.8
Public Const PplusMin As Double = 0.73
Public Const PplusMax As Double = 0.83
Public Const NMin As Double = 0.3
Public Const NMax As Double = 0.4

''SIRD����\����   2010/02/04 add Kameda
Type C_SIRD
    SpecSirdMax         As Double       ''�d�l�ʓ������
    SIRDCNT             As Double       ''SIRD����l
    JudgSird            As Boolean      ''SIRD���茋��
End Type


''���H�d�l,���H���э\����
''�����̒Ⴂ����MIN�l��������
Public Type Judg_Kakou
    top(2) As Double
    TAIL(2) As Double
    
    POS As String * 2
    DPTH(2) As Double   '�����グ�����̏ꍇ�f�[�^�͂ЂƂ������݂��Ȃ�
    WIDH(2) As Double   '�����グ�����̏ꍇ�f�[�^�͂ЂƂ������݂��Ȃ�
    ANGLE(2) As Double     '2009/09 SUMCO Akizuki
End Type

''���H���є��茋�ʍ\����
Public Type Judg_Kakou_Judg
    top As Boolean
    tTOP(2) As Boolean
    TAIL As Boolean
    tTAIL(2) As Boolean
    
    POS As Boolean          '�m�b�`�ʒu
    WIDH As Boolean         '�m�b�`��
    DPTH As Boolean         '�m�b�`�[��(TOP)
'    DPTH_BOT As Boolean     '�m�b�`�[��(BOT)
    ANGLE As Boolean        '�m�b�`�p�x
End Type

''���H���є���\����
Public Type type_KakouJudg
    Spec() As Judg_Kakou
    Jiltuseki As Judg_Kakou
    Judg As Judg_Kakou_Judg
'' 09/01/28 FAE)akiyama start
    BLOCKID As String
'' 09/01/28 FAE)akiyama end
End Type

'�U�ւ��s�����i�ԁ@2003/09/05
Public Type fHinban
    moto As String
    saki As String
End Type


'�T�v      :����FTIR������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Ftir          ,I  ,C_FTIR           ,����FTIR����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
'�@�@      :2001/07/19 ���� �M�� ����
Public Function CrystalFTIRJudg(Ftir As C_FTIR, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim Oi As C_Oi
    Dim Cs As C_Cs
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Oi.GuaranteeOi = Ftir.GuaranteeOi
    Oi.SpecOiMin = Ftir.SpecOiMin
    Oi.SpecOiMax = Ftir.SpecOiMax
    Oi.SpecORG = Ftir.SpecORG
    Oi.SpecOiAveMin = Ftir.SpecOiAveMin
    Oi.SpecOiAveMax = Ftir.SpecOiAveMax
    ReDim Oi.Oi(UBound(Ftir.Oi)) As Double
    For c0 = 0 To UBound(Ftir.Oi)
        Oi.Oi(c0) = Ftir.Oi(c0)
    Next
    Oi.ORG = Ftir.ORG
    
    FuncAns = CrystalOiJudg(Oi, ErrInfo)

    Ftir.JudgData = Oi.JudgData
    Ftir.JudgOi = Oi.JudgOi
    Ftir.JudgOrg = Oi.JudgOrg

'    If Ftir.GuaranteeOi.cJudg = JudgCodeC01 Then ''Oi����L��
'
'        ''ORG����
'        Ftir.JudgOrg = RangeDecision_nl(Ftir.ORG, 0, Ftir.SpecORG)
'
'        ''Oi����
'        If (InStr(ObjCodeGrp01, Ftir.GuaranteeOi.cObj) <> 0) And (GetCrystalJudgData(Ftir.GuaranteeOi, Ftir.Oi(), JData()) = FUNCTION_RETURN_SUCCESS) Then
'            Select Case Ftir.GuaranteeOi.cObj
'            Case ObjCode01, ObjCode02, ObjCode04 ''���S1�_�A�����l�AR/2
'                Ftir.JudgOi = RangeDecision_nl(JData(0), Ftir.SpecOiMin, Ftir.SpecOiMax)
'            Case ObjCode03 ''�S��
'                Ftir.JudgOi = JUDG_OK
'                For c0 = 0 To 4
'                    If JData(c0) <> -1 Then
'                        If RangeDecision_nl(JData(c0), Ftir.SpecOiMin, Ftir.SpecOiMax) = JUDG_NG Then
'                            Ftir.JudgOi = JUDG_NG
'                        End If
'                    End If
'                Next
'            End Select
'        Else
'            ''�Ώۃf�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Ftir.GuaranteeOi.cObj)
'        End If
'        Ftir.JudgOi = (Ftir.JudgOi And Ftir.JudgOrg)
''    Else
''        If InStr(JudgCodeC02, Ftir.GuaranteeOi.cJudg) = 0 Then
''            ''�������@�f�[�^����
''            ''�G���[���\���̂ɏ������B
''            FuncAns = SetErrInfo(ErrInfo, ZJ001, OI_JUDG, Ftir.GuaranteeOi.cJudg)
''        End If
'    End If
    
    If FuncAns = FUNCTION_RETURN_FAILURE Then Exit Function
    
    Cs.GuaranteeCs = Ftir.GuaranteeCs
    Cs.SpecCsMin = Ftir.SpecCsMin
    Cs.SpecCsMax = Ftir.SpecCsMax
    Cs.Cs = Ftir.Cs
    
    FuncAns = CrystalCsJudg(Cs, ErrInfo)

    Ftir.JudgCs = Cs.JudgCs

'    If Ftir.GuaranteeOi.cJudg = JudgCodeC01 Then ''Cs����L��
'        Ftir.JudgCs = RangeDecision_nl(Ftir.Cs, Ftir.SpecCsMin, Ftir.SpecCsMax)
''    Else
''        If InStr(JudgCodeC02, Ftir.GuaranteeCs.cJudg) = 0 Then
''            ''�������@�f�[�^����
''            ''�G���[���\���̂ɏ������B
''            FuncAns = SetErrInfo(ErrInfo, ZJ001, CS_JUDG, Ftir.GuaranteeCs.cJudg)
''        End If
'    End If

    CrystalFTIRJudg = FuncAns
End Function

'�T�v      :����GFA������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Gfa           ,I  ,C_GFA            ,����GFA����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function CrystalGFAJudg(Gfa As C_GFA, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim Oi As C_Oi
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Oi.GuaranteeOi = Gfa.GuaranteeOi
    Oi.SpecOiMin = Gfa.SpecOiMin
    Oi.SpecOiMax = Gfa.SpecOiMax
    Oi.SpecORG = Gfa.SpecORG
    Oi.SpecOiAveMin = Gfa.SpecOiAveMin
    Oi.SpecOiAveMax = Gfa.SpecOiAveMax
    ReDim Oi.Oi(UBound(Gfa.Ftir)) As Double
    For c0 = 0 To UBound(Gfa.Ftir)
        Oi.Oi(c0) = Gfa.Ftir(c0)
    Next
    Oi.ORG = Gfa.ORG
    
    FuncAns = CrystalOiJudg(Oi, ErrInfo)
    
    Gfa.JudgData = Oi.JudgData
    Gfa.JudgFtir = Oi.JudgOi
    Gfa.JudgOrg = Oi.JudgOrg
'    If Gfa.GuaranteeOi.cJudg = JudgCodeC01 Then ''GFA����L��
'
'        ''ORG����
'        Gfa.JudgOrg = RangeDecision_nl(Gfa.ORG, 0, Gfa.SpecORG)
'
'        ''FTIR����
'        If (InStr(ObjCodeGrp01, Gfa.GuaranteeOi.cObj) <> 0) And (GetCrystalJudgData(Gfa.GuaranteeOi, Gfa.Ftir(), JData()) = FUNCTION_RETURN_SUCCESS) Then
'            Select Case Gfa.GuaranteeOi.cObj
'            Case ObjCode01, ObjCode02, ObjCode04 ''���S1�_�A�����l�AR/2
'                Gfa.JudgFtir = RangeDecision_nl(JData(0), Gfa.SpecOiMin, Gfa.SpecOiMax)
'            Case ObjCode03 ''�S��
'                Gfa.JudgFtir = JUDG_OK
'                For c0 = 0 To 19
'                    If JData(c0) <> -1 Then
'                        If RangeDecision_nl(JData(c0), Gfa.SpecOiMin, Gfa.SpecOiMax) = JUDG_NG Then
'                            Gfa.JudgFtir = JUDG_NG
'                        End If
'                    End If
'                Next
'            End Select
'        Else
'            ''�Ώۃf�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, EZJ00, GFA_JUDG, Gfa.GuaranteeOi.cObj)
'        End If
'        Gfa.JudgFtir = (Gfa.JudgFtir And Gfa.JudgOrg)
'    Else
''        If InStr(JudgCodeC02, Gfa.GuaranteeOi.cJudg) = 0 Then
''            ''�������@�f�[�^����
''            ''�G���[���\���̂ɏ������B
''            FuncAns = SetErrInfo(ErrInfo, ZJ001, GFA_JUDG, Gfa.GuaranteeOi.cJudg)
''        End If
'    End If
    
    CrystalGFAJudg = FuncAns
End Function

'�T�v      :�������R������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Res           ,I  ,C_RES            ,�������R����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function CrystalRESJudg(Res As C_RES, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim JData(4) As Double
    Dim c0 As Integer
    Dim pt As Integer
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Dim iRet    As Integer
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Res.JudgData = -1
    Res.JudgRes = JUDG_NG
    Res.JudgRes1 = JUDG_NG
    Res.JudgRrg = JUDG_NG
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Res.JudgDkTmp = JUDG_NG
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    If Res.GuaranteeRes.cJudg = JudgCodeC01 Then ''RES����L��


'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06 start

'        If Trim(Res.GuaranteeRes.cCount) = "" Then
'            pt = 3
'        Else
'            pt = Val(Res.GuaranteeRes.cCount)
'        End If
'        Res.RRG = RoundUp((RGCal(Res.Res(), pt)), 4)
        
        
        ''RRG����
        Select Case Res.GuaranteeRes.cPos
          Case "B", "C", "D", "E", "F", "K", "S", "Y"
              Select Case Res.GuaranteeRes.cBunp
              Case "A", "B", "C", "M"
                 ''RRG�v�Z
                 Res.RRG = MENNAI_Cal(RES_JUDG, Res.Res(), Res.GuaranteeRes, Res.GuaranteeRes.cBunp)

              Case "", " "
                 ''�v�Z�敪���X�y�[�X�̏ꍇ�́A�v�Z�C������s��Ȃ�
                 Res.RRG = 0
                 Res.JudgRrg = JUDG_OK
                 GoTo Cal_Escp
              Case Else
                 ''RRG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
                 If Trim(Res.GuaranteeRes.cCount) = "" Then
                    pt = 3
                 Else
                    pt = val(Res.GuaranteeRes.cCount)
                 End If
                 Res.RRG = RoundUp((RGCal(Res.Res(), pt)), 4)

             End Select

          Case Else
             Select Case Res.GuaranteeRes.cBunp
             Case "A", "B", "C", "D", "E", "M", "N"
                 ''RRG�v�Z
                 Res.RRG = MENNAI_Cal(RES_JUDG, Res.Res(), Res.GuaranteeRes, Res.GuaranteeRes.cBunp)

             Case "", " "
                 ''�v�Z�敪���X�y�[�X�̏ꍇ�́A�v�Z�C������s��Ȃ�
                 Res.RRG = 0
                 Res.JudgRrg = JUDG_OK
                 GoTo Cal_Escp
             Case Else
                 ''RRG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
                 If Trim(Res.GuaranteeRes.cCount) = "" Then
                    pt = 3
                 Else
                    pt = val(Res.GuaranteeRes.cCount)
                 End If
                 Res.RRG = RoundUp((RGCal(Res.Res(), pt)), 4)

             End Select
        End Select

'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06 end

'2002/02/27 S.Sano RRG�̎d�l��0�̏ꍇ�́A������s�킸�K��OK�Ƃ���B
'2002/02/27 S.Sano �ʓ����z�v�Z�͍s���B
        If Res.SpecRrg = 0 Then                                     '2002/02/27 S.Sano
            Res.JudgRrg = JUDG_OK                                   '2002/02/27 S.Sano
        Else                                                        '2002/02/27 S.Sano
            If Res.RRG = -1 Then
                Res.JudgRrg = JUDG_NG
            Else
                Res.JudgRrg = RangeDecision_nl(Res.RRG, 0, Res.SpecRrg)
            End If
        End If                                                      '2002/02/27 S.Sano
        
'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06 start
Cal_Escp:
'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06 end
        
        ''RES����
        '-----TEST2004/10 N�ǉ�
        'If (InStr(ObjCodeGrp01, Res.GuaranteeRes.cObj) <> 0) And (GetCrystalJudgData(Res.GuaranteeRes, Res.Res(), JData()) = FUNCTION_RETURN_SUCCESS) Then
        If (InStr(ObjCodeGrp05, Res.GuaranteeRes.cObj) <> 0) And (GetCrystalJudgData(Res.GuaranteeRes, Res.Res(), JData()) = FUNCTION_RETURN_SUCCESS) Then
            Select Case Res.GuaranteeRes.cObj
            Case ObjCode01, ObjCode02, ObjCode04  ''���S1�_�A�����l�AR/2
                Res.JudgRes1 = RangeDecision_nl(JData(0), Res.SpecResMin, Res.SpecResMax)
                Res.JudgData = JData(0)
            'Case ObjCode03 ''�S��
            Case ObjCode03, ObjCode13 ''�S��A�_��
                Res.JudgRes = JUDG_OK
                Res.JudgRes1 = JUDG_OK
                For c0 = 0 To 4
                    If JData(c0) <> -1 Then
                        If RangeDecision_nl(JData(c0), Res.SpecResMin, Res.SpecResMax) = JUDG_NG Then
                            Res.JudgRes1 = JUDG_NG
                        End If
                    End If
                Next
                Res.JudgData = JudgMax(JData())
            End Select
        Else
            ''�Ώۃf�[�^����
            ''�G���[���\���̂ɏ������B
            FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
        End If
        
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        ''DK���x����
        If Trim(Res.DkTmpJsk) = "" Or Trim(Res.DkTmpSiyo) = "" Then
            Res.JudgDkTmp = JUDG_OK
        Else
            iRet = funCodeDBGetMatrixReturn(DKTMP_TBCMB005SYS, DKTMP_TBCMB005CLS, Res.DkTmpJsk, Res.DkTmpSiyo)
            If iRet = -1 Then
                FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
            ElseIf iRet = 0 Then
                Res.JudgDkTmp = JUDG_NG
            Else
                Res.JudgDkTmp = JUDG_OK
            End If
        End If
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'        Res.JudgRes = (Res.JudgRes1 And Res.JudgRrg)
        Res.JudgRes = (Res.JudgRes1 And Res.JudgRrg And Res.JudgDkTmp)
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
    Else
        Res.JudgRrg = JUDG_OK
        Res.JudgRes = JUDG_OK
        Res.JudgRes1 = JUDG_OK
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        Res.JudgDkTmp = JUDG_OK
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'        If InStr(JudgCodeC02, Res.GuaranteeRes.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, RES_JUDG, Res.GuaranteeRes.cJudg)
'        End If
    End If

    CrystalRESJudg = FuncAns
End Function

'�T�v      :����BMD������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Bmd           ,I  ,C_BMD            ,����BMD����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
'          :2001/07/04 ���� �M�� �C��
Public Function CrystalBMDJudg(BMD As C_BMD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    BMD.JudgBmd = JUDG_NG
    If BMD.GuaranteeBmd.cJudg = JudgCodeC01 Then ''BMD����L��
        
        ''BMD����
        If (InStr(ObjCodeGrp02, BMD.GuaranteeBmd.cObj) <> 0) Then
            Select Case BMD.GuaranteeBmd.cObj
            Case ObjCode05 ''�S�_�̕��ϒl
                BMD.JudgBmd = RangeDecision_nl(BMD.AVE, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
            Case ObjCode06, ObjCode10, ObjCode11 ''�S�_�̍ő�l�A�S�_�̍ŏ��l�AMAX(2,4�_��)�AMAX(2,3,4�_��)
                BMD.JudgBmd = RangeDecision_nl(BMD.max, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
            Case ObjCode08 ''�S�_�̍ŏ��l ******************************* �w���P��������s��
                BMD.JudgBmd = RangeDecision_nl(BMD.Min, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
            Case ObjCode07 ''�S�_�̕��ϒl�ƍő�l ************************ �w���P��������s��
                If RangeDecision_nl(BMD.AVE, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax) Then
                    BMD.JudgBmd = RangeDecision_nl(BMD.max, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
                Else
                    BMD.JudgBmd = JUDG_NG
                End If
            '2001/09/19 S.Sano Start
            Case ObjCode16 ''�S�_�̍ŏ��l�ƍő�l ************************ �w���P��������s��
                If RangeDecision_nl(BMD.Min, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax) Then
                    BMD.JudgBmd = RangeDecision_nl(BMD.max, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
                Else
                    BMD.JudgBmd = JUDG_NG
                End If
            '2001/09/19 S.Sano End
            End Select
        Else
            ''�Ώۃf�[�^����
            ''�G���[���\���̂ɏ������B
            FuncAns = SetErrInfo(ErrInfo, EZJ00, BMD_JUDG, BMD.GuaranteeBmd.cObj)
        End If
    Else
        BMD.JudgBmd = JUDG_OK
'        If InStr(JudgCodeC02, Bmd.GuaranteeBmd.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, BMD_JUDG, Bmd.GuaranteeBmd.cJudg)
'        End If
    End If
    
    CrystalBMDJudg = FuncAns
End Function

'�T�v      :����OSF������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Osf           ,I  ,C_OSF            ,����OSF����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
'          :2001/07/04 ���� �M�� �C��
Public Function CrystalOSFJudg(OSF As C_OSF, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim Index           As Integer
    Dim dAve            As Double
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    OSF.JudgOsf = JUDG_NG
    

    If Trim(OSF.GuaranteeOsf.cJudg) = JudgCodeC01 Then ''OSF����L��
        'OSF�͎d�l�l��Null�̏ꍇ�͑��̔���ƈقȂ�G���[�ƂȂ� 2004/12/22
        '�d�l�l��Null�`�F�b�N�ŃG���[�Ƃ�������ŃG���[�Ƃ���
        'Null�Ή� 08/11/06 ooba
''        If OSF.SpecOsfAveMax = -1 Or OSF.SpecOsfMax = -1 Then
''            Exit Function
''        End If
        ''OSF����
        If (InStr(ObjCodeGrp03, OSF.GuaranteeOsf.cObj) <> 0) Then
            Select Case OSF.GuaranteeOsf.cObj
            Case ObjCode05  ''�S�_�̕��ϒl
                OSF.JudgOsf = RangeDecision_nl(OSF.AVE, 0, OSF.SpecOsfAveMax)
            Case ObjCode06  ''�S�_�̍ő�l
                OSF.JudgOsf = RangeDecision_nl(OSF.max, 0, OSF.SpecOsfMax)
            Case ObjCode07 ''�S�_�̕��ϒl�ƍő�l
                If RangeDecision_nl(OSF.AVE, 0, OSF.SpecOsfAveMax) Then
                    OSF.JudgOsf = RangeDecision_nl(OSF.max, 0, OSF.SpecOsfMax)
                Else
                    OSF.JudgOsf = JUDG_NG
                End If
            End Select
        'Null�Ή�(�K�i��Null�̏ꍇ�͑Ώۺ��ޕs��) 08/11/06 ooba
        ElseIf OSF.SpecOsfAveMax = -1 And OSF.SpecOsfMax = -1 Then
            OSF.JudgOsf = JUDG_OK
        Else
            ''�Ώۃf�[�^����
            ''�G���[���\���̂ɏ������B
            FuncAns = SetErrInfo(ErrInfo, EZJ00, OSF_JUDG, OSF.GuaranteeOsf.cObj)
        End If

    Else
        OSF.JudgOsf = JUDG_OK
'        If InStr(JudgCodeC02, Osf.GuaranteeOsf.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, OSF_JUDG, Osf.GuaranteeOsf.cJudg)
'        End If
    End If
            
    CrystalOSFJudg = FuncAns
    
End Function

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
'�T�v      :����OSF������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Osf           ,I  ,C_OSF            ,����OSF����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2008/10/01 Systech �쐬  L/DL,OSF����ۼޯ��ǉ�
Public Function CrystalOSFJudg_02(OSF As C_OSF, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim Index           As Integer
    Dim Index2          As Integer
    Dim dAve            As Double
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    OSF.JudgOsfPtn = JUDG_OK
    
    If Trim(OSF.GuaranteeOsf.cJudg) = JudgCodeC01 Then ''OSF����L��
        'OSF�͎d�l�l��Null�̏ꍇ�͑��̔���ƈقȂ�G���[�ƂȂ� 2004/12/22
        '�d�l�l��Null�`�F�b�N�ŃG���[�Ƃ�������ŃG���[�Ƃ���
        
        If OSF.ARPTK = "1" Then
        'FG7
'            OSF.JudgOsf = JUDG_OK

'            If OSF.ARMIN = -1 Or _
'               OSF.ARMAX = -1 Then
'            Else
                '�㉺������
                For Index = 0 To 19
                    If OSF.OSF(Index) >= 0 Then
                        OSF.JudgOsfPtn = RangeDecision_nl(OSF.OSF(Index), OSF.ARMIN, OSF.ARMAX) '08/11/06 ooba
                        If OSF.JudgOsfPtn = JUDG_NG Then Exit For   '08/11/06 ooba
'                        If OSF.OSF(Index) >= OSF.ARMIN And _
'                           OSF.OSF(Index) <= OSF.ARMAX Then
'                        Else
'                            OSF.JudgOsfPtn = JUDG_NG
'                            Exit For
'                        End If
                    End If
                Next Index
'            End If
            
            If OSF.JudgOsfPtn = True Then
                If OSF.ARMHMX = -1 Then
                Else
                    '�ʓ���(MAX/MIN)����
''                    dAve = OSF.ArAveMax / OSF.ArAveMin
''                    dAve = (Fix((dAve * 10) + 0.9) / 10)    '������2�ʐ؂�グ
                    
                    If OSF.CALCMH <= OSF.ARMHMX Then
                    Else
                        OSF.JudgOsfPtn = JUDG_NG
                    End If
                End If
            End If
            
        ElseIf OSF.ARPTK = "2" Then
        '����(ArAN)
'            OSF.JudgOsf = JUDG_OK
    
            If OSF.ARMHMX = -1 Then
            Else
                '�ʓ���(MAX/MIN)����
                If OSF.CALCMH <= OSF.ARMHMX Then
                Else
                    OSF.JudgOsfPtn = JUDG_NG
                End If
            End If
            
            If OSF.JudgOsfPtn = True Then
                If OSF.ARMAX = -1 Then
                Else
                    '�������
                    If OSF.CALCMH = -1 Then     '08/11/06 ooba
                    'If OSF.OSF(Index) = 0 Then
                        For Index2 = 0 To 19
                            If OSF.ARMAX >= OSF.OSF(Index2) Then
                            Else
                                OSF.JudgOsfPtn = JUDG_NG
                                Exit For
                            End If
                        Next Index2
                    End If
                End If
            End If
        Else
        '����Ȃ�
'            OSF.JudgOsf = JUDG_OK
        
        End If
    Else
'        OSF.JudgOsf = JUDG_OK
    End If
            
    CrystalOSFJudg_02 = FuncAns
    
End Function
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

'�T�v      :����GD������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Gd            ,I  ,C_GD             ,����GD����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function CrystalGDJudg(GD As C_GD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    ''Den�����L�����f
    GD.JudgDen = JUDG_OK
    If GD.JudgFlagDen = "1" Then
        If GD.GuaranteeDen.cJudg = JudgCodeC01 Then ''Den���肠��
            GD.JudgDen = RangeDecision_nl(GD.Den, GD.SpecDenMin, GD.SpecDenMax)
        Else
            GD.JudgDen = JUDG_OK
'            If InStr(JudgCodeC02, Gd.GuaranteeDen.cJudg) = 0 Then
'                ''�������@�f�[�^����
'                ''�G���[���\���̂ɏ������B
'                FuncAns = SetErrInfo(ErrInfo, ZJ001, DEN_JUDG, Gd.GuaranteeDen.cJudg)
'            End If
        End If
    End If
    
    ''L/DL�����L�����f
    GD.JudgLdl = JUDG_OK
    If GD.JudgFlagLdl = "1" Then
        If GD.GuaranteeLdl.cJudg = JudgCodeC01 Then ''L/DL���肠��
            GD.JudgLdl = RangeDecision_nl(GD.Ldl, GD.SpecLdlMin, GD.SpecLdlMax)
        Else
            GD.JudgLdl = JUDG_OK
'            If InStr(JudgCodeC02, Gd.GuaranteeLdl.cJudg) = 0 Then
'                ''�������@�f�[�^����
'                ''�G���[���\���̂ɏ������B
'                FuncAns = SetErrInfo(ErrInfo, ZJ001, LDL_JUDG, Gd.GuaranteeLdl.cJudg)
'            End If
        End If
    End If
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
'    If GD.JudgLdl = JUDG_OK Then
    GD.JudgLdlPtn = JUDG_OK
    If GD.JudgFlagLdl = "1" Then
        If GD.GuaranteeLdl.cJudg = JudgCodeC01 Then ''L/DL���肠��
            If GD.GDPTK = "1" Then
                ' "0"�A����(MIN)�@���@�i__L/DL�A��0����(SX/WF)
                If GD.ZeroLdlMin = -1 Then
                    GD.JudgLdlPtn = JUDG_OK
                Else
                    If GD.LdlMin >= GD.ZeroLdlMin Then
                        GD.JudgLdlPtn = JUDG_OK
                    Else
                        GD.JudgLdlPtn = JUDG_NG
                    End If
                End If
            ElseIf GD.GDPTK = "2" Then
                ' "0"�A����(MAX)�@���@�i__L/DL�A��0���(SX/WF)
                If GD.ZeroLdlMax = -1 Then
                    GD.JudgLdlPtn = JUDG_OK
                Else
                    If GD.LdlMax <= GD.ZeroLdlMax Then
                        GD.JudgLdlPtn = JUDG_OK
                    Else
                        GD.JudgLdlPtn = JUDG_NG
                    End If
                End If
            Else
                ' ���薳��
                
            End If
        End If
    End If
'    End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    
    ''DVD2�����L�����f
    GD.JudgDvd2 = JUDG_OK
    If GD.JudgFlagDvd2 = "1" Then
        If GD.GuaranteeDvd2.cJudg = JudgCodeC01 Then ''Dvd2���肠��
'���ڒǉ��C�C���Ή� 2003.05.20 yakimura
'            GD.JudgDvd2 = RangeDecision_nl(GD.Dvd2, GD.SpecDvd2Min * 10!, GD.SpecDvd2Max * 10!)
            GD.JudgDvd2 = RangeDecision_nl(GD.Dvd2, GD.SpecDvd2Min, GD.SpecDvd2Max)
'���ڒǉ��C�C���Ή� 2003.05.20 yakimura
        Else
            GD.JudgDvd2 = JUDG_OK
'            If InStr(JudgCodeC02, Gd.GuaranteeDvd2.cJudg) = 0 Then
'                ''�������@�f�[�^����
'                ''�G���[���\���̂ɏ������B
'                FuncAns = SetErrInfo(ErrInfo, ZJ001, DVD2_JUDG, Gd.GuaranteeDvd2.cJudg)
'            End If
        End If
    End If
    
    CrystalGDJudg = FuncAns
End Function

'�T�v      :�������C�t�^�C��������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Lt            ,I  ,C_LT             ,�������C�t�^�C������\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function CrystalLTJudg(Lt As C_LT, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Lt.JudgLt = JUDG_OK
    If Lt.GuaranteeLt.cJudg = JudgCodeC01 Then
        If Lt.Lt < Lt.SpecLtMin Then
            Lt.JudgLt = JUDG_NG
        End If
    Else
'        If InStr(JudgCodeC02, Lt.GuaranteeLt.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, LT_JUDG, Lt.GuaranteeLt.cJudg)
'        End If
    End If
    
    CrystalLTJudg = FuncAns
End Function

'�T�v      :����������s���B�i�P�O�����Z�l�j
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Lt            ,I  ,C_LT             ,��������\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2011/07/21 T.Koi(SETsw)
Public Function CrystalLT10Judg(Lt As C_LT, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Lt.JudgLt10 = JUDG_OK
    If Lt.GuaranteeLt.cJudg = JudgCodeC01 Then
        If Lt.Lt10 < Lt.SpecLt10Min Then
            Lt.JudgLt10 = JUDG_NG
        End If
    End If
    
    CrystalLT10Judg = FuncAns
End Function


'�T�v      :����EPD������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Epd           ,I  ,C_EPD            ,����EPD����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function CrystalEPDJudg(EPD As C_EPD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    ''�G���[���\���̏�����
    CrystalEPDJudg = SetErrInfo(ErrInfo)
    
    EPD.JudgEpd = RangeDecision_nl(EPD.EPD, 0, EPD.SpecEpdMax)
    CrystalEPDJudg = FUNCTION_RETURN_SUCCESS
End Function

'�T�v      :X�� �����p<����>�̔�����s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :XY            ,I  ,C_XY             ,X�� ����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2009/08 SUMCO �H��(EPD����ɕ���āA�쐬)

Public Function CrystalXYJudg(XY As C_XY, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    
    ''�G���[���\���̏�����
    CrystalXYJudg = SetErrInfo(ErrInfo)
    
    ''����̎��{
    
    XY.JudgResult_X = RangeDecision_nl(XY.Spec_X, XY.SpecX_Min, XY.SpecX_Max)
    XY.JudgResult_Y = RangeDecision_nl(XY.Spec_Y, XY.SpecY_Min, XY.SpecY_Max)
    XY.JudgResult_XY = RangeDecision_nl(XY.Spec_XY, XY.SpecXY_Min, XY.SpecXY_Max)
    
    CrystalXYJudg = FUNCTION_RETURN_SUCCESS
    
End Function
'�T�v      :SIRD������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :SIRD          ,I  ,C_SIRD           ,SIRD����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2010/02/04 Kameda
Public Function CrystalSIRDJudg(SIRD As C_SIRD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    ''�G���[���\���̏�����
    CrystalSIRDJudg = SetErrInfo(ErrInfo)
    
    SIRD.JudgSird = RangeDecision_nl(SIRD.SIRDCNT, 0, SIRD.SpecSirdMax)
    CrystalSIRDJudg = FUNCTION_RETURN_SUCCESS
End Function

'Add Start 2011/01/24 SMPK Miyata
'�T�v      :����C������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Cudeco        ,I  ,C_CUDECO         ,����Cu-deco����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :
Public Function CrystalCJudg(CuDeco As C_C, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns         As FUNCTION_RETURN
    Dim ptnOK(3)        As String
    Dim bJudg           As Boolean


    '' �p�^�[��������
    '' �P�����ڂ̓p�^�[���敪�A�Q�����ȍ~���n�j�Ƃ���o�^�[�����т�ݒ肷��
    ''
    ''�@    �o�^�[�����т̎��
    ''          "1" : �����O�����w��    "2" : �f�B�X�N�����w��      "3" : �p�^�[�������w��
    ''          "4" : �s�� (�I���Ȃ�)   "5" : �o���h�����w��        "6" : P�o���h�����w��
    ''          "7" : B�o���h�w�薳��
    ''
    ''�@    �o�^�[�����т̎��
    ''          "0":None    "1":Ring    "2":Disk    "3":Disk & Ring
    ''          "5":PB-band "6":P-band  "7":B-band
    ptnOK(0) = "1 02"        '' �����O�����w��
    ptnOK(1) = "2 01"        '' �f�B�X�N�����w��
    ptnOK(2) = "3 0"         '' �p�^�[�������w��
    ptnOK(3) = "4 0123"      '' �s�� (�I���Ȃ�)
    
    ''�G���[���\���̏�����
    CrystalCJudg = SetErrInfo(ErrInfo)
    
    '*** �p�^�[������ ***
    bJudg = CudecoJudgPattern(CuDeco.HSXCPK, CuDeco.CPTNJSK, ptnOK)
    
    CuDeco.JudgC = bJudg

    CrystalCJudg = FUNCTION_RETURN_SUCCESS

End Function

'�T�v      :����CJ������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Cudeco        ,I  ,C_CUDECO         ,����Cu-deco����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :
Public Function CrystalCJJudg(CuDeco As C_CJ, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns         As FUNCTION_RETURN
    Dim ptnOK(3)        As String
    Dim bJudg           As Boolean
    Dim intMax          As Integer

    '' �p�^�[��������
    '' �P�����ڂ̓p�^�[���敪�A�Q�����ȍ~���n�j�Ƃ���o�^�[�����т�ݒ肷��
    ''
    ''�@    �o�^�[�����т̎��
    ''          "1" : �����O�����w��    "2" : �f�B�X�N�����w��      "3" : �p�^�[�������w��
    ''          "4" : �s�� (�I���Ȃ�)   "5" : �o���h�����w��        "6" : P�o���h�����w��
    ''          "7" : B�o���h�w�薳��
    ''
    ''�@    �o�^�[�����т̎��
    ''          "0":None    "1":Ring    "2":Disk    "3":Disk & Ring
    ''          "5":PB-band "6":P-band  "7":B-band
    ptnOK(0) = "1 02"        '' �����O�����w��
    ptnOK(1) = "2 01"        '' �f�B�X�N�����w��
    ptnOK(2) = "3 0"         '' �p�^�[�������w��
    ptnOK(3) = "4 0123"      '' �s�� (�I���Ȃ�)
    
    ''�G���[���\���̏�����
    CrystalCJJudg = SetErrInfo(ErrInfo)
    
    '*** �p�^�[������ ***
    bJudg = CudecoJudgPattern(CuDeco.HSXCJPK, CuDeco.CJPTNJSK, ptnOK)

    ' CJ Ring���a�E�O�a�̔���
    If bJudg Then
        '' �p�^�[�����т�[Ring] or [Disk & Ring]
        If (CuDeco.CJPTNJSK = CudecoJskPtnR) Or (CuDeco.CJPTNJSK = CudecoJskPtnDR) Then
            ' ����Ring���a�����l�������͂�150���傫���ꍇ
            If (CuDeco.CJALLMINRINC5 = -1) Or (CuDeco.CJALLMINRINC5 > 150) Then
                bJudg = False
            ' Ring���a���т������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJRINGNKJSK = -1) Or (CuDeco.CJRINGNKJSK > 150) Then
                bJudg = False
            ' ����Ring�O�a����l�������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJALLMAXRIGC5 = -1) Or (CuDeco.CJALLMAXRIGC5 > 150) Then
                bJudg = False
            ' Ring�O�a���т������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJRINGGKJSK = -1) Or (CuDeco.CJRINGGKJSK > 150) Then
                bJudg = False
            ' ����Ring���a�����l > Ring���a���т̏ꍇ
            ElseIf (CuDeco.CJALLMINRINC5 > CuDeco.CJRINGNKJSK) Then
                bJudg = False
            ' ����Ring�O�a����l > Ring�O�a����
            ElseIf (CuDeco.CJALLMAXRIGC5 < CuDeco.CJRINGGKJSK) Then
                bJudg = False
            End If
        End If
    End If
    
    ' CJ Disk���a�̔���
    If bJudg Then
        '' �p�^�[�����т�[Disk] or [Disk & Ring]
        If (CuDeco.CJPTNJSK = CudecoJskPtnD) Or (CuDeco.CJPTNJSK = CudecoJskPtnDR) Then
            ' ����Disk���a����l�������͂�150���傫���ꍇ
            If (CuDeco.CJALLMAXDIC5 = -1) Or (CuDeco.CJALLMAXDIC5 > 150) Then
                bJudg = False
            ' Disk���a���т������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJDISKJSK = -1) Or (CuDeco.CJDISKJSK > 150) Then
                bJudg = False
            ' ����Disk���a����l < Disk���a����
            ElseIf (CuDeco.CJALLMAXDIC5 < CuDeco.CJDISKJSK) Then
                bJudg = False
            End If
        End If
    End If

    'CJ �v�ZPi���̔���(����l�`�F�b�N)
    If bJudg Then
        '' �p�^�[�����т�[Disk] or [Ring] or [Disk & Ring]
        If (CuDeco.CJPTNJSK = CudecoJskPtnD) Or (CuDeco.CJPTNJSK = CudecoJskPtnR) Or (CuDeco.CJPTNJSK = CudecoJskPtnDR) Then
            If (CuDeco.CJPTNJSK = CudecoJskPtnD) Then        '[Disk]
                intMax = CuDeco.CJDMAXPIC5                  'Disk�̂݃p�^�[�� Pi������l
            ElseIf (CuDeco.CJPTNJSK = CudecoJskPtnR) Then    '[Ring]
                intMax = CuDeco.CJRMAXPIC5                  'Ring�̂݃p�^�[�� Pi������l
            Else                                            '[Disk & Ring]
                intMax = CuDeco.CJDRMAXPIC5                 'DiskRing�p�^�[�� Pi������l
            End If
            
            'Pi������l�������͂�150���傫���ꍇ
            If (intMax = -1) Or (intMax > 150) Then
                bJudg = False
            'Pi���v�Z�������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJPICALC = -1) Or (CuDeco.CJPICALC > 150) Then
                bJudg = False
            ''Pi������l < Pi���v�Z�̏ꍇ
            ElseIf (intMax < CuDeco.CJPICALC) Then
                bJudg = False
            End If
        End If
    End If

    CuDeco.JudgCJ = bJudg

    CrystalCJJudg = FUNCTION_RETURN_SUCCESS
    
End Function

'�T�v      :����CJLT������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Cudeco        ,I  ,C_CUDECO         ,����Cu-deco����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :
Public Function CrystalCJLTJudg(CuDeco As C_CJLT, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns         As FUNCTION_RETURN
    Dim ptnOK(6)        As String
    Dim bJudg           As Boolean

    '' �p�^�[��������
    '' �P�����ڂ̓p�^�[���敪�A�Q�����ȍ~���n�j�Ƃ���o�^�[�����т�ݒ肷��
    ''
    ''�@    �o�^�[�����т̎��
    ''          "1" : �����O�����w��    "2" : �f�B�X�N�����w��      "3" : �p�^�[�������w��
    ''          "4" : �s�� (�I���Ȃ�)   "5" : �o���h�����w��        "6" : P�o���h�����w��
    ''          "7" : B�o���h�w�薳��
    ''
    ''�@    �o�^�[�����т̎��
    ''          "0":None    "1":Ring    "2":Disk    "3":Disk & Ring
    ''          "5":PB-band "6":P-band  "7":B-band
    ptnOK(0) = "1 0567"         '' �����O�����w��
    ptnOK(1) = "2 0567"         '' �f�B�X�N�����w��
    ptnOK(2) = "3 0"            '' �p�^�[�������w��
    ptnOK(3) = "4 0567"         '' �s�� (�I���Ȃ�)
    ptnOK(4) = "5 0"            '' �o���h�����w��
    ptnOK(5) = "6 07"           '' �����O�����w��
    ptnOK(6) = "7 06"           '' B�o���h�w�薳��

    ''�G���[���\���̏�����
    CrystalCJLTJudg = SetErrInfo(ErrInfo)
    
    '*** �p�^�[������ ***
    bJudg = CudecoJudgPattern(CuDeco.HSXCJLTPK, CuDeco.CJLTPTNJSK, ptnOK)

    ' CJ(LT) �v�ZBand���̔���
    If bJudg Then
        '' �p�^�[�����т�[PB-band] or [P-band] or [B-band]
        ''Del Start 2011/05/13 Y.Hitomi �S�p�^�[�����ʉ�
'        If (CuDeco.CJLTPTNJSK = CudecoJskPtnPB_B) Or (CuDeco.CJLTPTNJSK = CudecoJskPtnP_B) Or (CuDeco.CJLTPTNJSK = CudecoJskPtnB_B) Then
        ''Del End   2011/05/13 Y.Hitomi
            ' Band������l�������͂�150���傫���ꍇ
            If (CuDeco.HSXCJLTBND = -1) Or (CuDeco.HSXCJLTBND > 150) Then
                bJudg = False
            ' Band�O�a���т������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJLTBANDGKJSK = -1) Or (CuDeco.CJLTBANDGKJSK > 150) Then
                bJudg = False
            ' Band���a���т������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJLTBANDNKJSK = -1) Or (CuDeco.CJLTBANDNKJSK > 150) Then
                bJudg = False
            ' Band�O�a���� < Band���a���т̏ꍇ
            ElseIf (CuDeco.CJLTBANDGKJSK < CuDeco.CJLTBANDNKJSK) Then
                bJudg = False
            ' Band������l < (Band�O�a����-Band���a����)�̏ꍇ
            ElseIf (CuDeco.HSXCJLTBND < (CuDeco.CJLTBANDGKJSK - CuDeco.CJLTBANDNKJSK)) Then
                bJudg = False
            End If
'        End If
    End If

    CuDeco.JudgCJLT = bJudg

    CrystalCJLTJudg = FUNCTION_RETURN_SUCCESS
    
End Function

'�T�v      :����CJ2������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Cudeco        ,I  ,C_CUDECO         ,����Cu-deco����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :
Public Function CrystalCJ2Judg(CuDeco As C_CJ2, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns         As FUNCTION_RETURN
    Dim ptnOK(3)        As String
    Dim bJudg           As Boolean
    Dim intMin          As Integer

    '' �p�^�[��������
    '' �P�����ڂ̓p�^�[���敪�A�Q�����ȍ~���n�j�Ƃ���o�^�[�����т�ݒ肷��
    ''
    ''�@    �o�^�[�����т̎��
    ''          "1" : �����O�����w��    "2" : �f�B�X�N�����w��      "3" : �p�^�[�������w��
    ''          "4" : �s�� (�I���Ȃ�)   "5" : �o���h�����w��        "6" : P�o���h�����w��
    ''          "7" : B�o���h�w�薳��
    ''
    ''�@    �o�^�[�����т̎��
    ''          "0":None    "1":Ring    "2":Disk    "3":Disk & Ring
    ''          "5":PB-band "6":P-band  "7":B-band
    ptnOK(0) = "1 02"        '' �����O�����w��
    ptnOK(1) = "2 01"        '' �f�B�X�N�����w��
    ptnOK(2) = "3 0"         '' �p�^�[�������w��
    ptnOK(3) = "4 0123"      '' �s�� (�I���Ȃ�)
    
    ''�G���[���\���̏�����
    CrystalCJ2Judg = SetErrInfo(ErrInfo)
    
    '*** �p�^�[������ ***
    bJudg = CudecoJudgPattern(CuDeco.HSXCJ2PK, CuDeco.CJ2PTNJSK, ptnOK)

    ' CJ2 Ring���a�E�O�a�̔���
    If bJudg Then
        '' �p�^�[�����т�[Ring]
        If (CuDeco.CJ2PTNJSK = CudecoJskPtnR) Then
            ' Ring�̂݃p�^�[�� Ring���a�����l�������͂�150���傫���ꍇ
            If (CuDeco.CJ2RMINRINC5 = -1) Or (CuDeco.CJ2RMINRINC5 > 150) Then
                bJudg = False
            ' Ring���a���т������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJ2RINGNKJSK = -1) Or (CuDeco.CJ2RINGNKJSK > 150) Then
                bJudg = False
            ' Ring�̂݃p�^�[�� Ring�O�a����l�������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJ2RMAXRIGC5 = -1) Or (CuDeco.CJ2RMAXRIGC5 > 150) Then
                bJudg = False
            ' Ring�O�a���т������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJ2RINGGKJSK = -1) Or (CuDeco.CJ2RINGGKJSK > 150) Then
                bJudg = False
            ' Ring�̂݃p�^�[�� Ring���a�����l > Ring���a���т̏ꍇ
            ElseIf (CuDeco.CJ2RMINRINC5 > CuDeco.CJ2RINGNKJSK) Then
                bJudg = False
            ' Ring�̂݃p�^�[�� Ring�O�a����l > Ring�O�a���т̏ꍇ
            ElseIf (CuDeco.CJ2RMAXRIGC5 < CuDeco.CJ2RINGGKJSK) Then
                bJudg = False
            End If
        '' �p�^�[�����т�[Disk & Ring]
        ElseIf (CuDeco.CJ2PTNJSK = CudecoJskPtnDR) Then
            ' DiskRing�p�^�[�� Ring���a�����l�������͂�150���傫���ꍇ
            If (CuDeco.CJ2DRMINRINC5 = -1) Or (CuDeco.CJ2DRMINRINC5 > 150) Then
                bJudg = False
            ' Ring���a���т������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJ2RINGNKJSK = -1) Or (CuDeco.CJ2RINGNKJSK > 150) Then
                bJudg = False
            ' DiskRing�p�^�[�� Ring�O�a����l�������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJ2DRMAXRIGC5 = -1) Or (CuDeco.CJ2DRMAXRIGC5 > 150) Then
                bJudg = False
            ' Ring�O�a���т������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJ2RINGGKJSK = -1) Or (CuDeco.CJ2RINGGKJSK > 150) Then
                bJudg = False
            ' DiskRing�p�^�[�� Ring���a�����l > Ring���a���т̏ꍇ
            ElseIf (CuDeco.CJ2DRMINRINC5 > CuDeco.CJ2RINGNKJSK) Then
                bJudg = False
            ' DiskRing�p�^�[�� Ring�O�a����l < Ring�O�a����
            ElseIf (CuDeco.CJ2DRMAXRIGC5 < CuDeco.CJ2RINGGKJSK) Then
                bJudg = False
            End If
        End If
    End If

    'CJ2 �v�ZPi���̔���(�����l�`�F�b�N)
    If bJudg Then
        '' �p�^�[�����т�[Disk] or [Ring] or [Disk & Ring]
        If (CuDeco.CJ2PTNJSK = CudecoJskPtnD) Or (CuDeco.CJ2PTNJSK = CudecoJskPtnR) Or (CuDeco.CJ2PTNJSK = CudecoJskPtnDR) Then
            If (CuDeco.CJ2PTNJSK = CudecoJskPtnD) Then       '[Disk]
                intMin = CuDeco.CJ2DMAXPIC5     ' Disk�̂݃p�^�[�� Pi������l(�����l�Ƃ��Ďg�p)
            ElseIf (CuDeco.CJ2PTNJSK = CudecoJskPtnR) Then   '[Ring]
                intMin = CuDeco.CJ2RMAXPIC5     ' Ring�̂݃p�^�[�� Pi������l(�����l�Ƃ��Ďg�p)
            Else                                            '[Disk & Ring]
                intMin = CuDeco.CJ2DRMAXPIC5    ' DiskRing�p�^�[�� Pi������l(�����l�Ƃ��Ďg�p)
            End If
            
            'Pi�������l�������͂�150���傫���ꍇ
            If (intMin = -1) Or (intMin > 150) Then
                bJudg = False
            'Pi���v�Z�������͂�150���傫���ꍇ
            ElseIf (CuDeco.CJ2PICALC = -1) Or (CuDeco.CJ2PICALC > 150) Then
                bJudg = False
            'Pi�������l > Pi���v�Z�̏ꍇ
            ElseIf (intMin > CuDeco.CJ2PICALC) Then
                bJudg = False
            End If
        End If
    End If

    CuDeco.JudgCJ2 = bJudg

    CrystalCJ2Judg = FUNCTION_RETURN_SUCCESS
    
End Function

'Add End   2011/01/24 SMPK Miyata

'�T�v      :�ΏۃR�[�h�ɏ]���Ĕ���Ώۃf�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Flag          ,I  ,GUARANTEE ,�ΏۃR�[�h
'          :d()           ,I  ,double    ,����l
'          :d1()          ,O  ,double    ,����Ώۃf�[�^
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :Flag.cObj�̒l,
'          :1       ,d1(0)=���S����l
'          :2       ,d1(0)=��������l
'          :3       ,d1()=�S����_
'          :4       ,d1(0)=R/2
'          :A       ,d1(0)=���ϒl
'          :B       ,d1(0)=�ő�l
'          :C       ,d1(0)=���ϒl,d1(1)=�ő�l
'          :D       ,d1(0)=�ŏ��l
'          :E       ,d1(0�`3)=������2�_�A�O����2�_(5�_�����1,2,4,5)
'          :F       ,d1(0)=2,4�_�ڂ̓��傫���l
'          :G       ,d1(0)=2,3,4�_�ڂ̓��傫���l
'����      :2001/06/06 ���� �M�� �쐬
Public Function GetCrystalJudgData(flag As Guarantee, d() As Double, d1() As Double) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim COUNT As Integer
    Dim High As Integer
    
    '' �z��̏�����擾���܂��B
    High = UBound(d)
    
    FuncAns = FUNCTION_RETURN_SUCCESS '' ����
    Select Case flag.cObj
    Case ObjCode01 ''���S����l
        d1(0) = d(0)
    Case ObjCode02 ''����l�̒����l
        d1(0) = JudgCenter(d())
    Case ObjCode03 ''�S����_
        DataCopy d(), d1()
    Case ObjCode04 ''R/2
        Select Case flag.cPos
        Case PosCode01
            If flag.cCount = "1" Then
                d1(0) = d(0)
            Else
                d1(0) = d(1)
            End If
        Case PosCode02, PosCode03, PosCode04, PosCode05, PosCode06, PosCode07, PosCode08
            d1(0) = d(1)
        Case PosCode09
            d1(0) = d(2)
        Case Else
            FuncAns = FUNCTION_RETURN_FAILURE '' �ُ�
        End Select
    Case ObjCode05 ''�S�_�̕��ϒl
        d1(0) = JudgAve(d())
    Case ObjCode06 ''�S�_�̍ő�l
        d1(0) = JudgMax(d())
    Case ObjCode07 ''�S�_�̕��ϒl�ƍő�l
        d1(0) = JudgAve(d())
        d1(1) = JudgMax(d())
    Case ObjCode08 ''�S�_�̍ŏ��l
        d1(0) = JudgMin(d())
    Case ObjCode09 ''������2�_�A�O����2�_(5�_�����1,2,4,5)
        DataCopy d(), d1()
        COUNT = 0
        For c0 = High To 0 Step -1
            If d1(c0) <> -1 Then
                d1(3 - COUNT) = d1(c0)
                COUNT = COUNT + 1
            End If
            If COUNT = 2 Then Exit For
        Next
    Case ObjCode10 ''MAX(2,4�_��)
        If (d(1) <> -1) And (d(3) <> -1) Then
            If d(1) >= d(3) Then
                d1(0) = d(1)
            Else
                d1(0) = d(3)
            End If
        Else
            FuncAns = FUNCTION_RETURN_FAILURE '' �ُ�
        End If
    Case ObjCode11 ''MAX(2,3,4�_��)
        If (d(1) <> -1) And (d(2) <> -1) And (d(3) <> -1) Then
            If d(1) >= d(2) Then
                If d(1) >= d(3) Then
                    d1(0) = d(1)
                Else
                    d1(0) = d(3)
                End If
            Else
                If d(2) >= d(3) Then
                    d1(0) = d(2)
                Else
                    d1(0) = d(3)
                End If
            End If
        Else
            ''���蓾�Ȃ��G���[
            FuncAns = FUNCTION_RETURN_FAILURE '' �ُ�
        End If
''    Case ObjCode12 ''���ۏ�
''    Case ObjCode13 ''�_��
''    Case ObjCode14 ''�`�󑪒�(���R�x�A���Ԃ�AWARP)
''    Case ObjCode15 ''�K�i�Ȃ�
    Case Else
        FuncAns = FUNCTION_RETURN_FAILURE '' �ُ�
    End Select
    
    GetCrystalJudgData = FuncAns
End Function

'�T�v      :����Oi������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Oi            ,I  ,C_Oi             ,����Oi����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/07/19 ���� �M�� �쐬
Public Function CrystalOiJudg(Oi As C_Oi, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim JData() As Double
    Dim c0 As Integer
    Dim pt As Integer
    
    ReDim JData(UBound(Oi.Oi())) As Double
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Oi.JudgData = -1
    Oi.JudgOi = JUDG_NG
    If Oi.GuaranteeOi.cJudg = JudgCodeC01 Then ''Oi����L��
        
'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06 start

'    If Trim(Oi.GuaranteeOi.cCount) = "" Then
'        pt = 1
'    Else
'        pt = Val(Oi.GuaranteeOi.cCount)
'    End If
'    Oi.ORG = RoundUp((RGCal(Oi.Oi(), pt)), 2)

        ''ORG����
        
        Select Case Oi.GuaranteeOi.cPos
          Case "B", "C", "D", "E", "F", "K", "Y"
              Select Case Oi.GuaranteeOi.cBunp
              Case "A", "B", "C"
                 ''ORG�v�Z
                 Oi.ORG = MENNAI_Cal(OI_JUDG, Oi.Oi(), Oi.GuaranteeOi, Oi.GuaranteeOi.cBunp)

              Case "", " "
                 ''�v�Z�敪���X�y�[�X�̏ꍇ�́A�v�Z�C������s��Ȃ�
                  Oi.ORG = 0
                  Oi.JudgOrg = JUDG_OK
                  GoTo Cal_Escp
              Case Else
                 ''ORG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
                 If Trim(Oi.GuaranteeOi.cCount) = "" Then
                    pt = 3
                 Else
                    pt = val(Oi.GuaranteeOi.cCount)
                 End If
                 Oi.ORG = RoundUp((RGCal(Oi.Oi(), pt)), 4)

             End Select

          Case Else

             Select Case Oi.GuaranteeOi.cBunp
             Case "A", "B", "C", "D", "E", "N"
                 ''ORG�v�Z
                 Oi.ORG = MENNAI_Cal(OI_JUDG, Oi.Oi(), Oi.GuaranteeOi, Oi.GuaranteeOi.cBunp)

             Case "", " "
                 ''�v�Z�敪���X�y�[�X�̏ꍇ�́A�v�Z�C������s��Ȃ�
                  Oi.ORG = 0
                  Oi.JudgOrg = JUDG_OK
                  GoTo Cal_Escp
             Case Else
                 ''ORG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
                 If Trim(Oi.GuaranteeOi.cCount) = "" Then
                    pt = 3
                 Else
                    pt = val(Oi.GuaranteeOi.cCount)
                 End If
                 Oi.ORG = RoundUp((RGCal(Oi.Oi(), pt)), 4)

             End Select
        End Select

'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06 end

'2002/02/27 S.Sano ORG�̎d�l��0�̏ꍇ�́A������s�킸�K��OK�Ƃ���B
'2002/02/27 S.Sano �ʓ����z�v�Z�͍s���B
        If Oi.SpecORG = 0 Then                                      '2002/02/27 S.Sano
            Oi.JudgOrg = JUDG_OK                                    '2002/02/27 S.Sano
        Else                                                        '2002/02/27 S.Sano
            If Oi.ORG = -1 Then
                Oi.JudgOrg = JUDG_NG
            Else
                Oi.JudgOrg = RangeDecision_nl(Oi.ORG, 0, Oi.SpecORG)
            End If
        End If                                                      '2002/02/27 S.Sano
        
'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06 start
Cal_Escp:
'' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06 end
        
        ''Oi����
        If (InStr(ObjCodeGrp01, Oi.GuaranteeOi.cObj) <> 0) And (GetCrystalJudgData(Oi.GuaranteeOi, Oi.Oi(), JData()) = FUNCTION_RETURN_SUCCESS) Then
            Select Case Oi.GuaranteeOi.cObj
            Case ObjCode01, ObjCode02, ObjCode04 ''���S1�_�A�����l�AR/2
                Oi.JudgOi = RangeDecision_nl(JData(0), Oi.SpecOiMin, Oi.SpecOiMax)
                Oi.JudgData = JData(0)
            Case ObjCode03 ''�S��
                Oi.JudgOi = JUDG_OK
                For c0 = 0 To UBound(JData())
                    If JData(c0) <> -1 Then
                        If RangeDecision_nl(JData(c0), Oi.SpecOiMin, Oi.SpecOiMax) = JUDG_NG Then
                            Oi.JudgOi = JUDG_NG
                        End If
                    End If
                Next
            End Select
            Oi.JudgData = JudgMax(JData())
        Else
            ''�Ώۃf�[�^����
            ''�G���[���\���̂ɏ������B
            FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
        End If
        Oi.JudgOi = (Oi.JudgOi And Oi.JudgOrg)
    Else
        Oi.JudgOrg = JUDG_OK
        Oi.JudgOi = JUDG_OK
'        If InStr(JudgCodeC02, Oi.GuaranteeOi.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, OI_JUDG, Oi.GuaranteeOi.cJudg)
'        End If
    End If
    
    CrystalOiJudg = FuncAns
End Function

'�T�v      :����Cs������s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :Cs            ,I  ,C_Cs             ,����Cs����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :2001/07/19 ���� �M�� �쐬
Public Function CrystalCsJudg(Cs As C_Cs, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''�G���[���\���̏�����
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    If Cs.GuaranteeCs.cJudg = JudgCodeC01 Then ''Cs����L��
        'BOT�ۏ؂͏������,TOP/BOT�ۏ؂͏㉺������ 09/01/08 ooba
        If Cs.SpecCsKHI = "6" Or Cs.SpecCsKHI = "9" Then
            Cs.JudgCs = RangeDecision_nl(Cs.Cs, Cs.SpecCsMin, Cs.SpecCsMax)
        Else
            Cs.JudgCs = RangeDecision_nl(Cs.Cs, -1, Cs.SpecCsMax)
        End If
    Else
        Cs.JudgCs = JUDG_OK
'        If InStr(JudgCodeC02, Cs.GuaranteeCs.cJudg) = 0 Then
'            ''�������@�f�[�^����
'            ''�G���[���\���̂ɏ������B
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, CS_JUDG, Cs.GuaranteeCs.cJudg)
'        End If
    End If
    
    CrystalCsJudg = FuncAns
End Function


#If NO_FURIKAECHECK = 0 Then
'�T�v      :�����ȍ~�ł̕i�ԐU�֎��ɁA�������̔�����s��
'���Ұ�    :�ϐ���        ,IO ,�^          ,����
'          :crynum        ,I  ,String      ,�����ԍ�
'          :ingotpos      ,I  ,Integer     ,�Ώ۔͈͂̊J�n�ʒu
'          :length        ,I  ,Integer     ,�Ώ۔͈͂̒���
'          :hin           ,I  ,tFullHinban ,�U�֐�̕i��
'          :judge_ok      ,O  ,Boolean     ,���茋��
'          :itemNG        ,O  ,String      ,����NG�ƂȂ�������
'          :�߂�l        ,O  ,FUNCTION_RETURN, ����̍���
'          :                   FUNCTION_RETURN_SUCCESS: �U�։�
'          :                   FUNCTION_RETURN_FAILURE: �U�֕s�������͎d�l�G���[
'����      :�����ۏ؂݂̂̍��ڂł��� GD/LT/Cs �ɂ��Ĕ��肷��
'����      :2002/03/xx ���� �M�� �쐬
Public Function SXLJudge(CRYNUM$, INGOTPOS%, Length%, HIN As tFullHinban, judge_ok As Boolean, itemNG$) As FUNCTION_RETURN
    Dim GD(1) As C_GD
    Dim Cs(1) As C_Cs
    Dim Lt(1) As C_LT
    Dim ErrInfo As ERROR_INFOMATION
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzccj.bas -- Function SXLJudge"

    SXLJudge = FUNCTION_RETURN_FAILURE
    If scmzc_getSXLGuarantee(HIN, GD(), Cs(), Lt()) = FUNCTION_RETURN_FAILURE Then
        '�d�l�擾�G���[
        GoTo proc_exit
    End If
    
    judge_ok = False
    itemNG$ = ""
    
    'GD����
    '�������т��擾����
    If scmzc_getSXLGD(CRYNUM$, INGOTPOS%, Length%, GD()) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'Top�ʒu�̔�����s��
    If CrystalGDJudg(GD(0), ErrInfo) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    'Bot�ʒu�̌������т��擾����
    ElseIf CrystalGDJudg(GD(1), ErrInfo) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'NG�Ȃ甲����
    SXLJudge = FUNCTION_RETURN_SUCCESS
    If (Not GD(0).JudgDen) Or (Not GD(1).JudgDen) Then
        itemNG$ = "DEN"
        GoTo proc_exit
    ElseIf (Not GD(0).JudgDvd2) Or (Not GD(1).JudgDvd2) Then
        itemNG$ = "DVD2"
        GoTo proc_exit
    ElseIf (Not GD(0).JudgLdl) Or (Not GD(1).JudgLdl) Then
        itemNG$ = "L/DL"
        GoTo proc_exit
    End If

    'Cs����
    '�������т��擾����
    SXLJudge = FUNCTION_RETURN_FAILURE
    If scmzc_getSXLCs(CRYNUM$, INGOTPOS%, Length%, Cs()) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'Top�ʒu�̔�����s��
    If (Cs(0).GuaranteeCs.cJudg = "H") And (Cs(0).SpecCsMin > 0#) Then
        'Cs��FromTo�ۏ؂̏ꍇ�́ATop��������s��
        If CrystalCsJudg(Cs(0), ErrInfo) = FUNCTION_RETURN_FAILURE Then
            GoTo proc_exit
        End If
    Else
        Cs(0).JudgCs = True
    End If

    'Bot�ʒu�̔�����s��
    If CrystalCsJudg(Cs(1), ErrInfo) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'NG�Ȃ甲����
    SXLJudge = FUNCTION_RETURN_SUCCESS
    If (Not Cs(0).JudgCs) Or (Not Cs(1).JudgCs) Then
        itemNG$ = "CS"
        GoTo proc_exit
    End If
    
    'LT����
    '�������т��擾����
    SXLJudge = FUNCTION_RETURN_FAILURE
    If scmzc_getSXLLt(CRYNUM$, INGOTPOS%, Length%, Lt()) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'Bot�ʒu�̔�����s��(LT��Bot�������̕ۏ�)
    Lt(0).JudgLt = True
    If CrystalLTJudg(Lt(1), ErrInfo) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'NG�Ȃ甲����
    SXLJudge = FUNCTION_RETURN_SUCCESS
    If (Not Lt(1).JudgLt) Then
        itemNG$ = "LT"
        GoTo proc_exit
    End If

    judge_ok = True

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WF���ł̕i�ԐU�֎��`�F�b�N���s��
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :orgSXL        ,I  ,c_cmzcSxls   ,�U�֑O��SXL�\��
'          :wfSmps()      ,I  ,typ_XSDCW    ,�V�T���v���Ǘ��iSXL�j
'          :Crynum        ,I  ,String       ,�����ԍ�
'          :lblMsg        ,I  ,Label        ,���b�Z�[�W�\���G���A
'          :needPreJudge  ,I  ,Boolean      ,���̑��̌�����������s��
'          :chkFrom       ,I  ,Integer      ,�`�F�b�N�͈�(mm)
'          :chkTo         ,I  ,Integer      ,�`�F�b�N�͈�(mm)
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :�`�F�b�N�Ώۂ́ACs,GD,LT
'����      :
'Public Function FurikaeCheck(orgSXL As c_cmzcSxls, WfSmps() As typ_XSDCW, CRYNUM$, lblMsg As Label, needPreJudge As Boolean, chkFrom As Integer, chkTo As Integer) As FUNCTION_RETURN
Public Function FurikaeCheck(orgSXL As c_cmzcSxls, WfSmps() As typ_XSDCW, CRYNUM$, lblMsg As Label, chkFrom As Integer, chkTo As Integer) As fHinban
    Dim HIN As tFullHinban '2002/03/14 S.Sano
    Dim c0 As Integer
    Dim c1 As Integer
    Dim judge_ok As Boolean
    Dim itemNG$
    Dim eqf As Boolean
    Dim hinban$
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim nSxl As Integer
    Dim fHin As fHinban

''    FurikaeCheck = FUNCTION_RETURN_FAILURE

    ReDim buff$(0)
'    For c0 = 1 To UBound(WfSmps) - 1
        pos1 = WfSmps(c0).INPOSCW
        pos2 = WfSmps(c0 + 1).INPOSCW
        If (pos1 >= chkFrom) And (pos2 <= chkTo) Then
            hinban = Trim(WfSmps(c0).HINBCW)
            If (hinban <> "Z") And (hinban <> "G") And (hinban <> vbNullString) Then
                '�i�Ԃ��ς���Ă��Ȃ���΁A�X�L�b�v����B
                eqf = True
                If (WfSmps(c0).SMPKBNCW = "U") Or (WfSmps(c0).SMPKBNCW = "B") Then
                    nSxl = orgSXL.UpperArea(pos1)
                Else
                    nSxl = orgSXL.LowerArea(pos1)
                End If
                If Abs(nSxl) <> 9999 Then
                    If hinban <> orgSXL(CStr(nSxl)).hinban Then
'                        eqf = False
                '�\���̂ɒl��ێ�����
                fHin.moto = hinban
                fHin.saki = orgSXL(CStr(nSxl)).hinban
                    End If
                End If

        '�i�Ԃ̐U�ւ͊֐��ōs�����ߐU�փ`�F�b�N�̋@�\���폜-------start iida 2003/09/05
'                If Not eqf Then
'                    If GetLastHinban(HINBAN$, hin) = FUNCTION_RETURN_FAILURE Then
'                        lblMsg.Caption = GetMsgStr("EHIN8", vbNullString) '03/06/06 �㓡
'                        Exit Function
'                    End If

'                    '��{����
'                    If needPreJudge Then
'                        If SXLPreJudge(CRYNUM$, pos1, pos2 - pos1, hin, judge_ok, itemNG$) = FUNCTION_RETURN_FAILURE Then
'                            lblMsg.Caption = GetMsgStr("EHIN8", "(" & itemNG & ")")   '03/06/06 �㓡
'                            Exit Function
'                        End If
'                        If Not judge_ok Then
'                        lblMsg.Caption = GetMsgStr("EHIN9", Trim(HINBAN$) & " " & itemNG$)    '03/06/06 �㓡
'                            Exit Function
'                        End If
'                    End If

                    'GD/Cs/LT����
'                    If SXLJudge(CRYNUM, pos1, pos2 - pos1, hin, judge_ok, itemNG$) = FUNCTION_RETURN_FAILURE Then
'                        lblMsg.Caption = GetMsgStr("EHIN8", vbNullString)   '03/06/06 �㓡
'                        Exit Function
'                    End If
'                    If Not judge_ok Then
'                        lblMsg.Caption = GetMsgStr("EHIN9", Trim(HINBAN$) & " " & itemNG$)        '03/06/06 �㓡
'                        Exit Function
'                    End If
'                End If
        '�i�ԐU�ւ͊֐��ōs�����ߐU�փ`�F�b�N�̋@�\���폜-------end iida 2003/09/05
            End If
        End If
'    Next
''    FurikaeCheck = fHinban
End Function
#End If

'�T�v      :���H���т̔�����s��
'���Ұ�    :�ϐ���        ,IO ,�^              ,����
'          :Kakou         ,IO ,type_KakouJudg  ,���H���є���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN, ����̐���
'����      :���H���тɂ��Ĕ��肷��
'����      :2002/04/17 ���� �M�� �쐬
Public Function FormJudg(Kakou As type_KakouJudg) As FUNCTION_RETURN
    Dim c0 As Integer
    
    FormJudg = FUNCTION_RETURN_FAILURE
    
    Kakou.Judg.top = False
    Kakou.Judg.tTOP(1) = False
    Kakou.Judg.tTOP(2) = False
    
    Kakou.Judg.TAIL = False
    Kakou.Judg.tTAIL(1) = False
    Kakou.Judg.tTAIL(2) = False
    
    Kakou.Judg.POS = False
    Kakou.Judg.WIDH = False
    Kakou.Judg.DPTH = False
    Kakou.Judg.ANGLE = False        '2009/09 SUMCO Akizuki
    
    
    ''Notch�ʒu�̋K�i����
    If InStr("A1", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("A1", Kakou.Spec())
    ElseIf InStr("A2", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("A2", Kakou.Spec())
    ElseIf InStr("B1B2B3B4", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("B1B2B3B4", Kakou.Spec())
    ElseIf InStr("B5B6B7B8", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("B5B6B7B8", Kakou.Spec())
    ElseIf InStr("C1", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("C1", Kakou.Spec())
    ElseIf InStr("C2", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("C2", Kakou.Spec())
    ElseIf InStr("D1D2", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("D1D2", Kakou.Spec())
    ElseIf InStr("D3D4", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("D3D4", Kakou.Spec())
    ElseIf InStr("D5D8", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("D5D8", Kakou.Spec())
    ElseIf InStr("D6D7", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("D6D7", Kakou.Spec())
    ElseIf InStr("ZZ", Kakou.Jiltuseki.POS) <> 0 Then      ''''2005/05/27 ADD
        Kakou.Judg.POS = Kakou_Pos_Judg("ZZ", Kakou.Spec())

    Else
        Exit Function
    End If
    
    
    '' ���a(TOP1,2)�̋K�i�`�F�b�N
    If (Kakou.Jiltuseki.top(1) = -1) And (Kakou.Jiltuseki.top(2) = -1) Then
        Exit Function
    End If
    Kakou.Judg.tTOP(1) = True
    Kakou.Judg.tTOP(2) = True
    
    '�e�i�Ԃ��ƂɁA�d�l�K�i�̃`�F�b�N���s��
    For c0 = 1 To UBound(Kakou.Spec())
        If Kakou.Spec(c0).top(1) = -1 Or Kakou.Spec(c0).top(2) = -1 Then Exit Function      '2003/12/12 SystemBrain Null�Ή�
        If Kakou.Jiltuseki.top(1) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.top(1), Kakou.Spec(c0).top(1), Kakou.Spec(c0).top(2)) = False Then
                Kakou.Judg.tTOP(1) = False
            End If
        End If
        If Kakou.Jiltuseki.top(2) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.top(2), Kakou.Spec(c0).top(1), Kakou.Spec(c0).top(2)) = False Then
                Kakou.Judg.tTOP(2) = False
            End If
        End If
    Next
    If (Kakou.Jiltuseki.top(1) <> -1) And (Kakou.Jiltuseki.top(2) <> -1) Then
        Kakou.Judg.top = (Kakou.Judg.tTOP(1) And Kakou.Judg.tTOP(2))
    ElseIf Kakou.Jiltuseki.top(1) <> -1 Then
        Kakou.Judg.top = Kakou.Judg.tTOP(1)
    ElseIf Kakou.Jiltuseki.top(2) <> -1 Then
        Kakou.Judg.top = Kakou.Judg.tTOP(2)
    End If
    
    
    
    '' ���a(BOT1,2)�̋K�i�`�F�b�N
    If (Kakou.Jiltuseki.TAIL(1) = -1) And (Kakou.Jiltuseki.TAIL(2) = -1) Then
        Exit Function
    End If
    
    
    Kakou.Judg.tTAIL(1) = True
    Kakou.Judg.tTAIL(2) = True
    
    For c0 = 1 To UBound(Kakou.Spec())
        If Kakou.Spec(c0).top(1) = -1 Or Kakou.Spec(c0).top(2) = -1 Then Exit Function      '2003/12/12 SystemBrain Null�Ή�
        If Kakou.Jiltuseki.TAIL(1) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.TAIL(1), Kakou.Spec(c0).top(1), Kakou.Spec(c0).top(2)) = False Then
                Kakou.Judg.tTAIL(1) = False
            End If
        End If
        If Kakou.Jiltuseki.TAIL(2) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.TAIL(2), Kakou.Spec(c0).top(1), Kakou.Spec(c0).top(2)) = False Then
                Kakou.Judg.tTAIL(2) = False
            End If
        End If
    Next
    If (Kakou.Jiltuseki.TAIL(1) <> -1) And (Kakou.Jiltuseki.TAIL(2) <> -1) Then
        Kakou.Judg.TAIL = (Kakou.Judg.tTAIL(1) And Kakou.Judg.tTAIL(2))
    ElseIf Kakou.Jiltuseki.TAIL(1) <> -1 Then
        Kakou.Judg.TAIL = Kakou.Judg.tTAIL(1)
    ElseIf Kakou.Jiltuseki.TAIL(2) <> -1 Then
        Kakou.Judg.TAIL = Kakou.Judg.tTAIL(2)
    End If
    
    If (Kakou.Jiltuseki.WIDH(1) = -1) And (Kakou.Jiltuseki.WIDH(2) = -1) Then
        Exit Function
    End If
    
    
    ''Notch���̋K�i�`�F�b�N
    Kakou.Judg.WIDH = True
    
    For c0 = 1 To UBound(Kakou.Spec())
        If Kakou.Spec(c0).WIDH(1) = -1 Or Kakou.Spec(c0).WIDH(2) = -1 Then Exit Function      '2003/12/12 SystemBrain Null�Ή�
        If Kakou.Jiltuseki.WIDH(1) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.WIDH(1), Kakou.Spec(c0).WIDH(1), Kakou.Spec(c0).WIDH(2)) = False Then
                Kakou.Judg.WIDH = False
            End If
        End If
        If Kakou.Jiltuseki.WIDH(2) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.WIDH(2), Kakou.Spec(c0).WIDH(1), Kakou.Spec(c0).WIDH(2)) = False Then
                Kakou.Judg.WIDH = False
            End If
        End If
    Next
    
    If (Kakou.Jiltuseki.DPTH(1) = -1) And (Kakou.Jiltuseki.DPTH(2) = -1) Then
        Exit Function
    End If
    
    
    '' Notch�[��(TOP�BOT)�̋K�i�`�F�b�N
    Kakou.Judg.DPTH = True
    
    For c0 = 1 To UBound(Kakou.Spec())
        If Kakou.Spec(c0).DPTH(1) = -1 Or Kakou.Spec(c0).DPTH(2) = -1 Then Exit Function      '2003/12/12 SystemBrain Null�Ή�
        If Kakou.Jiltuseki.DPTH(1) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.DPTH(1), Kakou.Spec(c0).DPTH(1), Kakou.Spec(c0).DPTH(2)) = False Then
                Kakou.Judg.DPTH = False
            End If
        End If
        If Kakou.Jiltuseki.DPTH(2) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.DPTH(2), Kakou.Spec(c0).DPTH(1), Kakou.Spec(c0).DPTH(2)) = False Then
                Kakou.Judg.DPTH = False
            End If
        End If
    Next

    
    
    '' Notch�p�x�̋K�i�`�F�b�N      2009/09 SUMOCO Akizuki
    Kakou.Judg.ANGLE = True
    
    For c0 = 1 To UBound(Kakou.Spec())
        If Kakou.Spec(c0).ANGLE(1) = -1 Or Kakou.Spec(c0).ANGLE(2) = -1 Then Exit Function      '2003/12/12 SystemBrain Null�Ή�
        If Kakou.Jiltuseki.ANGLE(1) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.ANGLE(1), Kakou.Spec(c0).ANGLE(1), Kakou.Spec(c0).ANGLE(2)) = False Then
                Kakou.Judg.ANGLE = False
            End If
        End If
        If Kakou.Jiltuseki.ANGLE(2) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.ANGLE(2), Kakou.Spec(c0).ANGLE(1), Kakou.Spec(c0).ANGLE(2)) = False Then
                Kakou.Judg.ANGLE = False
            End If
        End If
    Next

    FormJudg = FUNCTION_RETURN_SUCCESS
    
End Function

'�T�v      :���H���т̈ʒu������s��
'���Ұ�    :�ϐ���        ,IO ,�^              ,����
'          :sPos          ,I  ,String         ,�ʒu�O���[�v������
'          :Spec()        ,I  ,type_KakouJudg ,���H�d�l�\����
'          :�߂�l        ,O  ,Boolean         ,���茋��
'����      :���H���шʒu�ɂ��Ĕ��肷������֐�
'����      :2002/04/17 ���� �M�� �쐬
Public Function Kakou_Pos_Judg(sPos As String, Spec() As Judg_Kakou) As Boolean
    Dim c0 As Integer
    Dim tJudg As Boolean
    tJudg = True
    For c0 = 1 To UBound(Spec())
        If tJudg Then
            tJudg = (InStr(sPos, Spec(c0).POS) <> 0)
        End If
    Next
    Kakou_Pos_Judg = tJudg
End Function
'�T�v      :�����u���b�N�ΐ͔�����s���B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :COEF          ,I  ,C_COEF            ,����\����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,
'����      :
'����      :
Public Function CrystalCOEFJudg(COEF As C_COEF, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim sMin As Double
    Dim sMax As Double
    
    ''�G���[���\���̏�����
    CrystalCOEFJudg = SetErrInfo(ErrInfo)
    Select Case COEF.NP
        Case "p-"
            sMin = PminusMin
            sMax = PminusMax
        Case "p+"
            sMin = PplusMin
            sMax = PplusMax
        Case "n"
            sMin = NMin
            sMax = NMax
    End Select
    COEF.JudgCOEF = RangeDecision_nl(COEF.COEF, sMin, sMax)
    CrystalCOEFJudg = FUNCTION_RETURN_SUCCESS
End Function

