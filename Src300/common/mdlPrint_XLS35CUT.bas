Attribute VB_Name = "mdlPrint_XLS35CUT"
Option Explicit

'******************************************************************************
' @(S)
'           ���[(�ؒf�w����)�o�̓��C��(300mm)
'
' @(h) mdlPrint_XLS35CUT.bas             ver 1.0 ( 2008.09.22 SETsw kubota )
'
'CMBC016���H���o�ACMBC030��������ACMDC101���[�Ĕ��s�œ����t�@�C�����g�p����
'(�Е��ŕύX������t�@�C�������̂܂܃R�s�[����)
'******************************************************************************

Private Enum enmSample
    CUT_RS = 0                      '��R
    CUT_OI                          'Oi
    CUT_GFA                         'GFA    2012/06/11�ǉ� SETsw kubota
    CUT_O1                          'OSF1
    CUT_O2                          'OSF2
    CUT_O3                          'OSF3
    CUT_CO3                         '2�iOS
    CUT_C                           'C      C,CJ,CJ2���� 2010/10/25�ǉ� SETsw kubota
    CUT_CJ                          'CJ
    CUT_CJ2                         'CJ2
    CUT_B1                          'BMD1
    CUT_B2                          'BMD2
    CUT_B3                          'BMD3
    CUT_GD                          'GD
    CUT_LT                          'LT
    CUT_CS                          'CS
    CUT_EPD                         'EPD
    CUT_X                           'X��
    
    CUT_MAXCNT                      '�T���v�����
End Enum

'���[����
Private Type typPrintInfo_Meisai
    sBlockNo        As String       '�u���b�NID(�����ԍ���10���ځ`)
    sZuban          As String       '�}��
    sLen            As String       '�u���b�N����
    sCutPos         As String       '�ؒf�ʒu
    sSmpl(CUT_MAXCNT - 1)   As String     '�e���荀�ڂ̃T���v���w����(��R�`EPD)
    sMaisu          As String       '�T���v������
    
    '�T���v���}(3,3)  4��(0�`3),�e1/4�ɕ�����(0:����,1:�E��,2:����,3:�E��)
    sSmplPic(3, 3)  As String
'>>>>> �T���v�����\���Ή��@2009/01/26�@Marushita
    sSmpNo          As String       '�T���v����
'<<<<< �T���v�����\���Ή��@2009/01/26�@Marushita
'>>>>> �g�b�v�E�{�g����ʑΉ��@2009/11/12�@Marushita
'>>>>> �}���`�i�ԑΉ��@2009/11/18�@Marushita
    iSmpKbnT(CUT_MAXCNT - 1)  As Integer      '�T���v���敪TOP
    iSmpKbnB(CUT_MAXCNT - 1)  As Integer      '�T���v���敪BOT
'<<<<< �}���`�i�ԑΉ��@2009/11/18�@Marushita
    sSmpNoT         As String       '�T���v����(TOP)
    sSmpNoB         As String       '�T���v����(BOT)
'<<<<< �g�b�v�E�{�g����ʑΉ��@2009/11/12�@Marushita
'Add Start 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
    sFrsFlg         As String       'FRS���
    sFrsResult      As String       'FRS����
'Add End 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�

    sNotchPos       As String       'Notch�ʒu  2012/06/08�ǉ� SETsw kubota

End Type

'���[�S��
Private Type typPrintInfo
    '�w�b�_
    sXtalNo         As String       '�����ԍ�
    sDate           As String       '���s��
    sZuban          As String       '�}��(�i��)
    sType           As String       '�`���^
    sDia            As String       '���a
    sJiku           As String       '������
    sRsKikaku       As String       '�ϋK�i
    sOiKikaku       As String       'Oi�K�i
    sNeraiRs        As String       '�˂炢��R
    sCharge         As String       '�`���[�W��
    sPgid           As String       'PG-ID
    sBottom         As String       '�{�g����
    sPulWeight      As String       '����d��
    sTopCutWeight   As String       '�g�b�v�J�b�g�d��
    sFreeLen        As String       '�t���[��
    sPulLen         As String       '���㒷��
    sKataLen        As String       '���J�b�g����
    sOiDopPos       As String       '�ǃh�[�v�ʒu
    
    '����
    tMeisai()       As typPrintInfo_Meisai
    
    '�T���v�����݂ƌ`��
    sThick(CUT_MAXCNT - 1)  As String   '����
    sShape(CUT_MAXCNT - 1)  As String   '�`��
    
    sSmplNm(CUT_MAXCNT - 1) As String   '����ٖ�
    sPicStr(CUT_MAXCNT - 1) As String   '����ِ}�\������
    
    sMaisu          As String   '�T���v���������v
    
    sBarCode        As String   '�����ԍ��o�[�R�[�h�@ADD 2009/08/05 SSS.Marushita

End Type
Private mtPrintInfo As typPrintInfo

Private Const PRINTFILENAME     As String = "�ؒf�w����"
Private Const TEMPLATENAME      As String = "XLS35CUT3"
'>>>>>���y�[�W�Ή� 2009/08/05 SSS.Marushita
Private Const CON_X             As Long = 40        '�~�̐�����
Private Const CON_Y             As Long = 40        '�~�̐�����
'Private Const CON_X             As Long = 50        '�~�̐�����
'Private Const CON_Y             As Long = 50        '�~�̐�����
'<<<<<���y�[�W�Ή� 2009/08/05 SSS.Marushita

Private Const PIC_DUMMY         As String = "*"     '1/4,1/2�ŕ�����}�[�N

'�R�[�h�}�X�^�[���y�[�W�Ǘ��f�[�^�擾�p ADD 2009/08/05 SSS.Marushita
Public lPRINTPAGEROW            As Long          '���[1�y�[�W�̍s��
Public lPRINTMEISAIROW          As Long          '1�y�[�W���̖��א�
Public lSET_MEISAI_CNT          As Long          '�}���`�i�Ԓ�����̖��א�
'Add Start 2011/03/08 SMPK Nakamura FRS�V�X�e�����ǉ��Ή�
Public Const FRSKBN_NONE    As String = "-"        ' FRS���       -:FRS�Ȃ�
Public Const FRSKBN_0       As String = "0"        '               0:�ΏۊO
Public Const FRSKBN_1       As String = "1"        '               1:�]��
Public Const FRSKBN_2       As String = "2"        '               2:���p
Public Const FRSRSL_0       As String = "0"        ' FRS����       0:������
Public Const FRSRSL_1       As String = "1"        '               1:����OK
Public Const FRSRSL_2       As String = "2"        '               2:����NG
Public Const FRSRSL_3       As String = "3"        '               3:�Ĕ���
Public Const FRSRSL_4       As String = "4"        '               4:�����
Public Const FRSKBN_0_NAME  As String = "�ΏۊO"   ' FRS�敪����   �ΏۊO(���FLG[FRS]�F0)
Public Const FRSKBN_11_NAME As String = "�]��"     '               �]��(���FLG[FRS]�F>0�A����FLG[FRS]�F0)
Public Const FRSKBN_12_NAME As String = "����"     '               �Ĕ���(���FLG[FRS]�F>0�A����FLG[FRS]�F3)
Public Const FRSKBN_13_NAME As String = "�����"   '               �����(���FLG[FRS]�F>0�A����FLG[FRS]�F1 or 2)
'Add End 2011/03/08 SMPK Nakamura FRS�V�X�e�����ǉ��Ή�

Private Const NOTCH_ASTER       As String = "****"     '�m�b�`�ʒu��*�\��   2012/06/08�ǉ� SETsw kubota


'///////////////////////////////////////////////////
' @(f)
' �@�\    : Excel�ҏW�����
' �Ԃ�l  : �Ȃ�
' ������  :
' �@�\����:
'///////////////////////////////////////////////////
Public Function PrtExec_XLS35CUT(ByVal sXtalNo As String _
                      , ByRef frmInet As Form _
                      ) As Boolean

    Dim lCnt        As Long
    Dim bResult     As Boolean

    '�e���v���[�g�_�E�����[�h
    bResult = ActDownLoad(TEMPLATENAME, ".xls", frmInet, frmInet.Inet1)
    If bResult = False Then
        '�_�E�����[�h���s
        Call MsgOut(0, "���[�t�@�C���̃_�E�����[�h�Ɏ��s���܂���", ERR_DISP)
        Exit Function
    End If
    
    '����f�[�^�擾
    Call MsgOut(0, "����f�[�^�擾��", NORMAL_MSG)
    DoEvents
    If GetPrintInfo(sXtalNo) = False Then
        Exit Function
    End If
    Call MsgOut(0, "", NORMAL_MSG)

    'Excel�ҏW�����
    If PrtExec_CutSiji = False Then
        Exit Function
    End If
    
    PrtExec_XLS35CUT = True

End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\�@�@: ���[�o�͏��擾����
' �Ԃ�l�@: True  - ����
' �@�@�@    False - �ُ�
' �������@:
' �@�\����:
'///////////////////////////////////////////////////
Private Function GetPrintInfo(ByVal sXtalNo As String) As Boolean
    
    Dim sSql            As String
    Dim objDynaData     As Object
    Dim sXtalNoCnv      As String
    Dim lCnt            As Long
    Dim sWkZuban        As String
    Dim lBlkCnt         As Long     '�u���b�N�J�E���g�@2009/12/08�ǉ�
    Dim sBlock          As String   '�v���b�N���f�p�@�@2009/12/08�ǉ�
    Dim sZuban          As String   '�}�ԕҏW�p�@�@�@�@2009/12/08�ǉ�

    Dim sCsPos          As String   'XODCS�ʒu
    Dim lCsCnt          As Long
    
    Dim lCnt2           As Long     'XODCZ�u���b�N�`�F�b�N�p
    Dim iJissoku        As Integer  '�����T���v���`�F�b�N�p
    Dim iBlock          As Integer  '�u���b�N�`�F�b�N�p

'>>>>> �R�[�h�}�X�^�[���y�[�W�Ǘ��f�[�^�擾 ADD 2009/06/25 SSS.Marushita
    Dim tKoda9          As typKoda9Data
    '�Ǘ��R�[�h�擾
    If GetKanriCode("K", "01", TEMPLATENAME, tKoda9) = False Then
        Exit Function
    End If    '���[1�y�[�W�̍s���̃Z�b�g
    lPRINTPAGEROW = val(Trim$(tKoda9.sKCODE01A9))
    '1�y�[�W���̖��א��̃Z�b�g
    lPRINTMEISAIROW = val(Trim$(tKoda9.sKCODE02A9))
'<<<<< �R�[�h�}�X�^�[���y�[�W�Ǘ��f�[�^�擾 ADD 2009/06/25 SSS.Marushita
    
    '���f�[�^�擾(�w�b�_)
    sSql = "SELECT  NVL(C1.PUHINBC1 , ' ') PUHINBC1"                        '�}��
    sSql = sSql & ",NVL(E018.HSXTYPE ,' ') HSXTYPE"                         '�`���^
    'sSQL = sSQL & ",NVL(TO_CHAR(U001.QCOM_SXLDIADV) , ' ') QCOM_SXLDIADV"   '���a�敪�H
    sSql = sSql & ",NVL(TO_CHAR(E018.HSXD1CEN) , ' ') HSXD1CEN"             '���a
    sSql = sSql & ",NVL(E018.HSXCDIR ,' ') HSXCDIR"                         '������
    sSql = sSql & ",NVL(TO_CHAR(E018.HSXRMIN) ,' ') HSXRMIN"                '��R�K�i(Min)
    sSql = sSql & ",NVL(TO_CHAR(E018.HSXRMAX) ,' ') HSXRMAX"                '��R�K�i(Max)
    sSql = sSql & ",NVL(TO_CHAR(E019.HSXONMIN) ,' ') HSXONMIN"              '����Oi�K�i(Min)
    sSql = sSql & ",NVL(TO_CHAR(E019.HSXONMAX) ,' ') HSXONMAX"              '����Oi�K�i(Max)
    sSql = sSql & ",NVL(TO_CHAR(H001.AMRESIST) ,' ') AMRESIST"              '�˂炢��R
    sSql = sSql & ",NVL(TO_CHAR(C1.PUCHAGC1) ,' ') PUCHAGC1"                '�`���[�W��
    sSql = sSql & ",NVL(H001.PGID ,' ') PGID"                               'PG-ID
    sSql = sSql & ",NVL(H004.STATCLS ,' ') STATCLS"                         '�{�g����
    sSql = sSql & ",NVL(TO_CHAR(C1.WGHTTAC1) ,' ') WGHTTAC1"                '����d��
    sSql = sSql & ",NVL(TO_CHAR(C1.PUTCUTWC1) ,' ') PUTCUTWC1"              '�g�b�v�J�b�g�d��
    sSql = sSql & ",NVL(TO_CHAR(C1.PUFRELC1) ,' ') PUFRELC1"                '�t���[��
    sSql = sSql & ",NVL(TO_CHAR(C1.LENTKC1) ,' ') LENTKC1"                  '���㒷��
    'sSQL = sSQL & ",NVL(TO_CHAR(E8.KACUTLE8) ,' ') KACUTLE8"                '���J�b�g�����H
    sSql = sSql & ",NVL(TO_CHAR(C1.ADDOPPC1) ,' ') ADDOPPC1"                '�ǃh�[�v�ʒu
    sSql = sSql & ",NVL(C2.GNWKNTC2 ,' ') GNWKNTC2"                '�ǃh�[�v�ʒu
    sSql = sSql & "  FROM XSDC2    C2"
    sSql = sSql & "     , XSDC1    C1"
    sSql = sSql & "     , TBCME018 E018"
    sSql = sSql & "     , TBCME019 E019"
    sSql = sSql & "     , TBCMH001 H001"
    sSql = sSql & "     , TBCMH004 H004"
    sSql = sSql & " WHERE C2.CRYNUMC2 = '" & sXtalNo & "'"
    sSql = sSql & "   AND C2.XTALC2 = C1.XTALC1"
    sSql = sSql & "   AND C1.HISIJIC1 = H001.UPINDNO(+)"
    sSql = sSql & "   AND C2.CRYNUMC2 = H004.CRYNUM(+)"
    sSql = sSql & "   AND ( C1.PUHINBC1 = E018.HINBAN(+)"
    sSql = sSql & "   AND   C1.PUREVNUMC1 = E018.MNOREVNO(+)"
    sSql = sSql & "   AND   C1.PUFACTORYC1 = E018.FACTORY(+)"
    sSql = sSql & "   AND   C1.PUOPEC1 = E018.OPECOND(+) )"
    sSql = sSql & "   AND ( C1.PUHINBC1 = E019.HINBAN(+)"
    sSql = sSql & "   AND   C1.PUREVNUMC1 = E019.MNOREVNO(+)"
    sSql = sSql & "   AND   C1.PUFACTORYC1 = E019.FACTORY(+)"
    sSql = sSql & "   AND   C1.PUOPEC1 = E019.OPECOND(+) )"

'Debug.Print sSQL
    
    'SQL���s
    'If mdlCommon.DynSet(objDynaData, sSQL) = False Then
    If mdlCommon.DynSet2(objDynaData, sSql) = False Then
        Call MsgOut(100, sSql, ERR_DISP_LOG, "XSDC2,C1,E018,E019,H001,H004")
        Exit Function
    End If
    If objDynaData.EOF = True Then
        Call MsgOut(0, "(�ؒf�w����)�Y���f�[�^�����݂��܂���", ERR_DISP)
        Exit Function
    End If

    '���w�b�_�ҏW
    With mtPrintInfo
        
        Call GetXtalHensyu(sXtalNo, 1, sXtalNoCnv)          '�����ԍ��ҏW("-"����)
        .sXtalNo = sXtalNoCnv                               '�����ԍ�
        .sDate = Format$(Date, "yyyy.mm.dd")                '���s��
        .sZuban = objDynaData("PUHINBC1").Value             '�i��
        Call GetHinbanHensyu(objDynaData("PUHINBC1").Value, 1, .sZuban)
            
        .sType = LCase$(objDynaData("HSXTYPE").Value)       '�`���^
        .sDia = objDynaData("HSXD1CEN").Value               '���a
        .sJiku = objDynaData("HSXCDIR").Value               '������
        .sRsKikaku = objDynaData("HSXRMIN").Value _
           & " - " & objDynaData("HSXRMAX").Value           '�ϋK�i"
        .sOiKikaku = objDynaData("HSXONMIN").Value _
           & " - " & objDynaData("HSXONMAX").Value          '����Oi�K�i
        .sNeraiRs = objDynaData("AMRESIST").Value           '�˂炢��R
        .sCharge = objDynaData("PUCHAGC1").Value            '�`���[�W��
        .sPgid = objDynaData("PGID").Value                  'PG-ID
        
        '�{�g����(0,1�ȊO������H0����5�H)
        .sBottom = objDynaData("STATCLS").Value
        'If objDynaData("STATCLS").Value = "0" Then
        '    .sBottom = "��"
        'ElseIf objDynaData("STATCLS").Value = "1" Then
        '    .sBottom = "�~"
        'Else
        '    .sBottom = ""
        'End If
            
        .sPulWeight = objDynaData("WGHTTAC1").Value         '����d��
        .sTopCutWeight = objDynaData("PUTCUTWC1").Value     '�g�b�v�J�b�g�d��
        .sFreeLen = objDynaData("PUFRELC1").Value           '�t���[��
        .sPulLen = objDynaData("LENTKC1").Value             '���㒷��
        '.sKataLen = objDynaData("KACUTLE8").Value           '���J�b�g����
        .sOiDopPos = objDynaData("ADDOPPC1").Value          '�ǃh�[�v�ʒu
    
        .sBarCode = "*" & sXtalNo & "*"                     '�����ԍ��o�[�R�[�h ADD 2009/06/30 SSS.Marushita
    End With
    objDynaData.Close
    
    '���f�[�^�擾(�T���v�����݁E�`��)
    sSql = "SELECT  NVL(NAMEJA9 ,' ')   NAMEJA9"            '�T���v����
    sSql = sSql & ",NVL(KCODEA9 ,' ')   KCODEA9"            '�}�\��
    sSql = sSql & ",NVL(KCODE01A9 ,' ') KCODE01A9"          '����
    sSql = sSql & ",NVL(KCODE02A9 ,' ') KCODE02A9"          '�`��(200mm����)
    sSql = sSql & ",NVL(KCODE03A9 ,' ') KCODE03A9"          '�`��(200mm�ȏ�)
    sSql = sSql & "  FROM KODA9"
    sSql = sSql & " WHERE SYSCA9 = 'X'"
    sSql = sSql & "   AND SHUCA9 = 'HE'"
    sSql = sSql & " ORDER BY CTR01A9"

    'SQL���s
    'If mdlCommon.DynSet(objDynaData, sSQL) = False Then
    If mdlCommon.DynSet2(objDynaData, sSql) = False Then
        Call MsgOut(100, sSql, ERR_DISP_LOG, "XODC2,C1,E8,TBSSU001,002,103")
        Exit Function
    End If

    If objDynaData.RecordCount <> CUT_MAXCNT Then
        Call MsgOut(0, "(�ؒf�w����)�T���v�����݁E�`��R�[�h�ݒ茏���ُ�", ERR_DISP)
        Exit Function
    End If
    
    With mtPrintInfo
        For lCnt = 0 To objDynaData.RecordCount - 1
            '���݁iKODA9��'F','08'�Ŏ擾�H�j
            .sThick(lCnt) = objDynaData("KCODE01A9").Value
'>>>>> �đ����1.3mm�Ή��@2008/10/28�@SET.Marushita
            'If .sThick(lCnt) <> "1.1" And .sThick(lCnt) <> "1.2" Then
            If .sThick(lCnt) <> "1.1" And .sThick(lCnt) <> "1.2" And .sThick(lCnt) <> "1.3" Then
'<<<<< �đ����1.3mm�Ή��@2008/10/28�@SET.Marushita
                Call MsgOut(0, "(�ؒf�w����)�T���v�����݃R�[�h�ݒ�ُ�u" & .sThick(lCnt) & "�v", ERR_DISP)
                Exit Function
            End If
            '�`��(���a�敪�������̂�300mm�̂݃Z�b�g�H)
            'If val(mtPrintInfo.sDia) < 200 Then
                '.sShape(lCnt) = objDynaData("KCODE02A9").Value
            'Else
            .sShape(lCnt) = objDynaData("KCODE03A9").Value
            'End If
            If .sShape(lCnt) <> "1/4" And .sShape(lCnt) <> "1/2" And .sShape(lCnt) <> "4/4" Then
                Call MsgOut(0, "(�ؒf�w����)�T���v���`��R�[�h�ݒ�ُ�u" & .sShape(lCnt) & "�v", ERR_DISP)
                Exit Function
            End If
            .sSmplNm(lCnt) = objDynaData("NAMEJA9").Value       '����ٖ�
            .sPicStr(lCnt) = objDynaData("KCODEA9").Value       '����ِ}�\������
            objDynaData.MoveNext
        Next lCnt
    End With
    
    '���f�[�^�擾(����)
    sSql = "SELECT  NVL(CZ.CRYNUMCZ , ' ') CRYNUMCZ"            '�����ԍ�
    sSql = sSql & ",NVL(CZ.HINBCZ , ' ') HINBCZ"                '�}��
    sSql = sSql & ",NVL(TO_CHAR(CZ.INPOSCZ) , ' ') INPOSCZ"     '��������(Top)
    sSql = sSql & ",NVL(TO_CHAR(CZ.GNLCZ) , ' ') GNLCZ"         '�d�|����
'Add Start 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
    sSql = sSql & ",NVL(CS1.CRYINDOIFRSCS1,'-') as INDOIFRS "   'FRS���
    sSql = sSql & ",NVL(CS1.CRYRESOIFRSCS1,'0') as RESOIFRS "   'FRS����
'Add End 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
    sSql = sSql & ",NVL(E018.HSXDPDIR ,' ') HSXDPDIR"           '�i�r�w�a�ʒu���� 2012/06/08�ǉ� SETsw kubota

    sSql = sSql & "  FROM XSDCZ    CZ"
'Add Start 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
    sSql = sSql & ",XSDCS_1    CS1"
'Add End 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
    sSql = sSql & ",TBCME018 E018"                              '2012/06/08�ǉ� SETsw kubota
    
    'sSQL = sSQL & " WHERE CZ.CRYNUMCZ = '" & sXtalNo & "'"
    sSql = sSql & " WHERE CZ.RPCRYNUMCZ = '" & sXtalNo & "'"
    'sSQL = sSQL & "   AND CZ.CUTKCZ = '1'"
'Add Start 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
    sSql = sSql & "   AND CZ.CRYNUMCZ = CS1.CRYNUMCS1(+)"
    sSql = sSql & "   AND CZ.HINBCZ = CS1.HINBCS1(+)"
    sSql = sSql & "   AND CZ.INPOSCZ = CS1.INPOSCS1(+)"
'Add End 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
    
    '2012/06/08�ǉ� SETsw kubota
    sSql = sSql & "   AND CZ.HINBCZ    = E018.HINBAN(+)"
    sSql = sSql & "   AND CZ.REVNUMCZ  = E018.MNOREVNO(+)"
    sSql = sSql & "   AND CZ.FACTORYCZ = E018.FACTORY(+)"
    sSql = sSql & "   AND CZ.OPECZ     = E018.OPECOND(+)"
    
    sSql = sSql & " ORDER BY CZ.INPOSCZ"
    
    'SQL���s
    'If mdlCommon.DynSet(objDynaData, sSQL) = False Then
    If mdlCommon.DynSet2(objDynaData, sSql) = False Then
        Call MsgOut(100, sSql, ERR_DISP_LOG, "XSDCZ")
        Exit Function
    End If
    If objDynaData.EOF = True Then
        Call MsgOut(0, "(�ؒf�w����)�Y���f�[�^�����݂��܂���", ERR_DISP)
        Exit Function
    End If
    
'>>>>> �}���`�i�ԑΉ����C�@2009/12/08�@SSS.Marushita
    lBlkCnt = 0
    sBlock = ""
    ReDim mtPrintInfo.tMeisai(objDynaData.RecordCount)      '��]���ɍ���Ă���
    For lCnt = 1 To objDynaData.RecordCount
        '�u���b�N�������ꍇ�͔z��͂��̂܂�
        If sBlock = Mid$(objDynaData("CRYNUMCZ").Value, 10, 3) Then
            '�u���b�N�������ꍇ�͕i�Ԃ�ǉ�
            sZuban = ""                                           '�}��
            Call GetHinbanHensyu(objDynaData("HINBCZ").Value, 1, sZuban)
            mtPrintInfo.tMeisai(lBlkCnt - 1).sZuban = mtPrintInfo.tMeisai(lBlkCnt - 1).sZuban & " " & Trim$(sZuban)
            mtPrintInfo.tMeisai(lBlkCnt - 1).sLen = CStr(CDbl(objDynaData("GNLCZ").Value) + CDbl(mtPrintInfo.tMeisai(lBlkCnt - 1).sLen))                    '�u���b�N����
            '�}���`�i�ԍP�v�Ή� 2009/12/25 SSS.Marushita
            '�Ō�̃u���b�N�̏ꍇ(�Ō�̌�������(BOT)���Z�b�g) 2009/12/25
            If lCnt = objDynaData.RecordCount Then
                mtPrintInfo.tMeisai(lBlkCnt).sCutPos = CStr(CDbl(objDynaData("GNLCZ").Value) + CDbl(objDynaData("INPOSCZ").Value)) '��������(BOT)
            End If
            
            'Notch�ʒu�Ή� 2012/06/08 SETsw kubota
            If mtPrintInfo.tMeisai(lBlkCnt - 1).sNotchPos <> CnvMizoNotchDisp(objDynaData("HSXDPDIR").Value) Then
                '��i�ԂƎd�l���قȂ�ꍇ�A*�\��
                mtPrintInfo.tMeisai(lBlkCnt - 1).sNotchPos = NOTCH_ASTER
            End If
            
        Else
            lBlkCnt = lBlkCnt + 1
            With mtPrintInfo.tMeisai(lBlkCnt - 1)
                sZuban = ""                                       '�}�Ԕ��f�p���N���A
                Call GetHinbanHensyu(objDynaData("HINBCZ").Value, 1, sZuban)
                .sZuban = Trim$(sZuban)
                .sLen = objDynaData("GNLCZ").Value                    '�u���b�N����
                .sCutPos = objDynaData("INPOSCZ").Value               '��������(Top)
                .sBlockNo = Mid$(objDynaData("CRYNUMCZ").Value, 10, 3)    '�u���b�N�ԍ��H(�����ԍ���10���ځ`)
'Add Start 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
                .sFrsFlg = objDynaData("INDOIFRS").Value
                .sFrsResult = objDynaData("RESOIFRS").Value
'Add End 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�

                'Notch�ʒu�Ή� 2012/06/08 SETsw kubota
                .sNotchPos = CnvMizoNotchDisp(objDynaData("HSXDPDIR").Value)
                If .sNotchPos = "" Then     '�󔒂�*�\��
                    .sNotchPos = NOTCH_ASTER
                End If
                
                sBlock = Mid$(objDynaData("CRYNUMCZ").Value, 10, 3)    '�u���b�N�ԍ���r�p
            End With
            '�Ō�̃u���b�N�̏ꍇ
            If lCnt = objDynaData.RecordCount Then
                '����Bot�ʒu(Top�ʒu+����)��ۑ�
                mtPrintInfo.tMeisai(lBlkCnt).sCutPos = CStr(val(mtPrintInfo.tMeisai(lBlkCnt - 1).sCutPos) + val(mtPrintInfo.tMeisai(lBlkCnt - 1).sLen))
            End If
        End If
        objDynaData.MoveNext
    Next lCnt
    objDynaData.Close
    
    lSET_MEISAI_CNT = lBlkCnt       '�}���`�i�Ԓ�����̖��א�
'<<<<< �}���`�i�ԑΉ����C�@2009/12/08�@SSS.Marushita
    
    '���f�[�^�擾(�T���v���Ǘ�)
    sSql = "SELECT  NVL(CS.CRYNUMCS , ' ') CRYNUMCS"        '�����ԍ�
    sSql = sSql & ",NVL(CS.TBKBNCS , ' ') TBKBNCS"          'T/B�敪
    sSql = sSql & ",NVL(TO_CHAR(CS.INPOSCS) , ' ') INPOSCS" '��������
    '��R
    sSql = sSql & ",NVL(CS.CRYINDRSCS , ' ') CRYINDRSCS"    '���FLG(Rs)
    'Oi�܂���GFA(�ǂ����\���H�����\���H)
    sSql = sSql & ",NVL(CS.CRYINDOICS , ' ') CRYINDOICS"    '���FLG(Oi)
    'OSF�EOSF3(L1����L4�F���ׂĎg�p�H)
    sSql = sSql & ",NVL(CS.CRYINDL1CS , ' ') CRYINDL1CS"    '���FLG(L1)
    sSql = sSql & ",NVL(CS.CRYINDL2CS , ' ') CRYINDL2CS"    '���FLG(L2)
    sSql = sSql & ",NVL(CS.CRYINDL3CS , ' ') CRYINDL3CS"    '���FLG(L3)
    sSql = sSql & ",NVL(CS.CRYINDL4CS , ' ') CRYINDL4CS"    '���FLG(L4)
    'BMD(B1����B3�F���ׂĎg�p�H)
    sSql = sSql & ",NVL(CS.CRYINDB1CS , ' ') CRYINDB1CS"    '���FLG(B1)
    sSql = sSql & ",NVL(CS.CRYINDB2CS , ' ') CRYINDB2CS"    '���FLG(B2)
    sSql = sSql & ",NVL(CS.CRYINDB3CS , ' ') CRYINDB3CS"    '���FLG(B3)
    'DvD2(�\����GD�H)
    sSql = sSql & ",NVL(CS.CRYINDGDCS , ' ') CRYINDGDCS"    '���FLG(GD)
    'LT(T)
    sSql = sSql & ",NVL(CS.CRYINDTCS , ' ') CRYINDTCS"      '���FLG(T)
    'CS
    sSql = sSql & ",NVL(CS.CRYINDCSCS , ' ') CRYINDCSCS"    '���FLG(CS)
    'EPD
    sSql = sSql & ",NVL(CS.CRYINDEPCS , ' ') CRYINDEPCS"    '���FLG(EPD)
    'X��    2009/08/06�ǉ� SETsw kubota
    sSql = sSql & ",NVL(CS.CRYINDXCS , ' ') CRYINDXCS"      '���FLG(X��)

    'Add Start 2011/02/02 SMPK Miyata
    sSql = sSql & ",NVL(CS.CRYINDCCS , ' ') CRYINDCCS"      '���FLG(C)
    sSql = sSql & ",NVL(CS.CRYINDCJCS , ' ') CRYINDCJCS"    '���FLG(CJ)
    sSql = sSql & ",NVL(CS.CRYINDCJLTCS , ' ') CRYINDCJLTCS" '���FLG(CJLT)
    sSql = sSql & ",NVL(CS.CRYINDCJ2CS , ' ') CRYINDCJ2CS"  '���FLG(CJ2)
    'Add End   2011/02/02 SMPK Miyata

    '��R(FLG1,2�̂ǂ�����g�p�H)
    sSql = sSql & ",NVL(CS.CRYRESRS1CS , ' ') CRYRESRS1CS"  '����FLG(Rs1)
    sSql = sSql & ",NVL(CS.CRYRESRS2CS , ' ') CRYRESRS2CS"  '����FLG(Rs2)
    'Oi�܂���GFA(�ǂ����\���H�����\���H)
    sSql = sSql & ",NVL(CS.CRYRESOICS , ' ') CRYRESOICS"    '����FLG(Oi)
    'OSF�EOSF3(L1����L4�F���ׂĎg�p�H)
    sSql = sSql & ",NVL(CS.CRYRESL1CS , ' ') CRYRESL1CS"    '����FLG(L1)
    sSql = sSql & ",NVL(CS.CRYRESL2CS , ' ') CRYRESL2CS"    '����FLG(L2)
    sSql = sSql & ",NVL(CS.CRYRESL3CS , ' ') CRYRESL3CS"    '����FLG(L3)
    sSql = sSql & ",NVL(CS.CRYRESL4CS , ' ') CRYRESL4CS"    '����FLG(L4)
    'BMD(B1����B3�F���ׂĎg�p�H)
    sSql = sSql & ",NVL(CS.CRYRESB1CS , ' ') CRYRESB1CS"    '����FLG(B1)
    sSql = sSql & ",NVL(CS.CRYRESB2CS , ' ') CRYRESB2CS"    '����FLG(B2)
    sSql = sSql & ",NVL(CS.CRYRESB3CS , ' ') CRYRESB3CS"    '����FLG(B3)
    'DvD2(�\����GD�H)
    sSql = sSql & ",NVL(CS.CRYRESGDCS , ' ') CRYRESGDCS"    '����FLG(DvD2)
    'LT(T)
    sSql = sSql & ",NVL(CS.CRYRESTCS , ' ') CRYRESTCS"      '����FLG(LT)
    'CS
    sSql = sSql & ",NVL(CS.CRYRESCSCS , ' ') CRYRESCSCS"    '����FLG(CS)
    'EPD
    sSql = sSql & ",NVL(CS.CRYRESEPCS , ' ') CRYRESEPCS"    '����FLG(EPD)
    'X��    2009/08/06�ǉ� SETsw kubota
    sSql = sSql & ",NVL(CS.CRYRESXCS , ' ') CRYRESXCS"      '����FLG(X��)
    'Add Start 2011/02/02 SMPK Miyata
    sSql = sSql & ",NVL(CS.CRYRESCCS , ' ') CRYRESCCS"      '����FLG(C)
    sSql = sSql & ",NVL(CS.CRYRESCJCS , ' ') CRYRESCJCS"    '����FLG(CJ)
    sSql = sSql & ",NVL(CS.CRYRESCJLTCS , ' ') CRYRESCJLTCS" '����FLG(CJLT)
    sSql = sSql & ",NVL(CS.CRYRESCJ2CS , ' ') CRYRESCJ2CS"  '����FLG(CJ2)
    'Add End   2011/02/02 SMPK Miyata
'>>>>> ��\�T���v��ID�̎擾�Ή��@2009/01/26�@Marushita
    sSql = sSql & ",NVL(CS.REPSMPLIDCS , 0) REPSMPLIDCS"    '��\�T���v��ID
'<<<<< ��\�T���v��ID�̎擾�Ή��@2009/01/26�@Marushita
    
    'GFA�Ή� 2012/06/11 SETsw kubota
    sSql = sSql & ",NVL(E019.HSXONKWY , ' ') HSXONKWY"      '�i�r�w�_�f�Z�x�������@
    
    sSql = sSql & "  FROM XSDCS CS"
    sSql = sSql & "     , XSDC2 C2"
    sSql = sSql & "     , TBCME019 E019"    '2012/06/11�ǉ� SETsw kubota
    sSql = sSql & " WHERE C2.CRYNUMC2 = '" & sXtalNo & "'"
    sSql = sSql & "   AND C2.XTALC2 = CS.XTALCS"
    sSql = sSql & "   AND CS.LIVKCS <> '1'"
    sSql = sSql & "   AND CS.HINBCS = E019.HINBAN(+)"
    sSql = sSql & "   AND CS.REVNUMCS = E019.MNOREVNO(+)"
    sSql = sSql & "   AND CS.FACTORYCS = E019.FACTORY(+)"
    sSql = sSql & "   AND CS.OPECS = E019.OPECOND(+)"
    sSql = sSql & " ORDER BY CS.INPOSCS,CS.TBKBNCS"
    
    'SQL���s
    'If mdlCommon.DynSet(objDynaData, sSQL) = False Then
    If mdlCommon.DynSet2(objDynaData, sSql) = False Then
        Call MsgOut(100, sSql, ERR_DISP_LOG, "XODCS,XODC2")
        Exit Function
    End If
    If objDynaData.EOF = True Then
        Call MsgOut(0, "(�ؒf�w����)�Y���f�[�^�����݂��܂���", ERR_DISP)
        Exit Function
    End If
    
    
    For lCsCnt = 1 To objDynaData.RecordCount
        sCsPos = objDynaData("INPOSCS").Value
        
        'For lCnt = 0 To UBound(mtPrintInfo.tMeisai)
        For lCnt = 0 To lSET_MEISAI_CNT
            With mtPrintInfo.tMeisai(lCnt)
                If sCsPos = .sCutPos Then   '�ʒu�������ꍇ
                    '>>>>> �Ώۃu���b�N�`�F�b�N�s��Ή�  2009/11/18�@SSS.Marushita
                    ''>>>>> �Ώۃu���b�N�`�F�b�N�Ή�  2009/11/12�@SSS.Marushita
                    iBlock = 0
                    For lCnt2 = 0 To lSET_MEISAI_CNT
                    'For lCnt2 = 0 To UBound(mtPrintInfo.tMeisai)
                        '�u���b�N�������ꍇ�̂ݑΏۂƂ���
                        If Mid$(objDynaData("CRYNUMCS").Value, 10, 3) = mtPrintInfo.tMeisai(lCnt2).sBlockNo Then
                            iBlock = 1
                            Exit For
                        End If
                    Next lCnt2
                    '�u���b�N���������̂�����Ƃ��̂ݏ���
                    If iBlock = 1 Then
                        iJissoku = 0
                        '�e���荀�ڂɂ���
                        '���FLG='1'(����)�A����FLG='0'(���тȂ�)�̏ꍇ�A�J�E���g�A�b�v����
                        
                        '��R
                        If objDynaData("CRYINDRSCS").Value = "1" _
                        And objDynaData("CRYRESRS1CS").Value = "0" Then
                        'And objDynaData("CRYRESRS2CS").Value = "0" Then
                            
                            .sSmpl(CUT_RS) = CStr(val(.sSmpl(CUT_RS)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_RS) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_RS) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
                        
                        'Oi
                        If objDynaData("CRYINDOICS").Value = "1" _
                        And objDynaData("CRYRESOICS").Value = "0" Then
                            '>>>>> GFA�\���Ή� 2012/06/11 SETsw kubota ---------------------------
                            '.sSmpl(CUT_OI) = CStr(val(.sSmpl(CUT_OI)) + 1)
                            ''>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            ''CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            'If objDynaData("TBKBNCS").Value = "T" Then
                            '    .iSmpKbnT(CUT_OI) = 1
                            '    iJissoku = 1
                            ''CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            'ElseIf objDynaData("TBKBNCS").Value = "B" Then
                            '    .iSmpKbnB(CUT_OI) = 1
                            '    iJissoku = 2
                            ''<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'End If
                            If objDynaData("HSXONKWY").Value = "CD" Then
                                .sSmpl(CUT_OI) = CStr(val(.sSmpl(CUT_OI)) + 1)
                                'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                                If objDynaData("TBKBNCS").Value = "T" Then
                                    .iSmpKbnT(CUT_OI) = 1
                                    iJissoku = 1
                                'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                                ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                    .iSmpKbnB(CUT_OI) = 1
                                    iJissoku = 2
                                End If
                            ElseIf objDynaData("HSXONKWY").Value = "CG" Then
                                .sSmpl(CUT_GFA) = CStr(val(.sSmpl(CUT_GFA)) + 1)
                                'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                                If objDynaData("TBKBNCS").Value = "T" Then
                                    .iSmpKbnT(CUT_GFA) = 1
                                    iJissoku = 1
                                'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                                ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                    .iSmpKbnB(CUT_GFA) = 1
                                    iJissoku = 2
                                End If
                            End If
                            '<<<<< GFA�\���Ή� 2012/06/11 SETsw kubota ---------------------------
                        End If
                        
                        'OSF(L1)
                        If objDynaData("CRYINDL1CS").Value = "1" _
                        And objDynaData("CRYRESL1CS").Value = "0" Then
                            .sSmpl(CUT_O1) = CStr(val(.sSmpl(CUT_O1)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_O1) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_O1) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
                        
                        'OSF(L2)
                        If objDynaData("CRYINDL2CS").Value = "1" _
                        And objDynaData("CRYRESL2CS").Value = "0" Then
                            .sSmpl(CUT_O2) = CStr(val(.sSmpl(CUT_O2)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_O2) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_O2) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
                        
                        'OSF(L3)
                        If objDynaData("CRYINDL3CS").Value = "1" _
                        And objDynaData("CRYRESL3CS").Value = "0" Then
                            .sSmpl(CUT_O3) = CStr(val(.sSmpl(CUT_O3)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_O3) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_O3) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
                        
                        'BMD(B1)
                        If objDynaData("CRYINDB1CS").Value = "1" _
                        And objDynaData("CRYRESB1CS").Value = "0" Then
                            .sSmpl(CUT_B1) = CStr(val(.sSmpl(CUT_B1)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_B1) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_B1) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
                        
                        'BMD(B2)
                        If objDynaData("CRYINDB2CS").Value = "1" _
                        And objDynaData("CRYRESB2CS").Value = "0" Then
                            .sSmpl(CUT_B2) = CStr(val(.sSmpl(CUT_B2)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_B2) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_B2) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
                        
                        'BMD(B3)
                        If objDynaData("CRYINDB3CS").Value = "1" _
                        And objDynaData("CRYRESB3CS").Value = "0" Then
                            .sSmpl(CUT_B3) = CStr(val(.sSmpl(CUT_B3)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_B3) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_B3) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
                        
                        'DvD2(GD)
                        If objDynaData("CRYINDGDCS").Value = "1" _
                        And objDynaData("CRYRESGDCS").Value = "0" Then
                            .sSmpl(CUT_GD) = CStr(val(.sSmpl(CUT_GD)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_GD) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_GD) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
                        
                        'LT
                        If objDynaData("CRYINDTCS").Value = "1" _
                        And objDynaData("CRYRESTCS").Value = "0" Then
                            .sSmpl(CUT_LT) = CStr(val(.sSmpl(CUT_LT)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_LT) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_LT) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
                        
                        'CS
                        If objDynaData("CRYINDCSCS").Value = "1" _
                        And objDynaData("CRYRESCSCS").Value = "0" Then
                            .sSmpl(CUT_CS) = CStr(val(.sSmpl(CUT_CS)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_CS) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_CS) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
                        
                        'EPD
                        If objDynaData("CRYINDEPCS").Value = "1" _
                        And objDynaData("CRYRESEPCS").Value = "0" Then
                            .sSmpl(CUT_EPD) = CStr(val(.sSmpl(CUT_EPD)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_EPD) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_EPD) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
                        
                        '2�iOS(L4�H)
                        If objDynaData("CRYINDL4CS").Value = "1" _
                        And objDynaData("CRYRESL4CS").Value = "0" Then
                            .sSmpl(CUT_CO3) = CStr(val(.sSmpl(CUT_CO3)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_CO3) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_CO3) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
    
                        'X��    2009/08/06�ǉ� SETsw kubota
                        If objDynaData("CRYINDXCS").Value = "1" _
                        And objDynaData("CRYRESXCS").Value = "0" Then
                            .sSmpl(CUT_X) = CStr(val(.sSmpl(CUT_X)) + 1)
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_X) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_X) = 1
                                iJissoku = 2
                            End If
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If

                        'Add Start 2011/02/04 SMPK Miyata
                        'C
                        If objDynaData("CRYINDCCS").Value = "1" _
                        And objDynaData("CRYRESCCS").Value = "0" Then
                            .sSmpl(CUT_C) = CStr(val(.sSmpl(CUT_C)) + 1)
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_C) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_C) = 1
                                iJissoku = 2
                            End If
                        End If
                        'CJ
                        If objDynaData("CRYINDCJCS").Value = "1" _
                        And objDynaData("CRYRESCJCS").Value = "0" Then
                            .sSmpl(CUT_CJ) = CStr(val(.sSmpl(CUT_CJ)) + 1)
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_CJ) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_CJ) = 1
                                iJissoku = 2
                            End If
                        End If
                        'CJ2
                        If objDynaData("CRYINDCJ2CS").Value = "1" _
                        And objDynaData("CRYRESCJ2CS").Value = "0" Then
                            .sSmpl(CUT_CJ2) = CStr(val(.sSmpl(CUT_CJ2)) + 1)
                            'CS�̋敪���g�b�v�̎��A�敪�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_CJ2) = 1
                                iJissoku = 1
                            'CS�̋敪���{�g���̎��A�敪�ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_CJ2) = 1
                                iJissoku = 2
                            End If
                        End If
                        'Add End   2011/02/04 SMPK Miyata

'>>>>> ��\�T���v��ID�̃Z�b�g�@2009/01/26�@Marushita
                        If Trim(objDynaData("REPSMPLIDCS").Value) <> "0" Then
                            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                            'CS�̋敪���g�b�v�Ŏ����̎��A�g�b�v�ɃZ�b�g
                            If objDynaData("TBKBNCS").Value = "T" And iJissoku = 1 Then
                                .sSmpNoT = Format(val(CStr(objDynaData("REPSMPLIDCS").Value)), "000000")
                            'CS�̋敪���{�g���Ŏ����̎��A�{�g���ɃZ�b�g
                            ElseIf objDynaData("TBKBNCS").Value = "B" And iJissoku = 2 Then
                                .sSmpNoB = Format(val(CStr(objDynaData("REPSMPLIDCS").Value)), "000000")
                            End If
                            '.sSmpNo = Format(val(CStr(objDynaData("REPSMPLIDCS").Value)), "000000")
                            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                        End If
'<<<<< ��\�T���v��ID�̃Z�b�g�@2009/01/26�@Marushita
                    End If
                    'Next lCnt2
                    '<<<<< �Ώۃu���b�N�`�F�b�N�Ή�  2009/11/12�@SSS.Marushita
                    '<<<<< �Ώۃu���b�N�`�F�b�N�s��Ή�  2009/11/18�@SSS.Marushita
                End If
                
            End With
        
        Next lCnt
        objDynaData.MoveNext
    Next lCsCnt

    '�e�ʒu�̃T���v���}�ҏW�A�T���v�������𐔂���
    mtPrintInfo.sMaisu = "0"
    For lCnt = 0 To lSET_MEISAI_CNT
    'For lCnt = 0 To UBound(mtPrintInfo.tMeisai)
        Call GetSamplePic(lCnt)
        '�������v�v�Z
        mtPrintInfo.sMaisu = CStr(val(mtPrintInfo.sMaisu) _
                                + val(mtPrintInfo.tMeisai(lCnt).sMaisu))
    Next lCnt
    
    '����I��
    GetPrintInfo = True

End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\�@�@: Excel���[�ҏW���������
' �Ԃ�l�@: True  - ����I��
' �@�@�@�@  False - �ُ�I��
' ������  :
' �@�\����:
'///////////////////////////////////////////////////
Private Function PrtExec_CutSiji() As Boolean
    
    '��`��Object�ɕύX
    Dim xlApp           As Object               'EXCEL�֘A
    Dim xlBook          As Object               'EXCEL�֘A
    Dim xlSheet         As Object               'EXCEL�֘A
    'Dim xlApp           As Excel.Application    'EXCEL�֘A
    'Dim xlBook          As Excel.Workbook       'EXCEL�֘A
    'Dim xlSheet         As Excel.Worksheet      'EXCEL�֘A
    Dim objFSO          As Object               'FSO
    
    Dim szSavePath      As String               '�o�̓t�@�C���p�X
    Dim szTmpFileName   As String               '�e���v���[�g�t�@�C����(�p�X��)
    Dim szOutFileName   As String               '�o�̓t�@�C����(�p�X��)
    
    Dim szError         As String               '�G���[���b�Z�[�W
    
    Dim lCnt            As Long                 '���[�v�J�E���^
    Dim lSheetCnt       As Long                 '�V�[�g��

    Dim szSCell         As String               '�I���Z��(�J�n�ʒu)
    Dim szECell         As String               '�I���Z��(�I���ʒu)
    Dim szSCellTo       As String               '�R�s�[��Z��(�J�n�ʒu)
    Dim szECellTo       As String               '�R�s�[��Z��(�I���ʒu)
    
    Dim lOutputCnt      As String               '�o�͍��@���J�E���^
    Dim sGroupNo        As String               '�O���[�v��
    
    PrtExec_CutSiji = False
    
    Set xlApp = CreateObject("Excel.Application")
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Call MsgOut(0, "����t�@�C���쐬��", NORMAL_MSG)

    '�t�@�C�����擾
    szSavePath = App.Path & "\" & "REPORT"
    szTmpFileName = App.Path & "\" & TEMPLATENAME & ".xls"
    szOutFileName = szSavePath & "\" & PRINTFILENAME & "_" & Format$(Now(), "YYYYMMDDhhmmss") & ".xls"

    '�f�B���N�g�����ݗL���`�F�b�N
    If Not objFSO.FolderExists(szSavePath) Then
        '������΍��
        Call objFSO.CreateFolder(szSavePath)
    End If
    
    '�e���v���[�g���R�s�[
    objFSO.CopyFile szTmpFileName, szOutFileName
    
    '�t�@�C�����J��
    Set xlBook = xlApp.Workbooks.Open(szOutFileName)
    Set xlSheet = xlBook.Worksheets(1)
    
    '�x�����b�Z�[�W����
    xlApp.DisplayAlerts = False
'    xlApp.DisplayAlerts = True

    xlSheet.Activate

    '���o����
    xlApp.Visible = False                   'Excel���\��
    On Error GoTo FileDeleteErrorExit
    
    Call SetPrintData(xlSheet)
    
    '�����\��
    xlSheet.Cells(1, 1).Show
    
    '�����
    xlSheet.PrintOut
    
    '���t�@�C���ۑ�
    xlBook.SaveAs szOutFileName

    Call MsgOut(0, "�o�͂��������܂���", NORMAL_MSG)

    '�I��
    xlBook.Close
'    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set objFSO = Nothing

    '����I��
    PrtExec_CutSiji = True
    Exit Function
    
FileDeleteErrorExit:

    '���I��
'    xlBook.Close
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set objFSO = Nothing

    Call MsgOut(0, "", ERR_DISP_LOG, "")
    szError = "�d�w�b�d�k�t�@�C���o�͂Ɏ��s���܂����B" & vbCrLf & _
                "(" & Err.Number & ")" & Err.Description
    
    Call MsgBox(szError, vbCritical + vbOKOnly, "EXCEĻ�ُo��")

End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\�@�@: Excel���[�ҏW���C������
' �Ԃ�l�@: True  - ����I��
' �@�@�@�@  False - �ُ�I��
' �������@: xlSheet  - excel�V�[�g�I�u�W�F�N�g
' �@�@�@�@  lPageCnt - �y�[�W��
' �@�\����:
'         : ���y�[�W�Ή��@2009/06/29 SSS.Marushita
'///////////////////////////////////////////////////
Private Sub SetPrintData(ByRef xlSheet As Object)
'Private Sub SetPrintData(ByRef xlSheet As Excel.Worksheet)  ��`��Object�ɕύX

    Dim lCnt        As Long
    Dim lPicCnt     As Long
    Dim lSmplCnt    As Long
    Dim lBaseCol    As Long
    Dim sStrCell    As String
    Dim sEndCell    As String
    
    Dim szSCell     As String   ' �Z���ʒu�p�@ADD 2009/06/29 SSS.Marushita
    Dim szECell     As String   ' �Z���ʒu�p�@ADD 2009/06/29 SSS.Marushita

    Dim lPageCnt    As Long     ' �y�[�W���p�@ADD 2009/06/25 SSS.Marushita
    Dim lBaseRow    As Long     ' Row�ʒu�p�@ ADD 2009/06/25 SSS.Marushita

    Dim lEditCntMax As Long     ' ���ҏW���ő� ADD 2009/09/02 SETsw kubota

'>>>>> �F���ʊǗ��Ή��ǉ� 2011/11/30 SET.Abe
    Dim sGetColor   As String       '�F�Ǘ��R�[�h
    Dim typKODA9    As typKoda9Data 'KODA9��`��
'<<<<< �F���ʊǗ��Ή��ǉ� 2011/11/30 SET.Abe

    '�ǉ��y�[�W���擾
    lPageCnt = Int((lSET_MEISAI_CNT - 1) / lPRINTMEISAIROW)
    'lPageCnt = Int((UBound(mtPrintInfo.tMeisai) - 1) / lPRINTMEISAIROW)
    For lCnt = 1 To lPageCnt
        szSCell = "A" & CStr(lCnt * lPRINTPAGEROW) + 1
        szECell = "BC" & CStr((lCnt + 1) * lPRINTPAGEROW)
        '�y�[�W�̃R�s�[
        Call xlSheet.Range("A1", "BC" & CStr(lPRINTPAGEROW)).Copy(xlSheet.Range(szSCell, szECell))
    Next lCnt
    
    '�w�b�_�ҏW
    For lCnt = 0 To lPageCnt
        lBaseRow = lCnt * lPRINTPAGEROW
        With mtPrintInfo
'>>>>> �F���ʊǗ��Ή��ǉ� 2011/12/01 SET.Abe
            '���[�F�������ʋ��ʊ֐�(300mm�p)�ŐF�Ǘ��R�[�h���擾
            sGetColor = Fnc_GetColor_300(Replace(.sXtalNo, "-", ""))
'            Call MsgOut(0, "DEBUG �擾�F�Ǘ��R�[�h = " & sGetColor, ERR_DISP_LOG)
            '�Ǘ��R�[�h�e�[�u���擾���ʊ֐�(GetKanriCode)�ɂ��F�ԍ����擾
            Call GetKanriCode("X", "CO", sGetColor, typKODA9)
'            Call MsgOut(0, "DEBUG �擾�F�ԍ��P = " & typKODA9.sKCODE01A9, ERR_DISP_LOG)
'            Call MsgOut(0, "DEBUG �擾�F�ԍ��Q = " & typKODA9.sKCODE02A9, ERR_DISP_LOG)
            '�F�ԍ���0(��)�łȂ���
            If typKODA9.sKCODE01A9 <> "0" Then
                '���[�^�C�g�������̃Z��(B1�`G2) �̔w�i�F��F�ԍ��P(KCODE01A9)�ɐݒ�
                xlSheet.Range("B1:G2").Interior.ColorIndex = typKODA9.sKCODE01A9
            End If
            '�F�ԍ���0(��)�łȂ���
            If typKODA9.sKCODE02A9 <> "0" Then
                '���[�^�C�g���E���̃Z��(H1�`M2) �̔w�i�F��F�ԍ��Q(KCODE02A9)�ɐݒ�
                xlSheet.Range("H1:M2").Interior.ColorIndex = typKODA9.sKCODE02A9
            End If
'<<<<< �F���ʊǗ��Ή��ǉ� 2011/12/01 SET.Abe
            
            xlSheet.Cells(lBaseRow + 2, 37).Value = "'" & .sDate             '���s��
            xlSheet.Cells(lBaseRow + 4, 5).Value = "'" & .sXtalNo            '�����ԍ�
            xlSheet.Cells(lBaseRow + 4, 22).Value = "'" & .sZuban            '�i��
            xlSheet.Cells(lBaseRow + 5, 22).Value = "'" & .sType             '�`���^
            xlSheet.Cells(lBaseRow + 6, 22).Value = "'" & .sDia              '���a
            xlSheet.Cells(lBaseRow + 7, 22).Value = "'" & .sJiku             '������
            xlSheet.Cells(lBaseRow + 8, 22).Value = "'" & .sRsKikaku         '�ϋK�i
            xlSheet.Cells(lBaseRow + 9, 22).Value = "'" & .sOiKikaku         'Oi�K�i
        
            xlSheet.Cells(lBaseRow + 4, 31).Value = "'" & .sNeraiRs          '�˂炢��R
            xlSheet.Cells(lBaseRow + 5, 31).Value = "'" & .sCharge           '�`���[�W��
            xlSheet.Cells(lBaseRow + 6, 31).Value = "'" & .sPgid             'PG-ID
            xlSheet.Cells(lBaseRow + 7, 31).Value = "'" & .sBottom           '�{�g����
            xlSheet.Cells(lBaseRow + 8, 31).Value = "'" & .sPulWeight        '����d��
            xlSheet.Cells(lBaseRow + 9, 31).Value = "'" & .sTopCutWeight     '�g�b�v�J�b�g�d��
        
            xlSheet.Cells(lBaseRow + 4, 38).Value = "'" & .sFreeLen          '�t���[��
            xlSheet.Cells(lBaseRow + 5, 38).Value = "'" & .sPulLen           '���㒷��
            '�\���Ȃ�
            xlSheet.Cells(lBaseRow + 6, 35).Value = ""                       '���J�b�g����(�^�C�g��)
            xlSheet.Cells(lBaseRow + 6, 38).Value = "'" & .sKataLen          '���J�b�g����
            xlSheet.Cells(lBaseRow + 7, 38).Value = "'" & .sOiDopPos         '�ǃh�[�v�ʒu
            
            For lSmplCnt = CUT_RS To CUT_MAXCNT - 1
                xlSheet.Cells(lBaseRow + 20 + lSmplCnt, 2).Value = "'" & .sSmplNm(lSmplCnt)  '����ٖ�
                xlSheet.Cells(lBaseRow + 20 + lSmplCnt, 4).Value = "'" & .sThick(lSmplCnt)   '����
                xlSheet.Cells(lBaseRow + 20 + lSmplCnt, 6).Value = "'" & .sShape(lSmplCnt)   '�`��
            Next lSmplCnt
            
            xlSheet.Cells(lBaseRow + 40, 6).Value = "'" & .sMaisu            '�T���v���������v

            '�y�[�W�E�o�[�R�[�h�̕\���ǉ� ADD 2009/06/30 SSS.Marushita
            xlSheet.Cells(lBaseRow + 2, 53).Value = "'" & lCnt + 1 & "/" & lPageCnt + 1      '�y�[�W�\��
            xlSheet.Cells(lBaseRow + 4, 42).Value = "'" & .sBarCode          '�����ԍ��o�[�R�[�h

'        xlSheet.Cells(2, 37).Value = "'" & .sDate               '���s��
'        xlSheet.Cells(4, 5).Value = "'" & .sXtalNo              '�����ԍ�
'        xlSheet.Cells(4, 22).Value = "'" & .sZuban              '�i��
'        xlSheet.Cells(5, 22).Value = "'" & .sType               '�`���^
'        xlSheet.Cells(6, 22).Value = "'" & .sDia                '���a
'        xlSheet.Cells(7, 22).Value = "'" & .sJiku               '������
'        xlSheet.Cells(8, 22).Value = "'" & .sRsKikaku           '�ϋK�i
'        xlSheet.Cells(9, 22).Value = "'" & .sOiKikaku           'Oi�K�i
'
'        xlSheet.Cells(4, 31).Value = "'" & .sNeraiRs            '�˂炢��R
'        xlSheet.Cells(5, 31).Value = "'" & .sCharge             '�`���[�W��
'        xlSheet.Cells(6, 31).Value = "'" & .sPgid               'PG-ID
'        xlSheet.Cells(7, 31).Value = "'" & .sBottom             '�{�g����
'        xlSheet.Cells(8, 31).Value = "'" & .sPulWeight          '����d��
'        xlSheet.Cells(9, 31).Value = "'" & .sTopCutWeight       '�g�b�v�J�b�g�d��
'
'        xlSheet.Cells(4, 38).Value = "'" & .sFreeLen            '�t���[��
'        xlSheet.Cells(5, 38).Value = "'" & .sPulLen             '���㒷��
'        '�\���Ȃ�
'        xlSheet.Cells(6, 35).Value = ""                         '���J�b�g����(�^�C�g��)
'        xlSheet.Cells(6, 38).Value = "'" & .sKataLen            '���J�b�g����
'        xlSheet.Cells(7, 38).Value = "'" & .sOiDopPos           '�ǃh�[�v�ʒu
'
'        For lSmplCnt = CUT_RS To CUT_EPD
'            xlSheet.Cells(19 + lSmplCnt, 2).Value = "'" & .sSmplNm(lSmplCnt)    '����ٖ�
'            xlSheet.Cells(19 + lSmplCnt, 4).Value = "'" & .sThick(lSmplCnt)     '����
'            xlSheet.Cells(19 + lSmplCnt, 6).Value = "'" & .sShape(lSmplCnt)     '�`��
'        Next lSmplCnt
'
''>>>>> �T���v�����\���Ή��@2009/01/26�@Marushita
'        xlSheet.Cells(33, 6).Value = "'" & .sMaisu              '�T���v���������v
''        xlSheet.Cells(32, 6).Value = "'" & .sMaisu              '�T���v���������v
''<<<<< �T���v�����\���Ή��@2009/01/26�@Marushita
    
        End With
    Next lCnt
    
    '���וҏW
    For lCnt = 0 To lSET_MEISAI_CNT
    'For lCnt = 0 To UBound(mtPrintInfo.tMeisai)
'        With mtPrintInfo.tMeisai(lCnt)     '�g�p����Ă���ӏ�(��)�Ɉړ� 2010/10/25
            '15����(14����+�Ō�j
            '�J�n�ʒu�̒���(�ŏ��͌Œ�)
            If lCnt = 0 Then
                lBaseRow = 0
                lBaseCol = 8
            Else
                lBaseRow = Int((lCnt - 1) / lPRINTMEISAIROW) * lPRINTPAGEROW
                lBaseCol = (lCnt - Int((lCnt - 1) / lPRINTMEISAIROW) * lPRINTMEISAIROW) * 3 + 8
            End If
            '���y�[�W�̐擪�̎�
            If Int(lCnt / lPRINTMEISAIROW) > 0 And lBaseCol = 11 Then
                '�O�ōŏI�ʒu��擪�ɃZ�b�g
                xlSheet.Cells(lBaseRow + 11, 10).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sBlockNo  '�u���b�NID
                '>>>>> �}���`�i�ԍP�v�Ή��@2009/12/09�@SSS.Marushita
                xlSheet.Cells(lBaseRow + 12, 9).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sZuban    '�}��
                'xlSheet.Cells(lBaseRow + 12, 10).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sZuban    '�}��
                '�u���b�NID�A�}�Ԃ���ꂽ�ӏ��ɐF������
                If mtPrintInfo.tMeisai(lCnt - 1).sBlockNo <> "" Then
                    sStrCell = ConvXlsNumToA(9) & CStr(lBaseRow + 11)
                    sEndCell = ConvXlsNumToA(11) & CStr(lBaseRow + 12)
                    xlSheet.Range(sStrCell, sEndCell).Interior.Color = &HC0C0C0
                End If
                
                xlSheet.Cells(lBaseRow + 15, 9).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sLen         '�u���b�N����
                xlSheet.Cells(lBaseRow + 14, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sCutPos      '�ؒf�ʒu
                xlSheet.Cells(lBaseRow + 13, 10).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sNotchPos   'Notch�ʒu 2012/06/08�ǉ� SETsw kubota
                xlSheet.Cells(lBaseRow + 19, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sCutPos      '�ؒf�ʒu
'Add Start 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
                'FRS����
                If mtPrintInfo.tMeisai(lCnt - 1).sFrsFlg = FRSKBN_NONE Then
                    xlSheet.Cells(lBaseRow + 15, 11).Value = ""
                ElseIf mtPrintInfo.tMeisai(lCnt - 1).sFrsFlg = FRSKBN_0 Then
                    xlSheet.Cells(lBaseRow + 15, 11).Value = ""
                Else
                    If mtPrintInfo.tMeisai(lCnt - 1).sFrsResult = FRSRSL_0 Then
                        xlSheet.Cells(lBaseRow + 15, 11).Value = "��"
                    ElseIf mtPrintInfo.tMeisai(lCnt - 1).sFrsResult = FRSRSL_3 Then
                        xlSheet.Cells(lBaseRow + 15, 11).Value = FRSKBN_12_NAME
                    Else
                        xlSheet.Cells(lBaseRow + 15, 11).Value = ""
                    End If
                End If
'Add End 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
                For lSmplCnt = CUT_RS To CUT_MAXCNT - 1
                    '>>>>> �T���v���ʒu�g�b�v�{�g�������}���`�i�ԑΉ�  2009/11/18�@SSS.Marushita
                    xlSheet.Cells(lBaseRow + 20 + lSmplCnt, 8).Value = "'" & GetSampleStr(CStr(mtPrintInfo.tMeisai(lCnt - 1).iSmpKbnT(lSmplCnt)), _
                                                                                          CStr(mtPrintInfo.tMeisai(lCnt - 1).iSmpKbnB(lSmplCnt)))
                    '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                    'xlSheet.Cells(lBaseRow + 19 + lSmplCnt, 8).Value = "'" & GetSampleStr(CStr(mtPrintInfo.tMeisai(lCnt - 1).iSmpKbn(lSmplCnt)))
                    'xlSheet.Cells(lBaseRow + 19 + lSmplCnt, 8).Value = "'" & GetSampleStr(mtPrintInfo.tMeisai(lCnt - 1).sSmpl(lSmplCnt))
                    '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                    '<<<<< �T���v���ʒu�g�b�v�{�g�������}���`�i�ԑΉ�  2009/11/18�@SSS.Marushita
                Next lSmplCnt
                '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                xlSheet.Cells(lBaseRow + 38, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sSmpNoB       '�T���v����BOT
                xlSheet.Cells(lBaseRow + 39, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sSmpNoT       '�T���v����TOP
                'xlSheet.Cells(lBaseRow + 36, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sSmpNo       '�T���v����
                '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                xlSheet.Cells(lBaseRow + 40, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sMaisu       '�T���v���w��(��R)
                For lSmplCnt = 0 To val(mtPrintInfo.tMeisai(lCnt - 1).sMaisu) - 1
                    
'>>>>> X������Ή� 2009/09/02 SETsw kubota ------------------
'                    Call OvalWrite(xlSheet, lBaseRow + 39 + lSmplCnt * 5, 8)
                    If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 0) = "X" Then
                        'X���̏ꍇ�A��y�[�W�ڂ̍��v�̉��Ɂ��\��
                        Call OvalWrite(xlSheet, 42, 4)
                        xlSheet.Cells(43, 4).Value = mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 0)
                        lEditCntMax = 4     '�ʏ�̈ʒu��X����ҏW���Ȃ����A�ꖇ�������[�v
                    Else
                        lEditCntMax = 3
                        Call OvalWrite(xlSheet, lBaseRow + 42 + lSmplCnt * 5, 8)
                    End If
'<<<<< X������Ή� 2009/09/02 SETsw kubota ------------------
                    
                    For lPicCnt = 0 To 3
                        If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, lPicCnt) <> PIC_DUMMY Then
                            Select Case lPicCnt
                            Case 0  '����
                                If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 0) <> "X" Then   'X���͓�y�[�W�ڈȍ~���\���Ȃ�
                                    xlSheet.Cells(lBaseRow + 43 + lSmplCnt * 5, 8).Value = mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, lPicCnt)
                                    
                                    '�E�オ��łȂ���ΉE�Ɖ��Ɍr��������
                                    If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 1) <> "" Then
                                        sStrCell = ConvXlsNumToA(8) & CStr(lBaseRow + 42 + lSmplCnt * 5)
                                        sEndCell = ConvXlsNumToA(8) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                        With xlSheet.Range(sStrCell, sEndCell)
                                            .Borders(xlEdgeRight).LineStyle = xlContinuous  '�ʏ��
                                            .Borders(xlEdgeRight).Weight = xlThin
                                        End With
                                        sStrCell = ConvXlsNumToA(8) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                        sEndCell = ConvXlsNumToA(8) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                        With xlSheet.Range(sStrCell, sEndCell)
                                            .Borders(xlEdgeBottom).LineStyle = xlContinuous  '�ʏ��
                                            .Borders(xlEdgeBottom).Weight = xlThin
                                        End With
                                    End If
                                    '��������łȂ���Ή��r��������
                                    If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 2) <> "" Then
                                        sStrCell = ConvXlsNumToA(8) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                        sEndCell = ConvXlsNumToA(9) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                        With xlSheet.Range(sStrCell, sEndCell)
                                            .Borders(xlEdgeBottom).LineStyle = xlContinuous '�ʏ��
                                            .Borders(xlEdgeBottom).Weight = xlThin
                                        End With
                                    End If
                                End If
                            Case 1
                                xlSheet.Cells(lBaseRow + 43 + lSmplCnt * 5, 9).Value = mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, lPicCnt)
                            Case 2
                                xlSheet.Cells(lBaseRow + 44 + lSmplCnt * 5, 8).Value = mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, lPicCnt)
                                '�E������łȂ���ΉE�r��������
                                If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 3) <> "" Then
                                    sStrCell = ConvXlsNumToA(8) & CStr(lBaseRow + 44 + lSmplCnt * 5)
                                    sEndCell = ConvXlsNumToA(8) & CStr(lBaseRow + 45 + lSmplCnt * 5)
                                    With xlSheet.Range(sStrCell, sEndCell)
                                        .Borders(xlEdgeRight).LineStyle = xlContinuous  '�ʏ��
                                        .Borders(xlEdgeRight).Weight = xlThin
                                    End With
                                End If
                            Case 3
                                xlSheet.Cells(lBaseRow + 44 + lSmplCnt * 5, 9).Value = mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, lPicCnt)
                            End Select
                        End If
                    Next lPicCnt
                    
                    '4���ҏW�����甲����
                    'If lSmplCnt = 2 Then
                    If lSmplCnt = lEditCntMax Then
                        Exit For
                    End If
                Next
                
            End If
            
        With mtPrintInfo.tMeisai(lCnt)     '�ړ� 2010/10/25
            
            '�ŏI�ʒu�̎��A�u���b�NID�E�}�ԁE�F���E�u���b�N�����͕\�����Ȃ�
            If (lCnt Mod lPRINTMEISAIROW) = 0 And lCnt > 0 Then
            Else
                xlSheet.Cells(lBaseRow + 11, lBaseCol + 2).Value = "'" & .sBlockNo    '�u���b�NID
                '>>>>> �}���`�i�ԍP�v�Ή��@2009/12/09�@SSS.Marushita
                xlSheet.Cells(lBaseRow + 12, lBaseCol + 1).Value = "'" & .sZuban      '�}��
                'xlSheet.Cells(lBaseRow + 12, lBaseCol + 2).Value = "'" & .sZuban      '�}��
                'lBaseCol = lCnt * 3 + 9
                'xlSheet.Cells(11, lBaseCol + 1).Value = "'" & .sBlockNo    '�u���b�NID
                'xlSheet.Cells(12, lBaseCol + 1).Value = "'" & .sZuban      '�}��
                xlSheet.Cells(lBaseRow + 13, lBaseCol + 2).Value = "'" & .sNotchPos 'Notch�ʒu  2012/06/08�ǉ� SETsw kubota
                
                '�u���b�NID�A�}�Ԃ���ꂽ�ӏ��ɐF������
                If .sBlockNo <> "" Then
                    sStrCell = ConvXlsNumToA(lBaseCol + 1) & CStr(lBaseRow + 11)
                    sEndCell = ConvXlsNumToA(lBaseCol + 3) & CStr(lBaseRow + 12)
                    'sStrCell = ConvXlsNumToA(lBaseCol) & "11"
                    'sEndCell = ConvXlsNumToA(lBaseCol + 2) & "12"
                    xlSheet.Range(sStrCell, sEndCell).Interior.Color = &HC0C0C0
                End If
                
                xlSheet.Cells(lBaseRow + 15, lBaseCol + 1).Value = "'" & .sLen         '�u���b�N����
                'xlSheet.Cells(15, lBaseCol).Value = "'" & .sLen             '�u���b�N����
'Add Start 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
                'FRS����
                If .sFrsFlg = FRSKBN_NONE Then
                    xlSheet.Cells(lBaseRow + 17, lBaseCol + 1).Value = ""
                ElseIf .sFrsFlg = FRSKBN_0 Then
                    xlSheet.Cells(lBaseRow + 17, lBaseCol + 1).Value = ""
                Else
                    If .sFrsResult = FRSRSL_0 Then
                        xlSheet.Cells(lBaseRow + 17, lBaseCol + 1).Value = "��"
                    ElseIf .sFrsResult = FRSRSL_3 Then
                        xlSheet.Cells(lBaseRow + 17, lBaseCol + 1).Value = FRSKBN_12_NAME
                    Else
                        xlSheet.Cells(lBaseRow + 17, lBaseCol + 1).Value = ""
                    End If
                End If
'Add End 2011/03/08 SMPK Nakamura FRS�V�X�e�����Ή�
            End If
            xlSheet.Cells(lBaseRow + 14, lBaseCol).Value = "'" & .sCutPos       '�ؒf�ʒu
            xlSheet.Cells(lBaseRow + 19, lBaseCol).Value = "'" & .sCutPos       '�ؒf�ʒu
            'lBaseCol = lCnt * 3 + 8
            'xlSheet.Cells(14, lBaseCol).Value = "'" & .sCutPos          '�ؒf�ʒu
            'xlSheet.Cells(18, lBaseCol).Value = "'" & .sCutPos          '�ؒf�ʒu
            
            For lSmplCnt = CUT_RS To CUT_MAXCNT - 1
                '>>>>> �T���v���ʒu�g�b�v�{�g�������}���`�i�ԑΉ�  2009/11/18�@SSS.Marushita
                xlSheet.Cells(lBaseRow + 20 + lSmplCnt, lBaseCol).Value = "'" & GetSampleStr(CStr(.iSmpKbnT(lSmplCnt)), _
                                                                                             CStr(.iSmpKbnB(lSmplCnt)))
                '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                'xlSheet.Cells(lBaseRow + 19 + lSmplCnt, lBaseCol).Value = "'" & GetSampleStr(.sSmpl(lSmplCnt))
                '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
                'xlSheet.Cells(19 + lSmplCnt, lBaseCol).Value = "'" & GetSampleStr(.sSmpl(lSmplCnt))
                '<<<<< �T���v���ʒu�g�b�v�{�g�������}���`�i�ԑΉ�  2009/11/18�@SSS.Marushita
            Next lSmplCnt
            
'>>>>> �T���v�����̕\���@2009/01/26�@Marushita
            '>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
            xlSheet.Cells(lBaseRow + 38, lBaseCol).Value = "'" & .sSmpNoB      '�T���v����BOT
            xlSheet.Cells(lBaseRow + 39, lBaseCol).Value = "'" & .sSmpNoT      '�T���v����TOP
            ''''xlSheet.Cells(lBaseRow + 36, lBaseCol).Value = "'" & .sSmpNo       '�T���v����
            '<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
            xlSheet.Cells(lBaseRow + 40, lBaseCol).Value = "'" & .sMaisu       '�T���v���w��(����)
            'xlSheet.Cells(32, lBaseCol).Value = "'" & .sSmpNo           '�T���v����
            'xlSheet.Cells(33, lBaseCol).Value = "'" & .sMaisu           '�T���v���w��(��R)
'            xlSheet.Cells(32, lBaseCol).Value = "'" & .sMaisu           '�T���v���w��(��R)
'<<<<< �T���v�����̕\���@2009/01/26�@Marushita

            For lSmplCnt = 0 To val(.sMaisu) - 1
'>>>>> X������Ή� 2009/09/02 SETsw kubota ------------------
''>>>>> �T���v�����\���Ή��@2009/01/26�@Marushita
'                Call OvalWrite(xlSheet, lBaseRow + 39 + lSmplCnt * 5, lBaseCol)
'                'Call OvalWrite(xlSheet, 35 + lSmplCnt * 5, lBaseCol)
'                'Call OvalWrite(xlSheet, 34 + lSmplCnt * 5, lBaseCol)
''<<<<< �T���v�����\���Ή��@2009/01/26�@Marushita
                If .sSmplPic(lSmplCnt, 0) = "X" Then
                    'X���̏ꍇ�A��y�[�W�ڂ̍��v�̉��Ɂ��\��
                    Call OvalWrite(xlSheet, 42, 4)
                    xlSheet.Cells(43, 4).Value = .sSmplPic(lSmplCnt, 0)
                    lEditCntMax = 4     '�ʏ�̈ʒu��X����ҏW���Ȃ����A�ꖇ�������[�v
                Else
                    Call OvalWrite(xlSheet, lBaseRow + 42 + lSmplCnt * 5, lBaseCol)
                    lEditCntMax = 3
                End If
'<<<<< X������Ή� 2009/09/02 SETsw kubota ------------------
                
                For lPicCnt = 0 To 3
                    If .sSmplPic(lSmplCnt, lPicCnt) <> PIC_DUMMY Then
'                    If .sSmplPic(lSmplCnt, lPicCnt) <> "" Then
                    
                        Select Case lPicCnt
                        Case 0  '����
'>>>>> X������Ή� 2009/09/02 SETsw kubota ------------------
                            If .sSmplPic(lSmplCnt, 0) <> "X" Then
'<<<<< X������Ή� 2009/09/02 SETsw kubota ------------------
                                xlSheet.Cells(lBaseRow + 43 + lSmplCnt * 5, lBaseCol).Value = .sSmplPic(lSmplCnt, lPicCnt)
                                
                                '�E�オ��łȂ���ΉE�Ɖ��Ɍr��������
                                If .sSmplPic(lSmplCnt, 1) <> "" Then
                                    sStrCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 42 + lSmplCnt * 5)
                                    sEndCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                    With xlSheet.Range(sStrCell, sEndCell)
                                        .Borders(xlEdgeRight).LineStyle = xlContinuous  '�ʏ��
                                        .Borders(xlEdgeRight).Weight = xlThin
                                    End With
                                    sStrCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                    sEndCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                    With xlSheet.Range(sStrCell, sEndCell)
                                        .Borders(xlEdgeBottom).LineStyle = xlContinuous  '�ʏ��
                                        .Borders(xlEdgeBottom).Weight = xlThin
                                    End With
                                End If
                                '��������łȂ���Ή��r��������
                                If .sSmplPic(lSmplCnt, 2) <> "" Then
                                    sStrCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                    sEndCell = ConvXlsNumToA(lBaseCol + 1) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                    With xlSheet.Range(sStrCell, sEndCell)
                                        .Borders(xlEdgeBottom).LineStyle = xlContinuous '�ʏ��
                                        .Borders(xlEdgeBottom).Weight = xlThin
                                    End With
                                End If
'>>>>> X������Ή� 2009/09/02 SETsw kubota ------------------
                            End If
'<<<<< X������Ή� 2009/09/02 SETsw kubota ------------------
                        
                        Case 1
'>>>>> �T���v�����\���Ή��@2009/01/26�@Marushita
                            xlSheet.Cells(lBaseRow + 43 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(36 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(35 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
'<<<<< �T���v�����\���Ή��@2009/01/26�@Marushita
                        Case 2
'>>>>> �T���v�����\���Ή��@2009/01/26�@Marushita
                            xlSheet.Cells(lBaseRow + 44 + lSmplCnt * 5, lBaseCol).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(37 + lSmplCnt * 5, lBaseCol).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(36 + lSmplCnt * 5, lBaseCol).Value = .sSmplPic(lSmplCnt, lPicCnt)
'<<<<< �T���v�����\���Ή��@2009/01/26�@Marushita
                            '�E������łȂ���ΉE�r��������
                            If .sSmplPic(lSmplCnt, 3) <> "" Then
'>>>>> �T���v�����\���Ή��@2009/01/26�@Marushita
                                sStrCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 44 + lSmplCnt * 5)
                                sEndCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 45 + lSmplCnt * 5)
                                'sStrCell = ConvXlsNumToA(lBaseCol) & CStr(37 + lSmplCnt * 5)
                                'sEndCell = ConvXlsNumToA(lBaseCol) & CStr(38 + lSmplCnt * 5)
                                'sStrCell = ConvXlsNumToA(lBaseCol) & CStr(36 + lSmplCnt * 5)
                                'sEndCell = ConvXlsNumToA(lBaseCol) & CStr(37 + lSmplCnt * 5)
'<<<<< �T���v�����\���Ή��@2009/01/26�@Marushita
                                With xlSheet.Range(sStrCell, sEndCell)
                                    .Borders(xlEdgeRight).LineStyle = xlContinuous  '�ʏ��
                                    .Borders(xlEdgeRight).Weight = xlThin
                                End With
                            End If
                        Case 3
'>>>>> �T���v�����\���Ή��@2009/01/26�@Marushita
                            xlSheet.Cells(lBaseRow + 44 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(37 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(36 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
'<<<<< �T���v�����\���Ή��@2009/01/26�@Marushita
                        End Select
                    
                    End If
                Next lPicCnt
                
''>>>>> �T���v�����\���Ή��@2009/01/26�@Marushita
'                '4���ҏW�����甲�����3���ҏW�����甲����
'                If lSmplCnt = 2 Then
'                'If lSmplCnt = 3 Then
''<<<<< �T���v�����\���Ή��@2009/01/26�@Marushita
                If lSmplCnt = lEditCntMax Then
                    Exit For
                End If
            Next

        End With
    Next lCnt

End Sub


'///////////////////////////////////////////////////
' @(f)
' �@�\�@�@: �~��`������
' �Ԃ�l�@: True  - ����I��
' �@�@�@�@  False - �ُ�I��
' �������@: xlSheet  - excel�V�[�g�I�u�W�F�N�g
' �@�@�@�@  sSell - ��ƂȂ�Z��
' �@�\����:
'///////////////////////////////////////////////////
Private Function OvalWrite(ByRef xlSheet As Object _
                         , ByVal lRow As Long _
                         , ByVal lCol As Long _
                         ) As Boolean
'Private Function OvalWrite(ByRef xlSheet As Excel.Worksheet _�@  ��`��Object�ɕύX

'>>>>>�@Excel2007�Ή�(�~�̈ʒu���������ɑΉ�)�@2009/01/29�@Marushita
'Dim x As Object, MyZoom As Variant, VW As Variant
'>>>>>�@Excel���J���Ă��鎞�ɃI�u�W�F�N�g�G���[�ƂȂ���ɑΉ��@2009/05/18�@Marushita
Dim objShape As Object
    
    With xlSheet
    Set objShape = .Shapes.AddShape(msoShapeOval _
                        , .Cells(lRow, lCol).Left _
                        , .Cells(lRow, lCol).Top _
                        , CON_X _
                        , CON_Y _
                       )
    End With
    '�h��Ԃ��A���̐F�E�����̎w��
    objShape.Fill.Visible = msoFalse
    objShape.Line.ForeColor.SchemeColor = 0
    objShape.Line.Weight = 0.75
    
    '�Y�[�����������Ȃ��i�e���v���[�g��100%�ɂ��đΉ��j�@2009/05/18�@Marushita
'    '���݂̃Y�[���{����ۑ�
'    MyZoom = ActiveWindow.Zoom
'    '��ʂ̃Y�[���{����100%�ɂ���
'    VW = ActiveWindow.View
'    Application.ScreenUpdating = False
'    ActiveWindow.Zoom = 100
'    ActiveWindow.View = xlNormalView
'
'    With xlSheet
'         .Shapes.AddShape(msoShapeOval _
'                        , .Cells(lRow, lCol).Left _
'                        , .Cells(lRow, lCol).Top _
'                        , CON_X _
'                        , CON_Y _
'                       ).Select
'    End With
'    '�h��Ԃ��A���̐F�E�����̎w��
'    With Selection.ShapeRange
'        .Fill.Visible = msoFalse
'        .Line.ForeColor.SchemeColor = 0
'        .Line.Weight = 0.75
'    End With
'
'    '���̃Y�[���{���ɖ߂�
'    ActiveWindow.Zoom = MyZoom
'    ActiveWindow.View = VW
'
'    Application.ScreenUpdating = True
'<<<<<�@Excel2007�Ή�(�~�̈ʒu���������ɑΉ�)�@2009/01/29�@Marushita
'    '�Z����Left�ATop�AWidth�v���p�e�B�[�𗘗p���Ĉʒu����
'    With xlSheet
'        'Shape�̕`��
'        .Shapes.AddShape(msoShapeOval _
'                       , .Cells(lRow, lCol).Left _
'                       , .Cells(lRow, lCol).Top _
'                       , CON_X _
'                       , CON_Y _
'                       ).Fill.Visible = msoFalse
'
''        .Shapes.AddShape(msoShapeOval, BX, BY, EX, EY).Name = "aaa"
''        .Shapes("aaa").Fill.Visible = msoFalse
'    End With
'<<<<<�@Excel���J���Ă��鎞�ɃI�u�W�F�N�g�G���[�ƂȂ���ɑΉ��@2009/05/18�@Marushita

End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\�@�@: �T���v���\��������擾
' �Ԃ�l�@: �T���v���\��������
' �������@: �T���v����
' �@�\����: �T���v���\��������擾
'///////////////////////////////////////////////////
'�}���`�i�ԑΉ�
'Private Function GetSampleStr(ByVal sSampleNum As String) As String
Private Function GetSampleStr(ByVal sSampleNumT As String, ByVal sSampleNumB As String) As String
    
    '�T���v��������"��"��\������
    '�ύX�̉\���L��A1�Ȃ灛�A2�Ȃ灝��������
    'GetSampleStr = String(val(sSampleNum), "��")
'>>>>> �T���v���ʒu�g�b�v�{�g�������}���`�i�ԑΉ�  2009/11/12�@SSS.Marushita
    '�g�b�v�{�g���̔��f��ǉ�
    If sSampleNumT = "1" And sSampleNumB = "1" Then         '��������
        GetSampleStr = "�� ��"
    ElseIf sSampleNumT = "1" Then     '�g�b�v�݂̂���
        GetSampleStr = "�@ ��"
    ElseIf sSampleNumB = "1" Then     '�{�g���݂̂���
        GetSampleStr = "�� �@"
    End If
'<<<<< �T���v���ʒu�g�b�v�{�g�������}���`�i�ԑΉ�  2009/11/12�@SSS.Marushita
''>>>>> �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
'    '�g�b�v�{�g���̔��f��ǉ�
'    If val(sSampleNum) = 1 Then         '�g�b�v�݂̂���
'        GetSampleStr = "�@ ��"
'    ElseIf val(sSampleNum) = 2 Then     '�{�g���݂̂���
'        GetSampleStr = "�� �@"
'    ElseIf val(sSampleNum) = 3 Then     '��������
'        GetSampleStr = "�� ��"
'    End If
''<<<<< �T���v���ʒu�g�b�v�{�g�������Ή�  2009/11/12�@SSS.Marushita
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\�@�@: �T���v�������𐔂���
' �Ԃ�l�@: �T���v������
' �������@: �s
' �@�\����: �T���v�������𐔂���
'///////////////////////////////////////////////////
Private Sub GetSamplePic(ByVal lRow As Long)
    
    Dim lCnt            As Long
    Dim lCnt2           As Long
    Dim lThick          As Long
    Dim lSmpl_1_4(1)    As Long     '1/4�T���v���̐�
    Dim lSmpl_1_2(1)    As Long     '1/2�T���v���̐�
    Dim lSmpl_4_4(1)    As Long     '4/4�T���v���̐�
    
    Dim lPic_1_4()      As Long     '1/4�T���v���`��
    Dim lPic_1_2()      As Long     '1/4�T���v���`��
    Dim lPic_4_4()      As Long     '1/4�T���v���`��
    
    Dim lPicPos         As Long     '�ǂ̏ꏊ�ɏ�����
    
    With mtPrintInfo.tMeisai(lRow)
        '���݁A�`�󂪓����T���v�����܂Ƃ߂�
        For lCnt = CUT_RS To CUT_MAXCNT - 1
        
            '���݂�1.1,1.2�݂̂�z��
            If mtPrintInfo.sThick(lCnt) = "1.1" Then
                lThick = 0
'>>>>> �đ����1.3mm�Ή��@2008/10/28�@SET.Marushita
            'ElseIf mtPrintInfo.sThick(lCnt) = "1.2" Then
            '���݂�1.2,1.3�͓�����
            ElseIf mtPrintInfo.sThick(lCnt) = "1.2" Or _
                   mtPrintInfo.sThick(lCnt) = "1.3" Then
'<<<<< �đ����1.3mm�Ή��@2008/10/28�@SET.Marushita
                lThick = 1
            Else
                lThick = -1     '�G���[��
                Call MsgOut(0, "(�ؒf�w����)���݃G���[:���݁u" & mtPrintInfo.sThick(lCnt) & "�v", ERR_DISP)
            End If
            
            '�e�`��Ŏw�������J�E���g
            If val(.sSmpl(lCnt)) > 0 Then
                Select Case mtPrintInfo.sShape(lCnt)
                Case "1/4"
                    lSmpl_1_4(lThick) = lSmpl_1_4(lThick) + val(.sSmpl(lCnt))
                    
                    '�z����m��
                    If lSmpl_1_4(lThick) >= lSmpl_1_4(0) _
                    And lSmpl_1_4(lThick) >= lSmpl_1_4(1) Then
                        ReDim Preserve lPic_1_4(1, lSmpl_1_4(lThick) - 1)
                    End If
                    
                    For lCnt2 = lSmpl_1_4(lThick) - val(.sSmpl(lCnt)) To lSmpl_1_4(lThick) - 1
                        '�ǂ̃T���v������������ۑ�
                        lPic_1_4(lThick, lCnt2) = lCnt
                    Next lCnt2
                    
                Case "1/2"
                    lSmpl_1_2(lThick) = lSmpl_1_2(lThick) + val(.sSmpl(lCnt))
                
                    '�z����m��
                    If lSmpl_1_2(lThick) >= lSmpl_1_2(0) _
                    And lSmpl_1_2(lThick) >= lSmpl_1_2(1) Then
                        ReDim Preserve lPic_1_2(1, lSmpl_1_2(lThick) - 1)
                    End If
                    
                    For lCnt2 = lSmpl_1_2(lThick) - val(.sSmpl(lCnt)) To lSmpl_1_2(lThick) - 1
                        '�ǂ̃T���v������������ۑ�
                        lPic_1_2(lThick, lCnt2) = lCnt
                    Next lCnt2
                
                Case "4/4"
                    lSmpl_4_4(lThick) = lSmpl_4_4(lThick) + val(.sSmpl(lCnt))
                
                    '�z����m��
                    If lSmpl_4_4(lThick) >= lSmpl_4_4(0) _
                    And lSmpl_4_4(lThick) >= lSmpl_4_4(1) Then
                        ReDim Preserve lPic_4_4(1, lSmpl_4_4(lThick) - 1)
                    End If
                    
                    For lCnt2 = lSmpl_4_4(lThick) - val(.sSmpl(lCnt)) To lSmpl_4_4(lThick) - 1
                        '�ǂ̃T���v������������ۑ�
                        lPic_4_4(lThick, lCnt2) = lCnt
                    Next lCnt2
                
                End Select
            End If
        Next lCnt
        
        '�T���v�������̌v�Z�ƕ`����ҏW
        lPicPos = 0
        For lThick = 0 To 1       '����0:1.1�A1:1.2
            
            '���݂��ς������r������͕ҏW���Ȃ�
            If lPicPos Mod 4 > 0 Then
                lPicPos = lPicPos + 4 - lPicPos Mod 4
            End If
            
            For lCnt = 0 To lSmpl_1_4(lThick) - 1
                If lPicPos < 16 Then    '�ҏW��4���܂�
                    .sSmplPic(Fix(lPicPos / 4), lPicPos Mod 4) = mtPrintInfo.sPicStr(lPic_1_4(lThick, lCnt))
                End If
                lPicPos = lPicPos + 1
                If lPicPos < 16 Then    '�ҏW��4���܂�
                    .sSmplPic(Fix(lPicPos / 4), lPicPos Mod 4) = PIC_DUMMY  '1/4�ɂ���}�[�N������
                End If
            Next lCnt
            For lCnt = 0 To lSmpl_1_2(lThick) - 1
                '�c�ł͊���Ȃ�
                If lPicPos Mod 2 = 1 Then
                    lPicPos = lPicPos + 1
                End If
                If lPicPos < 16 Then    '�ҏW��4���܂�
                    .sSmplPic(Fix(lPicPos / 4), lPicPos Mod 4) = mtPrintInfo.sPicStr(lPic_1_2(lThick, lCnt))
                End If
                lPicPos = lPicPos + 2
                If lPicPos < 16 Then    '�ҏW��4���܂�
                    .sSmplPic(Fix(lPicPos / 4), lPicPos Mod 4) = PIC_DUMMY  '1/2�ɂ���}�[�N������
                End If
            Next lCnt
            For lCnt = 0 To lSmpl_4_4(lThick) - 1
                '�r������͕ҏW���Ȃ�
                If lPicPos Mod 4 > 0 Then
                    lPicPos = lPicPos + 4 - lPicPos Mod 4
                End If
                If lPicPos < 16 Then    '�ҏW��4���܂�
                    .sSmplPic(Fix(lPicPos / 4), lPicPos Mod 4) = mtPrintInfo.sPicStr(lPic_4_4(lThick, lCnt))
                End If
                lPicPos = lPicPos + 4
            Next lCnt
            
        Next lThick
        .sMaisu = Fix((lPicPos + 3) / 4)
        
    End With


End Sub


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �o�͗v�ۃ`�F�b�N
' �Ԃ�l  : True - �o�͂���  False - �o�͂��Ȃ�
' ������  : �Ȃ�
' �@�\����:
'///////////////////////////////////////////////////
Public Function ChkPrtYN() As Boolean

    Dim tKoda9          As typKoda9Data

    ChkPrtYN = False

    '�Ǘ��R�[�h�}�X�^�擾
    If GetKanriCode("K", "01", TEMPLATENAME, tKoda9) = False Then
        Exit Function
    End If
    
    If tKoda9.sKCODE05A9 = "1" Then
        'KCODE05��"1"�̏ꍇ�ɏo��
        ChkPrtYN = True
    End If

End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �o�͗v�ۃ`�F�b�N(�Đؒf��)
' �Ԃ�l  : True - �o�͂���  False - �o�͂��Ȃ�
' ������  : �Ȃ�
' �@�\����:
'///////////////////////////////////////////////////
Public Function ChkPrtYN_S() As Boolean

    Dim tKoda9          As typKoda9Data

    ChkPrtYN_S = False

    '�Ǘ��R�[�h�}�X�^�擾
    If GetKanriCode("K", "01", TEMPLATENAME, tKoda9) = False Then
        Exit Function
    End If
    
    If tKoda9.sKCODE04A9 = "1" Then
        'KCODE04��"1"�̏ꍇ�ɏo��
        ChkPrtYN_S = True
    End If

End Function



