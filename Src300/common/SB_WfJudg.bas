Attribute VB_Name = "SB_WfJudg"
Option Explicit

'''''Public typ_Param001b As DBDRV_scmzc_fcmlc001b_SXL
'''''Public Const MAXREC As Integer = 256
'''''Private intChkPos As Integer                                    ' �`�F�b�N�ʒu

''Public MaxLine As Integer
Public SelectSxlID As String
'''''Public typ_ww() As DBDRV_scmzc_fcmlc001b_SXL   '�҂��ꗗ���
'''''Public WFJudgExecOkFlag() As Boolean    'WF����������s�\�t���O

'����t���[�͂���Ȋ����ł����H
' ������t���[��
' �d�l�ۏؕ��@�Q�� --+--�Ȃ� --���сi�Y���ʒu�j--�����Ă��Ȃ��Ă�����OK
'�@�@�@�@�@�@�@�@�@�@|
'                   +--���� --���сi�Y���ʒu) --+--���� -- ����`�F�b�N --+-- OK
'                                              |                        |
'                                              |                        +-- MG
'                                              |
'                                              +--�Ȃ� --+-- �����w���T�E�U�ȊO�̏ꍇ ---NG
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�i�����w���́A�w���𗧂Ă鑤������ɗ��ĂĂ���ƍl���Ă���j


''
'' �萔��`
''
'''''Public lStfMst As Long
'''''Public intEnCmd As Integer
Private Const MAXCNT As Integer = 18                             ' �ő匏��
Private Const SXL_MAXSMP As Integer = 1 + 1 + 10                ' SXL���̍ő�T���v�������@'Add 2011/03/07 SMPK Miyata
                                                                '  - Top:MAX1���ABot:MAX1���A���Ԕ���:MAX10��
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
Private Const MAXCNT_EP As Integer = 6                             ' �ő匏��
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
Public Const SxlTop As Integer = 1                                 ' TOP��
Public Const SxlTail As Integer = 2                                ' TAIL��
Public Const SxlMidl As Integer = 3                                ' MIDLE��    'Add 2011/03/07 SMPK Miyata
'''''Public Const KSYSCLASS As String = "GP"                         ' �V�X�e���敪
Public Const MSYSCLASS As String = "NM"                         ' �V�X�e���敪
Public Const KCLASS As String = "01"                            ' �N���X
Public Const KCODE As String = "1"                              ' �R�[�h

'''''Private Const cnEnableColor As Long = &H80FF80                  ' �L���J���[
'''''Private Const cnEnableColor2 As Long = vbWindowBackground       ' �L���J���[
'''''Private Const cnDisenableColor As Long = &H80FF80               ' �����J���[
'''''Private Const cnDisenableGrayColor As Long = vbButtonFace       ' �����J���[�i�D�F�j
'''''Private Const cnWarningColor As Long = &H8080FF                 ' �x���J���[

Public Const WFRES As Integer = 0
Public Const WFOI As Integer = 1
Public Const WFBMD1 As Integer = 2
Public Const WFBMD2 As Integer = 3
Public Const WFBMD3 As Integer = 4
Public Const WFOSF1 As Integer = 5
Public Const WFOSF2 As Integer = 6
Public Const WFOSF3 As Integer = 7
Public Const WFOSF4 As Integer = 8
Public Const WFDS As Integer = 9
Public Const WFDZ As Integer = 10
Public Const WFSP As Integer = 11
Public Const WFDOI1 As Integer = 12
Public Const WFDOI2 As Integer = 13
Public Const WFDOI3 As Integer = 14
Public Const WFOT1 As Integer = 15
Public Const WFOT2 As Integer = 16
Public Const WFAOI As Integer = 17
Public Const WFGD As Integer = 18           '05/02/07 ooba
'''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
Public Const EPBMD1 As Integer = 0
Public Const EPBMD2 As Integer = 1
Public Const EPBMD3 As Integer = 2
Public Const EPOSF1 As Integer = 3
Public Const EPOSF2 As Integer = 4
Public Const EPOSF3 As Integer = 5
Public Const EPOT2 As Integer = 6
'''--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'Add 2010/01/07 SIRD�Ή� Y.Hitomi
Public Const WFSIRD As Integer = 8

Public Const OSWFRES As String = "RES"
Public Const OSWFOI As String = "OI"
Public Const OSWFBMD1 As String = "BMD1"
Public Const OSWFBMD2 As String = "BMD2"
Public Const OSWFBMD3 As String = "BMD3"
Public Const OSWFOSF1 As String = "OSF1"
Public Const OSWFOSF2 As String = "OSF2"
Public Const OSWFOSF3 As String = "OSF3"
Public Const OSWFOSF4 As String = "OSF4"
Public Const OSWFDS As String = "DSOD"
Public Const OSWFDZ As String = "DZ"
Public Const OSWFSP As String = "SPV"
Public Const OSWFDOI1 As String = "DOI1"
Public Const OSWFDOI2 As String = "DOI2"
Public Const OSWFDOI3 As String = "DOI3"
Public Const OSWFOT1 As String = "OT1"
Public Const OSWFOT2 As String = "OT2"
Public Const OSWFAOI As String = "AOI"      ''�c���_�f�ǉ��@03/12/15 ooba
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
Public Const OSEPBMD1 As String = "BMD1"
Public Const OSEPBMD2 As String = "BMD2"
Public Const OSEPBMD3 As String = "BMD3"
Public Const OSEPOSF1 As String = "OSF1"
Public Const OSEPOSF2 As String = "OSF2"
Public Const OSEPOSF3 As String = "OSF3"
Public Const OSEPOT2 As String = "OTHER2"
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'Add 2010/01/07 SIRD�Ή� Y.Hitomi
Public Const OSWFSIRD As String = "SIRD"


'''''' �R�[�h�}�X�^�[
'''''Public Type typ_CodeMaster
'''''    SYSCLASS As String * 2          ' �V�X�e���敪
'''''    Class As String * 2             ' �敪
'''''    CODE As String * 5              ' �R�[�h
'''''    INFO1 As String                 ' ���P
'''''    INFO2 As String                 ' ���Q
'''''    INFO3 As String                 ' ���R
'''''    INFO4 As String                 ' ���S
'''''    INFO5 As String                 ' ���T
'''''    INFO6 As String                 ' ���U
'''''    INFO7 As String                 ' ���V
'''''    INFO8 As String                 ' ���W
'''''    INFO9 As String                 ' ���X
'''''    NOTE As String                  ' ���l
'''''    TSTAFFID As String * 8          ' �o�^�Ј�ID
'''''    REGDATE As Date                 ' �o�^���t
'''''    KSTAFFID As String * 8          ' �X�V�Ј�ID
'''''    UPDDATE As Date                 ' �X�V���t
'''''End Type

''''''�e���я��
'''''Public Type typ_ALLRSLT
'''''    pos As Integer                    ' �������J�n�ʒu
'''''    NAIYO As String                   ' ���e
'''''    INFO1 As String                   ' ���P
'''''    INFO2 As String                   ' ���Q
'''''    INFO3 As String                   ' ���R
'''''    INFO4 As String                   ' ���S
'''''    OKNG  As String                   ' ���茋��
'''''    SMPLID As String                  ' �T���v���m��
'''''End Type

'Add Start 2011/03/10 SMPK Miyata
Public Type typ_TBCMY013_arry
    typ_y013midl()      As typ_TBCMY013
End Type

Public Type typ_TBCMY022_arry
    typ_y022midl()      As typ_TBCMY022
End Type
'Add End   2011/03/10 SMPK Miyata

'�S���\����
Public Type typ_AllTypesC
    StrStaffId          As String                               ' �X�^�b�tID
    strStaffName        As String                               ' �X�^�b�t��
'Chg Start 2011/03/09 SMPK Miyata
'    dblScut(2)          As Double                               ' �ăJ�b�g�ʒu
'    bOKNG(2)            As Boolean                              ' ���R����
'    COEF(2)             As Double                               ' �ΐ͌W��
'    JudgRes(2)          As Boolean                              ' ���R����    2002/01/15 S.Sano
'    JudgRrg(2)          As Boolean                              ' RRG����       2002/01/15 S.Sano
    dblScut()           As Double                               ' �ăJ�b�g�ʒu
    bOKNG()             As Boolean                              ' ���R����
    COEF()              As Double                               ' �ΐ͌W��
    JudgRes()           As Boolean                              ' ���R����
    JudgRrg()           As Boolean                              ' RRG����
'Chg End   2011/03/09 SMPK Miyata

'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    'AN���x����p
    JudgAntnp(12)        As Boolean                              ' AN���x����  'Cng 2011/08/12 Y.Hitomi
'    JudgAntnp(2)        As Boolean                              ' AN���x����
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    JudgDkTmp(12)        As Boolean                              ' DK���x����  'Cng 2011/08/12 Y.Hitomi
    DkTmpJsk(12)         As String                               ' DK���x(����)'Cng 2011/08/12 Y.Hitomi
'    JudgDkTmp(2)        As Boolean                              ' DK���x����
'    DkTmpJsk(2)         As String                               ' DK���x(����)
    DkTmpSiyo           As String                               ' DK���x(�d�l)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    typ_Param           As DBDRV_scmzc_fcmlc001b_SXL            ' SXL�Ǘ��i�҂��ꗗ����j
    typ_si              As type_DBDRV_scmzc_fcmlc001c_Siyou     ' ���i�d�l
    typ_y013top()       As typ_TBCMY013                         ' ���茋��(TOP)
    typ_y013tail()      As typ_TBCMY013                         ' ���茋��(TAIL)
    typ_y013midl_ary()  As typ_TBCMY013_arry                    ' ���茋��(MIDLE)   'Add 2011/03/07 SMPK Miyata
'Chg Start 2011/03/09 SMPK Miyata
'* VB��64k�����ɂ��typ_AllTypesC���̃T�C�Y��傫���ł��Ȃ��̂ŁA
'* typ_y013��ÓI�^�����I�^�ɕύX����
'    typ_y013(2, MAXCNT) As typ_TBCMY013                         ' ���茋��
'    typ_hage(2)         As typ_TBCMH004                         ' ���グ�I������
'    typ_rslt(2, MAXCNT) As typ_ALLRSLT                          ' �e���я��
    typ_y013()          As typ_TBCMY013                         ' ���茋��
    typ_hage()          As typ_TBCMH004                         ' ���グ�I������
    typ_rslt()          As typ_ALLRSLT                          ' �e���я��
'Chg End   2011/03/09 SMPK Miyata
    sMidErrMsg          As String                               ' ���Ԕ����`�F�b�N�G���[���b�Z�[�W  'Add 2011/05/10 SMPK Miyata
End Type

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
'VB��64k�����ɂ��typ_AllTypesC���ɒ�`�ł��Ȃ����߁A�ʍ\���̂Ƃ��č쐬����
'�S���\����(�G�s��)
Public Type typ_AllTypesC_EP
    typ_y022top()       As typ_TBCMY022                         ' �G�s��s���茋��(TOP)
    typ_y022tail()      As typ_TBCMY022                         ' �G�s��s���茋��(TAIL)
    typ_y022midl_ary()  As typ_TBCMY022_arry                    ' �G�s��s���茋��(MIDLE)   'Add 2011/03/10 SMPK Miyata
'Chg Start 2011/03/10 SMPK Miyata
'    typ_y022(2, MAXCNT_EP)      As typ_TBCMY022                 ' �G�s��s���茋��
'    typ_rslt(2, MAXCNT_EP)      As typ_ALLRSLT_EX               ' �e���я��
    typ_y022(SXL_MAXSMP, MAXCNT_EP)     As typ_TBCMY022          ' �G�s��s���茋��
    typ_rslt(SXL_MAXSMP, MAXCNT_EP)     As typ_ALLRSLT_EX        ' �e���я��
'Chg End   2011/03/10 SMPK Miyata
End Type
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

'�d�l�����x���\����
Type Judg_Spec_Wf
    rs      As Boolean
    Oi      As Boolean
    B1      As Boolean
    B2      As Boolean
    B3      As Boolean
    L1      As Boolean
    L2      As Boolean
    L3      As Boolean
    L4      As Boolean
    Dsod    As Boolean
    sp      As Boolean
    DZ      As Boolean
    Doi1    As Boolean
    Doi2    As Boolean
    Doi3    As Boolean
    OT1     As Boolean
    OT2     As Boolean
    AOI     As Boolean
    GD      As Boolean      'GD�ǉ��@05/01/27 ooba
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    B1E     As Boolean
    B2E     As Boolean
    B3E     As Boolean
    L1E     As Boolean
    L2E     As Boolean
    L3E     As Boolean
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'Add 2010/01/07 SIRD�Ή� Y.Hitomi
    SIRD    As Boolean
End Type

Public JudgSW       As Judg_Spec_Wf             '�d�l�����x���\����
Public typ_CType    As typ_AllTypesC            '�S���\����
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
Public typ_CType_EP As typ_AllTypesC_EP         '�S���\����(�G�s)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
Public TotalJudg    As Boolean                  '�g�[�^������
Public MidlJudg     As Boolean                  '���Ԕ�������   Add 2011/03/09 SMPK Miyata

'Chg Start 2011/03/07 SMPK Miyata
'Public typ_J015_WFGDJudg(2) As typ_TBCMJ015     'WF��������pGD���с@05/02/04 ooba
Public typ_J015_WFGDJudg() As typ_TBCMJ015     'WF��������pGD����
'Chg End   2011/03/07 SMPK Miyata
Public typ_J015_WFGDUpd() As typ_TBCMJ015       'TBCMJ015-UPDATE�pGD���с@05/02/07 ooba
Public iCntJ015upd As Integer                   'TBCMJ015-UPDATEں��ސ��@05/02/07 ooba

'Chg Start 2011/03/07 SMPK Miyata
'''Upd start 2005/06/21 (TCS)T.Terauchi  SPV9�_�Ή�
'Public typ_J016_WFSPVJudg(2) As typ_TBCMJ016     'WF��������pSPV����
'''Upd end   2005/06/21 (TCS)T.Terauchi  SPV9�_�Ή�
Public typ_J016_WFSPVJudg() As typ_TBCMJ016     'WF��������pSPV����
'Chg End   2011/03/07 SMPK Miyata

'Chg Start 2011/03/07 SMPK Miyata
'''��Add 2010/01/12 SIRD�Ή� Y.Hitomi
'Public typ_J022_WFSDJudg(2) As typ_TBCMJ022     'WF��������pSIRD����
'''��Add 2010/01/12 SIRD�Ή� Y.Hitomi
Public typ_J022_WFSDJudg() As typ_TBCMJ022     'WF��������pSIRD����
'Chg End   2011/03/07 SMPK Miyata

'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
'�����̍\���̂ɍ��ڒǉ������VB�̐����Ɉ���������̂ŁA�ʂŊǗ�����B
'�e���茋�ʏ��
'Chg Start 2011/03/09 SMPK Miyata
'Public typ_rslt_ex(2, MAXCNT) As typ_ALLRSLT_EX                          ' �e���я��
Public typ_rslt_ex(SXL_MAXSMP, MAXCNT) As typ_ALLRSLT_EX                  ' �e���я��
'Chg End   2011/03/09 SMPK Miyata
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------


'''''Public HErrMsg As String
'''''Public typ_rt(2) As typ_TBCMW009            '�������葪��l 2001/09/14 S.Sano
'''''Public bPPlus As Boolean                    'P+Flag 2001/12/18 S.Sano
'''''Public bNPlus As Boolean                    'N+Flag 2002/01/08 S.Sano
'Chg Start 2011/03/09 SMPK Miyata
'Public JiltusekiUmu(2, MAXCNT) As Boolean       '���їL����� 2001/12/19 S.Sano
Public JiltusekiUmu(SXL_MAXSMP, MAXCNT) As Boolean  '���їL�����
'Chg End   2011/03/09 SMPK Miyata
'''''Public MeasFlag(2) As Judg_Spec_Wf         '�d�l�����x���\����
''''Public Tokusai As String                    ' ���̃t���O    'del 2003/05/28 hitec)matsumoto �錾���Q������

'Chg Start 2011/03/09 SMPK Miyata
'Public TmpOsfData(1, 2, MAXCNT) As String  'OSF����/�ő�l�@2003/05/20 ooba
'Public TmpOsfMBNP(2, 2, MAXCNT) As String * 1  'OSF�ʓ����z�@2003/05/21 ooba
Public TmpOsfData(1, SXL_MAXSMP, MAXCNT) As String      'OSF����/�ő�l
Public TmpOsfMBNP(2, SXL_MAXSMP, MAXCNT) As String * 1  'OSF�ʓ����z
'Chg End   2011/03/09 SMPK Miyata
Public wiSmpGetFlg  As Integer              '����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
Public wiKcnt       As Integer              '�H���A��

'--------------- 2008/07/25 INSERT START  By Systech ---------------
Private pbGDJudgeTbl(3) As Boolean          ' GD���茋�ʑޔ�
'--------------- 2008/07/25 INSERT  END   By Systech ---------------

''Public STAFFIDBUFF  As String
''
''''�G���[���b�Z�[�W
''Public Const ESTAF = "ESTAF" ''�S���҃R�[�h�������ł��
''Public Const EIE00 = "EIE00" ''�S�Ẵf�[�^���͂��������Ă��܂���
''Public Const EIE01 = "EBLK1" ''�u���b�NID�̌������Ԉ���Ă��܂��
''Public Const EIM00 = "EIM00" ''�w���P����������с@�₢���킹���B
''Public Const EGET = "EGET" ''DB����̓Ǎ��Ɏ��s���܂����B
''Public Const EAPLY = "EAPLY" ''DB�ւ̏����Ɏ��s���܂����B
''Public Const EMAT1 = "EMAT1" '' �����ԍ��̌������Ԉ���Ă��܂��
''Public Const EMAT2 = "EMAT2" '' �w�肵�������ԍ��͖��o�^�ł��B
''Public Const KIE00 = "EBLK0" ''���͂��ꂽ�u���b�NID�ͤ���݂��܂���
''Public Const KDE01 = "KDE01" ''�w���P�����́A�C���[�W�\���ł��܂���B
''Public Const PWAIT = "PWAIT" ''���X���҂�������
''Public Const KC001 = "EKC01" ''�N���X�^���J�^���O���������s���܂����I
''Public Const TJE01 = "PJE01" ''��������NG�ł��B
''Public Const ESXL0 = "ESXL0" ''���͂��ꂽSXLID�́A���݂��܂���B"
''Public Const ESXL1 = "ESXL1" ''SXLID�̌������Ԉ���Ă��܂��B"
''Public Const EHIN1 = "EHIN1" ''�i�Ԃ̌������Ԉ���Ă��܂��B"
''Public Const EHIN0 = "EHIN0" ''�w��̕i�Ԃ͖��o�^�ł��B"



'''''' SXL�̑ΏۂƂȂ���ۯ��ۑ��p�\����
'''''Public Type typ_IntoBlock
'''''    SORTID As String
'''''    FULLID As String
'''''End Type

'''''' �u���b�N���
'''''Public Type typ_BlkInf
'''''    BLOCKID As String * 12      ' �u���b�NID
'''''    LENGTH As Integer           ' ����
'''''    REALLEN As Integer          ' ������
'''''    KRPROCCD As String * 5      ' ���݊Ǘ��H��
'''''    NOWPROC As String * 5       ' ���ݍH��
'''''    LPKRPROCCD As String * 5    ' �ŏI�ʉߊǗ��H��
'''''    LASTPASS As String * 5      ' �ŏI�ʉߍH��
'''''    RSTATCLS As String * 1      ' ������ԋ敪
'''''    SEED As String * 4          ' �V�[�h
'''''    COF As type_Coefficient     ' �ΐ͌W���v�Z
'''''    SAMPFLAG As Boolean         ' �T���v���擾�t���O
'''''End Type

''''''�J�b�g�ʒu�p�\����
'''''Public Type typ_CMKC001C
'''''    CRYNUM As String * 12       ' �����ԍ�
'''''    IngotPos As Integer         ' �������J�n�ʒu
'''''    LENGTH As Integer           ' ����
'''''End Type

'''''' �u���b�N���
'''''Public Type typ_BlkInf3
'''''    BLOCKID As String * 12      ' �u���b�NID
'''''    LENGTH As Integer           ' ����
'''''    REALLEN As Integer          ' ������
'''''    NOWPROC As String * 5       ' ���ݍH��
'''''    DELFLG As String * 1        ' �폜�敪
'''''    COF As type_Coefficient     ' �ΐ͌W���v�Z
'''''End Type

'''''Public tblHinMng() As typ_TBCME041                      ' �i�ԊǗ�
'''''Public tblWafSmp() As typ_TBCME044                      ' �v�e�T���v���Ǘ�
'''''Public tblBlkInf() As typ_BlkInf                        ' �u���b�N���e�[�u��
'''''Public tblTotal As typ_AllTypesC                        ' �O��ʂ���̏��ێ��\����
'''''Public tblWfSxlMng() As typ_TBCME042                    ' SXL�Ǘ��\����
'''''Public tblWfSxlMngS() As typ_TBCME042                   ' ����]���w���pSXL�Ǘ��\����
'''''Public tblWfSample() As typ_WfSampleGr                  ' WF�T���v���Ǘ�
'''''Public SxlIntoBlock() As typ_IntoBlock                  ' SXL�̑ΏۂƂȂ���ۯ��\����
'''''Public tblPrcList() As typ_TBCMB005                     ' �敪�p�R�[�h�}�X�^�[�\����
'''''Public tblHinbanRs() As type_DBDRV_scmzc_fcmlc001d_In   ' �i�ԏ��ێ��\����
'''''Public tblsiyou() As type_DBDRV_scmzc_fcmlc001d_WfSiyou ' �d�l���\����(�\���p)
'''''Public tblsmp() As type_DBDRV_scmzc_fcmlc001d_WfSmp     ' �T���v�����\����(�\���p)
'''''Public tblWfHantei As typ_TBCMW005                      ' WF�����������
'''''Public tblHuriHai() As typ_TBCMW006                     ' �U�֔p������
'''''Public tblSokuSizi() As typ_TBCMY003                    ' ����]�����@�w���\����
'''''Public tblSxlKSiji() As typ_TBCMY007                    ' �r�����m��w��
'''''Public NoTestHinList() As tFullHinban  ' �����̔������Ȃ��i��

'''''' �����w��
'''''Public Type typ_WafInd
'''''    BLOCKID As String * 12      ' �u���b�NID
'''''    BlockPos As Integer         ' �u���b�N�o
'''''''''    BkSampleId  As Variant      ' add 2003/03/28 hitec)matsumoto ���T���v��ID���擾
'''''    SAMPLEID    As Variant      ' add 2003/03/28 hitec)matsumoto �T���v��ID���擾
'''''    SAMPLEID2   As Variant      ' add 2003/03/28 hitec)matsumoto �T���v��ID2���擾
'''''    IngotPos As Integer         ' �����o
'''''    BkIngotPos  As Integer
'''''    LENGTH As Integer           ' ����
'''''    HINUP As tFullHinban        ' ��i��
'''''    HINDN As tFullHinban        ' ���i��
'''''    SMP As typ_WFSample         ' ��������
'''''    HinFlg As Boolean           ' �i�ԋ�؂�t���O
'''''    SMPFLG As Boolean           ' WF�T���v����؂�t���O
'''''    ERRDNFLG As Boolean         ' ���i�ԃG���[�t���O
'''''    SMPLKBN1 As String * 1      ' �T���v���敪�P
'''''    SMPLKBN2 As String * 1      ' �T���v���敪�Q
'''''End Type
'''''Public tblWafInd() As typ_WafInd        ' �����w���e�[�u��

'''''' �����E�F�n�[
'''''Public Type typ_LackMap
'''''    BLOCKID As String * 12      ' �u���b�NID
'''''    LACKPOSS As Double          ' �����ʒu(From)
'''''    LACKPOSE As Double          ' �����ʒu(To)
'''''    REJCAT As String * 1        ' �������R
'''''    LACKCNTS As Integer         ' ��������(From)
'''''    LACKCNTE As Integer         ' ��������(To)
'''''End Type
'''''Public tblLackMap() As typ_LackMap      ' �����E�F�n�[�e�[�u��


'''''' SXL�T���v�����
'''''Public Type typ_SxlSmp
'''''    strCRYNUM As String * 12          ' �����ԍ�
'''''    intINGOTPOS As Integer            ' �������J�n�ʒu
'''''    intLength As Integer              ' ����
'''''    strSXLID As String * 13           ' SXLID
'''''    strHINBAN As String * 12          ' �i��
'''''    strSMPLID As String * 16          ' �T���v��ID
'''''    intCount As Integer               ' ����
'''''    strSMPLUMU As String * 1          ' �T���v���L���敪
'''''    datREGDATE As Date                ' �o�^���t
'''''    datUPDDATE As Date                ' �X�V���t
'''''    strWFINDRS As String * 1          ' WF�����w���iRs)
'''''    strWFINDOI As String * 1          ' WF�����w���iOi)
'''''    strWFINDB1 As String * 1          ' WF�����w���iB1)
'''''    strWFINDB2 As String * 1          ' WF�����w���iB2�j
'''''    strWFINDB3 As String * 1          ' WF�����w���iB3)
'''''    strWFINDL1 As String * 1          ' WF�����w���iL1)
'''''    strWFINDL2 As String * 1          ' WF�����w���iL2)
'''''    strWFINDL3 As String * 1          ' WF�����w���iL3)
'''''    strWFINDL4 As String * 1          ' WF�����w���iL4)
'''''    strWFINDDS As String * 1          ' WF�����w���iDS)
'''''    strWFINDDZ As String * 1          ' WF�����w���iDZ)
'''''    strWFINDSP As String * 1          ' WF�����w���iSP)
'''''    strWFINDDO1 As String * 1         ' WF�����w���iDO1)
'''''    strWFINDDO2 As String * 1         ' WF�����w���iDO2)
'''''    strWFINDDO3 As String * 1         ' WF�����w���iDO3)
'''''    strWFRESRS As String * 1          ' WF�������сiRs)
'''''    strWFRESOI As String * 1          ' WF�������сiOi)
'''''    strWFRESB1 As String * 1          ' WF�������сiB1)
'''''    strWFRESB2 As String * 1          ' WF�������сiB2�j
'''''    strWFRESB3 As String * 1          ' WF�������сiB3)
'''''    strWFRESL1 As String * 1          ' WF�������сiL1)
'''''    strWFRESL2 As String * 1          ' WF�������сiL2)
'''''    strWFRESL3 As String * 1          ' WF�������сiL3)
'''''    strWFRESL4 As String * 1          ' WF�������сiL4)
'''''    strWFRESDS As String * 1          ' WF�������сiDS)
'''''    strWFRESDZ As String * 1          ' WF�������сiDZ)
'''''    strWFRESSP As String * 1          ' WF�������сiSP)
'''''    strWFRESDO1 As String * 1         ' WF�������сiDO1)
'''''    strWFRESDO2 As String * 1         ' WF�������сiDO2)
'''''    strWFRESDO3 As String * 1         ' WF�������сiDO3)
'''''End Type
'''''============================================================================================================================
'''''
''''''�T�v      :�p�����[�^�ݒ�
''''''���Ұ�    :�ϐ���        ,IO ,�^        ,����
''''''����      :�O��ʂ���̈�����ݒ肷��
''''''����      :
'''''Public Sub S_SetParamData()
'''''    typ_CType.typ_Param = typ_Param001b
'''''End Sub
'''''============================================================================================================================

'----------------------------------------------------------------------
'����
'----------------------------------------------------------------------
'�i��
'SelectBlkID
'tt(top,tail)
'�S���\����
'�d�l�����x���\����
'�d�l�����x���\����
'�g�[�^������
'�߂�l�i�z��j

'------------------------------------------------
' ��������
'------------------------------------------------

'�T�v      :���ђl�̑���������s���B
'���Ұ�    :�ϐ���          ,IO ,�^             :����
'          :sKeyID          ,I  ,String         :SXL-ID
'          :tNew_Hinban     ,I  ,String         :�U�֌��i��
'          :bTotalJudg      ,O  ,Boolean        :�g�[�^������
'          :iErr_Code       ,O  ,Integer        :�װ����(�߂�l�Ɠ���)
'          :sErr_Msg        ,O  ,String         :�װү���޺���
'          :typ_CType       ,O  ,typ_AllTypesC  :�S���\����(�\����)
'          :iSmpGetFlg      ,I  ,Integer        :����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
'          :sSamplID1       ,I  ,String         :TOP�����ID(�ȗ���)
'          :sSamplID2       ,I  ,String         :BOT�����ID(�ȗ���)
'          :iKcnt           ,I  ,Integer        :�H���A��(�ȗ���)
'          :�߂�l          ,O  ,Integer        :�擾�̐���(0:����I��, -1:�ُ�I��)
'����      :
'����      :2003/09/19 �V�K�쐬�@SB

Public Function funWfcSogoHantei(sKeyID As String, tNew_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_CType As typ_AllTypesC, _
                iSmpGetFlg As Integer, Optional sSamplID1 As String = vbNullString, Optional sSamplID2 As String = vbNullString, _
                Optional iKcnt As Integer = 0) As Integer
    
    Dim i       As Integer      'Add 2011/03/09 SMPK Miyata
    
    On Error GoTo Apl_down
    
    '�߂�l������
    funWfcSogoHantei = FUNCTION_RETURN_FAILURE
    
    '�O���[�o���ϐ��ɐݒ�
    wiSmpGetFlg = iSmpGetFlg
    wiKcnt = iKcnt
    
    '�����ݒ�
    sErr_Msg = "WFC��������(SetInitData)"
    If SetInitData(sKeyID, tNew_Hinban, typ_CType.typ_Param, iSmpGetFlg, sSamplID1, sSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '��ʏ��ݒ�
    sErr_Msg = "WFC��������(SetAllData)"
    If SetAllData(typ_CType, tNew_Hinban, iSmpGetFlg, iErr_Code, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
        
    TotalJudg = True
    MidlJudg = True             '���Ԕ�������   Add 2011/03/09 SMPK Miyata
    typ_CType.sMidErrMsg = ""   '���Ԕ����`�F�b�N�G���[���b�Z�[�W   Add 2011/05/10 SMPK Miyata

'    funWfcSogoHantei = FUNCTION_RETURN_FAILURE
'
    '�d�l�����w���擾
    sErr_Msg = "WFC��������(SpecJudgCheck)"
    SpecJudgCheck
    
    '2003/12/13 SystemBrain Null�Ή��ǉ���
    '�d�lNull�`�F�b�N
    sErr_Msg = "�d�lNull����"
    If funWfChkNull(typ_CType.typ_si, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    '2003/12/13 SystemBrain Null�Ή��ǉ���
    
    '���уf�[�^����(TOP)
    sErr_Msg = "WFC��������(����(TOP))"
    If WfAllJudg(typ_CType, tNew_Hinban, SxlTop) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '���уf�[�^����(TAIL)
    sErr_Msg = "WFC��������(����(TAIL))"
    If WfAllJudg(typ_CType, tNew_Hinban, SxlTail) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

'Add Start 2011/03/09 SMPK Miyata
    '���уf�[�^����(MIDLE)
    sErr_Msg = "WFC��������(����(MIDLE))"
    For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)
        If WfAllJudg(typ_CType, tNew_Hinban, i) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    Next i

    'Add Start 2011/07/19 Y.Hitomi
        '���Ԕ��������̃`�F�b�N
        Dim iMidCnt         As Integer       '���Ԕ����̖���
        Dim iSmpMai()       As Integer       '���������ۑ��z��
'Cng Start 2011/08/10 Y.Hitomi
            ' ���Ԕ����i�̏ꍇ
    If typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "2" Or typ_CType.typ_si.MSMPFLG = "3" Then
'        If typ_CType.typ_si.MSMPFLG = "1" Then
'Cng End   2011/08/10 Y.Hitomi
            ReDim iSxlPos(0)
            '���������̎擾
            If fncGetSmpMai(sKeyID, iSmpMai) = FUNCTION_RETURN_FAILURE Then
                MidlJudg = False
            Else
                
                For i = 0 To UBound(iSmpMai)
                    '�ŏI�ʒu�̖����`�F�b�N
                    If i = UBound(iSmpMai) Then
                    '�����̃`�F�b�N
'Cng Start 2011/10/25 Y.Hitomi
                        If iSmpMai(i) <= typ_CType.typ_si.MSMPTANIMAI Then
'                        If iSmpMai(i) >= typ_CType.typ_si.MSMPCONSTMAI And _
'                            iSmpMai(i) <= typ_CType.typ_si.MSMPTANIMAI Then
'Cng End 2011/10/25 Y.Hitomi
                        Else
                            If iSmpMai(i) > typ_CType.typ_si.MSMPTANIMAI And _
                                iSmpMai(i) < typ_CType.typ_si.MSMPCONSTMAI + typ_CType.typ_si.MSMPTANIMAI Then
                            Else
                                typ_CType.sMidErrMsg = "���Ԕ����������s�����Ă��܂��B���і���(" & CStr(iSmpMai(i)) & ")"
                                MidlJudg = False
                            End If
                        End If
                    Else
                    '�����̃`�F�b�N
'Cng Start 2011/08/25 Y.Hitomi
'                        If iSmpMai(i) >= typ_CType.typ_si.MSMPCONSTMAI And _
'                            iSmpMai(i) <= typ_CType.typ_si.MSMPTANIMAI Then
                    If iSmpMai(i) <= typ_CType.typ_si.MSMPTANIMAI Then
'Cng End   2011/08/25 Y.Hitomi
                        Else
                            typ_CType.sMidErrMsg = "���Ԕ����������s�����Ă��܂��B���і���(" & CStr(iSmpMai(i)) & ")"
                            MidlJudg = False
                            Exit For
                        End If
                    End If
                Next i
                
            End If
        End If
        
        Dim iMinMidCnt      As Integer       '���Ԕ����̕K�v��
        Dim iRstMidCnt      As Integer       '���Ԕ����̌���
            
'Cng Start 2011/08/10 Y.Hitomi
            ' ���Ԕ����i�̏ꍇ
    If typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "2" Or typ_CType.typ_si.MSMPFLG = "3" Then
'        If typ_CType.typ_si.MSMPFLG = "1" Then
'Cng End   2011/08/10 Y.Hitomi
                '���Ԕ����̕K�v�� = (SXL��WF���� - ���Ԕ������e�l(����)) / ���Ԕ����P��(����)
                iMinMidCnt = Fix((typ_CType.typ_Param.COUNT - typ_CType.typ_si.MSMPCONSTMAI) / typ_CType.typ_si.MSMPTANIMAI)
                '�}�C�i�X�̏ꍇ�A�O�Ƃ���
                If iMinMidCnt < 0 Then iMinMidCnt = 0
                
                '���Ԕ����̌���
                iRstMidCnt = (UBound(typ_CType.typ_Param.WFSMP) - SxlMidl) + 1
                If iRstMidCnt < iMinMidCnt Then
                    typ_CType.sMidErrMsg = "���Ԕ������т�����܂���B�@�d�l(" & iMinMidCnt & ") ����(" & iRstMidCnt & ")"
                    MidlJudg = False
                End If
            End If
            
'Add Start 2011/11/28 Y.Hitomi
        '���Ԕ���=�ۏ؂̏ꍇ�̂݁A����NG�Ƃ���B
    If typ_CType.typ_si.MSMPFLG = "2" Or typ_CType.typ_si.MSMPFLG = "0" Or typ_CType.typ_si.MSMPFLG = " " Then
                MidlJudg = True
        End If
'Add End   2011/11/28 Y.Hitomi


'Chg Start 2011/03/09 SMPK Miyata
'    bTotalJudg = TotalJudg
    bTotalJudg = TotalJudg And MidlJudg
'Chg End   2011/03/09 SMPK Miyata

    funWfcSogoHantei = FUNCTION_RETURN_SUCCESS
    
'------------------------------------------ �I������  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funWfcSogoHantei = -4
    iErr_Code = funWfcSogoHantei
    GoTo Apl_Exit
    
End Function

'�T�v      :��ʏ��f�[�^�ݒ�
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_CType     ,I  ,typ_AllTypesC ,�e���\����
'����      :��ʏ������\���̂ɐݒ肷��
'����      :
Private Function SetAllData(typ_CType As typ_AllTypesC, tNew_Hinban As tFullHinban, iSmpGetFlg As Integer, _
                                                                        iErr_Code As Integer, sErr_Msg As String) As FUNCTION_RETURN
    
    Dim typ_in      As type_DBDRV_scmzc_fcmlc001c_In     ' DB�A�N�Z�X���͗p
    Dim fret(2)     As FUNCTION_RETURN
    Dim RET         As FUNCTION_RETURN
    Dim typ_hi()    As typ_TBCMH004
    Dim sErrMsg     As String
    Dim FuncAns     As FUNCTION_RETURN ''2001/12/18 S.Sano
    Dim records()   As typ_TBCMH001
'Add Start 2011/03/07 SMPK Miyata
    Dim i           As Integer      '�J�E���^
    Dim iMidNo      As Integer      '���Ԕ���No
'Add End   2011/03/07 SMPK Miyata

    SetAllData = FUNCTION_RETURN_FAILURE
    
    'TOP��
    sErr_Msg = "WFC��������(TOP �����ް��ݒ�)"
    typ_in.HIN.hinban = typ_CType.typ_Param.hinban
    typ_in.HIN.factory = typ_CType.typ_Param.factory
    typ_in.HIN.mnorevno = typ_CType.typ_Param.REVNUM
    typ_in.HIN.opecond = typ_CType.typ_Param.opecond
    typ_in.SAMPLEID = typ_CType.typ_Param.WFSMP(SxlTop).REPSMPLIDCW
    typ_in.SXLID = typ_CType.typ_Param.SXLID
    typ_in.WFSMP = typ_CType.typ_Param.WFSMP(SxlTop)
    
    With typ_CType
        ReDim .typ_y013top(0)
        '�]�����茋�ʎ擾
        sErr_Msg = "WFC��������(TOP funWfcGetDataEtc)"

'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '' WF�d�l(SPV)�擾
        If funWfcGetDataEtc_SPV(tNew_Hinban, _
                                .typ_si, _
                                sErrMsg) <> FUNCTION_RETURN_SUCCESS Then
            'WF�d�l(SPV)�擾���s
            SetAllData = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------

        '�p�����[�^SxlTop�ǉ��@Add 2011/03/07 SMPK Miyata
        fret(SxlTop) = funWfcGetDataEtc(typ_in, SxlTop, tNew_Hinban, iSmpGetFlg, _
                                        .typ_si, _
                                        .typ_y013top(), _
                                        sErrMsg)
        If fret(SxlTop) = FUNCTION_RETURN_SUCCESS Then
            ' �]�����茋�ʐ���
            sErr_Msg = "WFC��������(TOP �]�����茋�ʐ���)"
            If SetMERInd(typ_CType, .typ_y013top(), SxlTop) <> True Then
                '�]�����茋�ʐ��񎸔s
                Exit Function
            End If
'''''            ' WF�����w���iRs)
'''''            If InStr("1345", .typ_Param.WFSMP(SxlTop).WFINDRSCW) <> 0 _
'''''            And .typ_Param.WFSMP(SxlTop).WFINDRSCW = "1" Then
'''''            End If
'''''            ' WF�����w���iOi)
'''''            If InStr("1345", .typ_Param.WFSMP(SxlTop).WFINDOICW) <> 0 _
'''''            And .typ_Param.WFSMP(SxlTop).WFINDOICW = "1" Then
'''''            End If
            '���グ�I�����ю擾
            ReDim typ_hi(0)
'��8���w���P�����������Ȃ� 2007/10/10 SETsw kubota
'            If Mid(.typ_Param.CRYNUM, 1, 1) <> "8" Then
                sErr_Msg = "WFC��������(TOP ���グ�I�����ю擾)"
                If s_cmmc001db_Sql(.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                    '���グ�I�����ю擾���s
                    Exit Function
                Else
                    If UBound(typ_hi) <> 0 Then
                        .typ_hage(SxlTop) = typ_hi(1)
                    Else
                        '���グ�I�����ю擾���s
                        SetAllData = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                End If
'            End If
        Else
            SetAllData = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    
            
    
        'TAIL��
        sErr_Msg = "WFC��������(TAIL �����ް��ݒ�)"
        typ_in.SAMPLEID = .typ_Param.WFSMP(SxlTail).REPSMPLIDCW
        typ_in.WFSMP = typ_CType.typ_Param.WFSMP(SxlTail)
    
        '�]�����茋�ʎ擾
        ReDim .typ_y013tail(0)
        sErr_Msg = "WFC��������(TAIL funWfcGetDataEtc)"
        '�p�����[�^SxlTail�ǉ��@Add 2011/03/07 SMPK Miyata
        fret(SxlTail) = funWfcGetDataEtc(typ_in, SxlTail, tNew_Hinban, iSmpGetFlg, _
                                         .typ_si, _
                                         .typ_y013tail(), _
                                         sErrMsg)
        If fret(SxlTail) = FUNCTION_RETURN_SUCCESS Then
            ' �]�����茋�ʐ���
            sErr_Msg = "WFC��������(TAIL �]�����茋�ʐ���)"
            If SetMERInd(typ_CType, .typ_y013tail(), SxlTail) <> True Then
                '�]�����茋�ʐ��񎸔s
                Exit Function
            End If
'''''            ' WF�����w���iRs)
'''''            If InStr("2345", .typ_Param.WFSMP(SxlTail).WFINDRSCW) <> 0 _
'''''            And .typ_Param.WFSMP(SxlTail).WFINDRSCW = "1" Then
'''''            End If
'''''            ' WF�����w���iOi)
'''''            If InStr("2345", .typ_Param.WFSMP(SxlTail).WFINDOICW) <> 0 _
'''''            And .typ_Param.WFSMP(SxlTail).WFINDOICW = "1" Then
'''''            End If
            '���グ�I�����ю擾
            ReDim typ_hi(0)
'��8���w���P�����������Ȃ� 2007/10/10 SETsw kubota
'            If Mid(.typ_Param.CRYNUM, 1, 1) <> "8" Then
                sErr_Msg = "WFC��������(TAIL ���グ�I�����ю擾)"
                If s_cmmc001db_Sql(.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                    '���グ�I�����ю擾���s
                    Exit Function
                Else
                    If UBound(typ_hi) <> 0 Then
                        .typ_hage(SxlTail) = typ_hi(1)
                    Else
                        '���グ�I�����ю擾���s
                        SetAllData = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                End If
'            End If
        Else
            SetAllData = FUNCTION_RETURN_FAILURE
            Exit Function
        End If

'Add Start 2011/03/07 SMPK Miyata
        For i = SxlMidl To UBound(.typ_Param.WFSMP)
            iMidNo = i - SxlMidl + 1
            
            If iMidNo > SXL_MAXSMP Then
                ' ���Ԕ����ő匏���I�[�o�[
                Exit Function
            End If

            'MIDLE��
            sErr_Msg = "WFC��������(MIDLE_" & iMidNo & " �����ް��ݒ�)"
            typ_in.SAMPLEID = .typ_Param.WFSMP(i).REPSMPLIDCW
            typ_in.WFSMP = typ_CType.typ_Param.WFSMP(i)
        
            '�]�����茋�ʎ擾
            ReDim Preserve .typ_y013midl_ary(iMidNo)
            sErr_Msg = "WFC��������(MIDLE_" & iMidNo & " funWfcGetDataEtc)"
            RET = funWfcGetDataEtc(typ_in, i, tNew_Hinban, iSmpGetFlg, _
                                    .typ_si, _
                                    .typ_y013midl_ary(iMidNo).typ_y013midl, _
                                    sErrMsg)
            If RET = FUNCTION_RETURN_SUCCESS Then

                ' �]�����茋�ʐ���
                sErr_Msg = "WFC��������(MIDLE_" & iMidNo & " �]�����茋�ʐ���)"
                If SetMERInd(typ_CType, .typ_y013midl_ary(iMidNo).typ_y013midl, i) <> True Then
                    '�]�����茋�ʐ��񎸔s
                    Exit Function
                End If
                
                '���グ�I�����ю擾
                ReDim typ_hi(0)
                sErr_Msg = "WFC��������(MIDLE_" & iMidNo & " ���グ�I�����ю擾)"
                If s_cmmc001db_Sql(.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                    '���グ�I�����ю擾���s
                    Exit Function
                Else
                    If UBound(typ_hi) <> 0 Then
                        .typ_hage(i) = typ_hi(1)
                    Else
                        '���グ�I�����ю擾���s
                        SetAllData = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                End If
            Else
                SetAllData = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        Next i
'Add End   2011/03/07 SMPK Miyata
    End With
    
''2001/12/18 S.Sano Start
    '' �o�{�����̔��f
    sErr_Msg = "WFC��������(P+�����̔��f)"
    '2004.09.09 Y.K �R�t���ύX�i���Ɏ擾�f�[�^�͎g�p���Ă��Ȃ��݂��������C�������j
'    If DBDRV_GetTBCMH001(records(), "where UPINDNO = '" & Left(SelectSxlID, 7) & "00" & "'") = FUNCTION_RETURN_SUCCESS Then
    If DBDRV_GetTBCMH001(records(), "where UPINDNO = '" & left(SelectSxlID, 7) & "0" & Mid(SelectSxlID, 9, 1) & "'") = FUNCTION_RETURN_SUCCESS Then

'�v�e�T���v�������ύX 2003.05.20 yakimura
'        If Left(SelectSxlID, 1) <> "8" Then
'            bPPlus = ((records(1).AMRESIST <= CDbl(GetCodeField("LG", "02", "P+", "INFO1"))) And (typ_CType.typ_si.HWFTYPE = "P"))
'            bNPlus = ((records(1).AMRESIST <= CDbl(GetCodeField("LG", "02", "N+", "INFO1"))) And (typ_CType.typ_si.HWFTYPE = "N"))
'        Else
'            bPPlus = False
'            bNPlus = False
'        End If
'�v�e�T���v�������ύX 2003.05.20 yakimura
    
    Else
        SetAllData = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
''2001/12/18 S.Sano End
    
    SetAllData = FUNCTION_RETURN_SUCCESS
End Function

'�T�v      :����]�����ʐݒ�
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_a         ,IO ,typ_AllTypesC ,�e���\����
'          :typ_y013()    ,I  ,typ_TBCMY013 ,����]�����ʏ��\����
'          :tt            ,I  ,Integer      ,TOP�ETAIL
'          :�߂�l        ,O  ,Integer      ,True:����I���@False:�ُ�I��
'����      :����]�����ʔz���DB�����������R�[�h�𐮗񂷂�
'����      :
Private Function SetMERInd(typ_CType As typ_AllTypesC, _
                          typ_y013() As typ_TBCMY013, _
                          tt As Integer) As Boolean
    Dim i As Integer
    
    With typ_CType
        For i = 1 To UBound(typ_y013)
            Select Case Trim(typ_y013(i).Spec)
            Case OSWFRES ' RES
                .typ_y013(tt, WFRES) = typ_y013(i)
            Case OSWFOI ' OI
                .typ_y013(tt, WFOI) = typ_y013(i)
            Case OSWFBMD1 ' BMD1
                .typ_y013(tt, WFBMD1) = typ_y013(i)
            Case OSWFBMD2 ' BMD2
                .typ_y013(tt, WFBMD2) = typ_y013(i)
            Case OSWFBMD3 ' BMD3
                .typ_y013(tt, WFBMD3) = typ_y013(i)
            Case OSWFOSF1 ' OSF1
                .typ_y013(tt, WFOSF1) = typ_y013(i)
            Case OSWFOSF2 ' OSF2
                .typ_y013(tt, WFOSF2) = typ_y013(i)
            Case OSWFOSF3 ' OSF3
                .typ_y013(tt, WFOSF3) = typ_y013(i)
'            Del 2010/01/07 SIRD�Ή� Y.Hitomi
'            Case OSWFOSF4 ' OSF4
'                .typ_y013(tt, WFOSF4) = typ_y013(i)
            Case OSWFDS ' DSOD
                .typ_y013(tt, WFDS) = typ_y013(i)
            Case OSWFDZ ' DZ
                .typ_y013(tt, WFDZ) = typ_y013(i)
            
        ''Upd start 2005/06/21 (TCS)T.Terauchi  SPV9�_�Ή�  SPV��SPV����(TBCMJ016)���擾����ׁA�R�����g
'            Case OSWFSP ' SPV
'                .typ_y013(tt, WFSP) = typ_y013(i)
        ''Upd end   2005/06/21 (TCS)T.Terauchi  SPV9�_�Ή�  SPV��SPV����(TBCMJ016)���擾����ׁA�R�����g
            
            Case OSWFDOI1 ' DOI1
                .typ_y013(tt, WFDOI1) = typ_y013(i)
            Case OSWFDOI2 ' DOI2
                .typ_y013(tt, WFDOI2) = typ_y013(i)
            Case OSWFDOI3 ' DOI3
                .typ_y013(tt, WFDOI3) = typ_y013(i)
            Case OSWFOT1 ' OT1
                .typ_y013(tt, WFOT1) = typ_y013(i)
            Case OSWFOT2 ' OT2
                .typ_y013(tt, WFOT2) = typ_y013(i)
            ''�c���_�f�ǉ��@03/12/15 ooba
            Case OSWFAOI ' AOI
                .typ_y013(tt, WFAOI) = typ_y013(i)
            
            'Add 2010/01/07 SIRD�Ή� Y.Hitomi
            Case OSWFSIRD ' SIRD
                .typ_y013(tt, WFSIRD) = typ_y013(i)

            End Select
        Next
    End With
    SetMERInd = True
End Function

Private Sub SpecJudgCheck()
    Dim IND As String * 4               '�����w��
    Dim c0  As Integer
    
    With typ_CType
'test Git 2014/09/24   

'�v�e�T���v�������ύX 2003.05.20 yakimura
'        JudgSW.rs = (.typ_si.HWFRHWYS = "X")
'        JudgSW.Oi = (.typ_si.HWFONHWS = "X")
'        JudgSW.B1 = (.typ_si.HWFBM1HS = "X")
'        JudgSW.B2 = (.typ_si.HWFBM2HS = "X")
'        JudgSW.B3 = (.typ_si.HWFBM3HS = "X")
'        JudgSW.L1 = (.typ_si.HWFOF1HS = "X")
'        JudgSW.L2 = (.typ_si.HWFOF2HS = "X")
'        JudgSW.L3 = (.typ_si.HWFOF3HS = "X")
'        JudgSW.L4 = (.typ_si.HWFOF4HS = "X")
'        JudgSW.Dsod = (.typ_si.HWFDSOHS = "X")
'        JudgSW.Dz = (.typ_si.HWFMKHWS = "X")
'        JudgSW.Doi1 = (.typ_si.HWFOS1HS = "X")
'        JudgSW.Doi2 = (.typ_si.HWFOS2HS = "X")
'        JudgSW.Doi3 = (.typ_si.HWFOS3HS = "X")
'        JudgSW.sp = (.typ_si.HWFSPVHS = "X") Or (.typ_si.HWFDLHWS = "X")
        
        JudgSW.rs = (.typ_si.HWFRHWYS = "H")
        JudgSW.Oi = (.typ_si.HWFONHWS = "H")
        JudgSW.B1 = (.typ_si.HWFBM1HS = "H")
        JudgSW.B2 = (.typ_si.HWFBM2HS = "H")
        JudgSW.B3 = (.typ_si.HWFBM3HS = "H")
        JudgSW.L1 = (.typ_si.HWFOF1HS = "H")
        JudgSW.L2 = (.typ_si.HWFOF2HS = "H")
        JudgSW.L3 = (.typ_si.HWFOF3HS = "H")
        JudgSW.L4 = (.typ_si.HWFOF4HS = "H")
        JudgSW.Dsod = (.typ_si.HWFDSOHS = "H")
        JudgSW.DZ = (.typ_si.HWFMKHWS = "H")
        JudgSW.Doi1 = (.typ_si.HWFOS1HS = "H")
        JudgSW.Doi2 = (.typ_si.HWFOS2HS = "H")
        JudgSW.Doi3 = (.typ_si.HWFOS3HS = "H")
        JudgSW.sp = (.typ_si.HWFSPVHS = "H") Or (.typ_si.HWFDLHWS = "H")
        JudgSW.AOI = (.typ_si.HWFZOHWS = "H")           '�c���_�f�ǉ��@03/12/09 ooba
        'GD�ǉ��@05/01/27 ooba
        JudgSW.GD = (.typ_si.HWFDENHS = "H") Or (.typ_si.HWFLDLHS = "H") Or _
                    (.typ_si.HWFDVDHS = "H")
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
        JudgSW.B1E = (.typ_si.HEPBM1HS = "H")
        JudgSW.B2E = (.typ_si.HEPBM2HS = "H")
        JudgSW.B3E = (.typ_si.HEPBM3HS = "H")
        JudgSW.L1E = (.typ_si.HEPOF1HS = "H")
        JudgSW.L2E = (.typ_si.HEPOF2HS = "H")
        JudgSW.L3E = (.typ_si.HEPOF3HS = "H")
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'�v�e�T���v�������ύX 2003.05.20 yakimura
        
'''''        For c0 = 1 To 2
'''''            IND = IIf(c0 = SxlTop, "1346", "2346")
'''''            MeasFlag(c0).B1 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDB1CW) <> 0)
'''''            MeasFlag(c0).B2 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDB2CW) <> 0)
'''''            MeasFlag(c0).B3 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDB3CW) <> 0)
'''''            MeasFlag(c0).Doi1 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDDO1CW) <> 0)
'''''            MeasFlag(c0).Doi2 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDDO2CW) <> 0)
'''''            MeasFlag(c0).Doi3 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDDO3CW) <> 0)
'''''            MeasFlag(c0).Dsod = (InStr(IND, .typ_Param.WFSMP(c0).WFINDDSCW) <> 0)
'''''            MeasFlag(c0).Dz = (InStr(IND, .typ_Param.WFSMP(c0).WFINDDZCW) <> 0)
'''''            MeasFlag(c0).L1 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDL1CW) <> 0)
'''''            MeasFlag(c0).L2 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDL2CW) <> 0)
'''''            MeasFlag(c0).L3 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDL3CW) <> 0)
'''''            MeasFlag(c0).L4 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDL4CW) <> 0)
'''''            MeasFlag(c0).Oi = (InStr(IND, .typ_Param.WFSMP(c0).WFINDOICW) <> 0)
'''''            MeasFlag(c0).rs = (InStr(IND, .typ_Param.WFSMP(c0).WFINDRSCW) <> 0)
'''''            MeasFlag(c0).sp = (InStr(IND, .typ_Param.WFSMP(c0).WFINDSPCW) <> 0)
'''''        Next
        
    End With
End Sub

'�T�v      :��������(�S)
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :typ_CType     ,I  ,typ_AllTypesC    ,�e���\����
'          :tNew_Hinban   ,I  ,tFullHinban      :�U�֌��i��
'          :tt            ,I  ,Integer          ,TopTail����p
'����      :�����w���ɏ]���A���є�����s��
'����      :
Public Function WfAllJudg(typ_CType As typ_AllTypesC, tNew_Hinban As tFullHinban, tt As Integer) As FUNCTION_RETURN
    
    Dim IND         As String * 4                  '�����w��
    Dim bJudg       As Boolean
    Dim i           As Integer
    Dim typTmList() As typ_TBCMB005
'Chg Start 2011/03/09 SMPK Miyata
'    Dim INGOTPOS(2) As Integer
    Dim INGOTPOS(SXL_MAXSMP) As Integer
'Chg End   2011/03/09 SMPK Miyata
    Dim vTemp       As Variant
    Dim sHinban12   As String                               '�i��(12��)
    Dim sSxlPos     As String       'SXL�ʒu(TOP/BOT)�@04/04/12 ooba

    i = 0
    WfAllJudg = FUNCTION_RETURN_FAILURE
    
    sHinban12 = tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond
    
    If tt = SxlTop Then
        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS
'Chg Start 2011/03/09 SMPK Miyata
'    Else
'        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
    ElseIf tt = SxlTail Then
        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
    Else
        INGOTPOS(tt) = typ_CType.typ_Param.WFSMP(tt).INPOSCW
'Chg End   2011/03/09 SMPK Miyata
    End If
    
    '�����w���ݒ�
    If tt = SxlTop Then
        IND = "123"
    Else
        IND = "123"
    End If

'Chg Start 2011/03/09 SMPK Miyata
'    sSxlPos = IIf(tt = SxlTop, "TOP", "BOT")        '04/04/12 ooba
    Select Case tt
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/09 SMPK Miyata
    
    '�����R�[�h���X�g�擾
    If GetCodeList(MSYSCLASS, KCLASS, typTmList()) <> FUNCTION_RETURN_SUCCESS Then
        '�����R�[�h���X�g�擾���s
        Exit Function
    End If
    
    With typ_CType
        '' WF�����w���iRs)*****************************************************************
'        If JudgSW.rs Then
        '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Cng Start 2011/08/10 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.rs And CheckKHN(.typ_si.HWFRKHNN, 1, sSxlPos) Then
        '�ʏ픲���F�ۏؕ��@=�ۏ� ���� �����p�x�Q���`�F�b�N�L��̏ꍇ�A�d�l�L�Ƃ���
        '���Ԕ����F�ۏؕ��@=�ۏ� ���� ���Ԕ����w���L�肩�@���Ԕ����t���O���ۏ؂̏ꍇ�A�d�l�L�Ƃ���
        If (JudgSW.rs And CheckKHN(.typ_si.HWFRKHNN, 1, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.rs And .typ_si.MSMPFLGWFR = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
           (tt >= SxlMidl)) Then
'        If (JudgSW.rs And CheckKHN(.typ_si.HWFRKHNN, 1, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
'           (JudgSW.rs And .typ_si.MSMPFLGWFR = "1" And (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Cng End  2011/08/10 Y.Hitomi
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDRSCW) <> 0 Then
'                If .typ_Param.WFSMP(tt).WFRESRS1CW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESRS1CW = "1") And (Trim(.typ_y013(tt, WFRES).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    '���R����
                    If WfCrResJudg(typ_CType, .typ_si, .bOKNG(tt), .dblScut(tt), tt) Then
                        JiltusekiUmu(tt, WFRES) = True '2001/12/19 S.Sano
                    End If
                Else
                    ' �T���v���������ꍇ�́ANG�Ƃ��ĕ\��
                    .bOKNG(tt) = False
                End If
                If .bOKNG(tt) = False Then
'Chg Start 2011/03/09 SMPK Miyata
'                    TotalJudg = False
                    If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
                    
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                    gsTbcmy028ErrCode = "00130"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
                End If
            Else 'If .typ_Param.WFSMP(tt).WFRESRS = "2" Then
                ' �w���������ꍇ�́ANG�Ƃ��ĕ\��
                .bOKNG(tt) = False
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata

'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00130"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
                        
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDRSCW) <> 0 Then
'                If .typ_Param.WFSMP(tt).WFRESRS1CW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESRS1CW = "1") And (Trim(.typ_y013(tt, WFRES).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    '���R����
                    If WfCrResJudg(typ_CType, .typ_si, .bOKNG(tt), .dblScut(tt), tt) Then
                        JiltusekiUmu(tt, WFRES) = True '2001/12/19 S.Sano
                    End If
                Else
                    ' �T���v���������ꍇ�́AOK�Ƃ��ĕ\��
                    .bOKNG(tt) = True
                End If
                'Add Start 2011/11/28 Y.Hitomi ���Ԕ���(�ۏ؁j�̏ꍇ�́A�Q�l�\��
                If sSxlPos = "MID" And JudgSW.rs And .bOKNG(tt) = False And _
                   (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") Then
'                If sSxlPos = "MID" And JudgSW.rs And .bOKNG(tt) = False Then
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
            Else
                .bOKNG(tt) = True
            End If
            
        End If

        '' WF�����w���iOi)*****************************************************************
'        If JudgSW.OI Then
        '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/08/10 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.Oi And CheckKHN(.typ_si.HWFONKHN, 2, sSxlPos) Then
        '�ʏ픲���F�ۏؕ��@=�ۏ� ���� �����p�x�Q���`�F�b�N�L��̏ꍇ�A�d�l�L�Ƃ���
        '���Ԕ����F�ۏؕ��@=�ۏ� ���� ���Ԕ���(�ۏ�)�̏ꍇ�A�d�l�L�Ƃ���
        If (JudgSW.Oi And CheckKHN(.typ_si.HWFONKHN, 2, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.Oi And .typ_si.MSMPFLGWFO = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
           (tt >= SxlMidl)) Then
'        If (JudgSW.Oi And CheckKHN(.typ_si.HWFONKHN, 2, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
'           (JudgSW.Oi And .typ_si.MSMPFLGWFO = "1" And (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Chg End   2011/08/10 Y.Hitomi
            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).pos = -1                                               ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())               ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                                       ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                                       ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                             ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                             ' ���S
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
            '5�Ԗڂ̏��FAN���x��ǉ�
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                             ' ���T
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' ���6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' ���7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' ���8
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                            ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                                     ' �i��(12��)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDOICW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFOI).SAMPLEID              ' �T���v���m��
'                If .typ_Param.WFSMP(tt).WFRESOICW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESOICW = "1") And (Trim(.typ_y013(tt, WFOI).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'OI���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                    'OI����
'Cng Start 2011/08/01 Y.Hitomi
                    If WfCrOiJudg(.typ_si, .typ_y013(tt, WFOI), bJudg, sSxlPos) Then
'                    If WfCrOiJudg(.typ_si, .typ_y013(tt, WFOI), bJudg) Then
'Cng Start 2011/08/01 Y.Hitomi
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA1)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' ���1
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA13)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' ���2
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA12)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")     ' ���3
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA11)
                        'ORG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                        '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' ���4
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' ���4
                        JiltusekiUmu(tt, WFOI) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(.typ_y013(tt, WFOI).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESOICW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                       ' �������J�n�ʒu
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata

'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00131"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
            
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDOICW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())           ' ���e
                .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                '5�Ԗڂ̏��FAN���x��ǉ�
                typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                typ_rslt_ex(tt, i).INFO5 = ""        ' ���5
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                typ_rslt_ex(tt, i).INFO6 = ""                                       ' ���6
                typ_rslt_ex(tt, i).INFO7 = ""                                       ' ���7
                typ_rslt_ex(tt, i).INFO8 = ""                                       ' ���8
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFOI).SAMPLEID              ' �T���v���m��
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
                .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                bJudg = False
'                If .typ_Param.WFSMP(tt).WFRESOICW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESOICW = "1") And (Trim(.typ_y013(tt, WFOI).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'OI���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                    'OI����
'Cng Start 2011/08/01 Y.Hitomi
                    If WfCrOiJudg(.typ_si, .typ_y013(tt, WFOI), bJudg, sSxlPos) Then
'                    If WfCrOiJudg(.typ_si, .typ_y013(tt, WFOI), bJudg) Then
'Cng Start 2011/08/01 Y.Hitomi
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())   ' ���e
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA1)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' ���1
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA13)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' ���2
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA12)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")     ' ���3
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA11)
                        'ORG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                        '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' ���4
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' ���4
                        JiltusekiUmu(tt, WFOI) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(.typ_y013(tt, WFOI).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESOICW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
                'Add Start 2011/11/28 Y.Hitomi ���Ԕ����̏ꍇ�́A�Q�l�\������
                If sSxlPos = "MID" And JudgSW.Oi And bJudg = False Then
                    .typ_rslt(tt, i).OKNG = "�Q�l"                                  ' ���茋��
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
                i = i + 1
            End If
        End If

        '' ���������w��(B1)*****************************************************************
        BMDDataSet 1, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(B2)*****************************************************************
        BMDDataSet 2, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(B3)*****************************************************************
        BMDDataSet 3, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(L1)*****************************************************************
        OSFDataSet 1, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(L2)*****************************************************************
        OSFDataSet 2, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(L3)*****************************************************************
        OSFDataSet 3, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        
        '' ���������w��(L4)*****************************************************************
    'Del 2010/01/07 SIRD�Ή� Y.Hitomi
'        OSFDataSet 4, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        
        '' WF�����w���iDsod)*****************************************************************
'        If JudgSW.Dsod Then
        '�ۏؕ��@�����ǉ��@04/04/12 ooba

'Chg Start 2011/08/10 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.Dsod And CheckKHN(.typ_si.HWFDSOKN, 13, sSxlPos) Then
        '�ʏ픲���F�ۏؕ��@=�ۏ� ���� �����p�x�Q���`�F�b�N�L��̏ꍇ�A�d�l�L�Ƃ���
        '���Ԕ����F�ۏؕ��@=�ۏ� ���� ���Ԕ���=�ۏ؂̏ꍇ�A�d�l�L�Ƃ���
        If (JudgSW.Dsod And CheckKHN(.typ_si.HWFDSOKN, 13, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.Dsod And .typ_si.MSMPFLGWFDS = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
            (tt >= SxlMidl)) Then
'        If (JudgSW.Dsod And CheckKHN(.typ_si.HWFDSOKN, 13, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
'           (JudgSW.Dsod And .typ_si.MSMPFLGWFDS = "1" And (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Chg End   2011/08/10 Y.Hitomi

            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).pos = -1              ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("DS", typTmList())               ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                                       ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                                       ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                             ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                             ' ���S
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
            '5�Ԗڂ̏��FAN���x��ǉ�
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' ���5
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' ���6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' ���7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' ���8
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                            ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                                     ' �i��(12��)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDDSCW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFDS).SAMPLEID              ' �T���v���m��
                'DS����擾
'                If .typ_Param.WFSMP(tt).WFRESDSCW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESDSCW = "1") And (Trim(.typ_y013(tt, WFDS).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'DS���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                    'DS����擾
                    If WfCrDsodjudg(.typ_si, .typ_y013(tt, WFDS), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(.typ_y013(tt, WFDS).MESDATA1)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' ���P
                        .typ_rslt(tt, i).INFO2 = ""                                 ' ���Q
                        .typ_rslt(tt, i).INFO3 = ""                                 ' ���R
'                        .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
                        'DSOD����ݕ\���ǉ��@04/07/28 ooba START ==========================================================================>
                        vTemp = CVar(IIf(Trim(.typ_y013(tt, WFDS).MESDATA4) = "", "-", Trim(.typ_y013(tt, WFDS).MESDATA4)) _
                                        & "  " & IIf(Trim(.typ_y013(tt, WFDS).MESDATA7) = "", "-", Trim(.typ_y013(tt, WFDS).MESDATA7))) & "   "
                        .typ_rslt(tt, i).INFO4 = vTemp                              ' ���S
                        'DSOD����ݕ\���ǉ��@04/07/28 ooba END ============================================================================>
                        JiltusekiUmu(tt, WFDS) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(.typ_y013(tt, WFDS).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESDSCW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata

'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00143"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
            
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDDSCW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("DS", typTmList())           ' ���e
                .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                '5�Ԗڂ̏��FAN���x��ǉ�
                typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                typ_rslt_ex(tt, i).INFO5 = ""                                       ' ���5
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                typ_rslt_ex(tt, i).INFO6 = ""                                       ' ���6
                typ_rslt_ex(tt, i).INFO7 = ""                                       ' ���7
                typ_rslt_ex(tt, i).INFO8 = ""                                       ' ���8
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFDS).SAMPLEID              ' �T���v���m��
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
                .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                bJudg = False
                'DS����擾
'                If .typ_Param.WFSMP(tt).WFRESDSCW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESDSCW = "1") And (Trim(.typ_y013(tt, WFDS).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'DS���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                    If WfCrDsodjudg(.typ_si, .typ_y013(tt, WFDS), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(.typ_y013(tt, WFDS).MESDATA1)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' ���P
                        .typ_rslt(tt, i).INFO2 = ""                                 ' ���Q
                        .typ_rslt(tt, i).INFO3 = ""                                 ' ���R
                        .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
                        JiltusekiUmu(tt, WFDS) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(.typ_y013(tt, WFDS).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESDSCW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = ""                                     ' ���R
                    .typ_rslt(tt, i).INFO4 = "����وُ�"                            ' ���S
                End If
                'Add Start 2011/11/28 Y.Hitomi ���Ԕ����̏ꍇ�́A�Q�l�\������
                If sSxlPos = "MID" And JudgSW.Dsod And bJudg = False Then
                    .typ_rslt(tt, i).OKNG = "�Q�l"
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
                i = i + 1
            End If
        End If
        
        
        '' WF�����w���iDZ)*****************************************************************
'        If JudgSW.DZ Then
        '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/08/10 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.DZ And CheckKHN(.typ_si.HWFMKKHN, 14, sSxlPos) Then
        '�ʏ픲���F�ۏؕ��@=�ۏ� ���� �����p�x�Q���`�F�b�N�L��̏ꍇ�A�d�l�L�Ƃ���
        '���Ԕ����F�ۏؕ��@=�ۏ� ���� ���Ԕ���=�ۏ؂̏ꍇ�A�d�l�L�Ƃ���
        If (JudgSW.DZ And CheckKHN(.typ_si.HWFMKKHN, 14, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.DZ And .typ_si.MSMPFLGWFDZ = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
           (tt >= SxlMidl)) Then
'        If (JudgSW.DZ And CheckKHN(.typ_si.HWFMKKHN, 14, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
'           (JudgSW.DZ And .typ_si.MSMPFLGWFDZ = "1" And (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Chg End   2011/08/10 Y.Hitomi

            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).pos = -1                                               ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("DZ", typTmList())               ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                                       ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                                       ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                             ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                             ' ���S
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
            '5�Ԗڂ̏��FAN���x��ǉ�
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' ���5
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
'����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' ���6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' ���7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' ���8
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                            ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                                     ' �i��(12��)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDDZCW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFDZ).SAMPLEID              ' �T���v���m��
                'DZ����擾
'                If .typ_Param.WFSMP(tt).WFRESDZCW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESDZCW = "1") And (Trim(.typ_y013(tt, WFDZ).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'DZ����擾
                    If WfCrDzjudg(.typ_si, .typ_y013(tt, WFDZ), bJudg) Then
                        'DZ���莸�s
                        .typ_rslt(tt, i).INFO3 = "����Err"                          ' ���Q
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA5)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")      ' ���1
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA6)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")      ' ���2
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA7)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")      ' ���3
                        .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
                        JiltusekiUmu(tt, WFDZ) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(.typ_y013(tt, WFDZ).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESDZCW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00144"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
            
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDDZCW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("DZ", typTmList())           ' ���e
                .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                '5�Ԗڂ̏��FAN���x��ǉ�
                typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                typ_rslt_ex(tt, i).INFO5 = ""                                       ' ���5
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                typ_rslt_ex(tt, i).INFO6 = ""                                       ' ���6
                typ_rslt_ex(tt, i).INFO7 = ""                                       ' ���7
                typ_rslt_ex(tt, i).INFO8 = ""                                       ' ���8
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFDZ).SAMPLEID              ' �T���v���m��
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
                .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                bJudg = False
                'DZ����擾
'                If .typ_Param.WFSMP(tt).WFRESDZCW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESDZCW = "1") And (Trim(.typ_y013(tt, WFDZ).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'DZ���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                    'DZ����擾
                    If WfCrDzjudg(.typ_si, .typ_y013(tt, WFDZ), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA5)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")      ' ���1
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA6)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")      ' ���2
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA7)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")      ' ���3
                        .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
                        JiltusekiUmu(tt, WFDZ) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(.typ_y013(tt, WFDZ).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESDZCW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
                'Add Start 2011/11/28 Y.Hitomi ���Ԕ����̏ꍇ�́A�Q�l�\������
                If sSxlPos = "MID" And JudgSW.DZ And bJudg = False Then
                    .typ_rslt(tt, i).OKNG = "�Q�l"
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
                i = i + 1
            End If
        End If
        
                
    ''Upd start 2005/06/21 (TCS)t.terauchi      SPV9�_�Ή�  Fe�Z�x�E�g�U���ɕ����ĕ\��

'        '' WF�����w���iSP)*****************************************************************
'        '�ۏؕ��@�����ǉ��@04/04/12 ooba
'        JudgSW.sp = ((.typ_si.HWFSPVHS = "H") And CheckKHN(.typ_si.HWFSPVKN, 15, sSxlPos)) _
'                    Or ((.typ_si.HWFDLHWS = "H") And CheckKHN(.typ_si.HWFDLKHN, 16, sSxlPos))
'        If JudgSW.sp Then
'
'            '��ʕ\�����e�ݒ�
'            .typ_rslt(tt, i).pos = -1                                               ' �������J�n�ʒu
'            .typ_rslt(tt, i).NAIYO = Search_CrCode("SP", typTmList())               ' ���e
'            .typ_rslt(tt, i).INFO1 = "�d�l�L"                                       ' ���P
'            .typ_rslt(tt, i).INFO2 = "������"                                       ' ���Q
'            .typ_rslt(tt, i).INFO3 = ""                                             ' ���R
'            .typ_rslt(tt, i).INFO4 = ""                                             ' ���S
'            .typ_rslt(tt, i).SMPLID = -1                                            ' �T���v���m��
'            .typ_rslt(tt, i).OKNG = "NG"                                            ' ���茋��
'            .typ_rslt(tt, i).hinban = sHinban12                                     ' �i��(12��)
'            bJudg = False
'            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
'                '��ʕ\�����e�ݒ�
'                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
'                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
'                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
'                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFSP).SAMPLEID              ' �T���v���m��
'                'SP����擾
''                If .typ_Param.WFSMP(tt).WFRESSPCW = "1" Then
'                If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(.typ_y013(tt, WFSP).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
'                    'LT���莸�s
'                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
'                    'SP����擾
'                    If WfCrSpvjudg(.typ_si, .typ_y013(tt, WFSP), bJudg) Then
'                        '��ʕ\�����e�ݒ�
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA5)
'                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' ���P
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA4)
'                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' ���Q
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA3)
'                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' ���R
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA2)
'                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' ���S
'                        JiltusekiUmu(tt, WFSP) = True '2001/12/19 S.Sano
'                    End If
'                ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
'                    '��ʕ\�����e�ݒ�
'                    bJudg = False
'                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
'                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
'                End If
'            End If
'            If bJudg = True Then
'                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
'            Else
'                .typ_rslt(tt, i).OKNG = "NG"                                        ' ���茋��
'                TotalJudg = False
'            End If
'            i = i + 1
'
'        Else
'            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 And .typ_Param.WFSMP(tt).WFRESSPCW = "1" Then
'
'                '��ʕ\�����e�ݒ�
'                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
'                .typ_rslt(tt, i).NAIYO = Search_CrCode("SP", typTmList())           ' ���e
'                .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
'                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
'                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
'                .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
'                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFSP).SAMPLEID              ' �T���v���m��
'                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
'                .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
'                'SP����擾
''                If .typ_Param.WFSMP(tt).WFRESSPCW = "1" Then
'                If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(.typ_y013(tt, WFSP).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
'                    'LT���莸�s
'                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
'                    'SP����擾
'                    If WfCrSpvjudg(.typ_si, .typ_y013(tt, WFSP), bJudg) Then
'                        '��ʕ\�����e�ݒ�
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA5)
'                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' ���P
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA4)
'                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' ���Q
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA3)
'                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' ���R
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA2)
'                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' ���S
'                        JiltusekiUmu(tt, WFSP) = True '2001/12/19 S.Sano
'                    End If
'                ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
'                    '��ʕ\�����e�ݒ�
'                    bJudg = False
'                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
'                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
'                End If
'                i = i + 1
'            End If
'        End If

        ''Fe�Z�x***************
'Chg Start 2011/03/10 SMPK Miyata
'        JudgSW.sp = (.typ_si.HWFSPVHS = "H") And CheckKHN(.typ_si.HWFSPVKN, 15, sSxlPos)
        JudgSW.sp = (.typ_si.HWFSPVHS = "H") And CheckKHN(.typ_si.HWFSPVKN, 15, sSxlPos) _
                    And (tt = SxlTop Or tt = SxlTail)
'Chg End   2011/03/10 SMPK Miyata
        
        If JudgSW.sp Then

            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).pos = -1                                               ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("SPFE", typTmList())             ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                                       ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                                       ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                             ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                             ' ���S
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
            '5�Ԗڂ̏��FAN���x��ǉ�
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' ���5
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' ���6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' ���7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' ���8
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                            ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                                     ' �i��(12��)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' �T���v���m��
                'SP����擾
                If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then
                    'SP���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                    'SP����擾
                    If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 1, sSxlPos) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_FE)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' ���P
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_FE)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' ���Q
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_FE)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")     ' ���R
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_FE)
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' ���S
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                    '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_FE)
                        typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' ���6
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_FE)
                        typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' ���7
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).STD_FE)
                        typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' ���8
                    '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00145"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        Else
            'If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 And .typ_Param.WFSMP(tt).WFRESSPCW = "1" Then
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                
                ''���т����鎞�̂ݕ\������
                '���ύX SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
'                If typ_J016_WFSPVJudg(tt).MAX_FE <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).MIN_FE <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).AVE_FE <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).CENTER_FE <> -1 Then
                If typ_J016_WFSPVJudg(tt).MAX_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).MIN_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).AVE_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).CENTER_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUA_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUAP_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).STD_FE <> -1 Then
                '���ύX SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                    
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                    .typ_rslt(tt, i).NAIYO = Search_CrCode("SPFE", typTmList())           ' ���e
                    .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
                    .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                    .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
                '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                    '5�Ԗڂ̏��FAN���x��ǉ�
                    typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                    typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                    typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                    typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                    typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                    typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                    typ_rslt_ex(tt, i).INFO5 = ""                                       ' ���5
                '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                    typ_rslt_ex(tt, i).INFO6 = ""                                       ' ���6
                    typ_rslt_ex(tt, i).INFO7 = ""                                       ' ���7
                    typ_rslt_ex(tt, i).INFO8 = ""                                       ' ���8
                '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                    .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' �T���v���m��
                    .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
                    .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                    'SP����擾
                    If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then      '2003/12/19 SystemBrain
                        'SP���莸�s
                        .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                        'SP����擾
                        If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 1, sSxlPos) Then
                            '��ʕ\�����e�ݒ�
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_FE)
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' ���P
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_FE)
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' ���Q
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_FE)
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")     ' ���R
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_FE)
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' ���S
                        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                        '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                            '5�Ԗڂ̏��FAN���x��ǉ�
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                            '3�`6���ڂ�AN���x
                            vTemp = Mid(vTemp, 3, 4)
                            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                            typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                        '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_FE)
                            typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' ���6
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_FE)
                            typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' ���7
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).STD_FE)
                            typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' ���8
                        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                        End If
                    ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                        '��ʕ\�����e�ݒ�
                        bJudg = False
                        .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                        .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                    End If
                    i = i + 1
                End If
            End If
        End If

    ''�g�U��***************
'Chg Start 2011/03/10 SMPK Miyata
'        JudgSW.sp = (.typ_si.HWFDLHWS = "H") And CheckKHN(.typ_si.HWFDLKHN, 16, sSxlPos)
        JudgSW.sp = (.typ_si.HWFDLHWS = "H") And CheckKHN(.typ_si.HWFDLKHN, 16, sSxlPos) _
                    And (tt = SxlTop Or tt = SxlTail)
'Chg End   2011/03/10 SMPK Miyata

        If JudgSW.sp Then
            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).pos = -1                                               ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("SPKL", typTmList())             ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                                       ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                                       ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                             ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                             ' ���S
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
            '5�Ԗڂ̏��FAN���x��ǉ�
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' ���5
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' ���6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' ���7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' ���8
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                            ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                                     ' �i��(12��)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' �T���v���m��
                
                If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then      '2003/12/19 SystemBrain
                    'SP���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                    'SP����擾
                    If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 2, sSxlPos) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_DIFF)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")      ' ���P
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_DIFF)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")      ' ���Q
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_DIFF)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")      ' ���R
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_DIFF)
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.0")      ' ���S
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                    '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_DIFF)
                        typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' ���6
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_DIFF)
                        typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' ���7
''                        vTemp = CVar(typ_J016_WFSPVJudg(tt).SPV_Fe_STD)
''                        typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' ���8
                    '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00145"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        
        Else
            'If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 And .typ_Param.WFSMP(tt).WFRESSPCW = "1" Then
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                
                ''���т����鎞�̂݁A�\������
                '���ύX SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
'                If typ_J016_WFSPVJudg(tt).MAX_DIFF <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).MIN_DIFF <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).AVE_DIFF <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).CENTER_DIFF <> -1 Then
                If typ_J016_WFSPVJudg(tt).MAX_DIFF <> -1 _
                    Or typ_J016_WFSPVJudg(tt).MIN_DIFF <> -1 _
                    Or typ_J016_WFSPVJudg(tt).AVE_DIFF <> -1 _
                    Or typ_J016_WFSPVJudg(tt).CENTER_DIFF <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUA_DIFF <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUAP_DIFF <> -1 Then
                '���ύX SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                    .typ_rslt(tt, i).NAIYO = Search_CrCode("SPKL", typTmList())         ' ���e
                    .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
                    .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                    .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
                '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                    '5�Ԗڂ̏��FAN���x��ǉ�
                    typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                    typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                    typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                    typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                    typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                    typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                    typ_rslt_ex(tt, i).INFO5 = ""                                       ' ���5
                '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                    typ_rslt_ex(tt, i).INFO6 = ""                                       ' ���6
                    typ_rslt_ex(tt, i).INFO7 = ""                                       ' ���7
                    typ_rslt_ex(tt, i).INFO8 = ""                                       ' ���8
                '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                    .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' �T���v���m��
                    .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
                    .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                    'SP����擾
                    If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then      '2003/12/19 SystemBrain
                        'SP���莸�s
                        .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                        'SP����擾
                        If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 2, sSxlPos) Then
                            '��ʕ\�����e�ݒ�
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_DIFF)
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")      ' ���P
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_DIFF)
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")      ' ���Q
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_DIFF)
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")      ' ���R
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_DIFF)
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.0")      ' ���S
                        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                        '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                            '5�Ԗڂ̏��FAN���x��ǉ�
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                            '3�`6���ڂ�AN���x
                            vTemp = Mid(vTemp, 3, 4)
                            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                            typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                        '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_DIFF)
                            typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' ���6
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_DIFF)
                            typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' ���7
''                            vTemp = CVar(typ_J016_WFSPVJudg(tt).SPV_Fe_STD)
''                            typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' ���8
                        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                        End If
                    ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                        '��ʕ\�����e�ݒ�
                        bJudg = False
                        .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                        .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                    End If
                    i = i + 1
                End If
            End If
        End If
        
    ''Upd end  2005/06/21 (TCS)t.terauchi      SPV9�_�Ή�   Fe�Z�x�E�g�U���ɕ����ĕ\��


'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
'Nr�Z�x(OtherRecords)�̎��ѕ\���s�̒ǉ��ɂ��ύX

        ''Nr�Z�x***************
'Chg Start 2011/03/10 SMPK Miyata
'        JudgSW.sp = (.typ_si.HWFNRHS = "H") And CheckKHN(.typ_si.HWFNRKN, 19, sSxlPos)
        JudgSW.sp = (.typ_si.HWFNRHS = "H") And CheckKHN(.typ_si.HWFNRKN, 19, sSxlPos) _
                    And (tt = SxlTop Or tt = SxlTail)
'Chg End   2011/03/10 SMPK Miyata

        If JudgSW.sp Then
            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).pos = -1                                               ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("SPNR", typTmList())             ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                                       ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                                       ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                             ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                             ' ���S
        '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
            '5�Ԗڂ̏��FAN���x��ǉ�
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' ���5
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' ���6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' ���7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' ���8
            .typ_rslt(tt, i).SMPLID = -1                                            ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                            ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                                     ' �i��(12��)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' �T���v���m��
                
                If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then      '2003/12/19 SystemBrain
                    'SP���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                    'SP����擾
                    If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 3, sSxlPos) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_NR)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")      ' ���P
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_NR)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")      ' ���Q
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_NR)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")      ' ���R
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_NR)
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")      ' ���S
                    '2.1.3 AN���x ���є��f�`�F�b�N
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_NR)
                        typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' ���6
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_NR)
                        typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' ���7
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).STD_NR)
                        typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' ���8
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00145"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                
                ''���т����鎞�̂݁A�\������
                If typ_J016_WFSPVJudg(tt).MAX_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).MIN_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).AVE_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).CENTER_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUA_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUAP_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).STD_NR <> -1 Then
                
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                    .typ_rslt(tt, i).NAIYO = Search_CrCode("SPNR", typTmList())         ' ���e
                    .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
                    .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                    .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
                    '5�Ԗڂ̏��FAN���x
                    typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                    typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                    typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                    typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                    typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                    typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                    typ_rslt_ex(tt, i).INFO5 = ""                                       ' ���5
                    typ_rslt_ex(tt, i).INFO6 = ""                                       ' ���6
                    typ_rslt_ex(tt, i).INFO7 = ""                                       ' ���7
                    typ_rslt_ex(tt, i).INFO8 = ""                                       ' ���8
                    .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' �T���v���m��
                    .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
                    .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                    'SP����擾
                    If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then
                        'SP���莸�s
                        .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                        'SP����擾
                        If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 3, sSxlPos) Then
                            '��ʕ\�����e�ݒ�
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_NR)
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")      ' ���P
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_NR)
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")      ' ���Q
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_NR)
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")      ' ���R
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_NR)
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")      ' ���S
                            '5�Ԗڂ̏��FAN���x
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                            '3�`6���ڂ�AN���x
                            vTemp = Mid(vTemp, 3, 4)
                            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                            typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_NR)
                            typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' ���6
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_NR)
                            typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' ���7
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).STD_NR)
                            typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' ���8
                        End If
                    ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                        '��ʕ\�����e�ݒ�
                        bJudg = False
                        .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                        .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                    End If
                    i = i + 1
                End If
            End If
        End If
        
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------


        '' ���������w��(DOI1)*****************************************************************
        DOIDataSet 1, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(DOI2)*****************************************************************
        DOIDataSet 2, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(DOI3)*****************************************************************
        DOIDataSet 3, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        
        
        ''�c���_�f���є���/�\�������ǉ��@03/12/09 ooba START ====================================>

        '' ���������w��(AOI)******************************************************************
'        If JudgSW.AOI Then
        '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/08/10 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.AOI And CheckKHN(.typ_si.HWFZOKHN, 17, sSxlPos) Then
        '�ʏ픲���F�ۏؕ��@=�ۏ� ���� �����p�x�Q���`�F�b�N�L��̏ꍇ�A�d�l�L�Ƃ���
        '���Ԕ����F�ۏؕ��@=�ۏ� ���� ���Ԕ���=�ۏ؂̏ꍇ�A�d�l�L�Ƃ���
        If (JudgSW.AOI And CheckKHN(.typ_si.HWFZOKHN, 17, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.AOI And .typ_si.MSMPFLGWFAOI = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
           (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Chg End   2011/08/10 Y.Hitomi
            
            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).pos = -1                                               ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("AO", typTmList())               ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                                       ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                                       ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                             ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                             ' ���S
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
            '5�Ԗڂ̏��FAN���x��ǉ�
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' ���5
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' ���6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' ���7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' ���8
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                            ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                                     ' �i��(12��)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDAOICW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFAOI).SAMPLEID             ' �T���v���m��
'                If .typ_Param.WFSMP(tt).WFRESAOICW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESAOICW = "1") And (Trim(.typ_y013(tt, WFAOI).SAMPLEID) <> "0") Then
                    'AOI���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                             ' ���Q
                    'AOI����
                    If WfCrAoiJudg(.typ_si, .typ_y013(tt, WFAOI), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA4)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")     ' ���1
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA5)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")     ' ���2
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA6)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")     ' ���3
                        .typ_rslt(tt, i).INFO4 = ""                                ' ���4
                        JiltusekiUmu(tt, WFAOI) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(.typ_y013(tt, WFAOI).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESAOICW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                       ' �������J�n�ʒu
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00142"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1

        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDAOICW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("AO", typTmList())           ' ���e
                .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                '5�Ԗڂ̏��FAN���x��ǉ�
                typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                typ_rslt_ex(tt, i).INFO5 = ""                                       ' ���5
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                typ_rslt_ex(tt, i).INFO6 = ""                                       ' ���6
                typ_rslt_ex(tt, i).INFO7 = ""                                       ' ���7
                typ_rslt_ex(tt, i).INFO8 = ""                                       ' ���8
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFAOI).SAMPLEID             ' �T���v���m��
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
                .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                bJudg = False
'                If .typ_Param.WFSMP(tt).WFRESAOICW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESAOICW = "1") And (Trim(.typ_y013(tt, WFAOI).SAMPLEID) <> "0") Then
                    'AOI���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                             ' ���Q
                    'AOI����
                    If WfCrAoiJudg(.typ_si, .typ_y013(tt, WFAOI), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(tt, i).NAIYO = Search_CrCode("AO", typTmList())  ' ���e
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA4)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")     ' ���1
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA5)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")     ' ���2
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA6)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")     ' ���3
                        .typ_rslt(tt, i).INFO4 = ""                                ' ���4
                        JiltusekiUmu(tt, WFAOI) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(.typ_y013(tt, WFAOI).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESAOICW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
                'Add Start 2011/11/28 Y.Hitomi ���Ԕ����̏ꍇ�́A�Q�l�\������
                If sSxlPos = "MID" And JudgSW.AOI And bJudg = False Then
                    .typ_rslt(tt, i).OKNG = "�Q�l"
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
                i = i + 1
            End If
        End If
        ''�c���_�f���є���/�\�������ǉ��@03/12/09 ooba END ======================================>
        
        ''GD���є���/�\�������ǉ��@05/02/04 ooba START =========================================>
'Cng Start 2011/11/28 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.GD And CheckKHN(.typ_si.HWFGDKHN, 18, sSxlPos) Then
        '�ʏ픲���F�ۏؕ��@=�ۏ� ���� �����p�x�Q���`�F�b�N�L��̏ꍇ�A�d�l�L�Ƃ���
        '���Ԕ����F�ۏؕ��@=�ۏ� ���� ���Ԕ����t���O���ۏ؂̏ꍇ�A�d�l�L�Ƃ���
        If (JudgSW.GD And CheckKHN(.typ_si.HWFGDKHN, 18, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.GD And .typ_si.MSMPFLGWFGD = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
           (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Cng End   2011/11/28 Y.Hitomi
            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).pos = -1                                               ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())               ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l�L"                                       ' ���P
            .typ_rslt(tt, i).INFO2 = "������"                                       ' ���Q
            .typ_rslt(tt, i).INFO3 = ""                                             ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                             ' ���S
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
            '5�Ԗڂ̏��FAN���x��ǉ�
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' ���5
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' ���6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' ���7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' ���8
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "NG"                                            ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                                     ' �i��(12��)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDGDCW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                If .typ_Param.WFSMP(tt).WFHSGDCW = "1" Then
                    '��������
                    .typ_rslt(tt, i).SMPLID = Format(Trim(typ_J015_WFGDJudg(tt).SMPLNO), "0000") & "       �y�����z"   ' �T���v���m��
                Else
                    'WF����
                    .typ_rslt(tt, i).SMPLID = typ_J015_WFGDJudg(tt).SMPLNO          ' �T���v���m��
                End If
                If (.typ_Param.WFSMP(tt).WFRESGDCW = "1") And (Trim(typ_J015_WFGDJudg(tt).SMPLNO) <> "") Then
                    'GD���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                    
                    '�v�e����/�������ю���
                    .typ_si.WFHSGDCW = .typ_Param.WFSMP(tt).WFHSGDCW    '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
                    
                    'GD����
                    If WfCrGdJudg(.typ_si, typ_J015_WFGDJudg(tt), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSDEN)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' ���P
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSLDL)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")        ' ���Q
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSDVD2)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' ���R
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech Start
''                        .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSZEROMN)
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' ���S
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSZEROMX)
                        .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & " , " & DBData2DispData(vTemp, "0")        ' ���S
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech End
                        JiltusekiUmu(tt, WFGD) = True
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(typ_J015_WFGDJudg(tt).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESGDCW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                       ' �������J�n�ʒu
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                If pbGDJudgeTbl(3) = False Then
                    gsTbcmy028ErrCode = "00148"
                ElseIf pbGDJudgeTbl(3) = False Then
                    gsTbcmy028ErrCode = "00147"
                Else
                    gsTbcmy028ErrCode = "00146"
                End If
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
            
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDGDCW) <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
                .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())           ' ���e
                .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
                .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
                .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
                .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                '5�Ԗڂ̏��FAN���x��ǉ�
                typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                typ_rslt_ex(tt, i).INFO5 = ""                                       ' ���5
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                typ_rslt_ex(tt, i).INFO6 = ""                                       ' ���6
                typ_rslt_ex(tt, i).INFO7 = ""                                       ' ���7
                typ_rslt_ex(tt, i).INFO8 = ""                                       ' ���8
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                If .typ_Param.WFSMP(tt).WFHSGDCW = "1" Then
                    '��������
                    .typ_rslt(tt, i).SMPLID = Format(Trim(typ_J015_WFGDJudg(tt).SMPLNO), "0000") & "       �y�����z"   ' �T���v���m��
                Else
                    'WF����
                    .typ_rslt(tt, i).SMPLID = typ_J015_WFGDJudg(tt).SMPLNO          ' �T���v���m��
                End If
                .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
                .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
                bJudg = False
                If (.typ_Param.WFSMP(tt).WFRESGDCW = "1") And (Trim(typ_J015_WFGDJudg(tt).SMPLNO) <> "") Then
                    'GD���莸�s
                    .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                    
                    '�v�e����/�������ю���
                    .typ_si.WFHSGDCW = .typ_Param.WFSMP(tt).WFHSGDCW    '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
                    
                    'GD����
                    If WfCrGdJudg(.typ_si, typ_J015_WFGDJudg(tt), bJudg) Then
                        '��ʕ\�����e�ݒ�
                        .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())   ' ���e
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSDEN)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' ���P
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSLDL)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")        ' ���Q
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSDVD2)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' ���R
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech Start
''                        .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSZEROMN)
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' ���S
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSZEROMX)
                        .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & " , " & DBData2DispData(vTemp, "0")        ' ���S
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech End
                        JiltusekiUmu(tt, WFGD) = True
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(typ_J015_WFGDJudg(tt).DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESGDCW = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "����وُ�"                            ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                     ' ���S
                End If
                'Add Start 2011/11/28 Y.Hitomi ���Ԕ����̏ꍇ�́A�Q�l�\������
                If sSxlPos = "MID" And JudgSW.GD And bJudg = False Then
                    .typ_rslt(tt, i).OKNG = "�Q�l"
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
                i = i + 1
            End If
        End If
        ''GD���є���/�\�������ǉ��@05/02/04 ooba END ===========================================>

'Chg Start 2011/03/10 SMPK Miyata
'        JudgSW.sp = (.typ_si.HWFNRHS = "H") And CheckKHN(.typ_si.HWFNRKN, 19, sSxlPos)
        JudgSW.sp = (.typ_si.HWFNRHS = "H") And CheckKHN(.typ_si.HWFNRKN, 19, sSxlPos) _
                    And (tt = SxlTop Or tt = SxlTail)

'Chg End   2011/03/10 SMPK Miyata

        ''��Add 2010/01/07 SIRD�Ή� Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If InStr(IND, .typ_Param.WFSMP(tt).WFINDL4CW) <> 0 Then
        If InStr(IND, .typ_Param.WFSMP(tt).WFINDL4CW) <> 0 And (tt = SxlTop Or tt = SxlTail) Then
'Chg End   2011/03/10 SMPK Miyata
            '��ʕ\�����e�ݒ�
            .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' �������J�n�ʒu
            .typ_rslt(tt, i).NAIYO = Search_CrCode("SD", typTmList())           ' ���e
            .typ_rslt(tt, i).INFO1 = "�d�l��"                                   ' ���P
            .typ_rslt(tt, i).INFO2 = "�����L"                                   ' ���Q
            .typ_rslt(tt, i).INFO3 = "���і�"                                   ' ���R
            .typ_rslt(tt, i).INFO4 = ""                                         ' ���S
            
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                       ' ���5
            typ_rslt_ex(tt, i).INFO6 = ""                                       ' ���6
            typ_rslt_ex(tt, i).INFO7 = ""                                       ' ���7
            typ_rslt_ex(tt, i).INFO8 = ""                                       ' ���8
            .typ_rslt(tt, i).SMPLID = typ_J022_WFSDJudg(tt).SMPLNO             ' �T���v���m��
            .typ_rslt(tt, i).OKNG = "OK"                                        ' ���茋��
            .typ_rslt(tt, i).hinban = sHinban12                                 ' �i��(12��)
            bJudg = False
            'SIRD����擾
            If (.typ_Param.WFSMP(tt).WFRESL4CW = "1") And (Trim(typ_J022_WFSDJudg(tt).SMPLNO) <> "") Then
                'SIRD���莸�s
                .typ_rslt(tt, i).INFO3 = "����Err"                              ' ���Q
                If WfCrSdjudg(.typ_si, typ_J022_WFSDJudg(tt), bJudg) Then
                    
                    '��ʕ\�����e�ݒ�
                    vTemp = CVar(typ_J022_WFSDJudg(tt).SIRDCNT)
                    .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' ���P
                    .typ_rslt(tt, i).INFO2 = ""                                 ' ���Q
                    .typ_rslt(tt, i).INFO3 = ""                                 ' ���R
                    .typ_rslt(tt, i).INFO4 = ""                                 ' ���S
                    
                    JiltusekiUmu(tt, WFSIRD) = True
                    
                    vTemp = CVar(typ_J022_WFSDJudg(tt).DKAN)
                    vTemp = Mid(vTemp, 3, 4)
                    typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                    typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                    typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                    typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                    typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                    typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                    typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")      ' ���5
                    
                End If
            ElseIf .typ_Param.WFSMP(tt).WFRESL4CW = "2" Then
                '��ʕ\�����e�ݒ�
                bJudg = False
                .typ_rslt(tt, i).INFO3 = ""                                     ' ���R
                .typ_rslt(tt, i).INFO4 = "����وُ�"                            ' ���S
            End If
            
''Del Start 2012/03/14 Y.Hitomi
 ''Add Start 2011/03/10 SMPK Miyata
 '            '�ۏؕ��@���ۏ؂��H
 '            If JudgSW.L4 = True Then
 ''Add End   2011/03/10 SMPK Miyata
''Del End 2012/03/14 Y.Hitomi
                If bJudg = True Then
                    .typ_rslt(tt, i).OKNG = "OK"                                ' ���茋��
                Else
                    .typ_rslt(tt, i).OKNG = "NG"                                ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                    If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
                End If

'Del Start 2012/03/13 Y.Hitomi
'            End If          'Add 2011/03/10 SMPK Miyata
'Del End   2012/03/13 Y.Hitomi
            
            i = i + 1
        End If
    ''��Add 2010/01/07 SIRD�Ή� Y.Hitomi
        
    End With
    WfAllJudg = FUNCTION_RETURN_SUCCESS
End Function
'''''============================================================================================================================
'''''
''''''�T�v      :���i�V�[�g�\��
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :typ_a         ,I  ,typ_AllTypesC ,�e���\����
''''''����      :�V�[�g�ɒl��ݒ肷��
''''''����      :
'''''Public Sub PutSeihinTop()
'''''    Dim i As Integer, j As Integer      ' ٰ�� ����
'''''
'''''    With f_cmbc039_2
'''''        For i = 1 To 4
'''''            .spdHinbanTop.col = i
'''''            .spdHinbanTop.row = 1
'''''            Select Case i
'''''            Case 1
'''''                '�i��
'''''                .spdHinbanTop.Value = typ_CType.typ_Param.hinban
'''''            Case 2
'''''                '�^�C�v
'''''                .spdHinbanTop.Value = typ_CType.typ_si.HWFTYPE
'''''            Case 3
'''''                '����
'''''                .spdHinbanTop.Value = typ_CType.typ_si.HWFCDIR
'''''            Case 4
'''''                '�����h�[�v
'''''                .spdHinbanTop.Value = typ_CType.typ_si.HWFCDOP
'''''            End Select
'''''        Next i
'''''    End With
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''�T�v      :���i�V�[�g�\��
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :typ_a         ,I  ,typ_AllTypesC ,�e���\����
''''''����      :�V�[�g�ɒl��ݒ肷��
''''''����      :
'''''Public Sub PutSeihinCenter()
'''''    Dim i As Integer, j As Integer      ' ٰ�� ����
'''''
'''''    'CENTER��
'''''    With f_cmbc039_2
'''''        For i = 1 To 9
'''''            .spdHinbanCen.col = i
'''''            .spdHinbanCen.row = 1
'''''            Select Case i
'''''            Case 1
'''''                '���R
''''''2001/12/26 S.Sano                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFRMIN, "0.0000") & " - " & DBData2DispData(typ_CType.typ_si.HWFRMAX, "0.0000")
'''''                .spdHinbanCen.Value = toRsStr(typ_CType.typ_si.HWFRMIN) & " - " & toRsStr(typ_CType.typ_si.HWFRMAX) '2001/12/26 S.Sano
'''''            Case 2
'''''                'Oi
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFONMIN, "0.00") & " - " & DBData2DispData(typ_CType.typ_si.HWFONMAX, "0.00")
'''''            Case 3
'''''                'BMD1
'''''                '.spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM1AN, "0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM1AX, "0")
'''''                '�ׂ��搔�ύX�@2003/05/19 osawa
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM1AN, "0.0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM1AX, "0.0")
'''''            Case 4
'''''                'BMD2
'''''                '.spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM2AN, "0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM2AX, "0")
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM2AN, "0.0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM2AX, "0.0")
'''''            Case 5
'''''                'BMD3
'''''                '.spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM3AN, "0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM3AX, "0")
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM3AN, "0.0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM3AX, "0.0")
'''''            Case 6
'''''                'OSF1
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFOF1AX, "0.00") & " , " & DBData2DispData(typ_CType.typ_si.HWFOF1MX, "0.0")
'''''            Case 7
'''''                'OSF2
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFOF2AX, "0.00") & " , " & DBData2DispData(typ_CType.typ_si.HWFOF2MX, "0.0")
'''''            Case 8
'''''                'OSF3
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFOF3AX, "0.00") & " , " & DBData2DispData(typ_CType.typ_si.HWFOF3MX, "0.0")
'''''            Case 9
'''''                'OSF4
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFOF4AX, "0.00") & " , " & DBData2DispData(typ_CType.typ_si.HWFOF4MX, "0.0")
'''''            End Select
'''''        Next i
'''''    End With
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''�T�v      :���i�V�[�g�\��
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :typ_a         ,I  ,typ_AllTypesC ,�e���\����
''''''����      :�V�[�g�ɒl��ݒ肷��
''''''����      :
'''''Public Sub PutSeihinTail()
'''''    Dim i As Integer, j As Integer      ' ٰ�� ����
'''''
'''''    'TAIL��
'''''    With f_cmbc039_2
'''''        For i = 1 To 6
'''''            .spdHinbanTail.col = i
'''''            .spdHinbanTail.row = 1
'''''            Select Case i
'''''            Case 1
'''''                'DS
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFDSOMN, "0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFDSOMX, "0")
'''''            Case 2
'''''                'DZ
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFMKMIN, "0.0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFMKMAX, "0.0")
'''''            Case 3
'''''                'SP
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFSPVMX, "0.00") & " , " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFDLMIN, "0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFDLMAX, "0")
'''''            Case 4
'''''                'D1
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFOS1MN, "0.0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFOS1MX, "0.0")
'''''            Case 5
'''''                'D2
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFOS2MN, "0.0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFOS2MX, "0.0")
'''''            Case 6
'''''                'D3
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFOS3MN, "0.0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFOS3MX, "0.0")
'''''            End Select
'''''        Next i
'''''    End With
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''�T�v      :���R�l�\��
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :typ_a         ,I  ,typ_AllTypesC ,�e���\����
''''''����      :���R�̈�ɒl��\������
''''''����      :
'''''Public Sub PutRs()
'''''
'''''    '���R�l�\��(TOP��)
'''''    PutRsTop
'''''
'''''    '���R�l�\��(TAIL��)
'''''    PutRsTail
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''�T�v      :���R�l�\��(TOP)
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :typ_a         ,I  ,typ_AllTypesC ,�e���\����
''''''          :bJudg         ,I  ,Boolean      ,���茋��
''''''          :dblScut       ,I  ,Double       ,�ăJ�b�g�ʒu
''''''          :dblCoef       ,I  ,Double       ,���s�ΐ�
''''''����      :���R�̈�ɒl��\������
''''''����      :
'''''Public Sub PutRsTop()
'''''    Dim bJudg As Boolean
'''''    Dim dblScut As Double
'''''    Dim dblCoef As Double
'''''
'''''''2001/12/18 S.Sano    bJudg = typ_CType.bOKNG(SxlTop)
''''''2002/03/04 S.Sano    bJudg = (typ_CType.bOKNG(SxlTop) Or bPPlus Or bNPlus) ''2001/12/18 S.Sano
'''''    dblScut = typ_CType.dblScut(SxlTop)
'''''    dblCoef = typ_CType.COEF(SxlTop)
'''''
'''''    With f_cmbc039_2
'''''        '' WF�����w���iRs)*****************************************************************
'''''        If JudgSW.rs Then
'''''            If InStr("1345", typ_CType.typ_Param.WFSMP(SxlTop).WFINDRS) <> 0 Then
'''''
'''''                If typ_CType.typ_Param.WFSMP(SxlTop).WFRESRS = "1" Then
'''''                    .txtSXLTop.Text = DBData2DispData(typ_CType.typ_Param.IngotPos, "0")            '�ʒu
'''''                    .txtRRGTop.Text = DBData2DispData(typ_CType.typ_y013(SxlTop, WFRES).MESDATA6, "0.00")  'RRG
'''''
''''''2002/03/04 S.Sano                    If typ_CType.JudgRrg(1) = False Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''                    If Not (typ_CType.JudgRrg(SxlTop) Or bPPlus Or bNPlus) Then '2002/03/04 S.Sano
'''''                    If Not (typ_CType.JudgRrg(SxlTop)) Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''                        CtrlEnabled .txtRRGTop, CTRL_DISABLE_WARNING, False  'RRG
'''''                    End If
'''''                    If dblCoef = -1 Or dblCoef = -9999 Or Mid(typ_CType.typ_Param.CRYNUM, 1, 1) = "8" Then
'''''                        .txtJHAll.Text = ""         '���s�ΐ̓u���b�N
'''''                    Else
'''''                        .txtJHAll.Text = DBData2DispData(dblCoef, "0.000")         '���s�ΐ̓u���b�N
'''''                    End If
'''''
'''''                    If Mid(typ_CType.typ_Param.CRYNUM, 1, 1) <> "8" Then
'''''                        '�ăJ�b�g�ʒu
''''''2002/03/04 S.Sano                        If typ_CType.JudgRes(1) Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''                        If typ_CType.JudgRes(SxlTop) Or bPPlus Or bNPlus Then '2002/03/04 S.Sano
'''''                        If typ_CType.JudgRes(SxlTop) Then '2002/03/04 S.Sano
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''                            .txtCutPosTop.Text = "OK"
'''''                        Else
''''''2001/08/30 S.Sano                            If dblScut < 0 Then
''''''2001/08/30 S.Sano                                .txtCutPosTop.Text = "0"
''''''2001/08/30 S.Sano                            ElseIf dblScut <= typ_CType.typ_Param.INGOTPOS Then
''''''2001/08/30 S.Sano                                .txtCutPosTop.Text = "0"
''''''2001/08/30 S.Sano                            ElseIf dblScut >= typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.INGOTPOS Then
''''''2001/08/30 S.Sano                                .txtCutPosTop.Text = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.INGOTPOS
''''''2001/08/30 S.Sano                            Else
''''''2001/08/30 S.Sano                                .txtCutPosTop.Text = DBData2DispData(dblScut, "0")
''''''2001/08/30 S.Sano                            End If
''''''2001/08/30 S.Sano Start
'''''                            Select Case dblScut
'''''                            Case -9999
'''''                                .txtCutPosTop.Text = ""
'''''                            Case Is <= typ_CType.typ_Param.IngotPos
'''''                                .txtCutPosTop.Text = typ_CType.typ_Param.IngotPos
'''''                            Case Is >= typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                                .txtCutPosTop.Text = typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                            Case Else
'''''                                .txtCutPosTail.Text = DBData2DispData(dblScut, "0")
'''''                            End Select
''''''2001/08/30 S.Sano End
'''''                            CtrlEnabled .txtCutPosTop, CTRL_DISABLE_WARNING, False  'TOP�ăJ�b�g
'''''                            intEnCmd = 1
'''''                        End If
'''''                    Else
''''''2001/08/30 S.Sano                        .txtCutPosTop.Text = "OK"
''''''2001/08/30 S.Sano Start
''''''2002/03/04 S.Sano                        If typ_CType.JudgRes(1) Then
'''''
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''                        If typ_CType.JudgRes(SxlTop) Or bPPlus Or bNPlus Then '2002/03/04 S.Sano
'''''                        If typ_CType.JudgRes(SxlTop) Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''                            .txtCutPosTop.Text = "OK"
'''''                        Else
'''''                            Select Case dblScut
'''''                            Case -9999
'''''                                .txtCutPosTop.Text = ""
'''''                            Case Is <= typ_CType.typ_Param.IngotPos
'''''                                .txtCutPosTop.Text = typ_CType.typ_Param.IngotPos
'''''                            Case Is >= typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                                .txtCutPosTop.Text = typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                            End Select
'''''                            CtrlEnabled .txtCutPosTop, CTRL_DISABLE_WARNING, False  'TOP�ăJ�b�g
'''''                            intEnCmd = 1
'''''                        End If
''''''2001/08/30 S.Sano End
'''''                    End If
'''''
'''''                    '���R
'''''                    If UBound(typ_CType.typ_y013top) > 0 Then
'''''                        With .spdMeasTop
'''''                            .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA1)
'''''                            .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA2)
'''''                            .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA3)
'''''                            .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA4)
'''''                            .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA5)
'''''                        End With
'''''                        RsSpreadSet .spdMeasTop, 1 '2002/01/25 S.Sano
'''''                    Else
'''''                        With .spdMeasTop
'''''                            .col = 1
'''''                            .row = 1:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "�d�l�L"
'''''                            .row = 2:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "�����L"
'''''                            .row = 3:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "���і�"
'''''                        End With
'''''                    End If
'''''                End If
'''''            Else
'''''                .txtSXLTop.Text = ""            '�ʒu
'''''                .txtRRGTop.Text = ""            'RRG
'''''                .txtJHAll.Text = ""
'''''                '�ăJ�b�g�ʒu
'''''
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''                If bPPlus Or bNPlus Then
''''''                    .txtCutPosTop.Text = "OK"
''''''                Else
'''''                    .txtCutPosTop.Text = "NG"
'''''                    CtrlEnabled .txtCutPosTop, CTRL_DISABLE_WARNING, False  'TOP�ăJ�b�g
''''''                End If
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''                '���R
'''''                With .spdMeasTop
'''''                    .col = 1
'''''                    .row = 1:
'''''                            .CellType = CellTypeStaticText
'''''                            .Value = "�d�l�L"
'''''                    .row = 2:
'''''                            .CellType = CellTypeStaticText
'''''                            .Value = "������"
'''''                    .row = 3: .Value = ""
'''''                    .row = 4: .Value = ""
'''''                    .row = 5: .Value = ""
'''''                End With
'''''            End If
'''''        Else
'''''            If InStr("1345", typ_CType.typ_Param.WFSMP(SxlTop).WFINDRS) <> 0 Then
'''''
'''''                If typ_CType.typ_Param.WFSMP(SxlTop).WFRESRS = "1" Then
'''''                    .txtSXLTop.Text = DBData2DispData(typ_CType.typ_Param.IngotPos, "0")            '�ʒu
'''''                    .txtRRGTop.Text = DBData2DispData(typ_CType.typ_y013(SxlTop, WFRES).MESDATA6, "0.00")  'RRG
'''''                    If dblCoef = -1 Or dblCoef = -9999 Then
'''''                        .txtJHAll.Text = ""         '���s�ΐ̓u���b�N
'''''                    Else
'''''                        .txtJHAll.Text = DBData2DispData(dblCoef, "0.000")         '���s�ΐ̓u���b�N
'''''                    End If
'''''
'''''                    '�ăJ�b�g�ʒu
'''''                    .txtCutPosTop.Text = "OK"
'''''
'''''                    '���R
'''''                    If UBound(typ_CType.typ_y013top) > 0 Then
'''''                        With .spdMeasTop
'''''                            .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA1)
'''''                            .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA2)
'''''                            .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA3)
'''''                            .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA4)
'''''                            .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA5)
'''''                        End With
'''''                        RsSpreadSet .spdMeasTop, 1 '2002/01/25 S.Sano
'''''                    Else
'''''                        With .spdMeasTop
'''''                            .col = 1
'''''                            .row = 1:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "�d�l��"
'''''                            .row = 2:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "�����L"
'''''                            .row = 3:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "���і�"
'''''                        End With
'''''                    End If
'''''                Else
'''''                    .txtSXLTop.Text = DBData2DispData(typ_CType.typ_Param.IngotPos, "0")            '�ʒu
'''''                    .txtRRGTop.Text = DBData2DispData(typ_CType.typ_y013(SxlTop, WFRES).MESDATA6, "0.00")  'RRG
'''''                    .txtJHAll.Text = ""
'''''                    '�ăJ�b�g�ʒu
'''''                    .txtCutPosTop.Text = "OK"
'''''                    '���R
'''''                    With .spdMeasTop
'''''                        .col = 1
'''''                        .row = 1:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "�d�l��"
'''''                        .row = 2:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "�����L"
'''''                        .row = 3:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "���і�"
'''''                        .row = 4: .Value = ""
'''''                        .row = 5: .Value = ""
'''''                    End With
'''''                End If
'''''            End If
'''''        End If
'''''    End With
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''�T�v      :���R�l�\��(TAIL)
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :typ_a         ,I  ,typ_AllTypesC ,�e���\����
''''''          :bJudg         ,I  ,Boolean      ,���茋��
''''''          :dblScut       ,I  ,Double       ,�ăJ�b�g�ʒu
''''''����      :���R�̈�ɒl��\������
''''''����      :
'''''Public Sub PutRsTail()
'''''    Dim bJudg As Boolean
'''''    Dim dblScut As Double
'''''
'''''''2001/12/18 S.Sano    bJudg = typ_CType.bOKNG(SxlTail)
''''''2002/03/04 S.Sano    bJudg = (typ_CType.bOKNG(SxlTail) Or bPPlus Or bNPlus) ''2001/12/18 S.Sano
'''''    dblScut = typ_CType.dblScut(SxlTail)
'''''
'''''
'''''    With f_cmbc039_2
'''''        '' WF�����w���iRs)*****************************************************************
'''''        If JudgSW.rs Then
'''''            If InStr("2345", typ_CType.typ_Param.WFSMP(SxlTail).WFINDRS) <> 0 Then
'''''
'''''                If typ_CType.typ_Param.WFSMP(SxlTail).WFRESRS = "1" Then
'''''                    .txtSXLTail.Text = DBData2DispData(typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH, "0")           '�ʒu
'''''                    .txtRRGTail.Text = DBData2DispData(typ_CType.typ_y013(SxlTail, WFRES).MESDATA6, "0.00")  'RRG
''''''2002/03/04 S.Sano                    If typ_CType.JudgRrg(1) = False Then
'''''
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''                    If Not (typ_CType.JudgRrg(SxlTail) Or bPPlus Or bNPlus) Then
'''''                    If Not (typ_CType.JudgRrg(SxlTail)) Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''                        CtrlEnabled .txtRRGTail, CTRL_DISABLE_WARNING, False  'RRG
'''''                    End If
'''''
'''''                    '�ăJ�b�g�ʒu
'''''                    If Mid(typ_CType.typ_Param.CRYNUM, 1, 1) <> "8" Then
''''''2002/03/04 S.Sano                        If typ_CType.JudgRes(2) = True Then
'''''
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''                        If typ_CType.JudgRes(SxlTail) Or bPPlus Or bNPlus Then '2002/03/04 S.Sano
'''''                        If typ_CType.JudgRes(SxlTail) Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''                            .txtCutPosTail.Text = "OK"
'''''                        Else
''''''2001/08/30 S.Sano                            If dblScut < 0 Then
''''''2001/08/30 S.Sano                                .txtCutPostail.Text = "0"
''''''2001/08/30 S.Sano                            ElseIf dblScut <= typ_CType.typ_Param.INGOTPOS Then
''''''2001/08/30 S.Sano                                .txtCutPostail.Text = "0"
''''''2001/08/30 S.Sano                            ElseIf dblScut >= typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.INGOTPOS Then
''''''2001/08/30 S.Sano                                .txtCutPostail.Text = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.INGOTPOS
''''''2001/08/30 S.Sano                            Else
''''''2001/08/30 S.Sano                                .txtCutPostail.Text = DBData2DispData(dblScut, "0")
''''''2001/08/30 S.Sano                            End If
''''''2001/08/30 S.Sano Start
'''''                            Select Case dblScut
'''''                            Case -9999
'''''                                .txtCutPosTail.Text = ""
'''''                            Case Is <= typ_CType.typ_Param.IngotPos
'''''                                .txtCutPosTail.Text = typ_CType.typ_Param.IngotPos
'''''                            Case Is >= typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                                .txtCutPosTail.Text = typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                            Case Else
'''''                                .txtCutPosTail.Text = DBData2DispData(dblScut, "0")
'''''                            End Select
''''''2001/08/30 S.Sano End
'''''                            CtrlEnabled .txtCutPosTail, CTRL_DISABLE_WARNING, False  'tail�ăJ�b�g
'''''                            intEnCmd = 1
'''''                        End If
'''''                    Else
''''''2001/08/30 S.Sano                        .txtCutPostail.Text = "OK"
''''''2001/08/30 S.Sano Start
''''''2002/03/04 S.Sano                        If typ_CType.JudgRes(2) = True Then
'''''
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''                        If typ_CType.JudgRes(SxlTail) Or bPPlus Or bNPlus Then '2002/03/04 S.Sano
'''''                        If typ_CType.JudgRes(SxlTail) Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''                            .txtCutPosTail.Text = "OK"
'''''                        Else
'''''                            Select Case dblScut
'''''                            Case -9999
'''''                                .txtCutPosTail.Text = ""
'''''                            Case Is <= typ_CType.typ_Param.IngotPos
'''''                                .txtCutPosTail.Text = typ_CType.typ_Param.IngotPos
'''''                            Case Is >= typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                                .txtCutPosTail.Text = typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                            End Select
'''''                            CtrlEnabled .txtCutPosTail, CTRL_DISABLE_WARNING, False  'tail�ăJ�b�g
'''''                            intEnCmd = 1
'''''                        End If
''''''2001/08/30 S.Sano End
'''''                    End If
'''''
'''''                    '���R
'''''                    If UBound(typ_CType.typ_y013tail) > 0 Then
'''''                        With .spdMeasTail
'''''                            .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA1)
'''''                            .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA2)
'''''                            .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA3)
'''''                            .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA4)
'''''                            .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA5)
'''''                        End With
'''''                        RsSpreadSet .spdMeasTail, 1 '2002/01/25 S.Sano
'''''                    Else
'''''                        With .spdMeasTail
'''''                            .col = 1
'''''                            .row = 1:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "�d�l�L"
'''''                            .row = 2:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "�����L"
'''''                            .row = 3:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "���і�"
'''''                        End With
'''''                    End If
'''''                End If
'''''            Else
'''''                .txtSXLTail.Text = ""            '�ʒu
'''''                .txtRRGTail.Text = ""            'RRG
'''''                .txtJHAll.Text = ""
'''''                '�ăJ�b�g�ʒu
'''''
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''                If bPPlus Or bNPlus Then
''''''                    .txtCutPosTail.Text = "OK"
''''''                Else
'''''                    .txtCutPosTail.Text = "NG"
'''''                    CtrlEnabled .txtCutPosTail, CTRL_DISABLE_WARNING, False  'Tail�ăJ�b�g
''''''                End If
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''                '���R
'''''                With .spdMeasTail
'''''                    .col = 1
'''''                    .row = 1:
'''''                            .CellType = CellTypeStaticText
'''''                            .Value = "�d�l�L"
'''''                    .row = 2:
'''''                            .CellType = CellTypeStaticText
'''''                            .Value = "������"
'''''                    .row = 3: .Value = ""
'''''                    .row = 4: .Value = ""
'''''                    .row = 5: .Value = ""
'''''                End With
'''''            End If
'''''        Else
'''''            If InStr("2345", typ_CType.typ_Param.WFSMP(SxlTail).WFINDRS) <> 0 Then
'''''
'''''                If typ_CType.typ_Param.WFSMP(SxlTail).WFRESRS = "1" Then
'''''                    .txtSXLTail.Text = DBData2DispData(typ_CType.typ_Param.IngotPos, "0")            '�ʒu
'''''                    .txtRRGTail.Text = DBData2DispData(typ_CType.typ_y013(SxlTail, WFRES).MESDATA6, "0.00")  'RRG
'''''
'''''                    '�ăJ�b�g�ʒu
'''''                    .txtCutPosTail.Text = "OK"
'''''
'''''                    '���R
'''''                    If UBound(typ_CType.typ_y013tail) > 0 Then
'''''                        With .spdMeasTail
'''''                            .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA1)
'''''                            .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA2)
'''''                            .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA3)
'''''                            .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA4)
'''''                            .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA5)
'''''                        End With
'''''                        RsSpreadSet .spdMeasTail, 1 '2002/01/25 S.Sano
'''''                    Else
'''''                        With .spdMeasTail
'''''                            .col = 1
'''''                            .row = 1:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "�d�l��"
'''''                            .row = 2:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "�����L"
'''''                            .row = 3:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "���і�"
'''''                        End With
'''''                    End If
'''''                Else
'''''                    .txtSXLTail.Text = DBData2DispData(typ_CType.typ_Param.IngotPos, "0")            '�ʒu
'''''                    .txtRRGTail.Text = DBData2DispData(typ_CType.typ_y013(SxlTail, WFRES).MESDATA6, "0.00")  'RRG
'''''                    .txtJHAll.Text = ""
'''''                    '�ăJ�b�g�ʒu
'''''                    .txtCutPosTop.Text = "OK"
'''''                    '���R
'''''                    With .spdMeasTop
'''''                        .col = 1
'''''                        .row = 1:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "�d�l��"
'''''                        .row = 2:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "�����L"
'''''                        .row = 3:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "���і�"
'''''                        .row = 4: .Value = ""
'''''                        .row = 5: .Value = ""
'''''                    End With
'''''                End If
'''''            End If
'''''        End If
'''''    End With
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''�T�v      :���ђl�\��(TOP)
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :typ_rslt()    ,I  ,typ_ALLRSLT  ,���я��\����
''''''          :tt            ,I  ,Integer      ,TopTail����p
''''''����      :���ї̈�ɒl��\������
''''''����      :
'''''Public Sub PutRslt(typ_rslt() As typ_ALLRSLT, tt As Integer)
'''''    Dim i, j As Integer
'''''    Dim va As vaSpread
'''''    Dim spMaxLine As Long
'''''
'''''    ''�ő�s���擾
'''''    spMaxLine = 0
'''''    Do While typ_rslt(tt, spMaxLine).OKNG <> ""
'''''        spMaxLine = spMaxLine + 1
'''''    Loop
'''''
'''''    If tt = SxlTop Then
'''''        Set va = f_cmbc039_2.spdKensaTop
'''''    Else
'''''        Set va = f_cmbc039_2.spdKensaTail
'''''    End If
'''''
'''''    SpCtrlInit va, spMaxLine
'''''    SpCtrlBlockEnabled va, 1, 1, spMaxLine, 5, CTRL_DISABLE
'''''
'''''
'''''    i = 1
'''''    Do While typ_rslt(tt, i - 1).OKNG <> ""
'''''        With typ_rslt(tt, i - 1)
'''''            va.row = i
'''''            For j = 1 To 8
'''''                va.col = j
'''''                Select Case j
'''''                Case 1
'''''                    '�ʒu
'''''                    va.Value = DBData2DispData(CVar(.pos), "0")
'''''                Case 2
'''''                    '���e
'''''                    va.Value = .NAIYO
'''''                Case 3
'''''                    '���P
'''''                    va.Value = .INFO1
'''''                Case 4
'''''                    '���Q
'''''                    va.Value = .INFO2
'''''                Case 5
'''''                    '���R
'''''                    va.Value = .INFO3
'''''                Case 6
'''''                    '���S
'''''                    va.Value = .INFO4
'''''                Case 7
'''''                    '����
'''''''2001/12/18 S.Sano                    If .OKNG = "NG" Then
'''''''2001/12/18 S.Sano                        SpCtrlEnabled va, va.Col, va.row, CTRL_DISABLE_WARNING
'''''''2001/12/18 S.Sano                        'va.BackColor = &H8080FF
'''''''2001/12/18 S.Sano                        intEnCmd = 1
'''''''2001/12/18 S.Sano                    End If
'''''''2001/12/18 S.Sano                    va.Value = .OKNG
'''''''2001/12/18 S.Sano Start
'''''
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''                    If bPPlus Or bNPlus Then
''''''                        va.Value = "OK"
''''''                    Else
'''''                        If .OKNG = "NG" Then
'''''                            SpCtrlEnabled va, va.col, va.row, CTRL_DISABLE_WARNING
'''''                            'va.BackColor = &H8080FF
'''''                            intEnCmd = 1
'''''                        End If
'''''                        va.Value = .OKNG
''''''                    End If
'''''''2001/12/18 S.Sano End
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''                Case 8
'''''                    '�ʒu
'''''                    va.Value = CStr(DBData2DispData(.SMPLID, "0"))
'''''                End Select
'''''            Next j
'''''        End With
'''''        i = i + 1
'''''    Loop
'''''
'''''    '�\�[�g����
'''''    If i <> 1 Then
'''''        With va
'''''            .MaxRows = i - 1                      '�@�i�ԁi�s���j
'''''            .row = 1                            ' �Z���u���b�N��ݒ�
'''''            .col = 1
'''''            .row2 = i - 1
'''''            .col2 = 8
'''''            .SortBy = SS_SORT_BY_ROW
'''''            .SortKey(1) = 7                    ' ��P�\�[�g�L�[��ݒ�
'''''            ' �����ɕ��בւ�
'''''            .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
'''''            .Action = SS_ACTION_SORT
'''''        End With
'''''    End If
'''''
'''''End Sub

'�T�v      :��R����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_CType     ,I  ,typ_AllTypesC                        :�S���\����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :dblScut       ,O  ,Double                               :�ăJ�b�g�ʒu
'          :tt            ,I  ,Integer                              :Top,Tail����p
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :��R������s���A���肪NG�������ꍇ�͍ăJ�b�g�ʒu��Ԃ�
'����      :
Public Function WfCrResJudg(typ_CType As typ_AllTypesC, _
                            typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            bJudg As Boolean, _
                            dblScut As Double, _
                            tt As Integer) As Boolean

    Dim ErrInfo     As ERROR_INFOMATION
    Dim rs          As W_RES
    Dim cc          As type_Coefficient
    Dim rp          As type_ResPosCal
    Dim COEF        As Double
    Dim wgtCharge   As Long         '�ΐ͌v�Z�p�p�����[�^
    Dim wgtTop      As Double       '�ΐ͌v�Z�p�p�����[�^
    Dim wgtTopCut   As Double       '�ΐ͌v�Z�p�p�����[�^
    Dim DM          As Double       '�ΐ͌v�Z�p�p�����[�^
    
    '��R��������ݒ�
    rs.GuaranteeRes.cMeth = typ_si.HWFRSPOH         ' �i�v�e���R����ʒu�Q��
    rs.GuaranteeRes.cCount = typ_si.HWFRSPOT        ' �i�v�e���R����ʒu�Q�_
    rs.GuaranteeRes.cPos = typ_si.HWFRSPOI          ' �i�v�e���R����ʒu�Q��
    rs.GuaranteeRes.cObj = typ_si.HWFRHWYT          ' �i�v�e���R�ۏؕ��@�Q��
    rs.GuaranteeRes.cJudg = typ_si.HWFRHWYS         ' �i�v�e���R�ۏؕ��@�Q��
    rs.GuaranteeCal = typ_si.HWFRMCAL               ' �i�v�e���R�ʓ��v�Z 2001/11/08 S.Sano
    rs.SpecResMin = typ_si.HWFRMIN                  ' �i�v�e���R����
    rs.SpecResMax = typ_si.HWFRMAX                  ' �i�v�e���R���
    rs.SpecResAveMin = typ_si.HWFRAMIN              ' �i�v�e���R���ω���
    rs.SpecResAveMax = typ_si.HWFRAMAX              ' �i�v�e���R���Ϗ��
    rs.SpecRrg = typ_si.HWFRMBNP                    ' �i�v�e���R�ʓ����z
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    '�`�F�b�N�pAN���x��ǉ�
    rs.Antnp = typ_si.HWFANTNP                      ' �i�v�e�`�m���x
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'--------------- 2008/08/25 INSERT START  By Systech --------------
    rs.DkTmpSiyo = typ_CType.DkTmpSiyo
    
'Cng Start 2011/09/26 Y.Hitomi
    If tt = SxlTop Or tt = SxlTail Then
        rs.DkTmpJsk = typ_CType.DkTmpJsk(tt)
    Else
        typ_CType.DkTmpJsk(tt) = GetWfDKTmpCode(False, typ_CType.typ_Param.WFSMP(1))
        rs.DkTmpJsk = typ_CType.DkTmpJsk(tt)
    End If
''Chg Start 2011/03/09 SMPK Miyata
'    If tt = SxlTop Or tt = SxlTail Then
'    rs.DkTmpJsk = typ_CType.DkTmpJsk(tt)
'    Else
'        '���Ԕ�����DK���x�����DK���x���тȂ��Ŕ���OK�ɂ���
'        rs.DkTmpJsk = ""
'    End If
''Chg End   2011/03/09 SMPK Miyata
'Cng End   2011/09/26 Y.Hitomi
'--------------- 2008/08/25 INSERT  END   By Systech --------------
    With typ_CType
        rs.Res(0) = NtoZ2(.typ_y013(tt, WFRES).MESDATA1)   ' ����l�P
        rs.Res(1) = NtoZ2(.typ_y013(tt, WFRES).MESDATA2)   ' ����l�Q
        rs.Res(2) = NtoZ2(.typ_y013(tt, WFRES).MESDATA3)   ' ����l�R
        rs.Res(3) = NtoZ2(.typ_y013(tt, WFRES).MESDATA4)   ' ����l�S
        rs.Res(4) = NtoZ2(.typ_y013(tt, WFRES).MESDATA5)   ' ����l�T
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
        '�`�F�b�N�pAN���x��ǉ�
        rs.ResAntnp = NtoZ2(Mid(.typ_y013(tt, WFRES).DKAN, 3, 4)) ' ����l�U
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    End With
    
    '��R����
    If WfRESJudg(rs, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrResJudg = False
        typ_CType.JudgRes(tt) = rs.JudgRes '2002/01/25 S.Sano
        typ_CType.JudgRrg(tt) = rs.JudgRrg '2002/01/25 S.Sano
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
        '�`�F�b�N�pAN���x��ǉ�
'Chg Start 2011/08/12 Y.Hitomi
        typ_CType.JudgAntnp(tt) = rs.JudgAntnp
'        If tt = SxlTop Or tt = SxlTail Then typ_CType.JudgAntnp(tt) = rs.JudgAntnp
'Chg End   2011/08/12 Y.Hitomi

    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'--------------- 2008/08/25 INSERT START  By Systech --------------
'Chg Start 2011/03/25 Y.Hitomi
        typ_CType.JudgDkTmp(tt) = rs.JudgDkTmp
'        If tt = SxlTop Or tt = SxlTail Then typ_CType.JudgDkTmp(tt) = rs.JudgDkTmp
'Chg End   2011/03/25 Y.Hitomi

'--------------- 2008/08/25 INSERT  END   By Systech --------------
        Exit Function
    End If
    typ_CType.JudgRes(tt) = rs.JudgRes '2002/01/25 S.Sano
    typ_CType.JudgRrg(tt) = rs.JudgRrg '2002/01/25 S.Sano
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    '�`�F�b�N�pAN���x��ǉ�
'Chg Start 2011/08/12 Y.Hitomi
    typ_CType.JudgAntnp(tt) = rs.JudgAntnp
'    If tt = SxlTop Or tt = SxlTail Then typ_CType.JudgAntnp(tt) = rs.JudgAntnp
'Chg End   2011/08/12 Y.Hitomi

'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'--------------- 2008/08/25 INSERT START  By Systech --------------
'Chg Start 2011/03/25 Y.Hitomi
    typ_CType.JudgDkTmp(tt) = rs.JudgDkTmp
'    If tt = SxlTop Or tt = SxlTail Then typ_CType.JudgDkTmp(tt) = rs.JudgDkTmp
'Chg End   2011/03/25 Y.Hitomi

'--------------- 2008/08/25 INSERT  END   By Systech --------------

'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
'    bJudg = (rs.JudgRes And rs.JudgRrg)  '2002/01/25 S.Sano
'--------------- 2008/08/25 UPDATE START  By Systech --------------
'    bJudg = (rs.JudgRes And rs.JudgRrg And rs.JudgAntnp)  '2002/01/25 S.Sano
    bJudg = (rs.JudgRes And rs.JudgRrg And rs.JudgAntnp And rs.JudgDkTmp)
'--------------- 2008/08/25 UPDATE  END   By Systech --------------
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    typ_CType.typ_y013(tt, WFRES).MESDATA6 = rs.RRG '2002/01/25 S.Sano
    If wiSmpGetFlg = 0 Then
        With typ_CType
            '���s�ΐ͗p�p�����[�^�擾 �}���`����Ή� �Q�Ɗ֐��ύX 2008/04/23 SETsw Nakada
            If GetCoeffParams_new(.typ_hage(tt).CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
'            If GetCoeffParams(.typ_hage(tt).CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
                Debug.Print "�ΐ͌v�Z�p�p�����[�^�̎擾�Ɏ��s����"
            End If
            
            '�ΐ͌W���v�Z
            cc.DUNMENSEKI = AreaOfCircle(DM)
            cc.TOPSMPLPOS = .typ_Param.INGOTPOS
            cc.BOTSMPLPOS = .typ_Param.INGOTPOS + .typ_Param.LENGTH
            cc.CHARGEWEIGHT = wgtCharge
            cc.TOPWEIGHT = wgtTop + wgtTopCut
            cc.TOPRES = .typ_y013(SxlTop, WFRES).MESDATA5
            cc.BOTRES = .typ_y013(SxlTail, WFRES).MESDATA5
            .COEF(tt) = CoefficientCalculation(cc)
    
            If rs.JudgRes <> True Then
                '�ΐ͌v�Z����ăJ�b�g�ʒu���v�Z
                rp.COEFFICIENT = .COEF(tt)
                rp.DUNMENSEKI = AreaOfCircle(DM)
                rp.CHARGEWEIGHT = wgtCharge
                rp.TOPWEIGHT = wgtTop + wgtTopCut
                rp.TOPSMPLPOS = IIf(tt = SxlTop, .typ_Param.INGOTPOS, .typ_Param.INGOTPOS + .typ_Param.LENGTH)
                rp.TOPRES = .typ_y013(tt, WFRES).MESDATA5
                rp.target = IIf(tt = SxlTop, .typ_si.HWFRMAX, .typ_si.HWFRMIN)
                dblScut = PosCalculation(rp)
            Else                                                                            '2002/01/25 S.Sano
                If tt = 1 Then                                                              '2002/01/25 S.Sano
                    dblScut = typ_CType.typ_Param.INGOTPOS                                  '2002/01/25 S.Sano
                Else                                                                        '2002/01/25 S.Sano
                    dblScut = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH     '2002/01/25 S.Sano
                End If                                                                      '2002/01/25 S.Sano
            End If
        End With
    End If
    
    WfCrResJudg = True
    
End Function

'�T�v      :OI����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_y013      ,I  ,typ_TBCMY013                         :OI���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :sSxlPos       ,I  ,String                               :�T���v�����("MID"�����Ԕ���)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :OI������s��
'����      :
'Cng Start 2011/08/01 Y.Hitomi
Public Function WfCrOiJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                           typ_y013 As typ_TBCMY013, _
                           bJudg As Boolean, _
                           Optional sSxlPos As String) As Boolean
'Public Function WfCrOiJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
'                           typ_y013 As typ_TBCMY013, _
'                           bJudg As Boolean) As Boolean
'Cng End 2011/08/01 Y.Hitomi
    
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim WOi     As W_OI                     'OI�\����
        
    bJudg = True
        
    'OI��������ݒ�
    WOi.GuaranteeOi.cMeth = typ_si.HWFONSPH   '�i�v�e�_�f�Z�x����ʒu�Q��
    WOi.GuaranteeOi.cCount = typ_si.HWFONSPT  '�i�v�e�_�f�Z�x����ʒu�Q�_
    WOi.GuaranteeOi.cPos = typ_si.HWFONSPI    '�i�v�e�_�f�Z�x����ʒu�Q��
    WOi.GuaranteeOi.cObj = typ_si.HWFONHWT    '�i�v�e�_�f�Z�x�ۏؕ��@�Q��
    WOi.GuaranteeOi.cJudg = typ_si.HWFONHWS   '�i�v�e�_�f�Z�x�ۏؕ��@�Q��
    WOi.GuaranteeCal = typ_si.HWFONMCL        '�i�v�e�_�f�Z�x�ʓ��v�Z 2001/11/08 S.Sano
    WOi.SpecOiMin = typ_si.HWFONMIN           '�iWF�_�f�Z�x����
    WOi.SpecOiMax = typ_si.HWFONMAX           '�iWF�_�f�Z�x���
    WOi.SpecORG = typ_si.HWFONMBP             '�iWF�_�f�Z�x�ʓ����z
    WOi.SpecOiAveMin = typ_si.HWFONAMN        '�iWF�_�f�Z�x���ω���
    WOi.SpecOiAveMax = typ_si.HWFONAMX        '�iWF�_�f�Z�x���Ϗ��
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    WOi.Antnp = typ_si.HWFANTNP               '�i�v�e�`�m���x
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    WOi.Oi(0) = NtoZ2(typ_y013.MESDATA1)               'Oi����l
    WOi.Oi(1) = NtoZ2(typ_y013.MESDATA2)               'Oi����l
    WOi.Oi(2) = NtoZ2(typ_y013.MESDATA3)               'Oi����l
    WOi.Oi(3) = NtoZ2(typ_y013.MESDATA4)               'Oi����l
    WOi.Oi(4) = NtoZ2(typ_y013.MESDATA5)               'Oi����l
    WOi.Oi(5) = NtoZ2(typ_y013.MESDATA6)               'Oi����l
    WOi.Oi(6) = NtoZ2(typ_y013.MESDATA7)               'Oi����l
    WOi.Oi(7) = NtoZ2(typ_y013.MESDATA8)               'Oi����l
    WOi.Oi(8) = NtoZ2(typ_y013.MESDATA9)               'Oi����l
    WOi.Oi(9) = NtoZ2(typ_y013.MESDATA10)              'Oi����l
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
    '�`�F�b�N�pAN���x��ǉ�
    WOi.OiAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))      'Oi����l
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    'OI����
    If WfOiJudg(WOi, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrOiJudg = False
        Exit Function
    End If
    
    typ_y013.MESDATA11 = WOi.ORG '2002/01/25 S.Sano
    typ_y013.MESDATA12 = WOi.OiMin '2002/01/25 S.Sano
    typ_y013.MESDATA13 = WOi.OiMax '2002/01/25 S.Sano
    
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
'Cng Start 2011/08/12 Y.Hitomi
'    If sSxlPos = "MID" Then
'        bJudg = (WOi.JudgOi And WOi.JudgOrg)
'    Else
'        bJudg = (WOi.JudgOi And WOi.JudgOrg And WOi.JudgAntnp)
'    End If
'    bJudg = (WOi.JudgOi And WOi.JudgOrg) '2002/01/25 S.Sano
    bJudg = (WOi.JudgOi And WOi.JudgOrg And WOi.JudgAntnp) '2002/01/25 S.Sano
'Cng Start 2011/08/12 Y.Hitomi
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    WfCrOiJudg = True
End Function

'�T�v      :BMD����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_y013      ,I  ,typ_TBCMY013                         :BMD���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :bmflg         ,I  ,Integer                              :BMD�׸�(1:BMD1, 2:BMD2, 3:BMD3)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :BMD������s��
'����      :
Public Function WfCrBmdJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y013 As typ_TBCMY013, _
                            bJudg As Boolean, _
                            bmflg As Integer) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim bm      As W_BMD                    'BMD�\����
    Dim c0      As Integer
    
    'Const keisu As Double = 10
    '' 2006/09/25 SMP)kondoh Del -s-
''    Const keisu As Double = 1        'BMD�ׂ��搔�ύX�Ή��@2003/05/19 osawa
    '' 2006/09/25 SMP)kondoh Del -e-
    '' 2006/09/25 SMP)kondoh Add -s-
    Dim keisu As Double
    Const keisu1 As Double = 10000
    Const keisu2 As Double = 10000
    Const keisu3 As Double = 10000
    Const keisu4 As Double = 10000
    Const keisu5 As Double = 10000
    Const keisu6 As Double = 333000
    Const keisu7 As Double = 10000
    Const keisu8 As Double = 10000
    Const keisu9 As Double = 10000 'Add 2012/07/20 Y.Hitomi

    '' 2006/09/25 SMP)kondoh Add -e-

    bJudg = True

    'BMD��������ݒ�
    Select Case bmflg
    Case 1
        bm.GuaranteeBmd.cMeth = typ_si.HWFBM1SH   '�i�v�e�a�l�c�P����ʒu�Q��
        bm.GuaranteeBmd.cCount = typ_si.HWFBM1ST  '�i�v�e�a�l�c�P����ʒu�Q�_
        bm.GuaranteeBmd.cPos = typ_si.HWFBM1SR    '�i�v�e�a�l�c�P����ʒu�Q��
        bm.GuaranteeBmd.cObj = typ_si.HWFBM1HT    '�i�v�e�a�l�c�P�ۏؕ��@�Q��
        bm.GuaranteeBmd.cJudg = typ_si.HWFBM1HS   '�i�v�e�a�l�c�P�ۏؕ��@�Q��
        bm.SpecBmdAveMin = typ_si.HWFBM1AN        '�i�v�e�a�l�c�P���ω���
        bm.SpecBmdAveMax = typ_si.HWFBM1AX        '�i�v�e�a�l�c�P���Ϗ��
        bm.SpecBmdMBP = typ_si.HWFBM1MBP          '�i�v�e�a�l�c�P�ʓ����z�@2003/05/20 ooba
        bm.SpecBmdMCL = NtoS(typ_si.HWFBM1MCL)    '�i�v�e�a�l�c�P�ʓ��v�Z�@2003/05/20 ooba
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        bm.Antnp = typ_si.HWFANTNP                '�i�v�e�`�m���x
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    Case 2
        bm.GuaranteeBmd.cMeth = typ_si.HWFBM2SH   '�i�v�e�a�l�c�Q����ʒu�Q��
        bm.GuaranteeBmd.cCount = typ_si.HWFBM2ST  '�i�v�e�a�l�c�Q����ʒu�Q�_
        bm.GuaranteeBmd.cPos = typ_si.HWFBM2SR    '�i�v�e�a�l�c�Q����ʒu�Q��
        bm.GuaranteeBmd.cObj = typ_si.HWFBM2HT    '�i�v�e�a�l�c�Q�ۏؕ��@�Q��
        bm.GuaranteeBmd.cJudg = typ_si.HWFBM2HS   '�i�v�e�a�l�c�Q�ۏؕ��@�Q��
        bm.SpecBmdAveMin = typ_si.HWFBM2AN        '�i�v�e�a�l�c�Q���ω���
        bm.SpecBmdAveMax = typ_si.HWFBM2AX        '�i�v�e�a�l�c�Q���Ϗ��
        bm.SpecBmdMBP = typ_si.HWFBM2MBP          '�i�v�e�a�l�c�Q�ʓ����z�@2003/05/20 ooba
        bm.SpecBmdMCL = NtoS(typ_si.HWFBM2MCL)    '�i�v�e�a�l�c�Q�ʓ��v�Z�@2003/05/20 ooba
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        bm.Antnp = typ_si.HWFANTNP                '�i�v�e�`�m���x
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    Case 3
        bm.GuaranteeBmd.cMeth = typ_si.HWFBM3SH   '�i�v�e�a�l�c�R����ʒu�Q��
        bm.GuaranteeBmd.cCount = typ_si.HWFBM3ST  '�i�v�e�a�l�c�R����ʒu�Q�_
        bm.GuaranteeBmd.cPos = typ_si.HWFBM3SR    '�i�v�e�a�l�c�R����ʒu�Q��
        bm.GuaranteeBmd.cObj = typ_si.HWFBM3HT    '�i�v�e�a�l�c�R�ۏؕ��@�Q��
        bm.GuaranteeBmd.cJudg = typ_si.HWFBM3HS   '�i�v�e�a�l�c�R�ۏؕ��@�Q��
        bm.SpecBmdAveMin = typ_si.HWFBM3AN        '�i�v�e�a�l�c�R���ω���
        bm.SpecBmdAveMax = typ_si.HWFBM3AX        '�i�v�e�a�l�c�R���Ϗ��
        bm.SpecBmdMBP = typ_si.HWFBM3MBP          '�i�v�e�a�l�c�R�ʓ����z�@2003/05/20 ooba
        bm.SpecBmdMCL = NtoS(typ_si.HWFBM3MCL)    '�i�v�e�a�l�c�R�ʓ��v�Z�@2003/05/20 ooba
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        bm.Antnp = typ_si.HWFANTNP                '�i�v�e�`�m���x
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    End Select

    '' 2006/09/25 SMP)kondoh Add -s-
    If bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "H" Then
        keisu = keisu1
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "H" Then
        keisu = keisu2
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu3
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu4
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "A" Then
        keisu = keisu5
    ElseIf bm.GuaranteeBmd.cMeth = "G" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu6
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu7
    ElseIf bm.GuaranteeBmd.cMeth = "8" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu8
    'Add Start 2012/07/20 Y.Hitomi
    ElseIf bm.GuaranteeBmd.cMeth = "P" Then
        keisu = keisu9
    'Add End 2012/07/20 Y.Hitomi
    
    Else
        bJudg = False
        WfCrBmdJudg = False
        Exit Function
    End If
    '' 2006/09/25 SMP)kondoh Add -e-

    With bm
        .BMD(0) = NtoZ2(typ_y013.MESDATA1)                   'BMD����l
        .BMD(1) = NtoZ2(typ_y013.MESDATA2)                   'BMD����l
        .BMD(2) = NtoZ2(typ_y013.MESDATA3)                   'BMD����l
        .BMD(3) = NtoZ2(typ_y013.MESDATA4)                   'BMD����l
        .BMD(4) = NtoZ2(typ_y013.MESDATA5)                   'BMD����l�@2003/05/20 ooba
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '�`�F�b�N�pAN���x��ǉ�
        .BmdAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------

        For c0 = 0 To 4                                      ' 2003/05/20 ooba
    '' 2006/09/25 SMP)kondoh Cng -s-
''            .BMD(c0) = IIf(.BMD(c0) <> -1, .BMD(c0) * keisu, -1)
            .BMD(c0) = IIf(.BMD(c0) <> -1, .BMD(c0) * CDbl(keisu / 10000), -1)
    '' 2006/09/25 SMP)kondoh Cng -e-
        Next
    End With
    
    'BMD����
    If WfBMDJudg(bm, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrBmdJudg = False
        Exit Function
    End If
    
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
'    If bm.JudgBmd <> True Then
    If bm.JudgBmd <> True Or bm.JudgAntnp <> True Then
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        bJudg = False
    End If
    
    typ_y013.MESDATA6 = bm.JudgDataAve            '�@��2003/05/20 ooba
    typ_y013.MESDATA7 = bm.JudgDataMax
    typ_y013.MESDATA8 = bm.JudgDataMin
    typ_y013.MESDATA9 = bm.JudgDataMBP            '�@��2003/05/20 ooba
     
    WfCrBmdJudg = True
End Function

'�T�v      :OSF����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_y013      ,I  ,typ_TBCMY013                         :OSF���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :osfflg        ,I  ,Integer                              :OSF�׸�(1:OSF1, 2:OSF2, 3:OSF3, 4:OSF4)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :OSF������s��
'����      :
Public Function WfCrOsfJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y013 As typ_TBCMY013, _
                            bJudg As Boolean, _
                            osfflg As Integer, _
                            TmpData() As String) As Boolean

    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim os      As W_OSF                    'OSF�\����
    Dim keisu   As Double
    Dim c0      As Integer
    
    Const keisu1 As Double = 1.8248175
    Const keisu2 As Double = 1.8518519
    Const keisu3 As Double = 1.9230769
    Const keisu4 As Double = 3.649635
    Const keisu5 As Double = 3.7037037
    Const keisu6 As Double = 3.8461538
    '' 2006/09/25 SMP)kondoh Add -s-
    Const keisu7 As Double = 7.6923077
    '' 2006/09/25 SMP)kondoh Add -e-
        
    bJudg = True

    'OSF��������ݒ�
    Select Case osfflg
    Case 1
        os.GuaranteeOsf.cMeth = typ_si.HWFOF1SH  '�i�v�e�n�r�e�P����ʒu�Q��
        os.GuaranteeOsf.cCount = typ_si.HWFOF1ST '�i�v�e�n�r�e�P����ʒu�Q�_
        os.GuaranteeOsf.cPos = typ_si.HWFOF1SR   '�i�v�e�n�r�e�P����ʒu�Q��
        os.GuaranteeOsf.cObj = typ_si.HWFOF1HT   '�i�v�e�a�l�c�P�ۏؕ��@�Q��
        os.GuaranteeOsf.cJudg = typ_si.HWFOF1HS  '�i�v�e�a�l�c�P�ۏؕ��@�Q��
        os.SpecOsfAveMax = typ_si.HWFOF1AX       '�i�v�e�n�r�e�P���Ϗ��
        os.SpecOsfMax = typ_si.HWFOF1MX          '�i�v�e�n�r�e�P���
        os.JudgDataPTK = NtoS(typ_si.HWFOSF1PTK) '�i�v�e�n�r�e�P�p�^���敪�@2003/05/17 ooba
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        os.Antnp = typ_si.HWFANTNP                '�i�v�e�`�m���x
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    Case 2
        os.GuaranteeOsf.cMeth = typ_si.HWFOF2SH  '�i�v�e�n�r�e�Q����ʒu�Q��
        os.GuaranteeOsf.cCount = typ_si.HWFOF2ST '�i�v�e�n�r�e�Q����ʒu�Q�_
        os.GuaranteeOsf.cPos = typ_si.HWFOF2SR   '�i�v�e�n�r�e�Q����ʒu�Q��
        os.GuaranteeOsf.cObj = typ_si.HWFOF2HT   '�i�v�e�a�l�c�Q�ۏؕ��@�Q��
        os.GuaranteeOsf.cJudg = typ_si.HWFOF2HS  '�i�v�e�a�l�c�Q�ۏؕ��@�Q��
        os.SpecOsfAveMax = typ_si.HWFOF2AX       '�i�v�e�n�r�e�Q���Ϗ��
        os.SpecOsfMax = typ_si.HWFOF2MX          '�i�v�e�n�r�e�Q���
        os.JudgDataPTK = NtoS(typ_si.HWFOSF2PTK) '�i�v�e�n�r�e�Q�p�^���敪�@2003/05/17 ooba
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        os.Antnp = typ_si.HWFANTNP                '�i�v�e�`�m���x
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    Case 3
        os.GuaranteeOsf.cMeth = typ_si.HWFOF3SH  '�i�v�e�n�r�e�R����ʒu�Q��
        os.GuaranteeOsf.cCount = typ_si.HWFOF3ST '�i�v�e�n�r�e�R����ʒu�Q�_
        os.GuaranteeOsf.cPos = typ_si.HWFOF3SR   '�i�v�e�n�r�e�R����ʒu�Q��
        os.GuaranteeOsf.cObj = typ_si.HWFOF3HT   '�i�v�e�a�l�c�R�ۏؕ��@�Q��
        os.GuaranteeOsf.cJudg = typ_si.HWFOF3HS  '�i�v�e�a�l�c�R�ۏؕ��@�Q��
        os.SpecOsfAveMax = typ_si.HWFOF3AX       '�i�v�e�n�r�e�R���Ϗ��
        os.SpecOsfMax = typ_si.HWFOF3MX          '�i�v�e�n�r�e�R���
        os.JudgDataPTK = NtoS(typ_si.HWFOSF3PTK) '�i�v�e�n�r�e�R�p�^���敪�@2003/05/17 ooba
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        os.Antnp = typ_si.HWFANTNP                '�i�v�e�`�m���x
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    Case 4
        os.GuaranteeOsf.cMeth = typ_si.HWFOF4SH  '�i�v�e�n�r�e�S����ʒu�Q��
        os.GuaranteeOsf.cCount = typ_si.HWFOF4ST '�i�v�e�n�r�e�S����ʒu�Q�_
        os.GuaranteeOsf.cPos = typ_si.HWFOF4SR   '�i�v�e�n�r�e�S����ʒu�Q��
        os.GuaranteeOsf.cObj = typ_si.HWFOF4HT   '�i�v�e�a�l�c�S�ۏؕ��@�Q��
        os.GuaranteeOsf.cJudg = typ_si.HWFOF4HS  '�i�v�e�a�l�c�S�ۏؕ��@�Q��
        os.SpecOsfAveMax = typ_si.HWFOF4AX       '�i�v�e�n�r�e�S���Ϗ��
        os.SpecOsfMax = typ_si.HWFOF4MX          '�i�v�e�n�r�e�S���
        os.JudgDataPTK = NtoS(typ_si.HWFOSF4PTK) '�i�v�e�n�r�e�S�p�^���敪�@2003/05/17 ooba
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        os.Antnp = typ_si.HWFANTNP                '�i�v�e�`�m���x
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    End Select
    
    If os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "3" Then
        keisu = keisu1
    ElseIf os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "5" Then
        keisu = keisu2
    ElseIf os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu3
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "3" Then
        keisu = keisu4
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "5" Then
        keisu = keisu5
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu6
    '' 2006/09/25 SMP)kondoh Add -s-
    ElseIf os.GuaranteeOsf.cMeth = "E" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu7
    '' 2006/09/25 SMP)kondoh Add -e-
    Else
        bJudg = False
        WfCrOsfJudg = False
        Exit Function
    End If

    With os
        .OSF(0) = NtoZ2(typ_y013.MESDATA1)                   'OSF����l
        .OSF(1) = NtoZ2(typ_y013.MESDATA2)                   'OSF����l
        .OSF(2) = NtoZ2(typ_y013.MESDATA3)                   'OSF����l
        .OSF(3) = NtoZ2(typ_y013.MESDATA4)                   'OSF����l
        .OSF(4) = NtoZ2(typ_y013.MESDATA5)                   'OSF����l
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '�`�F�b�N�pAN���x��ǉ�
        .OsfAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        For c0 = 0 To 4
            .OSF(c0) = IIf(.OSF(c0) <> -1, .OSF(c0) * keisu, -1)
        Next
        typ_y013.MESDATA6 = typ_y013.MESDATA6 * 100
        .OSFp(0) = Trim(typ_y013.MESDATA9)                   'OSF�p�^�[������(��)�@��2003/05/17 ooba
        .OSFp(1) = Trim(typ_y013.MESDATA12)                  'OSF�p�^�[������(��)
        .OSFp(2) = Trim(typ_y013.MESDATA15)                  'OSF�p�^�[������(��)�@��2003/05/17 ooba
    End With
    typ_y013.MESDATA6 = typ_y013.MESDATA6 * 100
    
    'OSF����
    If WfOSFJudg(os, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrOsfJudg = False
        Exit Function
    End If
    
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
'    If os.JudgOsf <> True Then
    If os.JudgOsf <> True Or os.JudgAntnp <> True Then
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        bJudg = False
    End If
    
    TmpData(0) = os.JudgDataAve                                  ' 2003/05/20 ooba
    TmpData(1) = os.JudgDataMax                                  ' 2003/05/20 ooba
'    typ_y013.MESDATA7 = os.JudgDataAve
'    typ_y013.MESDATA8 = os.JudgDataMax
     WfCrOsfJudg = True
End Function

'�T�v      :DSOD����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_y013      ,I  ,typ_TBCMY013                         :DSOD���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :CS������s��
'����      :
Public Function WfCrDsodjudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                             typ_y013 As typ_TBCMY013, _
                             bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim Dsod    As W_DSOD                   'W_DSOD�\����
    
    bJudg = True
        
    'DSOD��������ݒ�
    Dsod.GuaranteeDsod.cMeth = ""  '
    Dsod.GuaranteeDsod.cCount = "" '
    Dsod.GuaranteeDsod.cPos = ""  '
    Dsod.GuaranteeDsod.cObj = typ_si.HWFDSOHT    '�i�v�e�c�r�n�c�ۏؕ��@�Q��
    Dsod.GuaranteeDsod.cJudg = typ_si.HWFDSOHS   '�i�v�e�c�r�n�c�ۏؕ��@�Q��
    Dsod.SpecDsodMin = typ_si.HWFDSOMN           '�i�v�e�c�r�n�c����
    Dsod.SpecDsodMax = typ_si.HWFDSOMX           '�i�v�e�c�r�n�c���
    Dsod.JudgDataPTK = NtoS(typ_si.HWFDSOPTK)    '�i�v�e�c�r�n�c�p�^���敪�@04/07/28 ooba
    
    Dsod.Dsod = NtoZ2(typ_y013.MESDATA1)         'DSOD����l
    Dsod.Dsodp(0) = Trim(typ_y013.MESDATA4)      'DSOD�p�^������1�@04/07/28 ooba
    Dsod.Dsodp(1) = Trim(typ_y013.MESDATA7)      'DSOD�p�^������2�@04/07/28 ooba
    
    Dsod.Antnp = typ_si.HWFANTNP                        '�iWFAN���x(�d�l)�@06/12/22 ooba
    Dsod.DsodAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))    '�iWFAN���x(����)�@06/12/22 ooba
    
    'DSOD����
    If WfDSODJudg(Dsod, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrDsodjudg = False
        Exit Function
    End If
    
'    If Dsod.JudgDsod <> True Then
    If Dsod.JudgDsod <> True Or Dsod.JudgAntnp <> True Then  'AN���x���茋�ʒǉ��@06/12/22 ooba
        bJudg = False
    End If
    
    WfCrDsodjudg = True
End Function

'�T�v      :DZ����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_y013      ,I  ,typ_TBCMY013                         :DZ���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :DZ������s��
'����      :
Public Function WfCrDzjudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                           typ_y013 As typ_TBCMY013, _
                           bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim DZ      As W_DZ                     'DZ�\����
    
    bJudg = True
        
    'DZ��������ݒ�
    DZ.GuaranteeDz.cMeth = typ_si.HWFMKSPH   '�i�v�e�����בw����ʒu�Q��
    DZ.GuaranteeDz.cCount = typ_si.HWFMKSPT  '�i�v�e�����בw����ʒu�Q�_
    DZ.GuaranteeDz.cPos = typ_si.HWFMKSPR    '�i�v�e�����בw����ʒu�Q��
    DZ.GuaranteeDz.cObj = typ_si.HWFMKHWT    '�i�v�e�����בw�ۏؕ��@�Q��
    DZ.GuaranteeDz.cJudg = typ_si.HWFMKHWS   '�i�v�e�����בw�ۏؕ��@�Q��
    DZ.SpecDzMin = typ_si.HWFMKMIN           '�i�v�e�����בw����
    DZ.SpecDzMax = typ_si.HWFMKMAX           '�i�v�e�����בw���
    
    DZ.DZ(0) = NtoZ2(typ_y013.MESDATA1)               'DZ����l
    DZ.DZ(1) = NtoZ2(typ_y013.MESDATA2)               'DZ����l
    DZ.DZ(2) = NtoZ2(typ_y013.MESDATA3)               'DZ����l
    DZ.DZ(3) = NtoZ2(typ_y013.MESDATA4)               'DZ����l
    
    'DZ����
    If WfDZJudg(DZ, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrDzjudg = False
        Exit Function
    End If
    
    If DZ.JudgDz <> True Then
        bJudg = False
    End If
        
    typ_y013.MESDATA5 = DZ.JudgDataAve
    typ_y013.MESDATA6 = DZ.JudgDataMax
    typ_y013.MESDATA7 = DZ.JudgDataMin
    
    WfCrDzjudg = True
End Function

'�T�v      :SPVFE����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_y013      ,I  ,typ_TBCMY013                         :SPVFE���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :SPV������s��
'����      :
Public Function WfCrSpvjudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y013 As typ_TBCMY013, _
                            bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim sp      As W_SPV                    'SPV�\����
    
    bJudg = True
        
    'SPV��������ݒ�
    sp.GuaranteeSpv.cMeth = typ_si.HWFDLSPH    '�i�v�e�g�U������ʒu�Q��
    sp.GuaranteeSpv.cCount = typ_si.HWFDLSPT   '�i�v�e�g�U������ʒu�Q�_
    sp.GuaranteeSpv.cPos = typ_si.HWFDLSPI     '�i�v�e�g�U������ʒu�Q��
    sp.GuaranteeSpv.cObj = typ_si.HWFDLHWT     '�i�v�e�g�U���ۏؕ��@�Q��
    sp.GuaranteeSpv.cJudg = typ_si.HWFDLHWS    '�i�v�e�g�U���ۏؕ��@�Q��
    
    sp.GuaranteeSpvFe.cMeth = typ_si.HWFSPVSH  '�i�v�e�r�o�u�e�d����ʒu�Q��
    sp.GuaranteeSpvFe.cCount = typ_si.HWFSPVST '�i�v�e�r�o�u�e�d����ʒu�Q�_
    sp.GuaranteeSpvFe.cPos = typ_si.HWFSPVSI   '�i�v�e�r�o�u�e�d����ʒu�Q��
    sp.GuaranteeSpvFe.cObj = typ_si.HWFSPVHT   '�i�v�e�r�o�u�e�d�ۏؕ��@�Q��
    sp.GuaranteeSpvFe.cJudg = typ_si.HWFSPVHS  '�i�v�e�r�o�u�e�d�ۏؕ��@�Q��
    
    sp.SpecSpvMin = typ_si.HWFDLMIN            '�iWF�g�U������
    sp.SpecSpvMax = typ_si.HWFDLMAX            '�iWF�g�U�����
    sp.SpecSpvFeMax = typ_si.HWFSPVMX          '�iWFFe�Z�x���
    '----TEST2004/10
    sp.SpecSpvAvMax = typ_si.HWFSPVAM
    
    sp.Spv(0) = NtoZ2(typ_y013.MESDATA1)                'SPV����l
    sp.Spv(1) = NtoZ2(typ_y013.MESDATA2)                'SPV����l
    sp.Spv(2) = NtoZ2(typ_y013.MESDATA3)                'SPV����l
    sp.Spv(3) = NtoZ2(typ_y013.MESDATA4)                'SPV����l
    sp.Spv(4) = NtoZ2(typ_y013.MESDATA5)                'SPV����l
    
    'SPV����
    If WfSPVJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrSpvjudg = False
        Exit Function
    End If
    
    If sp.JudgSpv <> True Then
        bJudg = False
    End If
    
    WfCrSpvjudg = True
End Function

'�T�v      :DOI����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_y013      ,I  ,typ_TBCMY013                         :DOI���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :doiflg        ,I  ,Integer                              :DOI�׸�(1:DOI1, 2:DOI2, 3:DOI3)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :DOI������s��
'����      :
Public Function WfCrDoiJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y013 As typ_TBCMY013, _
                            bJudg As Boolean, _
                            doiflg As Integer) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim WFDOI   As W_DOI                    'DOI�\����

    bJudg = True

    'DOI��������ݒ�
    Select Case doiflg
    Case 1
        WFDOI.GuaranteeDoi.cMeth = typ_si.HWFOS1SH    '�i�v�e�_�f�͏o�P����ʒu�Q��
        WFDOI.GuaranteeDoi.cCount = typ_si.HWFOS1ST   '�i�v�e�_�f�͏o�P����ʒu�Q�_
        WFDOI.GuaranteeDoi.cPos = typ_si.HWFOS1SI     '�i�v�e�_�f�͏o�P����ʒu�Q��
        WFDOI.GuaranteeDoi.cObj = typ_si.HWFOS1HT     '�i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        WFDOI.GuaranteeDoi.cJudg = typ_si.HWFOS1HS    '�i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        WFDOI.SpecDoiMin = typ_si.HWFOS1MN            '�i�v�e�_�f�͏o�P����
        WFDOI.SpecDoiMax = typ_si.HWFOS1MX            '�i�v�e�_�f�͏o�P���
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        WFDOI.Antnp = typ_si.HWFANTNP                 '�i�v�e�`�m���x
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    Case 2
        WFDOI.GuaranteeDoi.cMeth = typ_si.HWFOS2SH    '�i�v�e�_�f�͏o�P����ʒu�Q��
        WFDOI.GuaranteeDoi.cCount = typ_si.HWFOS2ST   '�i�v�e�_�f�͏o�P����ʒu�Q�_
        WFDOI.GuaranteeDoi.cPos = typ_si.HWFOS2SI     '�i�v�e�_�f�͏o�P����ʒu�Q��
        WFDOI.GuaranteeDoi.cObj = typ_si.HWFOS2HT     '�i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        WFDOI.GuaranteeDoi.cJudg = typ_si.HWFOS2HS    '�i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        WFDOI.SpecDoiMin = typ_si.HWFOS2MN            '�i�v�e�_�f�͏o�P����
        WFDOI.SpecDoiMax = typ_si.HWFOS2MX            '�i�v�e�_�f�͏o�P���
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        WFDOI.Antnp = typ_si.HWFANTNP                 '�i�v�e�`�m���x
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    Case 3
        WFDOI.GuaranteeDoi.cMeth = typ_si.HWFOS3SH    '�i�v�e�_�f�͏o�P����ʒu�Q��
        WFDOI.GuaranteeDoi.cCount = typ_si.HWFOS3ST   '�i�v�e�_�f�͏o�P����ʒu�Q�_
        WFDOI.GuaranteeDoi.cPos = typ_si.HWFOS3SI     '�i�v�e�_�f�͏o�P����ʒu�Q��
        WFDOI.GuaranteeDoi.cObj = typ_si.HWFOS3HT     '�i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        WFDOI.GuaranteeDoi.cJudg = typ_si.HWFOS3HS    '�i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        WFDOI.SpecDoiMin = typ_si.HWFOS3MN            '�i�v�e�_�f�͏o�P����
        WFDOI.SpecDoiMax = typ_si.HWFOS3MX            '�i�v�e�_�f�͏o�P���
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        WFDOI.Antnp = typ_si.HWFANTNP                 '�i�v�e�`�m���x
    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    End Select
    
    WFDOI.Doi(0) = NtoZ2(typ_y013.MESDATA1)                    'DOI����l
    WFDOI.Doi(1) = NtoZ2(typ_y013.MESDATA2)                    'DOI����l
    WFDOI.Doi(2) = NtoZ2(typ_y013.MESDATA3)                    'DOI����l
    WFDOI.Doi(3) = NtoZ2(typ_y013.MESDATA4)                    'DOI����l
    WFDOI.Doi(4) = NtoZ2(typ_y013.MESDATA5)                    'DOI����l
    WFDOI.Doi(5) = NtoZ2(typ_y013.MESDATA6)                    'DOI����l�@-*-*-�@20010912 add
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '�`�F�b�N�pAN���x��ǉ�
    WFDOI.DoiAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    'DOI����
    If WfDOiJudg(WFDOI, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrDoiJudg = False
        Exit Function
    End If
    
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
'    If WFDOI.JudgDoi <> True Then
    If WFDOI.JudgDoi <> True Or WFDOI.JudgAntnp <> True Then
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        bJudg = False
    End If
        
    WfCrDoiJudg = True
End Function
'''''============================================================================================================================
'''''
''''''�T�v      :���я��f�[�^�ݒ�
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :typ_a         ,I  ,typ_AllTypesC ,�e���\����
''''''����      :�]���������\���̂ɐݒ肷��
''''''����      :
'''''Public Function JudgAllRsltData() As FUNCTION_RETURN
'''''
'''''    TotalJudg = True
'''''
''''''''''    typ_rtInit '2001/09/14 S.Sano
'''''
'''''    JudgAllRsltData = FUNCTION_RETURN_FAILURE
'''''
'''''    '�d�l�����x���擾
'''''    SpecJudgCheck
'''''
'''''
''''''''''    WFCJudgDialog.WFCErrorMessage SelectSxlID & " ******************"
'''''    '���уf�[�^����(TOP)
'''''    If WfAllJudg(SxlTop) = FUNCTION_RETURN_FAILURE Then
'''''        Exit Function
'''''    End If
'''''    '���уf�[�^����(TAIL)
'''''    If WfAllJudg(SxlTail) = FUNCTION_RETURN_FAILURE Then
'''''        Exit Function
'''''    End If
'''''
'''''    JudgAllRsltData = FUNCTION_RETURN_SUCCESS
'''''End Function

'�T�v      :AOI����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_y013      ,I  ,typ_TBCMY013                         :AOI���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :AOI������s��
'����      :03/12/09 ooba
Public Function WfCrAoiJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y013 As typ_TBCMY013, _
                            bJudg As Boolean) As Boolean
                            
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim WFAOI   As W_AOI                    'AOI�\����

    bJudg = True

    'AOI��������ݒ�
    WFAOI.GuaranteeAoi.cMeth = typ_si.HWFZOSPH    '�i�v�e�c���_�f����ʒu�Q��
    WFAOI.GuaranteeAoi.cCount = typ_si.HWFZOSPT   '�i�v�e�c���_�f����ʒu�Q�_
    WFAOI.GuaranteeAoi.cPos = typ_si.HWFZOSPI     '�i�v�e�c���_�f����ʒu�Q��
    WFAOI.GuaranteeAoi.cObj = typ_si.HWFZOHWT     '�i�v�e�c���_�f�ۏؕ��@�Q��
    WFAOI.GuaranteeAoi.cJudg = typ_si.HWFZOHWS    '�i�v�e�c���_�f�ۏؕ��@�Q��
    WFAOI.SpecAoiMin = typ_si.HWFZOMIN            '�i�v�e�c���_�f����
    WFAOI.SpecAoiMax = typ_si.HWFZOMAX            '�i�v�e�c���_�f���
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    WFAOI.Antnp = typ_si.HWFANTNP                 '�i�v�e�`�m���x
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    WFAOI.AOI(0) = NtoZ2(typ_y013.MESDATA4)       'AOI����l
    WFAOI.AOI(1) = NtoZ2(typ_y013.MESDATA5)       'AOI����l
    WFAOI.AOI(2) = NtoZ2(typ_y013.MESDATA6)       'AOI����l
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    '�`�F�b�N�pAN���x��ǉ�
    WFAOI.AoiAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    'AOI����
    If WfAOiJudg(WFAOI, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrAoiJudg = False
        Exit Function
    End If

'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
'    If WFAOI.JudgAoi <> True Then
    If WFAOI.JudgAoi <> True Or WFAOI.JudgAntnp <> True Then
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        bJudg = False
    End If
        
    WfCrAoiJudg = True
End Function

'�T�v      :SIRD(SD)����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_y013      ,I  ,typ_TBCMJ022                         :SD���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :SD������s��
'����      :2010/01/07 SIRD�Ή� Y.Hitomi
Public Function WfCrSdjudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                           typ_j022 As typ_TBCMJ022, _
                           bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim SD      As W_SD                     'SD�\����
    
    bJudg = True
        
    'SD��������ݒ�
    SD.GuaranteeSd.cObj = typ_si.HWFSIRDHT   '����]�ʕۏؕ��@�Q��
    SD.GuaranteeSd.cJudg = typ_si.HWFSIRDHS  '����]�ʕۏؕ��@�Q��
    SD.SpecSdMax = typ_si.HWFSIRDMX          '����]�ʏ��
    
    SD.SdMeasData = val((typ_j022.SIRDCNT))  'SD����l
    
    'DZ����
    If WfSDJudg(SD, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrSdjudg = False
        Exit Function
    End If
    
    If SD.JudgSD = False Then
        bJudg = False
    End If
            
    WfCrSdjudg = True
End Function
'�T�v      :GD����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_j015      ,I  ,typ_TBCMJ015                         :GD���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :GD������s��
'����      :05/01/31 ooba
Public Function WfCrGdJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_j015 As typ_TBCMJ015, _
                            bJudg As Boolean) As Boolean

    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim WFGD    As W_GD                     'GD�\����
    Dim iCnt    As Integer
    Dim bUpdFlg As Boolean                  'TBCMJ015-UPDATE�L���׸�
    Dim bDenData    As Boolean              'Den�����ް������׸ށ@05/10/25 ooba
    Dim bLdlData    As Boolean              'L/DL�����ް������׸ށ@05/10/25 ooba
    Dim bDvd2Data   As Boolean              'DVD2�����ް������׸ށ@05/10/25 ooba
    Dim SYORIKBN As String                  '�����o�^���̃t���O�@�@10/04/01 hama


'   Dim WFGD2       As W_GD                 'GD�\����   '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    SYORIKBN = ""
    bJudg = True
    
    'GD��������ݒ�
    WFGD.GuaranteeDen.cMeth = ""                '����ʒu_��
    WFGD.GuaranteeDen.cCount = ""               '����ʒu_�_
    WFGD.GuaranteeDen.cPos = ""                 '����ʒu_��
    WFGD.GuaranteeDen.cObj = typ_si.HWFDENHT    '�ۏؕ��@_��
    WFGD.GuaranteeDen.cJudg = typ_si.HWFDENHS   '�ۏؕ��@_��
    
    WFGD.GuaranteeLdl.cMeth = ""                '����ʒu_��
    WFGD.GuaranteeLdl.cCount = ""               '����ʒu_�_
    WFGD.GuaranteeLdl.cPos = ""                 '����ʒu_��
    WFGD.GuaranteeLdl.cObj = typ_si.HWFLDLHT    '�ۏؕ��@_��
    WFGD.GuaranteeLdl.cJudg = typ_si.HWFLDLHS   '�ۏؕ��@_��
    
    WFGD.GuaranteeDvd2.cMeth = ""               '����ʒu_��
    WFGD.GuaranteeDvd2.cCount = ""              '����ʒu_�_
    WFGD.GuaranteeDvd2.cPos = ""                '����ʒu_��
    WFGD.GuaranteeDvd2.cObj = typ_si.HWFDVDHT   '�ۏؕ��@_��
    WFGD.GuaranteeDvd2.cJudg = typ_si.HWFDVDHS  '�ۏؕ��@_��
    
    WFGD.JudgFlagDen = typ_si.HWFDENKU          '�iWFDen�����L��
    WFGD.JudgFlagLdl = typ_si.HWFLDLKU          '�iWFL/DL�����L��
    WFGD.JudgFlagDvd2 = typ_si.HWFDVDKU         '�iWFDVD2�����L��
    
    WFGD.SpecDenMin = typ_si.HWFDENMN           '�iWFDen����
    WFGD.SpecDenMax = typ_si.HWFDENMX           '�iWFDen���
    WFGD.SpecLdlMin = typ_si.HWFLDLMN           '�iWFL/DL����
    WFGD.SpecLdlMax = typ_si.HWFLDLMX           '�iWFL/DL���
    WFGD.SpecDvd2Min = typ_si.HWFDVDMNN         '�iWFDVD2����
    WFGD.SpecDvd2Max = typ_si.HWFDVDMXN         '�iWFDVD2���
    
'*** UPDATE �� Y.SIMIZU 2005/10/7 �iWFGDײݐ��ǉ�
    WFGD.SpecGdLine = typ_si.HWFGDLINE          '�iWFGDײݐ�
'*** UPDATE �� Y.SIMIZU 2005/10/7 �iWFGDײݐ��ǉ�

    WFGD.Antnp = typ_si.HWFANTNP                        '�iWFAN���x(�d�l)�@06/12/22 ooba
    WFGD.GdAntnp = NtoZ2(Mid(typ_j015.DKAN, 3, 4))      '�iWFAN���x(����)�@06/12/22 ooba
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
''    If typ_si.WFHSGDCW = "1" Then
''        ' ����
''        WFGD.GDPTK = typ_si.HSXGDPTK
''        WFGD.ZeroLdlMin = typ_si.HSXLDLRMN
''        WFGD.ZeroLdlMax = typ_si.HSXLDLRMX
''    Else
        ' WF
        WFGD.GDPTK = typ_si.HWFGDPTK
        WFGD.ZeroLdlMin = typ_si.HWFLDLRMN
        WFGD.ZeroLdlMax = typ_si.HWFLDLRMX
''    End If
    ' �\���̂�GD����(WF)�����A�����̎��т��ݒ肳���ꍇ������
    WFGD.LdlMin = typ_j015.MSZEROMN
    WFGD.LdlMax = typ_j015.MSZEROMX
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    
'*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݐ���3����4.5����5�łȂ��ꍇ�͔���װ
    If WFGD.SpecGdLine <> 3 And WFGD.SpecGdLine <> 4.5 And WFGD.SpecGdLine <> 5 Then
        bJudg = False
        WfCrGdJudg = False
        Exit Function
    End If
'*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݐ���3����4.5����5�łȂ��ꍇ�͔���װ

'*** UPDATE �� Y.SIMIZU 2005/10/7 GDײݐ����̎��т����邩����������
'    If ChkGD_Data(typ_j015, WFGD) <> FUNCTION_RETURN_SUCCESS Then
    If ChkGD_Data(typ_j015, WFGD, bDenData, bLdlData, bDvd2Data) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrGdJudg = False
        Exit Function
    End If
'*** UPDATE �� Y.SIMIZU 2005/10/7 GDײݐ����̎��т����邩����������
    
    '���茋���ް������݂��Ȃ��ꍇ�͌v�Z�ŋ��߂�
    If typ_j015.MSRSDEN = -1 Or typ_j015.MSRSLDL = -1 Or typ_j015.MSRSDVD2 = -1 Then
    '*** UPDATE �� Y.SIMIZU 2005/10/7 �d�l��GDײݐ�,DVD2�㉺����n���悤�ɕύX
'        If Calculate_GD(typ_j015) <> FUNCTION_RETURN_SUCCESS Then
        If Calculate_GD(typ_j015, WFGD.SpecGdLine, WFGD.SpecDvd2Min, WFGD.SpecDvd2Max) <> FUNCTION_RETURN_SUCCESS Then
    '*** UPDATE �� Y.SIMIZU 2005/10/7 �d�l��GDײݐ�,DVD2�㉺����n���悤�ɕύX
            bJudg = False
            WfCrGdJudg = False
            Exit Function
        End If
        
        If Not bDenData Then typ_j015.MSRSDEN = -1      '05/10/25 ooba
        If Not bLdlData Then typ_j015.MSRSLDL = -1      '05/10/25 ooba
        If Not bDvd2Data Then typ_j015.MSRSDVD2 = -1    '05/10/25 ooba
        
        'L/DL���������ǉ��@05/10/26 ooba START ==================================>
        If Len(CStr(typ_j015.MSRSLDL)) > 3 Then
            bJudg = False
            WfCrGdJudg = False
            Exit Function
        End If
        'L/DL���������ǉ��@05/10/26 ooba END ====================================>
        SYORIKBN = "1"
    End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
'' 2010/04/01 ��́�SXL�����[�Ƀp�^�[�������t����ׂɃp�^�[�����ʂ��c����������
''            ���ׂ̈ɂ��̏ꏊ�ɉ��L�̃��W�b�N��ǉ������̂��Ǝv����B
''            MSRSDEN/MSRSLDL/MSRSDVD2���ۑ��f�[�^�����邩�̃`�F�b�N�������̂���
''�@�@�@�@�@�@���肵�ĂȂ��ꍇ�́A���܂ł����Ă����̏�����ʉ߂��邱�ƂɂȂ�B
''�@�@�@�@�@�@�܂��A���̏������Ǝv���Ă��̏��������Ă���̂���WFGD2�̍\���̂�
''�@�@�@�@�@�@������̂��Ӗ����s���ł���B
''�@�@�@�@�@�@�p�^�[���o�^�p�Ƃ���̂�����Ƃ��Ȃ�̖ړI�Ȃ̂����s���m�ł���B
''�@�@�@�@�@�@�܂��AWFGdJudg��GD�p�^�[���̕���ɂ��g�����̔��f�����邱�Ǝ��̂�
''�@�@�@�@�@�@���������̂ł͂Ȃ��B
''
''       If WFGD.GDPTK = "1" Or WFGD.GDPTK = "2" Then
''          WFGD2 = WFGD
''
''           '�đ��茋�ʔ��f
'            WFGD2.Den = typ_j015.MSRSDEN
'            WFGD2.Dvd2 = typ_j015.MSRSDVD2
'            WFGD2.Ldl = typ_j015.MSRSLDL
''           WFGD2.LdlMin = typ_j015.MSZEROMN
''           WFGD2.LdlMax = typ_j015.MSZEROMX
''
''           'GD����
''           If WfGdJudg(WFGD2, typ_j015.HSFLG, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
'                bJudg = False
'                WfCrGdJudg = False
'                Exit Function
''            End If
''            ''L/DL�̔��茋�ʂ�GD����(WF)�ɔ��f
'            If WFGD2.JudgLdlPtn = True Then
''                typ_j015.PTNJUDGRES = "1"
''            Else
''                typ_j015.PTNJUDGRES = "9"
''            End If
''        Else
''            typ_j015.PTNJUDGRES = " "
''        End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
''
''        'TBCMJ015-UPDATE�pGD�����ް����
''        bUpdFlg = True
''        If iCntJ015upd > 0 Then
''            For iCnt = 1 To iCntJ015upd
''                '�����ް������݂���ꍇ
''                If typ_j015.SMPLNO = typ_J015_WFGDUpd(iCnt).SMPLNO Then
''                    bUpdFlg = False
''                   Exit For
''               End If
''           Next
''        End If
''        If bUpdFlg Then
''            iCntJ015upd = iCntJ015upd + 1
''            ReDim Preserve typ_J015_WFGDUpd(iCntJ015upd)
''            typ_J015_WFGDUpd(iCntJ015upd) = typ_j015
''        End If
'' 2010/3/31 GD�̃p�^�[�������GD����ɂ��ĕs��v������
'' ���̏ꏊ�Ńp�^�[�����N���A���邱�Ƃ��Ӗ����킩��Ȃ��B
'' �J�������͂Ȃ������Ƃ̎��Ȃ̂Ŏ�荇�����R�����g�ɂ��܂��B
'' ���W�b�N�I�ɖ�肪������_�܂ŘA�������������B
''     WFGD.GDPTK = " "
''---------------------------------------------------------
    
''    End If
    
    WFGD.Den = typ_j015.MSRSDEN                  'Den�v�Z�l
    WFGD.Ldl = typ_j015.MSRSLDL                  'L/DL�v�Z�l
    WFGD.Dvd2 = typ_j015.MSRSDVD2                'DVD2�v�Z�l
    WFGD.LdlMin = typ_j015.MSZEROMN
    WFGD.LdlMax = typ_j015.MSZEROMX              'GD����


'    If WfGdJudg(WFGD, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
    '�ۏ��׸ޒǉ��@06/12/22 ooba
    If WfGdJudg(WFGD, typ_j015.HSFLG, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrGdJudg = False
        
        If SYORIKBN = "1" Then
            ''L/DL�̔��茋�ʂ�GD����(WF)�ɔ��f
                   typ_j015.PTNJUDGRES = " "
              If WFGD.JudgLdlPtn = True Then
                   typ_j015.PTNJUDGRES = "1"
              Else
                   typ_j015.PTNJUDGRES = "9"
              End If

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

            ''TBCMJ015-UPDATE�pGD�����ް����
                  bUpdFlg = True
              If iCntJ015upd > 0 Then
                For iCnt = 1 To iCntJ015upd
                  '�����ް������݂���ꍇ
                  If typ_j015.SMPLNO = typ_J015_WFGDUpd(iCnt).SMPLNO Then
                      bUpdFlg = False
                      Exit For
                  End If
                Next
              End If
              If bUpdFlg Then
                 iCntJ015upd = iCntJ015upd + 1
                 ReDim Preserve typ_J015_WFGDUpd(iCntJ015upd)
                 typ_J015_WFGDUpd(iCntJ015upd) = typ_j015
              End If
        End If
        Exit Function
    Else
        If SYORIKBN = "1" Then
            ''L/DL�̔��茋�ʂ�GD����(WF)�ɔ��f
                   typ_j015.PTNJUDGRES = " "
              If WFGD.JudgLdlPtn = True Then
                   typ_j015.PTNJUDGRES = "1"
              Else
                   typ_j015.PTNJUDGRES = "9"
              End If

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

            ''TBCMJ015-UPDATE�pGD�����ް����
                  bUpdFlg = True
              If iCntJ015upd > 0 Then
                For iCnt = 1 To iCntJ015upd
                  '�����ް������݂���ꍇ
                  If typ_j015.SMPLNO = typ_J015_WFGDUpd(iCnt).SMPLNO Then
                      bUpdFlg = False
                      Exit For
                  End If
                Next
              End If
              If bUpdFlg Then
                 iCntJ015upd = iCntJ015upd + 1
                 ReDim Preserve typ_J015_WFGDUpd(iCntJ015upd)
                 typ_J015_WFGDUpd(iCntJ015upd) = typ_j015
              End If
        End If
      End If
'    If WFGD.JudgDen <> True Or WFGD.JudgLdl <> True Or WFGD.JudgDvd2 <> True Then
    'AN���x���茋�ʒǉ��@06/12/22 ooba
    If WFGD.JudgDen <> True Or WFGD.JudgLdl <> True Or WFGD.JudgDvd2 <> True _
                                                            Or WFGD.JudgAntnp <> True Then
        bJudg = False
    End If
    
'--------------- 2008/07/25 INSERT START  By Systech ---------------
    pbGDJudgeTbl(1) = WFGD.JudgDen
    pbGDJudgeTbl(2) = WFGD.JudgDvd2
    pbGDJudgeTbl(3) = WFGD.JudgLdl
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
    
    WfCrGdJudg = True
    
End Function

'�T�v      :GD�v�Z
'���Ұ�    :�ϐ���        ,IO ,�^                        :����
'          :tJ015         ,I  ,typ_TBCMJ015              :GD���э\����
'          :�߂�l        ,O  ,FUNCTION_RETURN           :���� = FUNCTION_RETURN_SUCCESS : OK
'                                                               FUNCTION_RETURN_FAILURE : NG
'����      :GD���茋�ʂ��v�Z�ŋ��߂�B
'����      :05/01/31 ooba
'*** UPDATE �� Y.SIMIZU 2005/10/7 �d�l��GDײݐ�,DVD2�㉺���������ɂ���
'Public Function Calculate_GD(tJ015 As typ_TBCMJ015) As FUNCTION_RETURN
Public Function Calculate_GD(tJ015 As typ_TBCMJ015, ByVal iNum As Single, ByVal dSpecDvd2Min As Double, ByVal dSpecDvd2Max As Double) As FUNCTION_RETURN
'*** UPDATE �� Y.SIMIZU 2005/10/7 �d�l��GDײݐ�,DVD2�㉺���������ɂ���

    Dim iCntX As Integer            '����X(5)
    Dim iCntY As Integer            '����Y(15)
    Dim iPoint As Integer           '����_��
    Dim iNoZero As Integer          'Den�̕��ϒl������_���猩�ľ�ۂłȂ��Ȃ����_�܂ł̌�
    Dim dSum As Double
    Dim dAveDen(15) As Double       '�e����_�̕��ϒl(Den)
    Dim dAveLDL(15) As Double       '�e����_�̕��ϒl(L/DL)
    
'*** UPDATE �� Y.SIMIZU 2005/10/7 �����Ƃ���,�d�l��GDײݐ����擾����
'    Dim iNum As Integer             '����l��(3or5)
'*** UPDATE �� Y.SIMIZU 2005/10/7 �����Ƃ���,�d�l��GDײݐ����擾����
    Dim iDen(5, 15) As Integer      '����lDen
    Dim iLDL(5, 15) As Integer      '����lL/DL
    Dim iDVD2(5) As Integer         '����lDVD2
    Dim bDVD2flg As Boolean         '����lDVD2�����׸�(True:�L�AFalse:��)
'*** UPDATE �� Y.SIMIZU 2005/10/7
    Dim iNum2 As Integer
'*** UPDATE �� Y.SIMIZU 2005/10/7
    
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    Dim iZeroCnt        As Integer          ' ZERO�J�E���^
    Dim dLDLSum(15)     As Double           ' L/DL���v
    Dim dLDLZero(15)    As Double           ' L/DL�A��0
    Dim iLDLZeroCnt     As Integer          '
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End


'*** UPDATE �� Y.SIMIZU 2005/10/7 GDײݐ�,DVD2�׸ނ͎d�l��GDײݐ����g�p����
'    '�ϐ���GD����l�A����l�����
'    If CalcGD_DataSet(tJ015, iNum, bDVD2flg, iDen, iLDL, iDVD2) <> FUNCTION_RETURN_SUCCESS Then
    If CalcGD_DataSet(tJ015, iDen, iLDL, iDVD2) <> FUNCTION_RETURN_SUCCESS Then
'*** UPDATE �� Y.SIMIZU 2005/10/7 GDײݐ�,DVD2�׸ނ͎d�l��GDײݐ����g�p����
        Calculate_GD = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    '������
'*** UPDATE �� Y.SIMIZU 2005/10/7 ײݐ�������l�����邩�𒲂ׂ�DVD2�̌v�Z���@��ς���
    '������
    bDVD2flg = True

    '3ײ݂̏ꍇ
    If iNum = 3 Then
        iNum2 = 3
    '4.5ײݖ���5ײ݂̏ꍇ
    Else
        iNum2 = 5
    End If
    
    For iCntX = 1 To iNum2
        '����l������Ȃ��ꍇ�͌v�Z�ɂ����DVD2������
        If iDVD2(iCntX) = -1 Then
            bDVD2flg = False
        End If
    Next
'*** UPDATE �� Y.SIMIZU 2005/10/7 ײݐ�������l�����邩�𒲂ׂ�DVD2�̌v�Z���@��ς���
    
'--DVD2�l�̌v�Z

    '����lDVD2�����݂���ꍇ
    If bDVD2flg = True Then
    '*** UPDATE �� Y.SIMIZU 2005/10/7 4.5ײݑΉ�
'        'DVD2�W�v
'        For iCntX = 1 To iNum
'            If iDVD2(iCntX) = -1 Then
'                GoTo MEAS_TEN1
'            End If
'            dSum = dSum + iDVD2(iCntX)
'        Next

        '3ײ݂̏ꍇ
        If iNum = 3 Then
            iNum2 = 3
        '4.5ײ݂̏ꍇ
        Else
            iNum2 = 5
        End If
        
        'DVD2�W�v
        For iCntX = 1 To iNum2
            If iDVD2(iCntX) = -1 Then
                GoTo MEAS_TEN1
            End If
            dSum = dSum + iDVD2(iCntX)
        Next
    '*** UPDATE �� Y.SIMIZU 2005/10/7 4.5ײݑΉ�
MEAS_TEN1:
        'DVD2�v�Z(����)
        '�����_��2�ʂŎl�̌ܓ��A��1�ʂŐ؂�̂�
        tJ015.MSRSDVD2 = Int(Round(dSum / (iCntX - 1), 1))
    End If
    
'*** UPDATE �� Y.SIMIZU 2005/10/7 4.5ײݑΉ�
'    '����lDen���DVD2�����߂�
'    '�e����_�̕��ς����߂�
'    For iCntY = 1 To 15
'        dSum = 0
'        For iCntX = 1 To iNum
'            If iDen(iCntX, iCntY) = -1 Then
'                GoTo MEAS_TEN2
'            End If
'            dSum = dSum + iDen(iCntX, iCntY)
'        Next
'        dAveDen(iCntY) = dSum / (iCntX - 1)
'    Next

    '����lDen���DVD2�����߂�
    '�e����_�̕��ς����߂�
    For iCntY = 1 To 15
        dSum = 0
        
        '����_7�܂�
        If iCntY <= 7 Then
            '3ײ݂̏ꍇ
            If iNum = 3 Then
                iNum2 = 3
            '4.5ײݖ���5ײ݂̏ꍇ
            Else
                iNum2 = 5
            End If
        '����_8����
        Else
            '3ײ݂̏ꍇ
            If iNum = 3 Then
                iNum2 = 3
            '4.5ײ݂̏ꍇ
            ElseIf iNum = 4.5 Then
                iNum2 = 4
            '5ײ݂̏ꍇ
            Else
                iNum2 = 5
            End If
        End If
        
        For iCntX = 1 To iNum2
            If iDen(iCntX, iCntY) = -1 Then
                GoTo MEAS_TEN2
            End If
            dSum = dSum + iDen(iCntX, iCntY)
        Next
        dAveDen(iCntY) = dSum / (iCntX - 1)
    Next
'*** UPDATE �� Y.SIMIZU 2005/10/7 4.5ײݑΉ�
MEAS_TEN2:
    iPoint = iCntY - 1

    '����_���猩��0�łȂ��Ȃ����_�܂ł̌����擾(DVD2�͈͂��擾)
    For iCntY = iPoint To 1 Step -1
        If dAveDen(iCntY) <> 0 Then
            Exit For
        End If
    Next
    iNoZero = iCntY
    
    '����lDVD2�����݂��Ȃ��ꍇ
    If bDVD2flg = False Then
        'DVD2�v�Z
        tJ015.MSRSDVD2 = Round(iNoZero * 2 * 10, 0)
    End If

'--Den�l�̌v�Z
    
    '��AVE�����߂�
    dSum = 0
    For iCntY = 1 To iPoint
        dSum = dSum + dAveDen(iCntY)
    Next
    
    'Den�v�Z
    If tJ015.MSRSDVD2 = 0 Then
        tJ015.MSRSDEN = 0
    Else
        tJ015.MSRSDEN = RoundUp((dSum * 10) / (tJ015.MSRSDVD2 / 20), 0)
    End If

'--L/DL�l�̌v�Z

    If iNoZero = iPoint Then
        tJ015.MSRSLDL = 0
    Else
    
        'L/DL�e����_�̕��ς����߂�
        For iCntY = iNoZero + 1 To iPoint
        '*** UPDATE �� Y.SIMIZU 2005/10/7 4.5ײݑΉ�
'            dSum = 0
'            For iCntX = 0 To iNum
'                dSum = dSum + iLDL(iCntX, iCntY)
'            Next
'            dAveLDL(iCntY) = dSum / (iCntX - 1)

            '����_7�܂�
            If iCntY <= 7 Then
                '3ײ݂̏ꍇ
                If iNum = 3 Then
                    iNum2 = 3
                '4.5ײݖ���5ײ݂̏ꍇ
                Else
                    iNum2 = 5
                End If
            '����_8����
            Else
                '3ײ݂̏ꍇ
                If iNum = 3 Then
                    iNum2 = 3
                '4.5ײ݂̏ꍇ
                ElseIf iNum = 4.5 Then
                    iNum2 = 4
                '5ײ݂̏ꍇ
                Else
                    iNum2 = 5
                End If
            End If

            dSum = 0
            For iCntX = 0 To iNum2
                dSum = dSum + iLDL(iCntX, iCntY)
            Next
            dAveLDL(iCntY) = dSum / (iCntX - 1)
        '*** UPDATE �� Y.SIMIZU 2005/10/7 4.5ײݑΉ�
            
        Next
    
        '��AVE�����߂�
        dSum = 0
        For iCntY = iNoZero + 1 To iPoint
            dSum = dSum + dAveLDL(iCntY)
        Next
        
        'L/DL�v�Z
        tJ015.MSRSLDL = RoundUp((dSum / (iPoint - iNoZero)) * 10, 0)
    End If

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        tJ015.MSZEROMN = 0     ' �A��0MIN
        tJ015.MSZEROMX = 0     ' �A��0MAX
        Erase dLDLSum
        Erase dLDLZero
        iZeroCnt = 0    ' ZERO�J�E���^
        iLDLZeroCnt = 0
    
        '' ����_�ʂ̃��C�����v�����߂�
        For iCntY = 1 To 15
            For iCntX = 1 To 3   '�ꗗ��
                dLDLSum(iCntY) = dLDLSum(iCntY) + iLDL(iCntX, iCntY)
            Next iCntX
            
            If dLDLSum(iCntY) = 0# Then
                iZeroCnt = iZeroCnt + 1
            End If
            If (dLDLSum(iCntY) <> 0# Or iCntY = 15) _
               And iZeroCnt > 0 Then
               iLDLZeroCnt = iLDLZeroCnt + 1
               dLDLZero(iLDLZeroCnt) = iZeroCnt
               iZeroCnt = 0
            End If
        Next iCntY
        
        For iCntY = 1 To iLDLZeroCnt
            If dLDLZero(iCntY) > tJ015.MSZEROMX Or iCntY = 1 Then
                tJ015.MSZEROMX = dLDLZero(iCntY)
            End If
            If dLDLZero(iCntY) < tJ015.MSZEROMN Or iCntY = 1 Then
                tJ015.MSZEROMN = dLDLZero(iCntY)
            End If
        Next iCntY
        'Center��0�ȊO�̏ꍇ�A�ŏ��l��0�ɂ���
        If dLDLSum(1) <> 0 Then
            tJ015.MSZEROMN = 0
        Else
            tJ015.MSZEROMN = dLDLZero(1)
        End If
        
'        ' GD���C�����Ɋ֌W�Ȃ��A1�`3���C���Ŕ��肷��
'        iZeroCnt = 0    ' ZERO�J�E���^
'        For iCntY = 1 To 15 '�ꗗ�c
'            '���͔͈͂̏ꍇ�̂ݎ擾
'            If dLDLSum(iCntY) = 0# Then
'                iZeroCnt = iZeroCnt + 1
'            End If
'            If (dLDLSum(iCntY) <> 0# Or iCntY = 15) _
'               And iZeroCnt > 0 Then
'                If iZeroCnt = 1 Then iZeroCnt = 0   ' 0��1�̏ꍇ�A�A��0�Ƃ���
'
'                If iZeroCnt > tJ015.MSZEROMX Then
'                    tJ015.MSZEROMX = iZeroCnt
'                End If
'
'                If tJ015.MSZEROMN = -1 Then
'                    tJ015.MSZEROMN = tJ015.MSZEROMX
'                End If
'
'                If iZeroCnt < tJ015.MSZEROMN Then
'                    tJ015.MSZEROMN = iZeroCnt
'                End If
'
'                iZeroCnt = 0
'            End If
'        Next iCntY
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

End Function

'�T�v      :GD����l�A����l�����
'���Ұ�    :�ϐ���        ,IO ,�^                        :����
'          :tJ015         ,I  ,typ_TBCMJ015              :GD���э\����
'          :iTnum         ,O  ,Integer                   :����l��(3or5)
'          :bTflg         ,O  ,Boolean                   :����lDVD2�����׸�(True:�L�AFalse:��)
'          :iTden()       ,O  ,Integer                   :����lDen(5,15)
'          :iTldl()       ,O  ,Integer                   :����lL/DL(5,15)
'          :iTdvd2()      ,O  ,Integer                   :����lDVD2(5)
'          :�߂�l        ,O  ,FUNCTION_RETURN           :���� = FUNCTION_RETURN_SUCCESS : OK
'                                                               FUNCTION_RETURN_FAILURE : NG
'����      :
'����      :05/01/31 ooba
'*** UPDATE �� Y.SIMIZU 2005/10/7 �d�l��GDײݐ����g�p����
'Private Function CalcGD_DataSet(tGDdata As typ_TBCMJ015, iTnum As Integer, bTflg As Boolean, _
'                                    iTden() As Integer, iTldl() As Integer, iTdvd2() As Integer) _
'                                                                                As FUNCTION_RETURN
'    '�����l�Ƃ��đ���l��=3���
'    iTnum = 3
Private Function CalcGD_DataSet(tGDdata As typ_TBCMJ015, _
                                    iTden() As Integer, iTldl() As Integer, iTdvd2() As Integer) _
                                                                                As FUNCTION_RETURN
'*** UPDATE �� Y.SIMIZU 2005/10/7 �d�l��GDײݐ����g�p����

'*** UPDATE �� Y.SIMIZU 2005/10/7 DVD2���v�Z���邩���׸ނ�GD�̎d�l���痧�Ă�悤�ɕύX
'    bTflg = False
'*** UPDATE �� Y.SIMIZU 2005/10/7 DVD2���v�Z���邩���׸ނ�GD�̎d�l���痧�Ă�悤�ɕύX

    'GD����l���
    With tGDdata
        iTden(1, 1) = .MS01DEN1         '����l01 Den1
        iTden(2, 1) = .MS01DEN2         '����l01 Den2
        iTden(3, 1) = .MS01DEN3         '����l01 Den3
        iTden(4, 1) = .MS01DEN4         '����l01 Den4
        iTden(5, 1) = .MS01DEN5         '����l01 Den5
        iTden(1, 2) = .MS02DEN1         '����l02 Den1
        iTden(2, 2) = .MS02DEN2         '����l02 Den2
        iTden(3, 2) = .MS02DEN3         '����l02 Den3
        iTden(4, 2) = .MS02DEN4         '����l02 Den4
        iTden(5, 2) = .MS02DEN5         '����l02 Den5
        iTden(1, 3) = .MS03DEN1         '����l03 Den1
        iTden(2, 3) = .MS03DEN2         '����l03 Den2
        iTden(3, 3) = .MS03DEN3         '����l03 Den3
        iTden(4, 3) = .MS03DEN4         '����l03 Den4
        iTden(5, 3) = .MS03DEN5         '����l03 Den5
        iTden(1, 4) = .MS04DEN1         '����l04 Den1
        iTden(2, 4) = .MS04DEN2         '����l04 Den2
        iTden(3, 4) = .MS04DEN3         '����l04 Den3
        iTden(4, 4) = .MS04DEN4         '����l04 Den4
        iTden(5, 4) = .MS04DEN5         '����l04 Den5
        iTden(1, 5) = .MS05DEN1         '����l05 Den1
        iTden(2, 5) = .MS05DEN2         '����l05 Den2
        iTden(3, 5) = .MS05DEN3         '����l05 Den3
        iTden(4, 5) = .MS05DEN4         '����l05 Den4
        iTden(5, 5) = .MS05DEN5         '����l05 Den5
        iTden(1, 6) = .MS06DEN1         '����l06 Den1
        iTden(2, 6) = .MS06DEN2         '����l06 Den2
        iTden(3, 6) = .MS06DEN3         '����l06 Den3
        iTden(4, 6) = .MS06DEN4         '����l06 Den4
        iTden(5, 6) = .MS06DEN5         '����l06 Den5
        iTden(1, 7) = .MS07DEN1         '����l07 Den1
        iTden(2, 7) = .MS07DEN2         '����l07 Den2
        iTden(3, 7) = .MS07DEN3         '����l07 Den3
        iTden(4, 7) = .MS07DEN4         '����l07 Den4
        iTden(5, 7) = .MS07DEN5         '����l07 Den5
        iTden(1, 8) = .MS08DEN1         '����l08 Den1
        iTden(2, 8) = .MS08DEN2         '����l08 Den2
        iTden(3, 8) = .MS08DEN3         '����l08 Den3
        iTden(4, 8) = .MS08DEN4         '����l08 Den4
        iTden(5, 8) = .MS08DEN5         '����l08 Den5
        iTden(1, 9) = .MS09DEN1         '����l09 Den1
        iTden(2, 9) = .MS09DEN2         '����l09 Den2
        iTden(3, 9) = .MS09DEN3         '����l09 Den3
        iTden(4, 9) = .MS09DEN4         '����l09 Den4
        iTden(5, 9) = .MS09DEN5         '����l09 Den5
        iTden(1, 10) = .MS10DEN1        '����l10 Den1
        iTden(2, 10) = .MS10DEN2        '����l10 Den2
        iTden(3, 10) = .MS10DEN3        '����l10 Den3
        iTden(4, 10) = .MS10DEN4        '����l10 Den4
        iTden(5, 10) = .MS10DEN5        '����l10 Den5
        iTden(1, 11) = .MS11DEN1        '����l11 Den1
        iTden(2, 11) = .MS11DEN2        '����l11 Den2
        iTden(3, 11) = .MS11DEN3        '����l11 Den3
        iTden(4, 11) = .MS11DEN4        '����l11 Den4
        iTden(5, 11) = .MS11DEN5        '����l11 Den5
        iTden(1, 12) = .MS12DEN1        '����l12 Den1
        iTden(2, 12) = .MS12DEN2        '����l12 Den2
        iTden(3, 12) = .MS12DEN3        '����l12 Den3
        iTden(4, 12) = .MS12DEN4        '����l12 Den4
        iTden(5, 12) = .MS12DEN5        '����l12 Den5
        iTden(1, 13) = .MS13DEN1        '����l13 Den1
        iTden(2, 13) = .MS13DEN2        '����l13 Den2
        iTden(3, 13) = .MS13DEN3        '����l13 Den3
        iTden(4, 13) = .MS13DEN4        '����l13 Den4
        iTden(5, 13) = .MS13DEN5        '����l13 Den5
        iTden(1, 14) = .MS14DEN1        '����l14 Den1
        iTden(2, 14) = .MS14DEN2        '����l14 Den2
        iTden(3, 14) = .MS14DEN3        '����l14 Den3
        iTden(4, 14) = .MS14DEN4        '����l14 Den4
        iTden(5, 14) = .MS14DEN5        '����l14 Den5
        iTden(1, 15) = .MS15DEN1        '����l15 Den1
        iTden(2, 15) = .MS15DEN2        '����l15 Den2
        iTden(3, 15) = .MS15DEN3        '����l15 Den3
        iTden(4, 15) = .MS15DEN4        '����l15 Den4
        iTden(5, 15) = .MS15DEN5        '����l15 Den5
        
        iTldl(1, 1) = .MS01LDL1         '����l01 L/DL1
        iTldl(2, 1) = .MS01LDL2         '����l01 L/DL2
        iTldl(3, 1) = .MS01LDL3         '����l01 L/DL3
        iTldl(4, 1) = .MS01LDL4         '����l01 L/DL4
        iTldl(5, 1) = .MS01LDL5         '����l01 L/DL5
        iTldl(1, 2) = .MS02LDL1         '����l02 L/DL1
        iTldl(2, 2) = .MS02LDL2         '����l02 L/DL2
        iTldl(3, 2) = .MS02LDL3         '����l02 L/DL3
        iTldl(4, 2) = .MS02LDL4         '����l02 L/DL4
        iTldl(5, 2) = .MS02LDL5         '����l02 L/DL5
        iTldl(1, 3) = .MS03LDL1         '����l03 L/DL1
        iTldl(2, 3) = .MS03LDL2         '����l03 L/DL2
        iTldl(3, 3) = .MS03LDL3         '����l03 L/DL3
        iTldl(4, 3) = .MS03LDL4         '����l03 L/DL4
        iTldl(5, 3) = .MS03LDL5         '����l03 L/DL5
        iTldl(1, 4) = .MS04LDL1         '����l04 L/DL1
        iTldl(2, 4) = .MS04LDL2         '����l04 L/DL2
        iTldl(3, 4) = .MS04LDL3         '����l04 L/DL3
        iTldl(4, 4) = .MS04LDL4         '����l04 L/DL4
        iTldl(5, 4) = .MS04LDL5         '����l04 L/DL5
        iTldl(1, 5) = .MS05LDL1         '����l05 L/DL1
        iTldl(2, 5) = .MS05LDL2         '����l05 L/DL2
        iTldl(3, 5) = .MS05LDL3         '����l05 L/DL3
        iTldl(4, 5) = .MS05LDL4         '����l05 L/DL4
        iTldl(5, 5) = .MS05LDL5         '����l05 L/DL5
        iTldl(1, 6) = .MS06LDL1         '����l06 L/DL1
        iTldl(2, 6) = .MS06LDL2         '����l06 L/DL2
        iTldl(3, 6) = .MS06LDL3         '����l06 L/DL3
        iTldl(4, 6) = .MS06LDL4         '����l06 L/DL4
        iTldl(5, 6) = .MS06LDL5         '����l06 L/DL5
        iTldl(1, 7) = .MS07LDL1         '����l07 L/DL1
        iTldl(2, 7) = .MS07LDL2         '����l07 L/DL2
        iTldl(3, 7) = .MS07LDL3         '����l07 L/DL3
        iTldl(4, 7) = .MS07LDL4         '����l07 L/DL4
        iTldl(5, 7) = .MS07LDL5         '����l07 L/DL5
        iTldl(1, 8) = .MS08LDL1         '����l08 L/DL1
        iTldl(2, 8) = .MS08LDL2         '����l08 L/DL2
        iTldl(3, 8) = .MS08LDL3         '����l08 L/DL3
        iTldl(4, 8) = .MS08LDL4         '����l08 L/DL4
        iTldl(5, 8) = .MS08LDL5         '����l08 L/DL5
        iTldl(1, 9) = .MS09LDL1         '����l09 L/DL1
        iTldl(2, 9) = .MS09LDL2         '����l09 L/DL2
        iTldl(3, 9) = .MS09LDL3         '����l09 L/DL3
        iTldl(4, 9) = .MS09LDL4         '����l09 L/DL4
        iTldl(5, 9) = .MS09LDL5         '����l09 L/DL5
        iTldl(1, 10) = .MS10LDL1        '����l10 L/DL1
        iTldl(2, 10) = .MS10LDL2        '����l10 L/DL2
        iTldl(3, 10) = .MS10LDL3        '����l10 L/DL3
        iTldl(4, 10) = .MS10LDL4        '����l10 L/DL4
        iTldl(5, 10) = .MS10LDL5        '����l10 L/DL5
        iTldl(1, 11) = .MS11LDL1        '����l11 L/DL1
        iTldl(2, 11) = .MS11LDL2        '����l11 L/DL2
        iTldl(3, 11) = .MS11LDL3        '����l11 L/DL3
        iTldl(4, 11) = .MS11LDL4        '����l11 L/DL4
        iTldl(5, 11) = .MS11LDL5        '����l11 L/DL5
        iTldl(1, 12) = .MS12LDL1        '����l12 L/DL1
        iTldl(2, 12) = .MS12LDL2        '����l12 L/DL2
        iTldl(3, 12) = .MS12LDL3        '����l12 L/DL3
        iTldl(4, 12) = .MS12LDL4        '����l12 L/DL4
        iTldl(5, 12) = .MS12LDL5        '����l12 L/DL5
        iTldl(1, 13) = .MS13LDL1        '����l13 L/DL1
        iTldl(2, 13) = .MS13LDL2        '����l13 L/DL2
        iTldl(3, 13) = .MS13LDL3        '����l13 L/DL3
        iTldl(4, 13) = .MS13LDL4        '����l13 L/DL4
        iTldl(5, 13) = .MS13LDL5        '����l13 L/DL5
        iTldl(1, 14) = .MS14LDL1        '����l14 L/DL1
        iTldl(2, 14) = .MS14LDL2        '����l14 L/DL2
        iTldl(3, 14) = .MS14LDL3        '����l14 L/DL3
        iTldl(4, 14) = .MS14LDL4        '����l14 L/DL4
        iTldl(5, 14) = .MS14LDL5        '����l14 L/DL5
        iTldl(1, 15) = .MS15LDL1        '����l15 L/DL1
        iTldl(2, 15) = .MS15LDL2        '����l15 L/DL2
        iTldl(3, 15) = .MS15LDL3        '����l15 L/DL3
        iTldl(4, 15) = .MS15LDL4        '����l15 L/DL4
        iTldl(5, 15) = .MS15LDL5        '����l15 L/DL5
        
        iTdvd2(1) = .MS01DVD2           '����l01 DVD2
        iTdvd2(2) = .MS02DVD2           '����l02 DVD2
        iTdvd2(3) = .MS03DVD2           '����l03 DVD2
        iTdvd2(4) = .MS04DVD2           '����l04 DVD2
        iTdvd2(5) = .MS05DVD2           '����l05 DVD2
    End With
    
'*** UPDATE �� Y.SIMIZU 2005/10/7 ײݐ����̑���l������ChkGD_Data�ōs��
'    '����lDVD2��������
'    For iCnt = 1 To 5
'        If iTdvd2(iCnt) <> -1 Then
'            bTflg = True
'            Exit For
'        End If
'    Next
    
'    '����l������
'    For iCnt = 1 To 15
'        '����l��=5
'        If iTden(5, iCnt) <> -1 Or iTldl(5, iCnt) <> -1 Then
'            iTnum = 5
'            Exit For
'        End If
'    Next
'*** UPDATE �� Y.SIMIZU 2005/10/7 ײݐ����̑���l������ChkGD_Data�ōs��
    
    CalcGD_DataSet = FUNCTION_RETURN_SUCCESS
    
End Function

'�T�v      :�d�l��GDײݐ�������l�����݂��邩����������
'���Ұ�    :�ϐ���          ,IO ,�^                 :����
'          :tGDdata         ,I  ,typ_TBCMJ015      :GD���э\����
'          :WFGD            ,O  ,W_GD              :GD�d�l�\����
'          :bDenChk         ,O  ,Boolean           :Den�����ް������׸ށ@05/10/25 ooba
'          :bLdlChk         ,O  ,Boolean           :L/DL�����ް������׸ށ@05/10/25 ooba
'          :bDvd2Chk        ,O  ,Boolean           :DVD2�����ް������׸ށ@05/10/25 ooba
'          :�߂�l          ,O  ,FUNCTION_RETURN    :���� = FUNCTION_RETURN_SUCCESS : OK
'                                                               FUNCTION_RETURN_FAILURE : NG
'����      :
'����      :05/10/07 Y.SIMIZU
Private Function ChkGD_Data(tGDdata As typ_TBCMJ015, WFGD As W_GD, _
                            bDenChk As Boolean, bLdlChk As Boolean, bDvd2Chk As Boolean) _
                            As FUNCTION_RETURN
    Dim iCnt        As Integer
    Dim iPoint      As Integer
    Dim iLine       As Integer
    Dim iTden(5, 15)     As Integer
    Dim iTldl(5, 15)     As Integer
    Dim iTdvd2(5)    As Integer
    
    'GD����l���
    With tGDdata
        iTden(1, 1) = .MS01DEN1         '����l01 Den1
        iTden(2, 1) = .MS01DEN2         '����l01 Den2
        iTden(3, 1) = .MS01DEN3         '����l01 Den3
        iTden(4, 1) = .MS01DEN4         '����l01 Den4
        iTden(5, 1) = .MS01DEN5         '����l01 Den5
        iTden(1, 2) = .MS02DEN1         '����l02 Den1
        iTden(2, 2) = .MS02DEN2         '����l02 Den2
        iTden(3, 2) = .MS02DEN3         '����l02 Den3
        iTden(4, 2) = .MS02DEN4         '����l02 Den4
        iTden(5, 2) = .MS02DEN5         '����l02 Den5
        iTden(1, 3) = .MS03DEN1         '����l03 Den1
        iTden(2, 3) = .MS03DEN2         '����l03 Den2
        iTden(3, 3) = .MS03DEN3         '����l03 Den3
        iTden(4, 3) = .MS03DEN4         '����l03 Den4
        iTden(5, 3) = .MS03DEN5         '����l03 Den5
        iTden(1, 4) = .MS04DEN1         '����l04 Den1
        iTden(2, 4) = .MS04DEN2         '����l04 Den2
        iTden(3, 4) = .MS04DEN3         '����l04 Den3
        iTden(4, 4) = .MS04DEN4         '����l04 Den4
        iTden(5, 4) = .MS04DEN5         '����l04 Den5
        iTden(1, 5) = .MS05DEN1         '����l05 Den1
        iTden(2, 5) = .MS05DEN2         '����l05 Den2
        iTden(3, 5) = .MS05DEN3         '����l05 Den3
        iTden(4, 5) = .MS05DEN4         '����l05 Den4
        iTden(5, 5) = .MS05DEN5         '����l05 Den5
        iTden(1, 6) = .MS06DEN1         '����l06 Den1
        iTden(2, 6) = .MS06DEN2         '����l06 Den2
        iTden(3, 6) = .MS06DEN3         '����l06 Den3
        iTden(4, 6) = .MS06DEN4         '����l06 Den4
        iTden(5, 6) = .MS06DEN5         '����l06 Den5
        iTden(1, 7) = .MS07DEN1         '����l07 Den1
        iTden(2, 7) = .MS07DEN2         '����l07 Den2
        iTden(3, 7) = .MS07DEN3         '����l07 Den3
        iTden(4, 7) = .MS07DEN4         '����l07 Den4
        iTden(5, 7) = .MS07DEN5         '����l07 Den5
        iTden(1, 8) = .MS08DEN1         '����l08 Den1
        iTden(2, 8) = .MS08DEN2         '����l08 Den2
        iTden(3, 8) = .MS08DEN3         '����l08 Den3
        iTden(4, 8) = .MS08DEN4         '����l08 Den4
        iTden(5, 8) = .MS08DEN5         '����l08 Den5
        iTden(1, 9) = .MS09DEN1         '����l09 Den1
        iTden(2, 9) = .MS09DEN2         '����l09 Den2
        iTden(3, 9) = .MS09DEN3         '����l09 Den3
        iTden(4, 9) = .MS09DEN4         '����l09 Den4
        iTden(5, 9) = .MS09DEN5         '����l09 Den5
        iTden(1, 10) = .MS10DEN1        '����l10 Den1
        iTden(2, 10) = .MS10DEN2        '����l10 Den2
        iTden(3, 10) = .MS10DEN3        '����l10 Den3
        iTden(4, 10) = .MS10DEN4        '����l10 Den4
        iTden(5, 10) = .MS10DEN5        '����l10 Den5
        iTden(1, 11) = .MS11DEN1        '����l11 Den1
        iTden(2, 11) = .MS11DEN2        '����l11 Den2
        iTden(3, 11) = .MS11DEN3        '����l11 Den3
        iTden(4, 11) = .MS11DEN4        '����l11 Den4
        iTden(5, 11) = .MS11DEN5        '����l11 Den5
        iTden(1, 12) = .MS12DEN1        '����l12 Den1
        iTden(2, 12) = .MS12DEN2        '����l12 Den2
        iTden(3, 12) = .MS12DEN3        '����l12 Den3
        iTden(4, 12) = .MS12DEN4        '����l12 Den4
        iTden(5, 12) = .MS12DEN5        '����l12 Den5
        iTden(1, 13) = .MS13DEN1        '����l13 Den1
        iTden(2, 13) = .MS13DEN2        '����l13 Den2
        iTden(3, 13) = .MS13DEN3        '����l13 Den3
        iTden(4, 13) = .MS13DEN4        '����l13 Den4
        iTden(5, 13) = .MS13DEN5        '����l13 Den5
        iTden(1, 14) = .MS14DEN1        '����l14 Den1
        iTden(2, 14) = .MS14DEN2        '����l14 Den2
        iTden(3, 14) = .MS14DEN3        '����l14 Den3
        iTden(4, 14) = .MS14DEN4        '����l14 Den4
        iTden(5, 14) = .MS14DEN5        '����l14 Den5
        iTden(1, 15) = .MS15DEN1        '����l15 Den1
        iTden(2, 15) = .MS15DEN2        '����l15 Den2
        iTden(3, 15) = .MS15DEN3        '����l15 Den3
        iTden(4, 15) = .MS15DEN4        '����l15 Den4
        iTden(5, 15) = .MS15DEN5        '����l15 Den5
        
        iTldl(1, 1) = .MS01LDL1         '����l01 L/DL1
        iTldl(2, 1) = .MS01LDL2         '����l01 L/DL2
        iTldl(3, 1) = .MS01LDL3         '����l01 L/DL3
        iTldl(4, 1) = .MS01LDL4         '����l01 L/DL4
        iTldl(5, 1) = .MS01LDL5         '����l01 L/DL5
        iTldl(1, 2) = .MS02LDL1         '����l02 L/DL1
        iTldl(2, 2) = .MS02LDL2         '����l02 L/DL2
        iTldl(3, 2) = .MS02LDL3         '����l02 L/DL3
        iTldl(4, 2) = .MS02LDL4         '����l02 L/DL4
        iTldl(5, 2) = .MS02LDL5         '����l02 L/DL5
        iTldl(1, 3) = .MS03LDL1         '����l03 L/DL1
        iTldl(2, 3) = .MS03LDL2         '����l03 L/DL2
        iTldl(3, 3) = .MS03LDL3         '����l03 L/DL3
        iTldl(4, 3) = .MS03LDL4         '����l03 L/DL4
        iTldl(5, 3) = .MS03LDL5         '����l03 L/DL5
        iTldl(1, 4) = .MS04LDL1         '����l04 L/DL1
        iTldl(2, 4) = .MS04LDL2         '����l04 L/DL2
        iTldl(3, 4) = .MS04LDL3         '����l04 L/DL3
        iTldl(4, 4) = .MS04LDL4         '����l04 L/DL4
        iTldl(5, 4) = .MS04LDL5         '����l04 L/DL5
        iTldl(1, 5) = .MS05LDL1         '����l05 L/DL1
        iTldl(2, 5) = .MS05LDL2         '����l05 L/DL2
        iTldl(3, 5) = .MS05LDL3         '����l05 L/DL3
        iTldl(4, 5) = .MS05LDL4         '����l05 L/DL4
        iTldl(5, 5) = .MS05LDL5         '����l05 L/DL5
        iTldl(1, 6) = .MS06LDL1         '����l06 L/DL1
        iTldl(2, 6) = .MS06LDL2         '����l06 L/DL2
        iTldl(3, 6) = .MS06LDL3         '����l06 L/DL3
        iTldl(4, 6) = .MS06LDL4         '����l06 L/DL4
        iTldl(5, 6) = .MS06LDL5         '����l06 L/DL5
        iTldl(1, 7) = .MS07LDL1         '����l07 L/DL1
        iTldl(2, 7) = .MS07LDL2         '����l07 L/DL2
        iTldl(3, 7) = .MS07LDL3         '����l07 L/DL3
        iTldl(4, 7) = .MS07LDL4         '����l07 L/DL4
        iTldl(5, 7) = .MS07LDL5         '����l07 L/DL5
        iTldl(1, 8) = .MS08LDL1         '����l08 L/DL1
        iTldl(2, 8) = .MS08LDL2         '����l08 L/DL2
        iTldl(3, 8) = .MS08LDL3         '����l08 L/DL3
        iTldl(4, 8) = .MS08LDL4         '����l08 L/DL4
        iTldl(5, 8) = .MS08LDL5         '����l08 L/DL5
        iTldl(1, 9) = .MS09LDL1         '����l09 L/DL1
        iTldl(2, 9) = .MS09LDL2         '����l09 L/DL2
        iTldl(3, 9) = .MS09LDL3         '����l09 L/DL3
        iTldl(4, 9) = .MS09LDL4         '����l09 L/DL4
        iTldl(5, 9) = .MS09LDL5         '����l09 L/DL5
        iTldl(1, 10) = .MS10LDL1        '����l10 L/DL1
        iTldl(2, 10) = .MS10LDL2        '����l10 L/DL2
        iTldl(3, 10) = .MS10LDL3        '����l10 L/DL3
        iTldl(4, 10) = .MS10LDL4        '����l10 L/DL4
        iTldl(5, 10) = .MS10LDL5        '����l10 L/DL5
        iTldl(1, 11) = .MS11LDL1        '����l11 L/DL1
        iTldl(2, 11) = .MS11LDL2        '����l11 L/DL2
        iTldl(3, 11) = .MS11LDL3        '����l11 L/DL3
        iTldl(4, 11) = .MS11LDL4        '����l11 L/DL4
        iTldl(5, 11) = .MS11LDL5        '����l11 L/DL5
        iTldl(1, 12) = .MS12LDL1        '����l12 L/DL1
        iTldl(2, 12) = .MS12LDL2        '����l12 L/DL2
        iTldl(3, 12) = .MS12LDL3        '����l12 L/DL3
        iTldl(4, 12) = .MS12LDL4        '����l12 L/DL4
        iTldl(5, 12) = .MS12LDL5        '����l12 L/DL5
        iTldl(1, 13) = .MS13LDL1        '����l13 L/DL1
        iTldl(2, 13) = .MS13LDL2        '����l13 L/DL2
        iTldl(3, 13) = .MS13LDL3        '����l13 L/DL3
        iTldl(4, 13) = .MS13LDL4        '����l13 L/DL4
        iTldl(5, 13) = .MS13LDL5        '����l13 L/DL5
        iTldl(1, 14) = .MS14LDL1        '����l14 L/DL1
        iTldl(2, 14) = .MS14LDL2        '����l14 L/DL2
        iTldl(3, 14) = .MS14LDL3        '����l14 L/DL3
        iTldl(4, 14) = .MS14LDL4        '����l14 L/DL4
        iTldl(5, 14) = .MS14LDL5        '����l14 L/DL5
        iTldl(1, 15) = .MS15LDL1        '����l15 L/DL1
        iTldl(2, 15) = .MS15LDL2        '����l15 L/DL2
        iTldl(3, 15) = .MS15LDL3        '����l15 L/DL3
        iTldl(4, 15) = .MS15LDL4        '����l15 L/DL4
        iTldl(5, 15) = .MS15LDL5        '����l15 L/DL5
        
        iTdvd2(1) = .MS01DVD2           '����l01 DVD2
        iTdvd2(2) = .MS02DVD2           '����l02 DVD2
        iTdvd2(3) = .MS03DVD2           '����l03 DVD2
        iTdvd2(4) = .MS04DVD2           '����l04 DVD2
        iTdvd2(5) = .MS05DVD2           '����l05 DVD2
    End With
    
    
    'GD�����ް����������@05/10/25 ooba START ===================================>
    bDenChk = False
    bLdlChk = False
    bDvd2Chk = False
    
    For iCnt = 1 To 5
        If iTdvd2(iCnt) <> -1 Then bDvd2Chk = True
        For iPoint = 1 To 15
            If iTden(iCnt, iPoint) <> -1 Then bDenChk = True
            If iTldl(iCnt, iPoint) <> -1 Then bLdlChk = True
        Next iPoint
    Next iCnt
    If bDenChk Then bDvd2Chk = True
    'GD�����ް����������@05/10/25 ooba END =====================================>
    
    
    'Den�̎d�l�������L��,�ۏؗL��̏ꍇ
    If WFGD.JudgFlagDen = "1" And WFGD.GuaranteeDen.cJudg = JudgCodeC01 Then
        'Den�̑���l��ײݐ������邩������
        For iPoint = 1 To 15
            '����_7�܂�
            If iPoint <= 7 Then
                '�d�l��3ײ݂̏ꍇ
                If WFGD.SpecGdLine = 3 Then
                    iLine = 3
                '�d�l��4.5ײݖ���5ײ݂̏ꍇ
                Else
                    iLine = 5
                End If
            '����_8����
            Else
                '�d�l��3ײ݂̏ꍇ
                If WFGD.SpecGdLine = 3 Then
                    iLine = 3
                '�d�l��4.5ײ݂̏ꍇ
                ElseIf WFGD.SpecGdLine = 4.5 Then
                    iLine = 4
                '�d�l��5ײ݂̏ꍇ
                Else
                    iLine = 5
                End If
            End If
            
            For iCnt = 1 To iLine
                'DEN�̑���l���Ȃ��ꍇ
                If iTden(iCnt, iPoint) = -1 Then
                    ChkGD_Data = FUNCTION_RETURN_FAILURE
                    '����װ(�����𔲂���)
                    Exit Function
                End If
            Next iCnt
        Next iPoint
    End If
    
    'LDL�̎d�l�������L��,�ۏؗL��̏ꍇ
    If WFGD.JudgFlagLdl = "1" And WFGD.GuaranteeLdl.cJudg = JudgCodeC01 Then
        'L/DL�̑���l��ײݐ������邩������
        For iPoint = 1 To 15
            '����_7�܂�
            If iPoint <= 7 Then
                '�d�l��3ײ݂̏ꍇ
                If WFGD.SpecGdLine = 3 Then
                    iLine = 3
                '�d�l��4.5ײݖ���5ײ݂̏ꍇ
                Else
                    iLine = 5
                End If
            '����_8����
            Else
                '�d�l��3ײ݂̏ꍇ
                If WFGD.SpecGdLine = 3 Then
                    iLine = 3
                '�d�l��4.5ײ݂̏ꍇ
                ElseIf WFGD.SpecGdLine = 4.5 Then
                    iLine = 4
                '�d�l��5ײ݂̏ꍇ
                Else
                    iLine = 5
                End If
            End If
            
            For iCnt = 1 To iLine
                'L/DL�̑���l���Ȃ��ꍇ
                If iTldl(iCnt, iPoint) = -1 Then
                    ChkGD_Data = FUNCTION_RETURN_FAILURE
                    '����װ(�����𔲂���)
                    Exit Function
                End If
            Next iCnt
        Next iPoint
    End If
            
    ChkGD_Data = FUNCTION_RETURN_SUCCESS
    
End Function

'�T�v      :�R�[�h���擾
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :code          ,   ,Variant      ,
'          :CodeData      ,   ,typ_CodeMaster ,
'          :�߂�l        ,O  ,String       ,
'����      :�R�[�h��񃊃X�g����Y���R�[�h�̏����擾����
'����      :
Private Function Search_CrCode(strCode As String, typ_CodeData() As typ_TBCMB005) As String
    Dim i As Integer
    
    '���X�g����Y���R�[�h�̏��P������
    For i = 1 To UBound(typ_CodeData)
        If strCode = Trim(typ_CodeData(i).CODE) Then
            Search_CrCode = typ_CodeData(i).INFO1
            Exit Function
        End If
    Next
    Search_CrCode = ""
End Function

Public Function NtoS(strWk As String) As String
    If Mid(strWk, 1, 1) = Chr(0) Then
        NtoS = " "
        Exit Function
    End If
    NtoS = strWk
End Function

Public Function NtoZ2(strWk As String) As Double
    If Trim(strWk) = "" Then
        NtoZ2 = -1
        Exit Function
    End If
    NtoZ2 = CDbl(strWk)
End Function

Public Sub BMDDataSet(BmdNo As Integer, UpDo As Integer, INGOTPOS As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)

    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4               '�����w��
    Dim typ_y013z       As typ_TBCMY013
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim SijiUmu         As String
    Dim WFBmSokuP       As String                   ' ����ʒu�Q�_
    Dim WFBMD           As Integer                  '2001/12/19 S.Sano
    Dim sSxlPos         As String                   'SXL�ʒu(TOP/BOT)�@04/04/12 ooba

    '�����w���ݒ�
    IND = IIf(UpDo = SxlTop, "123", "123")
    
'Chg Start 2011/03/09 SMPK Miyata
'    sSxlPos = IIf(UpDo = SxlTop, "TOP", "BOT")      '04/04/12 ooba
    Select Case UpDo
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/09 SMPK Miyata

    With typ_CType

        Select Case BmdNo
        Case 1
            WFBMD = WFBMD1 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.B1
            '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B1 And CheckKHN(.typ_si.HWFBM1KN, 7, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B1 And CheckKHN(.typ_si.HWFBM1KN, 7, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.B1 And .typ_si.MSMPFLGWFBM = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B1 And .typ_si.MSMPFLGWFBM = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B1"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDB1CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESB1CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFBMD1)
            WFBmSokuP = .typ_si.HWFBM1ST
        Case 2
            WFBMD = WFBMD2 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.B2
            '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B2 And CheckKHN(.typ_si.HWFBM2KN, 8, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B2 And CheckKHN(.typ_si.HWFBM2KN, 8, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.B2 And .typ_si.MSMPFLGWFBM = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B2 And .typ_si.MSMPFLGWFBM = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B2"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDB2CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESB2CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFBMD2)
            WFBmSokuP = .typ_si.HWFBM2ST
        Case 3
            WFBMD = WFBMD3 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.B3
            '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B3 And CheckKHN(.typ_si.HWFBM3KN, 9, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B3 And CheckKHN(.typ_si.HWFBM3KN, 9, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.B3 And .typ_si.MSMPFLGWFBM = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B3 And .typ_si.MSMPFLGWFBM = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B3"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDB3CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESB3CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFBMD3)
            WFBmSokuP = .typ_si.HWFBM3ST
        End Select
            typ_y013z = .typ_y013(UpDo, WFBMD) '2001/12/19 S.Sano
    
    
        '' WF�����w���iB1)*****************************************************************
        If JudgSpecCode Then
            '��ʕ\�����e�ݒ�
            .typ_rslt(UpDo, DispLineCount).pos = -1                                             ' �������J�n�ʒu
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())              ' ���e
            .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                                     ' ���P
            .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                                     ' ���Q
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                           ' ���R
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                           ' ���S
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '5�Ԗڂ̏��FAN���x��ǉ�
            typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
            typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
            typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
            typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
            typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
            typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
            typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                         ' ���5
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
            typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                         ' ���6
            typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                         ' ���7
            typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                         ' ���8
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------

            .typ_rslt(UpDo, DispLineCount).SMPLID = -1                                          ' �T���v���m��
            .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                          ' ���茋��
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                                   ' �i��(12��)
            bJudg = False
            If shiji <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                                 ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                                 ' ���R
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' �T���v���m��
                    
                'BMD1����
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'BMD1���莸�s
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                            ' ���Q
                    'BMD1����
                    If WfCrBmdJudg(.typ_si, typ_y013z, bJudg, BmdNo) Then
'                        '��ʕ\�����e�ݒ�
'                        vTemp = CVar(typ_y013z.MESDATA5)
'                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���1
'                        vTemp = CVar(typ_y013z.MESDATA6)
'                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' ���2
'                        vTemp = CVar(typ_y013z.MESDATA8)
'                        .typ_rslt(UpDo, DispLineCount).INFO3 = ""                               ' ���R
'                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' ���S

                        '��ʕ\�����e�ݒ�@�@2003/05/20 ooba
                        vTemp = CVar(typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���1
                        vTemp = CVar(typ_y013z.MESDATA7)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' ���2
                        vTemp = CVar(typ_y013z.MESDATA8)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' ���R
                        vTemp = CVar(typ_y013z.MESDATA9)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.0")    ' ���S
                        JiltusekiUmu(UpDo, WFBMD) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                        '5�Ԗڂ̏��FAN���x��ǉ�
                        vTemp = CVar(typ_y013z.DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    bJudg = False
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����وُ�"                          ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                End If
                
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' ���茋��
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                      ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If UpDo = SxlTop Or UpDo = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case BmdNo
                Case 1
                    gsTbcmy028ErrCode = "00136"
                Case 2
                    gsTbcmy028ErrCode = "00137"
                Case 3
                    gsTbcmy028ErrCode = "00138"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())          ' ���e
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                                 ' ���P
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                                 ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                                 ' ���R
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                       ' ���S
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                     ' ���5
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                     ' ���6
                typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                     ' ���7
                typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                     ' ���8
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' �T���v���m��
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' ���茋��
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                               ' �i��(12��)
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'BMD1���莸�s
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                            ' ���Q
                    'BMD1����
                    If WfCrBmdJudg(.typ_si, typ_y013z, bJudg, BmdNo) Then
'                        '��ʕ\�����e�ݒ�
'                        vTemp = CVar(typ_y013z.MESDATA5)
'                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���1
'                        vTemp = CVar(typ_y013z.MESDATA6)
'                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' ���2
'                        vTemp = CVar(typ_y013z.MESDATA8)
'                        .typ_rslt(UpDo, DispLineCount).INFO3 = ""                               ' ���R
'                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' ���S

                        '��ʕ\�����e�ݒ�@�@2003/05/20 ooba
                        vTemp = CVar(typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���1
                        vTemp = CVar(typ_y013z.MESDATA7)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' ���2
                        vTemp = CVar(typ_y013z.MESDATA8)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' ���R
                        vTemp = CVar(typ_y013z.MESDATA9)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.0")    ' ���S
                        JiltusekiUmu(UpDo, WFBMD) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                        vTemp = CVar(typ_y013z.DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    bJudg = False
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����وُ�"                          ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                End If
                
                'Add Start 2011/11/28 Y.Hitomi ���Ԕ����̏ꍇ�́A�Q�l�\������
                If sSxlPos = "MID" And bJudg = False Then
                    If BmdNo = 1 And JudgSW.B1 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    ElseIf BmdNo = 2 And JudgSW.B2 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    ElseIf BmdNo = 3 And JudgSW.B3 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    End If
                End If
                'Add End  2011/11/28 Y.Hitomi
    
                DispLineCount = DispLineCount + 1
            End If
        End If
    
        Select Case BmdNo
        Case 1
            .typ_y013(UpDo, WFBMD1) = typ_y013z
        Case 2
            .typ_y013(UpDo, WFBMD2) = typ_y013z
        Case 3
            .typ_y013(UpDo, WFBMD3) = typ_y013z
        End Select
    
    End With
    
End Sub

Public Sub OSFDataSet(OsfNo As Integer, UpDo As Integer, INGOTPOS As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)

    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                   '�����w��
    Dim typ_y013z       As typ_TBCMY013
    Dim AveMax(1)       As String                       '����/�ő唻��l�@2003/05/20 ooba
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim SijiUmu         As String
    Dim WFBmSokuP       As String                       ' ����ʒu�Q�_
    Dim WFBmSokuHou     As String                       ' �iWFOSF1����ʒu_��
    Dim WFBmSokuRyou    As String                       ' �iWFOSF1����ʒu_��
    Dim WFOSF           As Integer                      '2001/12/19 S.Sano
    Dim sSxlPos         As String                       'SXL�ʒu(TOP/BOT)�@04/04/12 ooba
    
    '�����w���ݒ�
    IND = IIf(UpDo = SxlTop, "123", "123")
        
'Chg Start 2011/03/09 SMPK Miyata
'    sSxlPos = IIf(UpDo = SxlTop, "TOP", "BOT")      '04/04/12 ooba
    Select Case UpDo
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/09 SMPK Miyata
    With typ_CType
        Select Case OsfNo
        Case 1
            WFOSF = WFOSF1 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.L1
            '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L1 And CheckKHN(.typ_si.HWFOF1KN, 3, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L1 And CheckKHN(.typ_si.HWFOF1KN, 3, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.L1 And .typ_si.MSMPFLGWFOF = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L1 And .typ_si.MSMPFLGWFOF = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata

            SCC = "L1"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDL1CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESL1CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFOSF1)
            WFBmSokuHou = .typ_si.HWFOF1SH
            WFBmSokuP = .typ_si.HWFOF1ST
            WFBmSokuRyou = .typ_si.HWFOF1SR
        Case 2
            WFOSF = WFOSF2 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.L2
            '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L2 And CheckKHN(.typ_si.HWFOF2KN, 4, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L2 And CheckKHN(.typ_si.HWFOF2KN, 4, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.L2 And .typ_si.MSMPFLGWFOF = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L2 And .typ_si.MSMPFLGWFOF = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L2"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDL2CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESL2CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFOSF2)
            WFBmSokuHou = .typ_si.HWFOF2SH
            WFBmSokuP = .typ_si.HWFOF2ST
            WFBmSokuRyou = .typ_si.HWFOF2SR
        Case 3
            WFOSF = WFOSF3 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.L3
            '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L3 And CheckKHN(.typ_si.HWFOF3KN, 5, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L3 And CheckKHN(.typ_si.HWFOF3KN, 5, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.L3 And .typ_si.MSMPFLGWFOF = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L3 And .typ_si.MSMPFLGWFOF = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L3"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDL3CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESL3CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFOSF3)
            WFBmSokuHou = .typ_si.HWFOF3SH
            WFBmSokuP = .typ_si.HWFOF3ST
            WFBmSokuRyou = .typ_si.HWFOF3SR
        Case 4
            WFOSF = WFOSF4 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.L4
            '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L4 And CheckKHN(.typ_si.HWFOF4KN, 6, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L4 And CheckKHN(.typ_si.HWFOF4KN, 6, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.L4 And .typ_si.MSMPFLGWFOF = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L4 And .typ_si.MSMPFLGWFOF = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L4"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDL4CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESL4CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFOSF4)
            WFBmSokuHou = .typ_si.HWFOF4SH
            WFBmSokuP = .typ_si.HWFOF4ST
            WFBmSokuRyou = .typ_si.HWFOF4SR
        End Select
        typ_y013z = .typ_y013(UpDo, WFOSF)
        
        
        '' WF�����w���iL1)*****************************************************************
        If JudgSpecCode Then
            '��ʕ\�����e�ݒ�
            .typ_rslt(UpDo, DispLineCount).pos = -1                                             ' �������J�n�ʒu
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())              ' ���e
            .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                                     ' ���P
            .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                                     ' ���Q
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                           ' ���R
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                           ' ���S
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
            typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
            typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
            typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
            typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
            typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
            typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                         ' ���5
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
            typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                         ' ���6
            typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                         ' ���7
            typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                         ' ���8
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(UpDo, DispLineCount).SMPLID = -1                                          ' �T���v���m��
            .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                          ' ���茋��
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                                   ' �i��(12��)
            bJudg = False
            If shiji <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                                 ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                                 ' ���R
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' �T���v���m��
                'OSF����擾
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'OSF���莸�s
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                            ' ���Q
                    'OSF1����擾
                    If WfCrOsfJudg(.typ_si, typ_y013z, bJudg, OsfNo, AveMax()) Then             ' AveMax�ǉ��@2003/05/20 ooba
'                        '��ʕ\�����e�ݒ�
'                        vTemp = CVar(typ_y013z.MESDATA7)
'                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���P
'                        vTemp = CVar(typ_y013z.MESDATA8)
'                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' ���Q
'                        vTemp = CVar(typ_y013z.MESDATA6)
'                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' ���R
'                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' ���S
                        
                        '��ʕ\�����e�ݒ�@�@2003/05/21 ooba
                        vTemp = CVar(AveMax(0))
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���P
                        vTemp = CVar(AveMax(1))
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' ���Q
                        vTemp = CVar(typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' ���R
                        vTemp = CVar(IIf(Trim(typ_y013z.MESDATA9) = "", "-", Trim(typ_y013z.MESDATA9)) _
                                        & "  " & IIf(Trim(typ_y013z.MESDATA12) = "", "-", Trim(typ_y013z.MESDATA12)) _
                                        & "  " & IIf(Trim(typ_y013z.MESDATA15) = "", "-", Trim(typ_y013z.MESDATA15)))
                        .typ_rslt(UpDo, DispLineCount).INFO4 = vTemp                            ' ���S
                        
                        JiltusekiUmu(UpDo, WFOSF) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                        vTemp = CVar(typ_y013z.DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����وُ�"                          ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                End If
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' ���茋��
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                      ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If UpDo = SxlTop Or UpDo = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case OsfNo
                Case 1
                    gsTbcmy028ErrCode = "00132"
                Case 2
                    gsTbcmy028ErrCode = "00133"
                Case 3
                    gsTbcmy028ErrCode = "00134"
                Case 4
                    gsTbcmy028ErrCode = "00135"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())          ' ���e
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                                 ' ���P
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                                 ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                                 ' ���R
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                       ' ���S
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                     ' ���5
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                     ' ���6
                typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                     ' ���7
                typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                     ' ���8
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' �T���v���m��
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' ���茋��
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                               ' �i��(12��)
                'OSF����擾
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'OSF���莸�s
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                            ' ���Q
                    'OSF����擾
                    If WfCrOsfJudg(.typ_si, typ_y013z, bJudg, OsfNo, AveMax()) Then             ' AveMax�ǉ��@2003/05/20 ooba
'                        '��ʕ\�����e�ݒ�
'                        vTemp = CVar(typ_y013z.MESDATA7)
'                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���P
'                        vTemp = CVar(typ_y013z.MESDATA8)
'                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' ���Q
'                        vTemp = CVar(typ_y013z.MESDATA6)
'                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' ���R
'                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' ���S
                        
                        '��ʕ\�����e�ݒ�@�@2003/05/21 ooba
                        vTemp = CVar(AveMax(0))
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���P
                        vTemp = CVar(AveMax(1))
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' ���Q
                        vTemp = CVar(typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' ���R
                        vTemp = CVar(IIf(Trim(typ_y013z.MESDATA9) = "", "-", Trim(typ_y013z.MESDATA9)) _
                                        & "  " & IIf(Trim(typ_y013z.MESDATA12) = "", "-", Trim(typ_y013z.MESDATA12)) _
                                        & "  " & IIf(Trim(typ_y013z.MESDATA15) = "", "-", Trim(typ_y013z.MESDATA15)))
                        .typ_rslt(UpDo, DispLineCount).INFO4 = vTemp                            ' ���S
                         JiltusekiUmu(UpDo, WFOSF) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                        vTemp = CVar(typ_y013z.DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����وُ�"                          ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                End If
                'Add Start 2011/11/28 Y.Hitomi ���Ԕ����̏ꍇ�́A�Q�l�\������
                If sSxlPos = "MID" And bJudg = False Then
                    If OsfNo = 1 And JudgSW.L1 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    ElseIf OsfNo = 2 And JudgSW.L2 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    ElseIf OsfNo = 3 And JudgSW.L3 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    ElseIf OsfNo = 4 And JudgSW.L4 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    End If
                End If
                'Add End  2011/11/28 Y.Hitomi
                
                DispLineCount = DispLineCount + 1
            End If
        End If
    
        Select Case OsfNo
        Case 1
            .typ_y013(UpDo, WFOSF1) = typ_y013z
            TmpOsfData(0, UpDo, WFOSF1) = AveMax(0)                                                 '�@��2003/05/20 ooba
            TmpOsfData(1, UpDo, WFOSF1) = AveMax(1)
            TmpOsfMBNP(0, UpDo, WFOSF1) = IIf(typ_y013z.MESDATA9 = "-", " ", typ_y013z.MESDATA9)
            TmpOsfMBNP(1, UpDo, WFOSF1) = IIf(typ_y013z.MESDATA12 = "-", " ", typ_y013z.MESDATA12)
            TmpOsfMBNP(2, UpDo, WFOSF1) = IIf(typ_y013z.MESDATA15 = "-", " ", typ_y013z.MESDATA15)  '�@��2003/05/21 ooba
        Case 2
            .typ_y013(UpDo, WFOSF2) = typ_y013z
            TmpOsfData(0, UpDo, WFOSF2) = AveMax(0)                                                 '�@��2003/05/20 ooba
            TmpOsfData(1, UpDo, WFOSF2) = AveMax(1)
            TmpOsfMBNP(0, UpDo, WFOSF2) = IIf(typ_y013z.MESDATA9 = "-", " ", typ_y013z.MESDATA9)
            TmpOsfMBNP(1, UpDo, WFOSF2) = IIf(typ_y013z.MESDATA12 = "-", " ", typ_y013z.MESDATA12)
            TmpOsfMBNP(2, UpDo, WFOSF2) = IIf(typ_y013z.MESDATA15 = "-", " ", typ_y013z.MESDATA15)  '�@��2003/05/21 ooba
        Case 3
            .typ_y013(UpDo, WFOSF3) = typ_y013z
            TmpOsfData(0, UpDo, WFOSF3) = AveMax(0)                                                 '�@��2003/05/20 ooba
            TmpOsfData(1, UpDo, WFOSF3) = AveMax(1)
            TmpOsfMBNP(0, UpDo, WFOSF3) = IIf(typ_y013z.MESDATA9 = "-", " ", typ_y013z.MESDATA9)
            TmpOsfMBNP(1, UpDo, WFOSF3) = IIf(typ_y013z.MESDATA12 = "-", " ", typ_y013z.MESDATA12)
            TmpOsfMBNP(2, UpDo, WFOSF3) = IIf(typ_y013z.MESDATA15 = "-", " ", typ_y013z.MESDATA15)  '�@��2003/05/21 ooba
        Case 4
            .typ_y013(UpDo, WFOSF4) = typ_y013z
            TmpOsfData(0, UpDo, WFOSF4) = AveMax(0)                                                 '�@��2003/05/20 ooba
            TmpOsfData(1, UpDo, WFOSF4) = AveMax(1)
            TmpOsfMBNP(0, UpDo, WFOSF4) = IIf(typ_y013z.MESDATA9 = "-", " ", typ_y013z.MESDATA9)
            TmpOsfMBNP(1, UpDo, WFOSF4) = IIf(typ_y013z.MESDATA12 = "-", " ", typ_y013z.MESDATA12)
            TmpOsfMBNP(2, UpDo, WFOSF4) = IIf(typ_y013z.MESDATA15 = "-", " ", typ_y013z.MESDATA15)  '�@��2003/05/21 ooba
        End Select
    
    End With
End Sub

Public Sub DOIDataSet(DoiNo As Integer, UpDo As Integer, INGOTPOS As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)

    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4               '�����w��
    Dim typ_y013z       As typ_TBCMY013
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim SijiUmu         As String
    Dim WFBmSokuP       As String                   ' ����ʒu�Q�_
    Dim WFDOI           As Integer                  '2001/12/19 S.Sano
    Dim sSxlPos         As String                   'SXL�ʒu(TOP/BOT)�@04/04/12 ooba
    
    '�����w���ݒ�
'    IND = IIf(UpDo = SxlTop, "12346", "123")
    IND = IIf(UpDo = SxlTop, "123", "123")
        
'Chg Start 2011/03/09 SMPK Miyata
'    sSxlPos = IIf(UpDo = SxlTop, "TOP", "BOT")      '04/04/12 ooba
    Select Case UpDo
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/09 SMPK Miyata
    
    With typ_CType
        Select Case DoiNo
        Case 1
            WFDOI = WFDOI1 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.Doi1
            '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.Doi1 And CheckKHN(.typ_si.HWFOS1KN, 10, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.Doi1 And CheckKHN(.typ_si.HWFOS1KN, 10, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.Doi1 And .typ_si.MSMPFLGWFDOI = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.Doi1 And .typ_si.MSMPFLGWFDOI = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "DO1"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDDO1CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESDO1CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFDOI1)
        Case 2
            WFDOI = WFDOI2 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.Doi2
            '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.Doi2 And CheckKHN(.typ_si.HWFOS2KN, 11, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.Doi2 And CheckKHN(.typ_si.HWFOS2KN, 11, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.Doi2 And .typ_si.MSMPFLGWFDOI = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.Doi2 And .typ_si.MSMPFLGWFDOI = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "DO2"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDDO2CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESDO2CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFDOI2)
        Case 3
            WFDOI = WFDOI3 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.Doi3
            '�ۏؕ��@�����ǉ��@04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.Doi3 And CheckKHN(.typ_si.HWFOS3KN, 12, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.Doi3 And CheckKHN(.typ_si.HWFOS3KN, 12, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.Doi3 And .typ_si.MSMPFLGWFDOI = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.Doi3 And .typ_si.MSMPFLGWFDOI = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "DO3"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDDO3CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESDO3CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFDOI3)
        End Select
        typ_y013z = .typ_y013(UpDo, WFDOI)

        
        '' WF�����w���iDOI)*****************************************************************
        If JudgSpecCode Then
            '��ʕ\�����e�ݒ�
            .typ_rslt(UpDo, DispLineCount).pos = -1                                             ' �������J�n�ʒu
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())              ' ���e
            .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                                     ' ���P
            .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                                     ' ���Q
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                           ' ���R
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                           ' ���S
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
            typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
            typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
            typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
            typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
            typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                         ' ���5
        '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
            typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                         ' ���6
            typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                         ' ���7
            typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                         ' ���8
        '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(UpDo, DispLineCount).SMPLID = -1                                          ' �T���v���m��
            .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                          ' ���茋��
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                                   ' �i��(12��)
            bJudg = False
            If shiji <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                                 ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                                 ' ���R
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' �T���v���m��
                'DOI����擾
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'DOI���莸�s
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                            ' ���Q
                    'DOI����擾
                    If WfCrDoiJudg(.typ_si, typ_y013z, bJudg, DoiNo) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(typ_y013z.MESDATA1 - typ_y013z.MESDATA4)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.0")    ' ���1
                        vTemp = CVar(typ_y013z.MESDATA2 - typ_y013z.MESDATA5)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' ���2
                        vTemp = CVar(typ_y013z.MESDATA3 - typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.0")    ' ���3
                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' ���S
                        JiltusekiUmu(UpDo, WFDOI) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                        vTemp = CVar(typ_y013z.DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����وُ�"                          ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                End If
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' ���茋��
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                      ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If UpDo = SxlTop Or UpDo = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case DoiNo
                Case 1
                    gsTbcmy028ErrCode = "00139"
                Case 2
                    gsTbcmy028ErrCode = "00140"
                Case 3
                    gsTbcmy028ErrCode = "00141"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())          ' ���e
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                                 ' ���P
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                                 ' ���Q
                .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                                 ' ���R
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                       ' ���S
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                     ' ���5
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            '����6(PUA�l)�A7(PUA%�l)�A8(STD�l)�ǉ��ɂ��ύX
                typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                     ' ���6
                typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                     ' ���7
                typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                     ' ���8
            '���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' �T���v���m��
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' ���茋��
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                               ' �i��(12��)
                'DOI����擾
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'DOI���莸�s
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                            ' ���Q
                    'DOI����擾
                    If WfCrDoiJudg(.typ_si, typ_y013z, bJudg, DoiNo) Then
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(typ_y013z.MESDATA1 - typ_y013z.MESDATA4)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.0")    ' ���1
                        vTemp = CVar(typ_y013z.MESDATA2 - typ_y013z.MESDATA5)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' ���2
                        vTemp = CVar(typ_y013z.MESDATA3 - typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.0")    ' ���3
                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' ���S
                        JiltusekiUmu(UpDo, WFDOI) = True '2001/12/19 S.Sano
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                        vTemp = CVar(typ_y013z.DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' ���5
                    '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����وُ�"                          ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                End If
                'Add Start 2011/11/28 Y.Hitomi ���Ԕ����̏ꍇ�́A�Q�l�\������
                If sSxlPos = "MID" And bJudg = False Then
                    If DoiNo = 1 And JudgSW.Doi1 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    ElseIf DoiNo = 2 And JudgSW.Doi2 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    ElseIf DoiNo = 3 And JudgSW.Doi3 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    End If
                End If
                'Add End  2011/11/28 Y.Hitomi
                DispLineCount = DispLineCount + 1
            End If
        End If
    End With
End Sub
'''''============================================================================================================================
'''''
''''''�T�v      :�T���v���h�c�̎擾
''''''���Ұ��@�@:�ϐ����@�@�@�@,IO ,�^       ,����
''''''�@�@      :iWafPos      ,I  ,Integer�@,�����w���e�[�u���ʒu
''''''�@�@      :sSampID1     ,I  ,String �@,�T���v���h�c�P
''''''�@�@      :sSampID2     ,I  ,String �@,�T���v���h�c�Q
''''''�@�@      :�߂�l�@�@�@�@,O  ,Boolean�@,�I���̗L��
''''''����      :�����w���T���v���̃T���v���h�c���擾����
''''''����      :2001/07/11�@ �쐬
''''''           2003/04/05   hitec)matsumoto bKyotuFlg�ǉ�
'''''Public Function GetSampleID(iWafPos As Integer, sSampID1 As String, sSampID2 As String, _
'''''                                                     Optional iKubun As Integer) As Boolean
'''''
'''''    Dim bBot As Boolean
'''''    Dim bTop As Boolean
'''''    Dim bBlk As Boolean
'''''    Dim TargetBlkPos As Integer
'''''    Dim p As Integer
'''''    Dim m As Integer
'''''    Dim i As Integer
'''''
'''''    Dim iHinbanRow  As Integer
'''''    Dim vUpHinban   As Variant
'''''
'''''
'''''    bBot = False
'''''    bTop = False
'''''    bBlk = False
'''''    p = iWafPos
'''''    With tblWafInd(iWafPos)
'''''        m = UBound(tblBlkInf)
''''''        For i = 1 To m
''''''            If .IngotPos = tblBlkInf(i).COF.TOPSMPLPOS Or _
''''''               i = m And .IngotPos = tblBlkInf(i).COF.BOTSMPLPOS Then
''''''                bBlk = True
''''''                Exit For
''''''            End If
''''''        Next i
'''''        For i = 1 To UBound(tblBlkInf)
'''''            If tblWafInd(p).BLOCKID = tblBlkInf(i).BLOCKID Then
'''''                TargetBlkPos = i
'''''                Exit For
'''''            End If
'''''        Next
'''''
'''''        bBot = False
'''''        bTop = False
'''''        Call GetSampleBT(.SMP.CRYINDRS, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDOI, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDB1, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDB2, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDB3, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDL1, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDL2, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDL3, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDL4, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDDS, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDDZ, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDSP, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDD1, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDD2, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDD3, bTop, bBot)
''''''=========================================2003/04/16 okazaki
''''''�㉺�i�Ԃ�Z
'''''        If iWafPos >= 1 Then
'''''            If Trim(tblWafInd(iWafPos).HINDN.hinban) = "Z" Or _
'''''               Trim(tblWafInd(iWafPos).HINUP.hinban) = "Z" Or _
'''''               iKubun = 3 Then      '�u���b�N���ς�鏉���\���s
'''''                    bTop = True
'''''                    bBot = True
'''''                    bBlk = False
'''''                    '�`�F�b�N�{�b�N�X�ǉ��ɂ��T���v���ؑւ̔����ǉ� (iWafPos - 1���g�p���邽�ߋ敪�R�̂Ƃ��̂ݔ��肷��) 2003/06/01 okazaki
'''''                    If Trim(tblWafInd(iWafPos).HINDN.hinban) <> "Z" And tblWafInd(iWafPos).HINDN.hinban = tblWafInd(iWafPos - 1).HINDN.hinban Then
'''''                        bTop = False
'''''                        bBot = False
'''''                        bBlk = False
'''''                    End If
'''''                    '2003/06/01 end
'''''            End If
'''''        End If
''''''=========================================2003/04/16 end
'''''        '' ������^�������T���v���i�ʁj
'''''        If bTop = True And bBot = True Then
'''''            If bBlk = True Then
'''''                If .BlockPos = 0 Then
'''''                    If tblBlkInf(TargetBlkPos - 1).NOWPROC = PROCD_WFC_SOUGOUHANTEI Then
'''''                        sSampID1 = Right(tblBlkInf(TargetBlkPos - 1).BLOCKID, 3) & "-" & GetWafPos(tblBlkInf(TargetBlkPos - 1).LENGTH) & "B"
'''''                        sSampID2 = Right(.BLOCKID, 3) & "-000T"
'''''                    Else
'''''                        sSampID1 = Right(.BLOCKID, 3) & "-000T"
'''''                        sSampID2 = ""
'''''                    End If
'''''                Else
'''''                    sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
'''''                    sSampID2 = Right(tblBlkInf(TargetBlkPos + 1).BLOCKID, 3) & "-000T"
'''''                End If
'''''            Else
'''''                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "U"
'''''                sSampID2 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "D"
'''''            End If
'''''            GetSampleID = False
'''''        '' �������T���v��
'''''        ElseIf bTop = True And bBot = False Then
'''''            If bBlk = True Then
'''''                sSampID1 = Right(.BLOCKID, 3) & "-000T"
'''''            Else
'''''                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "D"
'''''            End If
'''''            sSampID2 = ""
'''''            GetSampleID = False
'''''        '' ������T���v��
'''''        ElseIf bTop = False And bBot = True Then
'''''            If bBlk = True Then
'''''                If .BlockPos = 0 Then
'''''                    sSampID1 = Right(tblBlkInf(TargetBlkPos - 1).BLOCKID, 3) & "-" & GetWafPos(tblBlkInf(TargetBlkPos - 1).LENGTH) & "B"
'''''                Else
'''''                    sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
'''''                End If
'''''            Else
'''''                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "U"
'''''            End If
'''''            sSampID2 = ""
'''''            GetSampleID = False
'''''        '' ������^�������T���v���i���ʁj
'''''        ElseIf bTop = False And bBot = False Then
'''''            If bBlk = True Then
'''''                If .BlockPos = 0 Then
'''''                    If tblBlkInf(TargetBlkPos).NOWPROC = PROCD_WFC_SOUGOUHANTEI Then
'''''                        sSampID1 = Right(tblBlkInf(TargetBlkPos).BLOCKID, 3) & "-" & GetWafPos(tblBlkInf(TargetBlkPos - 1).LENGTH) & "B"
'''''                        sSampID2 = Right(.BLOCKID, 3) & "-000T"
'''''                    Else
'''''                        sSampID1 = Right(.BLOCKID, 3) & "-000T"
'''''                        sSampID2 = ""
'''''                        GetSampleID = False
'''''                        Exit Function
'''''                    End If
'''''                Else
'''''                    If tblBlkInf(TargetBlkPos + 1).NOWPROC = PROCD_WFC_SOUGOUHANTEI Then
'''''                        sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
'''''                        sSampID2 = Right(tblBlkInf(TargetBlkPos + 1).BLOCKID, 3) & "-000T"
'''''                    Else
'''''                        sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
'''''                        sSampID2 = ""
'''''                        GetSampleID = False
'''''                        Exit Function
'''''                    End If
'''''                End If
'''''            Else
'''''                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "U"
'''''                sSampID2 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "D"
'''''            End If
'''''            GetSampleID = True
'''''        End If
'''''    End With
'''''
'''''End Function
'''''============================================================================================================================
'''''
''''''�T�v      :�T���v���̃g�b�v���^�{�g�����敪�̎擾
''''''���Ұ��@�@:�ϐ����@�@�@�@,IO ,�^       ,����
''''''�@�@      :sSample      ,I  ,String �@,�T���v��
''''''�@�@      :bTop         ,O  ,Boolean�@,�g�b�v���敪�̗L��
''''''�@�@      :bBot         ,O  ,Boolean�@,�{�g�����敪�̗L��
''''''����      :�T���v���敪�̗L����Ԃ�
''''''����      :2001/07/11�@ �쐬
'''''Public Sub GetSampleBT(ByVal sSample As String, bTop As Boolean, bBot As Boolean)
'''''
'''''    Select Case sSample
'''''    Case "1"
'''''        bTop = True
'''''    Case "2"
'''''        bBot = True
'''''    Case "4"
'''''        bTop = True
'''''        bBot = True
'''''    End Select
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''�T�v      :�����E�F�n�[�e�[�u���̍쐬
''''''���Ұ��@�@:�ϐ����@�@�@�@  ,IO ,�^              ,����
''''''�@�@      :tblBlkInf      ,I  ,typ_BlkInf3   �@,�u���b�N�Ǘ��\����
''''''�@�@      :tmpLackWaf     ,I  ,typ_LackWaf�@   ,�������
''''''�@�@      :BlkInfPos      ,I  ,Integer       �@,�������S�̂̃u���b�N���ɑ΂���Ώۃu���b�N�̊J�n�ʒu
''''''�@�@      :BlkCnt         ,I  ,Integer�@       ,�Ώۃu���b�N��
''''''�@�@      :RftblLackMap ,O  ,typ_LackMap     �@,�E�F�n�[�e�[�u���\����
''''''����      :
''''''����      :
'''''Public Function LackMapMake(tblBlkInf() As typ_BlkInf3, tmpLackWaf() As typ_LackWaf, BlkInfPos As Integer, BlkCnt As Integer) As FUNCTION_RETURN
'''''
'''''    Dim bFlag As Boolean
'''''    Dim p As Integer
'''''    Dim m As Integer
'''''    Dim n As Integer
'''''    Dim i As Integer
'''''    Dim j As Integer
'''''    Dim k As Integer
'''''
'''''    '' �����E�F�n�[�e�[�u���̍쐬
'''''    k = 0
'''''    m = BlkCnt + BlkInfPos - 1
'''''    n = UBound(tmpLackWaf)
'''''    ReDim tblLackMap(n)
'''''
'''''    '' �u���b�N�̎n�܂肩��
'''''    For i = BlkInfPos To m
'''''        DoEvents
'''''        For j = 1 To n
'''''            DoEvents
'''''            If tblBlkInf(i).BLOCKID = tmpLackWaf(j).BLOCKID Then
'''''                If bFlag = False Then
'''''                    k = k + 1
'''''                    tblLackMap(k).BLOCKID = tmpLackWaf(j).BLOCKID
'''''                    p = tmpLackWaf(j).WAFERNO
'''''                    If p = -1 Then
'''''                        tblLackMap(k).LACKPOSS = 0
'''''                        tblLackMap(k).LACKCNTS = -1
'''''                        tblLackMap(k).LACKPOSE = tblBlkInf(i).REALLEN
'''''                        tblLackMap(k).LACKCNTE = -1
'''''                        Exit For
'''''                    End If
'''''                    tblLackMap(k).LACKPOSS = tmpLackWaf(j).TOP_POS
'''''                    tblLackMap(k).LACKCNTS = tmpLackWaf(j).WAFERNO
'''''                    bFlag = True
'''''                Else
'''''                    If tmpLackWaf(j).WAFERNO = p + 1 Then
'''''                        p = p + 1
'''''                        If bFlag = True And j = n Then
'''''                            tblLackMap(k).LACKPOSE = tmpLackWaf(j).TAIL_POS
'''''                            tblLackMap(k).LACKCNTE = tmpLackWaf(j).WAFERNO
'''''                        End If
'''''                    Else
'''''                        tblLackMap(k).LACKPOSE = tmpLackWaf(j - 1).TAIL_POS
'''''                        tblLackMap(k).LACKCNTE = tmpLackWaf(j - 1).WAFERNO
'''''                        k = k + 1
'''''                        tblLackMap(k).BLOCKID = tmpLackWaf(j).BLOCKID
'''''                        tblLackMap(k).LACKPOSS = tmpLackWaf(j).TOP_POS
'''''                        tblLackMap(k).LACKCNTS = tmpLackWaf(j).WAFERNO
'''''                        p = tmpLackWaf(j).WAFERNO
'''''                    End If
'''''                End If
'''''            Else
'''''                If bFlag = True Then
'''''                    tblLackMap(k).LACKPOSE = tmpLackWaf(j - 1).TAIL_POS
'''''                    tblLackMap(k).LACKCNTE = tmpLackWaf(j - 1).WAFERNO
'''''                    bFlag = False
'''''                    Exit For
'''''                End If
'''''            End If
'''''        Next j
'''''    Next i
'''''    ReDim Preserve tblLackMap(k)
'''''
'''''    For i = 1 To k
'''''        With tblLackMap(i)
'''''            If .LACKPOSS > 0 And .LACKPOSE = 0 Then
'''''                .LACKPOSE = .LACKPOSS
'''''            End If
'''''            If .LACKCNTS > 0 And .LACKCNTE = 0 Then
'''''                .LACKCNTE = .LACKCNTS
'''''            End If
'''''        End With
'''''    Next
'''''
'''''End Function
'''''============================================================================================================================
'''''
'''''Public Function NoTestCheck(lblMsg As Label) As FUNCTION_RETURN
'''''    Dim c0 As Long
'''''
'''''
'''''    Dim HIN(1) As tFullHinban
'''''    Dim Inf(1) As NoTest_Info
'''''
'''''    NoTestCheck = FUNCTION_RETURN_FAILURE
'''''    For c0 = 1 To 2
'''''        '���i�ԃZ�b�g
'''''        HIN(0).factory = tblTotal.typ_Param.factory
'''''        HIN(0).hinban = tblTotal.typ_Param.hinban
'''''        HIN(0).mnorevno = tblTotal.typ_Param.REVNUM
'''''        HIN(0).opecond = tblTotal.typ_Param.opecond
'''''        '�U�֐�i�ԃZ�b�g
'''''        If c0 = 1 Then
'''''            HIN(1) = tblWafInd(1).HINDN
'''''        Else
'''''            HIN(1) = tblWafInd(UBound(tblWafInd())).HINUP
'''''        End If
'''''        If Trim(HIN(1).hinban) = "Z" Then
'''''            Exit For
'''''        End If
'''''        If DBDRV_GetNoTestHinInfo(HIN(), Inf()) = FUNCTION_RETURN_FAILURE Then
'''''            Exit Function
'''''        End If
'''''
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESRS <> "1") Then
'''''        '���і�
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Res.HWFRHWYS = "X") Or (Inf(1).Res.HWFRHWYS = "S") Then
'''''            If (Inf(1).Res.HWFRHWYS = "H") Or (Inf(1).Res.HWFRHWYS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " RES���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESOI <> "1") Then
'''''        '���і�
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Oi.HWFONHWS = "X") Or (Inf(1).Oi.HWFONHWS = "S") Then
'''''            If (Inf(1).Oi.HWFONHWS = "H") Or (Inf(1).Oi.HWFONHWS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OI���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESB1 = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).BMD(0).HWFBMxHS = "X") Or (Inf(1).BMD(0).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(0).HWFBMxHS = "H") Or (Inf(1).BMD(0).HWFBMxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).BMD(0).HWFBMxET <> Inf(1).BMD(0).HWFBMxET Or _
'''''                   Inf(0).BMD(0).HWFBMxNS <> Inf(1).BMD(0).HWFBMxNS Or _
'''''                   Inf(0).BMD(0).HWFBMxSH <> Inf(1).BMD(0).HWFBMxSH Or _
'''''                   Inf(0).BMD(0).HWFBMxSR <> Inf(1).BMD(0).HWFBMxSR Or _
'''''                   Inf(0).BMD(0).HWFBMxST <> Inf(1).BMD(0).HWFBMxST Or _
'''''                   Inf(0).BMD(0).HWFBMxSZ <> Inf(1).BMD(0).HWFBMxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD1")  '03/06/06
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
'''''        '���і�
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).BMD(0).HWFBMxHS = "X") Or (Inf(1).BMD(0).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(0).HWFBMxHS = "H") Or (Inf(1).BMD(0).HWFBMxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD1���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESB2 = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).BMD(1).HWFBMxHS = "X") Or (Inf(1).BMD(1).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(1).HWFBMxHS = "H") Or (Inf(1).BMD(1).HWFBMxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).BMD(1).HWFBMxET <> Inf(1).BMD(1).HWFBMxET Or _
'''''                   Inf(0).BMD(1).HWFBMxNS <> Inf(1).BMD(1).HWFBMxNS Or _
'''''                   Inf(0).BMD(1).HWFBMxSH <> Inf(1).BMD(1).HWFBMxSH Or _
'''''                   Inf(0).BMD(1).HWFBMxSR <> Inf(1).BMD(1).HWFBMxSR Or _
'''''                   Inf(0).BMD(1).HWFBMxST <> Inf(1).BMD(1).HWFBMxST Or _
'''''                   Inf(0).BMD(1).HWFBMxSZ <> Inf(1).BMD(1).HWFBMxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD2")  '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).BMD(1).HWFBMxHS = "X") Or (Inf(1).BMD(1).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(1).HWFBMxHS = "H") Or (Inf(1).BMD(1).HWFBMxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD2���і�")  '03/06/06
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESB3 = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).BMD(2).HWFBMxHS = "X") Or (Inf(1).BMD(2).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(2).HWFBMxHS = "H") Or (Inf(1).BMD(2).HWFBMxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).BMD(2).HWFBMxET <> Inf(1).BMD(2).HWFBMxET Or _
'''''                   Inf(0).BMD(2).HWFBMxNS <> Inf(1).BMD(2).HWFBMxNS Or _
'''''                   Inf(0).BMD(2).HWFBMxSH <> Inf(1).BMD(2).HWFBMxSH Or _
'''''                   Inf(0).BMD(2).HWFBMxSR <> Inf(1).BMD(2).HWFBMxSR Or _
'''''                   Inf(0).BMD(2).HWFBMxST <> Inf(1).BMD(2).HWFBMxST Or _
'''''                   Inf(0).BMD(2).HWFBMxSZ <> Inf(1).BMD(2).HWFBMxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD3") '03/06/06
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).BMD(2).HWFBMxHS = "X") Or (Inf(1).BMD(2).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(2).HWFBMxHS = "H") Or (Inf(1).BMD(2).HWFBMxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD3���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESL1 = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).OSF(0).HWFOFxHS = "X") Or (Inf(1).OSF(0).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(0).HWFOFxHS = "H") Or (Inf(1).OSF(0).HWFOFxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).OSF(0).HWFOFxET <> Inf(1).OSF(0).HWFOFxET Or _
'''''                   Inf(0).OSF(0).HWFOFxNS <> Inf(1).OSF(0).HWFOFxNS Or _
'''''                   Inf(0).OSF(0).HWFOFxSH <> Inf(1).OSF(0).HWFOFxSH Or _
'''''                   Inf(0).OSF(0).HWFOFxSR <> Inf(1).OSF(0).HWFOFxSR Or _
'''''                   Inf(0).OSF(0).HWFOFxST <> Inf(1).OSF(0).HWFOFxST Or _
'''''                   Inf(0).OSF(0).HWFOFxSZ <> Inf(1).OSF(0).HWFOFxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF1")  '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).OSF(0).HWFOFxHS = "X") Or (Inf(1).OSF(0).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(0).HWFOFxHS = "H") Or (Inf(1).OSF(0).HWFOFxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF1���і�") '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESL2 = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).OSF(1).HWFOFxHS = "X") Or (Inf(1).OSF(1).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(1).HWFOFxHS = "H") Or (Inf(1).OSF(1).HWFOFxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).OSF(1).HWFOFxET <> Inf(1).OSF(1).HWFOFxET Or _
'''''                   Inf(0).OSF(1).HWFOFxNS <> Inf(1).OSF(1).HWFOFxNS Or _
'''''                   Inf(0).OSF(1).HWFOFxSH <> Inf(1).OSF(1).HWFOFxSH Or _
'''''                   Inf(0).OSF(1).HWFOFxSR <> Inf(1).OSF(1).HWFOFxSR Or _
'''''                   Inf(0).OSF(1).HWFOFxST <> Inf(1).OSF(1).HWFOFxST Or _
'''''                   Inf(0).OSF(1).HWFOFxSZ <> Inf(1).OSF(1).HWFOFxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF2") '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).OSF(1).HWFOFxHS = "X") Or (Inf(1).OSF(1).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(1).HWFOFxHS = "H") Or (Inf(1).OSF(1).HWFOFxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF2���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESL3 = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).OSF(2).HWFOFxHS = "X") Or (Inf(1).OSF(2).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(2).HWFOFxHS = "H") Or (Inf(1).OSF(2).HWFOFxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).OSF(2).HWFOFxET <> Inf(1).OSF(2).HWFOFxET Or _
'''''                   Inf(0).OSF(2).HWFOFxNS <> Inf(1).OSF(2).HWFOFxNS Or _
'''''                   Inf(0).OSF(2).HWFOFxSH <> Inf(1).OSF(2).HWFOFxSH Or _
'''''                   Inf(0).OSF(2).HWFOFxSR <> Inf(1).OSF(2).HWFOFxSR Or _
'''''                   Inf(0).OSF(2).HWFOFxST <> Inf(1).OSF(2).HWFOFxST Or _
'''''                   Inf(0).OSF(2).HWFOFxSZ <> Inf(1).OSF(2).HWFOFxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF3")  '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).OSF(2).HWFOFxHS = "X") Or (Inf(1).OSF(2).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(2).HWFOFxHS = "H") Or (Inf(1).OSF(2).HWFOFxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF3���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESL4 = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).OSF(3).HWFOFxHS = "X") Or (Inf(1).OSF(3).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(3).HWFOFxHS = "H") Or (Inf(1).OSF(3).HWFOFxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).OSF(3).HWFOFxET <> Inf(1).OSF(3).HWFOFxET Or _
'''''                   Inf(0).OSF(3).HWFOFxNS <> Inf(1).OSF(3).HWFOFxNS Or _
'''''                   Inf(0).OSF(3).HWFOFxSH <> Inf(1).OSF(3).HWFOFxSH Or _
'''''                   Inf(0).OSF(3).HWFOFxSR <> Inf(1).OSF(3).HWFOFxSR Or _
'''''                   Inf(0).OSF(3).HWFOFxST <> Inf(1).OSF(3).HWFOFxST Or _
'''''                   Inf(0).OSF(3).HWFOFxSZ <> Inf(1).OSF(3).HWFOFxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF4")   '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).OSF(3).HWFOFxHS = "X") Or (Inf(1).OSF(3).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(3).HWFOFxHS = "H") Or (Inf(1).OSF(3).HWFOFxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF4���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESDS = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Dsod.HWFDSOHS = "X") Or (Inf(1).Dsod.HWFDSOHS = "S") Then
'''''            If (Inf(1).Dsod.HWFDSOHS = "H") Or (Inf(1).Dsod.HWFDSOHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).Dsod.HWFDSOKE <> Inf(1).Dsod.HWFDSOKE Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " DSOD")  '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Dsod.HWFDSOHS = "X") Or (Inf(1).Dsod.HWFDSOHS = "S") Then
'''''            If (Inf(1).Dsod.HWFDSOHS = "H") Or (Inf(1).Dsod.HWFDSOHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " DSOD���і�")   '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESDZ = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Dz.HWFMKHWS = "X") Or (Inf(1).Dz.HWFMKHWS = "S") Then
'''''            If (Inf(1).Dz.HWFMKHWS = "H") Or (Inf(1).Dz.HWFMKHWS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).Dz.HWFMKSPH <> Inf(1).Dz.HWFMKSPH Or _
'''''                   Inf(0).Dz.HWFMKSPR <> Inf(1).Dz.HWFMKSPR Or _
'''''                   Inf(0).Dz.HWFMKSPT <> Inf(1).Dz.HWFMKSPT Or _
'''''                   Inf(0).Dz.HWFMKSZY <> Inf(1).Dz.HWFMKSZY Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " DZ")   '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Dz.HWFMKHWS = "X") Or (Inf(1).Dz.HWFMKHWS = "S") Then
'''''            If (Inf(1).Dz.HWFMKHWS = "H") Or (Inf(1).Dz.HWFMKHWS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " DZ���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESSP = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).SpvFe.HWFSPVHS = "X") Or (Inf(1).SpvFe.HWFSPVHS = "S") Then
'''''            If (Inf(1).SpvFe.HWFSPVHS = "H") Or (Inf(1).SpvFe.HWFSPVHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).SpvFe.HWFSPVSH <> Inf(1).SpvFe.HWFSPVSH Or _
'''''                   Inf(0).SpvFe.HWFSPVSI <> Inf(1).SpvFe.HWFSPVSI Or _
'''''                   Inf(0).SpvFe.HWFSPVST <> Inf(1).SpvFe.HWFSPVST Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " SPVFE")  '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).SpvFe.HWFSPVHS = "X") Or (Inf(1).SpvFe.HWFSPVHS = "S") Then
'''''            If (Inf(1).SpvFe.HWFSPVHS = "H") Or (Inf(1).SpvFe.HWFSPVHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " SPVFE���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESSP = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Spv.HWFDLHWS = "X") Or (Inf(1).Spv.HWFDLHWS = "S") Then
'''''            If (Inf(1).Spv.HWFDLHWS = "H") Or (Inf(1).Spv.HWFDLHWS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).Spv.HWFDLSPH <> Inf(1).Spv.HWFDLSPH Or _
'''''                   Inf(0).Spv.HWFDLSPI <> Inf(1).Spv.HWFDLSPI Or _
'''''                   Inf(0).Spv.HWFDLSPT <> Inf(1).Spv.HWFDLSPT Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " SPV�g�U��")  '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Spv.HWFDLHWS = "X") Or (Inf(1).Spv.HWFDLHWS = "S") Then
'''''            If (Inf(1).Spv.HWFDLHWS = "H") Or (Inf(1).Spv.HWFDLHWS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " SPV�g�U�����і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESDO1 = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Doi(0).HWFOSxHS = "X") Or (Inf(1).Doi(0).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(0).HWFOSxHS = "H") Or (Inf(1).Doi(0).HWFOSxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).Doi(0).HWFOSxNS <> Inf(1).Doi(0).HWFOSxNS Or _
'''''                   Inf(0).Doi(0).HWFOSxSH <> Inf(1).Doi(0).HWFOSxSH Or _
'''''                   Inf(0).Doi(0).HWFOSxSI <> Inf(1).Doi(0).HWFOSxSI Or _
'''''                   Inf(0).Doi(0).HWFOSxST <> Inf(1).Doi(0).HWFOSxST Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ��Oi1")  '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Doi(0).HWFOSxHS = "X") Or (Inf(1).Doi(0).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(0).HWFOSxHS = "H") Or (Inf(1).Doi(0).HWFOSxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ��Oi1���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESDO2 = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Doi(1).HWFOSxHS = "X") Or (Inf(1).Doi(1).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(1).HWFOSxHS = "H") Or (Inf(1).Doi(1).HWFOSxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).Doi(0).HWFOSxNS <> Inf(1).Doi(0).HWFOSxNS Or _
'''''                   Inf(0).Doi(0).HWFOSxSH <> Inf(1).Doi(0).HWFOSxSH Or _
'''''                   Inf(0).Doi(0).HWFOSxSI <> Inf(1).Doi(0).HWFOSxSI Or _
'''''                   Inf(0).Doi(0).HWFOSxST <> Inf(1).Doi(0).HWFOSxST Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ��Oi2")  '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Doi(1).HWFOSxHS = "X") Or (Inf(1).Doi(1).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(1).HWFOSxHS = "H") Or (Inf(1).Doi(1).HWFOSxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ��Oi2���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESDO3 = "1") Then
'''''        '���їL��
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Doi(2).HWFOSxHS = "X") Or (Inf(1).Doi(2).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(2).HWFOSxHS = "H") Or (Inf(1).Doi(2).HWFOSxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                If Inf(0).Doi(0).HWFOSxNS <> Inf(1).Doi(0).HWFOSxNS Or _
'''''                   Inf(0).Doi(0).HWFOSxSH <> Inf(1).Doi(0).HWFOSxSH Or _
'''''                   Inf(0).Doi(0).HWFOSxSI <> Inf(1).Doi(0).HWFOSxSI Or _
'''''                   Inf(0).Doi(0).HWFOSxST <> Inf(1).Doi(0).HWFOSxST Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ��Oi3")  '03/06/06 �㓡
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
''''''            If (Inf(1).Doi(2).HWFOSxHS = "X") Or (Inf(1).Doi(2).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(2).HWFOSxHS = "H") Or (Inf(1).Doi(2).HWFOSxHS = "S") Then
''''''�v�e�T���v�������ύX 2003.05.20 yakimura
'''''
'''''            '�����L��
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ��Oi3���і�")  '03/06/06 �㓡
'''''                Exit Function
'''''            End If
'''''        End If
'''''    Next
'''''
'''''    NoTestCheck = FUNCTION_RETURN_SUCCESS
'''''
'''''End Function

''�T�v      :�w��̌����ԍ��Ɋ܂܂��i�Ԃ̈ꗗ�𓾂�
''���Ұ�    :�ϐ���        ,IO ,�^          ,����
''          :cryno         ,I  ,String      ,�����ԍ�
''          :hinban()      ,O  ,tFullHinban ,�i�ԃ��X�g
''          :�߂�l        ,O  ,FUNCTION_RETURN,���o�̐���
''����      :
''����      :2001/06/27 �쐬  ���� (2002/07 s_cmzc010a.bas���ړ�)
'Public Function GetXlHinban(cryno$, HINBAN() As tFullHinban) As FUNCTION_RETURN
'Dim rs      As OraDynaset               '���oRecordDynaset
'Dim rsCnt   As Integer                  'ں��޶���
'Dim sql     As String                   'SQL��
'Dim i       As Integer                  'ٰ�߶���
'
'    '�G���[�n���h���̐ݒ�
'    On Error GoTo proc_err
''(2002/07)    gErr.Push "s_cmzc010a.bas -- Function GetXlHinban"
'    gErr.Push "-- Function GetXlHinban"
'
'    'SQL���̍쐬
'    sql = "Select CRYNUM, HINBAN, REVNUM, FACTORY, OPECOND from TBCME041 "
'    sql = sql & "Where(CRYNUM = '" & cryno & "')"
'
'    '�f�[�^�̒��o
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'    '''���o���R�[�h�����݂��Ȃ��ꍇ
'    If rs.EOF Then
'        ReDim HINBAN(0)                     '�z��̏�����
'        GetXlHinban = FUNCTION_RETURN_FAILURE   '�װ�ð��
'        GoTo proc_exit
'    End If
'
'    rsCnt = rs.RecordCount                  'ں��ސ��̶��Ă����
'    ReDim HINBAN(rsCnt - 1)                 '�z��̍Ē�`
'
'    '�z��ɒl���Z�b�g
'    rs.MoveFirst                            '�擪ں��ނɈړ�
'    For i = 0 To rsCnt - 1                  'ں��ސ���ٰ��
'        DoEvents
'        With HINBAN(i)
'            .HINBAN = rs!HINBAN             '�i��
'            .mnorevno = rs!REVNUM           '���i�ԍ������ԍ�
'            .FACTORY = rs!FACTORY           '�H��
'            .OPECOND = rs!OPECOND           '���Ə���
'        End With
'        rs.MoveNext                         '��ں��ނɈړ�
'    Next
'
'    GetXlHinban = FUNCTION_RETURN_SUCCESS   '����ð��
'
'
'proc_exit:
'    '�I��
'    gErr.Pop
'    Exit Function
'
'proc_err:
'    '�G���[�n���h��
'    gErr.HandleError
'    Resume proc_exit
'End Function
'''''============================================================================================================================
'''''
'''''Public Function DBData2DispData(data As Variant, Optional Formatstr As String) As Variant
'''''    If data = -1 Then
'''''        DBData2DispData = ""
'''''    Else
'''''        If Formatstr = "" Then
'''''            DBData2DispData = data
'''''        Else
'''''            DBData2DispData = Format(data, Formatstr)
'''''        End If
'''''    End If
'''''End Function
'''''============================================================================================================================
'''''
''''''�T�v      :SXL ID�̎擾
''''''���Ұ��@�@:�ϐ���         ,IO ,�^       ,����
''''''�@�@      :sBlockID �@�@�@,I  ,String �@,�u���b�NID
''''''�@�@      :iIngotPos�@�@�@,I  ,Integer�@,�������J�n�ʒu
''''''�@�@      :�߂�l         ,O  ,String �@,SXL ID
''''''����      :SXL ID��Ԃ�
''''''����      :2001/07/11�@��� �쐬
'''''Public Function GetSXLID(sBlockID As String, iIngotpos As Integer) As String
'''''
'''''    GetSXLID = Left(sBlockID, 10) & GetWafPos(iIngotpos)
'''''
'''''End Function
'''''============================================================================================================================
'''''
''''''�T�v      :�����ʒu������̎擾
''''''���Ұ��@�@:�ϐ���         ,IO ,�^       ,����
''''''�@�@      :iIngotPos�@�@�@,I  ,Integer�@,�������J�n�ʒu
''''''�@�@      :�߂�l         ,O  ,String �@,�����ʒu������
''''''����      :�����ʒu�������Ԃ�
''''''����      :2001/07/11�@��� �쐬
'''''Public Function GetWafPos(iIngotpos As Integer) As String
'''''
'''''    Dim i As Integer
'''''    Dim j As Integer
'''''
'''''    If iIngotpos >= 1000 Then
'''''        i = Int(iIngotpos / 100)
'''''        j = iIngotpos Mod 100
'''''        GetWafPos = Chr$(i - 10 + Asc("A")) & Format(j, "00")
'''''    Else
'''''        GetWafPos = Format(iIngotpos, "000")
'''''    End If
'''''
'''''End Function
'''''============================================================================================================================
'''''
''''''�T�v      :����]�����@�w���e�[�u���̍쐬
''''''���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
''''''�@�@      :pSXLMng�@�@�@,I  ,typ_TBCME042   �@,SXL�Ǘ�
''''''�@�@      :pWafSmp�@�@�@,I  ,typ_TBCME044   �@,WF�T���v���Ǘ�
''''''�@�@      :pMesInd�@�@�@,O  ,typ_TBCMY003   �@,����]�����@�w��
''''''�@�@      :�߂�l       ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
''''''����      :����]�����@�w���e�[�u�����쐬����
''''''����      :2001/07/23�@��� �쐬
'''''Public Function MakeMesIndTbl(pSXLMng() As typ_TBCME042, pWafSmp() As typ_TBCME044, pMesInd() As typ_TBCMY003) As FUNCTION_RETURN
'''''
'''''    Dim tmpSpWFSamp() As typ_SpWFSamp
'''''    Dim sHin As String
'''''    Dim sDKAN As String
'''''    Dim m As Integer
'''''    Dim n As Integer
'''''    Dim i As Integer
'''''    Dim j As Integer
'''''    Dim k As Integer
'''''
'''''    '' ����]�����@�w���p�̐��i�d�l���擾
'''''    j = 0
'''''    m = UBound(pSXLMng)
'''''    ReDim tmpSpWFSamp(m)
'''''    For i = 1 To m
''''''==================================== 2003/04/17 okazaki
'''''            sHin = RTrim$(pSXLMng(i).hinban)
'''''        If (sHin <> "" And sHin <> "G" And sHin <> "Z") Then
''''''=======================================================
'''''            j = j + 1
'''''            tmpSpWFSamp(j).HIN.hinban = pSXLMng(i).hinban
'''''            tmpSpWFSamp(j).HIN.mnorevno = pSXLMng(i).REVNUM
'''''            tmpSpWFSamp(j).HIN.factory = pSXLMng(i).factory
'''''            tmpSpWFSamp(j).HIN.opecond = pSXLMng(i).opecond
'''''            If scmzc_getWF(tmpSpWFSamp(j)) = FUNCTION_RETURN_FAILURE Then
'''''                MakeMesIndTbl = FUNCTION_RETURN_FAILURE
'''''                Exit Function
'''''            End If
'''''        End If
'''''    Next i
'''''    ReDim Preserve tmpSpWFSamp(j)
'''''
'''''    '' ����]�����@�w���e�[�u���̍쐬
'''''    k = 0
'''''    m = UBound(pWafSmp)
'''''    n = UBound(tmpSpWFSamp)
'''''
'''''    ReDim pMesInd(m * 17)   ''### Add.03/05/20 �㓡 ###
''''''''    ReDim pMesInd(m * 15)
'''''    For i = 1 To m
'''''        For j = 1 To n
'''''            If tmpSpWFSamp(j).HIN.hinban = pWafSmp(i).hinban Then
'''''                Exit For
'''''            End If
'''''        Next j
'''''        If j <= n Then
'''''            With tmpSpWFSamp(j)
'''''                sDKAN = IIf(.HWFIGKBN = "3", "R ", "V ") & Format(.HWFANTNP, "@@@@") & Format(.HWFANTIM, "@@@@")
'''''                If pWafSmp(i).WFINDRS <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "RES"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "RES"
'''''                    pMesInd(k).NETSU = ""
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFRSPOH & .HWFRSPOT & .HWFRSPOI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDOI <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "OI"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OI"
'''''                    pMesInd(k).NETSU = ""
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFONSPH & .HWFONSPT & .HWFONSPI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDB1 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "BMD"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "BMD1"
'''''                    pMesInd(k).NETSU = .HWFBM1NS
'''''                    pMesInd(k).ET = .HWFBM1SZ & Format(.HWFBM1ET, "00")
'''''                    pMesInd(k).MES = .HWFBM1SH & .HWFBM1ST & .HWFBM1SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDB2 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "BMD"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "BMD2"
'''''                    pMesInd(k).NETSU = .HWFBM2NS
'''''                    pMesInd(k).ET = .HWFBM2SZ & Format(.HWFBM2ET, "00")
'''''                    pMesInd(k).MES = .HWFBM2SH & .HWFBM2ST & .HWFBM2SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDB3 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "BMD"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "BMD3"
'''''                    pMesInd(k).NETSU = .HWFBM3NS
'''''                    pMesInd(k).ET = .HWFBM3SZ & Format(.HWFBM3ET, "00")
'''''                    pMesInd(k).MES = .HWFBM3SH & .HWFBM3ST & .HWFBM3SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDL1 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "OSF"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OSF1"
'''''                    pMesInd(k).NETSU = .HWFOF1NS
'''''                    pMesInd(k).ET = .HWFOF1SZ & Format(.HWFOF1ET, "00")
'''''                    pMesInd(k).MES = .HWFOF1SH & .HWFOF1ST & .HWFOF1SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDL2 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "OSF"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OSF2"
'''''                    pMesInd(k).NETSU = .HWFOF2NS
'''''                    pMesInd(k).ET = .HWFOF2SZ & Format(.HWFOF2ET, "00")
'''''                    pMesInd(k).MES = .HWFOF2SH & .HWFOF2ST & .HWFOF2SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDL3 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "OSF"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OSF3"
'''''                    pMesInd(k).NETSU = .HWFOF3NS
'''''                    pMesInd(k).ET = .HWFOF3SZ & Format(.HWFOF3ET, "00")
'''''                    pMesInd(k).MES = .HWFOF3SH & .HWFOF3ST & .HWFOF3SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDL4 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "OSF"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OSF4"
'''''                    pMesInd(k).NETSU = .HWFOF4NS
'''''                    pMesInd(k).ET = .HWFOF4SZ & Format(.HWFOF4ET, "00")
'''''                    pMesInd(k).MES = .HWFOF4SH & .HWFOF4ST & .HWFOF4SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDDS <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "DSOD"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "DSOD"
'''''                    pMesInd(k).NETSU = "G0"
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = ""
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDDZ <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "DZ"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "DZ"
'''''                    pMesInd(k).NETSU = .HWFMKNSW
'''''                    pMesInd(k).ET = .HWFMKSZY & Format(.HWFMKCET, "00")
'''''                    pMesInd(k).MES = .HWFMKSPH & .HWFMKSPT & .HWFMKSPR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDSP <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "SPV"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "SPV"
'''''                    pMesInd(k).NETSU = ""
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFSPVSH & .HWFSPVST & .HWFSPVSI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDDO1 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "DOI"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "DOI1"
'''''                    pMesInd(k).NETSU = .HWFOS1NS
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFOS1SH & .HWFOS1ST & .HWFOS1SI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDDO2 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "DOI"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "DOI2"
'''''                    pMesInd(k).NETSU = .HWFOS2NS
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFOS2SH & .HWFOS2ST & .HWFOS2SI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDDO3 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "DOI"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "DOI3"
'''''                    pMesInd(k).NETSU = .HWFOS3NS
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                '################################## Add,03/05/20 �㓡 ##########
'''''                If pWafSmp(i).WFINDOT1 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''''''                    pMesInd(k).OSITEM = "OTH"
'''''                    pMesInd(k).OSITEM = "OTH1"  'upd 2003/06/10 hitec)matsumoto
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OTHER1"
''''''''                    pMesInd(k).NETSU = .HWFOS3NS
''''''''                    pMesInd(k).ET = ""
''''''''                    pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
'''''                    pMesInd(k).NETSU = vbNullString
'''''                    pMesInd(k).ET = vbNullString
'''''                    pMesInd(k).MES = vbNullString
'''''''''                    pMesInd(k).DKAN = vbNullString  '03/05/22
'''''                    pMesInd(k).DKAN = sDKAN 'upd 2003/09/10 hitec)matsumoto
'''''                End If
'''''                If pWafSmp(i).WFINDOT2 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''''''                    pMesInd(k).OSITEM = "OTH"
'''''                    pMesInd(k).OSITEM = "OTH2"  'upd 2003/06/10 hitec)matsumoto
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OTHER2"
''''''''                    pMesInd(k).NETSU = .HWFOS3NS
''''''''                    pMesInd(k).ET = ""
''''''''                    pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
'''''                    pMesInd(k).NETSU = vbNullString
'''''                    pMesInd(k).ET = vbNullString
'''''                    pMesInd(k).MES = vbNullString
'''''''''                    pMesInd(k).DKAN = vbNullString  '03/05/22
'''''                    pMesInd(k).DKAN = sDKAN 'upd 2003/09/10 hitec)matsumoto
'''''                End If
'''''                '################################## End,03/05/20 �㓡 ##########
'''''            End With
'''''        End If
'''''    Next i
'''''    ReDim Preserve pMesInd(k)
'''''
'''''    MakeMesIndTbl = FUNCTION_RETURN_SUCCESS
'''''
'''''End Function
'''''============================================================================================================================
'''''
''''''�T�v      :�t�^�c���㉺�ɕ���
''''''����      :�t�^�c�T���v�����㉺�ɕ�������
''''''����      :2001/10/05�@��� �쐬
'''''Public Sub SeparateUD()
'''''    'Step3.2�ɂāA�@�\�p�~
'''''End Sub

'�T�v      :�ϐ�������
'���Ұ�    :�ϐ���        ,IO ,�^              ,����
'          :typ_A         ,IO ,typ_AllTypes ,�e���\����
'����      :
'����      :
Public Sub InitHensu2(typ_C As typ_AllTypesC)
    Dim i As Integer, j As Integer
    
'Chg Start 2011/03/10 SMPK Miyata
'    For i = 1 To 2
    For i = 1 To SXL_MAXSMP
'Chg End   2011/03/10 SMPK Miyata
        For j = 0 To MAXCNT
            With typ_C
                .typ_y013(i, j).SAMPLEID = "0"         ' �T���v��ID
                .typ_y013(i, j).MESDATA1 = "0"         ' ����f�[�^���̂P
                .typ_y013(i, j).MESDATA2 = "0"         ' ����f�[�^���̂Q
                .typ_y013(i, j).MESDATA3 = "0"         ' ����f�[�^���̂R
                .typ_y013(i, j).MESDATA4 = "0"         ' ����f�[�^���̂S
                .typ_y013(i, j).MESDATA5 = "0"         ' ����f�[�^���̂T
                .typ_y013(i, j).MESDATA6 = "0"         ' ����f�[�^���̂U
                .typ_y013(i, j).MESDATA7 = "0"         ' ����f�[�^���̂V
                .typ_y013(i, j).MESDATA8 = "0"         ' ����f�[�^���̂W
                .typ_y013(i, j).MESDATA9 = "0"         ' ����f�[�^���̂X
                .typ_y013(i, j).MESDATA10 = "0"        ' ����f�[�^���̂P�O
                .typ_y013(i, j).MESDATA11 = "0"        ' ����f�[�^���̂P�P
                .typ_y013(i, j).MESDATA12 = "0"        ' ����f�[�^���̂P�Q
                .typ_y013(i, j).MESDATA13 = "0"        ' ����f�[�^���̂P�R
                .typ_y013(i, j).MESDATA14 = "0"        ' ����f�[�^���̂P�S
                .typ_y013(i, j).MESDATA15 = "0"        ' ����f�[�^���̂P�T
            End With
        Next
    Next
End Sub

'�T�v      :�ϐ�������
'���Ұ�    :�ϐ���        ,IO ,�^              ,����
'          :typ_A_EP      ,IO ,typ_AllTypes_EP ,�e���\����
'����      :
'����      :2006/08/15 �V�K�쐬 �G�s��s�]���ǉ��Ή� SMP)kondoh
Public Sub InitHensu2_EP(typ_C_EP As typ_AllTypesC_EP)
    Dim i As Integer, j As Integer
    
    For i = 1 To 2
        For j = 0 To MAXCNT_EP
            With typ_C_EP
                .typ_y022(i, j).SAMPLEID = "0"          ' �T���v��ID
                .typ_y022(i, j).MESDATA1 = "0"         ' ����f�[�^���̂P
                .typ_y022(i, j).MESDATA2 = "0"         ' ����f�[�^���̂Q
                .typ_y022(i, j).MESDATA3 = "0"         ' ����f�[�^���̂R
                .typ_y022(i, j).MESDATA4 = "0"         ' ����f�[�^���̂S
                .typ_y022(i, j).MESDATA5 = "0"         ' ����f�[�^���̂T
                .typ_y022(i, j).MESDATA6 = "0"         ' ����f�[�^���̂U
                .typ_y022(i, j).MESDATA7 = "0"         ' ����f�[�^���̂V
                .typ_y022(i, j).MESDATA8 = "0"         ' ����f�[�^���̂W
                .typ_y022(i, j).MESDATA9 = "0"         ' ����f�[�^���̂X
                .typ_y022(i, j).MESDATA10 = "0"        ' ����f�[�^���̂P�O
                .typ_y022(i, j).MESDATA11 = "0"        ' ����f�[�^���̂P�P
                .typ_y022(i, j).MESDATA12 = "0"        ' ����f�[�^���̂P�Q
                .typ_y022(i, j).MESDATA13 = "0"        ' ����f�[�^���̂P�R
                .typ_y022(i, j).MESDATA14 = "0"        ' ����f�[�^���̂P�S
                .typ_y022(i, j).MESDATA15 = "0"        ' ����f�[�^���̂P�T
            End With
        Next
    Next
End Sub

'------------------------------------------------
' �d�lNull�`�F�b�N(WFC)
'------------------------------------------------

'�T�v      :WFC��������̊e�������ڂ̕ۏؕ��@��'H'�܂���'S'�̏ꍇ�A�d�l�l��Null(-1)���ǂ����𔻒f����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tSiyou        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�i�ԁA�d�l�A���������擾�p
'          :sErrMsg       ,IO ,String                               :�װү����
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���� = FUNCTION_RETURN_SUCCESS : OK
'                                                                           FUNCTION_RETURN_FAILURE : NG
'����      :
'����      :2003/12/13 �V�K�쐬�@�V�X�e���u���C��

Private Function funWfChkNull(tSiyou As type_DBDRV_scmzc_fcmlc001c_Siyou, sErrMsg As String) As FUNCTION_RETURN
    Dim dShiyo()    As Double
    Dim sHosyo      As String
    Dim cnt         As Integer
    
    '������
    funWfChkNull = FUNCTION_RETURN_SUCCESS
    
    '--------------- RS(���R) ---------------
    ReDim dShiyo(5)
    dShiyo(1) = tSiyou.HWFRMIN          ' �i�v�e���R����
    dShiyo(2) = tSiyou.HWFRMAX          ' �i�v�e���R���
    dShiyo(3) = tSiyou.HWFRAMIN         ' �i�v�e���R���ω���
    dShiyo(4) = tSiyou.HWFRAMAX         ' �i�v�e���R���Ϗ��
    dShiyo(5) = tSiyou.HWFRMBNP         ' �i�v�e���R�ʓ����z
    If fncJissekiHantei_nl(tSiyou.HWFRHWYS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(RS)"
        funWfChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00130"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
    '--------------- Oi(�_�f�Z�x) ---------------
    ReDim dShiyo(5)
    dShiyo(1) = tSiyou.HWFONMIN         ' �i�v�e�_�f�Z�x����
    dShiyo(2) = tSiyou.HWFONMAX         ' �i�v�e�_�f�Z�x���
    dShiyo(3) = tSiyou.HWFONMBP         ' �i�v�e�_�f�Z�x�ʓ����z
    dShiyo(4) = tSiyou.HWFONAMN         ' �i�v�e�_�f�Z�x���ω���
    dShiyo(5) = tSiyou.HWFONAMX         ' �i�v�e�_�f�Z�x���Ϗ��
    If fncJissekiHantei_nl(tSiyou.HWFONHWS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(Oi)"
        funWfChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00131"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
    '--------------- BMD1,BMD2,BMD3 ---------------
    'BMD�̎g�pNULL�`�F�b�N���폜�i�����OK�Ƃ���B�j�@          2003/12/19 tuku
''''    For cnt = 1 To 3
''''        ReDim dShiyo(1)
''''        If cnt = 1 Then         'BMD1
''''            sHosyo = tSiyou.HWFBM1HS            ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
''''            dShiyo(1) = tSiyou.HWFBM1MBP        ' �i�v�e�a�l�c�P�ʓ����z
''''        ElseIf cnt = 2 Then     'BMD2
''''            sHosyo = tSiyou.HWFBM2HS            ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
''''            dShiyo(1) = tSiyou.HWFBM2MBP        ' �i�v�e�a�l�c�Q�ʓ����z
''''        ElseIf cnt = 3 Then     'BMD3
''''            sHosyo = tSiyou.HWFBM3HS            ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
''''            dShiyo(1) = tSiyou.HWFBM3MBP        ' �i�v�e�a�l�c�R�ʓ����z
''''        End If
''''        If fncJissekiHantei_nl(sHosyo, dShiyo) = False Then
''''            sErrMsg = sErrMsg & "(BMD" & cnt & ")"
''''            funWfChkNull = FUNCTION_RETURN_FAILURE
''''            Exit Function
''''        End If
''''    Next cnt
    
    '--------------- OSF1,OSF2,OSF3,OSF4 ---------------
    '�`�F�b�N�Ȃ�
    
    '--------------- DSOD ---------------
    ReDim dShiyo(4)
    dShiyo(1) = tSiyou.HWFDSOMX         ' �i�v�e�c�r�n�c���
    dShiyo(2) = tSiyou.HWFDSOMN         ' �i�v�e�c�r�n�c����
    dShiyo(3) = tSiyou.HWFDSOAX         ' �i�v�e�c�r�n�c�̈���
    dShiyo(4) = tSiyou.HWFDSOAN         ' �i�v�e�c�r�n�c�̈扺��
    If fncJissekiHantei_nl(tSiyou.HWFDSOHS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(DSOD)"
        funWfChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00143"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
    '--------------- DZ�� ---------------
    '�`�F�b�N�Ȃ�
        
    '--------------- SPVFE ---------------
    '�`�F�b�N�Ȃ�
        
    '--------------- DOI1(�_�f�͏o1),DOI2(�_�f�͏o2),DOI3(�_�f�͏o3) ---------------
    '�`�F�b�N�Ȃ�
        
    '--------------- AOI(�c���_�f) ---------------
    '�`�F�b�N�Ȃ�

    '--------------- GD ---------------         'GD�ǉ��@05/01/27 ooba
    '�`�F�b�N�Ȃ�
'    ReDim dShiyo(2)     'Den
'    dShiyo(1) = tSiyou.HWFDENMX         ' �i�v�e�c�������
'    dShiyo(2) = tSiyou.HWFDENMN         ' �i�v�e�c��������
'    If fncJissekiHantei_nl(tSiyou.HWFDENHS, dShiyo) = False Then
'        sErrMsg = sErrMsg & "(GD_Den)"
'        funWfChkNull = FUNCTION_RETURN_FAILURE
'        Exit Function
'    End If
'
'    ReDim dShiyo(2)     'DVD2
'    dShiyo(1) = tSiyou.HWFDVDMXN        ' �i�v�e�c�u�c�Q���
'    dShiyo(2) = tSiyou.HWFDVDMNN        ' �i�v�e�c�u�c�Q����
'    If fncJissekiHantei_nl(tSiyou.HWFDVDHS, dShiyo) = False Then
'        sErrMsg = sErrMsg & "(GD_DVD2)"
'        funWfChkNull = FUNCTION_RETURN_FAILURE
'        Exit Function
'    End If
'
'    ReDim dShiyo(2)     'L/DL
'    dShiyo(1) = tSiyou.HWFLDLMX         ' �i�v�e�k�^�c�k���
'    dShiyo(2) = tSiyou.HWFLDLMN         ' �i�v�e�k�^�c�k����
'    If fncJissekiHantei_nl(tSiyou.HWFLDLHS, dShiyo) = False Then
'        sErrMsg = sErrMsg & "(GD_LDL)"
'        funWfChkNull = FUNCTION_RETURN_FAILURE
'        Exit Function
'    End If
    
End Function

''Upd start 2005/06/22 (TCS)t.terauchi  SPV9�_�Ή�
'�T�v      :SPV����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_j016      ,I  ,typ_TBCMJ016                         :SPV���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :iKubun        ,I  ,Integer                              :�敪(1:Fe�Z�x, 2:�g�U��, 3:Nr�Z�x)
'          :sSxlPos       ,I  ,String                               :SXL�ʒu(TOP/BOT)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :SPV������s��
'����      :2005/06/22 �V�K�쐬�@(TCS)t.terauchi
Public Function WfCrSpvJudg_New(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_j016 As typ_TBCMJ016, _
                            bJudg As Boolean, iKubun As Integer, sSxlPos As String) As Boolean
    Dim ErrInfo         As ERROR_INFOMATION         '�G���[���\����
    Dim sp              As W_SPV                    'SPV�\����
    Dim sSokutei_Fe     As String                   'Fe�Z�x�@������@
    Dim sSokutei_Diff   As String                   '�g�U���@������@
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
    Dim sSokutei_Nr     As String                   'Nr�Z�x�@������@
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------

    bJudg = True

    '������@�擾
    sSokutei_Fe = typ_si.HWFSPVSH & typ_si.HWFSPVST & typ_si.HWFSPVSI
    sSokutei_Diff = typ_si.HWFDLSPH & typ_si.HWFDLSPT & typ_si.HWFDLSPI
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
    sSokutei_Nr = typ_si.HWFNRSH & typ_si.HWFNRST & typ_si.HWFNRSI
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------

    'Fe�Z�x�Ɗg�U���̑�����@���Ⴄ�ꍇ
    If sSokutei_Fe <> sSokutei_Diff Then
        
        'Fe�Z�x�E�g�U�����ɁA�d�l�L��̏ꍇ����Err�Ƃ���
        If ((typ_si.HWFDLHWS = "H") And CheckKHN(typ_si.HWFDLKHN, 16, sSxlPos)) _
            And ((typ_si.HWFSPVHS = "H") And CheckKHN(typ_si.HWFSPVKN, 15, sSxlPos)) Then
        
            bJudg = False
            WfCrSpvJudg_New = False
            Exit Function
        End If
    End If

    'SPV��������ݒ�
    sp.GuaranteeSpvFe.cMeth = typ_si.HWFSPVSH   '�iWFSPVFE����ʒu�Q��
    sp.GuaranteeSpvFe.cCount = typ_si.HWFSPVST  '�iWFSPVFE����ʒu�Q�_
    sp.GuaranteeSpvFe.cPos = typ_si.HWFSPVSI    '�iWFSPVFE����ʒu�Q��
    sp.GuaranteeSpvFe.cObj = typ_si.HWFSPVHT    '�iWFSPVFE�ۏؕ��@�Q��
    sp.GuaranteeSpvFe.cJudg = typ_si.HWFSPVHS   '�iWFSPVFE�ۏؕ��@�Q��
    sp.GuaranteeSpv.cMeth = typ_si.HWFDLSPH     '�iWF�g�U������ʒu�Q��
    sp.GuaranteeSpv.cCount = typ_si.HWFDLSPT    '�iWF�g�U������ʒu�Q�_
    sp.GuaranteeSpv.cPos = typ_si.HWFDLSPI      '�iWF�g�U������ʒu�Q��
    sp.GuaranteeSpv.cObj = typ_si.HWFDLHWT      '�iWF�g�U���ۏؕ��@�Q��
    sp.GuaranteeSpv.cJudg = typ_si.HWFDLHWS     '�iWF�g�U���ۏؕ��@�Q��
    
    
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
    '���荀�ڂ�Nr�Z�x��ǉ�
    sp.GuaranteeSpvNr.cMeth = typ_si.HWFNRSH    '�iWFSPVNR����ʒu�Q��
    sp.GuaranteeSpvNr.cCount = typ_si.HWFNRST   '�iWFSPVNR����ʒu�Q�_
    sp.GuaranteeSpvNr.cPos = typ_si.HWFNRSI     '�iWFSPVNR����ʒu�Q��
    sp.GuaranteeSpvNr.cObj = typ_si.HWFNRHT     '�iWFSPVNR�ۏؕ��@�Q��
    sp.GuaranteeSpvNr.cJudg = typ_si.HWFNRHS    '�iWFSPVNR�ۏؕ��@�Q��
    sp.SpecSpvNrMax = typ_si.HWFNRMX            '�iWFNR�Z�x���
    sp.SpecSpvNrAvMax = typ_si.HWFNRAM          '�iWFNR���Ϗ��
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------

    sp.SpecSpvFeMax = typ_si.HWFSPVMX           '�iWFFe�Z�x���
    sp.SpecSpvAvMax = typ_si.HWFSPVAM           '�iWF���Ϗ��
    sp.SpecSpvMin = typ_si.HWFDLMIN             '�iWF�g�U������
    sp.SpecSpvMax = typ_si.HWFDLMAX             '�iWF�g�U�����

    'Fe�Z�x����
    If iKubun = "1" Then
        sp.Spv(0) = typ_j016.MAX_FE                                 'Fe�Z�x�|MAX
        sp.Spv(1) = typ_j016.MIN_FE                                 'Fe�Z�x�|MIN
        sp.Spv(2) = Format(typ_j016.AVE_FE, "0.00")                 'Fe�Z�x�|AVE
        sp.Spv(3) = typ_j016.CENTER_FE                              'Fe�Z�x�|�Z���^�[
    
        If sSokutei_Fe = "AMX" Then
            'SPV(Fe�Z�x MAP����)����
            If WfSPV_Fe_AMXJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            If sp.GuaranteeSpvFe.cJudg = JudgCodeW01 Then ''SPVFE�Z�x�@����L��
                ' SPV_Fe PUA�l��Fe�Z�xPUA���ȉ�
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUA_FE, -1, typ_si.HWFSPVPUG)
                End If
                ' SPV_Fe PUA%�l��Fe�Z�xPUA��
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUAP_FE, -1, typ_si.HWFSPVPUR)
                End If
                ' SPV_Fe STD�l��Fe�Z�x�W���΍��ȉ�
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.STD_FE, -1, typ_si.HWFSPVSTD)
                End If
            End If
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        ElseIf sSokutei_Fe = "V9T" Then
            'SPV(Fe�Z�x 9�_����)����
            If WfSPV_Fe_V9TJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
        Else
            bJudg = False
            WfCrSpvJudg_New = False
            Exit Function
        End If
    
    '�g�U������
    ElseIf iKubun = "2" Then
        sp.Spv(0) = typ_j016.MAX_DIFF                               '�g�U���|MAX
        sp.Spv(1) = typ_j016.MIN_DIFF                               '�g�U���|MIN
        sp.Spv(2) = Format(typ_j016.AVE_DIFF, "0.0")                '�g�U���|AVE
        sp.Spv(3) = typ_j016.CENTER_DIFF                            '�g�U���|�Z���^�[
    
        If sSokutei_Diff = "AMX" Then
            'SPV(�g�U�� MAP����)����
            If WfSPV_DIFF_AMXJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
            If sp.GuaranteeSpv.cJudg = JudgCodeW01 Then ''SPV�g�U���@����L��
                ' SPV_�g�U��PUA�l���g�U��PUA���ȏ�
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUA_DIFF, typ_si.HWFDLPUG, -1)
                End If
                ' SPV_�g�U��PUA%�l���g�U��PUA���ȏ�
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUAP_DIFF, typ_si.HWFDLPUR, -1)
                End If
            End If
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        ElseIf sSokutei_Diff = "V9T" Then
            'SPV(�g�U�� 9�_����)����
            If WfSPV_DIFF_V9TJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
        Else
            bJudg = False
            WfCrSpvJudg_New = False
            Exit Function
        End If
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
    'Nr����
    ElseIf iKubun = "3" Then

        sp.Spv(0) = typ_j016.SPV_Nr_MAX                             'Nr�Z�x�|MAX
        sp.Spv(2) = Format(typ_j016.SPV_Nr_AVE, "0.00")             'Nr�Z�x�|AVE
    
        If sSokutei_Nr = "AMX" Then
            'SPV(Fe�Z�x MAP����)����
            If WfSPV_Nr_AMXJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
            If sp.GuaranteeSpvNr.cJudg = JudgCodeW01 Then ''SPVNR�Z�x�@����L��
                ' SPV_Nr PUA�l��Nr�Z�xPUA���ȉ�
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUA_NR, -1, typ_si.HWFNRPUG)
                End If
                ' SPV_Nr PUA%�l��Nr�Z�xPUA��
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUAP_NR, -1, typ_si.HWFNRPUR)
                End If
                ' SPV_Nr STD�l��Nr�Z�x�W���΍��ȉ�
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.STD_NR, -1, typ_si.HWFNRSTD)
                End If
            End If
        ElseIf sSokutei_Fe = "V9T" Then
            'SPV(Fe�Z�x 9�_����)����
            If WfSPV_Nr_V9TJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
        Else
            bJudg = False
            WfCrSpvJudg_New = False
            Exit Function
        End If
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
    Else
        bJudg = False
        WfCrSpvJudg_New = False
        Exit Function
    End If

    If sp.JudgSpv <> True Then
        bJudg = False
    End If

    WfCrSpvJudg_New = True

End Function
'Upd end   2005/06/21 (TCS)t.terauchi  SPV9�_�Ή�

'�T�v      :Warp����
'���Ұ�    :�ϐ���        ,IO ,�^                       :����
'          :dWarpMax      ,I  ,Double                   :Warp���
'          :dMeas         ,I  ,Double                   :����l
'          :�߂�l        ,O  ,Boolean                  :True������OK,False������NG
'����      :
'����      :05/12/16 ooba
Public Function WfWarpJudg(dWarpMax As Double, dMeas As Double) As Boolean

    WfWarpJudg = True
    
    If dMeas = -1 Then Exit Function
    
    '�d�l�l��0orNULL�̏ꍇ�͔���OK�Ƃ���
    If dWarpMax = 0 Then dWarpMax = -1
    
    'Warp����(����l������l�Ȃ画��OK)
    WfWarpJudg = RangeDecision_nl(dMeas, -1, dWarpMax)
        
End Function

'�T�v      :�����p�x����
'���Ұ�    :�ϐ���        ,IO ,�^                       :����
'          :dKakuMin      ,I  ,Double                   :�����ʌX����
'          :dKakuMax      ,I  ,Double                   :�����ʌX���
'          :dMeas         ,I  ,Double                   :����l
'          :�߂�l        ,O  ,Boolean                  :True������OK,False������NG
'����      :
'����      :05/12/16 ooba
Public Function WfKakuJudg(dKakuMin As Double, dKakuMax As Double, dMeas As Double) As Boolean

    WfKakuJudg = True
    
    If dMeas = -1 Then Exit Function
    
    '�d�l�l��0orNULL�̏ꍇ�͔���OK�Ƃ���
    If dKakuMin = 0 Then dKakuMin = -1
    If dKakuMax = 0 Then dKakuMax = -1
    
    '�����p�x����(�����l������l������l�Ȃ画��OK)
    WfKakuJudg = RangeDecision_nl(dMeas, dKakuMin, dKakuMax)
        
End Function

'�T�v      :WF�d�l�\���̃N���A�֐�
'���Ұ�    :�ϐ���        ,IO ,�^                       :����
'����      :�����̃v���V�[�W���ɋL�q�����VB�̐����Ɉ���������̂ŁA�ʃv���V�[�W���ō쐬����B
'����      :06/06/12 �V�K�쐬
Public Sub Crear_type_Siyou_Spv()
    'typ_Ctype��������
    Dim clear_typeC(0) As typ_AllTypesC
    typ_CType = clear_typeC(0)
'Add Start 2011/03/09 SMPK Miyata
    ReDim typ_CType.dblScut(SXL_MAXSMP)             ' �ăJ�b�g�ʒu
    ReDim typ_CType.bOKNG(SXL_MAXSMP)               ' ���R����
    ReDim typ_CType.COEF(SXL_MAXSMP)                ' �ΐ͌W��
    ReDim typ_CType.JudgRes(SXL_MAXSMP)             ' ���R����
    ReDim typ_CType.JudgRrg(SXL_MAXSMP)             ' RRG����
    ReDim typ_CType.typ_y013(SXL_MAXSMP, MAXCNT)    ' ���茋��
    ReDim typ_CType.typ_hage(SXL_MAXSMP)            ' ���グ�I������
    ReDim typ_CType.typ_rslt(SXL_MAXSMP, MAXCNT)    ' �e���я��
'Add End   2011/03/09 SMPK Miyata

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    'typ_Ctype_EP��������
    Dim clear_typeC_EP(0) As typ_AllTypesC_EP
    typ_CType_EP = clear_typeC_EP(0)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'    Dim clear_typ_si_Spv As type_DBDRV_scmzc_fcmlc001c_Siyou_Spv
'    typ_si_Spv = clear_typ_si_Spv
End Sub

'�T�v      :BMD(EP)����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_y022      ,I  ,typ_TBCMY022                         :BMD���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :bmflg         ,I  ,Integer                              :BMD�׸�(1:BMD1, 2:BMD2, 3:BMD3)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :BMD������s��(�֐�WfCrBmdJudg����ɍ쐬)
'����      :�V�K�쐬 2006/08/15 �G�s��s�]���ǉ��Ή� SMP)kondoh
Public Function EpBmdJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y022 As typ_TBCMY022, _
                            bJudg As Boolean, _
                            bmflg As Integer) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim bm      As W_BMD                    'BMD�\����
    Dim c0      As Integer

    Dim keisu As Double
    Const keisu1 As Double = 10000
    Const keisu2 As Double = 10000
    Const keisu3 As Double = 10000
    Const keisu4 As Double = 10000
    Const keisu5 As Double = 10000
    Const keisu6 As Double = 333000
    Const keisu7 As Double = 10000
    Const keisu8 As Double = 10000
    Const keisu9 As Double = 10000 'Add 2012/07/20 Y.Hitomi
    
    bJudg = True

    'BMD(EP)��������ݒ�
    Select Case bmflg
    Case 1
        bm.GuaranteeBmd.cMeth = typ_si.HEPBM1SH   '�iEPBMD1����ʒu_��
        bm.GuaranteeBmd.cCount = typ_si.HEPBM1ST  '�iEPBMD1����ʒu_�_
        bm.GuaranteeBmd.cPos = typ_si.HEPBM1SR    '�iEPBMD1����ʒu_��
        bm.GuaranteeBmd.cObj = typ_si.HEPBM1HT    '�iEPBMD1�ۏؕ��@_��
        bm.GuaranteeBmd.cJudg = typ_si.HEPBM1HS   '�iEPBMD1�ۏؕ��@_��
        bm.SpecBmdAveMin = typ_si.HEPBM1AN        '�iEPBMD1���ω���
        bm.SpecBmdAveMax = typ_si.HEPBM1AX        '�iEPBMD1���Ϗ��
        bm.SpecBmdMBP = typ_si.HEPBM1MBP          '�iEPBMD1�ʓ����z
        bm.SpecBmdMCL = NtoS(typ_si.HEPBM1MCL)    '�iEPBMD1�ʓ��v�Z
        bm.Antnp = typ_si.HEPANTNP                '�iEPAN���x
    Case 2
        bm.GuaranteeBmd.cMeth = typ_si.HEPBM2SH   '�iEPBMD2����ʒu_��
        bm.GuaranteeBmd.cCount = typ_si.HEPBM2ST  '�iEPBMD2����ʒu_�_
        bm.GuaranteeBmd.cPos = typ_si.HEPBM2SR    '�iEPBMD2����ʒu_��
        bm.GuaranteeBmd.cObj = typ_si.HEPBM2HT    '�iEPBMD2�ۏؕ��@_��
        bm.GuaranteeBmd.cJudg = typ_si.HEPBM2HS   '�iEPBMD2�ۏؕ��@_��
        bm.SpecBmdAveMin = typ_si.HEPBM2AN        '�iEPBMD2���ω���
        bm.SpecBmdAveMax = typ_si.HEPBM2AX        '�iEPBMD2���Ϗ��
        bm.SpecBmdMBP = typ_si.HEPBM2MBP          '�iEPBMD2�ʓ����z
        bm.SpecBmdMCL = NtoS(typ_si.HEPBM2MCL)    '�iEPBMD2�ʓ��v�Z
        bm.Antnp = typ_si.HEPANTNP                '�iEPAN���x
    Case 3
        bm.GuaranteeBmd.cMeth = typ_si.HEPBM3SH   '�iEPBMD3����ʒu_��
        bm.GuaranteeBmd.cCount = typ_si.HEPBM3ST  '�iEPBMD3����ʒu_�_
        bm.GuaranteeBmd.cPos = typ_si.HEPBM3SR    '�iEPBMD3����ʒu_��
        bm.GuaranteeBmd.cObj = typ_si.HEPBM3HT    '�iEPBMD3�ۏؕ��@_��
        bm.GuaranteeBmd.cJudg = typ_si.HEPBM3HS   '�iEPBMD3�ۏؕ��@_��
        bm.SpecBmdAveMin = typ_si.HEPBM3AN        '�iEPBMD3���ω���
        bm.SpecBmdAveMax = typ_si.HEPBM3AX        '�iEPBMD3���Ϗ��
        bm.SpecBmdGsAveMin = typ_si.HEPBM3GSAN    '�iEPBMD3���ω���(�O��)�@09/05/07 ooba
        bm.SpecBmdGsAveMax = typ_si.HEPBM3GSAX    '�iEPBMD3���Ϗ��(�O��)�@09/05/07 ooba
        bm.SpecBmdMBP = typ_si.HEPBM3MBP          '�iEPBMD3�ʓ����z
        bm.SpecBmdMCL = NtoS(typ_si.HEPBM3MCL)    '�iEPBMD3�ʓ��v�Z
        bm.Antnp = typ_si.HEPANTNP                '�iEPAN���x
    End Select
    
    If bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "H" Then
        keisu = keisu1
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "H" Then
        keisu = keisu2
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu3
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu4
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "A" Then
        keisu = keisu5
    ElseIf bm.GuaranteeBmd.cMeth = "G" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu6
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu7
    ElseIf bm.GuaranteeBmd.cMeth = "8" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu8
    'Add Start 2012/07/20 Y.Hitomi
    ElseIf bm.GuaranteeBmd.cMeth = "P" Then
        keisu = keisu9
    'Add End 2012/07/20 Y.Hitomi
    Else
        bJudg = False
        EpBmdJudg = False
        Exit Function
    End If
    
    With bm
        .BMD(0) = NtoZ2(typ_y022.MESDATA1)                   'BMD����l
        .BMD(1) = NtoZ2(typ_y022.MESDATA2)                   'BMD����l
        .BMD(2) = NtoZ2(typ_y022.MESDATA3)                   'BMD����l
        .BMD(3) = NtoZ2(typ_y022.MESDATA4)                   'BMD����l
        .BMD(4) = NtoZ2(typ_y022.MESDATA5)                   'BMD����l
        .BmdAntnp = NtoZ2(Mid(typ_y022.DKAN, 3, 4))
        For c0 = 0 To 4
            .BMD(c0) = IIf(.BMD(c0) <> -1, .BMD(c0) * CDbl(keisu / 10000), -1)
        Next
    End With

    'BMD����
'    If WfBMDJudg(bm, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
    If WfBMDJudg(bm, ErrInfo, bmflg) <> FUNCTION_RETURN_SUCCESS Then    'BMDno�ǉ��@09/05/07 ooba
        bJudg = False
        EpBmdJudg = False
        Exit Function
    End If
    
    If bm.JudgBmd <> True Or bm.JudgAntnp <> True Then
        bJudg = False
    End If
    
    typ_y022.MESDATA6 = bm.JudgDataAve
    typ_y022.MESDATA7 = bm.JudgDataMax
    typ_y022.MESDATA8 = bm.JudgDataMin
    typ_y022.MESDATA9 = bm.JudgDataMBP
    
'' 2008/10/20 BMD�]��,�O��1�_�ۏ؋@�\�ǉ� ADD By Systech Start
    If bm.GuaranteeBmd.cObj = ObjCode18 Then
        If bm.BMD(0) <> -1 Then
            typ_y022.MESDATA9 = bm.BMD(0)
        End If
    End If
'' 2008/10/20 BMD�]��,�O��1�_�ۏ؋@�\�ǉ� ADD By Systech End
     
    EpBmdJudg = True
End Function

'�T�v      :OSF(EP)����
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :�d�l���\����
'          :typ_y022      ,I  ,typ_TBCMY022                         :OSF���э\����
'          :bJudg         ,O  ,Boolean                              :���茋��(0:����OK, 1:����NG)
'          :osfflg        ,I  ,Integer                              :OSF�׸�(1:OSF1, 2:OSF2, 3:OSF3)
'          :�߂�l        ,O  ,Boolean                              :True:����I��, False:�ُ�I��
'����      :OSF������s��(�֐�WfCrOsfJudg����ɍ쐬)
'����      :�V�K�쐬 2006/08/15 �G�s��s�]���ǉ��Ή� SMP)kondoh
Public Function EpOsfJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y022 As typ_TBCMY022, _
                            bJudg As Boolean, _
                            osfflg As Integer, _
                            TmpData() As String) As Boolean

    Dim ErrInfo As ERROR_INFOMATION         '�G���[���\����
    Dim os      As W_OSF                    'OSF�\����
    Dim keisu   As Double
    Dim c0      As Integer
    
    Const keisu1 As Double = 1.8248175
    Const keisu2 As Double = 1.8518519
    Const keisu3 As Double = 1.9230769
    Const keisu4 As Double = 3.649635
    Const keisu5 As Double = 3.7037037
    Const keisu6 As Double = 3.8461538
    Const keisu7 As Double = 7.6923077
        
    bJudg = True

    'OSF(EP)��������ݒ�
    Select Case osfflg
    Case 1
        os.GuaranteeOsf.cMeth = typ_si.HEPOF1SH     '�iEPOSF1����ʒu_��
        os.GuaranteeOsf.cCount = typ_si.HEPOF1ST    '�iEPOSF1����ʒu_�_
        os.GuaranteeOsf.cPos = typ_si.HEPOF1SR      '�iEPOSF1����ʒu_��
        os.GuaranteeOsf.cObj = typ_si.HEPOF1HT      '�iEPBMD1�ۏؕ��@_��
        os.GuaranteeOsf.cJudg = typ_si.HEPOF1HS     '�iEPBMD1�ۏؕ��@_��
        os.SpecOsfAveMax = typ_si.HEPOF1AX          '�iEPOSF1���Ϗ��
        os.SpecOsfMax = typ_si.HEPOF1MX             '�iEPOSF1���
        os.JudgDataPTK = NtoS(typ_si.HEPOSF1PTK)    '�iEPOSF1���݋敪
        os.Antnp = typ_si.HEPANTNP                  '�iEPAN���x
    Case 2
        os.GuaranteeOsf.cMeth = typ_si.HEPOF2SH     '�iEPOSF2����ʒu_��
        os.GuaranteeOsf.cCount = typ_si.HEPOF2ST    '�iEPOSF2����ʒu_�_
        os.GuaranteeOsf.cPos = typ_si.HEPOF2SR      '�iEPOSF2����ʒu_��
        os.GuaranteeOsf.cObj = typ_si.HEPOF2HT      '�iEPBMD2�ۏؕ��@_��
        os.GuaranteeOsf.cJudg = typ_si.HEPOF2HS     '�iEPBMD2�ۏؕ��@_��
        os.SpecOsfAveMax = typ_si.HEPOF2AX          '�iEPOSF2���Ϗ��
        os.SpecOsfMax = typ_si.HEPOF2MX             '�iEPOSF2���
        os.JudgDataPTK = NtoS(typ_si.HEPOSF2PTK)    '�iEPOSF2���݋敪
        os.Antnp = typ_si.HEPANTNP                  '�iEPAN���x
    Case 3
        os.GuaranteeOsf.cMeth = typ_si.HEPOF3SH     '�iEPOSF3����ʒu_��
        os.GuaranteeOsf.cCount = typ_si.HEPOF3ST    '�iEPOSF3����ʒu_�_
        os.GuaranteeOsf.cPos = typ_si.HEPOF3SR      '�iEPOSF3����ʒu_��
        os.GuaranteeOsf.cObj = typ_si.HEPOF3HT      '�iEPBMD3�ۏؕ��@_��
        os.GuaranteeOsf.cJudg = typ_si.HEPOF3HS     '�iEPBMD3�ۏؕ��@_��
        os.SpecOsfAveMax = typ_si.HEPOF3AX          '�iEPOSF3���Ϗ��
        os.SpecOsfMax = typ_si.HEPOF3MX             '�iEPOSF3���
        os.JudgDataPTK = NtoS(typ_si.HEPOSF3PTK)    '�iEPOSF3���݋敪
        os.Antnp = typ_si.HEPANTNP                  '�iEPAN���x
    End Select
    
    If os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "3" Then
        keisu = keisu1
    ElseIf os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "5" Then
        keisu = keisu2
    ElseIf os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu3
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "3" Then
        keisu = keisu4
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "5" Then
        keisu = keisu5
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu6
    ElseIf os.GuaranteeOsf.cMeth = "E" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu7
    Else
        bJudg = False
        EpOsfJudg = False
        Exit Function
    End If

    With os
        .OSF(0) = NtoZ2(typ_y022.MESDATA1)                   'OSF����l
        .OSF(1) = NtoZ2(typ_y022.MESDATA2)                   'OSF����l
        .OSF(2) = NtoZ2(typ_y022.MESDATA3)                   'OSF����l
        .OSF(3) = NtoZ2(typ_y022.MESDATA4)                   'OSF����l
        .OSF(4) = NtoZ2(typ_y022.MESDATA5)                   'OSF����l
        .OsfAntnp = NtoZ2(Mid(typ_y022.DKAN, 3, 4))
        For c0 = 0 To 4
            .OSF(c0) = IIf(.OSF(c0) <> -1, .OSF(c0) * keisu, -1)
        Next
        typ_y022.MESDATA6 = typ_y022.MESDATA6 * 100
        .OSFp(0) = Trim(typ_y022.MESDATA9)                   'OSF�p�^�[������(��)
        .OSFp(1) = Trim(typ_y022.MESDATA12)                  'OSF�p�^�[������(��)
        .OSFp(2) = Trim(typ_y022.MESDATA15)                  'OSF�p�^�[������(��)
    End With
    typ_y022.MESDATA6 = typ_y022.MESDATA6 * 100
    
    'OSF����
    If WfOSFJudg(os, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        EpOsfJudg = False
        Exit Function
    End If
    
    If os.JudgOsf <> True Or os.JudgAntnp <> True Then
        bJudg = False
    End If
    
    TmpData(0) = os.JudgDataAve
    TmpData(1) = os.JudgDataMax
     EpOsfJudg = True
End Function

'�T�v      :�G�s���ђl�̑���������s���B
'���Ұ�    :�ϐ���          ,IO ,�^             :����
'          :sKeyID          ,I  ,String         :SXL-ID
'          :tNew_Hinban     ,I  ,String         :�U�֌��i��
'          :bTotalJudg      ,O  ,Boolean        :�g�[�^������
'          :iErr_Code       ,O  ,Integer        :�װ����(�߂�l�Ɠ���)
'          :sErr_Msg        ,O  ,String         :�װү���޺���
'          :typ_CType       ,O  ,typ_AllTypesC  :�S���\����(�\����)
'          :typ_CType_EP    ,O  ,typ_AllTypesC_EP  :�S���\����(�\����)(�G�s�p)
'          :iSmpGetFlg      ,I  ,Integer        :����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
'          :sSamplID1       ,I  ,String         :TOP�����ID(�ȗ���)
'          :sSamplID2       ,I  ,String         :BOT�����ID(�ȗ���)
'          :iKcnt           ,I  ,Integer        :�H���A��(�ȗ���)
'          :�߂�l          ,O  ,Integer        :�擾�̐���(0:����I��, -1:�ُ�I��)
'����      :
'����      :2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh

Public Function funWfcSogoHantei_EP(sKeyID As String, tNew_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_CType As typ_AllTypesC, typ_CType_EP As typ_AllTypesC_EP, _
                iSmpGetFlg As Integer, Optional sSamplID1 As String = vbNullString, Optional sSamplID2 As String = vbNullString, _
                Optional iKcnt As Integer = 0) As Integer

    Dim i       As Integer      'Add 2011/03/09 SMPK Miyata

    On Error GoTo Apl_down
    
    '�߂�l������
    funWfcSogoHantei_EP = FUNCTION_RETURN_FAILURE
    
    '�O���[�o���ϐ��ɐݒ�
    wiSmpGetFlg = iSmpGetFlg
    wiKcnt = iKcnt
    
    '�����ݒ�
    sErr_Msg = "WFC��������(SetInitData)"
    If SetInitData(sKeyID, tNew_Hinban, typ_CType.typ_Param, iSmpGetFlg, sSamplID1, sSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '��ʏ��ݒ�
    sErr_Msg = "WFC��������(�G�s)(SetAllData_EP)"
    If SetAllData_EP(typ_CType, typ_CType_EP, tNew_Hinban, iSmpGetFlg, iErr_Code, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
        
    TotalJudg = True
    MidlJudg = True             '���Ԕ�������   Add 2011/03/09 SMPK Miyata

    '�d�l�����x���擾
    sErr_Msg = "WFC��������(�G�s)(SpecJudgCheck)"
    SpecJudgCheck

'''    '2003/12/13 SystemBrain Null�Ή��ǉ���
'''    '�d�lNull�`�F�b�N
'''    sErr_Msg = "�d�lNull����(�G�s)"
'''    If funWfChkNull(typ_CType.typ_si, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
'''        GoTo Apl_down
'''    End If
'''    '2003/12/13 SystemBrain Null�Ή��ǉ���

    '���уf�[�^����(TOP)
    sErr_Msg = "WFC��������(�G�s)(����(TOP))"
    If EPJudge(typ_CType, typ_CType_EP, tNew_Hinban, SxlTop) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

    '���уf�[�^����(TAIL)
    sErr_Msg = "WFC��������(�G�s)(����(TAIL))"
    If EPJudge(typ_CType, typ_CType_EP, tNew_Hinban, SxlTail) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

'Add Start 2011/03/10 SMPK Miyata
    '���уf�[�^����(MIDLE)
    sErr_Msg = "WFC��������(�G�s)(����(MIDLE))"
    For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)
        If EPJudge(typ_CType, typ_CType_EP, tNew_Hinban, i) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    Next i
'Add End   2011/03/10 SMPK Miyata

'Chg Start 2011/03/09 SMPK Miyata
'    bTotalJudg = TotalJudg
    bTotalJudg = TotalJudg And MidlJudg
'Chg End   2011/03/09 SMPK Miyata

    funWfcSogoHantei_EP = FUNCTION_RETURN_SUCCESS
    
'------------------------------------------ �I������  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funWfcSogoHantei_EP = -4
    iErr_Code = funWfcSogoHantei_EP
    GoTo Apl_Exit
    
End Function

'�T�v      :��ʏ��f�[�^�ݒ�(�G�s)
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_CType     ,I  ,typ_AllTypesC ,�e���\����
'          :typ_CType_EP  ,I  ,typ_AllTypesC_EP ,�e���\����
'����      :��ʏ������\���̂ɐݒ肷��
'����      :
Private Function SetAllData_EP(typ_CType As typ_AllTypesC, typ_CType_EP As typ_AllTypesC_EP, tNew_Hinban As tFullHinban, iSmpGetFlg As Integer, _
                                                    iErr_Code As Integer, sErr_Msg As String) As FUNCTION_RETURN
    
    Dim typ_in      As type_DBDRV_scmzc_fcmlc001c_In     ' DB�A�N�Z�X���͗p
    Dim fret(2)     As FUNCTION_RETURN
    Dim typ_hi()    As typ_TBCMH004
    Dim sErrMsg     As String
    Dim FuncAns     As FUNCTION_RETURN
    Dim records()   As typ_TBCMH001
'Add Start 2011/03/07 SMPK Miyata
    Dim i           As Integer      '�J�E���^
    Dim iMidNo      As Integer      '���Ԕ���No
'Add End   2011/03/07 SMPK Miyata

    SetAllData_EP = FUNCTION_RETURN_FAILURE
    
    typ_in.HIN.hinban = typ_CType.typ_Param.hinban
    typ_in.HIN.factory = typ_CType.typ_Param.factory
    typ_in.HIN.mnorevno = typ_CType.typ_Param.REVNUM
    typ_in.HIN.opecond = typ_CType.typ_Param.opecond
    typ_in.SXLID = typ_CType.typ_Param.SXLID

    With typ_CType_EP
        
        'TOP��
        sErr_Msg = "WFC��������(TOP �����ް��ݒ�)(�G�s)"
        typ_in.SAMPLEID = typ_CType.typ_Param.WFSMP(SxlTop).REPSMPLIDCW
        typ_in.WFSMP = typ_CType.typ_Param.WFSMP(SxlTop)
    
        '�]�����茋�ʎ擾
        ReDim .typ_y022top(0)
        
        sErr_Msg = "WFC��������(TOP funGet_TBCME050)"
        '' �G�s�d�l���擾
        If funGet_TBCME050(tNew_Hinban, typ_CType.typ_si, sErrMsg) <> FUNCTION_RETURN_SUCCESS Then
            '�G�s�d�l�擾���s
            SetAllData_EP = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        
        sErr_Msg = "WFC��������(TOP funGetTBCMY022_All)"
        '' �G�s����]������(���ђl)���擾(0���ł��G���[�ł͂Ȃ�)
        If funGetTBCMY022_All(typ_in, .typ_y022top()) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("EGET2", "Y022")
            SetAllData_EP = FUNCTION_RETURN_FAILURE
            Exit Function
        End If

        ' �]�����茋�ʐ���
        sErr_Msg = "WFC��������(�G�s)(TOP �]�����茋�ʐ���)"
        If SetMERInd_EP(typ_CType_EP, .typ_y022top(), SxlTop) <> True Then
            '�]�����茋�ʐ��񎸔s
            Exit Function
        End If
        '���グ�I�����ю擾
        ReDim typ_hi(0)
'��8���w���P�����������Ȃ� 2007/10/10 SETsw kubota
'        If Mid(typ_CType.typ_Param.CRYNUM, 1, 1) <> "8" Then
            sErr_Msg = "WFC��������(�G�s)(TOP ���グ�I�����ю擾)"
            If s_cmmc001db_Sql(typ_CType.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                '���グ�I�����ю擾���s
                Exit Function
            Else
                If UBound(typ_hi) <> 0 Then
                    typ_CType.typ_hage(SxlTop) = typ_hi(1)
                Else
                    '���グ�I�����ю擾���s
                    SetAllData_EP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
            End If
'        End If
    
        'TAIL��
        sErr_Msg = "WFC��������(TAIL �����ް��ݒ�)(�G�s)"
        typ_in.SAMPLEID = typ_CType.typ_Param.WFSMP(SxlTail).REPSMPLIDCW
        typ_in.WFSMP = typ_CType.typ_Param.WFSMP(SxlTail)
    
        ReDim .typ_y022tail(0)
        
        '' �G�s����]������(���ђl)���擾(0���ł��G���[�ł͂Ȃ�)
        If funGetTBCMY022_All(typ_in, .typ_y022tail()) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("EGET2", "Y022")
            SetAllData_EP = FUNCTION_RETURN_FAILURE
            Exit Function
        End If

        ' �]�����茋�ʐ���
        sErr_Msg = "WFC��������(�G�s)(TAIL �]�����茋�ʐ���)"
        If SetMERInd_EP(typ_CType_EP, .typ_y022tail(), SxlTail) <> True Then
            '�]�����茋�ʐ��񎸔s
            Exit Function
        End If

        '���グ�I�����ю擾
        ReDim typ_hi(0)
        
'��8���w���P�����������Ȃ� 2007/10/10 SETsw kubota
'        If Mid(typ_CType.typ_Param.CRYNUM, 1, 1) <> "8" Then
            sErr_Msg = "WFC��������(�G�s)(TAIL ���グ�I�����ю擾)"
            If s_cmmc001db_Sql(typ_CType.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                '���グ�I�����ю擾���s
                Exit Function
            Else
                If UBound(typ_hi) <> 0 Then
                    typ_CType.typ_hage(SxlTail) = typ_hi(1)
                Else
                    '���グ�I�����ю擾���s
                    SetAllData_EP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
            End If
'        End If

'Add Start 2011/03/10 SMPK Miyata
        For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)

            iMidNo = i - SxlMidl + 1
            
            If iMidNo > SXL_MAXSMP Then
                ' ���Ԕ����ő匏���I�[�o�[
                Exit Function
            End If

            'MIDLE��
            sErr_Msg = "WFC��������(MIDLE_" & iMidNo & " �����ް��ݒ�)(�G�s)"
            typ_in.SAMPLEID = typ_CType.typ_Param.WFSMP(i).REPSMPLIDCW
            typ_in.WFSMP = typ_CType.typ_Param.WFSMP(i)
    
            '�]�����茋�ʎ擾
            ReDim Preserve .typ_y022midl_ary(iMidNo)
            
            sErr_Msg = "WFC��������(MIDLE_" & iMidNo & " funGet_TBCME050)"
            '' �G�s�d�l���擾
            If funGet_TBCME050(tNew_Hinban, typ_CType.typ_si, sErrMsg) <> FUNCTION_RETURN_SUCCESS Then
                '�G�s�d�l�擾���s
                SetAllData_EP = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        
            sErr_Msg = "WFC��������(MIDLE_" & iMidNo & " funGetTBCMY022_All)"

            '' �G�s����]������(���ђl)���擾(0���ł��G���[�ł͂Ȃ�)
            If funGetTBCMY022_All(typ_in, .typ_y022midl_ary(iMidNo).typ_y022midl) = FUNCTION_RETURN_FAILURE Then
                sErrMsg = GetMsgStr("EGET2", "Y022")
                SetAllData_EP = FUNCTION_RETURN_FAILURE
                Exit Function
            End If

            ' �]�����茋�ʐ���
            sErr_Msg = "WFC��������(�G�s)(MIDLE_" & iMidNo & " �]�����茋�ʐ���)"
            If SetMERInd_EP(typ_CType_EP, .typ_y022midl_ary(iMidNo).typ_y022midl, i) <> True Then
                '�]�����茋�ʐ��񎸔s
                Exit Function
            End If

            '���グ�I�����ю擾
            ReDim typ_hi(0)
            
            sErr_Msg = "WFC��������(�G�s)(MIDLE_" & iMidNo & " ���グ�I�����ю擾)"
            If s_cmmc001db_Sql(typ_CType.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                '���グ�I�����ю擾���s
                Exit Function
            Else
                If UBound(typ_hi) <> 0 Then
                    typ_CType.typ_hage(i) = typ_hi(1)
                Else
                    '���グ�I�����ю擾���s
                    SetAllData_EP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
            End If
        Next i
'Add End   2011/03/10 SMPK Miyata
    End With
    
    '' �o�{�����̔��f
    sErr_Msg = "WFC��������(P+�����̔��f)"
    If DBDRV_GetTBCMH001(records(), "where UPINDNO = '" & left(SelectSxlID, 7) & "0" & Mid(SelectSxlID, 9, 1) & "'") = FUNCTION_RETURN_SUCCESS Then
    Else
        SetAllData_EP = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    SetAllData_EP = FUNCTION_RETURN_SUCCESS
End Function

'�T�v      :����]�����ʂ̃\�[�g
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_a         ,IO ,typ_AllTypesC ,�e���\����
'          :typ_y022()    ,I  ,typ_TBCMY022 ,����]�����ʏ��\����
'          :tt            ,I  ,Integer      ,TOP�ETAIL
'          :�߂�l        ,O  ,Integer      ,True:����I���@False:�ُ�I��
'����      :����]�����ʔz���DB�����������R�[�h�𐮗񂷂�
'����      :SB_WfJudg.SetMERInd����ɍ쐬
Private Function SetMERInd_EP(typ_CType_EP As typ_AllTypesC_EP, _
                          typ_y022() As typ_TBCMY022, _
                          tt As Integer) As Boolean
    Dim i As Integer
    
    With typ_CType_EP
        For i = 1 To UBound(typ_y022)
            Select Case Trim(typ_y022(i).Spec)
            Case OSEPBMD1 ' BMD1
                .typ_y022(tt, EPBMD1) = typ_y022(i)
            Case OSEPBMD2 ' BMD2
                .typ_y022(tt, EPBMD2) = typ_y022(i)
            Case OSEPBMD3 ' BMD3
                .typ_y022(tt, EPBMD3) = typ_y022(i)
            Case OSEPOSF1 ' OSF1
                .typ_y022(tt, EPOSF1) = typ_y022(i)
            Case OSEPOSF2 ' OSF2
                .typ_y022(tt, EPOSF2) = typ_y022(i)
            Case OSEPOSF3 ' OSF3
                .typ_y022(tt, EPOSF3) = typ_y022(i)
            Case OSEPOT2 ' OT2
                .typ_y022(tt, EPOT2) = typ_y022(i)
            End Select
        Next
    End With
    SetMERInd_EP = True
End Function

'�T�v      :��������(�G�s)
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :typ_CType     ,I  ,typ_AllTypesC    ,�e���\����
'          :typ_CType_EP  ,I  ,typ_AllTypesC_EP ,�e���\����(�G�s)
'          :tNew_Hinban   ,I  ,tFullHinban      :�U�֌��i��
'          :tt            ,I  ,Integer          ,TopTail����p
'����      :�����w���ɏ]���A�G�s���є�����s��
'����      :
Public Function EPJudge(typ_CType As typ_AllTypesC, typ_CType_EP As typ_AllTypesC_EP, tNew_Hinban As tFullHinban, tt As Integer) As FUNCTION_RETURN
    
    Dim IND         As String * 4                  '�����w��
    Dim bJudg       As Boolean
    Dim i           As Integer
    Dim typTmList() As typ_TBCMB005
'Chg Start 2011/03/10 SMPK Miyata
'    Dim INGOTPOS(2) As Integer
    Dim INGOTPOS(SXL_MAXSMP) As Integer
'Chg End   2011/03/10 SMPK Miyata
    Dim vTemp       As Variant
    Dim sHinban12   As String                               '�i��(12��)
    Dim sSxlPos     As String       'SXL�ʒu(TOP/BOT)�@04/04/12 ooba

    i = 0
    EPJudge = FUNCTION_RETURN_FAILURE
    
    sHinban12 = tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond
    
    If tt = SxlTop Then
        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS
'Chg Start 2011/03/10 SMPK Miyata
'    Else
'        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
    ElseIf tt = SxlTail Then
        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
    Else
        INGOTPOS(tt) = typ_CType.typ_Param.WFSMP(tt).INPOSCW
'Chg End   2011/03/10 SMPK Miyata
    End If
    
    '�����w���ݒ�
    If tt = SxlTop Then
        IND = "123"
    Else
        IND = "123"
    End If
    
'Chg Start 2011/03/10 SMPK Miyata
'    sSxlPos = IIf(tt = SxlTop, "TOP", "BOT")        '04/04/12 ooba
    Select Case tt
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/10 SMPK Miyata

    '�����R�[�h���X�g�擾
    If GetCodeList(MSYSCLASS, KCLASS, typTmList()) <> FUNCTION_RETURN_SUCCESS Then
        '�����R�[�h���X�g�擾���s
        Exit Function
    End If
    
        '' ���������w��(B1)*****************************************************************
        BMDDataSet_EP 1, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(B2)*****************************************************************
        BMDDataSet_EP 2, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(B3)*****************************************************************
        BMDDataSet_EP 3, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(L1)*****************************************************************
        OSFDataSet_EP 1, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(L2)*****************************************************************
        OSFDataSet_EP 2, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' ���������w��(L3)*****************************************************************
        OSFDataSet_EP 3, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        
    EPJudge = FUNCTION_RETURN_SUCCESS
End Function

Public Sub BMDDataSet_EP(BmdNo As Integer, UpDo As Integer, INGOTPOS As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)

    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4               '�����w��
    Dim typ_y022z       As typ_TBCMY022
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim SijiUmu         As String
    Dim EPBmSokuP       As String                   ' ����ʒu�Q�_
    Dim EPBMD           As Integer
    Dim sSxlPos         As String                   'SXL�ʒu(TOP/BOT)

    '�����w���ݒ�
    IND = IIf(UpDo = SxlTop, "123", "123")
    
'Chg Start 2011/03/10 SMPK Miyata
'    sSxlPos = IIf(UpDo = SxlTop, "TOP", "BOT")
    Select Case UpDo
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/10 SMPK Miyata

    With typ_CType_EP
        
        Select Case BmdNo
        Case 1
            EPBMD = EPBMD1
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B1E And CheckKHN_EP(typ_CType.typ_si.HEPBM1KN, 1, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B1E And CheckKHN_EP(typ_CType.typ_si.HEPBM1KN, 1, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.B1E And typ_CType.typ_si.MSMPFLGEPBM = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B1E And typ_CType.typ_si.MSMPFLGEPBM = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B1E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDB1CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESB1CW
            EPBmSokuP = typ_CType.typ_si.HEPBM1ST
        Case 2
            EPBMD = EPBMD2
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B2E And CheckKHN_EP(typ_CType.typ_si.HEPBM2KN, 2, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B2E And CheckKHN_EP(typ_CType.typ_si.HEPBM2KN, 2, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.B2E And typ_CType.typ_si.MSMPFLGEPBM = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B2E And typ_CType.typ_si.MSMPFLGEPBM = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B2E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDB2CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESB2CW
            EPBmSokuP = typ_CType.typ_si.HEPBM2ST
        Case 3
            EPBMD = EPBMD3
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B3E And CheckKHN_EP(typ_CType.typ_si.HEPBM3KN, 3, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B3E And CheckKHN_EP(typ_CType.typ_si.HEPBM3KN, 3, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.B3E And typ_CType.typ_si.MSMPFLGEPBM = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B3E And typ_CType.typ_si.MSMPFLGEPBM = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B3E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDB3CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESB3CW
            EPBmSokuP = typ_CType.typ_si.HEPBM3ST
        End Select
            typ_y022z = .typ_y022(UpDo, EPBMD)
    
    
        '' EP�����w���iBMDE)*****************************************************************
        If JudgSpecCode Then
            '��ʕ\�����e�����ݒ�
            .typ_rslt(UpDo, DispLineCount).pos = -1                                             ' �������J�n�ʒu
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())              ' ���e
            .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                                     ' ���1
            .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                                     ' ���2
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                           ' ���3
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                           ' ���4
            .typ_rslt(UpDo, DispLineCount).INFO5 = ""                                           ' ���5
            .typ_rslt(UpDo, DispLineCount).INFO6 = ""                                           ' ���6
            .typ_rslt(UpDo, DispLineCount).INFO7 = ""                                           ' ���7
            .typ_rslt(UpDo, DispLineCount).INFO8 = ""                                           ' ���8
            .typ_rslt(UpDo, DispLineCount).SMPLID = -1                                          ' �T���v���m��
            .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                          ' ���茋��
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                                   ' �i��(12��)
            bJudg = False
            If shiji <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                                 ' ���2
                .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                                 ' ���3
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y022z.SAMPLEID                      ' �T���v���m��
                    
                'BMDE����
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y022z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'BMDE���莸�s
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                            ' ���Q
                    'BMDE����
                    If EpBmdJudg(typ_CType.typ_si, typ_y022z, bJudg, BmdNo) Then
                        vTemp = CVar(typ_y022z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���1
                        vTemp = CVar(typ_y022z.MESDATA7)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' ���2
                        vTemp = CVar(typ_y022z.MESDATA8)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' ���R
                        vTemp = CVar(typ_y022z.MESDATA9)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.00")   ' ���S
                        JiltusekiUmu(UpDo, EPBMD) = True
                        '5�Ԗڂ̏��FAN���x
                        vTemp = CVar(typ_y022z.DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        .typ_rslt(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")      ' ���5
                    End If
                ElseIf SijiUmu = "2" Then
                    bJudg = False
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����وُ�"                          ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                End If
                
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' ���茋��
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                      ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If UpDo = SxlTop Or UpDo = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case BmdNo
                Case 1
                    gsTbcmy028ErrCode = "00152"
                Case 2
                    gsTbcmy028ErrCode = "00153"
                Case 3
                    gsTbcmy028ErrCode = "00154"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())          ' ���e
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                                 ' ���1
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                                 ' ���2
                .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                                 ' ���3
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                       ' ���4
                .typ_rslt(UpDo, DispLineCount).INFO5 = ""                                       ' ���5
                .typ_rslt(UpDo, DispLineCount).INFO6 = ""                                       ' ���6
                .typ_rslt(UpDo, DispLineCount).INFO7 = ""                                       ' ���7
                .typ_rslt(UpDo, DispLineCount).INFO8 = ""                                       ' ���8
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y022z.SAMPLEID                      ' �T���v���m��
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' ���茋��
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                               ' �i��(12��)
                If (SijiUmu = "1") And (Trim(typ_y022z.SAMPLEID) <> "0") Then
                    'BMDE���莸�s
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                            ' ���2
                    'BMD����
                    If EpBmdJudg(typ_CType.typ_si, typ_y022z, bJudg, BmdNo) Then
                        '��ʕ\�����e�ݒ�@�@2003/05/20 ooba
                        vTemp = CVar(typ_y022z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���1
                        vTemp = CVar(typ_y022z.MESDATA7)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' ���2
                        vTemp = CVar(typ_y022z.MESDATA8)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' ���3
                        vTemp = CVar(typ_y022z.MESDATA9)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.00")   ' ���4
                        JiltusekiUmu(UpDo, EPBMD) = True
                        vTemp = CVar(typ_y022z.DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        .typ_rslt(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")      ' ���5
                    End If
                ElseIf SijiUmu = "2" Then
                    bJudg = False
                    '��ʕ\�����e�ݒ�
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����وُ�"                          ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                End If
                'Add Start 2011/11/28 Y.Hitomi ���Ԕ����̏ꍇ�́A�Q�l�\������
                If sSxlPos = "MID" And bJudg = False Then
                    If BmdNo = 1 And JudgSW.B1E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    ElseIf BmdNo = 2 And JudgSW.B2E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    ElseIf BmdNo = 3 And JudgSW.B3E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    End If
                End If
                'Add End  2011/11/28 Y.Hitomi
                DispLineCount = DispLineCount + 1
            End If
        End If
    
        Select Case BmdNo
        Case 1
            .typ_y022(UpDo, EPBMD1) = typ_y022z
        Case 2
            .typ_y022(UpDo, EPBMD2) = typ_y022z
        Case 3
            .typ_y022(UpDo, EPBMD3) = typ_y022z
        End Select
    
    End With
    
End Sub

Public Sub OSFDataSet_EP(OsfNo As Integer, UpDo As Integer, INGOTPOS As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)

    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                   '�����w��
''    Dim typ_y013z       As typ_TBCMY013
    Dim typ_y022z       As typ_TBCMY022
    Dim AveMax(1)       As String                       '����/�ő唻��l
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim SijiUmu         As String
    Dim EPBmSokuP       As String                       ' ����ʒu�Q�_
    Dim EPBmSokuHou     As String                       ' �iWFOSF1����ʒu_��
    Dim EPBmSokuRyou    As String                       ' �iWFOSF1����ʒu_��
    Dim EPOSF           As Integer
    Dim sSxlPos         As String                       'SXL�ʒu(TOP/BOT)

    '�����w���ݒ�
    IND = IIf(UpDo = SxlTop, "123", "123")

'Chg Start 2011/03/10 SMPK Miyata
'    sSxlPos = IIf(UpDo = SxlTop, "TOP", "BOT")
    Select Case UpDo
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/10 SMPK Miyata

    With typ_CType_EP
        Select Case OsfNo
        Case 1
            EPOSF = EPOSF1
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L1E And CheckKHN_EP(typ_CType.typ_si.HEPOF1KN, 3, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L1E And CheckKHN_EP(typ_CType.typ_si.HEPOF1KN, 3, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.L1E And typ_CType.typ_si.MSMPFLGEPOF = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L1E And .typ_si.MSMPFLGEPBM = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L1E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDL1CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESL1CW
            EPBmSokuHou = typ_CType.typ_si.HEPOF1SH
            EPBmSokuP = typ_CType.typ_si.HEPOF1ST
            EPBmSokuRyou = typ_CType.typ_si.HEPOF1SR
        Case 2
            EPOSF = EPOSF2
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L2E And CheckKHN_EP(typ_CType.typ_si.HEPOF2KN, 4, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L2E And CheckKHN_EP(typ_CType.typ_si.HEPOF2KN, 4, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.L2E And typ_CType.typ_si.MSMPFLGEPOF = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L2E And .typ_si.MSMPFLGEPBM = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L2E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDL2CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESL2CW
            EPBmSokuHou = typ_CType.typ_si.HEPOF2SH
            EPBmSokuP = typ_CType.typ_si.HEPOF2ST
            EPBmSokuRyou = typ_CType.typ_si.HEPOF2SR
        Case 3
            EPOSF = EPOSF3
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L3E And CheckKHN_EP(typ_CType.typ_si.HEPOF3KN, 5, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L3E And CheckKHN_EP(typ_CType.typ_si.HEPOF3KN, 5, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '�ۏؕ��@=�ۏ� ���� ���Ԕ����i�ۏ؁j�̏ꍇ�A�d�l�L�Ƃ���
                JudgSpecCode = (JudgSW.L3E And typ_CType.typ_si.MSMPFLGEPOF = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L3E And .typ_si.MSMPFLGEPBM = "1")
'Chg End�@ 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L3E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDL3CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESL3CW
            EPBmSokuHou = typ_CType.typ_si.HEPOF3SH
            EPBmSokuP = typ_CType.typ_si.HEPOF3ST
            EPBmSokuRyou = typ_CType.typ_si.HEPOF3SR
        End Select
        typ_y022z = .typ_y022(UpDo, EPOSF)


        '' WF�����w���iOSFE)*****************************************************************
        If JudgSpecCode Then
            '��ʕ\�����e�ݒ�
            .typ_rslt(UpDo, DispLineCount).pos = -1                                             ' �������J�n�ʒu
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())              ' ���e
            .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l�L"                                     ' ���1
            .typ_rslt(UpDo, DispLineCount).INFO2 = "������"                                     ' ���2
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                           ' ���3
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                           ' ���4
            .typ_rslt(UpDo, DispLineCount).INFO5 = ""                                           ' ���5
            .typ_rslt(UpDo, DispLineCount).INFO6 = ""                                           ' ���6
            .typ_rslt(UpDo, DispLineCount).INFO7 = ""                                           ' ���7
            .typ_rslt(UpDo, DispLineCount).INFO8 = ""                                           ' ���8
            .typ_rslt(UpDo, DispLineCount).SMPLID = -1                                          ' �T���v���m��
            .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                          ' ���茋��
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                                   ' �i��(12��)
            bJudg = False
            If shiji <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                                 ' ���2
                .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                                 ' ���3
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y022z.SAMPLEID                      ' �T���v���m��
                'OSF����擾
                If (SijiUmu = "1") And (Trim(typ_y022z.SAMPLEID) <> "0") Then
                    'OSF���莸�s
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                            ' ���2
                    'OSF����擾
                    If EpOsfJudg(typ_CType.typ_si, typ_y022z, bJudg, OsfNo, AveMax()) Then             ' AveMax
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(AveMax(0))
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���1
                        vTemp = CVar(AveMax(1))
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' ���2
                        vTemp = CVar(typ_y022z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' ���3
                        vTemp = CVar(IIf(Trim(typ_y022z.MESDATA9) = "", "-", Trim(typ_y022z.MESDATA9)) _
                                        & "  " & IIf(Trim(typ_y022z.MESDATA12) = "", "-", Trim(typ_y022z.MESDATA12)) _
                                        & "  " & IIf(Trim(typ_y022z.MESDATA15) = "", "-", Trim(typ_y022z.MESDATA15)))
                        .typ_rslt(UpDo, DispLineCount).INFO4 = vTemp                            ' ���4
                        JiltusekiUmu(UpDo, EPOSF) = True
                        vTemp = CVar(typ_y022z.DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        .typ_rslt(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")      ' ���5
                    End If
                ElseIf SijiUmu = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����وُ�"                          ' ���R
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���S
                End If
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' ���茋��
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                      ' ���茋��
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If UpDo = SxlTop Or UpDo = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case OsfNo
                Case 1
                    gsTbcmy028ErrCode = "00149"
                Case 2
                    gsTbcmy028ErrCode = "00150"
                Case 3
                    gsTbcmy028ErrCode = "00151"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji <> 0 Then
                '��ʕ\�����e�ݒ�
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' �������J�n�ʒu
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())          ' ���e
                .typ_rslt(UpDo, DispLineCount).INFO1 = "�d�l��"                                 ' ���1
                .typ_rslt(UpDo, DispLineCount).INFO2 = "�����L"                                 ' ���2
                .typ_rslt(UpDo, DispLineCount).INFO3 = "���і�"                                 ' ���3
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                       ' ���4
                .typ_rslt(UpDo, DispLineCount).INFO5 = ""                                       ' ���5
                .typ_rslt(UpDo, DispLineCount).INFO6 = ""                                       ' ���6
                .typ_rslt(UpDo, DispLineCount).INFO7 = ""                                       ' ���7
                .typ_rslt(UpDo, DispLineCount).INFO8 = ""                                       ' ���8
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y022z.SAMPLEID                      ' �T���v���m��
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' ���茋��
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                               ' �i��(12��)
                'OSF����擾
                If (SijiUmu = "1") And (Trim(typ_y022z.SAMPLEID) <> "0") Then
                    'OSF���莸�s
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����Err"                            ' ���2
                    'OSF����擾
                    If EpOsfJudg(typ_CType.typ_si, typ_y022z, bJudg, OsfNo, AveMax()) Then             ' AveMax
                        '��ʕ\�����e�ݒ�
                        vTemp = CVar(AveMax(0))
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' ���1
                        vTemp = CVar(AveMax(1))
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' ���2
                        vTemp = CVar(typ_y022z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' ���3
                        vTemp = CVar(IIf(Trim(typ_y022z.MESDATA9) = "", "-", Trim(typ_y022z.MESDATA9)) _
                                        & "  " & IIf(Trim(typ_y022z.MESDATA12) = "", "-", Trim(typ_y022z.MESDATA12)) _
                                        & "  " & IIf(Trim(typ_y022z.MESDATA15) = "", "-", Trim(typ_y022z.MESDATA15)))
                        .typ_rslt(UpDo, DispLineCount).INFO4 = vTemp                            ' ���4
                         JiltusekiUmu(UpDo, EPOSF) = True
                        vTemp = CVar(typ_y022z.DKAN)
                        '3�`6���ڂ�AN���x
                        vTemp = Mid(vTemp, 3, 4)
                        .typ_rslt(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")      ' ���5
                    End If
                ElseIf SijiUmu = "2" Then
                    '��ʕ\�����e�ݒ�
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "����وُ�"                          ' ���3
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' ���4
                End If
                'Add Start 2011/11/28 Y.Hitomi ���Ԕ����̏ꍇ�́A�Q�l�\������
                If sSxlPos = "MID" And bJudg = False Then
                    If OsfNo = 1 And JudgSW.L1E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    ElseIf OsfNo = 2 And JudgSW.L2E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    ElseIf OsfNo = 3 And JudgSW.L3E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "�Q�l"                                      ' ���茋��
                        MidlJudg = False
                    End If
                End If
                'Add End  2011/11/28 Y.Hitomi
                DispLineCount = DispLineCount + 1
            End If
        End If

        Select Case OsfNo
        Case 1
            .typ_y022(UpDo, EPOSF1) = typ_y022z
            TmpOsfData(0, UpDo, EPOSF1) = AveMax(0)
            TmpOsfData(1, UpDo, EPOSF1) = AveMax(1)
            TmpOsfMBNP(0, UpDo, EPOSF1) = IIf(typ_y022z.MESDATA9 = "-", " ", typ_y022z.MESDATA9)
            TmpOsfMBNP(1, UpDo, EPOSF1) = IIf(typ_y022z.MESDATA12 = "-", " ", typ_y022z.MESDATA12)
            TmpOsfMBNP(2, UpDo, EPOSF1) = IIf(typ_y022z.MESDATA15 = "-", " ", typ_y022z.MESDATA15)
        Case 2
            .typ_y022(UpDo, WFOSF2) = typ_y022z
            TmpOsfData(0, UpDo, EPOSF2) = AveMax(0)
            TmpOsfData(1, UpDo, EPOSF2) = AveMax(1)
            TmpOsfMBNP(0, UpDo, EPOSF2) = IIf(typ_y022z.MESDATA9 = "-", " ", typ_y022z.MESDATA9)
            TmpOsfMBNP(1, UpDo, EPOSF2) = IIf(typ_y022z.MESDATA12 = "-", " ", typ_y022z.MESDATA12)
            TmpOsfMBNP(2, UpDo, EPOSF2) = IIf(typ_y022z.MESDATA15 = "-", " ", typ_y022z.MESDATA15)
        Case 3
            .typ_y022(UpDo, EPOSF3) = typ_y022z
            TmpOsfData(0, UpDo, EPOSF3) = AveMax(0)
            TmpOsfData(1, UpDo, EPOSF3) = AveMax(1)
            TmpOsfMBNP(0, UpDo, EPOSF3) = IIf(typ_y022z.MESDATA9 = "-", " ", typ_y022z.MESDATA9)
            TmpOsfMBNP(1, UpDo, EPOSF3) = IIf(typ_y022z.MESDATA12 = "-", " ", typ_y022z.MESDATA12)
            TmpOsfMBNP(2, UpDo, EPOSF3) = IIf(typ_y022z.MESDATA15 = "-", " ", typ_y022z.MESDATA15)
        End Select

    End With
End Sub

'------------------------------------------------
' �G�s����]�����ʎ擾
'------------------------------------------------

'�T�v      :�T���v���h�c����TBCMY022���������AEP��s�]�����ʂ��擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typIn         ,I  ,type_DBDRV_scmzc_fcmlc001c_In         ,���͗p
'          :records()     ,O  ,typ_TBCMY022 ,���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :SB_WfJudg_SQL.funGetTBCMY013����ɍ쐬
'����      :�V�K�쐬 2006/08/15 �G�s��s�]���ǉ��Ή� SMP)kondoh
Public Function funGetTBCMY022_All(typIn As type_DBDRV_scmzc_fcmlc001c_In, records() As typ_TBCMY022) As FUNCTION_RETURN
    
    Dim sql     As String       'SQL�S��
    Dim rs      As OraDynaset   'RecordSet
    Dim recCnt  As Long         '���R�[�h��
    Dim i       As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetTBCMY022_All"

    ''SQL��g�ݗ��Ă�
    sql = "select SAMPLEID, OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5, "
    sql = sql & "MESDATA6, MESDATA7, MESDATA8, MESDATA9, MESDATA10, MESDATA11, MESDATA12, MESDATA13, MESDATA14, MESDATA15, "
    sql = sql & "TXID, REGDATE, SENDFLAG, SENDDATE "
    sql = sql & "from TBCMY022 "
    sql = sql & "where ('" & typIn.WFSMP.EPINDB1CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDB1CW & "' and SPEC = '" & OSEPBMD1 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.EPINDB2CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDB2CW & "' and SPEC = '" & OSEPBMD2 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.EPINDB3CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDB3CW & "' and SPEC = '" & OSEPBMD3 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.EPINDL1CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDL1CW & "' and SPEC = '" & OSEPOSF1 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.EPINDL2CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDL2CW & "' and SPEC = '" & OSEPOSF2 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.EPINDL3CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDL3CW & "' and SPEC = '" & OSEPOSF3 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.WFINDOT2CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.WFSMPLIDOT2CW & "' and SPEC = '" & OSEPOT2 & "') "
    
    Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        Set rs = Nothing
        ReDim records(0)
        funGetTBCMY022_All = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
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
    Set rs = Nothing

    funGetTBCMY022_All = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    funGetTBCMY022_All = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'Add Start 2011/04/25 SMPK Miyata
'------------------------------------------------
' ���Ԕ������є���
'------------------------------------------------

'�T�v      :���Ԕ����̎��ђl������s���B
'���Ұ�    :�ϐ���          ,IO ,�^             :����
'          :sKeyID          ,I  ,String         :SXL-ID
'          :tNew_Hinban     ,I  ,String         :�U�֌��i��
'          :bTotalJudg      ,O  ,Boolean        :�g�[�^������
'          :iErr_Code       ,O  ,Integer        :�װ����(�߂�l�Ɠ���)
'          :sErr_Msg        ,O  ,String         :�װү���޺���
'          :typ_CType       ,O  ,typ_AllTypesC  :�S���\����(�\����)
'          :iSmpGetFlg      ,I  ,Integer        :����يǗ��擾�׸�(0:����َw��Ȃ�, 1:����َw�肠��)
'          :sSamplID1       ,I  ,String         :TOP�����ID(�ȗ���)
'          :sSamplID2       ,I  ,String         :BOT�����ID(�ȗ���)
'          :iKcnt           ,I  ,Integer        :�H���A��(�ȗ���)
'          :�߂�l          ,O  ,Integer        :�擾�̐���(0:����I��, -1:�ُ�I��)
'����      :
'����      :

Public Function funWfcMidleHantei(sKeyID As String, tNew_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_CType As typ_AllTypesC, _
                iSmpGetFlg As Integer, Optional sSamplID1 As String = vbNullString, Optional sSamplID2 As String = vbNullString, _
                Optional iKcnt As Integer = 0) As Integer
    
    Dim i       As Integer      'Add 2011/03/09 SMPK Miyata
    
    On Error GoTo Apl_down
    
    '�߂�l������
    funWfcMidleHantei = FUNCTION_RETURN_FAILURE
    
    '�O���[�o���ϐ��ɐݒ�
    wiSmpGetFlg = iSmpGetFlg
    wiKcnt = iKcnt

    tNew_Hinban = tMapHinG.HIN

    '�����ݒ�
    sErr_Msg = "WFC���Ԕ������є���(SetInitData)"
    If SetInitData(sKeyID, tNew_Hinban, typ_CType.typ_Param, iSmpGetFlg, sSamplID1, sSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

    '��ʏ��ݒ�
    sErr_Msg = "WFC���Ԕ������є���(SetAllData_Mid)"
    If SetAllData_Mid(typ_CType, tNew_Hinban, iSmpGetFlg, iErr_Code, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

    TotalJudg = True
    MidlJudg = True             '���Ԕ�������

    '�d�l�����w���擾
    sErr_Msg = "WFC���Ԕ������є���(SpecJudgCheck)"
    SpecJudgCheck

    '�d�lNull�`�F�b�N
    sErr_Msg = "�d�lNull����"
    If funWfChkNull(typ_CType.typ_si, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

    '���уf�[�^����(MIDLE)
    sErr_Msg = "WFC���Ԕ������є���(WfAllJudg(MIDLE))"
    
    For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)
    
        '�Y����ۯ��̒��Ԕ������H
        If typ_CType.typ_Param.WFSMP(i).INPOSCW >= tMapHinG.INPOSCS_S And _
           typ_CType.typ_Param.WFSMP(i).INPOSCW < tMapHinG.INPOSCS_E Then
        
            'WF���� (�S)
            If WfAllJudg(typ_CType, tNew_Hinban, i) = FUNCTION_RETURN_FAILURE Then
                GoTo Apl_down
            End If
        End If

    Next i

    '��ʏ��ݒ�
    sErr_Msg = "WFC���Ԕ������є���(�G�s)(SetAllData_Mid_EP)"
    If SetAllData_Mid_EP(typ_CType, typ_CType_EP, tNew_Hinban, iSmpGetFlg, iErr_Code, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

    '�d�l�����w���擾
    sErr_Msg = "WFC���Ԕ������є���(�G�s)(SpecJudgCheck)"
    SpecJudgCheck

    '���уf�[�^����(MIDLE)
    sErr_Msg = "WFC���Ԕ������є���(�G�s)(EPJudge(MIDLE))"
    
    For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)
        '�Y����ۯ��̒��Ԕ������H
        If typ_CType.typ_Param.WFSMP(i).INPOSCW >= tMapHinG.INPOSCS_S And _
           typ_CType.typ_Param.WFSMP(i).INPOSCW < tMapHinG.INPOSCS_E Then
            '��������(�G�s)
            If EPJudge(typ_CType, typ_CType_EP, tNew_Hinban, i) = FUNCTION_RETURN_FAILURE Then
                GoTo Apl_down
            End If
        End If
    Next i

    Dim iMinMidCnt      As Integer       '���Ԕ����̕K�v��
    Dim iRstMidCnt      As Integer       '���Ԕ����̌���
    
    ' ���Ԕ����i���H
    If typ_CType.typ_si.MSMPFLG = "1" Then
        '���Ԕ����̕K�v�� = (SXL��WF���� - ���Ԕ������e�l(����)) / ���Ԕ����P��(����)
        iMinMidCnt = Fix((typ_CType.typ_Param.COUNT - typ_CType.typ_si.MSMPCONSTMAI) / typ_CType.typ_si.MSMPTANIMAI)
        '�}�C�i�X�̏ꍇ�A�O�Ƃ���
        If iMinMidCnt < 0 Then iMinMidCnt = 0
        
        '���Ԕ����̌���
        iRstMidCnt = (UBound(typ_CType.typ_Param.WFSMP) - SxlMidl) + 1
        If iRstMidCnt < iMinMidCnt Then
            typ_CType.sMidErrMsg = "���Ԕ������т�����܂���B�@�d�l(" & iMinMidCnt & ") ����(" & iRstMidCnt & ")"
            MidlJudg = False
        End If
        
    End If

    bTotalJudg = TotalJudg And MidlJudg

    funWfcMidleHantei = FUNCTION_RETURN_SUCCESS
    
'------------------------------------------ �I������  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funWfcMidleHantei = -4
    iErr_Code = funWfcMidleHantei
    GoTo Apl_Exit
    
End Function


'�T�v      :��ʏ��f�[�^�ݒ�
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_CType     ,I  ,typ_AllTypesC ,�e���\����
'����      :��ʏ������\���̂ɐݒ肷��
'����      :
Private Function SetAllData_Mid(typ_CType As typ_AllTypesC, tNew_Hinban As tFullHinban, iSmpGetFlg As Integer, _
                                iErr_Code As Integer, sErr_Msg As String) As FUNCTION_RETURN
    
    Dim typ_in      As type_DBDRV_scmzc_fcmlc001c_In     ' DB�A�N�Z�X���͗p
    Dim fret(2)     As FUNCTION_RETURN
    Dim RET         As FUNCTION_RETURN
    Dim typ_hi()    As typ_TBCMH004
    Dim sErrMsg     As String
    Dim FuncAns     As FUNCTION_RETURN ''2001/12/18 S.Sano
    Dim records()   As typ_TBCMH001
    Dim i           As Integer      '�J�E���^
    Dim iMidNo      As Integer      '���Ԕ���No

    SetAllData_Mid = FUNCTION_RETURN_FAILURE
    
    typ_in.HIN = tNew_Hinban
    typ_in.SXLID = typ_CType.typ_Param.SXLID

    With typ_CType
        
        For i = SxlMidl To UBound(.typ_Param.WFSMP)
            iMidNo = i - SxlMidl + 1

            If iMidNo > SXL_MAXSMP Then
                ' ���Ԕ����ő匏���I�[�o�[
                Exit Function
            End If

            'MIDLE��
            sErr_Msg = "WFC���Ԕ������є���(MIDLE_" & iMidNo & " �����ް��ݒ�)"
            typ_in.SAMPLEID = .typ_Param.WFSMP(i).REPSMPLIDCW
            typ_in.WFSMP = typ_CType.typ_Param.WFSMP(i)

            '�]�����茋�ʎ擾
            ReDim Preserve .typ_y013midl_ary(iMidNo)
            sErr_Msg = "WFC���Ԕ������є���(MIDLE_" & iMidNo & " funWfcGetDataEtc)"
            RET = funWfcGetDataEtc(typ_in, i, tNew_Hinban, iSmpGetFlg, _
                                    .typ_si, _
                                    .typ_y013midl_ary(iMidNo).typ_y013midl, _
                                    sErrMsg)
            If RET = FUNCTION_RETURN_SUCCESS Then

                ' �]�����茋�ʐ���
                sErr_Msg = "WFC���Ԕ������є���(MIDLE_" & iMidNo & " �]�����茋�ʐ���)"
                If SetMERInd(typ_CType, .typ_y013midl_ary(iMidNo).typ_y013midl, i) <> True Then
                    '�]�����茋�ʐ��񎸔s
                    Exit Function
                End If

                '���グ�I�����ю擾
                ReDim typ_hi(0)
                sErr_Msg = "WFC���Ԕ������є���(MIDLE_" & iMidNo & " ���グ�I�����ю擾)"
                If s_cmmc001db_Sql(.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                    '���グ�I�����ю擾���s
                    Exit Function
                Else
                    If UBound(typ_hi) <> 0 Then
                        .typ_hage(i) = typ_hi(1)
                    Else
                        '���グ�I�����ю擾���s
                        SetAllData_Mid = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                End If
            Else
                SetAllData_Mid = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        Next i
    End With
    
    '' �o�{�����̔��f
    sErr_Msg = "WFC���Ԕ������є���(P+�����̔��f)"
    If DBDRV_GetTBCMH001(records(), "where UPINDNO = '" & left(SelectSxlID, 7) & "0" & Mid(SelectSxlID, 9, 1) & "'") = FUNCTION_RETURN_SUCCESS Then

    Else
        SetAllData_Mid = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    SetAllData_Mid = FUNCTION_RETURN_SUCCESS
End Function

'�T�v      :��ʏ��f�[�^�ݒ�(�G�s)
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :typ_CType     ,I  ,typ_AllTypesC ,�e���\����
'          :typ_CType_EP  ,I  ,typ_AllTypesC_EP ,�e���\����
'����      :��ʏ������\���̂ɐݒ肷��
'����      :
Private Function SetAllData_Mid_EP(typ_CType As typ_AllTypesC, typ_CType_EP As typ_AllTypesC_EP, tNew_Hinban As tFullHinban, iSmpGetFlg As Integer, _
                                                    iErr_Code As Integer, sErr_Msg As String) As FUNCTION_RETURN
    
    Dim typ_in      As type_DBDRV_scmzc_fcmlc001c_In     ' DB�A�N�Z�X���͗p
    Dim fret(2)     As FUNCTION_RETURN
    Dim typ_hi()    As typ_TBCMH004
    Dim sErrMsg     As String
    Dim FuncAns     As FUNCTION_RETURN
    Dim records()   As typ_TBCMH001
    Dim i           As Integer      '�J�E���^
    Dim iMidNo      As Integer      '���Ԕ���No

    SetAllData_Mid_EP = FUNCTION_RETURN_FAILURE
    
    typ_in.HIN = tNew_Hinban
    typ_in.SXLID = typ_CType.typ_Param.SXLID

    With typ_CType_EP
        
        For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)

            iMidNo = i - SxlMidl + 1
            
            If iMidNo > SXL_MAXSMP Then
                ' ���Ԕ����ő匏���I�[�o�[
                Exit Function
            End If

            'MIDLE��
            sErr_Msg = "WFC���Ԕ������є���(�G�s)(MIDLE_" & iMidNo & " �����ް��ݒ�)(�G�s)"
            typ_in.SAMPLEID = typ_CType.typ_Param.WFSMP(i).REPSMPLIDCW
            typ_in.WFSMP = typ_CType.typ_Param.WFSMP(i)
    
            '�]�����茋�ʎ擾
            ReDim Preserve .typ_y022midl_ary(iMidNo)
            
            sErr_Msg = "WFC��������(MIDLE_" & iMidNo & " funGet_TBCME050)"
            '' �G�s�d�l���擾
            If funGet_TBCME050(tNew_Hinban, typ_CType.typ_si, sErrMsg) <> FUNCTION_RETURN_SUCCESS Then
                '�G�s�d�l�擾���s
                SetAllData_Mid_EP = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        
            sErr_Msg = "WFC���Ԕ������є���(�G�s)(MIDLE_" & iMidNo & " funGetTBCMY022_All)"

            '' �G�s����]������(���ђl)���擾(0���ł��G���[�ł͂Ȃ�)
            If funGetTBCMY022_All(typ_in, .typ_y022midl_ary(iMidNo).typ_y022midl) = FUNCTION_RETURN_FAILURE Then
                sErrMsg = GetMsgStr("EGET2", "Y022")
                SetAllData_Mid_EP = FUNCTION_RETURN_FAILURE
                Exit Function
            End If

            ' �]�����茋�ʐ���
            sErr_Msg = "WFC���Ԕ������є���(�G�s)(MIDLE_" & iMidNo & " �]�����茋�ʐ���)"
            If SetMERInd_EP(typ_CType_EP, .typ_y022midl_ary(iMidNo).typ_y022midl, i) <> True Then
                '�]�����茋�ʐ��񎸔s
                Exit Function
            End If

            '���グ�I�����ю擾
            ReDim typ_hi(0)
            
            sErr_Msg = "WFC���Ԕ������є���(�G�s)(MIDLE_" & iMidNo & " ���グ�I�����ю擾)"
            If s_cmmc001db_Sql(typ_CType.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                '���グ�I�����ю擾���s
                Exit Function
            Else
                If UBound(typ_hi) <> 0 Then
                    typ_CType.typ_hage(i) = typ_hi(1)
                Else
                    '���グ�I�����ю擾���s
                    SetAllData_Mid_EP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
            End If
        Next i
    End With
    
    '' �o�{�����̔��f
    sErr_Msg = "WFC���Ԕ������є���(�G�s)(P+�����̔��f)"
    If DBDRV_GetTBCMH001(records(), "where UPINDNO = '" & left(SelectSxlID, 7) & "0" & Mid(SelectSxlID, 9, 1) & "'") = FUNCTION_RETURN_SUCCESS Then
    Else
        SetAllData_Mid_EP = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    SetAllData_Mid_EP = FUNCTION_RETURN_SUCCESS
End Function

'Add End   2011/04/25 SMPK Miyata

'�T�v      :�T���v���Ԃ̖������擾����
'���Ұ�    :�ϐ���        ,IO ,�^          ,����
'          :sSxlid        ,I  ,String      ,SXLID
'          :iMaisu()      ,O  ,Integer  �@ ,�T���v���Ԗ���
'          :�߂�l        ,O  ,FUNCTION_RETURN,���o�̐���
'����      :
'����      :2011/07/19 �쐬  Marushita
Public Function fncGetSmpMai(sSXLID As String, ByRef iMaisu() As Integer) As FUNCTION_RETURN
Dim rs      As OraDynaset               '���oRecordDynaset
Dim rsCnt   As Integer                  'ں��޶���
Dim sql     As String                   'SQL��
Dim i       As Integer                  'ٰ�߶���
Dim iMCnt   As Integer                  '�f�[�^����
Dim iKcnt   As Integer                  '��������
Dim iSflg   As Integer                  '�T���v���t���O
Dim sSmpId  As String                   '�T���v��ID

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "-- Function fncGetSmpMai"

    'SQL���̍쐬
    sql = "Select NVL(WFSTA,' ') AS WFSTA,  NVL(MSMPLEID,' ') AS MSMPLEID FROM TBCMY011 "
    sql = sql & "Where MSXLID = '" & sSXLID & "' "
    sql = sql & "AND   EXISTFLG = 'Y' "
    sql = sql & "ORDER BY LOTID, BLOCKSEQ "

    '�f�[�^�̒��o
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '''���o���R�[�h�����݂��Ȃ��ꍇ
    If rs.EOF Then
        ReDim iMaisu(0)                     '�z��̏�����
        fncGetSmpMai = FUNCTION_RETURN_FAILURE   '�װ�ð��
        GoTo proc_exit
    End If

    iKcnt = 0
    ReDim iMaisu(iKcnt)
    iMCnt = 0
    iSflg = 0
    sSmpId = ""
    rsCnt = rs.RecordCount                  'ں��ސ��̶��Ă����

    '�z��ɒl���Z�b�g
    rs.MoveFirst                            '�擪ں��ނɈړ�
    For i = 0 To rsCnt - 1                  'ں��ސ���ٰ��
        DoEvents
        '������SKIP
        If Trim(CStr(rs!MSMPLEID)) = "" And CStr(rs!WFSTA) = "4" Then
        Else
            '�����P�ʂ̔��f
            If Trim(CStr(rs!MSMPLEID)) <> "" And sSmpId <> Trim(CStr(rs!MSMPLEID)) Then
                '�擪�f�[�^�̔��f
                If sSmpId = "" Then
                    sSmpId = Trim(CStr(rs!MSMPLEID))
                    iMCnt = 1
                Else
                    '����T���v��ID�̃`�F�b�N(�T���v�����ς�������̓���T���v��ID���f�p����)
                    If iSflg = 0 Then
                        iMCnt = iMCnt + 1
                        sSmpId = Trim(CStr(rs!MSMPLEID))
                        iSflg = 1
                    Else
                        iMaisu(iKcnt) = iMCnt
                        iMCnt = 1
                        iKcnt = iKcnt + 1
                        ReDim Preserve iMaisu(iKcnt)
                        iSflg = 0
                    End If
                End If
            Else
                '�����T���v��ID�̔��f
                If Trim(CStr(rs!MSMPLEID)) <> "" And sSmpId = Trim(CStr(rs!MSMPLEID)) Then
                    iMCnt = iMCnt + 1
                Else
                    '�T���v��ID�Ȃ��̔��f
                    If Trim(CStr(rs!MSMPLEID)) = "" And CStr(rs!WFSTA) = "0" Then
                        '����T���v��ID�̃`�F�b�N(����T���v�����Ȃ��Ȃ����猏�����Z�b�g)
                        If iSflg = 1 Then
                            iMaisu(iKcnt) = iMCnt
                            iMCnt = 1
                            iKcnt = iKcnt + 1
                            ReDim Preserve iMaisu(iKcnt)
                            iSflg = 0
                        Else
                            iMCnt = iMCnt + 1
                        End If
                    End If
                End If
            End If
        End If
        rs.MoveNext                         '��ں��ނɈړ�
    Next
    If iMCnt > 0 Then
        iMaisu(iKcnt) = iMCnt
    End If

    fncGetSmpMai = FUNCTION_RETURN_SUCCESS   '����ð��


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

