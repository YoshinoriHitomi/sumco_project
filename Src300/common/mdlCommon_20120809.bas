Attribute VB_Name = "mdlCommon"

'///////////////////////////////////////////////////
' @(S)
'       ���ʊ֐�
'
' @(h)  mdlDWHCommon.bas ver 1.0 ( 1999.12.14 ���� �G�K )
'
'///////////////////////////////////////////////////
Option Explicit

Public Const constInch As Integer = 25    ''1����̇o�T�C�Y

''�w�i�F��Locked�̒萔��
Public Enum CtlKind
    NORMAL_CTL                      '' 0 : ���͉\�E��
    CHECK_CTL                       '' 1 : ���͕s�E�O���[
    RED_CTL                         '' 2 : ���͉\�E��
End Enum

''ү���ށ^���O�@��ʒ萔��
Public Enum MsgKind
    NORMAL_MSG                      '' 0 : �ʏ�(��ʕ\��)
    ERR_DISP                        '' 1 : ��ʕ\���װ
    ERR_LOG                         '' 2 : ���O�o�ʹװ
    ERR_DISP_LOG                    '' 3 : ��ʕ\���E���O�o�ʹװ
    DEBUG_DISP = 5                  '' 5 : ��ʕ\�����ޯ��
    DEBUG_LOG                       '' 6 : ���O�o�����ޯ��
    DEBUG_DISP_LOG                  '' 7 : ��ʕ\���E���O�o�����ޯ��
End Enum

''���������G���[��ʒ萔��
Public Enum InitKind
    NORMAL_RET                      '' 0 : ����
    EXITSUB_RET                     '' 1 : Exit Sub ���Ȃ���΂Ȃ�Ȃ�
    MAINMENU_RET                    '' 2 : GotoMainMenu ���Ȃ���΂Ȃ�Ȃ�
End Enum

''�o�͂��郁�b�Z�[�W�����i�}�X�N�j
Private Const MsgKindMask = 7       '' 0 : �ʏ�ү���ނ���ʕ\�������
                                    '' 1 : �ʏ�/��ʕ\���װ���o�͂����
                                    '' 2 : �ʏ�/���O�o�ʹװ���o�͂����
                                    '' 3 : �ʏ�/��ʕ\���װ/���O�o�ʹװ���o�͂����
                                    '' 7 : �ʏ�/��ʕ\���װ/���O�o�ʹװ/���ޯ�ނ��o�͂����

''�p�X
Private Const LogDir = "..\LOG\" ''���O�̃p�X

''�O���[�o���ϐ�
Public gobjOraSess      As Object   ''�I���N���Z�b�V�����I�u�W�F�N�g
Public gobjOraDB        As Object   ''�I���N���f�[�^�x�[�X�u�W�F�N�g

''�d�w�d�I�v�V����
Public gsFactryCd       As String   ''�H��R�[�h
Public gsCallCd         As String   ''�ďo�敪
Public gsHinban         As String   ''�i��
Public myFactryCd       As String   ''�H��R�[�h

Public gsCompName       As String   ''�R���s���[�^��
Public gsEXEName        As String   ''�d�w�d�t�@�C����

Public gbFTPFlg         As Boolean  ''FTP�]���t���O
Public mbMenuRet        As Boolean  ''���j���[�J�ڋ��t���O

'Public gsProcCode1      As String   ''  �����H���R�[�h1
'Public gsProcCode2      As String   ''  �����H���R�[�h2
'Public gsProcCode3      As String   ''  �����H���R�[�h3
'Public gsProcCode4      As String   ''  �����H���R�[�h4
'Public gsProcCode5      As String   ''  �����H���R�[�h5
''���W���[���ϐ�
Private msLogFile       As String   ''���O�t�@�C����
Private msMsgStr(100)   As String   ''���b�Z�[�W�z��
'' ========== �ϐ����� ===========
'' ���b�Z�[�W�z��́A���O�������������Ń��b�Z�[�W��������
'' ���O�o�́E��ʕ\���ɂĎg�p����B

Private Const SOKUTEI_MAX = 8       '' ��������ő�l
Private Const NULL_CHECK = 999999   '' �f�[�^�����`�F�b�N
'' 2000/04/24 �ǉ�
Public Type RRG_CALC                '' RRG�v�Z�f�[�^
    dTeikou As Double               '' ��R�l
    sRRGFlg As String               '' �v�Z�t���O
End Type
Public Type TYPE_RRG                '' RRG�Z�o��R�l�ꗗ
    iSampleNo As String             '' �T���v��No
    dTeikouDT(SOKUTEI_MAX) As RRG_CALC '' ��R�l(A�`I)
End Type

'Cs����v�Z�p���Ұ��@06/04/20 ooba
Public Type CS_SUITEI_TYPE
    sSiWeight           As String   ''����ޗ�(Kg)
    sTopWT              As String   ''į�ߏd��(Kg)
    sUpDm               As String   ''���a(mm)
    sCsHenseki          As String   ''����ݕΐ͌W��(����Ͻ��ɕێ�)
    sSamplePos          As String   ''����وʒu
    sResCs              As String   ''����ّ���l
    sInfPos             As String   ''����ʒu
End Type
    
' �X�V���
Public Enum CHANGE_TYPE
    ST_NORMAL                ' ���X�V
    ST_UPDATE                ' �X�V
    ST_INSERT                ' �ǉ�
    ST_DELETE                ' �폜
    ST_DELINS                ' �폜�ǉ�
End Enum
' Oi/Cs�e�[�u���f�[�^
Public Type ST_OICS
    sCrystalNo  As String     ' �����ԍ�
    sCryBuiNo   As String      ' ��������
    sMenPosIti  As String     ' �ʓ��ʒu
    sSampleNo   As String      ' �T���v��No
    sBuiKubun   As String      ' ���ʋ敪
    sCarbonAT   As String      ' �J�[�{��(AT)
    sSansoAT    As String       ' �_�f(AT)
    sCarbonPpma As String    ' �J�[�{��(ppma)
    sSansoPpma  As String     ' �_�f(ppma)
    sORGNo      As String         ' ORG
    sYMDData    As String       ' �������t
    sChgType    As CHANGE_TYPE  ' �ύX���
End Type
''===============================================================================

''�R���s���[�^���擾API
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpszName As String, lpcchBuffer As Long) As Long

'�H���A�ԍ쐬�p
Public wMaxKcnt As Integer
'�߂�t���O(f_cmbc039_2����F_cmbc039_1�Ŗ߂�Ƃ��ɐݒ�)
Public intModoru As Integer


'*ADD*  TCS)K.Kunori 2004.11.29 START >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    '2004/9/27tcs  yamauchi �ǉ�
    Public gsSysdate        As String   ''���ѓ��t
    
    '2004/11/12tcs tagawa �ǉ� start-------------------------------------
    Public gsSystemCd       As String   ''�V�X�e���敪(200�A300)
    
    '*DEL* �d���錾�ɂȂ��Ă���̂ō폜 TCS)K.Kunori 2004.11.29 START >>>
    '''Public Const SYSTEM_200 = "2"               ''200mm�V�X�e����
    '''Public Const SYSTEM_300 = "3"               ''300mm�V�X�e����
    '''
    '''''���і�
    '''Public Const SYSTEM_NAME_200 = "200mm�������ƃV�X�e��"
    '''Public Const SYSTEM_NAME_300 = "300mm�������ƃV�X�e��"
    '*DEL* �d���錾�ɂȂ��Ă���̂ō폜 TCS)K.Kunori 2004.11.29 END <<<
    ''2004/11/12tcs tagawa �ǉ� end--------------------------------------
    
    '*ADD* TCS)K.Kunori 2004.11.17 START >>>
    Public gobjOraSess2     As Object   ''�I���N���Z�b�V�����I�u�W�F�N�g(SQL�����O�쐬�p)
    Public gobjOraDB2       As Object   ''�I���N���f�[�^�x�[�X�u�W�F�N�g(SQL�����O�쐬�p)
    '*ADD* TCS)K.Kunori 2004.11.17 END <<<
    
    '2004/9/27tcs Suenaga �ǉ� start-------------------------------------
    Private mtrlNo  As String          '����No
    Private cryno   As String          '���i���b�gNo
    Private PROCCD  As String          '�H���R�[�h
    Private staffCd As String          '�S���҃R�[�h
    Private recW    As String          '����d��
    Private sendW   As String          '���o�d��
    Private lossW   As String          '���X�d��
    Private factCd  As String          '�H��R�[�h
    Private recCd   As String          '����H��R�[�h
    Private sendCd  As String          '���o�H��R�[�h
    Private disapp  As String          '���ŋ敪
    Private sikake  As String          '�d�|�敪
    Private sysCd   As String          '�V�X�e���敪�R�[�h
    Private conceK  As String          '�Z�x�敪
    Private conceT  As String          '�Z�x�l
    Private SENDFLG As String          '�������M�t���O
    Private occuFlg As String          '�����t���O
    Private conceM  As String          '���Z�x
    Private planFac As String          '�g�p�\��H��
    Private tanaKu  As String          '�I����敪
    '2004/9/27tcs Suenaga �ǉ� end-------------------------------------
    
    '*** UPDATE START T.TERAUCHI 2004/10/19 ���o�敪�ǉ�
    Private stowkkbb3 As String          '���o�敪
    '*** UPDATE END   T.TERAUCHI 2004/10/19
    
    Private sChgNo  As String           '�`���[�WNo(CC200/CC300�o�^�p)�@05/08/23 ooba
    
    '2004/9/17tcs Yamauchi �ǉ� start-------------------------------------
    
    ''�w�i�F
    Public Const COLOR_GRAY_SPR = &HC0C0C0      ''�D�F
    Public Const COLOR_PINK_SPR = &HFFC0FF      ''��ݸ
    Public Const COLOR_RED_SPR = &HFF&          ''��
    
    Public Const SYSTEM_200 = "2"               ''200mm�V�X�e����
    Public Const SYSTEM_300 = "3"               ''300mm�V�X�e����
    
    ''���і�
    Public Const SYSTEM_NAME_200 = "200mm�������ƃV�X�e��"
    Public Const SYSTEM_NAME_300 = "300mm�������ƃV�X�e��"
    
    ''ү���ށi�ʏ�j
    Public Const MSG0001 = "���͗��ɓ��͌�A���o�L�[������"
    Public Const MSG0002 = "���͏��Ŕp�����܂����A�X�����ł���"
    Public Const MSG0003 = "�p�����܂���"
    Public Const MSG0004 = "�S���ҁA������������͂��A���s�{�^������"
    Public Const MSG0005 = "�����A�d�ʁA�c����������͂��A���s�{�^������"
    Public Const MSG0006 = "�d�ʁA�d�|��ԁA�g�p�\��H�����͂��A���s�{�^������"
    Public Const MSG0007 = "���͏��ōX�V���܂����A�X�����ł���"
    Public Const MSG0008 = "�Ώۂ̃f�[�^��I�����A���s�L�[������"
    Public Const MSG0009 = "�X�V���܂���"
    Public Const MSG0010 = "�Ώۂ̃f�[�^��I�����A���s�L�[������"
    Public Const MSG0011 = "�\�����ׂ̕��o���s���܂��B�X�����ł��傤��"
    Public Const MSG0012 = "�ǉ����܂���"
    Public Const MSG0013 = "�\�����ׂ̐�������s���܂��B�X�����ł��傤��"
    Public Const MSG0014 = "���͏��ő��H��ɕ��o�������܂��B�X�����ł��傤��"
    Public Const MSG0015 = "���͏��̒I����/�u���ꏈ�����s���܂��B�X�����ł��傤��"
    Public Const MSG0016 = "���͏��ŕ��o���܂����A�X�����ł���"
    Public Const MSG0017 = "���͗��ɓ��͌�A���o�L�[������"
    
    ''ү���ށi�װ�p�j
    Public Const ERR0001 = "�����̍��v���u���b�N�S�̂̒����𒴂��Ă��܂�"
    Public Const ERR0002 = "�d�ʂ̍��v���u���b�N�S�̂̏d�ʂ𒴂��Ă��܂�"
    Public Const ERR0003 = "�d�ʂ̍��v�������薈�̏d�ʂ𒴂��Ă��܂�"
    Public Const ERR0004 = "���������͂���Ă��܂���"
    Public Const ERR0005 = "�d�ʂ����͂���Ă��܂���"
    Public Const ERR0006 = "�c�����������͂���Ă��܂���"
    
    '*** update start T.TERAUCHI 2004/10/20
    'Public Const ERR0007 = "���������������܂�"
    Public Const ERR0007 = "���i���b�gNo���̔Ԃ��邱�Ƃ��ł��܂���"
    '*** UPDATE END   T.TERAUCHI 2004/10/20
    
    Public Const ERR0008 = "�������b�g�ׁ̈A���̋@�\�͎g�p�ł��܂���"
    Public Const ERR0009 = "�Z�x�v�Z�̏�񂪕s�����Ă��܂�"
    Public Const ERR0010 = "�Z�x�v�Z�̏�񂪕s���ł�"
    Public Const ERR0011 = "�g�p�\��H�ꂪ�I������Ă��܂���"
    Public Const ERR0012 = "���o�H�ꂪ�I������Ă��܂���"
    Public Const ERR0013 = "�o�X�P�b�gNo�̓��͂Ɍ�肪����܂�"
    Public Const ERR0014 = "����Ԃ̓��͂Ɍ�肪����܂�"
    Public Const ERR0015 = "�����ԍ��̘A�Ԃ͂���ȏ�̔Ԃł��܂���"
    Public Const ERR0016 = "�����R�l�擾�̏�񂪕s�����Ă��܂�"
    Public Const ERR0017 = "�I�����ꂽ���b�g�́A������������邱�Ƃ͂ł��܂���"
    Public Const ERR0018 = "���͂��ꂽ�d�ʂ��s���ł�"
    Public Const ERR0019 = "���͂��ꂽ�u���b�N�������s���ł�"
    Public Const ERR0020 = "Cs�A�E�g�l�v�Z�̏�񂪕s�����Ă��܂�"
    Public Const ERR0021 = "Cs�A�E�g�l�v�Z�̏�񂪕s���ł�"
    Public Const ERR0022 = "�����ԍ���12�����͂��Ă�������"
    Public Const ERR0023 = "�u���b�N�`���I�����Ă�������"
    Public Const ERR0024 = "���C�t�^�C��10�����Z�l�v�Z�̏�񂪕s�����Ă��܂�"
    Public Const ERR0025 = "���C�t�^�C��10�����Z�l�v�Z�̏�񂪕s���ł�"
    Public Const ERR0026 = "�d�ɍނׁ̈A�ؒf�������邱�Ƃ͂ł��܂���"
    Public Const ERR0027 = "�����R���v�Z����ׂ̎����l������܂���"
    Public Const ERR0028 = "�����R���v�Z����ׂ̌�����񂪂���܂���"
    Public Const ERR0029 = "��R�l���p���Ώۂׁ̈A����������邱�Ƃ͂ł��܂���"
    Public Const ERR0030 = "���o�H�ꂪ�I������Ă���ׁA���s�������邱�Ƃ͂ł��܂���"
    Public Const ERR0031 = "�u���b�N���ʑΏۊO�ׁ̈A�Z�x�v�Z�͂ł��܂���"
    
    ''�H������
    Public Const PROCD_GENRYO_UKEIRE = "CB410"              ''�����������
    Public Const PROCD_GENRYO_SETUDAN = "CB510"             ''���������ؒf
    Public Const PROCD_ROT_KOSEI = "CB610"                  ''ۯč\��
    Public Const PROCD_GENRYO_SENJYO_UKEIRE = "CB220"       ''�������������
    Public Const PROCD_GENRYO_SENJYO_HARAIDASI = "CB225"    ''����������򕥏o
    Public Const PROCD_GENRYO_TANAIRE = "CB230"             ''�����I��
    Public Const PROCD_ZAIKO_SYUSEI = "RP10"                ''�݌ɏC��
    
    ''�~����
    Public Const CIRCULAR_CONSTANT = 3.14159265358979
    
    ''�V���R����d
    Public Const SPECIFIC_GRAVITY = 0.00233
    '2004/9/17tcs Yamauchi �ǉ� end-------------------------------------

'*ADD*  TCS)K.Kunori 2004.11.29 END <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<




'///////////////////////////////////////////////////
' @(f)
' �@�\    : �R���s���[�^���擾�E�ϐ��Z�b�g
'
' �Ԃ�l  : �R���s���[�^��
'
' ������  :
'
' �@�\����: �R���s���[�^�����擾���A�ϐ��Z�b�g
'
'///////////////////////////////////////////////////
Public Function GetCompName() As String
''>>>>> PC��20�����Ή� SETsw H.Iwamoto 2005/10/31
'    Dim sCompName As String * 9     ''���߭��������ޯ̧
    Dim sCompName As String * 20    ''���߭��������ޯ̧
''<<<<< PC��20�����Ή� SETsw H.Iwamoto 2005/10/31
    Dim lCompNameLen As Long        ''�ޯ̧���ޓn���A���߭����LEN���
    Dim bResult As Boolean          ''�擾���ʎ��
    
    ''�T�C�Y���Z�b�g
    lCompNameLen = LenB(sCompName) - 1
    ''�擾
    bResult = GetComputerName(sCompName, lCompNameLen)
    If bResult Then                 ''�擾����
        gsCompName = left(sCompName, lCompNameLen)  ''��۰��ٕϐ��ɃZ�b�g
    Else                            ''�擾���s
        gsCompName = ""                             ''��۰��ٕϐ��ɃZ�b�g
    End If

''>>>>> �抸������(2005/10/26����)��8�����Ԃ��B SETsw H.Iwamoto 2005/10/31
    gsCompName = IIf(Len(gsCompName) > 10, left(gsCompName, 8), gsCompName)
''<<<<< �抸������(2005/10/26����)��8�����Ԃ��B SETsw H.Iwamoto 2005/10/31
    
    ''�R���s���[�^����Ԃ�
    GetCompName = gsCompName
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : ���s�t�@�C�����擾�E�ϐ��Z�b�g
'
' �Ԃ�l  : ���s�t�@�C����
'
' ������  :
'
' �@�\����: ���s�t�@�C�������擾���A�ϐ��Z�b�g
'
'///////////////////////////////////////////////////
Public Function GetEXEName() As String
    gsEXEName = App.EXENAME         ''VB�̱��ع���ݵ�޼ު�Ă�����ş�ٖ��擾�E�ϐ����
    GetEXEName = Trim(gsEXEName)          ''���ş�ٖ���߂�
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �R�}���h���C�������擾�E�ϐ��Z�b�g
'
' �Ԃ�l  : True:���@False:��
'
' ������  :
'
' �@�\����: �R�}���h���C���������擾���E�g�[�N���؏o���E�ϐ��Z�b�g
'
'///////////////////////////////////////////////////
Public Function GetCmdLine() As Boolean
    Dim sCmdLine As String
    
    ''�R�}���h���C���擾
    sCmdLine = Command
    '' 0        1         2
    '' 1234567890123456789012
    ''"99_*******_***********"
    ''�H��R�[�h_�ďo�敪_�i��
    
    ''�Œ�ŃR�}���h���C��������؏o��
    gsFactryCd = left(sCmdLine, 2)    ''�H��R�[�h(2��)
    gsCallCd = Mid(sCmdLine, 4, 7)    ''�ďo�敪(7��)
    gsHinban = Mid(sCmdLine, 12, 11)  ''�i��(11��)
    myFactryCd = Mid(sCmdLine, 24, 2)
    
    If Len(gsFactryCd) <> 2 Then Exit Function
    If Len(gsCallCd) <> 7 Then Exit Function
    If gsHinban = "00000000000" Then gsHinban = ""
    GetCmdLine = True
End Function


'///////////////////////////////////////////////////
' @(f)
'
' �@�\      : �t�H�|���𒆉��ɕ\��
'
' �Ԃ�l    : �Ȃ�
'
' ������    : FrmName - �t�H�[����
'
' �@�\����  : �t�H�|���𒆉��ɕ\��
'
'///////////////////////////////////////////////////
Public Function FrmCenter(frmName As Form)
    With frmName
        If .WindowState <> 2 Then
            .left = (Screen.Width - .Width) / 2
            .top = (Screen.Height - .Height) / 2
            .Width = 12000
            .Height = 9000
        End If
    End With
End Function

'///////////////////////////////////////////////////
' @(f)
' �@�\    : �L�����Z���E�C���{�^���L���^��������
' �Ԃ�l  : �Ȃ�
' ������  : True:����L��
'           False:���䖳��
' �@�\����: �A�N�e�B�u�ɂȂ��Ă���t�H�[����
'          �L�����Z���E�C���{�^���𐧌�L���^�����ɂ���
'///////////////////////////////////////////////////
Public Sub F3F4Enabled(bEnabled As Boolean)
    Screen.ActiveForm.cmdF(3).Enabled = bEnabled
    Screen.ActiveForm.cmdF(4).Enabled = bEnabled
End Sub


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �t�H�[����̃R���g���[�����g�p�s�ɂ���
'
' �Ԃ�l  :
'
' ������  : �t�H�[��
'
' �@�\����: �w�肵���t�H�[���́u[�e�P]Ҳ��ƭ��v�{�^���ȊO
'           �̃R���g���[�����g�p�s�ɂ���
'
'///////////////////////////////////////////////////
Public Sub CtrlCancel(frmForm As Form)
    Dim iIdx As Integer
    Dim ctlControl As Control
    
    ''�t�H�[����̃R���g���[����S�Ďg�p�s�ɂ���
    For Each ctlControl In frmForm.Controls
        If TypeOf ctlControl Is TextBox Then
' Mod 2005/11/18 M.Makino CHECK_CTL -> CTRL_DISABLE_GRAY
            Call CtrlEnabled(ctlControl, CTRL_DISABLE_GRAY)
        ElseIf TypeOf ctlControl Is ComboBox Then
' Mod 2005/11/18 M.Makino CHECK_CTL -> CTRL_DISABLE_GRAY
            Call CtrlEnabled(ctlControl, CTRL_DISABLE_GRAY)
        ElseIf TypeOf ctlControl Is CommandButton Then
            ctlControl.Enabled = False
        End If
    Next ctlControl
    
    ''�u[�e�P]Ҳ��ƭ��v�{�^�����g�p�\�ɂ���
    frmForm.cmdF(1).Enabled = True
End Sub

'///////////////////////////////////////////////////
' @(f)
'
' �@�\      : ���b�Z�[�W������
'
' �Ԃ�l    : �Ȃ�
'
' ������    : �Ȃ�
'
' �@�\����  : ���b�Z�[�W�z��ݒ�
'
' ���l      :
'///////////////////////////////////////////////////
Private Sub MsgInit()
'   CSV���ڐ���
'   ү���޺���,���b�Z�[�W
    Const MsgData1 As String = _
    "01,���͗��ɓ��͌�A���s�L�[������," & _
    "02,�\�����ꂽ���e���m�F�����s�L�[������," & _
    "03,���͗��C����A���s�L�[������," & _
    "04,�I�������s�L�[������," & _
    "05,�O��ʃ{�^���Ŗ߂�܂�," & _
    "06,�s���ȕ��������͂���Ă��܂�," & _
    "07,�l�������_���ɂȂ��Ă��܂�," & _
    "08,�w�肵���S���҃R�[�h�͓o�^����Ă��܂���," & _
    "09,�w�肵�������ԍ��͓o�^����Ă��܂���," & _
    "10,�w�肵���T���v��No.�͓o�^����Ă��܂���," & _
    "11,����������܂���," & _
    "12,���͂���Ă܂���," & _
    "13,���l����͂��ĉ�����," & _
    "14,�l���O�ɂȂ��Ă��܂�," & _
    "15,�l���}�C�i�X�ɂȂ��Ă��܂�," & _
    "16,�w�肵�������ԍ��̎d�|�H�����Ⴂ�܂�," & _
    "17,�w�肵�������ԍ��̎d�|�H�����Ⴂ�܂�," & _
    "18,�˗��d�ʂ��d�|�d�ʂ𒴂��Ă��܂�," & _
    "19,�w�肵��������ނ͑������ł͂���܂���," & _
    "20,�w�肵��������ނ͓o�^����Ă��܂���,"
    Const MsgData2 As String = _
    "22,�d�|�H�����Ⴂ�܂�," & _
    "23,�z�[���h�������̌����ł�," & _
    "25,���o�����ς݂̌����ł�," & _
    "26,�����I�����s���Ă�������," & _
    "27,���s�ρF�����Ĕ��s����ꍇ�́A���͂��ĉ�����," & _
    "30,�\�����ꂽ���e���m�F����ݾق��C��������," & _
    "31,�H���R�[�h��I�����ĉ�����," & _
    "32,�W�v�敪��I�����ĉ�����," & _
    "33,�f�q�^�a�r��I�����ĉ�����," & _
    "40,�^�]�����@������c," & _
    "41,����w�����@������c," & _
    "42,���H�����[�@������c," & _
    "43,�Ċe�t���f�����@������c,"
    Const MsgData3 As String = _
    "50,���͂����f�[�^�Ɍ�肪����܂�," & _
    "51,���͂����l�Ɍ�肪����܂�," & _
    "52,���t�����������͂���Ă��܂���," & _
    "53,���t�͈͎̔w�肪����������܂���," & _
    "54,���ɓo�^�ς݂ł�," & _
    "55,�Y���f�[�^������܂���ł���," & _
    "56,�f�[�^���d�����Ă��܂�," & _
    "57,��������s���܂���," & _
    "58,�I������Ă��܂���," & _
    "60,�R���s���[�^���擾���s," & _
    "61,���O���������s," & _
    "62,���s�t�@�C�����擾���s," & _
    "63,���d�N�����܂���," & _
    "64,�����ײ݈������s�����Ă��܂�," & _
    "65,���s�t�@�C�����N���ł��܂���," & _
    "66,�������L�������𒴂��Ă��܂�," & _
    "67,�����_�����L�������𒴂��Ă��܂�," & _
    "68,���l���ő�l�𒴂��Ă��܂�," & _
    "69,���l���ŏ��l�����ł�,"
    Const MsgData4 As String = _
    "70,���R�[�h�������s," & _
    "71,���R�[�h�}�����s," & _
    "72,���R�[�h�X�V���s," & _
    "73,���R�[�h�폜���s," & _
    "100,�I���N���G���[," & _
    "00,,"
    '�f�[�^�𓝈ꉻ
    Const sMsgData As String = MsgData1 & MsgData2 & MsgData3 & MsgData4
    
    Dim iP1 As Integer
    Dim iP2 As Integer
    Dim iMsgCd As Integer
    
    iP2 = 1  '�ŏ��̈ʒu���w��
    Do
        'ү���޺��ޒ��o
        iP1 = InStr(iP2, sMsgData, ",")
        iMsgCd = val(Mid(sMsgData, iP2, iP1 - iP2))
        iP2 = iP1 + 1
        
        '���b�Z�[�W���o
        iP1 = InStr(iP2, sMsgData, ",")
        msMsgStr(iMsgCd) = Mid(sMsgData, iP2, iP1 - iP2)
        iP2 = iP1 + 1
        
'Debug.Print iMsgCd; ":"; msMsgStr(iMsgCd)
    Loop While iMsgCd   '���ނ��O�ɂȂ�܂�
End Sub


'///////////////////////////////////////////////////
' @(f)
'
' �@�\      : ���O�o�͊֘A�̏�����
'
' �Ԃ�l    : True:���@False:��
'
' ������    : LoadModule - [i]���[�h���W���[����(Kxxxxxx.EXE)
'             LogFile    - [i]���O�o�̓t�@�C����(Kxxxxxx.LOG�A�t�@�C�����̂�)
'
' �@�\����  : ���O�o�͊֘A�̏������A�v���O�����̋N������CALL����B
'///////////////////////////////////////////////////
Public Function LogInit() As Boolean
    Dim sLine As String '�s�f�[�^
    Dim FN1, FN2 '�t�@�C���ԍ�
    Dim sTmp As String 'work file
    Dim m, n
    
    On Error Resume Next
    
    ''���O�������s���Z�b�g
    LogInit = False
    
    ''���b�Z�[�W������============================
    Call MsgInit
    
    ''���O�t�@�C������۰��ٕϐ��Z�b�g
    msLogFile = LogDir & GetCompName() & ".Log" ''�R���s���[�^����۸�̧�ٖ��ɂ���
    
    ''�e���|�����t�@�C�����쐬
    m = Len(msLogFile)
    sTmp = left(msLogFile, m - 4) & ".tmp"      ''�g���q����O��".tmp"��t��
    
    ''�f�B���N�g�����݃`�F�b�N
    If Dir(LogDir, vbDirectory) = "" Then       ''���O�f�B���N�g��������
        MkDir (LogDir)                          ''���O�f�B���N�g���쐬
        If Dir(LogDir, vbDirectory) = "" Then   ''���O�f�B���N�g�����쐬�ł��Ȃ����
            GoTo Er
        End If
    End If
    
    ''���O�t�@�C���̃I�[�v��
    FN1 = FreeFile                              ''���g�p�̃t�@�C���ԍ����擾���܂�
    Err = 0
    Open msLogFile For Input As #FN1            ''���O�t�@�C���I�[�v��
    If Err <> 0 Then                            ''���O�t�@�C�����������
        LogInit = True                          ''����
        Exit Function
    End If
    
    ''�e���|�����t�@�C���폜����
    If Dir(sTmp) <> "" Then                     ''�e���|�����t�@�C�������ɑ��݂��Ă�����
        Kill sTmp                               ''�e���|�����t�@�C���폜
    End If
    
    ''�e���|�����t�@�C���I�[�v��
    FN2 = FreeFile                              ''���g�p�̃t�@�C���ԍ����擾���܂�
    Err = 0
    Open sTmp For Output As #FN2                ''�e���|�����t�@�C���I�[�v��
    If Err <> 0 Then
        Debug.Print "������̧�ق��J���Ȃ�:" & sTmp
        Close #FN1
        GoTo Er
    End If
    
    ''�ꃕ���ȓ��̃��O�̂݃e���|�����ɃR�s�[
    Do While Not EOF(FN1)                       ''�t�@�C���̏I�[�܂Ń��[�v
        Line Input #FN1, sLine                  ''���O�t�@�C���Ǎ�
        If IsDate(left(sLine, 19)) Then         ''�s�������t�Ȃ�
            If CDate(left(sLine, 19)) > CDate(DateAdd("m", -1, Now)) Then   '' 1�����O�ȓ��Ȃ�
                Print #FN2, sLine               ''�e���|�����ɏo��
            End If
        End If
    Loop
    
    Close #FN1
    Close #FN2
    Kill msLogFile                              ''���O�t�@�C���폜
    Name sTmp As msLogFile                      ''�e���|���������O�t�@�C���ɂ���
    ''���O����������Z�b�g
    LogInit = True
    On Error GoTo 0
    Exit Function
Er:
    Close #FN1
    Close #FN2
    On Error GoTo 0
End Function


'///////////////////////////////////////////////////
' @(f)
'
' �@�\      : ���b�Z�[�W���O�o��
'
' �Ԃ�l    : �Ȃ�
'
' ������    : ���b�Z�[�W
'
' �@�\����  : ���b�Z�[�W�����O�o�͂���
'
' ���l      :
'///////////////////////////////////////////////////
Private Sub MsgLog(Msg As String)
    On Error Resume Next
    Dim fno                             ''�t�@�C���ԍ�
    
    fno = FreeFile                      '' ���g�p�̃t�@�C���ԍ����擾����
    Err = 0
    Open msLogFile For Append As #fno   '' �I�[�v������
    If Err <> 0 Then
        Exit Sub
    End If
    Print #fno, Msg                     '' �o�͂���
    Close #fno                          '' ����
    On Error GoTo 0
End Sub


'///////////////////////////////////////////////////
' @(f)
'
' �@�\      : ���b�Z�[�W��ʕ\��
'
' �Ԃ�l    : �Ȃ�
'
' ������    : arg1:���b�Z�[�W
'
' �@�\����  : ���b�Z�[�W����ʕ\������
'
' ���l      :
'
'           �g�p�����F̫�я��lblMsg�Ƃ������۰ق��A
'                     �\��t���Ă��邱��
'
'///////////////////////////////////////////////////
Private Sub MsgDisp(Msg As String, Optional lForeColor As Long = 0)
    On Error Resume Next
'    Screen.ActiveForm.lblMsg.ForeColor = lForeColor
    Screen.ActiveForm.lblMsg = Msg
    Screen.ActiveForm.lblMsg.Refresh
    On Error GoTo 0
End Sub


'///////////////////////////////////////////////////
' @(f)
'
' �@�\      : �׸ٴװү���ނ��׸ٴװ���ނ�؏o��
'
' �Ԃ�l    : �׸ٴװ����
'             "ORA-????? "��������Ȃ��ꍇ ""��Ԃ�
'
' ������    : (�׸ٵ�޼ު��).LastServerErrText :"ORA-????? "���܂ޕ�����
'
' �@�\����  : �׸ٴװү���ނ��׸ٴװ���ނ�؏o��
'
' ���l      :
'///////////////////////////////////////////////////
Private Function GetStrOraErrCd(LastServerErrText As String) As String
    Dim vPnt
    Dim vLen
    vPnt = InStr(LastServerErrText, "ORA-")             ''�װ���ނ̐擪
    If vPnt < 1 Then
        GetStrOraErrCd = ""
        Exit Function
    End If
    vLen = InStr(vPnt, LastServerErrText, ":") - 1      ''�ݸ޽����
    If vLen < 1 Then
        GetStrOraErrCd = Mid(LastServerErrText, vPnt)
    Else
        GetStrOraErrCd = Mid(LastServerErrText, vPnt, vLen)
    End If
End Function


'///////////////////////////////////////////////////
' @(f)
'
' �@�\      : ���b�Z�[�W�ҏW�E��ʕ\���E���O�o��
'
' �Ԃ�l    : �Ȃ�
'
' ������    : arg1:ү���޺��� 100:�׸ٴװ 100�ȊO:�׸ٴװ�ȊO
'             arg2:�ǉ�ү����
'             arg3:ү���ޑ����@0:�ʏ�ү����
'                              1:��ʕ\���װү���ށi���͗��ԕ\���̴װ�Ȃǁj
'                              2:���O�o�ʹװү����
'                              3:��ʕ\���E���O�o�ʹװү���ށi�׸ٴװ�Ȃǁj
'                              5:��ʕ\�����ޯ��ү����
'                              6:���O�o�����ޯ��ү����
'             arg4:�׸ٴװ����۸�/��ʕ\��ð��ٖ�
'
' �@�\����  : ү���޺��ނ���ү���ނ�ҏW���ĉ�ʏo�͂��A
'             �ǉ�ү���ނ�ҏW����ү���ނɒǉ����ă��O�o�͂��A
'             ү���޺��ނ��װ�敪�̏ꍇ�A�x������炷�B
' ���l      :
'       �y���O�o�͌`���z
'       YYYY/MM/DD HH:NN:SS::LOADMODULE::MSGCD::Msg::AddMsg ���s
'       YYYY/MM/DD HH:NN:SS::LOADMODULE::MSGCD::Msg::AddMsg ���s
'           .
'           .
'       YYYY/MM/DD HH:NN:SS::LOADMODULE::MSGCD::Msg::AddMsg ���s
'       YYYY/MM/DD = �N����
'       HH:NN:SS   = �����b
'       LOADMODULE = ���[�h���W���[����
'       MsgCd      = ���b�Z�[�W�ԍ�
'       Msg        = �T�v���b�Z�[�W(�Œ蕶��)
'       AddMsg     = �ڍ׃��b�Z�[�W
'       �y���O�o�͗�z
'       1998/04/01 10:10:00::Kxxxxxx.EXE::AA250100::�A�v���P�[�V�����N��::���b�Z�[�W�ڍ�
'       1998/04/01 10:10:03::Kxxxxxx.EXE::AA250200::�A�v���P�[�V�����I��::���b�Z�[�W�ڍ�
'
'///////////////////////////////////////////////////
Public Sub MsgOut(ByVal iMsgCd As Integer, Optional ByVal sAddMsgStr As String = "", _
           Optional ByVal eMsgKind As MsgKind = 0, Optional ByVal TABLENAME As String = "Unknown")
    Dim sMsg As String                              ''���b�Z�[�W
    Dim sOraErrCd As String                         ''�׸ٴװ����
    
    '���b�Z�[�W������
    Call MsgInit

    'ү���ޑ�����ү���ޏo�͑����͈͊O�̏ꍇ�o�͂��Ȃ��i�J���^�p�J�n��A���ޯ��ү���ނ��o�͂��Ȃ��悤�ɂł���j
    If Not ((eMsgKind = NORMAL_MSG) Or _
            ((eMsgKind And MsgKindMask) <> 0)) Then
        Exit Sub                                    ''�I��
    End If
    
    If iMsgCd < 100 Then                            ''ү���޺��ނ��׸وȊO�Ȃ�
        ''�I���N���ȊO�̃��b�Z�[�W
        On Error Resume Next                        ''�װ�ׯ��
        sMsg = msMsgStr(iMsgCd)                     ''���b�Z�[�W�擾
        On Error GoTo 0                             ''�װ�ׯ�߉���
    Else                                            ''ү���޺��ނ��׸ٴװ�Ȃ�
        ''�I���N���̃G���[���b�Z�[�W
        If gobjOraSess.LastServerErr Then           ''�׸پ���ݵ�޼ު�ẴG���[�Ȃ��
            sMsg = gobjOraSess.LastServerErrText    ''�׸پ���ݵ�޼ު�Ĵװү���ނ��Z�b�g
            gobjOraSess.LastServerErrReset          ''�׸پ���ݵ�޼ު�Ĵװ�����Z�b�g
        ElseIf gobjOraDB.LastServerErr Then         ''�׸��ް��ް���޼ު�ẴG���[�Ȃ��
            ''�׸ٴװү���ނ��׸ٴװ���ނ�؏o��
            sOraErrCd = GetStrOraErrCd(gobjOraDB.LastServerErrText)
            If sOraErrCd <> "" Then                 ''�׸ٴװ���ނ������Ă����
                sMsg = "DB�G���[�i" & TABLENAME & ")" & sOraErrCd ''�w��̃t�H�[�}�b�g�ŕҏW
                sAddMsgStr = gobjOraDB.LastServerErrText & _
                             "::" & sAddMsgStr
            Else                                    ''�׸ٴװ���ނ������Ă��Ȃ����
                sMsg = gobjOraDB.LastServerErrText  ''�׸��ް��ް���޼ު�Ĵװү���ނ��Z�b�g
            End If
            gobjOraDB.LastServerErrReset            ''�׸��ް��ް���޼ު�Ĵװ�����Z�b�g
        ElseIf Err.Number Then                      ''����VB�̴װ�������Ȃ�
            sMsg = Error(Err.Number)                ''VB�̴װү���ނ��Z�b�g
        Else                                        ''���ʹװ����Ȃ��Ȃ��
            sMsg = "�׸ِ��펞�ɴװ�o�͂���"         ''�x��
        End If
    End If
    
    If (eMsgKind = NORMAL_MSG) Or _
       (eMsgKind And ERR_DISP) Then                     ''�ʏ�ү���ނ���ʕ\���r�b�g�������Ă����
        ''�G���[�Ȃ�ԕ\��
        If (eMsgKind = ERR_DISP) Or _
           (eMsgKind = ERR_DISP_LOG) Then
            If iMsgCd = 100 Then                        ''�I���N���G���[�̏ꍇ
                MsgDisp sMsg, vbRed                     ''���b�Z�[�W����ʕ\������
            Else
                MsgDisp sMsg & sAddMsgStr, vbRed        ''���b�Z�[�W & �ǉ����b�Z�[�W����ʕ\������
            End If
        ''����ȊO�͍��\��
        Else
            If iMsgCd = 100 Then                        ''�I���N���G���[�̏ꍇ
                MsgDisp sMsg                            ''���b�Z�[�W����ʕ\������
            Else
                MsgDisp sMsg & sAddMsgStr               ''���b�Z�[�W & �ǉ����b�Z�[�W����ʕ\������
            End If
        End If
    End If
    
    If eMsgKind And ERR_LOG Then                    ''���O�o�̓r�b�g�������Ă����
        MsgLog (Format(Now, "YYYY/MM/DD HH:NN:SS::") & App.EXENAME & "::" & _
            iMsgCd & "::" & sMsg & "::" & sAddMsgStr) ''���b�Z�[�W�����O�o�͂���
    End If
    
    If (eMsgKind = ERR_DISP) Or _
       (eMsgKind = ERR_LOG) Or _
       (eMsgKind = ERR_DISP_LOG) Then                       ''ү���ޑ������G���[�Ȃ�
        Beep
    End If
End Sub


'///////////////////////////////////////////////////
' @(f)
' �@�\    :�c�a�ɃR�l�N�g����
'
' �Ԃ�l  : ���� - true
'           �ُ� - false
'
' ������  : �Ȃ�
'
' �@�\����: �c�a�ɃR�l�N�g����
'           �ȸĐ�́A�����ײ݈����̍H�꺰�ނɂ�芷����
'
'///////////////////////////////////////////////////
Public Function OraConn() As Boolean
    Dim sDbName As String
    Dim sUID As String
    Dim sPWD As String
    
    Select Case gsFactryCd
    Case "10"               ''��c�H��
        sDbName = "NODA"
        sUID = "oracle"
        sPWD = "oracle"
    Case "30"               ''����H��
        sDbName = "IKNO"
        sUID = "oracle"
        sPWD = "oracle"
    Case "40"               ''�đ�H��
        sDbName = "YONE"
        sUID = "oracle"
        sPWD = "oracle"
    Case "42"               '�f�R�O�O����
        sDbName = "cm1"
        sUID = "cm1"
        sPWD = "cm1"
    Case "43"               '�f�R�O�O����
        sDbName = "cmt"
        sUID = "cm1"
        sPWD = "cm1"
    Case "90"               ''�e�X�g��
        sDbName = "TEST0"
        sUID = "oracle"
        sPWD = "oracle"
    Case "91"               ''�e�X�g��(�V) 2007/04/05�ǉ� SETsw kubota
                            ''�e�X�g��(�đ�) 2009/11/16�ǉ� SSS.Marushita
        sDbName = "CLA0X"
        sUID = "oracle"
        sPWD = "oracle"
    Case "92"               ''�e�X�g��(����) 2009/11/16�ǉ� SSS.Marushita
        sDbName = "CLA0X"
        sUID = "oracle"
        sPWD = "oracle"
    Case "93"               ''�e�X�g��(����A1) 2010/04/14�ǉ� SETsw kubota
        sDbName = "CLA1"
        sUID = "oracle"
        sPWD = "oracle"
    Case "94"               ''�e�X�g��(���A1) 2009/11/16�ǉ� SSS.Marushita
        sDbName = "CLA1"
        sUID = "oracle"
        sPWD = "oracle"
    Case "99"               ''��
        sDbName = "BOIS"
        sUID = "BOIS"
        sPWD = "BOIS"
    Case "AM"               ''���H�� 2009/06/02�ǉ� SSS.Marushita
        sDbName = "CLK0"
        sUID = "oracle"
        sPWD = "oracle"
    Case Else               ''�O��
        sDbName = "oracle"
        sUID = "oracle"
        sPWD = "oracle"
    End Select
    
    On Error GoTo ConnError
    
    ''�I���N���ڑ�
    Set gobjOraSess = CreateObject("OracleInProcServer.XOraSession")
    Set gobjOraDB = gobjOraSess.OpenDatabase(sDbName, sUID & "/" & sPWD, 0&)
    
    OraConn = True
    Exit Function
    
ConnError:
    OraConn = False
End Function

'///////////////////////////////////////////////////
' @(f)
' �@�\    :�c�a�ɃR�l�N�g����
'
' �Ԃ�l  : ���� - true
'           �ُ� - false
'
' ������  : �Ȃ�
'
' �@�\����: �c�a�ɃR�l�N�g����
'           �ȸĐ�́A�����ײ݈����̍H�꺰�ނɂ�芷����
'
'///////////////////////////////////////////////////
Public Function OraConn2() As Boolean
    Dim sDbName2 As String
    Dim sUID2 As String
    Dim sPWD2 As String
    
        sDbName2 = "DWH"
        sUID2 = "dwhmgr"
        sPWD2 = "dwhmgr"
    
    On Error GoTo ConnError2
    
    ''�I���N���ڑ�
    Set gobjOraSess = CreateObject("OracleInProcServer.XOraSession")
    Set gobjOraDB = gobjOraSess.OpenDatabase(sDbName2, sUID2 & "/" & sPWD2, 0&)
    
    OraConn2 = True
    Exit Function
    
ConnError2:
    OraConn2 = False
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    :�c�a�J��
'
' �Ԃ�l  : ���� - true
'           �ُ� - false
'
' �@�\����: �c�a�J��
'
'///////////////////////////////////////////////////
Public Function OraDisConn() As Boolean
    
    On Error GoTo ErrProc
    
    ''�I���N���ؒf
    gobjOraDB.Close
    
    ''���
    Set gobjOraDB = Nothing
    Set gobjOraSess = Nothing
    
    OraDisConn = True
    Exit Function
    
ErrProc:
    OraDisConn = False
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    :�I���N���_�C�i�Z�b�g�̍쐬
'
' �Ԃ�l  : ���� - true
'           �ُ� - false
'
' ������  : ARG1 - �_�C�i�Z�b�g�Z�b�g�I�u�W�F�N�g
'           ARG2 - SQL��
'           ARG3 - �_�C�i�Z�b�g�I�v�V����
'
' �@�\����: �I���N���_�C�i�Z�b�g�쐬
'
'///////////////////////////////////////////////////
Public Function DynSet(ByRef objOraDynaset As Object, sSqlStmt As String, Optional vOpt = &H4&) As Boolean
    On Error GoTo DynErr
    
    ''�I���N���_�C�i�Z�b�g�쐬
    Set objOraDynaset = gobjOraDB.CreateDynaset(sSqlStmt, vOpt)
    DynSet = True
    Exit Function
    
DynErr:
    DynSet = False
End Function
'///////////////////////////////////////////////////
' @(f)
' �@�\    :�I���N���_�C�i�Z�b�g�̍쐬
'
' �Ԃ�l  : ���� - true
'           �ُ� - false
'
' ������  : ARG1 - �_�C�i�Z�b�g�Z�b�g�I�u�W�F�N�g
'           ARG2 - SQL��
'           ARG3 - �_�C�i�Z�b�g�I�v�V����
'
' �@�\����: �I���N���_�C�i�Z�b�g�쐬
'
'///////////////////////////////////////////////////
Public Function DynSet2(ByRef objOraDynaset As Object, sSqlStmt As String, Optional vOpt = &H4&) As Boolean
    On Error GoTo DynErr
    
    ''�I���N���_�C�i�Z�b�g�쐬
    Set objOraDynaset = OraDB.CreateDynaset(sSqlStmt, vOpt)
    DynSet2 = True
    Exit Function
    
DynErr:
    DynSet2 = False
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �r�p�k�����s
'
' �Ԃ�l  : 0�ȏ�:��������
'           �@-1�F�ُ�
'
' ������  : ARG1 - SQL��
'
' �@�\����: �r�p�k�����s���A����������Ԃ�
'
'///////////////////////////////////////////////////
Public Function SqlExec(sSqlStmt As String) As Long
    On Error GoTo ErrProc
    
    ''�I���N���r�p�k���s
    SqlExec = gobjOraDB.DbExecuteSQL(sSqlStmt)
    
    Exit Function
    
ErrProc:
    SqlExec = -1
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �r�p�k�����s
'
' �Ԃ�l  : 0�ȏ�:��������
'           �@-1�F�ُ�
'
' ������  : ARG1 - SQL��
'
' �@�\����: �r�p�k�����s���A����������Ԃ�
'
'///////////////////////////////////////////////////
Public Function SqlExec2(sSqlStmt As String) As Long
    On Error GoTo ErrProc
    
    ''�I���N���r�p�k���s
    SqlExec2 = OraDB.DbExecuteSQL(sSqlStmt)
    
    Exit Function
    
ErrProc:
    SqlExec2 = -1
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �S���Җ��擾
'
' �Ԃ�l  : True:����
'           False:���s
'
' ������  : �S���Һ���
'
' �@�\����: �S���Һ��ނ���S���Җ����擾
'
'///////////////////////////////////////////////////
Public Function GetUserName(sUserCd As String, ByRef sUserName As String) As Boolean
    Dim sSqlStmt As String
    Dim objOraDyn As Object
    
    sUserName = vbNullString       ''�S���Җ��N���A�[
    ''�r�p�k���쐬
    sSqlStmt = "SELECT NVL(nameja9, ' ')                    "
    sSqlStmt = sSqlStmt & "FROM koda9                       "
    sSqlStmt = sSqlStmt & "WHERE sysca9 = 'K'               "
    sSqlStmt = sSqlStmt & "  AND shuca9 = '55'              "
    sSqlStmt = sSqlStmt & "  AND codea9 = '" & sUserCd & "' "
    
    ''�_�C�i�Z�b�g�쐬
    If DynSet(objOraDyn, sSqlStmt) = False Then
        ''�_�C�i�Z�b�g�쐬���s
        Call MsgOut(100, sSqlStmt, ERR_DISP_LOG)
        
        GetUserName = False
        Exit Function
    End If
    If objOraDyn.EOF Then
        ''�Y������S���Һ��ނ���������
        Call MsgOut(8, "", ERR_DISP)
        
        GetUserName = False
        Exit Function
    End If

    sUserName = objOraDyn(0)  ''�S���Җ��擾
    
    GetUserName = True        ''����������Ԃ�
    
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �i�ԕҏW
' �Ԃ�l  : True:����
'           False:���s
' ������  :
'
' �@�\����: �L�[���ڕҏW
'�@�@�@�@�@ �i�Ԃ��e�|�f�ŕ����ҏW�܂��͋t�ҏW���s���B
'
'///////////////////////////////////////////////////
'Public Function GetHinbanHensyu(sbHinban As String, sflg As Integer, ByRef sahinban As String) As Boolean
'
'    '�@"-" �Ȃ�����@"-"�L��ɕҏW
'        If sflg = 1 Then
'            sahinban = Format(sbHinban, "@@@-@@@@-@@@@")
'        End If
'
'    '�@"-" �L�肩��@"-"�����ɕҏW
'        If sflg = 2 Then
'            sahinban = Replace(sbHinban, "-", "")
'            'Mid(saHinban, 1, 6) = Mid(sbHinban, 1, 6)
'            'Mid(saHinban, 7, 2) = Mid(sbHinban, 8, 2)
'            'Mid(saHinban, 9, 3) = Mid(sbHinban, 11, 3)
'        End If
'
'    GetHinbanHensyu = True        ''����������Ԃ�
'
'End Function


Public Function GetHinbanHensyu(sbHinban As String, sFlg As Integer, ByRef sahinban As String) As Boolean
    If sbHinban = "G" Or sbHinban = "Z" Then
    '�@"-" �Ȃ�����@"-"�L��ɕҏW
        If sFlg = 1 Then
            sahinban = Format(sbHinban, "@")
'            sahinban = Format(sbHinban, "@@@-@@@@-@@@@")
        End If
        
    '�@"-" �L�肩��@"-"�����ɕҏW
        If sFlg = 2 Then
            sahinban = Replace(sbHinban, "-", "")
            'Mid(saHinban, 1, 6) = Mid(sbHinban, 1, 6)
            'Mid(saHinban, 7, 2) = Mid(sbHinban, 8, 2)
            'Mid(saHinban, 9, 3) = Mid(sbHinban, 11, 3)
        End If
    Else
    '�@"-" �Ȃ�����@"-"�L��ɕҏW
        If sFlg = 1 Then
            sahinban = Format(sbHinban, "@@@-@@@@-@")
'            sahinban = Format(sbHinban, "@@@-@@@@-@@@@")
        End If
        
    '�@"-" �L�肩��@"-"�����ɕҏW
        If sFlg = 2 Then
            sahinban = Replace(sbHinban, "-", "")
            'Mid(saHinban, 1, 6) = Mid(sbHinban, 1, 6)
            'Mid(saHinban, 7, 2) = Mid(sbHinban, 8, 2)
            'Mid(saHinban, 9, 3) = Mid(sbHinban, 11, 3)
        End If
    End If
    GetHinbanHensyu = True        ''����������Ԃ�
    
End Function
                                                                                                                                                                                         
'///////////////////////////////////////////////////
' @(f)
' �@�\    : ���ԕҏW
' �Ԃ�l  : True:����
'           False:���s
' ������  :
'
' �@�\����: �L�[���ڕҏW
'�@�@�@�@�@ ���Ԃ��e�|�f�ŕ����ҏW�܂��͋t�ҏW���s���B
'
'///////////////////////////////////////////////////
Public Function GetSeibanHensyu(sbSeiban As String, sFlg As Integer, ByRef saSeiban As String) As Boolean
    
    '�@"-" �Ȃ�����@"-"�L��ɕҏW
        If sFlg = 1 Then
            saSeiban = Format(sbSeiban, "@@-@@@@@")
        End If
        
    '�@"-" �L�肩��@"-"�����ɕҏW
        If sFlg = 2 Then
            saSeiban = Replace(sbSeiban, "-", "")
            'Mid(saSeiban, 1, 2) = Mid(sbSeiban, 1, 2)
            'Mid(saSeiban, 3, 5) = Mid(sbSeiban, 4, 5)
        End If
    
    GetSeibanHensyu = True        ''����������Ԃ�
    
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �����ԍ��ҏW
' �Ԃ�l  : True:����
'           False:���s
' ������  :
'
' �@�\����: �L�[���ڕҏW
'�@�@�@�@�@ �����ԍ����e�|�f�ŕ����ҏW�܂��͋t�ҏW���s���B
'
'///////////////////////////////////////////////////
Public Function GetXtalHensyu(sbXtal As String, sFlg As Integer, ByRef saXtal As String) As Boolean
Dim wXTAL As String
    
    '�@"-" �Ȃ�����@"-"�L��ɕҏW
        If sFlg = 1 Then
            wXTAL = sbXtal
            If Len(sbXtal) > 0 Then wXTAL = wXTAL & "000000000000"
            saXtal = Format(Mid(wXTAL, 1, 12), "@@@@-@@@@@-@@@")
        End If
        
    '�@"-" �L�肩��@"-"�����ɕҏW
        If sFlg = 2 Then
            saXtal = Replace(sbXtal, "-", "")
            'Mid(saXtal, 1, 4) = Mid(sbXtal, 1, 4)
            'Mid(saXtal, 5, 3) = Mid(sbXtal, 6, 3)
            'Mid(saXtal, 8, 2) = Mid(sbXtal, 10, 2)
            'Mid(saXtal, 10, 3) = Mid(sbXtal, 13, 3)
        End If
        
        sFlg = 0
        
    GetXtalHensyu = True        ''����������Ԃ�
    
End Function

' @(f)
'
' �@�\      : �}�E�X�|�C���^�ύX
'
' �Ԃ�l    : �Ȃ�
'
' ����      : ipID   -   0:�W��
'                       1:�����v
'
' �@�\����  : �}�E�X�|�C���^�ύX
'
Public Sub SetMousePointer(ipID%)
    Select Case ipID
    Case 0
        Screen.MousePointer = vbDefault    ''�}�E�X�|�C���^�W��
    Case 1
        Screen.MousePointer = vbHourglass  ''�}�E�X�|�C���^�����v
    Case Else
        Screen.MousePointer = vbDefault    ''�}�E�X�|�C���^�W��
    End Select
End Sub


' @(f)
'
' �@�\      : ���t�`�F�b�N
'
' �Ԃ�l  : OK - TRUE
'           NG - FALSE
'
' ����      : sDate - (String)���t
'             iKind - 0:yyyymmdd�`��
'             iKind - 1:yymmdd�`��
'             iKind - 2:���̂܂ܕϊ��\�Ȍ`��
'
' �@�\����  : ���t�ɕϊ��\���`�F�b�N����
'
Public Function DateCheck(sDate$, iKind%) As Boolean
    Dim sCheckDate As String
    
    Select Case iKind
    Case 0
        sCheckDate = Mid(sDate, 1, 4) & "/" & Mid(sDate, 5, 2) & "/" & Mid(sDate, 7)
    Case 1
        sCheckDate = Mid(sDate, 1, 2) & "/" & Mid(sDate, 3, 2) & "/" & Mid(sDate, 5)
    Case Else
        sCheckDate = sDate
    End Select
    
    If IsDate(sCheckDate) Then
        DateCheck = True
    Else
        DateCheck = False
    End If
End Function


' @(f)
' �@�\    : ���ԃ`�F�b�N
'
' �Ԃ�l  :  OK - TRUE
'            NG -FALSE
'
' ������  : ctlControlS : �R���g���[��(�J�n��)
'           ctlControlE : �R���g���[��(�I����)
'           sDateS      : �J�n��
'           sDateE      : �I����
'
' �@�\����: ���ԃ`�F�b�N���s��8���̔N������Ԃ�
'Update - 2000/02/15
Public Function KikanCheck(ctlControlS As Control, ctlControlE As Control, _
        ByRef sDateS$, ByRef sDateE$) As Boolean
    'xxxxxxxxxxxxxxxxxxxxxxx
    '   mdlDWHCommon.bas?
    'xxxxxxxxxxxxxxxxxxxxxxx
    Dim sDtS    As String       ''�W�v���ԊJ�n��
    Dim sDtE    As String       ''�W�v���ԏI����
    Dim sDtT    As String       ''�V�X�e�����t
    Dim sDtL    As String       ''�Y�����̌�����(�J�n��)
    Dim sDtLE   As String       ''�Y�����̌�����(�I����)
    Dim sWk     As String
    
    KikanCheck = False
    
    ''�V�X�e�����t�擾([yymmdd])
    sDtT = Format(Date, "yymmdd")

    ''�W�v���Ԏ擾(6��[yymmdd])
    sDtS = Trim(ctlControlS.text)
    sDtE = Trim(ctlControlE.text)
    
    
    ''�J�n���̌��`�F�b�N
    If Len(sDtS) = 0 Then
        If Len(sDtE) <> 0 Then
            '''�I�����̂ݓ��͂̓G���[
            Call MsgOut(0, "���Ԃ̊J�n������͂��ĉ�����", ERR_DISP)
            Call CtrlEnabled(ctlControlS, RED_CTL)
            Exit Function
        End If
        '''�����͓͂��������`������ݒ�
        sDtS = Mid(sDtT, 1, 4) & "00"
    
    ElseIf Mid(sDtS, 5) = "00" Then
        If Len(sDtE) <> 0 Then
            Call MsgOut(0, "���Ԃ̏I�����̓��͕͂K�v����܂���", ERR_DISP)
            Call CtrlEnabled(ctlControlE, RED_CTL)
            Exit Function
        End If
    
    ElseIf Not DateCheck(sDtS, 1) Then
    
    
    
    
    
    ElseIf Mid(sDtS, 5) <> "00" Then
        If Len(sDtE) = 0 Then
            sDtE = sDtT
        End If
    
    End If
    
    
    ''�J�n���̓��t�`�F�b�N
    If DateCheck(Mid(sDtS, 1, 4) & "01", 1) Then
        If Mid(sDtS, 1, 4) = Mid(sDtT, 1, 4) Then
            '''�Y�����������̏ꍇ�͌������𓖓��ɐݒ�
            sDtL = sDtT
        Else
            '''�Y�����̌������Z�o([yymmdd])
            sWk = DateAdd("m", 1, Mid(sDtS, 1, 2) & "/" & Mid(sDtS, 3, 2) & "/" & "01")
            sDtL = DateAdd("d", -1, sWk)
''            sDtLE = Format(sDtLE, "yy/mm/dd") '2003/10/24 tuku SUMCO�a�������e�ǉ�
            sDtL = Format(sDtL, "yy/mm/dd") '2004/11/26 ���t�t�H�[�}�b�g�s��Ή�
            sDtL = Mid(sDtL, 1, 2) & Mid(sDtL, 4, 2) & Mid(sDtL, 7)
'            sDtL = Mid(sDtL, 3, 2) & Mid(sDtL, 6, 2) & Mid(sDtL, 9)
        End If
        Call CtrlEnabled(ctlControlS, NORMAL_CTL)
    Else
        '''���t�ɕϊ��ł��Ȃ��ꍇ�̓G���[
        Call CtrlEnabled(ctlControlS, RED_CTL)
        Call MsgOut(52, "", ERR_DISP)
        Exit Function
    End If
    
   
    If Len(sDtE) = 0 Then
        '''�J�n���̓����ڂ�[00]�̏ꍇ�͊Y�����̏������猎����ݒ�
        '''(�����̏ꍇ�͏������瓖��)
        sDtE = sDtL
    End If
    
    
    ''�I�����̓��t�`�F�b�N
    If DateCheck(Mid(sDtE, 1, 4) & "01", 1) Then
        If Mid(sDtE, 1, 4) = Mid(sDtT, 1, 4) Then
            '''�Y�����������̏ꍇ�͌������𓖓��ɐݒ�
            sDtLE = sDtT
        Else
            '''�Y�����̌������Z�o([yymmdd])
            sWk = DateAdd("m", 1, Mid(sDtE, 1, 2) & "/" & Mid(sDtE, 3, 2) & "/" & "01")
            sDtLE = DateAdd("d", -1, sWk)
            sDtLE = Format(sDtLE, "yy/mm/dd") '2003/10/24 tuku SUMCO�a�������e�ǉ�
'            sDtLE = Mid(sDtLE, 3, 2) & Mid(sDtLE, 6, 2) & Mid(sDtLE, 9)
            sDtLE = Mid(sDtLE, 1, 2) & Mid(sDtLE, 4, 2) & Mid(sDtLE, 7)
        End If
        Call CtrlEnabled(ctlControlE, NORMAL_CTL)
    Else
        '''���t�ɕϊ��ł��Ȃ��ꍇ�̓G���[
        Call CtrlEnabled(ctlControlE, RED_CTL)
        Call MsgOut(52, "", ERR_DISP)
        Exit Function
    End If
    
    ''�W�v���Ԃ̓��t�ϊ�
    If Mid(sDtS, 5) = "00" Then
        '''�J�n���̓����ڂ�[00]�̏ꍇ�͊Y�����̏������猎����ݒ�
        '''(�����̏ꍇ�͏������瓖��)
        sDtS = Mid(sDtS, 1, 4) & "01"
        sDtE = sDtL
    
    ElseIf Mid(sDtS, 5) > Mid(sDtL, 5) Then
        '''�J�n�����������傫���ꍇ�͊J�n���Ɍ�������ݒ肵
        '''�I�����������͂̏ꍇ�͓�����ݒ肷��
        sDtS = sDtL
        If Len(sDtE) = 0 Then
            sDtE = sDtT
        ElseIf Mid(sDtE, 5) > Mid(sDtLE, 5) Then
            sDtE = sDtLE
        End If
    
    Else
        '''����ȊO�̏ꍇ�͏W�v���Ԃ̓��t�`�F�b�N���s��
        If Mid(sDtE, 5) > Mid(sDtLE, 5) Then
            '''�I�������������傫���ꍇ�͏I�����Ɍ�������ݒ�
            sDtE = sDtLE
        End If
'*********************
        If Mid(sDtE, 5) = "00" Then
        '''�J�n���̓����ڂ�[00]�̏ꍇ�͊Y�����̏������猎����ݒ�
        '''(�����̏ꍇ�͏������瓖��)
            sDtE = Mid(sDtE, 1, 4) & "01"
        End If
'**********************
        If Not DateCheck(sDtS, 1) Then
            '''���t�ɕϊ��ł��Ȃ��ꍇ�̓G���[(�J�n��)
            Call CtrlEnabled(ctlControlS, RED_CTL)
            Call MsgOut(52, "", ERR_DISP)
            Exit Function
        End If
        If Not DateCheck(sDtE, 1) Then
            '''���t�ɕϊ��ł��Ȃ��ꍇ�̓G���[(�I����)
            Call CtrlEnabled(ctlControlE, RED_CTL)
            Call MsgOut(52, "", ERR_DISP)
            Exit Function
        End If
    End If
    
    ''�N�����̌����킹
    ''�e�X�g�p��1900�N����Ή�����(2000�𑫂���������1900�N��ɑΉ��ł��Ȃ�)
'    sDtS = "20" & sDtS
'    sDtE = "20" & sDtE
    sDtS = DateChange(sDtS)
    sDtE = DateChange(sDtE)
    
    ''���t�͈̔̓`�F�b�N
    If val(sDtS) > val(sDtE) Then
        Call CtrlEnabled(ctlControlS, RED_CTL)
        Call CtrlEnabled(ctlControlE, RED_CTL)
        Call MsgOut(53, "", ERR_DISP)
        Exit Function
    End If
    
    sDateS = sDtS
    sDateE = sDtE
    
    KikanCheck = True

End Function


' @(f)
' �@�\    : ���t�ϊ�
'
' �Ԃ�l  : �ϊ�����t
'
' ������  : sDate - ���t
'
' �@�\����: 6��[yymmdd]�̓��t��8��[yyyymmdd]�ɂ���
'
Public Function DateChange$(sDate$)
    'xxxxxxxxxxxxxxxxxxxxxxx
    '   mdlDWHCommon.bas?
    'xxxxxxxxxxxxxxxxxxxxxxx
    If Mid(sDate, 1, 2) < 50 Then
        DateChange = "20" & sDate
    Else
        DateChange = "19" & sDate
    End If
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �`���[�W�ʊ֐�
'
' �Ԃ�l  : �`���[�W��
'
' ������  : ��������ޗ�
'           į�߶�ďd��
'           ����ďd��
'
' �@�\����: �`���[�W�ʂ̌v�Z
'
'  �`���[�W�� = ��������ޗ� - į�߶�ďd�� - ����ďd��
'
'///////////////////////////////////////////////////
Public Function CHARGEWEIGHT(lSuiteiChargeWeight As Long, _
                      lTopCutWeight As Long, _
                      lShoulderCutWeight As Long _
                      ) As Long
    CHARGEWEIGHT = lSuiteiChargeWeight - lTopCutWeight - lShoulderCutWeight
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �ΐ͒l�֐�
'
' �Ԃ�l  : �ΐ͒l
'
' ������  : TOP������R
'           BOT������R
'           �ؒf��d��
'           �`���[�W��
'
' �@�\����: �ΐ͒l�̌v�Z
'
'           log(TOP������R / BOT������R)
'  �ΐ͒l = �������������������������������� + 1
'           log(1 - �ؒf��d�� / �`���[�W��)
'
'           �g�p�����F�\�߃`���[�W�ʂ��v�Z���Ă���
'
'///////////////////////////////////////////////////
Public Function Henseki(dTopRes As Double, _
                 dBotRes As Double, _
                 lCutAfterWeight As Long, _
                 lChargeWeight As Long _
                 ) As Double
    Dim dVal As Double              ''�r���v�Z�p
    
    ''�[���`�F�b�N
    If dBotRes = 0 Then             ''BOT������R�[���`�F�b�N
        Call MsgOut(14, "BOT������R", ERR_DISP)
        Exit Function
    ElseIf dTopRes = 0 Then         ''TOP������R�[���`�F�b�N
        Call MsgOut(14, "TOP������R", ERR_DISP)
        Exit Function
    ElseIf lChargeWeight = 0 Then   ''�`���[�W�ʃ[���`�F�b�N
        Call MsgOut(14, "�`���[�W��", ERR_DISP)
        Exit Function
    End If
        
    dVal = 1 - lCutAfterWeight / lChargeWeight ''1 - �ؒf��d�� / �`���[�W��
    If dVal = 0 Then                ''�[���`�F�b�N
        Call MsgOut(14, "1-�ؒf��d��/����ޗ�", ERR_DISP)
        Exit Function
    ElseIf dVal < 0 Then
        Call MsgOut(14, "�ؒf��d�� > ����ޗ�", ERR_DISP)
        Exit Function
    End If
    
    dVal = Log(dVal)                ''log(1 - �ؒf��d�� / �`���[�W��)
    If dVal = 0 Then                ''�[���`�F�b�N
        Call MsgOut(14, "log(1-�ؒf��d��/����ޗ�)", ERR_DISP)
        Exit Function
    End If
    
    Henseki = Log(dTopRes / dBotRes) / dVal + 1
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : ����ʒu�֐�
'
' �Ԃ�l  : ����ʒu�i���i�ʒu�j
'
' ������  : �ڕW��R
'           TOP������R
'           �ΐ͒l
'           �ؒf��d��
'           �`���[�W��
'           �ؒf�㒷��
'           �[�����Z�h�~�G���[���O�o�̓t���O
'
' �@�\����: ����ʒu�̌v�Z
'                                           1
'                                       ����������
'                                       1 - �ΐ͒l
'             1 - (�ڕW��R / TOP������R)
'  ����ʒu = ����������������������������������������
'              �ؒf��d�� / �`���[�W�� / �ؒf�㒷��
'
'           �g�p�����F�\�߃`���[�W�ʂƕΐ͒l���v�Z���Ă���
'
'///////////////////////////////////////////////////
Public Function SuiteiIchi(dTargetRes As Double, _
                            dTopRes As Double, _
                            dHenseki As Double, _
                            lCutAfterWeight As Long, _
                            lChargeWeight As Long, _
                            iCutAfterSize As Integer, _
                            Optional bErrLogFlg As Boolean = True _
                            ) As Double
    
    If (1 - dHenseki) = 0 Then                 ''(1-�ΐ͒l)�[���`�F�b�N
        If bErrLogFlg Then Call MsgOut(14, "1-�ΐ͒l", ERR_DISP)
        Exit Function
    ElseIf dTopRes = 0 Then                    ''TOP������R�[���`�F�b�N
        If bErrLogFlg Then Call MsgOut(14, "TOP������R", ERR_DISP)
        Exit Function
    ElseIf lChargeWeight = 0 Then              ''�`���[�W�ʃ[���`�F�b�N
        If bErrLogFlg Then Call MsgOut(14, "�`���[�W��", ERR_DISP)
        Exit Function
    ElseIf iCutAfterSize = 0 Then              ''�ؒf�㒷���[���`�F�b�N
        If bErrLogFlg Then Call MsgOut(14, "�ؒf�㒷��", ERR_DISP)
        Exit Function
    End If
    
    SuiteiIchi = (CLng(1) - (dTargetRes / dTopRes) ^ (CLng(1) / (CLng(1) - dHenseki))) / _
                 (CLng(lCutAfterWeight) / CLng(lChargeWeight) / CLng(iCutAfterSize))
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : ����d�ʊ֐�
'
' �Ԃ�l  : ����d�ʁi���i�d�ʁj
'
' ������  : ����ʒu
'           �`���[�W��
'           �ؒf�㒷��
'
' �@�\����: ����d�ʂ̌v�Z
'
'  ����d�� = ����ʒu �~ �`���[�W�� / �ؒf�㒷��
'
'           �g�p�����F�\�߃`���[�W�ʂƐ���ʒu���v�Z���Ă���
'
'///////////////////////////////////////////////////
Public Function SuiteiWeight(sSuiteiIchi As Double, _
                            lChargeWeight As Long, _
                            iCutAfterSize As Integer _
                            ) As Double
    
    If iCutAfterSize = 0 Then                   ''�ؒf�㒷���[���`�F�b�N
        Call MsgOut(14, "�ؒf�㒷��", ERR_DISP)
        Exit Function
    End If
    
    SuiteiWeight = sSuiteiIchi * lChargeWeight / iCutAfterSize
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �����R���֐�
'
' �Ԃ�l  : �����R��
'
' ������  : ����ʒu
'           �ؒf��d��
'           �`���[�W��
'           �ؒf�㒷��
'           �ΐ͒l
'           TOP������R
'
' �@�\����: �����R���̌v�Z
'                                                        (1 - �ΐ͒l)
'  �` = (1 - ����ʒu �~ �ؒf��d�� / �`���[�W�� / �ؒf�㒷��)
'
'  �����R = TOP������R �~ �`
'
'           �g�p�����F�\�߃`���[�W�ʂƕΐ͒l�Ɛ���ʒu���v�Z���Ă���
'
'///////////////////////////////////////////////////
Public Function SuiteiRes(iSuiteiIchi As Integer, _
                        lCutAfterWeight As Long, _
                        lChargeWeight As Long, _
                        iCutAfterSize As Integer, _
                        dHenseki As Double, _
                        dTopRes As Double _
                        ) As Double
    Dim dA As Double
    Dim db As Double
    
    If lChargeWeight = 0 Then                   ''�`���[�W�ʃ[���`�F�b�N
        Call MsgOut(14, "�`���[�W��", ERR_DISP)
        Exit Function
    ElseIf iCutAfterSize = 0 Then               ''�ؒf�㒷���[���`�F�b�N
        Call MsgOut(14, "�ؒf�㒷��", ERR_DISP)
        Exit Function
    End If
    
    db = 1 - iSuiteiIchi * lCutAfterWeight / lChargeWeight / iCutAfterSize
    dA = CDbl(db) ^ (1 - dHenseki)
    SuiteiRes = dTopRes * dA
    SuiteiRes = val(Format(SuiteiRes, "######0.0######"))
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : ����1�o�d�ʊ֐�
'
' �Ԃ�l  : ����1�o�̏d��
'
' ������  : ���amm
'
' �@�\����: ����1�o����̏d�ʂ̌v�Z
'                            2
'                 �i���a / 2) �~ 3.14 �~ 2.33
'  ����1�o�̏d�� = ��������������������������
'                            1000
'
'///////////////////////////////////////////////////
Public Function WeightPar1mm(sChokkei As Single) As Double
    WeightPar1mm = (((sChokkei / 2) ^ 2) * 3.14 * 2.33) / 1000
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �����d�ʊ֐�
'
' �Ԃ�l  : �����̏d��
'
' ������  : ���amm
'           ����mm
'
' �@�\����: ���a�ƒ����ɂ�茋���d�ʂ��v�Z����
'
'  �����d�� = ����1�o�d�ʊ֐� �~ ����
'
'///////////////////////////////////////////////////
Public Function WeightCompute(sChokkei As Single, sNagasa As Single) As Double
    WeightCompute = WeightPar1mm(sChokkei) * sNagasa
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : ����˂炢��R�֐�
'
' �Ԃ�l  : ����˂炢��R
'
' ������  : �K�i��R���
'           ��������Ǘ�(�p�[�Z���g)
'
'
' �@�\����: ����˂炢��R�̌v�Z
'
' ����˂炢��R = (�K�i��R��� - �K�i��R��� �~ ��������Ǘ� �~ 0.01) �~ 0.97
'
'///////////////////////////////////////////////////
Public Function UperTergetRes(dUperRes As Double, iUperInPar As Integer) As Double
    UperTergetRes = (dUperRes - dUperRes * (iUperInPar * 0.01)) * 0.97
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �˂炢��R�ɑΉ�����s�����Z�x�֐�
'
' �Ԃ�l  : �s�����Z�x
'
' ������  : �˂炢��R
'           �W���`
'           �W���a
'
' �@�\����: �˂炢��R�ɑΉ�����s�����Z�x�v�Z
'
'                                     1       ((Log(�˂炢��R)�|�W���a)���W���`)
' �˂炢��R�ɑΉ�����s�����Z�x = ������ �~10
'                                   2.33
'///////////////////////////////////////////////////
Public Function DopantPar1g(sngTergetRes As Single, _
                            sngA As Single, sngB As Single) As Single
    If sngA = 0 Then   ''�W���`�[���`�F�b�N
        Call MsgOut(14, "�W���`", ERR_DISP)
        Exit Function
    End If
    DopantPar1g = (1 / 2.33) * 10 ^ ((Log(sngTergetRes) - sngB) / sngA)
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �h�[�p���g�ʂ̌v�Z
'
' �Ԃ�l  : �h�[�p���g��
'
' ������  : �`���[�W��
'           �˂炢��R�ɑΉ�����s�����Z�x
'           �s�����Z�x
'           �t�@�N�^�[
'           �ΐ͒l
'
' �@�\����: �s�����i�h�[�p���g�j�ʊ֐�
'
'                �`���[�W��(g)�~�˂炢��R�ɑΉ�����s�����Z�x
' �h�[�p���g�� = ���������������������������������������������� �~ �t�@�N�^�[ �~ �ΐ͒l
'                                   �s�����Z�x
'///////////////////////////////////////////////////
Public Function DopantWeight(sngCharge As Single, sngDopantPar1g As Single, _
                             sngDopant As Single, sngFactor As Single, _
                             Optional sngHenseki As Single = 1) As Single
    If sngDopant = 0 Then   ''�s�����Z�x�[���`�F�b�N
        Call MsgOut(14, "�s�����Z�x", ERR_DISP)
        Exit Function
    End If
    DopantWeight = ((sngCharge * sngDopantPar1g) / sngDopant) * sngFactor * sngHenseki
End Function

'///////////////////////////////////////////////////
' @(f)
' �@�\    : �h�[�p���g�ʂ̌v�Z�Q
'
' �Ԃ�l  : �h�[�p���g��(�l�������g�p�j
'
' ������  : �`���[�W��
'           �˂炢��R�ɑΉ�����s�����Z�x
'           ���q��
'           �ΐ͒l
'
' �@�\����: �s�����i�h�[�p���g�j�ʊ֐��Q
'
'                �`���[�W��(g)�~�˂炢��R�ɑΉ�����s�����Z�x �~ ���q��         �P
' �h�[�p���g�� = ���������������������������������������������������������� �~ ����������
'                                   0.6 �~ 10�O23                           �ΐ͒l
'
'///////////////////////////////////////////////////
Public Function DopantWeight2(sngCharge As Single, sngDopantPar1g As Single, _
                             sngBunsi As Single, _
                             Optional sngHenseki As Single = 1) As Single
    If sngHenseki = 0 Then   ''�ΐ͒l�[���`�F�b�N
        Call MsgOut(14, "�ΐ͒l", ERR_DISP)
        Exit Function
    End If
    DopantWeight2 = ((sngCharge * sngDopantPar1g * sngBunsi) / (0.6 * 10 ^ 23)) / sngHenseki
End Function

'///////////////////////////////////////////////////
' @(f)
' �@�\    : �h�[�p���g�ʂ̌v�Z�R
'
' �Ԃ�l  : �h�[�p���g��(�������g�g�p�j
'
' ������  : �������g��
'           �������g���σC�I���Z�x
'           �s�����Z�x
'           ���q��
'           �ΐ͒l
'
' �@�\����: �s�����i�h�[�p���g�j�ʊ֐��R
'
'                �������g��(g)�~�������g���σC�I���Z�x
' �h�[�p���g�� = ���������������������������������������� �~ �t�@�N�^�[ �~ �ΐ͒l
'                                   �s�����Z�x
'///////////////////////////////////////////////////
Public Function DopantWeight3(sngCharge As Single, sngDopantPar1g As Single, _
                             sngDopant As Single, sngFactor As Single, _
                             Optional sngHenseki As Single = 1) As Single
    If sngDopant = 0 Then   ''�s�����Z�x�[���`�F�b�N
        Call MsgOut(14, "�s�����Z�x", ERR_DISP)
        Exit Function
    End If
    DopantWeight3 = ((sngCharge * sngDopantPar1g) / sngDopant) * sngFactor * sngHenseki
End Function


' @(f)
'
' �@�\      : ���l����
'
' �Ԃ�l    : �Ȃ�
'
' ������    : ctlControl    -   �R���g���[��
'             sChar         -   ���l������
'
' �@�\����  : �w�肵���������MaxLength�ɑ���Ȃ������������߂�
'
' ���l      :
'
Public Sub FillUpString(ctlControl As Control, sChar As String)
    Dim iLength As Integer
    If TypeOf ctlControl Is TextBox Then
        iLength = LenB(StrConv(Trim(ctlControl.text), vbFromUnicode))
        If iLength < ctlControl.MaxLength Then
            ctlControl.text = Trim(ctlControl.text) & String(ctlControl.MaxLength - iLength, sChar)
        End If
    End If
End Sub


' @(f)
'
' �@�\      : ����RRG�Z�o����
'
' �Ԃ�l    : RRG�l
'
' ����      : CalcType  - �v�Z���(�ۏؕ��@)
'             TeikouDat - ��R�l�ꗗ
'
' �@�\����  :�@�v�Z���(�ۏؕ��@)�ɂ��RRG���Z�o���鏈��������
'
' ���l      : RRG�l��[999999]�̏ꍇ�ARRG�v�ZNG���͊Y���ۏؕ��@����
'
Public Function GetCalcRRG(iCalcType As Integer, tTeikouDat As TYPE_RRG) As Double
    Dim dMaxNum As Double
    Dim dMinNum As Double
    Dim dHeikin As Double
    Dim dChuou As Double
    Dim dRRG As Double
    
    '' �v�Z���t���O����                <--- 2000/04/24 �ǉ�
    If iCalcType < 1 Or iCalcType > 7 Then
        GetCalcRRG = NULL_CHECK
        Exit Function
    End If
    
    '' �ő�l�擾
    dMaxNum = GetMax(iCalcType, tTeikouDat)
    '' ��R�lMAX����
    If dMaxNum = NULL_CHECK Then
        GetCalcRRG = dMaxNum
        Exit Function
    End If
    
    '' �ŏ��l�擾
    dMinNum = GetMin(tTeikouDat)
    '' ��R�lMIN����
    If dMinNum = NULL_CHECK Then
        GetCalcRRG = dMinNum
        Exit Function
    End If
    
    '' ���ϒl�擾
    dHeikin = GetHeikin(tTeikouDat)
    
    '' �����l�擾
    dChuou = GetChuou(tTeikouDat, dHeikin)
    dRRG = 0
    
    '' �v�Z�s���菈��(���q���͕��ꂪ�O�ƂȂ�ꍇ)
    If dMaxNum = 0 Or dMinNum = 0 Or (dMaxNum - dMinNum) = 0 Then
        GetCalcRRG = NULL_CHECK
        Exit Function
    End If
    
    '' �ۏؕ��@�ɂ��RRG�Z�o�𕪊򂷂�
    Select Case iCalcType
    Case 1
        dRRG = (dMaxNum - dMinNum) / dMaxNum * 100
    Case 2
        dRRG = (dMaxNum - dMinNum) / dMinNum * 100
    Case 3
        dRRG = (dMaxNum - dMinNum) / dChuou * 100
    Case 4
        dRRG = (dMaxNum - dMinNum) / tTeikouDat.dTeikouDT(0).dTeikou * 100
        '' dRRG = (dMaxNum - dMinNum) / tTeikouDat.dTeikou(0) * 100
    Case 5
        dRRG = dMaxNum / tTeikouDat.dTeikouDT(0).dTeikou * 100
        '' dRRG = dMaxNum / tTeikouDat.dTeikou(0) * 100
    Case 6
        dRRG = dMaxNum / tTeikouDat.dTeikouDT(0).dTeikou * 100
        '' dRRG = dMaxNum / tTeikouDat.dTeikou(0) * 100
    Case 7
        dRRG = Abs(tTeikouDat.dTeikouDT(0).dTeikou - dHeikin) / (tTeikouDat.dTeikouDT(0).dTeikou + dHeikin) / 2 * 100
        '' dRRG = Abs(tTeikouDat.dTeikou(0) - dHeikin) / (tTeikouDat.dTeikou(0) + dHeikin) / 2 * 100
    Case Else
        dRRG = NULL_CHECK
    End Select

    ' �����킹����(4���ɒ���)
    GetCalcRRG = left(dRRG, 4)
End Function

' @(f)
'
' �@�\      : �ő�l�擾
'
' �Ԃ�l    : �ő�l
'
' ����      : CalcType  - �v�Z���(�ۏؕ��@)
'             TeikouDat - ��R�l�ꗗ
'
' �@�\����  :�@��R�l�̍ő�l���������鏈��������
'
' ���l      :
'
Private Function GetMax(iCalcType As Integer, tTeikouDat As TYPE_RRG) As Double
    Dim dMaxData As Double
    Dim iCntI As Integer
    Dim iNextCnt As Integer
    
    iNextCnt = 0
    iCntI = 0
    
    With tTeikouDat
        '' �ۏؕ��@�ɂ�蕪��
        Select Case iCalcType
        Case 5
            '' �t���O�`�F�b�N
            If .dTeikouDT(0).sRRGFlg = "1" Then
                If .dTeikouDT(1).sRRGFlg = "1" And .dTeikouDT(1).dTeikou <> NULL_CHECK Then
                    dMaxData = .dTeikouDT(1).dTeikou - .dTeikouDT(0).dTeikou
                    iNextCnt = 2
                Else
                    For iNextCnt = 2 To SOKUTEI_MAX
                    If .dTeikouDT(iNextCnt).sRRGFlg = "1" And .dTeikouDT(iNextCnt).dTeikou <> NULL_CHECK Then
                        dMaxData = .dTeikouDT(iNextCnt).dTeikou - .dTeikouDT(0).dTeikou
                        Exit For
                    End If
                    Next iNextCnt
                End If
            Else
                dMaxData = NULL_CHECK
            End If
            '' dMaxData = .dTeikou(1) - .dTeikou(0)
            For iCntI = iNextCnt To SOKUTEI_MAX
                '' �t���O�`�F�b�N                 <--- 2000/0424 �ύX
                If .dTeikouDT(iCntI).sRRGFlg = "1" Then
                    '' �ő�l���r����
                    If dMaxData < .dTeikouDT(iCntI).dTeikou - .dTeikouDT(0).dTeikou And .dTeikouDT(iCntI).dTeikou <> NULL_CHECK Then
                        '' �傫���ꍇ�A�ő�l���X�V����
                        dMaxData = .dTeikouDT(iCntI).dTeikou - .dTeikouDT(0).dTeikou
                    End If
                End If
                '' �v�Z�t���O�ǉ��ɂ��ύX        <--- 2000/04/24 �폜
                '' If dMaxData < .dTeikou(iCntI) - .dTeikou(0) And .dTeikou(iCntI) <> 999999 Then
                ''    '' �傫���ꍇ�A�ő�l���X�V����
                ''    dMaxData = .dTeikou(iCntI) - .dTeikou(0)
                '' End If
            Next iCntI
        Case 6
            '' �t���O�`�F�b�N                     <--- 2000/04/24 �ύX
            If .dTeikouDT(0).sRRGFlg = "1" Then
                If .dTeikouDT(1).sRRGFlg = "1" And .dTeikouDT(1).dTeikou <> NULL_CHECK Then
                   dMaxData = Abs(.dTeikouDT(1).dTeikou - .dTeikouDT(0).dTeikou)
                Else
                    For iNextCnt = 2 To SOKUTEI_MAX
                    If .dTeikouDT(iNextCnt).sRRGFlg = "1" And .dTeikouDT(iNextCnt).dTeikou <> NULL_CHECK Then
                        dMaxData = Abs(.dTeikouDT(iNextCnt).dTeikou - .dTeikouDT(0).dTeikou)
                        Exit For
                    End If
                    Next iNextCnt
                End If
            Else
                dMaxData = NULL_CHECK
            End If
            '' dMaxData = Abs(.dTeikou(1) - .dTeikou(0))
            For iCntI = 2 To SOKUTEI_MAX
                '' �t���O�`�F�b�N                  <--- 2000/04/24 �ύX
                If .dTeikouDT(iCntI).sRRGFlg = "1" Then
                    '' �ő�l���r����
                    If dMaxData < Abs(.dTeikouDT(iCntI).dTeikou - .dTeikouDT(0).dTeikou) And .dTeikouDT(iCntI).dTeikou <> NULL_CHECK Then
                        '' �傫���ꍇ�A�ő�l���X�V����
                        dMaxData = Abs(.dTeikouDT(iCntI).dTeikou - .dTeikouDT(0).dTeikou)
                    End If
                End If
                '' �v�Z�t���O�ǉ��ɂ��ύX        <--- 2000/04/24 �폜
                '' If dMaxData < Abs(.dTeikou(iCntI) - .dTeikou(0)) And .dTeikou(iCntI) <> 999999 Then
                ''     '' �傫���ꍇ�A�ő�l���X�V����
                ''     dMaxData = Abs(.dTeikou(iCntI) - .dTeikou(0))
                '' End If
            Next iCntI
        Case Else
            '' �t���O�`�F�b�N
            If .dTeikouDT(0).sRRGFlg = "1" Then
                dMaxData = .dTeikouDT(0).dTeikou
            Else
                dMaxData = NULL_CHECK
            End If
            '' dMaxData = .dTeikou(0)
            '' ��������J��Ԃ�(�ő�9��)
            For iCntI = 1 To SOKUTEI_MAX
                '' �t���O�`�F�b�N                   <--- 2000/04/24 �ύX
                If .dTeikouDT(iCntI).sRRGFlg = "1" Then
                    '' �ő�l���r����
                    If dMaxData < .dTeikouDT(iCntI).dTeikou And .dTeikouDT(iCntI).dTeikou <> NULL_CHECK Then
                        '' �傫���ꍇ�A�ő�l���X�V����
                        dMaxData = .dTeikouDT(iCntI).dTeikou
                    End If
                End If
                '' �v�Z�t���O�ǉ��ɂ��ύX          <--- 2000/04/24 �폜
                '' If dMaxData < .dTeikou(iCntI) And .dTeikou(iCntI) <> 999999 Then
                ''     '' �傫���ꍇ�A�ő�l���X�V����
                ''     dMaxData = .dTeikou(iCntI)
                '' End If
            Next iCntI
        End Select
    End With
    GetMax = dMaxData
End Function

' @(f)
'
' �@�\      : �ŏ��l�擾
'
' �Ԃ�l    : �ŏ��l
'
' ����      : TeikouDat - ��R�l�ꗗ
'
' �@�\����  :�@��R�l�̍ŏ��l���������鏈��������
'
' ���l      :
'
Private Function GetMin(tTeikouDat As TYPE_RRG) As Double
    Dim dMineData As Double
    Dim iCntI As Integer
    Dim bChkFlg As Boolean
    
    bChkFlg = False
    iCntI = 0
    dMineData = tTeikouDat.dTeikouDT(0).dTeikou
    '' dMineData = tTeikouDat.dTeikou(0)
    
    '' ��������J��Ԃ�(�ő�9��)
    For iCntI = 1 To SOKUTEI_MAX
        '' �t���O�`�F�b�N                   <--- 2000/04/24 �ύX
        If tTeikouDat.dTeikouDT(iCntI).sRRGFlg = "1" Then
            If tTeikouDat.dTeikouDT(iCntI).dTeikou <> NULL_CHECK Then
                '' ��R�l������ꍇ�A��R�l���r����
                If dMineData >= tTeikouDat.dTeikouDT(iCntI).dTeikou Then
                    '' �������ꍇ�A�ŏ��l���X�V����
                    dMineData = tTeikouDat.dTeikouDT(iCntI).dTeikou
                End If
                bChkFlg = True
            End If
        End If
        '' �v�Z�t���O�ǉ��ɂ��ύX          <--- 2000/04/24 �폜
        '' If tTeikouDat.dTeikou(iCntI) <> NULL_CHECK Then
        ''     '' ��R�l������ꍇ�A��R�l���r����
        ''     If dMineData > tTeikouDat.dTeikou(iCntI) Then
        ''         '' �������ꍇ�A�ŏ��l���X�V����
        ''         dMineData = tTeikouDat.dTeikou(iCntI)
        ''     End If
        '' End If
    Next iCntI
    If bChkFlg = False Then
        GetMin = NULL_CHECK
    Else
        GetMin = dMineData
    End If
End Function

' @(f)
'
' �@�\      : �����l�擾
'
' �Ԃ�l    : �����l
'
' ����      : TeikouDat - ��R�l�ꗗ
'             Heikin    - ���ϒl
'
' �@�\����  :�@��R�l�̒����l���擾���鏈��������
'
' ���l      :
'
Private Function GetChuou(tTeikouDat As TYPE_RRG, dHeikin As Double) As Double
    Dim dWorkIn(9) As Double
    Dim dJudgeDT As Double
    Dim iCntI As Integer
    Dim iJituCnt As Integer
    Dim iSortCnt As Integer
    Dim iChuouCnt As Integer
    iSortCnt = 0
    iJituCnt = 0
    iChuouCnt = 0
    
    ' ����f�[�^���J�E���g(�ő�9��)
    For iCntI = 0 To SOKUTEI_MAX
        '' �t���O�`�F�b�N                         <--- 2000/04/24 �ύX
        If tTeikouDat.dTeikouDT(iCntI).sRRGFlg = "1" Then
            '' �f�[�^�L������
            If tTeikouDat.dTeikouDT(iCntI).dTeikou <> NULL_CHECK Then
                '' ���̓f�[�^������ꍇ�A���[�N�G���A�ɑ��
                dWorkIn(iJituCnt) = tTeikouDat.dTeikouDT(iCntI).dTeikou
                iJituCnt = iJituCnt + 1
            End If
        End If
        '' �v�Z�t���O�ǉ��ɂ��ύX               <--- 2000/04/24 �폜
        '' If tTeikouDat.dTeikou(iCntI) <> NULL_CHECK Then
        ''     '' ���̓f�[�^������ꍇ�A���[�N�G���A�ɑ��
        ''     dWorkIn(iJituCnt) = tTeikouDat.dTeikou(iCntI)
        ''     iJituCnt = iJituCnt + 1
        '' End If
    Next iCntI

    '' �f�[�^����ёւ���
    Do
        iSortCnt = iSortCnt + 1
        ' ���������{����������
        If iSortCnt >= iJituCnt Then
            Exit Do
        End If
        dJudgeDT = dWorkIn(iSortCnt - 1)
        '' �f�[�^���r����
        If dJudgeDT > dWorkIn(iSortCnt) Then
            '' �f�[�^�������������ꍇ�A�f�[�^�����ւ���
            dJudgeDT = dWorkIn(iSortCnt)
            dWorkIn(iSortCnt) = dWorkIn(iSortCnt - 1)
            dWorkIn(iSortCnt - 1) = dJudgeDT
            iSortCnt = 0
        End If
    Loop

    '' ����������������
    If iJituCnt Mod 2 = 0 Then
        '' �����̏ꍇ�A���ςɋ߂�����ݒ肷��
        For iCntI = 0 To iJituCnt
        If dHeikin < dWorkIn(iCntI) Then
            iChuouCnt = iCntI
            Exit For
        End If
        Next iCntI
        '' ������r����
        If (dWorkIn(iCntI) - dHeikin) >= (dHeikin - dWorkIn(iCntI - 1)) Then
            GetChuou = dWorkIn(iCntI - 1)
        Else
            GetChuou = dWorkIn(iCntI)
        End If
    Else
        '' ��̏ꍇ�A�^�񒆂̒l��ݒ肷��
        iChuouCnt = iJituCnt \ 2
        GetChuou = dWorkIn(iChuouCnt)
    End If
End Function

' @(f)
'
' �@�\      : ���ϒl�擾
'
' �Ԃ�l    : ���ϒl
'
' ����      : TeikouDat - ��R�l�ꗗ
'
' �@�\����  :�@��R�l�̕��ϒl���擾���鏈��������
'
' ���l      :
'
Private Function GetHeikin(tTeikouDat As TYPE_RRG) As Double
    Dim dSumData As Double
    Dim iCnt As Integer
    Dim iKazu As Integer
    
    dSumData = 0
    iKazu = 0
    iCnt = 0
    
    '' �����R�����J��Ԃ�(A�`I)
    For iCnt = 0 To SOKUTEI_MAX
        '' �t���O�`�F�b�N                        <--- 2000/04/24 �ύX
        If tTeikouDat.dTeikouDT(iCnt).sRRGFlg = "1" Then
            ' ��R�l��[0]������
            If tTeikouDat.dTeikouDT(iCnt).dTeikou <> NULL_CHECK Then
                '' ��R�l���L��ꍇ�A���v�l���X�V����
                dSumData = dSumData + tTeikouDat.dTeikouDT(iCnt).dTeikou
                iKazu = iKazu + 1
            End If
        End If
        '' �v�Z�t���O�ǉ��ɂ��ύX              <--- 2000/04/24 �폜
        '' If tTeikouDat.dTeikou(iCnt) <> NULL_CHECK Then
        ''     '' ��R�l���L��ꍇ�A���v�l���X�V����
        ''     dSumData = dSumData + tTeikouDat.dTeikou(iCnt)
        ''     iKazu = iKazu + 1
        '' End If
    
    Next iCnt
    '' ���ϒl���Z�o����
    GetHeikin = dSumData / iKazu
End Function

' @(f)
'
' �@�\      :   �����d�|�H���e�[�u����������
'
' �Ԃ�l    :   �����E���s
'
' ����      :   sMateNum    -   �����ԍ�
'               sKoutei     -   �H���R�[�h
'               sWeight     -   �d�|�d��
'
' �@�\����  :�@ �w�肳�ꂽ�H���̎d�|�d�ʂ��������Ăяo�����ɕԂ�
'
' ���l      :
'
Public Function SelectSikakariWeightDat(sMateNum$, sKoutei$, sWeight$) As Boolean
    Dim sSql    As String
    Dim objDS   As Object
    SelectSikakariWeightDat = False
    sSql = "SELECT  NVL(TO_CHAR(siwb2,'FM999999999'),' ')   "
    sSql = sSql & "FROM     xodb2                           "
    sSql = sSql & "WHERE    polnob2 = '" & sMateNum & "'    "
    sSql = sSql & "  AND    wkktb2  = '" & sKoutei & "'     "
    ''  SQL�_�C�i�Z�b�g����
    ''  �G���[����FALSE��Ԃ�
    If DynSet(objDS, sSql) = False Then
        Call MsgOut(100, sSql, ERR_DISP_LOG, "xodb2")
        GoTo Er
    End If
    If objDS.EOF = False Then
        sWeight = objDS(0)
    End If
    Set objDS = Nothing
    SelectSikakariWeightDat = True
    Exit Function
Er:
    On Error Resume Next
    Set objDS = Nothing
End Function

' @(f)
'
' �@�\      :   �N�����{�^�����쏈��
'
' �Ԃ�l    :   �Ȃ�
'
' ����      :   sForm   -   �����Ώۃt�H�[��
'
' �@�\����  :�@ �����Ώۂ̃t�H�[���ɑ΂��ă{�^����L���E�����ɂ���
'
' ���l      :
'
Public Function InitCtrlAction(frmForm As Form, Optional bFlg As Boolean = False, Optional bAllFlg As Boolean = False) As Boolean
    Dim iIdx As Integer
    Dim ctlControl As Control
    If bAllFlg = True Then
        ''�@�t�H�[����̂��ׂẴ{�^����񋓂���
        For Each ctlControl In frmForm.Controls
            If TypeOf ctlControl Is CommandButton Then
                ctlControl.Enabled = bFlg
            End If
        Next ctlControl
    Else
        frmForm.cmdF(1).Enabled = bFlg
        frmForm.cmdF(2).Enabled = bFlg
        frmForm.cmdF(5).Enabled = bFlg
        frmForm.cmdF(6).Enabled = bFlg
        frmForm.cmdF(12).Enabled = bFlg
    End If
End Function

' @(f)
'
' �@�\      :   �X�v���b�h�o�͏���
'
' �Ԃ�l    :   �Ȃ�
'
' ����      :   sForm   -   �����Ώۃt�H�[��
'
' �@�\����  :�@ �X�v���b�h�V�[�g���b�r�u�t�@�C���ɏo�͂���B
'
' ���l      :
'
Function SPRD_PRT(msOBJ As Variant, msPGCAP As String, msNO As String) As Integer

Dim f_path As String    'CSV̧�ق��߽
Dim f_name(100)   As String    'CSV̧�ق̖��O
Dim f_time(100)   As Variant   'CSV̧�ق̎���
Dim old_f     As String    '�ŌÂ�CSV̧�ق̖��O
Dim old_t     As Variant   '�ŌÂ�CSV̧�ق̎���
Dim f_cnt     As Integer   '̧�ق̶���
Dim f_max     As Integer
Dim data_wk   As String
Dim i, j, rows_cnt, col_cnt As Long

On Error GoTo Err
       
    Screen.MousePointer = vbHourglass
    
    '�V�[�g�̃v���p�e�B�[�iMAXrows,MAXcol)�擾
    rows_cnt = msOBJ.MaxRows
    col_cnt = msOBJ.MaxCols
    Debug.Print rows_cnt, col_cnt
        
    '�o�͂b�r�u�t�@�C���̂n�o�d�m
    If Dir(App.Path & "\Copy", vbDirectory) = "" Then
        MkDir (App.Path & "\Copy")
    End If
    
    '2000.08.16 �ÎR
    f_max = 1
    f_name(f_max) = Dir(App.Path & "\Copy\*.csv")
            
    '�b�r�u�t�@�C��������
    Do While Not f_name(f_max) = ""
        f_max = f_max + 1
        f_name(f_max) = Dir
    Loop
    
    f_max = f_max - 1
    
    '�t�@�C�������T�O�̏ꍇ�A��ԌÂ��t�@�C�����폜����B
    '�t�@�C�����̏��(�P�O�O�j��ύX�������ꍇ�͈ȉ��̂h�e����f_max�̒l��ύX���Ă��������B
    If f_max = 50 Then
                    
        old_f = ""
        old_t = "99999999999999"
        
        For f_cnt = 1 To f_max
            If Mid(f_name(f_cnt), 11, 14) < old_t Then
                old_f = f_name(f_cnt)
                old_t = Mid(f_name(f_cnt), 11, 14)
            End If
        Next f_cnt
        
        '�t�@�C�����폜����B
        Kill App.Path & "\Copy\" & old_f

    End If
    '2000.08.16 �ÎR END
    
    f_path = App.Path & "\Copy\" & msPGCAP & "_" & msNO & "_" & Format(Now(), "YYYYMMDDhhnnss") & ".csv"
    
    Open f_path For Output As #1
    
    '�V�[�g��maxrows�����[�v
    For i = 0 To rows_cnt       '�O�̓w�b�_
        data_wk = ""
        msOBJ.row = i
        For j = 0 To col_cnt
            msOBJ.col = j
            data_wk = data_wk & Chr(34) & msOBJ.text & Chr(34) & ","
            DoEvents
        Next j
        
        '�o�͂b�r�u�t�@�C����1�s��������
        Print #1, data_wk
    Next i
    
    Screen.MousePointer = vbDefault
    
    '�o�͂b�r�u�t�@�C���̂b�k�n�r�d
    Close #1
    
On Error GoTo 0

    SPRD_PRT = 0
        
Exit Function
    
Err:
        
    Screen.MousePointer = vbDefault
    MsgLog ("SPREAD SHEET�̏o�͂Ɏ��s���܂����B::" & f_path & Chr(13) & Chr(10))

End Function

' @(f)
'
' �@�\      :   �����H����������
'
' �Ԃ�l    :   TRUE�F�����AFALSE�F���s
'
' ����      :   sGamen  -   ��ʃR�[�h
'               sKCode1 -   ��ƍH���P
'               sKCode2 -   ��ƍH���Q
'
' �@�\����  :   ��ʃR�[�h���珈���H���R�[�h�i��ƍH���R�[�h�j����������ƍH���P�Ɋi�[����
'
' ���l      :   �P��ʂ�2�񏈗����s���ꍇ�i���т��Q�j�܂��́A�����̉�ʂ��������̎d�|�ɑ΂���
'               ������s���悤�ȏꍇ�̍H���R�[�h�͍�ƍH���Q�ɓ����Ă���B
'               �������A�����������H�̏ꍇ�͍�ƍH���P��AGR�A��ƍH���Q��MGR�̍H���R�[�h��o�^���Ă��܂�
'
'Public Function GetMyProcessCode(sGamen As String) As Boolean
'    Dim sSql    As String
'    Dim objDS   As Object
'    Dim sCode   As String
'    GetMyProcessCode = False
'    sCode = Left$(sGamen, 6) & "0"
'    sSql = "SELECT  NVL(kcode01a9,' '), "
'    sSql = sSql & " NVL(kcode02a9,' '), "
'    sSql = sSql & " NVL(kcode03a9,' '), "
'    sSql = sSql & " NVL(kcode04a9,' '), "
'    sSql = sSql & " NVL(kcode05a9,' ')  "
'    sSql = sSql & " FROM koda9          "
'    sSql = sSql & " WHERE   codea9  =   '" & sCode & "' "
'    sSql = sSql & "   AND   shuca9  =   '95'            "
'    sSql = sSql & "   AND   sysca9  =   'K'             "
'    If DynSet(objDS, sSql) = False Then
'        Call MsgOut(100, sSql, ERR_DISP_LOG, "koda9")
'        Exit Function
'    End If
'    If objDS.EOF = False Then
'        Do Until objDS.EOF
'            gsProcCode1 = objDS(0)
'            gsProcCode2 = objDS(1)
'            gsProcCode3 = objDS(2)
'            gsProcCode4 = objDS(3)
'            gsProcCode5 = objDS(4)
'            objDS.MoveNext
'        Loop
'    Else
'        gsProcCode1 = ""
'        gsProcCode2 = ""
'        gsProcCode3 = ""
'        gsProcCode4 = ""
'        gsProcCode5 = ""
'    End If
'    GetMyProcessCode = True
'End Function


' @(f)
'
' �@�\      :   �L����������
'
' �Ԃ�l    :   �Ȃ�
'
' ����      :�@ �L�������A�f�[�^
'
' ���l      :
'
Function keta(ketasu As Integer, motodata As Variant) As Variant
Dim yukocnt As Integer
Dim ln As Integer
Dim lp As Integer
Dim ld As Integer
Dim lz As Integer
Dim work
Dim moji
Dim oflg As Integer

    '�����l�Z�b�g
    ld = 0
    lz = 0
    oflg = 0
    yukocnt = 0
    work = motodata
    If InStr(work, ".") = 0 Then
        work = work & ".0"
        motodata = work
    End If
    ln = Len(work)
    

    If Format(work, "###0.0####") = "0.0" Then
        keta = " "
        Exit Function
    End If

    '���f�[�^�����L�����������[�v���Ȃ���
    For lp = 1 To ln
        moji = Mid(work, lp, 1)
        Debug.Print moji
        If moji <> 0 And moji <> "." And moji <> "-" And moji <> "+" Then
            yukocnt = yukocnt + 1
        ElseIf yukocnt > 0 And moji <> "." And moji <> "-" And moji <> "+" Then
            yukocnt = yukocnt + 1
        End If
        
        If yukocnt >= ketasu Then
            oflg = 1
            Exit For
        End If
    Next lp
    
'' �����_�ʒu����
    For ld = 1 To ln
        moji = Mid(work, lp, 1)
        If moji = "." Then Exit For
    Next ld

    keta = Mid(motodata, 1, lp)

'' �L�������s�����h�O�h����
    If oflg = 0 Then
        For lz = 1 To ketasu - yukocnt
            keta = keta & "0"
            lp = lp + 1
        Next lz
    End If
    
'' �������s���������h�O�h����
    ld = ld - 1
    If ld > lp Then
        For lz = 1 To lp - ld
            keta = keta & "0"
        Next lz
    End If

End Function


'///////////////////////////////////////////////////
' @(f)
'
' �@�\      :   �n�e�ʒu�^�m�b�`�ʒu�^�b�e�ʒu�����ϊ�����
'
' �Ԃ�l    :   ����F" 0 0 0"�`"-1-1-1"
'               �ُ�F"ERR"
'
' ����      :   �n�e�ʒu�^�m�b�`�ʒu�^�b�e�ʒu
'               ���͗����ږ��@"�n�e�ʒu"�Ȃ�
'               �����͋��t���O�@True:�����͉�  False:�����͕s��
'
' ���l      :   �n�e�ʒu�^�m�b�`�ʒu�^�b�e�ʒu���͗��̌������U��������
'
'///////////////////////////////////////////////////
Public Function PosChkCnv(ctlControl As Control, _
                          sMsg As String, _
                          Optional bUnInput As Boolean = False) As String
    Dim sPos As String      ''�ʒu
    Dim sCnvPos As String   ''�ʒu�i�ϊ��j
    Dim sChr As String      ''1�����؂�o��
    Dim iIdx As Integer     ''�C���f�b�N�X
    Dim iNumCnt As Integer  ''���l�J�E���^
    Dim bHifen As Boolean   ''�n�C�t���t���O
    
'    PosChkCnv = "ERR"
    
    ''�R���g���[������
    If (TypeOf ctlControl Is TextBox) Or _
       (TypeOf ctlControl Is ComboBox) Then
        ''�l�擾
        sPos = ctlControl.text
    Else
        ''�l�擾
        sPos = ctlControl.Caption
    End If
    
    PosChkCnv = sPos
    ''�����`�F�b�N
    If bUnInput And (sPos = "") Then    ''�����͋��Ŗ����͂Ȃ�
        PosChkCnv = sPos
        Exit Function
    ElseIf Len(sPos) = 6 Then           ''�ŏ�����U���̏ꍇ
        ''�m�[�`�F�b�N
        Call CtrlEnabled(ctlControl, NORMAL_CTL)
        PosChkCnv = sPos
        Exit Function
    ElseIf Len(sPos) > 6 Then
        Call CtrlEnabled(ctlControl, RED_CTL)
        Call MsgOut(0, sMsg & "���U���ȏ�ł�", ERR_DISP)
        Exit Function
    ElseIf Len(sPos) < 3 Then
        Call CtrlEnabled(ctlControl, RED_CTL)
        Call MsgOut(0, sMsg & "���R�������ł�", ERR_DISP)
        Exit Function
    End If
    
    ''������`�F�b�N
    For iIdx = 1 To Len(sPos)
        Select Case Mid(sPos, iIdx, 1)
        Case " ", "-", "0" To "9"
        Case Else
            Call CtrlEnabled(ctlControl, RED_CTL)
            Call MsgOut(0, sMsg & "�ɕs���ȕ������܂܂�Ă��܂�", ERR_DISP)
            Exit Function
        End Select
    Next
    
    ''�ϊ�
    For iIdx = 1 To Len(sPos)
        sChr = Mid(sPos, iIdx, 1)               ''�؂�o��
        If bHifen Then                          ''�O���n�C�t���̏ꍇ
            If IsNumeric(sChr) Then             ''���l�Ȃ�
                sCnvPos = sCnvPos & sChr        ''���̐��l���i�[
                bHifen = False                  ''���񐔒l
                iNumCnt = iNumCnt + 1           ''���l���J�E���g
            Else                                ''�n�C�t���Q�A��
                Call CtrlEnabled(ctlControl, RED_CTL)
                Call MsgOut(0, sMsg & "�̕����̕��т��s���ł�", ERR_DISP)
                Exit Function
            End If
        Else                                    ''�O�����l�̏ꍇ
            If IsNumeric(sChr) Then             ''���l�Ȃ�
                sCnvPos = sCnvPos & " " & sChr  ''�󔒁����l�i�[
                bHifen = False                  ''���񐔒l���Z�b�g
                iNumCnt = iNumCnt + 1           ''���l���J�E���g
            Else                                ''���l�ȊO�Ȃ�
                sCnvPos = sCnvPos & sChr        ''���̕������i�[
                bHifen = True                   ''����n�C�t�����Z�b�g
            End If
        End If
    Next
    If bHifen Then  ''�Ōオ���l�ȊO�̏ꍇ
        Call CtrlEnabled(ctlControl, RED_CTL)
        Call MsgOut(0, sMsg & "�̕����̕��т��s���ł�", ERR_DISP)
        Exit Function
    ElseIf iNumCnt > 3 Then ''���l���R���𒴂�����
        Call CtrlEnabled(ctlControl, RED_CTL)
        Call MsgOut(0, sMsg & "�̐��l���R���ȏ�܂܂�Ă��܂�", ERR_DISP)
        Exit Function
    ElseIf iNumCnt < 3 Then ''���l���R������
        Call CtrlEnabled(ctlControl, RED_CTL)
        Call MsgOut(0, sMsg & "�̐��l���R�������ł�", ERR_DISP)
        Exit Function
    End If
    
    ''����
    Call CtrlEnabled(ctlControl, NORMAL_CTL)
    PosChkCnv = sCnvPos
End Function


'///////////////////////////////////////////////////
' @(f)
'
' �@�\      :   �e�L�X�g�{�b�N�X�������ߏ���
'
' �Ԃ�l    :   �������ʂ̕�����
'
' ����      :   �e�L�X�g�{�b�N�X�R���g���[��
'               �������߂��錅��
'               �������߂��镶��
'
' ���l      :   �������߂��錅���ȗ�����MaxLength�܂Ŗ��߂�
'               MaxLength���ݒ肳��ĂȂ��ꍇ�P�Q���܂Ŗ��߂�
'               �������߂��镶���ȗ�����"0"�Ŗ��߂�
'
'///////////////////////////////////////////////////
Public Function TextBoxDap(ctlTextBox As Control, Optional ByVal iColumn As Integer, Optional ByVal sChar As String) As String
    Dim iCol As Integer     ''�ݒ肷�錅��
    Dim iLackCnt As Integer ''�s������
    ''�������w�肳��Ă�����
    If iColumn Then
        iCol = iColumn
    ''MaxLength���ݒ肳��Ă�����
    ElseIf ctlTextBox.MaxLength Then
        iCol = ctlTextBox.MaxLength
    ''�����ȗ����AMaxLength���ݒ肳��ĂȂ��ꍇ
    Else
        iCol = 12
    End If
    ''�s�������Z�o
    iLackCnt = iCol - Len(ctlTextBox)
    ''�s�����Ă�����
    If iLackCnt > 0 Then
        ''�������߂��镶�����ȗ����ꂽ�ꍇ"0"�ɂ���
        If sChar = "" Then sChar = "0"
        ''�������߂���
        ctlTextBox = ctlTextBox & String(iLackCnt, sChar)
    End If
    TextBoxDap = ctlTextBox
End Function


' @(f)
'
' �@�\      : ORG�f�[�^�ݒ�
'
' �Ԃ�l    : ORG�l
'
' ����      : OiCsData  - ����Oi/Cs�f�[�^
'             BuiNumber - �������ʔԍ�
'             sKeisanFlg -�n�r�e�v�Z�t���O
'
' �@�\����  : ORG�f�[�^���v�Z���ݒ肷��
'
' ���l      : ORG�v�Z
'            �P���b�i���Ӂ|���S�j�b�����S�~�P�O�O  �����Ӗ��Ɍv�Z���A�v�Z���ʂ̍ő�l���g��
'            �Q���b�i���S�`�j�|�i���ӂ̕��ρj�b���i���S�`�j�~�P�O�O
'            �R���i�������|�������j�^�������~�P�O�O
'            �S���b�i���S�`�j�|�i�a�A�b�A�e�A�f�̕��ρj�b���i���S�`�j�~�P�O�O
'            �T���i�b���S�l�|���Ӓl�b�j���i���S�l�{���Ӓl�j�~�Q�O�O  �����Ӗ��Ɍv�Z���A�v�Z���ʂ̍ő�l���g��
'            �V���i�������|�������j�^�������~�P�O�O
'            �W���i���S�`�j�|�i���ӂ̂������j���i���S�`�j�~�P�O�O
''20000908 ���� ���̊֐��S�̂���蒼����
Public Function SetORGData(sOiCsDataIti() As ST_OICS, sBuiNumber As String, Optional sKeisanFlg As String = "1") As String
    Dim iCntIdx  As Integer ''�Z���^�[�l�̲��ޯ��
    Dim iIdx     As Integer ''���ޯ��
    Dim iPnt     As Integer ''����_"A"�`"I"�͂O�`�W
    Dim iCnt     As Integer ''���͂��ꂽ����_���J�E���g
    Dim dOrg     As Double  ''�ŏI�n�q�f�v�Z����
    Dim dOrgs(8) As Double  ''����_���̂n�q�f�v�Z����
    Dim sOi(8)   As String  ''�����ʂ̊e����_���̎_�f�F�ް������͒����O�̕�����
    Dim dAveOi   As Double  ''���ϒl
    Dim dMaxOi   As Double  ''�ő�l
    Dim dMinOi   As Double  ''�ŏ��l
    Dim bValGet  As Boolean ''�l�擾�t���O
    
    ''���ꕔ�ʂ̊e����_�̎_�f�l���擾���鏈��
    For iIdx = 0 To UBound(sOiCsDataIti)
        ''���ʂ�������������
        If val(sBuiNumber) = val(sOiCsDataIti(iIdx).sCryBuiNo) Then
            ''����_"A"�`"I"���O�`�W�ɕϊ�
            Select Case sOiCsDataIti(iIdx).sMenPosIti
            Case "A": iPnt = 0: iCntIdx = iIdx ''�Z���^�[�̲��ޯ���擾
            Case "B": iPnt = 1
            Case "C": iPnt = 2
            Case "D": iPnt = 3
            Case "E": iPnt = 4
            Case "F": iPnt = 5
            Case "G": iPnt = 6
            Case "H": iPnt = 7
            Case "I": iPnt = 8
            End Select
            ''����_��"A"�`"I"�Ȃ�
            Select Case sOiCsDataIti(iIdx).sMenPosIti
            Case "A", "B", "C", "D", "E", "F", "G", "H", "I"
                ''����_�̎_�f�l�擾
                sOi(iPnt) = sOiCsDataIti(iIdx).sSansoAT
            End Select
        End If
    Next
    
    ''OSF�v�Z�׸ނɂ��g�p����v�Z����؂�ւ���
    Select Case sKeisanFlg
    Case "3", "7", "8"
        ''�ő�l�E�ŏ��l���擾���鏈��
        dMaxOi = -999999  ''�ő�l�ɍŏ��̒l��ݒ�
        dMinOi = 999999   ''�ŏ��l�ɍő�̒l��ݒ�
        bValGet = False   ''�l���擾��ݒ�
            
        For iPnt = 0 To 8
            ''���̑���_�����͂���Ă�����
            If sOi(iPnt) <> "" Then
                ''���̑���l���ő�l���傫�����
                If val(sOi(iPnt)) > dMaxOi Then
                    dMaxOi = val(sOi(iPnt)) ''�ő�l�擾
                    bValGet = True          ''�l�擾��ݒ�
                End If
                ''���̑���l���ŏ��l��菬�������
                If val(sOi(iPnt)) < dMinOi Then
                    dMinOi = val(sOi(iPnt)) ''�ŏ��l�擾
                    bValGet = True          ''�l�擾��ݒ�
                End If
            End If
        Next
        If sKeisanFlg = "8" Then ''�n�r�e�v�Z�t���O��"8"�Ȃ�
            bValGet = False   ''�l���擾��ݒ�
            For iPnt = 1 To 8
                ''���̑���_�����͂���Ă�����
                If sOi(iPnt) <> "" Then
                    ''���̑���l���ŏ��l��菬�������
                    If val(sOi(iPnt)) < dMinOi Then
                        dMinOi = val(sOi(iPnt)) ''�ŏ��l�擾
                        bValGet = True          ''�l�擾��ݒ�
                    End If
                End If
            Next
        End If
        ''�ő�l�ƍŏ��l���擾�ł�����
        If bValGet Then
            ''�n�r�e�v�Z�t���O�ɂ��v�Z����ウ��
            If sKeisanFlg = "8" Then ''�n�r�e�v�Z�t���O��"8"�Ȃ�
                ''���S�l�͕K���K�v
                If sOi(0) = "" Then    ''�Ȃ����
                    Exit Function      ''������
                End If
                ''�i���S�|���Ӎŏ��l�j�^���S�~�P�O�O
                dOrg = (val(sOi(0)) - dMinOi) / val(sOi(0)) * 100
            ElseIf sKeisanFlg = "3" Then ''�n�r�e�v�Z�t���O��"3"�Ȃ�
                ''�i����ő�l�|����ŏ��l�j�^����ő�l�~�P�O�O
                dOrg = (dMaxOi - dMinOi) / dMaxOi * 100
            Else                     ''�n�r�e�v�Z�t���O��"7"�Ȃ�
                ''�i����ő�l�|����ŏ��l�j�^����ŏ��l�~�P�O�O
                dOrg = (dMaxOi - dMinOi) / dMinOi * 100
            End If
            ''�����߂�l�Ƃ���
            SetORGData = left(CStr(dOrg), 6)
        End If
        
    Case "2", "4"
        ''���S�l�͕K���K�v
        If sOi(0) = "" Then    ''�Ȃ����
            Exit Function      ''������
        End If
        ''���ϒl���擾���鏈��
        dAveOi = 0        ''���ϒl�N���A
        iCnt = 0          ''���͌����N���A
        For iPnt = 1 To 8 ''���Ӓl�F����_"B"�`"I"
            If sKeisanFlg = "2" Then ''�n�r�e�v�Z�t���O��"2"�Ȃ�
                ''���̑���_�����͂���Ă���A����_��"B"�`"I"�Ȃ�
                If (sOi(iPnt) <> "") Then
                    dAveOi = dAveOi + val(sOi(iPnt))  ''�Ώےl�����Z
                    iCnt = iCnt + 1                   ''�Ώۂ̓��͌����J�E���g
                End If
            End If
            If sKeisanFlg = "4" Then ''�n�r�e�v�Z�t���O��"4"�Ȃ�
                ''���̑���_�����͂���Ă���A����_��"B"�E"C"�E"F"�E"G"�Ȃ�
                If (sOi(iPnt) <> "") And ( _
                    (iPnt = 1) Or _
                    (iPnt = 2) Or _
                    (iPnt = 5) Or _
                    (iPnt = 6)) Then
                    dAveOi = dAveOi + val(sOi(iPnt))  ''�Ώےl�����Z
                    iCnt = iCnt + 1                   ''�Ώۂ̓��͌����J�E���g
                End If
            End If
        Next
        ''���ӂ̒l���擾�ł�����
        If iCnt > 0 Then
            ''���ϒl�Z�o
            If dAveOi <> 0 Then
                dAveOi = dAveOi / iCnt
            End If
            ''��Βl�i���S�l�j�|�i���ӂ̕��ϒl�j�����S�l�~�P�O�O
            dOrg = Abs(val(sOi(0)) - dAveOi) / val(sOi(0)) * 100
            ''�����߂�l�Ƃ���
            SetORGData = left(CStr(dOrg), 6)
        End If
        
    Case Else
        ''���S�l�͕K���K�v
        If sOi(0) = "" Then    ''�Ȃ����
            Exit Function      ''������
        End If
        ''���ӂ̑���_���̂n�q�f���v�Z���鏈��
        dOrg = -999999    ''�n�q�f�ɍŏ��l��ݒ�
        bValGet = False   ''�l���擾��ݒ�
        For iPnt = 1 To 8 ''���Ӓl�F����_"B"�`"I"
            ''���̑���_�����͂���Ă�����
            If sOi(iPnt) <> "" Then
                ''�n�r�e�v�Z�t���O�ɂ��v�Z����ウ��
                If sKeisanFlg = "5" Then ''�n�r�e�v�Z�t���O��"5"�Ȃ�
                    ''��Βl�i���S�l�|���Ӓl�j���i���S�l�{���Ӓl�j�~�Q�O�O
                    dOrgs(iPnt) = Abs(val(sOi(0)) - val(sOi(iPnt))) / (val(sOi(0)) + val(sOi(iPnt))) * 200
                Else                     ''�n�r�e�v�Z�t���O��"1"��""�Ȃ�
                    ''��Βl�i���Ӓl�|���S�l�j�����S�l�~�P�O�O     8/16 Yam�@100��������Round�������悤�ɏC��
                    dOrgs(iPnt) = Round((Abs(val(sOi(iPnt)) - val(sOi(0))) / val(sOi(0))) * 100, 1)
                End If
                ''��ԑ傫���n�q�f���擾
                If dOrg < dOrgs(iPnt) Then
                    dOrg = dOrgs(iPnt)
                    bValGet = True       ''�l�擾��ݒ�
                End If
            End If
        Next
        ''�ő�l���擾�ł�����
        If bValGet Then
            ''�����߂�l�Ƃ���
            SetORGData = left(CStr(dOrg), 6)
        End If
        
    End Select
    
    ''�Z���^�[�l���L���
    If sOi(0) <> "" Then
        ''�n�q�f��ݒ肷��
        sOiCsDataIti(iCntIdx).sORGNo = SetORGData
    End If
End Function

' @(f)
'
' �@�\      :   �}�C�i�X�␳�֐�
'
' �Ԃ�l    :   �}�C�i�X�␳��̐��l������
'
' ������    :   ARG1        - �Z���^�[�l
'               ARG2        - �}�C�i�X�␳�l
'
' �@�\����  :   �}�C�i�X�␳���s��������ɂĒl��Ԃ�
'
' ���l      :
'
Public Function ArgmentFormat(sDat As String) As String
    Dim iDat    As Integer
    Dim iDo     As Integer
    Dim iFun    As Integer
    If sDat <> "" Then
        iDat = CInt(val(sDat))
        iDo = iDat \ 60
        iFun = (iDat Mod 60)
        If iDo > -1 And iFun < 0 Then     '�}�C�i�X�̏ꍇ�̏C���@1/9 Yam
            iFun = Abs(iDat Mod 60)
            ArgmentFormat = Format$(iDo, "-#0��") & Format$(iFun, "00��")
        Else
            iFun = Abs(iDat Mod 60)
            ArgmentFormat = Format$(iDo, "#0��") & Format$(iFun, "00��")
        End If
    Else
        ArgmentFormat = ""
    End If
End Function


' @(f)
'
' �@�\      :   �n�e�ʒu�ϊ������i�V�\�������\���j
'
' �Ԃ�l    :   �ϊ���̂n�e�ʒu�i���\���j
'
' ������    :   ARG1        - ������
'               ARG2        - �n�e�ʒu�i�V�\���j
'
' �@�\����  :   �n�e�ʒu�̕ϊ����s��������ɂĒl��Ԃ�
'
' ���l      :
'
Public Function OfposChg(Xjiku As String, Ofposo As String) As String
    
    OfposChg = ""
    If Ofposo = "1" Then OfposChg = "110"
    If Ofposo = "2" Then OfposChg = "110"
    If Ofposo = "3" Then OfposChg = "110"
    If Ofposo = "4" Then OfposChg = "110"
    If Ofposo = "5" Then OfposChg = "100"
    If Ofposo = "6" Then OfposChg = "100"
    If Ofposo = "7" Then OfposChg = "100"
    If Ofposo = "8" Then OfposChg = "100"
    If Ofposo = "9" Then OfposChg = "110"
    If Ofposo = "10" Then OfposChg = "110"
    If Ofposo = "11" Then OfposChg = "2111"
    If Ofposo = "12" Then OfposChg = "2112"
    If Ofposo = "13" Then OfposChg = "111"
    If Ofposo = "14" Then OfposChg = "111"
    If Ofposo = "15" Then OfposChg = "111"
    If Ofposo = "16" Then OfposChg = "111"
    If Ofposo = "17" Then OfposChg = "211"
    If Ofposo = "18" Then OfposChg = "211"
    If Ofposo = "19" Then OfposChg = "211"
    If Ofposo = "20" Then OfposChg = "211"
    If Ofposo = "21" Then OfposChg = "111"
    If Ofposo = "22" Then OfposChg = "111"
    If Ofposo = "23" Then OfposChg = "110"
    If Ofposo = "24" Then OfposChg = "111"
    If Ofposo = "1" And Xjiku = "511" Then OfposChg = "1101"
    If Ofposo = "9" And Xjiku = "111" Then OfposChg = "1101"
    If Ofposo = "10" And Xjiku = "111" Then OfposChg = "1102"

End Function



' @(f)
'
' �@�\      :   �n�e�ʒu�}�ԍ����菈��(�n�e�p�^�[���j
'
' �Ԃ�l    :   �����̂n�e�ʒu�}�ԍ�
'
' ������    :   Xjiku        - ������
'               Ofpos        - �n�e�ʒu�i�V�\���j
'               Cfkaku       - �n�e�p�x�i�V�\���j
'               Cfpos        - �n�e�ʒu�i�V�\���j
'               Nopos        - �m�b�`�ʒu�i�V�\���j
'               Ocnflg       - OF,CF,ɯ��v��s�v
'               Cfkijyn      - �b�e�w���i�V�\���j
'
' �@�\����  :   �n�e�ʒu�}�ԍ��̔�����s��������ɂĒl��Ԃ�
'
' ���l      :
'
Public Function OfposBangoChg(Xjiku As String, Ofpos As String, Cfkaku As String, Cfpos As String, _
                              Nopos As String, Ocnflg As String, Cfkijyn As String) As String
    
    OfposBangoChg = "99"
    If Ocnflg = "" Or Ocnflg = " " Or Ocnflg = "0" Then OfposBangoChg = "00"
    If Ocnflg = "1" Then
        If Xjiku = "111" And Ofpos = "9" Then OfposBangoChg = "01" Else
        If Xjiku = "111" And Ofpos = "10" Then OfposBangoChg = "02" Else
        If Xjiku = "111" And Ofpos = "12" Then OfposBangoChg = "03" Else
        If Xjiku = "111" And Ofpos = "11" Then OfposBangoChg = "04" Else
        If Xjiku = "511" And Ofpos = "1" Then OfposBangoChg = "05" Else
        If Xjiku = "100" And Ofpos = "1" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Ofpos = "2" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Ofpos = "3" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Ofpos = "4" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Ofpos = "5" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Ofpos = "6" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Ofpos = "7" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Ofpos = "8" Then OfposBangoChg = "07" Else
        If Xjiku = "110" And Ofpos = "9" Then OfposBangoChg = "08" Else
        If Xjiku = "110" And Ofpos = "10" Then OfposBangoChg = "08" Else
        If Xjiku = "110" And Ofpos = "5" Then OfposBangoChg = "09" Else
        If Xjiku = "110" And Ofpos = "6" Then OfposBangoChg = "09" Else
        If Xjiku = "110" And Ofpos = "13" Then OfposBangoChg = "10" Else
        If Xjiku = "110" And Ofpos = "15" Then OfposBangoChg = "10" Else
        If Xjiku = "   " Or Xjiku = "" Then OfposBangoChg = "98"
    End If

    If Ocnflg = "3" Then
        If Xjiku = "111" And Nopos = "9" Then OfposBangoChg = "01" Else
        If Xjiku = "111" And Nopos = "10" Then OfposBangoChg = "02" Else
        If Xjiku = "111" And Nopos = "12" Then OfposBangoChg = "03" Else
        If Xjiku = "111" And Nopos = "11" Then OfposBangoChg = "04" Else
        If Xjiku = "511" And Nopos = "1" Then OfposBangoChg = "05" Else
        If Xjiku = "100" And Nopos = "1" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Nopos = "2" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Nopos = "3" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Nopos = "4" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Nopos = "5" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Nopos = "6" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Nopos = "7" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Nopos = "8" Then OfposBangoChg = "07" Else
        If Xjiku = "110" And Nopos = "9" Then OfposBangoChg = "08" Else
        If Xjiku = "110" And Nopos = "10" Then OfposBangoChg = "08" Else
        If Xjiku = "110" And Nopos = "5" Then OfposBangoChg = "09" Else
        If Xjiku = "110" And Nopos = "6" Then OfposBangoChg = "09" Else
        If Xjiku = "110" And Nopos = "13" Then OfposBangoChg = "10" Else
        If Xjiku = "110" And Nopos = "15" Then OfposBangoChg = "10" Else
        If Xjiku = "   " Or Xjiku = "" Then OfposBangoChg = "98"
    End If

    If Ocnflg = "2" And Cfkijyn = "1" Then
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 10800 Then OfposBangoChg = "11" Else
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 5400 Then OfposBangoChg = "12" Else
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 2700 Then OfposBangoChg = "13" Else
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 18900 Then OfposBangoChg = "14" Else
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 13500 Then OfposBangoChg = "15" Else
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 8100 Then OfposBangoChg = "16" Else
        If Xjiku = "111" And Ofpos = "12" And Cfkaku = 10800 Then OfposBangoChg = "17" Else
        If Xjiku = "111" And Ofpos = "12" And Cfkaku = 16200 Then OfposBangoChg = "18" Else
        If Xjiku = "111" And Ofpos = "11" And Cfkaku = 13500 Then OfposBangoChg = "19" Else
        If Xjiku = "111" And Ofpos = "11" And Cfkaku = 8100 Then OfposBangoChg = "20" Else
        If Xjiku = "100" And Ofpos = "1" And Cfkaku = 10800 Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "2" And Cfkaku = 10800 Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "3" And Cfkaku = 10800 Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "4" And Cfkaku = 10800 Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "1" And Cfkaku = 5400 Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "2" And Cfkaku = 5400 Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "3" And Cfkaku = 5400 Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "4" And Cfkaku = 5400 Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "1" And Cfkaku = 13500 Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "2" And Cfkaku = 13500 Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "3" And Cfkaku = 13500 Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "4" And Cfkaku = 13500 Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "1" And Cfkaku = 8100 Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "2" And Cfkaku = 8100 Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "3" And Cfkaku = 8100 Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "4" And Cfkaku = 8100 Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "1" And Cfkaku = 2700 Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "2" And Cfkaku = 2700 Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "3" And Cfkaku = 2700 Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "4" And Cfkaku = 2700 Then OfposBangoChg = "25"
    End If

    If Ocnflg = "2" And Cfkijyn = "2" Then
        If Xjiku = "111" And Ofpos = "9" And Cfpos = "10" Then OfposBangoChg = "11" Else
        If Xjiku = "111" And Ofpos = "9" And Cfpos = "12" Then OfposBangoChg = "12" Else
        If Xjiku = "111" And Ofpos = "12" And Cfpos = "11" Then OfposBangoChg = "17" Else
        If Xjiku = "111" And Ofpos = "12" And Cfpos = "9" Then OfposBangoChg = "18" Else
        If Xjiku = "100" And Ofpos = "2" And Cfpos = "1" Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "1" And Cfpos = "2" Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "3" And Cfpos = "4" Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "4" And Cfpos = "3" Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "2" And Cfpos = "3" Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "1" And Cfpos = "4" Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "3" And Cfpos = "1" Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "4" And Cfpos = "2" Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "2" And Cfpos = "6" Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "1" And Cfpos = "5" Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "3" And Cfpos = "8" Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "4" And Cfpos = "7" Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "2" And Cfpos = "7" Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "1" And Cfpos = "8" Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "3" And Cfpos = "6" Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "4" And Cfpos = "5" Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "2" And Cfpos = "5" Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "1" And Cfpos = "6" Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "3" And Cfpos = "7" Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "4" And Cfpos = "8" Then OfposBangoChg = "25"
    End If

End Function

'///////////////////////////////////////////////////
'
' @(f)
' �@�\         : ����{�^������
' �Ԃ�l       : �Ȃ�
' ������       : �t�H�[����
' �@�\����     : ����{�^������
'                <2001.02.23 kuro>
'
' �����è���� �F�@Copies   :�@�������
'                 FromPage :�@����J�n�y�[�W
'                 ToPage   :�@����I���y�[�W
'                 hDC      :�@�I�����ꂽ�v�����^�̃f���@�C�X�R���e�L�X�g
'
'///////////////////////////////////////////////////
'
Public Function mdlHard_Copy(Getform As Form)
    Dim BeginPage, EndPage, NumCopies, i
    
    With Getform
        .CommonDialog1.CancelError = True
    
        On Error GoTo ErrHandler
    
        '�_�C�A���O�{�b�N�X�́m�y�[�W�w��n�I�v�V�����{�^���𖳌��ɂ���
        .CommonDialog1.Flags = &H8&
    
        '�_�C�A���O�{�b�N�X�́m�I�����������n�I�v�V�����{�^���𖳌��ɂ���
        .CommonDialog1.Flags = &H4&
    
        '�m����n�_�C�A���O�{�b�N�X��\��
        .CommonDialog1.ShowPrinter
    
        '���[�U�[�̑I�������l���_�C�A���O�{�b�N�X����擾
        'BeginPage = .CommonDialog1.FromPage   ''�X�^�[�g�y�[�W
        'EndPage = .CommonDialog1.ToPage       ''�G���h�y�[�W
        NumCopies = .CommonDialog1.Copies      ''�������
        For i = 1 To NumCopies
            '����t�H�[�����v�����^�ɑ��M
            .PrintForm
            Printer.EndDoc
        Next i
        Exit Function
    End With

ErrHandler:
    '���[�U�[���m�L�����Z���n���N���b�N���܂���
    Exit Function
End Function

'///////////////////////////////////////////////////
' @(f)
' �@�\    : �o�[�W�����`�F�b�N�֐�
'
' �Ԃ�l  : True :�ް�ޮ݈�v
'           False:�ް�ޮݎ擾���s�^�ް�ޮ݂��Â�
' ������  :
'
' �@�\����: �c�a�ɓo�^����Ă���o�[�W������
'           ���W���[���̃o�[�W��������v���Ă��邩��r���A
'           ���W���[���̃o�[�W�������Â��ꍇ
'           �t�H�[���́uҲ��ƭ��v�{�^���ȊO�̺��۰ق�
'           �g�p�s�ɂ���
'
'///////////////////////////////////////////////////
Public Function VerChk(frmCurrent As Form) As Boolean
    Dim objOraDyn As Object         ''�_�C�i�Z�b�g�I�u�W�F�N�g
    Dim sSql As String              ''�r�p�k��
    Dim iMajorx As String           ''�c�a�擾���W���[�o�[�W����
    Dim iMajor As Integer           ''�c�a�擾���W���[�o�[�W����
    Dim iMinor As Integer           ''�c�a�擾�}�C�i�[�o�[�W����
    Dim iRevision As Integer        ''�c�a�擾���r�W����
    
    gbFTPFlg = False    ''FTP�N���t���O�N���A
        
    ''DB�o�^�ް�ޮݎ擾
    ''�r�p�k���쐬
    sSql = "SELECT   NVL(ctr01a9,0),    "
    sSql = sSql & "  NVL(ctr02a9,0),    "
    sSql = sSql & "  NVL(ctr03a9,0)     "
    sSql = sSql & "FROM  koda9         "
    sSql = sSql & "WHERE sysca9 = 'K'  "
    sSql = sSql & "AND   shuca9 = '01' "
    sSql = sSql & "AND   codea9 = '" & UCase(gsEXEName) & "'"
    
    If DynSet2(objOraDyn, sSql) = False Then
        ''�擾���s
        Call MsgOut(100, sSql, ERR_DISP_LOG, "koda9")
        VerChk = False
        GoTo Er
    End If
    If objOraDyn.EOF Then
        ''������Ȃ�����
        Call MsgOut(55, "�ް�ޮݏ��", ERR_DISP)
        VerChk = False
        GoTo Er
    End If
    'GetMyProcessCode frmCurrent.Caption
    
    iMajor = objOraDyn(0)
    iMinor = objOraDyn(1)
    iRevision = objOraDyn(2)
    
    If iMajor = 0 And iMinor = 0 And iRevision = 0 Then
       VerChk = True
       Exit Function
    End If
    
    ''�o�[�W�����s��v���Z�b�g
    VerChk = False
    ''���W���[�o�[�W�����`�F�b�N
    If iMajor <> val(App.Major) Then
        Call MsgOut(0, "Ҽެ��ް�ޮݕs��vҲ��ƭ����݉���", ERR_DISP)
        GoTo Er
    End If
    ''�}�C�i�[�o�[�W�����`�F�b�N
    If iMinor <> val(App.Minor) Then
        Call MsgOut(0, "ϲŰ�ް�ޮݕs��vҲ��ƭ����݉���", ERR_DISP)
        GoTo Er
    End If
    ''���r�W�����`�F�b�N
    If iRevision <> val(App.Revision) Then
        Call MsgOut(0, "��޼ޮ��ް�ޮݕs��vҲ��ƭ����݉���", ERR_DISP)
        GoTo Er
    End If
    ''�o�[�W������v���Z�b�g
    VerChk = True
     ''  �����H���R�[�h�擾
    Exit Function
Er:
    ''�o�[�W�����s��v
    Call CtrlCancel(frmCurrent)     ''Ҳ��ƭ��ȊO�̺��۰ق��g���Ȃ�����
    gbFTPFlg = True    ''FTP�N���t���O�Z�b�g
End Function

'///////////////////////////////////////////////////
' @(f)
' �@�\    : �r�����������Ԏ擾
'
' �Ԃ�l  : �r�����������ԁiDate�^�FYYYY/MM/DD�j
'
' ������  : �ϊ������������iDate�^�FYYYY/MM/DD�j
'
' �@�\����: �p�����[�^�̎�������r�����������Ԃɕϊ�����
'
'  �r������������ = �p�����[�^�̓��t �|�@��������
'
'           �G���[���������ꍇ�̓p�����[�^�̓��t��
'           ���̂܂ܖ߂��B
'
'///////////////////////////////////////////////////
Public Function CalcSumcoTime(tParmDate As Date) As Date

    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim vChoseiTime     As Variant          '��������

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR

    '�f�t�H���g�߂�l�ݒ�
    CalcSumcoTime = tParmDate

    'SUMCO���ԍ쐬�ׁ̈A�������Ԏ擾
    sql = "SELECT KCODE01A9"
    sql = sql & " FROM koda9 "
    sql = sql & " WHERE SYSCA9 = 'X'"
    sql = sql & "   AND SHUCA9 = '80'"
    sql = sql & "   AND CODEA9 = '1'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    '���݂��Ȃ����A�����I��
    If rs Is Nothing Then
        Exit Function
    End If
    If Not rs.EOF Then
        If IsNull(rs.Fields("KCODE01A9")) = True Then
            Exit Function
        Else
    '�������Ԏ擾
            vChoseiTime = CDate(rs.Fields("KCODE01A9"))
        End If
    End If
    rs.Close

    'SUMCO����=�p�����[�^�̓��t-KODA9.��������
    CalcSumcoTime = Format(tParmDate - CDate(vChoseiTime), "yyyy/mm/dd")

    Exit Function

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

PROC_ERR:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.Number
    Resume proc_exit
End Function

'///////////////////////////////////////////////////
' @(f)
' �@�\    : �T�[�o�[���Ԏ擾
'
' �Ԃ�l  : �T�[�o�[���ԁiDate�^�FYYYY/MM/DD�j
'
' �@�\����: ORACLE��茻�ݎ��Ԃ��擾����B
'
'
'///////////////////////////////////////////////////
Public Function getSvrTime() As Date
                                
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
                                
    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
                                
    '�f�t�H���g�߂�l�ݒ�(�[�������j
    getSvrTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
                                
    '�T�[�o�[���Ԏ擾
    sql = "SELECT SYSDATE"
    sql = sql & " FROM DUAL "
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
                                
    '���݂��Ȃ����A�����I��
    If rs Is Nothing Then
        Exit Function
    End If
    If Not rs.EOF Then
        If IsNull(rs.Fields("SYSDATE")) = True Then
            Exit Function
        Else
    '���Ԏ擾
            getSvrTime = CDate(rs.Fields("SYSDATE"))
        End If
    End If
    rs.Close
                                
'    'SUMCO����=�p�����[�^�̓��t-KODA9.��������
'    CalcSumcoTime = Format(tParmDate - CDate(vChoseiTime), "yyyy/mm/dd")
'
'    Exit Function
                                
proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function
                                
PROC_ERR:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.Number
    Resume proc_exit
End Function
                                
                                
'*ADD* Ӽޭ�ٓ��� TCS)K.Kunori 2004.11.29 START >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    '2004/9/27tcs Yamauchi �ǉ� start-------------------------------------
    '///////////////////////////////////////////////////
    ' @(f)
    ' �@�\    : �S���Җ��擾(�����������܂ޏꍇ�j
    '           300mm����
    '
    ' �Ԃ�l  : True:����
    '           False:���s
    '
    ' ������  : �S���Һ���
    '
    ' �@�\����: �S���Һ��ނ���S���Җ����擾
    '
    '///////////////////////////////////////////////////
    Public Function GetAuthorityUser_300(ByVal STAFFID As String, ByVal ProcID As String, _
                                                                ByRef sUserName As String) As Boolean
        Dim sSqlStmt As String
        Dim objOraDyn As Object
        
        sUserName = vbNullString       ''�S���Җ��ر
        
        '' �����Ŏw�肳�ꂽ�Ј�ID�ƍH�����ނŌ������s��
        sSqlStmt = ""
        sSqlStmt = sSqlStmt & "select   t1.JFMLNAME,t1.JFSTNAME         " & vbLf
        sSqlStmt = sSqlStmt & "from     TBCMB001 t1,TBCMB004 t4         " & vbLf
        sSqlStmt = sSqlStmt & "where    t1.EXECODE = t4.AUTHCODE        " & vbLf
        sSqlStmt = sSqlStmt & "and      t1.STAFFID = '" & STAFFID & "'  " & vbLf
        sSqlStmt = sSqlStmt & "and      t4.TRANID = '" & ProcID & "'    " & vbLf
        
        ''�_�C�i�Z�b�g�쐬
        If DynSet2(objOraDyn, sSqlStmt) = False Then
            ''�_�C�i�Z�b�g�쐬���s
            Call MsgOut(100, sSqlStmt, ERR_DISP_LOG)
            
            GetAuthorityUser_300 = False
            Exit Function
        End If
        If objOraDyn.EOF Then
            ''�Y������S���Һ��ނ���������
            Call MsgOut(0, "�F��S���҂ł͂���܂���", ERR_DISP)
            
            GetAuthorityUser_300 = False
            Exit Function
        End If
    
        sUserName = NulltoStr(objOraDyn(0)) & Space(1) & NulltoStr(objOraDyn(1))    ''�S���Җ��擾
        
        GetAuthorityUser_300 = True         ''����������Ԃ�
        
    End Function
    
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' �@�\      :   ���ѓ��t�擾
    '
    ' �Ԃ�l    :�@ ����
    '
    ' ����      :�@ �Ȃ�
    '
    ' �@�\����  :   ���ѓ��t���擾����
    '
    ' ���l      :   �擾�����l������د��ϐ��Ɋi�[
    '
    '///////////////////////////////////////////////////
    Public Function GetSysdate() As Boolean
    
        Dim sSql        As String
        Dim objOraDyn   As Object
        
    On Error GoTo ErrHand
    
        GetSysdate = False
        
        sSql = ""
        sSql = sSql & "select to_char(SYSDATE,'YYYY/MM/DD HH24:MI:SS')  " & vbLf
        sSql = sSql & "from dual                                        " & vbLf
        
        ''�޲ž��č쐬
        If DynSet2(objOraDyn, sSql) = False Then
            ''�޲ž��č쐬���s
            Call MsgOut(100, sSql, ERR_DISP_LOG)
            GetSysdate = False
            Exit Function
        End If
        
        ''����د��ϐ��Ɋi�[
        gsSysdate = objOraDyn.Fields(0).Value
        
        '�J��
        If Not objOraDyn Is Nothing Then objOraDyn.Close: Set objOraDyn = Nothing
        
        GetSysdate = True
        Exit Function
        
ErrHand:
    
        '�J��
        If Not objOraDyn Is Nothing Then objOraDyn.Close: Set objOraDyn = Nothing
    
        ''�װ
        Call MsgOut(100, "", ERR_DISP_LOG, "")
    
    End Function
    
    '///////////////////////////////////////////////////
    ' @(f)
    ' �@�\    : �R�}���h���C�������擾�E�ϐ��Z�b�g
    '
    ' �Ԃ�l  : True:���@False:��
    '
    ' ������  :
    '
    ' �@�\����: �R�}���h���C���������擾���E�g�[�N���؏o���E�ϐ��Z�b�g
    '
    '///////////////////////////////////////////////////
    Public Function GetCmdLine_Re() As Boolean
        Dim sCmdLine As String
        
        ''�R�}���h���C���擾
        sCmdLine = Command
        '' 0        1         2
        '' 1234567890123456789012
        ''"99_*******_***********"
        ''�H��R�[�h_�ďo�敪_�i��
        
        ''�Œ�ŃR�}���h���C��������؏o��
        gsFactryCd = left(sCmdLine, 2)    ''�H��R�[�h(2��)
        gsCallCd = Mid(sCmdLine, 4, 7)    ''�ďo�敪(7��)
        gsHinban = Mid(sCmdLine, 12, 11)  ''�i��(11��)
        myFactryCd = Mid(sCmdLine, 24, 2)
        
        If Len(gsFactryCd) <> 2 Then Exit Function
        If Len(gsCallCd) <> 7 Then Exit Function
        If gsHinban = "00000000000" Then gsHinban = ""
        
    '2004/11/12 TCS TAGAWA �ǉ� Start--------------------------------------------------
        Select Case gsFactryCd
            Case "10"               ''��c�H��
                gsSystemCd = SYSTEM_200
            Case "30"               ''����H��
                gsSystemCd = SYSTEM_200
            Case "40"               ''�đ�H��
                gsSystemCd = SYSTEM_200
            Case "42"               ''�R�O�O����
                gsSystemCd = SYSTEM_300
            Case "43"               ''�R�O�O����
                gsSystemCd = SYSTEM_300
            Case "90"               ''�e�X�g��
                gsSystemCd = SYSTEM_200
            Case "91"               ''�e�X�g��(�V) 2007/04/05�ǉ� SETsw kubota
                gsSystemCd = SYSTEM_200
            Case "92"               ''�e�X�g��(����) 2009/11/20�ǉ� SETsw kubota
                gsSystemCd = SYSTEM_200
            Case "AM"               ''���H�� 2009/06/02�ǉ� SSS.Marushita
                gsSystemCd = SYSTEM_200
            Case "93"               ''����A1�e�X�g 2010/04/14�ǉ� SETsw kubota
                gsSystemCd = SYSTEM_200
            Case "94"               ''���e�X�g 2009/06/03�ǉ� SSS.Marushita
                gsSystemCd = SYSTEM_200
            Case Else               ''�O��
                gsSystemCd = SYSTEM_200
        End Select
    '2004/11/12 tcs tagawa �ǉ�  end---------------------------------------------------
    
        GetCmdLine_Re = True
    End Function
    
    '///////////////////////////////////////////////////
    ' @(f) 2004/11/12 TCS TAGAWA
    '
    ' �@�\      : ̫�т̷��߼�ݐݒ�
    '
    ' �Ԃ�l    :�@�Ȃ�
    '
    ' ����      : frmForm       - ̫��
    '             sProgramId    - ��۸���ID
    
    '
    ' �@�\����  :�@̫�т̷��߼�ݐݒ�
    '
    ' ���l      :
    '
    '///////////////////////////////////////////////////
    Public Sub SetFormCaption(frmForm As Form, ByVal sProgramId As String)
    
        ''200mm�̏ꍇ
        If gsSystemCd = SYSTEM_200 Then
            frmForm.Caption = sProgramId & " - " & SYSTEM_NAME_200
        ''300mm�̏ꍇ
        Else
            frmForm.Caption = sProgramId & " - " & SYSTEM_NAME_300
        End If
    
    End Sub
    
    '�T�v      :�v���O�����N�����̏���������
    '����      :
    Public Function InitExe_Re() As FUNCTION_RETURN
        
        '' �v���O�����N�����̏���������
        DoEvents
        
        '' �p�����[�^������
        InitExe_Re = FUNCTION_RETURN_SUCCESS
       
        '' �G���[�o�̓I�u�W�F�N�g�쐬
        Init_ErrHandler_Re
        
        ''�R�}���h���C�������擾
        If GetCmdLine_Re() = False Then
            ''�R�}���h���C����������
            Call MsgOut(64, "", ERR_DISP_LOG)
            Exit Function
        End If
        
           ''���s�t�@�C�����̎擾
        If GetEXEName = "" Then
            ''�R�}���h���C����������
            Call MsgOut(62, "", ERR_DISP_LOG)
            Exit Function
        End If
     
        '' ���d�N���`�F�b�N
        If App.PrevInstance = True Then
            '' ���d�N�����Ă���ꍇ
            '' �G���[���b�Z�[�W�����O�o��
            MsgBox "���łɃv���O�������N������Ă��܂��B", vbOKOnly + vbInformation
            InitExe_Re = FUNCTION_RETURN_FAILURE
            End
        End If
        
        '' �f�[�^�x�[�X�ڑ�
        OraDBOpen
        
        '' �����I��
    
    End Function
    '�T�v      :�v���O�����N�����̏���������
    '����      :���d�N�������� 2008/04/30 �ǉ� Info.Kameda
    Public Function InitExe_Re_Ref() As FUNCTION_RETURN
        
        '' �v���O�����N�����̏���������
        DoEvents
        
        '' �p�����[�^������
        InitExe_Re_Ref = FUNCTION_RETURN_SUCCESS
       
        '' �G���[�o�̓I�u�W�F�N�g�쐬
        Init_ErrHandler_Re
        
        ''�R�}���h���C�������擾
        If GetCmdLine_Re() = False Then
            ''�R�}���h���C����������
            Call MsgOut(64, "", ERR_DISP_LOG)
            Exit Function
        End If
        
           ''���s�t�@�C�����̎擾
        If GetEXEName = "" Then
            ''�R�}���h���C����������
            Call MsgOut(62, "", ERR_DISP_LOG)
            Exit Function
        End If
     
        ''' ���d�N���`�F�b�N
        'If App.PrevInstance = True Then
        '    '' ���d�N�����Ă���ꍇ
        '    '' �G���[���b�Z�[�W�����O�o��
        '    MsgBox "���łɃv���O�������N������Ă��܂��B", vbOKOnly + vbInformation
        '    InitExe_Re = FUNCTION_RETURN_FAILURE
        '    End
        'End If
        
        '' �f�[�^�x�[�X�ڑ�
        OraDBOpen
        
        '' �����I��
    
    End Function
    Private Sub Init_ErrHandler_Re()
        Set gErr = New CErrHandler
        With gErr
            .AppTitle = App.Title
            .Destination = App.Path & "\Err.log"
            .DisplayMsgOnError = True
            .MaxProcStackItems = 20
            .IncludeExpandedInfo = False
        End With
    End Sub
    
    '*** UPDATE START T.TERAUCHI 2004/11/17 �ǉ�
    
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' �@�\    : ���s����SQL�������O�Ɏc��
    '
    ' �Ԃ�l  : ����
    '
    ' ������  :�@sHostName  -�}�V�����i�N�����j
    '           sAppName    -�v���O�������i�N�����j
    '           sFncName    -�֐����i�N�����j
    '           sSQL        -���s�N�G���[
    '           sMemo       -����
    '
    ' �@�\����: ���s����SQL����TBCMC003�e�[�u���ɏ�������
    '
    ' ���l    :
    '
    '///////////////////////////////////////////////////
    Public Function WriteDBLog_Re(ByVal sHostName As String, ByVal sAppName As String, _
                              ByVal sFncName As String, ByVal sSqlLog As String, _
                              ByVal sMemo As String) As Boolean
    
        Dim sSql        As String
        Dim iRet        As Long
        Dim sDbName     As String
        Dim sUID        As String
        Dim sPWD        As String
        Dim bErrFlag    As Boolean
        Dim sTableName  As String
        
    On Error GoTo ErrHand
    
        WriteDBLog_Re = False
        
        Select Case gsFactryCd
        Case "10"               ''��c�H��
            sDbName = "NODA"
            sUID = "oracle"
            sPWD = "oracle"
        Case "30"               ''����H��
            sDbName = "IKNO"
            sUID = "oracle"
            sPWD = "oracle"
        Case "40"               ''�đ�H��
            sDbName = "YONE"
            sUID = "oracle"
            sPWD = "oracle"
        Case "42"               '�f�R�O�O����
            sDbName = "cm1"
            sUID = "cm1"
            sPWD = "cm1"
        Case "43"               '�f�R�O�O����
            sDbName = "cmt"
            sUID = "cm1"
            sPWD = "cm1"
    '2001/02/24 FFC start
        'Case "43"               ''SUMCO
        '    sDbName = "CMT"
        '    sUID = "cm1"
        '    sPWD = "cm1"
        '    gsFactryCd = "40"
    '2001/02/24 FFC end
        Case "90"               ''�e�X�g��
            sDbName = "TEST0"
            sUID = "oracle"
            sPWD = "oracle"
        Case "91"               ''�e�X�g��(�V) 2007/04/05�ǉ� SETsw kubota
            sDbName = "CLA0X"
            sUID = "oracle"
            sPWD = "oracle"
        Case "99"               ''��
            sDbName = "BOIS"
            sUID = "BOIS"
            sPWD = "BOIS"
        Case Else               ''�O��
            sDbName = "oracle"
            sUID = "oracle"
            sPWD = "oracle"
        End Select
        
        ''�I���N���ڑ�
        Set gobjOraSess2 = CreateObject("OracleInProcServer.XOraSession")
        Set gobjOraDB2 = gobjOraSess2.OpenDatabase(sDbName, sUID & "/" & sPWD, 0&)
        
        If sHostName = "" Then
            sHostName = " "
        End If
        
        If sAppName = "" Then
            sAppName = " "
        End If
        
        If sFncName = "" Then
            sFncName = " "
        End If
        
        If sSqlLog = "" Then
            sSqlLog = " "
        End If
        
        If sMemo = "" Then
            sMemo = " "
        End If
        
        sSqlLog = Replace(sSqlLog, "'", "''")
        
        ''�e�[�u�����擾 2005/07/13 tuku
        If gsSystemCd = SYSTEM_200 Then
            sTableName = "KODZL"
        Else
            sTableName = "TBCMC003"
        End If
        
        sSql = ""
        sSql = sSql & "insert into " & sTableName & " (           " & vbLf
        sSql = sSql & "                     L_DATE                  " & vbLf    ''���O���Ƃ�������
        sSql = sSql & "                     ,SEQ                    " & vbLf    ''���O�̃V�[�P���X
        sSql = sSql & "                     ,HOSTNAME               " & vbLf    ''�N�����}�V����
        sSql = sSql & "                     ,APPNAME                " & vbLf    ''�N�����v���O������
        sSql = sSql & "                     ,FNCNAME                " & vbLf    ''�N�����֐���
        sSql = sSql & "                     ,SQL                    " & vbLf    ''SQL�����O
        sSql = sSql & "                     ,MEMO               )   " & vbLf    ''����
        sSql = sSql & "values(              sysdate                 " & vbLf    ''���O���Ƃ�������
        sSql = sSql & "                     ,log_seq.nextval        " & vbLf    ''���O�̃V�[�P���X
        sSql = sSql & "                     ,'" & sHostName & "'    " & vbLf    ''�N�����}�V����
        sSql = sSql & "                     ,'" & sAppName & "'     " & vbLf    ''�N�����v���O������
        sSql = sSql & "                     ,'" & sFncName & "'     " & vbLf    ''�N�����֐���
        sSql = sSql & "                     ,'" & sSqlLog & "'      " & vbLf    ''SQL�����O
        sSql = sSql & "                     ,'" & sMemo & "'    )   " & vbLf    ''����
            
        Set gobjOraSess2 = CreateObject("OracleInProcServer.XOraSession")
            
            
        ''��ݻ޸��݊J�n
        gobjOraSess2.BeginTrans
        
        ''�I���N���r�p�k���s
        iRet = gobjOraDB2.DbExecuteSQL(sSql)
            
        'SQL���s
        If iRet < 0 Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "TBCMC003")
            ''�װ����
            bErrFlag = True
            GoTo ErrHand
        ElseIf iRet = 0 Then
            Call MsgOut(71, "TBCMC003", ERR_DISP_LOG)
            ''�װ����
            bErrFlag = True
            GoTo ErrHand
            Exit Function
        End If
        
        ''�R�~�b�g����
        gobjOraSess2.CommitTrans
        
        ''�I���N���ؒf
        gobjOraDB2.Close
        
        ''���
        Set gobjOraDB2 = Nothing
        Set gobjOraSess2 = Nothing
        
        WriteDBLog_Re = True
        
        Exit Function
    
ErrHand:
    
        ''۰��ޯ�����
        gobjOraSess2.Rollback
        
        ''�I���N���ؒf
        gobjOraDB2.Close
        
        ''���
        Set gobjOraDB2 = Nothing
        Set gobjOraSess2 = Nothing
    
        ''VB�G���[
        If Not bErrFlag Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "")
        End If
        
    End Function
    '2004/9/27tcs Yamauchi �ǉ� end-------------------------------------
    
    '2004/9/17tcs Suenaga �ǉ� start-------------------------------------
    '///////////////////////////////////////////////////
    ' @(f)
    ' �@�\    : ���������Ǘ��̍X�V���s��
    '
    ' �Ԃ�l  : True:�����@False:���s
    '
    ' ������  :  genryoNo  - ����No
    '           genpinNo  - ���i���b�gNo
    '           kouteiCd  - �H���R�[�h
    '           tantoCd   - �S���҃R�[�h
    '           ukeireW   - ����d��
    '           haraiW    - ���o�d��
    '           losW      - ���X�d��
    '           koujyoCd  - �H��R�[�h
    '           ukeireCd  - ����H��R�[�h
    '           haraiCd   - ���o�H��R�[�h
    '           syoumetu  - ���ŋ敪(1:���i���� 2���i���b�gNo����)
    '           shikake   - �d�|�敪(1:�d�|����)
    '           sisutemu  - �V�X�e���敪�R�[�h
    '           noudoK    - �Z�x�敪
    '           noudoT    - �Z�x�l
    '           sousinFlg - �������M�t���O(1:���M 7:���M�Ȃ�)
    '           hasseiFlg - �����t���O(0:�����M 7:���M�Ȃ�)
    '           motoConce - ���Z�x
    '           yoteiFac  - �g�p�\��H��
    '           tanaire   - �I����敪(0:�I���ꖳ�� 1:�I����)
    '           sChg      - �`���[�WNo(CC200/CC300�o�^�p)�@05/08/23 ooba
    '           motoSyscd - �����擾��(200mm/300mm)�@2008/06/16 SET/miyatake
    '
    ' �@�\����: �����ԍ�(XODB1)�A�����d�|�H��(XODB2)�A�����H������(XODB#)
    '          �̍X�V���s��
    '
    '///////////////////////////////////////////////////
    Public Function Upd_XODB123(genryoNo, _
                                genpinNo, _
                                kouteiCd, _
                                tantoCd, _
                                ukeireW, _
                                HARAIW, _
                                losW, _
                                koujyoCd, _
                                ukeireCd, _
                                haraiCd, _
                                syoumetu, _
                                shikake, _
                                sisutemu, _
                                noudoK, _
                                noudoT, _
                                sousinFlg, _
                                hasseiFlg, _
                                motoConce, _
                                yoteiFac, _
                                tanaire, _
                                Optional sChg As String = "", _
                                Optional motoSyscd As String = "") As Boolean
        
        Dim objOraDyn As Object     '�_�C�i�Z�b�g�I�u�W�F�N�g
        Dim sSql      As String     'SQL��
        Dim sUserName As String     '�S���Җ�
        Dim nowdate   As String     '�V�X�e�����t
        Dim cyoku     As String     '���敪
        Dim year      As String     '���ѓ��@�N
        Dim month     As String     '�@�@�@�@��
        Dim day       As String     '�@�@�@�@��
        Dim hour      As String     '�@�@�@�@��
        Dim Min       As String     '�@�@�@�@��
        Dim before    As String     '�O�H��
        Dim after     As String     '���H��
        Dim sikakeK   As String     '�d�|�H��
        
        '�߂�l�ݒ�
        Upd_XODB123 = False
        
    '�G���[�n���h��
    On Error GoTo ErrHand
        
        '����No���Ȃ���
        If genryoNo = "" Then
            '���b�Z�[�W�\��
            Call MsgOut(0, "����No������܂���", ERR_DISP)
            '�����𔲂���
            Exit Function
        '����No������Ƃ�
        Else
            '�v���C�x�[�g�ϐ��Ɋi�[
            mtrlNo = genryoNo
        End If
        
    '---- UPD [�󕶎��ϊ��O��Trim��ǉ�] 2004/10/18 TCS)R.Kawaguchi START ----
        ''�p�����[�^��NULL�̎��A�󕶎��ɕϊ�
        '���i���b�gNo
        cryno = NulltoStr(Trim(genpinNo))
        '�H���R�[�h
        PROCCD = NulltoStr(Trim(kouteiCd))
        '�H���R�[�h��5����(CB220��)�̏ꍇ�A�E4�����݂̂��擾
        If Len(Trim(PROCCD)) = 5 Then
            PROCCD = Right(PROCCD, 4)
        End If
        '�S���҃R�[�h
        staffCd = NulltoStr(Trim(tantoCd))
        '����d��
        recW = NulltoStr(Trim(ukeireW))
        '���o�d��
        sendW = NulltoStr(Trim(HARAIW))
        '���X�d��
        lossW = NulltoStr(Trim(losW))
        '�H��R�[�h
        factCd = NulltoStr(Trim(koujyoCd))
        '����H��R�[�h
        recCd = NulltoStr(Trim(ukeireCd))
        '���o�H��R�[�h
        sendCd = NulltoStr(Trim(haraiCd))
        '���ŋ敪
        disapp = NulltoStr(Trim(syoumetu))
        '�d�|�敪
        sikake = NulltoStr(Trim(shikake))
        '�V�X�e���敪�R�[�h
        sysCd = NulltoStr(Trim(sisutemu))
        '�Z�x�敪
        conceK = NulltoStr(Trim(noudoK))
        '�Z�x�l
        conceT = NulltoStr(Trim(noudoT))
        '�������M�t���O
        SENDFLG = NulltoStr(Trim(sousinFlg))
        '�����t���O
        occuFlg = NulltoStr(Trim(hasseiFlg))
        '���Z�x
        conceM = NulltoStr(Trim(motoConce))
        '�g�p�\��H��
        planFac = NulltoStr(Trim(yoteiFac))
        '�I����敪
        tanaKu = NulltoStr(Trim(tanaire))
    '---- UPD [�󕶎��ϊ��O��Trim��ǉ�] 2004/10/18 TCS)R.Kawaguchi END ----
    
    '*** UPDATE START T.TERAUCHI 2004/10/19 ���o�敪�ǉ�
        '���o�敪
        stowkkbb3 = " "
    '*** UPDATE END   T.TERAUCHI 2004/10/19
    
        '�`���[�WNo(CC200/CC300�o�^�p)�@05/08/23 ooba
        sChgNo = sChg
    
        '���t���擾
    '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''    If Not GetSysdate() Then Exit Function
    
        '�v���C�x�[�g�ϐ��Ɋi�[
        nowdate = gsSysdate
    
        ''300mm�̓��͒S���҂̌����`�F�b�N
        If Not GetAuthorityUser_300(staffCd, "C" & PROCCD, sUserName) Then
            '�����𔲂���
            Exit Function
        End If
        
        '����No���擾�ł��Ȃ�������
        If SqlCheck = False Then
            ''���b�Z�[�W�\��
            Call MsgOut(0, "����No�̓��͂Ɍ�肪����܂�", ERR_DISP)
            '�����𔲂���
            Exit Function
        End If
        
    '---- UPD 2004/10/18 TCS)R.Kawaguchi START ----
    ''    '�T�[�o�[�V�X�e�����t���������ɕύX
    ''    nowdate = GetJITUDATE(nowdate)
    ''
    ''    '��������蒼�敪�𔻒�
    ''    cyoku = GetCYOKU(nowdate)
    ''
    ''    '���ѓ�����؂���
    ''    year = Mid(nowdate, 1, 4)     '�N
    ''    month = Mid(nowdate, 6, 2)    '��
    ''    day = Mid(nowdate, 9, 2)      '��
    ''    hour = Mid(nowdate, 12, 2)    '��
    ''    min = Mid(nowdate, 15, 2)     '��
    
       '�T�[�o�[�V�X�e�����t���������ɕύX
        nowdate = GetJITUDATE(Format(nowdate, "yyyymmddhhmmss"))
        
        '��������蒼�敪�𔻒�
        cyoku = GetCYOKU(gsSysdate)
        
        '���ѓ�����؂���
        year = Mid(nowdate, 1, 4)     '�N
        month = Mid(nowdate, 5, 2)    '��
        day = Mid(nowdate, 7, 2)      '��
        hour = Mid(nowdate, 9, 2)    '��
        Min = Mid(nowdate, 11, 2)     '��
    '---- UPD 2004/10/18 TCS)R.Kawaguchi END ----
        
        '���ѓ��`�F�b�N
        If CheckDateFormat_Re(year, month, day, hour, Min) = False Then
            '���ѓ����Ó��łȂ����A���b�Z�[�W�\��
            Call MsgOut(0, "���т̎��ԃt�H�[�}�b�g���s���ł�", ERR_DISP_LOG)
            '�����𔲂���
            Exit Function
        End If
        
        '�H���R�[�h����O�H���A���H���A�d�|�H���ݒ�
        ' upd �����݌ɓ����ɂ��C��  2008/06/16 SET/miyatake ===================> START
        'Call Settei(before, after, sikakeK)
        Call Settei(before, after, sikakeK, motoSyscd)
        ' upd �����݌ɓ����ɂ��C��  2008/06/16 SET/miyatake ===================> END
        
        '�����ԍ�(XODB1)�X�V
        If Not Upd_XODB1() Then GoTo ErrHand
        
        '�����d�|�H��(XODB2)�X�V
        If Not Upd_XODB2(sikakeK) Then GoTo ErrHand
        
        '�����d�|�H��(XODB2)�ǉ�
        If Not Ins_XODB2(after, cyoku) Then GoTo ErrHand
        
        '�����H������(XODB3)�ǉ�
        If Not Ins_XODB3(before, after, cyoku, staffCd, sUserName) Then GoTo ErrHand
        
        '�߂�l�ݒ�
         Upd_XODB123 = True
        Exit Function
    
    '�G���[��
ErrHand:
        Call MsgOut(72, "", ERR_DISP_LOG)
    End Function
    
    ' @(f)
    ' �@�\      : ����No�`�F�b�N
    '
    ' �Ԃ�l    : True:����No���� False:����No�񑶍�
    '
    ' ������    : �Ȃ�
    '
    ' �@�\����  : XODB1�Ɍ���No�����݂��邩�`�F�b�N����
    '
    Private Function SqlCheck() As Boolean
        Dim sSql As String      'SQL���i�[
        Dim objOraDyn As Object
    
        '�߂�l�ݒ�
        SqlCheck = False
        
        ''SQL���쐬
        sSql = ""
        sSql = sSql & " SELECT polnob1                    " & vbLf
        sSql = sSql & " FROM   xodb1                      " & vbLf
        sSql = sSql & " WHERE  polnob1 = '" & mtrlNo & "' "
        
        'SQL�����s
        If DynSet2(objOraDyn, sSql) = False Then
            '����No���擾�ł��Ȃ��������A�����𔲂���
            Exit Function
        End If
        
        '�߂�l�ݒ�
        SqlCheck = True
    End Function
    
    
    ' @(f)
    '
    ' �@�\    : �O�H���A���H���A�d�|�H���擾
    '
    ' �Ԃ�l  : �Ȃ�
    '
    ' ������  : before :�O�H��
    '           after  :���H��
    '           sikakeK:�d�|�H��
    '           motoSyscd:�����擾��(200mm/300mm)
    '
    ' �@�\����: �H���R�[�h����O�H���A���H���A�d�|�H���擾
    '
    ' ���l    :
    '
    Private Function Settei(ByRef before As String, ByRef after As String, ByRef sikakeK As String, Optional ByVal motoSyscd As String = "")
        
    '*** UPDATE START T.TERAUCHI 2004/10/18 �g�p���Ȃ��̂ō폜
    '    '�����������
    '        If procCd = "B410" Then
    '            If disapp = "1" Then          '�p��
    '                before = "B410"
    '                after = "ZZZZ"
    '            '����d�|
    '            Else                          '�ؒf�d�|
    '                before = "C450"
    '                after = "B510"
    '            End If
    '        End If
    '
    '        '���������ؒf
    '        If procCd = "B510" Then
    '            before = "B410"
    '            If disapp = "1" Then         '����No����
    '                after = "ZZZZ"
    '            ElseIf disapp = "2" Then     '���i���b�g����
    '                after = "ZZZZ"
    '            ElseIf sikake = "1" Then     '�ؒf�d�|
    '                after = "B510"
    '            Else                         '������d�|
    '                after = "B220"
    '            End If
    '        End If
    '
    '        '���b�g�\��
    '        If procCd = "B610" Then
    '            before = "B510"
    '            sikakeK = "B220"
    '            If disapp = "1" Then         '�������b�g�����O
    '                after = "ZZZZ"
    '            Else                         '�������b�g������
    '                after = "B220"
    '            End If
    '        End If
    '*** UPDATE END T.TERAUCHI 2004/10/18
       
            '�����
            If PROCCD = "B220" Then          '��򕥏o�d�|
                before = "B510"
                after = "B225"
                sikakeK = "B220"
            End If
            
            '������d�|
            If PROCCD = "B620" Then
                If sendCd <> "" Then         '���H�ꕥ�o
                    before = "B510"
                    after = "RP00"
                    sikakeK = "B220"
                ElseIf disapp = "1" Then     '�p��
                    before = "B510"
                    after = "ZZZZ"
                    sikakeK = "B220"
                                
                '*** UPDATE START T.TERAUCHI 2004/10/18 ������ꗗ�̔p�����A���эH����B220�Ƃ���
                    PROCCD = "B220"
                '*** UPDATE END   T.TERAUCHI 2004/10/18
    
                End If
            End If
            
            '��򕥏o
            If PROCCD = "B225" Then
                before = "B220"
                sikakeK = "B225"
                
            '*** UPDATE START T.TERAUCHI 2004/10/8 ���㓊���d�|�̏ꍇ
                after = "C200"
            '*** UPDATE END T.TERAUCHI 2004/10/8
            
                If sendCd <> "" Then         '���H�ꕥ�o
                    after = "RP00"
                ElseIf disapp = "1" Then     '�p��
                    after = "ZZZZ"
                ElseIf sikake = "1" Then     '�Đ��
                    after = "B225"
                End If
            End If
            
            '�����I����(���H��)
            If PROCCD = "B230" Then
                If recCd <> "" Then
                    before = "RP00"
                    sikakeK = "B230"
                    If sendCd <> "" Then     '���H����-���㓊���d�|
                        after = "RP00"
                    Else
                        after = "C200"       '���H����-���H�ꕥ�o
                    End If
                Else
                    If sendCd <> "" Then     '���H�ꕥ�o
                        before = "B225"
                        after = "RP00"
                        sikakeK = "C200"
                    End If
                End If
            End If
            
        '*** UPDATE START T.TERAUCHI 2004/10/18 ��������H��
            If PROCCD = "B240" Then
                before = "B240"
                
                If disapp = "1" Then     '�p��
                
                    after = "ZZZZ"
                    sikakeK = "C200"
                    
                '*** UPDATE START T.TERAUCHI 2004/10/19 ���o�敪��1(���X)��ǉ�
                    stowkkbb3 = "1"
                '*** UPDATE END   T.TERAUCHI 2004/10/19
                
                Else
                    after = "C200"
                    
            '*** UPDATE START T.TERAUCHI 2004/10/20 �����敪�Ή�
                    If sysCd <> "" Then
                        stowkkbb3 = sysCd
                    End If
            '*** UPDATE END   T.TERAUCHI 2004/10/20
            
            '*** UPDATE START T.TERAUCHI 2004/10/20
                    '�V�����C���̏ꍇ
                    If sikake = "1" Then
                        sikakeK = "C200"
                    '�V�����ǉ��̏ꍇ
                    Else
                        sikakeK = ""
                    End If
            '*** UPDATE END   T.TERAUCHI 2004/10/20
                
                End If
                
            End If
        
        '*** UPDATE END   T.TERAUCHI 2004/10/18
    
    
    
            '���㓊��
        '*** UPDATE START T.TERAUCHI 2004/10/18 ���㓊���̍H���ύX
    '        If procCd = "C200" Then           '����I���d�|
    '            before = "B225"
    '            after = "C300"
    '            sikakeK = "C200"
    '        End If
            '���㓊��
        '*** UPDATE START T.TERAUCHI 2004/10/18 ���㓊���̍H���ύX
    '        If procCd = "C200" Then           '����I���d�|
    '            before = "B225"
    '            after = "C300"
    '            sikakeK = "C200"
    '        End If
    
    ''Start 2004/10/22 Upd M.Yamauchi---------------------------------------
    '        If procCd = "C250" Then           '����I���d�|
    '            before = "C200"
    '            after = "C300"
    '            sikakeK = "C200"
    '        End If
    ' upd �����݌ɓ����ɂ��C��  2008/06/16 SET/miyatake ===================> START
    '        If PROCCD = "C200" Then           '����I���d�|
    '            before = "C100"
    '            after = "C300"
    '            sikakeK = "C200"
    '        End If
            If PROCCD = "C200" Then           '����I���d�|
                Dim work As String
                If motoSyscd = SYSTEM_200 Then
'>>>>> �H���R�[�h�ύX 2008/07/25 SET.Marushita
                    before = "B690"
                    'before = "B990"
'<<<<< �H���R�[�h�ύX 2008/07/25 SET.Marushita
                    after = "C300"
                    sikakeK = "C200"
                Else
                    before = "C100"
                    after = "C300"
                    sikakeK = "C200"
                End If
            End If
    ' upd �����݌ɓ����ɂ��C��  2008/06/16 SET/miyatake ===================> END
    ''End 2004/10/22---------------------------------------------------------

        '*** UPDATE END   T.TERAUCHI 2004/10/18
            
            '����I��
            If PROCCD = "C300" Then
'                If disapp = "1" Then          '����I������         '���ĉ��@05/08/23 ooba
                    
            '*** UPDATE START T.TERAUCHI 2004/11/02
    '            '*** UPDATE START T.TERAUCHI 2004/10/18
    '            '    before = "C200"
    '                before = "C250"
    '            '*** UPDATE END   T.TERAUCHI 2004/10/18
                    
                    before = "C200"
            
            '*** UPDAT END    T.TERAUCHI 2004/11/02
                    
                    after = "ZZZZ"
                    sikakeK = "C300"
'                End If
            End If
        
    End Function
    
    
    
    
    
    ' @(f)
    '
    ' �@�\    : �����ԍ��X�V����
    '
    ' �Ԃ�l  : True:���� False:���s
    '
    ' ������  : �Ȃ�
    '
    ' �@�\����: �����ԍ�(XODB1)�X�V����
    '
    ' ���l    :
    '
    Private Function Upd_XODB1() As Boolean
        Dim sSql    As String        'SQL���i�[
        Dim iRet    As Integer       '�f�[�^�X�V��
        Dim renban  As String        '�H���A��
        Dim objOraDyn As Object
    
        '�H���R�[�h��B410�̎��͏����𔲂���
        If PROCCD = "B410" Then
            '�߂�l�ݒ�
            Upd_XODB1 = True
            Exit Function
        End If
        
        '�H���R�[�h��B450�ŏ��ŋ敪��2�̎��A�����𔲂���
        If PROCCD = "B450" And disapp = "2" Then
            '�߂�l�ݒ�
            Upd_XODB1 = True
            Exit Function
        End If
    
    '�G���[�n���h��
    On Error GoTo ErrHand
        
        '�߂�l�ݒ�
        Upd_XODB1 = False
        
        '�H���A�Ԏ擾
        sSql = ""
        sSql = sSql & " SELECT kcntb1                               " & vbLf
        sSql = sSql & " FROM   xodb1                                " & vbLf
        sSql = sSql & " WHERE  polnob1 = '" & mtrlNo & "'           " & vbLf
        
        'SQL�����s
        If DynSet2(objOraDyn, sSql) = True Then
            '�擾�����f�[�^���i�[
            renban = NulltoStr(objOraDyn.Fields("kcntb1").Value)
        End If
        
        'SQL���쐬
        sSql = ""
        sSql = sSql & "UPDATE xodb1                                 " & vbLf
        
        '�H���A�Ԃ�NULL�̎�
        If renban = "" Then
            sSql = sSql & "SET kcntb1 = 1                           " & vbLf  '�H���A��
        Else
            sSql = sSql & "SET kcntb1 = kcntb1 + 1                  " & vbLf
        End If
        
        '�V�X�e�����t�擾
    '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''    If Not GetSysdate() Then Exit Function
        
        '*** UPDATE START T.TERAUCHI 2004/10/14 ���o�H�꺰�ނ��ύX
            If sendCd <> "" Then
                sSql = sSql & "    ,toworkb1 = '" & sendCd & "'"
            End If
        '*** UPDATE END   T.TERAUCHI 2004/10/14
        
        '*** UPDATE START T.TERAUCHI 2004/10/18 �C�����t���X�V
        sSql = sSql & "        ,rdayb1 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
        '*** UPDATE END   T.TERAUCHI 2004/10/18
        
        '*** UPDATE START T.TERAUCHI 2004/11/2 ���M���t�͐ݒ肵�Ȃ�
        'sSql = sSql & "       ,sdayb1 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf '���M���t
        '*** UPDATE END   T.TERAUCHI 2004/11/2
        
        sSql = sSql & "       ,sndkb1 = ' '                        " & vbLf   '���M�敪
        sSql = sSql & "WHERE  polnob1 = '" & mtrlNo & "'           " & vbLf
    
        '���s
        iRet = SqlExec2(sSql)
        If iRet < 0 Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB1")
            Exit Function
        ElseIf iRet = 0 Then
            Call MsgOut(71, "XODB1", ERR_DISP_LOG)
            Exit Function
        End If
    
        '�߂�l�ݒ�
        Upd_XODB1 = True
        Exit Function
    
    '�G���[��
ErrHand:
        ''���b�Z�[�W�\��
        Call MsgOut(100, "", ERR_DISP_LOG, "XODB1")
    End Function
    
    
    
    ' @(f)
    '
    ' �@�\    : �����d�|�H���X�V����
    '
    ' �Ԃ�l  : True:���� False:���s
    '
    ' ������  : sikakeK:�d�|�H���R�[�h
    
    '
    ' �@�\����: �����d�|�H��(XODB2)�X�V����
    '
    ' ���l    :
    '
    Private Function Upd_XODB2(sikakeK As String) As Boolean
        Dim sSql      As String        'SQL���i�[
        Dim KUBUN     As String        '�����R�[�h
        Dim syurui    As String        '������ރR�[�h
        Dim objOraDyn As Object
        Dim iRet      As Integer       '�f�[�^�X�V��
         
    '�G���[�n���h��
    On Error GoTo ErrHand
        
        '�d�|�H���R�[�h���Ȃ����͏����𔲂���
        If sikakeK = "" Then
            '�߂�l�ݒ�
            Upd_XODB2 = True
            Exit Function
        End If
        
        '�I����敪��1�̎��͏����𔲂���
        If tanaKu = "1" Then
            '�߂�l�ݒ�
            Upd_XODB2 = True
            Exit Function
        End If
        
        '����No�A�d�|�H���R�[�h�̑��݃`�F�b�N
        'SQL���쐬
        sSql = ""
        sSql = sSql & " SELECT polnob2,                   " & vbLf
        sSql = sSql & "        wkktb2                     " & vbLf
        sSql = sSql & " FROM   xodb2                      " & vbLf
        sSql = sSql & " WHERE  polnob2 = '" & mtrlNo & "' " & vbLf
        sSql = sSql & " AND    wkktb2 = '" & sikakeK & "' " & vbLf
        
        'SQL�����s
        If DynSet2(objOraDyn, sSql) = False Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB2")
            Exit Function
        Else
            If objOraDyn.EOF = True Then
                Upd_XODB2 = True
                '����No�A�d�|�H���R�[�h�����݂��Ȃ����͏����𔲂���
                Exit Function
            End If
        End If
        
        '�߂�l�ݒ�
        Upd_XODB2 = False
    
        'XODB1����XODB2�X�V�̂��߂̃p�����[�^�擾
        'SQL���쐬
        sSql = ""
        sSql = sSql & " SELECT pokubb1,                            " & vbLf  '�����敪
        sSql = sSql & "        pokidcb1                            " & vbLf  '������ރR�[�h
        sSql = sSql & " FROM   xodb1                               " & vbLf
        sSql = sSql & " WHERE  polnob1 = '" & mtrlNo & "'          " & vbLf
        
        'SQL�����s
        If DynSet2(objOraDyn, sSql) = True Then
            '�擾�����f�[�^���i�[
            KUBUN = NulltoStr(objOraDyn.Fields("pokubb1").Value)
            syurui = NulltoStr(objOraDyn.Fields("pokidcb1").Value)
        End If
        
        '���ŋ敪��1�̎�
        If disapp = "1" Then
            'SQL���쐬
            sSql = ""
            sSql = sSql & " UPDATE xodb2                           " & vbLf
            sSql = sSql & " SET    siwb2 = 0                       " & vbLf '�d�|�d��
            sSql = sSql & "        ,sikosub2 = 0                   " & vbLf '�d�|��
            
            '�V�X�e�����t�擾
            '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''        If Not GetSysdate() Then Exit Function
            
            '*** UPDATE START T.TERAUCHI 2004/10/18 �C�����t���X�V
            sSql = sSql & "        ,rdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
            '*** UPDATE END   T.TERAUCHI 2004/10/18
            
            sSql = sSql & "        ,gndayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf
            sSql = sSql & "        ,pokubb2 = '" & KUBUN & "'      " & vbLf '�����敪
            sSql = sSql & "        ,pokidcb2 = '" & syurui & "'    " & vbLf '������ރR�[�h
            
        '*** UPDATE START T.TERAUCHI 2004/10/25
        '    sSql = sSql & "        ,sdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf
        '*** UPDATE END   T.TERAUCHI 2004/10/25
            
            sSql = sSql & "        ,sndkb2 = ' '                   " & vbLf '���M�敪
            sSql = sSql & " WHERE  polnob2 = '" & mtrlNo & "'      " & vbLf
            sSql = sSql & " AND    wkktb2 = '" & sikakeK & "'      " & vbLf
        
        '���ŋ敪��1�ȊO�̂Ƃ�
        Else
            'SQL���쐬
            sSql = ""
            sSql = sSql & " UPDATE xodb2                           " & vbLf
            sSql = sSql & " SET    siwb2 = siwb2 - " & val(recW) & vbLf     '�d�|�d��
            sSql = sSql & "        ,sikosub2 = sikosub2 - 1        " & vbLf '�d�|��
            
            '�V�X�e�����t�擾
            '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''        If Not GetSysdate() Then Exit Function
        
            '*** UPDATE START T.TERAUCHI 2004/10/18 �C�����t���X�V
            sSql = sSql & "        ,rdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
            '*** UPDATE END   T.TERAUCHI 2004/10/18
        
            sSql = sSql & "        ,gndayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf
            sSql = sSql & "        ,pokubb2 = '" & KUBUN & "'      " & vbLf '�����敪
            sSql = sSql & "        ,pokidcb2 = '" & syurui & "'    " & vbLf '������ރR�[�h
            
        '*** UPDATE START T.TERAUCHI 2004/10/25 ���M���t�̕ύX�͂Ȃ�
        '    sSql = sSql & "        ,sdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf
        '*** UPDATE END   T.TERAUCHI 2004/10/25
                
            sSql = sSql & "        ,sndkb2 = ' '                   " & vbLf '���M�敪
            sSql = sSql & " WHERE  polnob2 = '" & mtrlNo & "'      " & vbLf
            sSql = sSql & " AND    wkktb2 = '" & sikakeK & "'      " & vbLf
        End If
        
        '���s
        iRet = SqlExec2(sSql)
        If iRet < 0 Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB2")
            Exit Function
        ElseIf iRet = 0 Then
            Call MsgOut(71, "XODB2", ERR_DISP_LOG)
            Exit Function
        End If
    
        '�߂�l�ݒ�
        Upd_XODB2 = True
        Exit Function
    
    '�G���[��
ErrHand:
        ''�װ
        Call MsgOut(100, "", ERR_DISP_LOG, "XODB2")
    End Function
    
    
    
    ' @(f)
    '
    ' �@�\    : �����d�|�H���ǉ�����
    '
    ' �Ԃ�l  : True:���� False:���s
    '
    ' ������  : after:���H���R�[�h
    '           cyoku:���敪
    '
    ' �@�\����: �����d�|�H��(XODB2)�ǉ�����
    '
    ' ���l    :
    '
    Private Function Ins_XODB2(after As String, cyoku As String) As Boolean
        Dim sSql      As String       'SQL���i�[
        Dim KUBUN     As String       '�����敪
        Dim syurui    As String       '������ރR�[�h
        Dim objOraDyn As Object
        Dim iRet    As Integer        '�f�[�^�ǉ���
        
        '���H���R�[�h��RP00�AZZZZ�AB510�̎�
        If after = "RP00" Or after = "ZZZZ" Or after = "B510" Then
            Ins_XODB2 = True
            '�����𔲂���
            Exit Function
        End If
        
        '�G���[�n���h��
        On Error GoTo ErrHand
    
        '�߂�l�ݒ�
        Ins_XODB2 = False
        
        'XODB1����XODB2�X�V�̂��߂̃p�����[�^�擾
        'SQL���쐬
        sSql = ""
        sSql = sSql & " SELECT pokubb1,                        " & vbLf  '�����敪
        sSql = sSql & "        pokidcb1                        " & vbLf  '������ރR�[�h
        sSql = sSql & " FROM   xodb1                           " & vbLf
        sSql = sSql & " WHERE  polnob1 = '" & mtrlNo & "'      " & vbLf
        
        'SQL�����s
        If DynSet2(objOraDyn, sSql) = True Then
            '�擾�����f�[�^���i�[
            KUBUN = NulltoStr(objOraDyn.Fields("pokubb1").Value)
            syurui = NulltoStr(objOraDyn.Fields("pokidcb1").Value)
        End If
        
    
        '����No�A���H���R�[�h�̑��݃`�F�b�N
        sSql = ""
        sSql = sSql & " SELECT polnob2,                            " & vbLf
        sSql = sSql & "        wkktb2                              " & vbLf
        sSql = sSql & " FROM   xodb2                               " & vbLf
        sSql = sSql & " WHERE  POLNOB2 = '" & mtrlNo & "'          " & vbLf
        sSql = sSql & " AND    WKKTB2 = '" & after & "'            " & vbLf
    
        'SQL�����s
        If DynSet2(objOraDyn, sSql) = False Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB2")
            Exit Function
        End If
        
        '�Y���f�[�^������ꍇ
        If objOraDyn.EOF = False Then
            
            'SQL���쐬
            sSql = ""
            sSql = sSql & " UPDATE xodb2                       " & vbLf
     '*** UPDATE START T.TERAUCHI 2004/10/19 ����d�ʂł͂Ȃ��A���o�d�ʂ�ݒ�
        '    sSql = sSql & " SET    siwb2 = siwb2 + " & val(recW) & vbLf '�d�|�d��
             sSql = sSql & " SET    siwb2 = siwb2 + " & val(sendW) & vbLf '�d�|�d��
     '*** UPDATE END   T.TERAUCHI 2004/10/19
            sSql = sSql & "        ,sikosub2 = sikosub2 + 1    " & vbLf '�d�|��
            
            '�V�X�e�����t�擾
            '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''        If Not GetSysdate() Then Exit Function
        
        
            '*** UPDATE START T.TERAUCHI 2004/10/18 �C�����t���X�V
            sSql = sSql & "        ,rdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
            '*** UPDATE END   T.TERAUCHI 2004/10/18
        
            sSql = sSql & "        ,gndayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
            sSql = sSql & "        ,pokubb2 = '" & KUBUN & "'      " & vbLf '�����敪
            sSql = sSql & "        ,pokidcb2 = '" & syurui & "'    " & vbLf '������ރR�[�h
        
        '*** UPDATE START T.TERAUCHI 2004/10/25 �C�����t�̍X�V�͂Ȃ�
        '    sSql = sSql & "        ,sdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
        '*** UPDATE END   T.TERAUCHI 2004/10/25
            
            sSql = sSql & "        ,sndkb2 = ' '                   " & vbLf '���M�敪
            sSql = sSql & " WHERE  POLNOB2 = '" & mtrlNo & "'      " & vbLf
            sSql = sSql & " AND    WKKTB2 = '" & after & "'        " & vbLf
    
            
        '�Y���f�[�^���Ȃ������ꍇ
        Else
            sSql = ""
            sSql = sSql & "INSERT INTO XODB2                       " & vbLf
            sSql = sSql & "            (polnob2,                   " & vbLf
            sSql = sSql & "            wkktb2,                     " & vbLf
            sSql = sSql & "            PLACB2,                     " & vbLf
    
    '*** UPDATE START T.TERAUCHI 2004/10/18 �o�^���t���ǉ�
            sSql = sSql & "            TDAYB2,                     " & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/10/18
    
            sSql = sSql & "            RDAYB2,                     " & vbLf
            sSql = sSql & "            SDAYB2,                     " & vbLf
            sSql = sSql & "            SNDKB2,                     " & vbLf
            sSql = sSql & "            SAKJB2,                     " & vbLf
            sSql = sSql & "            POKUBB2,                    " & vbLf
            sSql = sSql & "            POKIDCB2,                   " & vbLf
            sSql = sSql & "            SIWB2,                      " & vbLf
            sSql = sSql & "            SIKOSUB2,                   " & vbLf
            sSql = sSql & "            GNDAYB2,                    " & vbLf
            sSql = sSql & "            GNCYOKB2)                   " & vbLf
            sSql = sSql & " VALUES     ('" & mtrlNo & "',          " & vbLf '�����ԍ�
            
            If after = "" Then
                sSql = sSql & "        ' ',                        " & vbLf '�H���R�[�h
            Else
                sSql = sSql & "        '" & after & "',            " & vbLf '�H���R�[�h
            End If
            
            sSql = sSql & "            ' ',                        " & vbLf '���C���R�[�h
    '*** UPDATE START T.TERAUCHI 2004/10/18 �o�^���t�A�C�����t�ɒl�ݒ�
        '    sSql = sSql & "            null,                       " & vbLf '�C�����t
            sSql = sSql & "            to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')," & vbLf '�o�^���t"
            
        '*** UPDATE START T.TERAUCHI 2004/10/25 �C�����t�̓o�^�͂Ȃ�
        '    sSql = sSql & "            to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')," & vbLf '�C�����t"
            sSql = sSql & "            null,                       " & vbLf '�C�����t
        '*** UPDATE END   T.TERAUCHI 2004/10/25
    
    '*** UPDATE END   T.TERAUCHI 2004/10/18
    
            '�V�X�e�����t�擾
            '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''        If Not GetSysdate() Then Exit Function
        
        '*** UPDATE START T.TERAUCHI 2004/10/25 ���M���t�̓o�^�͂Ȃ�
        '    sSql = sSql & "            to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')," & vbLf '���M���t"
            sSql = sSql & "            null,                       " & vbLf '���M���t
        '*** UPDATE END   T.TERAUCHI 2004/10/25
            
            sSql = sSql & "            ' ',                        " & vbLf          '���M�敪
            sSql = sSql & "            '0',                        " & vbLf          '�폜�敪
            sSql = sSql & "            '" & KUBUN & "',            " & vbLf          '�����敪
            sSql = sSql & "            '" & syurui & "',           " & vbLf          '������ރR�[�h
            sSql = sSql & "            " & val(sendW) & ",         " & vbLf          '�d�|�d��
            sSql = sSql & "            1,                          " & vbLf          '�d�|��
            sSql = sSql & "            to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')," & vbLf '�ŐV�d�|���t
            sSql = sSql & "            '" & cyoku & "')            " & vbLf          '���敪
        End If
        
        '���s
        iRet = SqlExec2(sSql)
        If iRet < 0 Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB2")
            Exit Function
        End If
    
        '�߂�l�ݒ�
        Ins_XODB2 = True
        Exit Function
    
    '�G���[��
ErrHand:
        ''�װ
        Call MsgOut(100, "", ERR_DISP_LOG, "XODB2")
    End Function
    
    ' @(f)
    '
    ' �@�\    : �����H�����ђǉ�����
    '
    ' �Ԃ�l  : True:���� False:���s
    '
    ' ������  : before  :�O�H���R�[�h
    '          after    :���H���R�[�h
    '          cyoku    :���敪
    '          staffCd  :�S���҃R�[�h
    '          sUserName:�S���Җ�
    '
    ' �@�\����: �����H��(XODB3)�ǉ�����
    '
    ' ���l    :
    '
    Private Function Ins_XODB3(before As String, _
                                after As String, _
                                cyoku As String, _
                                staffCd As String, _
                                sUserName As String) As Boolean
        Dim sSql       As String         'SQL���i�[
        Dim iRet       As Integer        '�f�[�^�ǉ���
        Dim renban     As String         'XODB1�̍H���A��
        Dim KUBUN      As String         '       �����敪
        Dim syurui     As String         '       ������ރR�[�h
        Dim objOraDyn  As Object
        Dim year       As String         '�V�X�e�����t(�N)
        Dim month      As String         '           (��)
        Dim day        As String         '           (��)
        Dim hour       As String         '           (��)
        Dim Min        As String         '           (��)
        Dim sNowdate   As String         'SUMCO���ԑΉ��@05/08/23 ooba
    
    '�G���[�n���h��
    On Error GoTo ErrHand
    
        '�߂�l�ݒ�
        Ins_XODB3 = False
    
        'SUMCO���ԑΉ��@05/08/23 ooba START =======================================>
        sNowdate = gsSysdate
        '�T�[�o�[�V�X�e�����t�����ѓ��ɕύX
        sNowdate = GetJITUDATE(Format(sNowdate, "yyyymmddhhmmss"))
        '���ѓ�����؂���
        year = Mid(sNowdate, 1, 4)     '�N
        month = Mid(sNowdate, 5, 2)    '��
        day = Mid(sNowdate, 7, 2)      '��
        hour = Mid(sNowdate, 9, 2)     '��
        Min = Mid(sNowdate, 11, 2)     '��
        'SUMCO���ԑΉ��@05/08/23 ooba END =========================================>
        
        'XODB1����f�[�^�擾
        'SQL���쐬
        sSql = ""
        sSql = sSql & " SELECT kcntb1,                       " & vbLf   '�H���A��
        sSql = sSql & "        pokubb1,                      " & vbLf   '�����敪
        sSql = sSql & "        pokidcb1                      " & vbLf   '������ރR�[�h
        sSql = sSql & " FROM   xodb1                         " & vbLf
        sSql = sSql & " WHERE  polnob1 = '" & mtrlNo & "'    " & vbLf
    
        'SQL�����s
        If DynSet2(objOraDyn, sSql) = True Then
            '�擾�����f�[�^���i�[
            renban = NulltoStr(objOraDyn.Fields("kcntb1").Value)
            KUBUN = NulltoStr(objOraDyn.Fields("pokubb1").Value)
            syurui = NulltoStr(objOraDyn.Fields("pokidcb1").Value)
        End If
    
        '�����H������(XODB3)�X�V
        sSql = ""
        sSql = sSql & "insert into XODB3                        " & vbLf
        sSql = sSql & "            (POLNOB3,                    " & vbLf   '�����ԍ�
        sSql = sSql & "            KCNTB3,                      " & vbLf   '�H���A��
        sSql = sSql & "            CRSEQB3,                     " & vbLf   '�����A��
        sSql = sSql & "            TDAYB3,                      " & vbLf   '�o�^���t
        sSql = sSql & "            RDAYB3,                      " & vbLf   '�C�����t
        sSql = sSql & "            SDAYB3,                      " & vbLf   '���M���t
        sSql = sSql & "            SNDKB3,                      " & vbLf   '���M�敪
        sSql = sSql & "            SAKJB3,                      " & vbLf   '�폜�敪
        sSql = sSql & "            POKUBB3,                     " & vbLf   '�����敪
        sSql = sSql & "            POKIDCB3,                    " & vbLf   '������ރR�[�h
        sSql = sSql & "            POLTNB3,                     " & vbLf   '�������b�gNo
        sSql = sSql & "            MODKBB3,                     " & vbLf   '�ԍ��敪
        sSql = sSql & "            SUMKBB3,                     " & vbLf   '�W�v�敪
        sSql = sSql & "            WKKTB3,                      " & vbLf   '�H���R�[�h
        sSql = sSql & "            PLACB3,                      " & vbLf   '���C���R�[�h
        sSql = sSql & "            FRWB3,                       " & vbLf   '����d��
        sSql = sSql & "            TOWB3,                       " & vbLf   '���o�d��
        sSql = sSql & "            LOSWB3,                      " & vbLf   '���X�d��
        sSql = sSql & "            FRWKKTB3,                    " & vbLf   '����d�ʃR�[�h
        sSql = sSql & "            TOWKKTB3,                    " & vbLf   '���o�H���R�[�h
        sSql = sSql & "            TOWKKBB3,                    " & vbLf   '���o�H���敪
        sSql = sSql & "            TOWORKB3,                    " & vbLf   '���o�H��R�[�h
        sSql = sSql & "            TOPLACB3,                    " & vbLf   '���o���C���R�[�h
        sSql = sSql & "            CHGNB3,                      " & vbLf   '�`���[�WNo
        sSql = sSql & "            EYYB3,                       " & vbLf   '���ѓ��t(�N)
        sSql = sSql & "            EMMB3,                       " & vbLf   '���ѓ��t(��)
        sSql = sSql & "            EDDB3,                       " & vbLf   '���ѓ��t(��)
        sSql = sSql & "            ECYOKB3,                     " & vbLf   '���敪
        sSql = sSql & "            EHHB3,                       " & vbLf   '���ю���(��)
        sSql = sSql & "            EMIB3,                       " & vbLf   '���ю���(��)
        sSql = sSql & "            MANB3,                       " & vbLf   '�S����
        sSql = sSql & "            MANJB3,                      " & vbLf   '�S���Җ�
        sSql = sSql & "            DENKB3,                      " & vbLf   '�Z�x�敪
        sSql = sSql & "            DENSITYB3,                   " & vbLf   '�Z�x�l
        sSql = sSql & "            GSNDFLGB3,                   " & vbLf   '�������M�t���O
        sSql = sSql & "            HFLGB3,                      " & vbLf   '�����t���O
        sSql = sSql & "            mdensityb3,                  " & vbLf   '���Z�x
        sSql = sSql & "            plworkb3,                    " & vbLf   '�g�p�\��H��
        sSql = sSql & "            htkbnb3)                     " & vbLf   '�p��/�A���敪
        sSql = sSql & "VALUES      ('" & mtrlNo & "',           " & vbLf   '�����ԍ�
        
        '�H���A�Ԃ���̎�
        If renban = "" Then
            sSql = sSql & "        0,                           " & vbLf
        Else
            sSql = sSql & "        " & val(renban) & ",         " & vbLf   '�H���A��
        End If
        
        sSql = sSql & "            1,                           " & vbLf   '�����A��
        
        '�V�X�e�����t�擾
        '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''    If Not GetSysdate() Then Exit Function
        
        sSql = sSql & " to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss'), " & vbLf   '�o�^���t
    '*** UPDATE START T.TERAUCHI 2004/10/18
    '    sSql = sSql & "            null,                        " & vbLf   '�C�����t
    
        '*** UPDATE START T.TERAUCHI 2004/10/25 �C�����t�̓o�^�͂Ȃ�
    '    sSql = sSql & " to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss'), " & vbLf   '�C�����t
        sSql = sSql & "            null,                        " & vbLf   '�C�����t
        '*** UPDATE END   T.TERAUCHI 2004/10/25
    
    '*** UPDATE END   T.TERAUCHI 2004/10/18
    
        '*** UPDATE START T.TERAUCHI 2004/10/25 ���M���t�̓o�^�͂Ȃ�
    '    sSql = sSql & " to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss'), " & vbLf   '���M���t
        sSql = sSql & "            null,                        " & vbLf   '�C�����t
        '*** UPDATE END   T.TERAUCHI 2004/10/25
        
        sSql = sSql & "            ' ',                         " & vbLf   '���M�敪
        sSql = sSql & "            '0',                         " & vbLf   '�폜�敪
        sSql = sSql & "            '" & KUBUN & "',             " & vbLf   '�����敪
        sSql = sSql & "            '" & syurui & "',            " & vbLf   '������ރR�[�h
        sSql = sSql & "            ' ',                         " & vbLf   '�������b�g�ԍ�
        sSql = sSql & "            ' ',                         " & vbLf   '�ԍ��敪
        sSql = sSql & "            ' ',                         " & vbLf   '�W�v�敪
        sSql = sSql & "            '" & PROCCD & "',            " & vbLf   '�H���R�[�h
        sSql = sSql & "            '�@',                        " & vbLf   '���C���R�[�h
        
    ''Upd start t.terauchi 2005/04/15   B240�̎��A��������ƁA�݌ɏC���ŏ����𕪂���
        ''B240�ȊO�͂���܂Œʂ�
        If PROCCD <> "B240" Then
            sSql = sSql & "            " & val(recW) & ",           " & vbLf   '����d��
        ''B240�̎�
        Else
            ''�݌ɏC��(�p�p)�̎�
            If disapp = "1" Then
                sSql = sSql & "            0,          " & vbLf   '����d��
            ''�݌ɏC���̎�
            ElseIf lossW <> "0" Then
                sSql = sSql & "            0,           " & vbLf   '����d��
            ''��������̎�(����ݒ�)
            Else
                sSql = sSql & "            " & val(recW) & ",           " & vbLf   '����d��
            End If
        End If
    '*** UPDATE START T.TERAUCHI 2004/10/21 �p���̍ۂ̓��X�d�ʂɉ��Z����
    '    sSql = sSql & "            " & val(sendW) & ",          " & vbLf   '���o�d��
    '    sSql = sSql & "            " & val(lossW) & ",          " & vbLf   '���X�d��
        
        '�p���̂Ƃ�
'        2004/12/17 TCS NAKAJIMA update-start
         If disapp = "1" And PROCCD <> "C300" Then
'        If disapp = "1" Then
'        2004/12/17 TCS NAKAJIMA update-end
            sSql = sSql & "            0,                           " & vbLf   '���o�d��
            sSql = sSql & "            " & val(sendW) & ",          " & vbLf   '���X�d��
        Else
            
        ''Upd start t.terauchi 2005/04/15
            ''B240�ȊO�͂���܂Œʂ�
            If PROCCD <> "B240" Then
                sSql = sSql & "            " & val(sendW) & ",          " & vbLf   '���o�d��
            ''B240�̎�
            Else
                ''�݌ɏC���̎�
                If lossW <> "0" Then
                    sSql = sSql & "            0,          " & vbLf   '���o�d��
                ''��������̎�(����ݒ�)
                Else
                    sSql = sSql & "            " & val(sendW) & ",          " & vbLf   '���o�d��
                End If
            End If
            
            sSql = sSql & "            " & val(lossW) & ",          " & vbLf   '���X�d��
        End If
    '*** UPDATE END   T.TERAUCHI 2004/10/21
        
        ' upd �����݌ɓ����ɂ��C��  2008/06/17 SET/miyatake ===================> START
'        sSQL = sSQL & "            '" & before & "',            " & vbLf   '����H���R�[�h
'        sSQL = sSQL & "            '" & after & "',             " & vbLf   '���o�H���R�[�h
        If PROCCD = "C200" And recW < 0 Then
            '�}�C�i�X����(�ԋp��)�͎���ƕ��o���t�ɐݒ肷��
            sSql = sSql & "            '" & after & "',             " & vbLf   '����H���R�[�h
            sSql = sSql & "            '" & before & "',            " & vbLf   '���o�H���R�[�h
        Else
            sSql = sSql & "            '" & before & "',            " & vbLf   '����H���R�[�h
            sSql = sSql & "            '" & after & "',             " & vbLf   '���o�H���R�[�h
        End If
        ' upd �����݌ɓ����ɂ��C��  2008/06/17 SET/miyatake ===================> END
        
    '*** UPDATE START T.TERAUCHI 2004/10/19
    '    sSql = sSql & "            ' ',                         " & vbLf   '���o�敪
        sSql = sSql & "            '" & stowkkbb3 & "',          " & vbLf   '���o�敪
    '*** UPDATE END   T.TERAUCHI 2004/10/19
    
    '*** UPDATE START T.TERAUCHI 2004/10/18 ���o�H�꺰�ނ��ݒ肳��Ă��Ȃ��ꍇ�͎��H���ݒ肷��
        If sendCd <> "" Then
            sSql = sSql & "            '" & sendCd & "',            " & vbLf   '���o�H��R�[�h
        Else
            sSql = sSql & "            '" & factCd & "',            " & vbLf   '���o�H��R�[�h
        End If
    '*** UPDATE END   T.TERAUCHI 2004/10/18
        
        sSql = sSql & "            ' ',                         " & vbLf   '���o���C���R�[�h
        
'        sSql = sSql & "            ' ',                         " & vbLf   '�`���[�WNo
        '�`���[�WNo�Z�b�g�@05/08/23 ooba START ==============================================>
        If PROCCD <> "C300" And PROCCD <> "C200" Then
            sSql = sSql & "            ' ',                         " & vbLf   '�`���[�WNo
        Else
            sSql = sSql & "            '" & sChgNo & "',            " & vbLf   '�`���[�WNo
        End If
        '�`���[�WNo�Z�b�g�@05/08/23 ooba END ================================================>

        '�V�X�e�����t����N��؂���
'        year = Left(gsSysdate, 4)       '���ĉ��@05/08/23 ooba
        sSql = sSql & "            '" & year & "',              " & vbLf   '���ѓ��t(�N)
        '�V�X�e�����t���猎��؂���
'        month = Mid(gsSysdate, 6, 2)    '���ĉ��@05/08/23 ooba
        sSql = sSql & "            '" & month & "',             " & vbLf   '���ѓ��t(��)
        '�V�X�e�����t�������؂���
'        day = Mid(gsSysdate, 9, 2)      '���ĉ��@05/08/23 ooba
        sSql = sSql & "            '" & day & "',               " & vbLf   '���ѓ��t(��)
        sSql = sSql & "            '" & cyoku & "',             " & vbLf   '���敪
        '�V�X�e�����t���玞��؂���
'        hour = Mid(gsSysdate, 12, 2)    '���ĉ��@05/08/23 ooba
        sSql = sSql & "            '" & hour & "',              " & vbLf   '���ю���(��)
        '�V�X�e�����t���番��؂���
'        min = Mid(gsSysdate, 15, 2)     '���ĉ��@05/08/23 ooba
        sSql = sSql & "            '" & Min & "',               " & vbLf   '���ю���(��)
        sSql = sSql & "            '" & Right(staffCd, 7) & "',  " & vbLf  '�S����
        sSql = sSql & "            '" & sUserName & "',         " & vbLf   '�S���Җ�
        sSql = sSql & "            '" & conceK & "',            " & vbLf   '�Z�x�敪
        sSql = sSql & "            " & val(conceT) & ",         " & vbLf   '�Z�x�l
        sSql = sSql & "            '" & SENDFLG & "',           " & vbLf   '�������M�t���O
        sSql = sSql & "            '" & occuFlg & "',           " & vbLf   '�����t���O
        sSql = sSql & "            " & val(conceM) & ",         " & vbLf   '���Z�x
        sSql = sSql & "            '" & planFac & "',           " & vbLf   '�g�p�\��H��
        
        '���ŋ敪1�̎�
        If disapp = "1" Then
            sSql = sSql & "        '2')                         " & vbLf   '�p��/�A���敪
        '���ŋ敪2�̎�
        ElseIf disapp = "2" Then
            sSql = sSql & "        '9')                         " & vbLf
        '��L�ȊO�̎�
        Else
            sSql = sSql & "        '1')                         " & vbLf
        End If
    
        '���s
        iRet = SqlExec2(sSql)
        If iRet < 0 Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB3")
            Exit Function
        ElseIf iRet = 0 Then
            Call MsgOut(71, "XODB3", ERR_DISP_LOG)
            Exit Function
        End If
    
        '�߂�l�ݒ�
        Ins_XODB3 = True
        Exit Function
    
    '�G���[��
ErrHand:
        ''�װ
        Call MsgOut(100, "", ERR_DISP_LOG, "XODB3")
    End Function
    
    ' @(f)
    '
    ' �@�\    : ���ѓ��`�F�b�N
    '
    ' �Ԃ�l  : True:���� False:���s
    '
    ' ������  : before  :�O�H���R�[�h
    '          after    :���H���R�[�h
    '          cyoku    :���敪
    '          staffCd  :�S���҃R�[�h
    '          sUserName:�S���Җ�
    '
    ' �@�\����: ���ѓ������t�Ƃ��Đ��������`�F�b�N����
    '
    ' ���l    :
    '
    Public Function CheckDateFormat_Re(sYear As String, _
                                       sMonth As String, _
                                       sDay As String, _
                                       sHour As String, _
                                       sMinute As String) As Boolean
                                        
        ''�N�`�F�b�N2000�N������������΃G���[
        If sYear < "2000" Or Len(sYear) <> 4 Then
            CheckDateFormat_Re = False
            Exit Function
        End If
        
        ''���`�F�b�N�@12�����傫����΃G���[
        If sMonth > "12" Or Len(sMonth) <> 2 Then
            CheckDateFormat_Re = False
            Exit Function
        End If
        
        ''���`�F�b�N�@31�����傫����΃G���[
        If sDay > "31" Or Len(sDay) <> 2 Then
            CheckDateFormat_Re = False
            Exit Function
        End If
        
        ''���`�F�b�N�@23�����傫����΃G���[
        If sHour > "23" Or Len(sHour) <> 2 Then
            CheckDateFormat_Re = False
            Exit Function
        End If
        
        ''���`�F�b�N�@59�����傫����΃G���[
        If sMinute > "59" Or Len(sMinute) <> 2 Then
            CheckDateFormat_Re = False
            Exit Function
        End If
             
        '�߂�l�ݒ�
        CheckDateFormat_Re = True
    End Function
    
    ' @(f)
    ' �@�\      : ���ѓ��t�쐬
    '
    ' �Ԃ�l    : ���ѓ��t(YYYYMMDDhhmmss�^)
    '
    ' ������    : nowdate  -  ���t�f�[�^(YYYYMMDDhhmmss�^)
    '
    ' �@�\����  : ���ѓ��t���쐬���Ăяo�����ɕԂ�
    '
    '---- UPD [mdlSP_Re(200mm�p)�̃\�[�X���f] 2004/10/18 TCS)R.Kawaguchi START ----
    ''Private Function GetJITUDATE(nowdate As String) As String
    ''    Dim jitudate As String     '�ϊ��p�̓��t���i�[����ϐ�(YYYY/MM/DD�^)
    ''    Dim jitutime As String     '����p�̎������i�[����ϐ�
    ''
    ''    '����p�ɓ��t�A������؂�o��
    ''    jitudate = Left(nowdate, 10)
    ''    jitutime = Mid(nowdate, 11, 6)
    ''
    ''    '0:00����7:59�̊Ԃ̏ꍇ�͓��t��O���ɐݒ�
    ''    If jitutime >= "000000" And jitutime < "080000" Then
    ''        jitudate = Format(DateAdd("d", -1, jitudate), "YYYYMMDD")
    ''        jitudate = Replace(jitudate, "/", "")
    ''        GetJITUDATE = jitudate & jitutime
    ''    Else
    ''        GetJITUDATE = nowdate
    ''    End If
    ''End Function
    Public Function GetJITUDATE(ByVal systemdate As String) As String
    
        Dim jitudate As String     '�ϊ��p�̓��t���i�[����ϐ�(YYYY/MM/DD�^)
        Dim jitutime As String     '����p�̎������i�[����ϐ�
    
        '����p�ɓ��t�A������؂�o��
        jitudate = left(systemdate, 4) & "/" & Mid(systemdate, 5, 2) & "/" & Mid(systemdate, 7, 2)
        jitutime = Mid(systemdate, 9, 6)
    ''    jitudate = Format(systemdate, "YYYYMMDD")
    ''    jitutime = Format(systemdate, "hhmmss")
        '0:00����7:59�̊Ԃ̏ꍇ�͓��t��O���ɐݒ�
        If jitutime >= "000000" And jitutime < "080000" Then
    ''        jitudate = jitudate - 1 'DateAdd("d", -1, jitudate)
    '        jitudate = DateAdd("d", -1, jitudate)
            jitudate = Format(DateAdd("d", -1, jitudate), "YYYYMMDD")   ''�C��(2004/01/14)
            jitudate = Replace(jitudate, "/", "")
            GetJITUDATE = jitudate & jitutime
        Else
            GetJITUDATE = systemdate
        End If
    
    End Function
    '---- UPD [mdlSP_Re(200mm�p)�̃\�[�X���f] 2004/10/18 TCS)R.Kawaguchi END ----
    
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' �@�\      :   NULL�����ϊ�
    '
    ' �Ԃ�l    :�@ �ϊ�����ް�
    '
    ' ����      :�@ val - �ϊ������ް�
    '
    ' �@�\����  :   NULL�����ϊ�
    '
    ' ���l      :
    '
    '///////////////////////////////////////////////////
    Public Function NulltoStr(val As Variant) As String
        
        If IsNull(val) Then
            NulltoStr = ""
        Else
            NulltoStr = val
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
    
    ' @(f)
    ' �@�\      : SQL���l�ϊ��֐�
    '
    ' �Ԃ�l    : <���͐��l> or NULL
    '
    ' ������    : �ϊ��Ώې��l
    '
    ' �@�\����  : �n���ꂽ���l��NULL�ł����"NULL"�������łȂ���΂��̂܂܏o�͂���
    Private Function Cnv2Number(vinput) As String
        If IsNull(vinput) Or vinput = "NULL" Then
            vinput = ""
        End If
        
        If vinput = "" Then
            Cnv2Number = "NULL"
        Else
            Cnv2Number = vinput
        End If
    End Function
    '2004/9/17tcs Suenaga �ǉ� end-------------------------------------
    
    '2004/10/15tcs Yamauchi �ǉ� start-------------------------------------
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' �@�\    : �l�̌ܓ��A�؏�A�؎̂��s��
    '
    ' �Ԃ�l  : �l�̌ܓ������l
    '
    ' ������  : vValue          - �̓������ް�
    '           lExp            - �̓����錅�ʒu���P�O�̏搔�œn��
    '                             ��F�O�D�O�O�P���ڂ��l�̌ܓ�����Ȃ�u�|�R�v
    '           iIs             - �̓����鐔
    '                             �؏�̏ꍇ�͂O�A�؎̂̏ꍇ�͂P�O
    '
    ' �@�\����: �l�̌ܓ����s��
    '
    ' ���l    :
    '
    '///////////////////////////////////////////////////
    Public Function Round_Re(vValue As Variant, ByVal lExp As Long, Optional ByVal iIs As Integer = 5) As Double
    
        Dim lPeriod     As Long
        Dim vValueTemp  As Variant
      
        '�f�[�^������
        Round_Re = 0
        vValue = CDec(vValue)
        
        ' �����̔���
        If Not IsNumeric(vValue) Then Exit Function
            
        vValueTemp = vValue
        
        ' �s���I�h�̈ʒu���擾
        lPeriod = InStr(vValueTemp, Chr(46))
        If vValue < 0 Then
            lPeriod = lPeriod - 1
            vValueTemp = Mid(vValueTemp, 2)
        End If
    
        '�����_���Ȃ��A�����ȉ��l�̌ܓ��̎��͖���
        If lPeriod <= 0 And lExp < 0 Then
            Round_Re = CDbl(vValueTemp)
            Exit Function
        End If
        
        '�����_������A��������lExp���Z���Ƃ�����
        If lPeriod > 0 And lExp * -1 > Len(vValueTemp) - lPeriod Then
            Round_Re = CDbl(vValueTemp)
            Exit Function
        End If
        
        'lExp�������_�ȏ�̌�����菬�����ꍇ
        If lExp + 2 < lPeriod Then
            If lExp < 0 Then         '�̓��ʒu�������_�ȉ��̏ꍇ
                '�̓��ʒu�ȏ���擾
                Round_Re = CDbl(left(vValueTemp, lPeriod - (lExp + 1)))
                '�̓�
                If CInt(Mid(vValueTemp, lPeriod - lExp, 1)) >= iIs Then Round_Re = Round_Re + 10 ^ (lExp + 1)
            Else                         '�����_�ȏ�̏ꍇ
                '�̓��ʒu�ȏ���擾
                Round_Re = CDbl(left(vValueTemp, lPeriod - (lExp + 2))) * (10 ^ (lExp + 1))
                '�̓�
                If CInt(Mid(vValueTemp, lPeriod - (lExp + 1), 1)) >= iIs Then Round_Re = Round_Re + 10 ^ (lExp + 1)
            End If
        'lExp���ŏ�ʌ��̂Ƃ�
        ElseIf lExp + 2 = lPeriod Then
            '�̓�
            If CInt(left(vValueTemp, 1)) >= iIs Then
                Round_Re = 10 ^ (lExp + 1)
            Else
                Round_Re = 0
            End If
        '����ȏ�̎�
        Else
            Exit Function
        End If
        
        If vValue < 0 Then
            Round_Re = Round_Re * (-1)
        End If
    
    End Function
    
    '
    ' @(f)
    ' �@�\    : ���H��擾
    ' �Ԃ�l  : �Ȃ�
    ' ������  : sFactryCd       - �H�꺰��
    '           sSelfFactory    - ���H��
    ' �@�\����: �H��R�[�h���A���H����擾����
    '
    Public Sub GetSelfFactory(ByVal sFactryCd As String, sSelfFactory As String)
    
        Select Case sFactryCd
            Case "10"               ''��c�H��
                sSelfFactory = "I"
            Case "30"               ''����H��
                sSelfFactory = "I"
            Case "40"               ''�đ�H��
                sSelfFactory = "Y"
            Case "42"               ''�R�O�O����
                sSelfFactory = "Z"
            Case "43"               ''�R�O�O����
                sSelfFactory = "Z"
            Case "90"               ''�e�X�g��
                sSelfFactory = "Y"
            Case "91"               ''�e�X�g��(�đ�) 2007/04/05�ǉ� SETsw kubota
                sSelfFactory = "Y"
            Case "92"               ''�e�X�g��(����)
                sSelfFactory = "I"
            Case "93"               ''�e�X�g��(����A1) 2010/04/14 SETsw kubota
                sSelfFactory = "I"
            Case Else               ''�O��
                sSelfFactory = "I"
        End Select
    
    End Sub
    
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' �@�\    : 300�p�@XODCX�o�^���A��{���擾SQL
    '
    ' �Ԃ�l  : �Ȃ�
    '
    ' ������  :�@ssql       SQL�i�[�p�ϐ�
    '            sCrystalNo �����ԍ�
    '
    ' �@�\����:
    '
    ' ���l    :
    '
    '///////////////////////////////////////////////////
    Public Sub GetAssistSQL_300(sSql As String, sCrystalNo As String)
    
        sSql = ""
        sSql = sSql & "select                                                   " & vbLf
        sSql = sSql & "         T1.ADDOPPC1                                     " & vbLf    ''�ǉ��h�[�v�����ʒu
        sSql = sSql & "         ,T1.PUTCUTWC1                                   " & vbLf    ''�g�b�vWT
        sSql = sSql & "         ,T1.DTYPEC1�@�@�@�@                              " & vbLf    ''�h�[�v�^�C�v
        sSql = sSql & "         ,(T1.DIA1C1 + T1.DIA2C1 + T1.DIA3C1)            " & vbLf
        sSql = sSql & "         / 3 AS UPDMCX                                   " & vbLf    ''���グAV�a
        sSql = sSql & "         ,T2.HINBAN || TRIM(TO_CHAR(T2.NMNOREVNO,'00'))  " & vbLf
        sSql = sSql & "         || T2.NFACTORY || T2.NOPECOND AS HINBCX         " & vbLf    ''�i��
        
    '*** UPDATE START T.TERAUCHI 2004/11/27 �����h�[�v�A�����h�[�v�ʕύX�Ή�
    '    sSql = sSql & "         ,T2.DPNTCLS                                     " & vbLf    ''�����h�[�v
    '    sSql = sSql & "         ,T2.DOPANT                                      " & vbLf    ''�����h�[�v��
        sSql = sSql & "         ,T6.CRYDOP DPNTCLS                               " & vbLf    ''�����h�[�v
        sSql = sSql & "         ,T6.CRYDOPVL DOPANT                              " & vbLf    ''�����h�[�v��
    '*** UPDATE END   T.TERAUCHI 2004/11/27
        
        sSql = sSql & "         ,T2.PGID                                        " & vbLf    ''PGID
        sSql = sSql & "         ,T3.HSXTYPE                                     " & vbLf    ''�^�C�v
        sSql = sSql & "         ,(T3.HSXD1MIN + T3.HSXD1MAX) / 2 AS PRODMCX     " & vbLf    ''���i�a
        sSql = sSql & "         ,T1.SUICHARGE                                  " & vbLf     ''����`���[�W��
        sSql = sSql & "         ,T4.HSXLTHWS                                    " & vbLf    ''���C�t�^�C���d�l�L��
    
    '*** UPDATE START T.TERAUCHI 2004/10/21
        sSql = sSql & "         ,T5.CTR01A9 * 1000 AS CTR01A9" & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/10/21
    
    '*** UPDATE START T.TERAUCHI 2004/12/08
        sSql = sSql & "         ,T1.LENTKC1 AS LENTKC1" & vbLf                              ''��������
    '*** UPDATE END   T.TERAUCHI 2004/12/08
    
    '*** UPDATE START Marushita 2011/03/23 TSMC�i���ʑΉ�
        sSql = sSql & "         ,T7.MTRLCHKFLG AS MTRLCHKFLG" & vbLf                        ''���������`�F�b�N�t���O(�K�i)
    '*** UPDATE END   Marushita 2011/03/23
    
        sSql = sSql & "from     XSDC1 T1                                        " & vbLf
        sSql = sSql & "         ,TBCMH001 T2                                    " & vbLf
        sSql = sSql & "         ,TBCME018 T3                                    " & vbLf
        sSql = sSql & "         ,TBCME019 T4                                    " & vbLf
    
    '*** UPDATE START T.TERAUCHI 2004/10/21
        sSql = sSql & "             ,KODA9 T5 " & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/10/21
        
    '*** UPDATE START T.TERAUCHI 2004/11/27 �����h�[�v�A�����h�[�v�ʕύX�Ή�
        sSql = sSql & "             ,TBCMH002 T6 "
    '*** UPDATE END   T.TERAUCHI 2004/11/27
    '*** UPDATE START Marushita 2011/03/23 TSMC�i���ʑΉ�
        sSql = sSql & "         ,TBCME036 T7 " & vbLf
    '*** UPDATE END   Marushita 2011/03/23
    
    '*** UPDATE START T.TERAUCHI 2004/10/21
    '    sSql = sSql & "where    SUBSTRB(T1.XTALC1,1,9)                          " & vbLf
    '    sSql = sSql & "         = SUBSTRB('" & sCrystalNo & "',1,9)             " & vbLf
        sSql = sSql & "where    T1.XTALC1                          " & vbLf
        sSql = sSql & "         = '" & sCrystalNo & "'             " & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/10/21
        
        sSql = sSql & "and      SUBSTRB(T1.XTALC1,1,7) || '0' || SUBSTRB(T1.XTALC1,9,1)     " & vbLf
        sSql = sSql & "         = SUBSTRB(T2.UPINDNO,1,7) || '0' || SUBSTRB(T2.UPINDNO,9,1) " & vbLf
        sSql = sSql & "and      T2.HINBAN                                       " & vbLf
        sSql = sSql & "         = T3.HINBAN                                     " & vbLf
        sSql = sSql & "and      T2.NMNOREVNO                                    " & vbLf
        sSql = sSql & "         = T3.MNOREVNO                                   " & vbLf
        sSql = sSql & "and      T2.NFACTORY                                     " & vbLf
        sSql = sSql & "         = T3.FACTORY                                    " & vbLf
        sSql = sSql & "and      T2.NOPECOND                                     " & vbLf
        sSql = sSql & "         = T3.OPECOND                                    " & vbLf
        sSql = sSql & "and      T2.HINBAN                                       " & vbLf
        sSql = sSql & "         = T4.HINBAN                                     " & vbLf
        sSql = sSql & "and      T2.NMNOREVNO                                    " & vbLf
        sSql = sSql & "         = T4.MNOREVNO                                   " & vbLf
        sSql = sSql & "and      T2.NFACTORY                                     " & vbLf
        sSql = sSql & "         = T4.FACTORY                                    " & vbLf
        sSql = sSql & "and      T2.NOPECOND                                     " & vbLf
        sSql = sSql & "         = T4.OPECOND                                    " & vbLf
        
    '*** UPDATE START T.TERAUCHI 2004/10/21
        sSql = sSql & "AND          T5.SYSCA9 = 'K' " & vbLf
        sSql = sSql & "AND          T5.SHUCA9 = 'A7' " & vbLf
        sSql = sSql & "AND          T5.CODEA9 = '300'" & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/10/21
    
    '*** UPDATE START T.TERAUCHI 2004/11/27
        sSql = sSql & "and      SUBSTRB(T1.XTALC1,1,7) || '0' || SUBSTRB(T1.XTALC1,9,1)     " & vbLf
        sSql = sSql & "         = SUBSTRB(T6.UPINDNO,1,7) || '0' || SUBSTRB(T6.UPINDNO,9,1) " & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/11/27
    '*** UPDATE START Marushita 2011/03/23 TSMC�i���ʑΉ�
        sSql = sSql & "and      T2.HINBAN                                       " & vbLf
        sSql = sSql & "         = T7.HINBAN                                     " & vbLf
        sSql = sSql & "and      T2.NMNOREVNO                                    " & vbLf
        sSql = sSql & "         = T7.MNOREVNO                                   " & vbLf
        sSql = sSql & "and      T2.NFACTORY                                     " & vbLf
        sSql = sSql & "         = T7.FACTORY                                    " & vbLf
        sSql = sSql & "and      T2.NOPECOND                                     " & vbLf
        sSql = sSql & "         = T7.OPECOND                                    " & vbLf
    '*** UPDATE END   Marushita 2011/03/23 TSMC�i���ʑΉ�
    
    End Sub
    '2004/10/15tcs Yamauchi �ǉ� end-------------------------------------
    
    '---- ADD [���������Ǘ��A�����H�����э쐬�Ή�] 2004/10/29 TCS)R.Kawaguchi START ----
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' �@�\    : ���d�ʎ擾
    '
    ' �Ԃ�l  : ����
    '
    ' ������  : sDiameterKbn    - ���a�敪
    '           dKataWeight     - ���d��
    '
    ' �@�\����: ���d�ʎ擾
    '
    ' ���l    :
    '
    '///////////////////////////////////////////////////
    Public Function GetKataWeight(ByVal sDiameterKbn As String, dKataWeight As Double) As Boolean
    
        Dim sSql        As String
        Dim objOraDyn   As Object
    
    On Error GoTo ErrHand
    
        GetKataWeight = False
    
        sSql = ""
        sSql = sSql & "select   CTR01A9" & vbLf
        sSql = sSql & "from     KODA9" & vbLf
        sSql = sSql & "where    SYSCA9 = 'K'" & vbLf
        sSql = sSql & "and      SHUCA9 = 'A7'" & vbLf
        sSql = sSql & "and      CODEA9 = '" & sDiameterKbn & "'" & vbLf
    
        'SQL�����s
        If DynSet2(objOraDyn, sSql) = False Then
            ''�擾���s
            Call MsgOut(100, sSql, ERR_DISP_LOG, "KODA9")
            Set objOraDyn = Nothing
            Exit Function
        End If
    
        ''�ް��Ȃ�
        If objOraDyn.EOF Then
            Call MsgOut(55, "�Ǘ�����ð���", ERR_DISP)
            Set objOraDyn = Nothing
            Exit Function
        End If
    
        dKataWeight = objOraDyn.Fields("CTR01A9").Value
        
        '�J��
        If Not objOraDyn Is Nothing Then objOraDyn.Close: Set objOraDyn = Nothing
        
        GetKataWeight = True
        Exit Function
    
ErrHand:
        ''�װ
        Call MsgOut(100, "", ERR_DISP_LOG, "KODA9")
        '�J��
        If Not objOraDyn Is Nothing Then objOraDyn.Close: Set objOraDyn = Nothing
    
    End Function
    '---- ADD [���������Ǘ��A�����H�����э쐬�Ή�] 2004/10/29 TCS)R.Kawaguchi END ----
    
'>>>>>>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -----------------START
'>>>>>>>>>> Ins_TBCMC001_New�֐���s_cmzclabel.bas�Ɉړ��̂��߃R�����g�� -----
'
'    '*** UPDATE START T.TERAUCHI 2004/12/06 ����������┭�����A�����������x�����s�pӼޭ�ْǉ�
'    ' @(f)
'    '
'    ' �@�\    : ���x�����s�pӼޭ��
'    '
'    ' �Ԃ�l  : True:���� False:���s
'    '
'    ' ������  : after:���H���R�[�h
'    '           cyoku:���敪
'    '
'    ' �@�\����:�@����������┭�����A�����������x���𔭍s����
'    '
'    ' ���l    :
'    '       �����F
'    '           sProcCode   �H������
'    '           sEtcPrKind  ���̑����َ��
'    '           sStaffID    �v���S����
'    '           sPrKey01    ���[���ް�1
'    '           sSysdate    ������t
'    '           sRegDate    �o�^���t�@���o�^���t��PK�ׁ̈A���̏����ŕ������o�^����ꍇ�A
'    '                                   �Ăяo������1�b���炷���̐��䂪�K�v
'    '       �g�p��۸��сF
'    '           cmbc008     �ؽ�ٶ�۸ތ����i�グ
'    '           cmbc030     ������������
'    '           cmbc018     �ؒf�E����َw���Ɖ�
'    '
'    Public Function Ins_TBCMC001_New(sProcCode As String, sEtcPrKind As String, sStaffID As String, sPrKey01 As String, sSysDate As String) As Boolean
'        Dim sSql      As String       'SQL���i�[
'        Dim iRet    As Integer        '�f�[�^�ǉ���
'
'    '�G���[�n���h��
'    On Error GoTo ErrHand
'
'        '�߂�l�ݒ�
'        Ins_TBCMC001_New = False
'
'        '���߭�����ݒ�
'        gsCompName = GetCompName
'
'        '�o�^�p��ذ�ݒ�
'        sSql = ""
'        sSql = sSql & "insert into tbcmc001(" & vbLf    ''
'        sSql = sSql & "                 quedate                     " & vbLf    ''�L���[���t
'        sSql = sSql & "                 ,reqkind                    " & vbLf    ''����v���敪
'        sSql = sSql & "                 ,printkind                  " & vbLf    ''������
'        sSql = sSql & "                 ,endflg                     " & vbLf    ''�����敪
'        sSql = sSql & "                 ,status                     " & vbLf    ''�I���X�e�[�^�X
'        sSql = sSql & "                 ,blockidumu                 " & vbLf    ''�u���b�NID�L���敪
'        sSql = sSql & "                 ,proccode                   " & vbLf    ''�H���R�[�h
'        sSql = sSql & "                 ,etcprkind                  " & vbLf    ''���̑����x�����
'        sSql = sSql & "                 ,crynum                     " & vbLf    ''�����ԍ�
'        sSql = sSql & "                 ,ingotpos                   " & vbLf    ''�������ʒu
'        sSql = sSql & "                 ,smplno                     " & vbLf    ''�T���v��No
'        sSql = sSql & "                 ,mtrlnum                    " & vbLf    ''�����ԍ�
'        sSql = sSql & "                 ,smtrlnum                   " & vbLf    ''���������ԍ�
'        sSql = sSql & "                 ,blockid                    " & vbLf    ''�u���b�NID
'        sSql = sSql & "                 ,hinban                     " & vbLf    ''�i��
'        sSql = sSql & "                 ,revnum                     " & vbLf    ''���i�ԍ�����ԍ�
'        sSql = sSql & "                 ,factory                    " & vbLf    ''�H��
'        sSql = sSql & "                 ,opecond                    " & vbLf    ''���Ə���
'        sSql = sSql & "                 ,cryindrs                   " & vbLf    ''���������w��(Rs)
'        sSql = sSql & "                 ,cryindoi                   " & vbLf    ''���������w��(Oi)
'        sSql = sSql & "                 ,cryindb1                   " & vbLf    ''���������w��(B1)
'        sSql = sSql & "                 ,cryindb2                   " & vbLf    ''���������w��(B2)
'        sSql = sSql & "                 ,cryindb3                   " & vbLf    ''���������w��(B3)
'        sSql = sSql & "                 ,cryindl1                   " & vbLf    ''���������w��(L1)
'        sSql = sSql & "                 ,cryindl2                   " & vbLf    ''���������w��(L2)
'        sSql = sSql & "                 ,cryindl3                   " & vbLf    ''���������w��(L3)
'        sSql = sSql & "                 ,cryindl4                   " & vbLf    ''���������w��(L4)
'        sSql = sSql & "                 ,cryindcs                   " & vbLf    ''���������w��(Cs)
'        sSql = sSql & "                 ,cryindgd                   " & vbLf    ''���������w��(Gd)
'        sSql = sSql & "                 ,cryindt                    " & vbLf    ''���������w��(T)
'        sSql = sSql & "                 ,cryindep                   " & vbLf    ''���������w��(EPD)
'        sSql = sSql & "                 ,staffid                    " & vbLf    ''�v���S����
'        sSql = sSql & "                 ,machine                    " & vbLf    ''�v���}�V����
'        sSql = sSql & "                 ,regdate                    " & vbLf    ''�o�^���t
'        sSql = sSql & "                 ,upddate                    " & vbLf    ''�X�V���t
'        sSql = sSql & "                 ,prkey01                     " & vbLf    ''���[�L�[�f�[�^�P
'        sSql = sSql & "     )                                       " & vbLf
'        sSql = sSql & "values(                                      " & vbLf
'        sSql = sSql & "                 to_date('" & sSysDate & "','yyyy/mm/dd hh24:mi:ss')       " & vbLf    ''�L���[���t
'        sSql = sSql & "                 ,'0'                        " & vbLf    ''����v���敪
'        sSql = sSql & "                 ,'1'                        " & vbLf    ''������
'        sSql = sSql & "                 ,'0'                        " & vbLf    ''�����敪
'        sSql = sSql & "                 ,'0'                        " & vbLf    ''�I���X�e�[�^�X
'        sSql = sSql & "                 ,'0'                        " & vbLf    ''�u���b�NID�L���敪
'        sSql = sSql & "                 ,'" & sProcCode & "'        " & vbLf    ''�H���R�[�h
'        sSql = sSql & "                 ,'" & sEtcPrKind & "'       " & vbLf    ''���̑����x�����
'        sSql = sSql & "                 ,null                       " & vbLf    ''�����ԍ�
'        sSql = sSql & "                 ,null                       " & vbLf    ''�������ʒu
'        sSql = sSql & "                 ,null                       " & vbLf    ''�T���v��No
'        sSql = sSql & "                 ,null                       " & vbLf    ''�����ԍ�
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������ԍ�
'        sSql = sSql & "                 ,null                       " & vbLf    ''�u���b�NID
'        sSql = sSql & "                 ,null                       " & vbLf    ''�i��
'        sSql = sSql & "                 ,null                       " & vbLf    ''���i�ԍ�����ԍ�
'        sSql = sSql & "                 ,null                       " & vbLf    ''�H��
'        sSql = sSql & "                 ,null                       " & vbLf    ''���Ə���
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(Rs)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(Oi)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(B1)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(B2)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(B3)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(L1)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(L2)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(L3)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(L4)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(Cs)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(Gd)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(T)
'        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(Epd)
'        sSql = sSql & "                 ,'" & sStaffID & "'         " & vbLf    ''�v���S���Җ�
'        sSql = sSql & "                 ,'" & gsCompName & "'       " & vbLf    ''
'        sSql = sSql & "                 ,SYSDATE                    " & vbLf    ''�o�^���t
'        sSql = sSql & "                 ,SYSDATE                    " & vbLf    ''�X�V���t
'        sSql = sSql & "                 ,'" & sPrKey01 & "'         " & vbLf    ''���[�L�[�f�[�^�P
'        sSql = sSql & "                             )               " & vbLf
'
'        '���s
'        iRet = SqlExec2(sSql)
'
'        If iRet < 0 Then
'            Call MsgOut(100, sSql, ERR_DISP_LOG, "TBCMC001")
'            Exit Function
'        End If
'
'        '�߂�l�ݒ�
'        Ins_TBCMC001_New = True
'
'    '�G���[��
'ErrHand:
'        ''�װ
'        Call MsgOut(100, "", ERR_DISP_LOG, "TBCMC001")
'    End Function
'    '*** UPDATE END T.TERAUCHI 2004/12/06 ����������┭�����A�����������x�����s�pӼޭ�ْǉ�
''*ADD* Ӽޭ�ٓ��� TCS)K.Kunori 2004.11.29 END <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -----------------END

'�T�v      :Cs����v�Z
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:tCsSuitei    ,I  ,CS_SUITEI_TYPE   ,Cs����v�Z�p���Ұ�
'      �@�@:dCsSuitei    ,O  ,Double           ,Cs����l
'      �@�@:�߂�l       ,O   ,Boolean         ,�v�Z�̐���
'����      :Cs����l���v�Z����
'����      :06/04/20 ooba
Public Function GetCsSuiteiMain(tCsSuitei As CS_SUITEI_TYPE, dCsSuitei As Double) As Boolean
    Dim vGS         As Variant
    Dim vCSSC0      As Variant
    Dim vCS0        As Variant
    Dim vGT         As Variant
    Dim vCSTC0      As Variant
    
    Dim dSiWeight   As String   ''����ޗ�(Kg)
    Dim dTopWT      As String   ''į�ߏd��(Kg)
    Dim dUpDm       As String   ''���a(mm)
    Dim dCsHenseki  As String   ''����ݕΐ͌W��(����Ͻ��ɕێ�)
    Dim dSamplePos  As String   ''����وʒu
    Dim dResCs      As String   ''����ّ���l
    Dim dInfPos     As String   ''����ʒu
    
On Error GoTo ErrHand

    GetCsSuiteiMain = False

    ''�ϐ��i�[
    dSiWeight = CDbl(tCsSuitei.sSiWeight) / 1000    ''����ޗ�(Kg)
    dTopWT = CDbl(tCsSuitei.sTopWT) / 1000          ''į�ߏd��(Kg)
    dUpDm = CDbl(tCsSuitei.sUpDm)                   ''���a(mm)
    dCsHenseki = CDbl(tCsSuitei.sCsHenseki)         ''����ݕΐ͌W��(����Ͻ��ɕێ�)
    dSamplePos = CDbl(tCsSuitei.sSamplePos)         ''����وʒu
    dResCs = CDbl(tCsSuitei.sResCs)                 ''����ّ���l
    dInfPos = CDbl(tCsSuitei.sInfPos)               ''����ʒu

    ''GS = (���a / 20) ^ 2 * 3.14 * 2.33 * ����وʒu / (����ޗ� - TOP�d��) / 1000
    vGS = (dUpDm / 20) ^ 2 * 3.14 * 2.33 * dSamplePos / (dSiWeight - dTopWT) / 10000

    ''CSSC0 = ����ݕΐ͌W�� * (1 - GS) ^ (����ݕΐ͌W�� - 1)
    vCSSC0 = dCsHenseki * (1 - vGS) ^ (dCsHenseki - 1)

    ''CS0 = ����ّ���l / CSSC0
    vCS0 = dResCs / vCSSC0
    
    ''GT = (���a / 20) ^ 2 * 3.14 * 2.33 * ����ʒu / (����ޗ� - TOP�d��) / 1000
    vGT = (dUpDm / 20) ^ 2 * 3.14 * 2.33 * dInfPos / (dSiWeight - dTopWT) / 10000
    
    ''CSTC0 = ����ݕΐ͌W�� * (1 - GT) ^ (����ݕΐ͌W�� - 1)
    vCSTC0 = dCsHenseki * (1 - vGT) ^ (dCsHenseki - 1)
    
    ''����l = CS0 * CSTC0
    dCsSuitei = vCS0 * vCSTC0
    
    GetCsSuiteiMain = True

    Exit Function
    
ErrHand:

    ''�װ
'    Call MsgOut(100, "", ERR_DISP_LOG, "")

End Function

'�֐���    :�����R�[�h�l��
'�T�v      :�Ј����Ƃɓo�^����Ă��錠���i�P�O���j�b�g�j����
'           ��ʂ��Ƃɐݒ肳�ꂽ�Ă��錠�����j�b�g�R�[�h(1�`10���̒l)��
'           �Y�����錠���R�[�h�𔲂��o���Ė߂�l�ɐݒ肷��B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :staffID       ,I  ,String    ,�Ј�ID
'          :picname       ,I  ,String    ,��ʖ�
'          :�߂�l        ,O  ,String    ,�����R�[�h�i 1 : �Q�Ɓ� �X�V�~ ���F�~
'                                                      2 : �Q�Ɓ� �X�V�� ���F�~
'                                                      3 : �Q�Ɓ� �X�V�~ ���F��
'                                                      4 : �Q�Ɓ� �X�V�� ���F�� �j
'����      :������Ȃ������ꍇ�́A�߂�l�ɂO��Ԃ�
'
'
'�ύX����   2009/09 SUMCO Akizuki �����ݒ�̉��C

Public Function Getstaffauthority(STAFFID$, picname) As Integer
Dim dbIsMine As Boolean         '�c�a�I�[�v���t���O
Dim rs1 As OraDynaset           'KODA9(��ʃ}�X�^)�p�_�C�i�Z�b�g
Dim rs2 As OraDynaset           'KODA9(�Ј��}�X�^)�p�_�C�i�Z�b�g
Dim sql As String               '�r�p�k���i�[�̈�
Dim picauthority As Integer     '�������j�b�g�R�[�h

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "mdlCommon.bas -- Function Getstaffauthority"
    
    Getstaffauthority = 0

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    'KODA9(��ʃ}�X�^)���A���Y��ʂ̌������j�b�g�R�[�h�l��
    sql = ""
    sql = "select KCODE01A9 from KODA9 where SYSCA9='K' and SHUCA9='01' and CODEA9='" & picname & "'"
    Set rs1 = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs1("KCODE01A9")) Then
        '�������j�b�g�R�[�h���m�t�k�k�܂��̓X�y�[�X�̏ꍇ�A
        '(=�܂�A�����`�F�b�N�̐ݒ���|���Ă��Ȃ��ꍇ)�@�`�F�b�N���s�킸�S����\��Ԃ�
        Getstaffauthority = 4
    ElseIf Trim(rs1("KCODE01A9")) = "" Then
        Getstaffauthority = 4
    Else
        '�l�������������j�b�g�R�[�h
        picauthority = CInt(rs1("KCODE01A9"))
        
        'KODA9(�Ј��}�X�^)���A���Y�Ј��̌����R�[�h�l��
        sql = ""
        sql = "select KCODE03A9 from KODA9 where SYSCA9='K' and SHUCA9='55' and CODEA9='" & STAFFID & "'"
        Set rs2 = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
        If rs2.RecordCount = 0 Then
            ''�Y������Ј��̏�񂪌�����Ȃ������ꍇ�A�O��Ԃ�
            Getstaffauthority = 0
            
        '>>>2009/09/07 SUMCO Akizuki
        ElseIf rs2.RecordCount = 1 Then
            ''�����R�[�h���m�t�k�k�̏ꍇ�A����0[����s��]��
            If IsNull(rs2("KCODE03A9")) Then
                'Getstaffauthority = 4
                Getstaffauthority = 0
            
            ''�����R�[�h���X�y�[�X�̏ꍇ������0[����s��]��
            ElseIf Trim(rs2("KCODE03A9")) = "" Then
                'Getstaffauthority = 4
                Getstaffauthority = 0
            
            ''�����R�[�h�����j�b�g�R�[�h��菬���� = ��ݒ�Ȃ���̏ꍇ���A����0[����s��]��
            ElseIf (picauthority <= 0) Or (picauthority > Len(Trim(rs2("KCODE03A9")))) Then
                Getstaffauthority = 0
        '<<<
        
            Else
            '�����R�[�h���O�i�����l�j�̏ꍇ�A�����`�F�b�N���s�킸����0[����s��]��
            '2007/08/09 kaga
                'If CInt(Left(rs2("KCODE03A9"), picauthority)) = 0 Then
                If CInt(Mid(Trim(rs2("KCODE03A9")), picauthority, 1)) = 0 Then
                    Getstaffauthority = 0
                '�����R�[�h���O�i�����l�j�̈ȊO�ꍇ�A���Y��ʂɑ΂��錠���R�[�h��Ԃ�
                Else
                    '2007/08/09 kaga
                    'Getstaffauthority = CInt(Left(rs2("KCODE03A9"), picauthority))
                    Getstaffauthority = CInt(Mid(Trim(rs2("KCODE03A9")), picauthority, 1))
                End If
            End If
        End If
    End If
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

PROC_ERR:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

    
'--------------- 2008/08/25 INSERT START  By Systech ---------------
'�T�v      :�e�[�u���uTBCME036�v��������ɂ��������R�[�h��DK���x�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :bFlg          ,I  ,Boolean         ,TRUE=�d�l, FALSE=����
'          :xsdcs         ,I  ,typ_XSDCS       ,�V�T���v���Ǘ�
'          :�߂�l        ,O  ,String          ,DK���x
'����      :
'����      :
Public Function GetDKTmpCode(bFlg As Boolean, xsdcs As typ_XSDCS) As String
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim i           As Long
    
    GetDKTmpCode = ""

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "mdlCommon.bas -- Function GetDKTmpName"
    
    ''SQL��g�ݗ��Ă�
    If bFlg Or xsdcs.CRYINDRSCS = "1" Then
        ' DK���x(�d�l)���͏��FLG(Rs)���ʏ�̏ꍇ�A���Y�i�Ԃ�DK���x���擾
        sql = "SELECT NVL(HSXDKTMP, ' ') AS HSXDKTMP"
        sql = sql & " FROM TBCME036"
        sql = sql & " WHERE HINBAN = '" & xsdcs.HINBCS & "'"
        sql = sql & " AND MNOREVNO = " & xsdcs.REVNUMCS
        sql = sql & " AND FACTORY = '" & xsdcs.FACTORYCS & "'"
        sql = sql & " AND OPECOND = '" & xsdcs.OPECS & "'"
    Else
        ' DK���x(����)�̏ꍇ�A���f���i�Ԃ�DK���x���擾
        sql = "SELECT NVL(A.HSXDKTMP, ' ') AS HSXDKTMP"
        sql = sql & " FROM TBCME036 A, XSDCS B"
        sql = sql & " WHERE B.XTALCS = '" & xsdcs.XTALCS & "'"
        sql = sql & " AND B.CRYSMPLIDRSCS = '" & xsdcs.CRYSMPLIDRSCS & "'"
        sql = sql & " AND B.CRYINDRSCS = '1'"
        sql = sql & " AND A.HINBAN = B.HINBCS"
        sql = sql & " AND A.MNOREVNO = B.REVNUMCS"
        sql = sql & " AND A.FACTORY = B.FACTORYCS"
        sql = sql & " AND A.OPECOND = B.OPECS"
    End If
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        ' DK���x��Ԃ�
        GetDKTmpCode = rs("HSXDKTMP")
    End If
    rs.Close

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

PROC_ERR:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�e�[�u���uTBCME036�v��������ɂ��������R�[�h��DK���x�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :bFlg          ,I  ,Boolean         ,TRUE=�d�l, FALSE=����
'          :xsdcw         ,I  ,typ_XSDCW       ,�V�T���v���Ǘ�
'          :�߂�l        ,O  ,String          ,DK���x
'����      :
'����      :
Public Function GetWfDKTmpCode(bFlg As Boolean, xsdcw As typ_XSDCW) As String
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim i           As Long
    
    GetWfDKTmpCode = ""

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "mdlCommon.bas -- Function GetWfDKTmpCode"
    
    ''SQL��g�ݗ��Ă�
    If bFlg Or xsdcw.WFINDRSCW = "1" Then
        ' DK���x(�d�l)���͏��FLG(Rs)���ʏ�̏ꍇ�A���Y�i�Ԃ�DK���x���擾
        sql = "SELECT NVL(HSXDKTMP, ' ') AS HSXDKTMP"
        sql = sql & " FROM TBCME036"
        sql = sql & " WHERE HINBAN = '" & xsdcw.HINBCW & "'"
        sql = sql & " AND MNOREVNO = " & xsdcw.REVNUMCW
        sql = sql & " AND FACTORY = '" & xsdcw.FACTORYCW & "'"
        sql = sql & " AND OPECOND = '" & xsdcw.OPECW & "'"
    Else
        ' DK���x(����)�̏ꍇ�A���f���i�Ԃ�DK���x���擾
        sql = "SELECT NVL(A.HSXDKTMP, ' ') AS HSXDKTMP"
        sql = sql & " FROM TBCME036 A, XSDCW B"
        sql = sql & " WHERE B.XTALCW = '" & xsdcw.XTALCW & "'"
        sql = sql & " AND B.WFSMPLIDRSCW = '" & xsdcw.WFSMPLIDRSCW & "'"
        sql = sql & " AND B.WFINDRSCW = '1'"
        sql = sql & " AND A.HINBAN = B.HINBCW"
        sql = sql & " AND A.MNOREVNO = B.REVNUMCW"
        sql = sql & " AND A.FACTORY = B.FACTORYCW"
        sql = sql & " AND A.OPECOND = B.OPECW"
    End If
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        ' DK���x��Ԃ�
        GetWfDKTmpCode = rs("HSXDKTMP")
    End If
    rs.Close

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

PROC_ERR:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :DK���x���̂��A�擪�̐��l("��"�ȑO)��Ԃ�
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :szBuf         ,I  ,String       , DK���x����(�ėp�R�[�h�}�X�^�̃R�[�h���e)
'          :�߂�l        ,O  ,String       , DK���x���̐擪�̐��l
'����      :
'����      :
Public Function GetDKTmpDispName(szBuf As String) As String
    Dim i       As Integer
    
    i = InStr(1, szBuf, "��") - 1
    If i > 0 And Len(szBuf) >= i Then
        GetDKTmpDispName = left(szBuf, i)
    Else
        GetDKTmpDispName = szBuf
    End If

End Function
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'�T�v      :�a�ENotch�ʒu���ʂ�\���p�ɕϊ�
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :sDPDIR        ,I  ,String       ,�a�ENotch�ʒu����
'          :�߂�l        ,O  ,String       ,�\��������
'����      :�a�ENotch�ʒu���ʂ�\���p�ɕϊ�
'2009/05/22 SETsw kubota
Public Function CnvMizoNotchDisp(ByVal sDPDIR As String) As String

    CnvMizoNotchDisp = ""
    
    Select Case sDPDIR
    Case "B1", "B2", "B3", "B4", "D3", "D4"
        CnvMizoNotchDisp = "0"
    Case "B5", "B6", "B7", "B8", "D1", "D2"
        CnvMizoNotchDisp = "45"
    End Select

End Function

'///////////////////////////////////////////////////
' @(f)
'
' �@�\      : ���b�Z�[�W�ҏW�E��ʕ\���E���O�o��
'
' �Ԃ�l    : �Ȃ�
'
' ������    : arg1:ү���޺��� 100:�׸ٴװ 100�ȊO:�׸ٴװ�ȊO
'             arg2:�Z�b�V����
'             arg3:DB
'             arg4:�ǉ�ү����
'             arg5:ү���ޑ����@0:�ʏ�ү����
'                              1:��ʕ\���װү���ށi���͗��ԕ\���̴װ�Ȃǁj
'                              2:���O�o�ʹװү����
'                              3:��ʕ\���E���O�o�ʹװү���ށi�׸ٴװ�Ȃǁj
'                              5:��ʕ\�����ޯ��ү����
'                              6:���O�o�����ޯ��ү����
'             arg6:�׸ٴװ����۸�/��ʕ\��ð��ٖ�
'
' �@�\����  : MsgOut��DB�ڑ�Object�w���
' ���l      :
'2009/05/28�ǉ� SETsw kubota
'///////////////////////////////////////////////////
Public Sub MsgOut_DB(ByVal iMsgCd As Integer _
                   , ByRef objSess As Object _
                   , ByRef objDB As Object _
                   , Optional ByVal sAddMsgStr As String = "" _
                   , Optional ByVal eMsgKind As MsgKind = 0 _
                   , Optional ByVal TABLENAME As String = "Unknown" _
                   )
    Dim sMsg As String                              ''���b�Z�[�W
    Dim sOraErrCd As String                         ''�׸ٴװ����
    
    '���b�Z�[�W������
    Call MsgInit

    'ү���ޑ�����ү���ޏo�͑����͈͊O�̏ꍇ�o�͂��Ȃ��i�J���^�p�J�n��A���ޯ��ү���ނ��o�͂��Ȃ��悤�ɂł���j
    If Not ((eMsgKind = NORMAL_MSG) Or _
            ((eMsgKind And MsgKindMask) <> 0)) Then
        Exit Sub                                    ''�I��
    End If
    
    If iMsgCd < 100 Then                            ''ү���޺��ނ��׸وȊO�Ȃ�
        ''�I���N���ȊO�̃��b�Z�[�W
        On Error Resume Next                        ''�װ�ׯ��
        sMsg = msMsgStr(iMsgCd)                     ''���b�Z�[�W�擾
        On Error GoTo 0                             ''�װ�ׯ�߉���
    Else                                            ''ү���޺��ނ��׸ٴװ�Ȃ�
        ''�I���N���̃G���[���b�Z�[�W
        If objSess.LastServerErr Then           ''�׸پ���ݵ�޼ު�ẴG���[�Ȃ��
            sMsg = objSess.LastServerErrText    ''�׸پ���ݵ�޼ު�Ĵװү���ނ��Z�b�g
            objSess.LastServerErrReset          ''�׸پ���ݵ�޼ު�Ĵװ�����Z�b�g
        ElseIf objDB.LastServerErr Then         ''�׸��ް��ް���޼ު�ẴG���[�Ȃ��
            ''�׸ٴװү���ނ��׸ٴװ���ނ�؏o��
            sOraErrCd = GetStrOraErrCd(objDB.LastServerErrText)
            If sOraErrCd <> "" Then                 ''�׸ٴװ���ނ������Ă����
                sMsg = "DB�G���[�i" & TABLENAME & ")" & sOraErrCd ''�w��̃t�H�[�}�b�g�ŕҏW
                sAddMsgStr = objDB.LastServerErrText & _
                             "::" & sAddMsgStr
            Else                                    ''�׸ٴװ���ނ������Ă��Ȃ����
                sMsg = objDB.LastServerErrText  ''�׸��ް��ް���޼ު�Ĵװү���ނ��Z�b�g
            End If
            objDB.LastServerErrReset            ''�׸��ް��ް���޼ު�Ĵװ�����Z�b�g
        ElseIf Err.Number Then                      ''����VB�̴װ�������Ȃ�
            sMsg = Error(Err.Number)                ''VB�̴װү���ނ��Z�b�g
        Else                                        ''���ʹװ����Ȃ��Ȃ��
            sMsg = "�׸ِ��펞�ɴװ�o�͂���"         ''�x��
        End If
    End If
    
    If (eMsgKind = NORMAL_MSG) Or _
       (eMsgKind And ERR_DISP) Then                     ''�ʏ�ү���ނ���ʕ\���r�b�g�������Ă����
        ''�G���[�Ȃ�ԕ\��
        If (eMsgKind = ERR_DISP) Or _
           (eMsgKind = ERR_DISP_LOG) Then
            If iMsgCd = 100 Then                        ''�I���N���G���[�̏ꍇ
                MsgDisp sMsg, vbRed                     ''���b�Z�[�W����ʕ\������
            Else
                MsgDisp sMsg & sAddMsgStr, vbRed        ''���b�Z�[�W & �ǉ����b�Z�[�W����ʕ\������
            End If
        ''����ȊO�͍��\��
        Else
            If iMsgCd = 100 Then                        ''�I���N���G���[�̏ꍇ
                MsgDisp sMsg                            ''���b�Z�[�W����ʕ\������
            Else
                MsgDisp sMsg & sAddMsgStr               ''���b�Z�[�W & �ǉ����b�Z�[�W����ʕ\������
            End If
        End If
    End If
    
    If eMsgKind And ERR_LOG Then                    ''���O�o�̓r�b�g�������Ă����
        MsgLog (Format(Now, "YYYY/MM/DD HH:NN:SS::") & App.EXENAME & "::" & _
            iMsgCd & "::" & sMsg & "::" & sAddMsgStr) ''���b�Z�[�W�����O�o�͂���
    End If
    
    If (eMsgKind = ERR_DISP) Or _
       (eMsgKind = ERR_LOG) Or _
       (eMsgKind = ERR_DISP_LOG) Then                       ''ү���ޑ������G���[�Ȃ�
        Beep
    End If
End Sub

'�T�v      :���ʗp�ƭ��J�ڃ{�^������
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :iIndex        ,I  ,Integer      ,1:Ҳ��ƭ��J�� 2:����ƭ��J��
'����      :
'2009/08/12 SETsw kubota
Public Sub execSubClick(ByVal iIndex As Integer)
    
    Select Case iIndex
    Case 1      'Ҳ��ƭ�
        GotoMainMenu
    Case 2      '����ƭ�
        GotoSubMenu
    End Select

End Sub

