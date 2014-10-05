Attribute VB_Name = "mdlInetCheck"
Option Explicit
'///////////////////////////////////////////////////
' @(S)
'       Inet�R���g���[���g�p���̑��d�N���`�F�b�N����
'
' @(h)  mdlGetFile.bas ver 1.0      ( 2004.12.02 �E�c�@�� )
'
'///////////////////////////////////////////////////


'�N���X�����̓L���v�V��������^���ăE�C���h�E�̃n���h�����擾
Private Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long


'���d�N���`�F�b�N�𑱂��鎞��(�b)   ���̕b����҂��ďI�����Ȃ��ꍇ�A�G���[�ŏI��
Private Const MULTIBOOT_CHECKTIME As Long = 7


'**********************************************************************
' @(f)
'
' �@�\�@�@ : ����E�C���h�E���N���̔���
'
' �Ԃ�l�@ : True  �c ����E�C���h�E�������݂���
' �@�@�@�@   False �c ����E�C���h�E�������݂��Ȃ�
'
' �������@ : �����������E�C���h�E��
'
' �@�\���� : ����E�C���h�E���N���̔���
'
' ���l�@�@ :
'**********************************************************************
Private Function CheckWindowName(ByVal strWindowName As String) As Boolean
    
    Dim lnghwnd As Long
    
    '�E�C���h�E����^���ăn���h�����擾����
    lnghwnd = FindWindow(vbNullString, strWindowName)
    If lnghwnd = 0 Then
        '���ꖼ�̃E�C���h�E�Ȃ�
        CheckWindowName = False
    Else
        '���ꖼ�̃E�C���h�E���J���Ă���
        CheckWindowName = True
    End If

End Function


'**********************************************************************
' @(f)
'
' �@�\�@�@ : ���d�N���̔���
'
' �Ԃ�l�@ : True  �c ���d�N�����Ă��Ȃ�(����)
' �@�@�@�@   False �c ���d�N�����Ă���@(�ُ�)
'
' �������@ : �N���t�H�[��
'
' �@�\���� : ���d�N���̔���
'
' ���l�@�@ :
'**********************************************************************
Private Function CheckMultipleBoot(ByRef frmMain As Form) As Boolean
    
    Dim strWindowName   As String       '' exe�E�C���h�E��
    Dim strFormCaption  As String       '' �t�H�[���̃E�C���h�E��
    Dim lngCount        As Long         '' ���[�v�J�E���^
    Dim blnResult       As Boolean      '' �E�C���h�E���`�F�b�N����

    '' �Ԃ�l������
    CheckMultipleBoot = False

    '' �E�C���h�E����ێ����ύX
    strWindowName = App.Title
    App.Title = App.Title & "_Check"
    strFormCaption = frmMain.Caption
    frmMain.Caption = frmMain.Caption & " ���d�N���`�F�b�N��"
    
    '' �`�F�b�N�����b�Z�[�W�\��
    Call MsgOut(0, "���d�N���`�F�b�N���ł�", NORMAL_MSG)
    
    '' �t�H�[���̃L���v�V�������Ń`�F�b�N(�����c���Ă��邩�`�F�b�N)
    If CheckWindowName(strFormCaption) = True Then
        '' ���d�N��(�ُ�I��)
        Exit Function
    End If
    
    '' �O���exe���I������܂ő҂�
    For lngCount = 1 To MULTIBOOT_CHECKTIME
        
        '' 1�b�҂�
        Sleep (1000)
        
        '' �ēx���d�N���̃`�F�b�N
        blnResult = CheckWindowName(strWindowName)
    
        If blnResult = False Then
            '' �I�����Ă����烋�[�v���甲����
            Exit For
        End If
    
    Next lngCount
    
    '' �I���������ǂ�������
    If blnResult = True Then
        '' ���d�N���Ƃ݂Ȃ�(�ُ�I��)
        Exit Function
    End If

    '�E�C���h�E�������ɖ߂�
    App.Title = strWindowName
    frmMain.Caption = strFormCaption

    '' �`�F�b�N�����b�Z�[�W�N���A
    Call MsgOut(0, "", NORMAL_MSG)

    '' ���d�N���Ȃ�(����I��)
    CheckMultipleBoot = True

End Function


'**********************************************************************
' @(f)
'
' �@�\�@�@ : �v���O�����N�����̏���������(Inet�R���g���[���g�p��ʗp)
'
' �Ԃ�l�@ : �Ȃ�
'
' �������@ : �N���t�H�[��
'
' �@�\���� : Inet�R���g���[���g�p��ʗp��InitExe
'
' ���l�@�@ : Inet�R���g���[���g�p����exe�̏I���Ɏ��Ԃ�������A��d�N���G���[�ɂȂ錏�̑Ή���
'**********************************************************************
Public Function InitExe_Inet(ByRef frmMain As Form) As Integer
    
    DoEvents
    
    ''�����������s�F���C�����j���[�N���w��
    InitExe_Inet = MAINMENU_RET
    mbMenuRet = False       ''���j���[�J�ڕs����
    
    ''���O������
    If LogInit() = False Then
        ''���O���������s
        Call MsgOut(61, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    ''�R�}���h���C�������擾
    If GetCmdLine_Hikiage() = False Then
        ''�R�}���h���C����������
        Call MsgOut(64, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    ''���s�t�@�C�����擾
    If GetEXEName = "" Then
        ''���s�t�@�C�����擾���s
        Call MsgOut(62, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    
'�� InitExe() �Ƃ̑���_  *****************************************
    
    '' ���d�N���`�F�b�N
    If App.PrevInstance = True Then
        
        '' Inet�R���g���[���g�p�̏ꍇ�A�O���exe���I�����Ă��Ȃ��\��������ׁA�Ⴄ���@�ōēx�`�F�b�N
        If CheckMultipleBoot(frmMain) = False Then
            ''���d�N������
            Call MsgOut(63, "", ERR_DISP_LOG)
            Exit Function
        End If
        
    End If
    
'�� InitExe() �Ƃ̑���_  *****************************************
    
    
    ''�����������s�FF1�ȊO����s�w��
    InitExe_Inet = EXITSUB_RET
    mbMenuRet = True       ''���j���[�J�ڋ���
    
    ''�I���N���ڑ�
    If OraConn() = False Then
        ''�I���N���ڑ��G���[
        Call MsgOut(100, "", ERR_DISP_LOG)
        Call CtrlCancel(Screen.ActiveForm)      ''Ҳ��ƭ��ȊO�̺��۰ق��g���Ȃ�����
        Exit Function
    End If
    
    ''������������
    InitExe_Inet = NORMAL_RET
    
End Function


'**********************************************************************
' @(f)
'
' �@�\�@�@ : �v���O�����N�����̏���������(Inet�R���g���[���g�p��ʗp)
'
' �Ԃ�l�@ : �Ȃ�
'
' �������@ : �N���t�H�[��
'
' �@�\���� : Inet�R���g���[���g�p��ʗp��InitExe
'
' ���l�@�@ : Inet�R���g���[���g�p����exe�̏I���Ɏ��Ԃ�������A��d�N���G���[�ɂȂ錏�̑Ή���
'**********************************************************************
Public Function InitExe_Re_Inet(ByRef frmMain As Form) As Integer
    
    DoEvents
    
    ''�����������s�F���C�����j���[�N���w��
    InitExe_Re_Inet = MAINMENU_RET
    mbMenuRet = False       ''���j���[�J�ڕs����
    
    ''���O������
    If LogInit() = False Then
        ''���O���������s
        Call MsgOut(61, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    ''�R�}���h���C�������擾
    If GetCmdLine_Re() = False Then
        ''�R�}���h���C����������
        Call MsgOut(64, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    ''���s�t�@�C�����擾
    If GetEXEName = "" Then
        ''���s�t�@�C�����擾���s
        Call MsgOut(62, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    
'�� InitExe() �Ƃ̑���_  *****************************************
    
    '' ���d�N���`�F�b�N
    If App.PrevInstance = True Then
        
        '' Inet�R���g���[���g�p�̏ꍇ�A�O���exe���I�����Ă��Ȃ��\��������ׁA�Ⴄ���@�ōēx�`�F�b�N
        If CheckMultipleBoot(frmMain) = False Then
            ''���d�N������
            Call MsgOut(63, "", ERR_DISP_LOG)
            Exit Function
        End If
        
    End If
    
'�� InitExe() �Ƃ̑���_  *****************************************
    
    
    ''�����������s�FF1�ȊO����s�w��
    InitExe_Re_Inet = EXITSUB_RET
    mbMenuRet = True       ''���j���[�J�ڋ���
    
    ''�I���N���ڑ�
    If OraConn() = False Then
        ''�I���N���ڑ��G���[
        Call MsgOut(100, "", ERR_DISP_LOG)
        Call CtrlCancel(Screen.ActiveForm)      ''Ҳ��ƭ��ȊO�̺��۰ق��g���Ȃ�����
        Exit Function
    End If
    
    ''������������
    InitExe_Re_Inet = NORMAL_RET
    
End Function

