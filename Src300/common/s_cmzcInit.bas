Attribute VB_Name = "s_cmzcInit"
Option Explicit

Public Const SW_SHOW = 5
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public gErr As CErrHandler  '�G���[�n���h��(from VB SourceBook)

'�T�v      :�v���O�����N�����̏���������
'����      :
Public Function InitExe() As FUNCTION_RETURN
    
    '' �v���O�����N�����̏���������
    DoEvents
    
    '' �p�����[�^������
    InitExe = FUNCTION_RETURN_SUCCESS
   
    '' �G���[�o�̓I�u�W�F�N�g�쐬
    Init_ErrHandler
    
    ''�R�}���h���C�������擾
    If GetCmdLine() = False Then
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
        InitExe = FUNCTION_RETURN_FAILURE
        End
    End If
    
    '' �f�[�^�x�[�X�ڑ�
    OraDBOpen
    
    '' �����I��

End Function

Private Sub Init_ErrHandler()
    Set gErr = New CErrHandler
    With gErr
        .AppTitle = App.Title
        .Destination = App.Path & "\Err.log"
        .DisplayMsgOnError = True
        .MaxProcStackItems = 20
        .IncludeExpandedInfo = False
    End With
End Sub

Private Sub TerminateHandler()
    On Error Resume Next
    Set gErr = Nothing
    On Error GoTo 0
End Sub


'///////////////////////////////////////////////////
' @(f)
' �@�\    : ���C�����j���[�ɑJ�ڂ���
'
' �Ԃ�l  :
'
' ������  :
'
' �@�\����:
'
'///////////////////////////////////////////////////
Public Sub GotoMainMenu()
    Dim sCallCd As String
    sCallCd = "0000000"
    If gbFTPFlg = True Then             ''FTP�N���t���O�������Ă�����
        sCallCd = UCase(App.EXENAME)    ''�����W���[������n��
    End If
    Call ExitExe(sCallCd) ''���C�����j���[���N�����A�I��
End Sub

'///////////////////////////////////////////////////
' @(f)
' �@�\    : �T�u���j���[�ɑJ�ڂ���
'
' �Ԃ�l  :
'
' ������  :
'
' �@�\����:
'
'///////////////////////////////////////////////////
Public Sub GotoSubMenu()
    Dim sCallCd As String
    sCallCd = gsCallCd    ''�󂯎�����ďo�敪��n��
    Call ExitExe(sCallCd) ''�T�u���j���[���N�����A�I��
End Sub


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �I������
'
' �Ԃ�l  :
'
' ������  : �ďo�敪�i�ȗ������ꍇ�A�ďo�����j���[���N������j
'
' �@�\����: �I������
'
'///////////////////////////////////////////////////
Public Sub ExitExe(Optional sCallCd As String = "0000000")
    Dim sExeName As String          ''���s�t�@�C����
'    On Error GoTo Er
    On Error GoTo proc_err
    gErr.Push "s_cmzcCtl.bas -- Function ExitExe"
    sExeName = "XMAIN"
    
    DoEvents                            ''���j���[���܂��I�����ĂȂ��\��������̂ł�����Ƒ҂�
   ''���j���[�J�ڋ��Ȃ�
 '   If mbMenuRet = True Then
        ''�R�}���h���C���擾 �i�N�����̍H��R�[�h�Ń��j���[�ɖ߂�j
        gsFactryCd = Left(Command, 2)    ''�H��R�[�h(2��)
        myFactryCd = Mid(Command, 24, 2)   ''�H��R�[�h(2��)
        gsHinban = Mid(Command, 12, 11)
        If Len(gsHinban) <> 11 Then
            gsHinban = "00000000000"
        End If
        
        ''�N��
        If 0 = Shell(sExeName & " " & gsFactryCd & " " & sCallCd & " " & gsHinban & " " & myFactryCd, vbNormalFocus) Then
            ''�O���߂��Ă�����ُ�
            Call MsgOut(65, sExeName, ERR_DISP_LOG)
        End If
  '  End If
    
   
    ''���j���[�J�ڋ��Ȃ�
'    If mbMenuRet = True Then
'        ''�N��
'        If 0 = Shell(sExeName & " " & sCallCd, vbNormalFocus) Then
'            ''�O���߂��Ă�����ُ�
'            Call MsgOut(65, sExeName, ERR_DISP_LOG)
'
'            WriteDBLog " ", "���j���[�̋N���Ɏ��s���܂����B"
'        End If
'    End If

    
'Er: On Error Resume Next
'    ''�I���N���f�B�X�R�l�N�g
'    Call OraDisConn
'    On Error GoTo 0
'    End ''�I��

proc_exit:
    gErr.Pop

    '' �f�[�^�x�[�X�ڑ��I��
    OraDBClose
    '' �G���[�o�̓I�u�W�F�N�g�j��
    TerminateHandler
    

    End ''�I��
    
    
proc_err:
    gErr.HandleError
    Resume proc_exit


End Sub

