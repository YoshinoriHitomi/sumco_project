Attribute VB_Name = "mdlGetFile"
Option Explicit
'///////////////////////////////////////////////////
' @(S)
'       �t�@�C���擾����
'
' @(h)  mdlGetFile.bas ver 1.0      ( 2004.12.02 �E�c�@�� )
'
'///////////////////////////////////////////////////

''�萔-------------------------------------------

''�o�[�W���������֌W
Const FtpTimeOut = 20               ''�^�C���A�E�g
Const ExtBak = "bk"                 ''�g���qbak

Const INIFILENAME = "DownLoad2.ini" ''�O��_�E�����[�h���tINI�t�@�C��
Const INIFILESECTION = "OTHER"      ''�O��_�E�����[�h���tINI�t�@�C���Z�N�V����

''��`
''--------------API--------------------------------
''INI�t�@�C���֌W
Declare Function GetPrivateProfileString Lib "kernel32" _
     Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) _
     As Long                        ''INI�t�@�C���Ǎ���

Declare Function WritePrivateProfileString Lib "kernel32" _
     Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpString As Any, _
     ByVal lpFileName As String) _
     As Long                        ''INI�t�@�C��������

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


' �t�@�C���Ɋւ���o�[�W���������擾����֐�
Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" ( _
    ByVal lptstrFilename As String, _
    lpdwHandle As Long _
    ) As Long

' �t�@�C���Ɋւ���o�[�W���������擾����֐�
Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" ( _
    ByVal lptstrFilename As String, _
    ByVal dwHandle As Long, _
    ByVal dwLen As Long, _
    lpData As Any _
    ) As Long

' �o�[�W������񃊃\�[�X����I�����ꂽ�o�[�W���������擾����֐�
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" ( _
    pBlock As Any, _
    ByVal lpSubBlock As String, _
    lplpBuffer As Any, _
    puLen As Long _
    ) As Long
    
' ����ʒu����ʂ̈ʒu�Ƀ������u���b�N���ړ�����֐�
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    ByVal Souce As Long, _
    ByVal Length As Long _
    )
    

Public Type VS_FIXEDFILEINFO
    dwSignature         As Long
    dwStrucVersion      As Long
    dwFileVersionMS     As Long
    dwFileVersionLS     As Long
    dwProductVersionMS  As Long
    dwProductVersionLS  As Long
    deFileFlagsMask     As Long
    dwFileFlags         As Long
    dwFileOS            As Long
    dwFileType          As Long
    dwFileDateMS        As Long
    dwFileDateLS        As Long
End Type


'///////////////////////////////////////////////////
' @(f)
' �@�\    :INI�t�@�C���擾
' �Ԃ�l  : ������
' ������  : ARG1 - �Z�N�V������
'           ARG2 - �L�[��
'           ARG2 - �t�@�C����
' �@�\����:INI�t�@�C���擾
'///////////////////////////////////////////////////
Function GetIni(sec As String, Key As String) As String
    Dim strbuf As String * 256
    Dim strLen As Long
    Dim sIniName As String
    
    sIniName = App.Path & "\" & INIFILENAME
    
    strbuf = ""
    strLen = GetPrivateProfileString(sec, Key, "", strbuf, 256, sIniName)
    GetIni = Left(strbuf, strLen)
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    :INI�t�@�C����������
'
' �Ԃ�l  : ����
'
' ������  : ARG1 - �Z�N�V������
'           ARG2 - �L�[��
' �@�\����:INI�t�@�C����������
'
'///////////////////////////////////////////////////
Function SetIni(sec As String, Key As String) As Boolean
    
    Dim sData As String
    Dim sIniName As String
    
    sData = """" & Format$(Now(), "yyyy/mm/dd hh:nn:ss") & """"
    sIniName = App.Path & "\" & INIFILENAME
    
    SetIni = WritePrivateProfileString(sec, Key, sData, sIniName)

End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �_�E�����[�h����
' �Ԃ�l  : True  - ����I��
' �@�@�@�@  False - �ُ�I��
' ������  : sDownLoadFile - �_�E�����[�h�t�@�C����
' �@�@�@�@  sExt          - �g���q
' �@�@�@�@  frmMain       - �ďo�����t�H�[��
' �@�@�@�@  objInet       - Inet�R���g���[��
' �@�\����: �_�E�����[�h����
'///////////////////////////////////////////////////
Public Function ActDownLoad(ByVal sDownLoadFile As String _
                          , ByVal sExt As String _
                          , ByRef frmMain As Form _
                          , ByRef objInet As Inet _
                          ) As Boolean
    
    Dim sLastDLDay      As String   '�O��_�E�����[�h���t
    Dim bDownloadRes    As Boolean  '�_�E�����[�h����
    Dim bCheckRes       As Boolean  '�_�E�����[�h�`�F�b�N����
    Dim iResult         As Integer
    
    ActDownLoad = False
    
    ''�_�E�����[�h�̗v�s�v�𔻒�
    If Dir(App.Path & "\" & sDownLoadFile & sExt) = "" Then
        ''�t�@�C����������΃_�E���[�h���K�v
        bCheckRes = True
    Else
'>>>>> exe�_�E�����[�h�Ή� 2008/07/01 SETsw kubota -------------
'        ''�O��_�E�����[�h���t��INI�t�@�C�����Q�Ƃ���
'        sLastDLDay = GetIni(INIFILESECTION, sDownLoadFile)
'        ''�O��_�E�����[�h���t�ȍ~�ɍX�V����Ă��邩�𔻒�
'        If ChkDownload(sLastDLDay, sDownLoadFile, bCheckRes) = False Then
'            Exit Function
'        End If
        If UCase(sExt) = ".EXE" Then
            If ChkDownload_EXE(sDownLoadFile, bCheckRes, sExt) = False Then
                Exit Function
            End If
        Else
            ''�O��_�E�����[�h���t��INI�t�@�C�����Q�Ƃ���
            sLastDLDay = GetIni(INIFILESECTION, sDownLoadFile)
            ''�O��_�E�����[�h���t�ȍ~�ɍX�V����Ă��邩�𔻒�
            If ChkDownload(sLastDLDay, sDownLoadFile, bCheckRes) = False Then
                Exit Function
            End If
        End If
'<<<<< exe�_�E�����[�h�Ή� 2008/07/01 SETsw kubota -------------
    
    End If
    
    If bCheckRes = False Then
        '�_�E�����[�h����K�v���Ȃ��ꍇ�A����I��
        ActDownLoad = True
        Exit Function
    End If
    
    ''���b�Z�[�W�\��
    Call MsgOut(0, "�_�E�����[�h�J�n", NORMAL_MSG)
    
    ''�g���q���o�b�N�A�b�v�p�ɕύX����
    Call ReNameFiles(sDownLoadFile, sExt, sExt & ExtBak)
        
    ''�e�s�o�_�E�����[�h
    iResult = FtpGetFiles(objInet, sDownLoadFile & sExt, bDownloadRes)
        
    ''�_�E�����[�h��t�@�C������
    ''  ���s�F�o�b�N�A�b�v�t�@�C�����A
    ''  �����F�o�b�N�A�b�v�t�@�C���폜
    Call ReNameOrDeleteFiles(sDownLoadFile, bDownloadRes, sExt & ExtBak, sExt)
        
    ''�_�E�����[�h���s�Ȃ�
    If iResult < 0 Then
        Call MsgOut(0, "�_�E�����[�h���s", ERR_DISP)
        Exit Function
    ''�_�E�����[�h�����Ȃ�
    Else
        If UCase(sExt) <> ".EXE" Then
            ''�O��_�E�����[�h���tINI�t�@�C��������
            If SetIni(INIFILESECTION, sDownLoadFile) = False Then
                Call MsgOut(0, "�_�E�����[�h���tINI�t�@�C�������ݎ��s", ERR_DISP)
                Exit Function
            End If
        End If
    End If
        
    ActDownLoad = True

End Function


'////////////////////////////////////////////////////
' @(f)
' �@�\    : �O��_�E�����[�h���t�ȍ~�ɍX�V����Ă��邩�𔻒�
'
' �Ԃ�l  : -1:���s
'          >=0:�擾����
'
' ������  : sLastDLDay - �O��_�E�����[�h���t
' �@�@�@�@  sFileName  - �_�E�����[�h�Ώۃ��W���[����
' �@�@�@�@  bDownFlg   - �_�E�����[�h�v�s�v�t���O   True - �v  False - �s�v
'
' �@�\����: �O��_�E�����[�h���t�ȍ~�ɍX�V����Ă��邩�𔻒�
'///////////////////////////////////////////////////
Function ChkDownload(ByVal sLastDLDay As String _
                   , ByVal sFileName As String _
                   , ByRef bDownFlg As Boolean _
                   ) As Boolean
    
    Dim sSQL As String
    Dim objOraDyn As Object
    
    ChkDownload = False
    bDownFlg = False
    
    sSQL = "       SELECT codea9                    "   ''���[�h���W���[����
    sSQL = sSQL & "FROM   koda9                     "   ''�Ǘ��R�[�h�e�[�u��
    sSQL = sSQL & "WHERE  sysca9 = 'K'              "   ''
    sSQL = sSQL & "AND    shuca9 = '01'             "   ''�o�[�W�������
    sSQL = sSQL & "AND    codea9 = '" & sFileName & "' "
    If Trim$(sLastDLDay) <> "" Then ''���t���w�肳�ꂽ������ɓ����
        sSQL = sSQL & "AND    tdaya9 > TO_DATE(        '" _
                & Format(sLastDLDay, "yyyymmddhhnnss") _
                & "','yyyymmddhh24miss')            "         ''�o�^���t���O��_�E�����[�h���t���V��������
    End If
    
    ''�_�C�i�Z�b�g�쐬
'>>>>> 300mmDynSet2�Ή��@2008/11/21�@SET.Marushita
    'If DynSet(objOraDyn, sSQL) = False Then
    If DynSet2(objOraDyn, sSQL) = False Then
'<<<<< 300mmDynSet2�Ή��@2008/11/21�@SET.Marushita
        ''�_�C�i�Z�b�g�쐬���s
        Call MsgOut(100, sSQL, ERR_DISP_LOG, "kodea9")
        Exit Function
    End If
    
    If objOraDyn.EOF = False Then
        bDownFlg = True
    End If
    
    ChkDownload = True
    
End Function

'////////////////////////////////////////////////////
' @(f)
' �@�\    : �c�a�ƃ��[�J���t�@�C���̃o�[�W�������r���_�E�����[�h�v�s�v�𔻒�
'
' �Ԃ�l  : False - �ُ�
'           True  - ����
'
' ������  : sFileName  - �_�E�����[�h�Ώۃ��W���[����
' �@�@�@�@  bDownFlg   - �_�E�����[�h�v�s�v�t���O   True - �v  False - �s�v
'
' �@�\����:
'///////////////////////////////////////////////////
Function ChkDownload_EXE(ByVal sFileName As String _
                       , ByRef bDownFlg As Boolean _
                       , ByVal sExt As String _
                       ) As Boolean
    
    Dim sSQL        As String
    Dim objOraDyn   As Object
    Dim sMajor      As String
    Dim sMinor      As String
    Dim sRevision   As String
    Dim tKoda9      As typKoda9Data

    bDownFlg = False
    
    '�Ώۃt�@�C���̃��[�J���t�@�C��Ver�擾
    If GetFileVer(sFileName & sExt, sMajor, sMinor, sRevision) = False Then
        Exit Function
    End If
    
    '�Ώۃt�@�C���̂c�a�o�^Ver�擾
    If GetKanriCode("K", "01", sFileName, tKoda9) = False Then
        Exit Function
    End If

    'Ҽެ��AϲŰ�A��޼ޮ݂��r���Ĉ�ł��Ⴄ�ꍇ�A�_�E�����[�h�v
    If val(sMajor) <> val(tKoda9.sCTR01A9) _
    Or val(sMinor) <> val(tKoda9.sCTR02A9) _
    Or val(sRevision) <> val(tKoda9.sCTR03A9) Then
        bDownFlg = True
    End If
    
    ChkDownload_EXE = True
    
End Function



'///////////////////////////////////////////////////
' @(f)
' �@�\    : �e�s�o�_�E�����[�h����
' �Ԃ�l  : -1:�ُ�
'            0:����
' ������  : sFileName - �_�E�����[�h�t�@�C����
' �@�@�@�@  bResult   - �_�E�����[�h����
' �@�\����: �e�s�o�_�E�����[�h����
'///////////////////////////////////////////////////
Private Function FtpGetFiles(ByRef objInet As Inet, ByVal sFileName As String, ByRef bResult As Boolean) As Integer
    Dim sHost     As String ''�z�X�g
    Dim sUserId   As String ''���[�U�[
    Dim sPassword As String ''�p�X���[�h
    Dim sHostPath As String ''�z�X�g���[�h���W���[���p�X
    On Error GoTo Er
    
    bResult = False
    
    Select Case gsFactryCd
    Case "10"               ''��c�H��
        sHost = "CLB0"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case "30"               ''����H��
        sHost = "CLD0"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case "AM"               ''���H��
        sHost = "133.0.0.47"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case "40"               ''�đ�H��
        sHost = "CLE0"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case "42"               ''�R�O�O����
        sHost = "172.20.128.2"
        sUserId = "mqm"
        sPassword = "mqm0001"
        sHostPath = "/home2/cm1/tool/newvb/"
    Case "43"               ''�R�O�O��������
        sHost = "172.20.104.24"
        sUserId = "mqm"
        sPassword = "manager"
        sHostPath = "/home2/cm1/tool/newvb/"
    Case "90"               ''�e�X�g
        sHost = "CLA0"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case "91"               ''�V�e�X�g
        sHost = "172.20.104.24"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case Else               ''�O��
        sHost = "CLB0"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    End Select
    
    With objInet
        .URL = sHost
        .UserName = sUserId
        .Password = sPassword
        .RequestTimeout = FtpTimeOut
            
        Call MsgOut(0, "�e�s�o�擾��" & sFileName, DEBUG_DISP_LOG)
        .Execute , "GET " & sHostPath & sFileName _
                  & " """ & App.Path & "\" & sFileName & """"        ''�e�s�o�擾
        Do While .StillExecuting = True ''�I���҂�
            DoEvents
            If .ResponseCode Then   ''�G���[
                Exit Do ''���[�v�𔲂���
            End If
        Loop
        If .ResponseCode Then   ''�G���[
            bResult = False   ''���s
            '��ʕ\��
            Call MsgOut(0, "�_�E�����[�h���s�F" & sFileName, DEBUG_DISP_LOG)
            '���O�o��
            Call MsgOut(0, "�e�s�o�G���[:�G���[�R�[�h(" & .ResponseCode & ")" & _
                                                        .ResponseInfo, ERR_LOG)
            '���ɃG���[�o�͍ς�
            FtpGetFiles = -1
        Else
            bResult = True    ''����
        End If
        On Error Resume Next
        .Execute , " CLOSE"  '' �ڑ������
    End With
    On Error GoTo 0
    
    Call MsgOut(0, "", NORMAL_MSG)
    Exit Function
Er:
    With objInet
        ''��ʕ\��
        Call MsgOut(0, "�_�E�����[�h���s�F" & sFileName, DEBUG_DISP_LOG)
        ''���O�o��
        Call MsgOut(0, "�e�s�o�G���[:�G���[�R�[�h(" & .ResponseCode & ")" & _
                                                    .ResponseInfo, ERR_LOG)
        FtpGetFiles = -1
    End With
    Resume Next
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �g���q�ύX����
' �Ԃ�l  : -1:�ُ�
'           >0:��������
' ������  : �t�@�C�����z��
'           �ύX�O�g���q
'           �ύX��g���q
' �@�\����: �g���q�ύX����
'///////////////////////////////////////////////////
Private Function ReNameFiles(sFileName, sSExt As String, sDExt As String) As Integer
    On Error GoTo Er
    ''�t�@�C���̑��݃`�F�b�N
    If Dir(App.Path & "\" & sFileName & sSExt) <> "" Then   ''���̃t�@�C�����݂��
        ''�t�@�C�����̊g���q��ύX
        Name App.Path & "\" & sFileName & sSExt As App.Path & "\" & sFileName & sDExt
        ReNameFiles = ReNameFiles + 1
    End If
    Exit Function
Er:
    Call MsgOut(0, "̧�يg���q�ύX���s " & sFileName & sSExt & "��" & sDExt, ERR_DISP_LOG)
    Resume Next
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �t�@�C���폜����
' �Ԃ�l  : -1:�ُ�
'           >0:��������
' ������  : �t�@�C����
' �@�\����: �t�@�C���폜����
'///////////////////////////////////////////////////
Private Function DeleteFiles(sFileName)
    On Error GoTo Er
    ''�t�@�C���̑��݃`�F�b�N
    If Dir(App.Path & "\" & sFileName) <> "" Then    ''���̃t�@�C�����݂��
        ''�폜
        Kill App.Path & "\" & sFileName
        DeleteFiles = DeleteFiles + 1
    End If
    Exit Function
Er:
    Call MsgOut(0, "̧�ٍ폜���s " & sFileName, ERR_DISP_LOG)
    Resume Next
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �_�E�����[�h��t�@�C������
' �Ԃ�l  :
' ������  : �t�@�C�����z��
'           �o�b�N�A�b�v�t�@�C���g���q
'           ���s�t�@�C���g���q
' �@�\����: �_�E�����[�h���ʃt���O�ɂ��o�b�N�A�b�v�t�@�C���폜�^exe���A����
' ���l    :     �����F�o�b�N�A�b�v�t�@�C���폜
'               ���s�F�o�b�N�A�b�v�t�@�C��exe���A
'///////////////////////////////////////////////////
Private Sub ReNameOrDeleteFiles(ByVal sFileName As String _
                              , ByVal bDownloadRes As Boolean _
                              , ByVal sBakExt As String _
                              , ByVal sExeExt As String)
    On Error GoTo Er
    ''�t�@�C���̑��݃`�F�b�N
    If Dir(App.Path & "\" & sFileName & sBakExt) <> "" Then    ''���̃t�@�C�����݂��
        ''�_�E�����[�h�����Ȃ�
        If bDownloadRes = True Then
            ''�t�@�C���폜
            Kill App.Path & "\" & sFileName & sBakExt
        ''�_�E�����[�h���s�Ȃ�
        Else
            On Error Resume Next
            ''���s�����_�E�����[�h�r���̎c�[�t�@�C�����폜
            Kill App.Path & "\" & sFileName & sExeExt
            On Error GoTo 0
            ''�t�@�C�����̊g���q���o�b�N�A�b�v����d�w�d�t�@�C���ɕύX
            Name App.Path & "\" & sFileName & sBakExt As App.Path & "\" & sFileName & sExeExt
        End If
    End If
    Exit Sub
Er:
    Call MsgOut(0, "�޳�۰�ތ�̧�ُ������s " & sFileName, ERR_DISP_LOG)
    Resume Next
End Sub

'///////////////////////////////////////////////////
' @(f)
' �@�\    : �t�@�C���o�[�W�����擾
' �Ԃ�l  : True  - ����
' �@�@�@�@  False - �ُ�
' ������  : sFileName  - Ver�擾�Ώۃt�@�C����
' �@�@�@�@  sMajor     - ���W���[�o�[�W����
' �@�@�@�@  sMinor     - �}�C�i�[�o�[�W����
' �@�@�@�@  sRevision  - �}��
' �@�\����:
' ���l    : 2008/07/01 EXE�_�E�����[�h�Ή�
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetFileVer(ByVal sFileName As String _
                         , ByRef sMajor As String _
                         , ByRef sMinor As String _
                         , ByRef sRevision As String _
                         ) As Boolean
    
    Dim lngSizeOfVersionInfo    As Long
    Dim lngDummyHandle          As Long
    Dim bytDummyVersionInfo()   As Byte
    Dim lngPointerversionInfo   As Long
    Dim lngLengthVersioninfo    As Long
    Dim udtVSFixedFileInfo      As VS_FIXEDFILEINFO
    Dim lngWin32apiResultCode   As Long
    
    lngSizeOfVersionInfo = GetFileVersionInfoSize(sFileName, lngDummyHandle)
    If lngSizeOfVersionInfo > 0 Then
        ReDim bytDummyVersionInfo(lngSizeOfVersionInfo - 1)
        lngWin32apiResultCode = GetFileVersionInfo(sFileName, _
                                                    0, _
                                                    lngSizeOfVersionInfo, _
                                                    bytDummyVersionInfo(0))
        lngWin32apiResultCode = VerQueryValue(bytDummyVersionInfo(0), _
                                            "\", _
                                            lngPointerversionInfo, _
                                            lngLengthVersioninfo)
        
        Call MoveMemory(udtVSFixedFileInfo, lngPointerversionInfo, Len(udtVSFixedFileInfo))
        
        With udtVSFixedFileInfo
            sMajor = (.dwProductVersionMS \ (2 ^ 16)) And &HFFFF&
            sMinor = Format(.dwProductVersionMS And &HFFFF&, "#0")
            sRevision = Format(.dwProductVersionLS)
        End With
    
    Else
        sMajor = ""
        sMinor = ""
        sRevision = ""
        Exit Function
    End If

    GetFileVer = True

End Function




