Attribute VB_Name = "s_cmzcXSDC_WF_Basic"
'===============================================================================
'   �\���̒�`
'===============================================================================
'����{���\����
Public Type type_KIHON      '��{���\����
    STAFFID     As String   '�S���҃R�[�h
    NEWPROC     As String   '���H��
    NOWPROC     As String   '���H��
    DIAMETER    As Long     '���a
    ALLSCRAP    As String   '�S���X�N���b�v�i'Y'�F����A'N'�F�Ȃ��j
    FURYOUMU    As String   '�s�ǗL���i'Y'�F����A'N'�F�Ȃ��j
    CNTHINOLD   As Integer  '���������i�i�ԁF�O�H���j����
    CNTHINNOW   As Integer  '���������i�i�ԁF�Ǖi�j����
End Type

Public Kihon    As type_KIHON          '��{���
Public BlkOld As typ_XSDC2_Update      '��������(�u���b�N)�F�O�H��
Public BlkNow As typ_XSDC2_Update      '��������(�u���b�N)�F�Ǖi
Public HinOld() As typ_XSDCA_Update    '��������(�i��)�F�O�H��
Public HinNow() As typ_XSDCA_Update    '��������(�i��)�F�Ǖi
Public Furyou   As typ_XSDC4_Update    '�s�Ǔ���

Private blkInfo As typ_cmkc001f_Block
Private bMapErrFlg As Boolean           'WFϯ�߈ʒu�����׸�

Private HSXCTCEN As Double      ' �i�r�w�����ʌX�c���S
Private HSXCYCEN As Double      ' �i�r�w�����ʌX�����S

'WF�����v�Z�p�̃p�����[�^
Private SEEDDEG As Integer      ' SEED�X��
Private Loss0 As Integer        ' �X����0�x�̂Ƃ��̌X�����X
Private Loss4 As Integer        ' �X����4�x�̂Ƃ��̌X�����X
Private Mlt4 As Double          ' �X����4�x�̎��̌W��
Private Pitch As Double         ' ���C���\�[���C�����[���s�b�`

'�u���b�N�Ǘ�
Public Type typ_cmkc001f_Block
    'E040 �u���b�N�Ǘ�
    INGOTPOS As Integer         ' �������J�n�ʒu
    LENGTH As Integer           ' ����
    REALLEN As Integer          ' ������
    KRPROCCD As String * 5      ' ���݊Ǘ��H��
    NOWPROC As String * 5       ' ���ݍH��
    LPKRPROCCD As String * 5    ' �ŏI�ʉߊǗ��H��
    LASTPASS As String * 5      ' �ŏI�ʉߍH��
    DELCLS As String * 1        ' �폜�敪
    RSTATCLS As String * 1      ' ������ԋ敪
    LSTATCLS As String * 1      ' �ŏI��ԋ敪 */
    'E037 �������Ǘ�
    SEED As String              'SEED
End Type


'�d�l�擾�p
Public Type typ_cmkc001f_Disp
    '�i�ԊǗ�
    hinban As String * 8              ' �i��
    INGOTPOS As Integer               ' �������J�n�ʒu
    REVNUM As Integer                 ' ���i�ԍ������ԍ�
    factory As String * 1             ' �H��
    opecond As String * 1             ' ���Ə���
    LENGTH As Integer                 ' ����
    '���i�d�lSXL�f�[�^
    HSXD1CEN As Double                ' �i�r�w���a�P���S
    HSXRMIN As Double                 ' �i�r�w���R����
    HSXRMAX As Double                 ' �i�r�w���R���
    HSXRMBNP As Double                ' �i�r�w���R�ʓ����z
    HSXRHWYS As String * 1            ' �i�r�w���R�ۏؕ��@�Q��
    HSXONMIN As Double                ' �i�r�w�_�f�Z�x����
    HSXONMAX As Double                ' �i�r�w�_�f�Z�x���
    HSXONMBP As Double                ' �i�r�w�_�f�Z�x�ʓ����z
    HSXONHWS As String * 1            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    HSXCNMIN As Double                ' �i�r�w�Y�f�Z�x����
    HSXCNMAX As Double                ' �i�r�w�Y�f�Z�x���
    HSXCNHWS As String * 1            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    HSXTMMAX As Double                ' �i�r�w�]�ʖ��x���
    HSXBMnAN(1 To 3) As Double        ' �i�r�w�a�l�cn ���ω���
    HSXBMnAX(1 To 3) As Double        ' �i�r�w�a�l�cn ���Ϗ��
    HSXBMnHS(1 To 3) As String * 1    ' �i�r�w�a�l�cn �ۏؕ��@�Q��
    HSXOFnAX(1 To 4) As Double        ' �i�r�w�n�r�en���Ϗ��
    HSXOFnMX(1 To 4) As Double        ' �i�r�w�n�r�en���
    HSXOFnHS(1 To 4) As String * 1    ' �i�r�w�n�r�en �ۏؕ��@�Q��
    HSXDENMX As Integer               ' �i�r�w�c�������
    HSXDENMN As Integer               ' �i�r�w�c��������
    HSXDENHS As String * 1            ' �i�r�w�c�����ۏؕ��@�Q��
    HSXDVDMX As Integer               ' �i�r�w�c�u�c�Q���
    HSXDVDMN As Integer               ' �i�r�w�c�u�c�Q����
    HSXDVDHS As String * 1            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    HSXLDLMX As Integer               ' �i�r�w�k�^�c�k���
    HSXLDLMN As Integer               ' �i�r�w�k�^�c�k����
    HSXLDLHS As String * 1            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    HSXLTMIN As Integer               ' �i�r�w�k�^�C������
    HSXLTMAX As Integer               ' �i�r�w�k�^�C�����
    HSXLTHWS As String * 1            ' �i�r�w�k�^�C���ۏؕ��@�Q��
    HSXDPDIR As String * 2            ' �i�r�w�a�ʒu����
    HSXDPDRC As String * 1            ' �i�r�w�a�ʒu����
    HSXDWMIN As Double                ' �i�r�w�a�Љ���
    HSXDWMAX As Double                ' �i�r�w�a�Џ��
    HSXDDMIN As Double                ' �i�r�w�a�[����
    HSXDDMAX As Double                ' �i�r�w�a�[���
    HSXD1MIN As Double                ' �i�r�w���a�P����
    HSXD1MAX As Double                ' �i�r�w���a�P���
    HSXCTCEN As Double                ' �i�r�w�����ʌX�c���S
    HSXCYCEN As Double                ' �i�r�w�����ʌX�����S
    EPDUP As Integer                  ' ���������Ǘ� EPD�@���
End Type

'�݌Ɍ����o�^�p
Public Type typ_stock_info
    hinban As String * 8        ' �i��
    GENZAL As Long              ' ���ݒ���
    HARAIL As Long              ' �����o������
    FURYOL As Long              ' �s�ǒ���
    GENZAW As Long              ' ���ݏd��
    HARAIW As Long              '�����o���d��
    FuryoW As Long              ' �s�Ǐd��
    GENZAM As Long              ' ���ݖ���
    HARAIM As Long              ' �����o������
    FURYOM As Long              ' �s�ǖ���
    KCKNT  As Integer           '�H���A��
    REVNUM As Integer           ' ���i�����ԍ�
    factory As String           ' �H��
    OPE As String               ' ���Ə���
End Type

'�s�Ǐ��
Public Type typ_bad_info
    pos As Double              ' �i��
    LEN As Double
End Type

'�i�ԐU�֏��
Public Type typ_trans_info
    hinban As String * 8        ' �i��
    LEN As Long                 ' ����
    WAT As Long                 ' �d��
    MAI As Long                 ' ����
    KCKNT  As Integer           ' �H���A��
    REVNUM As Integer           ' ���i�����ԍ�
    factory As String           ' �H��
    OPE As String               ' ���Ə���
End Type

Public STOCKINFO() As typ_stock_info    'XSDC3Proc2()��XSDC3Proc3()�Ŏg�p
Public giInpos  As Integer
Public strSxlData   As String

'*******************************************************************************
'*    �֐���        : KihonProc
'*
'*    �����T�v      : 1.��ʂ���̊�{�������s��
'*                      �iDB�ւ̓o�^�yXSDC2,XSDC3,XSDC4,XSDCA,XSDCB�z�j
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function KihonProc() As FUNCTION_RETURN
'   �����ϐ�
    Dim i               As Integer
    Dim j               As Integer
    Dim intRtn          As Integer          '���A���
    Dim sSQL            As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim sErrMsg         As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    KihonProc = FUNCTION_RETURN_FAILURE
    'XSDCAProc�AXSDC2Proc�̏��ԓ���ւ�
'    �ᕪ�������i�i�ԁj�o�^��
    intRtn = XSDCAProc()
    If intRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDCAProc()�FXSDCA�o�^�G���["
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
'    �ᕪ�������i�u���b�N�j�o�^��
    intRtn = XSDC2Proc()
    If intRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDC2Proc()�FXSDC2�o�^�G���["
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
'    ��s�Ǔ���o�^��
    '�s�ǗL�������鎞
    If Kihon.FURYOUMU = "Y" Then
                                                ' �o�^���t
        Furyou.TDAYC4 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                                                ' �X�V���t
        Furyou.KDAYC4 = Format(Now(), "YYYY/MM/DD HH:NN:SS")

        '�s�ǒ����E�d�ʁE�������Ď擾-------------
        Furyou.PUCUTLC4 = CLng(BlkOld.GNLC2) - CLng(BlkNow.GNLC2)
        Furyou.PUCUTWC4 = CLng(BlkOld.GNWC2) - CLng(BlkNow.GNWC2)
        Furyou.PUCUTMC4 = CLng(BlkOld.GNMC2) - CLng(BlkNow.GNMC2)

        '�s�Ǔ���ǉ�
        intRtn = CreateXSDC4(Furyou, sErrMsg)

        '�s�Ǔ���ǉ��G���[
        If intRtn = FUNCTION_RETURN_FAILURE Then
            MsgBox sErrMsg
            Debug.Print "CreateXSDC4()�FXSDC4�o�^�G���["
            KihonProc = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    Debug.Print HinNow(0).SXLIDCA

    ' ��H�����ѓo�^��
    intRtn = XSDC3Proc()
    If intRtn = FUNCTION_RETURN_FAILURE Then
        Debug.Print "XSDC3Proc()�FXSDC3�o�^�G���["
        KihonProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    Select Case Kihon.NOWPROC
        Case "CW740", "CW760"
            ' ��݌Ɍ����o�^��
            intRtn = XSDC3Proc2()
            If intRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()�FXSDC3�݌Ɍ����o�^�G���["
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If

            ' ��U�֏��o�^��
            intRtn = XSDC3Proc3()
            If intRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()�FXSDC3�U�֏��o�^�G���["
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        Case "CC730"
            ' ��݌Ɍ����o�^��
            intRtn = XSDC3Proc4()
            If intRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()�FXSDC3�݌Ɍ����o�^�G���["
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If

            ' ��U�֏��o�^��
            intRtn = XSDC3Proc5()
            If intRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()�FXSDC3�U�֏��o�^�G���["
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
    End Select

    Debug.Print HinNow(0).SXLIDCA

    ' �ᕪ�������i�r�w�k�j�o�^��
    intRtn = XSDCBProc()
    If intRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDCBProc()�FXSDCB�o�^�G���["
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
    KihonProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    KihonProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        :
'*
'*    �����T�v      : 1.���������i�u���b�N�j�o�^�������s��(XSDC2)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC2Proc()
    ' �����ϐ�
    Dim i               As Integer
    Dim j               As Integer
    Dim intRtn          As Integer          ' ���A���
    Dim sSQL            As String           ' �r�p�k
    Dim rs              As OraDynaset       ' ���R�[�h�Z�b�g
    Dim sSqlWhere       As String           ' WHERE��
    Dim sErrMsg         As String
    Dim intSyoriKaisu   As Integer          ' ���ݏ�����
    Dim intHantei       As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    XSDC2Proc = FUNCTION_RETURN_FAILURE

    '���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
    Select Case Kihon.NOWPROC
        Case "CC730"
            intHantei = CInt(BlkNow.GNLC2)
        Case Else
            intHantei = CInt(BlkNow.GNMC2)
    End Select

    If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
        BlkNow.LSTATBC2 = "H"                   ' �ŏI��ԋ敪�i�p���j
        BlkNow.LDFRBC2 = "2"                    ' �i���敪�i�n�C�L�j
        BlkNow.LIVKC2 = "1"                     ' �����敪�i�����b�g�j
        BlkNow.GNWKNTC2 = " "                   ' ���ݍH��
        BlkNow.GNMACOC2 = "0"                   ' ���ݏ�����
    Else
        ' �����񐔎擾���W�b�N�ύX
        intSyoriKaisu = GetGNMACOC(BlkNow.CRYNUMC2, BlkNow.GNWKNTC2)
        If BlkNow.GNWKNTC2 = BlkNow.NEWKNTC2 Then
            intSyoriKaisu = intSyoriKaisu + 1
        End If
        BlkNow.GNMACOC2 = intSyoriKaisu                                 ' ���ݏ�����
        BlkNow.NEMACOC2 = GetGNMACOC(BlkNow.CRYNUMC2, BlkNow.NEWKNTC2)  ' �ŏI�ʉߏ�����
    End If


    BlkNow.KDAYC2 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
    BlkNow.PLANTCATC2 = sCmbMukesaki
    sSqlWhere = "WHERE CRYNUMC2 = '" & BlkNow.CRYNUMC2 & "' "

    intRtn = UpdateXSDC2(BlkNow, sSqlWhere)
    '���������i�u���b�N�j�X�V�G���[
    If intRtn = FUNCTION_RETURN_FAILURE Then
        MsgBox "XSDCB UPDATET ERROR"
        Exit Function
    End If

    XSDC2Proc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC2Proc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : XSDCAProc
'*
'*    �����T�v      : 1.���������i�i�ԁj�o�^�������s��(XSDCA)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDCAProc()
    ' �����ϐ�
    Dim i               As Integer
    Dim j               As Integer
    Dim intRtn          As Integer          ' ���A���
    Dim sSQL            As String           ' �r�p�k
    Dim rs              As OraDynaset       ' ���R�[�h�Z�b�g
    Dim sSqlWhere       As String           ' WHERE��
    Dim sErrMsg         As String
    Dim intLivFlg       As Integer          ' ���݃t���O
    Dim udtHinban()     As typ_XSDCA        ' ���������i�i�ԁj���[�N�̈�
    Dim udtHinban_UP()  As typ_XSDCA_Update ' ���������i�i�ԁj���[�N�̈�
    Dim intDataCnt      As Integer          ' �Y���f�[�^����
    Dim lngSumGNWCA     As Long
    Dim lngSumGNMCA     As Long
    Dim blChgFlg        As Boolean
    Dim intSyoriKaisu   As Integer          ' ���ݏ�����
    Dim intHantei       As Integer
    Dim lngGetLength    As Long             ' TBCME040���A�u���b�N�������擾����

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    XSDCAProc = FUNCTION_RETURN_FAILURE
    lngSumGNWCA = 0
    lngSumGNMCA = 0
    blChgFlg = False

    ' �i�Ԃ̏d�ʁE�����v�Z
    If Kihon.CNTHINNOW = Kihon.CNTHINOLD Then   ' �O�H���ƌ��ݍH���̌����������ŁA�e�����������ꍇ�͌v�Z���������Ȃ�
        For i = 0 To Kihon.CNTHINNOW - 1
            If CLng(HinNow(i).GNLCA) <> CLng(HinOld(i).GNLCA) Then
                blChgFlg = True
            End If
        Next
    Else            ' �O�H���ƌ��ݍH���̌������Ⴄ�ꍇ�́A�v�Z�������s��
        blChgFlg = True
    End If
    ' �d�ʁE�����v�Z����

    If blChgFlg = True Then
        ' CW740,CW760�p�ǉ�
        If Kihon.NOWPROC = "CW740" Or Kihon.NOWPROC = "CW760" Then
            For i = 0 To Kihon.CNTHINNOW - 1
                With HinNow(i)  ' BLKOLD��ɕύX
                    If Kihon.CNTHINNOW = 1 Then
                        HinNow(i).GNWCA = BlkOld.GNWC2
                        HinNow(i).GNLCA = BlkOld.GNLC2
                        .SUMITLCA = .GNLCA
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    ElseIf i = Kihon.CNTHINNOW - 1 Then
                        HinNow(i).GNWCA = CLng(BlkOld.GNWC2) - lngSumGNWCA
                        HinNow(i).GNLCA = CLng(BlkOld.GNLC2) - lngSumGNLCA
                        'Add Start 2010/10/14 Y.Hitomi
                        If HinNow(i).GNLCA <= 0 And HinNow(i).GNMCA = 1 Then
                            HinNow(i).GNLCA = 1
                        End If
                        'Add End 2010/10/14 Y.Hitomi
                        .SUMITLCA = .GNLCA
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    Else
                        HinNow(i).GNWCA = Round(CLng(BlkOld.GNWC2) * (CLng(HinNow(i).GNMCA) / CLng(BlkOld.GNMC2)))
                        HinNow(i).GNLCA = Round(CLng(BlkOld.GNLC2) * (CLng(HinNow(i).GNMCA) / CLng(BlkOld.GNMC2)))
                        'Add Start 2010/10/14 Y.Hitomi
                        If HinNow(i).GNLCA <= 0 And HinNow(i).GNMCA = 1 Then
                            HinNow(i).GNLCA = 1
                        End If
                        'Add End 2010/10/14 Y.Hitomi
                        lngSumGNWCA = lngSumGNWCA + CLng(HinNow(i).GNWCA)
                        lngSumGNLCA = lngSumGNLCA + CLng(HinNow(i).GNLCA)
                        .SUMITLCA = .GNLCA
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    End If
                End With
            Next
        Else
            If BlkOld.GNLC2 = BlkNow.GNLC2 Then ' �����������ꍇ��BLKOLD��B�������قȂ�ꍇ��BLKNOW��ɂ���
                BlkNow.GNWC2 = BlkOld.GNWC2
                BlkNow.GNMC2 = BlkOld.GNMC2
            Else
                BlkNow.GNWC2 = Round(CLng(BlkOld.GNWC2) * (CLng(BlkNow.GNLC2) / CLng(BlkOld.GNLC2)))
                BlkNow.GNMC2 = Round(CLng(BlkOld.GNMC2) * (CLng(BlkNow.GNLC2) / CLng(BlkOld.GNLC2)))
            End If
            For i = 0 To Kihon.CNTHINNOW - 1    ' BLKOLD��ɕύX
                With HinNow(i)
                    If Kihon.CNTHINNOW = 1 Then
                        HinNow(i).GNWCA = BlkNow.GNWC2
                        HinNow(i).GNMCA = BlkNow.GNMC2
                    ElseIf i = Kihon.CNTHINNOW - 1 Then
                        HinNow(i).GNWCA = CLng(BlkNow.GNWC2) - lngSumGNWCA
                        HinNow(i).GNMCA = CLng(BlkNow.GNMC2) - lngSumGNMCA
                    Else
                        HinNow(i).GNWCA = Round(CLng(BlkNow.GNWC2) * (CLng(HinNow(i).GNLCA) / CLng(BlkNow.GNLC2)))
                        HinNow(i).GNMCA = Round(CLng(BlkNow.GNMC2) * (CLng(HinNow(i).GNLCA) / CLng(BlkNow.GNLC2)))
                        lngSumGNWCA = lngSumGNWCA + CLng(HinNow(i).GNWCA)
                        lngSumGNMCA = lngSumGNMCA + CLng(HinNow(i).GNMCA)
                    End If
                End With
            Next
        End If
    End If

    ' XSDC2�̏d�ʁA������XSDCA�̍��v�ɂ��邽�߂�BlkNow�Čv�Z
    BlkNow.GNLC2 = 0
    BlkNow.GNWC2 = 0
    For i = 0 To Kihon.CNTHINNOW - 1
        BlkNow.GNLC2 = CLng(BlkNow.GNLC2) + CLng(HinNow(i).GNLCA)   '2003/05/24 clng�ǉ�
        BlkNow.GNWC2 = CLng(BlkNow.GNWC2) + CLng(HinNow(i).GNWCA)
    Next i

    ' �O�H���̕��������i�i�ԁj�ƗǕi���̕i�ԁE�ʒu���r����
    For i = 0 To Kihon.CNTHINOLD - 1

        intLivFlg = 0
        For j = 0 To Kihon.CNTHINNOW - 1
            If (HinOld(i).HINBCA = HinNow(j).HINBCA) And (HinOld(i).INPOSCA = HinNow(j).INPOSCA) Then
                intLivFlg = 1
            End If
        Next j

        ' �O�H���̕��������i�i�ԁj�ɂ����ėǕi���ɂȂ����͎̂����b�g�Ƃ���
        If intLivFlg = 0 Then
            sSqlWhere = "WHERE CRYNUMCA = '" & HinOld(i).CRYNUMCA & "' "
            sSqlWhere = sSqlWhere & "AND HINBCA = '" & HinOld(i).HINBCA & "' "
            sSqlWhere = sSqlWhere & "AND INPOSCA = '" & HinOld(i).INPOSCA & "' "
            ReDim udtHinban(0) As typ_XSDCA

            ' �f�[�^�̌������擾
            intRtn = SelCntXSDCA(sSqlWhere, intDataCnt)
            If intRtn = FUNCTION_RETURN_FAILURE Then    ' �G���[
                MsgBox "XSDCA SELECT ERROR"
                Exit Function
            Else                                        ' ����
                If intDataCnt = 0 Then
                    ' �O�H���̏��͕K������͂��Ȃ̂ŁA0���̓G���[
                    Exit Function
                ElseIf intDataCnt > 0 Then
                    ' �O�H��
                    intRtn = DBDRV_GetXSDCA(udtHinban(), sSqlWhere)

                    ' ���݂��Ȃ����G���[
                    If intRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox "XSDCA SELECT ERROR"
                        Exit Function
                    End If
                    ReDim udtHinban_UP(0) As typ_XSDCA_Update

                    udtHinban_UP(0).CRYNUMCA = HinOld(i).CRYNUMCA
                    udtHinban_UP(0).INPOSCA = HinOld(i).INPOSCA
                    udtHinban_UP(0).HINBCA = HinOld(i).HINBCA

                    ' �����敪�Ɏ����b�g���Z�b�g
                    udtHinban_UP(0).LIVKCA = "1"              ' �����敪
                    udtHinban_UP(0).LSTATBCA = "H"            ' �ŏI��ԋ敪
                    udtHinban_UP(0).LDFRBCA = "2"             ' �i���敪
                    udtHinban_UP(0).KANKCA = "0"              ' �����敪
                    udtHinban_UP(0).SUMITBCA = "0"
                    udtHinban_UP(0).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
                    udtHinban_UP(0).PLANTCATCA = sCmbMukesaki

                    '���������i�i�ԁj���X�V
                    intRtn = UpdateXSDCA(udtHinban_UP(0), sSqlWhere)

                    '���������i�i�ԁj�X�V�G���[
                    If intRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox "XSDCA UPDATE ERROR"
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    ' ���������i�i�ԁj���J��Ԃ�
    For i = 0 To Kihon.CNTHINNOW - 1
        ' �����ԍ��A�i�ԁA�ʒu�Ō���
        sSqlWhere = "WHERE CRYNUMCA = '" & HinNow(i).CRYNUMCA & "' "
        sSqlWhere = sSqlWhere & "AND HINBCA = '" & HinNow(i).HINBCA & "' "
        sSqlWhere = sSqlWhere & "AND INPOSCA = '" & HinNow(i).INPOSCA & "' "

        ' �f�[�^�̌������擾
        intRtn = SelCntXSDCA(sSqlWhere, intDataCnt)
        If intRtn = FUNCTION_RETURN_FAILURE Then    ' �G���[
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        Else                                        ' ����
            ' �f�[�^������ꍇ�͍X�V����
            If intDataCnt > 0 Then
                intRtn = DBDRV_GetXSDCA(udtHinban, sSqlWhere)
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If

                ' ���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
                Select Case Kihon.NOWPROC
                    Case "CC730"
                        intHantei = CInt(BlkNow.GNLC2)
                    Case Else
                        intHantei = CInt(HinNow(i).GNMCA) ' 0����p���ɂ��鏈�����A�u���b�N�P�ʂł͂Ȃ��i�ԒP�ʂɕύX
                End Select

                If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                    HinNow(i).LIVKCA = "1"              ' �����敪
                    HinNow(i).LSTATBCA = "H"            ' �ŏI��ԋ敪
                    HinNow(i).LDFRBCA = "2"             ' �i���敪
                    HinNow(i).GNWKNTCA = " "            ' ���ݍH��
                    HinNow(i).GNMACOCA = "0"            ' ���ݏ�����
                Else
                    HinNow(i).LIVKCA = "0"              ' �����敪�i�����b�g�j
                    HinNow(i).LSTATBCA = "T"            ' �ŏI��ԋ敪�i�ʏ�j
                    HinNow(i).LDFRBCA = "0"             ' �i���敪�i�ʏ�j

                    ' �����񐔎擾���W�b�N�ύX
                    intSyoriKaisu = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).GNWKNTCA)
                    If HinNow(i).GNWKNTCA = HinNow(i).NEWKNTCA Then
                          intSyoriKaisu = intSyoriKaisu + 1
                    End If
                    HinNow(i).GNMACOCA = intSyoriKaisu    '���ݏ�����
                    HinNow(i).NEMACOCA = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).NEWKNTCA)   '�ŏI�ʉߏ�����
                End If
                ' �����敪�t���O�ύX
                HinNow(i).KANKCA = "0"              ' �����敪
                HinNow(i).SUMITBCA = "0"
                HinNow(i).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
                HinNow(i).PLANTCATCA = sCmbMukesaki

                ' �Ǖi���Œu������
                intRtn = UpdateXSDCA(HinNow(i), sSqlWhere)

                ' ���������i�i�ԁj�X�V�G���[
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCA UPDATE ERROR"
                    Exit Function
                End If
            ' ���݂��Ȃ����ǉ�
            ElseIf intDataCnt = 0 Then
                '���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
                Select Case Kihon.NOWPROC
                    Case "CC730"
                        intHantei = CInt(BlkNow.GNLC2)
                    Case Else
                        intHantei = CInt(HinNow(i).GNMCA) ' 0����p���ɂ��鏈�����A�u���b�N�P�ʂł͂Ȃ��i�ԒP�ʂɕύX
                End Select

                If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                    HinNow(i).LIVKCA = "1"              ' �����敪
                    HinNow(i).LSTATBCA = "H"            ' �ŏI��ԋ敪
                    HinNow(i).LDFRBCA = "2"             ' �i���敪
                    HinNow(i).GNWKNTCA = " "            ' ���ݍH��
                    HinNow(i).GNMACOCA = "0"            ' ���ݏ�����
                Else
                    HinNow(i).LIVKCA = "0"              ' �����敪�i�����b�g�j
                    HinNow(i).LSTATBCA = "T"            ' �ŏI��ԋ敪�i�ʏ�j
                    HinNow(i).LDFRBCA = "0"             ' �i���敪�i�ʏ�j

                    ' �����񐔎擾���W�b�N�ύX
                    intSyoriKaisu = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).GNWKNTCA)
                    If HinNow(i).GNWKNTCA = HinNow(i).NEWKNTCA Then
                          intSyoriKaisu = intSyoriKaisu + 1
                    End If
                    HinNow(i).GNMACOCA = intSyoriKaisu  ' ���ݏ�����
                    HinNow(i).NEMACOCA = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).NEWKNTCA)   '�ŏI�ʉߏ�����
                End If

                ' �����敪�t���O�ύX
                HinNow(i).KANKCA = "0"                  ' �����敪
                                                        ' �o�^���t
                HinNow(i).TDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                HinNow(i).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
                HinNow(i).SUMITBCA = "0"
                HinNow(i).PLANTCATCA = sCmbMukesaki

                intRtn = CreateXSDCA(HinNow(i), sErrMsg)

                '���������i�i�ԁj�X�V�G���[
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox sErrMsg
                    Exit Function
                End If
            End If
        End If
    Next i

    XSDCAProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDCAProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : XSDC3Proc
'*
'*    �����T�v      : 1.�H�����ѓo�^�������s��(XSDC3)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC3Proc()
    ' �����ϐ�
    Dim i               As Integer
    Dim j               As Integer
    Dim intRtn          As Integer          ' ���A���
    Dim sSQL            As String           ' �r�p�k
    Dim rs              As OraDynaset       ' ���R�[�h�Z�b�g
    Dim sSqlWhere       As String           ' WHERE��
    Dim lngFULC3        As Long             ' ���������i�i�ԁj�̕s�ǒ���
    Dim lngFUWC3        As Long             ' ���������i�i�ԁj�̕s�Ǐd��
    Dim lngFUMC3        As Long             ' ���������i�i�ԁj�̕s�ǖ���
    Dim sErrMsg         As String
    Dim udtKoutei       As typ_XSDC3_Update ' �H������
    Dim rsKCNTC         As OraDynaset       ' ���R�[�h�Z�b�g
    Dim intNextCnt      As Integer
    Dim intOldCnt       As Integer
    Dim blNewRec        As Boolean          ' �O�H���̖������R�[�h���������ꍇ
    Dim sSUMITLC3       As String           ' SUMIT����
    Dim sSUMITWC3       As String           ' SUMIT�d��
    Dim sSUMITMC3       As String           ' SUMIT����
    Dim dtmSumcoTime    As Date             ' SUMCO����
    Dim vChoseiTime     As Variant          ' ��������
    Dim intLoopCnt      As Integer
    Dim intHantei       As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    blNewRec = False                        ' �t���O������
    XSDC3Proc = FUNCTION_RETURN_FAILURE

    ' �H�����т���u���b�N�h�c�A�i�Ԃ���v����H���A�Ԃ̍ő���擾
    sSQL = "SELECT MAX(KCNTC3) as wKCNTC3 "
    sSQL = sSQL & " FROM XSDC3 "
    sSQL = sSQL & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' ���݂��Ȃ����A�G���[
    If rs Is Nothing Then
        MsgBox "XSDC3 MAX KCNT SELECT ERROR"
        Exit Function
    End If

    If rs.EOF = False Then
        If IsNull(rs.Fields("wKCNTC3")) = True Then
            intNextCnt = 1
        Else
            intNextCnt = CInt(rs.Fields("wKCNTC3")) + 1
        End If
    Else
        intNextCnt = 1
    End If
    rs.Close

    ' �O�H�����т̖������R�[�h�����邩�`�F�b�N���A��������t���O�����Ă�
    For i = 0 To Kihon.CNTHINNOW - 1
        ' �H�����т���O�H���̃f�[�^��ǂݍ���
        sSQL = "SELECT STATIMEC3, STOTIMEC3 , TOLC3,TOWC3,TOMC3,WKKTC3,MACOC3 "
        sSQL = sSQL & " FROM XSDC3 "
        sSQL = sSQL & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sSQL = sSQL & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
        sSQL = sSQL & " AND KCNTC3 = " & intNextCnt - 1 & ""  ' intOldCnt�͎g���Ȃ��̂ŁAintNextCnt - 1��ς��Ɏg�p

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' ���݂��Ȃ����A�G���[
        If rs Is Nothing Then
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        End If
        If rs.RecordCount = 0 Then
            blNewRec = True  ' �O�H�����Ȃ��ꍇ�̓t���O�����Ă�
        End If
        rs.Close
    Next

    If Kihon.NOWPROC = "CC730" Then
        blNewRec = True
    End If

    For i = 0 To Kihon.CNTHINNOW - 1
        ' �s�Ǔ��󂩂�u���b�N�h�c�A�i�ԁA�J�n�ʒu����v����s�ǒ������擾����
        intOldCnt = 0
        lngFULC3 = 0
        lngFUWC3 = 0
        lngFUMC3 = 0

        ' �H�����т���u���b�N�h�c�A�i�Ԃ���v����H���A�Ԃ̍ő���擾
        sSQL = "SELECT MAX(KCNTC3) as wKCNTC3 "
        sSQL = sSQL & " FROM XSDC3 "
        sSQL = sSQL & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sSQL = sSQL & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' ���݂��Ȃ����A�G���[
        If rs Is Nothing Then
            MsgBox "XSDC3 MAX KCNT SELECT ERROR"
            Exit Function
        End If

        If rs.EOF = False Then
            If IsNull(rs.Fields("wKCNTC3")) = True Then
                intOldCnt = 0
            Else
                intOldCnt = CInt(rs.Fields("wKCNTC3"))
            End If
        Else
            intOldCnt = 0
        End If
        rs.Close

        ' �H�����т���O�H���̂̃f�[�^��ǂݍ���
        sSQL = "SELECT STATIMEC3, STOTIMEC3 , TOLC3,TOWC3,TOMC3, "
        sSQL = sSQL & " SUMITLC3, SUMITWC3, SUMITMC3"
        sSQL = sSQL & " FROM XSDC3 "
        sSQL = sSQL & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sSQL = sSQL & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
        sSQL = sSQL & " AND KCNTC3 = " & intOldCnt & ""

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' ���݂��Ȃ����A�G���[
        If rs Is Nothing Then
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        End If
        If rs.RecordCount = 0 Then
            lngFULC3 = 0
            lngFUWC3 = 0
            lngFUMC3 = 0
            sSUMITLC3 = "0" 'SUMIT����
            sSUMITWC3 = "0" 'SUMIT�d��
            sSUMITMC3 = "0" 'SUMIT����
        Else
            If IsNull(rs.Fields("STATIMEC3")) = True Then
                ' ��������Ȃ�
            Else
                udtKoutei.STATIMEC3 = rs.Fields("STATIMEC3")
            End If
                                                    ' �������ԏI��
            If IsNull(rs.Fields("STOTIMEC3")) = True Then
                ' ��������Ȃ�
            Else
                udtKoutei.STOTIMEC3 = rs.Fields("STOTIMEC3")
            End If

            If IsNull(rs.Fields("TOLC3")) = True Then   ' �s�ǒ���
                lngFULC3 = 0
                udtKoutei.FRLC3 = "0"
            Else
                udtKoutei.FRLC3 = CLng(rs.Fields("TOLC3"))
                lngFULC3 = CLng(rs.Fields("TOLC3"))
            End If
            If IsNull(rs.Fields("TOWC3")) = True Then   ' �s�Ǐd��
                lngFUWC3 = 0
                udtKoutei.FRWC3 = "0"
            Else
                udtKoutei.FRWC3 = CLng((rs.Fields("TOWC3")))
                lngFUWC3 = CLng((rs.Fields("TOWC3")))
            End If
            If IsNull(rs.Fields("TOMC3")) = True Then   ' �s�ǖ���
                lngFUMC3 = 0
                udtKoutei.FRMC3 = "0"
            Else
                udtKoutei.FRMC3 = CLng((rs.Fields("TOMC3")))
                lngFUMC3 = CLng((rs.Fields("TOMC3")))
            End If
            If IsNull(rs.Fields("SUMITLC3")) = True Then   ' SUMIT����
                sSUMITLC3 = "0"
            Else
                sSUMITLC3 = CLng((rs.Fields("SUMITLC3")))
            End If
            If IsNull(rs.Fields("SUMITWC3")) = True Then   ' SUMIT�d��
                sSUMITWC3 = "0"
            Else
                sSUMITWC3 = CLng((rs.Fields("SUMITWC3")))
            End If
            If IsNull(rs.Fields("SUMITMC3")) = True Then   ' SUMIT����
                sSUMITMC3 = "0"
            Else
                sSUMITMC3 = CLng((rs.Fields("SUMITMC3")))
            End If
        End If

        ' �����񐔎擾���W�b�N�ύX
        If IsNull(HinOld(0).NEWKNTCA) = True Then          ' (����j�H��
            udtKoutei.FRWKKTC3 = "0"
        Else
            udtKoutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
        End If

        If IsNull(HinOld(0).NEMACOCA) = True Then          '�i����j������
            udtKoutei.FRMACOC3 = "0"
        Else
            udtKoutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
        End If

        ' ���������i�i�ԁj����H�����т�ǉ�
        udtKoutei.CRYNUMC3 = HinNow(i).CRYNUMCA            ' ��ۯ�ID������ԍ�
        udtKoutei.INPOSC3 = HinNow(i).INPOSCA              ' �������J�n�ʒu
        udtKoutei.KCNTC3 = intNextCnt                      ' �H���A��
        udtKoutei.HINBC3 = HinNow(i).HINBCA                ' �i��
        udtKoutei.REVNUMC3 = HinNow(i).REVNUMCA            ' ���i�ԍ������ԍ�
        udtKoutei.FACTORYC3 = HinNow(i).FACTORYCA          ' �H��
        udtKoutei.OPEC3 = HinNow(i).OPECA                  ' ���Ə���
        udtKoutei.PLANTCATC3 = HinNow(i).PLANTCATCA        ' ����

        '���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
        Select Case Kihon.NOWPROC
            Case "CC730"
                intHantei = CInt(BlkNow.GNLC2)
            Case Else
                intHantei = CInt(BlkNow.GNMC2)
        End Select

        If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
            udtKoutei.LENC3 = 0                            ' ����
        Else
            udtKoutei.LENC3 = HinNow(i).GNLCA              ' ����
        End If

        udtKoutei.XTALC3 = HinNow(i).XTALCA                ' �����ԍ�
        udtKoutei.SXLIDC3 = HinNow(i).SXLIDCA              ' SXLID

        Select Case Kihon.NOWPROC                          ' CW740�CCW760�H���ŁA�Ǘ��H���Ɍ��ݍH���{�R����������
            Case "CW740", "CW760", "CC730"
                udtKoutei.KNKTC3 = left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
                      CStr(CInt(Right(Kihon.NOWPROC, 1)) + 3)   ' �Ǘ��H��(���ݍH��+3)
            Case Else
                udtKoutei.KNKTC3 = HinNow(i).GNKKNTCA           ' �Ǘ��H��
        End Select
        udtKoutei.WKKTC3 = Kihon.NOWPROC                   ' �H��
        udtKoutei.WKKBC3 = HinNow(i).GNWKKBCA              ' ��Ƌ敪
        udtKoutei.MACOC3 = HinNow(i).NEMACOCA              ' ������
        udtKoutei.MODKBC3 = ""                             ' �ԍ��敪
        udtKoutei.SUMKBC3 = "0"                            ' �W�v�敪
        udtKoutei.FRKNKTC3 = " "                           ' (���)�Ǘ��H��
        udtKoutei.FRWKKBC3 = " "                           ' (���)��Ƌ敪
        udtKoutei.TOWNKTC3 = " "                           ' (���o)�Ǘ��H��

        ' ���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
        Select Case Kihon.NOWPROC
            Case "CC730"
                intHantei = CInt(BlkNow.GNLC2)
            Case Else
                intHantei = CInt(BlkNow.GNMC2)
        End Select

        If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
            udtKoutei.TOWKKTC3 = " "                       ' (���o)�H��
            udtKoutei.TOMACOC3 = "0"                       ' (���o)������
        Else
            udtKoutei.TOWKKTC3 = HinNow(i).GNWKNTCA        ' (���o)�H��
        End If

        udtKoutei.TOMACOC3 = HinNow(i).GNMACOCA            ' (���o)������
        udtKoutei.LOSWC3 = ""                              ' ���X����
        udtKoutei.LOSLC3 = ""                              ' ���X�d��
        udtKoutei.LOSMC3 = ""                              ' ���X����

        If blNewRec = True Then                            ' �O�H���������f�[�^�����݂��Ă���ꍇ�́A���o�����ʂ�������ʂɓ����
            udtKoutei.FRLC3 = HinNow(i).GNLCA              ' �������<=���o����
            udtKoutei.FRWC3 = HinNow(i).GNWCA              ' ����d��<=���o�d��
            udtKoutei.FRMC3 = HinNow(i).GNMCA              ' �������<=���o����
            udtKoutei.TOLC3 = HinNow(i).GNLCA              ' ���o����
            udtKoutei.TOWC3 = HinNow(i).GNWCA              ' ���o�d�ʁi�֐��j
            udtKoutei.TOMC3 = HinNow(i).GNMCA              ' ���o�����i�֐��j
            udtKoutei.FULC3 = 0                            ' �s�ǒ���
            udtKoutei.FUWC3 = 0                            ' �s�Ǐd��
            udtKoutei.FUMC3 = 0                            ' �s�ǖ���
        Else
            Select Case Kihon.NOWPROC
                Case "CC730"
                    intHantei = CInt(BlkNow.GNLC2)
                Case Else
                    intHantei = CInt(BlkNow.GNMC2)
            End Select

            If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                udtKoutei.TOLC3 = 0                             ' ���o����
                udtKoutei.TOWC3 = 0                             ' ���o�d�ʁi�֐��j
                udtKoutei.TOMC3 = 0                             ' ���o�����i�֐��j
            Else
                udtKoutei.TOLC3 = HinNow(i).GNLCA               ' ���o����
                udtKoutei.TOWC3 = HinNow(i).GNWCA               ' ���o�d�ʁi�֐��j
                udtKoutei.TOMC3 = HinNow(i).GNMCA               ' ���o�����i�֐��j
            End If
            udtKoutei.FULC3 = lngFULC3 - CLng(udtKoutei.TOLC3)  ' �s�ǒ���
            udtKoutei.FUWC3 = lngFUWC3 - CLng(udtKoutei.TOWC3)  ' �s�Ǐd��
            udtKoutei.FUMC3 = lngFUMC3 - CLng(udtKoutei.TOMC3)  ' �s�ǖ���
        End If
        If udtKoutei.TOLC3 = "" Then
            udtKoutei.TOLC3 = "0"
        End If
        If udtKoutei.TOWC3 = "" Then
            udtKoutei.TOWC3 = "0"
        End If
        If udtKoutei.TOMC3 = "" Then
            udtKoutei.TOMC3 = "0"
        End If
        ' SUMIT�����ɍH���ʂɒl���Z�b�g����--------------------
        udtKoutei.SUMITLC3 = 0                                  ' SUMIT����
        udtKoutei.SUMITWC3 = 0                                  ' SUMIT�d��
        udtKoutei.SUMITMC3 = 0                                  ' SUMIT����

        For intLoopCnt = 0 To Kihon.CNTHINOLD - 1
            If (udtKoutei.CRYNUMC3 = HinOld(intLoopCnt).CRYNUMCA) _
                And (udtKoutei.INPOSC3 = HinOld(intLoopCnt).INPOSCA) Then
                    udtKoutei.SUMITLC3 = HinOld(intLoopCnt).SUMITLCA     ' SUMIT����=�O�H��SUMIT����
                    udtKoutei.SUMITWC3 = HinOld(intLoopCnt).SUMITWCA     ' SUMIT�d��=�O�H��SUMIT�d��
                    udtKoutei.SUMITMC3 = HinOld(intLoopCnt).SUMITMCA     ' SUMIT����=�O�H��SUMIT����
                    Exit For
            End If
        Next
        udtKoutei.MOTHINC3 = " "                                ' �U�֕i��(��)
        udtKoutei.XTWORKC3 = "42"                               ' �����H��
        udtKoutei.WFWORKC3 = " "                                ' ���ʐ���
        udtKoutei.HOLDCC3 = " "                                 ' �z�[���h�R�[�h
        udtKoutei.HOLDBC3 = "0"                                 ' �z�[���h�敪
        udtKoutei.LDFRCC3 = " "                                 ' �i���R�[�h

        '���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
        Select Case Kihon.NOWPROC
            Case "CC730"
                intHantei = CInt(BlkNow.GNLC2)
            Case Else
                intHantei = CInt(BlkNow.GNMC2)
        End Select

        If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
            udtKoutei.LDFRBC3 = "2"                             ' �i���敪�i�n�C�L�j
        Else
            udtKoutei.LDFRBC3 = "0"                             ' �i���敪
        End If
        udtKoutei.TSTAFFC3 = Kihon.STAFFID                      ' �o�^�Ј�ID

        udtKoutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �o�^���t
        udtKoutei.KSTAFFC3 = Kihon.STAFFID                      ' �X�V�Ј�ID
        udtKoutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
        udtKoutei.SUMDAYC3 = CalcSumcoTime(udtKoutei.KDAYC3)
        udtKoutei.SUMITBC3 = "0"                                ' SUMIT���M�t���O
        udtKoutei.SNDKC3 = "0"                                  ' ���M�t���O
        udtKoutei.MODMACOC3 = "00"                              ' �ԍ��̏�����
        udtKoutei.KAKUCC3 = " "                                 ' �m��R�[�h

        ' ��ʂŎg�p���Ă���f�[�^�̂ݍX�V���s��
        udtKoutei.PLANTCATC3 = sCmbMukesaki

        Select Case Kihon.NOWPROC
            Case "CW750"                                        ' ��������
                If udtKoutei.SXLIDC3 = Trim(f_cmbc039_2.txtSxlID.text) Then
                    intRtn = CreateXSDC3(udtKoutei, sErrMsg)
                    ' �H�����ђǉ��G���[
                    If intRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox sErrMsg
                        Exit Function
                    End If
                End If
            Case "CW760"    ' �Ĕ���
                If (SIngotP <= udtKoutei.INPOSC3) And (udtKoutei.INPOSC3 < EIngotP) Then
                    intRtn = CreateXSDC3(udtKoutei, sErrMsg)
                    ' �H�����ђǉ��G���[
                    If intRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox sErrMsg
                        Exit Function
                    End If
                End If
            Case "CW800"    ' �V���O���m��
                If udtKoutei.SXLIDC3 = strSxlData Then
                    intRtn = CreateXSDC3(udtKoutei, sErrMsg)

                    ' �H�����ђǉ��G���[
                    If intRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox sErrMsg
                        Exit Function
                    End If
                End If
            Case Else
                intRtn = CreateXSDC3(udtKoutei, sErrMsg)

                ' �H�����ђǉ��G���[
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox sErrMsg
                    Exit Function
                End If
        End Select
    Next i

    XSDC3Proc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : XSDCBProc
'*
'*    �����T�v      : 1.���������iudtSXL�j�o�^�������s��(XSDCB)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDCBProc()
    ' �����ϐ�
    Dim i               As Integer
    Dim intRtn          As Integer          ' ���A���
    Dim sSQL            As String           ' �r�p�k
    Dim rs              As OraDynaset       ' ���R�[�h�Z�b�g
    Dim sSqlWhere       As String           ' WHERE��
    Dim lngGNLCA        As Long             ' ���������i�i�ԁj�̍��v����
    Dim lngGNMCA        As Long             ' ���������i�i�ԁj�̍��v����
    Dim sErrMsg         As String
    Dim udtSXL()        As typ_XSDCB_Update ' ��������(udtSXL)
    Dim udtWSXL()       As typ_XSDCB        ' ��������(udtSXL)
    Dim intDataCnt      As Integer          ' �Y���f�[�^����
    Dim sBlockId        As String
    Dim lngRYOMAI       As Long             ' �H�����̗Ǖi����
    Dim lngFRYMAI       As Long             ' �H�����̕s�Ǖi����
    Dim lngLen          As Long             ' ����
    Dim lngMAI          As Long             ' ����
    Dim lngMAI800       As Long             ' CW800����
    Dim lngFUR          As Long             ' �s�ǖ���
    Dim lngFURKEI       As Long             ' �s�ǖ������v
    Dim lngSAM          As Long             ' �T���v������
    Dim lngSIJ          As Long             ' �T���v�����w������
    Dim lngSAMFUR       As Long             ' �T���v�����w���s�ǖ���
    Dim intLoopBkHinGet As Integer          ' ���i��
    Dim m               As Integer          ' �J�E���^

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    XSDCBProc = FUNCTION_RETURN_FAILURE

    For i = 0 To Kihon.CNTHINNOW - 1
        ' �H�����т��瓯���r�w�k�h�c�̒����A�����A�s�ǖ����̍��v���擾
        intRtn = XSDCBSum(Kihon.NOWPROC, HinNow(i).SXLIDCA, lngLen, lngMAI, lngMAI800, lngFUR, lngFURKEI, lngSAM, wSAMSIJ, lngSAMFUR)

        ' ���������i�i�ԁj�F�Ǖi�̂r�w�k�h�c�ŕ��������iudtSXL�j������
        sSqlWhere = "WHERE SXLIDCB = '" & HinNow(i).SXLIDCA & "' "
        ReDim udtWSXL(0) As typ_XSDCB

        ' �f�[�^�̌������擾
        intRtn = SelCntXSDCB(sSqlWhere, intDataCnt)
        If intRtn = FUNCTION_RETURN_FAILURE Then            ' �G���[
            MsgBox "XSDCB SELECT ERROR"
            Exit Function
        Else                                                ' ����
            ' �f�[�^�����݂���ꍇ��UPDATE
            If intDataCnt > 0 Then
                intRtn = DBDRV_GetXSDCB(udtWSXL(), sSqlWhere)
                If intRtn = FUNCTION_RETURN_FAILURE Then    ' �G���[
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If

                ReDim udtSXL(0) As typ_XSDCB_Update

                ' ���������iudtSXL�j���X�V
                udtSXL(0).LENCB = lngLen

                ' �X�V�����i�ԕύX����
                udtSXL(0).HINBCB = HinNow(i).HINBCA         ' �i��
                udtSXL(0).MAICB = lngMAI
                udtSXL(0).KCNTCB = BlkNow.KCNTC2            ' �H���A��

                ' �V���O���m�莞�A�ŏI��ԋ敪='S'�ɂ���
                udtSXL(0).LIVKCB = "0"
                udtSXL(0).KANKCB = "0"                      ' �����敪
                udtSXL(0).LSTCCB = "T"                      ' �ŏI��ԋ敪
                udtSXL(0).LDFRBCB = "0"                     ' �i���敪
                udtSXL(0).LENCB = lngLen                    ' ����

                If Kihon.NOWPROC = PROCD_udtSXL_KAKUTEI Then
                    udtSXL(0).LSTCCB = "S"
                End If

                ' �H���ɂ��U���i�Ƃ肠������ʕ��j
                Select Case Kihon.NOWPROC
                     Case "CW740"
                        udtSXL(0).SXLNMAICB = lngFUR       ' �p��WF����
                        udtSXL(0).NEWKNTCB = "CW740"       ' �ŏI�ʉߍH��
                        udtSXL(0).GNWKNTCB = "CW750"       ' ���ݍH��
                     Case "CW750"
                        udtSXL(0).SRMAICB = lngSIJ         ' �T���v�����w������
                        udtSXL(0).SNMAICB = lngSAMFUR      ' �T���v�����w���s�ǖ���
                        udtSXL(0).STMAICB = lngSAM         ' �T���v������

                        If SelectSxlID039 = HinNow(i).SXLIDCA Then
                           udtSXL(0).NEWKNTCB = "CW750"    ' �ŏI�ʉߍH��
                           udtSXL(0).GNWKNTCB = "CW800"    ' ���ݍH��
                        End If

                        ' ����SXL�ȊO�ͤ�ŏI��ԋ敪��ύX���Ȃ�
                        If Trim(HinNow(i).SXLIDCA) <> SelectSxlID039 Then
                            udtSXL(0).LSTCCB = udtWSXL(1).LSTCCB  ' �ŏI��ԋ敪
                        End If
                     Case "CW760"
                        udtSXL(0).SXLNMAICB = lngFUR       ' �p��WF����

                        ' Z�i�Ԃ̏ꍇ�͔p���Ƃ���B
                        For m = 1 To UBound(tblWfSxlMng())
                            If HinNow(i).SXLIDCA = tblWfSxlMng(m).SXLID Then
                            udtSXL(0).NEWKNTCB = "CW760"       ' �ŏI�ʉߍH��
                            udtSXL(0).GNWKNTCB = "CW750"       ' ���ݍH��
                            If Trim(udtWSXL(1).FACTORYCB) = "" Then udtSXL(0).FACTORYCB = HinNow(i).FACTORYCA   ' �H��
                            If Trim(udtWSXL(1).OPECB) = "" Then udtSXL(0).OPECB = HinNow(i).OPECA               ' ���Ə���
                            If Trim(udtWSXL(1).MOTHINCB) = "" Then
                                udtSXL(0).MOTHINCB = vbNullString ' ������
                                udtSXL(0).INPOSCB = udtWSXL(1).INPOSCB
                                If udtSXL(0).INPOSCB = "" Then udtSXL(0).INPOSCB = HinNow(i).INPOSCA
                                For intLoopBkHinGet = 0 To Kihon.CNTHINOLD - 1
                                    If (CInt(HinOld(intLoopBkHinGet).INPOSCA) <= CInt(udtSXL(0).INPOSCB)) And (CInt(udtSXL(0).INPOSCB) <= CInt(HinOld(intLoopBkHinGet).INPOSCA) + CInt(HinOld(intLoopBkHinGet).GNLCA)) Then
                                        udtSXL(0).MOTHINCB = HinOld(intLoopBkHinGet).HINBCA
                                        Exit For
                                    End If
                                Next

                                ' �����Y��HINOLD�����������玩���̕i�Ԃ����i�ԂƂ���
                                If udtSXL(0).MOTHINCB = vbNullString Then
                                    udtSXL(0).MOTHINCB = udtSXL(0).HINBCB
                                End If
                            End If
                               ' �p���̏ꍇ
                               If Trim(tblWfSxlMng(m).hinban) = "Z" Then
                                   udtSXL(0).NEWKNTCB = "CW760"         ' �ŏI�ʉߍH��
                                   udtSXL(0).GNWKNTCB = "TX860"         ' ���ݍH��
                                   udtSXL(0).LSTCCB = "H"               ' �ŏI��ԋ敪
                               End If
                               Exit For
                           End If
                        Next m
                    Case "CW800"
                        udtSXL(0).SXLRMAICB = lngMAI800                 ' SXL�w���i�Ǖi�j
                        udtSXL(0).WFCNMAICB = lngFURKEI                 ' WFC����������
                End Select
                udtSXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                udtSXL(0).PLANTCATCB = sCmbMukesaki
                If sKanrenFlg = "1" Then udtSXL(0).KBLKFLGCB = sKanrenFlg      ' �֘A��ۯ��׸ށ@08/01/31 ooba

                intRtn = UpdateXSDCB(udtSXL(0), sSqlWhere)

                ' ���������iudtSXL�j�X�V�G���[
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCB UPDATET ERROR"
                    Exit Function
                End If
             ' ���݂��Ȃ����A�ǉ�
             ElseIf intDataCnt = 0 Then
                ReDim udtSXL(0) As typ_XSDCB_Update
                udtSXL(0).SXLIDCB = HinNow(i).SXLIDCA                   ' SXLID
                udtSXL(0).KCNTCB = BlkNow.KCNTC2                        ' �H���A��
                udtSXL(0).XTALCB = HinNow(i).XTALCA                     ' �����ԍ�
                udtSXL(0).INPOSCB = HinNow(i).INPOSCA                   ' �������J�n�ʒu
                udtSXL(0).LENCB = lngLen                                ' ����
                udtSXL(0).HINBCB = HinNow(i).HINBCA                     ' �i��
                udtSXL(0).REVNUMCB = HinNow(i).REVNUMCA                 ' �d�b�ԍ������ԍ�
                udtSXL(0).FACTORYCB = HinNow(i).FACTORYCA               ' �H��
                udtSXL(0).OPECB = HinNow(i).OPECA                       ' ���Ə���
                udtSXL(0).MAICB = lngMAI                                ' ������
                udtSXL(0).WSRMAICB = 0                                  ' WS��㖇��
                udtSXL(0).WSNMAICB = 0                                  ' WS��򌇗�����
                udtSXL(0).WFCMAICB = 0                                  ' �������
                udtSXL(0).SXLRMAICB = 0                                 ' SXL�w��(�Ǖi)
                udtSXL(0).SXLEMAICB = 0                                 ' SXL�m�薇��
                udtSXL(0).FURIMAICB = ""                                ' �U�֖���
                udtSXL(0).XTWORKCB = "42"                               ' �����H��
                udtSXL(0).WFWORKCB = " "                                ' �E�F�[�n����
                udtSXL(0).FURYCCB = " "                                 ' �s�Ǘ��R
                udtSXL(0).LSTCCB = "T"                                  ' �̎��ԋ敪

                ' �V���O���m�莞�A�ŏI��ԋ敪='S'�ɂ���
                If Kihon.NOWPROC = PROCD_udtSXL_KAKUTEI Then
                   udtSXL(0).LSTCCB = "S"
                End If

                udtSXL(0).LUFRCCB = " "                                 ' �i��R�[�h
                udtSXL(0).LUFRBCB = " "                                 ' �i��敪
                udtSXL(0).LDERCCB = " "                                 ' �i���R�[�h
                udtSXL(0).LDFRBCB = "0"                                 ' �i���敪
                udtSXL(0).HOLDCCB = " "                                 ' �z�[���h�R�[�h
                udtSXL(0).HOLDBCB = " "                                 ' �z�[���h�敪
                udtSXL(0).EXKUBCB = " "                                 ' ��O�敪
                udtSXL(0).HENPKCB = " "                                 ' �ԕi�敪
                udtSXL(0).KANKCB = "0"                                  ' �����敪
                udtSXL(0).LIVKCB = "0"                                  ' �����敪
                udtSXL(0).NFCB = "0"                                    ' ���ɋ敪
                udtSXL(0).SAKJCB = "0"                                  ' �폜�敪
                udtSXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �o�^���t
                udtSXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
                udtSXL(0).SUMITCB = "0"                                 ' SUMIT���M�t���O
                udtSXL(0).SNDKCB = "0"                                  ' �ԕi�敪
                udtSXL(0).SNDAYCB = ""                                  ' ���M���t
                udtSXL(0).PLANTCATCB = HinNow(i).PLANTCATCA             ' ����
                If sKanrenFlg = "1" Then udtSXL(0).KBLKFLGCB = sKanrenFlg      ' �֘A��ۯ��׸ށ@08/01/31 ooba


                ' �H���ɂ��U���i�Ƃ肠������ʕ��j
                Select Case Kihon.NOWPROC
                    Case "CW740"
                        udtSXL(0).SXLNMAICB = lngFUR                    ' �p��WF����
                        udtSXL(0).NEWKNTCB = "CW740"                    ' �ŏI�ʉߍH��
                        udtSXL(0).GNWKNTCB = "CW750"                    ' ���ݍH��
                    Case "CW750"
                        udtSXL(0).SRMAICB = lngSIJ                      ' �T���v�����w������
                        udtSXL(0).SNMAICB = lngSAMFUR                   ' �T���v�����w���s�ǖ���
                        udtSXL(0).STMAICB = lngSAM                      ' �T���v������

                        ' �u���b�N�P�ʂŕύX����Ă��܂����߃R�����g��
                        If SelectSxlID039 = HinNow(i).SXLIDCA Then
                           udtSXL(0).NEWKNTCB = "CW750"                 ' �ŏI�ʉߍH��
                           udtSXL(0).GNWKNTCB = "CW800"                 ' ���ݍH��
                        End If
                    Case "CW760"
                        udtSXL(0).SXLNMAICB = lngFUR                    ' �p��WF����
                        udtSXL(0).NEWKNTCB = "CW760"                    ' �ŏI�ʉߍH��
                        udtSXL(0).GNWKNTCB = "CW750"                    ' ���ݍH��

                        ' Z�i�Ԃ̏ꍇ�͔p���Ƃ���B
                        For m = 1 To UBound(tblWfSxlMng())
                           If HinNow(i).SXLIDCA = tblWfSxlMng(m).SXLID Then
                               If Trim(tblWfSxlMng(m).hinban) = "Z" Then
                                   udtSXL(0).NEWKNTCB = "CW760"         ' �ŏI�ʉߍH��
                                   udtSXL(0).GNWKNTCB = "TX860 "        ' ���ݍH��
                                   udtSXL(0).LSTCCB = "H"               ' �ŏI��ԋ敪

                                   ' �o�^�̏ꍇ�̂ݐݒ�
                                   udtSXL(0).FURYCCB = tblWfSxlMng(m).BDCAUS        ' �s�Ǘ��R
                                   udtSXL(0).RLENCB = tblWfSxlMng(m).LENGTH         ' ���_����
                                   udtSXL(0).SHOLDCLSCB = tblWfSxlMng(m).HOLDCLS    ' �z�[���h�敪
                               End If
                               Exit For
                           End If
                        Next m
                    Case "CW800"
                        udtSXL(0).SXLRMAICB = lngMAI                    ' SXL�w���i�Ǖi�j
                        udtSXL(0).WFCNMAICB = lngFURKEI                 ' WFC����������
                End Select

                ' �O�ް����Ȃ��A���i�Ԃ��擾�ł��Ȃ��̂ŁAHINOLD����Y���ʒu�̕i�Ԃ��擾���A��������i�ԂƂ���@---------------
                udtSXL(0).MOTHINCB = vbNullString '������
                ' CW740,CW760�݂̂ɕύX
                If Kihon.NOWPROC = "CW740" Or Kihon.NOWPROC = "CW760" Then
                    For intLoopBkHinGet = 0 To Kihon.CNTHINOLD - 1
                        If (CInt(HinOld(intLoopBkHinGet).INPOSCA) <= CInt(udtSXL(0).INPOSCB)) And (CInt(udtSXL(0).INPOSCB) <= CInt(HinOld(intLoopBkHinGet).INPOSCA) + CInt(HinOld(intLoopBkHinGet).GNLCA)) Then
                            udtSXL(0).MOTHINCB = HinOld(intLoopBkHinGet).HINBCA
                            Exit For
                        End If
                    Next
                   If udtSXL(0).MOTHINCB = vbNullString Then ' �����Y��HINOLD�����������玩���̕i�Ԃ����i�ԂƂ���
                       udtSXL(0).MOTHINCB = udtSXL(0).HINBCB
                   End If
                End If
                intRtn = CreateXSDCB(udtSXL(0), sErrMsg)

                ' ���������iudtSXL�j�ǉ��G���[
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox sErrMsg
                    Exit Function
                End If
             End If
        End If
    Next i

''''' sta�F���������i�i�ԁj�̍��v������0�ɂȂ������A
''''' ���������iudtSXL�j�������b�g�ɂ���
    For i = 0 To Kihon.CNTHINOLD - 1
        ' �H�����т��瓯���r�w�k�h�c�̒����A�����A�s�ǖ����̍��v���擾
        intRtn = XSDCBSum(Kihon.NOWPROC, HinOld(i).SXLIDCA, lngLen, lngMAI, pMAI800, lngFUR, lngFURKEI, lngSAM, wSAMSIJ, lngSAMFUR)

        ' ���������i�i�ԁj�F�r�w�k�h�c�ŕ��������iudtSXL�j������
        sSqlWhere = "WHERE SXLIDCB = '" & HinOld(i).SXLIDCA & "' "
        ReDim udtWSXL(0) As typ_XSDCB

        ' �f�[�^�̌������擾
        intRtn = SelCntXSDCB(sSqlWhere, intDataCnt)
        If intRtn = FUNCTION_RETURN_FAILURE Then    ' �G���[
            MsgBox "XSDCB SELECT ERROR"
            Exit Function
        Else                                        ' ����
            ' �f�[�^�����݂���ꍇ��UPDATE
            If intDataCnt > 0 Then
                intRtn = DBDRV_GetXSDCB(udtWSXL(), sSqlWhere)
                If intRtn = FUNCTION_RETURN_FAILURE Then  '�G���[
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If

                ReDim udtSXL(0) As typ_XSDCB_Update

                ' ���������iudtSXL�j���X�V
                udtSXL(0).LENCB = lngLen
                udtSXL(0).MAICB = lngMAI
                udtSXL(0).KCNTCB = BlkNow.KCNTC2

                '������0�̎��A�����b�g�Ƃ���
                If (lngMAI = 0 And Kihon.NOWPROC <> PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Or _
                        (lngLen = 0 And Kihon.NOWPROC = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Then    '05/03/29 ooba
                     udtSXL(0).LIVKCB = "1"                 ' �����敪
                     udtSXL(0).KANKCB = "2"                 ' �����敪
                     udtSXL(0).LSTCCB = "H"                 ' �ŏI��ԋ敪
                     udtSXL(0).LDFRBCB = "2"                ' �i���敪
                 Else
                     udtSXL(0).LIVKCB = "0"
                     udtSXL(0).KANKCB = "0"                 ' �����敪
                     udtSXL(0).LSTCCB = "T"                 ' �ŏI��ԋ敪
                     udtSXL(0).LDFRBCB = "0"                ' �i���敪
                End If
                udtSXL(0).KANKCB = "0"                      ' �����敪

                ' �H���ɂ��U���i�Ƃ肠������ʕ��j
                Select Case Kihon.NOWPROC
                    Case "CW740"
                        udtSXL(0).SXLNMAICB = lngFUR                ' �p��WF����
                    Case "CW750"
                        udtSXL(0).SRMAICB = lngSIJ                  ' �T���v�����w������
                        udtSXL(0).SNMAICB = lngSAMFUR               ' �T���v�����w���s�ǖ���
                        udtSXL(0).STMAICB = lngSAM                  ' �T���v������

                        ' ����SXL�ȊO�ͤ�ŏI��ԋ敪��ύX���Ȃ�
                        If Trim(HinOld(i).SXLIDCA) <> SelectSxlID039 Then
                            udtSXL(0).LSTCCB = udtWSXL(1).LSTCCB    ' �ŏI��ԋ敪
                        End If
                    Case "CW760"
                        udtSXL(0).SXLNMAICB = lngFUR               ' �p��WF����

                        ' Z�i�Ԃ̏ꍇ�͔p���Ƃ���B
                        For m = 1 To UBound(tblWfSxlMng())
                            If HinOld(i).SXLIDCA = tblWfSxlMng(m).SXLID Then
                                If Trim(udtWSXL(1).FACTORYCB) = "" Then udtSXL(0).FACTORYCB = HinOld(i).FACTORYCA   ' �H��
                                If Trim(udtWSXL(1).OPECB) = "" Then udtSXL(0).OPECB = HinOld(i).OPECA               ' ���Ə���
                                ' �p���̏ꍇ
                                If Trim(tblWfSxlMng(m).hinban) = "Z" Then
                                    udtSXL(0).NEWKNTCB = "CW760"    ' �ŏI�ʉߍH��
                                    udtSXL(0).GNWKNTCB = "TX860"    ' ���ݍH��
                                    udtSXL(0).LSTCCB = "H"          ' �ŏI��ԋ敪
                                End If
                                Exit For
                            End If
                        Next m
                    Case "CW800"
                        udtSXL(0).SXLRMAICB = lngMAI         ' SXL�w���i�Ǖi�j
                        udtSXL(0).WFCNMAICB = lngFURKEI      ' WFC����������
                End Select

                ' �V���O���m�莞�A�ŏI��ԋ敪='S'�ɂ���
                If Kihon.NOWPROC = PROCD_udtSXL_KAKUTEI Then
                   udtSXL(0).LSTCCB = "S"
                End If

                udtSXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                udtSXL(0).PLANTCATCB = sCmbMukesaki
                If sKanrenFlg = "1" Then udtSXL(0).KBLKFLGCB = sKanrenFlg      ' �֘A��ۯ��׸ށ@08/01/31 ooba

                intRtn = UpdateXSDCB(udtSXL(0), sSqlWhere)

                ' ���������iudtSXL�j�X�V�G���[
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCB UPDATET ERROR"
                    Exit Function
                End If
             ' ���݂��Ȃ����A�ǉ�
             ElseIf intDataCnt = 0 Then
                ReDim udtSXL(0) As typ_XSDCB_Update
                udtSXL(0).SXLIDCB = HinOld(i).SXLIDCA      ' SXLID
                udtSXL(0).KCNTCB = BlkNow.KCNTC2           ' �H���A��
                udtSXL(0).XTALCB = HinOld(i).XTALCA        ' �����ԍ�
                udtSXL(0).INPOSCB = HinOld(i).INPOSCA      ' �������J�n�ʒu
                udtSXL(0).LENCB = lngLen                   ' ����
                udtSXL(0).HINBCB = HinOld(i).HINBCA        ' �i��
                udtSXL(0).REVNUMCB = HinOld(i).REVNUMCA    ' �d�b�ԍ������ԍ�
                udtSXL(0).FACTORYCB = HinOld(i).FACTORYCA  ' �H��
                udtSXL(0).OPECB = HinOld(i).OPECA          ' ���Ə���
                udtSXL(0).MAICB = lngMAI                   ' ������
                udtSXL(0).WSRMAICB = 0                     ' WS��㖇��
                udtSXL(0).WSNMAICB = 0                     ' WS��򌇗�����
                udtSXL(0).WFCMAICB = 0                     ' �������
                udtSXL(0).WSNMAICB = 0                     ' WS��򌇗�����
                udtSXL(0).WFCMAICB = 0                     ' �������
                udtSXL(0).SXLEMAICB = 0                    ' SXL�m�薇��

                udtSXL(0).FURIMAICB = ""                   ' �U�֖���
                udtSXL(0).XTWORKCB = "42"                  ' �����H��
                udtSXL(0).WFWORKCB = " "                   ' �E�F�[�n����
                udtSXL(0).FURYCCB = " "                    ' �s�Ǘ��R
                udtSXL(0).LSTCCB = "T"                     ' �̎��ԋ敪
                udtSXL(0).LUFRCCB = " "                    ' �i��R�[�h
                udtSXL(0).LUFRBCB = " "                    ' �i��敪
                udtSXL(0).LDERCCB = " "                    ' �i���R�[�h

                ' ������0�̎��A�p���Ƃ���
                If wLENCB = 0 Then
                    udtSXL(0).LDFRBCB = "2"                ' �i���敪
                Else
                    udtSXL(0).LDFRBCB = "0"
                End If

                udtSXL(0).HOLDCCB = " "                    ' �z�[���h�R�[�h
                udtSXL(0).HOLDBCB = " "                    ' �z�[���h�敪
                udtSXL(0).EXKUBCB = " "                    ' ��O�敪
                udtSXL(0).HENPKCB = " "                    ' �ԕi�敪

                ' ������0�̎��A�����b�g�Ƃ���
                If (lngMAI = 0 And Kihon.NOWPROC <> PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Or _
                        (lngLen = 0 And Kihon.NOWPROC = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Then    '05/03/29 ooba
                    udtSXL(0).LIVKCB = "1"                 ' �����敪
                    udtSXL(0).KANKCB = "2"                 ' �����敪
                    udtSXL(0).LSTCCB = "H"                 ' �ŏI��ԋ敪
                    udtSXL(0).LDFRBCB = "2"                ' �i���敪
                Else
                    udtSXL(0).LIVKCB = "0"
                    udtSXL(0).KANKCB = "0"                 ' �����敪
                    udtSXL(0).LSTCCB = "T"                 ' �ŏI��ԋ敪
                    udtSXL(0).LDFRBCB = "0"                ' �i���敪
                End If

                ' �H���ɂ��U���i�Ƃ肠������ʕ��j
                Select Case Kihon.NOWPROC
                    Case "CW740"
                        udtSXL(0).SXLNMAICB = lngFUR         ' �p��WF����
                    Case "CW750"
                        udtSXL(0).SRMAICB = lngSIJ           ' �T���v�����w������
                        udtSXL(0).SNMAICB = lngSAMFUR        ' �T���v�����w���s�ǖ���
                        udtSXL(0).STMAICB = lngSAM           ' �T���v������
                    Case "CW760"
                        udtSXL(0).SXLNMAICB = lngFUR         ' �p��WF����

                        ''Z�i�Ԃ̏ꍇ�͔p���Ƃ���B
                        For m = 1 To UBound(tblWfSxlMng())
                           If HinOld(i).SXLIDCA = tblWfSxlMng(m).SXLID Then
                               If Trim(tblWfSxlMng(m).hinban) = "Z" Then
                                   udtSXL(0).NEWKNTCB = "CW760"               ' �ŏI�ʉߍH��
                                   udtSXL(0).GNWKNTCB = "TX860"               ' ���ݍH��
                                   udtSXL(0).LSTCCB = "H"                     ' �ŏI��ԋ敪

                                   ' �o�^�̏ꍇ�̂ݐݒ�
                                   udtSXL(0).FURYCCB = tblWfSxlMng(m).BDCAUS  ' �s�Ǘ��R
                                   udtSXL(0).RLENCB = tblWfSxlMng(m).LENGTH   ' ���_����
                                   udtSXL(0).SHOLDCLSCB = tblWfSxlMng(m).HOLDCLS  ''�z�[���h�敪
                               End If
                               Exit For
                           End If
                        Next m
                    Case "CW800"
                        udtSXL(0).SXLRMAICB = lngMAI         ' SXL�w���i�Ǖi�j
                        udtSXL(0).WFCNMAICB = lngFURKEI      ' WFC����������
                End Select

                ' �����敪�t���O�ύX
                udtSXL(0).KANKCB = "0"                 ' �����敪

                ' �V���O���m�莞�A�ŏI��ԋ敪='S'�ɂ���
                If Kihon.NOWPROC = PROCD_udtSXL_KAKUTEI Then
                   udtSXL(0).LSTCCB = "S"
                End If

                udtSXL(0).NFCB = "0"                                    ' ���ɋ敪
                udtSXL(0).SAKJCB = "0"                                  ' �폜�敪
                udtSXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �o�^���t
                udtSXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
                udtSXL(0).SUMITCB = "0"                                 ' SUMIT���M�t���O
                udtSXL(0).SNDKCB = "0"                                  ' �ԕi�敪
                udtSXL(0).SNDAYCB = ""                                  ' ���M���t
                udtSXL(0).PLANTCATCB = sCmbMukesaki
                If sKanrenFlg = "1" Then udtSXL(0).KBLKFLGCB = sKanrenFlg      ' �֘A��ۯ��׸ށ@08/01/31 ooba

                intRtn = CreateXSDCB(udtSXL(0), sErrMsg)

                ' ���������iudtSXL�j�ǉ��G���[
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox sErrMsg
                    Exit Function
                End If
             End If
        End If
    Next i

    XSDCBProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDCBProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'************************************************************************************************
'*    �֐���        : XSDCBSum
'*
'*    �����T�v      : 1.�H�����т���w�肳�ꂽ�H���A�r�w�k�h�c�i�u���b�N�h�c�A�ʒu�A�i�ԁj�̒����A
'*                      �����A�s�ǖ������W�v�����M�p�̃e�[�u������w�肳�ꂽ�r�w�k�h�c�j��
'*                      �T���v�������A�T���v�����w�������A�T���v�����w���s�ǖ������W�v����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    pKKTC          I   string   �H��
'*                    pSXLID         I   string   �r�w�k�h�c
'*                    pLEN           O   NUMBER   ����
'*                    pMAI           O   NUMBER   ����
'*                    pMAI800        O   NUMBER   CW800����
'*                    pFUR           O   NUMBER   �s�ǖ���
'*                    pFURKEI        O   NUMBER   �s�ǖ������v
'*                    pSAM           O   NUMBER   �T���v������
'*                    pSAMNUK        O   NUMBER   �T���v�����w������
'*                    pSAMFUR        O   NUMBER   �T���v�����w���s�ǖ���
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************************
Public Function XSDCBSum(ByVal pKKTC, ByVal pSXLID, ByRef pLEN, ByRef pMAI, ByRef pMAI800, ByRef pFUR, ByRef pFURKEI, ByRef pSAM, ByRef pSAMSIJ, ByRef pSAMFUR)

    ' �����ϐ�
    Dim i               As Integer
    Dim intRtn          As Integer          ' ���A���
    Dim sSQL            As String           ' �r�p�k
    Dim rs              As OraDynaset       ' ���R�[�h�Z�b�g
    Dim sCRYNUMCA       As String           ' �u���b�N�h�c
    Dim lngINPOSCA      As Long             ' �J�n�ʒu
    Dim sHINBCA         As String           ' �i��
    Dim lngLen          As Long             ' ����
    Dim lngMAI          As Long             ' ����
    Dim lngMAI800       As Long             ' CW800��ʉ߂�������
    Dim lngFUR          As Long             ' �s�ǖ���
    Dim lngFURKEI       As Long             ' �s�ǖ������v
    Dim sKCNTC3         As String           ' �H���A�ԍő�
    Dim sSAMFUR         As String           ' �T���v�������w���s�ǖ���
    Dim rsXsdca         As OraDynaset
    Dim rsMain          As OraDynaset

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    ' �p�����[�^������
    pLEN = 0
    pMAI = 0
    pMAI800 = 0
    pFUR = 0
    pFURKE = 0
    pSAM = 0
    pSAMSIJ = 0
    pSAMFUR = 0

    ' ���������i�i�ԁj����p�����[�^�̂r�w�k�h�c�̒����A�������擾
    sSQL = "SELECT SUM(GNLCA) AS wLEN, SUM(GNMCA) AS lngMAI "
    sSQL = sSQL & " FROM XSDCA "
    sSQL = sSQL & " WHERE CRYNUMCA like '" & left(pSXLID, 9) & "%' "    '���ޯ�����ڒǉ� 09/05/25 ooba
    sSQL = sSQL & " AND SXLIDCA = '" & pSXLID & "' "
    sSQL = sSQL & " AND LIVKCA = '0' "

    Set rsXsdca = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' ���݂��Ȃ����A����
    If rsXsdca.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rsXsdca.Close
        GoTo CW800_CAL
    End If

    ' ���o���ʂ��i�[����
    If IsNull(rsXsdca.Fields("wLEN")) = True Then
        pLEN = 0
    Else
        pLEN = rsXsdca.Fields("wLEN")                   ' ����
    End If
    If IsNull(rsXsdca.Fields("lngMAI")) = True Then
        pMAI = 0
    Else
        pMAI = rsXsdca.Fields("lngMAI")                 ' ����
    End If

    rsXsdca.Close

CW800_CAL:
    ' ���������i�i�ԁj���瓯���r�w�k�h�c�̃u���b�N�h�c�A�J�n�ʒu�A�i�Ԃ��擾
    sSQL = "SELECT CRYNUMCA, INPOSCA, HINBCA "
    sSQL = sSQL & " FROM XSDCA "
    sSQL = sSQL & " WHERE CRYNUMCA like '" & left(pSXLID, 9) & "%' "    '���ޯ�����ڒǉ� 09/05/25 ooba
    sSQL = sSQL & " AND SXLIDCA = '" & pSXLID & "' "
    sSQL = sSQL & " AND LIVKCA = '0' "
    sSQL = sSQL & " ORDER BY CRYNUMCA,INPOSCA"

    Set rsMain = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' ���݂��Ȃ����A����
    If rsMain.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rsMain.Close
        GoTo SAMPLE_CAL
    End If

    Do Until rsMain.EOF
        ' ���o���ʂ��i�[����
        sCRYNUMCA = rsMain.Fields("CRYNUMCA")
        lngINPOSCA = rsMain.Fields("INPOSCA")
        sHINBCA = rsMain.Fields("HINBCA")

        ' �擾�����u���b�N�h�c��J�n�ʒu��i�ԂōH�����т̊Y���H���ŁA�H���A�Ԃ̍ő���擾����
        sSQL = "SELECT TOMC3 AS lngMAI800,FUMC3 AS lngFUR "
        sSQL = sSQL & " FROM XSDC3 "
        sSQL = sSQL & " WHERE CRYNUMC3 = '" & sCRYNUMCA & "' "
        sSQL = sSQL & " AND INPOSC3 = " & lngINPOSCA & ""
        sSQL = sSQL & " AND KCNTC3  = (SELECT MAX(KCNTC3)"
        sSQL = sSQL & "                  FROM XSDC3"
        sSQL = sSQL & "                 WHERE CRYNUMC3 = '" & sCRYNUMCA & "' "
        sSQL = sSQL & "                   AND HINBC3 = '" & sHINBCA & "'"
        sSQL = sSQL & "                   AND INPOSC3 = '" & lngINPOSCA & "'"
        sSQL = sSQL & "                   AND WKKTC3 = '" & pKKTC & "' "
        sSQL = sSQL & "                   AND (SUMKBC3 = '0' "
        sSQL = sSQL & "                    OR SUMKBC3 = ' ' "
        sSQL = sSQL & "                    OR SUMKBC3 is null)) "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' ���݂��Ȃ����A���s
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
            GoTo SAMPLE_CAL
        End If

        ' ���o���ʂ��i�[����
        If IsNull(rs.Fields("lngMAI800")) = True Then
            pMAI800 = pMAI800 + 0
        Else
            pMAI800 = pMAI800 + CInt(rs.Fields("lngMAI800"))     '����
        End If
        If IsNull(rs.Fields("lngFUR")) = True Then
            pFUR = pFUR + 0
        Else
            pFUR = pFUR + CInt(rs.Fields("lngFUR"))              '�s�ǒ���
        End If

        ' �擾�����u���b�N�h�c��J�n�ʒu��i�ԂōH�����т̕s�Ǎ��v���擾����
        sSQL = "SELECT SUM(FUMC3) AS lngFURKEI "
        sSQL = sSQL & " FROM XSDC3 "
        sSQL = sSQL & " WHERE CRYNUMC3 = '" & sCRYNUMCA & "' "
        sSQL = sSQL & " AND INPOSC3 = " & lngINPOSCA & ""
        sSQL = sSQL & " AND HINBC3 = '" & sHINBCA & "' "
        sSQL = sSQL & " AND SUMKBC3 = '1' "
        sSQL = sSQL & " AND MODKBC3 = '0' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' ���݂��Ȃ����A���s
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
            GoTo SAMPLE_CAL
        End If

        ' ���o���ʂ��i�[����
        If IsNull(rs.Fields("lngFURKEI")) = True Then
            pFURKEI = pFURKEI + 0 '�s�ǒ���
        Else
            pFURKEI = pFURKEI + CInt(rs.Fields("lngFURKEI")) '�s�ǒ���
        End If
        rsMain.MoveNext
    Loop

    rs.Close
    rsMain.Close

SAMPLE_CAL:
    ' �]�����ʎ�M���R�[�h�����T���v���������擾����
    ' ���T���v�������@-�@�]�����ʎ�M���R�[�h���iY013)
    ' �擾�����ύX(��߰�މ�) 09/05/25 ooba
    sSQL = "SELECT COUNT(SAMPLEID) AS wSAM "
    sSQL = sSQL & " FROM TBCMY013 Y013"
    sSQL = sSQL & " WHERE  SAMPLEID in ( "
    sSQL = sSQL & " SELECT REPSMPLIDCW FROM XSDCW"
    sSQL = sSQL & " WHERE SXLIDCW = '" & pSXLID & "'"
    sSQL = sSQL & " )"
''    sSql = sSql & " SELECT E044.REPSMPLIDCW "
''    sSql = sSql & " FROM XSDCW E044 "
''    sSql = sSql & "  ,("
''    sSql = sSql & "    SELECT"
''    sSql = sSql & "      XTALCB as CRYNUM"
''    sSql = sSql & "     ,INPOSCB as INGOTPOS"
''    sSql = sSql & "     ,RLENCB as LENGTH"
''    sSql = sSql & "    FROM"
''    sSql = sSql & "      XSDCB"
''    sSql = sSql & "    WHERE SXLIDCB = '" & pSXLID & "'"
''    sSql = sSql & "   ) E042"
''    sSql = sSql & " WHERE (E044.XTALCW = E042.CRYNUM "
''    sSql = sSql & " AND  E044.INPOSCW = E042.INGOTPOS "
''    sSql = sSql & " AND E044.SMPKBNCW = 'T' ) "
''    sSql = sSql & " OR (E044.XTALCW = E042.CRYNUM"
''    sSql = sSql & " AND E044.INPOSCW = E042.INGOTPOS + E042.LENGTH "
''    sSql = sSql & " AND E044.SMPKBNCW = 'B' ))"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' ���݂��Ȃ����A���s
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
    End If

    ' ���o���ʂ��i�[����
    pSAM = rs.Fields("wSAM") ' �T���v������

    rs.Close

    ' �����w�����R�[�h�����T���v���w���������擾����
    ' ���T���v�����w�������i�Ǖi�j�@-�@�����w�����R�[�h���iY003)
    ' �擾�����ύX(��߰�މ�) 09/05/25 ooba
    sSQL = "SELECT COUNT(SAMPLEID) AS wSIJ"
    sSQL = sSQL & " FROM TBCMY003 "
    sSQL = sSQL & " WHERE SAMPLEID in ( "
    sSQL = sSQL & " SELECT REPSMPLIDCW FROM XSDCW"
    sSQL = sSQL & " WHERE SXLIDCW = '" & pSXLID & "'"
    sSQL = sSQL & " )"
''    sSql = sSql & " SELECT E044.REPSMPLIDCW "
''    sSql = sSql & " FROM XSDCW E044 "
''    sSql = sSql & "  ,("
''    sSql = sSql & "    SELECT"
''    sSql = sSql & "      XTALCB as CRYNUM"
''    sSql = sSql & "     ,INPOSCB as INGOTPOS"
''    sSql = sSql & "     ,RLENCB as LENGTH"
''    sSql = sSql & "    FROM"
''    sSql = sSql & "      XSDCB"
''    sSql = sSql & "    WHERE SXLIDCB = '" & pSXLID & "'"
''    sSql = sSql & "   ) E042"
''    sSql = sSql & " WHERE (E044.XTALCW = E042.CRYNUM "
''    sSql = sSql & " AND E044.INPOSCW = E042.INGOTPOS "
''    sSql = sSql & " AND E044.SMPKBNCW = 'T' ) "
''    sSql = sSql & " OR (E044.XTALCW = E042.CRYNUM "
''    sSql = sSql & " AND E044.INPOSCW = E042.INGOTPOS + E042.LENGTH "
''    sSql = sSql & " AND  E044.SMPKBNCW = 'B' )) "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' ���݂��Ȃ����A���s
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
    End If

    ' ���o���ʂ��i�[����
'    pSIJ = rs.Fields("wSIJ") ' �T���v�����w������
    pSAMSIJ = rs.Fields("wSIJ") ' �T���v�����w������ 09/05/25 ooba

    rs.Close

    ' C�����������T���v�����w���s�ǖ������擾����
    ' ���T���v�������w���s�ǖ����@-�@C���������@-�iY012�j

    ' �Ώۂ̃u���b�NID�擾
    sSQL = "SELECT DISTINCT(CRYNUMCA) "
    sSQL = sSQL & " FROM XSDCA"
    sSQL = sSQL & " WHERE CRYNUMCA like '" & left(pSXLID, 9) & "%' "    '���ޯ�����ڒǉ� 09/05/25 ooba
    sSQL = sSQL & " AND SXLIDCA = '" & pSXLID & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' ���݂��Ȃ����A���s
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
    End If

    Do Until rs.EOF
        ' �������COUNT(�u���b�NID���ƃ��[�v��SUM����j
        ' ���o���ʂ��i�[����
        sCRYNUMCA = rs.Fields("CRYNUMCA") ' �u���b�NID

        sSQL = "SELECT COUNT(Y012.LOTID) AS sSAMFUR "
        sSQL = sSQL & " FROM TBCMY012 Y012 "
        sSQL = sSQL & "  ,("
        sSQL = sSQL & "    SELECT"
        sSQL = sSQL & "      XTALCB as CRYNUM"
        sSQL = sSQL & "     ,INPOSCB as INGOTPOS"
        sSQL = sSQL & "     ,RLENCB as LENGTH"
        sSQL = sSQL & "    FROM"
        sSQL = sSQL & "      XSDCB"
        sSQL = sSQL & "    WHERE SXLIDCB = '" & pSXLID & "'"
        sSQL = sSQL & "   ) E042"
        sSQL = sSQL & " ,(SELECT CRYNUM, INGOTPOS, LENGTH, BLOCKID "
        sSQL = sSQL & " FROM TBCME040 "
        sSQL = sSQL & " WHERE BLOCKID =  '" & sCRYNUMCA & "' ) E040 "
        sSQL = sSQL & " WHERE Y012.LOTID = E040.BLOCKID "
        sSQL = sSQL & " AND E042.INGOTPOS <= Y012.TOP_POS / 10 + E040.INGOTPOS "
        sSQL = sSQL & " AND E042.INGOTPOS + E042.LENGTH  >= Y012.TOP_POS / 10 + E040.INGOTPOS "
        sSQL = sSQL & " AND REJCAT = 'C' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' ���݂��Ȃ����A���s
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
        End If

        ' ���o���ʂ��i�[����
        pSAMFUR = rs.Fields("sSAMFUR") ' �T���v�������w���s�ǖ���
        rs.MoveNext
    Loop

    rs.Close

    XSDCBSum = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    XSDCBSum = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : WfCount
'*
'*    �����T�v      : 1.WF�������v�Z����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                   sSelectBlkID    ,I  ,Integer  ,�u���b�NID
'*                   intBlkLen      ,I  ,Integer  ,�u���b�N����
'*                   intWfCnt       ,O  ,Integer  ,����
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function WfCount(ByVal sSelectBlkID As String, ByVal intBlkLen As Integer, ByRef intWFcnt As Integer) As FUNCTION_RETURN
    Dim udtRec()    As typ_cmkc001f_Disp
    Dim RET         As FUNCTION_RETURN
    Dim lngRecCnt   As Long
    Dim i           As Long
    Dim j           As Integer
    Dim s           As String
    Dim intWfNum    As Integer '����

    ' �����v�Z�֐��p�p�����[�^�iHSXCTCEN & HSXCYCEN�j
    ' �d�l�E���т�ǂݍ���
    RET = DBDRV_fcmkc001f_Disp(Trim(sSelectBlkID), blkInfo, udtRec)   ' SelectBlkId=�u���b�NID,blkInfo=�u���b�N�Ǘ��\����
    If RET = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    ElseIf UBound(udtRec) Then
        HSXCTCEN = udtRec(1).HSXCTCEN
        HSXCYCEN = udtRec(1).HSXCYCEN
    End If

    ' WF�����v�Z�p�̊�{�l���擾����
    Loss0 = val(GetCodeField("LG", "01", "LOSS0", "INFO1"))
    Loss4 = val(GetCodeField("LG", "01", "LOSS4", "INFO1"))
    Mlt4 = val(GetCodeField("LG", "01", "MLT4", "INFO1"))
    Pitch = val(GetCodeField("LG", "01", "PITCH", "INFO1"))

    ' �����v�Z�֐��p�p�����[�^�iSEEDDEG�j�擾
    ' �����グ�����̏ꍇ
    s = GetCodeField("SC", "28", left$(blkInfo.SEED, 1), "INFO3")
    If left$(s, 1) = "4" Then
        SEEDDEG = 4
    Else
        SEEDDEG = 0
    End If

    ' �����擾�֐�
    intWFcnt = GetWfCount(val(intBlkLen), SEEDDEG, HSXCTCEN, HSXCYCEN)  'intWfCount=����
    WfCount = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    Exit Function

End Function

'*******************************************************************************
'*    �֐���        : GetWfCount
'*
'*    �����T�v      : 1.WF�������v�Z����
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^        ,����
'*                   blkLen        ,I  ,Integer   ,�u���b�N����
'*                   seedDeg       ,I  ,Integer   ,������SEED�X��
'*                   dblHinDegT       ,I  ,Double    ,�i�ԌX���i�c�j
'*                   dblHinDegY       ,I  ,Double    ,�i�ԌX���i���j
'*
'*    �߂�l        : WF����
'*
'*******************************************************************************
Public Function GetWfCount(ByVal BlkLen%, ByVal SEEDDEG%, ByVal dblHinDegT As Double, ByVal dblHinDegY As Double) As Integer
    Dim intHinDeg   As Integer
    Dim s           As String
    Dim intWFcnt    As Integer

    If Pitch = 0# Then
        GetWfCount = 0
        Exit Function
    End If

    ' �i�ԌX���𓾂�
    ' �����ŏI�����o���A�i�ԌX���̋��ߕ��ύX
    If (Abs(dblHinDegT) = 2.83) And (Abs(dblHinDegY) = 2.83) Then
        intHinDeg = 4
    ElseIf (Abs(dblHinDegT) = 4) And (dblHinDegY = 0) Then
        intHinDeg = 4
    ElseIf (dblHinDegT = 0) And (Abs(dblHinDegY) = 4) Then
        intHinDeg = 4
    Else
        intHinDeg = 0
    End If

    ' WF�������v�Z����
    If SEEDDEG = intHinDeg Then
        ' �ʏ�i�̏ꍇ
        intWFcnt = Format(((BlkLen - Loss0) / Pitch) + 0.4, "0")
    Else
        intWFcnt = Format(((BlkLen * Mlt4 - Loss4) / Pitch) + 0.4, "0")
    End If

    GetWfCount = intWFcnt
End Function

'�T�v      :�����ŏI���o���� �\���p�c�a�h���C�o
'���Ұ��@�@:�ϐ���       ,IO ,�^                   ,����
'      �@�@:sBlockID_in�@ ,I  ,String               ,�u���b�NID
'      �@�@:udtBlkInfo�@�@�@,O  ,typ_cmkc001f_Block   ,�u���b�N���
'      �@�@:udtRecords�@�@�@,O  ,typ_cmkc001f_Disp    ,���i�d�l�擾�p
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN      ,�ǂݍ��݂̐���
'*******************************************************************************
'*    �֐���        : DBDRV_fcmkc001f_Disp
'*
'*    �����T�v      : 1.�����ŏI���o���� �\���p�c�a�h���C�o
'*
'*    �p�����[�^    : �ϐ���       ,IO ,�^                   ,����
'*               �@�@sBlockID_in�@  ,I  ,String               ,�u���b�NID
'*      �@�@     �@�@udtBlkInfo�@�@�@ ,O  ,typ_cmkc001f_Block   ,�u���b�N���
'*      �@�@     �@�@udtRecords�@�@�@ ,O  ,typ_cmkc001f_Disp    ,���i�d�l�擾�p
'*
'*    �߂�l        : WF����
'*
'*******************************************************************************
Public Function DBDRV_fcmkc001f_Disp(sBlockID_in As String, udtBlkInfo As typ_cmkc001f_Block, udtRecords() As typ_cmkc001f_Disp) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim intRecCnt   As Integer
    Dim i           As Long
    Dim n           As Integer

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_SUCCESS

    ' �u���b�N���𓾂�
    sSQL = "Select BLK.INGOTPOS, BLK.LENGTH, BLK.REALLEN, BLK.KRPROCCD, BLK.NOWPROC, BLK.LPKRPROCCD, " & _
          "BLK.LASTPASS, BLK.DELCLS, BLK.RSTATCLS, BLK.LSTATCLS, CRY.SEED " & _
          "From TBCME040 BLK, TBCME037 CRY " & _
          "Where (BLOCKID='" & sBlockID_in & "') and (BLK.CRYNUM=CRY.CRYNUM)"
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If
    With udtBlkInfo
        .INGOTPOS = rs("INGOTPOS")      ' �������J�n�ʒu
        .LENGTH = rs("LENGTH")          ' ����
        .REALLEN = rs("REALLEN")        ' ������
        .KRPROCCD = rs("KRPROCCD")      ' ���݊Ǘ��H��
        .NOWPROC = rs("NOWPROC")        ' ���ݍH��
        .LPKRPROCCD = rs("LPKRPROCCD")  ' �ŏI�ʉߊǗ��H��
        .LASTPASS = rs("LASTPASS")      ' �ŏI�ʉߍH��
        .DELCLS = rs("DELCLS")          ' �폜�敪
        .RSTATCLS = rs("RSTATCLS")      ' ������ԋ敪
        .LSTATCLS = rs("LSTATCLS")      ' �ŏI��ԋ敪
        .SEED = rs("SEED")              ' SEED
    End With
    rs.Close

    ' ���i�d�l�𓾂�
    sSQL = "select "
    sSQL = sSQL & "BH.E041HINBAN, "           ' �i��
    sSQL = sSQL & "BH.E041INGOTPOS, "         ' �������J�n�ʒu
    sSQL = sSQL & "BH.E041REVNUM, "           ' ���i�ԍ������ԍ�
    sSQL = sSQL & "BH.E041FACTORY, "          ' �H��
    sSQL = sSQL & "BH.E041OPECOND, "          ' ���Ə���
    sSQL = sSQL & "BH.E041LENGTH, "           ' ����

    ' ���i�d�lSXL�f�[�^
    sSQL = sSQL & "S.E018HSXD1CEN, "          ' �i�r�w���a�P���S
    sSQL = sSQL & "S.E018HSXRMIN, "           ' �i�r�w���R����
    sSQL = sSQL & "S.E018HSXRMAX, "           ' �i�r�w���R���
    sSQL = sSQL & "S.E018HSXRMBNP, "          ' �i�r�w���R�ʓ����z
    sSQL = sSQL & "S.E018HSXRHWYS, "          ' �i�r�w���R�ۏؕ��@�Q��
    sSQL = sSQL & "S.E019HSXONMIN, "          ' �i�r�w�_�f�Z�x����
    sSQL = sSQL & "S.E019HSXONMAX, "          ' �i�r�w�_�f�Z�x���
    sSQL = sSQL & "S.E019HSXONMBP, "          ' �i�r�w�_�f�Z�x�ʓ����z
    sSQL = sSQL & "S.E019HSXONHWS, "          ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    sSQL = sSQL & "S.E019HSXCNMIN, "          ' �i�r�w�Y�f�Z�x����
    sSQL = sSQL & "S.E019HSXCNMAX, "          ' �i�r�w�Y�f�Z�x���
    sSQL = sSQL & "S.E019HSXCNHWS, "          ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    sSQL = sSQL & "S.E019HSXTMMAXN, "         ' �i�r�w�]�ʖ��x���             ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sSQL = sSQL & "S.E020HSXBM1AN, "          ' �i�r�w�a�l�c�P���ω���
    sSQL = sSQL & "S.E020HSXBM1AX, "          ' �i�r�w�a�l�c�P���Ϗ��
    sSQL = sSQL & "S.E020HSXBM1HS, "          ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    sSQL = sSQL & "S.E020HSXBM2AN, "          ' �i�r�w�a�l�c�Q���ω���
    sSQL = sSQL & "S.E020HSXBM2AX, "          ' �i�r�w�a�l�c�Q���Ϗ��
    sSQL = sSQL & "S.E020HSXBM2HS, "          ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    sSQL = sSQL & "S.E020HSXBM3AN, "          ' �i�r�w�a�l�c�R���ω���
    sSQL = sSQL & "S.E020HSXBM3AX, "          ' �i�r�w�a�l�c�R���Ϗ��
    sSQL = sSQL & "S.E020HSXBM3HS, "          ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
    sSQL = sSQL & "S.E020HSXOF1AX, "          ' �i�r�w�n�r�e�P���Ϗ��
    sSQL = sSQL & "S.E020HSXOF1MX, "          ' �i�r�w�n�r�e�P���
    sSQL = sSQL & "S.E020HSXOF1HS, "          ' �i�r�w�n�r�e�P �ۏؕ��@�Q��
    sSQL = sSQL & "S.E020HSXOF2AX, "          ' �i�r�w�n�r�e�Q���Ϗ��
    sSQL = sSQL & "S.E020HSXOF2MX, "          ' �i�r�w�n�r�e�Q���
    sSQL = sSQL & "S.E020HSXOF2HS, "          ' �i�r�w�n�r�e�Q �ۏؕ��@�Q��
    sSQL = sSQL & "S.E020HSXOF3AX, "          ' �i�r�w�n�r�e�R���Ϗ��
    sSQL = sSQL & "S.E020HSXOF3MX, "          ' �i�r�w�n�r�e�R���
    sSQL = sSQL & "S.E020HSXOF3HS, "          ' �i�r�w�n�r�e�R �ۏؕ��@�Q��
    sSQL = sSQL & "S.E020HSXOF4AX, "          ' �i�r�w�n�r�e�S���Ϗ��
    sSQL = sSQL & "S.E020HSXOF4MX, "          ' �i�r�w�n�r�e�S���
    sSQL = sSQL & "S.E020HSXOF4HS, "          ' �i�r�w�n�r�e�S �ۏؕ��@�Q��
    sSQL = sSQL & "S.E020HSXDENMX, "          ' �i�r�w�c�������
    sSQL = sSQL & "S.E020HSXDENMN, "          ' �i�r�w�c��������
    sSQL = sSQL & "S.E020HSXDENHS, "          ' �i�r�w�c�����ۏؕ��@�Q��
    sSQL = sSQL & "S.E020HSXDVDMXN, "         ' �i�r�w�c�u�c�Q���           ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sSQL = sSQL & "S.E020HSXDVDMNN, "         ' �i�r�w�c�u�c�Q����           ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sSQL = sSQL & "S.E020HSXDVDHS, "          ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    sSQL = sSQL & "S.E020HSXLDLMX, "          ' �i�r�w�k�^�c�k���
    sSQL = sSQL & "S.E020HSXLDLMN, "          ' �i�r�w�k�^�c�k����
    sSQL = sSQL & "S.E020HSXLDLHS, "          ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    sSQL = sSQL & "S.E019HSXLTMIN, "          ' �i�r�w�k�^�C������
    sSQL = sSQL & "S.E019HSXLTMAX, "          ' �i�r�w�k�^�C�����
    sSQL = sSQL & "S.E019HSXLTHWS, "          ' �i�r�w�k�^�C���ۏؕ��@�Q��
    sSQL = sSQL & "S.E018HSXDPDIR, "          ' �i�r�w�a�ʒu����
    sSQL = sSQL & "S.E018HSXDPDRC, "          ' �i�r�w�a�ʒu����
    sSQL = sSQL & "S.E018HSXDWMIN, "          ' �i�r�w�a�Љ���
    sSQL = sSQL & "S.E018HSXDWMAX, "          ' �i�r�w�a�Џ��
    sSQL = sSQL & "S.E018HSXDDMIN, "          ' �i�r�w�a�[����
    sSQL = sSQL & "S.E018HSXDDMAX, "          ' �i�r�w�a�[���
    sSQL = sSQL & "S.E018HSXD1MIN, "          ' �i�r�w���a�P����
    sSQL = sSQL & "S.E018HSXD1MAX, "          ' �i�r�w���a�P���
    sSQL = sSQL & "S.E018HSXCTCEN, "          ' �i�r�w�����ʌX�c���S
    sSQL = sSQL & "S.E018HSXCYCEN, "          ' �i�r�w�����ʌX�����S
    sSQL = sSQL & "U.EPDUP "                  ' ���������Ǘ� EPD�@���
    sSQL = sSQL & " from VECME009 BH, VECME001 S, TBCME036 U "
    sSQL = sSQL & " where BH.E040BLOCKID='" & sBlockID_in & "' "
    sSQL = sSQL & " and S.E018HINBAN=BH.E041HINBAN "
    sSQL = sSQL & " and S.E018MNOREVNO=BH.E041REVNUM "
    sSQL = sSQL & " and S.E018FACTORY=BH.E041FACTORY "
    sSQL = sSQL & " and S.E018OPECOND=BH.E041OPECOND "
    sSQL = sSQL & " and U.HINBAN=BH.E041HINBAN "
    sSQL = sSQL & " and U.MNOREVNO=BH.E041REVNUM "
    sSQL = sSQL & " and U.FACTORY=BH.E041FACTORY "
    sSQL = sSQL & " and U.OPECOND=BH.E041OPECOND "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        ReDim udtRecords(0)
        rs.Close
        GoTo proc_exit
    End If

    intRecCnt = rs.RecordCount
    ReDim udtRecords(intRecCnt)
    For i = 1 To intRecCnt
        With udtRecords(i)
            ' �i�ԊǗ�
            .hinban = rs("E041HINBAN")              ' �i��
            .INGOTPOS = rs("E041INGOTPOS")          ' �������J�n�ʒu
            .REVNUM = rs("E041REVNUM")              ' ���i�ԍ������ԍ�
            .factory = rs("E041FACTORY")            ' �H��
            .opecond = rs("E041OPECOND")            ' ���Ə���
            .LENGTH = rs("E041LENGTH")              ' ����

            ' ���i�d�lSXL�f�[�^
            .HSXD1CEN = rs("E018HSXD1CEN")          ' �i�r�w���a�P���S
            .HSXRMIN = rs("E018HSXRMIN")            ' �i�r�w���R����
            .HSXRMAX = rs("E018HSXRMAX")            ' �i�r�w���R���
            .HSXRMBNP = rs("E018HSXRMBNP")          ' �i�r�w���R�ʓ����z
            .HSXRHWYS = rs("E018HSXRHWYS")          ' �i�r�w���R�ۏؕ��@�Q��
            .HSXONMIN = rs("E019HSXONMIN")          ' �i�r�w�_�f�Z�x����
            .HSXONMAX = rs("E019HSXONMAX")          ' �i�r�w�_�f�Z�x���
            .HSXONMBP = rs("E019HSXONMBP")          ' �i�r�w�_�f�Z�x�ʓ����z
            .HSXONHWS = rs("E019HSXONHWS")          ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
            .HSXCNMIN = rs("E019HSXCNMIN")          ' �i�r�w�Y�f�Z�x����
            .HSXCNMAX = rs("E019HSXCNMAX")          ' �i�r�w�Y�f�Z�x���
            .HSXCNHWS = rs("E019HSXCNHWS")          ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
            .HSXTMMAX = rs("E019HSXTMMAXN")         ' �i�r�w�]�ʖ��x���           ���ڒǉ��C�C���Ή� 2003.05.20 yakimura

            For n = 1 To 3
                .HSXBMnHS(n) = rs("E020HSXBM" & n & "HS")  ' �i�r�w�a�l�cn �ۏؕ��@�Q��
            Next

            For n = 1 To 4
                If IsNull(rs("E020HSXOF" & n & "AX")) = False Then .HSXOFnAX(n) = rs("E020HSXOF" & n & "AX")   ' �i�r�w�n�r�en ���Ϗ��         '05/03/29 ooba NULL�Ή�
                If IsNull(rs("E020HSXOF" & n & "MX")) = False Then .HSXOFnMX(n) = rs("E020HSXOF" & n & "MX")   ' �i�r�w�n�r�en ���             '05/03/29 ooba NULL�Ή�
                .HSXOFnHS(n) = rs("E020HSXOF" & n & "HS")   ' �i�r�w�n�r�en �ۏؕ��@�Q��
            Next

            .HSXDENMX = rs("E020HSXDENMX")          ' �i�r�w�c�������
            .HSXDENMN = rs("E020HSXDENMN")          ' �i�r�w�c��������
            .HSXDENHS = rs("E020HSXDENHS")          ' �i�r�w�c�����ۏؕ��@�Q��
            .HSXDVDMX = rs("E020HSXDVDMXN")         ' �i�r�w�c�u�c�Q���        ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
            .HSXDVDMN = rs("E020HSXDVDMNN")         ' �i�r�w�c�u�c�Q����        ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
            .HSXDVDHS = rs("E020HSXDVDHS")          ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
            .HSXLDLMX = rs("E020HSXLDLMX")          ' �i�r�w�k�^�c�k���
            .HSXLDLMN = rs("E020HSXLDLMN")          ' �i�r�w�k�^�c�k����
            .HSXLDLHS = rs("E020HSXLDLHS")          ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
            .HSXLTMIN = rs("E019HSXLTMIN")          ' �i�r�w�k�^�C������
            .HSXLTMAX = rs("E019HSXLTMAX")          ' �i�r�w�k�^�C�����
            .HSXLTHWS = rs("E019HSXLTHWS")          ' �i�r�w�k�^�C���ۏؕ��@�Q��
            .HSXDPDIR = rs("E018HSXDPDIR")          ' �i�r�w�a�ʒu����
            .HSXDPDRC = rs("E018HSXDPDRC")          ' �i�r�w�a�ʒu����
            .HSXDWMIN = rs("E018HSXDWMIN")          ' �i�r�w�a�Љ���
            .HSXDWMAX = rs("E018HSXDWMAX")          ' �i�r�w�a�Џ��
            .HSXDDMIN = rs("E018HSXDDMIN")          ' �i�r�w�a�[����
            .HSXDDMAX = rs("E018HSXDDMAX")          ' �i�r�w�a�[���
            .HSXD1MIN = rs("E018HSXD1MIN")          ' �i�r�w���a�P����
            .HSXD1MAX = rs("E018HSXD1MAX")          ' �i�r�w���a�P���
            If IsNull(rs("E018HSXCTCEN")) = False Then .HSXCTCEN = rs("E018HSXCTCEN")       ' �i�r�w�����ʌX�c���S      '05/03/29 ooba NULL�Ή�
            If IsNull(rs("E018HSXCYCEN")) = False Then .HSXCYCEN = rs("E018HSXCYCEN")       ' �i�r�w�����ʌX�����S      '05/03/29 ooba NULL�Ή�
            .EPDUP = rs("EPDUP")                    ' ���������Ǘ� EPD�@���
        End With
        rs.MoveNext
    Next
    rs.Close

proc_exit:
    '�I��
'    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Resume proc_exit
End Function

'******************************************************************************************
'*    �֐���        : SelCntXSDC4
'*
'*    �����T�v      : 1.���������i�s�Ǔ���j����A�w�肵�������ɊY������f�[�^�̍s�����擾
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@�@�@�@          sStrWhere      ,I  ,String   ,SELECT������
'*�@�@�@�@�@          intCnt        ,O  ,Integer  ,����
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'******************************************************************************************
Public Function SelCntXSDC4(ByVal sStrWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    Dim sSQL        As String           ' �r�p�k
    Dim rs          As OraDynaset       ' ���R�[�h�Z�b�g
    Dim sSqlWhere   As String           ' WHERE��

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    SelCntXSDC4 = FUNCTION_RETURN_FAILURE

    sSQL = "      SELECT count(*) cnt "
    sSQL = sSQL & "  FROM XSDC4 " & sStrWhere

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' ���݂��Ȃ����A�G���[
    If rs Is Nothing Then
        SelCntXSDC4 = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    If rs.RecordCount = 0 Then
        intCnt = 0
    Else
        intCnt = CInt(rs("cnt"))
    End If
    rs.Close

    SelCntXSDC4 = FUNCTION_RETURN_SUCCESS
    Exit Function

proc_exit:
    ' �I��
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    SelCntXSDC4 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*************************************************************************************
'*    �֐���        : SelCntXSDCA
'*
'*    �����T�v      : 1.���������i�i�ԁj����A�w�肵�������ɊY������f�[�^�̍s�����擾
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@�@�@�@          sStrWhere      ,I  ,String   ,SELECT������
'*�@�@�@�@�@          intCnt        ,O  ,Integer  ,����
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*************************************************************************************
Public Function SelCntXSDCA(ByVal sStrWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    Dim sSQL        As String           ' �r�p�k
    Dim rs          As OraDynaset       ' ���R�[�h�Z�b�g
    Dim sSqlWhere   As String           ' WHERE��

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    SelCntXSDCA = FUNCTION_RETURN_FAILURE

    sSQL = "      SELECT count(*) cnt "
    sSQL = sSQL & "  FROM XSDCA " & sStrWhere

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' ���݂��Ȃ����A�G���[
    If rs Is Nothing Then
        SelCntXSDCA = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    If rs.RecordCount = 0 Then
        intCnt = 0
    Else
        intCnt = CInt(rs("cnt"))
    End If
    rs.Close

    SelCntXSDCA = FUNCTION_RETURN_SUCCESS
    Exit Function

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    SelCntXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'*************************************************************************************
'*    �֐���        : SelCntXSDCB
'*
'*    �����T�v      : 1.���������iSXL�j����A�w�肵�������ɊY������f�[�^�̍s�����擾
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@�@�@�@          sStrWhere      ,I  ,String   ,SELECT������
'*�@�@�@�@�@          intCnt        ,O  ,Integer  ,����
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*************************************************************************************
Public Function SelCntXSDCB(ByVal sStrWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    Dim sSQL        As String           ' �r�p�k
    Dim rs          As OraDynaset       ' ���R�[�h�Z�b�g
    Dim sSqlWhere   As String           ' WHERE��

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    SelCntXSDCB = FUNCTION_RETURN_FAILURE

    sSQL = "      SELECT count(*) cnt "
    sSQL = sSQL & "  FROM XSDCB " & sStrWhere

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' ���݂��Ȃ����A�G���[
    If rs Is Nothing Then
        SelCntXSDCB = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    If rs.RecordCount = 0 Then
        intCnt = 0
    Else
        intCnt = CInt(rs("cnt"))
    End If
    rs.Close

    SelCntXSDCB = FUNCTION_RETURN_SUCCESS
    Exit Function

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    SelCntXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*************************************************************************************
'*    �֐���        : clearType
'*
'*    �����T�v      : 1.�VDB�\���̏�����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*************************************************************************************
Public Sub clearType()
    On Error Resume Next

    ' ��{���
    With Kihon
        .ALLSCRAP = ""
        .CNTHINNOW = 0
        .CNTHINOLD = 0
        .DIAMETER = 0
        .FURYOUMU = ""
        .NEWPROC = ""
        .NOWPROC = ""
        .STAFFID = ""
    End With

    ' ��������(�u���b�N)�F�O�H��
    With BlkOld
        .CRYNUMC2 = ""
        .KCNTC2 = ""
        .XTALC2 = ""
        .INPOSC2 = ""
        .NEKKNTC2 = ""
        .NEWKNTC2 = ""
        .NEWKKBC2 = ""
        .NEMACOC2 = ""
        .GNKKNTC2 = ""
        .GNWKNTC2 = ""
        .GNWKKBC2 = ""
        .GNMACOC2 = ""
        .GNDAYC2 = ""
        .GNLC2 = ""
        .GNWC2 = ""
        .GNMC2 = ""
        .SUMITLC2 = ""
        .SUMITWC2 = ""
        .SUMITMC2 = ""
        .CHGC2 = ""
        .KAKOUBC2 = ""
        .KEIDAYC2 = ""
        .GNTKUBC2 = ""
        .GNTNOC2 = ""
        .XTWORKC2 = ""
        .WFWORKC2 = ""
        .LSTATBC2 = ""
        .RSTATBC2 = ""
        .LUFRCC2 = ""
        .LUFRBC2 = ""
        .LDFRCC2 = ""
        .LDFRBC2 = ""
        .HOLDCC2 = ""
        .HOLDBC2 = ""
        .EXKUBC2 = ""
        .HENPKC2 = ""
        .LIVKC2 = ""
        .KANKC2 = ""
        .NFC2 = ""
        .SAKJC2 = ""
        .TDAYC2 = ""
        .KDAYC2 = ""
        .SUMITBC2 = ""
        .SNDKC2 = ""
        .SNDDAYC2 = ""
    End With
    With BlkNow
        .CRYNUMC2 = ""
        .KCNTC2 = ""
        .XTALC2 = ""
        .INPOSC2 = ""
        .NEKKNTC2 = ""
        .NEWKNTC2 = ""
        .NEWKKBC2 = ""
        .NEMACOC2 = ""
        .GNKKNTC2 = ""
        .GNWKNTC2 = ""
        .GNWKKBC2 = ""
        .GNMACOC2 = ""
        .GNDAYC2 = ""
        .GNLC2 = ""
        .GNWC2 = ""
        .GNMC2 = ""
        .SUMITLC2 = ""
        .SUMITWC2 = ""
        .SUMITMC2 = ""
        .CHGC2 = ""
        .KAKOUBC2 = ""
        .KEIDAYC2 = ""
        .GNTKUBC2 = ""
        .GNTNOC2 = ""
        .XTWORKC2 = ""
        .WFWORKC2 = ""
        .LSTATBC2 = ""
        .RSTATBC2 = ""
        .LUFRCC2 = ""
        .LUFRBC2 = ""
        .LDFRCC2 = ""
        .LDFRBC2 = ""
        .HOLDCC2 = ""
        .HOLDBC2 = ""
        .EXKUBC2 = ""
        .HENPKC2 = ""
        .LIVKC2 = ""
        .KANKC2 = ""
        .NFC2 = ""
        .SAKJC2 = ""
        .TDAYC2 = ""
        .KDAYC2 = ""
        .SUMITBC2 = ""
        .SNDKC2 = ""
        .SNDDAYC2 = ""
    End With

    ReDim HinOld(0) As typ_XSDCA_Update
    ReDim HinNow(0) As typ_XSDCA_Update

    With Furyou
        .XTALC4 = ""
        .INPOSC4 = ""
        .KCKNTC4 = ""
        .HINBC4 = ""
        .REVNUMC4 = ""
        .FACTORYC4 = ""
        .OPEC4 = ""
        .KNKTC4 = ""
        .WKKTC4 = ""
        .WKKDC4 = ""
        .MACOC4 = ""
        .SXLIDC4 = ""
        .FCODEC4 = ""
        .PUCUTLC4 = ""
        .PUCUTWC4 = ""
        .PUCUTMC4 = ""
        .FKUBC4 = ""
        .TDAYC4 = ""
        .KDAYC4 = ""
        .SUMITBC3 = ""
        .SNDKC3 = ""
        .SNDDAYC3 = ""
    End With

End Sub

'*******************************************************************************
'*    �֐���        : calculateWfNum
'*
'*    �����T�v      : 1.WF�������v�Z����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^        ,����
'*                    blkLen        ,I  ,Integer   ,�u���b�N����
'*                    seedDeg       ,I  ,Integer   ,������SEED�X��
'*                    dblHinDegT       ,I  ,Double    ,�i�ԌX���i�c�j
'*                    dblHinDegY       ,I  ,Double    ,�i�ԌX���i���j
'*
'*    �߂�l        : WF����
'*
'*******************************************************************************
Private Function calculateWfNum(ByVal BlkLen%, ByVal SEEDDEG%, ByVal dblHinDegT As Double, ByVal dblHinDegY As Double) As Integer
    Dim intHinDeg   As Integer
    Dim s           As String
    Dim intWFcnt    As Integer

    If Pitch = 0# Then
        calculateWfNum = 0
        Exit Function
    End If

    ' �i�ԌX���𓾂�
    ' �����ŏI�����o���A�i�ԌX���̋��ߕ��ύX
    If (Abs(dblHinDegT) = 2.83) And (Abs(dblHinDegY) = 2.83) Then
        intHinDeg = 4
    ElseIf (Abs(dblHinDegT) = 4) And (dblHinDegY = 0) Then
        intHinDeg = 4
    ElseIf (dblHinDegT = 0) And (Abs(dblHinDegY) = 4) Then
        intHinDeg = 4
    Else
        intHinDeg = 0
    End If

    ' WF�������v�Z����
    If SEEDDEG = intHinDeg Then
        ' �ʏ�i�̏ꍇ
        intWFcnt = Format(((BlkLen - Loss0) / Pitch) + 0.4, "0")
    Else
        intWFcnt = Format(((BlkLen * Mlt4 - Loss4) / Pitch) + 0.4, "0")
    End If
    If intWFcnt < 0 Then intWFcnt = 0
    calculateWfNum = intWFcnt
End Function

'*******************************************************************************
'*    �֐���        : XSDC3Proc2
'*
'*    �����T�v      : 1.�H�����ѓo�^�������s��(�݌Ɍ����FCW740,CW760�p)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^        ,����
'*�@�@�@�@�@�@�@�@�@�@�Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC3Proc2() As FUNCTION_RETURN
    ' �����ϐ�
    Dim i, j, k         As Integer
    Dim intRtn          As Integer          ' ���A���
    Dim sSQL            As String           ' �r�p�k
    Dim rs              As OraDynaset       ' ���R�[�h�Z�b�g
    Dim sSqlWhere       As String           ' WHERE��
    Dim sErrMsg         As String           ' �G���[���b�Z�[�W
    Dim udtKoutei       As typ_XSDC3_Update ' �H������
    Dim rsKCNTC         As OraDynaset       ' ���R�[�h�Z�b�g

    Dim udtWSTOCKINFO() As typ_stock_info   ' ���ݍH���̏��
    Dim vGetData        As Variant          ' ��ʎ捞�pwork
    Dim sOldHinban      As String           ' ���i��
    Dim sNowHinban      As String           ' ���i��
    Dim vBlkId          As Variant          ' ��ʎ捞�pwork
    Dim sOldBlkID       As String           ' ���u���b�NID
    Dim vREVNUM         As Variant          ' ��ʎ捞�pwork
    Dim vFACTORY        As Variant          ' ��ʎ捞�pwork
    Dim vOPE            As Variant          ' ��ʎ捞�pwork
    Dim intREVNUM       As Integer          ' ���i�����ԍ�
    Dim sFACTORY        As String           ' �H��
    Dim sOPE            As String           ' ���Ə���
    Dim sBlkId          As String           ' �u���b�NID

    Dim intMapSt        As Integer          ' �}�b�v�J�n�ʒu
    Dim intMapEd        As Integer          ' �}�b�v�I���ʒu
    Dim blHinFlg        As Boolean          ' �i�Ԕ�r�p�t���O
    Dim lngTMaisu       As Long             ' ���v����
    Dim intGetHinInpos  As Integer          ' �������ʒu
    Dim objGamenSpd     As Object           ' ���ID
    Dim intHantei       As Integer

    ' �G���[�n���h���̐ݒ�
    On Error GoTo 0

    ' �����ݒ�
    XSDC3Proc2 = FUNCTION_RETURN_FAILURE

    ReDim STOCKINFO(0)
    ReDim udtWSTOCKINFO(0)

   ' �O�H���������v
    For i = 0 To Kihon.CNTHINOLD - 1
        If (Kihon.NOWPROC = "CW760") _
           And ((SIngotP > CLng(HinOld(i).INPOSCA)) Or (HinOld(i).INPOSCA >= EIngotP)) Then
            ' �����Ȃ�
        Else
            ' �O�H��������0�̎��A�����I��
            If HinOld(i).GNMCA <= 0 Then
                XSDC3Proc2 = FUNCTION_RETURN_SUCCESS
                Exit Function
            End If

            ReDim Preserve STOCKINFO(UBound(STOCKINFO) + 1)  ' �z��̒ǉ�

            ' �s�Ǥ�����̏����ݒ�
            STOCKINFO(UBound(STOCKINFO)).hinban = HinOld(i).HINBCA
            STOCKINFO(UBound(STOCKINFO)).GENZAL = CLng(HinOld(i).GNLCA)
            STOCKINFO(UBound(STOCKINFO)).FURYOL = 0
            STOCKINFO(UBound(STOCKINFO)).HARAIL = CLng(HinOld(i).GNLCA)
            STOCKINFO(UBound(STOCKINFO)).GENZAW = CLng(HinOld(i).GNWCA)
            STOCKINFO(UBound(STOCKINFO)).FuryoW = 0
            STOCKINFO(UBound(STOCKINFO)).HARAIW = CLng(HinOld(i).GNWCA)
            STOCKINFO(UBound(STOCKINFO)).GENZAM = CLng(HinOld(i).GNMCA)
            STOCKINFO(UBound(STOCKINFO)).FURYOM = 0
            STOCKINFO(UBound(STOCKINFO)).HARAIM = CLng(HinOld(i).GNMCA)
            STOCKINFO(UBound(STOCKINFO)).KCKNT = CLng(HinOld(i).KCKNTCA)
        End If
    Next i

    ' �Ŕ����w����ʂ���i�Ԃ̕����o���ƌ������}�b�v�ʒu���ڂ��狁�߂�
    ' STOCKINFO�z��Ɋi�[���邪STOCKINFO�̕i�Ԃ�HinOld�̕i�Ԃ̓o�^�����ƈ�v���Ă���Ƃ͌���Ȃ�
    If Kihon.NOWPROC = "CW740" Then
        Set objGamenSpd = f_cmbc036_2.sprExamine    ' �����ύX���گ��
    Else
        Set objGamenSpd = f_cmbc039_3.sprExamine    ' �Ĕ������گ��
    End If

    ' �i�Ԃ�1��ǉ��������Ƃɂ���̕ύX
    With objGamenSpd
        bMapErrFlg = False                          ' WFϯ�߈ʒu�����׸ޏ�����

        ' �G�s��s�]���ǉ��Ή�
        .GetText 39, 1, vBlkId                      ' �u���b�NID
        sOldBlkID = CStr(Trim(vBlkId))
        For i = 1 To .MaxRows Step 2                ' ���گ�ނ����ް������(2�s���m�F)
            ' �G�s��s�]���ǉ��Ή�
            .GetText 40, i, vGetData                ' �Â��i�Ԏ擾
            sOldHinban = Trim(CStr(vGetData))
            .GetText 2, i, vGetData                 ' �V�����i�Ԏ擾

            ' �i�Ԃ�"Z"�̎��͐V�i��=���i��
            If Trim(CStr(vGetData)) = "Z" Then
                sNowHinban = sOldHinban
            Else
                sNowHinban = Trim(CStr(vGetData))
            End If
            .GetText 5, i, vGetData                 ' �����ʒu
            intGetHinInpos = val(vGetData)
            .GetText 6, i, vGetData                 ' �}�b�v�J�n�ʒu
            intMapSt = val(vGetData)
            .GetText 6, i + 1, vGetData             ' �}�b�v�I���ʒu
            intMapEd = val(vGetData)

            ' �G�s��s�]���ǉ��Ή�
            .GetText 39, i, vBlkId                  ' �u���b�NID
            If vBlkId = "" Then                     ' �u���b�N��NULL��������A�O��̃u���b�N���g�p
                vBlkId = Mid(BlkNow.CRYNUMC2, 1, 9) & sOldBlkID
            Else
                vBlkId = Mid(BlkNow.CRYNUMC2, 1, 9) & vBlkId
            End If

            ' �Ǖi�����̓}�b�v�ʒu�ł͂Ȃ��e�[�u������擾����
            sBlkId = vBlkId
            intREVNUM = gtSprWfMap(i).REVNUM        ' ���i�����ԍ�
            sFACTORY = gtSprWfMap(i).factory        ' �H��
            sOPE = gtSprWfMap(i).opecond            ' ���Ə���

            If ((Kihon.NOWPROC = "CW760") Or (Kihon.NOWPROC = "CW740")) And (vBlkId <> BlkNow.CRYNUMC2) Then
                ' �����Ȃ�
            Else
                ' SXL�͈͊O�ʹװ�Ƃ���
'                If (Kihon.NOWPROC = "CW760") And ((SIngotP > intGetHinInpos) Or (intGetHinInpos >= EIngotP)) Then
'2010/05/10 Change Y.Hitomi �ŏI�}�b�v�ʒu�ɘa�i999.9��1000�j�Ή�
                If (Kihon.NOWPROC = "CW760") And ((SIngotP > intGetHinInpos) Or (intGetHinInpos > EIngotP)) Then
                    bMapErrFlg = True
                End If

                ' �Ǖi�����̓}�b�v�ʒu�ł͂Ȃ��e�[�u������擾����
                sSQL = "SELECT COUNT(*) AS SXLCNT"
                sSQL = sSQL & " FROM TBCMY011 "
                sSQL = sSQL & " WHERE LOTID = '" & sBlkId & "'"
                sSQL = sSQL & " AND (WFSTA ='0' OR WFSTA = '1') "
                sSQL = sSQL & " AND BLOCKSEQ >= " & intMapSt & ""
                sSQL = sSQL & " AND BLOCKSEQ <= " & intMapEd & ""

                Debug.Print sSQL

                Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)

                ' �݂���Ȃ�������G���[
                If rs.RecordCount = 0 Then
                    SXLCnt = 0
                Else ' ����������A�Ǖi�������擾����
                    SXLCnt = val(rs("SXLCNT"))
                End If
                Debug.Print SXLCnt

                blHinFlg = False ' �����̔z��ɓ����i�Ԃ��o�^����Ă��邩���׸�

                ' udtWSTOCKINFO()��1����J�n
                For j = 1 To UBound(udtWSTOCKINFO)
                    If (udtWSTOCKINFO(j).hinban = sOldHinban) Then  ' ���ɓo�^���Ă���i��
                        blHinFlg = True
                        udtWSTOCKINFO(j).HARAIM = udtWSTOCKINFO(j).HARAIM + SXLCnt
                    End If
                Next j

                If (blHinFlg = False) Then   ' udtWSTOCKINFO()�̔z��ɕi�Ԃ��o�^�Ȃ��������V�K��wSTOCKINFO()�ɓo�^
                    ReDim Preserve udtWSTOCKINFO(UBound(udtWSTOCKINFO) + 1)     ' �z��̒ǉ�
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).hinban = sOldHinban    ' �i��
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).HARAIM = 0             ' �z�񏉊��ݒ�
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).FURYOM = 0             ' �z�񏉊��ݒ�
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).HARAIM = SXLCnt        ' ��ʂ��ް�
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).REVNUM = intREVNUM     ' ��ʂ��ް�
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).factory = sFACTORY     ' ��ʂ��ް�
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).OPE = sOPE             ' ��ʂ��ް�
                End If
            End If
        Next i
    End With

    ' STOCKINF��wSTOCKINF�̓ˍ���������STOCKINF�ɕi�ԁA�����A�d�ʁA�����̕����o���ƕs�ǂ��i�[
    ' HinOld�ɂȂ��f�[�^�͂Ȃ��Ȃ�
    ' STOCKINFO()�͓Y��0����J�n
    ' STOCKINFO()�̕i�Ԃ�HinOld���ް��A�ް���HinNow���ް�
    For i = 1 To UBound(STOCKINFO)
        STOCKINFO(i).HARAIM = 0
        STOCKINFO(i).FURYOM = 0
        For j = 1 To UBound(udtWSTOCKINFO)
            If (STOCKINFO(i).hinban = udtWSTOCKINFO(j).hinban) Then    ' �i�Ԃ����������o�^����
                STOCKINFO(i).HARAIM = udtWSTOCKINFO(j).HARAIM
                STOCKINFO(i).FURYOM = udtWSTOCKINFO(j).FURYOM
                STOCKINFO(i).REVNUM = udtWSTOCKINFO(j).REVNUM
                STOCKINFO(i).factory = udtWSTOCKINFO(j).factory
                STOCKINFO(i).OPE = udtWSTOCKINFO(j).OPE
                lngTMaisu = udtWSTOCKINFO(j).HARAIM + udtWSTOCKINFO(j).FURYOM   ' �����̍��v
                If (lngTMaisu > 0) Then   ' ���������邩�m�F
                    STOCKINFO(i).HARAIW = udtWSTOCKINFO(j).HARAIM / lngTMaisu * CLng(STOCKINFO(i).HARAIW)   ' �s�ǁA���o�͖����̔䗦�ŎZ�o
                    STOCKINFO(i).FuryoW = CLng(STOCKINFO(i).GENZAW) - STOCKINFO(i).HARAIW
                    STOCKINFO(i).HARAIL = udtWSTOCKINFO(j).HARAIM / lngTMaisu * CLng(STOCKINFO(i).HARAIL)   ' �s�ǁA���o�͖����̔䗦�ŎZ�o
                    STOCKINFO(i).FURYOL = CLng(STOCKINFO(i).GENZAL) - STOCKINFO(i).HARAIL
                 End If
            End If
        Next j
    Next i

    ' �s�ǂ�����ꍇ���݌Ɍ����̍쐬
    For i = 1 To UBound(STOCKINFO)
        If STOCKINFO(i).FURYOM > 0 Then
            udtKoutei.CRYNUMC3 = HinNow(0).CRYNUMCA     ' �u���b�N�h�c
            giInpos = giInpos + 1
            udtKoutei.INPOSC3 = giInpos                 ' �ʒu
            udtKoutei.KCNTC3 = STOCKINFO(i).KCKNT + 1   ' �H���A��
            udtKoutei.HINBC3 = STOCKINFO(i).hinban      ' �i��
            udtKoutei.REVNUMC3 = STOCKINFO(i).REVNUM    ' ���i�����ԍ�
            udtKoutei.FACTORYC3 = STOCKINFO(i).factory  ' �H��
            udtKoutei.OPEC3 = STOCKINFO(i).OPE          ' ���Ə���
            udtKoutei.LENC3 = STOCKINFO(i).HARAIL       ' ����
            udtKoutei.XTALC3 = HinNow(0).XTALCA         ' �����ԍ�
            udtKoutei.SXLIDC3 = ""                      ' SXLID

            udtKoutei.KNKTC3 = left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 1)   ' �Ǘ��H��(���ݍH��+1)
            udtKoutei.WKKTC3 = Kihon.NOWPROC            ' �H��
            udtKoutei.WKKBC3 = ""                       ' ��Ƌ敪
            udtKoutei.MACOC3 = HinNow(0).NEMACOCA       ' ������
            udtKoutei.MODKBC3 = ""                      ' �ԍ��敪
            udtKoutei.SUMKBC3 = ""                      ' �W�v�敪
            udtKoutei.FRKNKTC3 = ""                     ' (���)�Ǘ��H��

            If IsNull(HinOld(0).NEWKNTCA) = True Then   '(����j�H��
                udtKoutei.FRWKKTC3 = ""
            Else
                udtKoutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
            End If

            udtKoutei.FRWKKBC3 = ""                     ' (���)��Ƌ敪

            If IsNull(HinOld(0).NEMACOCA) = True Then   '�i����j������
                udtKoutei.FRMACOC3 = "0"
            Else
                udtKoutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If

            Select Case Kihon.NOWPROC
                Case "CC730"
                    intHantei = CInt(BlkNow.GNLC2)
                Case Else
                    intHantei = CInt(BlkNow.GNMC2)
            End Select

            If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                udtKoutei.TOWKKTC3 = " "                            ' (���o)�H��
                udtKoutei.TOMACOC3 = "0"                            ' (���o)������
            Else
                udtKoutei.TOWKKTC3 = HinNow(0).GNWKNTCA             ' (���o)�H��
                udtKoutei.TOMACOC3 = HinNow(0).GNMACOCA             ' (���o)������
            End If

            udtKoutei.FRLC3 = STOCKINFO(i).GENZAL                   ' �������
            udtKoutei.FRWC3 = STOCKINFO(i).GENZAW                   ' ����d��
            udtKoutei.FRMC3 = STOCKINFO(i).GENZAM                   ' �������
            udtKoutei.FULC3 = STOCKINFO(i).FURYOL                   ' �s�ǒ���
            udtKoutei.FUWC3 = STOCKINFO(i).FuryoW                   ' �s�Ǐd��
            udtKoutei.FUMC3 = STOCKINFO(i).FURYOM                   ' �s�ǖ���
            udtKoutei.LOSWC3 = ""                                   ' ���X����

            udtKoutei.LOSLC3 = ""                                   ' ���X�d��
            udtKoutei.LOSMC3 = ""                                   ' ���X����
            udtKoutei.TOLC3 = STOCKINFO(i).HARAIL                   ' ���o����
            udtKoutei.TOWC3 = STOCKINFO(i).HARAIW                   ' ���o�d��
            udtKoutei.TOMC3 = STOCKINFO(i).HARAIM                   ' ���o����
            udtKoutei.SUMITLC3 = ""                                 ' SUMIT����
            udtKoutei.SUMITWC3 = ""                                 ' SUMIT�d��
            udtKoutei.SUMITMC3 = ""                                 ' SUMIT����
            udtKoutei.MOTHINC3 = ""                                 ' �U�֕i��(��)
            udtKoutei.XTWORKC3 = "42"                               ' �����H��

            udtKoutei.WFWORKC3 = ""                                 ' ���ʐ���
            udtKoutei.HOLDCC3 = " "                                 ' �z�[���h�R�[�h
            udtKoutei.HOLDBC3 = "0"                                 ' �z�[���h�敪
            udtKoutei.LDFRCC3 = ""                                  ' �i���R�[�h
            udtKoutei.LDFRBC3 = "0"                                 ' �i���敪�i�n�C�L�j
            udtKoutei.TSTAFFC3 = Kihon.STAFFID                      ' �o�^�Ј�ID
            udtKoutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �o�^���t

            udtKoutei.KSTAFFC3 = ""                                 ' �X�V�Ј�ID
            udtKoutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
            udtKoutei.SUMITBC3 = ""                                 ' SUMIT���M�t���O
            udtKoutei.SNDKC3 = ""                                   ' ���M�t���O
            udtKoutei.MODMACOC3 = ""                                ' �ԍ��̏�����
            udtKoutei.KAKUCC3 = ""                                  ' �m��R�[�h
            udtKoutei.SUMDAYC3 = CalcSumcoTime(udtKoutei.KDAYC3)    ' SUMCO����
            udtKoutei.PAYCLASSC3 = ""                               ' �]����H��t���O
            udtKoutei.PLANTCATC3 = HinNow(0).PLANTCATCA             ' ����
            intRtn = CreateXSDC3(udtKoutei, sErrMsg)                ' �H�����тɍ݌Ɍ����o�^
            If intRtn = FUNCTION_RETURN_FAILURE Then                ' �H�����ђǉ��G���[
                MsgBox sErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : XSDC3Proc3
'*
'*    �����T�v      : 1.�H�����ѓo�^�������s��(�i�ԐU�֏��FCW740,CW760�p)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC3Proc3() As FUNCTION_RETURN
    ' �����ϐ�
    Dim i, j            As Integer
    Dim intRtn          As Integer          ' ���A���
    Dim sSQL            As String           ' �r�p�k
    Dim sSqlWhere       As String           ' WHERE��
    Dim sErrMsg         As String
    Dim udtKoutei       As typ_XSDC3_Update ' �H������

    Dim lngLen          As Long
    Dim lngCHKPOS       As Long

    Dim udtWOINF()      As typ_trans_info   ' �O�i�ԕ��ёւ��p
    Dim udtWNINF()      As typ_trans_info   ' ��i�ԕ��ёւ��p
    Dim udtWWINF()      As typ_trans_info   ' ���ёւ��p���[�N
    Dim intBuf          As Integer
    Dim intOINFrecCnt   As Integer
    Dim intNINFrecCnt   As Integer
    Dim intOINFFLG      As Integer
    Dim intNINFFLG      As Integer
    Dim intCnt          As Integer
    Dim intWNINFMAX     As Integer
    Dim intWOINFMAX     As Integer
    Dim intHantei       As Integer

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    ' �����ݒ�
    XSDC3Proc3 = FUNCTION_RETURN_FAILURE

    ReDim udtWOINF(UBound(STOCKINFO))      ' �i�Ԃ��ƂɃ\�[�g�p
    ReDim udtWNINF(Kihon.CNTHINNOW)        ' �i�Ԃ��ƂɃ\�[�g�p
    ReDim udtWWINF(1)                      ' �i�Ԃ��ƂɃ\�[�g�p

    ' �݌Ɍ�������荞��
    For i = 1 To UBound(STOCKINFO)
        If Trim(STOCKINFO(i).hinban) <> "" Then
            udtWOINF(i).hinban = STOCKINFO(i).hinban
            udtWOINF(i).LEN = STOCKINFO(i).HARAIL       ' �O�H���������v
            udtWOINF(i).WAT = STOCKINFO(i).HARAIW       ' �O�H���d�ʍ��v
            udtWOINF(i).MAI = STOCKINFO(i).HARAIM       ' �O�H���������v
        End If
    Next i

    ' �Ǖi������荞��
    For j = 0 To Kihon.CNTHINNOW - 1
        If (Kihon.NOWPROC = "CW760") _
            And ((SIngotP > HinNow(j).INPOSCA) Or (HinNow(j).INPOSCA >= EIngotP)) Then
                ' �����Ȃ�
        Else
            udtWNINF(j).hinban = HinNow(j).HINBCA
            udtWNINF(j).REVNUM = HinNow(j).REVNUMCA
            udtWNINF(j).factory = HinNow(j).FACTORYCA
            udtWNINF(j).OPE = HinNow(j).OPECA
            udtWNINF(j).LEN = HinNow(j).GNLCA           ' ��H���������v
            udtWNINF(j).WAT = HinNow(j).GNWCA           ' ��H���d�ʍ��v
            udtWNINF(j).MAI = HinNow(j).GNMCA           ' ��H���������v
            udtWNINF(j).KCKNT = HinNow(j).KCKNTCA       ' ��H���A��
        End If
    Next j

    ' �����i�ԓ��m�A��������ł�����
    For i = 1 To UBound(STOCKINFO)
        For j = 0 To Kihon.CNTHINNOW - 1
           If udtWOINF(i).hinban = udtWNINF(j).hinban Then
                ' ���������̐����𗼕��������
                If udtWOINF(i).MAI <= udtWNINF(j).MAI Then
                    udtWNINF(j).LEN = udtWNINF(j).LEN - udtWOINF(i).LEN
                    udtWNINF(j).WAT = udtWNINF(j).WAT - udtWOINF(i).WAT
                    udtWNINF(j).MAI = udtWNINF(j).MAI - udtWOINF(i).MAI
                    udtWOINF(i).LEN = 0
                    udtWOINF(i).WAT = 0
                    udtWOINF(i).MAI = 0
                Else
                    udtWOINF(i).LEN = udtWOINF(i).LEN - udtWNINF(j).LEN
                    udtWOINF(i).WAT = udtWOINF(i).WAT - udtWNINF(j).WAT
                    udtWOINF(i).MAI = udtWOINF(i).MAI - udtWNINF(j).MAI
                    If udtWOINF(i).MAI < 0 Then
                        udtWOINF(i).MAI = 0
                    End If
                    udtWNINF(j).LEN = 0
                    udtWNINF(j).WAT = 0
                    udtWNINF(j).MAI = 0
                End If
            End If
        Next
    Next

    For i = 0 To UBound(udtWOINF) - 2
        For j = i + 1 To UBound(udtWOINF) - 1
            If (StrComp(udtWOINF(i).hinban, udtWOINF(j).hinban, _
                vbTextCompare)) = 1 Then ' �i�Ԃ̓��֕K�v
                udtWWINF(0) = udtWOINF(j)
                udtWOINF(j) = udtWOINF(i)
                udtWOINF(i) = udtWWINF(0)
            End If
        Next j
    Next i

    ' wNINF�̕i�Ԃ��\�[�g����
    For i = 0 To UBound(udtWNINF) - 2
        For j = i + 1 To UBound(udtWNINF) - 1
            If (StrComp(udtWNINF(i).hinban, udtWNINF(j).hinban, _
                vbTextCompare)) = 1 Then ' �i�Ԃ̓��֕K�v
                udtWWINF(0) = udtWNINF(j)
                udtWNINF(j) = udtWNINF(i)
                udtWNINF(i) = udtWWINF(0)
            End If
        Next j
    Next i

    ' �󂫂̔z��폜����(�z��̃f�[�^���l�߂�)
    For i = 0 To intWOINFMAX
        If udtWOINF(i).MAI <= 0 Then
            intCnt = i
            Call HairetuOpe_Mai(udtWOINF(), intCnt, -1)
        End If
    Next i

    ' �󂫂̔z��폜����(�z��̃f�[�^���l�߂�)
    For i = 0 To intWNINFMAX
        If udtWNINF(i).MAI <= 0 Then
            intCnt = i
            Call HairetuOpe_Mai(udtWNINF(), intCnt, -1)
        End If
    Next i

    ' �i�ԓ��֏����쐬����
    i = 0 ' �O�i�Ԃ̈ʒu
    j = 0 ' ��i�Ԃ̈ʒu
    Do
        ' ������˂����킹�Đ��ʂ������łȂ�������傫���l�̕i�Ԃ𕪊�����
        If (udtWOINF(i).MAI = udtWNINF(j).MAI) Then         ' �i�Ԓ����������������Ƃ����ɐi��
        ElseIf (udtWOINF(i).MAI > udtWNINF(j).MAI) Then     ' �i�Ԓ������قȂ鎞
            intCnt = i
            Call HairetuOpe(udtWOINF(), intCnt, 1)          ' �z��̒ǉ�
            udtWOINF(i + 1).hinban = udtWOINF(i).hinban
            udtWOINF(i + 1).LEN = udtWOINF(i).LEN - udtWNINF(j).LEN
            udtWOINF(i + 1).WAT = udtWOINF(i).WAT - udtWNINF(j).WAT
            udtWOINF(i + 1).MAI = udtWOINF(i).MAI - udtWNINF(j).MAI
            udtWOINF(i).LEN = udtWNINF(j).LEN
            udtWOINF(i).WAT = udtWNINF(j).WAT
            udtWOINF(i).MAI = udtWNINF(j).MAI
        ElseIf (udtWOINF(i).MAI < udtWNINF(j).MAI) Then     ' �i�Ԑ��ʂ��قȂ鎞
            intCnt = j
            Call HairetuOpe(udtWNINF(), intCnt, 1)
            udtWNINF(j + 1).hinban = udtWNINF(i).hinban
            udtWNINF(j + 1).LEN = udtWNINF(j).LEN - udtWOINF(i).LEN
            udtWNINF(j + 1).WAT = udtWNINF(j).WAT - udtWOINF(i).WAT
            udtWNINF(j + 1).MAI = udtWNINF(j).MAI - udtWOINF(i).MAI
            udtWNINF(j).LEN = udtWOINF(i).LEN
            udtWNINF(j).WAT = udtWOINF(i).WAT
            udtWNINF(j).MAI = udtWOINF(i).MAI
        End If
        intOINFrecCnt = UBound(udtWOINF())
        intNINFrecCnt = UBound(udtWNINF())
        i = i + 1
        j = j + 1
        If (i > intOINFrecCnt) Then
            Exit Do

        End If
        If (j > intNINFrecCnt) Then
            Exit Do
        End If
        If (udtWOINF(i).MAI) <= 0 Then
            Exit Do
        End If
        If (udtWNINF(j).MAI) <= 0 Then
            Exit Do
        End If
    Loop

    intOINFrecCnt = UBound(udtWOINF())
    For i = 0 To intOINFrecCnt
        If (StrComp(udtWNINF(i).hinban, udtWOINF(i).hinban, vbTextCompare) <> 0) Then  ' �i�Ԃ��قȂ鎞�U�֏��ɓo�^����
            If Trim(udtWNINF(i).hinban) <> "" And udtWNINF(i).LEN > 0 Then
                ' SXL�͈͊O�ʹװ�Ƃ���
                If bMapErrFlg Then
                    MsgBox "����P�`�F�b�N�G���[ (SXL�ʒu �F " & SIngotP & "�|" & EIngotP & ")"
                    Exit Function
                End If

                udtKoutei.CRYNUMC3 = HinNow(0).CRYNUMCA     ' �u���b�N�h�c
                giInpos = giInpos + 1
                udtKoutei.INPOSC3 = giInpos                 ' �ʒu
                udtKoutei.KCNTC3 = udtWNINF(i).KCKNT        ' �H���A��
                udtKoutei.HINBC3 = udtWNINF(i).hinban       ' �i��
                udtKoutei.REVNUMC3 = udtWNINF(i).REVNUM     ' ���i�����ԍ�
                udtKoutei.FACTORYC3 = udtWNINF(i).factory   ' �H��
                udtKoutei.OPEC3 = udtWNINF(i).OPE           ' ���Ə���
                udtKoutei.LENC3 = udtWNINF(i).LEN           ' �������
                udtKoutei.XTALC3 = HinNow(0).XTALCA         ' �����ԍ�
                udtKoutei.SXLIDC3 = ""                      ' SXLID

                udtKoutei.KNKTC3 = left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
                  CStr(CInt(Right(Kihon.NOWPROC, 1)) + 2)   ' �Ǘ��H��(���ݍH��+2)
                udtKoutei.WKKTC3 = Kihon.NOWPROC            ' �H��
                udtKoutei.WKKBC3 = ""                       ' ��Ƌ敪
                udtKoutei.MACOC3 = HinNow(0).NEMACOCA       ' ������
                udtKoutei.MODKBC3 = ""                      ' �ԍ��敪
                udtKoutei.SUMKBC3 = ""                      ' �W�v�敪
                udtKoutei.FRKNKTC3 = ""                     ' (���)�Ǘ��H��
                If IsNull(HinOld(0).NEWKNTCA) = True Then   '(����j�H��
                    udtKoutei.FRWKKTC3 = ""
                Else
                    udtKoutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
                End If
                udtKoutei.FRWKKBC3 = ""                     ' (���)��Ƌ敪
                If IsNull(HinOld(0).NEMACOCA) = True Then   '�i����j������
                    udtKoutei.FRMACOC3 = "0"
                Else
                    udtKoutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
                End If

                Select Case Kihon.NOWPROC
                    Case "CC730"
                        intHantei = CInt(BlkNow.GNLC2)
                    Case Else
                        intHantei = CInt(BlkNow.GNMC2)
                End Select

                If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                    udtKoutei.TOWKKTC3 = " "                            ' (���o)�H��
                    udtKoutei.TOMACOC3 = "0"                            '(���o)������
                Else
                    udtKoutei.TOWKKTC3 = HinNow(0).GNWKNTCA             ' (���o)�H��
                    udtKoutei.TOMACOC3 = HinNow(0).GNMACOCA             ' (���o)������
                End If
                udtKoutei.FRLC3 = udtWNINF(i).LEN                       '�������
                udtKoutei.FRWC3 = udtWNINF(i).WAT                       '����d��
                udtKoutei.FRMC3 = udtWNINF(i).MAI                       '�������
                udtKoutei.FULC3 = 0                                     '�s�ǒ���
                udtKoutei.FUWC3 = 0                                     '�s�Ǐd��
                udtKoutei.FUMC3 = 0                                     '�s�ǖ���
                udtKoutei.LOSWC3 = ""                                   ' ���X����

                udtKoutei.LOSLC3 = ""                                   ' ���X�d��
                udtKoutei.LOSMC3 = ""                                   ' ���X����
                udtKoutei.TOLC3 = udtWNINF(i).LEN                       '���o����
                udtKoutei.TOWC3 = udtWNINF(i).WAT                       '���o�d��
                udtKoutei.TOMC3 = udtWNINF(i).MAI                       '���o����
                udtKoutei.SUMITLC3 = ""                                 ' SUMIT����
                udtKoutei.SUMITWC3 = ""                                 ' SUMIT�d��
                udtKoutei.SUMITMC3 = ""                                 ' SUMIT����
                udtKoutei.MOTHINC3 = udtWOINF(i).hinban                 '���i��
                udtKoutei.XTWORKC3 = "42"                               ' �����H��

                udtKoutei.WFWORKC3 = ""                                 ' ���ʐ���
                udtKoutei.HOLDCC3 = " "                                 ' �z�[���h�R�[�h
                udtKoutei.HOLDBC3 = "0"                                 ' �z�[���h�敪
                udtKoutei.LDFRCC3 = ""                                  ' �i���R�[�h
                udtKoutei.LDFRBC3 = "0"                                 ' �i���敪�i�n�C�L�j
                udtKoutei.TSTAFFC3 = Kihon.STAFFID                      ' �o�^�Ј�ID
                udtKoutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �o�^���t

                udtKoutei.KSTAFFC3 = ""                                 ' �X�V�Ј�ID
                udtKoutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
                udtKoutei.SUMITBC3 = ""                                 ' SUMIT���M�t���O
                udtKoutei.SNDKC3 = ""                                   ' ���M�t���O
                udtKoutei.MODMACOC3 = ""                                ' �ԍ��̏�����
                udtKoutei.KAKUCC3 = ""                                  ' �m��R�[�h
                udtKoutei.SUMDAYC3 = CalcSumcoTime(udtKoutei.KDAYC3)    ' SUMCO����
                udtKoutei.PAYCLASSC3 = ""                               ' �]����H��t���O
                udtKoutei.PLANTCATC3 = HinNow(0).PLANTCATCA             ' ����

                intRtn = CreateXSDC3(udtKoutei, sErrMsg)                ' �H�����тɍ݌Ɍ����o�^
                If intRtn = FUNCTION_RETURN_FAILURE Then                ' �H�����ђǉ��G���[
                    MsgBox sErrMsg
                    Exit Function
                End If
            End If
        End If
    Next i

    XSDC3Proc3 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC3Proc3 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : XSDC3Proc4
'*
'*    �����T�v      : 1.�H�����ѓo�^�������s��(�݌Ɍ����FCC730�p)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC3Proc4() As FUNCTION_RETURN
    ' �����ϐ�
    Dim i, j                    As Integer
    Dim intRtn                  As Integer          ' ���A���
    Dim sSQL                    As String           ' �r�p�k
    Dim rs                      As OraDynaset       ' ���R�[�h�Z�b�g
    Dim sSqlWhere               As String           ' WHERE��
    Dim sErrMsg                 As String
    Dim udtKoutei               As typ_XSDC3_Update ' �H������
    Dim rsKCNTC                 As OraDynaset       ' ���R�[�h�Z�b�g
    Dim intNextCnt              As Integer
    Dim lngLen                  As Long
    Dim lngCHKPOS               As Long
    Dim intBadcnt               As Integer
    Dim udtBADINFO()            As typ_bad_info
    Dim udtWSTOCKINFO()         As typ_stock_info
    Dim intLoopCnt              As Integer
    Dim vGetMaxPos              As Variant
    Dim vGetData                As Variant
    Dim sOldHinban, sNowHinban  As String
    Dim intSXLCnt               As Integer
    Dim intMapSt, intMapEd      As Integer
    Dim blHinFlg                As Boolean
    Dim lngTMaisu               As Long
    Dim vNukisiFlg              As Variant
    Dim intHantei               As Integer

    ' �G���[�n���h���̐ݒ�
    On Error GoTo 0

    ' �����ݒ�
    XSDC3Proc4 = FUNCTION_RETURN_FAILURE

    ReDim STOCKINFO(Kihon.CNTHINOLD)
    ReDim udtWSTOCKINFO(0)

    ' HinOld����O�H������,�d��,�������v�擾(������0)
    For i = 0 To Kihon.CNTHINOLD - 1
        FRLC3Sum = FRLC3Sum + CLng(HinOld(i).GNLCA)     ' �O�H���������v
        FRWC3Sum = FRWC3Sum + CLng(HinOld(i).GNWCA)     ' �O�H���d�ʍ��v
        FRMC3Sum = FRMC3Sum + CLng(HinOld(i).GNMCA)     ' �O�H���������v

        ' �s�Ǥ�����̏����ݒ�
        STOCKINFO(i).hinban = HinOld(i).HINBCA
        STOCKINFO(i).FURYOL = 0
        STOCKINFO(i).HARAIL = CLng(HinOld(i).GNLCA)
        STOCKINFO(i).FuryoW = CLng(HinOld(i).GNWCA)     ' �s�Ǐd�ʂɕ����d�ʂ����ɑ�����Č�Ōv�Z����
        STOCKINFO(i).HARAIW = CLng(HinOld(i).GNWCA)
        STOCKINFO(i).FURYOM = CLng(HinOld(i).GNMCA)     ' �s�ǖ����ɕ������������ɑ������Ōv�Z����
        STOCKINFO(i).HARAIM = CLng(HinOld(i).GNMCA)
        STOCKINFO(i).KCKNT = CLng(HinOld(i).KCKNTCA)
        STOCKINFO(i).REVNUM = HinOld(i).REVNUMCA        ' ���i�����ԍ�
        STOCKINFO(i).factory = HinOld(i).FACTORYCA      ' �H��
        STOCKINFO(i).OPE = HinOld(i).OPECA              ' ���i�����ԍ�
    Next i

    ' �Ŕ����w����ʂ���i�Ԃ̕����o���ƌ������}�b�v�ʒu���ڂ��狁�߂�
    ' STOCKINFO�z��Ɋi�[���邪STOCKINFO�̕i�Ԃ�HinOld�̕i�Ԃ̓o�^�����ƈ�v���Ă���Ƃ͌���Ȃ�
    intBadcnt = 0  ' �s�ǐ������ݒ�

    ' �s�ǂ��擪�ɂȂ����m�F
    If ((CLng(HinNow(0).INPOSCA) - CLng(HinOld(0).INPOSCA)) > 0) Then ' �O��J�n�ʒu���r���č�������Εs�ǈʒu�o�^
        intBadcnt = intBadcnt + 1
        ReDim Preserve udtBADINFO(intBadcnt)
        udtBADINFO(intBadcnt).pos = CLng(HinOld(0).INPOSCA)
        udtBADINFO(intBadcnt).LEN = CLng(HinNow(0).INPOSCA) - CLng(HinOld(0).INPOSCA)
    End If

    ' �s�ǒ������i�ԊԂɂȂ����m�F
    For i = 0 To Kihon.CNTHINNOW - 2
        If (CLng(HinNow(i + 1).INPOSCA) > (CLng(HinNow(i).INPOSCA) + CLng(HinNow(i).GNLCA))) Then ' �i�ԊԂɕs�ǗL
            intBadcnt = intBadcnt + 1 ' �s�ǈʒu�̓o�^
            ReDim Preserve udtBADINFO(intBadcnt)
            udtBADINFO(intBadcnt).pos = CLng(HinNow(i).INPOSCA) + CLng(HinNow(i).GNLCA)
            udtBADINFO(intBadcnt).LEN = CLng(HinNow(i + 1).INPOSCA) - CLng(HinNow(i).INPOSCA) - CLng(HinNow(i).GNLCA)
        End If
    Next i

    ' �s�ǂ��Ō�ɂȂ����O�̊m�F(�������J�n�ʒu+�����Ŕ�r)
    If ((CLng(HinOld(Kihon.CNTHINOLD - 1).INPOSCA) + CLng(HinOld(Kihon.CNTHINOLD - 1).GNLCA)) _
        <> (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))) Then ' �I���ʒu�̊m�F
        intBadcnt = intBadcnt + 1
        ReDim Preserve udtBADINFO(intBadcnt)
        udtBADINFO(intBadcnt).pos = (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))
        udtBADINFO(intBadcnt).LEN = CLng(HinOld(Kihon.CNTHINOLD - 1).INPOSCA) + CLng(HinOld(Kihon.CNTHINOLD - 1).GNLCA) - (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))
    End If

    If (intBadcnt = 0) Then  ' �O�ƌ�ŕs�ǂȂ� �����I��
        XSDC3Proc4 = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    ' �s�ǈʒu��U�֑O�̌����ʒu���m�F���ĕs�ǈʒu�ɑ�������i�Ԃ�o�^����
    If BlkOld.GNLC2 < BlkNow.GNLC2 Then
        For i = 1 To intBadcnt
            For j = 0 To Kihon.CNTHINOLD - 1
                STOCKINFO(j).FURYOL = STOCKINFO(j).FURYOL + udtBADINFO(i).LEN   ' �s�ǂ̒���(HinOld(i)���������Ă���i�Ԃ̒����s��)
                STOCKINFO(j).HARAIL = STOCKINFO(j).HARAIL - udtBADINFO(i).LEN   ' �Ǖi����
            Next j
        Next i
        For i = 0 To Kihon.CNTHINOLD - 1
            If ((STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) > 0) Then           ' �s�ǐ������݂�����s�Ǐd����s�ǔ䗦�ŋ��߂�
                If i = Kihon.CNTHINOLD - 1 Then
                    STOCKINFO(i).HARAIW = Round((STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL)) * STOCKINFO(i).HARAIW)  'STOCKINFO(i).HARAIW�͓��͍ς�  �f2003/08/06 hitec)matsumoto ROUND�ǉ�
                    STOCKINFO(i).FuryoW = STOCKINFO(i).FuryoW - STOCKINFO(i).HARAIW
                    STOCKINFO(i).HARAIM = Round((STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL)) * STOCKINFO(i).HARAIM)  'STOCKINFO(i).HARAIM�͓��͍ς�  �f2003/08/06 hitec)matsumoto ROUND�ǉ�
                    STOCKINFO(i).FURYOM = STOCKINFO(i).FURYOM - STOCKINFO(i).HARAIM
                Else
                    STOCKINFO(i).HARAIW = HinOld(i).GNWCA
                    STOCKINFO(i).FuryoW = 0
                    STOCKINFO(i).HARAIM = HinOld(i).GNMCA
                    STOCKINFO(i).FURYOM = 0
                End If
            End If
        Next i
    Else
        ' STOCKINFO�̕����͊��ɓ��͍ς�
        For i = 1 To intBadcnt
            For j = 0 To Kihon.CNTHINOLD - 1
                If (udtBADINFO(i).pos >= CLng(HinOld(j).INPOSCA) And _
                    udtBADINFO(i).pos < CLng(HinOld(j).INPOSCA) + CLng(HinOld(j).GNLCA)) Then
                    STOCKINFO(j).FURYOL = STOCKINFO(j).FURYOL + udtBADINFO(i).LEN   ' �s�ǂ̒���(HinOld(i)���������Ă���i�Ԃ̒����s��)
                    STOCKINFO(j).HARAIL = STOCKINFO(j).HARAIL - udtBADINFO(i).LEN   ' �Ǖi����
                End If
            Next j
        Next i

        ' �d�ʂƖ����̕s�ǂƕ����o���̒l��ݒ肷��
        ' STOCKINFO(i).HARAIW�͓��͍ς�
        For i = 0 To Kihon.CNTHINOLD - 1
            If ((STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) > 0) Then           ' �s�ǐ������݂�����s�Ǐd����s�ǔ䗦�ŋ��߂�
                STOCKINFO(i).HARAIW = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIW
                STOCKINFO(i).FuryoW = STOCKINFO(i).FuryoW - STOCKINFO(i).HARAIW
                STOCKINFO(i).HARAIM = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIM
                STOCKINFO(i).FURYOM = STOCKINFO(i).FURYOM - STOCKINFO(i).HARAIM
            End If
        Next i
    End If

    ' �s�ǂ�����ꍇ���݌Ɍ����̍쐬
    For i = 0 To Kihon.CNTHINOLD - 1
        If STOCKINFO(i).FURYOL <> 0 Then
            udtKoutei.CRYNUMC3 = HinNow(0).CRYNUMCA                 ' �u���b�N�h�c
            giInpos = giInpos + 1
            udtKoutei.INPOSC3 = giInpos                             ' �ʒu
            udtKoutei.KCNTC3 = STOCKINFO(i).KCKNT + 1               ' �H���A��
            udtKoutei.HINBC3 = HinOld(i).HINBCA                     ' �i��
            udtKoutei.REVNUMC3 = HinOld(i).REVNUMCA                 ' ���i�����ԍ�
            udtKoutei.FACTORYC3 = HinOld(i).FACTORYCA               ' �H��
            udtKoutei.OPEC3 = HinOld(i).OPECA                       ' ���Ə���
            udtKoutei.LENC3 = STOCKINFO(i).HARAIL                   ' ����
            udtKoutei.XTALC3 = HinOld(i).XTALCA                     ' �����ԍ�
            udtKoutei.SXLIDC3 = ""                                  ' SXLID

            udtKoutei.KNKTC3 = left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 1)               ' �Ǘ��H��(���ݍH��+1)

            udtKoutei.WKKTC3 = Kihon.NOWPROC                        ' �H��
            udtKoutei.WKKBC3 = ""                                   ' ��Ƌ敪
            udtKoutei.MACOC3 = HinNow(0).NEMACOCA                   ' ������
            udtKoutei.MODKBC3 = ""                                  ' �ԍ��敪
            udtKoutei.SUMKBC3 = ""                                  ' �W�v�敪
            udtKoutei.FRKNKTC3 = ""                                 ' (���)�Ǘ��H��

            If IsNull(HinOld(0).NEWKNTCA) = True Then               '(����j�H��
                udtKoutei.FRWKKTC3 = ""
            Else
                udtKoutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
            End If
            udtKoutei.FRWKKBC3 = ""                                 ' (���)��Ƌ敪
            If IsNull(HinOld(0).NEMACOCA) = True Then               '�i����j������
                udtKoutei.FRMACOC3 = "0"
            Else
                udtKoutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If

            Select Case Kihon.NOWPROC
                Case "CC730"
                    intHantei = CInt(BlkNow.GNLC2)
                Case Else
                    intHantei = CInt(BlkNow.GNMC2)
            End Select

            If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                udtKoutei.TOWKKTC3 = " "                            ' (���o)�H��
                udtKoutei.TOMACOC3 = "0"                            ' (���o)������
            Else
                udtKoutei.TOWKKTC3 = HinNow(0).GNWKNTCA             ' (���o)�H��
                udtKoutei.TOMACOC3 = HinNow(0).GNMACOCA             ' (���o)������
            End If
            udtKoutei.FRLC3 = HinOld(i).GNLCA                       ' �������
            udtKoutei.FRWC3 = HinOld(i).GNWCA                       ' ����d��
            udtKoutei.FRMC3 = HinOld(i).GNMCA                       ' �������
            udtKoutei.FULC3 = STOCKINFO(i).FURYOL                   ' �s�ǒ���
            udtKoutei.FUWC3 = STOCKINFO(i).FuryoW                   ' �s�Ǐd��
            udtKoutei.FUMC3 = STOCKINFO(i).FURYOM                   ' �s�ǖ���
            udtKoutei.LOSWC3 = ""                                   ' ���X����

            udtKoutei.LOSLC3 = ""                                   ' ���X�d��
            udtKoutei.LOSMC3 = ""                                   ' ���X����
            udtKoutei.TOLC3 = STOCKINFO(i).HARAIL                   ' ���o����
            udtKoutei.TOWC3 = STOCKINFO(i).HARAIW                   ' ���o�d��
            udtKoutei.TOMC3 = STOCKINFO(i).HARAIM                   ' ���o����
            udtKoutei.SUMITLC3 = ""                                 ' SUMIT����
            udtKoutei.SUMITWC3 = ""                                 ' SUMIT�d��
            udtKoutei.SUMITMC3 = ""                                 ' SUMIT����
            udtKoutei.MOTHINC3 = " "                                ' ���i��
            udtKoutei.XTWORKC3 = "42"                               ' �����H��

            udtKoutei.WFWORKC3 = ""                                 ' ���ʐ���
            udtKoutei.HOLDCC3 = " "                                 ' �z�[���h�R�[�h
            udtKoutei.HOLDBC3 = "0"                                 ' �z�[���h�敪
            udtKoutei.LDFRCC3 = ""                                  ' �i���R�[�h
            udtKoutei.LDFRBC3 = "0"                                 ' �i���敪�i�n�C�L�j
            udtKoutei.TSTAFFC3 = Kihon.STAFFID                      ' �o�^�Ј�ID
            udtKoutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �o�^���t

            udtKoutei.KSTAFFC3 = ""                                 ' �X�V�Ј�ID
            udtKoutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
            udtKoutei.SUMITBC3 = ""                                 ' SUMIT���M�t���O
            udtKoutei.SNDKC3 = ""                                   ' ���M�t���O
'           udtKoutei.SNDDAYC3 = ""                                 ' ���M���t
            udtKoutei.MODMACOC3 = ""                                ' �ԍ��̏�����
            udtKoutei.KAKUCC3 = ""                                  ' �m��R�[�h
            udtKoutei.SUMDAYC3 = CalcSumcoTime(udtKoutei.KDAYC3)    ' SUMCO����
            udtKoutei.PAYCLASSC3 = ""                               ' �]����H��t���O

            intRtn = CreateXSDC3(udtKoutei, sErrMsg)                ' �H�����тɍ݌Ɍ����o�^
            If intRtn = FUNCTION_RETURN_FAILURE Then                ' �H�����ђǉ��G���[
                MsgBox sErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc4 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    �֐���        : XSDC3Proc5
'*
'*    �����T�v      : 1.�H�����ѓo�^�������s��(�i�ԐU�֏��FCC730�p)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC3Proc5() As FUNCTION_RETURN
    ' �����ϐ�
    Dim i, j            As Integer
    Dim intRtn          As Integer                  ' ���A���
    Dim sSQL            As String                   ' �r�p�k
    Dim sSqlWhere       As String                   ' WHERE��
    Dim sErrMsg         As String
    Dim udtKoutei       As typ_XSDC3_Update         ' �H������

    Dim lngLen          As Long
    Dim lngCHKPOS       As Long

    Dim udtWCHKPOS()    As typ_trans_info           ' �O�i�ԕ��ёւ��p
    Dim udtWNINF()      As typ_trans_info           ' ��i�ԕ��ёւ��p
    Dim udtWWINF()      As typ_trans_info           ' ���ёւ��p���[�N
    Dim intBuf          As Integer
    Dim intOINFrecCnt   As Integer
    Dim intNINFrecCnt   As Integer
    Dim intOINFFLG      As Integer
    Dim intNINFFLG      As Integer
    Dim intPoint        As Integer
    Dim intNINFMAX      As Integer
    Dim intOINFMAX      As Integer
    Dim intHantei       As Integer

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    ' �����ݒ�
    XSDC3Proc5 = FUNCTION_RETURN_FAILURE

    ReDim udtWCHKPOS(UBound(STOCKINFO))             ' �i�Ԃ��ƂɃ\�[�g�p
    ReDim udtWNINF(Kihon.CNTHINNOW)                 ' �i�Ԃ��ƂɃ\�[�g�p
    ReDim udtWWINF(1)                               ' �i�Ԃ��ƂɃ\�[�g�p

    ' �݌Ɍ�������荞��
    For i = 0 To UBound(STOCKINFO) - 1
        udtWCHKPOS(i).hinban = STOCKINFO(i).hinban
        udtWCHKPOS(i).LEN = STOCKINFO(i).HARAIL     ' �O�H���������v
        udtWCHKPOS(i).WAT = STOCKINFO(i).HARAIW     ' �O�H���d�ʍ��v
        udtWCHKPOS(i).MAI = STOCKINFO(i).HARAIM     ' �O�H���������v
    Next i

    ' �Ǖi������荞��
    For j = 0 To Kihon.CNTHINNOW - 1
        udtWNINF(j).hinban = HinNow(j).HINBCA
        udtWNINF(j).LEN = HinNow(j).GNLCA           ' ��H���������v
        udtWNINF(j).WAT = HinNow(j).GNWCA           ' ��H���d�ʍ��v
        udtWNINF(j).MAI = HinNow(j).GNMCA           ' ��H���������v
        udtWNINF(j).KCKNT = HinNow(j).KCKNTCA       ' ��H���A��
    Next j

    ' �����i�ԓ��m�A��������ł�����
    For i = 0 To UBound(STOCKINFO) - 1
        For j = 0 To Kihon.CNTHINNOW - 1
           If udtWCHKPOS(i).hinban = udtWNINF(j).hinban Then
                ' ���������̐����𗼕��������
                If udtWCHKPOS(i).LEN <= udtWNINF(j).LEN Then
                    udtWNINF(j).LEN = udtWNINF(j).LEN - udtWCHKPOS(i).LEN
                    udtWNINF(j).WAT = udtWNINF(j).WAT - udtWCHKPOS(i).WAT
                    udtWNINF(j).MAI = udtWNINF(j).MAI - udtWCHKPOS(i).MAI
                    udtWCHKPOS(i).LEN = 0
                    udtWCHKPOS(i).WAT = 0
                    udtWCHKPOS(i).MAI = 0
                Else
                    udtWCHKPOS(i).LEN = udtWCHKPOS(i).LEN - udtWNINF(j).LEN
                    udtWCHKPOS(i).WAT = udtWCHKPOS(i).WAT - udtWNINF(j).WAT
                    udtWCHKPOS(i).MAI = udtWCHKPOS(i).MAI - udtWNINF(j).MAI
                    If udtWCHKPOS(i).MAI < 0 Then
                        udtWCHKPOS(i).MAI = 0
                    End If
                    udtWNINF(j).LEN = 0
                    udtWNINF(j).WAT = 0
                    udtWNINF(j).MAI = 0
                End If
            End If
        Next
    Next

    For i = 0 To UBound(udtWCHKPOS) - 2
        For j = i + 1 To UBound(udtWCHKPOS) - 1
            If (StrComp(udtWCHKPOS(i).hinban, udtWCHKPOS(j).hinban, _
                vbTextCompare)) = 1 Then ' �i�Ԃ̓��֕K�v
                udtWWINF(0) = udtWCHKPOS(j)
                udtWCHKPOS(j) = udtWCHKPOS(i)
                udtWCHKPOS(i) = udtWWINF(0)
            End If
        Next j
    Next i

    ' wNINF�̕i�Ԃ��\�[�g����
    For i = 0 To UBound(udtWNINF) - 2
        For j = i + 1 To UBound(udtWNINF) - 1
            If (StrComp(udtWNINF(i).hinban, udtWNINF(j).hinban, _
                vbTextCompare)) = 1 Then ' �i�Ԃ̓��֕K�v
                udtWWINF(0) = udtWNINF(j)
                udtWNINF(j) = udtWNINF(i)
                udtWNINF(i) = udtWWINF(0)
            End If
        Next j
    Next i

    ' �󂫂̔z��폜����(�z��̃f�[�^���l�߂�)
    For i = 0 To intOINFMAX
        If udtWCHKPOS(i).LEN <= 0 Then
            intPoint = i
            Call HairetuOpe(udtWCHKPOS(), intPoint, -1)
        End If
    Next i

    ' �󂫂̔z��폜����(�z��̃f�[�^���l�߂�)
    For i = 0 To intNINFMAX
        If udtWNINF(i).LEN <= 0 Then
            intPoint = i
            Call HairetuOpe(udtWNINF(), intPoint, -1)
        End If
    Next i

    ' �i�ԓ��֏����쐬����
    i = 0 ' �O�i�Ԃ̈ʒu
    j = 0 ' ��i�Ԃ̈ʒu
    Do
        ' ������˂����킹�Đ��ʂ������łȂ�������傫���l�̕i�Ԃ𕪊�����
        If (udtWCHKPOS(i).LEN = udtWNINF(j).LEN And udtWCHKPOS(i).hinban = udtWNINF(j).hinban) Then   ' �i�Ԓ����������������Ƃ����ɐi��
        ElseIf (udtWCHKPOS(i).LEN >= udtWNINF(j).LEN) Then  ' �i�Ԗ������قȂ鎞
            intPoint = i
            Call HairetuOpe(udtWCHKPOS(), intPoint, 1)      ' �z��̒ǉ�
            udtWCHKPOS(i + 1).hinban = udtWCHKPOS(i).hinban
            udtWCHKPOS(i + 1).LEN = udtWCHKPOS(i).LEN - udtWNINF(j).LEN
            udtWCHKPOS(i + 1).WAT = udtWCHKPOS(i).WAT - udtWNINF(j).WAT
            udtWCHKPOS(i + 1).MAI = udtWCHKPOS(i).MAI - udtWNINF(j).MAI
            If udtWCHKPOS(i + 1).MAI < 0 Then
                udtWCHKPOS(i + 1).MAI = 0
            End If
            udtWCHKPOS(i).LEN = udtWNINF(j).LEN
            udtWCHKPOS(i).WAT = udtWNINF(j).WAT
            udtWCHKPOS(i).MAI = udtWNINF(j).MAI
            Debug.Print "HINBAN=", i, udtWCHKPOS(i).hinban
            Debug.Print "LEN=", i, udtWCHKPOS(i).LEN
        ElseIf (udtWCHKPOS(i).LEN < udtWNINF(j).LEN) Then   ' �i�Ԑ��ʂ��قȂ鎞
            intPoint = j
            Call HairetuOpe(udtWNINF(), intPoint, 1)
            udtWNINF(j + 1).hinban = udtWNINF(i).hinban
            udtWNINF(j + 1).LEN = udtWNINF(j).LEN - udtWCHKPOS(i).LEN
            udtWNINF(j + 1).WAT = udtWNINF(j).WAT - udtWCHKPOS(i).WAT
            udtWNINF(j + 1).MAI = udtWNINF(j).MAI - udtWCHKPOS(i).MAI
            udtWNINF(j).LEN = udtWCHKPOS(i).LEN
            udtWNINF(j).WAT = udtWCHKPOS(i).WAT
            udtWNINF(j).MAI = udtWCHKPOS(i).MAI
            Debug.Print "HINBAN=", i, udtWNINF(i).hinban
            Debug.Print "LEN=", i, udtWNINF(i).LEN
        End If

        intOINFrecCnt = UBound(udtWCHKPOS())
        intNINFrecCnt = UBound(udtWNINF())
        i = i + 1
        j = j + 1
        If (i > intOINFrecCnt) Then
            Exit Do

        End If
        If (j > intNINFrecCnt) Then
            Exit Do
        End If

        If (udtWCHKPOS(i).LEN) <= 0 Then
            Exit Do
        End If

        If (udtWNINF(j).LEN) <= 0 Then
            Exit Do
        End If
    Loop

    intOINFrecCnt = UBound(udtWCHKPOS())
    For i = 0 To intOINFrecCnt - 1
        If (StrComp(udtWNINF(i).hinban, udtWCHKPOS(i).hinban, vbTextCompare) <> 0 _
            And Len(Trim(udtWNINF(i).hinban) > 0)) And (udtWNINF(i).LEN > 0) Then ' �i�Ԃ��قȂ鎞�U�֏��ɓo�^����   'upd 2003/05/31 hitec)matsumoto udtWNINF(i).LEN > 0�ǉ�

            udtKoutei.CRYNUMC3 = HinNow(0).CRYNUMCA     ' �u���b�N�h�c
            giInpos = giInpos + 1
            udtKoutei.INPOSC3 = giInpos                 ' �ʒu
            udtKoutei.KCNTC3 = udtWNINF(i).KCKNT        ' �H���A��
            udtKoutei.HINBC3 = udtWNINF(i).hinban       ' �i��
            udtKoutei.REVNUMC3 = HinNow(0).REVNUMCA     ' ���i�����ԍ�
            udtKoutei.FACTORYC3 = HinNow(0).FACTORYCA   ' �H��
            udtKoutei.OPEC3 = HinNow(0).OPECA           ' ���Ə���
            udtKoutei.LENC3 = udtWNINF(i).LEN           ' �������
            udtKoutei.XTALC3 = HinNow(0).XTALCA         ' �����ԍ�
            udtKoutei.SXLIDC3 = ""                      ' SXLID

            udtKoutei.KNKTC3 = left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 2)   ' �Ǘ��H��(���ݍH��+2)
            udtKoutei.WKKTC3 = Kihon.NOWPROC            ' �H��
            udtKoutei.WKKBC3 = ""                       ' ��Ƌ敪
            udtKoutei.MACOC3 = HinNow(0).NEMACOCA       ' ������
            udtKoutei.MODKBC3 = ""                      ' �ԍ��敪
            udtKoutei.SUMKBC3 = ""                      ' �W�v�敪
            udtKoutei.FRKNKTC3 = ""                     ' (���)�Ǘ��H��

            If IsNull(HinOld(0).NEWKNTCA) = True Then   ' (����j�H��
                udtKoutei.FRWKKTC3 = ""
            Else
                udtKoutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
            End If

            udtKoutei.FRWKKBC3 = ""                     ' (���)��Ƌ敪

            If IsNull(HinOld(0).NEMACOCA) = True Then   '�i����j������
                udtKoutei.FRMACOC3 = "0"
            Else
                udtKoutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If

            Select Case Kihon.NOWPROC
                Case "CC730"
                    intHantei = CInt(BlkNow.GNLC2)
                Case Else
                    intHantei = CInt(BlkNow.GNMC2)
            End Select

            If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                udtKoutei.TOWKKTC3 = " "                            ' (���o)�H��
                udtKoutei.TOMACOC3 = "0"                            ' (���o)������
            Else
                udtKoutei.TOWKKTC3 = HinNow(0).GNWKNTCA             ' (���o)�H��
                udtKoutei.TOMACOC3 = HinNow(0).GNMACOCA             ' (���o)������
            End If
            udtKoutei.FRLC3 = udtWNINF(i).LEN                       ' �������
            udtKoutei.FRWC3 = udtWNINF(i).WAT                       ' ����d��
            udtKoutei.FRMC3 = udtWNINF(i).MAI                       ' �������
            udtKoutei.FULC3 = 0                                     ' �s�ǒ���
            udtKoutei.FUWC3 = 0                                     ' �s�Ǐd��
            udtKoutei.FUMC3 = 0                                     ' �s�ǖ���
            udtKoutei.LOSWC3 = ""                                   ' ���X����

            udtKoutei.LOSLC3 = ""                                   ' ���X�d��
            udtKoutei.LOSMC3 = ""                                   ' ���X����
            udtKoutei.TOLC3 = udtWNINF(i).LEN                       ' ���o����
            udtKoutei.TOWC3 = udtWNINF(i).WAT                       ' ���o�d��
            udtKoutei.TOMC3 = udtWNINF(i).MAI                       ' ���o����
            udtKoutei.SUMITLC3 = ""                                 ' SUMIT����
            udtKoutei.SUMITWC3 = ""                                 ' SUMIT�d��
            udtKoutei.SUMITMC3 = ""                                 ' SUMIT����
            udtKoutei.MOTHINC3 = udtWCHKPOS(i).hinban               ' ���i��
            udtKoutei.XTWORKC3 = "42"                               ' �����H��

            udtKoutei.WFWORKC3 = ""                                 ' ���ʐ���
            udtKoutei.HOLDCC3 = " "                                 ' �z�[���h�R�[�h
            udtKoutei.HOLDBC3 = "0"                                 ' �z�[���h�敪
            udtKoutei.LDFRCC3 = ""                                  ' �i���R�[�h
            udtKoutei.LDFRBC3 = "0"                                 ' �i���敪�i�n�C�L�j
            udtKoutei.TSTAFFC3 = Kihon.STAFFID                      ' �o�^�Ј�ID
            udtKoutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �o�^���t

            udtKoutei.KSTAFFC3 = ""                                 ' �X�V�Ј�ID
            udtKoutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
            udtKoutei.SUMITBC3 = ""                                 ' SUMIT���M�t���O
            udtKoutei.SNDKC3 = ""                                   ' ���M�t���O
            udtKoutei.MODMACOC3 = ""                                ' �ԍ��̏�����
            udtKoutei.KAKUCC3 = ""                                  ' �m��R�[�h
            udtKoutei.SUMDAYC3 = CalcSumcoTime(udtKoutei.KDAYC3)    ' SUMCO����
            udtKoutei.PAYCLASSC3 = ""                               ' �]����H��t���O

            intRtn = CreateXSDC3(udtKoutei, sErrMsg)                ' �H�����тɍ݌Ɍ����o�^
            If intRtn = FUNCTION_RETURN_FAILURE Then                ' �H�����ђǉ��G���[
                MsgBox sErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc5 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
    Exit Function
proc_err:
    ' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC3Proc5 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'**************************************************************************************
'*    �֐���        : HairetuOpe
'*
'*    �����T�v      : 1.HinNum�Ԗڂ̔z����󂫂ɂ���(�z��f�[�^�����ɂ��炵�ċ󂯂�)
'*                    2.HinNum�Ԗڂ̔z����폜����(�z��f�[�^��O�ɂ߂�)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*                    udtHinInf        ,I  ,typ_trans_info ,�i�ԐU�֏��
'*                    intHinNum        ,I  ,Integer        ,�i�ԐU�֏��
'*                    intHINFLG        ,I  ,Integer        ,�i�ԐU�֏��
'*
'*    �߂�l        : �Ȃ�
'*
'**************************************************************************************
Public Function HairetuOpe(udtHinInf() As typ_trans_info, intHinNum As Integer, intHINFLG As Integer)
    Dim intRecCnt   As Integer
    Dim i, j        As Integer
    Dim intSflg     As Integer

    intSflg = 0
    intRecCnt = UBound(udtHinInf())

    If (intHINFLG = 1) Then    ' HinNum�Ԗڂ̔z����󂫂ɂ���(�z��f�[�^�����ɂ��炵�ċ󂯂�)
        For i = intHinNum + 1 To intRecCnt ' �����̔z��ɋ󂫏ꏊ��T��
            If (udtHinInf(i).LEN <= 0) Then    ' i�Ԗڂɋ󂫂��������̂Ńf�[�^�����炷
                For j = i To intHinNum + 1 Step -1
                    udtHinInf(j).hinban = udtHinInf(j - 1).hinban
                    udtHinInf(j).LEN = udtHinInf(j - 1).LEN
                    udtHinInf(j).WAT = udtHinInf(j - 1).WAT
                    udtHinInf(j).MAI = udtHinInf(j - 1).MAI
                    udtHinInf(j).KCKNT = udtHinInf(j - 1).KCKNT
                Next j
                intSflg = 1
                Exit For
            End If
        Next i
        If (intSflg = 0) Then  ' �󂫌����炸
            ReDim Preserve udtHinInf(intRecCnt + 1)
            For i = intRecCnt + 1 To intHinNum + 1 Step -1
                udtHinInf(i).hinban = udtHinInf(i - 1).hinban
                udtHinInf(i).LEN = udtHinInf(i - 1).LEN
                udtHinInf(i).WAT = udtHinInf(i - 1).WAT
                udtHinInf(i).MAI = udtHinInf(i - 1).MAI
                udtHinInf(i).KCKNT = udtHinInf(i - 1).KCKNT
            Next i
        End If

        ' intHinNum+1�Ԗڂ��󂫂ɂ���
        udtHinInf(intHinNum + 1).hinban = ""
        udtHinInf(intHinNum + 1).LEN = 0
        udtHinInf(intHinNum + 1).MAI = 0
        udtHinInf(intHinNum + 1).WAT = 0
        udtHinInf(intHinNum + 1).KCKNT = 0
    Else    ' HinNum�Ԗڂ̔z����폜����(�z��f�[�^��O�ɂ߂�)
        i = intHinNum
        udtHinInf(intHinNum).hinban = ""
        udtHinInf(intHinNum).LEN = 0
        udtHinInf(intHinNum).MAI = 0
        udtHinInf(intHinNum).WAT = 0
        udtHinInf(intHinNum).KCKNT = 0

        For j = intHinNum + 1 To intRecCnt
            If (udtHinInf(j).LEN > 0) Then ' HinNum�ȍ~�Ńf�[�^�����݂��Ă�����
                udtHinInf(i).hinban = udtHinInf(j).hinban
                udtHinInf(i).LEN = udtHinInf(j).LEN
                udtHinInf(i).MAI = udtHinInf(j).MAI
                udtHinInf(i).WAT = udtHinInf(j).WAT
                udtHinInf(i).KCKNT = udtHinInf(j).KCKNT
                udtHinInf(j).hinban = ""
                udtHinInf(j).LEN = 0
                udtHinInf(j).MAI = 0
                udtHinInf(j).WAT = 0
                udtHinInf(j).KCKNT = 0
                i = i + 1
            Else
                udtHinInf(j).hinban = ""
                udtHinInf(j).LEN = 0
                udtHinInf(j).MAI = 0
                udtHinInf(j).WAT = 0
                udtHinInf(j).KCKNT = 0
             End If
        Next j
    End If
End Function

'**************************************************************************************
'*    �֐���        : HairetuOpe_Mai
'*
'*    �����T�v      : 1.HinNum�Ԗڂ̔z����󂫂ɂ���(�z��f�[�^�����ɂ��炵�ċ󂯂�)
'*                    2.HinNum�Ԗڂ̔z����폜����(�z��f�[�^��O�ɂ߂�)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*                    udtHinInf        ,I  ,typ_trans_info ,�i�ԐU�֏��
'*                    intHinNum        ,I  ,Integer        ,�i�ԐU�֏��
'*                    intHINFLG        ,I  ,Integer        ,�i�ԐU�֏��
'*
'*    �߂�l        : �Ȃ�
'*
'**************************************************************************************
Public Function HairetuOpe_Mai(udtHinInf() As typ_trans_info, intHinNum As Integer, intHINFLG As Integer)
    Dim intRecCnt   As Integer
    Dim i, j        As Integer
    Dim intSflg     As Integer

    intSflg = 0
    intRecCnt = UBound(udtHinInf())

    If (intHINFLG = 1) Then    ' HinNum�Ԗڂ̔z����󂫂ɂ���(�z��f�[�^�����ɂ��炵�ċ󂯂�)
        For i = intHinNum + 1 To intRecCnt ' �����̔z��ɋ󂫏ꏊ��T��
            If (udtHinInf(i).MAI <= 0) Then    ' i�Ԗڂɋ󂫂��������̂Ńf�[�^�����炷
                For j = i To intHinNum + 1 Step -1
                    udtHinInf(j).hinban = udtHinInf(j - 1).hinban
                    udtHinInf(j).LEN = udtHinInf(j - 1).LEN
                    udtHinInf(j).WAT = udtHinInf(j - 1).WAT
                    udtHinInf(j).MAI = udtHinInf(j - 1).MAI
                    udtHinInf(j).KCKNT = udtHinInf(j - 1).KCKNT
                Next j
                intSflg = 1
                Exit For
            End If
        Next i
        If (intSflg = 0) Then  ' �󂫌����炸
            ReDim Preserve udtHinInf(intRecCnt + 1)
            For i = intRecCnt + 1 To intHinNum + 1 Step -1
                udtHinInf(i).hinban = udtHinInf(i - 1).hinban
                udtHinInf(i).LEN = udtHinInf(i - 1).LEN
                udtHinInf(i).WAT = udtHinInf(i - 1).WAT
                udtHinInf(i).MAI = udtHinInf(i - 1).MAI
                udtHinInf(i).KCKNT = udtHinInf(i - 1).KCKNT
            Next i
        End If

        ' intHinNum+1�Ԗڂ��󂫂ɂ���
        udtHinInf(intHinNum + 1).hinban = ""
        udtHinInf(intHinNum + 1).LEN = 0
        udtHinInf(intHinNum + 1).MAI = 0
        udtHinInf(intHinNum + 1).WAT = 0
        udtHinInf(intHinNum + 1).KCKNT = 0
    Else    ' HinNum�Ԗڂ̔z����폜����(�z��f�[�^��O�ɂ߂�)
        i = intHinNum
        udtHinInf(intHinNum).hinban = ""
        udtHinInf(intHinNum).LEN = 0
        udtHinInf(intHinNum).MAI = 0
        udtHinInf(intHinNum).WAT = 0
        udtHinInf(intHinNum).KCKNT = 0
        For j = intHinNum + 1 To intRecCnt
            If (udtHinInf(j).MAI > 0) Then ' HinNum�ȍ~�Ńf�[�^�����݂��Ă�����
                udtHinInf(i).hinban = udtHinInf(j).hinban
                udtHinInf(i).LEN = udtHinInf(j).LEN
                udtHinInf(i).MAI = udtHinInf(j).MAI
                udtHinInf(i).WAT = udtHinInf(j).WAT
                udtHinInf(i).KCKNT = udtHinInf(j).KCKNT
                udtHinInf(j).hinban = ""
                udtHinInf(j).LEN = 0
                udtHinInf(j).MAI = 0
                udtHinInf(j).WAT = 0
                udtHinInf(j).KCKNT = 0
                i = i + 1
            Else
                udtHinInf(j).hinban = ""
                udtHinInf(j).LEN = 0
                udtHinInf(j).MAI = 0
                udtHinInf(j).WAT = 0
                udtHinInf(j).KCKNT = 0
             End If
        Next j
    End If
End Function
