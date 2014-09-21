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
    HSXTMMAX As Double                ' �i�r�w�]�ʖ��x���             ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
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

'add 2003/03/29 hitec)sada --------------
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
'add 2003/03/29 hitec)sada ---------
Public giInpos  As Integer  ' add 2003/04/09 hitec)matsumoto
Public strSxlData   As String   'Add. 03/05/01 �㓡

'�T�v      :��ʂ���̊�{�������s��
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function KihonProc() As FUNCTION_RETURN

'   �����ϐ�
    Dim i               As Integer
    Dim j               As Integer
    Dim iRtn            As Integer          '���A���
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim wErrMsg         As String
        
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    KihonProc = FUNCTION_RETURN_FAILURE
    '########################################### 2003/05/23 okazaki
    'XSDCAProc�AXSDC2Proc�̏��ԓ���ւ�
'    �ᕪ�������i�i�ԁj�o�^��
    iRtn = XSDCAProc()
    If iRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDCAProc()�FXSDCA�o�^�G���["
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
'    �ᕪ�������i�u���b�N�j�o�^��
    iRtn = XSDC2Proc()
    If iRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDC2Proc()�FXSDC2�o�^�G���["
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
    '########################################### 2003/05/23 okazaki
'    ��s�Ǔ���o�^��
    '�s�ǗL�������鎞
    If Kihon.FURYOUMU = "Y" Then
                                                ' �o�^���t
        Furyou.TDAYC4 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                                                ' �X�V���t
        Furyou.KDAYC4 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
        'add start 2003/06/01 hitec)matsumoto �s�ǒ����E�d�ʁE�������Ď擾-------------
'''        Furyou.PUCUTLC4 = CLng(BlkNow.GNLC2) - CLng(BlkOld.GNLC2)
'''        Furyou.PUCUTWC4 = CLng(BlkNow.GNWC2) - CLng(BlkOld.GNWC2)
'''        Furyou.PUCUTMC4 = CLng(BlkNow.GNMC2) - CLng(BlkOld.GNMC2)
        Furyou.PUCUTLC4 = CLng(BlkOld.GNLC2) - CLng(BlkNow.GNLC2)
        Furyou.PUCUTWC4 = CLng(BlkOld.GNWC2) - CLng(BlkNow.GNWC2)
        Furyou.PUCUTMC4 = CLng(BlkOld.GNMC2) - CLng(BlkNow.GNMC2)
        'add end 2003/06/01 hitec)matsumoto -------------
        '�s�Ǔ���ǉ�
        iRtn = CreateXSDC4(Furyou, wErrMsg)
        '�s�Ǔ���ǉ��G���[
        If iRtn = FUNCTION_RETURN_FAILURE Then
            MsgBox wErrMsg
            Debug.Print "CreateXSDC4()�FXSDC4�o�^�G���["
            KihonProc = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    Debug.Print HinNow(0).SXLIDCA
'    ��H�����ѓo�^��
    iRtn = XSDC3Proc()
    If iRtn = FUNCTION_RETURN_FAILURE Then
        Debug.Print "XSDC3Proc()�FXSDC3�o�^�G���["
        KihonProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    Select Case Kihon.NOWPROC
        Case "CW740", "CW760"
        '    ��݌Ɍ����o�^��
            iRtn = XSDC3Proc2()
            If iRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()�FXSDC3�݌Ɍ����o�^�G���["
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        '    ��U�֏��o�^��
            iRtn = XSDC3Proc3()
            If iRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()�FXSDC3�U�֏��o�^�G���["
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        Case "CC730"
        '    ��݌Ɍ����o�^��
            iRtn = XSDC3Proc4()
            If iRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()�FXSDC3�݌Ɍ����o�^�G���["
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        '    ��U�֏��o�^��
            iRtn = XSDC3Proc5()
            If iRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()�FXSDC3�U�֏��o�^�G���["
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
    End Select
    Debug.Print HinNow(0).SXLIDCA
'    �ᕪ�������i�r�w�k�j�o�^��
    iRtn = XSDCBProc()
    If iRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDCBProc()�FXSDCB�o�^�G���["
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
    KihonProc = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    KihonProc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

'�T�v      :���������i�u���b�N�j�o�^�������s��
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :

Public Function XSDC2Proc()

'   �����ϐ�
    Dim i               As Integer
    Dim j               As Integer
    Dim iRtn            As Integer          '���A���
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim sqlWhere        As String           'WHERE��
    Dim wErrMsg         As String
    Dim intSyoriKaisu   As Integer          '���ݏ�����
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    XSDC2Proc = FUNCTION_RETURN_FAILURE
    
    '���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
''''    If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
'''    If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
    Select Case Kihon.NOWPROC
    Case "CC730"
        iHantei = CInt(BlkNow.GNLC2)
    Case Else
        iHantei = CInt(BlkNow.GNMC2)
    End Select
    If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
        BlkNow.LSTATBC2 = "H"                   '�ŏI��ԋ敪�i�p���j
        BlkNow.LDFRBC2 = "2"                    '�i���敪�i�n�C�L�j
        BlkNow.LIVKC2 = "1"                     '�����敪�i�����b�g�j
    '2002/12/13 ooba �����敪�t���O�ύX
    '   BlkNow.KANKC2 = "2"                     '�����敪�i�I���j
        '2002/12/02 ooba
        BlkNow.GNWKNTC2 = " "                   '���ݍH��
        BlkNow.GNMACOC2 = "0"                   '���ݏ�����
    Else
        '2002/11/24 tuku �����񐔎擾���W�b�N�ύX
        intSyoriKaisu = GetGNMACOC(BlkNow.CRYNUMC2, BlkNow.GNWKNTC2)
        If BlkNow.GNWKNTC2 = BlkNow.NEWKNTC2 Then
            intSyoriKaisu = intSyoriKaisu + 1
        End If
        BlkNow.GNMACOC2 = intSyoriKaisu                             '���ݏ�����
'        BlkNow.NEMACOC2 = GetNEMACOC2(BlkNow.CRYNUMC2)               '�ŏI�ʉߏ�����
        
        '2002/12/02 tuku
        BlkNow.NEMACOC2 = GetGNMACOC(BlkNow.CRYNUMC2, BlkNow.NEWKNTC2)   '�ŏI�ʉߏ�����
    End If
    
                                                ' �X�V���t
    BlkNow.KDAYC2 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
    
    sqlWhere = "WHERE CRYNUMC2 = '" & BlkNow.CRYNUMC2 & "' "

    iRtn = UpdateXSDC2(BlkNow, sqlWhere)
    '���������i�u���b�N�j�X�V�G���[
    If iRtn = FUNCTION_RETURN_FAILURE Then
        MsgBox "XSDCB UPDATET ERROR"
        Exit Function
    End If
    
    XSDC2Proc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC2Proc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

'�T�v      :���������i�i�ԁj�o�^�������s��
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :

Public Function XSDCAProc()

'   �����ϐ�
    Dim i               As Integer
    Dim j               As Integer
    Dim iRtn            As Integer          '���A���
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim sqlWhere        As String           'WHERE��
    Dim wErrMsg         As String
    Dim LivFlg          As Integer          '���݃t���O
    Dim wHinban()         As typ_XSDCA        '���������i�i�ԁj���[�N�̈�
    Dim wHinban_UP()      As typ_XSDCA_Update '���������i�i�ԁj���[�N�̈�
    Dim intDataCnt      As Integer          '�Y���f�[�^����
    Dim lngSumGNWCA     As Long
    Dim lngSumGNMCA     As Long
    Dim bChgFlg         As Boolean
    Dim intSyoriKaisu As Integer    '���ݏ�����
    Dim iHantei         As Integer  'add 2003/05/27 hitec)matsumoto
    Dim lGetLength      As Long     'add 2003/06/02 hitec)matsumoto TBCME040���A�u���b�N�������擾����
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    XSDCAProc = FUNCTION_RETURN_FAILURE
    lngSumGNWCA = 0
    lngSumGNMCA = 0
    bChgFlg = False
    '�i�Ԃ̏d�ʁE�����v�Z
    If Kihon.CNTHINNOW = Kihon.CNTHINOLD Then   '�O�H���ƌ��ݍH���̌����������ŁA�e�����������ꍇ�͌v�Z���������Ȃ�
        For i = 0 To Kihon.CNTHINNOW - 1
            If CLng(HinNow(i).GNLCA) <> CLng(HinOld(i).GNLCA) Then
                bChgFlg = True
            End If
        Next
    Else            '�O�H���ƌ��ݍH���̌������Ⴄ�ꍇ�́A�v�Z�������s��
        bChgFlg = True
    End If
    '�d�ʁE�����v�Z����

    If bChgFlg = True Then
' VVVVV 2003/04/30 ALT BY HITEC)��c�FCW740,CW760�p�ǉ�
        If Kihon.NOWPROC = "CW740" Or Kihon.NOWPROC = "CW760" Then
            For i = 0 To Kihon.CNTHINNOW - 1
                With HinNow(i)  'upd 203/05/19 hitec)matsumoto BLKOLD��ɕύX
                    If Kihon.CNTHINNOW = 1 Then
                        HinNow(i).GNWCA = BlkOld.GNWC2
                        HinNow(i).GNLCA = BlkOld.GNLC2
                        .SUMITLCA = .GNLCA   '' 03/05/18 matsumoto
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    ElseIf i = Kihon.CNTHINNOW - 1 Then
                        HinNow(i).GNWCA = CLng(BlkOld.GNWC2) - lngSumGNWCA
                        HinNow(i).GNLCA = CLng(BlkOld.GNLC2) - lngSumGNLCA
                        .SUMITLCA = .GNLCA   '' 03/05/18 matsumoto
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    Else
                        HinNow(i).GNWCA = Round(CLng(BlkOld.GNWC2) * (CLng(HinNow(i).GNMCA) / CLng(BlkOld.GNMC2)))
                        HinNow(i).GNLCA = Round(CLng(BlkOld.GNLC2) * (CLng(HinNow(i).GNMCA) / CLng(BlkOld.GNMC2)))
                        lngSumGNWCA = lngSumGNWCA + CLng(HinNow(i).GNWCA)
                        lngSumGNLCA = lngSumGNLCA + CLng(HinNow(i).GNLCA)
                        .SUMITLCA = .GNLCA   '' 03/05/18 matsumoto
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    End If
                End With
            Next
        Else
            If BlkOld.GNLC2 = BlkNow.GNLC2 Then 'upd 2003/06/01 hitec)matsumoto �����������ꍇ��BLKOLD��B�������قȂ�ꍇ��BLKNOW��ɂ���
                BlkNow.GNWC2 = BlkOld.GNWC2
                BlkNow.GNMC2 = BlkOld.GNMC2
            Else
                BlkNow.GNWC2 = Round(CLng(BlkOld.GNWC2) * (CLng(BlkNow.GNLC2) / CLng(BlkOld.GNLC2)))
                BlkNow.GNMC2 = Round(CLng(BlkOld.GNMC2) * (CLng(BlkNow.GNLC2) / CLng(BlkOld.GNLC2)))
            End If
            For i = 0 To Kihon.CNTHINNOW - 1    'upd 203/05/19 hitec)matsumoto BLKOLD��ɕύX
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
    '########################################### 2003/05/23 okazaki
    'XSDC2�̏d�ʁA������XSDCA�̍��v�ɂ��邽�߂�BlkNow�Čv�Z
    BlkNow.GNLC2 = 0
    BlkNow.GNWC2 = 0
    For i = 0 To Kihon.CNTHINNOW - 1
        BlkNow.GNLC2 = CLng(BlkNow.GNLC2) + CLng(HinNow(i).GNLCA)   '2003/05/24 clng�ǉ�
        BlkNow.GNWC2 = CLng(BlkNow.GNWC2) + CLng(HinNow(i).GNWCA)
    Next i
    '########################################### 2003/05/23 end

    '�O�H���̕��������i�i�ԁj�ƗǕi���̕i�ԁE�ʒu���r����
    For i = 0 To Kihon.CNTHINOLD - 1
    
        LivFlg = 0
        For j = 0 To Kihon.CNTHINNOW - 1
            If (HinOld(i).HINBCA = HinNow(j).HINBCA) And (HinOld(i).INPOSCA = HinNow(j).INPOSCA) Then
                LivFlg = 1
            End If
        Next j
        '�O�H���̕��������i�i�ԁj�ɂ����ėǕi���ɂȂ����͎̂����b�g�Ƃ���
        If LivFlg = 0 Then
            sqlWhere = "WHERE CRYNUMCA = '" & HinOld(i).CRYNUMCA & "' "
            sqlWhere = sqlWhere & "AND HINBCA = '" & HinOld(i).HINBCA & "' "
            sqlWhere = sqlWhere & "AND INPOSCA = '" & HinOld(i).INPOSCA & "' "
            ReDim wHinban(0) As typ_XSDCA
            
            '�f�[�^�̌������擾
            iRtn = SelCntXSDCA(sqlWhere, intDataCnt)
            If iRtn = FUNCTION_RETURN_FAILURE Then  '�G���[
                MsgBox "XSDCA SELECT ERROR"
                Exit Function
            Else                                    '����
                If intDataCnt = 0 Then
                    '�O�H���̏��͕K������͂��Ȃ̂ŁA0���̓G���[
''''                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                ElseIf intDataCnt > 0 Then
                    '�O�H��
                    iRtn = DBDRV_GetXSDCA(wHinban(), sqlWhere)
                    '���݂��Ȃ����G���[
                    If iRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox "XSDCA SELECT ERROR"
                        Exit Function
                    End If
                    ReDim wHinban_UP(0) As typ_XSDCA_Update
                    
                    wHinban_UP(0).CRYNUMCA = HinOld(i).CRYNUMCA
                    wHinban_UP(0).INPOSCA = HinOld(i).INPOSCA
                    wHinban_UP(0).HINBCA = HinOld(i).HINBCA
                    
                    '�����敪�Ɏ����b�g���Z�b�g
                    wHinban_UP(0).LIVKCA = "1"              ' �����敪
                    wHinban_UP(0).LSTATBCA = "H"            ' �ŏI��ԋ敪
                    wHinban_UP(0).LDFRBCA = "2"             ' �i���敪
                    ' 2002/12/13 ooba �����敪�t���O�ύX
                    wHinban_UP(0).KANKCA = "0"              ' �����敪
'                    wHinban_UP(0).KANKCA = "2"              ' �����敪

                    
                    wHinban_UP(0).SUMITBCA = "0"
''''                    wHinban_UP(0).SUMITLCA = "0"    'del 2003/05/18 hitec)matsumoto
''''                    wHinban_UP(0).SUMITMCA = "0"
''''                    wHinban_UP(0).SUMITWCA = "0"
                                                        ' �X�V���t
                    wHinban_UP(0).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                    
                    '���������i�i�ԁj���X�V
                    iRtn = UpdateXSDCA(wHinban_UP(0), sqlWhere)
                    '���������i�i�ԁj�X�V�G���[
                    If iRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox "XSDCA UPDATE ERROR"
                        Exit Function
                    End If
                End If
            End If
        
        End If
    
    Next i
    
    '���������i�i�ԁj���J��Ԃ�
    For i = 0 To Kihon.CNTHINNOW - 1
        '�����ԍ��A�i�ԁA�ʒu�Ō���
        sqlWhere = "WHERE CRYNUMCA = '" & HinNow(i).CRYNUMCA & "' "
        sqlWhere = sqlWhere & "AND HINBCA = '" & HinNow(i).HINBCA & "' "
        sqlWhere = sqlWhere & "AND INPOSCA = '" & HinNow(i).INPOSCA & "' "

        '�f�[�^�̌������擾
        iRtn = SelCntXSDCA(sqlWhere, intDataCnt)
        If iRtn = FUNCTION_RETURN_FAILURE Then  '�G���[
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        Else                                    '����
            '�f�[�^������ꍇ�͍X�V����
            If intDataCnt > 0 Then
                iRtn = DBDRV_GetXSDCA(wHinban, sqlWhere)
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If

                '���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
''''                If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
                Select Case Kihon.NOWPROC
                Case "CC730"
                    iHantei = CInt(BlkNow.GNLC2)
                Case Else
''''                    iHantei = CInt(BlkNow.GNMC2)
                    iHantei = CInt(HinNow(i).GNMCA) 'upd 2003/06/05 hitec)matsumoto 0����p���ɂ��鏈�����A�u���b�N�P�ʂł͂Ȃ��i�ԒP�ʂɕύX
                End Select
                If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                    HinNow(i).LIVKCA = "1"              ' �����敪
                    HinNow(i).LSTATBCA = "H"            ' �ŏI��ԋ敪
                    HinNow(i).LDFRBCA = "2"             ' �i���敪
'                    HinNow(i).KANKCA = "2"              ' �����敪
                    '2002/12/02 ooba
                    HinNow(i).GNWKNTCA = " "            '���ݍH��
                    HinNow(i).GNMACOCA = "0"            '���ݏ�����
                Else
                    HinNow(i).LIVKCA = "0"               ' �����敪�i�����b�g�j
                    HinNow(i).LSTATBCA = "T"            ' �ŏI��ԋ敪�i�ʏ�j
                    HinNow(i).LDFRBCA = "0"             ' �i���敪�i�ʏ�j
 '                    HinNow(i).KANKCA = "0"              ' �����敪
                    '2002/11/24 tuku �����񐔎擾���W�b�N�ύX
                    intSyoriKaisu = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).GNWKNTCA)
                    If HinNow(i).GNWKNTCA = HinNow(i).NEWKNTCA Then
                          intSyoriKaisu = intSyoriKaisu + 1
                    End If
                    HinNow(i).GNMACOCA = intSyoriKaisu    '���ݏ�����
'                    HinNow(i).NEMACOCA = GetNEMACOC2(HinNow(i).CRYNUMCA)               '�ŏI�ʉߏ�����

                    '2002/12/02 ooba
                    HinNow(i).NEMACOCA = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).NEWKNTCA)   '�ŏI�ʉߏ�����
                End If
                ' 2002/12/13 ooba �����敪�t���O�ύX
                HinNow(i).KANKCA = "0"              ' �����敪
                
                HinNow(i).SUMITBCA = "0"
''''                HinNow(i).SUMITLCA = "0"    'del 2003/05/18 hitec)matsumoto
''''                HinNow(i).SUMITMCA = "0"
''''                HinNow(i).SUMITWCA = "0"

                                                    ' �X�V���t
                HinNow(i).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                
''''                '�V���O���m�莞�A�����敪='2'�E�����敪='1'�ɂ���
''''                If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
''''                    HinNow(i).LIVKCA = "1"
''''                    HinNow(i).KANKCA = "2"
''''                End If
                
                '�Ǖi���Œu������
                iRtn = UpdateXSDCA(HinNow(i), sqlWhere)
                '���������i�i�ԁj�X�V�G���[
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCA UPDATE ERROR"
                    Exit Function
                End If
            '���݂��Ȃ����ǉ�
            ElseIf intDataCnt = 0 Then
                '���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
''''                If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''                If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
                Select Case Kihon.NOWPROC
                Case "CC730"
                    iHantei = CInt(BlkNow.GNLC2)
                Case Else
''''                    iHantei = CInt(BlkNow.GNMC2)
                    iHantei = CInt(HinNow(i).GNMCA) 'upd 2003/06/05 hitec)matsumoto 0����p���ɂ��鏈�����A�u���b�N�P�ʂł͂Ȃ��i�ԒP�ʂɕύX
                End Select
                If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                    HinNow(i).LIVKCA = "1"              ' �����敪
                    HinNow(i).LSTATBCA = "H"            ' �ŏI��ԋ敪
                    HinNow(i).LDFRBCA = "2"             ' �i���敪
'                    HinNow(i).KANKCA = "2"              ' �����敪
                    '2002/12/02 ooba
                    HinNow(i).GNWKNTCA = " "            '���ݍH��
                    HinNow(i).GNMACOCA = "0"            '���ݏ�����
                Else
                    HinNow(i).LIVKCA = "0"               ' �����敪�i�����b�g�j
                    HinNow(i).LSTATBCA = "T"            ' �ŏI��ԋ敪�i�ʏ�j
                    HinNow(i).LDFRBCA = "0"             ' �i���敪�i�ʏ�j
'                    HinNow(i).KANKCA = "0"              ' �����敪
                    '2002/11/24 tuku �����񐔎擾���W�b�N�ύX
                    intSyoriKaisu = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).GNWKNTCA)
                    If HinNow(i).GNWKNTCA = HinNow(i).NEWKNTCA Then
                          intSyoriKaisu = intSyoriKaisu + 1
                    End If
                    HinNow(i).GNMACOCA = intSyoriKaisu                      '���ݏ�����
'                    HinNow(i).NEMACOCA = GetNEMACOC2(HinNow(i).CRYNUMCA)    '�ŏI�ʉߏ�����

                    '2002/12/02 ooba
                    HinNow(i).NEMACOCA = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).NEWKNTCA)   '�ŏI�ʉߏ�����
                End If
                ' 2002/12/13 ooba �����敪�t���O�ύX
                HinNow(i).KANKCA = "0"              ' �����敪
                                                    ' �o�^���t
                HinNow(i).TDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                                                    ' �X�V���t
                HinNow(i).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                
''''                '�V���O���m�莞�A�����敪='2'�E�����敪='1'�ɂ���
''''                If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
''''                    HinNow(i).LIVKCA = "1"
''''                    HinNow(i).KANKCA = "2"
''''                End If
                HinNow(i).SUMITBCA = "0"
'''                HinNow(i).SUMITLCA = "0"     'del 2003/05/18 hitec)matsumoto
'''                HinNow(i).SUMITMCA = "0"
'''                HinNow(i).SUMITWCA = "0"
                
                iRtn = CreateXSDCA(HinNow(i), wErrMsg)
                '���������i�i�ԁj�X�V�G���[
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox wErrMsg
                    Exit Function
                End If
            End If
        End If
    Next i
    
    XSDCAProc = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDCAProc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function


'�T�v      :�H�����ѓo�^�������s��
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :

Public Function XSDC3Proc()

'   �����ϐ�
    Dim i               As Integer
    Dim j               As Integer
    Dim iRtn            As Integer          '���A���
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim sqlWhere        As String           'WHERE��
    Dim wFULC3          As Long             '���������i�i�ԁj�̕s�ǒ���
    Dim wFUWC3          As Long             '���������i�i�ԁj�̕s�Ǐd��
    Dim wFUMC3          As Long             '���������i�i�ԁj�̕s�ǖ���
    Dim wErrMsg         As String
    Dim Koutei          As typ_XSDC3_Update    '�H������
    Dim rsKCNTC         As OraDynaset       '���R�[�h�Z�b�g
    Dim intNextCnt      As Integer
    Dim intOldCnt       As Integer
    'add start 2003/03/28 hitec)matsumoto ------------------
    Dim bNewRec         As Boolean          '�O�H���̖������R�[�h���������ꍇ
    Dim sSUMITLC3       As String           'SUMIT����
    Dim sSUMITWC3       As String           'SUMIT�d��
    Dim sSUMITMC3       As String           'SUMIT����
    Dim dSumcoTime      As Date             'SUMCO����
    Dim vChoseiTime     As Variant          '��������
    'add end   2003/03/28 hitec)matsumoto ------------------
    'add 03/05/17 �㓡 -------------------------------------
    Dim iLoopCnt        As Integer
    '--------------------------------------end 03/05/17 ---
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    bNewRec = False 'add 2003/03/27 hitec)matsumoto �t���O������
    XSDC3Proc = FUNCTION_RETURN_FAILURE
    
''''    'SUMCO���ԍ쐬�ׁ̈A�������Ԏ擾  add   2003/04/01 hitec)matsumoto ----
''''    sql = "SELECT KCODE01A9"
''''    sql = sql & " FROM koda9 "
''''    sql = sql & " WHERE SYSCA9 = 'X'"
''''    sql = sql & "   AND SHUCA9 = '80'"
''''    sql = sql & "   AND CODEA9 = '1'"
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''    '���݂��Ȃ����A�G���[
''''    If rs Is Nothing Then
''''        MsgBox "koda9 KCODE01A9 SELECT ERROR"
''''        Exit Function
''''    End If
''''    If Not rs.EOF Then
''''        If IsNull(rs.Fields("KCODE01A9")) = True Then
''''            MsgBox "koda9 KCODE01A9 SELECT ERROR"
''''            Exit Function
''''        Else
''''            vChoseiTime = CDate(rs.Fields("KCODE01A9"))
''''        End If
''''    End If
''''    rs.Close
''''    '----------------------------------------------------------------------
    
    '�H�����т���u���b�N�h�c�A�i�Ԃ���v����H���A�Ԃ̍ő���擾
    sql = "SELECT MAX(KCNTC3) as wKCNTC3 "
    sql = sql & " FROM XSDC3 "
    sql = sql & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
'        sql = sql & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
''''        sql = sql & " AND ((SUMKBC3 ='0') OR (SUMKBC3 = ' ') OR (SUMKBC3 IS NULL)) "

    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '���݂��Ȃ����A�G���[
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
    
    'add 2003/03/27 hitec)matsumoto �O�H�����т̖������R�[�h�����邩�`�F�b�N���A��������t���O�����Ă�-----------------------
    For i = 0 To Kihon.CNTHINNOW - 1
        '�H�����т���O�H���̃f�[�^��ǂݍ���
        sql = "SELECT STATIMEC3, STOTIMEC3 , TOLC3,TOWC3,TOMC3,WKKTC3,MACOC3 "
        sql = sql & " FROM XSDC3 "
        sql = sql & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sql = sql & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
''''        sql = sql & " AND KCNTC3 = " & intOldCnt & ""
        sql = sql & " AND KCNTC3 = " & intNextCnt - 1 & ""  'upd 2003/05/31 hitec)matsumoto intOldCnt�͎g���Ȃ��̂ŁAintNextCnt - 1��ς��Ɏg�p
''''        sql = sql & " AND ((SUMKBC3 ='0') OR (SUMKBC3 = ' ') OR (SUMKBC3 IS NULL)) "


        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '���݂��Ȃ����A�G���[
        If rs Is Nothing Then
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        End If
        If rs.RecordCount = 0 Then
            bNewRec = True  '�O�H�����Ȃ��ꍇ�̓t���O�����Ă�
        End If
        rs.Close
    Next
    'add end 2003/03/27 hitec)matsumoto ---------------------------
    If Kihon.NOWPROC = "CC730" Then
        bNewRec = True
    End If
    For i = 0 To Kihon.CNTHINNOW - 1
        '�s�Ǔ��󂩂�u���b�N�h�c�A�i�ԁA�J�n�ʒu����v����s�ǒ������擾����
        intOldCnt = 0
        wFULC3 = 0
        wFUWC3 = 0
        wFUMC3 = 0
'            Koutei.FRWKKTC3 = " "
'            Koutei.FRMACOC3 = 0
''''        For j = 0 To Kihon.CNTHINNOW - 1
''''            If HinNow(i).CRYNUMCA = Furyou.XTALC4 And HinNow(i).HINBCA = Furyou.HINBC4 And HinNow(i).INPOSCA = Furyou.INPOSC4 Then
''''                wFULC3 = Furyou.PUCUTLC4
''''                wFUWC3 = Furyou.PUCUTWC4
''''                wFUMC3 = Furyou.PUCUTMC4
''''            End If
''''        Next j
''''
''''        If Furyou.HINBC4 = "Z" Then '�p��
''''            wFULC3 = Furyou.PUCUTLC4
''''            wFUWC3 = Furyou.PUCUTWC4
''''            wFUMC3 = Furyou.PUCUTMC4
''''        End If
    
''''' 02/09/20 Add By ��c@HITEC  sta
''''        '�H�����т���u���b�N�h�c�A�i�Ԃ���v����H���A�Ԃ̍ő���擾
''''        sql = "SELECT MAX(KCNTC3) as wKCNTC3 "
''''        sql = sql & " FROM XSDC3 "
''''        sql = sql & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
'''''        sql = sql & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
''''''''        sql = sql & " AND ((SUMKBC3 ='0') OR (SUMKBC3 = ' ') OR (SUMKBC3 IS NULL)) "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '���݂��Ȃ����A�G���[
''''        If rs Is Nothing Then
''''            MsgBox "XSDC3 MAX KCNT SELECT ERROR"
''''            Exit Function
''''        End If
''''
''''        If rs.EOF = False Then
''''            If IsNull(rs.Fields("wKCNTC3")) = True Then
''''                intNextCnt = 1
''''                intOldCnt = 0
''''            Else
''''                intNextCnt = CInt(rs.Fields("wKCNTC3")) + 1
''''                intOldCnt = CInt(rs.Fields("wKCNTC3"))
''''            End If
''''        Else
''''            intNextCnt = 1
''''            intOldCnt = 0
''''        End If
''''        rs.Close

        '�H�����т���u���b�N�h�c�A�i�Ԃ���v����H���A�Ԃ̍ő���擾
        sql = "SELECT MAX(KCNTC3) as wKCNTC3 "
        sql = sql & " FROM XSDC3 "
        sql = sql & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sql = sql & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
''''        sql = sql & " AND ((SUMKBC3 ='0') OR (SUMKBC3 = ' ') OR (SUMKBC3 IS NULL)) "
    
        
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '���݂��Ȃ����A�G���[
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
    
        '�H�����т���O�H���̂̃f�[�^��ǂݍ���
        sql = "SELECT STATIMEC3, STOTIMEC3 , TOLC3,TOWC3,TOMC3, "
        'add start 2003/03/28 hitec)matsumoto -------
        sql = sql & " SUMITLC3, SUMITWC3, SUMITMC3"
        'add end   2003/03/28 hitec)matsumoto -------
        sql = sql & " FROM XSDC3 "
        sql = sql & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sql = sql & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
        sql = sql & " AND KCNTC3 = " & intOldCnt & ""
''''        sql = sql & " AND ((SUMKBC3 ='0') OR (SUMKBC3 = ' ') OR (SUMKBC3 IS NULL)) "


        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '���݂��Ȃ����A�G���[
        If rs Is Nothing Then
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        End If
        If rs.RecordCount = 0 Then
            'upd end 2003/03/27 hitec)matsumoto �O�H���������́A���o����������������ɓ����------------
''''            Koutei.FRLC3 = ""                       ' ��������N���A
''''            Koutei.FRWC3 = ""                       ' ����d�ʃN���A
''''            Koutei.FRMC3 = ""                       ' ��������N���A
            'upd end 2003/03/27 hitec)matsumoto -------------
            wFULC3 = 0
            wFUWC3 = 0
            wFUMC3 = 0
            'add start 2003/03/28 hitec)matsumoto --------
            sSUMITLC3 = "0" 'SUMIT����
            sSUMITWC3 = "0" 'SUMIT�d��
            sSUMITMC3 = "0" 'SUMIT����
            'add end 2003/03/28 hietc)matsumoto -----------
        Else
            If IsNull(rs.Fields("STATIMEC3")) = True Then
                '��������Ȃ�
            Else
                Koutei.STATIMEC3 = rs.Fields("STATIMEC3")
            End If
                                                    ' �������ԏI��
            If IsNull(rs.Fields("STOTIMEC3")) = True Then
                '��������Ȃ�
            Else
                Koutei.STOTIMEC3 = rs.Fields("STOTIMEC3")
            End If
            
            If IsNull(rs.Fields("TOLC3")) = True Then   '�s�ǒ���
                wFULC3 = 0
                Koutei.FRLC3 = "0"
            Else
                Koutei.FRLC3 = CLng(rs.Fields("TOLC3"))
                wFULC3 = CLng(rs.Fields("TOLC3"))
            End If
            If IsNull(rs.Fields("TOWC3")) = True Then   '�s�Ǐd��
                wFUWC3 = 0
                Koutei.FRWC3 = "0"
            Else
                Koutei.FRWC3 = CLng((rs.Fields("TOWC3")))
                wFUWC3 = CLng((rs.Fields("TOWC3")))
            End If
            If IsNull(rs.Fields("TOMC3")) = True Then   '�s�ǖ���
                wFUMC3 = 0
                Koutei.FRMC3 = "0"
            Else
                Koutei.FRMC3 = CLng((rs.Fields("TOMC3")))
                wFUMC3 = CLng((rs.Fields("TOMC3")))
            End If

            'add start 2003/03/28 hitec)matsumoto --------
            If IsNull(rs.Fields("SUMITLC3")) = True Then   'SUMIT����
                sSUMITLC3 = "0"
            Else
                sSUMITLC3 = CLng((rs.Fields("SUMITLC3")))
            End If
            If IsNull(rs.Fields("SUMITWC3")) = True Then   'SUMIT�d��
                sSUMITWC3 = "0"
            Else
                sSUMITWC3 = CLng((rs.Fields("SUMITWC3")))
            End If
            If IsNull(rs.Fields("SUMITMC3")) = True Then   'SUMIT����
                sSUMITMC3 = "0"
            Else
                sSUMITMC3 = CLng((rs.Fields("SUMITMC3")))
            End If
'            '2002/11/24 tuku �����񐔎擾���W�b�N�ύX
'            If IsNull(rs.Fields("WKKTC3")) = True Then   '(����j�H��
'                Koutei.FRWKKTC3 = "0"
'            Else
'                Koutei.FRWKKTC3 = CStr((rs.Fields("WKKTC3")))
'            End If
'
'            If IsNull(rs.Fields("MACOC3")) = True Then   '�i����j������
'                Koutei.FRMACOC3 = "0"
'            Else
'                Koutei.FRMACOC3 = CLng((rs.Fields("MACOC3")))
'            End If
        End If
        
        '2002/11/29 ooba �����񐔎擾���W�b�N�ύX
        If IsNull(HinOld(0).NEWKNTCA) = True Then   '(����j�H��
            Koutei.FRWKKTC3 = "0"
        Else
            Koutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
            'add end 2003/03/28 hitec)matsumoto --------
        End If
        
        If IsNull(HinOld(0).NEMACOCA) = True Then   '�i����j������
            Koutei.FRMACOC3 = "0"
        Else
            Koutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
        End If
        '���������i�i�ԁj����H�����т�ǉ�
        Koutei.CRYNUMC3 = HinNow(i).CRYNUMCA    ' ��ۯ�ID������ԍ�
        Koutei.INPOSC3 = HinNow(i).INPOSCA      ' �������J�n�ʒu
            
        
        
'''        '�H���A�Ԃ�MAX���擾
'''        sql = ""
'''        sql = "SELECT MAX(KCNTC3) KCNTC3 FROM XSDC3"
'''        sql = sql & "  WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "'"
'''        sql = sql & "    AND INPOSC3 = " & HinNow(i).INPOSCA
'''        sql = sql & "  group by CRYNUMC3,INPOSC3 "
'''
'''        Set rsKCNTC = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'''
'''        If rsKCNTC Is Nothing Then
'''            XSDC3Proc = FUNCTION_RETURN_FAILURE
'''            Exit Function
'''        End If
'''        '�H���A�Ԃ�MAX+1�̒l���g�p
'''        If rsKCNTC.EOF = False Then
'''            If IsNull(rsKCNTC("KCNTC3")) = True Then
'''                lngProcCnt = 0
'''            Else
'''                lngProcCnt = rsKCNTC("KCNTC3") + 1
'''            End If
'''        Else
'''            lngProcCnt = 0
'''        End If
        
'''        rsKCNTC.Close
        Koutei.KCNTC3 = intNextCnt       ' �H���A��
        Koutei.HINBC3 = HinNow(i).HINBCA        ' �i��
        Koutei.REVNUMC3 = HinNow(i).REVNUMCA    ' ���i�ԍ������ԍ�
        Koutei.FACTORYC3 = HinNow(i).FACTORYCA  ' �H��
        Koutei.OPEC3 = HinNow(i).OPECA          ' ���Ə���
        '���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
''''        If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''        If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
        Select Case Kihon.NOWPROC
        Case "CC730"
            iHantei = CInt(BlkNow.GNLC2)
        Case Else
            iHantei = CInt(BlkNow.GNMC2)
        End Select
        If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
            Koutei.LENC3 = 0                    ' ����
        Else
            Koutei.LENC3 = HinNow(i).GNLCA      ' ����
        End If
        Koutei.XTALC3 = HinNow(i).XTALCA        ' �����ԍ�
        Koutei.SXLIDC3 = HinNow(i).SXLIDCA      ' SXLID
        Select Case Kihon.NOWPROC   'upd 2003/04/05 hitec)matsumoto  CW740�CCW760�H���ŁA�Ǘ��H���Ɍ��ݍH���{�R����������
            Case "CW740", "CW760", "CC730"
                Koutei.KNKTC3 = Left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
                      CStr(CInt(Right(Kihon.NOWPROC, 1)) + 3) ' �Ǘ��H��(���ݍH��+3)
            Case Else
                Koutei.KNKTC3 = HinNow(i).GNKKNTCA      ' �Ǘ��H��
        End Select
        Koutei.WKKTC3 = Kihon.NOWPROC           ' �H��
        Koutei.WKKBC3 = HinNow(i).GNWKKBCA      ' ��Ƌ敪
        Koutei.MACOC3 = HinNow(i).NEMACOCA      ' ������
        Koutei.MODKBC3 = ""                     ' �ԍ��敪
        Koutei.SUMKBC3 = "0"                    ' �W�v�敪
        Koutei.FRKNKTC3 = " "                   ' (���)�Ǘ��H��
'        Koutei.FRWKKTC3 = HinOld(0).NEWKNTCA    ' (���)�H��
        Koutei.FRWKKBC3 = " "                   ' (���)��Ƌ敪
'        Koutei.FRWKKTC3 = HinOld(0).NEWKNTCA    ' (���)�H��
        Koutei.TOWNKTC3 = " "                   ' (���o)�Ǘ��H��
        '���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
''''        If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''        If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
        Select Case Kihon.NOWPROC
        Case "CC730"
            iHantei = CInt(BlkNow.GNLC2)
        Case Else
            iHantei = CInt(BlkNow.GNMC2)
        End Select
        If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
            Koutei.TOWKKTC3 = " "               ' (���o)�H��
            '2002/12/02 ooba
            Koutei.TOMACOC3 = "0"               '(���o)������
        Else
            Koutei.TOWKKTC3 = HinNow(i).GNWKNTCA     ' (���o)�H��
        End If
        Koutei.TOMACOC3 = HinNow(i).GNMACOCA    ' (���o)������
        
''''        Koutei.FRLC3 = ""                       ' ��������N���A
''''        Koutei.FRWC3 = ""                       ' ����d�ʃN���A
''''        Koutei.FRMC3 = ""                       ' ��������N���A
''''        For j = 0 To Kihon.CNTHINOLD - 1
''''            If (HinNow(i).CRYNUMCA = HinOld(j).CRYNUMCA) And (HinNow(i).INPOSCA = HinOld(j).INPOSCA) And (HinNow(i).KCKNTCA = HinOld(j).KCKNTCA) Then
''''                Koutei.FRLC3 = HinOld(i).GNLCA  ' �������
''''                Koutei.FRWC3 = HinOld(i).GNWCA  ' ����d��
''''                Koutei.FRMC3 = HinOld(i).GNMCA  ' �������
''''                Exit For
''''            End If
''''        Next j

        Koutei.LOSWC3 = ""                      ' ���X����
        Koutei.LOSLC3 = ""                      ' ���X�d��
        Koutei.LOSMC3 = ""                      ' ���X����
        If bNewRec = True Then  'add 2003/03/27 hitec)matsumoto �O�H���������f�[�^�����݂��Ă���ꍇ�́A���o�����ʂ�������ʂɓ����
            Koutei.FRLC3 = HinNow(i).GNLCA      ' �������<=���o����
            Koutei.FRWC3 = HinNow(i).GNWCA      ' ����d��<=���o�d��
            Koutei.FRMC3 = HinNow(i).GNMCA      ' �������<=���o����
            Koutei.TOLC3 = HinNow(i).GNLCA      ' ���o����
            Koutei.TOWC3 = HinNow(i).GNWCA      ' ���o�d�ʁi�֐��j
            Koutei.TOMC3 = HinNow(i).GNMCA      ' ���o�����i�֐��j
            Koutei.FULC3 = 0                    ' �s�ǒ���
            Koutei.FUWC3 = 0                    ' �s�Ǐd��
            Koutei.FUMC3 = 0                    ' �s�ǖ���
        Else
''''            If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''            If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
            Select Case Kihon.NOWPROC
            Case "CC730"
                iHantei = CInt(BlkNow.GNLC2)
            Case Else
                iHantei = CInt(BlkNow.GNMC2)
            End Select
            If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                Koutei.TOLC3 = 0                    ' ���o����
                Koutei.TOWC3 = 0                    ' ���o�d�ʁi�֐��j
                Koutei.TOMC3 = 0                    ' ���o�����i�֐��j
            Else
                Koutei.TOLC3 = HinNow(i).GNLCA      ' ���o����
                Koutei.TOWC3 = HinNow(i).GNWCA      ' ���o�d�ʁi�֐��j
                Koutei.TOMC3 = HinNow(i).GNMCA      ' ���o�����i�֐��j
            End If
            Koutei.FULC3 = wFULC3 - CLng(Koutei.TOLC3)                  ' �s�ǒ���
            Koutei.FUWC3 = wFUWC3 - CLng(Koutei.TOWC3)                  ' �s�Ǐd��
            Koutei.FUMC3 = wFUMC3 - CLng(Koutei.TOMC3)                  ' �s�ǖ���
        End If
        If Koutei.TOLC3 = "" Then
            Koutei.TOLC3 = "0"
        End If
        If Koutei.TOWC3 = "" Then
            Koutei.TOWC3 = "0"
        End If
        If Koutei.TOMC3 = "" Then
            Koutei.TOMC3 = "0"
        End If
        'upd start 2003/03/28 hitec)matsumoto SUMIT�����ɍH���ʂɒl���Z�b�g����--------------------
        Koutei.SUMITLC3 = 0                     ' SUMIT����
        Koutei.SUMITWC3 = 0                     ' SUMIT�d��
        Koutei.SUMITMC3 = 0                     ' SUMIT����
'----------------------------------------------------- 03/05/13 �㓡 ----------------------------
'''        Select Case Kihon.NOWPROC
'''            Case "CW740", "CW760"
'''                Koutei.SUMITLC3 = HinOld(i).SUMITLCA    ' SUMIT����=���o����
'''                Koutei.SUMITWC3 = HinOld(i).SUMITWCA    ' SUMIT�d��=���o�d��
'''                Koutei.SUMITMC3 = HinOld(i).SUMITMCA    ' SUMIT����=���o����
'''            Case "CW750", "CW800"
'''                Koutei.SUMITLC3 = HinNow(i).SUMITLCA    ' SUMIT����=���o����
'''                Koutei.SUMITWC3 = HinNow(i).SUMITWCA    ' SUMIT�d��=���o�d��
'''                Koutei.SUMITMC3 = HinNow(i).SUMITMCA    ' SUMIT����=���o����
'''        End Select
        For iLoopCnt = 0 To Kihon.CNTHINOLD - 1     '' 03/05/17 �㓡
            If (Koutei.CRYNUMC3 = HinOld(iLoopCnt).CRYNUMCA) _
                And (Koutei.INPOSC3 = HinOld(iLoopCnt).INPOSCA) Then
                    Koutei.SUMITLC3 = HinOld(iLoopCnt).SUMITLCA     ' SUMIT����=�O�H��SUMIT����
                    Koutei.SUMITWC3 = HinOld(iLoopCnt).SUMITWCA     ' SUMIT�d��=�O�H��SUMIT�d��
                    Koutei.SUMITMC3 = HinOld(iLoopCnt).SUMITMCA     ' SUMIT����=�O�H��SUMIT����
                    Exit For
            End If
        Next
'----------------------------------------------------------------------------- END 03/05/13 ------
        'upd end 2003/03/28 hitec)matsumoto ---------------
        Koutei.MOTHINC3 = " "                   ' �U�֕i��(��)
        Koutei.XTWORKC3 = "42"                  ' �����H��
        Koutei.WFWORKC3 = " "                   ' ���ʐ���
                                                ' �������ԊJ�n
''''        Koutei.ETIMEC3 = ""                     ' ���ю��Ԃ͓���Ȃ�
        Koutei.HOLDCC3 = " "                    ' �z�[���h�R�[�h
        Koutei.HOLDBC3 = "0"                    ' �z�[���h�敪
        Koutei.LDFRCC3 = " "                    ' �i���R�[�h
        '���������i�u���b�N�j�̗Ǖi����<=0 or �S���X�N���b�v�̎�
''''        If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''        If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
        Select Case Kihon.NOWPROC
        Case "CC730"
            iHantei = CInt(BlkNow.GNLC2)
        Case Else
            iHantei = CInt(BlkNow.GNMC2)
        End Select
        If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
            Koutei.LDFRBC3 = "2"                ' �i���敪�i�n�C�L�j
        Else
            Koutei.LDFRBC3 = "0"                ' �i���敪
        End If
        Koutei.TSTAFFC3 = Kihon.STAFFID         ' �o�^�Ј�ID
                                                ' �o�^���t
        Koutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
        Koutei.KSTAFFC3 = Kihon.STAFFID         ' �X�V�Ј�ID
                                                ' �X�V���t
        Koutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
        'SUMCO����=�H������.�X�V���t-KODA9.�������� add 2003/04/01 hitec)matsumoto --
''''        dSumcoTime = CDate(Koutei.KDAYC3) - CDate(vChoseiTime)
''''        Koutei.SUMDAYC3 = Format(dSumcoTime, "YYYY/MM/DD")
        Koutei.SUMDAYC3 = CalcSumcoTime(Koutei.KDAYC3)
        '-------------------------------------------------------------------------------
        Koutei.SUMITBC3 = "0"                   ' SUMIT���M�t���O
        Koutei.SNDKC3 = "0"                     ' ���M�t���O
'        Koutei.SNDDAYC3 = ""                   ' ���M���t
        Koutei.MODMACOC3 = "00"                 ' �ԍ��̏�����
        Koutei.KAKUCC3 = " "                    ' �m��R�[�h
        'upd start 2003/03/25 hitec)matsumoto ��ʂŎg�p���Ă���f�[�^�̂ݍX�V���s���B -------------
        Select Case Kihon.NOWPROC
            Case "CW750"    '��������
                If Koutei.SXLIDC3 = Trim(f_cmbc039_2.txtSxlID.Text) Then
                    iRtn = CreateXSDC3(Koutei, wErrMsg)
                    '�H�����ђǉ��G���[
                    If iRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox wErrMsg
                        Exit Function
                    End If
                End If
            Case "CW760"    '�Ĕ���
                If (SIngotP <= Koutei.INPOSC3) And (Koutei.INPOSC3 < EIngotP) Then
                    iRtn = CreateXSDC3(Koutei, wErrMsg)
                    '�H�����ђǉ��G���[
                    If iRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox wErrMsg
                        Exit Function
                    End If
                End If
            Case "CW800"    '�V���O���m��   03/05/01 Add.�㓡
                If Koutei.SXLIDC3 = strSxlData Then
                    iRtn = CreateXSDC3(Koutei, wErrMsg)
                    '�H�����ђǉ��G���[
                    If iRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox wErrMsg
                        Exit Function
                    End If
                End If
            Case Else
                iRtn = CreateXSDC3(Koutei, wErrMsg)
                '�H�����ђǉ��G���[
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox wErrMsg
                    Exit Function
                End If
        End Select
        'upd end  2003/03/25 hitec)matsumoto ------------------------------------------
    Next i

    XSDC3Proc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function


'''''�T�v      :���������i�r�w�k�j�o�^�������s��
'''''���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'''''      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'''''����      :
''''
''''Public Function XSDCBProc()
''''
'''''   �����ϐ�
''''    Dim i               As Integer
''''    Dim iRtn            As Integer          '���A���
''''    Dim sql             As String           '�r�p�k
''''    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
''''    Dim sqlWhere        As String           'WHERE��
''''    Dim wGNLCA          As Long             '���������i�i�ԁj�̍��v����
''''    Dim wGNMCA          As Long             '���������i�i�ԁj�̍��v����
''''    Dim wLENCB          As Long             '���v����
''''    Dim wMAICB          As Long             '���v����
''''    Dim wPUCUTMC4       As Long             '�s�Ǔ���̍��v�s�ǖ���
''''    Dim wPUCUTMCB       As Long             '���v�s�ǖ���
''''    Dim wErrMsg         As String
''''    Dim SXL()           As typ_XSDCB_Update '��������(�r�w�k)
''''    Dim wSXL()          As typ_XSDCB        '��������(�r�w�k)
''''    Dim intDataCnt      As Integer          '�Y���f�[�^����
''''    Dim strBlockID      As String
''''
''''    '�G���[�n���h���̐ݒ�
''''    On Error GoTo proc_err
''''
''''    XSDCBProc = FUNCTION_RETURN_FAILURE
''''
''''    For i = 0 To Kihon.CNTHINNOW - 1
''''        '���������i�i�ԁj���瓯���r�w�k�h�c�̒����̍��v���擾
''''        sql = "SELECT SUM(GNLCA) AS wGNLCA, "
''''        sql = sql & " SUM(GNMCA) AS wGNMCA "
''''        sql = sql & " FROM XSDCA "
''''        sql = sql & " WHERE SXLIDCA = '" & HinNow(i).SXLIDCA & "' "
''''        sql = sql & " AND LIVKCA = '0' "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '���݂��Ȃ����A�G���[
''''        If rs Is Nothing Then
''''            MsgBox "XSDCA SELECT ERROR"
''''            Exit Function
''''        End If
''''
''''        '���o���ʂ��i�[����
''''        If IsNull(rs.Fields("wGNLCA")) = False Then
''''            wLENCB = rs.Fields("wGNLCA")
''''        Else
''''            wLENCB = 0
''''        End If
''''        If IsNull(rs.Fields("wGNMCA")) = False Then
''''            wMAICB = rs.Fields("wGNMCA")
''''        Else
''''            wMAICB = 0
''''        End If
''''
''''        '�s�Ǔ��󂩂瓯���r�w�k�h�c�̕s�ǖ����̍��v���擾
''''        sql = "SELECT SUM(PUCUTMC4) AS wPUCUTMC4 "
''''        sql = sql & " FROM XSDC4 "
''''        sql = sql & " WHERE SXLIDC4 = '" & HinNow(i).SXLIDCA & "' "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '���݂��Ȃ����A�G���[
''''        If rs Is Nothing Then
''''            MsgBox "XSDC4 SELECT ERROR"
''''            Exit Function
''''        End If
''''
''''        '���o���ʂ��i�[����
''''        If IsNull(rs.Fields("wPUCUTMC4")) = False Then
''''            wPUCUTMCB = rs.Fields("wPUCUTMC4")
''''        Else
''''            wPUCUTMCB = 0
''''        End If
''''
''''        '���������i�i�ԁj�F�Ǖi�̂r�w�k�h�c�ŕ��������i�r�w�k�j������
''''        sqlWhere = "WHERE SXLIDCB = '" & HinNow(i).SXLIDCA & "' "
''''        ReDim wSXL(0) As typ_XSDCB
''''
''''        '�f�[�^�̌������擾
''''        iRtn = SelCntXSDCB(sqlWhere, intDataCnt)
''''        If iRtn = FUNCTION_RETURN_FAILURE Then  '�G���[
''''            MsgBox "XSDCB SELECT ERROR"
''''            Exit Function
''''        Else                                    '����
''''            '�f�[�^�����݂���ꍇ��UPDATE
''''            If intDataCnt > 0 Then
''''                iRtn = DBDRV_GetXSDCB(wSXL(), sqlWhere)
''''                If iRtn = FUNCTION_RETURN_FAILURE Then  '�G���[
''''                    MsgBox "XSDCA SELECT ERROR"
''''                    Exit Function
''''                End If
''''
''''                ReDim SXL(0) As typ_XSDCB_Update
''''
''''                '���������i�r�w�k�j���X�V
''''                SXL(0).LENCB = wLENCB
''''                SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''
''''                iRtn = UpdateXSDCB(SXL(0), sqlWhere)
''''                '���������i�r�w�k�j�X�V�G���[
''''                If iRtn = FUNCTION_RETURN_FAILURE Then
''''                    MsgBox "XSDCB UPDATET ERROR"
''''                    Exit Function
''''                End If
''''             '���݂��Ȃ����A�ǉ�
''''             ElseIf intDataCnt = 0 Then
''''                 ReDim SXL(0) As typ_XSDCB_Update
''''                 SXL(0).SXLIDCB = HinNow(i).SXLIDCA      ' SXLID
''''                 SXL(0).KCNTCB = HinNow(i).KCKNTCA       ' �H���A��
''''                 SXL(0).XTALCB = HinNow(i).XTALCA        ' �����ԍ�
''''                 SXL(0).INPOSCB = HinNow(i).INPOSCA      ' �������J�n�ʒu
''''                 SXL(0).LENCB = wLENCB                   ' ����
''''                 SXL(0).HINBCB = HinNow(i).HINBCA        ' �i��
''''                 SXL(0).REVNUMCB = HinNow(i).REVNUMCA    ' �d�b�ԍ������ԍ�
''''                 SXL(0).FACTORYCB = HinNow(i).FACTORYCA  ' �H��
''''                 SXL(0).OPECB = HinNow(i).OPECA          ' ���Ə���
''''                 SXL(0).MAICB = wMAICB                   ' ������
''''                 SXL(0).WSRMAICB = 0                     ' WS��㖇��
''''                 SXL(0).WSNMAICB = 0                     ' WS��򌇗�����
''''                 SXL(0).WFCMAICB = 0                     ' �������
''''
''''                 SXL(0).SXLRMAICB = 0                    ' SXL�w��(�Ǖi)
''''                 SXL(0).SXLNMAICB = 0                    ' SXL�w��(�s��)
''''                 SXL(0).WFCNMAICB = 0                    ' WFC����������
''''                 SXL(0).SXLEMAICB = 0                    ' SXL�m�薇��
''''                 SXL(0).SRMAICB = 0                      ' �T���v�����w��(�Ǖi)
''''                 SXL(0).SNMAICB = 0                      ' �T���v�����w��(�s��)
''''                 SXL(0).STMAICB = 0                      ' �T���v������
''''                 '�H���ɂ��U���i�Ƃ肠������ʕ��j
''''                 Select Case Kihon.NOWPROC
''''                     Case "CW740"
''''                         SXL(0).SRMAICB = wMAICB         ' �T���v�����w��(�Ǖi)
''''                         SXL(0).SNMAICB = wPUCUTMCB      ' �T���v�����w��(�s��)
''''                     Case "CW800"
''''                         SXL(0).SXLEMAICB = wMAICB       ' SXL�m�薇��
''''                 End Select
''''                 SXL(0).FURIMAICB = ""                 ' �U�֖���
''''                 SXL(0).XTWORKCB = "42"                  ' �����H��
''''                 SXL(0).WFWORKCB = " "                   ' �E�F�[�n����
''''                 SXL(0).FURYCCB = " "                     ' �s�Ǘ��R
''''                 SXL(0).LSTCCB = "T"                     ' �̎��ԋ敪
''''                 SXL(0).LUFRCCB = " "                    ' �i��R�[�h
''''                 SXL(0).LUFRBCB = " "                    ' �i��敪
''''                 SXL(0).LDERCCB = " "                    ' �i���R�[�h
''''                 SXL(0).LDFRBCB = "0"                    ' �i���敪
''''                 SXL(0).HOLDCCB = " "                    ' �z�[���h�R�[�h
''''                 SXL(0).HOLDBCB = " "                    ' �z�[���h�敪
''''                 SXL(0).EXKUBCB = " "                    ' ��O�敪
''''                 SXL(0).HENPKCB = " "                    ' �ԕi�敪
''''''''                 '�V���O���m�莞�A�����敪='2'�E�����敪='1'�ɂ���
''''''''                 If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
''''''''                    SXL(0).KANKCB = "2"                ' �����敪
''''''''                    SXL(0).LIVKCB = "1"                 ' �����敪
''''''''                 Else
''''                    SXL(0).KANKCB = "0"                ' �����敪
''''                    SXL(0).LIVKCB = "0"                 ' �����敪
''''''''                 End If
''''                 SXL(0).NFCB = "0"                       ' ���ɋ敪
''''                 SXL(0).SAKJCB = "0"                     ' �폜�敪
''''                                                      ' �o�^���t
''''                 SXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''                                                      ' �X�V���t
''''                 SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''                 SXL(0).SUMITCB = "0"                    ' SUMIT���M�t���O
''''                 SXL(0).SNDKCB = "0"                    ' �ԕi�敪
''''                 SXL(0).SNDAYCB = ""                   ' ���M���t
''''
''''                 iRtn = CreateXSDCB(SXL(0), wErrMsg)
''''                 '���������i�r�w�k�j�ǉ��G���[
''''                 If iRtn = FUNCTION_RETURN_FAILURE Then
''''                     MsgBox wErrMsg
''''                     Exit Function
''''                 End If
''''             End If
''''        End If
''''    Next i
''''
''''''''' 02/09/20 Add By ��c@HITEC  sta�F���������i�i�ԁj�̍��v������0�ɂȂ������A
'''''''''                                 ���������i�r�w�k�j�������b�g�ɂ���
''''    For i = 0 To Kihon.CNTHINOLD - 1
''''        '���������i�i�ԁj�O�H�����玀���b�g�̓����r�w�k�h�c�̒����̍��v���擾
''''        sql = "SELECT SUM(GNLCA) AS wGNLCA, "
''''        sql = sql & " SUM(GNMCA) AS wGNMCA, "
''''        sql = sql & " max(CRYNUMCA) AS CRYNUMCA"
''''        sql = sql & " FROM XSDCA "
''''        sql = sql & " WHERE SXLIDCA = '" & HinOld(i).SXLIDCA & "' "
''''        sql = sql & " AND LIVKCA = '0' "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '���݂��Ȃ����A�G���[
''''        If rs Is Nothing Then
''''            MsgBox "XSDCA SELECT ERROR"
''''            Exit Function
''''        End If
''''
''''        '���o���ʂ��i�[����
''''        If IsNull(rs.Fields("wGNLCA")) = False Then
''''            wLENCB = rs.Fields("wGNLCA")
''''        Else
''''            wLENCB = 0
''''        End If
''''        If IsNull(rs.Fields("wGNMCA")) = False Then
''''            wMAICB = rs.Fields("wGNMCA")
''''        Else
''''            wMAICB = 0
''''        End If
''''
''''        '�s�Ǔ��󂩂瓯���r�w�k�h�c�̕s�ǖ����̍��v���擾
''''        sql = "SELECT SUM(PUCUTMC4) AS wPUCUTMC4 "
''''        sql = sql & " FROM XSDC4 "
''''        sql = sql & " WHERE SXLIDC4 = '" & HinOld(i).SXLIDCA & "' "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '���݂��Ȃ����A�G���[
''''        If rs Is Nothing Then
''''            MsgBox "XSDC4 SELECT ERROR"
''''            Exit Function
''''        End If
''''
''''        '���o���ʂ��i�[����
''''        If IsNull(rs.Fields("wPUCUTMC4")) = False Then
''''            wPUCUTMCB = rs.Fields("wPUCUTMC4")
''''        Else
''''            wPUCUTMCB = 0
''''        End If
''''
''''        '���������i�i�ԁj�F�r�w�k�h�c�ŕ��������i�r�w�k�j������
''''        sqlWhere = "WHERE SXLIDCB = '" & HinOld(i).SXLIDCA & "' "
''''        ReDim wSXL(0) As typ_XSDCB
''''
''''        '�f�[�^�̌������擾
''''        iRtn = SelCntXSDCB(sqlWhere, intDataCnt)
''''        If iRtn = FUNCTION_RETURN_FAILURE Then  '�G���[
''''            MsgBox "XSDCB SELECT ERROR"
''''            Exit Function
''''        Else                                    '����
''''            '�f�[�^�����݂���ꍇ��UPDATE
''''            If intDataCnt > 0 Then
''''                iRtn = DBDRV_GetXSDCB(wSXL(), sqlWhere)
''''                If iRtn = FUNCTION_RETURN_FAILURE Then  '�G���[
''''                    MsgBox "XSDCA SELECT ERROR"
''''                    Exit Function
''''                End If
''''
''''                ReDim sxl(0) As typ_XSDCB_Update
''''
''''                '���������i�r�w�k�j���X�V
''''                SXL(0).LENCB = wLENCB
''''                '������0�̎��A�����b�g�Ƃ���
''''                If wLENCB = 0 Then
''''                    SXL(0).LDFRBCB = "2"
''''                    SXL(0).LIVKCB = "1"                 ' �����敪
''''                    SXL(0).LDFRBCB = "H"
''''                    SXL(0).KANKCB = "2"
''''                End If
''''
''''                SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''
''''                iRtn = UpdateXSDCB(SXL(0), sqlWhere)
''''                '���������i�r�w�k�j�X�V�G���[
''''                If iRtn = FUNCTION_RETURN_FAILURE Then
''''                    MsgBox "XSDCB UPDATET ERROR"
''''                    Exit Function
''''                End If
''''             '���݂��Ȃ����A�ǉ�
''''             ElseIf intDataCnt = 0 Then
''''                 ReDim SXL(0) As typ_XSDCB_Update
''''                 SXL(0).SXLIDCB = HinNow(i).SXLIDCA      ' SXLID
''''                 SXL(0).KCNTCB = HinNow(i).KCKNTCA       ' �H���A��
''''                 SXL(0).XTALCB = HinNow(i).XTALCA        ' �����ԍ�
''''                 SXL(0).INPOSCB = HinNow(i).INPOSCA      ' �������J�n�ʒu
''''                 SXL(0).LENCB = wLENCB                   ' ����
''''                 SXL(0).HINBCB = HinNow(i).HINBCA        ' �i��
''''                 SXL(0).REVNUMCB = HinNow(i).REVNUMCA    ' �d�b�ԍ������ԍ�
''''                 SXL(0).FACTORYCB = HinNow(i).FACTORYCA  ' �H��
''''                 SXL(0).OPECB = HinNow(i).OPECA          ' ���Ə���
''''                 SXL(0).MAICB = wMAICB                   ' ������
''''                 SXL(0).WSRMAICB = 0                     ' WS��㖇��
''''                 SXL(0).WSNMAICB = 0                     ' WS��򌇗�����
''''                 SXL(0).WFCMAICB = 0                     ' �������
''''                 SXL(0).SXLRMAICB = 0                    ' SXL�w��(�Ǖi)
''''                 SXL(0).SXLNMAICB = 0                    ' SXL�w��(�s��)
''''                 SXL(0).WFCNMAICB = 0                    ' WFC����������
''''                 SXL(0).SXLEMAICB = 0                    ' SXL�m�薇��
''''                 SXL(0).SRMAICB = 0                      ' �T���v�����w��(�Ǖi)
''''                 SXL(0).SNMAICB = 0                      ' �T���v�����w��(�s��)
''''                 SXL(0).STMAICB = 0                      ' �T���v������
''''                 '�H���ɂ��U���i�Ƃ肠������ʕ��j
''''                 Select Case Kihon.NOWPROC
''''                     Case "CW740"
''''                         SXL(0).SRMAICB = wMAICB         ' �T���v�����w��(�Ǖi)
''''                         SXL(0).SNMAICB = wPUCUTMCB      ' �T���v�����w��(�s��)
''''                     Case "CW800"
''''                         SXL(0).SXLEMAICB = wMAICB       ' SXL�m�薇��
''''                 End Select
''''                 SXL(0).FURIMAICB = ""                 ' �U�֖���
''''                 SXL(0).XTWORKCB = "42"                  ' �����H��
''''                 SXL(0).WFWORKCB = " "                   ' �E�F�[�n����
''''                 SXL(0).FURYCCB = " "                     ' �s�Ǘ��R
''''                 SXL(0).LSTCCB = "T"                     ' �̎��ԋ敪
''''                 SXL(0).LUFRCCB = " "                    ' �i��R�[�h
''''                 SXL(0).LUFRBCB = " "                    ' �i��敪
''''                 SXL(0).LDERCCB = " "                    ' �i���R�[�h
''''                '������0�̎��A�p���Ƃ���
''''                 If wLENCB = 0 Then
''''                     SXL(0).LDFRBCB = "2"                ' �i���敪
''''                 Else
''''                     SXL(0).LDFRBCB = "0"
''''                 End If
''''                 SXL(0).HOLDCCB = " "                    ' �z�[���h�R�[�h
''''                 SXL(0).HOLDBCB = " "                    ' �z�[���h�敪
''''                 SXL(0).EXKUBCB = " "                    ' ��O�敪
''''                 SXL(0).HENPKCB = " "                    ' �ԕi�敪
''''                 '������0�̎��A�����b�g�Ƃ���
''''                 If wLENCB = 0 Then
''''                     SXL(0).LIVKCB = "1"                 ' �����敪
''''                     SXL(0).KANKCB = "2"                    ' �����敪
''''                 Else
''''                     SXL(0).LIVKCB = "0"
''''                     SXL(0).KANKCB = "0"                    ' �����敪
''''                 End If
''''                 SXL(0).KANKCB = "0"                    ' �����敪
''''                 SXL(0).NFCB = "0"                       ' ���ɋ敪
''''                 SXL(0).SAKJCB = "0"                     ' �폜�敪
''''                                                      ' �o�^���t
''''                 SXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''                                                      ' �X�V���t
''''                 SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''                 SXL(0).SUMITCB = "0"                    ' SUMIT���M�t���O
''''                 SXL(0).SNDKCB = "0"                    ' �ԕi�敪
''''                 SXL(0).SNDAYCB = ""                   ' ���M���t
''''
''''                 iRtn = CreateXSDCB(SXL(0), wErrMsg)
''''                 '���������i�r�w�k�j�ǉ��G���[
''''                 If iRtn = FUNCTION_RETURN_FAILURE Then
''''                     MsgBox wErrMsg
''''                     Exit Function
''''                 End If
''''             End If
''''        End If
''''    Next i
''''
''''''''' 02/09/20 Add                end
''''
''''    XSDCBProc = FUNCTION_RETURN_SUCCESS
''''
''''proc_exit:
''''    '' �I��
'''''    gErr.Pop
''''    Exit Function
''''
''''proc_err:
''''    '' �G���[�n���h��
''''    Debug.Print "====== Error SQL ======"
''''    Debug.Print sql
''''    XSDCBProc = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    sErrMsg = GetMsgStr("EXXXX", sDBName)
''''    Resume proc_exit
''''
''''End Function
''''




'�T�v      :���������i�r�w�k�j�o�^�������s��
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :

Public Function XSDCBProc()
    
'   �����ϐ�
    Dim i               As Integer
    Dim iRtn            As Integer          '���A���
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim sqlWhere        As String           'WHERE��
    Dim wGNLCA          As Long             '���������i�i�ԁj�̍��v����
    Dim wGNMCA          As Long             '���������i�i�ԁj�̍��v����
'    Dim wLENCB          As Long             '���v����
'    Dim wMAICB          As Long             '���v����
'    Dim wPUCUTMC4       As Long             '�s�Ǔ���̍��v�s�ǖ���
'    Dim wPUCUTMCB       As Long             '���v�s�ǖ���
    Dim wErrMsg         As String
    Dim SXL()           As typ_XSDCB_Update '��������(�r�w�k)
    Dim wSXL()          As typ_XSDCB        '��������(�r�w�k)
    Dim intDataCnt      As Integer          '�Y���f�[�^����
    Dim strBlockID      As String
''''' 02/09/21 Add bY ��c@hitec  sta
    Dim wRYOMAI         As Long             '�H�����̗Ǖi����
    Dim wFRYMAI         As Long             '�H�����̕s�Ǖi����
    Dim wLen            As Long             '����
    Dim wMAI            As Long             '����
    Dim wMAI800         As Long             'CW800����
    Dim wFUR            As Long             '�s�ǖ���
    Dim wFURKEI         As Long             '�s�ǖ������v
    Dim wSAM            As Long             '�T���v������
    Dim wSIJ            As Long             '�T���v�����w������
    Dim wSAMFUR         As Long             '�T���v�����w���s�ǖ���
''''' 02/09/21 Add                end
    Dim iLoopBkHinGet   As Integer          '���i�� 'add 2003/04/03 hitec)matsumoto
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    XSDCBProc = FUNCTION_RETURN_FAILURE
    
    For i = 0 To Kihon.CNTHINNOW - 1
''''' 02/09/21 Dlt bY ��c@hitec  sta
'        '���������i�i�ԁj���瓯���r�w�k�h�c�̒����̍��v���擾
'        sql = "SELECT SUM(GNLCA) AS wGNLCA, "
'        sql = sql & " SUM(GNMCA) AS wGNMCA "
'        sql = sql & " FROM XSDCA "
'        sql = sql & " WHERE SXLIDCA = '" & HinNow(i).SXLIDCA & "' "
'        sql = sql & " AND LIVKCA = '0' "
'
'        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'        '���݂��Ȃ����A�G���[
'        If rs Is Nothing Then
'            MsgBox "XSDCA SELECT ERROR"
'            Exit Function
'        End If
'
'        '���o���ʂ��i�[����
'        If IsNull(rs.Fields("wGNLCA")) = False Then
'            wLENCB = rs.Fields("wGNLCA")
'        Else
'            wLENCB = 0
'        End If
'        If IsNull(rs.Fields("wGNMCA")) = False Then
'            wMAICB = rs.Fields("wGNMCA")
'        Else
'            wMAICB = 0
'        End If
'
'        '�s�Ǔ��󂩂瓯���r�w�k�h�c�̕s�ǖ����̍��v���擾
'        sql = "SELECT SUM(PUCUTMC4) AS wPUCUTMC4 "
'        sql = sql & " FROM XSDC4 "
'        sql = sql & " WHERE SXLIDC4 = '" & HinNow(i).SXLIDCA & "' "
'
'        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'        '���݂��Ȃ����A�G���[
'        If rs Is Nothing Then
'            MsgBox "XSDC4 SELECT ERROR"
'            Exit Function
'        End If
'
'        '���o���ʂ��i�[����
'        If IsNull(rs.Fields("wPUCUTMC4")) = False Then
'            wPUCUTMCB = rs.Fields("wPUCUTMC4")
'        Else
'            wPUCUTMCB = 0
'        End If
'
''''' 02/09/21 Dlt                end

''''' 02/09/21 Add bY ��c@hitec  sta
'        '�H�����т��瓯���r�w�k�h�c�̒����A�����A�s�ǖ����̍��v���擾
        iRtn = XSDCBSum(Kihon.NOWPROC, HinNow(i).SXLIDCA, wLen, wMAI, wMAI800, wFUR, wFURKEI, wSAM, wSAMSIJ, wSAMFUR)
''''' 02/09/21 Add                end

        '���������i�i�ԁj�F�Ǖi�̂r�w�k�h�c�ŕ��������i�r�w�k�j������
        sqlWhere = "WHERE SXLIDCB = '" & HinNow(i).SXLIDCA & "' "
        ReDim wSXL(0) As typ_XSDCB

        '�f�[�^�̌������擾
        iRtn = SelCntXSDCB(sqlWhere, intDataCnt)
        If iRtn = FUNCTION_RETURN_FAILURE Then  '�G���[
            MsgBox "XSDCB SELECT ERROR"
            Exit Function
        Else                                    '����
            '�f�[�^�����݂���ꍇ��UPDATE
            If intDataCnt > 0 Then
                iRtn = DBDRV_GetXSDCB(wSXL(), sqlWhere)
                If iRtn = FUNCTION_RETURN_FAILURE Then  '�G���[
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If
                
                ReDim SXL(0) As typ_XSDCB_Update
            
                '���������i�r�w�k�j���X�V
'''''                SXL(0).LENCB = wLENCB
                SXL(0).LENCB = wLen
' VVVVV 2003/05/02 ADD BY HITEC)��c�F�X�V�����i�ԕύX����
                SXL(0).HINBCB = HinNow(i).HINBCA        ' �i��
' ^^^^^ 2003/05/02 ADD BY HITEC)��c  END
                SXL(0).MAICB = wMAI
                SXL(0).KCNTCB = BlkNow.KCNTC2           ' �H���A��
                '�V���O���m�莞�A�ŏI��ԋ敪='S'�ɂ���
                SXL(0).LIVKCB = "0"
                SXL(0).KANKCB = "0"                 ' �����敪
                SXL(0).LSTCCB = "T"                 ' �ŏI��ԋ敪
                SXL(0).LDFRBCB = "0"                ' �i���敪
                'add start  2003/06/09 hitec)matsumoto ---------------
                SXL(0).INPOSCB = HinNow(i).INPOSCA      ' �������J�n�ʒu
                SXL(0).LENCB = wLen                     ' ����
                'add end    2003/06/09 hitec)matsumoto ---------------
                If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
                    SXL(0).LSTCCB = "S"
                End If
                 '�H���ɂ��U���i�Ƃ肠������ʕ��j
                Select Case Kihon.NOWPROC
                     Case "CW740"
                         SXL(0).SXLNMAICB = wFUR         ' �p��WF����
                         'add start 2003/03/25 hitec)matsumoto
                         SXL(0).NEWKNTCB = "CW740"       ' �ŏI�ʉߍH��
                         SXL(0).GNWKNTCB = "CW750"       ' ���ݍH��
                         'add end 2003/03/25 hitec)matsumoto
                     Case "CW750"
                         SXL(0).SRMAICB = wSIJ           ' �T���v�����w������
                         SXL(0).SNMAICB = wSAMFUR        ' �T���v�����w���s�ǖ���
                         SXL(0).STMAICB = wSAM           ' �T���v������
                         'add start 2003/10/22 tuku �� 03/11/05 �u���b�N�P�ʂŕύX����Ă��܂����߃R�����g��
'                         SXL(0).NEWKNTCB = "CW750"       ' �ŏI�ʉߍH��
'                         SXL(0).GNWKNTCB = "CW800"       ' ���ݍH��
                         'add end 2003/10/22 tuku
                         
                     Case "CW760"
                         SXL(0).SXLNMAICB = wFUR         ' �p��WF����
                         'add start 2003/10/22
                         SXL(0).NEWKNTCB = "CW760"       ' �ŏI�ʉߍH��
                         SXL(0).GNWKNTCB = "CW750"       ' ���ݍH��
                         'add end 2003/10/22 tuku
                     Case "CW800"
                         SXL(0).SXLRMAICB = wMAI800      ' SXL�w���i�Ǖi�j
                         SXL(0).WFCNMAICB = wFURKEI      ' WFC����������
                End Select
                SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                
                iRtn = UpdateXSDCB(SXL(0), sqlWhere)
                '���������i�r�w�k�j�X�V�G���[
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCB UPDATET ERROR"
                    Exit Function
                End If
             '���݂��Ȃ����A�ǉ�
             ElseIf intDataCnt = 0 Then
                 ReDim SXL(0) As typ_XSDCB_Update
                 SXL(0).SXLIDCB = HinNow(i).SXLIDCA      ' SXLID
                 SXL(0).KCNTCB = BlkNow.KCNTC2           ' �H���A��
                 SXL(0).XTALCB = HinNow(i).XTALCA        ' �����ԍ�
                 SXL(0).INPOSCB = HinNow(i).INPOSCA      ' �������J�n�ʒu
                 SXL(0).LENCB = wLen                     ' ����
                 SXL(0).HINBCB = HinNow(i).HINBCA        ' �i��
                 SXL(0).REVNUMCB = HinNow(i).REVNUMCA    ' �d�b�ԍ������ԍ�
                 SXL(0).FACTORYCB = HinNow(i).FACTORYCA  ' �H��
                 SXL(0).OPECB = HinNow(i).OPECA          ' ���Ə���
                 SXL(0).MAICB = wMAI                     ' ������
                 SXL(0).WSRMAICB = 0                     ' WS��㖇��
                 SXL(0).WSNMAICB = 0                     ' WS��򌇗�����
                 SXL(0).WFCMAICB = 0                     ' �������
                 SXL(0).SXLRMAICB = 0                    ' SXL�w��(�Ǖi)
                 SXL(0).SXLEMAICB = 0                    ' SXL�m�薇��
                 '�H���ɂ��U���i�Ƃ肠������ʕ��j
                 Select Case Kihon.NOWPROC
                     Case "CW740"
                         SXL(0).SXLNMAICB = wFUR         ' �p��WF����
                         'add start 2003/03/25 hitec)matsumoto
                         SXL(0).NEWKNTCB = "CW740"       ' �ŏI�ʉߍH��
                         SXL(0).GNWKNTCB = "CW750"       ' ���ݍH��
                         'add end 2003/03/25 hitec)matsumoto
                     Case "CW750"
                         SXL(0).SRMAICB = wSIJ           ' �T���v�����w������
                         SXL(0).SNMAICB = wSAMFUR        ' �T���v�����w���s�ǖ���
                         SXL(0).STMAICB = wSAM           ' �T���v������
                         'add start 2003/10/22 tuku�@�� 03/11/05 �u���b�N�P�ʂŕύX����Ă��܂����߃R�����g��
'                         SXL(0).NEWKNTCB = "CW750"       ' �ŏI�ʉߍH��
'                         SXL(0).GNWKNTCB = "CW800"       ' ���ݍH��
                         'add end 2003/10/22 tuku
                     Case "CW760"
                         SXL(0).SXLNMAICB = wFUR         ' �p��WF����
                         'add start 2003/10/22
                         SXL(0).NEWKNTCB = "CW760"       ' �ŏI�ʉߍH��
                         SXL(0).GNWKNTCB = "CW750"       ' ���ݍH��
                         'add end 2003/10/22 tuku
                     Case "CW800"
                         SXL(0).SXLRMAICB = wMAI         ' SXL�w���i�Ǖi�j
                         SXL(0).WFCNMAICB = wFURKEI      ' WFC����������
                 End Select
                 SXL(0).FURIMAICB = ""                   ' �U�֖���
                 SXL(0).XTWORKCB = "42"                  ' �����H��
                 SXL(0).WFWORKCB = " "                   ' �E�F�[�n����
                 SXL(0).FURYCCB = " "                     ' �s�Ǘ��R
                 SXL(0).LSTCCB = "T"                     ' �̎��ԋ敪
                 '�V���O���m�莞�A�ŏI��ԋ敪='S'�ɂ���
                 If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
                    SXL(0).LSTCCB = "S"
                 End If
                 SXL(0).LUFRCCB = " "                    ' �i��R�[�h
                 SXL(0).LUFRBCB = " "                    ' �i��敪
                 SXL(0).LDERCCB = " "                    ' �i���R�[�h
                 SXL(0).LDFRBCB = "0"                    ' �i���敪
                 SXL(0).HOLDCCB = " "                    ' �z�[���h�R�[�h
                 SXL(0).HOLDBCB = " "                    ' �z�[���h�敪
                 SXL(0).EXKUBCB = " "                    ' ��O�敪
                 SXL(0).HENPKCB = " "                    ' �ԕi�敪
''''                 '�V���O���m�莞�A�����敪='2'�E�����敪='1'�ɂ���
''''                 If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
''''                    SXL(0).KANKCB = "2"                ' �����敪
''''                    SXL(0).LIVKCB = "1"                ' �����敪
''''                 Else
                    SXL(0).KANKCB = "0"                  ' �����敪
                    SXL(0).LIVKCB = "0"                  ' �����敪
''''                 End If
                 SXL(0).NFCB = "0"                       ' ���ɋ敪
                 SXL(0).SAKJCB = "0"                     ' �폜�敪
                                                         ' �o�^���t
                 SXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                                                      ' �X�V���t
                 SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                 SXL(0).SUMITCB = "0"                    ' SUMIT���M�t���O
                 SXL(0).SNDKCB = "0"                     ' �ԕi�敪
                 SXL(0).SNDAYCB = ""                     ' ���M���t
            
                 'add start 2003/04/03 hitec)matsumoto �O�ް����Ȃ��A���i�Ԃ��擾�ł��Ȃ��̂ŁAHINOLD����Y���ʒu�̕i�Ԃ��擾���A��������i�ԂƂ���@---------------
                 SXL(0).MOTHINCB = vbNullString '������
' VVVVV 2003/04/30 ALT BY HITEC)��c�FCW740,CW760�݂̂ɕύX
                 If Kihon.NOWPROC = "CW740" Or Kihon.NOWPROC = "CW760" Then
                    For iLoopBkHinGet = 0 To Kihon.CNTHINOLD - 1
                        If (CInt(HinOld(iLoopBkHinGet).INPOSCA) <= CInt(SXL(0).INPOSCB)) And (CInt(SXL(0).INPOSCB) <= CInt(HinOld(iLoopBkHinGet).INPOSCA) + CInt(HinOld(iLoopBkHinGet).GNLCA)) Then
                             SXL(0).MOTHINCB = HinOld(iLoopBkHinGet).HINBCA
                             Exit For
                        End If
                    Next
                    If SXL(0).MOTHINCB = vbNullString Then '�����Y��HINOLD�����������玩���̕i�Ԃ����i�ԂƂ���
                        SXL(0).MOTHINCB = SXL(0).HINBCB
                    End If
                 End If
' ^^^^^^ 2003/04/30 ALT BY HITEC)��c  END
                 'add end   2003/04/03 hitec)matsumoto ---------------
            
                 iRtn = CreateXSDCB(SXL(0), wErrMsg)
                 '���������i�r�w�k�j�ǉ��G���[
                 If iRtn = FUNCTION_RETURN_FAILURE Then
                     MsgBox wErrMsg
                     Exit Function
                 End If
             End If
        End If
    Next i
    
''''' 02/09/20 Add By ��c@HITEC  sta�F���������i�i�ԁj�̍��v������0�ɂȂ������A
'''''                                 ���������i�r�w�k�j�������b�g�ɂ���
    For i = 0 To Kihon.CNTHINOLD - 1
''''' 02/09/21 Dlt bY ��c@hitec  sta
'        '���������i�i�ԁj�O�H�����玀���b�g�̓����r�w�k�h�c�̒����̍��v���擾
'        sql = "SELECT SUM(GNLCA) AS wGNLCA, "
'        sql = sql & " SUM(GNMCA) AS wGNMCA, "
'        sql = sql & " max(CRYNUMCA) AS CRYNUMCA"
'        sql = sql & " FROM XSDCA "
'        sql = sql & " WHERE SXLIDCA = '" & HinOld(i).SXLIDCA & "' "
'        sql = sql & " AND LIVKCA = '0' "
'
'        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'        '���݂��Ȃ����A�G���[
'        If rs Is Nothing Then
'            MsgBox "XSDCA SELECT ERROR"
'            Exit Function
'        End If
'
'        '���o���ʂ��i�[����
'        If IsNull(rs.Fields("wGNLCA")) = False Then
'            wLENCB = rs.Fields("wGNLCA")
'        Else
'            wLENCB = 0
'        End If
'        If IsNull(rs.Fields("wGNMCA")) = False Then
'            wMAICB = rs.Fields("wGNMCA")
'        Else
'            wMAICB = 0
'        End If
'
'        '�s�Ǔ��󂩂瓯���r�w�k�h�c�̕s�ǖ����̍��v���擾
'        sql = "SELECT SUM(PUCUTMC4) AS wPUCUTMC4 "
'        sql = sql & " FROM XSDC4 "
'        sql = sql & " WHERE SXLIDC4 = '" & HinOld(i).SXLIDCA & "' "
'
'        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'        '���݂��Ȃ����A�G���[
'        If rs Is Nothing Then
'            MsgBox "XSDC4 SELECT ERROR"
'            Exit Function
'        End If
'
'        '���o���ʂ��i�[����
'        If IsNull(rs.Fields("wPUCUTMC4")) = False Then
'            wPUCUTMCB = rs.Fields("wPUCUTMC4")
'        Else
'            wPUCUTMCB = 0
'        End If
    
''''' 02/09/21 Dlt                end

''''' 02/09/21 Add bY ��c@hitec  sta
'        '�H�����т��瓯���r�w�k�h�c�̒����A�����A�s�ǖ����̍��v���擾
        iRtn = XSDCBSum(Kihon.NOWPROC, HinOld(i).SXLIDCA, wLen, wMAI, pMAI800, wFUR, wFURKEI, wSAM, wSAMSIJ, wSAMFUR)
''''' 02/09/21 Add                end
        
        '���������i�i�ԁj�F�r�w�k�h�c�ŕ��������i�r�w�k�j������
        sqlWhere = "WHERE SXLIDCB = '" & HinOld(i).SXLIDCA & "' "
        ReDim wSXL(0) As typ_XSDCB
        
        '�f�[�^�̌������擾
        iRtn = SelCntXSDCB(sqlWhere, intDataCnt)
        If iRtn = FUNCTION_RETURN_FAILURE Then  '�G���[
            MsgBox "XSDCB SELECT ERROR"
            Exit Function
        Else                                    '����
            '�f�[�^�����݂���ꍇ��UPDATE
            If intDataCnt > 0 Then
                iRtn = DBDRV_GetXSDCB(wSXL(), sqlWhere)
                If iRtn = FUNCTION_RETURN_FAILURE Then  '�G���[
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If
                
                ReDim SXL(0) As typ_XSDCB_Update
            
                '���������i�r�w�k�j���X�V
'''                SXL(0).LENCB = wLENCB
                SXL(0).LENCB = wLen
''''' VVVVV 2003/05/02 ADD BY HITEC)��c�F�X�V�����i�ԕύX����
''''                SXL(0).HINBCB = HinNow(i).HINBCA        ' �i��
''''' ^^^^^ 2003/05/02 ADD BY HITEC)��c  END
                SXL(0).MAICB = wMAI
                SXL(0).KCNTCB = BlkNow.KCNTC2
                '�H���ɂ��U���i�Ƃ肠������ʕ��j
                Select Case Kihon.NOWPROC
                     Case "CW740"
                         SXL(0).SXLNMAICB = wFUR         ' �p��WF����
                     Case "CW750"
                         SXL(0).SRMAICB = wSIJ           ' �T���v�����w������
                         SXL(0).SNMAICB = wSAMFUR        ' �T���v�����w���s�ǖ���
                         SXL(0).STMAICB = wSAM           ' �T���v������
                     Case "CW760"
                         SXL(0).SXLNMAICB = wFUR         ' �p��WF����
                     Case "CW800"
                         SXL(0).SXLRMAICB = wMAI         ' SXL�w���i�Ǖi�j
                         SXL(0).WFCNMAICB = wFURKEI      ' WFC����������
                End Select
                '������0�̎��A�����b�g�Ƃ���
''''                If wLen = 0 Then
'                If wMAI = 0 Then    'upd 2003/06/05 hitec)matsumoto
                If (wMAI = 0 And Kihon.NOWPROC <> PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Or _
                        (wLen = 0 And Kihon.NOWPROC = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Then    '05/03/29 ooba
                     SXL(0).LIVKCB = "1"                 ' �����敪
                     SXL(0).KANKCB = "2"                 ' �����敪
                     SXL(0).LSTCCB = "H"                 ' �ŏI��ԋ敪
                     SXL(0).LDFRBCB = "2"                ' �i���敪
                 Else
                     SXL(0).LIVKCB = "0"
                     SXL(0).KANKCB = "0"                 ' �����敪
                     SXL(0).LSTCCB = "T"                 ' �ŏI��ԋ敪
                     SXL(0).LDFRBCB = "0"                ' �i���敪
                End If
                ' 2002/12/13 ooba �����敪�t���O�ύX
                SXL(0).KANKCB = "0"                 ' �����敪

                 '�V���O���m�莞�A�ŏI��ԋ敪='S'�ɂ���
                 If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
                    SXL(0).LSTCCB = "S"
                 End If
                SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
            
                iRtn = UpdateXSDCB(SXL(0), sqlWhere)
                '���������i�r�w�k�j�X�V�G���[
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCB UPDATET ERROR"
                    Exit Function
                End If
             '���݂��Ȃ����A�ǉ�
             ElseIf intDataCnt = 0 Then
                 ReDim SXL(0) As typ_XSDCB_Update
                 SXL(0).SXLIDCB = HinOld(i).SXLIDCA      ' SXLID
                 SXL(0).KCNTCB = BlkNow.KCNTC2       ' �H���A��
                 SXL(0).XTALCB = HinOld(i).XTALCA        ' �����ԍ�
                 SXL(0).INPOSCB = HinOld(i).INPOSCA      ' �������J�n�ʒu
                 SXL(0).LENCB = wLen                     ' ����
                 SXL(0).HINBCB = HinOld(i).HINBCA        ' �i��
                 SXL(0).REVNUMCB = HinOld(i).REVNUMCA    ' �d�b�ԍ������ԍ�
                 SXL(0).FACTORYCB = HinOld(i).FACTORYCA  ' �H��
                 SXL(0).OPECB = HinOld(i).OPECA          ' ���Ə���
                 SXL(0).MAICB = wMAI                     ' ������
                 SXL(0).WSRMAICB = 0                     ' WS��㖇��
                 SXL(0).WSNMAICB = 0                     ' WS��򌇗�����
                 SXL(0).WFCMAICB = 0                     ' �������
                 SXL(0).WSNMAICB = 0                     ' WS��򌇗�����
                 SXL(0).WFCMAICB = 0                     ' �������
                 SXL(0).SXLEMAICB = 0                    ' SXL�m�薇��
                 '�H���ɂ��U���i�Ƃ肠������ʕ��j
                 Select Case Kihon.NOWPROC
                     Case "CW740"
                         SXL(0).SXLNMAICB = wFUR         ' �p��WF����
                     Case "CW750"
                         SXL(0).SRMAICB = wSIJ           ' �T���v�����w������
                         SXL(0).SNMAICB = wSAMFUR        ' �T���v�����w���s�ǖ���
                         SXL(0).STMAICB = wSAM           ' �T���v������
                     Case "CW760"
                         SXL(0).SXLNMAICB = wFUR         ' �p��WF����
                     Case "CW800"
                         SXL(0).SXLRMAICB = wMAI         ' SXL�w���i�Ǖi�j
                         SXL(0).WFCNMAICB = wFURKEI      ' WFC����������
                 End Select
                 SXL(0).FURIMAICB = ""                   ' �U�֖���
                 SXL(0).XTWORKCB = "42"                  ' �����H��
                 SXL(0).WFWORKCB = " "                   ' �E�F�[�n����
                 SXL(0).FURYCCB = " "                    ' �s�Ǘ��R
                 SXL(0).LSTCCB = "T"                     ' �̎��ԋ敪
                 SXL(0).LUFRCCB = " "                    ' �i��R�[�h
                 SXL(0).LUFRBCB = " "                    ' �i��敪
                 SXL(0).LDERCCB = " "                    ' �i���R�[�h
                '������0�̎��A�p���Ƃ���
                 If wLENCB = 0 Then
                     SXL(0).LDFRBCB = "2"                ' �i���敪
                 Else
                     SXL(0).LDFRBCB = "0"
                 End If
                 SXL(0).HOLDCCB = " "                    ' �z�[���h�R�[�h
                 SXL(0).HOLDBCB = " "                    ' �z�[���h�敪
                 SXL(0).EXKUBCB = " "                    ' ��O�敪
                 SXL(0).HENPKCB = " "                    ' �ԕi�敪
                 '������0�̎��A�����b�g�Ƃ���
''''                 If wLENCB = 0 Then
'                 If wMAI = 0 Then    'upd 2003/06/05 hitec)matsumoto
                If (wMAI = 0 And Kihon.NOWPROC <> PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Or _
                        (wLen = 0 And Kihon.NOWPROC = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Then    '05/03/29 ooba
                     SXL(0).LIVKCB = "1"                 ' �����敪
                     SXL(0).KANKCB = "2"                 ' �����敪
                     SXL(0).LSTCCB = "H"                 ' �ŏI��ԋ敪
                     SXL(0).LDFRBCB = "2"                ' �i���敪
                 Else
                     SXL(0).LIVKCB = "0"
                     SXL(0).KANKCB = "0"                 ' �����敪
                     SXL(0).LSTCCB = "T"                 ' �ŏI��ԋ敪
                     SXL(0).LDFRBCB = "0"                ' �i���敪
                 End If
                 ' 2002/12/13 ooba �����敪�t���O�ύX
                 SXL(0).KANKCB = "0"                 ' �����敪
                 '�V���O���m�莞�A�ŏI��ԋ敪='S'�ɂ���
                 If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
                    SXL(0).LSTCCB = "S"
                 End If
''''                 SXL(0).KANKCB = "0"                     ' �����敪
                 SXL(0).NFCB = "0"                       ' ���ɋ敪
                 SXL(0).SAKJCB = "0"                     ' �폜�敪
                                                         ' �o�^���t
                 SXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                                                      ' �X�V���t
                 SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                 SXL(0).SUMITCB = "0"                    ' SUMIT���M�t���O
                 SXL(0).SNDKCB = "0"                     ' �ԕi�敪
                 SXL(0).SNDAYCB = ""                    ' ���M���t
            
                 iRtn = CreateXSDCB(SXL(0), wErrMsg)
                 '���������i�r�w�k�j�ǉ��G���[
                 If iRtn = FUNCTION_RETURN_FAILURE Then
                     MsgBox wErrMsg
                     Exit Function
                 End If
             End If
        End If
    Next i

''''' 02/09/20 Add                end
    
    XSDCBProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDCBProc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

'�T�v      :�H�����т���w�肳�ꂽ�H���A�r�w�k�h�c�i�u���b�N�h�c�A�ʒu�A�i�ԁj�̒����A�����A�s�ǖ������W�v����
'          :��M�p�̃e�[�u������w�肳�ꂽ�r�w�k�h�c�j�̃T���v�������A�T���v�����w�������A�T���v�����w���s�ǖ���
'           ���W�v����
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'           pKKTC          I   string            �H��
'           pSXLID         I   string            �r�w�k�h�c
'           pLEN           O   NUMBER            ����
'           pMAI           O   NUMBER            ����
'           pMAI800        O   NUMBER            CW800����
'           pFUR           O   NUMBER            �s�ǖ���
'           pFURKEI        O   NUMBER            �s�ǖ������v
'           pSAM           O   NUMBER            �T���v������
'           pSAMNUK        O   NUMBER            �T���v�����w������
'           pSAMFUR        O   NUMBER            �T���v�����w���s�ǖ���
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :

Public Function XSDCBSum(ByVal pKKTC, ByVal pSXLID, ByRef pLEN, ByRef pMAI, ByRef pMAI800, ByRef pFUR, ByRef pFURKEI, ByRef pSAM, ByRef pSAMSIJ, ByRef pSAMFUR)
    
'   �����ϐ�
    Dim i               As Integer
    Dim iRtn            As Integer          '���A���
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim wCRYNUMCA       As String           '�u���b�N�h�c
    Dim wINPOSCA        As Long             '�J�n�ʒu
    Dim wHINBCA         As String           '�i��
    Dim wLen            As Long             '����
    Dim wMAI            As Long             '����
    Dim wMAI800         As Long             'CW800��ʉ߂�������
    Dim wFUR            As Long             '�s�ǖ���
    Dim wFURKEI         As Long             '�s�ǖ������v
    Dim wKCNTC3         As String           '�H���A�ԍő�
    Dim wSAMFUR         As String           '�T���v�������w���s�ǖ���
    Dim rsXsdca         As OraDynaset
    Dim rsMain          As OraDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
        
    '�p�����[�^�m������
    pLEN = 0
    pMAI = 0
    pMAI800 = 0
    pFUR = 0
    pFURKE = 0
    pSAM = 0
    pSAMSIJ = 0
    pSAMFUR = 0

    '���������i�i�ԁj����p�����[�^�̂r�w�k�h�c�̒����A�������擾
    sql = "SELECT SUM(GNLCA) AS wLEN, SUM(GNMCA) AS wMAI "
    sql = sql & " FROM XSDCA "
    sql = sql & " WHERE SXLIDCA = '" & pSXLID & "' "
    sql = sql & " AND LIVKCA = '0' "

    Set rsXsdca = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '���݂��Ȃ����A����
    If rsXsdca.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rsXsdca.Close
        GoTo CW800_CAL
    End If
    
    '���o���ʂ��i�[����
    If IsNull(rsXsdca.Fields("wLEN")) = True Then
        pLEN = 0
    Else
        pLEN = rsXsdca.Fields("wLEN")                '����
    End If
    If IsNull(rsXsdca.Fields("wMAI")) = True Then
        pMAI = 0
    Else
        pMAI = rsXsdca.Fields("wMAI")                '����
    End If
    
    rsXsdca.Close

CW800_CAL:
    
    '���������i�i�ԁj���瓯���r�w�k�h�c�̃u���b�N�h�c�A�J�n�ʒu�A�i�Ԃ��擾
    sql = "SELECT CRYNUMCA, INPOSCA, HINBCA "
    sql = sql & " FROM XSDCA "
    sql = sql & " WHERE SXLIDCA = '" & pSXLID & "' "
    sql = sql & " AND LIVKCA = '0' "
    sql = sql & " ORDER BY CRYNUMCA,INPOSCA"

    Set rsMain = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '���݂��Ȃ����A����
    If rsMain.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rsMain.Close
        GoTo SAMPLE_CAL
    End If

    Do Until rsMain.EOF
        '���o���ʂ��i�[����
        wCRYNUMCA = rsMain.Fields("CRYNUMCA")
        wINPOSCA = rsMain.Fields("INPOSCA")
        wHINBCA = rsMain.Fields("HINBCA")

        '�擾�����u���b�N�h�c��J�n�ʒu��i�ԂōH�����т̊Y���H���ŁA�H���A�Ԃ̍ő���擾����
''''        sql = "SELECT MAX(KCNTC3) AS wKCNTC3 "
''''        sql = sql & " FROM XSDC3 "
''''        sql = sql & " WHERE CRYNUMC3 = '" & wCRYNUMCA & "' "
''''        sql = sql & " AND INPOSC3 = " & wINPOSCA & ""
''''        sql = sql & " AND HINBC3 = '" & wHINBCA & "' "
''''        sql = sql & " AND WKKTC3 = '" & pKKTC & "' "
''''''''        sql = sql & " AND LIVKC3 = '0' "
''''        sql = sql & " AND ((SUMKBC3 = '0') "
''''        sql = sql & "  OR (SUMKBC3 = ' ') "
''''        sql = sql & "  OR (SUMKBC3 is null)) "
''''''''        sql = sql & " AND KKCNTC3  = '0' "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '���݂��Ȃ����A���s
''''        If rs.RecordCount = 0 Then
''''            XSDCBSum = FUNCTION_RETURN_FAILURE
''''            rs.Close
''''            GoTo SAMPLE_CAL
''''        End If
''''
''''        '���o���ʂ��i�[����
''''        If IsNull(rs.Fields("wKCNTC3")) = True Then
''''            XSDCBSum = FUNCTION_RETURN_FAILURE
''''            wKCNTC3 = 0
''''            rs.Close
''''            GoTo SAMPLE_CAL
''''        Else
''''            wKCNTC3 = rs.Fields("wKCNTC3")
''''        End If
        
        '�擾�����u���b�N�h�c��J�n�ʒu��i�ԁA�H���A�ԂōH�����т��璷���A�����A�s�ǖ������擾����
''''        sql = "SELECT SUM(LENC3) AS wLEN, SUM(TOMC3) AS wMAI,SUM(FUMC3) AS wFUR "
        sql = "SELECT TOMC3 AS wMAI800,FUMC3 AS wFUR "
        sql = sql & " FROM XSDC3 "
        sql = sql & " WHERE CRYNUMC3 = '" & wCRYNUMCA & "' "
        sql = sql & " AND INPOSC3 = " & wINPOSCA & ""
''''        sql = sql & " AND HINBC3 = '" & wHINBCA & "' "
''''        sql = sql & " AND WKKTC3 = '" & pKKTC & "' "
''''        sql = sql & " AND LIVKC3 = '0' "
''''        sql = sql & " AND (SUMKBC3 = '0' "
''''        sql = sql & "  OR SUMKBC3 = ' ' "
''''        sql = sql & "  OR SUMKBC3 is null) "
        sql = sql & " AND KCNTC3  = (SELECT MAX(KCNTC3)"
        sql = sql & "                  FROM XSDC3"
        sql = sql & "                 WHERE CRYNUMC3 = '" & wCRYNUMCA & "' "
        sql = sql & "                   AND HINBC3 = '" & wHINBCA & "'"
        sql = sql & "                   AND INPOSC3 = '" & wINPOSCA & "'"
        sql = sql & "                   AND WKKTC3 = '" & pKKTC & "' "
        sql = sql & "                   AND (SUMKBC3 = '0' "
        sql = sql & "                    OR SUMKBC3 = ' ' "
        sql = sql & "                    OR SUMKBC3 is null)) "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '���݂��Ȃ����A���s
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
            GoTo SAMPLE_CAL
        End If
        
        '���o���ʂ��i�[����
'''''        If IsNull(rs.Fields("wLEN")) = True Then
'''''            pLEN = pLEN + 0
'''''        Else
'''''            pLEN = pLEN + rs.Fields("wLEN")         '����
'''''        End If
        If IsNull(rs.Fields("wMAI800")) = True Then
            pMAI800 = pMAI800 + 0
        Else
            pMAI800 = pMAI800 + CInt(rs.Fields("wMAI800"))     '����
        End If
        If IsNull(rs.Fields("wFUR")) = True Then
            pFUR = pFUR + 0
        Else
            pFUR = pFUR + CInt(rs.Fields("wFUR"))              '�s�ǒ���
        End If
        
        '�擾�����u���b�N�h�c��J�n�ʒu��i�ԂōH�����т̕s�Ǎ��v���擾����
        sql = "SELECT SUM(FUMC3) AS wFURKEI "
        sql = sql & " FROM XSDC3 "
        sql = sql & " WHERE CRYNUMC3 = '" & wCRYNUMCA & "' "
        sql = sql & " AND INPOSC3 = " & wINPOSCA & ""
        sql = sql & " AND HINBC3 = '" & wHINBCA & "' "
        sql = sql & " AND SUMKBC3 = '1' "
        sql = sql & " AND MODKBC3 = '0' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '���݂��Ȃ����A���s
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
            GoTo SAMPLE_CAL
        End If
        
        '���o���ʂ��i�[����
        If IsNull(rs.Fields("wFURKEI")) = True Then
            pFURKEI = pFURKEI + 0 '�s�ǒ���
        Else
            pFURKEI = pFURKEI + CInt(rs.Fields("wFURKEI")) '�s�ǒ���    'upd 2003/05/20
        End If
        rsMain.MoveNext

    Loop

    rs.Close
    rsMain.Close

SAMPLE_CAL:
    '�]�����ʎ�M���R�[�h�����T���v���������擾����
    '���T���v�������@-�@�]�����ʎ�M���R�[�h���iY013)
'    sql = "SELECT COUNT(SAMPLEID) AS wSAM "
'    sql = sql & " FROM TBCMY013 "
'    sql = sql & " WHERE  SAMPLEID in ( "
'    sql = sql & " SELECT E044.SMPLID "
'    sql = sql & " FROM TBCME044 E044 "
'    sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH "
'    sql = sql & "  FROM TBCME042 "
'    sql = sql & " WHERE SXLID = '" & pSXLID & "') E042 "
'    sql = sql & " WHERE (E044.CRYNUM = E042.CRYNUM "
'    sql = sql & " AND  E044.INGOTPOS = E042.INGOTPOS "
'    sql = sql & " AND SMPKBN = 'T' ) "
'    sql = sql & " OR (E044.CRYNUM = E042.CRYNUM"
'    sql = sql & " AND E044.INGOTPOS = E042.INGOTPOS + E042.LENGTH "
'    sql = sql & " AND SMPKBN = 'B' ))"

    sql = "SELECT COUNT(SAMPLEID) AS wSAM "
    sql = sql & " FROM TBCMY013 Y013"
    sql = sql & " WHERE  SAMPLEID in ( "
    sql = sql & " SELECT E044.REPSMPLIDCW "
    sql = sql & " FROM XSDCW E044 "
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
    sql = sql & "  ,("
    sql = sql & "    SELECT"
    sql = sql & "      XTALCB as CRYNUM"
    sql = sql & "     ,INPOSCB as INGOTPOS"
    sql = sql & "     ,RLENCB as LENGTH"
    sql = sql & "    FROM"
    sql = sql & "      XSDCB"
    sql = sql & "    WHERE SXLIDCB = '" & pSXLID & "'"
    sql = sql & "   ) E042"
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
'    sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH "
'    sql = sql & "  FROM TBCME042 "
'    sql = sql & " WHERE SXLID = '" & pSXLID & "') E042 "
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
    sql = sql & " WHERE (E044.XTALCW = E042.CRYNUM "
    sql = sql & " AND  E044.INPOSCW = E042.INGOTPOS "
    sql = sql & " AND E044.SMPKBNCW = 'T' ) "
    sql = sql & " OR (E044.XTALCW = E042.CRYNUM"
    sql = sql & " AND E044.INPOSCW = E042.INGOTPOS + E042.LENGTH "
    sql = sql & " AND E044.SMPKBNCW = 'B' ))"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '���݂��Ȃ����A���s
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
'        GoTo PROC_EXIT
    End If

    '���o���ʂ��i�[����
    pSAM = rs.Fields("wSAM") '�T���v������

    rs.Close

    '�����w�����R�[�h�����T���v���w���������擾����
    '���T���v�����w�������i�Ǖi�j�@-�@�����w�����R�[�h���iY003)
'    sql = "SELECT COUNT(SAMPLEID) AS wSIJ"
'    sql = sql & " FROM TBCMY003 "
'    sql = sql & " WHERE SAMPLEID in ( "
'    sql = sql & " SELECT E044.SMPLID "
'    sql = sql & " FROM TBCME044 E044 "
'    sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH "
'    sql = sql & " FROM TBCME042 "
'    sql = sql & " WHERE SXLID = '" & pSXLID & "') E042 "
'    sql = sql & " WHERE (E044.CRYNUM = E042.CRYNUM "
'    sql = sql & " AND E044.INGOTPOS = E042.INGOTPOS "
'    sql = sql & " AND SMPKBN = 'T' ) "
'    sql = sql & " OR (E044.CRYNUM = E042.CRYNUM "
'    sql = sql & " AND E044.INGOTPOS = E042.INGOTPOS + E042.LENGTH "
'    sql = sql & " AND  SMPKBN = 'B' )) "

    sql = "SELECT COUNT(SAMPLEID) AS wSIJ"
    sql = sql & " FROM TBCMY003 "
    sql = sql & " WHERE SAMPLEID in ( "
    sql = sql & " SELECT E044.REPSMPLIDCW "
    sql = sql & " FROM XSDCW E044 "
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
    sql = sql & "  ,("
    sql = sql & "    SELECT"
    sql = sql & "      XTALCB as CRYNUM"
    sql = sql & "     ,INPOSCB as INGOTPOS"
    sql = sql & "     ,RLENCB as LENGTH"
    sql = sql & "    FROM"
    sql = sql & "      XSDCB"
    sql = sql & "    WHERE SXLIDCB = '" & pSXLID & "'"
    sql = sql & "   ) E042"
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
'    sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH "
'    sql = sql & " FROM TBCME042 "
'    sql = sql & " WHERE SXLID = '" & pSXLID & "') E042 "
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
    sql = sql & " WHERE (E044.XTALCW = E042.CRYNUM "
    sql = sql & " AND E044.INPOSCW = E042.INGOTPOS "
    sql = sql & " AND E044.SMPKBNCW = 'T' ) "
    sql = sql & " OR (E044.XTALCW = E042.CRYNUM "
    sql = sql & " AND E044.INPOSCW = E042.INGOTPOS + E042.LENGTH "
    sql = sql & " AND  E044.SMPKBNCW = 'B' )) "

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '���݂��Ȃ����A���s
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
'        GoTo PROC_EXIT
    End If

    '���o���ʂ��i�[����
    pSIJ = rs.Fields("wSIJ") '�T���v�����w������

    rs.Close

    'C�����������T���v�����w���s�ǖ������擾����
    '���T���v�������w���s�ǖ����@-�@C���������@-�iY012�j

    '�Ώۂ̃u���b�NID�擾
    sql = "SELECT DISTINCT(CRYNUMCA) "
    sql = sql & " FROM XSDCA"
    sql = sql & " WHERE SXLIDCA = '" & pSXLID & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '���݂��Ȃ����A���s
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
'        GoTo PROC_EXIT
    End If

    Do Until rs.EOF
        '�������COUNT(�u���b�NID���ƃ��[�v��SUM����j

        '���o���ʂ��i�[����
        wCRYNUMCA = rs.Fields("CRYNUMCA") '�u���b�NID

        sql = "SELECT COUNT(Y012.LOTID) AS wSAMFUR "
        sql = sql & " FROM TBCMY012 Y012 "
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
        sql = sql & "  ,("
        sql = sql & "    SELECT"
        sql = sql & "      XTALCB as CRYNUM"
        sql = sql & "     ,INPOSCB as INGOTPOS"
        sql = sql & "     ,RLENCB as LENGTH"
        sql = sql & "    FROM"
        sql = sql & "      XSDCB"
        sql = sql & "    WHERE SXLIDCB = '" & pSXLID & "'"
    sql = sql & "   ) E042"
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
'        sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH "
'        sql = sql & " FROM TBCME042 "
'        sql = sql & " WHERE SXLID = '" & pSXLID & "' ) E042 "
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
        sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH, BLOCKID "
        sql = sql & " FROM TBCME040 "
        sql = sql & " WHERE BLOCKID =  '" & wCRYNUMCA & "' ) E040 "
        sql = sql & " WHERE Y012.LOTID = E040.BLOCKID "
        sql = sql & " AND E042.INGOTPOS <= Y012.TOP_POS / 10 + E040.INGOTPOS "
        sql = sql & " AND E042.INGOTPOS + E042.LENGTH  >= Y012.TOP_POS / 10 + E040.INGOTPOS "
        sql = sql & " AND REJCAT = 'C' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '���݂��Ȃ����A���s
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
'            GoTo PROC_EXIT
        End If

        '���o���ʂ��i�[����
        pSAMFUR = rs.Fields("wSAMFUR") '�T���v�������w���s�ǖ���

        rs.MoveNext

    Loop

    rs.Close

    XSDCBSum = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    XSDCBSum = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function









'###  �����擾�֐�  ###########################

'�T�v      :WF�������v�Z����
'���Ұ�    :�ϐ���          ,IO ,�^        ,����
'          :SelectBlkID     ,I  ,Integer   ,�u���b�NID
'          :intBlkLen       ,I  ,Integer   ,�u���b�N����
'          :intWfCnt        ,O  ,Integer   ,����
'          :�߂�l          ,O  ,Integer   ,WF����
'����      :
'����      :2002/09/12 ADD hitec)N.MATSUMOTO
Public Function WfCount(ByVal SelectBlkID As String, ByVal intBlkLen As Integer, ByRef intWfCnt As Integer) As FUNCTION_RETURN


Dim rec() As typ_cmkc001f_Disp
Dim ret As FUNCTION_RETURN
Dim recCnt As Long
Dim i As Long
Dim j As Integer
Dim s As String
Dim intWfNum    As Integer '����

    '###�@�����v�Z�֐��p�p�����[�^�iHSXCTCEN & HSXCYCEN�j�擾 ###
    
    ''�d�l�E���т�ǂݍ���
    ret = DBDRV_fcmkc001f_Disp(Trim(SelectBlkID), blkInfo, rec)   'SelectBlkId=�u���b�NID,blkInfo=�u���b�N�Ǘ��\����
    If ret = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    ElseIf UBound(rec) Then
        HSXCTCEN = rec(1).HSXCTCEN
        HSXCYCEN = rec(1).HSXCYCEN
    End If
    
    '########################################################
    
    
    '###  WF�����v�Z�p�̊�{�l���擾����  ###
    Loss0 = val(GetCodeField("LG", "01", "LOSS0", "INFO1"))
    Loss4 = val(GetCodeField("LG", "01", "LOSS4", "INFO1"))
    Mlt4 = val(GetCodeField("LG", "01", "MLT4", "INFO1"))
    Pitch = val(GetCodeField("LG", "01", "PITCH", "INFO1"))
    '######################################
    
    
    '###�@�����v�Z�֐��p�p�����[�^�iSEEDDEG�j�擾 ###
    
'��8���w���P�����������Ȃ� 2007/10/10 SETsw kubota
'    If Left(Trim(SelectBlkID), 1) = "8" Then
'        '�w���P�����̏ꍇ
'        If DBDRV_getSEEDDEG(Trim(SelectBlkID), SEEDDEG) = FUNCTION_RETURN_FAILURE Then
'            GoTo proc_exit
'        End If
'    Else
        '�����グ�����̏ꍇ
        s = GetCodeField("SC", "28", Left$(blkInfo.SEED, 1), "INFO3")
        If Left$(s, 1) = "4" Then
            SEEDDEG = 4
        Else
            SEEDDEG = 0
        End If
'    End If
    
    '#############################################


    '###  �����擾�֐�  ###########################
    intWfCnt = GetWfCount(val(intBlkLen), SEEDDEG, HSXCTCEN, HSXCYCEN)  'intWfCount=����
    WfCount = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

End Function
'2002/09/12 ADD hitec)N.MATSUMOTO

'2002/09/12 ADD hitec)N.MATSUMOTO End

'###  �����擾�֐�  ###########################

'�T�v      :WF�������v�Z����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :blkLen        ,I  ,Integer   ,�u���b�N����
'          :seedDeg       ,I  ,Integer   ,������SEED�X��
'          :hinDegT       ,I  ,Double    ,�i�ԌX���i�c�j
'          :hinDegY       ,I  ,Double    ,�i�ԌX���i���j
'          :�߂�l        ,O  ,Integer   ,WF����
'����      :
'����      :2001/8/30 �쐬  �쑺
Public Function GetWfCount(ByVal BlkLen%, ByVal SEEDDEG%, ByVal hinDegT As Double, ByVal hinDegY As Double) As Integer
Dim hinDeg As Integer
Dim s As String
Dim WfCnt As Integer

    If Pitch = 0# Then
        GetWfCount = 0
        Exit Function
    End If

    ''�i�ԌX���𓾂�
    '�����ŏI�����o���A�i�ԌX���̋��ߕ��ύX
    If (Abs(hinDegT) = 2.83) And (Abs(hinDegY) = 2.83) Then
        hinDeg = 4
    ElseIf (Abs(hinDegT) = 4) And (hinDegY = 0) Then
        hinDeg = 4
    ElseIf (hinDegT = 0) And (Abs(hinDegY) = 4) Then
        hinDeg = 4
    Else
        hinDeg = 0
    End If
    
    ''WF�������v�Z����
    If SEEDDEG = hinDeg Then
        '�ʏ�i�̏ꍇ
        WfCnt = Format(((BlkLen - Loss0) / Pitch) + 0.4, "0")
    Else
        WfCnt = Format(((BlkLen * Mlt4 - Loss4) / Pitch) + 0.4, "0")
    End If
''''    If WfCnt < 0 Then WfCnt = 0
    GetWfCount = WfCnt
End Function
'##########################################################


'2002/09/12 ADD hitec)N.MATSUMOTO Start
'###�@�����v�Z�֐��p�p�����[�^�iHSXCTCEN & HSXCYCEN�j�擾 ###

'�T�v      :�����ŏI���o���� �\���p�c�a�h���C�o
'���Ұ��@�@:�ϐ���       ,IO ,�^                   ,����
'      �@�@:BlockID_in�@ ,I  ,String               ,�u���b�NID
'      �@�@:blkInfo�@�@�@,O  ,typ_cmkc001f_Block   ,�u���b�N���
'      �@�@:records�@�@�@,O  ,typ_cmkc001f_Disp    ,���i�d�l�擾�p
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN      ,�ǂݍ��݂̐���
Public Function DBDRV_fcmkc001f_Disp(BlockID_in As String, blkInfo As typ_cmkc001f_Block, records() As typ_cmkc001f_Disp) As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Integer
    Dim i As Long
    Dim n As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_SUCCESS
    
    ''�u���b�N���𓾂�
    sql = "Select BLK.INGOTPOS, BLK.LENGTH, BLK.REALLEN, BLK.KRPROCCD, BLK.NOWPROC, BLK.LPKRPROCCD, " & _
          "BLK.LASTPASS, BLK.DELCLS, BLK.RSTATCLS, BLK.LSTATCLS, CRY.SEED " & _
          "From TBCME040 BLK, TBCME037 CRY " & _
          "Where (BLOCKID='" & BlockID_in & "') and (BLK.CRYNUM=CRY.CRYNUM)"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If
    With blkInfo
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
    
    
    
    ''���i�d�l�𓾂�
    sql = "select "
    sql = sql & "BH.E041HINBAN, "           ' �i��
    sql = sql & "BH.E041INGOTPOS, "         ' �������J�n�ʒu
    sql = sql & "BH.E041REVNUM, "           ' ���i�ԍ������ԍ�
    sql = sql & "BH.E041FACTORY, "          ' �H��
    sql = sql & "BH.E041OPECOND, "          ' ���Ə���
    sql = sql & "BH.E041LENGTH, "           ' ����
    '���i�d�lSXL�f�[�^
    sql = sql & "S.E018HSXD1CEN, "          ' �i�r�w���a�P���S
    sql = sql & "S.E018HSXRMIN, "           ' �i�r�w���R����
    sql = sql & "S.E018HSXRMAX, "           ' �i�r�w���R���
    sql = sql & "S.E018HSXRMBNP, "          ' �i�r�w���R�ʓ����z
    sql = sql & "S.E018HSXRHWYS, "          ' �i�r�w���R�ۏؕ��@�Q��
    sql = sql & "S.E019HSXONMIN, "          ' �i�r�w�_�f�Z�x����
    sql = sql & "S.E019HSXONMAX, "          ' �i�r�w�_�f�Z�x���
    sql = sql & "S.E019HSXONMBP, "          ' �i�r�w�_�f�Z�x�ʓ����z
    sql = sql & "S.E019HSXONHWS, "          ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    sql = sql & "S.E019HSXCNMIN, "          ' �i�r�w�Y�f�Z�x����
    sql = sql & "S.E019HSXCNMAX, "          ' �i�r�w�Y�f�Z�x���
    sql = sql & "S.E019HSXCNHWS, "          ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    sql = sql & "S.E019HSXTMMAXN, "          ' �i�r�w�]�ʖ��x���             ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sql = sql & "S.E020HSXBM1AN, "          ' �i�r�w�a�l�c�P���ω���
    sql = sql & "S.E020HSXBM1AX, "          ' �i�r�w�a�l�c�P���Ϗ��
    sql = sql & "S.E020HSXBM1HS, "          ' �i�r�w�a�l�c�P�ۏؕ��@�Q��
    sql = sql & "S.E020HSXBM2AN, "          ' �i�r�w�a�l�c�Q���ω���
    sql = sql & "S.E020HSXBM2AX, "          ' �i�r�w�a�l�c�Q���Ϗ��
    sql = sql & "S.E020HSXBM2HS, "          ' �i�r�w�a�l�c�Q�ۏؕ��@�Q��
    sql = sql & "S.E020HSXBM3AN, "          ' �i�r�w�a�l�c�R���ω���
    sql = sql & "S.E020HSXBM3AX, "          ' �i�r�w�a�l�c�R���Ϗ��
    sql = sql & "S.E020HSXBM3HS, "          ' �i�r�w�a�l�c�R�ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF1AX, "          ' �i�r�w�n�r�e�P���Ϗ��
    sql = sql & "S.E020HSXOF1MX, "          ' �i�r�w�n�r�e�P���
    sql = sql & "S.E020HSXOF1HS, "          ' �i�r�w�n�r�e�P �ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF2AX, "          ' �i�r�w�n�r�e�Q���Ϗ��
    sql = sql & "S.E020HSXOF2MX, "          ' �i�r�w�n�r�e�Q���
    sql = sql & "S.E020HSXOF2HS, "          ' �i�r�w�n�r�e�Q �ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF3AX, "          ' �i�r�w�n�r�e�R���Ϗ��
    sql = sql & "S.E020HSXOF3MX, "          ' �i�r�w�n�r�e�R���
    sql = sql & "S.E020HSXOF3HS, "          ' �i�r�w�n�r�e�R �ۏؕ��@�Q��
    sql = sql & "S.E020HSXOF4AX, "          ' �i�r�w�n�r�e�S���Ϗ��
    sql = sql & "S.E020HSXOF4MX, "          ' �i�r�w�n�r�e�S���
    sql = sql & "S.E020HSXOF4HS, "          ' �i�r�w�n�r�e�S �ۏؕ��@�Q��
    sql = sql & "S.E020HSXDENMX, "          ' �i�r�w�c�������
    sql = sql & "S.E020HSXDENMN, "          ' �i�r�w�c��������
    sql = sql & "S.E020HSXDENHS, "          ' �i�r�w�c�����ۏؕ��@�Q��
    sql = sql & "S.E020HSXDVDMXN, "          ' �i�r�w�c�u�c�Q���           ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sql = sql & "S.E020HSXDVDMNN, "          ' �i�r�w�c�u�c�Q����           ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    sql = sql & "S.E020HSXDVDHS, "          ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    sql = sql & "S.E020HSXLDLMX, "          ' �i�r�w�k�^�c�k���
    sql = sql & "S.E020HSXLDLMN, "          ' �i�r�w�k�^�c�k����
    sql = sql & "S.E020HSXLDLHS, "          ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    sql = sql & "S.E019HSXLTMIN, "          ' �i�r�w�k�^�C������
    sql = sql & "S.E019HSXLTMAX, "          ' �i�r�w�k�^�C�����
    sql = sql & "S.E019HSXLTHWS, "          ' �i�r�w�k�^�C���ۏؕ��@�Q��
    sql = sql & "S.E018HSXDPDIR, "          ' �i�r�w�a�ʒu����
    sql = sql & "S.E018HSXDPDRC, "          ' �i�r�w�a�ʒu����
    sql = sql & "S.E018HSXDWMIN, "          ' �i�r�w�a�Љ���
    sql = sql & "S.E018HSXDWMAX, "          ' �i�r�w�a�Џ��
    sql = sql & "S.E018HSXDDMIN, "          ' �i�r�w�a�[����
    sql = sql & "S.E018HSXDDMAX, "          ' �i�r�w�a�[���
    sql = sql & "S.E018HSXD1MIN, "          ' �i�r�w���a�P����
    sql = sql & "S.E018HSXD1MAX, "          ' �i�r�w���a�P���
    sql = sql & "S.E018HSXCTCEN, "          ' �i�r�w�����ʌX�c���S
    sql = sql & "S.E018HSXCYCEN, "          ' �i�r�w�����ʌX�����S
    sql = sql & "U.EPDUP "                  ' ���������Ǘ� EPD�@���
    sql = sql & " from VECME009 BH, VECME001 S, TBCME036 U "
    sql = sql & " where BH.E040BLOCKID='" & BlockID_in & "' "
    sql = sql & " and S.E018HINBAN=BH.E041HINBAN "
    sql = sql & " and S.E018MNOREVNO=BH.E041REVNUM "
    sql = sql & " and S.E018FACTORY=BH.E041FACTORY "
    sql = sql & " and S.E018OPECOND=BH.E041OPECOND "
    sql = sql & " and U.HINBAN=BH.E041HINBAN "
    sql = sql & " and U.MNOREVNO=BH.E041REVNUM "
    sql = sql & " and U.FACTORY=BH.E041FACTORY "
    sql = sql & " and U.OPECOND=BH.E041OPECOND "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        ReDim records(0)
        rs.Close
        GoTo proc_exit
    End If
    
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            '�i�ԊǗ�
            .hinban = rs("E041HINBAN")              ' �i��
            .INGOTPOS = rs("E041INGOTPOS")          ' �������J�n�ʒu
            .REVNUM = rs("E041REVNUM")              ' ���i�ԍ������ԍ�
            .factory = rs("E041FACTORY")            ' �H��
            .opecond = rs("E041OPECOND")            ' ���Ə���
            .LENGTH = rs("E041LENGTH")              ' ����
            '���i�d�lSXL�f�[�^
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
            .HSXTMMAX = rs("E019HSXTMMAXN")           ' �i�r�w�]�ʖ��x���           ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
            For n = 1 To 3
'''                .HSXBMnAN(n) = rs("E020HSXBM" & n & "AN") * 10 ' �i�r�w�a�l�cn ���ω���
'''                .HSXBMnAX(n) = rs("E020HSXBM" & n & "AX") * 10 ' �i�r�w�a�l�cn ���Ϗ��
                .HSXBMnHS(n) = rs("E020HSXBM" & n & "HS")  ' �i�r�w�a�l�cn �ۏؕ��@�Q��
            Next
            For n = 1 To 4
'''                .HSXOFnAX(n) = rs("E020HSXOF" & n & "AX")   ' �i�r�w�n�r�en ���Ϗ��
'''                .HSXOFnMX(n) = rs("E020HSXOF" & n & "MX")   ' �i�r�w�n�r�en ���
                If IsNull(rs("E020HSXOF" & n & "AX")) = False Then .HSXOFnAX(n) = rs("E020HSXOF" & n & "AX")   ' �i�r�w�n�r�en ���Ϗ��         '05/03/29 ooba NULL�Ή�
                If IsNull(rs("E020HSXOF" & n & "MX")) = False Then .HSXOFnMX(n) = rs("E020HSXOF" & n & "MX")   ' �i�r�w�n�r�en ���             '05/03/29 ooba NULL�Ή�
                .HSXOFnHS(n) = rs("E020HSXOF" & n & "HS")   ' �i�r�w�n�r�en �ۏؕ��@�Q��
            Next
            .HSXDENMX = rs("E020HSXDENMX")          ' �i�r�w�c�������
            .HSXDENMN = rs("E020HSXDENMN")          ' �i�r�w�c��������
            .HSXDENHS = rs("E020HSXDENHS")          ' �i�r�w�c�����ۏؕ��@�Q��
            .HSXDVDMX = rs("E020HSXDVDMXN")          ' �i�r�w�c�u�c�Q���        ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
            .HSXDVDMN = rs("E020HSXDVDMNN")          ' �i�r�w�c�u�c�Q����        ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
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
'''            .HSXCTCEN = rs("E018HSXCTCEN")          ' �i�r�w�����ʌX�c���S
'''            .HSXCYCEN = rs("E018HSXCYCEN")          ' �i�r�w�����ʌX�����S
            If IsNull(rs("E018HSXCTCEN")) = False Then .HSXCTCEN = rs("E018HSXCTCEN")       ' �i�r�w�����ʌX�c���S      '05/03/29 ooba NULL�Ή�
            If IsNull(rs("E018HSXCYCEN")) = False Then .HSXCYCEN = rs("E018HSXCYCEN")       ' �i�r�w�����ʌX�����S      '05/03/29 ooba NULL�Ή�
            .EPDUP = rs("EPDUP")                    ' ���������Ǘ� EPD�@���
        End With
        rs.MoveNext
    Next
    rs.Close


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'2002/09/13 Add hitec)N.MATSUMOTO  Start
'�T�v      :���������i�s�Ǔ���j����A�w�肵�������ɊY������f�[�^�̍s�����擾
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'�@�@�@�@�@�FstrWhere     ,I  ,String           ,SELECT������
'�@�@�@�@�@�FintCnt       ,O  ,Integer          ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :

Public Function SelCntXSDC4(ByVal strWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim sqlWhere        As String           'WHERE��

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    SelCntXSDC4 = FUNCTION_RETURN_FAILURE
    
    sql = "      SELECT count(*) cnt "
    sql = sql & "  FROM XSDC4 " & strWhere

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '���݂��Ȃ����A�G���[
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
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    SelCntXSDC4 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/13 Add hitec)N.MATSUMOTO  End


'2002/09/13 Add hitec)N.MATSUMOTO  Start
'�T�v      :���������i�i�ԁj����A�w�肵�������ɊY������f�[�^�̍s�����擾
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'�@�@�@�@�@�FstrWhere     ,I  ,String           ,SELECT������
'�@�@�@�@�@�FintCnt       ,O  ,Integer          ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :

Public Function SelCntXSDCA(ByVal strWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim sqlWhere        As String           'WHERE��

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    SelCntXSDCA = FUNCTION_RETURN_FAILURE
    
    sql = "      SELECT count(*) cnt "
    sql = sql & "  FROM XSDCA " & strWhere

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '���݂��Ȃ����A�G���[
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
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    SelCntXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/13 Add hitec)N.MATSUMOTO  End

'2002/09/13 Add hitec)N.MATSUMOTO  Start
'�T�v      :���������iSXL�j����A�w�肵�������ɊY������f�[�^�̍s�����擾
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'�@�@�@�@�@�FstrWhere     ,I  ,String           ,SELECT������
'�@�@�@�@�@�FintCnt       ,O  ,Integer          ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :

Public Function SelCntXSDCB(ByVal strWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim sqlWhere        As String           'WHERE��

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    SelCntXSDCB = FUNCTION_RETURN_FAILURE
    
    sql = "      SELECT count(*) cnt "
    sql = sql & "  FROM XSDCB " & strWhere

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '���݂��Ȃ����A�G���[
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
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    SelCntXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/13 Add hitec)N.MATSUMOTO  End


'**********************************************
'�@�VDB�\���̏�����
'�@ADD hitec)N.MATSUMOTO
'**********************************************
Public Sub clearType()

    On Error Resume Next

    With Kihon  '��{���
        .ALLSCRAP = ""
        .CNTHINNOW = 0
        .CNTHINOLD = 0
        .DIAMETER = 0
        .FURYOUMU = ""
        .NEWPROC = ""
        .NOWPROC = ""
        .STAFFID = ""
    End With
    
    With BlkOld      '��������(�u���b�N)�F�O�H��
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


'''''�T�v      :�����v�Z�֐�
'''''���Ұ�    :�ϐ���          ,IO ,�^        ,����
'''''          :strBlockId      ,I  ,Integer   ,�u���b�NID
'''''          :intLen          ,I  ,Integer   ,����
'''''          :�߂�l          ,O  ,Integer   ,WF����
'''''����      :
'''''����      :2002/09/11 ADD hitec)N.MATSUMOTO
''''Public Function GetWfNum(ByVal strBlockID As String, ByVal intLen As Integer, ByRef intWfNum As Integer) As FUNCTION_RETURN
''''
''''    Dim rs      As OraDynaset
''''    Dim sql     As String
''''    Dim intRtn  As Integer
''''
''''    '' �G���[�n���h���̐ݒ�
''''    On Error GoTo proc_err
''''
''''    '���i�d�l�𓾂�
''''    sql = "SELECT S.E018HSXCTCEN,S.E018HSXCYCEN "
''''    sql = sql & " from VECME009 BH, VECME001 S, TBCME036 U "
''''    sql = sql & " where BH.E040BLOCKID='" & strBlockID & "' "
''''    sql = sql & " and S.E018HINBAN=BH.E041HINBAN "
''''    sql = sql & " and S.E018MNOREVNO=BH.E041REVNUM "
''''    sql = sql & " and S.E018FACTORY=BH.E041FACTORY "
''''    sql = sql & " and S.E018OPECOND=BH.E041OPECOND "
''''    sql = sql & " and U.HINBAN=BH.E041HINBAN "
''''    sql = sql & " and U.MNOREVNO=BH.E041REVNUM "
''''    sql = sql & " and U.FACTORY=BH.E041FACTORY "
''''    sql = sql & " and U.OPECOND=BH.E041OPECOND "
''''
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''    If rs.RecordCount = 0 Then
''''        rs.Close
''''        GoTo proc_exit
''''    End If
''''
''''    ' �i�r�w�����ʌX�c���S
''''    If IsNull(rs("E018HSXCTCEN")) = False Then
''''        HSXCTCEN = 0
''''    Else
''''        HSXCTCEN = rs("E018HSXCTCEN")
''''    End If
''''
''''    ' �i�r�w�����ʌX�����S
''''    If IsNull(rs("E018HSXCYCEN")) = False Then
''''        HSXCYCEN = 0
''''    Else
''''        HSXCYCEN = rs("E018HSXCYCEN")
''''    End If
''''
''''    If Left(Trim(strBlockID), 1) = "8" Then
''''        '�w���P�����̏ꍇ
''''        If DBDRV_getSEEDDEG(Trim(strBlockID), SEEDDEG) = FUNCTION_RETURN_FAILURE Then
''''            rs.Close
''''            GoTo proc_exit
''''        End If
''''    Else
''''''        '�u���b�N�Ǘ�TBCME037����擾
''''''
''''''        '�����グ�����̏ꍇ
''''''        s = GetCodeField("SC", "28", Left$(blkInfo.SEED, 1), "INFO3")
''''''        If Left$(s, 1) = "4" Then
''''''            SEEDDEG = 4
''''''        Else
''''''            SEEDDEG = 0
''''''        End If
''''    End If
''''
''''    'WF�������v�Z���A�l��Ԃ�
''''    intWfNum = calculateWfNum(intLen, SEEDDEG, HSXCTCEN, HSXCYCEN)
''''
''''    GetWfNum = FUNCTION_RETURN_SUCCESS
''''
''''proc_exit:
''''    '' �I��
''''    gErr.Pop
''''    Exit Function
''''
''''proc_err:
''''    '' �G���[�n���h��
''''    Debug.Print "====== Error SQL ======"
''''    Debug.Print sql
''''    gErr.HandleError
''''    GetWfNum = FUNCTION_RETURN_FAILURE
''''    Resume proc_exit
''''
''''End Function

'�T�v      :WF�������v�Z����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :blkLen        ,I  ,Integer   ,�u���b�N����
'          :seedDeg       ,I  ,Integer   ,������SEED�X��
'          :hinDegT       ,I  ,Double    ,�i�ԌX���i�c�j
'          :hinDegY       ,I  ,Double    ,�i�ԌX���i���j
'          :�߂�l        ,O  ,Integer   ,WF����
'����      :
'����      :2001/8/30 �쐬  �쑺
Private Function calculateWfNum(ByVal BlkLen%, ByVal SEEDDEG%, ByVal hinDegT As Double, ByVal hinDegY As Double) As Integer
Dim hinDeg As Integer
Dim s As String
Dim WfCnt As Integer

    If Pitch = 0# Then
        calculateWfNum = 0
        Exit Function
    End If

    ''�i�ԌX���𓾂�
    '�����ŏI�����o���A�i�ԌX���̋��ߕ��ύX
    If (Abs(hinDegT) = 2.83) And (Abs(hinDegY) = 2.83) Then
        hinDeg = 4
    ElseIf (Abs(hinDegT) = 4) And (hinDegY = 0) Then
        hinDeg = 4
    ElseIf (hinDegT = 0) And (Abs(hinDegY) = 4) Then
        hinDeg = 4
    Else
        hinDeg = 0
    End If
    
    ''WF�������v�Z����
    If SEEDDEG = hinDeg Then
        '�ʏ�i�̏ꍇ
        WfCnt = Format(((BlkLen - Loss0) / Pitch) + 0.4, "0")
    Else
        WfCnt = Format(((BlkLen * Mlt4 - Loss4) / Pitch) + 0.4, "0")
    End If
    If WfCnt < 0 Then WfCnt = 0
    calculateWfNum = WfCnt
End Function
'2002/09/11 ADD hitec)N.MATSUMOTO End


'�T�v      :�H�����ѓo�^�������s��(�݌Ɍ����FCW740,CW760�p)
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :�H������(XSDC3)�ɍ݌Ɍ����̓o�^�������s��
'����      :2003/04/27  HITEC)��c�F�v�e�����̓}�b�v�ʒu�ł͂Ȃ���ʂ��璼�ڎ擾����
'                                  �i��="Z"�͕s�ǂɂ��Ȃ�

Public Function XSDC3Proc2() As FUNCTION_RETURN

'   �����ϐ�
    Dim i, j, k         As Integer
    Dim iRtn            As Integer          '���A���
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim sqlWhere        As String           'WHERE��
    Dim wErrMsg         As String           '�G���[���b�Z�[�W
    Dim Koutei          As typ_XSDC3_Update '�H������
    Dim rsKCNTC         As OraDynaset       '���R�[�h�Z�b�g
                                                
    Dim wSTOCKINFO()    As typ_stock_info   '���ݍH���̏��
    Dim vGetData        As Variant          '��ʎ捞�pwork
    Dim sOldHinban      As String           '���i��
    Dim sNowHinban      As String           '���i��
    Dim vBlkId          As Variant          '��ʎ捞�pwork
    Dim sOldBlkID       As String           '���u���b�NID
    Dim vREVNUM         As Variant          '��ʎ捞�pwork
    Dim vFACTORY        As Variant          '��ʎ捞�pwork
    Dim vOPE            As Variant          '��ʎ捞�pwork
    Dim iREVNUM         As Integer          '���i�����ԍ�
    Dim sFACTORY        As String           '�H��
    Dim sOPE            As String           '���Ə���
    Dim sBlkId          As String           '�u���b�NID
    
    Dim iMapSt          As Integer          '�}�b�v�J�n�ʒu
    Dim iMapEd          As Integer          '�}�b�v�I���ʒu
    Dim bHinFlg         As Boolean          '�i�Ԕ�r�p�t���O
    Dim lTMaisu         As Long             '���v����
    Dim iGetHinInpos    As Integer          '�������ʒu
    Dim oGamenSpd       As Object           '���ID
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto

    '�G���[�n���h���̐ݒ�
    'On Error GoTo PROC_ERR
    On Error GoTo 0
    
    '�����ݒ�
    XSDC3Proc2 = FUNCTION_RETURN_FAILURE

    ReDim STOCKINFO(0)
    ReDim wSTOCKINFO(0)
        
   '�O�H���������v
    For i = 0 To Kihon.CNTHINOLD - 1
        If (Kihon.NOWPROC = "CW760") _
           And ((SIngotP > CLng(HinOld(i).INPOSCA)) Or (HinOld(i).INPOSCA >= EIngotP)) Then
                '�����Ȃ�
        Else
            ' �O�H��������0�̎��A�����I��
            If HinOld(i).GNMCA <= 0 Then
                XSDC3Proc2 = FUNCTION_RETURN_SUCCESS
                Exit Function
            End If
                
            ReDim Preserve STOCKINFO(UBound(STOCKINFO) + 1)  '�z��̒ǉ�
            '�s�Ǥ�����̏����ݒ�
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
    
        
'�Ŕ����w����ʂ���i�Ԃ̕����o���ƌ������}�b�v�ʒu���ڂ��狁�߂�
'STOCKINFO�z��Ɋi�[���邪STOCKINFO�̕i�Ԃ�HinOld�̕i�Ԃ̓o�^�����ƈ�v���Ă���Ƃ͌���Ȃ�

    If Kihon.NOWPROC = "CW740" Then
        Set oGamenSpd = f_cmbc036_2.sprExamine    '�����ύX���گ��
    Else
        Set oGamenSpd = f_cmbc039_3.sprExamine    '�Ĕ������گ��
    End If
    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX-------start iida 2003/09/06
    With oGamenSpd
'        .GetText 31, 1, vBlkId          '�u���b�NID
        ''�c���_�f�������ڒǉ��ɂ��ύX�@04/01/09 ooba
'        .GetText 32, 1, vBlkId          '�u���b�NID
        'GD�ǉ��ɂ��ύX�@05/01/31 ooba
'        .GetText 33, 1, vBlkId          '�u���b�NID
        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
        .GetText 39, 1, vBlkId          '�u���b�NID
        sOldBlkID = CStr(Trim(vBlkId))
        For i = 1 To .MaxRows Step 2    '���گ�ނ����ް������(2�s���m�F)
''''''            .GetText 28, i, vNukisiFlg
''''''            If (vNukisiFlg = "1") Then
            ' ���i�Ԃ̓o�^
'            .GetText 32, i, vGetData    '�Â��i�Ԏ擾
            ''�c���_�f�������ڒǉ��ɂ��ύX�@04/01/09 ooba
'            .GetText 33, i, vGetData    '�Â��i�Ԏ擾
            'GD�ǉ��ɂ��ύX�@05/01/31 ooba
'            .GetText 34, i, vGetData    '�Â��i�Ԏ擾
        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
            .GetText 40, i, vGetData    '�Â��i�Ԏ擾
            sOldHinban = Trim(CStr(vGetData))
            .GetText 2, i, vGetData     '�V�����i�Ԏ擾
            ' �i�Ԃ�"Z"�̎��͐V�i��=���i��
            If Trim(CStr(vGetData)) = "Z" Then
                sNowHinban = sOldHinban
            Else
                sNowHinban = Trim(CStr(vGetData))
            End If
            .GetText 5, i, vGetData     '�����ʒu
            iGetHinInpos = val(vGetData)
            .GetText 6, i, vGetData     '�}�b�v�J�n�ʒu
            iMapSt = val(vGetData)
            .GetText 6, i + 1, vGetData '�}�b�v�I���ʒu
            iMapEd = val(vGetData)
'            .GetText 31, i, vBlkId      '�u���b�NID
            ''�c���_�f�������ڒǉ��ɂ��ύX�@04/01/09 ooba
'            .GetText 32, i, vBlkId      '�u���b�NID
            'GD�ǉ��ɂ��ύX�@05/01/31 ooba
'            .GetText 33, i, vBlkId      '�u���b�NID
        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
            .GetText 39, i, vBlkId      '�u���b�NID
            If vBlkId = "" Then         '�u���b�N��NULL��������A�O��̃u���b�N���g�p
                vBlkId = Mid(BlkNow.CRYNUMC2, 1, 9) & sOldBlkID
            Else
                vBlkId = Mid(BlkNow.CRYNUMC2, 1, 9) & vBlkId
            End If
' VVVVV 2003/04/27 ALT BY HITEC)��c�F�Ǖi�����̓}�b�v�ʒu�ł͂Ȃ��e�[�u������擾����
            sBlkId = vBlkId
'''''            SXLCnt = iMapEd - iMapSt + 1        '�}�b�v����
' ^^^^^ 2003/04/27 ALT BY HITEC)��c  END
            iREVNUM = gtSprWfMap(i).REVNUM      '���i�����ԍ�
            sFACTORY = gtSprWfMap(i).factory    '�H��
            sOPE = gtSprWfMap(i).opecond        '���Ə���
            
'''            For k = 0 To Kihon.CNTHINNOW - 1    '�i�Ԃ��r���A�Y���ް��̌������J�n�ʒu���擾
'''                If sNowHinban = HinNow(k).HINBCA Then
'''                    iGetHinInpos = HinNow(k).INPOSCA
'''                    Exit For
'''                End If
'''            Next
                              
            If (((Kihon.NOWPROC = "CW760") Or (Kihon.NOWPROC = "CW740")) And (vBlkId <> BlkNow.CRYNUMC2)) Or ((Kihon.NOWPROC = "CW760") And ((SIngotP > iGetHinInpos) Or (iGetHinInpos >= EIngotP))) Then
                '�����Ȃ�
            Else
' VVVVV 2003/04/27 ALT BY HITEC)��c�F�Ǖi�����̓}�b�v�ʒu�ł͂Ȃ��e�[�u������擾����
                sql = "SELECT COUNT(*) AS SXLCNT"
                sql = sql & " FROM TBCMY011 "
                sql = sql & " WHERE LOTID = '" & sBlkId & "'"
                sql = sql & " AND (WFSTA ='0' OR WFSTA = '1') "
                sql = sql & " AND BLOCKSEQ >= " & iMapSt & ""
                sql = sql & " AND BLOCKSEQ <= " & iMapEd & ""
                
                Debug.Print sql
                
                Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
                ''�݂���Ȃ�������G���[
                If rs.RecordCount = 0 Then
                    SXLCnt = 0
                Else ''����������A�Ǖi�������擾����
                    SXLCnt = val(rs("SXLCNT"))
                End If
                Debug.Print SXLCnt
                
' ^^^^^ 2003/04/27 ALT BY HITEC)��c  END
                bHinFlg = False '�����̔z��ɓ����i�Ԃ��o�^����Ă��邩���׸�
                'wSTOCKINFO()��1����J�n
                For j = 1 To UBound(wSTOCKINFO)
                    If (wSTOCKINFO(j).hinban = sOldHinban) Then  '���ɓo�^���Ă���i��
                        bHinFlg = True
' VVVVV 2003/04/27 ALT BY HITEC)��c�F�i��="Z"�͕s�ǂɂ��Ȃ�
'''''                        If (sNowHinban = "Z") Then   'Z�o�^�̎�
'''''                            wSTOCKINFO(j).FURYOM = wSTOCKINFO(j).FURYOM + SXLCnt
'''''                        Else
                        wSTOCKINFO(j).HARAIM = wSTOCKINFO(j).HARAIM + SXLCnt
'''''                        End If
' ^^^^^ 2003/04/27 ALT BY HITEC)��c  END
                    End If
                Next j
    
                If (bHinFlg = False) Then   'wSTOCKINFO()�̔z��ɕi�Ԃ��o�^�Ȃ��������V�K��wSTOCKINFO()�ɓo�^
                    ReDim Preserve wSTOCKINFO(UBound(wSTOCKINFO) + 1)  '�z��̒ǉ�
                    wSTOCKINFO(UBound(wSTOCKINFO)).hinban = sOldHinban  '�i��
                    wSTOCKINFO(UBound(wSTOCKINFO)).HARAIM = 0           '�z�񏉊��ݒ�
                    wSTOCKINFO(UBound(wSTOCKINFO)).FURYOM = 0           '�z�񏉊��ݒ�
' VVVVV 2003/04/27 ALT BY HITEC)��c�F�i��="Z"�͕s�ǂɂ��Ȃ�
'''''                    If (sNowHinban = "Z") Then   'Z�o�^�̎�
'''''                        wSTOCKINFO(UBound(wSTOCKINFO)).FURYOM = SXLCnt  '��ʂ��ް�
'''''                    Else
                    wSTOCKINFO(UBound(wSTOCKINFO)).HARAIM = SXLCnt  '��ʂ��ް�
'''''                    End If
' ^^^^^ 2003/04/27 ALT BY HITEC)��c  END
                    wSTOCKINFO(UBound(wSTOCKINFO)).REVNUM = iREVNUM     '��ʂ��ް�
                    wSTOCKINFO(UBound(wSTOCKINFO)).factory = sFACTORY   '��ʂ��ް�
                    wSTOCKINFO(UBound(wSTOCKINFO)).OPE = sOPE           '��ʂ��ް�
                End If
            End If
        Next i
    End With
    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX-------end iida 2003/09/06
    'STOCKINF��wSTOCKINF�̓ˍ���������STOCKINF�ɕi�ԁA�����A�d�ʁA�����̕����o���ƕs�ǂ��i�[
    'HinOld�ɂȂ��f�[�^�͂Ȃ��Ȃ�
    'STOCKINFO()�͓Y��0����J�n
    'STOCKINFO()�̕i�Ԃ�HinOld���ް��A�ް���HinNow���ް�
    For i = 1 To UBound(STOCKINFO)
        STOCKINFO(i).HARAIM = 0
        STOCKINFO(i).FURYOM = 0
        For j = 1 To UBound(wSTOCKINFO)
            If (STOCKINFO(i).hinban = wSTOCKINFO(j).hinban) Then    '�i�Ԃ����������o�^����
                '''STOCKINFO(i).hinban = HinOld(i).HINBCA
                STOCKINFO(i).HARAIM = wSTOCKINFO(j).HARAIM
                STOCKINFO(i).FURYOM = wSTOCKINFO(j).FURYOM
'''''                STOCKINFO(i).KCKNT = HinOld(i).KCKNTCA  '�A�Ԃ�HinOld����擾
                STOCKINFO(i).REVNUM = wSTOCKINFO(j).REVNUM
                STOCKINFO(i).factory = wSTOCKINFO(j).factory
                STOCKINFO(i).OPE = wSTOCKINFO(j).OPE
                lTMaisu = wSTOCKINFO(j).HARAIM + wSTOCKINFO(j).FURYOM   '�����̍��v
                If (lTMaisu > 0) Then   '���������邩�m�F
                    STOCKINFO(i).HARAIW = wSTOCKINFO(j).HARAIM / lTMaisu * CLng(STOCKINFO(i).HARAIW)        '�s�ǁA���o�͖����̔䗦�ŎZ�o
                    STOCKINFO(i).FuryoW = CLng(STOCKINFO(i).GENZAW) - STOCKINFO(i).HARAIW
                    STOCKINFO(i).HARAIL = wSTOCKINFO(j).HARAIM / lTMaisu * CLng(STOCKINFO(i).HARAIL)       '�s�ǁA���o�͖����̔䗦�ŎZ�o
                    STOCKINFO(i).FURYOL = CLng(STOCKINFO(i).GENZAL) - STOCKINFO(i).HARAIL
                 End If
            End If
        Next j
    Next i
   
    '�s�ǂ�����ꍇ���݌Ɍ����̍쐬
    For i = 1 To UBound(STOCKINFO)
        If STOCKINFO(i).FURYOM > 0 Then
            Koutei.CRYNUMC3 = HinNow(0).CRYNUMCA    '�u���b�N�h�c
            giInpos = giInpos + 1
            Koutei.INPOSC3 = giInpos                '�ʒu
            Koutei.KCNTC3 = STOCKINFO(i).KCKNT + 1  '�H���A��
            Koutei.HINBC3 = STOCKINFO(i).hinban     '�i��
            Koutei.REVNUMC3 = STOCKINFO(i).REVNUM   '���i�����ԍ�
            Koutei.FACTORYC3 = STOCKINFO(i).factory '�H��
            Koutei.OPEC3 = STOCKINFO(i).OPE         '���Ə���
            Koutei.LENC3 = STOCKINFO(i).HARAIL      '����
            Koutei.XTALC3 = HinNow(0).XTALCA        '�����ԍ�
            Koutei.SXLIDC3 = ""                     ' SXLID
            
            Koutei.KNKTC3 = Left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 1) ' �Ǘ��H��(���ݍH��+1)
            Koutei.WKKTC3 = Kihon.NOWPROC           ' �H��
            Koutei.WKKBC3 = ""                      ' ��Ƌ敪
            Koutei.MACOC3 = HinNow(0).NEMACOCA      ' ������
            Koutei.MODKBC3 = ""                     ' �ԍ��敪
            Koutei.SUMKBC3 = ""                     ' �W�v�敪
            Koutei.FRKNKTC3 = ""                    ' (���)�Ǘ��H��
            If IsNull(HinOld(0).NEWKNTCA) = True Then   '(����j�H��
                Koutei.FRWKKTC3 = ""
            Else
                Koutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
                'add end 2003/03/28 hitec)matsumoto --------
            End If
            Koutei.FRWKKBC3 = ""                    ' (���)��Ƌ敪
            If IsNull(HinOld(0).NEMACOCA) = True Then   '�i����j������
                Koutei.FRMACOC3 = "0"
            Else
                Koutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If
            
''''            If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''            If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
            Select Case Kihon.NOWPROC
            Case "CC730"
                iHantei = CInt(BlkNow.GNLC2)
            Case Else
                iHantei = CInt(BlkNow.GNMC2)
            End Select
            If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                Koutei.TOWKKTC3 = " "               ' (���o)�H��
                Koutei.TOMACOC3 = "0"               '(���o)������
            Else
                Koutei.TOWKKTC3 = HinNow(0).GNWKNTCA    ' (���o)�H��
                Koutei.TOMACOC3 = HinNow(0).GNMACOCA    ' (���o)������
            End If
            Koutei.FRLC3 = STOCKINFO(i).GENZAL      ' �������
            Koutei.FRWC3 = STOCKINFO(i).GENZAW      '����d��
            Koutei.FRMC3 = STOCKINFO(i).GENZAM      '�������
            Koutei.FULC3 = STOCKINFO(i).FURYOL      '�s�ǒ���
            Koutei.FUWC3 = STOCKINFO(i).FuryoW      '�s�Ǐd��
            Koutei.FUMC3 = STOCKINFO(i).FURYOM      '�s�ǖ���
            Koutei.LOSWC3 = ""                      ' ���X����
            
            Koutei.LOSLC3 = ""                      ' ���X�d��
            Koutei.LOSMC3 = ""                      ' ���X����
            Koutei.TOLC3 = STOCKINFO(i).HARAIL      '���o����
            Koutei.TOWC3 = STOCKINFO(i).HARAIW      '���o�d��
            Koutei.TOMC3 = STOCKINFO(i).HARAIM      '���o����
            Koutei.SUMITLC3 = ""                    ' SUMIT����
            Koutei.SUMITWC3 = ""                    ' SUMIT�d��
            Koutei.SUMITMC3 = ""                    ' SUMIT����
            Koutei.MOTHINC3 = ""                    ' �U�֕i��(��)
            Koutei.XTWORKC3 = "42"                  ' �����H��
            
            Koutei.WFWORKC3 = ""                    ' ���ʐ���
'           Koutei.STATIMEC3 = Null                 ' �����J�n�I��
'           Koutei.STOTIMEC3 = Null                 ' �������ԏI��
'           Koutei.ETIMEC3 = ""                     ' ���ю��Ԃ͓���Ȃ�
            Koutei.HOLDCC3 = " "                    ' �z�[���h�R�[�h
            Koutei.HOLDBC3 = "0"                    ' �z�[���h�敪
            Koutei.LDFRCC3 = ""                     ' �i���R�[�h
            Koutei.LDFRBC3 = "0"                    ' �i���敪�i�n�C�L�j
            Koutei.TSTAFFC3 = Kihon.STAFFID         ' �o�^�Ј�ID
            Koutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �o�^���t
            
            Koutei.KSTAFFC3 = ""                    ' �X�V�Ј�ID
            Koutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
            Koutei.SUMITBC3 = ""                    ' SUMIT���M�t���O
            Koutei.SNDKC3 = ""                      ' ���M�t���O
'           Koutei.SNDDAYC3 = ""                    ' ���M���t
            Koutei.MODMACOC3 = ""                   ' �ԍ��̏�����
            Koutei.KAKUCC3 = ""                     ' �m��R�[�h
            Koutei.SUMDAYC3 = CalcSumcoTime(Koutei.KDAYC3) 'SUMCO����
            Koutei.PAYCLASSC3 = ""                  '�@�]����H��t���O
'            Koutei.SUMITSNDC3 = ""                  ' SUMIT���M���t
            
'            Koutei.SSENDNOC3 = ""
            
            iRtn = CreateXSDC3(Koutei, wErrMsg)     '�H�����тɍ݌Ɍ����o�^
            If iRtn = FUNCTION_RETURN_FAILURE Then  '�H�����ђǉ��G���[
                MsgBox wErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

''''
'HinInf():���삷��z��
'HinNum:���삷��z��ʒu
'HinFlg:-1�Ȃ�z��폜 1�Ȃ�z��ǉ�

'�T�v      :�H�����ѓo�^�������s��(�i�ԐU�֏��FCW740,CW760�p)
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :�H������(XSDC3)�ɕi�ԐU�֏��̓o�^�������s��

Public Function XSDC3Proc3() As FUNCTION_RETURN

'   �����ϐ�
    Dim i, j            As Integer
    Dim iRtn            As Integer          '���A���
    Dim sql             As String           '�r�p�k
    Dim sqlWhere        As String           'WHERE��
    Dim wErrMsg         As String
    Dim Koutei          As typ_XSDC3_Update    '�H������
    
    Dim wLen            As Long
    Dim wCHKPOS         As Long
        
    Dim wOINF()         As typ_trans_info   '�O�i�ԕ��ёւ��p
    Dim wNINF()         As typ_trans_info   '��i�ԕ��ёւ��p
    Dim wWINF()         As typ_trans_info   '���ёւ��p���[�N  ' 2003/04/17 add by t.t
    Dim ibuf            As Integer
    Dim wOINFrecCnt     As Integer
    Dim wNINFrecCnt     As Integer
    Dim wOINFFLG        As Integer
    Dim wNINFFLG        As Integer
    Dim iCnt             As Integer
    Dim wNINFMAX        As Integer           ' 2003/04/17 add by t.t
    Dim wOINFMAX        As Integer           ' 2003/04/17 add by t.t
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    '�����ݒ�
    XSDC3Proc3 = FUNCTION_RETURN_FAILURE

'' 2003/04/17 add by t.t   start
    ReDim wOINF(UBound(STOCKINFO))      '�i�Ԃ��ƂɃ\�[�g�p
    ReDim wNINF(Kihon.CNTHINNOW)        '�i�Ԃ��ƂɃ\�[�g�p
    ReDim wWINF(1)                      '�i�Ԃ��ƂɃ\�[�g�p
        
    ' �݌Ɍ�������荞��
    For i = 1 To UBound(STOCKINFO)
        If Trim(STOCKINFO(i).hinban) <> "" Then
            wOINF(i).hinban = STOCKINFO(i).hinban
            wOINF(i).LEN = STOCKINFO(i).HARAIL       ' �O�H���������v
            wOINF(i).WAT = STOCKINFO(i).HARAIW       ' �O�H���d�ʍ��v
            wOINF(i).MAI = STOCKINFO(i).HARAIM       ' �O�H���������v
        End If
    Next i
    
    ' �Ǖi������荞��
    For j = 0 To Kihon.CNTHINNOW - 1
        If (Kihon.NOWPROC = "CW760") _
            And ((SIngotP > HinNow(j).INPOSCA) Or (HinNow(j).INPOSCA >= EIngotP)) Then
                '�����Ȃ�
        Else
            wNINF(j).hinban = HinNow(j).HINBCA
            '04/01/13 ooba �ǉ� START ===================>
            wNINF(j).REVNUM = HinNow(j).REVNUMCA
            wNINF(j).factory = HinNow(j).FACTORYCA
            wNINF(j).OPE = HinNow(j).OPECA
            '04/01/13 ooba �ǉ� END =====================>
            wNINF(j).LEN = HinNow(j).GNLCA           ' ��H���������v
            wNINF(j).WAT = HinNow(j).GNWCA           ' ��H���d�ʍ��v
            wNINF(j).MAI = HinNow(j).GNMCA           ' ��H���������v
            wNINF(j).KCKNT = HinNow(j).KCKNTCA       ' ��H���A��
        End If
    Next j
        
    '�����i�ԓ��m�A��������ł�����
    For i = 1 To UBound(STOCKINFO)
        For j = 0 To Kihon.CNTHINNOW - 1
           If wOINF(i).hinban = wNINF(j).hinban Then
                '���������̐����𗼕��������
                If wOINF(i).MAI <= wNINF(j).MAI Then
                    wNINF(j).LEN = wNINF(j).LEN - wOINF(i).LEN
                    wNINF(j).WAT = wNINF(j).WAT - wOINF(i).WAT
                    wNINF(j).MAI = wNINF(j).MAI - wOINF(i).MAI
                    wOINF(i).LEN = 0
                    wOINF(i).WAT = 0
                    wOINF(i).MAI = 0
                Else
                    wOINF(i).LEN = wOINF(i).LEN - wNINF(j).LEN
                    wOINF(i).WAT = wOINF(i).WAT - wNINF(j).WAT
                    wOINF(i).MAI = wOINF(i).MAI - wNINF(j).MAI
                    If wOINF(i).MAI < 0 Then
                        wOINF(i).MAI = 0
                    End If
                    wNINF(j).LEN = 0
                    wNINF(j).WAT = 0
                    wNINF(j).MAI = 0
                End If
            End If
        Next
    Next
        
    For i = 0 To UBound(wOINF) - 2
        For j = i + 1 To UBound(wOINF) - 1
            If (StrComp(wOINF(i).hinban, wOINF(j).hinban, _
                vbTextCompare)) = 1 Then '�i�Ԃ̓��֕K�v
                wWINF(0) = wOINF(j)
                wOINF(j) = wOINF(i)
                wOINF(i) = wWINF(0)
            End If
        Next j
    Next i
'' 2003/04/17 add by t.t   end

'' 2003/04/17 add by t.t   start
    'wNINF�̕i�Ԃ��\�[�g����
    For i = 0 To UBound(wNINF) - 2
        For j = i + 1 To UBound(wNINF) - 1
            If (StrComp(wNINF(i).hinban, wNINF(j).hinban, _
                vbTextCompare)) = 1 Then '�i�Ԃ̓��֕K�v
                wWINF(0) = wNINF(j)
                wNINF(j) = wNINF(i)
                wNINF(i) = wWINF(0)
            End If
        Next j
    Next i
'' 2003/04/17 add by t.t   end

    '�󂫂̔z��폜����(�z��̃f�[�^���l�߂�)
    For i = 0 To wOINFMAX
        If wOINF(i).MAI <= 0 Then
            iCnt = i
            Call HairetuOpe_Mai(wOINF(), iCnt, -1)
        End If
    Next i

    '�󂫂̔z��폜����(�z��̃f�[�^���l�߂�)
    For i = 0 To wNINFMAX
        If wNINF(i).MAI <= 0 Then
            iCnt = i
            Call HairetuOpe_Mai(wNINF(), iCnt, -1)
        End If
    Next i
    
    '�i�ԓ��֏����쐬����
    i = 0 '�O�i�Ԃ̈ʒu
    j = 0 '��i�Ԃ̈ʒu
    Do
        '������˂����킹�Đ��ʂ������łȂ�������傫���l�̕i�Ԃ𕪊�����
        If (wOINF(i).MAI = wNINF(j).MAI) Then   '�i�Ԓ����������������Ƃ����ɐi��
        ElseIf (wOINF(i).MAI > wNINF(j).MAI) Then   '�i�Ԓ������قȂ鎞
            iCnt = i
            Call HairetuOpe(wOINF(), iCnt, 1)    '�z��̒ǉ�
            wOINF(i + 1).hinban = wOINF(i).hinban
            wOINF(i + 1).LEN = wOINF(i).LEN - wNINF(j).LEN
            wOINF(i + 1).WAT = wOINF(i).WAT - wNINF(j).WAT
            wOINF(i + 1).MAI = wOINF(i).MAI - wNINF(j).MAI
            wOINF(i).LEN = wNINF(j).LEN
            wOINF(i).WAT = wNINF(j).WAT
            wOINF(i).MAI = wNINF(j).MAI
'''''        ElseIf (wOINF(i).LEN < wNINF(j).LEN) Then   '�i�Ԑ��ʂ��قȂ鎞   '2003/04/17 rep by tt
        ElseIf (wOINF(i).MAI < wNINF(j).MAI) Then   '�i�Ԑ��ʂ��قȂ鎞
            iCnt = j
            Call HairetuOpe(wNINF(), iCnt, 1)
            wNINF(j + 1).hinban = wNINF(i).hinban
            wNINF(j + 1).LEN = wNINF(j).LEN - wOINF(i).LEN
            wNINF(j + 1).WAT = wNINF(j).WAT - wOINF(i).WAT
            wNINF(j + 1).MAI = wNINF(j).MAI - wOINF(i).MAI
            wNINF(j).LEN = wOINF(i).LEN
            wNINF(j).WAT = wOINF(i).WAT
            wNINF(j).MAI = wOINF(i).MAI
        End If
        wOINFrecCnt = UBound(wOINF())
        wNINFrecCnt = UBound(wNINF())
        i = i + 1
        j = j + 1
        If (i > wOINFrecCnt) Then
            Exit Do
        
        End If
        If (j > wNINFrecCnt) Then
            Exit Do
        End If
        
'''''        If (wOINF(i).LEN) <= 0 Then   '2003/04/17 rep by tt
        If (wOINF(i).MAI) <= 0 Then
            Exit Do
        End If
        
'''''        If (wNINF(j).LEN) <= 0 Then   '2003/04/17 rep by tt
        If (wNINF(j).MAI) <= 0 Then
            Exit Do
        End If
    Loop
    
    wOINFrecCnt = UBound(wOINF())
'    For i = 0 To wOINFrecCnt - 1 'Step 1
    For i = 0 To wOINFrecCnt     '04/01/13 ooba
        If (StrComp(wNINF(i).hinban, wOINF(i).hinban, vbTextCompare) <> 0) Then  '�i�Ԃ��قȂ鎞�U�֏��ɓo�^����
            If Trim(wNINF(i).hinban) <> "" And wNINF(i).LEN > 0 Then
                Koutei.CRYNUMC3 = HinNow(0).CRYNUMCA    '�u���b�N�h�c
                giInpos = giInpos + 1
                Koutei.INPOSC3 = giInpos                '�ʒu
                Koutei.KCNTC3 = wNINF(i).KCKNT          ' �H���A��
                Koutei.HINBC3 = wNINF(i).hinban         '�i��
'''                Koutei.REVNUMC3 = HinNow(i).REVNUMCA    '���i�����ԍ�
'''                Koutei.FACTORYC3 = HinNow(i).FACTORYCA  '�H��
'''                Koutei.OPEC3 = HinNow(i).OPECA          '���Ə���
                Koutei.REVNUMC3 = wNINF(i).REVNUM       '���i�����ԍ�       ''04/01/13 ooba
                Koutei.FACTORYC3 = wNINF(i).factory     '�H��               ''04/01/13 ooba
                Koutei.OPEC3 = wNINF(i).OPE             '���Ə���           ''04/01/13 ooba
                Koutei.LENC3 = wNINF(i).LEN             '�������
                Koutei.XTALC3 = HinNow(0).XTALCA        '�����ԍ�
                Koutei.SXLIDC3 = ""                     ' SXLID
                
                Koutei.KNKTC3 = Left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
                  CStr(CInt(Right(Kihon.NOWPROC, 1)) + 2) ' �Ǘ��H��(���ݍH��+2)
                Koutei.WKKTC3 = Kihon.NOWPROC           ' �H��
                Koutei.WKKBC3 = ""                      ' ��Ƌ敪
                Koutei.MACOC3 = HinNow(0).NEMACOCA      ' ������
                Koutei.MODKBC3 = ""                     ' �ԍ��敪
                Koutei.SUMKBC3 = ""                     ' �W�v�敪
                Koutei.FRKNKTC3 = ""                    ' (���)�Ǘ��H��
                If IsNull(HinOld(0).NEWKNTCA) = True Then   '(����j�H��
                    Koutei.FRWKKTC3 = ""
                Else
                    Koutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
                    'add end 2003/03/28 hitec)matsumoto --------
                End If
                Koutei.FRWKKBC3 = ""                    ' (���)��Ƌ敪
                If IsNull(HinOld(0).NEMACOCA) = True Then   '�i����j������
                    Koutei.FRMACOC3 = "0"
                Else
                    Koutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
                End If
                
''''                If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''                If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
                Select Case Kihon.NOWPROC
                Case "CC730"
                    iHantei = CInt(BlkNow.GNLC2)
                Case Else
                    iHantei = CInt(BlkNow.GNMC2)
                End Select
                If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                    Koutei.TOWKKTC3 = " "               ' (���o)�H��
                    Koutei.TOMACOC3 = "0"               '(���o)������
                Else
'''                    Koutei.TOWKKTC3 = HinNow(i).GNWKNTCA    ' (���o)�H��
'''                    Koutei.TOMACOC3 = HinNow(i).GNMACOCA    ' (���o)������
                    Koutei.TOWKKTC3 = HinNow(0).GNWKNTCA    ' (���o)�H��        ''04/01/13 ooba
                    Koutei.TOMACOC3 = HinNow(0).GNMACOCA    ' (���o)������    ''04/01/13 ooba
                End If
                Koutei.FRLC3 = wNINF(i).LEN             '�������
                Koutei.FRWC3 = wNINF(i).WAT             '����d��
                Koutei.FRMC3 = wNINF(i).MAI             '�������
                Koutei.FULC3 = 0                        '�s�ǒ���
                Koutei.FUWC3 = 0                        '�s�Ǐd��
                Koutei.FUMC3 = 0                        '�s�ǖ���
                Koutei.LOSWC3 = ""                      ' ���X����
                
                Koutei.LOSLC3 = ""                      ' ���X�d��
                Koutei.LOSMC3 = ""                      ' ���X����
                Koutei.TOLC3 = wNINF(i).LEN             '���o����
'''                Koutei.TOWC3 = wNINF(0).WAT             '���o�d��
'''                Koutei.TOMC3 = wNINF(0).MAI             '���o����
                Koutei.TOWC3 = wNINF(i).WAT             '���o�d��           ''04/01/13 ooba
                Koutei.TOMC3 = wNINF(i).MAI             '���o����           ''04/01/13 ooba
                Koutei.SUMITLC3 = ""                    ' SUMIT����
                Koutei.SUMITWC3 = ""                    ' SUMIT�d��
                Koutei.SUMITMC3 = ""                    ' SUMIT����
                Koutei.MOTHINC3 = wOINF(i).hinban       '���i��
                Koutei.XTWORKC3 = "42"                  ' �����H��
                
                Koutei.WFWORKC3 = ""                    ' ���ʐ���
    '           Koutei.STATIMEC3 = Null                 ' �����J�n�I��
    '           Koutei.STOTIMEC3 = Null                 ' �������ԏI��
    '           Koutei.ETIMEC3 = ""                     ' ���ю��Ԃ͓���Ȃ�
                Koutei.HOLDCC3 = " "                    ' �z�[���h�R�[�h
                Koutei.HOLDBC3 = "0"                    ' �z�[���h�敪
                Koutei.LDFRCC3 = ""                     ' �i���R�[�h
                Koutei.LDFRBC3 = "0"                    ' �i���敪�i�n�C�L�j
                Koutei.TSTAFFC3 = Kihon.STAFFID         ' �o�^�Ј�ID
                Koutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �o�^���t
                
                Koutei.KSTAFFC3 = ""                    ' �X�V�Ј�ID
                Koutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
                Koutei.SUMITBC3 = ""                    ' SUMIT���M�t���O
                Koutei.SNDKC3 = ""                      ' ���M�t���O
    '           Koutei.SNDDAYC3 = ""                    ' ���M���t
                Koutei.MODMACOC3 = ""                   ' �ԍ��̏�����
                Koutei.KAKUCC3 = ""                     ' �m��R�[�h
                Koutei.SUMDAYC3 = CalcSumcoTime(Koutei.KDAYC3) 'SUMCO����
                Koutei.PAYCLASSC3 = ""                  '�@�]����H��t���O
    '            Koutei.SUMITSNDC3 = ""                  ' SUMIT���M���t
                
    '            Koutei.SSENDNOC3 = ""
               
                iRtn = CreateXSDC3(Koutei, wErrMsg)     '�H�����тɍ݌Ɍ����o�^
                If iRtn = FUNCTION_RETURN_FAILURE Then  '�H�����ђǉ��G���[
                    MsgBox wErrMsg
                    Exit Function
                End If
            End If
        End If
    Next i

    XSDC3Proc3 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC3Proc3 = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

'�T�v      :�H�����ѓo�^�������s��(�݌Ɍ����FCC730�p)
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :�H������(XSDC3)�ɍ݌Ɍ����̓o�^�������s��

Public Function XSDC3Proc4() As FUNCTION_RETURN

'   �����ϐ�
    Dim i, j            As Integer
    Dim iRtn            As Integer          '���A���
    Dim sql             As String           '�r�p�k
    Dim rs              As OraDynaset       '���R�[�h�Z�b�g
    Dim sqlWhere        As String           'WHERE��
    Dim wErrMsg         As String
    Dim Koutei          As typ_XSDC3_Update    '�H������
    Dim rsKCNTC         As OraDynaset       '���R�[�h�Z�b�g
    Dim intNextCnt      As Integer
    
    Dim wLen As Long
    Dim wCHKPOS As Long
    Dim badcnt As Integer
    Dim BADINFO() As typ_bad_info
    Dim wSTOCKINFO() As typ_stock_info
    Dim iLoopCnt As Integer
    Dim vGetMaxPos  As Variant
    Dim vGetData  As Variant
    Dim sOldHinban, sNowHinban As String
    Dim iSXLcnt As Integer
    Dim iMapSt, iMapEd As Integer
    Dim bHinFlg As Boolean
    Dim lTMaisu As Long
    Dim vNukisiFlg  As Variant
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto

    '�G���[�n���h���̐ݒ�
    'On Error GoTo PROC_ERR
    On Error GoTo 0
    
    '�����ݒ�
    XSDC3Proc4 = FUNCTION_RETURN_FAILURE

    ReDim STOCKINFO(Kihon.CNTHINOLD)
    ReDim wSTOCKINFO(0)
    
   'HinOld����O�H������,�d��,�������v�擾(������0)
    For i = 0 To Kihon.CNTHINOLD - 1
        FRLC3Sum = FRLC3Sum + CLng(HinOld(i).GNLCA)    ' �O�H���������v
        FRWC3Sum = FRWC3Sum + CLng(HinOld(i).GNWCA)    ' �O�H���d�ʍ��v
        FRMC3Sum = FRMC3Sum + CLng(HinOld(i).GNMCA)    ' �O�H���������v
        '�s�Ǥ�����̏����ݒ�
        STOCKINFO(i).hinban = HinOld(i).HINBCA
        STOCKINFO(i).FURYOL = 0
        STOCKINFO(i).HARAIL = CLng(HinOld(i).GNLCA)
        STOCKINFO(i).FuryoW = CLng(HinOld(i).GNWCA) '�s�Ǐd�ʂɕ����d�ʂ����ɑ�����Č�Ōv�Z����
        STOCKINFO(i).HARAIW = CLng(HinOld(i).GNWCA)
        STOCKINFO(i).FURYOM = CLng(HinOld(i).GNMCA) '�s�ǖ����ɕ������������ɑ������Ōv�Z����
        STOCKINFO(i).HARAIM = CLng(HinOld(i).GNMCA)
        STOCKINFO(i).KCKNT = CLng(HinOld(i).KCKNTCA)
        STOCKINFO(i).REVNUM = HinOld(i).REVNUMCA        ' ���i�����ԍ�
        STOCKINFO(i).factory = HinOld(i).FACTORYCA      ' �H��
        STOCKINFO(i).OPE = HinOld(i).OPECA              ' ���i�����ԍ�
    Next i
        
'�Ŕ����w����ʂ���i�Ԃ̕����o���ƌ������}�b�v�ʒu���ڂ��狁�߂�
'STOCKINFO�z��Ɋi�[���邪STOCKINFO�̕i�Ԃ�HinOld�̕i�Ԃ̓o�^�����ƈ�v���Ă���Ƃ͌���Ȃ�
    
    badcnt = 0  '�s�ǐ������ݒ�
'    '�s�ǂ��擪�ɂȂ����m�F
    If ((CLng(HinNow(0).INPOSCA) - CLng(HinOld(0).INPOSCA)) > 0) Then '�O��J�n�ʒu���r���č�������Εs�ǈʒu�o�^
        badcnt = badcnt + 1
        ReDim Preserve BADINFO(badcnt)
        BADINFO(badcnt).pos = CLng(HinOld(0).INPOSCA)
        BADINFO(badcnt).LEN = CLng(HinNow(0).INPOSCA) - CLng(HinOld(0).INPOSCA)
    End If
'
    '�s�ǒ������i�ԊԂɂȂ����m�F
    For i = 0 To Kihon.CNTHINNOW - 2
        If (CLng(HinNow(i + 1).INPOSCA) > (CLng(HinNow(i).INPOSCA) + CLng(HinNow(i).GNLCA))) Then '�i�ԊԂɕs�ǗL
            badcnt = badcnt + 1 '�s�ǈʒu�̓o�^
            ReDim Preserve BADINFO(badcnt)
            BADINFO(badcnt).pos = CLng(HinNow(i).INPOSCA) + CLng(HinNow(i).GNLCA)
            BADINFO(badcnt).LEN = CLng(HinNow(i + 1).INPOSCA) - CLng(HinNow(i).INPOSCA) - CLng(HinNow(i).GNLCA)
        End If
    Next i
'
    '�s�ǂ��Ō�ɂȂ����O�̊m�F(�������J�n�ʒu+�����Ŕ�r)
    If ((CLng(HinOld(Kihon.CNTHINOLD - 1).INPOSCA) + CLng(HinOld(Kihon.CNTHINOLD - 1).GNLCA)) _
        <> (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))) Then '�I���ʒu�̊m�F  'upd 2003/05/31 hitec)matsumoto �u���v���u�����v�ɕύX
        badcnt = badcnt + 1
        ReDim Preserve BADINFO(badcnt)
        BADINFO(badcnt).pos = (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))
        BADINFO(badcnt).LEN = CLng(HinOld(Kihon.CNTHINOLD - 1).INPOSCA) + CLng(HinOld(Kihon.CNTHINOLD - 1).GNLCA) - (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))
    End If
'
    If (badcnt = 0) Then  '�O�ƌ�ŕs�ǂȂ� �����I��
        XSDC3Proc4 = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

'    '�s�ǈʒu��U�֑O�̌����ʒu���m�F���ĕs�ǈʒu�ɑ�������i�Ԃ�o�^����
            
    'add start 2003/05/31 hitec)matsumoto --------------
    If BlkOld.GNLC2 < BlkNow.GNLC2 Then
        For i = 1 To badcnt
            For j = 0 To Kihon.CNTHINOLD - 1
                STOCKINFO(j).FURYOL = STOCKINFO(j).FURYOL + BADINFO(i).LEN   '�s�ǂ̒���(HinOld(i)���������Ă���i�Ԃ̒����s��)
                STOCKINFO(j).HARAIL = STOCKINFO(j).HARAIL - BADINFO(i).LEN  '�Ǖi����
            Next j
        Next i
        For i = 0 To Kihon.CNTHINOLD - 1
            If ((STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) > 0) Then  '�s�ǐ������݂�����s�Ǐd����s�ǔ䗦�ŋ��߂�
                If i = Kihon.CNTHINOLD - 1 Then
                    'STOCKINFO(i).HARAIW = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIW 'STOCKINFO(i).HARAIW�͓��͍ς�
                    STOCKINFO(i).HARAIW = Round((STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL)) * STOCKINFO(i).HARAIW)  'STOCKINFO(i).HARAIW�͓��͍ς�  �f2003/08/06 hitec)matsumoto ROUND�ǉ�
                    STOCKINFO(i).FuryoW = STOCKINFO(i).FuryoW - STOCKINFO(i).HARAIW 'STOCKINFO(i).FURYOW�͓��͍ς�
                    'STOCKINFO(i).HARAIM = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIM 'STOCKINFO(i).HARAIM�͓��͍ς�
                    STOCKINFO(i).HARAIM = Round((STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL)) * STOCKINFO(i).HARAIM)  'STOCKINFO(i).HARAIM�͓��͍ς�  �f2003/08/06 hitec)matsumoto ROUND�ǉ�
                    STOCKINFO(i).FURYOM = STOCKINFO(i).FURYOM - STOCKINFO(i).HARAIM 'STOCKINFO(i).FURYOM�͓��͍ς�
                Else
                    STOCKINFO(i).HARAIW = HinOld(i).GNWCA
                    STOCKINFO(i).FuryoW = 0
                    STOCKINFO(i).HARAIM = HinOld(i).GNMCA
                    STOCKINFO(i).FURYOM = 0
                End If
            End If
        Next i
    Else
        'STOCKINFO�̕����͊��ɓ��͍ς�
        For i = 1 To badcnt
            For j = 0 To Kihon.CNTHINOLD - 1
                If (BADINFO(i).pos >= CLng(HinOld(j).INPOSCA) And _
                    BADINFO(i).pos < CLng(HinOld(j).INPOSCA) + CLng(HinOld(j).GNLCA)) Then
                    STOCKINFO(j).FURYOL = STOCKINFO(j).FURYOL + BADINFO(i).LEN   '�s�ǂ̒���(HinOld(i)���������Ă���i�Ԃ̒����s��)
                    STOCKINFO(j).HARAIL = STOCKINFO(j).HARAIL - BADINFO(i).LEN  '�Ǖi����
                End If
            Next j
        Next i
    
        '�d�ʂƖ����̕s�ǂƕ����o���̒l��ݒ肷��
        For i = 0 To Kihon.CNTHINOLD - 1
            If ((STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) > 0) Then  '�s�ǐ������݂�����s�Ǐd����s�ǔ䗦�ŋ��߂�
                STOCKINFO(i).HARAIW = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIW 'STOCKINFO(i).HARAIW�͓��͍ς�
                STOCKINFO(i).FuryoW = STOCKINFO(i).FuryoW - STOCKINFO(i).HARAIW 'STOCKINFO(i).FURYOW�͓��͍ς�
                STOCKINFO(i).HARAIM = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIM 'STOCKINFO(i).HARAIM�͓��͍ς�
                STOCKINFO(i).FURYOM = STOCKINFO(i).FURYOM - STOCKINFO(i).HARAIM 'STOCKINFO(i).FURYOM�͓��͍ς�
            End If
        Next i
    End If
    'add end 2003/05/31 hitec)matsumoto --------------
    
    '�s�ǂ�����ꍇ���݌Ɍ����̍쐬
    For i = 0 To Kihon.CNTHINOLD - 1
        If STOCKINFO(i).FURYOL <> 0 Then
            Koutei.CRYNUMC3 = HinNow(0).CRYNUMCA    '�u���b�N�h�c
            giInpos = giInpos + 1
            Koutei.INPOSC3 = giInpos                '�ʒu
            Koutei.KCNTC3 = STOCKINFO(i).KCKNT + 1  '�H���A��
            Koutei.HINBC3 = HinOld(i).HINBCA        '�i��
            Koutei.REVNUMC3 = HinOld(i).REVNUMCA    '���i�����ԍ�
            Koutei.FACTORYC3 = HinOld(i).FACTORYCA  '�H��
            Koutei.OPEC3 = HinOld(i).OPECA          '���Ə���
            Koutei.LENC3 = STOCKINFO(i).HARAIL      '����
            Koutei.XTALC3 = HinOld(i).XTALCA        '�����ԍ�
            Koutei.SXLIDC3 = ""                     ' SXLID
            
            Koutei.KNKTC3 = Left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 1) ' �Ǘ��H��(���ݍH��+1)
            Koutei.WKKTC3 = Kihon.NOWPROC           ' �H��
            Koutei.WKKBC3 = ""                      ' ��Ƌ敪
            Koutei.MACOC3 = HinNow(0).NEMACOCA      ' ������
            Koutei.MODKBC3 = ""                     ' �ԍ��敪
            Koutei.SUMKBC3 = ""                     ' �W�v�敪
            Koutei.FRKNKTC3 = ""                    ' (���)�Ǘ��H��
            If IsNull(HinOld(0).NEWKNTCA) = True Then   '(����j�H��
                Koutei.FRWKKTC3 = ""
            Else
                Koutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
                'add end 2003/03/28 hitec)matsumoto --------
            End If
            Koutei.FRWKKBC3 = ""                    ' (���)��Ƌ敪
            If IsNull(HinOld(0).NEMACOCA) = True Then   '�i����j������
                Koutei.FRMACOC3 = "0"
            Else
                Koutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If
''''            If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''            If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
            Select Case Kihon.NOWPROC
            Case "CC730"
                iHantei = CInt(BlkNow.GNLC2)
            Case Else
                iHantei = CInt(BlkNow.GNMC2)
            End Select
            If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                Koutei.TOWKKTC3 = " "               ' (���o)�H��
                Koutei.TOMACOC3 = "0"               '(���o)������
            Else
                Koutei.TOWKKTC3 = HinNow(0).GNWKNTCA    ' (���o)�H��
                Koutei.TOMACOC3 = HinNow(0).GNMACOCA    ' (���o)������
            End If
            Koutei.FRLC3 = HinOld(i).GNLCA          '�������
            Koutei.FRWC3 = HinOld(i).GNWCA          '����d��
            Koutei.FRMC3 = HinOld(i).GNMCA          '�������
            Koutei.FULC3 = STOCKINFO(i).FURYOL       '�s�ǒ���
            Koutei.FUWC3 = STOCKINFO(i).FuryoW       '�s�Ǐd��
            Koutei.FUMC3 = STOCKINFO(i).FURYOM       '�s�ǖ���
            Koutei.LOSWC3 = ""                      ' ���X����
            
            Koutei.LOSLC3 = ""                      ' ���X�d��
            Koutei.LOSMC3 = ""                      ' ���X����
            Koutei.TOLC3 = STOCKINFO(i).HARAIL       '���o����
            Koutei.TOWC3 = STOCKINFO(i).HARAIW       '���o�d��
            Koutei.TOMC3 = STOCKINFO(i).HARAIM       '���o����
            Koutei.SUMITLC3 = ""                    ' SUMIT����
            Koutei.SUMITWC3 = ""                    ' SUMIT�d��
            Koutei.SUMITMC3 = ""                    ' SUMIT����
            Koutei.MOTHINC3 = " "       '���i��
            Koutei.XTWORKC3 = "42"                  ' �����H��
            
            Koutei.WFWORKC3 = ""                    ' ���ʐ���
'           Koutei.STATIMEC3 = Null                 ' �����J�n�I��
'           Koutei.STOTIMEC3 = Null                 ' �������ԏI��
'           Koutei.ETIMEC3 = ""                     ' ���ю��Ԃ͓���Ȃ�
            Koutei.HOLDCC3 = " "                    ' �z�[���h�R�[�h
            Koutei.HOLDBC3 = "0"                    ' �z�[���h�敪
            Koutei.LDFRCC3 = ""                     ' �i���R�[�h
            Koutei.LDFRBC3 = "0"                    ' �i���敪�i�n�C�L�j
            Koutei.TSTAFFC3 = Kihon.STAFFID         ' �o�^�Ј�ID
            Koutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �o�^���t
            
            Koutei.KSTAFFC3 = ""                    ' �X�V�Ј�ID
            Koutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
            Koutei.SUMITBC3 = ""                    ' SUMIT���M�t���O
            Koutei.SNDKC3 = ""                      ' ���M�t���O
'           Koutei.SNDDAYC3 = ""                    ' ���M���t
            Koutei.MODMACOC3 = ""                   ' �ԍ��̏�����
            Koutei.KAKUCC3 = ""                     ' �m��R�[�h
            Koutei.SUMDAYC3 = CalcSumcoTime(Koutei.KDAYC3) 'SUMCO����
            Koutei.PAYCLASSC3 = ""                  '�@�]����H��t���O
'            Koutei.SUMITSNDC3 = ""                  ' SUMIT���M���t
            
'            Koutei.SSENDNOC3 = ""
            
            iRtn = CreateXSDC3(Koutei, wErrMsg)     '�H�����тɍ݌Ɍ����o�^
            If iRtn = FUNCTION_RETURN_FAILURE Then  '�H�����ђǉ��G���[
                MsgBox wErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc4 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

'�T�v      :�H�����ѓo�^�������s��(�i�ԐU�֏��FCC730�p)
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :�H������(XSDC3)�ɕi�ԐU�֏��̓o�^�������s��

Public Function XSDC3Proc5() As FUNCTION_RETURN

'   �����ϐ�
    Dim i, j            As Integer
    Dim iRtn            As Integer          '���A���
    Dim sql             As String           '�r�p�k
    Dim sqlWhere        As String           'WHERE��
    Dim wErrMsg         As String
    Dim Koutei          As typ_XSDC3_Update    '�H������
    
    Dim wLen            As Long
    Dim wCHKPOS         As Long
        
    Dim wOINF()         As typ_trans_info   '�O�i�ԕ��ёւ��p
    Dim wNINF()         As typ_trans_info   '��i�ԕ��ёւ��p
    Dim wWINF()         As typ_trans_info   '���ёւ��p���[�N
    Dim ibuf            As Integer
    Dim wOINFrecCnt     As Integer
    Dim wNINFrecCnt     As Integer
    Dim wOINFFLG        As Integer
    Dim wNINFFLG        As Integer
    Dim iPoint          As Integer
    Dim wNINFMAX        As Integer
    Dim wOINFMAX        As Integer
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    '�����ݒ�
    XSDC3Proc5 = FUNCTION_RETURN_FAILURE
    
    ReDim wOINF(UBound(STOCKINFO))      '�i�Ԃ��ƂɃ\�[�g�p
    ReDim wNINF(Kihon.CNTHINNOW)        '�i�Ԃ��ƂɃ\�[�g�p
    ReDim wWINF(1)                      '�i�Ԃ��ƂɃ\�[�g�p
        
    ' �݌Ɍ�������荞��
    For i = 0 To UBound(STOCKINFO) - 1
        wOINF(i).hinban = STOCKINFO(i).hinban
        wOINF(i).LEN = STOCKINFO(i).HARAIL       ' �O�H���������v
        wOINF(i).WAT = STOCKINFO(i).HARAIW       ' �O�H���d�ʍ��v
        wOINF(i).MAI = STOCKINFO(i).HARAIM       ' �O�H���������v
    Next i
    
    ' �Ǖi������荞��
    For j = 0 To Kihon.CNTHINNOW - 1
        wNINF(j).hinban = HinNow(j).HINBCA
        wNINF(j).LEN = HinNow(j).GNLCA           ' ��H���������v
        wNINF(j).WAT = HinNow(j).GNWCA           ' ��H���d�ʍ��v
        wNINF(j).MAI = HinNow(j).GNMCA           ' ��H���������v
        wNINF(j).KCKNT = HinNow(j).KCKNTCA       ' ��H���A��
    Next j
        
    '�����i�ԓ��m�A��������ł�����
    For i = 0 To UBound(STOCKINFO) - 1
        For j = 0 To Kihon.CNTHINNOW - 1
           If wOINF(i).hinban = wNINF(j).hinban Then
                '���������̐����𗼕��������
                If wOINF(i).LEN <= wNINF(j).LEN Then
                    wNINF(j).LEN = wNINF(j).LEN - wOINF(i).LEN
                    wNINF(j).WAT = wNINF(j).WAT - wOINF(i).WAT
                    wNINF(j).MAI = wNINF(j).MAI - wOINF(i).MAI
                    wOINF(i).LEN = 0
                    wOINF(i).WAT = 0
                    wOINF(i).MAI = 0
                Else
                    wOINF(i).LEN = wOINF(i).LEN - wNINF(j).LEN
                    wOINF(i).WAT = wOINF(i).WAT - wNINF(j).WAT
                    wOINF(i).MAI = wOINF(i).MAI - wNINF(j).MAI
                    If wOINF(i).MAI < 0 Then
                        wOINF(i).MAI = 0
                    End If
                    wNINF(j).LEN = 0
                    wNINF(j).WAT = 0
                    wNINF(j).MAI = 0
                End If
            End If
        Next
    Next
        
    For i = 0 To UBound(wOINF) - 2
        For j = i + 1 To UBound(wOINF) - 1
            If (StrComp(wOINF(i).hinban, wOINF(j).hinban, _
                vbTextCompare)) = 1 Then '�i�Ԃ̓��֕K�v
                wWINF(0) = wOINF(j)
                wOINF(j) = wOINF(i)
                wOINF(i) = wWINF(0)
            End If
        Next j
    Next i
    
    'wNINF�̕i�Ԃ��\�[�g����
    For i = 0 To UBound(wNINF) - 2
        For j = i + 1 To UBound(wNINF) - 1
            If (StrComp(wNINF(i).hinban, wNINF(j).hinban, _
                vbTextCompare)) = 1 Then '�i�Ԃ̓��֕K�v
                wWINF(0) = wNINF(j)
                wNINF(j) = wNINF(i)
                wNINF(i) = wWINF(0)
            End If
        Next j
    Next i
        
        
    '�󂫂̔z��폜����(�z��̃f�[�^���l�߂�)
    For i = 0 To wOINFMAX
        If wOINF(i).LEN <= 0 Then
            iPoint = i
            Call HairetuOpe(wOINF(), iPoint, -1)
        End If
    Next i
        
    '�󂫂̔z��폜����(�z��̃f�[�^���l�߂�)
    For i = 0 To wNINFMAX
        If wNINF(i).LEN <= 0 Then
            iPoint = i
            Call HairetuOpe(wNINF(), iPoint, -1)
        End If
    Next i
    
    '�i�ԓ��֏����쐬����
    i = 0 '�O�i�Ԃ̈ʒu
    j = 0 '��i�Ԃ̈ʒu
    Do
        '������˂����킹�Đ��ʂ������łȂ�������傫���l�̕i�Ԃ𕪊�����
        If (wOINF(i).LEN = wNINF(j).LEN And wOINF(i).hinban = wNINF(j).hinban) Then   '�i�Ԓ����������������Ƃ����ɐi��
        ElseIf (wOINF(i).LEN >= wNINF(j).LEN) Then   '�i�Ԗ������قȂ鎞
            iPoint = i
            Call HairetuOpe(wOINF(), iPoint, 1)    '�z��̒ǉ�
            wOINF(i + 1).hinban = wOINF(i).hinban
            wOINF(i + 1).LEN = wOINF(i).LEN - wNINF(j).LEN
            wOINF(i + 1).WAT = wOINF(i).WAT - wNINF(j).WAT
            wOINF(i + 1).MAI = wOINF(i).MAI - wNINF(j).MAI
            If wOINF(i + 1).MAI < 0 Then
                wOINF(i + 1).MAI = 0
            End If
            wOINF(i).LEN = wNINF(j).LEN
            wOINF(i).WAT = wNINF(j).WAT
            wOINF(i).MAI = wNINF(j).MAI
            Debug.Print "HINBAN=", i, wOINF(i).hinban
            Debug.Print "LEN=", i, wOINF(i).LEN
        ElseIf (wOINF(i).LEN < wNINF(j).LEN) Then   '�i�Ԑ��ʂ��قȂ鎞
            iPoint = j
            Call HairetuOpe(wNINF(), iPoint, 1)
            wNINF(j + 1).hinban = wNINF(i).hinban
            wNINF(j + 1).LEN = wNINF(j).LEN - wOINF(i).LEN
            wNINF(j + 1).WAT = wNINF(j).WAT - wOINF(i).WAT
            wNINF(j + 1).MAI = wNINF(j).MAI - wOINF(i).MAI
            wNINF(j).LEN = wOINF(i).LEN
            wNINF(j).WAT = wOINF(i).WAT
            wNINF(j).MAI = wOINF(i).MAI
            Debug.Print "HINBAN=", i, wNINF(i).hinban
            Debug.Print "LEN=", i, wNINF(i).LEN
        End If
        wOINFrecCnt = UBound(wOINF())
        wNINFrecCnt = UBound(wNINF())
        i = i + 1
        j = j + 1
        If (i > wOINFrecCnt) Then
            Exit Do
        
        End If
        If (j > wNINFrecCnt) Then
            Exit Do
        End If
        
        If (wOINF(i).LEN) <= 0 Then
            Exit Do
        End If

        If (wNINF(j).LEN) <= 0 Then
            Exit Do
        End If
    Loop

    wOINFrecCnt = UBound(wOINF())
    For i = 0 To wOINFrecCnt - 1
''''        wNINF(i).HINBAN
        If (StrComp(wNINF(i).hinban, wOINF(i).hinban, vbTextCompare) <> 0 _
            And Len(Trim(wNINF(i).hinban) > 0)) And (wNINF(i).LEN > 0) Then '�i�Ԃ��قȂ鎞�U�֏��ɓo�^����   'upd 2003/05/31 hitec)matsumoto wNINF(i).LEN > 0�ǉ�
            
            Koutei.CRYNUMC3 = HinNow(0).CRYNUMCA    '�u���b�N�h�c
            giInpos = giInpos + 1
            Koutei.INPOSC3 = giInpos                '�ʒu
            Koutei.KCNTC3 = wNINF(i).KCKNT          ' �H���A��
            Koutei.HINBC3 = wNINF(i).hinban         '�i��
            Koutei.REVNUMC3 = HinNow(0).REVNUMCA    '���i�����ԍ�
            Koutei.FACTORYC3 = HinNow(0).FACTORYCA  '�H��
            Koutei.OPEC3 = HinNow(0).OPECA          '���Ə���
            Koutei.LENC3 = wNINF(i).LEN             '�������
            Koutei.XTALC3 = HinNow(0).XTALCA        '�����ԍ�
            Koutei.SXLIDC3 = ""                     ' SXLID
            
            Koutei.KNKTC3 = Left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 2) ' �Ǘ��H��(���ݍH��+2)
            Koutei.WKKTC3 = Kihon.NOWPROC           ' �H��
            Koutei.WKKBC3 = ""                      ' ��Ƌ敪
            Koutei.MACOC3 = HinNow(0).NEMACOCA      ' ������
            Koutei.MODKBC3 = ""                     ' �ԍ��敪
            Koutei.SUMKBC3 = ""                     ' �W�v�敪
            Koutei.FRKNKTC3 = ""                    ' (���)�Ǘ��H��
            If IsNull(HinOld(0).NEWKNTCA) = True Then   '(����j�H��
                Koutei.FRWKKTC3 = ""
            Else
                Koutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
                'add end 2003/03/28 hitec)matsumoto --------
            End If
            Koutei.FRWKKBC3 = ""                    ' (���)��Ƌ敪
            If IsNull(HinOld(0).NEMACOCA) = True Then   '�i����j������
                Koutei.FRMACOC3 = "0"
            Else
                Koutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If
            
''''            If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''            If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   '2003/05/05 hitec)matsumoto
            Select Case Kihon.NOWPROC
            Case "CC730"
                iHantei = CInt(BlkNow.GNLC2)
            Case Else
                iHantei = CInt(BlkNow.GNMC2)
            End Select
            If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                Koutei.TOWKKTC3 = " "               ' (���o)�H��
                Koutei.TOMACOC3 = "0"               '(���o)������
            Else
                Koutei.TOWKKTC3 = HinNow(0).GNWKNTCA    ' (���o)�H��
                Koutei.TOMACOC3 = HinNow(0).GNMACOCA    ' (���o)������
            End If
            Koutei.FRLC3 = wNINF(i).LEN             '�������
            Koutei.FRWC3 = wNINF(i).WAT             '����d��
            Koutei.FRMC3 = wNINF(i).MAI             '�������
            Koutei.FULC3 = 0                        '�s�ǒ���
            Koutei.FUWC3 = 0                        '�s�Ǐd��
            Koutei.FUMC3 = 0                        '�s�ǖ���
            Koutei.LOSWC3 = ""                      ' ���X����
            
            Koutei.LOSLC3 = ""                      ' ���X�d��
            Koutei.LOSMC3 = ""                      ' ���X����
            Koutei.TOLC3 = wNINF(i).LEN             '���o����
            Koutei.TOWC3 = wNINF(i).WAT             '���o�d��
            Koutei.TOMC3 = wNINF(i).MAI             '���o����
            Koutei.SUMITLC3 = ""                    ' SUMIT����
            Koutei.SUMITWC3 = ""                    ' SUMIT�d��
            Koutei.SUMITMC3 = ""                    ' SUMIT����
            Koutei.MOTHINC3 = wOINF(i).hinban       '���i��
            Koutei.XTWORKC3 = "42"                  ' �����H��
            
            Koutei.WFWORKC3 = ""                    ' ���ʐ���
'           Koutei.STATIMEC3 = Null                 ' �����J�n�I��
'           Koutei.STOTIMEC3 = Null                 ' �������ԏI��
'           Koutei.ETIMEC3 = ""                     ' ���ю��Ԃ͓���Ȃ�
            Koutei.HOLDCC3 = " "                    ' �z�[���h�R�[�h
            Koutei.HOLDBC3 = "0"                    ' �z�[���h�敪
            Koutei.LDFRCC3 = ""                     ' �i���R�[�h
            Koutei.LDFRBC3 = "0"                    ' �i���敪�i�n�C�L�j
            Koutei.TSTAFFC3 = Kihon.STAFFID         ' �o�^�Ј�ID
            Koutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �o�^���t
            
            Koutei.KSTAFFC3 = ""                    ' �X�V�Ј�ID
            Koutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' �X�V���t
            Koutei.SUMITBC3 = ""                    ' SUMIT���M�t���O
            Koutei.SNDKC3 = ""                      ' ���M�t���O
'           Koutei.SNDDAYC3 = ""                    ' ���M���t
            Koutei.MODMACOC3 = ""                   ' �ԍ��̏�����
            Koutei.KAKUCC3 = ""                     ' �m��R�[�h
            Koutei.SUMDAYC3 = CalcSumcoTime(Koutei.KDAYC3) 'SUMCO����
            Koutei.PAYCLASSC3 = ""                  '�@�]����H��t���O
'            Koutei.SUMITSNDC3 = ""                  ' SUMIT���M���t
            
'            Koutei.SSENDNOC3 = ""
            
            iRtn = CreateXSDC3(Koutei, wErrMsg)     '�H�����тɍ݌Ɍ����o�^
            If iRtn = FUNCTION_RETURN_FAILURE Then  '�H�����ђǉ��G���[
                MsgBox wErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc5 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
'    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print Err.Description & "�F" & Err.MAIber
    XSDC3Proc5 = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

Public Function HairetuOpe(HinInf() As typ_trans_info, HinNum As Integer, HINFLG As Integer)
    Dim recCnt As Integer
    Dim i, j As Integer
    Dim sflg As Integer
    
    sflg = 0
    recCnt = UBound(HinInf())
    
    If (HINFLG = 1) Then    'HinNum�Ԗڂ̔z����󂫂ɂ���(�z��f�[�^�����ɂ��炵�ċ󂯂�)
        For i = HinNum + 1 To recCnt '�����̔z��ɋ󂫏ꏊ��T��
            If (HinInf(i).LEN <= 0) Then    'i�Ԗڂɋ󂫂��������̂Ńf�[�^�����炷
                For j = i To HinNum + 1 Step -1
                    HinInf(j).hinban = HinInf(j - 1).hinban
                    HinInf(j).LEN = HinInf(j - 1).LEN
                    HinInf(j).WAT = HinInf(j - 1).WAT
                    HinInf(j).MAI = HinInf(j - 1).MAI
                    HinInf(j).KCKNT = HinInf(j - 1).KCKNT
                Next j
                sflg = 1
                Exit For
            End If
        Next i
        If (sflg = 0) Then  '�󂫌����炸
            ReDim Preserve HinInf(recCnt + 1)
            For i = recCnt + 1 To HinNum + 1 Step -1
                HinInf(i).hinban = HinInf(i - 1).hinban
                HinInf(i).LEN = HinInf(i - 1).LEN
                HinInf(i).WAT = HinInf(i - 1).WAT
                HinInf(i).MAI = HinInf(i - 1).MAI
                HinInf(i).KCKNT = HinInf(i - 1).KCKNT
            Next i
        End If
        'HinNum+1�Ԗڂ��󂫂ɂ���
        HinInf(HinNum + 1).hinban = ""
        HinInf(HinNum + 1).LEN = 0
        HinInf(HinNum + 1).MAI = 0
        HinInf(HinNum + 1).WAT = 0
        HinInf(HinNum + 1).KCKNT = 0

    Else    'HinNum�Ԗڂ̔z����폜����(�z��f�[�^��O�ɂ߂�)
        i = HinNum
        HinInf(HinNum).hinban = ""
        HinInf(HinNum).LEN = 0
        HinInf(HinNum).MAI = 0
        HinInf(HinNum).WAT = 0
        HinInf(HinNum).KCKNT = 0
        For j = HinNum + 1 To recCnt
            If (HinInf(j).LEN > 0) Then 'HinNum�ȍ~�Ńf�[�^�����݂��Ă�����
                HinInf(i).hinban = HinInf(j).hinban
                HinInf(i).LEN = HinInf(j).LEN
                HinInf(i).MAI = HinInf(j).MAI
                HinInf(i).WAT = HinInf(j).WAT
                HinInf(i).KCKNT = HinInf(j).KCKNT
                HinInf(j).hinban = ""
                HinInf(j).LEN = 0
                HinInf(j).MAI = 0
                HinInf(j).WAT = 0
                HinInf(j).KCKNT = 0
                i = i + 1
            Else
                HinInf(j).hinban = ""
                HinInf(j).LEN = 0
                HinInf(j).MAI = 0
                HinInf(j).WAT = 0
                HinInf(j).KCKNT = 0
             End If
        Next j
End If

End Function

Public Function HairetuOpe_Mai(HinInf() As typ_trans_info, HinNum As Integer, HINFLG As Integer)
    Dim recCnt As Integer
    Dim i, j As Integer
    Dim sflg As Integer
    
    sflg = 0
    recCnt = UBound(HinInf())
    
    If (HINFLG = 1) Then    'HinNum�Ԗڂ̔z����󂫂ɂ���(�z��f�[�^�����ɂ��炵�ċ󂯂�)
        For i = HinNum + 1 To recCnt '�����̔z��ɋ󂫏ꏊ��T��
            If (HinInf(i).MAI <= 0) Then    'i�Ԗڂɋ󂫂��������̂Ńf�[�^�����炷
                For j = i To HinNum + 1 Step -1
                    HinInf(j).hinban = HinInf(j - 1).hinban
                    HinInf(j).LEN = HinInf(j - 1).LEN
                    HinInf(j).WAT = HinInf(j - 1).WAT
                    HinInf(j).MAI = HinInf(j - 1).MAI
                    HinInf(j).KCKNT = HinInf(j - 1).KCKNT
                Next j
                sflg = 1
                Exit For
            End If
        Next i
        If (sflg = 0) Then  '�󂫌����炸
            ReDim Preserve HinInf(recCnt + 1)
            For i = recCnt + 1 To HinNum + 1 Step -1
                HinInf(i).hinban = HinInf(i - 1).hinban
                HinInf(i).LEN = HinInf(i - 1).LEN
                HinInf(i).WAT = HinInf(i - 1).WAT
                HinInf(i).MAI = HinInf(i - 1).MAI
                HinInf(i).KCKNT = HinInf(i - 1).KCKNT
            Next i
        End If
        'HinNum+1�Ԗڂ��󂫂ɂ���
        HinInf(HinNum + 1).hinban = ""
        HinInf(HinNum + 1).LEN = 0
        HinInf(HinNum + 1).MAI = 0
        HinInf(HinNum + 1).WAT = 0
        HinInf(HinNum + 1).KCKNT = 0

    Else    'HinNum�Ԗڂ̔z����폜����(�z��f�[�^��O�ɂ߂�)
        i = HinNum
        HinInf(HinNum).hinban = ""
        HinInf(HinNum).LEN = 0
        HinInf(HinNum).MAI = 0
        HinInf(HinNum).WAT = 0
        HinInf(HinNum).KCKNT = 0
        For j = HinNum + 1 To recCnt
            If (HinInf(j).MAI > 0) Then 'HinNum�ȍ~�Ńf�[�^�����݂��Ă�����
                HinInf(i).hinban = HinInf(j).hinban
                HinInf(i).LEN = HinInf(j).LEN
                HinInf(i).MAI = HinInf(j).MAI
                HinInf(i).WAT = HinInf(j).WAT
                HinInf(i).KCKNT = HinInf(j).KCKNT
                HinInf(j).hinban = ""
                HinInf(j).LEN = 0
                HinInf(j).MAI = 0
                HinInf(j).WAT = 0
                HinInf(j).KCKNT = 0
                i = i + 1
            Else
                HinInf(j).hinban = ""
                HinInf(j).LEN = 0
                HinInf(j).MAI = 0
                HinInf(j).WAT = 0
                HinInf(j).KCKNT = 0
             End If
        Next j
End If

End Function

