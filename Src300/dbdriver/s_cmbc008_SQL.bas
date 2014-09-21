Attribute VB_Name = "s_cmbc008_SQL"
Option Explicit

Public Type type_DBDRV_cmgc001f1
    CRYNUM As String * 12   ' �����ԍ�
    INGOTPOS As Integer     ' �J�n�ʒu
    HINBAN As String * 8    ' Top�i�ԁ��i��
    REVNUM As Integer       ' ���i�ԍ������ԍ�
    factory As String * 1   ' �H��
    opecond As String * 1   ' ���Ə���
End Type

Public Type type_DBDRV_cmgc001f2
    PALTNUM As String * 4   ' �p���b�g�ԍ�
    BDCODE As String * 3    ' �s�Ǘ��R�R�[�h���i���敪
    BLOCKID As String * 12  ' �u���b�NID
End Type

Public Type type_DBDRV_cmgc001f3
    PGID As String * 8      ' �o�f�|�h�c
    CRYNUM As String * 12   ' �����ԍ�
End Type

Public Type type_DBDRV_cmgc001f4
    DMTOP1 As Double        ' ���a
    DMTOP2 As Double        ' ���a
    DMTAIL1 As Double       ' ���a
    DMTAIL2 As Double       ' ���a
    INGOTPOS As Integer     ' �J�n�ʒu
    TRANCNT As Integer      ' ������
    NCHPOS As String * 2    ' �m�b�`�ʒu
    CRYNUM As String * 12   ' �����ԍ�
End Type

Public Type type_DBDRV_cmgc001f5
    CRYNUM As String * 12   ' �����ԍ�
    MAGTYPE As String * 2   ' ����^�C�v�������@
End Type

Public Type type_DBDRV_cmgc001f6
'2002/04/25 S.Sano    TYPE As String * 1      ' �i�r�w�^�C�v���^�C�v
    HSXCDIR As String * 1   ' �i�r�w�����ʕ��ʁ�����
    HINBAN As String * 8    ' Top�i�ԁ��i��
End Type

Public Type type_DBDRV_cmgc001f7
    DPNTCLS As String * 7   ' �h�[�p���g��ށ��h�[�p���g�@�����ԍ��O7��+"00"�����グ�w����
    CRYNUM As String * 12   ' �����ԍ�
    TYPE As String * 2      ' �^�C�v'2002/04/25 S.Sano
End Type

Public Type type_DBDRV_cmgc001f8
    SMPLNO As Integer       ' �T���v����
    CRYNUM As String * 12   ' �����ԍ�
    INGOTPOS As Integer     ' �J�n�ʒu
    SMPKBN As String * 1    ' �T���v���敪
End Type

Public Type type_DBDRV_cmgc001f9
    CRYNUM As String * 12   ' �����ԍ�
    TRANCNT As Integer      ' ������
    TRANCOND As String * 1  ' ��������
    INGOTPOS As Integer     ' �J�n�ʒu
    SMPLNO As Long          ' �T���v����        Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPKBN As String * 1    ' �T���v���敪
    JudgData As Double      ' Top�T���v�����ɑΉ����錟���Ώےl
End Type

Public Type type_DBDRV_cmgc001f10
    CRYNUM As String * 12   ' �����ԍ�
    PRCMCN As String * 1    ' ����@
    SEED As String * 4      ' SEED
End Type

' �N���X�^���J�^���O�����i��
Public Type type_DBDRV_scmzc_fcmgc001f_Hinban
    ' �������
    MAGTYPE As String * 2   ' ����^�C�v
    ' ���i�d�lSXL�f�[�^1
    HSXD1MIN As Double      ' �i�r�w���a�P����
    HSXD1MAX As Double      ' �i�r�w���a�P���
    ' ���i�d�lSXL�f�[�^1
    HSXTYPE As String * 1   ' �i�r�w�^�C�v
    HSXDOP As String * 1    ' �i�r�w�h�[�p���g
    HSXCDIR As String * 1   ' �i�r�w�����ʕ���
    HSXRMIN As Double       ' �i�r�w���R����
    HSXRMAX As Double       ' �i�r�w���R���
    ' ���i�d�lSXL�f�[�^2
    HSXONMIN As Double      ' �i�r�w�_�f�Z�x����
    HSXONMAX As Double      ' �i�r�w�_�f�Z�x���
    HSXCNMIN As Double      ' �i�r�w�Y�f�Z�x����
    HSXCNMAX As Double      ' �i�r�w�Y�f�Z�x���
    ' ���i�d�lSXL�f�[�^1
    HSXDPDIR As String * 2  ' �i�r�w�a�ʒu����
    ' ���i�d�lSXL�f�[�^3
    HSXDVDMX As Integer     ' �i�r�w�c�u�c�Q���
    HSXDVDMN As Integer     ' �i�r�w�c�u�c�Q����
    HSXOS1AX As Double      ' �i�r�w�n�r�e�P���Ϗ��
    HSXOS1MX As Double      ' �i�r�w�n�r�e�P���
    HSXOS2AX As Double      ' �i�r�w�n�r�e�Q���Ϗ��
    HSXOS2MX As Double      ' �i�r�w�n�r�e�Q���
    HSXOS3AX As Double      ' �i�r�w�n�r�e�R���Ϗ��
    HSXOS3MX As Double      ' �i�r�w�n�r�e�R���
    HSXOS4AX As Double      ' �i�r�w�n�r�e�S���Ϗ��
    HSXOS4MX As Double      ' �i�r�w�n�r�e�S���
    HSXBM1AN As Double      ' �i�r�w�a�l�c�P���ω���
    HSXBM1AX As Double      ' �i�r�w�a�l�c�P���Ϗ��
    HSXBM2AN As Double      ' �i�r�w�a�l�c�Q���ω���
    HSXBM2AX As Double      ' �i�r�w�a�l�c�Q���Ϗ��
    HSXBM3AN As Double      ' �i�r�w�a�l�c�R���ω���
    HSXBM3AX As Double      ' �i�r�w�a�l�c�R���Ϗ��
    HSXLTMIN As Integer     ' �i�r�w�k�^�C������
    HSXLTMAX As Integer     ' �i�r�w�k�^�C�����

    SGLENGTH As Integer     ' �Œፇ�i����
End Type

Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
    ' �����p���
    CRYNUM As String * 12   ' �����ԍ�
    INGOTPOS As Integer     ' �J�n�ʒu
    TOPSMPKBN As String * 1 ' Top�T���v���敪
    BOTSMPKBN As String * 1 ' Bot�T���v���敪
    REVNUM As Integer       ' ���i�ԍ������ԍ�
    factory As String * 1   ' �H��
    opecond As String * 1   ' ���Ə���
    ' �u���b�N�Ǘ�
    BLOCKID As String * 12  ' �u���b�NID
    LENGTH As Integer       ' �������u���b�N��
    ' �i�ԊǗ�
    HINBAN As String * 8    ' Top�i�ԁ��i��
    ' �N���X�^���J�^���O�������
    PALTNUM As String * 4   ' �p���b�g�ԍ�
    BDCODE As String * 3    ' �s�Ǘ��R�R�[�h���i���敪
    ' �������
    DIAMETER As Double      ' ���a
    PGID As String * 8      ' �o�f�|�h�c
    ' ������R����
    TOPRES As Double        ' Top�T���v�����ɑΉ����錟���Ώےl��Top�����
    BOTRES As Double        ' Bot�T���v�����ɑΉ����錟���Ώےl��Bot�����
    TOPRESSMP As Integer        ' Top�T���v����
    BOTRESSMP As Integer        ' Bot�T���v����
    TOPIND As String        '��ԋ敪
    BOTIND As String        '��ԋ敪
    ' Oi����
    TOPOI   As Double       ' Top�T���v�����ɑΉ����錟���Ώےl��TopOi
    BOTOI   As Double       ' Top�T���v�����ɑΉ����錟���Ώےl��BotOi
    TOPOISMP As Long        ' Top�T���v����     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    BOTOISMP As Long        ' Bot�T���v����     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    ' Cs����
    BOTCS   As Double       ' Cs�����l��Cs
    BOTCSSMP As Long        ' Bot�T���v����     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    ' ���H�����o������
    PRCMCN  As String * 1   ' ����@
    SEED    As String * 4   ' SEED
    ' ������H����
    NCHPOS As String * 2    ' �m�b�`�ʒu
    DMTOP1 As Double        ' ���a
    DMTOP2 As Double        ' ���a
    DMTAIL1 As Double       ' ���a
    DMTAIL2 As Double       ' ���a
    ' �������
    MAGTYPE As String * 2   ' ����^�C�v�������@
    ' ���i�d�lSXL�ް��P
    TYPE As String * 1      ' �i�r�w�^�C�v���^�C�v
    HSXCDIR As String * 3   ' �i�r�w�����ʕ��ʁ�����
    ' ���グ��������
    DPNTCLS As String * 7   ' �h�[�p���g��ށ��h�[�p���g�@�����ԍ��O7��+"00"�����グ�w����
    ' �u���b�N�Ǘ�
    TOPPOS As Integer       ' �������J�n�ʒu��Top����
    BOTPOS As Integer       ' �������J�n�ʒu�{������Bot����
    UPDDATE As Date         ' �X�V���t��������t
    ' �����T���v���Ǘ�
    TOPSMPLNO As Long       ' Top�T���v����     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    BOTSMPLNO As Long       ' Bot�T���v����     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    ' GD����
    TOPDVD2 As Integer      ' ���茋�� DVD2��DVD2(Top)
    BOTDVD2 As Integer      ' ���茋�� DVD2��DVD2(Bot)
    TOPDVD2SMP As Long        ' Top�T���v����   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    BOTDVD2SMP As Long        ' Bot�T���v����   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    ' OSF����
    ' HTPRC = �M�������@�KKSP = �������ב���ʒu�KKSET = �������ב������ + �I��ET��
    ' ���A�����d�l��T���A���̎d�l�̔ԍ�(OSF1�Ƃ�OSF2)�����߁A�Ή�����ꏊ�֊i�[����B
    TOPOSF(3) As Double     ' �v�Z���� Max��OSF(Top)
    BOTOSF(3) As Double     ' �v�Z���� Max��OSF(Bot)
    TOPOSFSMP(3) As Long        ' Top�T���v����     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    BOTOSFSMP(3) As Long        ' Bot�T���v����     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    ' BMD����
    ' HTPRC = �M�������@�KKSP = �������ב���ʒu�KKSET = �������ב������ + �I��ET��
    ' ���A�����d�l��T���A���̎d�l�̔ԍ�(OSF1�Ƃ�OSF2)�����߁A�Ή�����ꏊ�֊i�[����B
    TOPBMD(2) As Double     ' Max��OSF(Top)
    BOTBMD(2) As Double     ' Max��OSF(Bot)
    TOPBMDSMP(2) As Long        ' Top�T���v����     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    BOTBMDSMP(2) As Long        ' Bot�T���v����     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    ' ���C�t�^�C��
    TOPLT As Integer        ' �v�Z���ʁ����C�t�^�C��(Top)
    BOTLT As Integer        ' �v�Z���ʁ����C�t�^�C��(Bot)
    TOPLTSMP As Long        ' Top�T���v����         Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    BOTLTSMP As Long        ' Bot�T���v����         Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    '---- ADD [���������V�X�e���Ή�] 2004/10/22 TCS)R.Kawaguchi START ----
    ' ����p�^�[��
    HIKIAGEPTRN As String   '����p�^�[�����O���[�v�w�����{�ŏI����{���{�O���[�v�����㏇��
'---- ADD [���������V�X�e���Ή�] 2004/10/22 TCS)R.Kawaguchi END ----
    '�z�[���h�f�[�^�ǉ�    2006/03
    HOLDKT As String
    BIKOU As String
    HLDCMNT As String
    HLDTRCLS As String
    HLDCAUSE As String
    AGRSTATUS           As String           ' ���F�m�F�敪      add SETkimizuka
    STOP                As String           ' ��~      add SETkimizuka
    CAUSE               As String           ' ��~���R  add SETkimizuka
    PRINTNO             As String           ' ��s�]��  add SETkimizuka
End Type

'---- ADD [���������V�X�e���Ή�] 2004/10/22 TCS)R.Kawaguchi START ----
Public Type type_DBDRV_xsdc1
    CRYNUM As String * 12   ' �����ԍ�
    HIKIAGEPTRN As String   ' ����p�^�[��
End Type
'---- ADD [���������V�X�e���Ή�] 2004/10/22 TCS)R.Kawaguchi END ----

' �i�ԓ��͎�
Public Function DBDRV_scmzc_fcmgc001f_Hinban(HINBAN As String, _
                                             Zyouken As type_DBDRV_scmzc_fcmgc001f_Hinban) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim fullHinban As tFullHinban

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_scmzc_fcmgc001f_Hinban"

    DBDRV_scmzc_fcmgc001f_Hinban = FUNCTION_RETURN_SUCCESS

    '12���i�Ԃ����߂�
    If GetLastHinban(HINBAN, fullHinban) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmgc001f_Hinban = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '�����������擾
    sql = "select "
    sql = sql & " T.MAGTYPE, "          ' ����^�C�v
    sql = sql & " V.E018HSXD1MIN, "     ' �i�r�w���a�P����
    sql = sql & " V.E018HSXD1MAX, "     ' �i�r�w���a�P���
    sql = sql & " V.E018HSXTYPE, "      ' �i�r�w�^�C�v
    sql = sql & " V.E018HSXDOP, "       ' �i�r�w�h�[�p���g
    sql = sql & " V.E018HSXCDIR, "      ' �i�r�w�����ʕ���
    sql = sql & " V.E018HSXRMIN, "      ' �i�r�w���R����
    sql = sql & " V.E018HSXRMAX, "      ' �i�r�w���R���
    sql = sql & " V.E018HSXDPDIR, "     ' �i�r�w�a�ʒu����
    sql = sql & " V.E019HSXONMIN, "     ' �i�r�w�_�f�Z�x����
    sql = sql & " V.E019HSXONMAX, "     ' �i�r�w�_�f�Z�x���
    sql = sql & " V.E019HSXCNMIN, "     ' �i�r�w�Y�f�Z�x����
    sql = sql & " V.E019HSXCNMAX, "     ' �i�r�w�Y�f�Z�x���
    sql = sql & " V.E019HSXLTMIN, "     ' �i�r�w�k�^�C������
    sql = sql & " V.E019HSXLTMAX, "     ' �i�r�w�k�^�C�����
    sql = sql & " V.E020HSXDVDMXN, "     ' �i�r�w�c�u�c�Q���   �v�e�T���v�������ύX 2003.05.20 yakimura
    sql = sql & " V.E020HSXDVDMNN, "     ' �i�r�w�c�u�c�Q����   �v�e�T���v�������ύX 2003.05.20 yakimura
    sql = sql & " V.E020HSXOF1AX, "     ' �i�r�w�n�r�e�P���Ϗ��
    sql = sql & " V.E020HSXOF1MX, "     ' �i�r�w�n�r�e�P���
    sql = sql & " V.E020HSXOF2AX, "     ' �i�r�w�n�r�e�Q���Ϗ��
    sql = sql & " V.E020HSXOF2MX, "     ' �i�r�w�n�r�e�Q���
    sql = sql & " V.E020HSXOF3AX, "     ' �i�r�w�n�r�e�R���Ϗ��
    sql = sql & " V.E020HSXOF3MX, "     ' �i�r�w�n�r�e�R���
    sql = sql & " V.E020HSXOF4AX, "     ' �i�r�w�n�r�e�S���Ϗ��
    sql = sql & " V.E020HSXOF4MX, "     ' �i�r�w�n�r�e�S���
    sql = sql & " V.E020HSXBM1AN, "     ' �i�r�w�a�l�c�P���ω���
    sql = sql & " V.E020HSXBM1AX, "     ' �i�r�w�a�l�c�P���Ϗ��
    sql = sql & " V.E020HSXBM2AN, "     ' �i�r�w�a�l�c�Q���ω���
    sql = sql & " V.E020HSXBM2AX, "     ' �i�r�w�a�l�c�Q���Ϗ��
    sql = sql & " V.E020HSXBM3AN, "     ' �i�r�w�a�l�c�R���ω���
    sql = sql & " V.E020HSXBM3AX  "     ' �i�r�w�a�l�c�R���Ϗ��
    sql = sql & " from VECME001 V, TBCMB012 T, TBCME018 S"
    With fullHinban
        sql = sql & " where V.E018HINBAN='" & .HINBAN & "' "
        sql = sql & " and V.E018MNOREVNO=" & .mnorevno & " "
        sql = sql & " and V.E018FACTORY='" & .factory & "' "
        sql = sql & " and V.E018OPECOND='" & .opecond & "' "
        sql = sql & " and S.HINBAN='" & .HINBAN & "' "
        sql = sql & " and S.MNOREVNO=" & .mnorevno & " "
        sql = sql & " and S.FACTORY='" & .factory & "' "
        sql = sql & " and S.OPECOND='" & .opecond & "' "
        sql = sql & " and trim(T.MKCONDNO)=trim(S.MCNO) "
    End With
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        DBDRV_scmzc_fcmgc001f_Hinban = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    'NULL�Ή� ----- START ----- 2003/12/09
    With Zyouken
        .MAGTYPE = rs("MAGTYPE")        ' ����^�C�v
        .HSXD1MIN = fncNullCheck(rs("E018HSXD1MIN"))  ' �i�r�w���a�P����
        .HSXD1MAX = fncNullCheck(rs("E018HSXD1MAX"))  ' �i�r�w���a�P���
        .HSXTYPE = rs("E018HSXTYPE")    ' �i�r�w�^�C�v
        .HSXDOP = rs("E018HSXDOP")      ' �i�r�w�h�[�p���g
        .HSXCDIR = rs("E018HSXCDIR")    ' �i�r�w�����ʕ���
        .HSXRMIN = fncNullCheck(rs("E018HSXRMIN"))    ' �i�r�w���R����
        .HSXRMAX = fncNullCheck(rs("E018HSXRMAX"))    ' �i�r�w���R���
        .HSXDPDIR = rs("E018HSXDPDIR")  ' �i�r�w�a�ʒu����
        .HSXONMIN = fncNullCheck(rs("E019HSXONMIN"))  ' �i�r�w�_�f�Z�x���
        .HSXONMAX = fncNullCheck(rs("E019HSXONMAX"))  ' �i�r�w�Y�f�Z�x����
        .HSXCNMIN = fncNullCheck(rs("E019HSXCNMIN"))  ' �i�r�w�Y�f�Z�x���
        .HSXCNMAX = fncNullCheck(rs("E019HSXCNMAX"))  ' �i�r�w�a�ʒu����
        .HSXLTMIN = fncNullCheck(rs("E019HSXLTMIN"))  ' �i�r�w�k�^�C������
        .HSXLTMAX = fncNullCheck(rs("E019HSXLTMAX"))  ' �i�r�w�k�^�C�����
        .HSXDVDMX = fncNullCheck(rs("E020HSXDVDMXN"))  ' �i�r�w�c�u�c�Q���   �v�e�T���v�������ύX 2003.05.20 yakimura
        .HSXDVDMN = fncNullCheck(rs("E020HSXDVDMNN"))  ' �i�r�w�c�u�c�Q����   �v�e�T���v�������ύX 2003.05.20 yakimura
        .HSXOS1AX = fncNullCheck(rs("E020HSXOF1AX"))  ' �i�r�w�n�r�e�P���Ϗ��
        .HSXOS1MX = fncNullCheck(rs("E020HSXOF1MX"))  ' �i�r�w�n�r�e�P���
        .HSXOS2AX = fncNullCheck(rs("E020HSXOF2AX"))  ' �i�r�w�n�r�e�Q���Ϗ��
        .HSXOS2MX = fncNullCheck(rs("E020HSXOF2MX"))  ' �i�r�w�n�r�e�Q���
        .HSXOS3AX = fncNullCheck(rs("E020HSXOF3AX"))  ' �i�r�w�n�r�e�R���Ϗ��
        .HSXOS3MX = fncNullCheck(rs("E020HSXOF3MX"))  ' �i�r�w�n�r�e�R���
        .HSXOS4AX = fncNullCheck(rs("E020HSXOF4AX"))  ' �i�r�w�n�r�e�S���Ϗ��
        .HSXOS4MX = fncNullCheck(rs("E020HSXOF4MX"))  ' �i�r�w�n�r�e�S��� HSXOS4AX��HSXOS4MX�ɕύX 2003/12/09
        .HSXBM1AN = fncNullCheck(rs("E020HSXBM1AN"))  ' �i�r�w�a�l�c�P���ω���
        .HSXBM1AX = fncNullCheck(rs("E020HSXBM1AX"))  ' �i�r�w�a�l�c�P���Ϗ��
        .HSXBM2AN = fncNullCheck(rs("E020HSXBM2AN"))  ' �i�r�w�a�l�c�Q���ω���
        .HSXBM2AX = fncNullCheck(rs("E020HSXBM2AX"))  ' �i�r�w�a�l�c�Q���Ϗ��
        .HSXBM3AN = fncNullCheck(rs("E020HSXBM3AN"))  ' �i�r�w�a�l�c�R���ω���
        .HSXBM3AX = fncNullCheck(rs("E020HSXBM3AX"))  ' �i�r�w�a�l�c�R���Ϗ��
    End With
    'NULL�Ή� -----  END  ----- 2003/12/09
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

' �����\��
Public Function DBDRV_scmzc_fcmgc001f_INITDISP(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim c0 As Integer
    Dim c1 As Integer
    Dim recCount As Integer
    Dim temp0 As String
    Dim CodeData As String
    Dim MaxRec As Integer
    Dim i As Integer
    Dim BlockIdBuf  As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_scmzc_fcmgc001f_INITDISP"

    DBDRV_scmzc_fcmgc001f_INITDISP = FUNCTION_RETURN_SUCCESS

    '�����\���u���b�NID���擾
    'sql = "select distinct BLOCKID, CRYNUM, INGOTPOS, LENGTH, REALLEN, UPDDATE "    ' �u���b�NID
    'sql = sql & "from TBCME040 where "
    'sql = sql & "DELCLS = '0' and HOLDCLS = '0' "
    'sql = sql & "and RSTATCLS = 'G'"
    sql = " SELECT DISTINCT CRYNUMC2, INPOSC2, GNLC2, KDAYC2, XTALC2,"
    sql = sql & "H2.PGID, H2.TYPE, H2.DPNTCLS,"
    sql = sql & "HINBCA, C.PALTNUM, C.BDCODE, PUPTNC1, "
    sql = sql & "A.REPSMPLIDCS AS ATOP,  "
    sql = sql & "B.REPSMPLIDCS AS BBOT "
    sql = sql & ",HOLDBC2, HOLDCC2, HOLDKTC2 "
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/26
    ' ������~���ڒǉ� add SETkimizuka Start  09/03/18
'    sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUS "
'    sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
'    sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSE "
'    sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNO "
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUS "
    sql = sql & " , NVL(TO_CHAR(Y4.STOPY4),' ') as STOP "
    sql = sql & " , DECODE(TRIM(Y4.CAUSEY4),NULL,' ',TRIM(Y4.CAUSEY4) || ':' || NAMEJA9) as CAUSE "
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNO "
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
   ' ������~���ڒǉ� add SETkimizuka End    09/03/18
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/26
    sql = sql & "FROM XSDC2, TBCMH002 H2, XSDCA, XSDCS A, XSDCS B, TBCMG007 C, XSDC1 "
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/26
    sql = sql & "    ,XODY3,XODY4 Y4,KODA9  "
    '' ������~���ڒǉ� add SETkimizuka Start  09/03/18
    'sql = sql & "    ,(SELECT XTALNOY3 as XTALNO,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
    'sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
    'sql = sql & "      FROM XODY3  "
    'sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND STOPY4 <> '2'  AND WKKTY4 in " & CreateWkktSQL(WATCH_PROCCD) & ") "
    'sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
    'sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
    'sql = sql & "                FROM XODY3,XODY4,XODY5 "
    'sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
    'sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
    'sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
    'sql = sql & "      WHERE  "
    'sql = sql & "       LIVKY3    = '0' "
    'sql = sql & "       GROUP BY XTALNOY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9) Y4 "
    ' ������~���ڒǉ� add SETkimizuka Start  09/03/18
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/26
    sql = sql & "WHERE "
    sql = sql & "LIVKC2  <> '1' AND "
    sql = sql & "GNWKNTC2 = 'CB320' AND "
    sql = sql & "SUBSTR(CRYNUMC2,1,9) = H2.UPINDNO AND "
    sql = sql & "CRYNUMC2 = CRYNUMCA AND "
    sql = sql & "INPOSC2 = INPOSCA AND "
    sql = sql & "LIVKCA <> '1' AND "
    sql = sql & "CRYNUMC2 = A.CRYNUMCS AND "
    sql = sql & "CRYNUMC2 = B.CRYNUMCS AND "
    sql = sql & "A.SMPKBNCS = 'T' AND "
    sql = sql & "B.SMPKBNCS = 'B' AND "
    sql = sql & "CRYNUMC2 = C.CRYNUM AND "
    sql = sql & "C.TRANCNT=(SELECT MAX(TRANCNT) FROM TBCMG007 WHERE CRYNUM=C.CRYNUM) AND "
    sql = sql & "XTALC2 = XTALC1 "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/26
    'sql = sql & "   AND CRYNUMCA     = Y4.XTALNO(+) "            'add 09/03/18 SETkimizuka
    sql = sql & " AND CRYNUMCA = XTALNOY3(+) "
    sql = sql & " AND LIVKY3(+) = '0' "
    sql = sql & " AND LIVKY4(+) = '0' "
    sql = sql & " AND XTALNOY3 = XTALNOY4(+) "
    sql = sql & " AND RCNTY3 = RCNTY4(+) "
    sql = sql & " AND SYSCA9(+) = 'X' AND SHUCA9(+) = '30' AND CAUSEY4 = CODEA9(+) "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/26
    
    Debug.Print sql
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount = 0 Then
        DBDRV_scmzc_fcmgc001f_INITDISP = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '������~���ڒǉ��ɔ����C�� upd 09/04/23 Start SETkimizuka
    '���o���ʂ��i�[����
    ReDim Res(0) As type_DBDRV_scmzc_fcmgc001f_Kensaku
'    DataInit Res()
    TotalLength = 0
    For c0 = 1 To recCount
        If rs("CRYNUMC2") <> BlockIdBuf Then
            i = i + 1
            ReDim Preserve Res(i)
            DataInit2 Res(), i
            
            With Res(i)
                .BLOCKID = rs("CRYNUMC2")    '
                .LENGTH = rs("GNLC2")
                .CRYNUM = rs("XTALC2")
                .INGOTPOS = rs("INPOSC2")
                .UPDDATE = rs("KDAYC2")
                .TOPPOS = .INGOTPOS
                .BOTPOS = .INGOTPOS + .LENGTH
                TotalLength = TotalLength + .LENGTH
            
                .DPNTCLS = rs("DPNTCLS")
                .TYPE = rs("TYPE")
                .PGID = rs("PGID")
                
                .HINBAN = rs("HINBCA")
            
                .TOPSMPLNO = rs("ATOP")
                .BOTSMPLNO = rs("BBOT")
                .BDCODE = rs("BDCODE")
                .PALTNUM = rs("PALTNUM")
            
                .HIKIAGEPTRN = rs("PUPTNC1")
                
                If IsNull(rs("HOLDBC2")) = False Then .HLDTRCLS = rs("HOLDBC2")    '2006/03
                If IsNull(rs("HOLDKTC2")) = False Then .HOLDKT = rs("HOLDKTC2")  '2006/03
                If IsNull(rs("HOLDCC2")) = False Then .HLDCAUSE = rs("HOLDCC2")    '2006/03
                
                ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/26
                '.AGRSTATUS = rs("AGRSTATUS")
                '.STOP = rs("STOP")
                'If Trim(rs("CAUSE")) <> "" And InStr(Res(i).CAUSE, rs("CAUSE")) = 0 Then
                '    Res(i).CAUSE = Res(i).CAUSE & rs("CAUSE") & vbTab
                'End If
                If rs("STOP") <> "2" And rs("WKKTY4") = "CB320" Then
                    .AGRSTATUS = rs("AGRSTATUS")
                    .STOP = rs("STOP")
                    If Trim(rs("CAUSE")) <> "" And InStr(Res(i).CAUSE, rs("CAUSE")) = 0 Then
                        Res(i).CAUSE = Res(i).CAUSE & rs("CAUSE") & vbTab
                    End If
                End If
                ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/26
                If Trim(rs("PRINTNO")) <> "" And InStr(Res(i).PRINTNO, rs("PRINTNO")) = 0 Then
                    Res(i).PRINTNO = Res(i).PRINTNO & rs("PRINTNO") & vbTab
                End If
                
                BlockIdBuf = rs("CRYNUMC2")
            End With
        Else
            ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/26
            'If Trim(rs("CAUSE")) <> "" And InStr(Res(i).CAUSE, rs("CAUSE")) = 0 Then
            '    Res(i).CAUSE = Res(i).CAUSE & rs("CAUSE") & vbTab
            'End If
            If rs("STOP") <> "2" And rs("WKKTY4") = "CB320" Then
                If Trim(Res(i).AGRSTATUS) = "" Or rs("AGRSTATUS") < Res(i).AGRSTATUS Then
                    Res(i).AGRSTATUS = rs("AGRSTATUS")
                    Res(i).STOP = rs("STOP")
                End If
                If Trim(rs("CAUSE")) <> "" And InStr(Res(i).CAUSE, rs("CAUSE")) = 0 Then
                    Res(i).CAUSE = Res(i).CAUSE & rs("CAUSE") & vbTab
                End If
            End If
            ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/26
            If Trim(rs("PRINTNO")) <> "" And InStr(Res(i).PRINTNO, rs("PRINTNO")) = 0 Then
                Res(i).PRINTNO = Res(i).PRINTNO & rs("PRINTNO") & vbTab
            End If
        End If
        
        rs.MoveNext
    Next
    rs.Close
'    '���o���ʂ��i�[����
'    ReDim Res(recCount) As type_DBDRV_scmzc_fcmgc001f_Kensaku
'    DataInit Res()
'    TotalLength = 0
'    For c0 = 1 To recCount
'        With Res(c0)
'            '.BLOCKID = rs("BLOCKID")    ' �u���b�NID
'            .BLOCKID = rs("CRYNUMC2")    '
'            '.LENGTH = rs("REALLEN")
'            .LENGTH = rs("GNLC2")
'            '.CRYNUM = rs("CRYNUM")
'            .CRYNUM = rs("XTALC2")
'            '.INGOTPOS = rs("INGOTPOS")
'            .INGOTPOS = rs("INPOSC2")
'            '.UPDDATE = rs("UPDDATE")
'            .UPDDATE = rs("KDAYC2")
'            .TOPPOS = .INGOTPOS
'            .BOTPOS = .INGOTPOS + .LENGTH
'            TotalLength = TotalLength + .LENGTH
'
'            .DPNTCLS = rs("DPNTCLS")
'            .TYPE = rs("TYPE")
'            .PGID = rs("PGID")
'
'            .hinban = rs("HINBCA")
'
'            .TOPSMPLNO = rs("ATOP")
'            .BOTSMPLNO = rs("BBOT")
'            .BDCODE = rs("BDCODE")
'            .PALTNUM = rs("PALTNUM")
'
'            .HIKIAGEPTRN = rs("PUPTNC1")
'
'            If IsNull(rs("HOLDBC2")) = False Then .HLDTRCLS = rs("HOLDBC2")    '2006/03
'            If IsNull(rs("HOLDKTC2")) = False Then .HOLDKT = rs("HOLDKTC2")  '2006/03
'            If IsNull(rs("HOLDCC2")) = False Then .HLDCAUSE = rs("HOLDCC2")    '2006/03
'        End With
'
'        rs.MoveNext
'    Next
'    rs.Close
    '������~���ڒǉ��ɔ����C�� upd 09/04/23 End SETkimizuka
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

Public Sub DataInit(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)

    Dim c0 As Integer
    Dim c1 As Integer
    Dim recCount As Integer

    recCount = UBound(Res())
    For c0 = 1 To recCount
        With Res(c0)
            .BLOCKID = ""           ' �u���b�NID
            .LENGTH = -1            ' �������u���b�N��
            '�i�ԊǗ�
            .HINBAN = ""            ' Top�i�ԁ��i��
            '�N���X�^���J�^���O�������
            .PALTNUM = ""           ' �p���b�g�ԍ�
            .BDCODE = ""            ' �s�Ǘ��R�R�[�h���i���敪
            '�������
            .DIAMETER = -1          ' ���a
            .PGID = ""              ' �o�f�|�h�c
            '������R����
            .TOPRES = -1            ' Top�T���v�����ɑΉ����錟���Ώےl��Top�����
            .BOTRES = -1            ' Bot�T���v�����ɑΉ����錟���Ώےl��Bot�����
            'Oi����
            .TOPOI = -1             ' Top�T���v�����ɑΉ����錟���Ώےl��TopOi
            .BOTOI = -1             ' Bot�T���v�����ɑΉ����錟���Ώےl��BotOi
            'Cs����
            .BOTCS = -1             ' Cs�����l��Cs
            '������H����
            .NCHPOS = ""            ' �m�b�`�ʒu
            '�������
            .MAGTYPE = ""           ' ����^�C�v�������@
            '���i�d�lSXL�ް��P
            .TYPE = ""              ' �i�r�w�^�C�v���^�C�v
            .HSXCDIR = ""           ' �i�r�w�����ʕ��ʁ�����
            ' ���グ��������
            .DPNTCLS = ""           ' �h�[�p���g��ށ��h�[�p���g�@�����ԍ��O7��+"00"�����グ�w����
            '�u���b�N�Ǘ�
            .TOPPOS = -1            ' �������J�n�ʒu��Top����
            .BOTPOS = -1            ' �������J�n�ʒu�{������Bot����
            .UPDDATE = -1           ' �X�V���t��������t
            '�����T���v���Ǘ�
            .TOPSMPLNO = -1         ' �T���v����
            .BOTSMPLNO = -1         ' �T���v����
            'GD����
            .TOPDVD2 = -1           ' ���茋�� DVD2��DVD2(Top)
            .BOTDVD2 = -1           ' ���茋�� DVD2��DVD2(Bot)
            'OSF����
            'HTPRC = �M�������@�KKSP = �������ב���ʒu�KKSET = �������ב������ + �I��ET��
            '���A�����d�l��T���A���̎d�l�̔ԍ�(OSF1�Ƃ�OSF2)�����߁A�Ή�����ꏊ�֊i�[����B
            For c1 = 0 To 3
                .TOPOSF(c1) = -1    ' �v�Z���� Max��OSF(Top)
                .BOTOSF(c1) = -1    ' �v�Z���� Max��OSF(Bot)
            Next
            'BMD����
            'HTPRC = �M�������@�KKSP = �������ב���ʒu�KKSET = �������ב������ + �I��ET��
            '���A�����d�l��T���A���̎d�l�̔ԍ�(OSF1�Ƃ�OSF2)�����߁A�Ή�����ꏊ�֊i�[����B
            For c1 = 0 To 2
                .TOPBMD(c1) = -1    ' Max��OSF(Top)
                .BOTBMD(c1) = -1    ' Max��OSF(Bot)
            Next
            '���C�t�^�C��
            .TOPLT = -1             ' �v�Z���ʁ����C�t�^�C��(Top)
            .BOTLT = -1             ' �v�Z���ʁ����C�t�^�C��(Bot)
        End With
    Next

End Sub

Public Sub DataInit2(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku, Initnum As Integer)

    Dim c1 As Integer

        With Res(Initnum)
            .BLOCKID = ""           ' �u���b�NID
            .LENGTH = -1            ' �������u���b�N��
            '�i�ԊǗ�
            .HINBAN = ""            ' Top�i�ԁ��i��
            '�N���X�^���J�^���O�������
            .PALTNUM = ""           ' �p���b�g�ԍ�
            .BDCODE = ""            ' �s�Ǘ��R�R�[�h���i���敪
            '�������
            .DIAMETER = -1          ' ���a
            .PGID = ""              ' �o�f�|�h�c
            '������R����
            .TOPRES = -1            ' Top�T���v�����ɑΉ����錟���Ώےl��Top�����
            .BOTRES = -1            ' Bot�T���v�����ɑΉ����錟���Ώےl��Bot�����
            'Oi����
            .TOPOI = -1             ' Top�T���v�����ɑΉ����錟���Ώےl��TopOi
            .BOTOI = -1             ' Bot�T���v�����ɑΉ����錟���Ώےl��BotOi
            'Cs����
            .BOTCS = -1             ' Cs�����l��Cs
            '������H����
            .NCHPOS = ""            ' �m�b�`�ʒu
            '�������
            .MAGTYPE = ""           ' ����^�C�v�������@
            '���i�d�lSXL�ް��P
            .TYPE = ""              ' �i�r�w�^�C�v���^�C�v
            .HSXCDIR = ""           ' �i�r�w�����ʕ��ʁ�����
            ' ���グ��������
            .DPNTCLS = ""           ' �h�[�p���g��ށ��h�[�p���g�@�����ԍ��O7��+"00"�����グ�w����
            '�u���b�N�Ǘ�
            .TOPPOS = -1            ' �������J�n�ʒu��Top����
            .BOTPOS = -1            ' �������J�n�ʒu�{������Bot����
            .UPDDATE = -1           ' �X�V���t��������t
            '�����T���v���Ǘ�
            .TOPSMPLNO = -1         ' �T���v����
            .BOTSMPLNO = -1         ' �T���v����
            'GD����
            .TOPDVD2 = -1           ' ���茋�� DVD2��DVD2(Top)
            .BOTDVD2 = -1           ' ���茋�� DVD2��DVD2(Bot)
            'OSF����
            'HTPRC = �M�������@�KKSP = �������ב���ʒu�KKSET = �������ב������ + �I��ET��
            '���A�����d�l��T���A���̎d�l�̔ԍ�(OSF1�Ƃ�OSF2)�����߁A�Ή�����ꏊ�֊i�[����B
            For c1 = 0 To 3
                .TOPOSF(c1) = -1    ' �v�Z���� Max��OSF(Top)
                .BOTOSF(c1) = -1    ' �v�Z���� Max��OSF(Bot)
            Next
            'BMD����
            'HTPRC = �M�������@�KKSP = �������ב���ʒu�KKSET = �������ב������ + �I��ET��
            '���A�����d�l��T���A���̎d�l�̔ԍ�(OSF1�Ƃ�OSF2)�����߁A�Ή�����ꏊ�֊i�[����B
            For c1 = 0 To 2
                .TOPBMD(c1) = -1    ' Max��OSF(Top)
                .BOTBMD(c1) = -1    ' Max��OSF(Bot)
            Next
            '���C�t�^�C��
            .TOPLT = -1             ' �v�Z���ʁ����C�t�^�C��(Top)
            .BOTLT = -1             ' �v�Z���ʁ����C�t�^�C��(Bot)
        End With

End Sub

' �i�ԊǗ�
Public Sub GETTBCME041(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim buf() As type_DBDRV_cmgc001f1
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    '�����ԍ��ƌ����ʒu����i�Ԃ����߂�
    sql = "select blk.CRYNUM, blk.INGOTPOS, hin.HINBAN, hin.REVNUM, hin.FACTORY, hin.OPECOND "
    sql = sql & "from TBCME040 blk, TBCME041 hin "
    sql = sql & "where (blk.CRYNUM = hin.CRYNUM) "
    sql = sql & "and ((blk.INGOTPOS >= hin.INGOTPOS) and (blk.INGOTPOS < (hin.INGOTPOS + hin.LENGTH))) "

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
        ReDim buf(recCount) As type_DBDRV_cmgc001f1
        For c0 = 1 To recCount
            buf(c0).CRYNUM = rs("CRYNUM")
            buf(c0).factory = rs("FACTORY")
            buf(c0).HINBAN = rs("HINBAN")
            buf(c0).INGOTPOS = rs("INGOTPOS")
            buf(c0).opecond = rs("OPECOND")
            buf(c0).REVNUM = rs("REVNUM")
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (Res(c0).INGOTPOS = buf(c1).INGOTPOS) Then
                    Res(c0).HINBAN = buf(c1).HINBAN
                    Res(c0).REVNUM = buf(c1).REVNUM     ' ���i�ԍ������ԍ�
                    Res(c0).factory = buf(c1).factory   ' �H��
                    Res(c0).opecond = buf(c1).opecond   ' ���Ə���
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).HINBAN = " "
                Res(c0).REVNUM = -1     ' ���i�ԍ������ԍ�
                Res(c0).factory = " "   ' �H��
                Res(c0).opecond = " "   ' ���Ə���
            End If
        Next
    End If
    rs.Close
    On Error GoTo 0

End Sub

' �N���X�^���J�^���O�������
Public Sub GETTBCMG007(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim buf() As type_DBDRV_cmgc001f2
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    '�u���b�NID����p���b�g�ԍ��A�s�Ǘ��R�R�[�h(�i���敪)�����߂�
    MaxRec = UBound(Res())
    For c0 = 1 To MaxRec
        sql = "select CRYNUM, PALTNUM, BDCODE from TBCMG007 G"
        sql = sql & " where TRANCNT=(select max(TRANCNT) from TBCMG007 where CRYNUM=G.CRYNUM)"
        sql = sql & " and crynum = '" & Res(c0).BLOCKID & "' "
        DoEvents
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        recCount = rs.RecordCount
        If recCount <> 0 Then
            'ReDim buf(recCount) As type_DBDRV_cmgc001f2
            'For c0 = 1 To recCount
            '    buf(c0).BDCODE = rs("BDCODE")
            '    buf(c0).BLOCKID = rs("CRYNUM")
            '    buf(c0).PALTNUM = rs("PALTNUM")
            '    rs.MoveNext
            'Next
            'rs.Close
            'MaxRec = UBound(Res())
            'For c0 = 1 To MaxRec
            '    For c1 = 1 To recCount
            '        If (Res(c0).BLOCKID = buf(c1).BLOCKID) Then
            '            Res(c0).BDCODE = buf(c1).BDCODE
            '            Res(c0).PALTNUM = buf(c1).PALTNUM
            '            OKFlag = True
            '            Exit For
            '        End If
            '    Next
            '    If Not OKFlag Then
            '        Res(c0).BDCODE = " "
            '        Res(c0).PALTNUM = " "
            '    End If
            'Next
            Res(c0).BDCODE = rs("BDCODE")
            Res(c0).PALTNUM = rs("PALTNUM")
        End If
    Next
    rs.Close
    On Error GoTo 0

End Sub

' �������
Public Sub GETTBCME037(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim buf() As type_DBDRV_cmgc001f3
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    '���a�APG-ID�����߂�
    sql = "select CRYNUM, DIAMETER, PGID from TBCME037"

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
        ReDim buf(recCount) As type_DBDRV_cmgc001f3
        For c0 = 1 To recCount
            buf(c0).CRYNUM = rs("CRYNUM")
            buf(c0).PGID = rs("PGID")
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
                If (Res(c0).CRYNUM = buf(c1).CRYNUM) Then
                    Res(c0).PGID = buf(c1).PGID
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).PGID = " "
            End If
        Next
    End If
    rs.Close
    On Error GoTo 0

End Sub

' �������
Public Sub GETTBCMB012(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'    Dim sMKCONDNO As String * 12         ' �������No.
'
'    '�G���[�n���h���̐ݒ�
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '�����ԍ����琻����������߂�
'        sql1 = "select trim(PRODCOND) from TBCME037 where "
'        sql1 = sql1 & "CRYNUM = '" & Res(c0).CRYNUM & "'"
'
'        '����������玥��^�C�v(�����@)�����߂�
'        sql = "select MAGTYPE from TBCMB012 where "
'        sql = sql & "trim(MKCONDNO) = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).MAGTYPE = " "
'        Else
'            Res(c0).MAGTYPE = rs("MAGTYPE")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    CRYNUM As String * 12         ' �����ԍ�
'    MAGTYPE As String * 2   '����^�C�v�������@
'End Type
    Dim buf() As type_DBDRV_cmgc001f5
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    '�����ԍ����琻����������߂�
    '����������玥��^�C�v(�����@)�����߂�
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     XTALC2,"
    sql = sql & "     A.MAGTYPE "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, TBCMB012 A , TBCME037 B "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     XTALC2 = B.CRYNUM AND"
    sql = sql & "     TRIM(A.MKCONDNO) = TRIM(B.PRODCOND) "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                Res(c0).MAGTYPE = rs("MAGTYPE")
                Exit For
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0
End Sub

' ���i�d�lSXL�ް��P
Public Sub GETTBCME018(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim buf() As type_DBDRV_cmgc001f6
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    '�i�Ԃ���iSX�^�C�v�A�i�r�w�����ʕ��ʂ����߂�
    sql = "select HINBAN, HSXTYPE, HSXCDIR from TBCME018"

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
        ReDim buf(recCount) As type_DBDRV_cmgc001f6
        For c0 = 1 To recCount
            buf(c0).HINBAN = rs("HINBAN")
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
                If (Res(c0).HINBAN = buf(c1).HINBAN) Then
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).TYPE = " "
            End If
        Next
    End If
    rs.Close
    On Error GoTo 0

End Sub

' ���グ��������
Public Sub GETTBCMH002(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    '�G���[�n���h���̐ݒ�
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '�h�[�p���g���(�h�[�p���g)�����߂�
'        sql = "select DPNTCLS from TBCMH002 where "
'        sql = sql & "UPINDNO = '" & Left(Res(c0).CRYNUM, 7) & "00' "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).DPNTCLS = " "
'        Else
'            Res(c0).DPNTCLS = rs("DPNTCLS")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    DPNTCLS As String * 7   '�h�[�p���g��ށ��h�[�p���g�@�����ԍ��O7��+"00"�����グ�w��No.
'    CRYNUM As String * 12         ' �����ԍ�
'End Type
    Dim buf() As type_DBDRV_cmgc001f7
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    '�h�[�p���g���(�h�[�p���g)�����߂�
    sql = "select UPINDNO, DPNTCLS, TYPE from TBCMH002"
    'sql = sql & "UPINDNO = '" & Left(Res(c0).CRYNUM, 7) & "00' "

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
        ReDim buf(recCount) As type_DBDRV_cmgc001f7
        For c0 = 1 To recCount
            buf(c0).CRYNUM = rs("UPINDNO")
            buf(c0).DPNTCLS = rs("DPNTCLS")
            buf(c0).TYPE = rs("TYPE") '2002/04/25 S.Sano
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
'2004.09.10 Y.K ���`���[�W�w��No�X���Ή�
'                If (Left(Res(c0).CRYNUM, 8) & "0" = Trim(buf(c1).CRYNUM)) Then
                If (Left(Res(c0).CRYNUM, 9) = Trim(buf(c1).CRYNUM)) Then
                    Res(c0).DPNTCLS = buf(c1).DPNTCLS
                    Res(c0).TYPE = buf(c1).TYPE '2002/04/25 S.Sano
                    OKFlag = True
                    Exit For
                End If
'2004.09.10 Y.K ���`���[�W�w��No�X���Ή�
'�c�ʈ����ł��\���������ꍇ�̓��W�b�N�ƂȂ�i�T���v���j
'''                If (Left(Res(c0).CRYNUM, 8) = Left(Trim(buf(c1).CRYNUM), 8)) Then
'''                    If (IsNumeric(Mid(Res(c0).CRYNUM, 9, 1)) = True) Then
'''                        If (Mid(Res(c0).CRYNUM, 9, 1) = Mid(Trim(buf(c1).CRYNUM), 9, 1)) Then
'''                            Res(c0).DPNTCLS = buf(c1).DPNTCLS
'''                            Res(c0).TYPE = buf(c1).TYPE '2002/04/25 S.Sano
'''                            OKFlag = True
'''                            Exit For
'''                        End If
'''                    Else
'''                        If ("A" = Mid(Trim(buf(c1).CRYNUM), 9, 1)) Then
'''                            Res(c0).DPNTCLS = buf(c1).DPNTCLS
'''                            Res(c0).TYPE = buf(c1).TYPE '2002/04/25 S.Sano
'''                            OKFlag = True
'''                            Exit For
'''                        End If
'''                    End If
'''                End If
            Next
            If Not OKFlag Then
                Res(c0).DIAMETER = " "
            End If
        Next
    End If
    rs.Close
    On Error GoTo 0

End Sub

' �����T���v���Ǘ�
Public Sub GETTBCME043(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    '�G���[�n���h���̐ݒ�
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '�T���v��No�����߂�
'        sql = "select SMPLNO, SMPKBN from TBCME043 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and INGOTPOS = '" & Res(c0).INGOTPOS & "' "
'        sql = sql & "and ((SMPKBN = 'T') or "
'        sql = sql & "((SMPKBN = 'B') and "
'        sql = sql & "(CRYINDRS='3' or CRYINDOI='3' or CRYINDB1='3' or "
'        sql = sql & "CRYINDB2='3' or CRYINDB3='3' or CRYINDL1='3' or "
'        sql = sql & "CRYINDL2='3' or CRYINDL3='3' or CRYINDL4='3' or "
'        sql = sql & "CRYINDCS='3' or CRYINDGD='3' or CRYINDT='3' or CRYINDEP='3')))"
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).TOPSMPLNO = -1
'            Res(c0).TOPSMPKBN = " "
'        Else
'            Res(c0).TOPSMPLNO = rs("SMPLNO")
'            Res(c0).TOPSMPKBN = rs("SMPKBN")
'        End If
'        rs.Close
'
'        '�T���v��No�����߂�
'        sql = "select SMPLNO, SMPKBN from TBCME043 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and INGOTPOS = '" & Res(c0).INGOTPOS + Res(c0).LENGTH & "' "
'        sql = sql & "and ((SMPKBN = 'B') or "
'        sql = sql & "((SMPKBN = 'T') and "
'        sql = sql & "(CRYINDRS='3' or CRYINDOI='3' or CRYINDB1='3' or "
'        sql = sql & "CRYINDB2='3' or CRYINDB3='3' or CRYINDL1='3' or "
'        sql = sql & "CRYINDL2='3' or CRYINDL3='3' or CRYINDL4='3' or "
'        sql = sql & "CRYINDCS='3' or CRYINDGD='3' or CRYINDT='3' or CRYINDEP='3')))"
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTSMPLNO = -1
'            Res(c0).BOTSMPKBN = " "
'        Else
'            Res(c0).BOTSMPLNO = rs("SMPLNO")
'            Res(c0).BOTSMPKBN = rs("SMPKBN")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    SMPLNO As Integer    '�T���v��No
'    CRYNUM As String * 12         ' �����ԍ�
'    INGOTPOS As Integer         ' �J�n�ʒu
'    SMPKBN As String * 1       ' �T���v���敪
'End Type
    Dim buf() As type_DBDRV_cmgc001f8
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next
    '�T���v��No�����߂�
'    sql = "select E040CRYNUM, E043INGOTPOS, E043SMPKBN, E043SMPLNO from VECME010 order by E040CRYNUM, E043INGOTPOS"
    sql = "select E040CRYNUM, E043INPOSCS, E043SMPKBNCS, E043REPSMPLIDCS from VECME010 order by E040CRYNUM, E043INPOSCS"

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
            Res(c0).TOPSMPLNO = rs("E043REPSMPLIDCS")
            Res(c0).TOPSMPKBN = rs("E043SMPKBNCS")
        ReDim buf(recCount) As type_DBDRV_cmgc001f8
        For c0 = 1 To recCount
            buf(c0).CRYNUM = rs("E040CRYNUM")
'            buf(c0).INGOTPOS = rs("E043INGOTPOS")
'            buf(c0).SMPLNO = rs("E043SMPLNO")
'            buf(c0).SMPKBN = rs("E043SMPKBN")
            buf(c0).INGOTPOS = rs("E043INPOSCS")
            buf(c0).SMPLNO = rs("E043REPSMPLIDCS")
            buf(c0).SMPKBN = rs("E043SMPKBNCS")
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And (Res(c0).INGOTPOS = buf(c1).INGOTPOS) Then
                    Res(c0).TOPSMPLNO = buf(c1).SMPLNO
                    Res(c0).TOPSMPKBN = buf(c1).SMPKBN
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).TOPSMPLNO = -1
                Res(c0).TOPSMPKBN = " "
            End If
            For c1 = 1 To recCount
                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) Then
                    Res(c0).BOTSMPLNO = buf(c1).SMPLNO
                    Res(c0).BOTSMPKBN = buf(c1).SMPKBN
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).BOTSMPLNO = -1
                Res(c0).BOTSMPKBN = " "
            End If
        Next
    End If
    rs.Close
    On Error GoTo 0
End Sub

' ������R����
Public Sub GETTBCMJ002(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    '�G���[�n���h���̐ݒ�
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '�����񐔂̂����Ƃ��傫���l�����߂�
'        sql1 = "select max(TRANCNT) from TBCMJ002 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '����f�[�^�����߂�
'        sql = "select JUDGDATA from TBCMJ002 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).TOPRES = -1
'        Else
'            Res(c0).TOPRES = rs("JUDGDATA")
'        End If
'        rs.Close
'
'        '�����񐔂̂����Ƃ��傫���l�����߂�
'        sql1 = "select max(TRANCNT) from TBCMJ002 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '����f�[�^�����߂�
'        sql = "select JUDGDATA from TBCMJ002 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTRES = -1
'        Else
'            Res(c0).BOTRES = rs("JUDGDATA")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    CRYNUM As String * 12         ' �����ԍ�
'    TRANCNT As Integer              ' ������
'    INGOTPOS As Integer         ' �J�n�ʒu
'    SMPLNO As Integer    '�T���v��No
'    SMPKBN As String * 1       ' �T���v���敪
'    JUDGDATA As Double        'Top�T���v��No.�ɑΉ����錟���Ώےl
'End Type
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    ''����f�[�^�����߂�
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, JUDGDATA from TBCMJ002 RS "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=RS.CRYNUM and POSITION=RS.POSITION and SMPKBN=RS.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("JUDGDATA")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c0 = 1 To MaxRec
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c0).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                   (Res(c0).TOPSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).TOPSMPKBN = buf(c1).SMPKBN) Then

    '                iTRANCNT = buf(c0).TRANCNT
    '                Res(c0).TOPRES = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).TOPRES = -1
    '        End If

    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                   (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then

    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).BOTRES = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).BOTRES = -1
    '        End If
    '    Next
    'End If
    'rs.Close
    
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDRSCS AS TRS,"
    sql = sql & "     CRYINDRSCS,"
    sql = sql & "     A.MEAS1 "
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ002 A"
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     CRYSMPLIDRSCS = A.SMPLNO AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ002"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO )"
    sql = sql & " UNION  "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDRSCS AS TRS,"
    sql = sql & "     CRYINDRSCS,"
    sql = sql & "     A.MEAS1 "
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ002 A"
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     CRYSMPLIDRSCS = A.SMPLNO AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ002"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") And rs("CRYINDRSCS") <> "3" Then
                If rs("SMPKBNCS") = "T" Then
                    Res(c0).TOPRES = rs("MEAS1")
                    Res(c0).TOPRESSMP = rs("TRS")
                    Exit For
                ElseIf rs("SMPKBNCS") = "B" Then
                    Res(c0).BOTRES = rs("MEAS1")
                    Res(c0).BOTRESSMP = rs("TRS")
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' Oi����
Public Sub GETTBCMJ003(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    '�G���[�n���h���̐ݒ�
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '�����񐔂̂����Ƃ��傫���l�����߂�
'        sql1 = "select max(TRANCNT) from TBCMJ003 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
''        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '����f�[�^�����߂�
'        sql = "select JUDGDATA from TBCMJ003 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
''        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).TOPOI = -1
'        Else
'            Res(c0).TOPOI = rs("JUDGDATA")
'        End If
'        rs.Close
'
'        '�����񐔂̂����Ƃ��傫���l�����߂�
'        sql1 = "select max(TRANCNT) from TBCMJ003 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
''        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '����f�[�^�����߂�
'        sql = "select JUDGDATA from TBCMJ003 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
''        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTOI = -1
'        Else
'            Res(c0).BOTOI = rs("JUDGDATA")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    CRYNUM As String * 12         ' �����ԍ�
'    TRANCNT As Integer              ' ������
'    INGOTPOS As Integer         ' �J�n�ʒu
'    SMPLNO As Integer    '�T���v��No
'    SMPKBN As String * 1       ' �T���v���敪
'    JUDGDATA As Double        'Top�T���v��No.�ɑΉ����錟���Ώےl
'End Type
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    ''����f�[�^�����߂�
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, JUDGDATA from TBCMJ003 OI "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=OI.CRYNUM and POSITION=OI.POSITION and SMPKBN=OI.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("JUDGDATA")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c0 = 1 To MaxRec
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                   (Res(c0).TOPSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).TOPSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).TOPOI = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).TOPOI = -1
    '        End If
            
   '         iTRANCNT = 0
   '         For c1 = 1 To recCount
   '             If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                   (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                    
   '                 iTRANCNT = buf(c1).TRANCNT
   '                 Res(c0).BOTOI = buf(c1).JudgData
   '                 OKFlag = True
   '             End If
   '         Next
   '         If Not OKFlag Then
   '             Res(c0).BOTOI = -1
   '         End If
   '     Next
   ' End If
   ' rs.Close
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDOICS AS TRS,"
    sql = sql & "     A.OIMEAS1"
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ003 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     CRYSMPLIDOICS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND"
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ003 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    sql = sql & " AND A.TRANCOND = 0 "  'GFA��FTIR���Z�l�擾�ُ�Ή� 2011/02/28 SETsw kubota
    sql = sql & " UNION "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDOICS AS TRS,"
    sql = sql & "     A.OIMEAS1"
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ003 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     CRYSMPLIDOICS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND"
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ003 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    sql = sql & " AND A.TRANCOND = 0 "  'GFA��FTIR���Z�l�擾�ُ�Ή� 2011/02/28 SETsw kubota
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    rs.MoveFirst
    recCount = UBound(Res)
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If rs("SMPKBNCS") = "T" Then
                    'Res(c0).TOPOI = rs("OIMEAS1")
                    If IsNull(rs("OIMEAS1")) = False Then Res(c0).TOPOI = rs("OIMEAS1") Else Res(c0).TOPOI = -1   ' OI_NULL�Ή��@2005/03/08 TUKU
                    Res(c0).TOPOISMP = rs("TRS")
                    Exit For
                Else
                    'Res(c0).BOTOI = rs("OIMEAS1")
                    If IsNull(rs("OIMEAS1")) = False Then Res(c0).BOTOI = rs("OIMEAS1") Else Res(c0).BOTOI = -1   ' OI_NULL�Ή��@2005/03/08 TUKU
                    Res(c0).BOTOISMP = rs("TRS")
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' Cs����
Public Sub GETTBCMJ004(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    '�G���[�n���h���̐ݒ�
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'
'        '�����񐔂̂����Ƃ��傫���l�����߂�
'        sql1 = "select max(TRANCNT) from TBCMJ004 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
''        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '����f�[�^�����߂�
'        sql = "select CSMEAS from TBCMJ004 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
''        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTOI = -1
'        Else
'            Res(c0).BOTOI = rs("CSMEAS")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    CRYNUM As String * 12         ' �����ԍ�
'    TRANCNT As Integer              ' ������
'    INGOTPOS As Integer         ' �J�n�ʒu
'    SMPLNO As Integer    '�T���v��No
'    SMPKBN As String * 1       ' �T���v���敪
'    JUDGDATA As Double        'Top�T���v��No.�ɑΉ����錟���Ώےl
'End Type
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    ''����f�[�^�����߂�
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, CSMEAS from TBCMJ004 CS "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=CS.CRYNUM and POSITION=CS.POSITION and SMPKBN=CS.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("CSMEAS")
    '        rs.MoveNext
    '    Next
    '    rs.Close
   '     MaxRec = UBound(Res())
    '    For c0 = 1 To MaxRec
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                   (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).BOTCS = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).BOTCS = -1
    '        End If
    '    Next
    'End If
    'rs.Close
    sql = " SELECT DISTINCT "
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDCSCS AS TRS,"
    sql = sql & "     A.CSMEAS"
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ004 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     CRYSMPLIDCSCS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND"
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ004 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                'Res(c0).BOTCS = rs("CSMEAS")
                If IsNull(rs("CSMEAS")) = False Then Res(c0).BOTCS = rs("CSMEAS") Else Res(c0).BOTCS = -1  ' OI_NULL�Ή��@2005/03/08 TUKU
                Res(c0).BOTCSSMP = rs("TRS")
                Exit For
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' OSF����
Public Sub GETTBCMJ005(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'    Dim c1 As Integer
'
'    '�G���[�n���h���̐ݒ�
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        For c1 = 0 To 3
'            '�����񐔂̂����Ƃ��傫���l�����߂�
'            sql1 = "select max(TRANCNT) from TBCMJ005 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'            sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'            sql1 = sql1 & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'
'            '����f�[�^�����߂�
'            sql = "select CALCMAX from TBCMJ005 where "
'            sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'            sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'            sql = sql & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'            sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'            DoEvents
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                Res(c0).TOPOSF(c1) = -1
'            Else
'                Res(c0).TOPOSF(c1) = rs("CALCMAX")
'            End If
'            rs.Close
'
'            '�����񐔂̂����Ƃ��傫���l�����߂�
'            sql1 = "select max(TRANCNT) from TBCMJ005 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'            sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'            sql1 = sql1 & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'
'            '����f�[�^�����߂�
'            sql = "select CALCMAX from TBCMJ005 where "
'            sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'            sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'            sql1 = sql1 & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'            sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'            DoEvents
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                Res(c0).BOTOSF(c1) = -1
'            Else
'                Res(c0).BOTOSF(c1) = rs("CALCMAX")
'            End If
'            rs.Close
'        Next
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    CRYNUM As String * 12         ' �����ԍ�
'    TRANCNT As Integer              ' ������
'    INGOTPOS As Integer         ' �J�n�ʒu
'    SMPLNO As Integer    '�T���v��No
'    SMPKBN As String * 1       ' �T���v���敪
'    JUDGDATA As Double        'Top�T���v��No.�ɑΉ����錟���Ώےl
'End Type
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim c2 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    '����f�[�^�����߂�
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, CALCMAX from TBCMJ005 J "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=J.CRYNUM and POSITION=J.POSITION and SMPKBN=J.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).TRANCOND = rs("TRANCOND")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("CALCMAX")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c2 = 0 To 3
    '        For c0 = 1 To MaxRec
    '            iTRANCNT = 0
    '            For c1 = 1 To recCount
    '                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                       (buf(c1).TRANCNT > iTRANCNT) And _
                       (buf(c0).TRANCOND = Trim(Str(c2 + 1))) And _
                       (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                       (Res(c0).TOPSMPLNO = buf(c1).SMPLNO) And _
                       (Res(c0).TOPSMPKBN = buf(c1).SMPKBN) Then
                        
   '                     iTRANCNT = buf(c1).TRANCNT
   '                     Res(c0).TOPOSF(c2) = buf(c1).JudgData
   '                     OKFlag = True
   '                 End If
   '             Next
   '             If Not OKFlag Then
   '                 Res(c0).TOPOSF(c2) = -1
   '             End If
                
   '             iTRANCNT = 0
   '             For c1 = 1 To recCount
   '                 If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                       (buf(c1).TRANCNT > iTRANCNT) And _
                       (buf(c1).TRANCOND = Trim(Str(c2 + 1))) And _
                       (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                       (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                       (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                        
    ''                    iTRANCNT = buf(c1).TRANCNT
    '                    Res(c0).BOTOSF(c2) = buf(c1).JudgData
    '                    OKFlag = True
    '                End If
    '            Next
    '            If Not OKFlag Then
    '                Res(c0).BOTOSF(c2) = -1
    '            End If
    '        Next
    '    Next
    'End If
    'rs.Close
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDL1CS AS TRS1,CRYSMPLIDL2CS AS TRS2,CRYSMPLIDL3CS AS TRS3,CRYSMPLIDL4CS AS TRS4,"
    sql = sql & "     A.CALCMAX, A.TRANCOND,A.TRANCNT"
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ005 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ005"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO  AND"
    sql = sql & "                     TRANCOND=A.TRANCOND ) AND"
    sql = sql & " (A.SMPLNO = CRYSMPLIDL1CS OR A.SMPLNO = CRYSMPLIDL2CS OR A.SMPLNO = CRYSMPLIDL3CS OR A.SMPLNO = CRYSMPLIDL4CS) "
    sql = sql & " UNION "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDL1CS AS TRS1,CRYSMPLIDL2CS AS TRS2,CRYSMPLIDL3CS AS TRS3,CRYSMPLIDL4CS AS TRS4,"
    sql = sql & "     A.CALCMAX, A.TRANCOND,A.TRANCNT"
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ005 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ005"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO  AND"
    sql = sql & "                     TRANCOND=A.TRANCOND ) AND"
    sql = sql & " (A.SMPLNO = CRYSMPLIDL1CS OR A.SMPLNO = CRYSMPLIDL2CS OR A.SMPLNO = CRYSMPLIDL3CS OR A.SMPLNO = CRYSMPLIDL4CS) "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If rs("SMPKBNCS") = "T" Then
                    Select Case rs("TRANCOND")
                        Case "1"
                            Res(c0).TOPOSF(0) = rs("CALCMAX")
                            Res(c0).TOPOSFSMP(0) = rs("TRS1")
                        Case "2"
                            Res(c0).TOPOSF(1) = rs("CALCMAX")
                            Res(c0).TOPOSFSMP(1) = rs("TRS2")
                        Case "3"
                            Res(c0).TOPOSF(2) = rs("CALCMAX")
                            Res(c0).TOPOSFSMP(2) = rs("TRS3")
                        Case "4"
                            Res(c0).TOPOSF(3) = rs("CALCMAX")
                            Res(c0).TOPOSFSMP(3) = rs("TRS4")
                    End Select
                    Exit For
                ElseIf rs("SMPKBNCS") = "B" Then
                    Select Case rs("TRANCOND")
                        Case "1"
                            Res(c0).BOTOSF(0) = rs("CALCMAX")
                            Res(c0).BOTOSFSMP(0) = rs("TRS1")
                        Case "2"
                            Res(c0).BOTOSF(1) = rs("CALCMAX")
                            Res(c0).BOTOSFSMP(1) = rs("TRS2")
                        Case "3"
                            Res(c0).BOTOSF(2) = rs("CALCMAX")
                            Res(c0).BOTOSFSMP(2) = rs("TRS3")
                        Case "4"
                            Res(c0).BOTOSF(3) = rs("CALCMAX")
                            Res(c0).BOTOSFSMP(3) = rs("TRS4")
                    End Select
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' GD����
Public Sub GETTBCMJ006(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    '�G���[�n���h���̐ݒ�
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '�����񐔂̂����Ƃ��傫���l�����߂�
'        sql1 = "select max(TRANCNT) from TBCMJ006 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '����f�[�^�����߂�
'        sql = "select MSRSDVD2 from TBCMJ006 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).TOPDVD2 = -1
'        Else
'            Res(c0).TOPDVD2 = rs("MSRSDVD2")
'        End If
'        rs.Close
'
'        '�����񐔂̂����Ƃ��傫���l�����߂�
'        sql1 = "select max(TRANCNT) from TBCMJ006 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '����f�[�^�����߂�
'        sql = "select MSRSDVD2 from TBCMJ006 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTDVD2 = -1
'        Else
'            Res(c0).BOTDVD2 = rs("MSRSDVD2")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    '����f�[�^�����߂�
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, MSRSDVD2 from TBCMJ006 J "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=J.CRYNUM and POSITION=J.POSITION and SMPKBN=J.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("MSRSDVD2")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c0 = 1 To MaxRec
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                   (Res(c0).TOPSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).TOPSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).TOPDVD2 = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).TOPDVD2 = -1
    '        End If
            
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                   (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).BOTDVD2 = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).BOTDVD2 = -1
    '        End If
    '    Next
    'End If
    'rs.Close
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDGDCS AS TRS,"
    sql = sql & "     A.MSRSDVD2 "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ006 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     CRYSMPLIDGDCS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND"
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ006 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    sql = sql & " UNION "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDGDCS AS TRS,"
    sql = sql & "     A.MSRSDVD2 "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ006 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     CRYSMPLIDGDCS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND"
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ006 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If rs("SMPKBNCS") = "T" Then
                    Res(c0).TOPDVD2 = rs("MSRSDVD2")
                    Res(c0).TOPDVD2SMP = rs("TRS")
                    Exit For
                ElseIf rs("SMPKBNCS") = "B" Then
                    Res(c0).BOTDVD2 = rs("MSRSDVD2")
                    Res(c0).BOTDVD2SMP = rs("TRS")
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' ���C�t�^�C��
Public Sub GETTBCMJ007(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    '�G���[�n���h���̐ݒ�
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '�����񐔂̂����Ƃ��傫���l�����߂�
'        sql1 = "select max(TRANCNT) from TBCMJ007 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '����f�[�^�����߂�
'        sql = "select CALCMEAS from TBCMJ007 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).TOPLT = -1
'        Else
'            Res(c0).TOPLT = rs("CALCMEAS")
'        End If
'        rs.Close
'
'        '�����񐔂̂����Ƃ��傫���l�����߂�
'        sql1 = "select max(TRANCNT) from TBCMJ007 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '����f�[�^�����߂�
'        sql = "select CALCMEAS from TBCMJ007 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTLT = -1
'        Else
'            Res(c0).BOTLT = rs("CALCMEAS")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    '����f�[�^�����߂�
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, CALCMEAS from TBCMJ007 J "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=J.CRYNUM and POSITION=J.POSITION and SMPKBN=J.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("CALCMEAS")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c0 = 1 To MaxRec
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                   (Res(c0).TOPSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).TOPSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).TOPLT = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).TOPLT = -1
    '        End If
            
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                   (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).BOTLT = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).BOTLT = -1
    '        End If
    '    Next
    'End If
    'rs.Close
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDTCS AS TRS,"
    sql = sql & "     A.CALCMEAS "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ007 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     CRYSMPLIDTCS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND "
    sql = sql & " A.TRANCOND = '0' AND "
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ007 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    sql = sql & " UNION "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDTCS AS TRS,"
    sql = sql & "     A.CALCMEAS "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ007 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     CRYSMPLIDTCS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND "
    sql = sql & " A.TRANCOND = '0' AND "
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ007 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO) "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If rs("SMPKBNCS") = "T" Then
                    Res(c0).TOPLT = rs("CALCMEAS")
                    Res(c0).TOPLTSMP = rs("TRS")
                    Exit For
                ElseIf rs("SMPKBNCS") = "B" Then
                    Res(c0).BOTLT = rs("CALCMEAS")
                    Res(c0).BOTLTSMP = rs("TRS")
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' BMD����
Public Sub GETTBCMJ008(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'    Dim c1 As Integer
'
'    '�G���[�n���h���̐ݒ�
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        For c1 = 0 To 2
'            '�����񐔂̂����Ƃ��傫���l�����߂�
'            sql1 = "select max(TRANCNT) from TBCMJ008 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'            sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'            sql1 = sql1 & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'
'            '����f�[�^�����߂�
'            sql = "select MEASMAX from TBCMJ008 where "
'            sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'            sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'            sql = sql & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'            sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'            DoEvents
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                Res(c0).TOPBMD(c1) = -1
'            Else
'                Res(c0).TOPBMD(c1) = rs("MEASMAX")
'            End If
'            rs.Close
'
'            '�����񐔂̂����Ƃ��傫���l�����߂�
'            sql1 = "select max(TRANCNT) from TBCMJ008 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'            sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'            sql1 = sql1 & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'
'            '����f�[�^�����߂�
'            sql = "select MEASMAX from TBCMJ008 where "
'            sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'            sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'            sql = sql & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'            sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'            DoEvents
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                Res(c0).BOTBMD(c1) = -1
'            Else
'                Res(c0).BOTBMD(c1) = rs("MEASMAX")
'            End If
'            rs.Close
'        Next
'    Next
'    On Error GoTo 0
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim c2 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    '����f�[�^�����߂�
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, MEASMAX from TBCMJ008 J "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=J.CRYNUM and POSITION=J.POSITION and SMPKBN=J.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).TRANCOND = rs("TRANCOND")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("MEASMAX")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c2 = 0 To 2
    '        For c0 = 1 To MaxRec
    '            iTRANCNT = 0
    '            For c1 = 1 To recCount
    '                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                       (buf(c1).TRANCNT > iTRANCNT) And _
                       (buf(c1).TRANCOND = Trim(Str(c2 + 1))) And _
                       (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                       (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                       (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                        
    '                    iTRANCNT = buf(c1).TRANCNT
    '                    Res(c0).TOPBMD(c2) = buf(c1).JudgData
    '                    OKFlag = True
    '                End If
    '            Next
    '            If Not OKFlag Then
    '                Res(c0).BOTBMD(c2) = -1
    '            End If
                
    '            iTRANCNT = 0
    '            For c1 = 1 To recCount
    '                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                       (buf(c1).TRANCNT > iTRANCNT) And _
                       (buf(c1).TRANCOND = Trim(Str(c2 + 1))) And _
                       (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                       (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                       (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
    '
    '                    iTRANCNT = buf(c1).TRANCNT
    '                    Res(c0).BOTBMD(c2) = buf(c1).JudgData
    '                    OKFlag = True
    '                End If
    '            Next
    '            If Not OKFlag Then
    '                Res(c0).BOTBMD(c2) = -1
    '            End If
    '        Next
    '    Next
    'End If
    'rs.Close
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDB1CS AS TRS1,CRYSMPLIDB2CS AS TRS2,CRYSMPLIDB3CS AS TRS3,"
    sql = sql & "     A.MEASMAX, A.TRANCOND,A.TRANCNT"
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ008 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ008"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO  AND"
    sql = sql & "                     TRANCOND=A.TRANCOND ) AND"
    sql = sql & " (A.SMPLNO = CRYSMPLIDB1CS OR A.SMPLNO = CRYSMPLIDB2CS OR A.SMPLNO = CRYSMPLIDB3CS )  "
    sql = sql & " UNION "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDB1CS AS TRS1,CRYSMPLIDB2CS AS TRS2,CRYSMPLIDB3CS AS TRS3,"
    sql = sql & "     A.MEASMAX, A.TRANCOND,A.TRANCNT"
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ008 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ008"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO  AND"
    sql = sql & "                     TRANCOND=A.TRANCOND ) AND"
    sql = sql & " (A.SMPLNO = CRYSMPLIDB1CS OR A.SMPLNO = CRYSMPLIDB2CS OR A.SMPLNO = CRYSMPLIDB3CS )  "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If rs("SMPKBNCS") = "T" Then
                    Select Case rs("TRANCOND")
                        Case "1"
                            Res(c0).TOPBMD(0) = rs("MEASMAX")
                            Res(c0).TOPBMDSMP(0) = rs("TRS1")
                        Case "2"
                            Res(c0).TOPBMD(1) = rs("MEASMAX")
                            Res(c0).TOPBMDSMP(1) = rs("TRS2")
                        Case "3"
                            Res(c0).TOPBMD(2) = rs("MEASMAX")
                            Res(c0).TOPBMDSMP(2) = rs("TRS3")
                    End Select
                    Exit For
                ElseIf rs("SMPKBNCS") = "B" Then
                    Select Case rs("TRANCOND")
                        Case "1"
                            Res(c0).BOTBMD(0) = rs("MEASMAX")
                            Res(c0).BOTBMDSMP(0) = rs("TRS1")
                        Case "2"
                            Res(c0).BOTBMD(1) = rs("MEASMAX")
                            Res(c0).BOTBMDSMP(1) = rs("TRS2")
                        Case "3"
                            Res(c0).BOTBMD(2) = rs("MEASMAX")
                            Res(c0).BOTBMDSMP(2) = rs("TRS3")
                    End Select
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' ���H���o���ю���
Public Sub GETTBCMI001(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f10
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim MaxRecCode As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim c2 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     XTALC2,"
    sql = sql & "     PRCMCN,"
    sql = sql & "     SEED "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, TBCMI001 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     XTALC2 = CRYNUM AND"
    sql = sql & "     A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMI001 WHERE CRYNUM=A.CRYNUM)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                Res(c0).PRCMCN = rs("PRCMCN")
                Res(c0).SEED = rs("SEED")
                Res(c0).HSXCDIR = GetCodeField("SC", "28", Left(Res(c0).SEED, 1), "INFO2")
                Exit For
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub
' ������H����
Public Sub GETTBCMI002(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim buf() As type_DBDRV_cmgc001f4
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    '�G���[�n���h���̐ݒ�
'    On Error Resume Next

    '�m�b�`�ʒu�����߂�
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     XTALC2,"
    sql = sql & "     INGOTPOS,NCHPOS, DMTOP1, DMTOP2, DMTAIL1, DMTAIL2 "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, TBCMI002 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     XTALC2 = CRYNUM AND"
    sql = sql & "     A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMI002 WHERE CRYNUM=A.CRYNUM)"
    sql = sql & "order by INGOTPOS "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If Res(c0).PRCMCN = "A" Then
                            Res(c0).NCHPOS = rs("NCHPOS")
                            Res(c0).DMTOP1 = rs("DMTOP1")
                            Res(c0).DMTOP2 = rs("DMTOP2")
                            Res(c0).DMTAIL1 = rs("DMTAIL1")
                            Res(c0).DMTAIL2 = rs("DMTAIL2")
                            Res(c0).DIAMETER = (Res(c0).DMTOP1 + Res(c0).DMTOP2 + Res(c0).DMTAIL1 + Res(c0).DMTAIL2) / 4
                            Exit For
                ElseIf Res(c0).PRCMCN = "M" Then
                        If Res(c0).INGOTPOS >= rs("INGOTPOS") Then
                            Res(c0).NCHPOS = rs("NCHPOS")
                            Res(c0).DMTOP1 = rs("DMTOP1")
                            Res(c0).DMTOP2 = rs("DMTOP2")
                            Res(c0).DMTAIL1 = rs("DMTAIL1")
                            Res(c0).DMTAIL2 = rs("DMTAIL2")
                            Res(c0).DIAMETER = (Res(c0).DMTOP1 + Res(c0).DMTOP2 + Res(c0).DMTAIL1 + Res(c0).DMTAIL2) / 4
                            Exit For
                        End If
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' �p���������s��
Public Function DBDRV_scmzc_fcmgc001f_Haiki(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku) As FUNCTION_RETURN

    Dim BlockMng As typ_TBCME040
    Dim sql As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_scmzc_fcmgc001f_Haiki"

    DBDRV_scmzc_fcmgc001f_Haiki = FUNCTION_RETURN_FAILURE
    
    ' kuramoto�ύX
    '�u���b�N�Ǘ����X�V����
    sql = "update TBCME040 set "
    sql = sql & "KRPROCCD='" & MGPRCD_KAKUAGE & "', "             ' ���݊Ǘ��H��
    sql = sql & "NOWPROC='" & PROCD_KAKUAGE & "', "               ' ���ݍH��
    sql = sql & "DELCLS='1', "
    sql = sql & "LSTATCLS = 'H', "
    sql = sql & "RSTATCLS = 'T', "
    sql = sql & "UPDDATE=sysdate, "                     ' �X�V���t
    sql = sql & "SENDFLAG='0' "                         ' ���M�t���O
    sql = sql & "where "
    sql = sql & "CRYNUM='" & rec.CRYNUM & "' "
    sql = sql & "and INGOTPOS=" & rec.INGOTPOS & " "
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmgc001f_Haiki = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    '�u���b�N�Ǘ����X�V����
'    With BlockMng
'        .CryNum = rec.CryNum                    ' �����ԍ�
'        .INGOTPOS = rec.INGOTPOS                ' �������J�n�ʒu
'        .LENGTH = rec.LENGTH                    ' ����
'        .REALLEN = rec.LENGTH                   ' ������
'        .BLOCKID = rec.BLOCKID                  ' �u���b�NID
'        .KRPROCCD = MGPRCD_KAKUAGE              ' ���݊Ǘ��H��
'        .NOWPROC = PROCD_KAKUAGE                ' ���ݍH��
'        '.LPKRPROCCD = MGPRCD_KAKUAGE            ' �ŏI�ʉߊǗ��H��  --- �ŏI�ʉߍH���́AG�ɗ��Ƃ����H�����c��
'        '.LASTPASS = PROCD_KAKUAGE               ' �ŏI�ʉߍH��      --- �ŏI�ʉߍH���́AG�ɗ��Ƃ����H�����c��
'        .DELCLS = "1"                           ' �폜�敪
'        .LSTATCLS = "H"                         ' �ŏI��ԋ敪
'        .RSTATCLS = "T"                         ' ������ԋ敪
'        .HOLDCLS = "0"                          ' �z�[���h�敪
'        .BDCAUS = "   "                         ' �s�Ǘ��R
'    End With
'
'    If DBDRV_BlockMng_Upd(BlockMng) = FUNCTION_RETURN_FAILURE Then
'        Exit Function
'    End If

    DBDRV_scmzc_fcmgc001f_Haiki = FUNCTION_RETURN_SUCCESS
    
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

' �������g�������s��
Public Function DBDRV_scmzc_fcmgc001f_Remerto(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku) As FUNCTION_RETURN

    Dim BlockMng As typ_TBCME040
    Dim HIN As tFullHinban
    Dim sql As String
    Dim sWhere      As String       'ADD 2004/10/22 TCS)R.Kawaguchi
    Dim rec_xodc2() As typ_XSDC2    'ADD 2004/10/22 TCS)R.Kawaguchi
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_scmzc_fcmgc001f_Remerto"
    
    DBDRV_scmzc_fcmgc001f_Remerto = FUNCTION_RETURN_FAILURE

    ' kuramoto�ύX
    '�u���b�N�Ǘ����X�V����
    sql = "update TBCME040 set "
    sql = sql & "KRPROCCD='" & MGPRCD_RIMERUTO_UKEIRE & "', "     ' ���݊Ǘ��H��
    sql = sql & "NOWPROC='" & PROCD_RIMERUTO_UKEIRE & "', "       ' ���ݍH��
    sql = sql & "LSTATCLS = 'T', "
    sql = sql & "RSTATCLS = 'M', "
    sql = sql & "UPDDATE=sysdate, "                     ' �X�V���t
    sql = sql & "SENDFLAG='0' "                         ' ���M�t���O
    sql = sql & "where "
    sql = sql & "CRYNUM='" & rec.CRYNUM & "' "
    sql = sql & "and INGOTPOS=" & rec.INGOTPOS & " "
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmgc001f_Remerto = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '�u���b�N�Ǘ����X�V����
'    With BlockMng
'        .CryNum = rec.CryNum                    ' �����ԍ�
'        .INGOTPOS = rec.INGOTPOS                ' �������J�n�ʒu
'        .LENGTH = rec.LENGTH                    ' ����
'        .REALLEN = rec.LENGTH                   ' ������
'        .BLOCKID = rec.BLOCKID                  ' �u���b�NID
'        .KRPROCCD = MGPRCD_RIMERUTO_UKEIRE      ' ���݊Ǘ��H��
'        .NOWPROC = PROCD_RIMERUTO_UKEIRE        ' ���ݍH��
'        '.LPKRPROCCD = MGPRCD_KAKUAGE            ' �ŏI�ʉߊǗ��H��  --- �ŏI�ʉߍH���́AG�ɗ��Ƃ����H�����c��
'        '.LASTPASS = PROCD_KAKUAGE               ' �ŏI�ʉߍH��      --- �ŏI�ʉߍH���́AG�ɗ��Ƃ����H�����c��
'        .DELCLS = "0"                           ' �폜�敪
'        .LSTATCLS = "T"                         ' �ŏI��ԋ敪
'        .RSTATCLS = "M"                         ' ������ԋ敪
'        .HOLDCLS = "0"                          ' �z�[���h�敪
'        .BDCAUS = "   "                         ' �s�Ǘ��R
'    End With
'    If DBDRV_BlockMng_Upd(BlockMng) = FUNCTION_RETURN_FAILURE Then
'        Exit Function
'    End If
    
    '�i�Ԃ�'Z'�ɕς���
    With rec
        HIN.HINBAN = "Z"
        HIN.mnorevno = 0
        HIN.factory = "Y"
        HIN.opecond = "1"
        If ChangeAreaHinban(.CRYNUM, .INGOTPOS, .LENGTH, HIN) = FUNCTION_RETURN_FAILURE Then
            Exit Function
        End If
    End With

'---- ADD [���������V�X�e���Ή�] 2004/10/22 TCS)R.Kawaguchi START ----
    ''�u���b�N�Ǘ��̏��(�d�ʁA�������J�n�ʒu)���擾
    'WHERE������
    sWhere = "WHERE CRYNUMC2 = '" & rec.BLOCKID & "'"
    Call DBDRV_GetXSDC2(rec_xodc2(), sWhere)
    '�Y���f�[�^�����̏ꍇ
    If UBound(rec_xodc2) = 0 Then
        GoTo proc_exit
    End If

    ''���������Ǘ�(XODCX)�쐬����
    If InsXODCX(rec, rec_xodc2(1)) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    ''�����H������(XODB3)�쐬����
    If InsXODB3(rec, rec_xodc2(1)) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    
    '*** UPDATE START T.TERAUCHI 2004/12/06 �����������ٔ��s�p�����ǉ��Ή�
        '�����������ٔ��s����
        
    '*** UPDATE START T.TERAUCHI 2005/01/18 �����������ٔ��s�p�H�����ޕύX�Ή�
    '   If Ins_TBCMC001_New("RP10", "81", f_cmbc008_1.txtStaffID.Text, rec.BLOCKID & "0", gsSysdate) = False Then
        If Ins_TBCMC001_New(Right(PROCD_KAKUAGE, 4), "81", f_cmbc008_1.txtStaffID.Text, rec.BLOCKID & "0", gsSysdate) = False Then
    '*** UPDATE END   T.TERAUCHI 2005/01/18
                
            GoTo proc_exit
        End If
    '*** UPDATE END   T.TERAUCHI 2004/12/06
    
'---- ADD [���������V�X�e���Ή�] 2004/10/22 TCS)R.Kawaguchi END ----

    DBDRV_scmzc_fcmgc001f_Remerto = FUNCTION_RETURN_SUCCESS
    
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

'�u���b�N�̗��_�����𓾂�
Private Function GetBlockLength(blkID$) As Integer
Dim sql$
Dim rs As OraDynaset

    sql = "select LENGTH from TBCME040 where BLOCKID='" & blkID & "'"
    Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
    If rs.RecordCount = 0 Then
        GetBlockLength = 0
    Else
        GetBlockLength = rs("LENGTH")
    End If
    rs.Close
    Set rs = Nothing
End Function

' �i�グ�������s��
Public Function DBDRV_scmzc_fcmgc001f_Kakuage(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku) As FUNCTION_RETURN

    Dim sql As String
    Dim fullHinban As tFullHinban
    Dim BlockMng As typ_TBCME040
    Dim CC() As typ_TBCMG008

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_scmzc_fcmgc001f_Kakuage"
    DBDRV_scmzc_fcmgc001f_Kakuage = FUNCTION_RETURN_FAILURE

    '12���i�Ԃ����߂�
    If GetLastHinban(NewHinBan, fullHinban) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If

    'XSDCS���X�V����
    If ChangeXSDCSHinban(rec.BLOCKID, fullHinban) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If

    '�i�ԊǗ����X�V����
    If ChangeAreaHinban(rec.CRYNUM, rec.INGOTPOS, GetBlockLength(rec.BLOCKID), fullHinban) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    
    ' kuramoto�ύX
    '�u���b�N�Ǘ����X�V����
    sql = "update TBCME040 set "
    sql = sql & "KRPROCCD='" & MGPRCD_KESSYOU_SOUGOUHANTEI & "', "   ' ���݊Ǘ��H��
    sql = sql & "NOWPROC='" & PROCD_KESSYOU_SOUGOUHANTEI & "', "     ' ���ݍH��
    sql = sql & "LPKRPROCCD='" & MGPRCD_KAKUAGE & "', "              ' �ŏI�ʉߊǗ��H��
    sql = sql & "LASTPASS='" & PROCD_KAKUAGE & "', "                 ' �ŏI�ʉߍH��
    sql = sql & "LSTATCLS = 'T', "
    sql = sql & "RSTATCLS = 'T', "
    sql = sql & "UPDDATE=sysdate, "                     ' �X�V���t
    sql = sql & "SUMMITSENDFLAG='0', "                  ' SUMMIT���M�t���O
    sql = sql & "SENDFLAG='0' "                         ' ���M�t���O
    sql = sql & "where "
    sql = sql & "CRYNUM='" & rec.CRYNUM & "' "
    sql = sql & "and INGOTPOS=" & rec.INGOTPOS & " "
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmgc001f_Kakuage = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '�u���b�N�Ǘ����X�V����
'    With BlockMng
'        .CryNum = rec.CryNum                    ' �����ԍ�
'        .INGOTPOS = rec.INGOTPOS                ' �������J�n�ʒu
'        .LENGTH = rec.LENGTH                    ' ����
'        .REALLEN = rec.LENGTH                   ' ������
'        .BLOCKID = rec.BLOCKID                  ' �u���b�NID
'        .KRPROCCD = MGPRCD_KESSYOU_SOUGOUHANTEI ' ���݊Ǘ��H��
'        .NOWPROC = PROCD_KESSYOU_SOUGOUHANTEI   ' ���ݍH��
'        .LPKRPROCCD = MGPRCD_KAKUAGE            ' �ŏI�ʉߊǗ��H��
'        .LASTPASS = PROCD_KAKUAGE               ' �ŏI�ʉߍH��
'        .DELCLS = "0"                           ' �폜�敪
'        .LSTATCLS = "T"                         ' �ŏI��ԋ敪
'        .RSTATCLS = "T"                         ' ������ԋ敪
'        .HOLDCLS = "0"                          ' �z�[���h�敪
'        .BDCAUS = "   "                         ' �s�Ǘ��R
'    End With
'    If DBDRV_BlockMng_Upd(BlockMng) = FUNCTION_RETURN_FAILURE Then
'        GoTo proc_exit
'    End If

    If DBDRV_PutTBCMG008(rec, fullHinban) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmgc001f_Kakuage = FUNCTION_RETURN_SUCCESS

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

Public Function DBDRV_PutTBCMG008(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku, fullHinban As tFullHinban) As FUNCTION_RETURN

    Dim rs As OraDynaset    'RecordSet
    Dim sql As String
    Dim CC() As typ_TBCMG008
    Dim InsertFlag As Boolean

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_PutTBCMG008"
    DBDRV_PutTBCMG008 = FUNCTION_RETURN_FAILURE

    '�����񐔂̂����Ƃ��傫���l�����߂�
    sql = "where "
    sql = sql & "CRYNUM = '" & rec.BLOCKID & "' "
    sql = sql & "and TRANCNT = any("
    sql = sql & "select max(TRANCNT) from TBCMG008 where CRYNUM = '" & rec.BLOCKID & "'"
    sql = sql & ") "
    If DBDRV_GetTBCMG008(CC(), sql) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If

    InsertFlag = False
    If UBound(CC()) = 0 Then
        ReDim CC(1) As typ_TBCMG008
        InsertFlag = True
    End If
    With CC(1)
        .CRYNUM = rec.BLOCKID
        .KRPROCCD = MGPRCD_KAKUAGE          ' �Ǘ��H���R�[�h
        .PROCCODE = PROCD_KAKUAGE           ' �H���R�[�h
'        .OHINBAN = rec.hinban              ' ���i��
'        .OMNOREVNO = rec.REVNUM            ' �����i�ԍ������ԍ�
'        .OFACTORY = rec.factory            ' ���H��
'        .OOPECOND = rec.opecond            ' �����Ə���
        
        .OHINBAN = "G       "               ' ���i��
        .OMNOREVNO = 0                      ' �����i�ԍ������ԍ�
        .OFACTORY = " "                     ' ���H��
        .OOPECOND = " "                     ' �����Ə���
        
        .NHINBAN = fullHinban.HINBAN        ' �V�i��
        .NMNOREVNO = fullHinban.mnorevno    ' �V���i�ԍ������ԍ�
        .NFACTORY = fullHinban.factory      ' �V�H��
        .NOPECOND = fullHinban.opecond      ' �V���Ə���
    End With

    sql = " insert into TBCMG008 ( "
    sql = sql & "CRYNUM, "      ' �����ԍ�
    sql = sql & "KRPROCCD, "    ' �Ǘ��H���R�[�h
    sql = sql & "PROCCODE, "    ' �H���R�[�h
    sql = sql & "NHINBAN, "     ' �V�i��
    sql = sql & "NMNOREVNO, "   ' �V���i�ԍ������ԍ�
    sql = sql & "NFACTORY, "    ' �V�H��
    sql = sql & "NOPECOND, "    ' �V���Ə���
    sql = sql & "OHINBAN, "     ' ���i��
    sql = sql & "OMNOREVNO, "   ' �����i�ԍ������ԍ�
    sql = sql & "OFACTORY, "    ' ���H��
    sql = sql & "OOPECOND, "    ' �����Ə���
    sql = sql & "KSTAFFID, "    ' �X�V�Ј��h�c
    sql = sql & "UPDDATE, "     ' �X�V���t
    sql = sql & "SENDFLAG, "    ' ���M�t���O
    sql = sql & "SENDDATE,"      ' ���M���t
    sql = sql & "TRANCNT, "     ' ������
    sql = sql & "TSTAFFID, "    ' �o�^�Ј�ID
    sql = sql & "REGDATE "     ' �o�^���t
    sql = sql & ")"
    With CC(1)
        sql = sql & " values ( "
        sql = sql & "'" & .CRYNUM & "'," ' �����ԍ�
        sql = sql & "'" & .KRPROCCD & "'," ' �Ǘ��H���R�[�h
        sql = sql & "'" & .PROCCODE & "'," ' �H���R�[�h
        sql = sql & "'" & .NHINBAN & "'," ' �V�i��
        sql = sql & .NMNOREVNO & ","  ' �V���i�ԍ������ԍ�
        sql = sql & "'" & .NFACTORY & "'," ' �V�H��
        sql = sql & "'" & .NOPECOND & "'," ' �V���Ə���
        sql = sql & "'" & .OHINBAN & "'," ' ���i��
        sql = sql & .OMNOREVNO & "," ' �����i�ԍ������ԍ�
        sql = sql & "'" & .OFACTORY & "'," ' ���H��
        sql = sql & "'" & .OOPECOND & "'," ' �����Ə���
        sql = sql & "'" & STAFFIDBUFF & "'," ' �X�V�Ј��h�c
        sql = sql & "sysdate," ' �X�V���t
        sql = sql & "'0'," ' ���M�t���O
        sql = sql & "sysdate," ' ���M���t
        If InsertFlag Then
            sql = sql & "1," ' ������
            sql = sql & "'" & STAFFIDBUFF & "'," ' �o�^�Ј�ID
            sql = sql & "sysdate" ' �o�^���t
        Else
            sql = sql & .TRANCNT + 1 & "," ' ������
            sql = sql & "'" & .TSTAFFID & "',"  ' �o�^�Ј�ID
            sql = sql & "sysdate"  ' �o�^���t
        End If
        sql = sql & ")"
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        GoTo proc_exit
    End If

    DBDRV_PutTBCMG008 = FUNCTION_RETURN_SUCCESS
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


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMG008�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMG008 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMG008_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMG008(records() As typ_TBCMG008, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, TRANCNT, KRPROCCD, PROCCODE, NHINBAN, NMNOREVNO, NFACTORY, NOPECOND, OHINBAN, OMNOREVNO, OFACTORY," & _
              " OOPECOND, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMG008"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMG008 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ��i�i�グ�j
            .TRANCNT = rs("TRANCNT")         ' ������
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .NHINBAN = rs("NHINBAN")         ' �V�i��
            .NMNOREVNO = rs("NMNOREVNO")     ' �V���i�ԍ������ԍ�
            .NFACTORY = rs("NFACTORY")       ' �V�H��
            .NOPECOND = rs("NOPECOND")       ' �V���Ə���
            .OHINBAN = rs("OHINBAN")         ' ���i��
            .OMNOREVNO = rs("OMNOREVNO")     ' �����i�ԍ������ԍ�
            .OFACTORY = rs("OFACTORY")       ' ���H��
            .OOPECOND = rs("OOPECOND")       ' �����Ə���
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMG008 = FUNCTION_RETURN_SUCCESS
End Function


'---- ADD [���������Ǘ��쐬�Ή�] 2004/10/22 TCS)R.Kawaguchi START ----

' @(f)
' �@�\      : XSDC1�����֐�
'
' �Ԃ�l    : �Ȃ�
'
' ������    : �������ʊi�[�\����
'
' �@�\����  : �e�����ԍ��̈��グ�p�^�[�����擾����
Public Sub GETXSDC1(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)

    Dim buf()       As type_DBDRV_xsdc1
    Dim sql         As String
    Dim rs          As OraDynaset    'RecordSet
    Dim recCount    As Integer
    Dim MaxRec      As Integer
    Dim c0          As Integer
    Dim c1          As Integer
    Dim OKFlag      As Boolean

    '�G���[�n���h���̐ݒ�
    On Error Resume Next

    '����f�[�^�����߂�
    sql = ""
    sql = "select XTALC1, PUPTNC1 from XSDC1 "

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
        ReDim buf(recCount) As type_DBDRV_xsdc1
        For c0 = 1 To recCount
            buf(c0).CRYNUM = rs("XTALC1")
            buf(c0).HIKIAGEPTRN = NulltoStr(rs("PUPTNC1"))
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
                
            '*** UPDATE START T.TERAUCHI 2004/12/06 �������u���b�NID���猋���ԍ��ɕύX
            '    If (Res(c0).BLOCKID = buf(c1).CRYNUM) Then
                If (Res(c0).CRYNUM = buf(c1).CRYNUM) Then
            '*** UPDATE END   T.TERAUCHI 2004/12/06
                     
                    Res(c0).HIKIAGEPTRN = buf(c1).HIKIAGEPTRN
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).HIKIAGEPTRN = " "
            End If
            OKFlag = False
        Next

    End If
    On Error GoTo 0

End Sub

' @(f)
' �@�\      : XODCX�쐬�֐�
'
' �Ԃ�l    : FUNCTION_RETURN_FAILURE�F�ُ�
'             FUNCTION_RETURN_SUCCESS�F����
'
' ������    : �������ʊi�[�\����
'
' �@�\����  : �Y���f�[�^�̐��������Ǘ����쐬����
Private Function InsXODCX(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku, _
                            rec_xodc2 As typ_XSDC2) As FUNCTION_RETURN

    Dim objDS       As Object
    Dim sSql        As String
    Dim sDopType    As String
    Dim sCSDop      As String       'CS�h�[�v�L��
    Dim sNDop       As String       '���f�h�[�v�L��
    Dim sWhere      As String
    Dim sUserID     As String
    Dim sSCNTRL     As String       '���ʺ��۰ٺ��� ADD 2011/03/24 TSMC�i���ʑΉ�
    
'*** UPDATE START T.TERAUCHI 2004/12/07 ײ���юd�l�L��
    Dim sLTUmu      As String
'*** UPDATE END   T.TERAUCHI 2004/12/07
'*** UPDATE START TAGAWA 2004/12/16
    Dim sFlag       As String
'*** UPDATE END  TAGAWA 2004/12/16

On Error GoTo PROC_ERR

    InsXODCX = FUNCTION_RETURN_FAILURE
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function InsXODCX"
    
    '�o�^�Ј��h�c
    sUserID = f_cmbc008_1.txtStaffID.Text
    
    With rec
        
        '///����������{���擾SQL�쐬
        Call GetAssistSQL_300(sSql, .CRYNUM)
        If DynSet2(objDS, sSql) = False Then
            If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
            GoTo proc_exit
        End If
        '�Y���f�[�^�����̏ꍇ
        If objDS.EOF = True Then
            If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
            GoTo proc_exit
        End If

        '///CS�h�[�v�L���A���f�h�[�v�L���ݒ�
'*** UPDATE START T.TERAUCHI 2004/12/07 ���������啶���ϊ��Ή�
'        sDopType = NulltoStr(objDS.Fields("DTYPEC1").Value)
'        If sDopType = " " Or sDopType = "p" Then
'            sCSDop = "2"
'            sNDop = "1"
'        ElseIf sDopType = "n" Then
'            sCSDop = "1"
'            sNDop = "2"
'        Else
'            sCSDop = " "
'            sNDop = " "
'        End If
        sDopType = UCase(NulltoStr(objDS.Fields("DTYPEC1").Value))
'*** UPDATE START TAGAWA 2004/12/16**************************
''        If sDopType = " " Or sDopType = "P" Then
''            sCSDop = "2"
''            sNDop = "1"
''        ElseIf sDopType = "N" Then
''            sCSDop = "1"
''            sNDop = "2"
''        Else
''            sCSDop = " "
''            sNDop = " "
''        End If
        ''�����h�[�v�擾
        sFlag = UCase(Trim(NulltoStr(objDS.Fields("DPNTCLS").Value)))
        ''Cs�h�[�v�̎�
        If sFlag = "C" Then
            sCSDop = "2"
            sNDop = "1"
        ''���f�h�[�v�̎�
        ElseIf sFlag = "N" Then
            sCSDop = "1"
            sNDop = "2"
        ''�v�h�[�v�̎�
        ElseIf sFlag = "M" Then
            sCSDop = "2"
            sNDop = "2"
        ''���̑�
        Else
            sCSDop = "1"
            sNDop = "1"
        End If
'*** UPDATE END TAGAWA 2004/12/16**************************
'*** UPDATE END   T.TERAUCHI 2004/12/07

    '*** UPDATE START T.TERAUCHI 2004/12/07 ײ���юd�l�L������
        ''���C�t�^�C���d�l�L��
        If objDS.Fields("HSXLTHWS").Value = "H" Then
            ''�L��
            sLTUmu = "2"
        Else
            ''����
            sLTUmu = "1"
        End If
    '*** UPDATE END   T.TERAUCHI 2004/12/07
    '*** UPDATE START Marushita 2011/03/24 TSMC�i���ʑΉ�
        ''���������`�F�b�N�t���O�̔��f
        If NulltoStr(objDS.Fields("MTRLCHKFLG").Value) = "1" Then
            ''�i��NULL���̎��ʃR���g���[���R�[�h�Z�b�g(��3��)
            If NulltoStr(objDS.Fields("HINBCX").Value) = "" Then
                sSCNTRL = "   "
            Else
                ''���ʃR���g���[���R�[�h�Z�b�g(�i��3��)
                sSCNTRL = Left(objDS.Fields("HINBCX").Value, 3)
            End If
        Else
            ''���ʃR���g���[���R�[�h�Z�b�g(��3��)
            sSCNTRL = "   "
        End If
    '*** UPDATE END   Marushita 2011/03/24
        '///���������Ǘ��쐬
        sSql = ""
        sSql = sSql & "INSERT INTO xodcx(" & vbLf
        sSql = sSql & "crynumcx" & vbLf    ''�u���b�NID
        sSql = sSql & ",mtrlnumcx" & vbLf   ''������
        sSql = sSql & ",wkktcx" & vbLf      ''�H���R�[�h
        sSql = sSql & ",workcx" & vbLf      ''�H��R�[�h
        sSql = sSql & ",hdaycx" & vbLf      ''��������
        sSql = sSql & ",weightcx" & vbLf    ''�d��
        sSql = sSql & ",htkbncx" & vbLf     ''�p��/�K���敪
        sSql = sSql & ",divumucx" & vbLf    ''�����L��
        sSql = sSql & ",toworkcx" & vbLf    ''���o��H��R�[�h
        sSql = sSql & ",frworkcx" & vbLf    ''�����H��R�[�h
        sSql = sSql & ",hinbcx" & vbLf      ''�i��
        sSql = sSql & ",typecx" & vbLf      ''�^�C�v
        sSql = sSql & ",dptypecx" & vbLf    ''�h�[�v�^�C�v
        sSql = sSql & ",tposcx" & vbLf      ''�ʒuL(�g�b�v��)
        sSql = sSql & ",lencx" & vbLf       ''�u���b�N����
        sSql = sSql & ",siweightcx" & vbLf  ''�d���ݏd��
        sSql = sSql & ",updmcx" & vbLf      ''����AV�a
        sSql = sSql & ",prodmcx" & vbLf     ''���i�a
        sSql = sSql & ",tdopposcx" & vbLf   ''�ǉ��h�[�v�����ʒuL
        sSql = sSql & ",wdopumucx" & vbLf   ''W�h�[�v(P/N����)�L��
        sSql = sSql & ",csdopumucx" & vbLf  ''CS�h�[�v�L��
        sSql = sSql & ",ndopumucx" & vbLf   ''���f�h�[�v�L��
        sSql = sSql & ",ltspecumucx" & vbLf ''���C�t�^�C���d�l�L��
        sSql = sSql & ",csspecumucx" & vbLf ''CS�d�l�L��
        sSql = sSql & ",topwcx" & vbLf      ''�g�b�vWT
        sSql = sSql & ",dmkcx" & vbLf       ''���a�敪
        sSql = sSql & ",xtalcx" & vbLf      ''�����ԍ�
        sSql = sSql & ",livkcx" & vbLf      ''�����敪
        sSql = sSql & ",unifgcx" & vbLf     ''����FLG
        sSql = sSql & ",twarifgcx" & vbLf   ''�c��FLG
        sSql = sSql & ",refusefgcx" & vbLf  ''�����FLG
        sSql = sSql & ",tstafidcx" & vbLf   ''�o�^�Ј�ID
        sSql = sSql & ",tdaycx" & vbLf      ''�o�^���t
    '*** UPDATE START T.TERAUCHI 2004/12/07 �X�V�ҁA�X�V������ǉ�
        sSql = sSql & ",kstafidcx" & vbLf   ''�X�V��
        sSql = sSql & ",kdaycx" & vbLf      ''�X�V����
    '*** UPDATE END   T.TERAUCHI 2004/12/07
        sSql = sSql & ",crydopcx" & vbLf    ''�����h�[�v
        sSql = sSql & ",crydopvlcx" & vbLf  ''�����h�[�v��
        sSql = sSql & ",bkformcx" & vbLf    ''�u���b�N�`��
        sSql = sSql & ",pgidcx" & vbLf      ''PG-ID
        sSql = sSql & ",blktypcx" & vbLf    ''�u���b�N���
        sSql = sSql & ",tkacutwcx" & vbLf   ''T�T���v���O�d��
    '*** UPDATE START TAGAWA 2004/12/16***************
        sSql = sSql & ",denflgcx" & vbLf     ''�d�ɍރt���O
    '*** UPDATE END   TAGAWA 2004/12/16***************
    '*** UPDATE START T.TERAUCHI 2004/12/07 į�ߎ�o��WT�ǉ�
        sSql = sSql & ",toptwcx" & vbLf     ''į�ߎ�o��WT
    '*** UPDATE END   T.TERAUCHI 2004/12/07
    '*** UPDATE START Marushita 2011/03/24 TSMC�i���ʑΉ�
        sSql = sSql & ",scntrlcx" & vbLf    ''���ʺ��۰ٺ���
    '*** UPDATE END   Marushita 2011/03/24

        sSql = sSql & ")values(" & vbLf
        sSql = sSql & "'" & .BLOCKID & "0" & "'" & vbLf                         ''��ۯ�ID
        sSql = sSql & ",' '" & vbLf                                             ''����No
        sSql = sSql & ",'" & Right(PROCD_KOUNYU_TAN_KESSYOU, 4) & "'" & vbLf    ''�H������B410
        sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''�H�꺰��
        sSql = sSql & ",sysdate" & vbLf                                         ''��������
        sSql = sSql & "," & rec_xodc2.GNWC2 & vbLf                              ''��������(��ۯ�)����ݏd��
        sSql = sSql & ",'1'" & vbLf                                             ''�p���E�K���敪
        sSql = sSql & ",'1'" & vbLf                                             ''�����L��
        
    '*** UPDATE START T.TERAUCHI 2004/12/07 ���o�H�꺰�ނ�ݒ�
    '    sSQL = sSQL & ",' '" & vbLf                                             ''���o��H�꺰��
        sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''���o��H�꺰��
    '*** UPDATE END   T.TERAUCHI 2004/12/07
    
        sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''�����H�꺰��
        sSql = sSql & ",'" & objDS.Fields("HINBCX").Value & "'" & vbLf          ''�i��
        sSql = sSql & ",'" & objDS.Fields("HSXTYPE").Value & "'  " & vbLf       ''�^�C�v
        sSql = sSql & ",'" & sDopType & "'" & vbLf                              ''�h�[�v�^�C�v
        sSql = sSql & ", " & rec_xodc2.INPOSC2 & vbLf                           ''��������(��ۯ�)��������J�n�ʒu
        sSql = sSql & ", " & .LENGTH & vbLf                                     ''�u���b�N����
        sSql = sSql & ", " & ConvNum(objDS.Fields("SUICHARGE").Value) & vbLf    ''�d���ݏd��
        sSql = sSql & ", " & ConvNum(objDS.Fields("UPDMCX").Value) & vbLf       ''����AV�a
        sSql = sSql & ", " & ConvNum(objDS.Fields("PRODMCX").Value) & vbLf      ''���i�a
        sSql = sSql & ", " & ConvNum(objDS.Fields("ADDOPPC1").Value) & vbLf     ''�ǉ��h�[�v�����ʒuL
        sSql = sSql & ",'1'" & vbLf                                             ''W�h�[�v(P/N����)�L��
        sSql = sSql & ",'" & sCSDop & "'" & vbLf                                ''CS�h�[�v�L��
        sSql = sSql & ",'" & sNDop & "'" & vbLf                                 ''���f�h�[�v�L��
        
    '*** UPDATE START T.TERAUCHI 2004/12/07 ײ���юd�l�L���́A���茋�ʂ�蔻��
    '    sSQL = sSQL & ",'" & objDS.Fields("HSXLTHWS").Value & "'" & vbLf        ''���C�t�^�C���g�p�L��
        sSql = sSql & ",'" & sLTUmu & "'" & vbLf                                ''���C�t�^�C���g�p�L��
    '*** UPDATE END   T.TERAUCHI 2004/12/07
        
        sSql = sSql & ",'2'" & vbLf                                             ''CS�g�p�L��
        
    '*** UPDATE START T.TERAUCHI 2004/12/07 TOPWT�����d�ʂɕύX
    '    sSQL = sSQL & "," & ConvNum(objDS.Fields("PUTCUTWC1").Value) & vbLf     ''�g�b�vWT
        sSql = sSql & "," & ConvNum(objDS.Fields("CTR01A9").Value) & vbLf       '�f�g�b�vWT
    '*** UPDATE END   T.TERAUCHI 2004/12/07
        
        sSql = sSql & ",'300'" & vbLf                                           ''���a�敪
        sSql = sSql & ",'" & .CRYNUM & "'" & vbLf                               ''���㌋���ԍ�
        sSql = sSql & ",'0'" & vbLf                                             ''�����敪
        sSql = sSql & ",'0'" & vbLf                                             ''����FLG
        sSql = sSql & ",'0'" & vbLf                                             ''�c��FLG
        sSql = sSql & ",'0'" & vbLf                                             ''�����FLG
        sSql = sSql & ",'" & sUserID & "'" & vbLf                               ''�o�^�Ј�ID
        sSql = sSql & ",sysdate" & vbLf                                         ''�o�^���t
    '*** UPDATE START T.TERAUCHI 2004/12/07 �X�V�ҁA�X�V�����ǉ��Ή�
        sSql = sSql & ",'" & sUserID & "'" & vbLf                               ''�X�V�Ј�ID
        sSql = sSql & ",sysdate" & vbLf                                         ''�X�V���t
    '*** UPDATE END   T.TERAUCHI 2004/12/07
        sSql = sSql & ",'" & objDS.Fields("DPNTCLS").Value & "'" & vbLf         ''�����h�[�v
        sSql = sSql & "," & ConvNum(objDS.Fields("DOPANT").Value) & vbLf        ''�����h�[�v��
        sSql = sSql & ",'3'" & vbLf                                             ''�u���b�N�`��
        sSql = sSql & ",'" & objDS.Fields("PGID").Value & "'" & vbLf            ''PG-ID
        sSql = sSql & ",'A'" & vbLf                                             ''�u���b�N���
        sSql = sSql & ",0" & vbLf                                               ''T�T���v���O�d��
    '*** UPDATE START TAGAWA 2004/12/16***************
        sSql = sSql & ",'1'" & vbLf                                             ''�d�ɍރt���O
    '*** UPDATE END   TAGAWA 2004/12/16******************
    '*** UPDATE START T.TERAUCHI 2004/12/07
        sSql = sSql & "," & ConvNum(objDS.Fields("PUTCUTWC1").Value) & vbLf     ''į�ߎ�o��WT
    '*** UPDATE END   T.TERAUCHI 2004/12/07
    '*** UPDATE START Marushita 2011/03/24 TSMC�i���ʑΉ�
        sSql = sSql & ",'" & sSCNTRL & "'" & vbLf                               ''���ʺ��۰ٺ���
    '*** UPDATE END   Marushita 2011/03/24
    
        sSql = sSql & ")"
        
        If SqlExec2(sSql) = -1 Then
            If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
            GoTo proc_exit
        End If
        
    End With
    
    InsXODCX = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing

    '�I��
    gErr.Pop
    Exit Function

PROC_ERR:
    
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit

End Function

' @(f)
' �@�\      : XODB3�쐬�֐�
'
' �Ԃ�l    : FUNCTION_RETURN_FAILURE�F�ُ�
'             FUNCTION_RETURN_SUCCESS�F����
'
' ������    : �������ʊi�[�\����
'
' �@�\����  : �Y���f�[�^�̐��������Ǘ����쐬����
Private Function InsXODB3(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku, _
                            rec_xodc2 As typ_XSDC2) As FUNCTION_RETURN

    Dim objDS       As Object
    Dim sSql        As String
    Dim sUserID     As String
    Dim sUserName   As String
    Dim iIdx        As Integer
    Dim iRenban     As Integer
    Dim sYear       As String
    Dim sMonth      As String
    Dim sDay        As String
    Dim sHour       As String
    Dim sMin        As String
    Dim sNowdate    As String
    Dim sCyoku      As String

On Error GoTo PROC_ERR

    InsXODB3 = FUNCTION_RETURN_FAILURE
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function InsXODB3"
    
    '****** �o�^���̍쐬 *****
    '' �o�^�Ј��h�c�A�Ј���
    sUserID = f_cmbc008_1.txtStaffID.Text
    sUserName = f_cmbc008_1.txtJfName.Text
    
    '' �V�X�e�����t�A���ѓ��t���̐ݒ�
    If Not GetSysdate Then
        GoTo proc_exit
    End If
    sNowdate = gsSysdate
    '�T�[�o�[�V�X�e�����t�����ѓ��ɕύX
    sNowdate = GetJITUDATE(Format(sNowdate, "yyyymmddhhmmss"))
    '���ѓ���蒼�敪�𔻒�
    sCyoku = GetCYOKU(gsSysdate)
    '���ѓ�����؂���
    sYear = Mid(sNowdate, 1, 4)     '�N
    sMonth = Mid(sNowdate, 5, 2)    '��
    sDay = Mid(sNowdate, 7, 2)      '��
    sHour = Mid(sNowdate, 9, 2)     '��
    sMin = Mid(sNowdate, 11, 2)     '��
        
    '' �H���A�Ԃ̎擾
    iRenban = 0
    sSql = ""
    sSql = sSql & " SELECT NVL(MAX(kcntb3),0) maxcnt     " & vbLf   '�H���A��
    sSql = sSql & " FROM   xodb3                         " & vbLf
    
'*** UPDATE START T.TERAUCHI 2004/12/06 ���i���b�gNo��13���Ƃ���
'    sSQL = sSQL & " WHERE  polnob3 = '" & rec.BLOCKID & "'" & vbLf
    sSql = sSql & " WHERE  polnob3 = '" & rec.BLOCKID & "0" & "'" & vbLf
'*** UPDATE END   T.TERAUCHI 2004/12/06
    
    'SQL�����s
    If DynSet2(objDS, sSql) = False Then
        If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
        GoTo proc_exit
    End If
    '�擾�����f�[�^���i�[
    iRenban = objDS.Fields("maxcnt").Value + 1
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
        
    With rec
        
        '****** �����H�����э쐬 ******
        sSql = ""
        sSql = sSql & "insert into XODB3(                       " & vbLf
        sSql = sSql & "             POLNOB3                     " & vbLf   '�����ԍ�
        sSql = sSql & "            ,KCNTB3                      " & vbLf   '�H���A��
        sSql = sSql & "            ,CRSEQB3                     " & vbLf   '�����A��
        sSql = sSql & "            ,TDAYB3                      " & vbLf   '�o�^���t
        sSql = sSql & "            ,RDAYB3                      " & vbLf   '�C�����t
        sSql = sSql & "            ,SDAYB3                      " & vbLf   '���M���t
        sSql = sSql & "            ,SNDKB3                      " & vbLf   '���M�敪
        sSql = sSql & "            ,SAKJB3                      " & vbLf   '�폜�敪
        sSql = sSql & "            ,POKUBB3                     " & vbLf   '�����敪
        sSql = sSql & "            ,POKIDCB3                    " & vbLf   '������ރR�[�h
        sSql = sSql & "            ,POLTNB3                     " & vbLf   '�������b�gNo
        sSql = sSql & "            ,MODKBB3                     " & vbLf   '�ԍ��敪
        sSql = sSql & "            ,SUMKBB3                     " & vbLf   '�W�v�敪
        sSql = sSql & "            ,WKKTB3                      " & vbLf   '�H���R�[�h
        sSql = sSql & "            ,PLACB3                      " & vbLf   '���C���R�[�h
        sSql = sSql & "            ,FRWB3                       " & vbLf   '����d��
        sSql = sSql & "            ,TOWB3                       " & vbLf   '���o�d��
        sSql = sSql & "            ,LOSWB3                      " & vbLf   '���X�d��
        sSql = sSql & "            ,FRWKKTB3                    " & vbLf   '����H���R�[�h
        sSql = sSql & "            ,TOWKKTB3                    " & vbLf   '���o�H���R�[�h
        sSql = sSql & "            ,TOWKKBB3                    " & vbLf   '���o�敪
        sSql = sSql & "            ,TOWORKB3                    " & vbLf   '���o�H��R�[�h
        sSql = sSql & "            ,TOPLACB3                    " & vbLf   '���o���C���R�[�h
        sSql = sSql & "            ,CHGNB3                      " & vbLf   '�`���[�WNo
        sSql = sSql & "            ,EYYB3                       " & vbLf   '���ѓ��t(�N)
        sSql = sSql & "            ,EMMB3                       " & vbLf   '���ѓ��t(��)
        sSql = sSql & "            ,EDDB3                       " & vbLf   '���ѓ��t(��)
        sSql = sSql & "            ,ECYOKB3                     " & vbLf   '���敪
        sSql = sSql & "            ,EHHB3                       " & vbLf   '���ю���(��)
        sSql = sSql & "            ,EMIB3�@                     " & vbLf   '���ю���(��)
        sSql = sSql & "            ,MANB3                       " & vbLf   '�S����
        sSql = sSql & "            ,MANJB3                      " & vbLf   '�S���Җ�
        sSql = sSql & "            ,DENKB3                      " & vbLf   '�Z�x�敪
        sSql = sSql & "            ,DENSITYB3                   " & vbLf   '�Z�x�l
        sSql = sSql & "            ,GSNDFLGB3                   " & vbLf   '�������M�t���O
        sSql = sSql & "            ,HFLGB3                      " & vbLf   '�����t���O
        sSql = sSql & "            ,htkbnb3                     " & vbLf   '���i�敪
        sSql = sSql & "            ,plworkb3                    " & vbLf   '�g�p�\��H��
        sSql = sSql & "            ,mdensityb3                  " & vbLf   '���Z�x�l
        sSql = sSql & "            ,gsdayb3                     " & vbLf
        sSql = sSql & ")VALUES(                                 " & vbLf
        
    '*** UPDATE START T.TERAUCHI 2004/12/06 ���i���b�gNo��13���Ƃ���
    '    sSQL = sSQL & " '" & .BLOCKID & "'                      " & vbLf   '���������ԍ�
        sSql = sSql & " '" & .BLOCKID & "0" & "'                      " & vbLf  '���������ԍ�
    '*** UPDATE END   T.TERAUCHI 2004/12/06
        
        sSql = sSql & "," & iRenban & "                         " & vbLf   '�H���A��
        sSql = sSql & ",1                                       " & vbLf   '�����A��
        sSql = sSql & ",to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf '�o�^���t
        sSql = sSql & ",null                                    " & vbLf   '�C�����t
        sSql = sSql & ",null                                    " & vbLf   '���M���t
        sSql = sSql & ",' '                                     " & vbLf   '���M�敪
        sSql = sSql & ",'0'                                     " & vbLf   '�폜�敪
        sSql = sSql & ",'2'                                     " & vbLf   '�����敪
        sSql = sSql & ",'888'                                   " & vbLf   '������ރR�[�h
        sSql = sSql & ",' '                                     " & vbLf   '�������b�g�ԍ�
        sSql = sSql & ",' '                                     " & vbLf   '�ԍ��敪
        sSql = sSql & ",' '                                     " & vbLf   '�W�v�敪
        
    '*** UPDATE START T.TERAUCHI 2004/12/09 �H���ύX
    '    sSQL = sSQL & ",'" & Right(PROCD_RIMERUTO_UKEIRE, 4) & "'" & vbLf  '�H���R�[�h
        sSql = sSql & ",'" & Right(PROCD_KAKUAGE, 4) & "'" & vbLf  '�H���R�[�h  B320
    '*** UPDATE END   T.TERAUCHI 2004/12/09
        
        sSql = sSql & ",' '                                     " & vbLf   '���C���R�[�h
        sSql = sSql & "," & rec_xodc2.GNWC2 & "                 " & vbLf   '��ʎd�|�d��
        sSql = sSql & "," & rec_xodc2.GNWC2 & "                 " & vbLf   '��ʎd�|�d��
        sSql = sSql & ",0                                       " & vbLf   '���X�d��
        
    '*** UPDATE START T.TERAUCHI 2004/12/09 �H���ύX
    '    sSQL = sSQL & ",'" & Right(PROCD_RIMERUTO_UKEIRE, 4) & "'" & vbLf  '����H���R�[�h
        sSql = sSql & ",'" & Right(PROCD_KAKUAGE, 4) & "'" & vbLf  '����H���R�[�h�@B320
    '*** UPDATE END   T.TERAUCHI 2004/12/09
        
        sSql = sSql & ",'" & Right(PROCD_KOUNYU_TAN_KESSYOU, 4) & "'" & vbLf '���o�H���R�[�h('B410')
        sSql = sSql & ",' '                                     " & vbLf   '���o�敪
        sSql = sSql & ",'" & gsFactryCd & "'                    " & vbLf   '���o�H��R�[�h
        sSql = sSql & ",' '                                     " & vbLf   '���o���C���R�[�h
        sSql = sSql & ",' '                                     " & vbLf   '�`���[�WNo
        sSql = sSql & ",'" & sYear & "'                         " & vbLf   '���ѓ��t(�N)
        sSql = sSql & ",'" & sMonth & "'                        " & vbLf   '���ѓ��t(��)
        sSql = sSql & ",'" & sDay & "'                          " & vbLf   '���ѓ��t(��)
        sSql = sSql & ",'" & sCyoku & "'                        " & vbLf   '���敪
        sSql = sSql & ",'" & sHour & "'                         " & vbLf   '���ю���(��)
        sSql = sSql & ",'" & sMin & "'                          " & vbLf   '���ю���(��)
        sSql = sSql & ",'" & sUserID & "'                       " & vbLf   '�S����
        sSql = sSql & ",'" & sUserName & "'                     " & vbLf   '�S���Җ�
        sSql = sSql & ",' '                                     " & vbLf   '�Z�x�敪
        sSql = sSql & ",NULL                                    " & vbLf   '�Z�x�l
        sSql = sSql & ",'7'                                     " & vbLf   '�������M�t���O
        sSql = sSql & ",'0'                                     " & vbLf   '�����t���O
        sSql = sSql & ",'1'                                     " & vbLf   '���i�敪
        sSql = sSql & ",'" & gsFactryCd & "'                    " & vbLf   '�g�p�\��H��
        sSql = sSql & ",NULL                                    " & vbLf   '���Z�x
        sSql = sSql & ",NULL                                    " & vbLf '
        sSql = sSql & ")"
        
        If SqlExec2(sSql) = -1 Then
            GoTo proc_exit
        End If
        
    End With
    
    InsXODB3 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing

    '�I��
    gErr.Pop
    Exit Function

PROC_ERR:
    
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit

End Function

' @(f)
' �@�\      : SQL���l�ϊ��֐�
'
' �Ԃ�l    : <���͐��l> or NULL
'
' ������    : �ϊ��Ώې��l
'
' �@�\����  : �n���ꂽ���l��NULL�ł����"NULL"�������łȂ���΂��̂܂܏o�͂���
Private Function ConvNum(vinput) As String
    If IsNull(vinput) Or vinput = "NULL" Then
        vinput = ""
    End If
    
    If vinput = "" Then
        ConvNum = "NULL"
    Else
        ConvNum = vinput
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
Public Function DBDRV_SELECT_HOLD(pTblDispData As type_DBDRV_scmzc_fcmgc001f_Kensaku) As FUNCTION_RETURN

    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long
    Dim sCryNum As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_XSDC1_SQL.bas -- Function DBDRV_SELECT_HOLD"

    With pTblDispData
        sCryNum = Left(.BLOCKID, 9) & "000"
        ''SQL��g�ݗ��Ă�
        sql = "SELECT HLDCMNT FROM TBCMJ012 "
        sql = sql & " WHERE CRYNUM = '" & sCryNum & "'"
        sql = sql & " ORDER BY TRANCNT"
        '�f�[�^�𒊏o����
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        
        If rs Is Nothing Then
            DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        If rs.RecordCount > 0 Then
           rs.MoveLast
            If IsNull(rs("HLDCMNT")) = False Then .HLDCMNT = rs("HLDCMNT")
        End If
    End With
    rs.Close

    DBDRV_SELECT_HOLD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

PROC_ERR:
    '�G���[�n���h��
    DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

' 2007/09/18 SPK Tsutsumi Add Start
Public Function GetMukeCode() As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Long      '���R�[�h��
    Dim i  As Long
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "f_cmbc008_0.frm -- Function Getstaffauthority"
    
    GetMukeCode = FUNCTION_RETURN_FAILURE
    
    sql = "Select CODEA9,NAMEJA9 "
    sql = sql & "from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '20' "
    sql = sql & "and (CODEA9 = '14' "
    sql = sql & "or CODEA9 = '15' "
    sql = sql & "or CODEA9 = '16' "
'    sql = sql & "or CODEA9 = 'ALL') "
    sql = sql & "or CODEA9 = 'ZX' "         '08/07/01 ooba
    sql = sql & "or CODEA9 = 'ZZ') "        '08/07/01 ooba
    sql = sql & "order by CODEA9 "          '08/07/01 ooba

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If
    
    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim s_Mukesaki(recCnt)
    
    If recCnt = 0 Then
        Exit Function
    End If
    
    For i = 1 To recCnt
        With s_Mukesaki(i)
            If IsNull(rs.Fields("CODEA9")) = False Then
                .sMukeCode = rs.Fields("CODEA9")    ' ����R�[�h
            End If
            
            If IsNull(rs.Fields("NAMEJA9")) = False Then
                .sMukeName = rs.Fields("NAMEJA9")  ' ���於
            End If
        End With
        rs.MoveNext
    Next
    rs.Close

    GetMukeCode = FUNCTION_RETURN_SUCCESS
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

Public Function ChgMukesaki(sZuban As String) As String
    Dim lLp As Long
    Dim sBuf As String
    Dim rs As OraDynaset
    Dim sql As String
    Dim gsMuke4 As String
    Dim gsMuke5 As String
    Dim gsMuke6 As String
    Dim sCScode As String           '�ڋq���ށ@08/07/01 ooba
    Dim sTECHcode As String         'TECHXIV�i�ڋq���ށ@08/07/01 ooba
    
    sBuf = ""
    
    sql = "Select hinban,MAX(MNOREVNO), SUM(NVL(TRIM(E1.KFCTFLAG1),'')) FLAG1, SUM(NVL(TRIM(E1.KFCTFLAG2),'')) FLAG2, SUM(NVL(TRIM(E1.KFCTFLAG3),'')) FLAG3 "
    sql = sql & ", MAX(E1.KMGCSCOD) CSCODE "                '08/07/01 ooba
    sql = sql & "from TBCME001 E1 "
    sql = sql & "where E1.HINBAN = '" & Trim(sZuban) & "' "
    sql = sql & "and E1.OPECOND = '1' "
    sql = sql & "group by hinban"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    If IsNull(rs("FLAG1")) = False Then gsMuke4 = CStr(rs("FLAG1"))   '����S��
    If IsNull(rs("FLAG2")) = False Then gsMuke5 = CStr(rs("FLAG2"))   '����T��
    If IsNull(rs("FLAG3")) = False Then gsMuke6 = CStr(rs("FLAG3"))   '����U��
    sCScode = rs("CSCODE")                                            '�w�Ǘ��ڋq���ށ@08/07/01 ooba
    
    rs.Close

    'TECHXIV�i�ڋq���ގ擾
    sTECHcode = GetSSComboStrA9("X", "21", "CODEA9")
    
    For lLp = 1 To UBound(s_Mukesaki)
        Select Case lLp
            Case 1
                If gsMuke4 <> "" Then
                    sBuf = s_Mukesaki(lLp).sMukeCode
                    Exit For
                End If
            Case 2
                If gsMuke5 <> "" Then
                    sBuf = s_Mukesaki(lLp).sMukeCode
                    Exit For
                End If
            Case 3
                If gsMuke6 <> "" Then
                    sBuf = s_Mukesaki(lLp).sMukeCode
                    Exit For
                End If
            'TECHXIV�i�@08/07/01 ooba
            Case 4
                If gsMuke4 = "" And gsMuke5 = "" And gsMuke6 = "" Then
                    'TECHXIV�i�����@08/07/01 ooba
                    If InStr(1, sTECHcode, sCScode) > 0 Then
                        sBuf = s_Mukesaki(lLp).sMukeCode
                        Exit For
                    End If
                End If
            'Bar�o�וi�@08/07/01 ooba
            Case Else
                If gsMuke4 = "" And gsMuke5 = "" And gsMuke6 = "" Then
                    'TECHXIV�i�����@08/07/01 ooba
                    If InStr(1, sTECHcode, sCScode) > 0 Then
                    Else
                        sBuf = s_Mukesaki(lLp).sMukeCode
                        Exit For
                    End If
                End If
        End Select
    Next lLp
    
    If sBuf = "" Then
' 2007/10/10 SPK Tsutsumi Add Start
        '�S���E�T���E�U���ɉ����t���O�������Ă��Ȃ��ꍇ�ABar�o��
        ChgMukesaki = "ZZ"
'        f_cmbc008_1.lblMsg.Caption = "����擾�G���[ TBCME001"
' 2007/10/10 SPK Tsutsumi Add End
    Else
        ChgMukesaki = sBuf
    End If
            
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
' 2007/09/18 SPK Tsutsumi Add End
