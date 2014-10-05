Attribute VB_Name = "s_cmbc017_SQL"
Option Explicit

' ����������H���ѓ���

' ���i�d�l
Public Type typ_HinSpec2
    HIN As tFullHinban          ' �i��
    HSXTYPE As String * 1       ' �^�C�v
    HSXCDIR As String * 1       ' ����
    HSXD1CEN As Double          ' ���a
    
    HSXDPDIR As String * 2      ' �m�b�`�ʒu
    HSXDPMIN As Double          ' �m�b�`�p�x�̔���(Notch�ʒu�K�i ����)�@2009/09 SUMCO Akizuki
    HSXDPMAX As Double          ' �m�b�`�p�x�̔���(Notch�ʒu�K�i ���)�@2009/09 SUMCO Akizuki
    
    HSXDPAMN As Integer         ' �m�b�`�p�x(����)
    HSXDPAMX As Integer         ' �m�b�`�p�x(���)
    HSXDPACN As Integer         ' �m�b�`�p�x(���S) 2005/08
    
    HSXDWMIN As Double          ' �m�b�`��(����)
    HSXDWMAX As Double          ' �m�b�`��(���)
    
    HSXDDMIN As Double          ' �m�b�`�[��(����)
    HSXDDMAX As Double          ' �m�b�`�[��(���)
    HSXDDCEN As Double          ' �m�b�`�[��(���S)  2005/08
End Type

' �҂��ꗗ
Public Type typ_DispData
    CRYNUM As String
    HIN As String
    DIAK As String
    GNDAY As Date
    GNL As Double
    GNW As Double
    PRIORITY As String
    PUPTN As String
    NOUKI As Date
    MUKE As String
    HLDCAUSE As String
    HOLDKT As String
    BIKOU As String
    HLDCMNT As String
    HLDTRCLS As String
    PUHINB As String   '2005/10
    XTALCA As String   '2005/10
    RPCRYNUMCA As String   '2005/10
    NEWKNTCA As String   '2005/11
    KIKBN    As String  '�����ʋ敪 2006/11/14 SETsw kubota
    PLANTCATCA As String    '���� 2007/08/21 SPK Tsutsumi Add
    DPDIR   As String   '�m�b�`�ʒu���� 2008/01/09
    AGRSTATUS As String             ' ���F�m�F�敪 add SETkimizuka
    STOP    As String               ' ��~ add SETkimizuka
    CAUSE   As String               ' ��~���R add SETkimizuka
    PRINTNO As String               ' ��s�]�� add SETkimizuka
End Type

' 2007/08/17 SPK Tsutsumi Add Start
Public Type typ_Mukesaki
    sMukeCode As String     '' ����R�[�h
    sMukeName As String     '' ���於
End Type

Public s_Mukesaki() As typ_Mukesaki
' 2007/08/17 SPK Tsutsumi Add End

'' �X�g�b�J�Ή� 2006/11/08 SETsw J.W -->
Public gsGrTim As String     ' ���펞��
Public gsNchLength As String '�m�b�`����
'' �X�g�b�J�Ή� 2006/11/08 SETsw J.W -->


'�T�v      :����������H���ѓ��͗p �����ԍ����͎��c�a�h���C�o
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'�@�@      :sCryNum�@�@�@,I  ,String         �@,�����ԍ�
'�@�@      :pCryInf�@�@�@,O  ,typ_TBCME037   �@,�������
'�@�@      :pHinDsn�@�@�@,O  ,typ_TBCME039   �@,�i�Ԑ݌v
'�@�@      :pCutIns�@�@�@,O  ,typ_TBCME045   �@,�ؒf�w��
'�@�@      :pProcBR�@�@�@,O  ,typ_TBCMI001   �@,���H���o����
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'�@�@      :�߂�l       ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmic001c_Disp(ByVal sCryNum As String, ByVal sBlockId As String, pCryInf As typ_TBCME037, _
                                           pHinDsn() As typ_TBCME039, pCutIns() As typ_TBCME045, _
                                           pProcBR As typ_TBCMI001, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpProcBR() As typ_TBCMI001
    Dim rs  As OraDynaset
    Dim rs2 As OraDynaset   ' add 2006/11/09 SETsw J.W
    Dim sql As String
    Dim sDbName As String
    Dim recCnt As Long
    Dim i As Long
    Dim sans As String
    Dim sNowproc As String
    '2004.09.08 Y.K �R�t���ύX
    Dim sSijiNo As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc017_SQL.bas -- Function DBDRV_scmzc_fcmic001c_Disp"
    sErrMsg = ""

    '2004.09.08 Y.K �R�t���ύX  <=== START
    sDbName = "XSDC1"
    sSijiNo = ""
    sSijiNo = F_Get_SijiNoGet(sCryNum)
    '2004.09.08 Y.K �R�t���ύX  == > END

    '' �������̎擾(s_cmzcTBCME037_SQL.bas ���K�v)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)

    '' �H���`�F�b�N
    If DBDRV_get_xGR(sCryNum, sans) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET")
        Exit Function
    End If
    If sans = "A" Then
        'AGR�̏ꍇ
        If pCryInf.PROCCD <> PROCD_KENNSAKU_KAKOU Then
            sErrMsg = GetMsgStr("EPRC0")
            DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Else
        'MGR�̏ꍇ
        If GetTBCME040_NOWPROC(sBlockId, sNowproc) = FUNCTION_RETURN_FAILURE Then
            sDbName = "E040"
            sErrMsg = GetMsgStr("ENG11", sDbName)
            DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

'        If (pCryInf.PROCCD <> PROCD_KESSYOU_SOUGOUHANTEI) And (pCryInf.PROCCD <> PROCD_SETUDAN) Then
        If (sNowproc <> PROCD_KENNSAKU_KAKOU) Then    '�d�|����H�������ύX�@2002/11/28
            sErrMsg = GetMsgStr("EPRC2")
            DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    '' �i�Ԑ݌v�̎擾(s_cmzcTBCME039_SQL.bas ���K�v)
    sDbName = "E039"
    '2004.09.08 Y.K �R�t���ύX
'    sql = " where substr(CRYNUM,1,7)='" & Left(sCrynum, 7) & "' order by INGOTPOS"
    sql = " where substr(CRYNUM,1,9)='" & sSijiNo & "' order by INGOTPOS"
    If DBDRV_GetTBCME039(pHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinDsn) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �ؒf�w���̎擾 ----> 2005/11 XSDCA�ɕύX
    'sDbName = "E045"
    'sql = "select "
    'sql = sql & "CRYNUM, INGOTPOS, TRANCNT "
    'sql = sql & " from TBCME045 T1"
    'sql = sql & " where CRYNUM='" & sCryNum & "'"
    'sql = sql & " and TRANCNT=any(select max(TRANCNT) from TBCME045 T2 where CRYNUM='" & sCryNum & "'"
    'sql = sql & " and T1.INGOTPOS=T2.INGOTPOS ) "
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCnt = rs.RecordCount
    'If recCnt = 0 Then
    '    rs.Close
    '    sErrMsg = GetMsgStr("EGET2", sDbName)
    '    DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
    '    GoTo proc_exit
    'End If

    'ReDim pCutIns(recCnt)
    'For i = 1 To recCnt
    '    With pCutIns(i)
    '        .CRYNUM = rs("CRYNUM")          ' �����ԍ�
    '        .INGOTPOS = rs("INGOTPOS")      ' �������J�n�ʒu
    '        .TRANCNT = rs("TRANCNT")        ' ������
    '    End With
    '    rs.MoveNext
    'Next i
    'rs.Close

    'For i = 1 To recCnt
    '    With pCutIns(i)
    '        sql = "select "
    '        sql = sql & "LENGTH, HINBAN, REVNUM, FACTORY, OPECOND"
    '        sql = sql & " from TBCME045"
    '        sql = sql & " where CRYNUM='" & .CRYNUM & "'"
    '        sql = sql & " and INGOTPOS=" & .INGOTPOS
    '        sql = sql & " and TRANCNT=" & .TRANCNT
    '        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '        If rs.RecordCount = 0 Then
    '            rs.Close
    '            sErrMsg = GetMsgStr("EGET2", sDbName)
    '            DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
    '            GoTo proc_exit
    '        End If
    '        .LENGTH = rs("LENGTH")      ' ����
    '        .hinban = rs("HINBAN")      ' �i��
    '        .REVNUM = rs("REVNUM")      ' ���i�ԍ������ԍ�
    '        .factory = rs("FACTORY")    ' �H��
    '        .opecond = rs("OPECOND")    ' ���Ə���
    '    End With
    '    rs.Close
    'Next i
    '�f���X�V SPT�p���э쐬���@�ύX 2006/04/17 SMP���c

    sDbName = "XSDCZ"
    sql = "select DISTINCT "
    sql = sql & "HINBCZ, REVNUMCZ, FACTORYCZ, OPECZ "
    sql = sql & " from XSDCZ "
    sql = sql & " where RPCRYNUMCZ ='" & sBlockId & "'"
    sql = sql & " and GNWKNTCZ = 'CC400'"

'    sDbName = "XSDCA"
'    sql = "select DISTINCT "
'    sql = sql & "HINBCA, REVNUMCA, FACTORYCA, OPECA "
'    sql = sql & " from XSDCA "
'    sql = sql & " where RPCRYNUMCA ='" & sCrynum & "'"
'    sql = sql & " and GNWKNTCA = 'CC400'"

    '�f���X�V SPT�p���э쐬���@�ύX 2006/04/17 SMP���c
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ReDim pCutIns(recCnt)
    For i = 1 To recCnt
        With pCutIns(i)
        '�f���X�V SPT�p���э쐬���@�ύX 2006/04/17 SMP���c
            .hinban = rs("HINBCZ")      ' �i��
            .REVNUM = rs("REVNUMCZ")    ' ���i�ԍ������ԍ�
            .factory = rs("FACTORYCZ")  ' �H��
            .opecond = rs("OPECZ")      ' ���Ə���

'            .hinban = rs("HINBCA")      ' �i��
'            .REVNUM = rs("REVNUMCA")      ' ���i�ԍ������ԍ�
'            .factory = rs("FACTORYCA")    ' �H��
'            .opecond = rs("OPECA")    ' ���Ə���
        '�f���X�V SPT�p���э쐬���@�ύX 2006/04/17 SMP���c
        End With
        rs.MoveNext
    Next i
    rs.Close
    '' ���H���o���т̎擾(s_cmzcTBCMI001_SQL.bas ���K�v)
    sDbName = "I001"
    sql = " where CRYNUM='" & sCryNum & "'"
    sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT)"
    sql = sql & " from TBCMI001 where CRYNUM='" & sCryNum & "')"
    If DBDRV_GetTBCMI001(tmpProcBR(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpProcBR) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pProcBR = tmpProcBR(1)

    '' �X�g�b�J�Ή� 2006/11/09 SETsw J.W -->
    '' ���펞�ԥ�m�b�`�̎擾
    gsNchLength = ""
    gsGrTim = ""

    sql = "      select UPDDATE"
    sql = sql & "     , -1 NOTCH"  '<== �Y���J�������������ߋ� (2006/11/09 J.W)
    sql = sql & "     , TO_CHAR(CYGRTIM,'FM000000') GRTIM"
    sql = sql & "  from TBCMF002"
    sql = sql & " where INGOTNO='" & sCryNum & "'"
    sql = sql & "   and TRANCNT=(select MAX(TRANCNT) from TBCMF002 where INGOTNO='" & sCryNum & "')" & vbCrLf
    sql = sql & " union "
    sql = sql & "select NVL(UPDDATE,REGDATE) UPDDATE"   '�V�e�[�u���X�V���t��NULL�̏ꍇ�A�o�^���t
    sql = sql & "     , TRWLEN NOTCH"
    sql = sql & "     , TO_CHAR(CYGRTIM,'FM000000') GRTIM"
    sql = sql & "  from TBCMF010"
    sql = sql & " where CRYNUM='" & sBlockId & "'"
    sql = sql & "   and PROCNUM=(select MAX(PROCNUM) from TBCMF010 where CRYNUM='" & sBlockId & "')"
    sql = sql & " order by UPDDATE desc"

    Set rs2 = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs2.RecordCount > 0 Then
        gsNchLength = NulltoStr(rs2("NOTCH").Value)
        '' �m�b�`�������}�C�i�X(TBCMF002����擾�����ꍇ�͋󗓂ɂ���) 2006/11/10 SETsw J.W
        If (val(gsNchLength) < 0) Then
            gsNchLength = ""
        End If
        gsGrTim = Format(NulltoStr(rs2("GRTIM").Value), "@@:@@:@@")
        If Left(gsGrTim, 1) = "0" Then
            gsGrTim = Mid(gsGrTim, 2)
        End If
    End If
    rs2.Close
    '' �X�g�b�J�Ή� 2006/11/09 SETsw J.W <--

    DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("EGET2", sDbName)
    DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2004.09.08�@Y.K�@�R�t���ύX
'�w�茋���ԍ��̎w��No�擾����
'�w�r�c�b�P�̎w�����擾����
'�A���A�擾�ł��Ȃ��ꍇ�́A�����ԍ��V���{�f�O�f�{�����ԍ��X����Ԃ�
Private Function F_Get_SijiNoGet(sCryNum As String) As String
  Dim sSql As String
  Dim rs As OraDynaset    'RecordSet

    sSql = ""
    sSql = sSql & "SELECT"
    sSql = sSql & "  hisijiC1 "
    sSql = sSql & "FROM"
    sSql = sSql & "  XSDC1 C1 "
    sSql = sSql & "WHERE"
    sSql = sSql & "  substr(C1.XTALC1,1,9) = '" & Mid(sCryNum, 1, 9) & "'"

    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)

    If (rs.RecordCount = 1) Then
        If (IsNull(rs.Fields("hisijiC1")) = False) Then
            F_Get_SijiNoGet = rs.Fields("hisijiC1")
        Else
            F_Get_SijiNoGet = Mid(sCryNum, 1, 7) & "0" & Mid(sCryNum, 9, 1)
        End If
    Else
        F_Get_SijiNoGet = Mid(sCryNum, 1, 7) & "0" & Mid(sCryNum, 9, 1)
    End If

End Function


'�T�v      :����������H���ѓ��͗p ���i�d�l��p�c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'�@�@      :pHinSpec�@�@�@,IO ,typ_HinSpec2   �@,���i�d�l
'�@�@      :�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmic001c_GetSpec(pHinSpec As typ_HinSpec2) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc017_SQL.bas -- Function DBDRV_scmzc_fcmic001c_GetSpec"

    '' ���i�d�l�̎擾
    sql = "select "
    sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXDPDIR, HSXDPMIN, HSXDPMAX, "
    sql = sql & "HSXDWMIN, HSXDWMAX, HSXDDMIN, HSXDDMAX, HSXDDCEN, HSXDPACN "  '2009/09
    sql = sql & " from TBCME018"
    sql = sql & " where HINBAN='" & pHinSpec.HIN.hinban & "'"
    sql = sql & " and MNOREVNO=" & pHinSpec.HIN.mnorevno
    sql = sql & " and FACTORY='" & pHinSpec.HIN.factory & "'"
    sql = sql & " and OPECOND='" & pHinSpec.HIN.opecond & "'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        DBDRV_scmzc_fcmic001c_GetSpec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
'NULL�Ή� ----- START ----- 2003/12/10
    With pHinSpec
        .HSXTYPE = rs("HSXTYPE")                          ' �^�C�v
        .HSXCDIR = rs("HSXCDIR")                          ' ����
        .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))          ' ���a
        .HSXDPDIR = rs("HSXDPDIR")                        ' �m�b�`�ʒu
        .HSXDWMIN = fncNullCheck(rs("HSXDWMIN"))          ' �m�b�`��(����)
        .HSXDWMAX = fncNullCheck(rs("HSXDWMAX"))          ' �m�b�`��(���)
        .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))          ' �m�b�`�[��(����)
        .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))          ' �m�b�`�[��(���)
        .HSXDDCEN = fncNullCheck(rs("HSXDDCEN"))          ' �m�b�`�[��(���S) 2005/08
        .HSXDPACN = fncNullCheck(rs("HSXDPACN"))          ' �m�b�`�p�x(���S) 2005/08
        
        '�l�Ɂu-1�v�����邽�߁ANull�̏ꍇ��[999]��Ԃ�     ' �m�b�`�ʒu(����) 2009/09 Akizuki
        If IsNull(rs("HSXDPMIN")) Then
            .HSXDPMIN = 999
        Else
            .HSXDPMIN = rs("HSXDPMIN")
        End If
        
        '�l�Ɂu-1�v�����邽�߁ANull�̏ꍇ��[999]��Ԃ�   �@' �m�b�`�ʒu(���) 2009/09 Akizuki
        If IsNull(rs("HSXDPMAX")) Then
            .HSXDPMAX = 999
        Else
            .HSXDPMAX = rs("HSXDPMAX")
        End If
    End With
    
    rs.Close
'NULL�Ή� -----  END  ----- 2003/12/10

    DBDRV_scmzc_fcmic001c_GetSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmic001c_GetSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :����������H���ѓ��͗p ���s���c�a�h���C�o
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:sCryNum�@�@�@,I  ,String         �@,�����ԍ�
'      �@�@:pPlshPR�@�@�@,I  ,typ_TBCMI002   �@,������H����
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function DBDRV_scmzc_fcmic001c_Exec(sCryNum As String, pPlshPR As typ_TBCMI002, sErrMsg As String) As FUNCTION_RETURN

    Dim sql As String
    Dim sDbName As String
    Dim sans As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc017_SQL.bas -- Function DBDRV_scmzc_fcmic001c_Exec"
    sErrMsg = ""

    '�H���R�[�h�ݒ胍�W�b�N����   2002/11/27 tuku START
    'AGR MGR �̃`�F�b�N
    If DBDRV_get_xGR(sCryNum, sans) = FUNCTION_RETURN_FAILURE Then
        sDbName = "I001"
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        Exit Function
    End If
    If sans = "A" Then
        'AGR�̏ꍇ�������iTBCME037)�̂ݍX�V
        '' �������̍X�V
        sDbName = "E037"
        sql = "update TBCME037 set "
        sql = sql & "KRPROCCD='" & MGPRCD_SETUDAN & "', "
        sql = sql & "PROCCD='" & nextCd & "', "
        sql = sql & "LPKRPROCCD='" & MGPRCD_KENNSAKU_KAKOU & "', "
        sql = sql & "LASTPASS='" & nowCd & "', "
        sql = sql & "DIAMETER=" & pPlshPR.DMTOP1 & ", "
        sql = sql & "UPDDATE=sysdate, "
        sql = sql & "SENDFLAG='0'"
        sql = sql & " where CRYNUM='" & sCryNum & "'"
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    ElseIf sans = "M" Then
        'MGR�̏ꍇ�������iTBCME037)�ƃu���b�N�Ǘ��iTBCME040)���X�V
        '' �������̍X�V
        sDbName = "E037"
        sql = "update TBCME037 set "
        sql = sql & "KRPROCCD='" & MGPRCD_KESSYOU_SOUGOUHANTEI & "', "
        sql = sql & "PROCCD='" & nextCd & "', "
        sql = sql & "LPKRPROCCD='" & MGPRCD_KENNSAKU_KAKOU & "', "
        sql = sql & "LASTPASS='" & nowCd & "', "
        sql = sql & "DIAMETER=" & pPlshPR.DMTOP1 & ", "
        sql = sql & "UPDDATE=sysdate, "
        sql = sql & "SENDFLAG='0'"
        sql = sql & " where CRYNUM='" & sCryNum & "'"
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        ''�u���b�N�Ǘ��̍X�V
        '' �u���b�N�Ǘ��e�[�u���̍X�V
        sDbName = "E040"
        sql = "update TBCME040 set "
        sql = sql & "KRPROCCD='" & MGPRCD_KESSYOU_SOUGOUHANTEI & "', "      ' ���݊Ǘ��H��
        sql = sql & "NOWPROC='" & nextCd & "', "                        ' ���ݍH��
        sql = sql & "LPKRPROCCD='" & MGPRCD_KENNSAKU_KAKOU & "', "                ' �ŏI�ʉߊǗ��H��
        sql = sql & "LASTPASS='" & nowCd & "', "                        ' �ŏI�ʉߍH��
        sql = sql & "UPDDATE=sysdate "                                 ' �X�V���t
        sql = sql & "where CRYNUM='" & sCryNum & "' "
        sql = sql & "and  BLOCKID='" & pPlshPR.CRYNUM & "' "
        sql = sql & "and INGOTPOS=" & pPlshPR.INGOTPOS
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    '                           2002/11/27 tuku END



    '' ������H���т̑}��
    sDbName = "I002"
    With pPlshPR
        sql = "insert into TBCMI002 "
        sql = sql & "(CRYNUM, TRANCNT, KRPROCCD, PROCCODE, "
        sql = sql & "DMTOP1, DMTOP2, DMTAIL1, DMTAIL2, NCHPOS, NCHDPTH, NCHWIDTH, "
        sql = sql & "BDLNTOP, BDCDTOP, BDLNTAIL, BDCDTAIL, "
        sql = sql & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, INGOTPOS, LENGTH, "
        sql = sql & "GOUKI, NCHWTAIL ,BLOCKID, "             '2006/02/01 tuku
        sql = sql & "NCHLENGTH, CYGRTIM, NCHANGLE )"          ' �X�g�b�J�Ή� 2006/11/09 SETsw J.W
        sql = sql & " select '"
        sql = sql & sCryNum & "', "
        sql = sql & "nvl(max(TRANCNT),0)+1, '"
        sql = sql & .KRPROCCD & "', '"
        sql = sql & .PROCCODE & "', "
        sql = sql & .DMTOP1 & ", "
        sql = sql & .DMTOP2 & ", "
        sql = sql & .DMTAIL1 & ", "
        sql = sql & .DMTAIL2 & ", '"
        sql = sql & .NCHPOS & "', "
        sql = sql & .NCHDPTH & ", "
        sql = sql & .NCHWIDTH & ", "
        sql = sql & .BDLNTOP & ", '"
        sql = sql & .BDCDTOP & "', "
        sql = sql & .BDLNTAIL & ", '"
        sql = sql & .BDCDTAIL & "', '"
        sql = sql & .TSTAFFID & "', "
        sql = sql & "sysdate, '"
        sql = sql & .KSTAFFID & "', "
        sql = sql & "sysdate, "
        sql = sql & "'0', "
        sql = sql & "sysdate,"
        sql = sql & .INGOTPOS & ", "
        sql = sql & .LENGTH & " ,  "
        sql = sql & .GOUKI & " , "                      '2003/06/12 osawa ���@�ǉ�
        sql = sql & .NCHWTAIL & ", '"                  '2004/05/25
        sql = sql & .BLOCKID & "',"                    '2006/02/01 tuku ��ۯ�ID�ǉ�
        sql = sql & .NCHLENGTH & ", "                  ' �X�g�b�J�Ή� 2006/11/09 SETsw J.W
        sql = sql & .CYGRTIM & ", "                    ' �X�g�b�J�Ή� 2006/11/09 SETsw J.W
        sql = sql & .NCHANGLE                          ' 2009/09 SUMCO Akizuki Notch�p�x�ǉ�"
        sql = sql & " from TBCMI002"
        sql = sql & " where CRYNUM='" & sCryNum & "' and INGOTPOS=" & .INGOTPOS
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
    DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'�T�v      :INGOTPOS,LENGTH�̎擾�h���C�o
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :BLOCKID        ,I   ,String            ,�����ԍ�or�u���b�NID
'          :iIngotpos      ,O   ,Integer           ,�������J�n�ʒu
'          :iLength        ,O   ,Integer           ,����
'      �@�@:�߂�l          , O  , FUNCTION_RETURN�@, �ǂݍ��݂̐���
'����      :
'����      :2002/04/17 ���� �M�� �쐬
Public Function scmzc_getIngotposLength(BLOCKID As String, iIngotpos As Integer, iLength As Integer) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    Dim AGRFlag As Boolean
    Dim Ans As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc017_SQL.bas -- Function scmzc_getIngotposLength"
    scmzc_getIngotposLength = FUNCTION_RETURN_FAILURE

    '�����グ�����̏ꍇ
    '���H�����o�����т���AGR��MGR�������߂�
    If DBDRV_get_xGR(BLOCKID, Ans) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    AGRFlag = (Trim(Ans) = "A")

    If AGRFlag Then
        'AGR�̏ꍇ
        'INGOTPOS=0�ŉ��H���т�����т����߂�
        sql = "select UPLENGTH from TBCMI001 "
        sql = sql & "where CRYNUM = '" & Left(BLOCKID, 9) & "000" & "' and "
        sql = sql & "TRANCNT = any(select max(TRANCNT) from TBCMI001 where CRYNUM = '" & Left(BLOCKID, 9) & "000" & "')"

        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            scmzc_getIngotposLength = FUNCTION_RETURN_SUCCESS
            GoTo proc_exit
        End If
        iIngotpos = 0
        iLength = rs("UPLENGTH")
        rs.Close
    Else
        'MGR�̏ꍇ
        '�u���b�N�Ǘ�����INGOTPOS�����߂�
        '���̃u���b�N�̏���ؒf����INGOTPOS�����߂�
        sql = "select INGOTPOS,LENGTH from TBCME040 "
        sql = sql & "where CRYNUM = '" & Left(BLOCKID, 9) & "000" & "' and "
        sql = sql & "BLOCKID = '" & BLOCKID & "' "
        sql = sql & "order by INGOTPOS"

        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            GoTo proc_exit
        End If
        iIngotpos = rs("INGOTPOS")
        iLength = rs("LENGTH")
        rs.Close

    End If

    scmzc_getIngotposLength = FUNCTION_RETURN_SUCCESS
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getIngotposLength = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME037�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME037 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcF_TBCME037_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME037(records() As typ_TBCME037, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME037"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME037 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .DELCLS = rs("DELCLS")           ' �폜�敪
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCD = rs("PROCCD")           ' �H���R�[�h
            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
            .RPHINBAN = rs("RPHINBAN")       ' �˂炢�i��
            .RPREVNUM = rs("RPREVNUM")       ' �˂炢�i�Ԑ��i�ԍ������ԍ�
            .RPFACT = rs("RPFACT")           ' �˂炢�i�ԍH��
            .RPOPCOND = rs("RPOPCOND")       ' �˂炢�i�ԑ��Ə���
            .PRODCOND = rs("PRODCOND")       ' �������
            .PGID = rs("PGID")               ' �o�f�|�h�c
            .UPLENGTH = rs("UPLENGTH")       ' ���グ����
            .TOPLENG = rs("TOPLENG")         ' �s�n�o����
            .BODYLENG = rs("BODYLENG")       ' ��������
            .BOTLENG = rs("BOTLENG")         ' �a�n�s����
            .FREELENG = rs("FREELENG")       ' �t���[��
            .DIAMETER = rs("DIAMETER")       ' ���a
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .SEED = rs("SEED")               ' �V�[�h
            .ADDDPCLS = rs("ADDDPCLS")       ' �ǉ��h�[�v���
            .ADDDPPOS = rs("ADDDPPOS")       ' �ǉ��h�[�v�ʒu
            .ADDDPVAL = rs("ADDDPVAL")       ' �ǉ��h�[�v��
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME037 = FUNCTION_RETURN_SUCCESS
End Function




'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME039�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME039 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcF_TBCME039_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME039(records() As typ_TBCME039, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACT, OPCOND, LENGTH, USECLASS, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME039"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME039 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' �����ԍ�
            .FACT = rs("FACT")               ' �H��
            .OPCOND = rs("OPCOND")           ' ���Ə���
            .LENGTH = rs("LENGTH")           ' ����
            .USECLASS = rs("USECLASS")       ' �g�p�敪
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME039 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMI001�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMI001 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMI001_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMI001(records() As typ_TBCMI001, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, TRANCNT, KRPROCCD, PROCCODE, UPLENGTH, FREELENG, UPWEIGHT, SEED, PRCMCN, TSTAFFID, REGDATE," & _
              " KSTAFFID, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMI001"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMI001 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .TRANCNT = rs("TRANCNT")         ' ������
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .UPLENGTH = rs("UPLENGTH")       ' ���グ����
            .FREELENG = rs("FREELENG")       ' �t���[��
            .UPWEIGHT = rs("UPWEIGHT")       ' ���グ�d��
            .SEED = rs("SEED")               ' �V�[�h
            .PRCMCN = rs("PRCMCN")           ' ����@
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMI001 = FUNCTION_RETURN_SUCCESS
End Function

'
'�T�v    : TBCME040���t�B�[���h�lLENGTH�̎擾
'���Ұ�  :�ϐ���        ,IO  ,�^                                     ,����
'
'        :��ؒl         ,O   ,FUNCTION_RETURN                        ,�ǂݍ��ݐ���
'����    :
'����    :2002.8 �ǉ� H.Kakizawa    2005/11 C2�ɕύX
Public Function GetTBCME040_NOWPROC(ByVal pBlockid As String, pNowproc As String) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    '�����l
    GetTBCME040_NOWPROC = FUNCTION_RETURN_FAILURE

    '������т̉��H�敪���擾
    sql = ""
    'sql = sql & "select NOWPROC from TBCME040 "
    'sql = sql & "where BLOCKID = '" & pBlockid & "' "
    sql = sql & "select GNWKNTC2 from XSDC2 "
    sql = sql & "where CRYNUMC2 = '" & pBlockid & "' "

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP) '�f�[�^�𒊏o����

    If rs.RecordCount = 0 Then '���R�[�h���Ȃ��ꍇ�͔�
        rs.Close
        Exit Function
    Else
        'pNowproc = rs.Fields("NOWPROC")
        pNowproc = rs.Fields("GNWKNTC2")
        rs.Close
    End If

    GetTBCME040_NOWPROC = FUNCTION_RETURN_SUCCESS

End Function

'�T�v      :���H���o�ꗗ�p ��ʕ\�����c�a�h���C�o
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pCutMap�@�@�@,O  ,typ_CutMap     �@,�ؒf�w���ꗗ
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmic001d_Disp(pDispData() As typ_DispData, pCrynum As String, pHinb As String) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Long
    Dim i As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc017_SQL.bas -- Function DBDRV_scmzc_fcmic001d_Disp"
    sql = ""
    'sql = sql & "SELECT PUHINBC1, PUPTNC1, DIA1C1, XTALCA, CRYNUMCA, HINBCA||LTRIM(to_char(REVNUMCA,'00'))||FACTORYCA||OPECA as HINBAN12, GNLCA, GNWCA, GNDAYCA "
    sql = sql & "SELECT PUHINBC1, PUPTNC1, DIA1C1, XTALCA, CRYNUMCA, HINBCA||LTRIM(to_char(REVNUMCA,'00'))||FACTORYCA||OPECA as HINBAN12,  "
''���폜 START SPT�p���э쐬���@�ύX 2006/05/18 SMP-OKAMOTO
    sql = sql & " GNLCA, GNWCA, GNDAYCA, CRYNUMCA as RPCRYNUMCA, NEWKNTCA "
'    sql = sql & " GNLCA, GNWCA, GNDAYCA, RPCRYNUMCA, NEWKNTCA "
''���폜 END   SPT�p���э쐬���@�ύX 2006/05/18 SMP-OKAMOTO
    sql = sql & "  ,HOLDBC2, HOLDCC2, HOLDKTC2 "
    sql = sql & "  ,PLANTCATCA"     ' 2007/08/21 SPK Tsutsumi Add
    sql = sql & " , HSXDPDIR"       ' 2008/01/09 �m�b�`�ʒu����
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/29
    ' ������~���ڒǉ� add SETkimizuka Start  09/03/25
    'sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUS "
    'sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
    'sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSE "
    'sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNO "
    ' ������~���ڒǉ� add SETkimizuka End    09/03/25
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUS "
    sql = sql & " , NVL(TO_CHAR(Y4.STOPY4),' ') as STOP "
    sql = sql & " , DECODE(TRIM(Y4.CAUSEY4),NULL,' ',TRIM(Y4.CAUSEY4) || ':' || NAMEJA9) as CAUSE "
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNO "
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/29
    sql = sql & " from XSDC1, XSDCA, XSDC2 "
    sql = sql & " �@�@,TBCME018 "   ' 2008/01/09 �m�b�`�ʒu����
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/29
    sql = sql & "    ,XODY3,XODY4 Y4,KODA9  "
    ' ������~���ڒǉ� add SETkimizuka Start  09/03/25
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
    ' ������~���ڒǉ� add SETkimizuka End  09/03/25
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/29
    sql = sql & " where XTALC1 = XTALC2 "
    sql = sql & " AND CRYNUMCA = CRYNUMC2 "
    sql = sql & " AND CRYNUMCA LIKE '" & pCrynum & "%'"
    sql = sql & " AND GNWKNTCA = 'CC400'  "
    sql = sql & " AND LIVKCA = '0'  "
    sql = sql & " AND HINBCA||LTRIM(to_char(REVNUMCA,'00'))||FACTORYCA||OPECA  LIKE '" & pHinb & "%'"
    sql = sql & " AND HINBCA = HINBAN "         ' 2008/01/09 �m�b�`�ʒu����
    sql = sql & " AND REVNUMCA = MNOREVNO "     ' 2008/01/09 �m�b�`�ʒu����
    sql = sql & " AND FACTORYCA = FACTORY "     ' 2008/01/09 �m�b�`�ʒu����
    sql = sql & " AND OPECA = OPECOND "         ' 2008/01/09 �m�b�`�ʒu����
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/29
    'sql = sql & " AND CRYNUMCA    = Y4.XTALNO(+) "            'add 09/03/25 SETkimizuka
    sql = sql & " AND CRYNUMCA = XTALNOY3(+) "
    sql = sql & " AND LIVKY3(+) = '0' "
    sql = sql & " AND LIVKY4(+) = '0' "
    sql = sql & " AND XTALNOY3 = XTALNOY4(+) "
    sql = sql & " AND RCNTY3 = RCNTY4(+) "
    sql = sql & " AND SYSCA9(+) = 'X' AND SHUCA9(+) = '30' AND CAUSEY4 = CODEA9(+) "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/29
    sql = sql & " order by CRYNUMCA,INPOSCA "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
    End If

    ReDim pDispData(recCnt)
    For i = 1 To recCnt
        With pDispData(i)
            .CRYNUM = rs("CRYNUMCA")  '
            .HIN = rs("HINBAN12")
            .PUPTN = rs("PUPTNC1")
            .GNL = rs("GNLCA")
            .GNW = rs("GNWCA")
            .GNDAY = rs("GNDAYCA")
            .DIAK = rs("DIA1C1")
            .HLDCAUSE = rs("HOLDCC2")
            .HLDTRCLS = rs("HOLDBC2")
            If IsNull(rs("HOLDKTC2")) = False Then .HOLDKT = rs("HOLDKTC2")
            .PUHINB = rs("PUHINBC1")  '2005/10
            .XTALCA = rs("XTALCA")
            If IsNull(rs("RPCRYNUMCA")) = False Then .RPCRYNUMCA = rs("RPCRYNUMCA") '2005/10
            .NEWKNTCA = rs("NEWKNTCA")
            If IsNull(rs("PLANTCATCA")) = False Then .MUKE = rs("PLANTCATCA")  ' 2007/08/21 SPK Tsutsumi Add
            .DPDIR = rs("HSXDPDIR")     '2008/01/09  �m�b�`�ʒu����
            ' ������~���ڒǉ� add SETkimizuka Start  09/03/25
            ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/29
            '.STOP = rs("STOP")                   '��~�敪
            '.AGRSTATUS = rs("AGRSTATUS")       '���F�m�F�敪
            'If Trim(rs("CAUSE")) <> "" Then
            '    .CAUSE = rs("CAUSE") & vbTab       '��~���R
            'End If
            If rs("STOP") <> "2" And rs("WKKTY4") = "CC400" Then
                .STOP = rs("STOP")                   '��~�敪
                .AGRSTATUS = rs("AGRSTATUS")       '���F�m�F�敪
                If Trim(rs("CAUSE")) <> "" Then
                    .CAUSE = rs("CAUSE") & vbTab       '��~���R
                End If
            End If
            ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/29
            If Trim(rs("PRINTNO")) <> "" Then
                .PRINTNO = rs("PRINTNO") & vbTab       '��s�]��
            End If
            ' ������~���ڒǉ� add SETkimizuka End    09/03/25
        End With
        rs.MoveNext
    Next i
    rs.Close
    DBDRV_scmzc_fcmic001d_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmic001d_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

Public Function DBDRV_SELECT_HOLD(pTblDispData As typ_DispData) As FUNCTION_RETURN

    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long
    Dim sCryNum As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc017_SQL.bas -- Function DBDRV_SELECT_HOLD"

    With pTblDispData

        sCryNum = Left(.CRYNUM, 9) & "000"
        ''SQL��g�ݗ��Ă�
        sql = "SELECT HLDCMNT FROM TBCMJ012 "
        sql = sql & " WHERE CRYNUM = '" & sCryNum & "'"
        'sql = sql & " AND   XTALC2 = CRYNUM   "
        'sql = sql & " AND   INPOSC2 = INGOTPOS   "
        'sql = sql & " AND TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ012 WHERE CRYNUM = '" & pCrynum & "')"
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

proc_err:
    '�G���[�n���h��
    DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
Public Function DBDRV_SELECT_BLOCK(pBlockData() As typ_XSDC2, sXtal As String, sMgr As String) As FUNCTION_RETURN

    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long
    Dim sCryNum As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc017_SQL.bas -- Function DBDRV_SELECT_BLOCK"

        ''SQL��g�ݗ��Ă�
        '�f���X�V SPT�p���э쐬���@�ύX 2006/04/17 SMP���c
        '' 2007/08/22 SPK Tsutsumi Add Start
        sql = "SELECT CRYNUMCZ, GNLCZ, GNWCZ, INPOSCZ, PLANTCATCZ FROM XSDCZ "
'        sql = "SELECT CRYNUMCZ, GNLCZ, GNWCZ, INPOSCZ FROM XSDCZ "
        '' 2007/08/22 SPK Tsutsumi Add Start
        sql = sql & " WHERE RPCRYNUMCZ = '" & sXtal & "'"
        sql = sql & " AND LIVKCZ = '0'"
        sql = sql & " AND GNWKNTCZ = 'CC400'"
''���X�V START SPT�p���э쐬���@�ύX 2006/05/25 SMP-OKAMOTO
        ''�������ʒu�Ń\�[�g����
        sql = sql & " ORDER BY INPOSCZ "
'        sql = sql & " ORDER BY CRYNUMCZ "
''���X�V END   SPT�p���э쐬���@�ύX 2006/05/25 SMP-OKAMOTO

'        sql = "SELECT CRYNUMC2, GNLC2, GNWC2, INPOSC2 FROM XSDC2 "
'        If sMgr = "M" Then  'MGR�̏ꍇ
'            sql = sql & " WHERE CRYNUMC2 = '" & sXtal & "'"
'        Else                'AGR�̏ꍇ
'            sql = sql & " WHERE RPCRYNUMC2 = '" & sXtal & "'"
'        End If
'        sql = sql & " AND LIVKC2 = '0'"
'        sql = sql & " AND GNWKNTC2 = 'CC400'"
'        sql = sql & " ORDER BY CRYNUMC2 "
        '�f���X�V SPT�p���э쐬���@�ύX 2006/04/17 SMP���c

        '�f�[�^�𒊏o����
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

        If rs Is Nothing Then
            DBDRV_SELECT_BLOCK = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        recCnt = rs.RecordCount
        If recCnt = 0 Then
            rs.Close
        End If
        ''���ǉ� START SPT�p���э쐬���@�ύX 2006/06/28 SMP-OKAMOTO
        sCryNum = ""
        i = 0
        ReDim pBlockData(i)
        Do Until rs.EOF
            ''�u���b�NID�œZ�߂�
            If sCryNum <> CStr(rs("CRYNUMCZ")) Then
                i = i + 1
                ReDim Preserve pBlockData(i)
                sCryNum = CStr(rs("CRYNUMCZ"))
                pBlockData(i).CRYNUMC2 = rs("CRYNUMCZ")
                pBlockData(i).GNLC2 = rs("GNLCZ")
                pBlockData(i).GNWC2 = rs("GNWCZ")
                pBlockData(i).INPOSC2 = rs("INPOSCZ")

                If IsNull(rs("PLANTCATCZ")) = False Then pBlockData(i).PLANTCATC2 = rs("PLANTCATCZ") ' 2007/09/10 SPK Tsutsumi Add Start
            Else
                pBlockData(i).GNLC2 = CLng(pBlockData(i).GNLC2) + CLng(rs("GNLCZ"))
                pBlockData(i).GNWC2 = CLng(pBlockData(i).GNWC2) + CLng(rs("GNWCZ"))

                If IsNull(rs("PLANTCATCZ")) = False Then pBlockData(i).PLANTCATC2 = rs("PLANTCATCZ") ' 2007/09/10 SPK Tsutsumi Add Start
            End If
            rs.MoveNext
        Loop
        ''���ǉ� END   SPT�p���э쐬���@�ύX 2006/06/28 SMP-OKAMOTO
        ''���폜 START SPT�p���э쐬���@�ύX 2006/06/28 SMP-OKAMOTO
'        ReDim pBlockData(recCnt)
'        For i = 1 To recCnt
'            With pBlockData(i)
'                '�f���X�V SPT�p���э쐬���@�ύX 2006/04/17 SMP���c
'                .CRYNUMC2 = rs("CRYNUMCZ")
'                .GNLC2 = rs("GNLCZ")
'                .GNWC2 = rs("GNWCZ")
'                .INPOSC2 = rs("INPOSCZ")
'
''                .CRYNUMC2 = rs("CRYNUMC2")  '
''                .GNLC2 = rs("GNLC2")
''                .GNWC2 = rs("GNWC2")
''                .INPOSC2 = rs("INPOSC2")
'                '�f���X�V SPT�p���э쐬���@�ύX 2006/04/17 SMP���c
'            End With
'            rs.MoveNext
'        Next i
        ''���폜 END   SPT�p���э쐬���@�ύX 2006/06/28 SMP-OKAMOTO
        rs.Close
    DBDRV_SELECT_BLOCK = FUNCTION_RETURN_SUCCESS


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_SELECT_BLOCK = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :����������H���� ��ۯ��Ǘ��c�a�X�V�p�c�a�h���C�o
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:sCryNum�@�@�@,I  ,String         �@,�����ԍ�(BLOCKID)
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function DBDRV_TBCME040_UPDATE(sCryNum As String, sErrMsg As String) As FUNCTION_RETURN

    Dim sql As String
    Dim sDbName As String
    Dim sans As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc017_SQL.bas -- Function DBDRV_TBCME040_UPDATE"
    sErrMsg = ""

    '' �u���b�N�Ǘ��e�[�u���̍X�V
    sDbName = "E040"
    sql = "update TBCME040 set "
    sql = sql & "KRPROCCD='" & MGPRCD_SETUDAN & "', "      ' ���݊Ǘ��H��
    sql = sql & "NOWPROC='" & nextCd & "', "                        ' ���ݍH��
    sql = sql & "LPKRPROCCD='" & MGPRCD_KENNSAKU_KAKOU & "', "                ' �ŏI�ʉߊǗ��H��
    sql = sql & "LASTPASS='" & nowCd & "', "                        ' �ŏI�ʉߍH��
    sql = sql & "UPDDATE=sysdate "                                 ' �X�V���t
    '�f���X�V SPT�p���э쐬���@�ύX 2006/04/21 SMP���c
    sql = sql & "where CRYNUM ='" & sCryNum & "' "          ' �����ԍ�
    sql = sql & "and   nowproc = 'CC400'"                   ' ���ݍH��
'    sql = sql & "where BLOCKID ='" & sCryNum & "' "
    '�f���X�V SPT�p���э쐬���@�ύX 2006/04/21 SMP���c
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_TBCME040_UPDATE = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    DBDRV_TBCME040_UPDATE = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
    DBDRV_TBCME040_UPDATE = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

''���ǉ� START SPT�p���э쐬���@�ύX 2006/08/01 SMP-OKAMOTO
Public Function DBDRV_SELECT_HINBAN(BLOCKID, pHinban, ByRef Hinban12() As String) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc017_SQL.bas -- Function DBDRV_SELECT_HINBAN"

    sql = "select distinct A.hinban12 " & _
          "From (  " & _
          "select hinbcz||ltrim(to_char(revnumcz,'00'))||factorycz||opecz as hinban12,inposcz " & _
          "From XSDCZ  " & _
          "Where (RPCRYNUMCZ =  '" & BLOCKID & "')" & _
          " and (LIVKCZ <> '1' )" & _
          " AND hinbcz||ltrim(to_char(revnumcz,'00'))||factorycz||opecz LIKE '" & pHinban & "%'" & _
          " AND (trim(hinbcz) <> 'Z' AND trim(hinbcz) <> 'G')" & _
          " order by INPOSCZ ) A "

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim Hinban12(0)
        DBDRV_SELECT_HINBAN = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim Hinban12(recCnt)
    For i = 1 To recCnt
        If IsNull(rs("hinban12")) = False Then Hinban12(i) = rs("hinban12")
        rs.MoveNext
    Next
    rs.Close

    DBDRV_SELECT_HINBAN = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function
''���ǉ� END   SPT�p���э쐬���@�ύX 2006/08/01 SMP-OKAMOTO

'2007/08/17 SPK Tsutsumi Add Start
Public Function GetMukeCode() As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Long      '���R�[�h��
    Dim i  As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmbc016_0.frm -- Function Getstaffauthority"

    GetMukeCode = FUNCTION_RETURN_FAILURE

    sql = "Select CODEA9,NAMEJA9 "
    sql = sql & "from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '20' "
    sql = sql & "and (CODEA9 = '14' "
    sql = sql & "or CODEA9 = '15' "
    sql = sql & "or CODEA9 = '16' "
    sql = sql & "or CODEA9 = 'ZX' "         '08/07/01 ooba
    sql = sql & "or CODEA9 = 'ZZ') "

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
            If IsNull(rs.Fields("CODEA9")) = False Then .sMukeCode = rs.Fields("CODEA9")    ' ����R�[�h
            If IsNull(rs.Fields("NAMEJA9")) = False Then .sMukeName = rs.Fields("NAMEJA9")  ' ���於
        End With
        rs.MoveNext
    Next
    rs.Close

    GetMukeCode = FUNCTION_RETURN_SUCCESS
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :�S�i�Ԃ̉��H�d�l�f�[�^�̎擾�h���C�o
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :HIN()          ,I   ,tFullHinban       ,�i�ԃ��X�g
'          :Spec()         ,O   ,Judg_Kakou        ,���H�d�l
'      �@�@:�߂�l          , O  , FUNCTION_RETURN�@, �ǂݍ��݂̐���
'����      :
'
'����      :2002/04/17 ���� �M�� �쐬
'           2009/09    SUMCO Akizuki scmzc_getKakouSpec�����ɍ쐬
'                                    Notch�K�i�����ǉ�

Public Function scmzc_getKakouSpec_cmbc017(HIN() As tFullHinban, Spec() As Judg_Kakou_cmbc017) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Integer
    Dim c0 As Integer
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcjudg_SQL.bas -- Function scmzc_getKakouSpec_cmbc017"
    scmzc_getKakouSpec_cmbc017 = FUNCTION_RETURN_FAILURE
    
    '���߂��S�i�Ԃ̉��H�d�l�����߂�
    sql = "select HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXDPACN, HSXDPMIN, HSXDPMAX, HSXDPDIR, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDWCEN, HSXDWMIN, HSXDWMAX from TBCME018 "
    sql = sql & "Where " & SQLMake_HINBAN(HIN())

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim Spec(recCnt)
    For c0 = 1 To recCnt
        Spec(c0).TOP(0) = fncNullCheck(rs("HSXD1CEN"))
        Spec(c0).TOP(1) = fncNullCheck(rs("HSXD1MIN"))
        Spec(c0).TOP(2) = fncNullCheck(rs("HSXD1MAX"))
        Spec(c0).TAIL(0) = fncNullCheck(rs("HSXD2CEN"))
        Spec(c0).TAIL(1) = fncNullCheck(rs("HSXD2MIN"))
        Spec(c0).TAIL(2) = fncNullCheck(rs("HSXD2MAX"))
        Spec(c0).DPTH(0) = fncNullCheck(rs("HSXDDCEN"))
        
        Spec(c0).POS = rs("HSXDPDIR")
        Spec(c0).DPTH(1) = fncNullCheck(rs("HSXDDMIN"))
        Spec(c0).DPTH(2) = fncNullCheck(rs("HSXDDMAX"))
        Spec(c0).WIDH(0) = fncNullCheck(rs("HSXDWCEN"))
        Spec(c0).WIDH(1) = fncNullCheck(rs("HSXDWMIN"))
        Spec(c0).WIDH(2) = fncNullCheck(rs("HSXDWMAX"))
        
        Spec(c0).ANGLE(0) = fncNullCheck(rs("HSXDPACN"))   '2009/09 SUMCO Akizuki

'       �d�l�K�i�f�[�^�ɢ-1������邽�߁ANull�`�F�b�N�ł̢-1��u������p�~    2009/09 SUMCO Akizuki
        If IsNull(rs("HSXDPMIN")) Then
            Spec(c0).ANGLE(1) = 999
        Else
            Spec(c0).ANGLE(1) = rs("HSXDPMIN")
        End If
        
'       �d�l�K�i�f�[�^�ɢ-1������邽�߁ANull�`�F�b�N�ł̢-1��u������p�~    2009/09 SUMCO Akizuki
        If IsNull(rs("HSXDPMAX")) Then
            Spec(c0).ANGLE(2) = 999
        Else
            Spec(c0).ANGLE(2) = rs("HSXDPMAX")
        End If

        rs.MoveNext
    Next
    
    rs.Close

    scmzc_getKakouSpec_cmbc017 = FUNCTION_RETURN_SUCCESS
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getKakouSpec_cmbc017 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :���H���т̎擾�h���C�o
'���Ұ��@�@:�ϐ���          ,IO     , �^               , ����
'          :BLOCKID         ,I      ,String            ,�����ԍ�or�u���b�NID
'          :Jiltuseki       ,O      ,Judg_Kakou        ,���H����
'      �@�@:�߂�l          ,O      , FUNCTION_RETURN�@, �ǂݍ��݂̐���
'����      :
'
'����      :2002/04/17 ���� �M�� �쐬
'           2009/09 SUMCO Akizuki
'               ���ʊ֐�s_mzccjude_SQL(scmzc_getKakouJiltuseki)���Q�l
'               �w�i�F��������������֐����g�p���āA�e�������������߂ɍ쐬

Public Function scmzc_getKakouJiltuseki_cmbc017 _
(BLOCKID As String, Jiltuseki As Judg_Kakou_cmbc017) As FUNCTION_RETURN
    
    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Integer
    Dim c0 As Integer
    Dim AGRFlag As Boolean
    Dim Ans As String
    Dim tINGOTPOS As Integer
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcjudg_SQL.bas -- Function scmzc_getKakouJiltuseki_cmbc017"
    scmzc_getKakouJiltuseki_cmbc017 = FUNCTION_RETURN_FAILURE
    
    '�Ώۃu���b�N�̉��H���т̏�����
    For c0 = 1 To 2
        Jiltuseki.TAIL(c0) = -1
        Jiltuseki.TOP(c0) = -1
        Jiltuseki.DPTH(c0) = -1
        Jiltuseki.WIDH(c0) = -1
    Next
    
    Jiltuseki.POS = ""
        '�����グ�����̏ꍇ
        sql = "select DMTOP1, DMTOP2, DMTAIL1, DMTAIL2, NCHPOS, NCHDPTH, NCHWIDTH, NCHANGLE from TBCMI002 "
        sql = sql & "where CRYNUM='" & Left(BLOCKID, 9) & "000" & "'"
        
        '��ۯ��Ǘ�(TBCME040)�Q�ƒ�~�@05/10/03 ooba
        sql = sql & " and (select INPOSC2 from XSDC2 where CRYNUMC2 = '" & BLOCKID & "') between INGOTPOS and INGOTPOS+LENGTH-1 "
        sql = sql & "order by INGOTPOS desc, TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum=1"
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        recCnt = rs.RecordCount
        If recCnt = 0 Then
            rs.Close
            scmzc_getKakouJiltuseki_cmbc017 = FUNCTION_RETURN_SUCCESS
            GoTo proc_exit
        End If
        Jiltuseki.TAIL(1) = rs("DMTAIL1")
        Jiltuseki.TAIL(2) = rs("DMTAIL2")
        Jiltuseki.TOP(1) = rs("DMTOP1")
        Jiltuseki.TOP(2) = rs("DMTOP2")
        Jiltuseki.DPTH(1) = rs("NCHDPTH")
        Jiltuseki.DPTH(2) = -1
        Jiltuseki.WIDH(1) = rs("NCHWIDTH")
        Jiltuseki.WIDH(2) = -1
        Jiltuseki.POS = rs("NCHPOS")
        '2009/09 SUMCO Akizuki
        Jiltuseki.ANGLE(1) = rs("NCHANGLE")
        Jiltuseki.ANGLE(2) = 999
        rs.Close

    scmzc_getKakouJiltuseki_cmbc017 = FUNCTION_RETURN_SUCCESS
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getKakouJiltuseki_cmbc017 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
