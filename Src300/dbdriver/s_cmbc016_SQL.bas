Attribute VB_Name = "s_cmbc016_SQL"
Option Explicit

' �������H���o

' �ؒf�d�l (by SUMCO)
Public Type typ_CutSpec1
    hin As tFullHinban          ' �i��
    HSXTYPE As String * 1       ' �^�C�v
    HSXCDIR As String * 1       ' ����
    HSXD1CEN As Double          ' ���a
    HSXCDOP As String * 1       ' �����h�[�v  '4/2 Yam
    HSXDPDIR As String * 2      ' �m�b�`�ʒu  '3/20 Yam
    HSXDDMIN As Double          ' �m�b�`�[���i�l�h�m�j
    HSXDDMAX As Double          ' �m�b�`�[���i�l�`�w�j
    HSXSDSLP As String * 1      ' �V�[�h�X��
    HSXCTCEN As Double          ' �V�[�h�X���p�i�X�c���S�@N(3,2)�j4/2 Yam
    HSXCYCEN As Double          ' �V�[�h�X���p�i�X�����S�jN(3,2)) 4/2 Yam
End Type

' ���i�d�l
Public Type typ_HinSpec1
    hin As tFullHinban          ' �i��
    HSXTYPE As String * 1       ' �^�C�v
    HSXCDIR As String * 1       ' ����
    HSXD1CEN As Double          ' ���a
    HSXDOP As String * 1        ' �����h�[�v
    HSXDPDIR As String * 2      ' �m�b�`�ʒu
    HSXDDMIN As Double          ' �m�b�`�[���i�l�h�m�j
    HSXDDMAX As Double          ' �m�b�`�[���i�l�`�w�j
    HSXSDSLP As Integer         ' �V�[�h�X��
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 start
    TOPREG As Integer           ' TOP�K��
    TAILREG As Double           ' TAIL�K��
    BTMSPRT As Integer          ' �{�g���͏o�K��
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 end
End Type

' �ؒf�w��
Public Type typ_CutInd
    INGOTPOS As Integer         ' �J�b�g�ʒu
    TRANCNT As Integer          ' ������
    LENGTH As Integer           ' ����
    PROCCODE As String * 5      ' �H���R�[�h
    BDCAUS As String * 3        ' �敪
    HINUP As tFullHinban        ' ��i��
    HINDN As tFullHinban        ' ���i��
    BLOCKID As String * 12      ' �u���b�NID
    SMP As typ_SXLSample        ' ��������
    PALTNUM As String * 4       ' �p���b�g�ԍ�
    ERRUPFLG As Boolean         ' ��i�ԃG���[�t���O
    ERRDNFLG As Boolean         ' ���i�ԃG���[�t���O
    RECOMMEND(1 To 13) As String * 1    '�����ߌ���(Rs�`EPD)
End Type

'�����ߌ����̊e�����o�̈Ӗ�
Public Enum RCMD_COL
    RCMD_RS = 1
    RCMD_OI
    RCMD_B1
    RCMD_B2
    RCMD_B3
    RCMD_L1
    RCMD_L2
    RCMD_L3
    RCMD_L4
    RCMD_CS
    RCMD_GD
    RCMD_LT
    RCMD_EPD
End Enum

'�D�揇�ʊi�[�p�ϐ�
Public CUT_PRIORITY As String * 1


'�T�v      :�������H���o�p �����ԍ����͎��c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sCryNum �@�@�@,I  ,String         �@,�����ԍ�
'      �@�@:pCryInf �@�@�@,I  ,typ_TBCME037   �@,�������
'      �@�@:pHinDsn �@�@�@,O  ,typ_TBCME039   �@,�i�Ԑ݌v
'      �@�@:pPupEnd �@�@�@,O  ,typ_TBCMH004   �@,���グ�I������
'      �@�@:pHinSpec�@�@�@,O  ,typ_HinSpec1   �@,���i�d�l
'      �@�@:pCutInd �@�@�@,O  ,typ_CutInd     �@,�ؒf�w��
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmic001b_Disp(sCryNum As String, _
                                           pCryInf As typ_TBCME037, pHinDsn() As typ_TBCME039, _
                                           pPupEnd As typ_TBCMH004, pHinSpec() As typ_HinSpec1, _
                                           pCutInd() As typ_CutInd, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpPupEnd() As typ_TBCMH004
    Dim rs As OraDynaset
    Dim sql As String
    Dim sDbName As String
    Dim sHin As String
    Dim recCnt As Long
    Dim i As Long
    Dim j As Long
    Dim ctcen As Double
    Dim cycen As Double

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function DBDRV_scmzc_fcmic001b_Disp"
    sErrMsg = ""

    '' �������̎擾(s_cmzcTBCME037_SQL.bas ���K�v)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)

    '' �H���`�F�b�N
    If pCryInf.PROCCD <> PROCD_KAKOU_HARAIDASI Then
        sErrMsg = GetMsgStr("EPRC0")
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' ���グ�I�����т̎擾(s_cmzcTBCMH004_SQL.bas ���K�v)
    sDbName = "H004"
    If DBDRV_GetTBCMH004(tmpPupEnd(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpPupEnd) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pPupEnd = tmpPupEnd(1)

    '' �i�Ԑ݌v�̎擾(s_cmzcTBCME039_SQL.bas ���K�v)
    sDbName = "E039"
    sql = " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' order by INGOTPOS"
    If DBDRV_GetTBCME039(pHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    recCnt = UBound(pHinDsn)
    If recCnt = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' ���i�d�l�̎擾
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 start
    sDbName = "E018"
    j = 0
    ReDim pHinSpec(recCnt)
    For i = 1 To recCnt
        sHin = Trim(pHinDsn(i).HINBAN)
        If sHin <> "G" And sHin <> "Z" Then
            sql = "select "
            sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXDOP, HSXDPDIR, HSXDDMIN, HSXDDMAX, HSXCTCEN, HSXCYCEN"
            sql = sql & " ,NVL(TOPREG,0) TOPREG, NVL(TAILREG,0) TAILREG, NVL(BTMSPRT,0) BTMSPRT "
            sql = sql & " from TBCME018 E018,TBCME036 E036"
            sql = sql & " where E018.HINBAN='" & pHinDsn(i).HINBAN & "'"
            sql = sql & " and E018.MNOREVNO=" & pHinDsn(i).REVNUM
            sql = sql & " and E018.FACTORY='" & pHinDsn(i).FACT & "'"
            sql = sql & " and E018.OPECOND='" & pHinDsn(i).OPCOND & "'"
            sql = sql & " and E036.HINBAN='" & pHinDsn(i).HINBAN & "'"
            sql = sql & " and E036.MNOREVNO=" & pHinDsn(i).REVNUM
            sql = sql & " and E036.FACTORY='" & pHinDsn(i).FACT & "'"
            sql = sql & " and E036.OPECOND='" & pHinDsn(i).OPCOND & "'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sErrMsg = GetMsgStr("EGET2", sDbName)
                DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            j = j + 1
            With pHinSpec(j)
                .hin.HINBAN = pHinDsn(i).HINBAN
                .hin.mnorevno = pHinDsn(i).REVNUM
                .hin.factory = pHinDsn(i).FACT
                .hin.opecond = pHinDsn(i).OPCOND
                .HSXTYPE = rs("HSXTYPE")    ' �^�C�v
                .HSXCDIR = rs("HSXCDIR")    ' ����
                .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))  ' ���a
                .HSXDOP = rs("HSXDOP")      ' �����h�[�v
                .HSXDPDIR = rs("HSXDPDIR")          ' �i�r�w�a�ʒu����
                .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))          ' �i�r�w�a�[����
                .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))          ' �i�r�w�a�[���
                ctcen = Abs(fncNullCheck(rs("HSXCTCEN")))
                cycen = Abs(fncNullCheck(rs("HSXCYCEN")))
                .TOPREG = rs("TOPREG")              ' TOP�K��
                .TAILREG = rs("TAILREG")            ' TAIL�K��
                .BTMSPRT = rs("BTMSPRT")            ' �{�g���͏o�K��
                If ((ctcen = 2.83) And (cycen = 2.83)) _
                Or ((ctcen = 4) And (cycen = 0)) _
                Or ((ctcen = 0) And (cycen = 4)) Then
                    .HSXSDSLP = 4
                Else
                    .HSXSDSLP = 0
                End If
            End With
            rs.Close
        End If
    Next i
    ReDim Preserve pHinSpec(j)
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 end

    '' �u���b�N�݌v�̎擾
    sDbName = "E038"
    sql = "select INGOTPOS, LENGTH from TBCME038"
    sql = sql & " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' order by INGOTPOS"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    ReDim pCutInd(recCnt)
    For i = 1 To recCnt
        With pCutInd(i)
            .INGOTPOS = rs("INGOTPOS")      ' �J�b�g�ʒu
            .LENGTH = rs("LENGTH")          ' ����
        End With
        rs.MoveNext
    Next i
    rs.Close

    DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_SUCCESS

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
    DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�������H���o�p ���i�d�l��p�c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:pHinSpec�@�@�@,IO ,typ_HinSpec1   �@,���i�d�l
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmic001b_GetSpec(pHinSpec As typ_HinSpec1) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim ctcen As Double
    Dim cycen As Double

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function DBDRV_scmzc_fcmic001b_GetSpec"

    '' ���i�d�l�̎擾
    sql = "select "
    sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXDOP, HSXDPDIR, HSXDDMIN, HSXDDMAX, HSXCTCEN, HSXCYCEN"
    sql = sql & " from TBCME018"
    sql = sql & " where HINBAN='" & pHinSpec.hin.HINBAN & "'"
    sql = sql & " and MNOREVNO=" & pHinSpec.hin.mnorevno
    sql = sql & " and FACTORY='" & pHinSpec.hin.factory & "'"
    sql = sql & " and OPECOND='" & pHinSpec.hin.opecond & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        DBDRV_scmzc_fcmic001b_GetSpec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With pHinSpec
        .HSXTYPE = rs("HSXTYPE")            ' �^�C�v
        .HSXCDIR = rs("HSXCDIR")            ' ����
        .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))          ' ���a
        .HSXDOP = rs("HSXDOP")              ' �����h�[�v
        .HSXDPDIR = rs("HSXDPDIR")          ' �i�r�w�a�ʒu����
        .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))          ' �i�r�w�a�[����
        .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))          ' �i�r�w�a�[���
        ctcen = Abs(fncNullCheck(rs("HSXCTCEN")))
        cycen = Abs(fncNullCheck(rs("HSXCYCEN")))
        If ((ctcen = 2.83) And (cycen = 2.83)) _
        Or ((ctcen = 4) And (cycen = 0)) _
        Or ((ctcen = 0) And (cycen = 4)) Then
            .HSXSDSLP = 4
        Else
            .HSXSDSLP = 0
        End If
    End With
    rs.Close

    DBDRV_scmzc_fcmic001b_GetSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmic001b_GetSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�������H���o�p ���s���c�a�h���C�o
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pCryInf�@�@�@,I  ,typ_TBCME037   �@,�������
'      �@�@:pProcBR�@�@�@,I  ,typ_TBCMI001   �@,���H���o����
'      �@�@:pCutInd�@�@�@,I  ,typ_CutInd     �@,�ؒf�w��
'      �@�@:pNotCut�@�@�@,I  ,typ_CutInd     �@,�ؒf�w���i���ؒf���j
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function DBDRV_scmzc_fcmic001b_Exec(pCryInf As typ_TBCME037, pProcBR As typ_TBCMI001, _
                                           pCutInd() As typ_CutInd, pNotCut() As typ_CutInd, _
                                           newLength As Integer, sErrMsg As String) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim tmpBlkMng(3) As typ_TBCME040
    Dim sql As String
    Dim sDbName As String
    Dim bFlag As Boolean
    Dim recCnt As Long
    Dim i As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function DBDRV_scmzc_fcmic001b_Exec"
    sErrMsg = ""

    '' WriteDBLog " ", "Start"
    
    '' �������̍X�V
    sDbName = "E037"
    With pCryInf

''''''''' pCryInf �ɓ����Ă�����e���g�p����
'        sql = "update TBCME037 set "
'        sql = sql & "KRPROCCD='" & MGPRCD_KENNSAKU_KAKOU & "', "
'        sql = sql & "PROCCD='" & PROCD_KENNSAKU_KAKOU & "', "
'        sql = sql & "LPKRPROCCD='" & MGPRCD_KAKOU_HARAIDASI & "', "
'        sql = sql & "LASTPASS='" & PROCD_KAKOU_HARAIDASI & "', "
'        sql = sql & "BODYLENG=" & .BODYLENG & ", "
'        sql = sql & "FREELENG=" & .FREELENG & ", "
'        sql = sql & "SEED='" & .SEED & "', "
'        sql = sql & "UPDDATE=sysdate, "
'        sql = sql & "SENDFLAG='0'"
'        sql = sql & " where CRYNUM='" & .CRYNUM & "'"
        
        sql = "update TBCME037 set "
        sql = sql & "KRPROCCD='" & .KRPROCCD & "', "
        sql = sql & "PROCCD='" & .PROCCD & "', "
        sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "
        sql = sql & "LASTPASS='" & .LASTPASS & "', "
        sql = sql & "BODYLENG=" & .BODYLENG & ", "
        sql = sql & "FREELENG=" & .FREELENG & ", "
        sql = sql & "SEED='" & .SEED & "', "
        sql = sql & "UPDDATE=sysdate, "
        sql = sql & "SENDFLAG='0'"
        sql = sql & " where CRYNUM='" & .Crynum & "'"
    End With
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �u���b�N�݌v�̍X�V
    sDbName = "E038"
    sql = "update TBCME038 set "
    sql = sql & "USECLASS='1', "
    sql = sql & "UPDDATE=sysdate, "
    sql = sql & "SENDFLAG='0'"
    sql = sql & " where substr(CRYNUM,1,7)='" & Left(pCryInf.Crynum, 7) & "'"
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �i�Ԑ݌v�̍X�V
    sDbName = "E039"
    sql = "update TBCME039 set "
    sql = sql & "USECLASS='1', "
    sql = sql & "UPDDATE=sysdate, "
    sql = sql & "SENDFLAG='0'"
    sql = sql & " where substr(CRYNUM,1,7)='" & Left(pCryInf.Crynum, 7) & "'"
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    '' �i�Ԑ݌v����ш��グ�w���I���̍X�V�i�݌v���ƍŉ��ؒf�ʒu������Ă�����X�V����j
'    If newLength > 0 Then
        '�i�Ԑ݌v�̍ŉ��i�Ԃ��A�ŉ��ؒf�ʒu�܂łɐL�΂�
'        sDBName = "E039"
'        sql = "update TBCME039 set "
'        sql = sql & "LENGTH=" & newLength & ", "
'        sql = sql & "UPDDATE=sysdate, "
'        sql = sql & "SENDFLAG='0'"
'        sql = sql & " where substr(CRYNUM,1,7)='" & Left(pCryInf.CRYNUM, 7) & "'"
'        sql = sql & "and INGOTPOS=(select max(INGOTPOS) from TBCME039 where "
'        sql = sql & "substr(CRYNUM,1,7)='" & Left(pCryInf.CRYNUM, 7) & "')"
'        WriteDBLog sql, sDBName
'        If OraDB.ExecuteSQL(sql) <= 0 Then
'            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
'            DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
        
        '����I���� SUMMITSENDFLAG ���N���A����
'        sDBName = "H004"
'        sql = "update TBCMH004 set SUMMITSENDFLAG='0' "
'        sql = sql & "where CRYNUM='" & pCryInf.CRYNUM & "'"
'        WriteDBLog sql, sDBName
'        If OraDB.ExecuteSQL(sql) <= 0 Then
'            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
'            DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    End If

    '' ���s����i�ؒf�s�j�Ȃ�t���O�𗧂Ă�
    bFlag = False
    If UBound(pNotCut) = 1 Then
        If pNotCut(1).INGOTPOS = 0 Then
            bFlag = True
        End If
    End If

    If bFlag = False Then
        '' �ؒf�w���̑}��
        sDbName = "E045"
        recCnt = UBound(pCutInd)
        For i = 1 To recCnt
            With pCutInd(i)
                sql = "insert into TBCME045 "
                sql = sql & "(CRYNUM, INGOTPOS, TRANCNT, LENGTH, KRPROCCD, PROCCODE, STAFFID, "
                sql = sql & "HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, STATCLS, BLOCKID, "
                sql = sql & "CRYINDRS, CRYINDOI, CRYINDB1, CRYINDB2, CRYINDB3, CRYINDL1, "
                sql = sql & "CRYINDL2, CRYINDL3, CRYINDL4, CRYINDCS, CRYINDGD, CRYINDT, "
                sql = sql & "CRYINDEP, PRIORITY, PALTNUM, REGDATE, UPDDATE, SENDFLAG, SENDDATE)"
                sql = sql & " select '"
                sql = sql & pCryInf.Crynum & "', "
                sql = sql & .INGOTPOS & ", "
                sql = sql & "nvl(max(TRANCNT),0)+1, "
                sql = sql & .LENGTH & ", '"
                sql = sql & MGPRCD_KAKOU_HARAIDASI & "', '"
                sql = sql & PROCD_KAKOU_HARAIDASI & "', '"
                sql = sql & pProcBR.TSTAFFID & "', '"
                sql = sql & .HINDN.HINBAN & "', "
                sql = sql & .HINDN.mnorevno & ", '"
                sql = sql & .HINDN.factory & "', '"
                sql = sql & .HINDN.opecond & "', '"
                sql = sql & .BDCAUS & "', "
                sql = sql & "'0', '"
                sql = sql & pCryInf.Crynum & "', '"
                sql = sql & .SMP.CRYINDRS & "', '"
                sql = sql & .SMP.CRYINDOI & "', '"
                sql = sql & .SMP.CRYINDB1 & "', '"
                sql = sql & .SMP.CRYINDB2 & "', '"
                sql = sql & .SMP.CRYINDB3 & "', '"
                sql = sql & .SMP.CRYINDL1 & "', '"
                sql = sql & .SMP.CRYINDL2 & "', '"
                sql = sql & .SMP.CRYINDL3 & "', '"
                sql = sql & .SMP.CRYINDL4 & "', '"
                sql = sql & .SMP.CRYINDCS & "', '"
                sql = sql & .SMP.CRYINDGD & "', '"
                sql = sql & .SMP.CRYINDT & "', '"
                sql = sql & .SMP.CRYINDEP & "', "
                '�ؒf�D�揇�ʂ̊i�[
                sql = sql & "'" & CUT_PRIORITY & "', '"
                sql = sql & .PALTNUM & "', "
                sql = sql & "sysdate, "
                sql = sql & "sysdate, "
                sql = sql & "'0', "
                sql = sql & "sysdate"
                sql = sql & " from TBCME045"
                sql = sql & " where CRYNUM='" & pCryInf.Crynum & "'"
                sql = sql & " and INGOTPOS=" & .INGOTPOS
            End With
            '' WriteDBLog sql, sDbName
            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        Next i

        '' �ؒf�w���̑}���i���ؒf���j
        recCnt = UBound(pNotCut)
        For i = 1 To recCnt
            With pNotCut(i)
                sql = "insert into TBCME045 "
                sql = sql & "(CRYNUM, INGOTPOS, TRANCNT, LENGTH, KRPROCCD, PROCCODE, STAFFID, "
                sql = sql & "HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, STATCLS, BLOCKID, "
                sql = sql & "CRYINDRS, CRYINDOI, CRYINDB1, CRYINDB2, CRYINDB3, CRYINDL1, "
                sql = sql & "CRYINDL2, CRYINDL3, CRYINDL4, CRYINDCS, CRYINDGD, CRYINDT, "
                sql = sql & "CRYINDEP, PRIORITY, PALTNUM, REGDATE, UPDDATE, SENDFLAG, SENDDATE)"
                sql = sql & " select '"
                sql = sql & pCryInf.Crynum & "', "
                sql = sql & .INGOTPOS & ", "
                sql = sql & "nvl(max(TRANCNT),0)+1, "
                sql = sql & .LENGTH & ", '"
                sql = sql & MGPRCD_KAKOU_HARAIDASI & "', '"
                sql = sql & PROCD_KAKOU_HARAIDASI & "', '"
                sql = sql & pProcBR.TSTAFFID & "', "
                sql = sql & "'        ', "
                sql = sql & "0, "
                sql = sql & "' ', "
                sql = sql & "' ', '"
                sql = sql & .BDCAUS & "', "
                sql = sql & "'0', '"
                sql = sql & pCryInf.Crynum & "', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                '�ؒf�D�揇�ʂ̊i�[
                sql = sql & "'" & CUT_PRIORITY & "', "
                sql = sql & "'    ', "
                sql = sql & "sysdate, "
                sql = sql & "sysdate, "
                sql = sql & "'0', "
                sql = sql & "sysdate"
                sql = sql & " from TBCME045"
                sql = sql & " where CRYNUM='" & pCryInf.Crynum & "'"
                sql = sql & " and INGOTPOS=" & .INGOTPOS
            End With
            '' WriteDBLog sql, sDbName
            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        Next i
    Else
        '' �u���b�N�Ǘ��̑}��
        sDbName = "E040"
        With tmpBlkMng(1)
            '' TOP
            .Crynum = pCryInf.Crynum
            .INGOTPOS = -99
            .LENGTH = pCryInf.TOPLENG
            .REALLEN = .LENGTH
            .BLOCKID = Left(pCryInf.Crynum, 9) & "TOP"
            .KRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .NOWPROC = nextCd
            .LPKRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .LASTPASS = nowCd
            .DELCLS = "1"
            .LSTATCLS = "H"
            .RSTATCLS = "T"
            .HOLDCLS = "0"
            .BDCAUS = "TOP"
        End With
        With tmpBlkMng(2)
            '' BOT
            .Crynum = pCryInf.Crynum
            .INGOTPOS = -100
            .LENGTH = pCryInf.BOTLENG
            .REALLEN = .LENGTH
            .BLOCKID = Left(pCryInf.Crynum, 9) & "BOT"
            .KRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .NOWPROC = nextCd
            .LPKRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .LASTPASS = nowCd
            .DELCLS = "1"
            .LSTATCLS = "H"
            .RSTATCLS = "T"
            .HOLDCLS = "0"
            .BDCAUS = "BOT"
        End With
        With tmpBlkMng(3)
            '' TOP�����ؒf��
            .Crynum = pCryInf.Crynum
            .INGOTPOS = 0
            .LENGTH = pCryInf.BODYLENG
            .REALLEN = .LENGTH
            .BLOCKID = Left(pCryInf.Crynum, 9) & "0$1"
            .KRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .NOWPROC = nextCd
            .LPKRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .LASTPASS = nowCd
            .DELCLS = "1"
            .LSTATCLS = "H"
            .RSTATCLS = "T"
            .HOLDCLS = "0"
            .BDCAUS = pNotCut(1).BDCAUS
        End With
        For i = 1 To 3
            If DBDRV_BlockMng_Ins(tmpBlkMng(i)) = FUNCTION_RETURN_FAILURE Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        Next i
    End If

    '' ���H���o���т̑}��
    sDbName = "I001"
    With pProcBR
        sql = "insert into TBCMI001 "
        sql = sql & "(CRYNUM, TRANCNT, KRPROCCD, PROCCODE, "
        sql = sql & "UPLENGTH, FREELENG, UPWEIGHT, SEED, PRCMCN, "
        sql = sql & "TSTAFFID, REGDATE, KSTAFFID, SUMMITSENDFLAG, SENDFLAG, SENDDATE)"
        sql = sql & " select '"
        sql = sql & .Crynum & "', "
        sql = sql & "nvl(max(TRANCNT),0)+1, '"
        sql = sql & .KRPROCCD & "', '"
        sql = sql & .PROCCODE & "', "
        sql = sql & .UPLENGTH & ", "
        sql = sql & .FREELENG & ", "
        sql = sql & .UPWEIGHT & ", '"
        sql = sql & .SEED & "', '"
        sql = sql & .PRCMCN & "', '"
        sql = sql & .TSTAFFID & "', "
        sql = sql & "sysdate, '"
        sql = sql & .KSTAFFID & "', "
        sql = sql & "'0', "
        sql = sql & "'0', "
        sql = sql & "sysdate"
        sql = sql & " from TBCMI001"
        sql = sql & " where CRYNUM='" & .Crynum & "'"
    End With
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    '' WriteDBLog " ", "End"
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
    DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv SUMCO�쐬���� vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'�T�v      :�������H���o�p �����ԍ����͎��c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sCryNum �@�@�@,I  ,String         �@,�����ԍ�
'      �@�@:pCryInf �@�@�@,I  ,typ_TBCME037   �@,�������
'      �@�@:pHinDsn �@�@�@,O  ,typ_TBCME039   �@,�i�Ԑ݌v
'      �@�@:pPupEnd �@�@�@,O  ,typ_TBCMH004   �@,���グ�I������
'      �@�@:pHinSpec�@�@�@,O  ,typ_HinSpec1   �@,���i�d�l
'      �@�@:pCutInd �@�@�@,O  ,typ_CutInd     �@,�ؒf�w��
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function fcmic001b_Disp(sCryNum As String, _
                                           pCryInf As typ_TBCME037, pHinDsn() As typ_TBCME039, _
                                           pPupEnd As typ_TBCMH004, pHinSpec() As typ_CutSpec1, _
                                           pCutInd() As typ_CutInd, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpPupEnd() As typ_TBCMH004
    Dim rs As OraDynaset
    Dim sql As String
    Dim sDbName As String
    Dim sHin As String
    Dim recCnt As Long
    Dim i As Long
    Dim j As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function fcmic001b_Disp"
    sErrMsg = ""

    '' �������̎擾(s_cmzcTBCME037_SQL.bas ���K�v)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ECRY0")
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("ECRY0")
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)

    '' �H���`�F�b�N
    If pCryInf.PROCCD <> PROCD_KAKOU_HARAIDASI Then
        sErrMsg = GetMsgStr("EPRC0")
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' ���グ�I�����т̎擾(s_cmzcTBCMH004_SQL.bas ���K�v)
    sDbName = "H004"
    If DBDRV_GetTBCMH004(tmpPupEnd(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpPupEnd) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pPupEnd = tmpPupEnd(1)

    '' �i�Ԑ݌v�̎擾(s_cmzcTBCME039_SQL.bas ���K�v)
    sDbName = "E039"
    sql = " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' order by INGOTPOS"
    If DBDRV_GetTBCME039(pHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    recCnt = UBound(pHinDsn)
    If recCnt = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' ���i�d�l�̎擾
    sDbName = "E018"
    j = 0
    ReDim pHinSpec(recCnt)
    For i = 1 To recCnt
        sHin = Trim(pHinDsn(i).HINBAN)
        If sHin <> "G" And sHin <> "Z" Then
            sql = "select "
            'sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXDOP"
            sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXCDOP "  '4/3 Yam
            sql = sql & " from TBCME018 A,TBCME020 B"
            'sql = sql & " from TBCME018"
            sql = sql & " where A.HINBAN='" & pHinDsn(i).HINBAN & "'"
            sql = sql & " and A.MNOREVNO=" & pHinDsn(i).REVNUM
            sql = sql & " and A.FACTORY='" & pHinDsn(i).FACT & "'"
            sql = sql & " and A.OPECOND='" & pHinDsn(i).OPCOND & "'"
            sql = sql & " and B.HINBAN='" & pHinDsn(i).HINBAN & "'"    '4/3 Yam
            sql = sql & " and B.MNOREVNO=" & pHinDsn(i).REVNUM
            sql = sql & " and B.FACTORY='" & pHinDsn(i).FACT & "'"
            sql = sql & " and B.OPECOND='" & pHinDsn(i).OPCOND & "'"
            
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sErrMsg = GetMsgStr("EGET2", sDbName)
                fcmic001b_Disp = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            j = j + 1
            With pHinSpec(j)
                .hin.HINBAN = pHinDsn(i).HINBAN
                .hin.mnorevno = pHinDsn(i).REVNUM
                .hin.factory = pHinDsn(i).FACT
                .hin.opecond = pHinDsn(i).OPCOND
                .HSXTYPE = rs("HSXTYPE")    ' �^�C�v
                .HSXCDIR = rs("HSXCDIR")    ' ����
                .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))  ' ���a
'                .HSXDOP = rs("HSXDOP")      ' �����h�[�v
                .HSXCDOP = rs("HSXCDOP")     ' �����h�[�v  4/2 Yam
            End With
            rs.Close
        End If
    Next i
    ReDim Preserve pHinSpec(j)

    '' �u���b�N�݌v�̎擾
    sDbName = "E038"
    sql = "select INGOTPOS, LENGTH from TBCME038"
    sql = sql & " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' order by INGOTPOS"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    ReDim pCutInd(recCnt)
    For i = 1 To recCnt
        With pCutInd(i)
            .INGOTPOS = rs("INGOTPOS")      ' �J�b�g�ʒu
            .LENGTH = rs("LENGTH")          ' ����
        End With
        rs.MoveNext
    Next i
    rs.Close

    fcmic001b_Disp = FUNCTION_RETURN_SUCCESS

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
    fcmic001b_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�������H���o�p ���i�d�l��p�c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:pHinSpec�@�@�@,IO ,typ_HinSpec1   �@,���i�d�l
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function fcmic001b_GetSpec(pHinSpec As typ_CutSpec1) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim ctcen As Double
    Dim cycen As Double

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function fcmic001b_GetSpec"

    '' ���i�d�l�̎擾
    sql = "select "
    sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXCDOP,"     '4/2 Yam
    sql = sql & "HSXDPDIR, HSXDDMIN, HSXDDMAX, HSXSDSLP,"   '3/7 Yam
    sql = sql & "HSXCTCEN, HSXCYCEN "  '4/2 Yam
    sql = sql & " from TBCME018 A,TBCME020 B"
    sql = sql & " where A.HINBAN='" & pHinSpec.hin.HINBAN & "'"
    sql = sql & " and A.MNOREVNO=" & pHinSpec.hin.mnorevno
    sql = sql & " and A.FACTORY='" & pHinSpec.hin.factory & "'"
    sql = sql & " and A.OPECOND='" & pHinSpec.hin.opecond & "'"
    sql = sql & " and B.HINBAN='" & pHinSpec.hin.HINBAN & "'"
    sql = sql & " and B.MNOREVNO=" & pHinSpec.hin.mnorevno
    sql = sql & " and B.FACTORY='" & pHinSpec.hin.factory & "'"
    sql = sql & " and B.OPECOND='" & pHinSpec.hin.opecond & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        fcmic001b_GetSpec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With pHinSpec
        .HSXTYPE = rs("HSXTYPE")       ' �^�C�v
        .HSXCDIR = rs("HSXCDIR")       ' ����
        .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))     ' ���a
        .HSXCDOP = rs("HSXCDOP")       ' �����h�[�v  4/2 Yam
        .HSXDPDIR = rs("HSXDPDIR")     ' �m�b�`�ʒu
        .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))     ' �m�b�`�[���i�l�h�m�j3/7 Yam
        .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))     ' �m�b�`�[���i�l�`�w�j
        .HSXSDSLP = rs("HSXSDSLP")     ' �V�[�h�X��
        .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))     ' �V�[�h�X���p�i�X�c���S�j4/2 Yam
        .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))     ' �V�[�h�X���p�i�X�c���S�j4/2 Yam
    End With
    rs.Close

    fcmic001b_GetSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    fcmic001b_GetSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ SUMCO�쐬���� ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


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
            .Crynum = rs("CRYNUM")           ' �����ԍ�
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
            .Crynum = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .HINBAN = rs("HINBAN")           ' �i��
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

'�T�v      :�e�[�u���uTBCMH004�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMH004 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMH004_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMH004(records() As typ_TBCMH004, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, KRPROCCD, PROCCODE, LENGTOP, LENGTKDO, LENGTAIL, LENGFREE, DM1, DM2, DM3, WGHTTOP, WGHTTKDO," & _
              " WGHTTAIL, WGHTFREE, WGTOPCUT, UPWEIGHT, CHARGE, SEED, STATCLS, JDGECODE, PWTIME, ADDDPPOS, ADDDPCLS, ADDDPVAL," & _
              " ADDDPNAM, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMH004"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMH004 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Crynum = rs("CRYNUM")           ' �����ԍ�
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .LENGTOP = rs("LENGTOP")         ' �����iTOP�j
            .LENGTKDO = rs("LENGTKDO")       ' �����i�����j
            .LENGTAIL = rs("LENGTAIL")       ' �����iTAIL�j
            .LENGFREE = rs("LENGFREE")       ' �t���[����
            .DM1 = rs("DM1")                 ' �������a�P
            .DM2 = rs("DM2")                 ' �������a�Q
            .DM3 = rs("DM3")                 ' �������a�R
            .WGHTTOP = rs("WGHTTOP")         ' �d�ʁiTOP�j
            .WGHTTKDO = rs("WGHTTKDO")       ' �d�ʁi�����j
            .WGHTTAIL = rs("WGHTTAIL")       ' �d�ʁiTAIL)
            .WGHTFREE = rs("WGHTFREE")       ' �d�ʁi�t���[�����j
            .WGTOPCUT = rs("WGTOPCUT")       ' �g�b�v�J�b�g�d��
            .UPWEIGHT = rs("UPWEIGHT")       ' ���グ�d��
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .SEED = rs("SEED")               ' �V�[�h
            .STATCLS = rs("STATCLS")         ' BOT�󋵋敪
            .JDGECODE = rs("JDGECODE")       ' ����R�[�h
            .PWTIME = rs("PWTIME")           ' �p���[����
            .ADDDPPOS = rs("ADDDPPOS")       ' �ǉ��h�[�v�ʒu
            .ADDDPCLS = rs("ADDDPCLS")       ' �ǉ��h�[�p���g���
            .ADDDPVAL = rs("ADDDPVAL")       ' �ǉ��h�[�v��
            .ADDDPNAM = rs("ADDDPNAM")       ' �ǉ��h�[�v��
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMH004 = FUNCTION_RETURN_SUCCESS
End Function



