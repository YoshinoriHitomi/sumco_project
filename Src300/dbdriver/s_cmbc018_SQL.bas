Attribute VB_Name = "s_cmbc018_SQL"
Option Explicit

' �ؒf�҂��ꗗ

' �ؒf�҂��ꗗ
Public Type typ_CutMap
    PRIORITY    As String * 1       ' �D�揇��
    BLOCKID     As String * 12      ' �u���b�NID
    REGDATE     As Date             ' �o�^���t
    KENSAKU     As Integer          ' ������H��
    CRYNUM      As String           ' �����ԍ�
    TRANCNT     As Integer          ' ������
    PRCMCN      As String           ' ����@
End Type
' �ؒf�d�l (by SUMCO)
Public Type typ_CutSpec1
    HIN As tFullHinban          ' �i��
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
    HIN As tFullHinban          ' �i��
    HSXTYPE As String * 1       ' �^�C�v
    HSXCDIR As String * 1       ' ����
    HSXD1CEN As Double          ' ���a
    HSXDOP As String * 1        ' �����h�[�v
    HSXDPDIR As String * 2      ' �m�b�`�ʒu
    HSXDDMIN As Double          ' �m�b�`�[���i�l�h�m�j
    HSXDDMAX As Double          ' �m�b�`�[���i�l�`�w�j
    HSXSDSLP As Integer         ' �V�[�h�X��
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


'�T�v      :�ؒf�w���ꗗ�p ��ʕ\�����c�a�h���C�o
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pCutMap�@�@�@,O  ,typ_CutMap     �@,�ؒf�w���ꗗ
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmic001d_Disp(pCutMap() As typ_CutMap) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Long
    Dim i As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_scmzc_fcmic001d_Disp"

    '================= �b��Ή� 2001/10/17 T.Nomura ==================
    '�����Ȑؒf�w�����u�ؒf�ρv�Ƃ���
    sql = "update TBCME045 CUT set" & _
          " STATCLS='1', " & _
          " UPDDATE=sysdate, " & _
          " SENDFLAG='9' " & _
          "Where (Cut.STATCLS = 0)" & _
          "  and ((select NOWPROC from TBCME040 where BLOCKID=CUT.BLOCKID)<>'CC450')"
    OraDB.ExecuteSQL sql
    '=================================================================

    '' �ؒf�w���̎擾
    sql = ""
    sql = sql & "select B.PRIORITY, B.BLOCKID, B.MaxDate, decode(A.CRYNUM,null,0,1) as KENSAKU,"
    sql = sql & "       B.CRYNUM, B.TRANCNT, B.PRCMCN"
    sql = sql & "  from (select distinct CRYNUM from TBCMI002) A, "
    sql = sql & "       (select E045.PRIORITY, E045.BLOCKID, nvl(max(E045.REGDATE),to_date('1900','YYYY')) as MaxDate,"
    sql = sql & "               E045.CRYNUM, I001.TRANCNT, I001.PRCMCN "
    sql = sql & "          from TBCME045 E045, "
    sql = sql & "               (select CRYNUM,max(TRANCNT) as TRANCNT,PRCMCN from TBCMI001 group by CRYNUM,PRCMCN) I001"
    sql = sql & "         where (E045.STATCLS='0')"
    sql = sql & "           and (substr(E045.BLOCKID,1,9)||'000' = I001.CRYNUM(+))"
    sql = sql & "      group by PRIORITY, BLOCKID, E045.CRYNUM, I001.TRANCNT, I001.PRCMCN order by BLOCKID) B"
    sql = sql & " where (substr(B.BLOCKID,1,9)||'000' = A.CRYNUM(+))"
    sql = sql & " order by B.PRIORITY, B.BLOCKID"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
    End If

    ReDim pCutMap(recCnt)
    For i = 1 To recCnt
        With pCutMap(i)
            .PRIORITY = rs("PRIORITY")  ' �D�揇��
            .BLOCKID = rs("BLOCKID")    ' �u���b�NID
            .REGDATE = rs("MaxDate")    ' �o�^���t
            If rs("PRCMCN") = "M" Then  'MGR�Ȃ疢����Őؒf��
                .KENSAKU = 1    ' ������H��
            Else                        'AGR�Ȃ猤����H�ς��ǂ����Ŕ��f����
                .KENSAKU = rs("KENSAKU")    ' ������H��
            End If
            .CRYNUM = rs("CRYNUM")      ' �����ԍ�
            .TRANCNT = rs("TRANCNT")    ' ������
            .PRCMCN = IIf(IsNull(rs("PRCMCN")), "", (rs("PRCMCN")))     ' ����@
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
' �ؒf

'�T�v      :�ؒf�p ��ʕ\�����c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sBlockID�@�@�@,I  ,String         �@,�u���b�NID
'      �@�@:pCryInf �@�@�@,O  ,typ_TBCME037   �@,�������
'      �@�@:pBlkMng �@�@�@,O  ,typ_TBCME040   �@,�u���b�N�Ǘ�
'      �@�@:pHinMng �@�@�@,O  ,typ_TBCME041   �@,�i�ԊǗ��i�������͕i�Ԑ݌v�j
'      �@�@:pCrySmp �@�@�@,O  ,typ_XSDCS   �@   ,�V�T���v���Ǘ��i�u���b�N�j
'      �@�@:pProcBR �@�@�@,O  ,typ_TBCMI001   �@,���H���o����
'      �@�@:pCutInd �@�@�@,O  ,typ_CutInd     �@,�ؒf�w��
'      �@�@:pNotCut �@�@�@,O  ,typ_CutInd     �@,�ؒf�w���i���ؒf���j
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmic001e_Disp(ByVal sBlockID As String, pCryInf As typ_TBCME037, _
                                           pBlkMng() As typ_TBCME040, pHinMng() As typ_TBCME041, _
                                           pCrySmp() As typ_XSDCS, pProcBR As typ_TBCMI001, _
                                           pCutInd() As typ_CutInd, pNotCut() As typ_CutInd, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpHinDsn() As typ_TBCME039
    Dim tmpProcBR() As typ_TBCMI001
    Dim rs As OraDynaset
    Dim sql As String
    Dim sCryNum As String
    Dim sDBName As String
    Dim recCnt As Long
    Dim i As Long
    '----2002/05/10 �ǉ�-------
    Dim newLength As Integer
    Dim cutLength As Integer
    Dim desLength As Integer
    '--------------------------

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_scmzc_fcmic001e_Disp"
    sErrMsg = ""

    '' �������̎擾(s_cmzcTBCME037_SQL.bas ���K�v)
    sCryNum = Left(sBlockID, 9) & "000"
    sDBName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)

    '' �u���b�N�Ǘ��̎擾(s_cmzcTBCME040_SQL.bas ���K�v)
    sDBName = "E040"
    sql = " where CRYNUM='" & sCryNum & "' and INGOTPOS>=0 order by INGOTPOS"
    If DBDRV_GetTBCME040(pBlkMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �i�ԊǗ��̎擾(s_cmzcTBCME041_SQL.bas ���K�v)
    sDBName = "E041"
    sql = " where CRYNUM='" & sCryNum & "' order by INGOTPOS"
    If DBDRV_GetTBCME041(pHinMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinMng) = 0 Then
        '' �i�Ԑ݌v�̎擾(s_cmzcTBCME039_SQL.bas ���K�v)
        sDBName = "E039"
        sql = " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' order by INGOTPOS"
        If DBDRV_GetTBCME039(tmpHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("EGET2", sDBName)
            DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        recCnt = UBound(tmpHinDsn)
        If recCnt = 0 Then
            sErrMsg = GetMsgStr("EGET2", sDBName)
            DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        ReDim pHinMng(recCnt)
        For i = 1 To recCnt
            With pHinMng(i)
                .CRYNUM = sCryNum
                .INGOTPOS = tmpHinDsn(i).INGOTPOS
                .hinban = tmpHinDsn(i).hinban
                .REVNUM = tmpHinDsn(i).REVNUM
                .factory = tmpHinDsn(i).FACT
                .opecond = tmpHinDsn(i).OPCOND
                .LENGTH = tmpHinDsn(i).LENGTH
                .REGDATE = tmpHinDsn(i).REGDATE
                .UPDDATE = tmpHinDsn(i).UPDDATE
                .SENDFLAG = tmpHinDsn(i).SENDFLAG
                .SENDDATE = tmpHinDsn(i).SENDDATE
            End With
        Next i
    End If

    '' �����T���v���Ǘ��̎擾
    sDBName = "E043"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME043(pCrySmp(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' ���H���o���т̎擾(s_cmzcTBCMI001_SQL.bas ���K�v)
    sDBName = "I001"
    sql = " where CRYNUM='" & sCryNum & "'"
    sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMI001"
    sql = sql & " where CRYNUM='" & sCryNum & "')"
    If DBDRV_GetTBCMI001(tmpProcBR(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpProcBR) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pProcBR = tmpProcBR(1)

    '' �ؒf�w���̎擾
    sDBName = "E045"
    sql = "select INGOTPOS, TRANCNT from TBCME045"
    sql = sql & " where BLOCKID='" & sBlockID & "' and INGOTPOS>=0 and STATCLS='0'"
    sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCME045"
    sql = sql & " where BLOCKID='" & sBlockID & "') order by INGOTPOS"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    ReDim pCutInd(recCnt)
    For i = 1 To recCnt
        With pCutInd(i)
            .INGOTPOS = rs("INGOTPOS")      ' �������J�n�ʒu
            .TRANCNT = rs("TRANCNT")        ' ������
        End With
        rs.MoveNext
    Next i
    rs.Close

    For i = 1 To recCnt
        With pCutInd(i)
            sql = "select "
            sql = sql & "LENGTH, PROCCODE, BDCAUS, "
            sql = sql & "HINBAN, REVNUM, FACTORY, OPECOND, BLOCKID, "
            sql = sql & "CRYINDRS, CRYINDOI, CRYINDB1, CRYINDB2, CRYINDB3, "
            sql = sql & "CRYINDL1, CRYINDL2, CRYINDL3, CRYINDL4, CRYINDCS, "
            sql = sql & "CRYINDGD, CRYINDT, CRYINDEP, PALTNUM"
            sql = sql & " from TBCME045"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and INGOTPOS=" & pCutInd(i).INGOTPOS
            sql = sql & " and TRANCNT=" & pCutInd(i).TRANCNT
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sErrMsg = GetMsgStr("EGET2", sDBName)
                DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            .LENGTH = rs("LENGTH")
            .PROCCODE = rs("PROCCODE")
            .BDCAUS = rs("BDCAUS")
            .HINDN.hinban = rs("HINBAN")
            .HINDN.mnorevno = rs("REVNUM")
            .HINDN.factory = rs("FACTORY")
            .HINDN.opecond = rs("OPECOND")
            .BLOCKID = rs("BLOCKID")
            .SMP.CRYINDRS = rs("CRYINDRS")
            .SMP.CRYINDOI = rs("CRYINDOI")
            .SMP.CRYINDB1 = rs("CRYINDB1")
            .SMP.CRYINDB2 = rs("CRYINDB2")
            .SMP.CRYINDB3 = rs("CRYINDB3")
            .SMP.CRYINDL1 = rs("CRYINDL1")
            .SMP.CRYINDL2 = rs("CRYINDL2")
            .SMP.CRYINDL3 = rs("CRYINDL3")
            .SMP.CRYINDL4 = rs("CRYINDL4")
            .SMP.CRYINDCS = rs("CRYINDCS")
            .SMP.CRYINDGD = rs("CRYINDGD")
            .SMP.CRYINDT = rs("CRYINDT")
            .SMP.CRYINDEP = rs("CRYINDEP")
            .PALTNUM = rs("PALTNUM")
            rs.Close
        End With
    Next i

    '' �ؒf�w���i���ؒf���j�̎擾
    sql = "select INGOTPOS, TRANCNT, LENGTH, BDCAUS from TBCME045"
    sql = sql & " where BLOCKID='" & sBlockID & "' and INGOTPOS<=-99 and STATCLS='0'"
    sql = sql & " and TRANCNT=" & pCutInd(1).TRANCNT & " order by INGOTPOS"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim pNotCut(recCnt)
    For i = 1 To recCnt
        With pNotCut(i)
            .INGOTPOS = rs("INGOTPOS")      ' �������J�n�ʒu
            .TRANCNT = rs("TRANCNT")        ' ������
            .LENGTH = rs("LENGTH")          ' ����
            .BDCAUS = rs("BDCAUS")          ' �敪
        End With
        rs.MoveNext
    Next i
    rs.Close
    
    ''���ؒf���̐������������擾
    sql = "select C.CRYNUM, min(BODYLENG) as BODYLENG, min(INGOTPOS) as FIRSTCUT, max(INGOTPOS) as LASTCUT "
    sql = sql & "from TBCME045 C, TBCME037 XL "
    sql = sql & "where C.INGOTPOS>=0 and C.STATCLS='0'"
    sql = sql & "  and C.CRYNUM=XL.CRYNUM "
    sql = sql & "  and C.CRYNUM='" & sCryNum & "' "
    sql = sql & "group by C.CRYNUM"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    For i = 1 To UBound(pNotCut)
        If pNotCut(i).INGOTPOS = -99 Then
            pNotCut(i).LENGTH = rs("FIRSTCUT")
        Else
            pNotCut(i).LENGTH = rs("BODYLENG") - rs("LASTCUT")
        End If
    Next
    rs.Close
    Set rs = Nothing
    
    '----2002/05/10--------------
    '' �i�ԊǗ��̍X�V�i�݌v���ƍŉ��ؒf�ʒu������Ă����ꍇ�j
    '' �ŉ��ʒu�̌������ʂ����͂ł���悤�ɕi�ԊǗ��̍ŉ��ʒu���u���b�N�Ǘ��ƍ��킹��
    If Right$(sBlockID, 3) = "000" Then
        desLength = tmpHinDsn(UBound(tmpHinDsn)).INGOTPOS + tmpHinDsn(UBound(tmpHinDsn)).LENGTH
        cutLength = pCutInd(UBound(pCutInd)).INGOTPOS + pCutInd(UBound(pCutInd)).LENGTH

        newLength = 0
        If desLength < cutLength Then
            newLength = cutLength - tmpHinDsn(UBound(tmpHinDsn)).INGOTPOS
        End If

        If newLength > 0 Then
            '�i�Ԑ݌v�̍ŉ��i�Ԃ��A�ŉ��ؒf�ʒu�܂ŐL�΂�
            pHinMng(UBound(tmpHinDsn)).LENGTH = newLength
        End If
    End If
    '----------------------------

    
    DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("EGET2", sDBName)
    DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�ؒf�p ���s���c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sCryNum �@�@�@,I  ,String         �@,�����ԍ�
'      �@�@:pBlkMng �@�@�@,I  ,typ_TBCME040   �@,�u���b�N�Ǘ�
'      �@�@:pBlkOld �@�@�@,I  ,typ_TBCME040   �@,�ύX�O�u���b�N�Ǘ�
'      �@�@:pHinMng �@�@�@,I  ,typ_TBCME041   �@,�i�ԊǗ�
'      �@�@:pHinOld �@�@�@,I  ,typ_TBCME041   �@,�ύX�O�i�ԊǗ�
'      �@�@:pCrySmp �@�@�@,IO ,typ_XSDCS   �@   ,�V�T���v���Ǘ��i�u���b�N�j
'      �@�@:pCryOld �@�@�@,I  ,typ_XSDCS   �@   ,�ύX�O�V�T���v���Ǘ��i�u���b�N�j
'      �@�@:pCryCat �@�@�@,I  ,typ_TBCMG007   �@,�N���X�^���J�^���O�������
'      �@�@:pCutRslt�@�@�@,I  ,typ_TBCMI003   �@,�ؒf����
'      �@�@:pCutInd �@�@�@,I  ,typ_CutInd     �@,�ؒf�w��
'      �@�@:pNotCut �@�@�@,I  ,typ_CutInd     �@,�ؒf�w���i���ؒf���j
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function DBDRV_scmzc_fcmic001e_Exec(sCryNum As String, _
                                           pBlkMng() As typ_TBCME040, pBlkOld() As typ_TBCME040, _
                                           pHinMng() As typ_TBCME041, pHinOld() As typ_TBCME041, _
                                           pCrySmp() As typ_XSDCS, pCryOld() As typ_XSDCS, _
                                           pCryCat() As typ_TBCMG007, pCutRslt As typ_TBCMI003, _
                                           pCutInd() As typ_CutInd, pNotCut() As typ_CutInd, sErrMsg As String) As FUNCTION_RETURN

    Dim sql As String
    Dim sDBName As String
    Dim recCnt As Long
    Dim i As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_scmzc_fcmic001e_Exec"
    sErrMsg = ""
    
    '' WriteDBLog " ", "Start"

    '' �������̍X�V
    sDBName = "E037"
    sql = "update TBCME037 set "
    sql = sql & "KRPROCCD='" & MGPRCD_KESSYOU_SOUGOUHANTEI & "', "
    sql = sql & "PROCCD='" & nextCd & "', "
    sql = sql & "LPKRPROCCD='" & MGPRCD_SETUDAN & "', "
    sql = sql & "LASTPASS='" & nowCd & "', "
    sql = sql & "UPDDATE=sysdate, "
    sql = sql & "SENDFLAG='0'"
    sql = sql & " where CRYNUM='" & sCryNum & "'"
    '' WriteDBLog sql, sDBName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
        DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �u���b�N�Ǘ��̑}���^�X�V
    sDBName = "E040"
    If DBDRV_BlockMng_UpdIns(pBlkOld(), pBlkMng()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
        DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �i�ԊǗ��̑}���^�X�V
    sDBName = "E041"
    If DBDRV_Hinban_UpdIns(pHinOld(), pHinMng()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
        DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �T���v�����̎擾
    recCnt = UBound(pCrySmp)
    For i = 1 To recCnt
        If pCrySmp(i).REPSMPLIDCS = 0 Then
            pCrySmp(i).REPSMPLIDCS = GetNewID_SampleNo()
        End If
    Next i

    '' �����T���v���Ǘ��̑}���^�X�V
    sDBName = "E043"
    If DBDRV_CrySmp_UpdIns(pCryOld(), pCrySmp()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
        DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �����������葪��l�̍X�V
    sDBName = "J014"
    recCnt = UBound(pCrySmp)
    For i = 1 To recCnt
        With pCrySmp(i)
            If .KTKBNCS = "1" Then
                sql = "update TBCMJ014 set SMPKBN='" & .SMPKBNCS & "'"
                sql = sql & " where CRYNUM='" & .XTALCS & "' and POSITION=" & .INPOSCS
                '' WriteDBLog sql, sDBName
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
                    DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

    '' �ؒf�w���̍X�V
    sDBName = "E045"
    recCnt = UBound(pCutInd)
    For i = 1 To recCnt
        With pCutInd(i)
            sql = "update TBCME045 set "
            sql = sql & "STATCLS='1', "
            sql = sql & "UPDDATE=sysdate, "
            sql = sql & "SENDFLAG='0', "
            sql = sql & "SENDDATE=sysdate"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and INGOTPOS=" & .INGOTPOS
            sql = sql & " and TRANCNT=" & .TRANCNT
        End With
        '' WriteDBLog sql, sDBName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
            DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' �ؒf�w���̍X�V�i���ؒf���j
    recCnt = UBound(pNotCut)
    For i = 1 To recCnt
        With pNotCut(i)
            sql = "update TBCME045 set "
            sql = sql & "STATCLS='1', "
            sql = sql & "UPDDATE=sysdate, "
            sql = sql & "SENDFLAG='0', "
            sql = sql & "SENDDATE=sysdate"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and INGOTPOS=" & .INGOTPOS
            sql = sql & " and TRANCNT=" & .TRANCNT
        End With
        '' WriteDBLog sql, sDBName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
            DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' �N���X�^���J�^���O������т̑}��
    sDBName = "G007"
    recCnt = UBound(pCryCat)
    For i = 1 To recCnt
        With pCryCat(i)
            sql = "insert into TBCMG007 "
            sql = sql & "(CRYNUM, TRANCNT, KRPROCCD, PROCCODE, BDCODE, PALTNUM, "
            sql = sql & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE)"
            sql = sql & " select '"
            sql = sql & .CRYNUM & "', "
            sql = sql & "nvl(max(TRANCNT),0)+1, '"
            sql = sql & .KRPROCCD & "', '"
            sql = sql & .PROCCODE & "', '"
            sql = sql & .BDCODE & "', '"
            sql = sql & .PALTNUM & "', '"
            sql = sql & .TSTAFFID & "', "
            sql = sql & "sysdate, '"
            sql = sql & .KSTAFFID & "', "
            sql = sql & "sysdate, "
            sql = sql & "'0', "
            sql = sql & "sysdate "
            sql = sql & " from TBCMG007"
            sql = sql & " where CRYNUM='" & .CRYNUM & "'"
        End With
        '' WriteDBLog sql, sDBName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
            DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' �ؒf���т̑}��
    sDBName = "I003"
    With pCutRslt
        sql = "insert into TBCMI003 "
        sql = sql & "(CRYNUM, TRANCNT, KRPROCCD, PROCCODE, TSTAFFID, "
        sql = sql & "REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, GOUKI)" ' YAM
        sql = sql & " select '"
        sql = sql & .CRYNUM & "', "
        sql = sql & "nvl(max(TRANCNT),0)+1, '"
        sql = sql & .KRPROCCD & "', '"
        sql = sql & .PROCCODE & "', '"
        sql = sql & .TSTAFFID & "', "
        sql = sql & "sysdate, '"
        sql = sql & .KSTAFFID & "', "
        sql = sql & "sysdate, "
        sql = sql & "'0', "
        sql = sql & "sysdate, '"    ' Yam
        sql = sql & .GOUKI & "'"    ' Yam
        sql = sql & " from TBCMI003"
        sql = sql & " where CRYNUM='" & .CRYNUM & "'"
    End With
    '' WriteDBLog sql, sDBName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
        DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_SUCCESS

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
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
    DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�u���b�NID�p�A�Ԃ̎擾
'���Ұ��@�@:�ϐ���       ,IO ,�^       ,����
'�@�@      :CryNum       ,I  ,String   ,�����ԍ�
'�@�@      :�߂�l       ,O  ,String �@,�u���b�NID�A��(max)
'����      :�u���b�NID�̍ő�A�Ԃ��擾����
'����      :2001/09/26�@���{ �쐬
Public Function DBDRV_GetBlockNum(CRYNUM As String) As Integer
    
    Dim rs As OraDynaset
    Dim sql As String
    Dim sNum As String
    

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_GetBlockNum"

    DBDRV_GetBlockNum = 0

    sql = "select "
    sql = sql & "nvl(max(substr(BLOCKID,12,1)),'0') as NUM "
    sql = sql & "from TBCME040 "
    sql = sql & "where BLOCKID like '" & Left$(CRYNUM, 10) & "$_'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs Is Nothing Then
        rs.Close
        GoTo proc_exit
    End If
    
    If rs.RecordCount = 0 Then
        DBDRV_GetBlockNum = 0
    Else
        sNum = rs("NUM")
        If StrComp(sNum, "9", vbTextCompare) = 1 Then
            DBDRV_GetBlockNum = Asc(sNum) - 55
        Else
            DBDRV_GetBlockNum = Val(sNum)
        End If
    End If
        
    rs.Close

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
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

'�T�v      :�e�[�u���uTBCME040�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME040 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCME040_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME040(records() As typ_TBCME040, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, REALLEN, BLOCKID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS," & _
              " RSTATCLS, HOLDCLS, BDCAUS, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE," & _
              " PASSFLAG "   '02/07/05 hama
    
    sqlBase = sqlBase & "From TBCME040"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME040 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .LENGTH = rs("LENGTH")           ' ����
            .REALLEN = rs("REALLEN")         ' ������
            .BLOCKID = rs("BLOCKID")         ' �u���b�NID
            .KRPROCCD = rs("KRPROCCD")       ' ���݊Ǘ��H��
            .NOWPROC = rs("NOWPROC")         ' ���ݍH��
            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
            .DELCLS = rs("DELCLS")           ' �폜�敪
            .LSTATCLS = rs("LSTATCLS")       ' �ŏI��ԋ敪
            .RSTATCLS = rs("RSTATCLS")       ' ������ԋ敪
            .HOLDCLS = rs("HOLDCLS")         ' �z�[���h�敪
            .BDCAUS = rs("BDCAUS")           ' �s�Ǘ��R
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
            .PASSFLAG = " "   ' �ʉ߃t���O�̃X�y�[�X�N���A '02/07/05 hama
             If rs("PASSFLAG") = "1" Then
                .PASSFLAG = rs("PASSFLAG")   ' �ʉ߃t���O '02/07/05 hama
            End If

        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME040 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME041�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME041 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCME041_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME041(records() As typ_TBCME041, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACTORY, OPECOND, LENGTH, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME041"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME041 = FUNCTION_RETURN_FAILURE
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
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .LENGTH = rs("LENGTH")           ' ����
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME041 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uXSDCS�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_XSDCS    ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCME043_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME043(records() As typ_XSDCS, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
'    sqlBase = "Select CRYNUM, INGOTPOS, SMPKBN, SMPLNO, HINBAN, REVNUM, FACTORY, OPECOND, KTKBN, CRYINDRS, CRYINDOI, CRYINDB1," & _
'              " CRYINDB2, CRYINDB3, CRYINDL1, CRYINDL2, CRYINDL3, CRYINDL4, CRYINDCS, CRYINDGD, CRYINDT, CRYINDEP, CRYRESRS," & _
'              " CRYRESOI, CRYRESB1, CRYRESB2, CRYRESB3, CRYRESL1, CRYRESL2, CRYRESL3, CRYRESL4, CRYRESCS, CRYRESGD, CRYREST," & _
'              " CRYRESEP, SMPLNUM, SMPLPAT, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCME043"
    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, " & _
              " BLKKTFLAGCS, CRYSMPLIDRSCS, CRYSMPLIDRS1CS, CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, CRYRESRS2CS, CRYSMPLIDOICS, " & _
              " CRYINDOICS, CRYRESOICS, CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, " & _
              " CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS,  CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, " & _
              " CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, " & _
              " CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, CRYRESTCS, CRYSMPLIDEPCS, CRYINDEPCS,CRYRESEPCS, SMPLNUMCS, " & _
              " SMPLPATCS, TSTAFFCS, TDAYCS, KSTAFFCS, KDAYCS, SNDKCS, SNDDAYCS "
    sqlBase = sqlBase & "From XSDCS"

    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME043 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������ʒu
            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
            .SMPLNO = rs("SMPLNO")           ' �T���v��No
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .KTKBN = rs("KTKBN")             ' �m��敪
            .CRYINDRS = rs("CRYINDRS")       ' ���������w���iRs)
            .CRYINDOI = rs("CRYINDOI")       ' ���������w���iOi)
            .CRYINDB1 = rs("CRYINDB1")       ' ���������w���iB1)
            .CRYINDB2 = rs("CRYINDB2")       ' ���������w���iB2�j
            .CRYINDB3 = rs("CRYINDB3")       ' ���������w���iB3)
            .CRYINDL1 = rs("CRYINDL1")       ' ���������w���iL1)
            .CRYINDL2 = rs("CRYINDL2")       ' ���������w���iL2)
            .CRYINDL3 = rs("CRYINDL3")       ' ���������w���iL3)
            .CRYINDL4 = rs("CRYINDL4")       ' ���������w���iL4)
            .CRYINDCS = rs("CRYINDCS")       ' ���������w���iCs)
            .CRYINDGD = rs("CRYINDGD")       ' ���������w���iGD)
            .CRYINDT = rs("CRYINDT")         ' ���������w���iT)
            .CRYINDEP = rs("CRYINDEP")       ' ���������w���iEPD)
            .CRYRESRS = rs("CRYRESRS")       ' �����������сiRs)
            .CRYRESOI = rs("CRYRESOI")       ' �����������сiOi)
            .CRYRESB1 = rs("CRYRESB1")       ' �����������сiB1)
            .CRYRESB2 = rs("CRYRESB2")       ' �����������сiB2�j
            .CRYRESB3 = rs("CRYRESB3")       ' �����������сiB3)
            .CRYRESL1 = rs("CRYRESL1")       ' �����������сiL1)
            .CRYRESL2 = rs("CRYRESL2")       ' �����������сiL2)
            .CRYRESL3 = rs("CRYRESL3")       ' �����������сiL3)
            .CRYRESL4 = rs("CRYRESL4")       ' �����������сiL4)
            .CRYRESCS = rs("CRYRESCS")       ' �����������сiCs)
            .CRYRESGD = rs("CRYRESGD")       ' �����������сiGD)
            .CRYREST = rs("CRYREST")         ' �����������сiT)
            .CRYRESEP = rs("CRYRESEP")       ' �����������сiEPD)
            .SMPLNUM = rs("SMPLNUM")         ' �T���v������
            .SMPLPAT = rs("SMPLPAT")         ' �T���v���p�^�[��
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME043 = FUNCTION_RETURN_SUCCESS
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
    gErr.Push "s_cmbc018_SQL.bas -- Function fcmic001b_GetSpec"

    '' ���i�d�l�̎擾
    sql = "select "
    sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXCDOP,"     '4/2 Yam
    sql = sql & "HSXDPDIR, HSXDDMIN, HSXDDMAX, HSXSDSLP,"   '3/7 Yam
    sql = sql & "HSXCTCEN, HSXCYCEN "  '4/2 Yam
    sql = sql & " from TBCME018 A,TBCME020 B"
    sql = sql & " where A.HINBAN='" & pHinSpec.HIN.hinban & "'"
    sql = sql & " and A.MNOREVNO=" & pHinSpec.HIN.mnorevno
    sql = sql & " and A.FACTORY='" & pHinSpec.HIN.factory & "'"
    sql = sql & " and A.OPECOND='" & pHinSpec.HIN.opecond & "'"
    sql = sql & " and B.HINBAN='" & pHinSpec.HIN.hinban & "'"
    sql = sql & " and B.MNOREVNO=" & pHinSpec.HIN.mnorevno
    sql = sql & " and B.FACTORY='" & pHinSpec.HIN.factory & "'"
    sql = sql & " and B.OPECOND='" & pHinSpec.HIN.opecond & "'"
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

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMI002�v����Y�����錋���ԍ��̃��R�[�h������
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :crynum        ,I  ,string           ,�����ԍ�
'          :recCount      ,O  ,Integer          ,���R�[�h��
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,���o�̐���
'����      :
'����      :2002/08/09 H.FURUYA
Public Function DBDRV_GetTBCMI002(CRYNUM As String, recCount As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim rs As OraDynaset    'RecordSet

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_GetTBCMI002"
    DBDRV_GetTBCMI002 = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�
    sql = "Select * From TBCMI002 where CRYNUM ='" & Left(CRYNUM, 9) & "000'"

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    '���R�[�h�����Z�b�g
    recCount = rs.RecordCount

    '�������Z�b�g
    DBDRV_GetTBCMI002 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_GetTBCMI002 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMI003�v����Y�����錋���ԍ��̃��R�[�h������
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :crynum        ,I  ,string           ,�����ԍ�
'          :recCount      ,O  ,Integer          ,���R�[�h��
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,���o�̐���
'����      :
'����      :2002/08/09 H.FURUYA
Public Function DBDRV_GetTBCMI003(CRYNUM As String, recCount As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim rs As OraDynaset    'RecordSet

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_GetTBCMI003"
    '�����l�Z�b�g
    DBDRV_GetTBCMI003 = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�
    sql = "Select * From TBCMI003 where CRYNUM ='" & Left(CRYNUM, 9) & "000'"

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    '���R�[�h�����Z�b�g
    recCount = rs.RecordCount


    '�������Z�b�g
    DBDRV_GetTBCMI003 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_GetTBCMI003 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

