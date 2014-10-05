Attribute VB_Name = "s_cmbc034_SQL"
Option Explicit

' �v�e�Z���^�[���o

' �u���b�N�ꗗ
Public Type typ_BlkMap
    BLOCKID             As String * 12      ' �u���b�NID
    HIN(1 To 5)         As tFullHinban      ' �i��
    WFINDDATE           As String * 10      ' �ŏI�������t
    CRYNUM              As String * 12      ' �����ԍ�
    INGOTPOS            As Integer          ' �C���S�b�g���ʒu
    LENGTH              As Integer          ' �u���b�N����
    REALLEN             As Integer          ' �u���b�N������
    HINREALLEN(1 To 5)  As Integer          ' �i�Ԏ�����
    HinLen(1 To 5)      As Integer          ' �i�Ԓ���
    DIAMETER            As Double           ' ���a 2002/05/01 S.Sano
    sBlockID            As String * 12      ' �擪�u���b�NID
    BLOCKORDER          As Integer          ' �u���b�N����
    HOLDCLS             As String * 1       ' �z�[���h���  --- 2001/09/19 kuramoto �ǉ� ---
    PASSFLAG            As String * 1       ' �ʉ߃t���O�@�@--- 200/04/16 Yam
    AGRSTATUS           As String           ' ���F�m�F�敪      add SETkimizuka
    STOP                As String           ' ��~      add SETkimizuka
    CAUSE               As String           ' ��~���R  add SETkimizuka
    PRINTNO             As String           ' ��s�]��  add SETkimizuka
End Type

''�u���b�N���i�ԏ��(�\���i�Ԏ擾�p)�@�@--- 2007/07/17 �}���`�u���b�N�Ή� shindo
Public Type typ_WkBlkMap
    BLOCKID             As String * 12      ' �u���b�NID
    HINCNT As Integer
    HIN()         As tFullHinban      ' �i��
    HINREALLEN()  As Integer          ' �i�Ԏ�����
    HinLen()      As Integer          ' �i�Ԓ���
    INPOSCA() As Integer '�������J�n�ʒu
End Type

'�i�ԏ��--- 2007/07/17 �}���`�u���b�N�Ή� shindo
Public Wk_tblBlkMap() As typ_WkBlkMap

'�u���b�N���i�ԏ��
Public Type typ_BlkHinMap
    BLOCKID             As String * 12      ' �u���b�NID
    HIN                 As tFullHinban      ' �i��
    REALLEN             As Integer          ' �i�Ԏ�����
    HinLen              As Integer          ' ���i��
    PASSFLAG            As String * 1       ' �ʉ߃t���O
    INPOSCA             As Integer          ' �������J�n�ʒu�@--- 2007/07/17 shindo �ǉ� ---
    PLANTCATCA          As String           ' ���� 2007/09/12 SPK Tsutsumi Add
End Type

'�u���b�N�̏��
Public Type typ_BlkData
    CRYNUM              As String * 12      ' �����ԍ�
    BLOCKID             As String * 12      ' �u���b�NID
    INGOTPOS            As Integer          ' �C���S�b�g���ʒu
    LENGTH              As Integer          ' �u���b�N����
    REALLEN             As Integer          ' �u���b�N������
    sBlockID            As String * 12      ' ���o�擪�u���b�NID
    BLOCKORDER          As Integer          ' �u���b�N����
    DIAMETER            As Double           ' ���a 2002/05/01 S.Sano
    WFINDDATE           As String * 10      ' �ŏI�������t
    HOLDCLS             As String * 1       ' �z�[���h���
    AGRSTATUS           As String           ' ���F�m�F�敪      add SETkimizuka
    STOP                As String           ' ��~      add SETkimizuka
    CAUSE               As String           ' ��~���R  add SETkimizuka
    PRINTNO             As String           ' ��s�]��  add SETkimizuka
End Type


Public Type typ_KOSEIHIN
    KHINBAN As String * 10                  '�\���i��
    KHINPOS As Integer                      '�\���i��_�����J�n�ʒu
    KHINLEN As Integer                       '�\���i��_���ݒ���
End Type

''''''�T�v      :�v�e�Z���^�[���o ��ʕ\�����c�a�h���C�o
''''''���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
''''''      �@�@:pCryInf�@�@�@,O  ,typ_TBCME037   �@,�������
''''''      �@�@:pBlkMng�@�@�@,O  ,typ_TBCME040   �@,�u���b�N�Ǘ�
''''''      �@�@:pSXLMng�@�@�@,O  ,typ_TBCME042   �@,SXL�Ǘ�
''''''      �@�@:pBsInd �@�@�@,O  ,typ_TBCMW001   �@,�����w������
''''''      �@�@:pBlkForm�@ �@,O  ,typ_BlkForm    �@,�u���b�N�O�`���
''''''      �@�@:pBlkBad�@�@�@,O  ,typ_BlkBadPos  �@,�s�ǈʒu
''''''      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
''''''����      :
''''''����      :2001/07/12 �쐬 ���{
'''''Public Function DBDRV_scmzc_fcmkc001h_Disp(pCryInf() As typ_TBCME037, pBlkMng() As typ_TBCME040, _
'''''                                           pSXLMng() As typ_TBCME042, pBsInd() As typ_TBCMW001) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''
'''''    '' �G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_scmzc_fcmkc001h_Disp"
'''''
'''''    '' �������̎擾(s_cmzcTBCME037_SQL.bas ���K�v)
'''''    If DBDRV_GetTBCME037(pCryInf()) = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    '' �u���b�N�Ǘ��̎擾(s_cmzcTBCME040_SQL.bas ���K�v)
'''''    sql = " where NOWPROC='" & PROCD_WFC_HARAIDASI & "'"
'''''    sql = sql & " and LSTATCLS='T' order by CRYNUM, INGOTPOS"
'''''    If DBDRV_GetTBCME040(pBlkMng(), sql) = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    '' SXL�Ǘ��̎擾(s_cmzcTBCME042_SQL.bas ���K�v)
'''''    If DBDRV_GetTBCME042(pSXLMng(), sql) = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    '' �����w�����т̎擾(s_cmzcTBCMW001_SQL.bas ���K�v)
'''''    sql = " where TRANCNT=" & "any(select max(TRANCNT)"
'''''    sql = sql & " from TBCMW001 group by CRYNUM) order by CRYNUM, INGOTPOS"
'''''    If DBDRV_GetTBCMW001(pBsInd(), sql) = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''
'''''    DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '' �I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '' �G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_FAILURE
'''''    Resume proc_exit
'''''
'''''End Function


''''''
'''''' �����Ƃ̓����ɂ��33_SQL�́uFunction DBDRV_scmzc_fcmkc001h_Disp22�v�ɕύX�ڍs
''''''
''''''�T�v      :�v�e�Z���^�[���o ��ʕ\�����c�a�h���C�o (Step3.3��)
''''''���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
''''''      �@�@:pBlkData()   ,O  ,typ_BlkData      ,�u���b�N���
''''''      �@�@:pBlkHinMap() ,O  ,typ_BlkHinMap    ,�u���b�N�i�ԏ��
''''''      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
''''''����      :
''''''����      :2002/04/22 �쐬 �쑺
'''''Public Function DBDRV_scmzc_fcmkc001h_Disp2(pBlkData() As typ_BlkData, pBlkHinMap() As typ_BlkHinMap) As FUNCTION_RETURN
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''Dim recCnt As Long
'''''Dim i As Long
'''''Dim sBlkId As String
'''''Dim blkOrder As Integer
'''''    Dim Jiltuseki As Judg_Kakou '2002/05/01 S.Sano
'''''
'''''    '' �G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_scmzc_fcmkc001h_Disp2"
'''''
'''''    ''�u���b�N�̏����擾����
'''''    sql = "select B.CRYNUM, B.BLOCKID, B.INGOTPOS, B.LENGTH, B.REALLEN, B2.BLOCKID as SBLOCKID"
'''''    sql = sql & ", nvl("
'''''    sql = sql & "    (select DMTOP1 from TBCMI002 I2"
'''''    sql = sql & "     where CRYNUM=B.CRYNUM"
'''''    sql = sql & "       and INGOTPOS=(select max(INGOTPOS) from TBCMI002 where CRYNUM=B.CRYNUM and INGOTPOS<=B.INGOTPOS)"
'''''    sql = sql & "       and TRANCNT=(select max(TRANCNT) from TBCMI002 where CRYNUM=I2.CRYNUM and INGOTPOS=I2.INGOTPOS)"
'''''    sql = sql & "    )"
'''''    sql = sql & "    , (select DIAMETER from TBCME037 where CRYNUM=B.CRYNUM)"
'''''    sql = sql & "  ) as DIAM"
'''''    sql = sql & ", (select max(UPDDATE) from TBCMW001 where CRYNUM=B2.CRYNUM and INGOTPOS=B2.INGOTPOS) as NUKISHI_AT"
'''''    sql = sql & ", nvl((select HLDTRCLS from TBCMJ012 J12"
'''''    sql = sql & "       where CRYNUM=B.CRYNUM and INGOTPOS=B.INGOTPOS"
'''''    sql = sql & "         and TRANCNT=(select max(TRANCNT) from TBCMJ012 where CRYNUM=J12.CRYNUM and INGOTPOS=J12.INGOTPOS)"
'''''    sql = sql & "      ), '0'"
'''''    sql = sql & "  ) as HOLDCLS "
'''''    sql = sql & "from TBCME040 B, TBCME040 B2 "
'''''    sql = sql & "where B.DELCLS='0' and B.NOWPROC='CC720'"
'''''    sql = sql & "  and B2.CRYNUM=B.CRYNUM"
'''''    sql = sql & "  and B2.INGOTPOS=nvl("
'''''    sql = sql & "        (select max(BLK.INGOTPOS) from TBCME040 BLK, TBCME042 SXL"
'''''    sql = sql & "         where BLK.CRYNUM=B.CRYNUM and BLK.INGOTPOS<=B.INGOTPOS"
'''''    sql = sql & "           and SXL.CRYNUM=BLK.CRYNUM and SXL.INGOTPOS=BLK.INGOTPOS"
'''''    sql = sql & "        ), B.INGOTPOS) "
'''''    sql = sql & "order by B.CRYNUM, B.INGOTPOS"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
'''''    recCnt = rs.RecordCount
'''''    If recCnt <= 0 Then
'''''        ReDim pBlkData(0)
'''''    Else
'''''        ReDim pBlkData(1 To recCnt)
'''''        sBlkId = vbNullString
'''''        blkOrder = 0
'''''        For i = 1 To recCnt
'''''            With pBlkData(i)
'''''                .Crynum = rs("CRYNUM")
'''''                .BLOCKID = rs("BLOCKID")
'''''                .INGOTPOS = rs("INGOTPOS")
'''''                .LENGTH = rs("LENGTH")
'''''                .REALLEN = rs("REALLEN")
'''''                .sBlockID = rs("SBLOCKID")
'''''                If sBlkId <> .sBlockID Then
'''''                    sBlkId = .sBlockID
'''''                    blkOrder = 1
'''''                Else
'''''                    blkOrder = blkOrder + 1
'''''                End If
'''''                .BLOCKORDER = blkOrder
'''''                .DIAMETER = rs("DIAM")
'''''                If (vbNullString & rs("NUKISHI_AT")) = vbNullString Then
'''''                    .WFINDDATE = vbNullString
'''''                Else
'''''                    .WFINDDATE = Format$(rs("NUKISHI_AT"), "yyyy/mm/dd")
'''''                End If
'''''                .HOLDCLS = rs("HOLDCLS")
'''''            End With
'''''            rs.MoveNext
'''''        Next
''''''2002/05/01 S.Sano Start
'''''        rs.Close
'''''        For i = 1 To recCnt
'''''            With pBlkData(i)
'''''            If scmzc_getKakouJiltuseki(.BLOCKID, Jiltuseki) = FUNCTION_RETURN_SUCCESS Then
'''''                .DIAMETER = (Jiltuseki.TAIL(1) + Jiltuseki.TAIL(2) + Jiltuseki.TOP(1) + Jiltuseki.TOP(2)) / 4
'''''            End If
'''''            End With
'''''        Next
''''''2002/05/01 S.Sano End
'''''    End If
'''''
'''''
'''''
'''''    ''�u���b�N���̕i�ԍ\�����擾���� (�u���b�NID, �i��, ������, ���i��)
'''''    sql = "select BLOCKID, HINBAN, REVNUM, FACTORY, OPECOND, sum(REALLEN) as REALLEN, sum(HINLEN) as HINLEN "
'''''    sql = sql & ", PASSFLAG "
'''''    sql = sql & "from ("
'''''    sql = sql & "  select BLOCKID, HINBAN, REVNUM, FACTORY, OPECOND, REALLEN, HINFROM"
'''''    sql = sql & "  , REALLEN"
'''''    sql = sql & "    - case when BD1FROM<HINTO and BD1TO>HINFROM then least(HINTO,BD1TO)-greatest(HINFROM,BD1FROM) else 0 end"
'''''    sql = sql & "    - case when BD2FROM<HINTO and BD2TO>HINFROM then least(HINTO,BD2TO)-greatest(HINFROM,BD2FROM) else 0 end"
'''''    sql = sql & "    - case when BD3FROM<HINTO and BD3TO>HINFROM then least(HINTO,BD3TO)-greatest(HINFROM,BD3FROM) else 0 end"
'''''    sql = sql & "    - case when BD4FROM<HINTO and BD4TO>HINFROM then least(HINTO,BD4TO)-greatest(HINFROM,BD4FROM) else 0 end"
'''''    sql = sql & "    - case when BD5FROM<HINTO and BD5TO>HINFROM then least(HINTO,BD5TO)-greatest(HINFROM,BD5FROM) else 0 end"
'''''    sql = sql & "    as HINLEN"
'''''    sql = sql & "  , PASSFLAG"
'''''    sql = sql & "  from"
'''''    sql = sql & "  ("
'''''    sql = sql & "    select HINS.BLOCKID, HINBAN, REVNUM, FACTORY, OPECOND, HINS.HINFROM, HINS.HINTO, HINS.REALLEN"
'''''    sql = sql & "    , BD1FROM, BD1TO, BD2FROM, BD2TO, BD3FROM, BD3TO, BD4FROM, BD4TO, BD5FROM, BD5TO"
'''''    sql = sql & "    , HINS.PASSFLAG"
'''''    sql = sql & "    from"
'''''    sql = sql & "    (select BLK.CRYNUM, BLK.INGOTPOS, BLK.BLOCKID, SXL.HINBAN, REVNUM, FACTORY, OPECOND"
'''''    sql = sql & "      , greatest(BLK.INGOTPOS,SXL.INGOTPOS) as HINFROM"
'''''    sql = sql & "      , least(BLK.INGOTPOS+BLK.REALLEN,SXL.INGOTPOS+SXL.LENGTH) as HINTO"
'''''    sql = sql & "      , greatest(0,least(BLK.INGOTPOS+BLK.REALLEN,SXL.INGOTPOS+SXL.LENGTH) - greatest(BLK.INGOTPOS,SXL.INGOTPOS)) as REALLEN"
'''''    sql = sql & "      , BLK.PASSFLAG"
'''''    sql = sql & "      from TBCME040 BLK, TBCME042 SXL"
'''''    sql = sql & "      where BLK.DELCLS='0' and BLK.NOWPROC='CC720'"
'''''    sql = sql & "        and SXL.CRYNUM=BLK.CRYNUM"
'''''    sql = sql & "        and SXL.INGOTPOS<BLK.INGOTPOS+BLK.LENGTH"
'''''    sql = sql & "        and SXL.INGOTPOS+SXL.LENGTH>BLK.INGOTPOS"
'''''    sql = sql & "    ) HINS,"
'''''    sql = sql & "    (select B.CRYNUM, B.INGOTPOS"
'''''    sql = sql & "      , B.INGOTPOS + case when PART1=9999 then B.REALLEN-J.P1BDLEN else PART1 end as BD1FROM"
'''''    sql = sql & "      , B.INGOTPOS + case when PART1=9999 then B.REALLEN-J.P1BDLEN else PART1 end + P1BDLEN as BD1TO"
'''''    sql = sql & "      , B.INGOTPOS + case when PART2=9999 then B.REALLEN-J.P2BDLEN else PART2 end as BD2FROM"
'''''    sql = sql & "      , B.INGOTPOS + case when PART2=9999 then B.REALLEN-J.P2BDLEN else PART2 end + P2BDLEN as BD2TO"
'''''    sql = sql & "      , B.INGOTPOS + case when PART3=9999 then B.REALLEN-J.P3BDLEN else PART3 end as BD3FROM"
'''''    sql = sql & "      , B.INGOTPOS + case when PART3=9999 then B.REALLEN-J.P3BDLEN else PART3 end + P3BDLEN as BD3TO"
'''''    sql = sql & "      , B.INGOTPOS + case when PART4=9999 then B.REALLEN-J.P4BDLEN else PART4 end as BD4FROM"
'''''    sql = sql & "      , B.INGOTPOS + case when PART4=9999 then B.REALLEN-J.P4BDLEN else PART4 end + P4BDLEN as BD4TO"
'''''    sql = sql & "      , B.INGOTPOS + case when PART5=9999 then B.REALLEN-J.P5BDLEN else PART5 end as BD5FROM"
'''''    sql = sql & "      , B.INGOTPOS + case when PART5=9999 then B.REALLEN-J.P5BDLEN else PART5 end + P5BDLEN as BD5TO"
'''''    sql = sql & "      from TBCMJ010 J, TBCME040 B"
'''''    sql = sql & "      where B.DELCLS='0' and B.NOWPROC='CC720'"
'''''    sql = sql & "        and J.CRYNUM=B.CRYNUM and J.INGOTPOS=B.INGOTPOS"
'''''    sql = sql & "        and J.TRANCNT=(select max(TRANCNT) from TBCMJ010 where CRYNUM=J.CRYNUM and INGOTPOS=J.INGOTPOS)"
'''''    sql = sql & "    ) BADS"
'''''    sql = sql & "    where HINS.CRYNUM=BADS.CRYNUM"
'''''    sql = sql & "      and HINS.INGOTPOS=BADS.INGOTPOS"
'''''    sql = sql & "  )"
'''''    sql = sql & ")"
'''''    sql = sql & "group by BLOCKID, HINBAN, REVNUM, FACTORY, OPECOND "
'''''    sql = sql & ", PASSFLAG "
'''''    sql = sql & "order by BLOCKID, min(HINFROM)"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
'''''    recCnt = rs.RecordCount
'''''    If recCnt <= 0 Then
'''''        ReDim pBlkHinMap(0)
'''''    Else
'''''        ReDim pBlkHinMap(1 To recCnt)
'''''        For i = 1 To recCnt
'''''            With pBlkHinMap(i)
'''''                .BLOCKID = rs("BLOCKID")
'''''                .HIN.hinban = rs("HINBAN")
'''''                .HIN.mnorevno = rs("REVNUM")
'''''                .HIN.factory = rs("FACTORY")
'''''                .HIN.opecond = rs("OPECOND")
'''''                .REALLEN = rs("REALLEN")
'''''                .HINLEN = rs("HINLEN")
'''''                .PASSFLAG = vbNullString & rs("PASSFLAG")
'''''            End With
'''''            rs.MoveNext
'''''        Next
'''''    End If
'''''    rs.Close '2002/05/01 S.Sano
'''''
'''''    DBDRV_scmzc_fcmkc001h_Disp2 = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '' �I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '' �G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    DBDRV_scmzc_fcmkc001h_Disp2 = FUNCTION_RETURN_FAILURE
'''''    Resume proc_exit
'''''
'''''End Function



'�T�v      :�v�e�Z���^�[���o ���s���c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sStaffID�@�@�@,I  ,String         �@,�Ј�ID
'      �@�@:pBlkMap �@�@�@,I  ,typ_BlkMap     �@,�u���b�N�ꗗ
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_scmzc_fcmkc001h_Exec(ByVal sStaffID As String, pBlkMap() As typ_BlkMap, sErrMsg As String) As FUNCTION_RETURN

    Dim sql     As String
    Dim sDbName As String
    Dim recCnt  As Long
    Dim iPos    As Integer
    Dim i       As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_scmzc_fcmkc001h_Exec"
    sErrMsg = ""

    recCnt = UBound(pBlkMap)
    For i = 1 To recCnt
        '' �u���b�N�V�K���̑}��
        If DBDRV_BlockNewInf_Ins(pBlkMap(i), sDbName) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        With pBlkMap(i)
            '' �u���b�N�Ǘ��̍X�V
            sDbName = "E040"
            sql = "update TBCME040 set "
            sql = sql & "LPKRPROCCD='" & MGPRCD_WFC_HARAIDASI & "', "
            sql = sql & "LASTPASS  ='" & PROCD_WFC_HARAIDASI & "', "
            sql = sql & "DELCLS    ='1', "
            sql = sql & "LSTATCLS  ='W', "
            sql = sql & "UPDDATE   =sysdate, "
            sql = sql & "SENDFLAG  ='0'"
            sql = sql & " where CRYNUM='" & .CRYNUM & "' and INGOTPOS=" & .INGOTPOS

            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If

''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
'            '' SXL�Ǘ��̍X�V
'            sDbName = "E042"
'            iPos = .INGOTPOS + .LENGTH
'            sql = "update TBCME042 set "
'            sql = sql & "KRPROCCD  ='" & MGPRCD_WFC_SOUGOUHANTEI & "', "
'            sql = sql & "NOWPROC   ='" & PROCD_WFC_SOUGOUHANTEI & "', "
'            sql = sql & "LPKRPROCCD='" & MGPRCD_WFC_HARAIDASI & "', "
'            sql = sql & "LASTPASS  ='" & PROCD_WFC_HARAIDASI & "', "
'            sql = sql & "UPDDATE   =sysdate, "
'            sql = sql & "SENDFLAG  ='0'"
'            sql = sql & " where CRYNUM='" & .CRYNUM & "'"
'            sql = sql & " and INGOTPOS>=" & .INGOTPOS
'            sql = sql & " and INGOTPOS<" & iPos
'            If OraDB.ExecuteSQL(sql) < 0 Then
'                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
'                DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
'                GoTo proc_exit
'            End If
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{

            '' WF���o���т̑}��
            sDbName = "J011"
            sql = "insert into TBCMJ011 "
            sql = sql & "(CRYNUM,  INGOTPOS, LENGTH,         KRPROCCD, PROCCODE, "
            sql = sql & " BLOCKID, SBLOCKID, BLOCKORDER,     TSTAFFID, REGDATE, "

            '2007/08/31 SPK Tsutsumi Add Start
            sql = sql & " KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE,PLANTCAT)"
'            sql = sql & " KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE)"
            '2007/08/31 SPK Tsutsumi Add End

            sql = sql & " values ('"
            sql = sql & .CRYNUM & "', "                 ' �����ԍ�
            sql = sql & .INGOTPOS & ", "                ' �C���S�b�g���ʒu
            sql = sql & .LENGTH & ", '"                 ' ����
            sql = sql & MGPRCD_WFC_HARAIDASI & "', '"   ' �Ǘ��H���R�[�h
            sql = sql & PROCD_WFC_HARAIDASI & "', '"    ' �H���R�[�h
            sql = sql & .BLOCKID & "', '"               ' �u���b�NID
            sql = sql & .sBlockID & "', "               ' �擪�u���b�NID
            sql = sql & .BLOCKORDER & ", '"             ' �u���b�N����
            sql = sql & sStaffID & "', "                ' �o�^�Ј�ID
            sql = sql & "sysdate, '"                    ' �o�^���t
            sql = sql & sStaffID & "', "                ' �X�V�Ј�ID
            sql = sql & "sysdate, "                     ' �X�V���t
            sql = sql & "'0', "                         ' SUMMIT���M�t���O
            sql = sql & "'0', "                         ' ���M�t���O

            '2007/08/31 SPK Tsutsumi Add Start
            sql = sql & "sysdate,'"                      ' ���M���t
            sql = sql & sCmbMukesaki & "')"              ' ����
'            sql = sql & "sysdate)"                      ' ���M���t
            '2007/08/31 SPK Tsutsumi Add End

            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End With
    Next i

    '' 2003/04/22 ooba  WF�Z���^�[���o���ł̑��M�������~����
    '' ����]���w���̑��M��\�񂷂�
'    sDBName = "Y003"
'    sql = "update TBCMY003 set SENDFLAG='0' "
'    sql = sql & "where substr(SAMPLEID,1,12) in ("
'    recCnt = UBound(pBlkMap)
'    For i = 1 To recCnt
'        If i = recCnt Then
'            sql = sql & "'" & pBlkMap(i).BLOCKID & "'"
'        Else
'            sql = sql & "'" & pBlkMap(i).BLOCKID & "',"
'        End If
'    Next i
'    sql = sql & ")"
'    If OraDB.ExecuteSQL(sql) < 0 Then   '0���̓G���[�Ƃ��Ȃ�
'        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
'        DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
'        GoTo proc_exit
'    End If

    '�֘A��ۯ����o�^�@07/12/21 ooba START =====================================>
    If recCnt > 1 Then
        sDbName = "Y023"
        If DBDRV_KanrenBlk(pBlkMap()) = FUNCTION_RETURN_FAILURE Then

            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    '�֘A��ۯ����o�^�@07/12/21 ooba END =======================================>
    
    DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_SUCCESS

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
    DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�����֐��F�u���b�N�V�K���̍쐬�i�����w���t�j
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pBlkMap�@�@�@,I  ,typ_BlkMap     �@,�u���b�N�ꗗ
'      �@�@:sDBName�@�@�@,O  ,String         �@,DB����
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Private Function DBDRV_BlockNewInf_Ins(pBlkMap As typ_BlkMap, sDbName As String) As FUNCTION_RETURN

    Dim rs          As OraDynaset
    Dim sql         As String
    Dim CRYSTALMEN  As String
    Dim SEED        As Integer
    Dim TANMEN      As String * 3
    Dim WARPRANK    As String * 1
    Dim Ans         As String
    Dim MainHin     As tFullHinban      '��\�i�ԁ@05/11/25 ooba
    Dim SubHin      As tFullHinban      '��ޑ�\�i�ԁ@05/11/25 ooba
    Dim c0 As Integer
    Dim c1 As Integer
    Dim KOSEIHIN() As typ_KOSEIHIN
    Dim LENKEI As Integer
    Dim LENSA As Integer
    Dim KOSCNT As Integer               '�\���i�Ԑ�
    Dim FRKOSCNT As Integer             '�u���b�N�ɕR�t���i�Ԑ�
    Dim GNLC2 As Integer                '�u���b�N���ݒ��� 07/08/01 shindo
    Dim REALLC2 As Integer              '�u���b�N������ 07/08/01 shindo



    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_BlockNewInf_Ins"

    '' �V�[�h�X���̎擾
'��8���w���P�����������Ȃ� 2007/10/01 SETsw kubota
'    If Left(pBlkMap.CRYNUM, 1) = "8" Then
'        '' �w���P�����̏ꍇ
'        '�u���b�N�V�K���A�V�[�h�X���̋��ߕ���ύX
'        sDbName = "G002"
'        If DBDRV_getSEEDDEG(Trim(pBlkMap.BLOCKID), SEED) = FUNCTION_RETURN_FAILURE Then
'            DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    Else
        '' �w���P�����ȊO�̏ꍇ
        sDbName = "H004"
        If DBDRV_getSEED(pBlkMap.CRYNUM, SEED) = FUNCTION_RETURN_FAILURE Then
            DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
'    End If

''    '' �����ʂ̎擾
''    sDbName = "E022"
''    sql = "select HWFCDIR from TBCME022"
''    sql = sql & " where HINBAN='" & pBlkMap.HIN(1).hinban & "'"
''    sql = sql & " and MNOREVNO=" & pBlkMap.HIN(1).mnorevno
''    sql = sql & " and FACTORY='" & pBlkMap.HIN(1).factory & "'"
''    sql = sql & " and OPECOND='" & pBlkMap.HIN(1).opecond & "'"
''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''    If rs.RecordCount <= 0 Then
''        rs.Close
''        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
''        GoTo proc_exit
''    End If
''    CRYSTALMEN = rs("HWFCDIR")
''    rs.Close
''
''    If CRYSTALMEN = "B" Then
''        CRYSTALMEN = "100"
''    ElseIf CRYSTALMEN = "C" Then
''        CRYSTALMEN = "511"
''    ElseIf CRYSTALMEN = "D" Then
''        CRYSTALMEN = "110"
''    Else
''        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
''        GoTo proc_exit
''    End If

    '�u���b�N�V�K���
    '' �[�ʊp�x�����߂�
'    If DBDRV_getTANMEN(pBlkMap, ans) = FUNCTION_RETURN_FAILURE Then
    If DBDRV_getTANMEN(pBlkMap, SubHin, Ans) = FUNCTION_RETURN_FAILURE Then     '05/11/25 ooba
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    TANMEN = Ans
    '' ���[�v�����N�����߂�
'    If DBDRV_getWARPRANK(pBlkMap, ans) = FUNCTION_RETURN_FAILURE Then
    If DBDRV_getWARPRANK(pBlkMap, MainHin, Ans) = FUNCTION_RETURN_FAILURE Then  '05/11/25 ooba
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    WARPRANK = Ans

    '��ޑ�\�i�ԂŌ����ʂ��擾�@05/11/25 ooba START ================================>
    '' �����ʂ̎擾
    sDbName = "E022"
    sql = "select HWFCDIR from TBCME022"
    sql = sql & " where HINBAN='" & SubHin.hinban & "'"
    sql = sql & " and MNOREVNO=" & SubHin.mnorevno
    sql = sql & " and FACTORY='" & SubHin.factory & "'"
    sql = sql & " and OPECOND='" & SubHin.opecond & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then
        rs.Close
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    CRYSTALMEN = rs("HWFCDIR")
    rs.Close

    If CRYSTALMEN = "B" Then
        CRYSTALMEN = "100"
    ElseIf CRYSTALMEN = "C" Then
        CRYSTALMEN = "511"
    ElseIf CRYSTALMEN = "D" Then
        CRYSTALMEN = "110"
    Else
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    '��ޑ�\�i�ԂŌ����ʂ��擾�@05/11/25 ooba END ==================================>


    'Null �Ή��ɔ��Ȃ��C����Null�̏ꍇ�́h�O�h�ƌ��Ȃ��B�@�_�@����16�N10��8��
    'If DBDRV_NULLChk(pBlkMap) = FUNCTION_RETURN_FAILURE Then
    '    DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
    '    GoTo PROC_EXIT
    'End If

 '�\���i�Ԃ̏����擾 07/07/19 SHINDO STR=======================================>
    For c0 = 0 To UBound(Wk_tblBlkMap())

    If pBlkMap.BLOCKID = Wk_tblBlkMap(c0).BLOCKID Then
        '�u���b�N�ɕR�t���i�Ԑ����擾
        FRKOSCNT = UBound(Wk_tblBlkMap(c0).HIN)
        '�i�ԏ����擾
        For c1 = 1 To UBound(Wk_tblBlkMap(c0).HIN)
    ReDim Preserve KOSEIHIN(c1)
            With KOSEIHIN(c1)
                .KHINBAN = Wk_tblBlkMap(c0).HIN(c1).hinban + Format(Wk_tblBlkMap(c0).HIN(c1).mnorevno, "00")
                .KHINPOS = Wk_tblBlkMap(c0).INPOSCA(c1)
                .KHINLEN = Wk_tblBlkMap(c0).HinLen(c1)
            End With
        Next c1
    End If
    Next c0
  '�u���b�N�̎������A���ݒ������擾  07/08/01 shindo
       sDbName = "XSDC2"
       sql = "select GNLC2,REALLC2"
       sql = sql & " from XSDC2"
       sql = sql & " where CRYNUMC2='" & pBlkMap.BLOCKID & "'"
       Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
       If rs.RecordCount = 0 Then
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
       End If
       GNLC2 = rs("GNLC2")
       REALLC2 = rs("REALLC2")
       rs.Close

    '�������J�n�ʒu���u���b�N���J�n�ʒu�ɕύX

'07/08/01 shindo DELL_STR
    '�����̍��v�Z�o
'        LENKEI = 0
'        LENSA = 0
'07/08/01 shindo DELL_END

        For c1 = 1 To FRKOSCNT
            With KOSEIHIN(c1)
                .KHINPOS = .KHINPOS - pBlkMap.INGOTPOS
'07/08/01 shindo DELL
'               LENKEI = LENKEI + .KHINLEN
            End With
        Next c1
'07/08/01 shindo DELL
'        '�u�b�N�̐��i�����Ǝ������̍����Z�o
'        LENSA = pBlkMap.REALLEN - LENKEI

        If FRKOSCNT <= 5 Then
            KOSCNT = FRKOSCNT
            With KOSEIHIN(KOSCNT)
'07/08/01 shindo UPDATE_STR
'                .KHINLEN = .KHINLEN + LENSA
                .KHINLEN = .KHINLEN + (REALLC2 - GNLC2)
'07/08/01 shindo UPDATE_END
            End With
        Else
            KOSCNT = 5
            For c1 = 6 To FRKOSCNT
                    KOSEIHIN(5).KHINLEN = KOSEIHIN(5).KHINLEN + KOSEIHIN(c1).KHINLEN
            Next c1
'07/08/01 shindo UPDATE_STR
'            KOSEIHIN(5).KHINLEN = KOSEIHIN(5).KHINLEN + LENSA
            KOSEIHIN(5).KHINLEN = KOSEIHIN(5).KHINLEN + (REALLC2 - GNLC2)
'07/08/01 shindo UPDATE_END
        End If


 '�\���i�Ԃ̏����擾 07/07/19 SHINDO END=======================================<

    '' �u���b�N�V�K���̑}��
    sDbName = "Y001"
    With pBlkMap
        sql = "insert into TBCMY001 ("
        sql = sql & "BLOCKID, "         ' �u���b�NID
        sql = sql & "BLOCKLEN, "        ' �u���b�N�̒���
        sql = sql & "MAINHINBAN, "      ' ��\�i��
        sql = sql & "PNTYPE, "          ' �^�C�v
        sql = sql & "ROUP, "            ' ���R����l
        sql = sql & "ROLOW, "           ' ���R�����l
        sql = sql & "OIUP, "            ' �_�f�Z�x����l
        sql = sql & "OILOW, "           ' �_�f�Z�x�����l
        sql = sql & "TANMEN, "          ' �[�ʊp�x
        sql = sql & "WARPRANK, "        ' ���[�v�����N
        sql = sql & "CRYSTALMEN, "      ' ������
        sql = sql & "SLPCEN, "          ' �X���S
        sql = sql & "SLPLOW, "          ' �X����
        sql = sql & "SLPUP, "           ' �X���
        sql = sql & "INSPMETH, "        ' �������@
        sql = sql & "INSPFREQ, "        ' �����p�x
        sql = sql & "SLPDRC, "          ' �X����
        sql = sql & "SLPDRCAPP, "       ' �X���ʎw��
        sql = sql & "SLPHEIDRC, "       ' �X�c����
        sql = sql & "SLPHEICEN, "       ' �X�c���S
        sql = sql & "SLPHEILOW, "       ' �X�c����
        sql = sql & "SLPHEIUP, "        ' �X�c���
        sql = sql & "SLPWIDDRC, "       ' �X������
        sql = sql & "SLPWIDCEN, "       ' �X�����S
        sql = sql & "SLPWIDLOW, "       ' �X������
        sql = sql & "SLPWIDUP, "        ' �X�����
        sql = sql & "SEED, "            ' ���㎞�g�p�����V�|�h�X��
        sql = sql & "TXID, "            ' �g�����U�N�V����ID
        sql = sql & "SBLOCKID, "        ' �擪�u���b�NID
        sql = sql & "BLOCKORDER, "      ' �u���b�N����
        sql = sql & "REGDATE, "         ' �o�^���t
        sql = sql & "SENDFLAG, "        ' ���M�t���O
'2007/07/17 UPDATE_STR �}���`�u���b�N�Ή��@SHINDO
'        sql = sql & "SENDDATE)"         ' ���M���t
'****************
        sql = sql & "SENDDATE,"         ' ���M���t
        sql = sql & "PLANTCAT, "        ' ����  2007/08/31 SPK Tsutsumi Add
        sql = sql & "HINCNT, "          ' �\���i�Ԑ�"
        sql = sql & "MULUTIHINBAN1, "   ' �\���i�Ԃ��̂P�i��"
        sql = sql & "TOPICHI1, "        ' �\���i�Ԃ��̂PTop�ʒu(mm)"
        sql = sql & "TAILICHI1, "       ' �\���i�Ԃ��̂PTail�ʒu(mm)"
        sql = sql & "HINBANLEN1, "      ' �\���i�Ԃ��̂P����(mm)"
        sql = sql & "MULUTIHINBAN2, "   ' �\���i�Ԃ��̂Q�i��"
        sql = sql & "TOPICHI2, "        ' �\���i�Ԃ��̂QTop�ʒu(mm)"
        sql = sql & "TAILICHI2, "       ' �\���i�Ԃ��̂QTail�ʒu(mm)"
        sql = sql & "HINBANLEN2, "      ' �\���i�Ԃ��̂Q����(mm)"
        sql = sql & "MULUTIHINBAN3, "   ' �\���i�Ԃ��̂R�i��"
        sql = sql & "TOPICHI3, "        ' �\���i�Ԃ��̂RTop�ʒu(mm)"
        sql = sql & "TAILICHI3, "       ' �\���i�Ԃ��̂RTail�ʒu(mm)"
        sql = sql & "HINBANLEN3, "      ' �\���i�Ԃ��̂R����(mm)"
        sql = sql & "MULUTIHINBAN4, "   ' �\���i�Ԃ��̂S�i��"
        sql = sql & "TOPICHI4, "        ' �\���i�Ԃ��̂STop�ʒu(mm)"
        sql = sql & "TAILICHI4, "       ' �\���i�Ԃ��̂STail�ʒu(mm)"
        sql = sql & "HINBANLEN4, "      ' �\���i�Ԃ��̂S����(mm)"
        sql = sql & "MULUTIHINBAN5, "   ' �\���i�Ԃ��̂T�i��"
        sql = sql & "TOPICHI5, "        ' �\���i�Ԃ��̂TTop�ʒu(mm)"
        sql = sql & "TAILICHI5, "       ' �\���i�Ԃ��̂TTail�ʒu(mm)"
        sql = sql & "HINBANLEN5)"       ' �\���i�Ԃ��̂T����(mm)"

'2007/07/17 UPDATE_END �}���`�u���b�N�Ή��@SHINDO

        sql = sql & " select '"
        sql = sql & .BLOCKID & "', "                                            ' �u���b�NID
        sql = sql & .REALLEN & ", '"                                            ' �u���b�N�̒���
''        sql = sql & .HIN(1).hinban & Format(.HIN(1).mnorevno, "00") & "', "     ' ��\�i��
''        sql = sql & "E021HWFTYPE, "                                             ' �^�C�v
''        sql = sql & "case when E021HWFRMAX>=99999.9 then '99999.9'"
''        sql = sql & " when E021HWFRMAX>=9999.995 then to_char(round(E021HWFRMAX,2),'fm99990.0')"
''        sql = sql & " when E021HWFRMAX>=999.9995 then to_char(round(E021HWFRMAX,3),'fm9990.00')"
''        sql = sql & " when E021HWFRMAX>=99.99995 then to_char(round(E021HWFRMAX,4),'fm990.000')"
''        sql = sql & " when E021HWFRMAX>=10.00000 then to_char(round(E021HWFRMAX,5),'fm90.0000')"
''        sql = sql & " when E021HWFRMAX>=0.0 then to_char(E021HWFRMAX,'fm0.00000')"
''        sql = sql & " when nvl(E021HWFRMAX,0) = 0 then '0.0000'"
''        sql = sql & " else '-1.0000'"
''        sql = sql & "end as RMAX,"                                              ' ���R����l
''        sql = sql & "case when E021HWFRMIN>=99999.9 then '99999.9'"
''        sql = sql & " when E021HWFRMIN>=9999.995 then to_char(round(E021HWFRMIN,2),'fm99990.0')"
''        sql = sql & " when E021HWFRMIN>=999.9995 then to_char(round(E021HWFRMIN,3),'fm9990.00')"
''        sql = sql & " when E021HWFRMIN>=99.99995 then to_char(round(E021HWFRMIN,4),'fm990.000')"
''        sql = sql & " when E021HWFRMIN>=10.00000 then to_char(round(E021HWFRMIN,5),'fm90.0000')"
''        sql = sql & " when E021HWFRMIN>=0.0 then to_char(E021HWFRMIN,'fm0.00000')"
''        sql = sql & " when nvl(E021HWFRMIN,0) = 0 then '0.0000'"
''        sql = sql & " else '-1.0000'"
''        sql = sql & "end as RMIN,"                                              ' ���R�����l
''        sql = sql & "nvl(to_char(abs(E025HWFONMAX),'fm90.00'),'0.00'), "        ' �_�f�Z�x����l"
''        sql = sql & "nvl(to_char(abs(E025HWFONMIN),'fm90.00'),'0.00'), "        ' �_�f�Z�x�����l
''        sql = sql & "'" & TANMEN & "', "                                        ' �[�ʊp�x
''        sql = sql & "'" & WARPRANK & "', '"                                     ' ���[�v�����N
''        sql = sql & CRYSTALMEN & "', "                                          ' ������
''        sql = sql & "nvl(to_char(abs(E022HWFCSCEN),'fm0.00'),'0.00'), "         ' �X���S
''        sql = sql & "nvl(to_char(E022HWFCSMIN,'fm0.00'),'0.00'), "              ' �X����
''        sql = sql & "nvl(to_char(E022HWFCSMAX,'fm0.00'),'0.00'), "              ' �X���
''        sql = sql & "E022HWFCKWAY, "                                            ' �������@
''        sql = sql & "E022HWFCKHNM || E022HWFCKHNN || E022HWFCKHNH || E022HWFCKHNU, "    ' �����p�x�i���A���A�ہA�E�̏��ő����j
''        sql = sql & "E022HWFCSDIR, "                                            ' �X����
''        sql = sql & "E022HWFCSDIS, "                                            ' �X���ʎw��
''        sql = sql & "E022HWFCTDIR, "                                            ' �X�c����
''        sql = sql & "nvl(to_char(E022HWFCTCEN,'fm0.00'),'0.00'), "              ' �X�c���S
''        sql = sql & "nvl(to_char(E022HWFCTMIN,'fm0.00'),'0.00'), "              ' �X�c����
''        sql = sql & "nvl(to_char(E022HWFCTMAX,'fm0.00'),'0.00'), "              ' �X�c���
''        sql = sql & "E022HWFCYDIR, "                                            ' �X������
''        sql = sql & "nvl(to_char(E022HWFCYCEN,'fm0.00'),'0.00'), "              ' �X�����S
''        sql = sql & "nvl(to_char(E022HWFCYMIN,'fm0.00'),'0.00'), "              ' �X������
''        sql = sql & "nvl(to_char(E022HWFCYMAX,'fm0.00'),'0.00'), "              ' �X�����
''        sql = sql & SEED & ", "                                                 ' ���㎞�g�p�����V�|�h�X��
''        sql = sql & "'TX850I', '"                                               ' �g�����U�N�V����ID
''        sql = sql & .sBlockID & "', "                                           ' �擪�u���b�NID
''        sql = sql & .BLOCKORDER & ", "                                          ' �u���b�N����
''        sql = sql & "sysdate, "                                                 ' �o�^���t
''        sql = sql & "'0', "                                                     ' ���M�t���O
''        sql = sql & "sysdate"                                                   ' ���M���t
''        sql = sql & " from VECME001"
''        sql = sql & " where E018HINBAN='" & .HIN(1).hinban & "'"
''        sql = sql & " and E018MNOREVNO=" & .HIN(1).mnorevno
''        sql = sql & " and E018FACTORY='" & .HIN(1).factory & "'"
''        sql = sql & " and E018OPECOND='" & .HIN(1).opecond & "'"

        '��\�i�ԁA��ޑ�\�i�Ԃ̎d�l���擾�@05/11/25 ooba START ==============================>
        sql = sql & MainHin.hinban & Format(MainHin.mnorevno, "00") & "', "     ' ��\�i��
        sql = sql & "MAIN.E021HWFTYPE, "                                        ' �^�C�v
        sql = sql & "case when MAIN.E021HWFRMAX>=99999.9 then '99999.9'"
        sql = sql & " when MAIN.E021HWFRMAX>=9999.995 then to_char(round(MAIN.E021HWFRMAX,2),'fm99990.0')"
        sql = sql & " when MAIN.E021HWFRMAX>=999.9995 then to_char(round(MAIN.E021HWFRMAX,3),'fm9990.00')"
        sql = sql & " when MAIN.E021HWFRMAX>=99.99995 then to_char(round(MAIN.E021HWFRMAX,4),'fm990.000')"
        sql = sql & " when MAIN.E021HWFRMAX>=10.00000 then to_char(round(MAIN.E021HWFRMAX,5),'fm90.0000')"
        sql = sql & " when MAIN.E021HWFRMAX>=0.0 then to_char(MAIN.E021HWFRMAX,'fm0.00000')"
        sql = sql & " when nvl(MAIN.E021HWFRMAX,0) = 0 then '0.0000'"
        sql = sql & " else '-1.0000'"
        sql = sql & "end as RMAX,"                                              ' ���R����l
        sql = sql & "case when MAIN.E021HWFRMIN>=99999.9 then '99999.9'"
        sql = sql & " when MAIN.E021HWFRMIN>=9999.995 then to_char(round(MAIN.E021HWFRMIN,2),'fm99990.0')"
        sql = sql & " when MAIN.E021HWFRMIN>=999.9995 then to_char(round(MAIN.E021HWFRMIN,3),'fm9990.00')"
        sql = sql & " when MAIN.E021HWFRMIN>=99.99995 then to_char(round(MAIN.E021HWFRMIN,4),'fm990.000')"
        sql = sql & " when MAIN.E021HWFRMIN>=10.00000 then to_char(round(MAIN.E021HWFRMIN,5),'fm90.0000')"
        sql = sql & " when MAIN.E021HWFRMIN>=0.0 then to_char(MAIN.E021HWFRMIN,'fm0.00000')"
        sql = sql & " when nvl(MAIN.E021HWFRMIN,0) = 0 then '0.0000'"
        sql = sql & " else '-1.0000'"
        sql = sql & "end as RMIN,"                                              ' ���R�����l
        sql = sql & "nvl(to_char(abs(MAIN.E025HWFONMAX),'fm90.00'),'0.00'), "   ' �_�f�Z�x����l"
        sql = sql & "nvl(to_char(abs(MAIN.E025HWFONMIN),'fm90.00'),'0.00'), "   ' �_�f�Z�x�����l
        sql = sql & "'" & TANMEN & "', "                                        ' �[�ʊp�x
        sql = sql & "'" & WARPRANK & "', '"                                     ' ���[�v�����N
        sql = sql & CRYSTALMEN & "', "                                          ' ������
        sql = sql & "nvl(to_char(abs(SUB.E022HWFCSCEN),'fm0.00'),'0.00'), "     ' �X���S
        sql = sql & "nvl(to_char(SUB.E022HWFCSMIN,'fm0.00'),'0.00'), "          ' �X����
        sql = sql & "nvl(to_char(SUB.E022HWFCSMAX,'fm0.00'),'0.00'), "          ' �X���
        sql = sql & "SUB.E022HWFCKWAY, "                                        ' �������@
        sql = sql & "SUB.E022HWFCKHNM || SUB.E022HWFCKHNN || SUB.E022HWFCKHNH || SUB.E022HWFCKHNU, "    ' �����p�x�i���A���A�ہA�E�̏��ő����j
        sql = sql & "SUB.E022HWFCSDIR, "                                        ' �X����
        sql = sql & "SUB.E022HWFCSDIS, "                                        ' �X���ʎw��
        sql = sql & "SUB.E022HWFCTDIR, "                                        ' �X�c����
        sql = sql & "nvl(to_char(SUB.E022HWFCTCEN,'fm0.00'),'0.00'), "          ' �X�c���S
        sql = sql & "nvl(to_char(SUB.E022HWFCTMIN,'fm0.00'),'0.00'), "          ' �X�c����
        sql = sql & "nvl(to_char(SUB.E022HWFCTMAX,'fm0.00'),'0.00'), "          ' �X�c���
        sql = sql & "SUB.E022HWFCYDIR, "                                        ' �X������
        sql = sql & "nvl(to_char(SUB.E022HWFCYCEN,'fm0.00'),'0.00'), "          ' �X�����S
        sql = sql & "nvl(to_char(SUB.E022HWFCYMIN,'fm0.00'),'0.00'), "          ' �X������
        sql = sql & "nvl(to_char(SUB.E022HWFCYMAX,'fm0.00'),'0.00'), "          ' �X�����
        sql = sql & SEED & ", "                                                 ' ���㎞�g�p�����V�|�h�X��
        sql = sql & "'TX850I', '"                                               ' �g�����U�N�V����ID
        sql = sql & .sBlockID & "', "                                           ' �擪�u���b�NID
        sql = sql & .BLOCKORDER & ", "                                          ' �u���b�N����
        sql = sql & "sysdate, "                                                 ' �o�^���t
        sql = sql & "'0', "                                                     ' ���M�t���O
        sql = sql & "sysdate,"                                                  ' ���M���t
        sql = sql & "'" & sCmbMukesaki & "', "                                  ' ���� 2007/08/31 SPK Tsutsumi Add

    '�\���i�Ԃ�ǉ� 07/07/19 SHINDO STR=======================================>
        sql = sql & KOSCNT & ""
       For c0 = 1 To 5
        If c0 <= KOSCNT Then
            sql = sql & ",'" & KOSEIHIN(c0).KHINBAN & "'"
            sql = sql & ",'" & KOSEIHIN(c0).KHINPOS & "'"
            sql = sql & ",'" & KOSEIHIN(c0).KHINPOS + KOSEIHIN(c0).KHINLEN & "'"
            sql = sql & ",'" & KOSEIHIN(c0).KHINLEN & "'"
        Else
            sql = sql & ",NULL"
            sql = sql & ",NULL"
            sql = sql & ",NULL"
            sql = sql & ",NULL"
        End If
       Next c0

   '�\���i�Ԃ�ǉ� 07/07/19 SHINDO STR=======================================>


        sql = sql & " from ("
        sql = sql & "select "
        sql = sql & "E021HWFTYPE, "
        sql = sql & "E021HWFRMAX, "
        sql = sql & "E021HWFRMIN, "
        sql = sql & "E025HWFONMAX, "
        sql = sql & "E025HWFONMIN "
        sql = sql & "from VECME001 "
        sql = sql & "where E018HINBAN='" & MainHin.hinban & "' "
        sql = sql & "and E018MNOREVNO=" & MainHin.mnorevno & " "
        sql = sql & "and E018FACTORY='" & MainHin.factory & "' "
        sql = sql & "and E018OPECOND='" & MainHin.opecond & "' "
        sql = sql & ") MAIN, "
        sql = sql & "("
        sql = sql & "select "
        sql = sql & "E022HWFCSCEN, "
        sql = sql & "E022HWFCSMIN, "
        sql = sql & "E022HWFCSMAX, "
        sql = sql & "E022HWFCKWAY, "
        sql = sql & "E022HWFCKHNM, "
        sql = sql & "E022HWFCKHNN, "
        sql = sql & "E022HWFCKHNH, "
        sql = sql & "E022HWFCKHNU, "
        sql = sql & "E022HWFCSDIR, "
        sql = sql & "E022HWFCSDIS, "
        sql = sql & "E022HWFCTDIR, "
        sql = sql & "E022HWFCTCEN, "
        sql = sql & "E022HWFCTMIN, "
        sql = sql & "E022HWFCTMAX, "
        sql = sql & "E022HWFCYDIR, "
        sql = sql & "E022HWFCYCEN, "
        sql = sql & "E022HWFCYMIN, "
        sql = sql & "E022HWFCYMAX "
        sql = sql & "from VECME001 "
        sql = sql & "where E018HINBAN='" & SubHin.hinban & "' "
        sql = sql & "and E018MNOREVNO=" & SubHin.mnorevno & " "
        sql = sql & "and E018FACTORY='" & SubHin.factory & "' "
        sql = sql & "and E018OPECOND='" & SubHin.opecond & "' "
        sql = sql & ") SUB "
        '��\�i�ԁA��ޑ�\�i�Ԃ̎d�l���擾�@05/11/25 ooba END ================================>





    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    sDbName = ""
    DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�����֐��F�[�ʊp�x�����߂�
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pBlkMap�@�@�@,I  ,typ_BlkMap     �@,�u���b�N�ꗗ
'      �@�@:SubHinban�@�@,O  ,tFullHinban      ,��ޑ�\�i�ԁ@05/11/25 ooba
'      �@�@:ans    �@�@�@,O  ,String         �@,�[�ʊp�x
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :
'����      :2002/04/17  ���� �M�� �쐬
Public Function DBDRV_getTANMEN(pBlkMap As typ_BlkMap, SubHinban As tFullHinban, Ans As String) As FUNCTION_RETURN
    Dim rs              As OraDynaset
    Dim sql             As String
    Dim SQLHIN          As String
    Dim tHin(5)         As tFullHinban       ' �i��
    Dim tHSXCSCEN(5)    As String * 3
    Dim c0              As Integer
    Dim c1              As Integer
    Dim RecCount        As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_getTANMEN"
    DBDRV_getTANMEN = FUNCTION_RETURN_FAILURE

    SQLHIN = SQLMake_HINBAN(pBlkMap.HIN())
    'NULL�Ή��̂��߁AHSXCSMAX�EHSXCSMIN�̍��ڂ�ǉ�
    sql = "select HSXCSCEN, HSXCSMIN, HSXCSMAX, HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME018 where "
    sql = sql & "(ABS(HSXCSMAX - HSXCSMIN) = (select MIN(ABS(HSXCSMAX - HSXCSMIN)) from TBCME018 where " & SQLHIN & ")) and "
    sql = sql & SQLHIN
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    RecCount = rs.RecordCount
    If RecCount <= 0 Then
        rs.Close
        GoTo proc_exit
    End If

    For c0 = 1 To RecCount
        If IsNull(rs("HSXCSCEN")) Or IsNull(rs("HSXCSMIN")) Or IsNull(rs("HSXCSMAX")) Then
            DBDRV_getTANMEN = FUNCTION_RETURN_FAILURE
            GoTo proc_err
        End If
        tHSXCSCEN(c0) = fncNullCheck(rs("HSXCSCEN"))
        tHin(c0).factory = rs("FACTORY")
        tHin(c0).hinban = rs("HINBAN")
        tHin(c0).mnorevno = rs("MNOREVNO")
        tHin(c0).opecond = rs("OPECOND")
        rs.MoveNext
    Next
    rs.Close
    '�������݂����ꍇ�A�ł��㑤�̕i�Ԃ��̗p���[�ʊp�x�����߂�B
    Ans = ""
'    For c0 = 1 To RecCount
    For c0 = 1 To 5     '06/01/19 ooba
        If Trim(pBlkMap.HIN(c0).hinban) <> "" Then
            For c1 = 1 To RecCount
                If (pBlkMap.HIN(c0).factory = tHin(c1).factory) And _
                   (pBlkMap.HIN(c0).hinban = tHin(c1).hinban) And _
                   (pBlkMap.HIN(c0).mnorevno = tHin(c1).mnorevno) And _
                   (pBlkMap.HIN(c0).opecond = tHin(c1).opecond) Then
                    Ans = tHSXCSCEN(c1)
                    SubHinban.hinban = tHin(c1).hinban          '05/11/25 ooba START =====>
                    SubHinban.mnorevno = tHin(c1).mnorevno
                    SubHinban.factory = tHin(c1).factory
                    SubHinban.opecond = tHin(c1).opecond        '05/11/25 ooba END =======>
                    Exit For
                End If
            Next
        End If
        If Ans <> "" Then Exit For
    Next
    DBDRV_getTANMEN = FUNCTION_RETURN_SUCCESS
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

'�T�v      :�����֐��F���[�v�����N�����߂�
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pBlkMap�@�@�@,I  ,typ_BlkMap     �@,�u���b�N�ꗗ
'      �@�@:MainHinban�@ ,O  ,tFullHinban      ,��\�i�ԁ@05/11/25 ooba
'      �@�@:ans    �@�@�@,O  ,String         �@,�[�ʊp�x
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :
'����      :2002/04/17  ���� �M�� �쐬
Public Function DBDRV_getWARPRANK(pBlkMap As typ_BlkMap, MainHinban As tFullHinban, Ans As String) As FUNCTION_RETURN
    Dim rs  As OraDynaset
    Dim sql As String
    Dim SQLHIN          As String           '05/11/25 ooba START ========>
    Dim tHin(5)         As tFullHinban
    Dim c0              As Integer
    Dim c1              As Integer
    Dim RecCount        As Integer          '05/11/25 ooba END ==========>

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_getWARPRANK"
    DBDRV_getWARPRANK = FUNCTION_RETURN_FAILURE

    '�������@06/04/28 ooba
    For c0 = 1 To 5
        tHin(c0).hinban = ""
        tHin(c0).mnorevno = 0
        tHin(c0).factory = ""
        tHin(c0).opecond = ""
    Next c0

''    sql = "select max(HWFWARPR) as maxHWFWARPR from TBCME027"
''    sql = sql & " where " & SQLMake_HINBAN(pBlkMap.HIN())
''
''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''    If rs.RecordCount <= 0 Then
''        rs.Close
''        GoTo proc_exit
''    End If
''    ans = rs("maxHWFWARPR")
''    rs.Close

    'ܰ���ݸ���ő�̕i�Ԃ��\�i�ԂƂ���@05/11/25 ooba START ==============================>
    SQLHIN = SQLMake_HINBAN(pBlkMap.HIN())

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, HWFWARPR from TBCME027 where "
''    sql = sql & "HWFWARPR = (select MAX(HWFWARPR) from TBCME027 where " & SQLHIN & ") and "
''    sql = sql & SQLHIN
    '�@������׸�(0:��׽�ڒ�����,1:��׽�ڒ��L��)���ő�ȕi�Ԃ̒���           06/01/19 ooba
    '�����p�̋K�i��(�����ʌX���-�����ʌX����)���ŏ��̕i�Ԃ̒���            06/07/19 kondoh Add
    'ܰ���ݸ���ő�̕i��                                                    06/01/19 ooba
    sql = sql & "HWFWARPR = (select MAX(HWFWARPR) from TBCME027 "
    sql = sql & "            where (HINBAN, MNOREVNO, FACTORY, OPECOND) in "

''06/07/19 SMP)kondoh START Add =========================================================>
    sql = sql & "               ("
    sql = sql & "               select HINBAN, MNOREVNO, FACTORY, OPECOND "
    sql = sql & "               from TBCME018 "
    sql = sql & "               where ABS(HSXCSMAX - HSXCSMIN) = "
    sql = sql & "                       (select MIN(ABS(HSXCSMAX - HSXCSMIN)) from TBCME018 "
    sql = sql & "                      where (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
    sql = sql & "                           (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
    sql = sql & "                           where decode(GLASS,null,'0',' ','0',GLASS) = "
    sql = sql & "                               (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
    sql = sql & "                               from TBCME036 where " & SQLHIN
    sql = sql & "                               ) "
    sql = sql & "                           and " & SQLHIN
    sql = sql & "                           ) "
    sql = sql & "                       ) "
    sql = sql & "                and (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
    sql = sql & "                  (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
    sql = sql & "                   where decode(GLASS,null,'0',' ','0',GLASS) = "
    sql = sql & "                         (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
    sql = sql & "                          from TBCME036 where " & SQLHIN
    sql = sql & "                         )"
    sql = sql & "                   and " & SQLHIN
    sql = sql & "                  ) "
    sql = sql & "               ) "
''06/07/19 SMP)kondoh END Add =========================================================>

''06/07/19 SMP)kondoh START Del =========================================================>
''    sql = sql & "                  (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
''    sql = sql & "                   where decode(GLASS,null,'0',' ','0',GLASS) = "
''    sql = sql & "                         (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
''    sql = sql & "                          from TBCME036 where " & SQLHIN
''    sql = sql & "                         ) "
''    sql = sql & "                   and " & SQLHIN
''    sql = sql & "                  ) "
''06/07/19 SMP)kondoh END Del =========================================================>

    sql = sql & "           ) "
    sql = sql & "and (HINBAN, MNOREVNO, FACTORY, OPECOND) in "

''06/07/19 SMP)kondoh START Add =========================================================>
    sql = sql & "               ("
    sql = sql & "               select HINBAN, MNOREVNO, FACTORY, OPECOND "
    sql = sql & "               from TBCME018 "
    sql = sql & "               where ABS(HSXCSMAX - HSXCSMIN) = "
    sql = sql & "                       (select MIN(ABS(HSXCSMAX - HSXCSMIN)) from TBCME018 "
    sql = sql & "                      where (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
    sql = sql & "                           (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
    sql = sql & "                           where decode(GLASS,null,'0',' ','0',GLASS) = "
    sql = sql & "                               (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
    sql = sql & "                               from TBCME036 where " & SQLHIN
    sql = sql & "                               ) "
    sql = sql & "                           and " & SQLHIN
    sql = sql & "                           ) "
    sql = sql & "                       ) "
    sql = sql & "                and (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
    sql = sql & "                  (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
    sql = sql & "                   where decode(GLASS,null,'0',' ','0',GLASS) = "
    sql = sql & "                         (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
    sql = sql & "                          from TBCME036 where " & SQLHIN
    sql = sql & "                         )"
    sql = sql & "                   and " & SQLHIN
    sql = sql & "                  ) "
    sql = sql & "               ) "
''06/07/19 SMP)kondoh END Add =========================================================>

''06/07/19 SMP)kondoh START Del =========================================================>
''    sql = sql & "    (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
''    sql = sql & "     where decode(GLASS,null,'0',' ','0',GLASS) = "
''    sql = sql & "           (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
''    sql = sql & "            from TBCME036 where " & SQLHIN
''    sql = sql & "           ) "
''    sql = sql & "     and " & SQLHIN
''    sql = sql & "    )"
''06/07/19 SMP)kondoh END Del =========================================================>

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    RecCount = rs.RecordCount
    If RecCount <= 0 Then
        rs.Close
        GoTo proc_exit
    End If

    Ans = rs("HWFWARPR")

    For c0 = 1 To RecCount
        tHin(c0).hinban = rs("HINBAN")
        tHin(c0).factory = rs("FACTORY")
        tHin(c0).mnorevno = rs("MNOREVNO")
        tHin(c0).opecond = rs("OPECOND")
        rs.MoveNext
    Next
    rs.Close


''06/07/19 SMP)kondoh START Del =========================================================>
''    ''06/04/28 ooba START =========================================================>
''    '�@�𖞂�����������ߋK�i����Ԍ�����(�iWF�����2�������ԏ�����)�i��
''    SQLHIN = SQLMake_HINBAN(tHin())
''
''    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, HWFNP2MAX "
''    sql = sql & "from TBCME026 "
''    sql = sql & "where nvl(HWFNP2MAX,999.99) = (select min(nvl(HWFNP2MAX,999.99)) "
''    sql = sql & "                               from TBCME026 "
''    sql = sql & "                               where " & SQLHIN & ") "
''    sql = sql & "and " & SQLHIN
''
''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''    RecCount = rs.RecordCount
''    If RecCount <= 0 Then
''        rs.Close
''        GoTo proc_exit
''    End If
''
''    For c0 = 1 To RecCount
''        tHin(c0).hinban = rs("HINBAN")
''        tHin(c0).factory = rs("FACTORY")
''        tHin(c0).mnorevno = rs("MNOREVNO")
''        tHin(c0).opecond = rs("OPECOND")
''        rs.MoveNext
''    Next
''    rs.Close
''    ''06/04/28 ooba END ===========================================================>
''06/07/19 SMP)kondoh END Del =========================================================>

    MainHinban.hinban = ""
    '�������݂����ꍇ�A�ł��㑤�̕i�Ԃ��\�i�ԂƂ���B
'    For c0 = 1 To RecCount
    For c0 = 1 To 5     '06/01/19 ooba
        If Trim(pBlkMap.HIN(c0).hinban) <> "" Then
            For c1 = 1 To RecCount
                If (pBlkMap.HIN(c0).hinban = tHin(c1).hinban) And _
                   (pBlkMap.HIN(c0).mnorevno = tHin(c1).mnorevno) And _
                   (pBlkMap.HIN(c0).factory = tHin(c1).factory) And _
                   (pBlkMap.HIN(c0).opecond = tHin(c1).opecond) Then

                    MainHinban.hinban = tHin(c1).hinban
                    MainHinban.mnorevno = tHin(c1).mnorevno
                    MainHinban.factory = tHin(c1).factory
                    MainHinban.opecond = tHin(c1).opecond
                    Exit For
                End If
            Next
        End If
        If Trim(MainHinban.hinban) <> "" Then Exit For
    Next
    'ܰ���ݸ���ő�̕i�Ԃ��\�i�ԂƂ���@05/11/25 ooba END ================================>

    DBDRV_getWARPRANK = FUNCTION_RETURN_SUCCESS
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

'�T�v      :�����֐��F�d�l�l��NULL�������ꍇ�AInsert�������s�ł��Ȃ��悤�ɃG���[�ŏI������
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pBlkMap�@�@�@,I  ,typ_BlkMap     �@,�u���b�N�ꗗ
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :
'����      :2003/12/12  �V�X�e���u���C�� �쐬
Public Function DBDRV_NULLChk(pBlkMap As typ_BlkMap) As FUNCTION_RETURN
    Dim rs  As OraDynaset
    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_NULLChk"
    DBDRV_NULLChk = FUNCTION_RETURN_FAILURE

    'NUMBER�^�̃f�[�^��ǂݍ��݁ANULL�������ꍇ�̓G���[�Ƃ���
    sql = ""
    With pBlkMap
        sql = sql & "select E021HWFRMAX, E021HWFRMIN,"
        sql = sql & "       E025HWFONMAX, E025HWFONMIN,"
        sql = sql & "       E022HWFCSCEN, E022HWFCSMIN, E022HWFCSMAX,"
        sql = sql & "       E022HWFCTCEN, E022HWFCTMIN, E022HWFCTMAX,"
        sql = sql & "       E022HWFCYCEN, E022HWFCYMIN, E022HWFCYMAX from VECME001"
        sql = sql & " where E018HINBAN='" & .HIN(1).hinban & "'"
        sql = sql & " and E018MNOREVNO=" & .HIN(1).mnorevno
        sql = sql & " and E018FACTORY='" & .HIN(1).factory & "'"
        sql = sql & " and E018OPECOND='" & .HIN(1).opecond & "'"
    End With

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '�擾�G���[�A�܂��͂P�ł�NULL����������G���[�Ƃ���
    If rs.RecordCount <= 0 Or _
       IsNull(rs("E021HWFRMAX")) Or IsNull(rs("E021HWFRMIN")) Or _
       IsNull(rs("E025HWFONMAX")) Or IsNull(rs("E025HWFONMIN")) Or _
       IsNull(rs("E022HWFCSCEN")) Or IsNull(rs("E022HWFCSMIN")) Or IsNull(rs("E022HWFCSMAX")) Or _
       IsNull(rs("E022HWFCTCEN")) Or IsNull(rs("E022HWFCTMIN")) Or IsNull(rs("E022HWFCTMAX")) Or _
       IsNull(rs("E022HWFCYCEN")) Or IsNull(rs("E022HWFCYMIN")) Or IsNull(rs("E022HWFCYMAX")) Then

        DBDRV_NULLChk = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If

    DBDRV_NULLChk = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    rs.Close
    gErr.Pop
    Exit Function
proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
''''''------------------------------------------------
'''''' DB�A�N�Z�X�֐�
''''''------------------------------------------------
'''''
''''''�T�v      :�e�[�u���uTBCME037�v��������ɂ��������R�[�h�𒊏o����
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :records()     ,O  ,typ_TBCME037 ,���o���R�[�h
''''''          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
''''''          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
''''''          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
''''''����      :
''''''����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcF_TBCME037_SQL.bas���ړ�)
'''''Public Function DBDRV_GetTBCME037(records() As typ_TBCME037, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
'''''Dim sql As String       'SQL�S��
'''''Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
'''''Dim rs As OraDynaset    'RecordSet
'''''Dim recCnt As Long      '���R�[�h��
'''''Dim i As Long
'''''
'''''    ''SQL��g�ݗ��Ă�
'''''    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
'''''              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
'''''              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
'''''    sqlBase = sqlBase & "From TBCME037"
'''''    sql = sqlBase
'''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
'''''        sql = sql & " " & sqlWhere & " " & sqlOrder
'''''    End If
'''''
'''''    ''�f�[�^�𒊏o����
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'''''    If rs Is Nothing Then
'''''        ReDim records(0)
'''''        DBDRV_GetTBCME037 = FUNCTION_RETURN_FAILURE
'''''        Exit Function
'''''    End If
'''''
'''''    ''���o���ʂ��i�[����
'''''    recCnt = rs.RecordCount
'''''    ReDim records(recCnt)
'''''    For i = 1 To recCnt
'''''        With records(i)
'''''            .Crynum = rs("CRYNUM")           ' �����ԍ�
'''''            .DELCLS = rs("DELCLS")           ' �폜�敪
'''''            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
'''''            .PROCCD = rs("PROCCD")           ' �H���R�[�h
'''''            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
'''''            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
'''''            .RPHINBAN = rs("RPHINBAN")       ' �˂炢�i��
'''''            .RPREVNUM = rs("RPREVNUM")       ' �˂炢�i�Ԑ��i�ԍ������ԍ�
'''''            .RPFACT = rs("RPFACT")           ' �˂炢�i�ԍH��
'''''            .RPOPCOND = rs("RPOPCOND")       ' �˂炢�i�ԑ��Ə���
'''''            .PRODCOND = rs("PRODCOND")       ' �������
'''''            .PGID = rs("PGID")               ' �o�f�|�h�c
'''''            .UPLENGTH = rs("UPLENGTH")       ' ���グ����
'''''            .TOPLENG = rs("TOPLENG")         ' �s�n�o����
'''''            .BODYLENG = rs("BODYLENG")       ' ��������
'''''            .BOTLENG = rs("BOTLENG")         ' �a�n�s����
'''''            .FREELENG = rs("FREELENG")       ' �t���[��
'''''            .DIAMETER = rs("DIAMETER")       ' ���a
'''''            .CHARGE = rs("CHARGE")           ' �`���[�W��
'''''            .SEED = rs("SEED")               ' �V�[�h
'''''            .ADDDPCLS = rs("ADDDPCLS")       ' �ǉ��h�[�v���
'''''            .ADDDPPOS = rs("ADDDPPOS")       ' �ǉ��h�[�v�ʒu
'''''            .ADDDPVAL = rs("ADDDPVAL")       ' �ǉ��h�[�v��
'''''            .REGDATE = rs("REGDATE")         ' �o�^���t
'''''            .UPDDATE = rs("UPDDATE")         ' �X�V���t
'''''            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
'''''            .SENDDATE = rs("SENDDATE")       ' ���M���t
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    DBDRV_GetTBCME037 = FUNCTION_RETURN_SUCCESS
'''''End Function
'''''
'''''
''''''------------------------------------------------
'''''' DB�A�N�Z�X�֐�
''''''------------------------------------------------
'''''
''''''�T�v      :�e�[�u���uTBCME040�v��������ɂ��������R�[�h�𒊏o����
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :records()     ,O  ,typ_TBCME040 ,���o���R�[�h
''''''          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
''''''          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
''''''          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
''''''����      :
''''''����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCME040_SQL.bas���ړ�)
'''''Public Function DBDRV_GetTBCME040(records() As typ_TBCME040, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
'''''Dim sql As String       'SQL�S��
'''''Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
'''''Dim rs As OraDynaset    'RecordSet
'''''Dim recCnt As Long      '���R�[�h��
'''''Dim i As Long
'''''
'''''    ''SQL��g�ݗ��Ă�
'''''    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, REALLEN, BLOCKID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS," & _
'''''              " RSTATCLS, HOLDCLS, BDCAUS, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE," & _
'''''              " PASSFLAG "   '02/07/05 hama
'''''
'''''    sqlBase = sqlBase & "From TBCME040"
'''''    sql = sqlBase
'''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
'''''        sql = sql & " " & sqlWhere & " " & sqlOrder
'''''    End If
'''''
'''''    ''�f�[�^�𒊏o����
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'''''    If rs Is Nothing Then
'''''        ReDim records(0)
'''''        DBDRV_GetTBCME040 = FUNCTION_RETURN_FAILURE
'''''        Exit Function
'''''    End If
'''''
'''''    ''���o���ʂ��i�[����
'''''    recCnt = rs.RecordCount
'''''    ReDim records(recCnt)
'''''    For i = 1 To recCnt
'''''        With records(i)
'''''            .Crynum = rs("CRYNUM")           ' �����ԍ�
'''''            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
'''''            .LENGTH = rs("LENGTH")           ' ����
'''''            .REALLEN = rs("REALLEN")         ' ������
'''''            .BLOCKID = rs("BLOCKID")         ' �u���b�NID
'''''            .KRPROCCD = rs("KRPROCCD")       ' ���݊Ǘ��H��
'''''            .NOWPROC = rs("NOWPROC")         ' ���ݍH��
'''''            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
'''''            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
'''''            .DELCLS = rs("DELCLS")           ' �폜�敪
'''''            .LSTATCLS = rs("LSTATCLS")       ' �ŏI��ԋ敪
'''''            .RSTATCLS = rs("RSTATCLS")       ' ������ԋ敪
'''''            .HOLDCLS = rs("HOLDCLS")         ' �z�[���h�敪
'''''            .BDCAUS = rs("BDCAUS")           ' �s�Ǘ��R
'''''            .REGDATE = rs("REGDATE")         ' �o�^���t
'''''            .UPDDATE = rs("UPDDATE")         ' �X�V���t
'''''            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
'''''            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
'''''            .SENDDATE = rs("SENDDATE")       ' ���M���t
'''''            .PASSFLAG = " "   ' �ʉ߃t���O�̃X�y�[�X�N���A '02/07/05 hama
'''''             If rs("PASSFLAG") = "1" Then
'''''                .PASSFLAG = rs("PASSFLAG")   ' �ʉ߃t���O '02/07/05 hama
'''''            End If
'''''
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    DBDRV_GetTBCME040 = FUNCTION_RETURN_SUCCESS
'''''End Function
'''''
'''''
''''''------------------------------------------------
'''''' DB�A�N�Z�X�֐�
''''''------------------------------------------------
'''''
''''''�T�v      :�e�[�u���uTBCME042�v��������ɂ��������R�[�h�𒊏o����
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :records()     ,O  ,typ_TBCME042 ,���o���R�[�h
''''''          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
''''''          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
''''''          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
''''''����      :
''''''����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCME042_SQL.bas���ړ�)
'''''Public Function DBDRV_GetTBCME042(records() As typ_TBCME042, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
'''''Dim sql As String       'SQL�S��
'''''Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
'''''Dim rs As OraDynaset    'RecordSet
'''''Dim recCnt As Long      '���R�[�h��
'''''Dim i As Long
'''''
'''''    ''SQL��g�ݗ��Ă�
'''''    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, SXLID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS, HOLDCLS," & _
'''''              " HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, COUNT, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE, " & _
'''''              " PASSFLAG "   '02/04/16 Yam
'''''    sqlBase = sqlBase & "From TBCME042"
'''''    sql = sqlBase
'''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
'''''        sql = sql & " " & sqlWhere & " " & sqlOrder
'''''    End If
'''''
'''''    ''�f�[�^�𒊏o����
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'''''    If rs Is Nothing Then
'''''        ReDim records(0)
'''''        DBDRV_GetTBCME042 = FUNCTION_RETURN_FAILURE
'''''        Exit Function
'''''    End If
'''''
'''''    ''���o���ʂ��i�[����
'''''    recCnt = rs.RecordCount
'''''    ReDim records(recCnt)
'''''    For i = 1 To recCnt
'''''        With records(i)
'''''            .Crynum = rs("CRYNUM")           ' �����ԍ�
'''''            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
'''''            .LENGTH = rs("LENGTH")           ' ����
'''''            .SXLID = rs("SXLID")             ' SXLID
'''''            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H��
'''''            .NOWPROC = rs("NOWPROC")         ' ���ݍH��
'''''            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
'''''            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
'''''            .DELCLS = rs("DELCLS")           ' �폜�敪
'''''            .LSTATCLS = rs("LSTATCLS")       ' �ŏI��ԋ敪
'''''            .HOLDCLS = rs("HOLDCLS")         ' �z�[���h�敪
'''''            .hinban = rs("HINBAN")           ' �i��
'''''            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
'''''            .factory = rs("FACTORY")         ' �H��
'''''            .opecond = rs("OPECOND")         ' ���Ə���
'''''            .BDCAUS = rs("BDCAUS")           ' �s�Ǘ��R
'''''            .COUNT = rs("COUNT")             ' ����
'''''            .REGDATE = rs("REGDATE")         ' �o�^���t
'''''            .UPDDATE = rs("UPDDATE")         ' �X�V���t
'''''            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
'''''            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
'''''            .SENDDATE = rs("SENDDATE")       ' ���M���t
'''''            .PASSFLAG = " "   ' �ʉ߃t���O�̃X�y�[�X�N���A '02/04/16 Yam
'''''            If rs("PASSFLAG") = "1" Then
'''''                .PASSFLAG = rs("PASSFLAG")   ' �ʉ߃t���O '02/04/05 Yam
'''''            End If
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    DBDRV_GetTBCME042 = FUNCTION_RETURN_SUCCESS
'''''End Function
'''''
'''''
''''''------------------------------------------------
'''''' DB�A�N�Z�X�֐�
''''''------------------------------------------------
'''''
''''''�T�v      :�e�[�u���uTBCMW001�v��������ɂ��������R�[�h�𒊏o����
''''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''''''          :records()     ,O  ,typ_TBCMW001 ,���o���R�[�h
''''''          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
''''''          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
''''''          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
''''''����      :
''''''����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCMW001_SQL.bas���ړ�)
'''''Public Function DBDRV_GetTBCMW001(records() As typ_TBCMW001, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
'''''Dim sql As String       'SQL�S��
'''''Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
'''''Dim rs As OraDynaset    'RecordSet
'''''Dim recCnt As Long      '���R�[�h��
'''''Dim i As Long
'''''
'''''    ''SQL��g�ݗ��Ă�
'''''    sqlBase = "Select CRYNUM, INGOTPOS, TRANCNT, CRYLEN, KRPROCCD, PROCCODE, BLOCKID, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
'''''              " SENDFLAG, SENDDATE "
'''''    sqlBase = sqlBase & "From TBCMW001"
'''''    sql = sqlBase
'''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
'''''        sql = sql & " " & sqlWhere & " " & sqlOrder
'''''    End If
'''''
'''''    ''�f�[�^�𒊏o����
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'''''    If rs Is Nothing Then
'''''        ReDim records(0)
'''''        DBDRV_GetTBCMW001 = FUNCTION_RETURN_FAILURE
'''''        Exit Function
'''''    End If
'''''
'''''    ''���o���ʂ��i�[����
'''''    recCnt = rs.RecordCount
'''''    ReDim records(recCnt)
'''''    For i = 1 To recCnt
'''''        With records(i)
'''''            .Crynum = rs("CRYNUM")           ' �����ԍ�
'''''            .INGOTPOS = rs("INGOTPOS")       ' �C���S�b�g�ʒu
'''''            .TRANCNT = rs("TRANCNT")         ' ������
'''''            .CRYLEN = rs("CRYLEN")           ' ����
'''''            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
'''''            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
'''''            .BLOCKID = rs("BLOCKID")         ' �u���b�NID
'''''            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
'''''            .REGDATE = rs("REGDATE")         ' �o�^���t
'''''            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
'''''            .UPDDATE = rs("UPDDATE")         ' �X�V���t
'''''            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
'''''            .SENDDATE = rs("SENDDATE")       ' ���M���t
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    DBDRV_GetTBCMW001 = FUNCTION_RETURN_SUCCESS
'''''End Function
'''''

'�T�v      :�֘A��ۯ��R�t�R��(TBCMY023)�o�^
'���Ұ��@�@:�ϐ���      ,IO ,�^                 ,����
'      �@�@:KblkData()  ,I  ,typ_BlkMap         ,�֘A��ۯ��ް�
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@  ,�������݂̐���
'����      :
'����      :07/12/21 ooba
Public Function DBDRV_KanrenBlk(KblkData() As typ_BlkMap) As FUNCTION_RETURN

    Dim sql             As String
    Dim i               As Long
    Dim rs              As OraDynaset
    Dim lRecCnt         As Long             'ں��ސ�
    Dim KanrenData()    As typ_TBCMY023     '�֘A��ۯ��R�t�R���ް�
    Dim iTrnCnt         As Integer          '������


    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_KanrenBlk"

    DBDRV_KanrenBlk = FUNCTION_RETURN_FAILURE

    '�����񐔎擾
    sql = "SELECT NVL(MAX(TRANCNT),0) MAXCNT FROM TBCMY023"
    sql = sql & " WHERE CRYNUM = '" & KblkData(1).CRYNUM & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        iTrnCnt = 1
    Else
        iTrnCnt = rs("MAXCNT") + 1          '������(�ő�) + 1
    End If
    rs.Close


    lRecCnt = UBound(KblkData)              '�o�^ں��ސ�
    ReDim KanrenData(lRecCnt)

    '�֘A��ۯ������֘A��ۯ��R�t�R��(TBCMY023)�ɓo�^
    For i = 1 To lRecCnt
        With KanrenData(i)
            .CRYNUM = KblkData(i).CRYNUM        '�����ԍ�
            .TRANCNT = iTrnCnt                  '������
            .BLOCKID = KblkData(i).BLOCKID      '��ۯ�ID
            .PROCCAT = "N"                      '�����敪(N:�V�K)
            .TXID = "TX879I"                    '��ݻ޸���ID
            
            sql = "INSERT INTO TBCMY023"
            sql = sql & " (CRYNUM,"
            sql = sql & " TRANCNT,"
            sql = sql & " BLOCKID,"
            sql = sql & " PROCCAT,"
            sql = sql & " TXID,"
            sql = sql & " REGDATE,"
            sql = sql & " SENDFLAG,"
            sql = sql & " SENDDATE,"
            sql = sql & " PLANTCAT,"
            sql = sql & " SUMITFLAG,"
            sql = sql & " SUMITSND,"
            sql = sql & " SSENDNO) "
            sql = sql & " VALUES"
            sql = sql & " ('" & .CRYNUM & "',"          '�����ԍ�
            sql = sql & .TRANCNT & ","                  '������
            sql = sql & " '" & .BLOCKID & "',"          '��ۯ�ID
            sql = sql & " '" & .PROCCAT & "',"          '�����敪
            sql = sql & " '" & .TXID & "',"             '��ݻ޸���ID
            sql = sql & " SYSDATE,"                     '�o�^���t
            sql = sql & " '5',"                         '���M�׸�(5:WF���M�ΏۊO)
            sql = sql & " NULL, "                       '���M���t
            sql = sql & "  '" & sCmbMukesaki & "', "    '����
            sql = sql & " '0',"                         'SUMIT���M�׸�
            sql = sql & " NULL,"                        'SUMIT���M���t
            sql = sql & " NULL) "                       '���M���A��
        End With
        
        '�o�^���s
        If OraDB.ExecuteSQL(sql) <= 0 Then
            GoTo proc_exit
        End If
    Next i
    
    DBDRV_KanrenBlk = FUNCTION_RETURN_SUCCESS

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

