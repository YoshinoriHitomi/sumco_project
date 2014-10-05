Attribute VB_Name = "s_cmbc035_SQL"
Option Explicit

' �������ύX

' �u���b�N���
Public Type typ_BlkInf2
    BLOCKID As String * 12      ' �u���b�NID
    LENGTH As Integer           ' ����
    REALLEN As Integer          ' ������
    NOWPROC As String * 5       ' ���ݍH��
    DELFLG As String * 1        ' �폜�敪
    TOPBDLN As Integer          ' TOP�s�ǒ���
    TOPBDCS As String * 3       ' TOP�s�Ǘ��R
    TAILBDLN As Integer         ' TAIL�s�ǒ���
    TAILBDCS As String * 3      ' TAIL�s�Ǘ��R
    COF As type_Coefficient     ' �ΐ͌W���v�Z
End Type

'�T�v      :�������ύX�p �u���b�N�h�c���͎��c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sBlockID�@�@�@,I  ,String         �@,�u���b�NID
'�@�@      :pCryInf �@�@�@,O  ,typ_TBCME037   �@,�������
'�@�@      :pHinDsn �@�@�@,O  ,typ_TBCME039   �@,�i�Ԑ݌v
'�@�@      :pHinMng �@�@�@,O  ,typ_TBCME041   �@,�i�ԊǗ�
'      �@�@:pSXLMng �@�@�@,O  ,typ_TBCME042   �@,SXL�Ǘ�
'      �@�@:pWafSmp �@�@�@,O  ,typ_XSDCW   �@   ,�V�T���v���Ǘ��iSXL�j
'�@�@      :pBlkInf �@�@�@,O  ,typ_BlkInf2    �@,�u���b�N���
'�@�@      :pHinSpec�@�@�@,O  ,typ_HinSpec    �@,���i�d�l
'�@�@      :pBlkID  �@�@�@,O  ,String         �@,���o�P�ʃu���b�NID
'      �@�@:dNeraiRes �@�@,O  ,Double         �@,�˂炢�i�Ԃ̔��R����l�iP+�̔��f�p�j
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:sPuptn�@�@    ,O  ,String         �@,���������  2004/12/08 �ǉ�
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :
'����      :2001/07/10 ���� �쐬
Public Function DBDRV_scmzc_fcmkc001i_Disp(ByVal sBlockId As String, pCryInf As typ_TBCME037, _
                                           pHinDsn() As typ_TBCME039, pHinMng() As typ_TBCME041, _
                                           pSXLMng() As typ_TBCME042, pWafSmp() As typ_XSDCW, _
                                           pBlkInf() As typ_BlkInf2, pHinSpec() As typ_HinSpec, _
                                           pBlkID() As String, dNeraiRes As Double, sErrMsg As String, sPuptn As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpStrRslt() As typ_TBCMY009
    Dim rs As OraDynaset
    Dim sql As String
    Dim sDbName As String
    Dim sCryNum As String
    Dim sHin As String
    Dim sBlk As String
    Dim dMenseki As Double
    Dim dTopWght As Double
    Dim dCharge As Double
    Dim dMeas(4) As Double
    Dim bFlag As Boolean
    Dim recCnt As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc035_SQL.bas -- Function DBDRV_scmzc_fcmkc001i_Disp"
    sErrMsg = ""

'����������ݒǉ��Ή�(2004/12/08) kubota
    sPuptn = ""
    sDbName = "XSDC1"
    sql = "select PUPTNC1"
    sql = sql & "  from XSDC1,XSDC2"
    sql = sql & " where CRYNUMC2 = '" & Trim$(sBlockId) & "'"
    sql = sql & "   and XTALC1   = XTALC2"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt <> 0 Then
        sPuptn = rs("PUPTNC1")
    End If
    rs.Close
'����������ݒǉ��Ή�(2004/12/08) kubota

    '' �u���b�N�Ǘ��̎擾
    sDbName = "E040"
    sCryNum = Left(sBlockId, 9) & "000"
    sql = "select INGOTPOS, LENGTH, REALLEN, BLOCKID, "
    sql = sql & "KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, LSTATCLS"
    sql = sql & " from TBCME040 where CRYNUM='" & sCryNum & "'"
    sql = sql & " and INGOTPOS>=0 and LENGTH>0 order by INGOTPOS"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    bFlag = False
    ReDim pBlkInf(recCnt)
    For i = 1 To recCnt
        With pBlkInf(i)
            .COF.TOPSMPLPOS = rs("INGOTPOS")
            .LENGTH = rs("LENGTH")
            .REALLEN = rs("REALLEN")
            .BLOCKID = rs("BLOCKID")
            .NOWPROC = rs("NOWPROC")
            .COF.BOTSMPLPOS = .COF.TOPSMPLPOS + .LENGTH
            .DELFLG = "0"
            .TOPBDLN = 0
            .TOPBDCS = ""
            .TAILBDLN = 0
            .TAILBDCS = ""
            If .BLOCKID = sBlockId Then
                '' �H���`�F�b�N
                If rs("LSTATCLS") <> "W" Then
                    sErrMsg = GetMsgStr("EPRC2")
                    DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
                bFlag = True
            End If
        End With
        rs.MoveNext
    Next i
    rs.Close

    '' �u���b�NID���݃`�F�b�N
    If bFlag = False Then
        sErrMsg = GetMsgStr("EBLK0")
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �������ύX�̎擾
    sDbName = "W002"
    For i = 1 To recCnt
        With pBlkInf(i)
            sql = "select CRYLEN, TOPBDLN, TOPBDCS, TAILBDLN, TAILBDCS from TBCMW002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and INGOTPOS=" & .COF.TOPSMPLPOS
            sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMW002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and INGOTPOS=" & .COF.TOPSMPLPOS & ")"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount > 0 Then
                .REALLEN = rs("CRYLEN")
                .TOPBDLN = rs("TOPBDLN")
                .TOPBDCS = rs("TOPBDCS")
                .TAILBDLN = rs("TAILBDLN")
                .TAILBDCS = rs("TAILBDCS")
            End If
            rs.Close
        End With
    Next i

    '' �������̎擾(s_cmzcTBCME037_SQL.bas ���K�v)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)

    '' �i�Ԑ݌v�̎擾(s_cmzcTBCME039_SQL.bas ���K�v)
    sDbName = "E039"
    '2004.09.08 Y.K �R�t���ύX
'    sql = " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' and LENGTH>0 order by INGOTPOS"
    sql = " where substr(CRYNUM,1,9)='" & Left(sCryNum, 7) & "0" & Mid(sCryNum, 9, 1) & "' and LENGTH>0 order by INGOTPOS"
    If DBDRV_GetTBCME039(pHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinDsn) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �i�ԊǗ��̎擾(s_cmzcTBCME041_SQL.bas ���K�v)
    sDbName = "E041"
    sql = " where CRYNUM='" & sCryNum & "' order by INGOTPOS"
    If DBDRV_GetTBCME041(pHinMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' SXL�Ǘ��̎擾(s_cmzcTBCME042_SQL.bas ���K�v)
    sDbName = "E042"
    If DBDRV_GetTBCME042(pSXLMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pSXLMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' WF�T���v���Ǘ��̎擾(s_cmzcTBCME044_SQL.bas ���K�v)
    sDbName = "E044"
' �V�T���v���Ǘ�(�u���b�N)�ǉ��ɂ��C��  2003/10/06 Takada ===================> START
    sql = " where XTALCW='" & sCryNum & "' and LIVKCW='0' order by INPOSCW, TBKBNCW"
''    sql = " where XTALCW='" & sCryNum & "' order by INPOSCW, TBKBNCW"
' �V�T���v���Ǘ�(�u���b�N)�ǉ��ɂ��C��  2003/10/06 Takada ===================> END
    If DBDRV_GetTBCME044(pWafSmp(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' ���グ�I�����т̎擾
    sDbName = "H004"
    sql = "select (DM1+DM2+DM3)/3.0 as DM, WGHTTOP, CHARGE from TBCMH004 where CRYNUM='" & sCryNum & "'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    dMenseki = AreaOfCircle(rs("DM"))
    dTopWght = rs("WGHTTOP")
    dCharge = rs("CHARGE")
    rs.Close

    '' ������R���т̎擾
    sDbName = "J002"
    recCnt = UBound(pBlkInf)
    For i = 1 To recCnt
        With pBlkInf(i)
            .COF.DUNMENSEKI = dMenseki      ' �f�ʐ�
            .COF.CHARGEWEIGHT = dCharge     ' �`���[�W��
            .COF.TOPWEIGHT = dTopWght       ' �g�b�v�d��

            '' �g�b�v�����R�����l�̎擾
            sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.TOPSMPLPOS & " and SMPKBN='T'"
            sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.TOPSMPLPOS & " and SMPKBN='T')"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount > 0 Then
                dMeas(0) = rs("MEAS1")
                dMeas(1) = rs("MEAS2")
                dMeas(2) = rs("MEAS3")
                dMeas(3) = rs("MEAS4")
                dMeas(4) = rs("MEAS5")
                .COF.TOPRES = JudgCenter(dMeas())
            Else
                .COF.TOPRES = -9999
            End If
            rs.Close

            '' �{�g�������R�����l�̎擾
            sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='B'"
            sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='B')"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
                sql = sql & " where CRYNUM='" & sCryNum & "'"
                sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='T'"
                sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
                sql = sql & " where CRYNUM='" & sCryNum & "'"
                sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='T')"
                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            End If
            If rs.RecordCount > 0 Then
                dMeas(0) = rs("MEAS1")
                dMeas(1) = rs("MEAS2")
                dMeas(2) = rs("MEAS3")
                dMeas(3) = rs("MEAS4")
                dMeas(4) = rs("MEAS5")
                .COF.BOTRES = JudgCenter(dMeas())
            Else
                .COF.BOTRES = -9999
            End If
            rs.Close
        End With
    Next i

    '' �u���b�N�V�K���̎擾
    sDbName = "Y001"
    sql = "select SBLOCKID from TBCMY001 where BLOCKID='" & sBlockId & "'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    sBlk = rs("SBLOCKID")
    rs.Close

    sql = "select BLOCKID from TBCMY001"
    sql = sql & " where SBLOCKID='" & sBlk & "'"
    sql = sql & " order by SBLOCKID, BLOCKORDER"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt <= 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    ReDim pBlkID(recCnt)
    For i = 1 To recCnt
        pBlkID(i) = rs("BLOCKID")
        rs.MoveNext
    Next i
    rs.Close
    
    '' ���i�d�l�̎擾
    sDbName = "VE004"
    recCnt = UBound(pHinMng)
    ReDim pHinSpec(recCnt)
    k = 0
    For i = 1 To recCnt
        With pHinMng(i)
            sHin = RTrim$(.hinban)
            If sHin <> "" And sHin <> "G" And sHin <> "Z" Then
                For j = 1 To k
                    If pHinSpec(j).HIN.hinban = .hinban Then
                        pHinSpec(j).LENGTH = pHinSpec(j).LENGTH + .LENGTH
                        Exit For
                    End If
                Next j
                If j > k Then
                    k = k + 1
                    pHinSpec(k).INGOTPOS = .INGOTPOS
                    pHinSpec(k).HIN.hinban = .hinban
                    pHinSpec(k).HIN.mnorevno = .REVNUM
                    pHinSpec(k).HIN.factory = .factory
                    pHinSpec(k).HIN.opecond = .opecond
                    pHinSpec(k).LENGTH = .LENGTH
                    
                    ''�c���_�f�d�l�`�F�b�N�@03/12/09 ooba START ==============================>
                    iChkAoi = ChkAoiSiyou(pHinSpec(k).HIN)
                    If iChkAoi < 0 Then
                        sErrMsg = "�c���_�f(AOi)�d�l�G���["
                        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                    ''�c���_�f�d�l�`�F�b�N�@03/12/09 ooba END ================================>
                    
                    If DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec(k)) = FUNCTION_RETURN_FAILURE Then
                        sErrMsg = GetMsgStr("EGET") & sDbName
                        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
            End If
        End With
    Next i
    ReDim Preserve pHinSpec(k)

    '' �˂炢�i�Ԃ̔��R����l���擾
    sql = "select HSXRMAX"
    sql = sql & " from TBCME037 E37, TBCME018 E18"
    sql = sql & " where (E37.CRYNUM='" & Left$(sBlockId, 9) & "000')"
    sql = sql & " and (E37.RPHINBAN=E18.HINBAN) and (E37.RPREVNUM=E18.MNOREVNO)"
    sql = sql & " and (E37.RPFACT=E18.FACTORY) and (E37.RPOPCOND=E18.OPECOND)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        dNeraiRes = rs("HSXRMAX")
    Else
        dNeraiRes = 0#      '�����܂ł͂��Ȃ��͂�
    End If
    rs.Close

    DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_SUCCESS

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
    DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�������ύX�p ���s���c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sStaffID�@�@�@,I  ,String         �@,�Ј�ID
'�@�@      :pBlkInf �@�@�@,I  ,typ_BlkInf2    �@,�u���b�N���
'�@�@      :pHinMng �@�@�@,I  ,typ_TBCME041   �@,�i�ԊǗ�
'      �@�@:pSXLMng �@�@�@,I  ,typ_TBCME042   �@,SXL�Ǘ�
'      �@�@:pSXLOld �@�@�@,I  ,typ_TBCME042   �@,�ύX�OSXL�Ǘ�
'      �@�@:pWafSmp �@�@�@,I  ,typ_XSDCW   �@   ,�V�T���v���Ǘ��iSXL�j
'      �@�@:pWafOld �@�@�@,I  ,typ_XSDCW   �@   ,�ύX�O�V�T���v���Ǘ��iSXL�j
'      �@�@:pTrnScr �@�@�@,I  ,typ_TBCMW006   �@,�U�֔p������
'      �@�@:pMesInd �@�@�@,I  ,typ_TBCMY003   �@,����]�����@�w��
'      �@�@:pSXLDcd �@�@�@,I  ,typ_TBCMY007   �@,SXL�m��w��
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2001/07/10 ���� �V�K�쐬
'      �@�@:2001/07/11 ���{ �ύX
'      �@�@:2003/04/12 HITEC)��c�FTBCMY003,TBCMY007�̑��M�t���O��'0'=>'3'�ɕύX

Public Function DBDRV_scmzc_fcmkc001i_Exec(SSTAFFID As String, pBlkInf() As typ_BlkInf2, _
                                           pHinMng() As typ_TBCME041, pSXLMng() As typ_TBCME042, _
                                           pSXLOld() As typ_TBCME042, pWafSmp() As typ_XSDCW, _
                                           pWafOld() As typ_XSDCW, pTrnScr() As typ_TBCMW006, _
                                           pMesInd() As typ_TBCMY003, pSXLDcd() As typ_TBCMY007, sErrMsg As String) As FUNCTION_RETURN

    Dim sql As String
    Dim sDbName As String
    Dim sCryNum As String
    Dim sAllScrap As String
    Dim recCnt As Long
    Dim i As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc035_SQL.bas -- Function DBDRV_scmzc_fcmkc001i_Exec"
    sErrMsg = ""

    '' WriteDBLog " ", "Start"

    '' SXL�Ǘ��̑}���^�X�V(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    sDbName = "E042"
    If DBDRV_SXL_UpdIns(pSXLOld(), pSXLMng()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' WF�T���v���Ǘ��̑}���^�X�V(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    sDbName = "E044"
    If DBDRV_WfSmp_UpdIns(pWafOld(), pWafSmp()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    recCnt = UBound(pBlkInf)
    For i = 1 To recCnt
        With pBlkInf(i)
            sCryNum = Left(.BLOCKID, 9) & "000"
            '' �u���b�N�Ǘ��̍X�V
            sDbName = "E040"
            sql = "update TBCME040 set "
            sql = sql & "REALLEN='" & .REALLEN & "', "
            sql = sql & "UPDDATE=sysdate, "
            sql = sql & "SENDFLAG='0'"
            sql = sql & " where CRYNUM='" & sCryNum & "' and INGOTPOS=" & .COF.TOPSMPLPOS
            '' WriteDBLog sql, sDbName
            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If

            '' �������ύX���т̑}��
            sDbName = "W002"
            sql = "insert into TBCMW002 "
            sql = sql & "(CRYNUM, INGOTPOS, TRANCNT, CRYLEN, KRPROCCD, "
            sql = sql & "PROCCODE, BLOCKID, DELFLG, TOPBDLN, TOPBDCS, "
            sql = sql & "TAILBDLN, TAILBDCS, TSTAFFID, REGDATE, KSTAFFID, "
            sql = sql & "UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE)"
            sql = sql & " select '"
            sql = sql & sCryNum & "', "
            sql = sql & .COF.TOPSMPLPOS & ", "
            sql = sql & "nvl(max(TRANCNT),0)+1, "
            sql = sql & .REALLEN & ", '"
            sql = sql & MGPRCD_KESSYOU_SIYOUJOUHOU_HENKOU & "', '"
            sql = sql & PROCD_KESSYOU_SIYOUJOUHOU_HENKOU & "', '"
            sql = sql & .BLOCKID & "', '"
            sql = sql & .DELFLG & "', "
            sql = sql & .TOPBDLN & ", '"
            sql = sql & .TOPBDCS & "', "
            sql = sql & .TAILBDLN & ", '"
            sql = sql & .TAILBDCS & "', '"
            sql = sql & SSTAFFID & "', "
            sql = sql & "sysdate, '"
            sql = sql & SSTAFFID & "', "
            sql = sql & "sysdate, "
            sql = sql & "'0', "
            sql = sql & "'0', "
            sql = sql & "sysdate"
            sql = sql & " from TBCMW002"
            sql = sql & " where CRYNUM='" & sCryNum & "' and INGOTPOS=" & .COF.TOPSMPLPOS
            '' WriteDBLog sql, sDbName
            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If

            '' �������̑}��
            sDbName = "Y012"
            sql = "delete TBCMY012 where LOTID='" & .BLOCKID & "' and BLOCKSEQ<0"
            Call OraDB.ExecuteSQL(sql)
            If .TOPBDLN >= .REALLEN Or .TAILBDLN >= .REALLEN Or _
               .TOPBDLN + .TAILBDLN >= .REALLEN Then
                sql = "insert into TBCMY012 "
                sql = sql & "(LOTID, BLOCKSEQ, REJPCS, TOP_POS, REJCAT, "
                sql = sql & "REJDTTM, REJPROC, ALLSCRAP, LENFROM, LENTO, "
                sql = sql & "TXID, CHKFLG, REGDATE, SENDFLAG, SENDDATE)"
                sql = sql & " values ('"
                sql = sql & .BLOCKID & "', "
                sql = sql & "-1, "
                sql = sql & "-1, "
                sql = sql & "0, "
                sql = sql & "'A', "
                sql = sql & "sysdate, '"
                sql = sql & PROCD_KESSYOU_SIYOUJOUHOU_HENKOU & "', "
                sql = sql & "'Y', "
                sql = sql & "0, "
                sql = sql & .REALLEN & ", "
                sql = sql & "'      ', "
                sql = sql & "'1', "         ' 1129 �`�F�b�N�t���O�̓`�F�b�N�ς�
                sql = sql & "sysdate, "
                sql = sql & "'0', "
                sql = sql & "sysdate)"
                '' WriteDBLog sql, sDbName
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                    DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            Else
                If .TOPBDLN > 0 Then
                    sql = "insert into TBCMY012 "
                    sql = sql & "(LOTID, BLOCKSEQ, REJPCS, TOP_POS, REJCAT, "
                    sql = sql & "REJDTTM, REJPROC, ALLSCRAP, LENFROM, LENTO, "
                    sql = sql & "TXID, CHKFLG, REGDATE, SENDFLAG, SENDDATE)"
                    sql = sql & " values ('"
                    sql = sql & .BLOCKID & "', "
                    sql = sql & "-1, "
                    sql = sql & "1, "
                    sql = sql & "0, "
                    sql = sql & "'A', "
                    sql = sql & "sysdate, '"
                    sql = sql & PROCD_KESSYOU_SIYOUJOUHOU_HENKOU & "', "
                    sql = sql & "'N', "
                    sql = sql & "0, "
                    sql = sql & .TOPBDLN & ", "
                    sql = sql & "'      ', "
                    sql = sql & "'1', "         ' 1129 �`�F�b�N�t���O�̓`�F�b�N�ς�
                    sql = sql & "sysdate, "
                    sql = sql & "'0', "
                    sql = sql & "sysdate)"
                    '' WriteDBLog sql, sDbName
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
                If .TAILBDLN > 0 Then
                    sql = "insert into TBCMY012 "
                    sql = sql & "(LOTID, BLOCKSEQ, REJPCS, TOP_POS, REJCAT, "
                    sql = sql & "REJDTTM, REJPROC, ALLSCRAP, LENFROM, LENTO, "
                    sql = sql & "TXID, CHKFLG, REGDATE, SENDFLAG, SENDDATE)"
                    sql = sql & " values ('"
                    sql = sql & .BLOCKID & "', "
                    sql = sql & "-2, "
                    sql = sql & "1, "
                    sql = sql & .REALLEN - .TAILBDLN & ", "
                    sql = sql & "'A', "
                    sql = sql & "sysdate, '"
                    sql = sql & PROCD_KESSYOU_SIYOUJOUHOU_HENKOU & "', "
                    sql = sql & "'N', "
                    sql = sql & .REALLEN - .TAILBDLN & ", "
                    sql = sql & .REALLEN & ", "
                    sql = sql & "'      ', "
                    sql = sql & "'1', "         ' 1129 �`�F�b�N�t���O�̓`�F�b�N�ς�
                    sql = sql & "sysdate, "
                    sql = sql & "'0', "
                    sql = sql & "sysdate)"
                    '' WriteDBLog sql, sDbName
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
            End If
        End With
    Next i

    '' �U�֔p�����т̑}��(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    sDbName = "W006"
    recCnt = UBound(pTrnScr)
    For i = 1 To recCnt
        If DBDRV_Furikae_Ins(pTrnScr(i)) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' ����]�����@�w���̑}��(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    sDbName = "Y003"
    If DBDRV_SokuSizi_Ins(pMesInd()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �u���b�N�ύX���̑}��
    If DBDRV_BlkChg_Ins(pBlkInf(), pHinMng(), sDbName) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' SXL�m��w���̑}��
    sDbName = "Y007"
    recCnt = UBound(pSXLDcd)
    For i = 1 To recCnt
        With pSXLDcd(i)
            sql = "insert into TBCMY007 "
            sql = sql & "(SXL_ID, SAMPLE_FROM, SAMPLE_TO, BLOCKID, HINBAN, KUBUN, "
'            sql = sql & "TXID, REGDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE)"
            '���R�ް��o�^�ǉ��@04/04/09 ooba START =======================================>
            sql = sql & "TXID, REGDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE, "
            sql = sql & "MESDATA1TOP, "      ' ����l�P(Top)  center
            sql = sql & "MESDATA2TOP, "      ' ����l�Q(Top)  R/2
            sql = sql & "MESDATA3TOP, "      ' ����l�R(Top)  Inside 10mm
            sql = sql & "MESDATA4TOP, "      ' ����l�S(Top)  Inside   6mm
            sql = sql & "MESDATA5TOP, "      ' ����l�T(Top)  Inside   3mm
            sql = sql & "MESDATA1BOT, "      ' ����l�P(Tail)  center
            sql = sql & "MESDATA2BOT, "      ' ����l�Q(Tail)  R/2
            sql = sql & "MESDATA3BOT, "      ' ����l�R(Tail)  Inside 10mm
            sql = sql & "MESDATA4BOT, "      ' ����l�S(Tail)  Inside   6mm
            sql = sql & "MESDATA5BOT )"      ' ����l�T(Tail)  Inside   3mm
            '���R�ް��o�^�ǉ��@04/04/09 ooba END =========================================>
            sql = sql & " values ('"
            sql = sql & .SXL_ID & "', '"        ' SXL-ID
            sql = sql & .SAMPLE_FROM & "', '"   ' �T���v��ID (From)
            sql = sql & .SAMPLE_TO & "', '"     ' �T���v��ID (To)
            sql = sql & .BLOCKID & "', '"       ' �u���b�N�h�c
            sql = sql & .hinban & "', "         ' �m��i��
            sql = sql & "'S ', "                ' �敪�R�[�h
            sql = sql & "'TX853I', "            ' �g�����U�N�V����ID
            sql = sql & "sysdate, "             ' �o�^���t
            sql = sql & "'0', "                 ' SUMMIT���M�t���O
            
' vvvvv 2003.04.12 ALT BY HITEC)��c�F���M�t���O'0'=>'3'�ɕύX
'''''            sql = sql & "'0', "                 ' ���M�t���O
            sql = sql & "'3', "                 ' ���M�t���O
' ^^^^^ 2003.04.12 ALT BY HITEC)��c  END
'            sql = sql & "sysdate)"              ' ���M���t
            '���R�ް��o�^�ǉ��@04/04/09 ooba START =======================================>
            sql = sql & "sysdate, "              ' ���M���t
            sql = sql & " '" & .MESDATA1TOP & "', "      ' ����l�P(Top)  center
            sql = sql & " '" & .MESDATA2TOP & "', "      ' ����l�Q(Top)  R/2
            sql = sql & " '" & .MESDATA3TOP & "', "      ' ����l�R(Top)  Inside 10mm
            sql = sql & " '" & .MESDATA4TOP & "', "      ' ����l�S(Top)  Inside   6mm
            sql = sql & " '" & .MESDATA5TOP & "', "      ' ����l�T(Top)  Inside   3mm
            sql = sql & " '" & .MESDATA1BOT & "', "      ' ����l�P(Tail)  center
            sql = sql & " '" & .MESDATA2BOT & "', "      ' ����l�Q(Tail)  R/2
            sql = sql & " '" & .MESDATA3BOT & "', "      ' ����l�R(Tail)  Inside 10mm
            sql = sql & " '" & .MESDATA4BOT & "', "      ' ����l�S(Tail)  Inside   6mm
            sql = sql & " '" & .MESDATA5BOT & "' ) "     ' ����l�T(Tail)  Inside   3mm
            '���R�ް��o�^�ǉ��@04/04/09 ooba END =========================================>
            '' WriteDBLog sql, sDbName
            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End With
    Next i

    DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_SUCCESS

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
    DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�����֐��F�u���b�N�ύX���̍쐬
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'�@�@      :pBlkInf�@�@�@,I  ,typ_BlkInf2    �@,�u���b�N���
'�@�@      :pHinMng�@�@�@,I  ,typ_TBCME041   �@,�i�ԊǗ�
'      �@�@:sDBName�@�@�@,O  ,String         �@,DB����
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2001/07/25  �쐬 ���{
Private Function DBDRV_BlkChg_Ins(pBlkInf() As typ_BlkInf2, pHinMng() As typ_TBCME041, sDbName As String) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim CRYSTALMEN As String
    Dim SEED As Integer
    Dim TRANCNT As Long
    Dim m As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc035_SQL.bas -- Function DBDRV_BlkChg_Ins"

    m = UBound(pBlkInf)
    n = UBound(pHinMng)
    For i = 1 To m
        With pBlkInf(i)
            '' �i�Ԃ̌���
            For j = 1 To n
                If .COF.TOPSMPLPOS >= pHinMng(j).INGOTPOS And _
                   .COF.TOPSMPLPOS < pHinMng(j).INGOTPOS + pHinMng(j).LENGTH Then
                    Exit For
                End If
            Next j
            If RTrim$(pHinMng(j).hinban) <> "Z" Then
                '' �V�[�h�X���̎擾(s_cmzcDBdriverCOM_SQL.bas ���K�v)
                sDbName = "H004"
                If DBDRV_getSEED(Left(pBlkInf(i).BLOCKID, 9) & "000", SEED) = FUNCTION_RETURN_FAILURE Then
                    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If

                '' �����ʂ̎擾
                sDbName = "E022"
                sql = "select HWFCDIR from TBCME022"
                sql = sql & " where HINBAN='" & pHinMng(j).hinban & "'"
                sql = sql & " and MNOREVNO=" & pHinMng(j).REVNUM
                sql = sql & " and FACTORY='" & pHinMng(j).factory & "'"
                sql = sql & " and OPECOND='" & pHinMng(j).opecond & "'"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount <= 0 Then
                    rs.Close
                    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
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
                    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If

                '' �����񐔍ő�l�̎擾
                sDbName = "Y005"
                sql = "select nvl(max(TRANCNT),0)+1 as M"
                sql = sql & " from TBCMY005"
                sql = sql & " where BLOCKID='" & .BLOCKID & "'"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
                TRANCNT = rs("M")
                rs.Close

                '' �u���b�N�ύX���̑}��
                sql = "insert into TBCMY005 ("
                sql = sql & "BLOCKID, "           ' �u���b�NID
                sql = sql & "TRANCNT, "           ' ������
                sql = sql & "DELFLG, "            ' �폜�w��
                sql = sql & "BLOCKLEN, "          ' �u���b�N�̒���
                sql = sql & "MAINHINBAN, "        ' ��\�i��
                sql = sql & "PNTYPE, "            ' �^�C�v
                sql = sql & "ROUP, "              ' ���R����l
                sql = sql & "ROLOW, "             ' ���R�����l
                sql = sql & "OIUP, "              ' �_�f�Z�x����l
                sql = sql & "OILOW, "             ' �_�f�Z�x�����l
                sql = sql & "TANMEN, "            ' �[�ʊp�x
                sql = sql & "WARPRANK, "          ' ���[�v�����N
                sql = sql & "CRYSTALMEN, "        ' ������
                sql = sql & "SLPCEN, "            ' �X���S
                sql = sql & "SLPLOW, "            ' �X����
                sql = sql & "SLPUP, "             ' �X���
                sql = sql & "INSPMETH, "          ' �������@
                sql = sql & "INSPFREQ, "          ' �����p�x
                sql = sql & "SLPDRC, "            ' �X����
                sql = sql & "SLPDRCAPP, "         ' �X���ʎw��
                sql = sql & "SLPHEIDRC, "         ' �X�c����
                sql = sql & "SLPHEICEN, "         ' �X�c���S
                sql = sql & "SLPHEILOW, "         ' �X�c����
                sql = sql & "SLPHEIUP, "          ' �X�c���
                sql = sql & "SLPWIDDRC, "         ' �X������
                sql = sql & "SLPWIDCEN, "         ' �X�����S
                sql = sql & "SLPWIDLOW, "         ' �X������
                sql = sql & "SLPWIDUP, "          ' �X�����
                sql = sql & "SEED, "              ' ���㎞�g�p�����V�|�h�X��
                sql = sql & "TXID, "              ' �g�����U�N�V����ID
                sql = sql & "REGDATE, "           ' �o�^���t
                sql = sql & "SENDFLAG, "          ' ���M�t���O
                sql = sql & "SENDDATE)"           ' ���M���t
                sql = sql & " select '"
                sql = sql & .BLOCKID & "', "                            ' �u���b�NID
                sql = sql & TRANCNT & ", '"                             ' ������
                sql = sql & .DELFLG & "', '"                            ' �폜�w��
                sql = sql & .REALLEN & "', '"                           ' �u���b�N�̒���
                                                                        ' ��\�i��
                sql = sql & pHinMng(j).hinban & Format(pHinMng(j).REVNUM, "00") & "', "
                sql = sql & "E021HWFTYPE, "                             ' �^�C�v
                sql = sql & "case when E021HWFRMAX>=99999.9 then '99999.9'"
                sql = sql & " when E021HWFRMAX>=9999.995 then to_char(round(E021HWFRMAX,2),'fm99990.0')"
                sql = sql & " when E021HWFRMAX>=999.9995 then to_char(round(E021HWFRMAX,3),'fm9990.00')"
                sql = sql & " when E021HWFRMAX>=99.99995 then to_char(round(E021HWFRMAX,4),'fm990.000')"
                sql = sql & " when E021HWFRMAX>=10.00000 then to_char(round(E021HWFRMAX,5),'fm90.0000')"
                sql = sql & " when E021HWFRMAX>=0.0 then to_char(E021HWFRMAX,'fm0.00000')"
                sql = sql & " else '-1.0000'"
                sql = sql & "end as RMAX,"                              ' ���R����l
                sql = sql & "case when E021HWFRMIN>=99999.9 then '99999.9'"
                sql = sql & " when E021HWFRMIN>=9999.995 then to_char(round(E021HWFRMIN,2),'fm99990.0')"
                sql = sql & " when E021HWFRMIN>=999.9995 then to_char(round(E021HWFRMIN,3),'fm9990.00')"
                sql = sql & " when E021HWFRMIN>=99.99995 then to_char(round(E021HWFRMIN,4),'fm990.000')"
                sql = sql & " when E021HWFRMIN>=10.00000 then to_char(round(E021HWFRMIN,5),'fm90.0000')"
                sql = sql & " when E021HWFRMIN>=0.0 then to_char(E021HWFRMIN,'fm0.00000')"
                sql = sql & " else '-1.0000'"
                sql = sql & "end as RMIN,"                              ' ���R�����l
                sql = sql & "to_char(abs(E025HWFONMAX),'fm90.00'), "    ' �_�f�Z�x����l
                sql = sql & "to_char(abs(E025HWFONMIN),'fm90.00'), "    ' �_�f�Z�x�����l
                sql = sql & "'0', "                                     ' �[�ʊp�x
                sql = sql & "E027HWFWARPR, '"                           ' ���[�v�����N
                sql = sql & CRYSTALMEN & "', "                          ' ������
                sql = sql & "to_char(abs(E022HWFCSCEN),'fm0.00'), "     ' �X���S
                sql = sql & "to_char(E022HWFCSMIN,'fm0.00'), "          ' �X����
                sql = sql & "to_char(E022HWFCSMAX,'fm0.00'), "          ' �X���
                sql = sql & "E022HWFCKWAY, "                            ' �������@
                                                                        ' �����p�x�i���A���A�ہA�E�̏��ő����j
                sql = sql & "E022HWFCKHNM || E022HWFCKHNN || E022HWFCKHNH || E022HWFCKHNU, "
                sql = sql & "E022HWFCSDIR, "                            ' �X����
                sql = sql & "E022HWFCSDIS, "                            ' �X���ʎw��
                sql = sql & "E022HWFCTDIR, "                            ' �X�c����
'''                sql = sql & "to_char(E022HWFCTCEN,'fm0.00'), "          ' �X�c���S
'''                sql = sql & "to_char(E022HWFCTMIN,'fm0.00'), "          ' �X�c����
'''                sql = sql & "to_char(E022HWFCTMAX,'fm0.00'),"           ' �X�c���
                sql = sql & "to_char(nvl(E022HWFCTCEN,0),'fm0.00'), "   ' �X�c���S      '05/03/29 ooba NULL�Ή�
                sql = sql & "to_char(nvl(E022HWFCTMIN,0),'fm0.00'), "   ' �X�c����      '05/03/29 ooba NULL�Ή�
                sql = sql & "to_char(nvl(E022HWFCTMAX,0),'fm0.00'),"    ' �X�c���      '05/03/29 ooba NULL�Ή�
                sql = sql & "E022HWFCYDIR, "                            ' �X������
'''                sql = sql & "to_char(E022HWFCYCEN,'fm0.00'), "          ' �X�����S
'''                sql = sql & "to_char(E022HWFCYMIN,'fm0.00'), "          ' �X������
'''                sql = sql & "to_char(E022HWFCYMAX,'fm0.00'), '"         ' �X�����
                sql = sql & "to_char(nvl(E022HWFCYCEN,0),'fm0.00'), "   ' �X�����S      '05/03/29 ooba NULL�Ή�
                sql = sql & "to_char(nvl(E022HWFCYMIN,0),'fm0.00'), "   ' �X������      '05/03/29 ooba NULL�Ή�
                sql = sql & "to_char(nvl(E022HWFCYMAX,0),'fm0.00'), '"  ' �X�����      '05/03/29 ooba NULL�Ή�
                sql = sql & SEED & "', "                                ' ���㎞�g�p�����V�|�h�X��
                sql = sql & "'TX852I', "                                ' �g�����U�N�V����ID
                sql = sql & "sysdate, "                                 ' �o�^���t
                sql = sql & "'0', "                                     ' ���M�t���O
                sql = sql & "sysdate "                                  ' ���M���t
                sql = sql & " from VECME001"
                sql = sql & " where E018HINBAN='" & pHinMng(j).hinban & "'"
                sql = sql & " and E018MNOREVNO=" & pHinMng(j).REVNUM
                sql = sql & " and E018FACTORY='" & pHinMng(j).factory & "'"
                sql = sql & " and E018OPECOND='" & pHinMng(j).opecond & "'"
                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

    sDbName = ""
    DBDRV_BlkChg_Ins = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�����֐��F������т̃`�F�b�N
'���Ұ��@�@:�ϐ���      ,IO ,�^       ,����
'�@�@      :blkID �@�@�@,I  ,String �@,�u���b�NID
'      �@�@:�߂�l      ,O  ,Boolean�@,�o�^�̗L��
'����      :
'����      :2001/08/30  �쐬 �쑺
Public Function wasUkeire(ByVal blkID$) As Boolean

    Dim rs As OraDynaset
    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc035_SQL.bas -- Function wasUkeire"

    '' �������o�u���b�N�̂����ꂩ���������(TBCMY009)�ɓo�^����Ă��邩���`�F�b�N����
    wasUkeire = False
    sql = "select Y009.LOTID "
    sql = sql & "from TBCMY009 Y009, TBCMY001 Y001 "
    sql = sql & "Where Y009.LOTID = Y001.BLOCKID "
    sql = sql & "and Y001.SBLOCKID=(select SBLOCKID from TBCMY001 where BLOCKID='" & blkID & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        wasUkeire = True
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

'(2002/07 s_cmzcF_cmkc001g_SQL.bas���R�s�[)
'�T�v      :�����w���p ���i�d�l��p�c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:pHinSpec�@�@�@,IO ,typ_HinSpec    �@,���i�d�l
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec As typ_HinSpec) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim sOT1    As String   '03/05/24 �㓡
    Dim sOT2    As String
    Dim sMAI1    As String   '04/06/25
    Dim sMAI2    As String
    Dim rtn     As FUNCTION_RETURN

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001g_SQL.bas -- Function DBDRV_scmzc_fcmkc001g_GetSpec"

    '' ���i�d�l�̎擾
    With pHinSpec
        sql = "select "
        sql = sql & "E021HWFRMIN, E021HWFRMAX, E021HWFRHWYS, "
        sql = sql & "E024HWFMKHWS, E025HWFONHWS, E025HWFOS1HS, E025HWFOS2HS, E025HWFOS3HS, "
        sql = sql & "E026HWFDSOHS, E028HWFSPVHS, E028HWFDLHWS, E029HWFOF1HS, E029HWFOF2HS, "
        sql = sql & "E029HWFOF3HS, E029HWFOF4HS, E029HWFBM1HS, E029HWFBM2HS, E029HWFBM3HS"
        sql = sql & " from VECME004"
        sql = sql & " where E018HINBAN='" & .HIN.hinban & "'"
        sql = sql & " and E018MNOREVNO=" & .HIN.mnorevno
        sql = sql & " and E018FACTORY='" & .HIN.factory & "'"
        sql = sql & " and E018OPECOND='" & .HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

'        .HWFRMIN = rs("E021HWFRMIN")
'        .HWFRMAX = rs("E021HWFRMAX")
        .HWFRMIN = fncNullCheck(rs("E021HWFRMIN"))      'Null�Ή� 2003/12/10
        .HWFRMAX = fncNullCheck(rs("E021HWFRMAX"))      'Null�Ή� 2003/12/10
        .HWFRHWYS = rs("E021HWFRHWYS")
        .HWFMKHWS = rs("E024HWFMKHWS")
        .HWFONHWS = rs("E025HWFONHWS")
        .HWFOS1HS = rs("E025HWFOS1HS")
        .HWFOS2HS = rs("E025HWFOS2HS")
        .HWFOS3HS = rs("E025HWFOS3HS")
        .HWFDSOHS = rs("E026HWFDSOHS")
        .HWFSPVHS = rs("E028HWFSPVHS")
        .HWFDLHWS = rs("E028HWFDLHWS")
        .HWFOF1HS = rs("E029HWFOF1HS")
        .HWFOF2HS = rs("E029HWFOF2HS")
        .HWFOF3HS = rs("E029HWFOF3HS")
        .HWFOF4HS = rs("E029HWFOF4HS")
        .HWFBM1HS = rs("E029HWFBM1HS")
        .HWFBM2HS = rs("E029HWFBM2HS")
        .HWFBM3HS = rs("E029HWFBM3HS")
        rtn = scmzc_getE036(pHinSpec.HIN, sOT1, sOT2, sMAI1, sMAI2)   '03/05/24
        If rtn = FUNCTION_RETURN_FAILURE Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        .HWFOTHER1 = sOT1 '### 03/05/24
        .HWFOTHER2 = sOT2
        .HWFOTHER1MAI = sMAI1   '04/06/25
        .HWFOTHER2MAI = sMAI2   '04/06/25
        
        rs.Close
        
        ''�c���_�f�d�l�擾�@03/12/09 ooba START ==============================>
        sql = "select HWFZOHWS from TBCME025 "
        sql = sql & "where HINBAN  ='" & .HIN.hinban & "' "
        sql = sql & "and MNOREVNO= " & .HIN.mnorevno & " "
        sql = sql & "and FACTORY ='" & .HIN.factory & "' "
        sql = sql & "and OPECOND ='" & .HIN.opecond & "' "
        
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        
        If IsNull(rs("HWFZOHWS")) = False Then .HWFZOHWS = rs("HWFZOHWS") '�iWF�c���_�f�ۏؕ��@_��
        rs.Close
        ''�c���_�f�d�l�擾�@03/12/09 ooba END ================================>
        
        '' GD�d�l�擾�@05/01/25 ooba START ==================================>
        sql = "select "
        sql = sql & "HWFDENHS, "        '�iWFDen�ۏؕ��@_��
        sql = sql & "HWFLDLHS, "        '�iWFL/DL�ۏؕ��@_��
        sql = sql & "HWFDVDHS "         '�iWFDVD2�ۏؕ��@_��
        sql = sql & "from TBCME026 "
        sql = sql & "where HINBAN = '" & .HIN.hinban & "' "
        sql = sql & "and MNOREVNO = " & .HIN.mnorevno & " "
        sql = sql & "and FACTORY = '" & .HIN.factory & "' "
        sql = sql & "and OPECOND = '" & .HIN.opecond & "' "
        
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        
        If IsNull(rs("HWFDENHS")) = False Then .HWFDENHS = rs("HWFDENHS")   '�iWFDen�ۏؕ��@_��
        If IsNull(rs("HWFLDLHS")) = False Then .HWFLDLHS = rs("HWFLDLHS")   '�iWFL/DL�ۏؕ��@_��
        If IsNull(rs("HWFDVDHS")) = False Then .HWFDVDHS = rs("HWFDVDHS")   '�iWFDVD2�ۏؕ��@_��
        
        rs.Close
        '' GD�d�l�擾�@05/01/25 ooba END ====================================>
        
    End With

    DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

' vvvvv 2003.04.11 ADD BY HITEC)��c�Fcmbc036_SQL�����ɍ쐬
'��{�����p�����[�^�쐬
'�����FfrmFormID=������ʂ̔���i2:�������ύX�j
Public Function MakeParameter(ByVal strCryNum As String) As FUNCTION_RETURN

    Dim lng     As Long
    Dim dat     As Variant
    Dim lRowCnt As Long
    Dim rsMain      As OraDynaset
    Dim sql     As String
    Dim intCnt  As Integer
    Dim errTbl  As String
    Dim sErrMsg As String
    Dim lngBeginIngotpos    As Long
    Dim lngEndIngotpos      As Long
    Dim strIngotpos As String
    Dim varIngotpos As Variant
    
    With f_cmbc035_1.sprExamine
        .GetText 3, 1, varIngotpos
        lngBeginIngotpos = CInt(Trim(varIngotpos))
        .GetText 3, .MaxRows, varIngotpos
        lngEndIngotpos = CInt(Trim(varIngotpos))
    End With
    
    '�\���̍쐬
    If cmbc035_1_CreateTable(strCryNum, lngBeginIngotpos, lngEndIngotpos, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        MakeParameter = FUNCTION_RETURN_FAILURE
        f_cmbc035_1.lblMsg.Caption = sErrMsg
        Exit Function
    End If
    MakeParameter = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

End Function

Public Function cmbc035_1_CreateTable(ByVal strCryNum As String, ByVal lngBeginIngotpos As Long, ByVal lngEndIngotpos As Long, ByRef strErrMsg As String) As FUNCTION_RETURN

    Dim rs  As OraDynaset
    Dim errTbl  As String
    Dim strBlockID()  As String
    Dim strDBName   As String
    Dim bNoData     As Boolean
    Dim intLoopCnt  As Integer
    Dim sql     As String
    
    bNoData = False

    '�u���b�N�Ǘ�����u���b�N�h�c���擾
    sql = "SELECT * from TBCME040 "
    sql = sql & " WHERE CRYNUM='" & strCryNum & "'"
    sql = sql & "   AND INGOTPOS>=" & lngBeginIngotpos & " AND (INGOTPOS + LENGTH) <=" & lngEndIngotpos
    Debug.Print sql

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc035_1_CreateTable = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    '�u���b�NID���擾
    giInpos = 9000
    intLoopCnt = 0
    Do While Not rs.EOF
        ReDim Preserve strBlockID(intLoopCnt) As String
        If IsNull(rs("BLOCKID")) = True Then
            strBlockID(intLoopCnt) = ""
        Else
            strBlockID(intLoopCnt) = rs("BLOCKID")            '�u���b�NID
        End If
        
        '��{���\����
        With Kihon
            .STAFFID = Trim(f_cmbc035_1.txtStaffID.Text)
''''            .NEWPROC = PROCD_WFC_SOUGOUHANTEI
            .NEWPROC = "CRV01"  'upd 2003/05/31 hitec)matsumoto
            '---------------------------2003/04/13 okazaki
            .NOWPROC = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU
            .DIAMETER = 0      '--------------�ۗ�
            .ALLSCRAP = "N" '�S���X�N���b�v
        End With
        
        '���������i�u���b�N�j����O�H�����ю擾
        strDBName = "XSDC2"
        If cmbc035_1_CreateXSDC2(strBlockID(intLoopCnt), bNoData) = FUNCTION_RETURN_FAILURE Then
            If bNoData = True Then
                cmbc035_1_CreateTable = FUNCTION_RETURN_SUCCESS '�����͍s��Ȃ����A����ŕԂ�
                Debug.Print "cmbc035_1_CreateXSDC2(" & strBlockID(intLoopCnt) & "," & bNoData & "):XSDC2�O�H�����і���"
                Exit Function
            Else
                cmbc035_1_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "cmbc035_1_CreateXSDC2(" & strBlockID(intLoopCnt) & "," & bNoData & "):XSDC2�O�H�����ѓǍ��݃G���["
                Exit Function
            End If
        End If
        
        '���������i�i�ԁj����O�H�����ю擾
        strDBName = "XSDCA"
        If cmbc035_1_CreateXSDCA(strBlockID(intLoopCnt), bNoData) = FUNCTION_RETURN_FAILURE Then
            If bNoData = True Then
                cmbc035_1_CreateTable = FUNCTION_RETURN_SUCCESS '�����͍s��Ȃ����A����ŕԂ�
                Debug.Print "XSDCA�F�O�H�����і���"
                Exit Function
            Else
                cmbc035_1_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "XSDCA�F�O�H�����ѓǍ��݃G���["
                Exit Function
            End If
        End If
        
        '���ݍH�����э쐬
        If cmbc035_1_CreateNowProc(strBlockID(intLoopCnt), lngBeginIngotpos, lngEndIngotpos) = FUNCTION_RETURN_FAILURE Then
            cmbc035_1_CreateTable = FUNCTION_RETURN_FAILURE
            strErrMsg = GetMsgStr("EAPLY")
            Debug.Print "XSDC2,XSDCA�F���ݍH�����э쐬�G���["
            Exit Function
        End If
        
        '��{����
''''        giInpos = 900   'del 2003/05/27 hitec)matsumoto
        If KihonProc = FUNCTION_RETURN_FAILURE Then
            cmbc035_1_CreateTable = FUNCTION_RETURN_FAILURE
            strErrMsg = GetMsgStr("EAPLY")
            Debug.Print "��{�����ُ�I��"
            Exit Function
        End If
        intLoopCnt = intLoopCnt + 1
        rs.MoveNext
    Loop
    rs.Close
                
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function
                
End Function


'���������i�i�ԁj�O�H�����ю擾���\���̍쐬 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc035_1_CreateXSDCA(ByVal strBlockID As String, ByRef bNoData As Boolean) As FUNCTION_RETURN

    Dim iLoopCnt    As Integer
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim intProcNo   As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0

    '�u���b�NID�𓾂�
    sql = "SELECT * from XSDCA"
    sql = sql & " WHERE CRYNUMCA='" & strBlockID & "'"
    sql = sql & "   AND LIVKCA= '0'"   '�����敪

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc035_1_CreateXSDCA = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    rs.MoveFirst
    iLoopCnt = 0
    
    Do While Not rs.EOF
        ReDim Preserve HinOld(iLoopCnt)
        ReDim Preserve HinNow(iLoopCnt)
        With HinOld(iLoopCnt)
            If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
            If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
            If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
            If IsNull(rs.Fields("REVNUMCA")) = False Then .REVNUMCA = rs.Fields("REVNUMCA")
            If IsNull(rs.Fields("FACTORYCA")) = False Then .FACTORYCA = rs.Fields("FACTORYCA")
            If IsNull(rs.Fields("OPECA")) = False Then .OPECA = rs.Fields("OPECA")
            If IsNull(rs.Fields("KCKNTCA")) = False Then .KCKNTCA = rs.Fields("KCKNTCA")
            If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLIDCA = rs.Fields("SXLIDCA")
            If IsNull(rs.Fields("XTALCA")) = False Then .XTALCA = rs.Fields("XTALCA")
            If IsNull(rs.Fields("NEKKNTCA")) = False Then .NEKKNTCA = rs.Fields("NEKKNTCA")
            If IsNull(rs.Fields("NEWKNTCA")) = False Then .NEWKNTCA = rs.Fields("NEWKNTCA")
            If IsNull(rs.Fields("NEWKKBCA")) = False Then .NEWKKBCA = rs.Fields("NEWKKBCA")
            If IsNull(rs.Fields("NEMACOCA")) = False Then .NEMACOCA = rs.Fields("NEMACOCA")
            If IsNull(rs.Fields("GNKKNTCA")) = False Then .GNKKNTCA = rs.Fields("GNKKNTCA")
            If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
            If IsNull(rs.Fields("GNWKKBCA")) = False Then .GNWKKBCA = rs.Fields("GNWKKBCA")
            If IsNull(rs.Fields("GNMACOCA")) = False Then .GNMACOCA = rs.Fields("GNMACOCA")
            If IsNull(rs.Fields("GNDAYCA")) = False Then .GNDAYCA = rs.Fields("GNDAYCA")
            If IsNull(rs.Fields("GNLCA")) = False Then .GNLCA = rs.Fields("GNLCA")
            If IsNull(rs.Fields("GNWCA")) = False Then .GNWCA = rs.Fields("GNWCA")
            If IsNull(rs.Fields("GNMCA")) = False Then .GNMCA = rs.Fields("GNMCA")
            If IsNull(rs.Fields("SUMITLCA")) = False Then .SUMITLCA = rs.Fields("SUMITLCA")
            If IsNull(rs.Fields("SUMITWCA")) = False Then .SUMITWCA = rs.Fields("SUMITWCA")
            If IsNull(rs.Fields("SUMITMCA")) = False Then .SUMITMCA = rs.Fields("SUMITMCA")
            If IsNull(rs.Fields("CHGCA")) = False Then .CHGCA = rs.Fields("CHGCA")
            If IsNull(rs.Fields("KAKOUBCA")) = False Then .KAKOUBCA = rs.Fields("KAKOUBCA")
            If IsNull(rs.Fields("KEIDAYCA")) = False Then .KEIDAYCA = rs.Fields("KEIDAYCA")
            If IsNull(rs.Fields("GNTKUBCA")) = False Then .GNTKUBCA = rs.Fields("GNTKUBCA")
            If IsNull(rs.Fields("GNTNOCA")) = False Then .GNTNOCA = rs.Fields("GNTNOCA")
            If IsNull(rs.Fields("XTWORKCA")) = False Then .XTWORKCA = rs.Fields("XTWORKCA")
            If IsNull(rs.Fields("WFWORKCA")) = False Then .WFWORKCA = rs.Fields("WFWORKCA")
            If IsNull(rs.Fields("LSTATBCA")) = False Then .LSTATBCA = rs.Fields("LSTATBCA")
            If IsNull(rs.Fields("RSTATBCA")) = False Then .RSTATBCA = rs.Fields("RSTATBCA")
            If IsNull(rs.Fields("LUFRCCA")) = False Then .LUFRCCA = rs.Fields("LUFRCCA")
            If IsNull(rs.Fields("LUFRBCA")) = False Then .LUFRBCA = rs.Fields("LUFRBCA")
            If IsNull(rs.Fields("LDFRCCA")) = False Then .LDFRCCA = rs.Fields("LDFRCCA")
            If IsNull(rs.Fields("LDFRBCA")) = False Then .LDFRBCA = rs.Fields("LDFRBCA")
            If IsNull(rs.Fields("HOLDCCA")) = False Then .HOLDCCA = rs.Fields("HOLDCCA")
            If IsNull(rs.Fields("HOLDBCA")) = False Then .HOLDBCA = rs.Fields("HOLDBCA")
            If IsNull(rs.Fields("EXKUBCA")) = False Then .EXKUBCA = rs.Fields("EXKUBCA")
            If IsNull(rs.Fields("HENPKCA")) = False Then .HENPKCA = rs.Fields("HENPKCA")
            If IsNull(rs.Fields("LIVKCA")) = False Then .LIVKCA = rs.Fields("LIVKCA")
            If IsNull(rs.Fields("KANKCA")) = False Then .KANKCA = rs.Fields("KANKCA")
            If IsNull(rs.Fields("NFCA")) = False Then .NFCA = rs.Fields("NFCA")
            If IsNull(rs.Fields("SAKJCA")) = False Then .SAKJCA = rs.Fields("SAKJCA")
            If IsNull(rs.Fields("TDAYCA")) = False Then .TDAYCA = rs.Fields("TDAYCA")
            If IsNull(rs.Fields("KDAYCA")) = False Then .KDAYCA = rs.Fields("KDAYCA")
            If IsNull(rs.Fields("SUMITBCA")) = False Then .SUMITBCA = rs.Fields("SUMITBCA")
            If IsNull(rs.Fields("SNDKCA")) = False Then .SNDKCA = rs.Fields("SNDKCA")
            If IsNull(rs.Fields("SNDDAYCA")) = False Then .SNDDAYCA = rs.Fields("SNDDAYCA")
        End With
        '�Ǖi�����Z�b�g
        With Kihon
            .CNTHINOLD = iLoopCnt + 1
        End With
        iLoopCnt = iLoopCnt + 1
        rs.MoveNext
    Loop
    
    rs.Close
    cmbc035_1_CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc035_1_CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'���������i�u���b�N�j�O�H�����ю擾���\���̍쐬 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc035_1_CreateXSDC2(ByVal strBlockID As String, ByRef bNoData As Boolean) As FUNCTION_RETURN

    Dim rs      As OraDynaset
    Dim sql     As String
    Dim intProcNo   As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False
    
    '�u���b�NID�𓾂�
    sql = "SELECT * from XSDC2 "
    sql = sql & " WHERE CRYNUMC2='" & strBlockID & "'"
    sql = sql & "   AND LIVKC2= '0'"   '�����敪

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc035_1_CreateXSDC2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    rs.MoveFirst
    If rs.EOF = False Then
        With BlkOld
            If IsNull(rs.Fields("CRYNUMC2")) = False Then .CRYNUMC2 = rs.Fields("CRYNUMC2")
            If IsNull(rs.Fields("KCNTC2")) = False Then .KCNTC2 = rs.Fields("KCNTC2")       '�H���A��
            If IsNull(rs.Fields("XTALC2")) = False Then .XTALC2 = rs.Fields("XTALC2")
            If IsNull(rs.Fields("INPOSC2")) = False Then .INPOSC2 = rs.Fields("INPOSC2")
            If IsNull(rs.Fields("NEKKNTC2")) = False Then .NEKKNTC2 = rs.Fields("NEKKNTC2")
            If IsNull(rs.Fields("NEWKNTC2")) = False Then .NEWKNTC2 = rs.Fields("NEWKNTC2")
            If IsNull(rs.Fields("NEWKKBC2")) = False Then .NEWKKBC2 = rs.Fields("NEWKKBC2")
            If IsNull(rs.Fields("NEMACOC2")) = False Then .NEMACOC2 = rs.Fields("NEMACOC2")
            If IsNull(rs.Fields("GNKKNTC2")) = False Then .GNKKNTC2 = rs.Fields("GNKKNTC2")
            If IsNull(rs.Fields("GNWKNTC2")) = False Then .GNWKNTC2 = rs.Fields("GNWKNTC2")
            If IsNull(rs.Fields("GNWKKBC2")) = False Then .GNWKKBC2 = rs.Fields("GNWKKBC2")
            If IsNull(rs.Fields("GNMACOC2")) = False Then .GNMACOC2 = rs.Fields("GNMACOC2")
            If IsNull(rs.Fields("GNDAYC2")) = False Then .GNDAYC2 = rs.Fields("GNDAYC2")
            If IsNull(rs.Fields("GNLC2")) = False Then .GNLC2 = rs.Fields("GNLC2")          '���ݒ���
            If IsNull(rs.Fields("GNWC2")) = False Then .GNWC2 = rs.Fields("GNWC2")          '���ݏd��
            If IsNull(rs.Fields("GNMC2")) = False Then .GNMC2 = rs.Fields("GNMC2")          '���ݖ���
            If IsNull(rs.Fields("SUMITLC2")) = False Then .SUMITLC2 = rs.Fields("SUMITLC2")
            If IsNull(rs.Fields("SUMITWC2")) = False Then .SUMITWC2 = rs.Fields("SUMITWC2")
            If IsNull(rs.Fields("SUMITMC2")) = False Then .SUMITMC2 = rs.Fields("SUMITMC2")
            If IsNull(rs.Fields("CHGC2")) = False Then .CHGC2 = rs.Fields("CHGC2")
            If IsNull(rs.Fields("KAKOUBC2")) = False Then .KAKOUBC2 = rs.Fields("KAKOUBC2")
            If IsNull(rs.Fields("KEIDAYC2")) = False Then .KEIDAYC2 = rs.Fields("KEIDAYC2")
            If IsNull(rs.Fields("GNTKUBC2")) = False Then .GNTKUBC2 = rs.Fields("GNTKUBC2")
            If IsNull(rs.Fields("GNTNOC2")) = False Then .GNTNOC2 = rs.Fields("GNTNOC2")
            If IsNull(rs.Fields("XTWORKC2")) = False Then .XTWORKC2 = rs.Fields("XTWORKC2")
            If IsNull(rs.Fields("WFWORKC2")) = False Then .WFWORKC2 = rs.Fields("WFWORKC2")
            If IsNull(rs.Fields("LSTATBC2")) = False Then .LSTATBC2 = rs.Fields("LSTATBC2")
            If IsNull(rs.Fields("RSTATBC2")) = False Then .RSTATBC2 = rs.Fields("RSTATBC2")
            If IsNull(rs.Fields("LUFRCC2")) = False Then .LUFRCC2 = rs.Fields("LUFRCC2")
            If IsNull(rs.Fields("LUFRBC2")) = False Then .LUFRBC2 = rs.Fields("LUFRBC2")
            If IsNull(rs.Fields("LDFRCC2")) = False Then .LDFRCC2 = rs.Fields("LDFRCC2")
            If IsNull(rs.Fields("LDFRBC2")) = False Then .LDFRBC2 = rs.Fields("LDFRBC2")
            If IsNull(rs.Fields("HOLDCC2")) = False Then .HOLDCC2 = rs.Fields("HOLDCC2")
            If IsNull(rs.Fields("HOLDBC2")) = False Then .HOLDBC2 = rs.Fields("HOLDBC2")
            If IsNull(rs.Fields("EXKUBC2")) = False Then .EXKUBC2 = rs.Fields("EXKUBC2")
            If IsNull(rs.Fields("HENPKC2")) = False Then .HENPKC2 = rs.Fields("HENPKC2")
            If IsNull(rs.Fields("LIVKC2")) = False Then .LIVKC2 = rs.Fields("LIVKC2")
            If IsNull(rs.Fields("KANKC2")) = False Then .KANKC2 = rs.Fields("KANKC2")
            If IsNull(rs.Fields("NFC2")) = False Then .NFC2 = rs.Fields("NFC2")
            If IsNull(rs.Fields("SAKJC2")) = False Then .SAKJC2 = rs.Fields("SAKJC2")
            If IsNull(rs.Fields("TDAYC2")) = False Then .TDAYC2 = rs.Fields("TDAYC2")
            If IsNull(rs.Fields("KDAYC2")) = False Then .KDAYC2 = rs.Fields("KDAYC2")
            If IsNull(rs.Fields("SUMITBC2")) = False Then .SUMITBC2 = rs.Fields("SUMITBC2")
            If IsNull(rs.Fields("SNDKC2")) = False Then .SNDKC2 = rs.Fields("SNDKC2")
            If IsNull(rs.Fields("SNDDAYC2")) = False Then .SNDDAYC2 = rs.Fields("SNDDAYC2")
        End With
    End If
    
    rs.Close
    cmbc035_1_CreateXSDC2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc035_1_CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO

'���ݍH���\���̍쐬 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc035_1_CreateNowProc(ByVal strBlockID As String, ByVal lngBeginIngotpos As Long, ByVal lngEndIngotpos As Long) As FUNCTION_RETURN

    Dim rs      As OraDynaset
    Dim sql     As String
    Dim intProcNo   As Integer
    Dim intHinOldCnt    As Integer
    Dim intLengthCnt    As Integer
    Dim intLoopCnt      As Integer
    Dim dblDiameter     As Double
    Dim intNum          As Integer
    Dim strCryNum       As String
    Dim strLstatcls     As String
    Dim intBlkLength    As Integer  '�u���b�N�Ǘ��f�[�^�̒���
    Dim intBlkIngotPos  As Integer  '�u���b�N�Ǘ��f�[�^�̈ʒu
    Dim intSxlLength    As Integer  '�V���O���Ǘ��f�[�^�̒���
    Dim intSxlIngotPos  As Integer  '�V���O���Ǘ��f�[�^�̈ʒu
    Dim bFlg            As Boolean
    Dim sp              As Integer  '��������p
    Dim ep              As Integer  '��������p
    Dim sbp             As Integer  '��������p
    Dim ebp             As Integer  '��������p
    Dim intLength       As Integer  '����
    Dim intIngotPos     As Integer  '�ʒu
    Dim lngSumGNWCA     As Long     'add 2003/05/20 hitec)matsumoto
    Dim lngSumGNMCA     As Long     'add 2003/05/20 hitec)matsumoto
    Dim bChgFlg         As Boolean  'add 2003/05/20 hitec)matsumoto
    Dim i               As Integer  'add 2003/05/20 hitec)matsumoto

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0
    
    intBlkLength = 0
    intBlkIngotPos = 0
    intSxlLength = 0
    intSxlIngotPos = 0
    strCryNum = ""

    '�u���b�N�Ǘ����璷�����擾
    sql = "SELECT * from TBCME040 "
    sql = sql & " WHERE BLOCKID='" & strBlockID & "'"
''''    sql = sql & "   AND INGOTPOS=0"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc035_1_CreateNowProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    intLoopCnt = 0
    If rs.EOF = False Then
        If IsNull(rs("CRYNUM")) = False Then strCryNum = rs("CRYNUM")               '�����ԍ�
        If IsNull(rs("LENGTH")) = False Then intBlkLength = rs("LENGTH")            '����
        If IsNull(rs("INGOTPOS")) = False Then intBlkIngotPos = rs("INGOTPOS")      '�ʒu
    End If

    rs.Close

    '�u���b�N�Ǘ��Ŏ擾�������������ƂɃV���O���Ǘ�����f�[�^���擾
    sql = "SELECT * from TBCME042 "
    sql = sql & " WHERE CRYNUM='" & strCryNum & "'"
    '�����[�v���Ŕ���
    sql = sql & "   AND ((INGOTPOS >=" & lngBeginIngotpos & ") And (INGOTPOS + LENGTH <= " & lngEndIngotpos & "))"
    sql = sql & "   AND LSTATCLS<>'H'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then   '�Y���f�[�^0���̏ꍇ�A�S���X�N���b�v�̏���
        '�O�H�����т��A���ݍH�����тɃR�s�[
        BlkNow = BlkOld
        BlkNow.GNLC2 = "0"
        BlkNow.GNWC2 = "0"
        BlkNow.GNMC2 = "0"
        BlkNow.GNWKNTC2 = Kihon.NEWPROC
        BlkNow.NEWKNTC2 = Kihon.NOWPROC
        For intHinOldCnt = 0 To Kihon.CNTHINOLD - 1
            ReDim Preserve HinNow(intHinOldCnt) As typ_XSDCA_Update
            HinNow(intHinOldCnt) = HinOld(intHinOldCnt)
            HinNow(intHinOldCnt).GNLCA = "0"    '�S���X�N���b�v=������0
            HinNow(intHinOldCnt).GNWCA = "0"    '�d�� = 0
            HinNow(intHinOldCnt).GNMCA = "0"    '���� = 0
            HinNow(intHinOldCnt).GNWKNTCA = Kihon.NEWPROC
            HinNow(intHinOldCnt).NEWKNTCA = Kihon.NOWPROC
        Next
        Kihon.CNTHINNOW = 1
        Kihon.ALLSCRAP = "Y"
        
        '�O�H���̒����ƌ��ݍH���̒���������ׁA�s�ǂ����݂��邩����
        If CInt(BlkNow.GNLC2) = CInt(BlkOld.GNLC2) Then '�s�ǂȂ�
            '��{���\����
            With Kihon
                .FURYOUMU = "N"
            End With
        Else                                            '�s�ǂ���
            '��{���\����
            With Kihon
                .FURYOUMU = "Y"
            End With
            '�s�Ǎ\���̂��쐬
            With Furyou
                .XTALC4 = BlkNow.CRYNUMC2   '�u���b�NID
                .INPOSC4 = BlkNow.INPOSC2   '�������J�n�ʒu
                .KCKNTC4 = BlkNow.KCNTC2    '�H���A��
                .HINBC4 = "Z"               '�i��
    '            .REVNUMC4                   '���i�ԍ������ԍ�
    '            .FACTORYC4                  '�H��
    '            .OPEC4                      '���Ə���
     '           .WKKTC4 = PROCD_NUKISI_HENKOU
                .WKKTC4 = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU      ' 2003/04/12 okazaki
                '�u�������v�B�o�^�O�ɂ�����x�s�ǒ����E�d�ʁE���������߂Ȃ��� start -------------
                .PUCUTLC4 = CLng(BlkNow.GNLC2) - CLng(BlkOld.GNLC2) '�s�ǒ���(�O�H��-���ݍH���i�Ǖi�j)
                .PUCUTWC4 = CLng(BlkNow.GNWC2) - CLng(BlkOld.GNWC2)    'upd 2003/05/31 hitec)matsumoto �d�ʂ͍Čv�Z���Ȃ�
                .PUCUTMC4 = 0 'upd 2003/05/31 hitec)matsumoto �����͍Čv�Z���Ȃ�
                .SUMITBC3 = "0"
            End With
        End If
        rs.Close
        cmbc035_1_CreateNowProc = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    '�O�H���̍\���̂����ݍH���̍\���̂փR�s�[
    BlkNow = BlkOld
    '�H���A�ԂɁ{�P����
    With BlkNow
        .KCNTC2 = CInt(.KCNTC2) + 1         '�H���A��
        .NEWKNTC2 = Kihon.NOWPROC         '�O�H��
        .GNWKNTC2 = Kihon.NEWPROC           '���ݍH��
        .SUMITLC2 = "0"                     'SUMMIT����
        .SUMITMC2 = "0"                   'SUMMIT����
        .SUMITWC2 = "0"                   'SUMMIT�d��
        .SUMITBC2 = "0"
    End With
    
    intLoopCnt = 0
    BlkNow.GNLC2 = 0    '���ݍH���i�u���b�N�j�̒������N���A���Ă���
    BlkNow.GNWC2 = 0    '���ݍH���i�u���b�N�j�̒������N���A���Ă���
    BlkNow.GNMC2 = 0    '���ݍH���i�u���b�N�j�̒������N���A���Ă���
    
    Do While Not rs.EOF
        ReDim Preserve HinNow(intLoopCnt) As typ_XSDCA_Update
        '�O�H���̍\���̂����ݍH���̍\���̂փR�s�[
''''        HinNow(intLoopCnt) = HinOld(intHinOldCnt)
        
        If IsNull(rs("CRYNUM")) = False Then strCryNum = rs("CRYNUM")               '�����ԍ�
        If IsNull(rs("LENGTH")) = False Then intSxlLength = rs("LENGTH")            '����
        If IsNull(rs("INGOTPOS")) = False Then intSxlIngotPos = rs("INGOTPOS")      '�ʒu
        
        '-- �u���b�N�ƃV���O���̈ʒu�֌W�𔻒肵�A�������Z�o --------
        sp = intSxlIngotPos         '�V���O���J�n�ʒu
        ep = sp + intSxlLength      '�V���O���I�[�ʒu
        sbp = intBlkIngotPos        '�u���b�N�J�n�ʒu
        ebp = sbp + intBlkLength    '�u���b�N�I�[�ʒu
        
        '' �u���b�N��SXL�̒��Ɋ��S�Ɋ܂܂�Ă���ꍇ ---------
        If sp <= sbp And ep >= ebp Then
        
            intLength = intBlkLength                    '�u���b�N�Ǘ��̒������g�p
            intIngotPos = intBlkIngotPos
            
        '' �u���b�N��SXL�̊J�n�ʒu����ɂ���A���I�[�ʒu���������ꍇ ---------
        ElseIf sp >= sbp And ep <= ebp Then
            
            intLength = intSxlLength                  '�V���O���Ǘ��̒������g�p
            intIngotPos = intSxlIngotPos
            
        '' �u���b�N���ꕔSXL�ɂ������Ă���ꍇ
        '' (�u���b�N���㑤�B�������u���b�N�̏I�[��SXL�̊J�n�ʒu����v���Ȃ�����) ------------
        ElseIf sp > sbp And sp < ebp And sp <> ebp Then
            
            intLength = ebp - sp                        '�u���b�N�̏I�[�ʒu - �V���O���̊J�n�ʒu
            intIngotPos = intSxlIngotPos
        
        '' �u���b�N���ꕔSXL�ɂ������Ă���ꍇ
        '' (�u���b�N�������B������SXL�̏I�[�ƃu���b�N�̊J�n�ʒu����v���Ȃ�����) ----------
        ElseIf sp < sbp And ep > sbp And ep <> sbp Then
            
            intLength = ep - sbp                        '�V���O���̏I�[�ʒu - �u���b�N�̊J�n�ʒu
            intIngotPos = intBlkIngotPos
            
        Else
        
            GoTo LoopNext

        End If
        '----------------------------------------------------
        
        '���ݍH���ҏW
        With HinNow(intLoopCnt)
            If IsNull(rs("CRYNUM")) = False Then .XTALCA = rs("CRYNUM")
            .CRYNUMCA = strBlockID         '�u���b�NID
            If IsNull(rs("HINBAN")) = False Then .HINBCA = rs("HINBAN")         '�i��
            If IsNull(rs("REVNUM")) = False Then .REVNUMCA = rs("REVNUM")       '���i�ԍ������ԍ�
            If IsNull(rs("FACTORY")) = False Then .FACTORYCA = rs("FACTORY")    '�H��
            If IsNull(rs("OPECOND")) = False Then .OPECA = rs("OPECOND")        '���Ə���
            .INPOSCA = intIngotPos    '�������J�n�ʒu
            .GNLCA = intLength          '����
            BlkNow.GNLC2 = CStr(CLng(BlkNow.GNLC2) + CLng(HinNow(intLoopCnt).GNLCA))  '����
            If IsNull(rs("SXLID")) = False Then .SXLIDCA = rs("SXLID")          '�V���O��ID
            .SUMITBCA = 0
            .SUMITLCA = 0
            .SUMITMCA = 0
            .SUMITWCA = 0
            .NEWKNTCA = Kihon.NOWPROC   '�O�H��
            .GNWKNTCA = Kihon.NEWPROC   '���ݍH��
            .KCKNTCA = BlkNow.KCNTC2    '�H���A��
            .NEMACOCA = BlkNow.NEMACOC2 '�ŏI�ʉߏ�����
            .GNMACOCA = BlkNow.GNMACOC2 '���ݏ�����
''''            .XTALCA = strCryNum         '�����ԍ�
            '���ݏd�ʂ����߂�
            If GetDiameter(strBlockID, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '���a�����߂�
                dblDiameter = 0
''''                GoTo proc_wxit
            End If
            '��{���̒��a�Z�b�g
            Kihon.DIAMETER = dblDiameter
            
            '�擾�������a�����ɏd�ʂ����߂�
            .GNWCA = CStr(CLng(WeightOfCylinder(dblDiameter, CDbl(.GNLCA))))
            
            '���ݖ��������߂�
            If WfCount(strBlockID, CLng(.GNLCA), intNum) = FUNCTION_RETURN_FAILURE Then
                .GNMCA = 0
''''                GoTo proc_wxit
            Else
                .GNMCA = intNum
            End If
        End With
        
        With BlkNow
            '���ݏd�ʂ����߂�
            If GetDiameter(strBlockID, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '���a�����߂�
                dblDiameter = 0
    ''''                GoTo proc_wxit
            End If
            '��{���̒��a�Z�b�g
            Kihon.DIAMETER = dblDiameter
            '�擾�������a�����ɏd�ʂ����߂�
            .GNWC2 = CStr(CLng(WeightOfCylinder(dblDiameter, CDbl(.GNLC2))))
            '���ݖ��������߂�
            If WfCount(strBlockID, CLng(.GNLC2), intNum) = FUNCTION_RETURN_FAILURE Then
                .GNMC2 = 0
''''                GoTo proc_wxit
            Else
                .GNMC2 = intNum
            End If
            
        End With
        intLoopCnt = intLoopCnt + 1
        '�Ǖi�����Z�b�g
        With Kihon
            .CNTHINNOW = intLoopCnt
        End With

LoopNext:

        rs.MoveNext
    Loop
    
    rs.Close
    
    '�O�H���̒����ƌ��ݍH���̒���������ׁA�s�ǂ����݂��邩����
    If CInt(BlkNow.GNLC2) = CInt(BlkOld.GNLC2) Then '�s�ǂȂ�
        '��{���\����
        With Kihon
            .FURYOUMU = "N"
        End With
    Else                                            '�s�ǂ���
        '��{���\����
        With Kihon
            .FURYOUMU = "Y"
        End With
        '�s�Ǎ\���̂��쐬
        With Furyou
            .XTALC4 = BlkNow.CRYNUMC2   '�u���b�NID
            .INPOSC4 = BlkNow.INPOSC2   '�������J�n�ʒu
            .KCKNTC4 = BlkNow.KCNTC2    '�H���A��
            .HINBC4 = "Z"               '�i��
''            .REVNUMC4 =                '���i�ԍ������ԍ�
''            .FACTORYC4                  '�H��
''            .OPEC4                      '���Ə���
'            .WKKTC4 = PROCD_NUKISI_HENKOU
            .WKKTC4 = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU    ' 2003/04/12 okazaki
            '�u�������v�B�o�^�O�ɂ�����x�s�ǒ����E�d�ʁE���������߂Ȃ��� start -------------
            .PUCUTLC4 = CLng(BlkNow.GNLC2) - CLng(BlkOld.GNLC2) '�s�ǒ���(�O�H��-���ݍH���i�Ǖi�j)
            .PUCUTWC4 = CLng(BlkNow.GNWC2) - CLng(BlkOld.GNWC2)    'upd 2003/05/31 hitec)matsumoto �d�ʂ͍Čv�Z���Ȃ�
            .PUCUTMC4 = 0 'upd 2003/05/31 hitec)matsumoto �����͍Čv�Z���Ȃ�
        End With
    End If
    
    cmbc035_1_CreateNowProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc035_1_CreateNowProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

' ^^^^^ 2003.04.11 ADD BY HITEC)��c  END

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

'�T�v      :�e�[�u���uTBCME042�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME042 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCME042_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME042(records() As typ_TBCME042, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, SXLID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS, HOLDCLS," & _
              " HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, COUNT, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE, " & _
              " PASSFLAG "   '02/04/16 Yam
    sqlBase = sqlBase & "From TBCME042"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME042 = FUNCTION_RETURN_FAILURE
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
            .SXLID = rs("SXLID")             ' SXLID
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H��
            .NOWPROC = rs("NOWPROC")         ' ���ݍH��
            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
            .DELCLS = rs("DELCLS")           ' �폜�敪
            .LSTATCLS = rs("LSTATCLS")       ' �ŏI��ԋ敪
            .HOLDCLS = rs("HOLDCLS")         ' �z�[���h�敪
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .BDCAUS = rs("BDCAUS")           ' �s�Ǘ��R
            .Count = rs("COUNT")             ' ����
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
            .PASSFLAG = " "   ' �ʉ߃t���O�̃X�y�[�X�N���A '02/04/16 Yam
            If rs("PASSFLAG") = "1" Then
                .PASSFLAG = rs("PASSFLAG")   ' �ʉ߃t���O '02/04/05 Yam
            End If
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME042 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uXSDCW�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_XSDCW    ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCME044_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME044(records() As typ_XSDCW, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
Dim tHIN As tFullHinban '03/05/24
Dim sOT1    As String   '03/05/24
Dim sOT2    As String
Dim sMAI1    As String   '04/06/25
Dim sMAI2    As String
Dim rtn     As FUNCTION_RETURN
    ''SQL��g�ݗ��Ă�   '03/05/24
'    sqlBase = "Select CRYNUM, INGOTPOS, SMPKBN, SMPLID, HINBAN, REVNUM, FACTORY, OPECOND, KTKBN, WFINDRS, WFINDOI, WFINDB1," & _
'              " WFINDB2, WFINDB3, WFINDL1, WFINDL2, WFINDL3, WFINDL4, WFINDDS, WFINDDZ, WFINDSP, WFINDDO1, WFINDDO2, WFINDDO3," & _
'              " NVL(WFINDOT1,'0') as DOT1, NVL(WFINDOT2,'0') as DOT2," & _
'              " WFRESRS, WFRESOI, WFRESB1, WFRESB2, WFRESB3, WFRESL1, WFRESL2, WFRESL3, WFRESL4, WFRESDS, WFRESDZ, WFRESSP," & _
'              " WFRESDO1, WFRESDO2, WFRESDO3,NVL(WFRESOT1,'0') as SOT1, NVL(WFRESOT2,'0') as SOT2, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCME044"

    'GD���ڒǉ��@05/01/25 ooba
    sqlBase = "Select SXLIDCW, SMPKBNCW, TBKBNCW, REVNUMCW, XTALCW, INPOSCW, REPSMPLIDCW, HINBCW, FACTORYCW, OPECW, KTKBNCW, " & _
              " SMCRYNUMCW, WFSMPLIDRSCW, WFSMPLIDRS1CW, WFSMPLIDRS2CW, WFINDRSCW, WFRESRS1CW, WFRESRS2CW, WFSMPLIDOICW, WFINDOICW, " & _
              " WFRESOICW, WFSMPLIDB1CW, WFINDB1CW, WFRESB1CW, WFSMPLIDB2CW, WFINDB2CW, WFRESB2CW, WFSMPLIDB3CW, WFINDB3CW, " & _
              " WFRESB3CW, WFSMPLIDL1CW, WFINDL1CW, WFRESL1CW, WFSMPLIDL2CW, WFINDL2CW, WFRESL2CW, WFSMPLIDL3CW, WFINDL3CW, WFRESL3CW, " & _
              " WFSMPLIDL4CW, WFINDL4CW, WFRESL4CW, WFSMPLIDDSCW, WFINDDSCW, WFRESDSCW, WFSMPLIDDZCW, WFINDDZCW, WFRESDZCW, " & _
              " WFSMPLIDSPCW, WFINDSPCW, WFRESSPCW, WFSMPLIDDO1CW, WFINDDO1CW, WFRESDO1CW, WFSMPLIDDO2CW, WFINDDO2CW, WFRESDO2CW, " & _
              " WFSMPLIDDO3CW, WFINDDO3CW, WFRESDO3CW, WFSMPLIDOT1CW, NVL(WFINDOT1CW,'0') as DOT1, NVL(WFRESOT1CW,'0') as SOT1, " & _
              " WFSMPLIDOT2CW, NVL(WFINDOT2CW,'0') as DOT2, NVL(WFRESOT2CW,'0') as SOT2, WFSMPLIDAOICW, WFINDAOICW, WFRESAOICW, SMPLNUMCW, " & _
              " WFSMPLIDGDCW, WFINDGDCW, WFRESGDCW, WFHSGDCW, " & _
              " SMPLPATCW, TSTAFFCW, TDAYCW, KSTAFFCW, KDAYCW, SNDKCW, SNDDAYCW "
    sqlBase = sqlBase & "From XSDCW"

    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME044 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .SXLIDCW = rs("SXLIDCW")
            .SMPKBNCW = rs("SMPKBNCW")           ' �T���v���敪
            .TBKBNCW = rs("TBKBNCW")
            .REPSMPLIDCW = rs("REPSMPLIDCW")           ' �T���v��ID
            .XTALCW = rs("XTALCW")           ' �����ԍ�
            .INPOSCW = rs("INPOSCW")       ' �������ʒu
            .HINBCW = rs("HINBCW")           ' �i��
            .REVNUMCW = rs("REVNUMCW")           ' ���i�ԍ������ԍ�
            .FACTORYCW = rs("FACTORYCW")         ' �H��
            .OPECW = rs("OPECW")         ' ���Ə���
            .KTKBNCW = rs("KTKBNCW")             ' �m��敪
            .SMCRYNUMCW = rs("SMCRYNUMCW")
            .WFSMPLIDRSCW = rs("WFSMPLIDRSCW")
            If Not IsNull(rs("WFSMPLIDRS1CW")) Then .WFSMPLIDRS1CW = rs("WFSMPLIDRS1CW")
            If Not IsNull(rs("WFSMPLIDRS2CW")) Then .WFSMPLIDRS2CW = rs("WFSMPLIDRS2CW")
            .WFINDRSCW = rs("WFINDRSCW")         ' WF�����w���iRs)
            .WFRESRS1CW = rs("WFRESRS1CW")         ' WF�������сiRs)
            If Not IsNull(rs("WFRESRS2CW")) Then .WFRESRS2CW = rs("WFRESRS2CW")
            .WFSMPLIDOICW = rs("WFSMPLIDOICW")
            .WFINDOICW = rs("WFINDOICW")         ' WF�����w���iOi)
            .WFRESOICW = rs("WFRESOICW")         ' WF�������сiOi)
            .WFSMPLIDB1CW = rs("WFSMPLIDB1CW")
            .WFINDB1CW = rs("WFINDB1CW")         ' WF�����w���iB1)
            .WFRESB1CW = rs("WFRESB1CW")
            .WFSMPLIDB2CW = rs("WFSMPLIDB2CW")
            .WFINDB2CW = rs("WFINDB2CW")         ' WF�����w���iB2�j
            .WFRESB2CW = rs("WFRESB2CW")         ' WF�������сiB2�j
            .WFSMPLIDB3CW = rs("WFSMPLIDB3CW")
            .WFINDB3CW = rs("WFINDB3CW")         ' WF�����w���iB3)
            .WFRESB3CW = rs("WFRESB3CW")         ' WF�������сiB3)
            .WFSMPLIDL1CW = rs("WFSMPLIDL1CW")
            .WFINDL1CW = rs("WFINDL1CW")         ' WF�����w���iL1)
            .WFRESL1CW = rs("WFRESL1CW")         ' WF�������сiL1)
            .WFSMPLIDL2CW = rs("WFSMPLIDL2CW")
            .WFINDL2CW = rs("WFINDL2CW")         ' WF�����w���iL2)
            .WFRESL2CW = rs("WFRESL2CW")         ' WF�������сiL2)
            .WFSMPLIDL3CW = rs("WFSMPLIDL3CW")
            .WFINDL3CW = rs("WFINDL3CW")         ' WF�����w���iL3)
            .WFRESL3CW = rs("WFRESL3CW")         ' WF�������сiL3)
            .WFSMPLIDL4CW = rs("WFSMPLIDL4CW")
            .WFINDL4CW = rs("WFINDL4CW")         ' WF�����w���iL4)
            .WFRESL4CW = rs("WFRESL4CW")         ' WF�������сiL4)
            .WFSMPLIDDSCW = rs("WFSMPLIDDSCW")
            .WFINDDSCW = rs("WFINDDSCW")         ' WF�����w���iDS)
            .WFRESDSCW = rs("WFRESDSCW")         ' WF�������сiDS)
            .WFSMPLIDDZCW = rs("WFSMPLIDDZCW")
            .WFINDDZCW = rs("WFINDDZCW")         ' WF�����w���iDZ)
            .WFRESDZCW = rs("WFRESDZCW")         ' WF�������сiDZ)
            .WFSMPLIDSPCW = rs("WFSMPLIDSPCW")
            .WFINDSPCW = rs("WFINDSPCW")         ' WF�����w���iSP)
            .WFRESSPCW = rs("WFRESSPCW")         ' WF�������сiSP)
            .WFSMPLIDDO1CW = rs("WFSMPLIDDO1CW")
            .WFINDDO1CW = rs("WFINDDO1CW")       ' WF�����w���iDO1)
            .WFRESDO1CW = rs("WFRESDO1CW")       ' WF�������сiDO1)
            .WFSMPLIDDO2CW = rs("WFSMPLIDDO2CW")
            .WFINDDO2CW = rs("WFINDDO2CW")       ' WF�����w���iDO2)
            .WFRESDO2CW = rs("WFRESDO2CW")       ' WF�������сiDO2)
            .WFSMPLIDDO3CW = rs("WFSMPLIDDO3CW")
            .WFINDDO3CW = rs("WFINDDO3CW")       ' WF�����w���iDO3)
            .WFRESDO3CW = rs("WFRESDO3CW")       ' WF�������сiDO3)
            tHIN.hinban = .HINBCW   ''03/05/24
            tHIN.factory = .FACTORYCW
            tHIN.mnorevno = .REVNUMCW
            tHIN.opecond = .OPECW
            rtn = scmzc_getE036(tHIN, sOT1, sOT2, sMAI1, sMAI2)
            If rtn = FUNCTION_RETURN_FAILURE Then
                rs.Close
                DBDRV_GetTBCME044 = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
            If sOT1 = "1" Then
                .WFINDOT1CW = rs!DOT1 '03/05/23
            Else
                .WFINDOT1CW = 0 '03/05/23
            End If
            If sOT2 = "1" Then
                .WFINDOT2CW = rs!DOT2 '03/05/23
            Else
                .WFINDOT2CW = 0 '03/05/23
            End If
            '#####################################################03/05/23 �㓡
            .WFRESOT1CW = rs("SOT1")       ' WF�������сiOT1)
            .WFRESOT2CW = rs("SOT2")       ' WF�������сiOT2)
            '#####################################################03/05/23 �㓡
            If Not IsNull(rs("WFSMPLIDAOICW")) Then .WFSMPLIDAOICW = rs("WFSMPLIDAOICW")
            If Not IsNull(rs("WFINDAOICW")) Then .WFINDAOICW = rs("WFINDAOICW")
            If Not IsNull(rs("WFRESAOICW")) Then .WFRESAOICW = rs("WFRESAOICW")
            If Not IsNull(rs("SMPLNUMCW")) Then .SMPLNUMCW = rs("SMPLNUMCW")
            If Not IsNull(rs("SMPLPATCW")) Then .SMPLPATCW = rs("SMPLPATCW")
            If Not IsNull(rs("TSTAFFCW")) Then .TSTAFFCW = rs("TSTAFFCW")
            .TDAYCW = rs("TDAYCW")         ' �o�^���t
            If Not IsNull(rs("KSTAFFCW")) Then .KSTAFFCW = rs("KSTAFFCW")
            .KDAYCW = rs("KDAYCW")         ' �X�V���t
            If Not IsNull(rs("SNDKCW")) Then .SNDKCW = rs("SNDKCW")       ' ���M�t���O
            If Not IsNull(rs("SNDDAYCW")) Then .SNDDAYCW = rs("SNDDAYCW")       ' ���M���t
            '' GD���ڎ擾�@05/01/25 ooba START ==================================================>
            If Not IsNull(rs("WFSMPLIDGDCW")) Then .WFSMPLIDGDCW = rs("WFSMPLIDGDCW")
            If Not IsNull(rs("WFINDGDCW")) Then .WFINDGDCW = rs("WFINDGDCW")    ' WF�����w�� (GD)
            If Not IsNull(rs("WFRESGDCW")) Then .WFRESGDCW = rs("WFRESGDCW")    ' WF�������� (GD)
            If Not IsNull(rs("WFHSGDCW")) Then .WFHSGDCW = rs("WFHSGDCW")       ' WF�����ۏ� (GD)
            '' GD���ڎ擾�@05/01/25 ooba END ====================================================>
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME044 = FUNCTION_RETURN_SUCCESS
End Function



