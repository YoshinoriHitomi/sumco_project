Attribute VB_Name = "s_cmmc001db2"
Option Explicit

'Type type_HinbanSyutoku   ''s_cmmc001db_sql�@��OUT�p
'    CRYNUM      As String * 12  ''�����ԍ�
'    INGOTPOS    As Integer      ''�g�b�v�T���v���ʒu
'    HINBAN      As String * 12  ''�{�g���T���v���ʒu
'    BLOCKID     As String * 12  ''�`���[�W��
'    LENGTH      As Integer      ''�g�b�v�d��
'End Type

'�T�v      :
'���Ұ�    :�ϐ���        ,IO ,�^             ,����
'          :sCryNum       ,I   ,String        ,���͗p
'          :pTbcmj002()   ,O   ,typ_TBCMJ002  ,��R���я��\����
'����      :
'����      :2001/06/28�@���с@�쐬
'�@�@      :2001/08/08�@S.Sano����
'�@�@      :2006/04/27�@�E�c�@����      �T���v���L���Ή�
'Public Function s_cmmc001db2_sql(all As typ_AllTypes) As FUNCTION_RETURN
    
Public Function s_cmmc001db2_sql(sCrynum As String, ADDDPPOS As Integer, FREELENG As Integer, INGOTPOS As Integer, typ_rsz() As typ_TBCMJ002) As FUNCTION_RETURN
    Dim temp() As typ_TBCMJ002
    Dim sql As String
    Dim recCnt As Long      '���R�[�h��
    Dim c0 As Integer
    Dim c1 As Integer
'    Dim ret As Integer
'    Dim pos(2) As Long
    Dim rs As OraDynaset
'    Dim rsz() As typ_TBCMJ002
    Dim MaxMin As String
    
    s_cmmc001db2_sql = FUNCTION_RETURN_FAILURE
    
    ReDim typ_rsz(2)
    
'    sql = "select " & _
'        "INGOTPOS " & _
'        " from TBCME040 " & _
'        " where CRYNUM = '" & sCryNum & "' and " & _
'            "INGOTPOS = ANY (SELECT MIN(INGOTPOS) FROM TBCME040 " & _
'                        "WHERE CRYNUM = '" & sCryNum & "')"
'
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'    pos(1) = rs("INGOTPOS")
'    rs.Close
'
'    sql = "select " & _
'        "INGOTPOS, LENGTH " & _
'        " from TBCME040 " & _
'        " where CRYNUM = '" & sCryNum & "' and " & _
'            "INGOTPOS = ANY (SELECT MAX(INGOTPOS) FROM TBCME040 " & _
'                        "WHERE CRYNUM = '" & sCryNum & "')"
'
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'    pos(2) = CDbl(rs("INGOTPOS")) + CDbl(rs("LENGTH"))
'    rs.Close
'
'    For ii = 1 To 2
'        sql = " where CRYNUM = '" & sCryNum & "' and " & _
'            "POSITION = " & pos(ii) & " and " & _
'            "TRANCNT = ANY (SELECT MAX(TRANCNT) FROM TBCMJ002 " & _
'                    "WHERE CRYNUM = '" & sCryNum & "' and " & _
'                        "POSITION = " & pos(ii) & ")"
'        ReDim rsz(0)
'        '�������Ȃ���΂���
'        If DBDRV_GetTBCMJ002(rsz(), sql, "") <> FUNCTION_RETURN_SUCCESS Then
'            s_cmmc001db2_sql = FUNCTION_RETURN_FAILURE
'            Exit Function
'        End If
'        If UBound(rsz) = 0 Then
'            s_cmmc001db2_sql = FUNCTION_RETURN_FAILURE
'            Exit Function
'        End If
'        all.typ_rsz(ii) = rsz(1)
'    Next ii
    


    For c0 = 1 To 2
        If c0 = 1 Then
            MaxMin = "MIN"
        Else
            MaxMin = "MAX"
        End If
        sql = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
                  " OPECOND, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
                  " SENDFLAG, SENDDATE "
        sql = sql & "From TBCMJ002 "
        
        If ADDDPPOS > 0 And ADDDPPOS < FREELENG Then
            If INGOTPOS < ADDDPPOS Then
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION < '" & ADDDPPOS & "') and "
                sql = sql & "TRANCNT = ANY (SELECT MAX(TRANCNT) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION < '" & ADDDPPOS & "'))"
            Else
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION > '" & ADDDPPOS & "') and "
                sql = sql & "TRANCNT = ANY (SELECT MAX(TRANCNT) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION > '" & ADDDPPOS & "'))"
            End If
        Else
            sql = sql & "where CRYNUM = '" & sCrynum & "' and "
            sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
            
'>>>>> �T���v���L���Ή� SETsw kubota(2006/04/27)
            '�T���v���L�̒���Top�ʒu�ABot�ʒu���擾
            'sql = sql & "where CRYNUM = '" & sCrynum & "') and "
            sql = sql & "where CRYNUM = '" & sCrynum & "'"
            sql = sql & "  and SMPLUMU = '0'"
            sql = sql & "  and (POSITION,TRANCNT) in"
            sql = sql & "      (SELECT POSITION,MAX(TRANCNT) FROM TBCMJ002 where CRYNUM = '" & sCrynum & "' group by POSITION)"
            sql = sql & ") and "
'<<<<< �T���v���L���Ή� SETsw kubota(2006/04/27)
            
            sql = sql & "TRANCNT = ANY (SELECT MAX(TRANCNT) FROM TBCMJ002 "
            sql = sql & "where CRYNUM = '" & sCrynum & "' and "
            sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
            
'>>>>> �T���v���L���Ή� SETsw kubota(2006/04/27)
            '�T���v���L�̒���Top�ʒu�ABot�ʒu���擾
            'sql = sql & "where CRYNUM = '" & sCrynum & "'))"
            sql = sql & "where CRYNUM = '" & sCrynum & "'"
            sql = sql & "  and SMPLUMU = '0'"
            sql = sql & "  and (POSITION,TRANCNT) in"
            sql = sql & "      (SELECT POSITION,MAX(TRANCNT) FROM TBCMJ002 where CRYNUM = '" & sCrynum & "' group by POSITION)"
            sql = sql & "))"
'<<<<< �T���v���L���Ή� SETsw kubota(2006/04/27)
        
        End If
        
'Debug.Print sql
        ''�f�[�^�𒊏o����
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs Is Nothing Then
            s_cmmc001db2_sql = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        ''���o���ʂ��i�[����
        recCnt = rs.RecordCount
            
        If recCnt <> 0 Then

            ReDim temp(recCnt) As typ_TBCMJ002
            For c1 = 1 To recCnt
                With temp(c1)
                    .CRYNUM = rs("CRYNUM")           ' �����ԍ�
                    .POSITION = rs("POSITION")       ' �ʒu
                    .SMPKBN = rs("SMPKBN")           ' �T���v���敪
                    .TRANCOND = rs("TRANCOND")       ' ��������
                    .TRANCNT = rs("TRANCNT")         ' ������
                    .SMPLNO = rs("SMPLNO")           ' �T���v���m��
                    .SMPLUMU = rs("SMPLUMU")         ' �T���v���L��
                    .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
                    .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
                    .hinban = rs("HINBAN")           ' �i��
                    .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
                    .factory = rs("FACTORY")         ' �H��
                    .opecond = rs("OPECOND")         ' ���Ə���
                    .GOUKI = rs("GOUKI")             ' ���@
                    .TYPE = rs("TYPE")               ' �^�C�v
                    .MEAS1 = rs("MEAS1")             ' ����l�P
                    .MEAS2 = rs("MEAS2")             ' ����l�Q
                    .MEAS3 = rs("MEAS3")             ' ����l�R
                    .MEAS4 = rs("MEAS4")             ' ����l�S
                    .MEAS5 = rs("MEAS5")             ' ����l�T
                    .EFEHS = rs("EFEHS")             ' �����ΐ�
                    .RRG = rs("RRG")                 ' �q�q�f
                    .JudgData = rs("JUDGDATA")       ' �����Ώےl
                    .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
                    .REGDATE = rs("REGDATE")         ' �o�^���t
                    .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
                    .UPDDATE = rs("UPDDATE")         ' �X�V���t
                    .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
                    .SENDDATE = rs("SENDDATE")       ' ���M���t
                End With
                rs.MoveNext
            Next
            
            c1 = 1
            If recCnt > 1 Then
                Select Case c0
                Case 1
                    If temp(2).SMPKBN = "B" Then
                        c1 = 2
                    End If
                Case 2
                    If temp(2).SMPKBN = "T" Then
                        c1 = 2
                    End If
                End Select
            End If
            
            With typ_rsz(c0)
                .CRYNUM = temp(c1).CRYNUM       ' �����ԍ�
                .POSITION = temp(c1).POSITION   ' �ʒu
                .SMPKBN = temp(c1).SMPKBN       ' �T���v���敪
                .TRANCOND = temp(c1).TRANCOND   ' ��������
                .TRANCNT = temp(c1).TRANCNT     ' ������
                .SMPLNO = temp(c1).SMPLNO       ' �T���v���m��
                .SMPLUMU = temp(c1).SMPLUMU     ' �T���v���L��
                .KRPROCCD = temp(c1).KRPROCCD   ' �Ǘ��H���R�[�h
                .PROCCODE = temp(c1).PROCCODE   ' �H���R�[�h
                .hinban = temp(c1).hinban       ' �i��
                .REVNUM = temp(c1).REVNUM       ' ���i�ԍ������ԍ�
                .factory = temp(c1).factory     ' �H��
                .opecond = temp(c1).opecond     ' ���Ə���
                .GOUKI = temp(c1).GOUKI         ' ���@
                .TYPE = temp(c1).TYPE           ' �^�C�v
                .MEAS1 = temp(c1).MEAS1         ' ����l�P
                .MEAS2 = temp(c1).MEAS2         ' ����l�Q
                .MEAS3 = temp(c1).MEAS3         ' ����l�R
                .MEAS4 = temp(c1).MEAS4         ' ����l�S
                .MEAS5 = temp(c1).MEAS5         ' ����l�T
                .EFEHS = temp(c1).EFEHS         ' �����ΐ�
                .RRG = temp(c1).RRG             ' �q�q�f
                .JudgData = temp(c1).JudgData   ' �����Ώےl
                .TSTAFFID = temp(c1).TSTAFFID   ' �o�^�Ј�ID
                .REGDATE = temp(c1).REGDATE     ' �o�^���t
                .KSTAFFID = temp(c1).KSTAFFID   ' �X�V�Ј�ID
                .UPDDATE = temp(c1).UPDDATE     ' �X�V���t
                .SENDFLAG = temp(c1).SENDFLAG   ' ���M�t���O
                .SENDDATE = temp(c1).SENDDATE   ' ���M���t
            End With
        Else
            With typ_rsz(c0)
                .CRYNUM = ""       ' �����ԍ�
            End With
        End If
        rs.Close
    Next
    s_cmmc001db2_sql = FUNCTION_RETURN_SUCCESS
End Function
