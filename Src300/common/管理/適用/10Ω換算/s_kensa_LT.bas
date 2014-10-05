Attribute VB_Name = "s_kensa_LT"
' ���C�t�^�C������_���i�V�f�[�^�͂P�O�_�Œ�j
Public Const SS_SOKUETI_TENSU = 10
' ���C�t�^�C������_���i���f�[�^�͂T�_�Œ�j
Public Const SS_SOKUETI_TENSU_OLD = 5
' �z�񏉊����l
Public Const DEF_PARAM_VALUE_LT = -1

''�@LT�֐��߂�l��`
Public Enum FUNC_RET_LT         ''�֐��̖߂�l
    FUNC_RET_LT_NOSAMPLE = FUNCTION_RETURN_FAILURE + 2  '' �T���v����
    FUNC_RET_LT_NODATA = FUNCTION_RETURN_SUCCESS + 1    '' LT�f�[�^�Ȃ�
    FUNC_RET_LT_SUCCESS = FUNCTION_RETURN_SUCCESS       '' ����
    FUNC_RET_LT_FAILURE = FUNCTION_RETURN_FAILURE       '' �ُ�
    FUNC_RET_LT_CALCFAIL = FUNCTION_RETURN_FAILURE - 1  '' ���C�t�^�C�����茋�ʂ̎Z�o�G���[
End Enum

'' ���C�t�^�C������l
Public Type typ_LTMEAS
    CRYNUMCS As String * 12         '�u���b�NID
    XTALCS As String * 12           '�����ԍ�
    HINBCS As String * 8            '�i��
    REVNUMCS As Integer             '���i�ԍ������ԍ�
    FACTORYCS As String * 1         '�H��
    OPECS As String * 1             '���Ə���
    MEAS1 As Integer                '����l1
    MEAS2 As Integer                '����l2
    MEAS3 As Integer                '����l3
    MEAS4 As Integer                '����l4
    MEAS5 As Integer                '����l5
    MEAS6 As Integer                '����l6
    MEAS7 As Integer                '����l7
    MEAS8 As Integer                '����l8
    MEAS9 As Integer                '����l9
    MEAS10 As Integer               '����l10
    LTSPIFLG As String * 1          '����ʒu����t���O
End Type



'�T�v      :���C�t�^�C�����茋�ʂ��Z�o����
'���Ұ�    :�ϐ���        ,IO  ,�^              ,����
'          :iCalcMeas     ,O   ,Integer         ,���C�t�^�C�����茋��
'          :sCrynum       ,I   ,String          ,���������ԍ�
'          :iSmplIDLt     ,I   ,Long            ,�T���v��ID(LT)     Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O   ,FUNC_RET_LT     ,LT�֐��߂�l��`�̒ʂ�
'����      :
'����      :05/11/24 (SET)M.makino
Public Function KNS_GetLtCalcMeas(ByRef iCalcMeas As Integer, _
                                  sCryNum As String, iSmplIDLt As Long) As FUNC_RET_LT
    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tLtMeas     As typ_LTMEAS
    Dim tHinInf     As tFullHinban
    Dim tSXLData    As typ_TBCME019
    Dim iLtParam()  As Integer
    Dim iOldFlg     As Integer
    Dim sHsxLtspi   As String
    Dim sXtal        As String

    KNS_GetLtCalcMeas = FUNC_RET_LT_FAILURE

    '' ���C�t�^�C���f�[�^�̎擾
    iRet = DBDRV_KNS_GetLTMeas(tLtMeas, sCryNum, iSmplIDLt)
    If iRet <> FUNC_RET_LT_SUCCESS Then
        KNS_GetLtCalcMeas = iRet
        Exit Function
    End If

    '' �V�`���̃f�[�^�͑���ʒu���擾����
    If (Trim(tLtMeas.LTSPIFLG) <> "") Then
        iOldFlg = 0
    
        '' Z�i�ԁAG�i�Ԃ̏ꍇ�͂˂炢�i�Ԃɒu��������
        If (Trim(tLtMeas.HINBCS) = "Z") Or (Trim(tLtMeas.HINBCS) = "G") Then
            iRet = DBDRV_KNS_GetNeraiZuban(tHinInf, tLtMeas.XTALCS)
            If iRet <> FUNC_RET_LT_SUCCESS Then Exit Function
        Else
            tHinInf.hinban = tLtMeas.HINBCS
            tHinInf.mnorevno = tLtMeas.REVNUMCS
            tHinInf.factory = tLtMeas.FACTORYCS
            tHinInf.opecond = tLtMeas.OPECS
        End If

        '' ���i�d�lSXL�f�[�^�Q�擾(TBCME019)
        iRet = DBDRV_KNS_GetSXLData(tSXLData, tHinInf)
        If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
        
        sHsxLtspi = tSXLData.HSXLTSPI
    '' ���`���͑���ʒu���擾����K�v���Ȃ�
    Else
        iOldFlg = 1
        sHsxLtspi = ""
    End If
    
    ReDim iLtParam(KNS_GetMeasureNum_LT(iOldFlg) - 1)
    
    ''������
    For Index = 0 To UBound(iLtParam)
        iLtParam(Index) = DEF_PARAM_VALUE_LT
    Next Index
    
    iLtParam(0) = tLtMeas.MEAS1
    iLtParam(1) = tLtMeas.MEAS2
    iLtParam(2) = tLtMeas.MEAS3
    iLtParam(3) = tLtMeas.MEAS4
    iLtParam(4) = tLtMeas.MEAS5
    If iOldFlg = 0 Then
        iLtParam(5) = tLtMeas.MEAS6
        iLtParam(6) = tLtMeas.MEAS7
        iLtParam(7) = tLtMeas.MEAS8
        iLtParam(8) = tLtMeas.MEAS9
        iLtParam(9) = tLtMeas.MEAS10
    End If
    
    '' ���茋�ʂ̎Z�o
    iRet = KNS_CalculateMeasResult_LT(iCalcMeas, iLtParam(), sHsxLtspi, iOldFlg)
    If iRet <> FUNCTION_RETURN_SUCCESS Then
        KNS_GetLtCalcMeas = FUNC_RET_LT_CALCFAIL
        Exit Function
    End If

    KNS_GetLtCalcMeas = FUNC_RET_LT_SUCCESS
    
End Function

'�T�v      :�擾�������i�d�lSXL�f�[�^��葪��_�����擾����i���C�t�^�C�����сj
'���Ұ�    :�ϐ���        ,IO ,�^             ,����
'          :iOldFlg       ,I  ,Integer        ,���f�[�^�t���O   (���f�[�^[5�_����]��1��ݒ肷��)
'          :�߂�l        ,O  ,Integer        ,����_��
'����      :
' Mod Start 2005/11/14 M.Makino
'Private Function KNS_GetMeasureNum_LT(tHinInf As tFullHinban) As Integer
Public Function KNS_GetMeasureNum_LT(iOldFlg As Integer) As Integer
' Mod End   2005/11/14 M.Makino

    Dim Index   As Integer
    Dim strMN   As String
    Dim iNum    As Integer

' Mod Start 2005/11/14 M.Makino
'    '' ���C�t�^�C���͂T�_�Œ�
'    iNum = 5
    If iOldFlg = 1 Then
        iNum = SS_SOKUETI_TENSU_OLD
    Else
        iNum = SS_SOKUETI_TENSU
    End If
' Mod End   2005/11/14 M.Makino

    KNS_GetMeasureNum_LT = iNum

End Function

'�T�v      :���茋�ʂ��v�Z����i���C�t�^�C�����сj
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :iResult       ,O   ,Integer   ,�v�Z����
'          :iParam()      ,I   ,Integer   ,����l�z��
'          :sHsxLtspi     ,I   ,String    ,����ʒu         (�V�f�[�^[10�_����]��3,5,A�̂ǂꂩ��ݒ肷��)
'          :iOldFlg       ,I   ,Integer   ,���f�[�^�t���O   (���f�[�^[5�_����]��1��ݒ肷��)
'          :�߂�l        ,O   ,Integer   ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :2005/11/07 �q�� �ύX�@10�_����Ή�
Public Function KNS_CalculateMeasResult_LT(iResult As Integer, iParam() As Integer, _
                    sHsxLtspi As String, iOldFlg As Integer) As Integer
    Dim Index   As Integer
    Dim iAve    As Integer

    On Error GoTo Err
    KNS_CalculateMeasResult_LT = FUNCTION_RETURN_FAILURE

' Mod Start 2005/11/14 M.Makino
'    '' �p�����[�^���̓`�F�b�N
'    For Index = 0 To UBound(iParam)
'        If iParam(Index) = DEF_PARAM_VALUE_LT Then
'            Exit Function
'        End If
'    Next Index
'
'    ''�R�C�S�C�T�_�̑���_��AVE�����߂�
'    iAve = RoundDown((iParam(2) + iParam(3) + iParam(4)) / 3#, 0)
'
'    '' ����_�Q��AVE�l���r�A�l�̏��������𑪒茋�ʂƂ���
'    If iAve < iParam(1) Then
'        iResult = iAve
'    Else
'        iResult = iParam(1)
'    End If

    
    '' ���f�[�^�̏ꍇ�i�T�_����j
    If iOldFlg = 1 Then
        '' �p�����[�^���̓`�F�b�N
        For Index = 0 To KNS_GetMeasureNum_LT(iOldFlg) - 1
            If iParam(Index) = DEF_PARAM_VALUE_LT Then
                Exit Function
            End If
        Next Index
        ''�R�C�S�C�T�_�̑���_��AVE�����߂�
        iAve = RoundDown((iParam(2) + iParam(3) + iParam(4)) / 3#, 0)

        '' ����_�Q��AVE�l���r�A�l�̏��������𑪒茋�ʂƂ���
        If iAve < iParam(1) Then
            iResult = iAve
        Else
            iResult = iParam(1)
        End If

    '' �V�f�[�^�̏ꍇ�i�P�O�_����j
    Else
        '' �p�����[�^���̓`�F�b�N
        For Index = 0 To KNS_GetMeasureNum_LT(iOldFlg) - 1
            If iParam(Index) = DEF_PARAM_VALUE_LT Then
                Exit Function
            End If
        Next Index

        ''' [A:Ce,Inside3mm]�̏ꍇ
        If Trim(sHsxLtspi) = "3" Then
            ''�W�C�X�C�P�O�_�̑���_��AVE�����߂�
            iAve = RoundDown((iParam(7) + iParam(8) + iParam(9)) / 3#, 0)

        ''' [A:Ce,Inside5mm]�̏ꍇ
        ElseIf Trim(sHsxLtspi) = "5" Then
            ''�T�C�U�C�V�_�̑���_��AVE�����߂�
            iAve = RoundDown((iParam(4) + iParam(5) + iParam(6)) / 3#, 0)

        ''' [A:Ce,Inside10mm]�̏ꍇ
        ElseIf Trim(sHsxLtspi) = "A" Then
            ''�Q�C�R�C�S�_�̑���_��AVE�����߂�
            iAve = RoundDown((iParam(1) + iParam(2) + iParam(3)) / 3#, 0)

' Mod Start 2005/12/13 M.Makino
'        ''' ���̑��̏ꍇ�̓G���[
'        Else
'            Exit Function

        ''' ���̑��̏ꍇ��[A:Ce,Inside10mm]�̎d�l�Ƃ���
        Else
            ''�Q�C�R�C�S�_�̑���_��AVE�����߂�
            iAve = RoundDown((iParam(1) + iParam(2) + iParam(3)) / 3#, 0)
' Mod End   2005/12/13 M.Makino

        End If
    
        '' ����_�P��AVE�l���r�A�l�̏��������𑪒茋�ʂƂ���
        If iAve < iParam(0) Then
            iResult = iAve
        Else
            iResult = iParam(0)
        End If
    End If
' Mod End   2005/11/14 M.Makino

    KNS_CalculateMeasResult_LT = FUNCTION_RETURN_SUCCESS
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
End Function

'�T�v      :�V�T���v���Ǘ��A���C�t�^�C���e�[�u�����烉�C�t�^�C������l�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records       ,O  ,typ_LTMEAS   ,���o���R�[�h
'          :sCrynum       ,I  ,String       ,���������ԍ�
'          :iSmplIDLt     ,I  ,Long         ,�T���v���ԍ�   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O  ,FUNC_RET_LT  ,���o�̐���
'����      :
'����      :2005/11/24 Create (SET)M.Makino
Public Function DBDRV_KNS_GetLTMeas(records As typ_LTMEAS, sCryNum As String, _
                                    iSmplIDLt As Long) As FUNC_RET_LT
    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet

    DBDRV_KNS_GetLTMeas = FUNC_RET_LT_FAILURE

    ''SQL��g�ݗ��Ă�
    sql = sql & "select nvl(T1.CRYNUMCS, '')  CRYNUMCS" '�u���b�NID
    sql = sql & ", nvl(T1.XTALCS, '')   XTALCS"         '�����ԍ�
    sql = sql & ", nvl(T1.HINBCS, '')    HINBCS"        '�i��
    sql = sql & ", nvl(T1.REVNUMCS, 0)   REVNUMCS"      '���i�ԍ������ԍ�
    sql = sql & ", nvl(T1.FACTORYCS, '') FACTORYCS"     '�H��
    sql = sql & ", nvl(T1.OPECS, '')     OPECS"         '���Ə���
    sql = sql & ", nvl(T2.MEAS1, -1) MEAS1"             '����l�P
    sql = sql & ", nvl(T2.MEAS2, -1) MEAS2"             '����l�Q
    sql = sql & ", nvl(T2.MEAS3, -1) MEAS3"             '����l�R
    sql = sql & ", nvl(T2.MEAS4, -1) MEAS4"             '����l�S
    sql = sql & ", nvl(T2.MEAS5, -1) MEAS5"             '����l�T
    sql = sql & ", nvl(T2.MEAS6, -1) MEAS6"             '����l�U
    sql = sql & ", nvl(T2.MEAS7, -1) MEAS7"             '����l�V
    sql = sql & ", nvl(T2.MEAS8, -1) MEAS8"             '����l�W
    sql = sql & ", nvl(T2.MEAS9, -1) MEAS9"             '����l�X
    sql = sql & ", nvl(T2.MEAS10, -1) MEAS10"           '����l�P�O
    sql = sql & ", LTSPIFLG"                            '����ʒu����t���O
    sql = sql & ", SMPLUMU"                             '�T���v���L��
    sql = sql & " from XSDCS T1, TBCMJ007 T2"
    sql = sql & " where T1.CRYNUMCS = '" & sCryNum & "'"
    sql = sql & " and T1.CRYSMPLIDTCS = " & iSmplIDLt
    sql = sql & " and T1.XTALCS = T2.CRYNUM"
'    sql = sql & " and T1.INPOSCS = T2.POSITION"
    sql = sql & " and T1.TBKBNCS = T2.SMPKBN"
    sql = sql & " and T2.TRANCOND = '0'"
    sql = sql & " and T2.TRANCNT = "
    sql = sql & "("
    sql = sql & " select max(T2.TRANCNT)"
    sql = sql & " from XSDCS T1, TBCMJ007 T2"
    sql = sql & " where T1.CRYNUMCS = '" & sCryNum & "'"
    sql = sql & " and T1.CRYSMPLIDTCS = " & iSmplIDLt
    sql = sql & " and T1.XTALCS = T2.CRYNUM"
'    sql = sql & " and T1.INPOSCS = T2.POSITION"
    sql = sql & " and T1.TBKBNCS = T2.SMPKBN"
    sql = sql & " and T2.TRANCOND = '0'"
    sql = sql & ")"

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    If rs Is Nothing Then
        Exit Function
    End If

    ' �Y������f�[�^�������ꍇ
    If rs.EOF Then
        DBDRV_KNS_GetLTMeas = FUNC_RET_LT_NODATA
        Exit Function
    End If
    
    '' �T���v�����̏ꍇ�̓T���v�����R�[�h��Ԃ�
    If rs.Fields("SMPLUMU").Value <> "0" Then
        DBDRV_KNS_GetLTMeas = FUNC_RET_LT_NOSAMPLE
        Exit Function
    End If


    With records
        .CRYNUMCS = rs.Fields("CRYNUMCS").Value   '�u���b�NID
        .XTALCS = rs.Fields("XTALCS").Value       '�����ԍ�
        .HINBCS = rs.Fields("HINBCS").Value       '�i��
        .REVNUMCS = rs.Fields("REVNUMCS").Value   '���i�ԍ������ԍ�
        .FACTORYCS = rs.Fields("FACTORYCS").Value '�H��
        .OPECS = rs.Fields("OPECS").Value         '���Ə���
        .MEAS1 = rs.Fields("MEAS1").Value         '����l�P
        .MEAS2 = rs.Fields("MEAS2").Value         '����l�Q
        .MEAS3 = rs.Fields("MEAS3").Value         '����l�R
        .MEAS4 = rs.Fields("MEAS4").Value         '����l�S
        .MEAS5 = rs.Fields("MEAS5").Value         '����l�T
        .MEAS6 = rs.Fields("MEAS6").Value         '����l�U
        .MEAS7 = rs.Fields("MEAS7").Value         '����l�V
        .MEAS8 = rs.Fields("MEAS8").Value         '����l�W
        .MEAS9 = rs.Fields("MEAS9").Value         '����l�X
        .MEAS10 = rs.Fields("MEAS10").Value       '����l�P�O
        .LTSPIFLG = Trim(CStr(NulltoStr(rs.Fields("LTSPIFLG").Value)))  '����ʒu����t���O
    End With

    rs.Close

    DBDRV_KNS_GetLTMeas = FUNC_RET_LT_SUCCESS
End Function

'�T�v      :�˂炢�i�Ԃ��擾����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,tFullHinban  ,���o���R�[�h
'          :sXtal         ,I  ,String       ,�����ԍ�
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2005/11/24 Create (SET)M.Makino
Public Function DBDRV_KNS_GetNeraiZuban(records As tFullHinban, sXtal As String) As FUNCTION_RETURN
    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet

    DBDRV_KNS_GetNeraiZuban = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�
    sql = sql & "select RPHINBAN, RPREVNUM, RPFACT, RPOPCOND"
    sql = sql & " from TBCME037"
    sql = sql & " where CRYNUM = '" & sXtal & "'"

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    If rs Is Nothing Then
        Exit Function
    End If

    ' �Y������f�[�^�������ꍇ
    If rs.EOF Then Exit Function

    With records
        .hinban = Trim(CStr(NulltoStr(rs.Fields("RPHINBAN").Value)))
        .mnorevno = Trim(CStr(NulltoStr(rs.Fields("RPREVNUM").Value)))
        .factory = Trim(CStr(NulltoStr(rs.Fields("RPFACT").Value)))
        .opecond = Trim(CStr(NulltoStr(rs.Fields("RPOPCOND").Value)))
    End With

    rs.Close

    DBDRV_KNS_GetNeraiZuban = FUNCTION_RETURN_SUCCESS
End Function


'�T�v      :�i�ԓ����瑪��_���擾����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :siyou         ,O  ,typ_TBCME019 ,����ʒu�i�[�\����
'          :tHinInf       ,I  ,tFullHinban  ,�i�ԓ�
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2005/11/24 Create (SET)T.Takasaki
Public Function DBDRV_KNS_GetSXLData(siyou As typ_TBCME019, tHinInf As tFullHinban) As FUNCTION_RETURN
    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet

    DBDRV_KNS_GetSXLData = FUNCTION_RETURN_FAILURE
    
    ''SQL��g�ݗ��Ă�
    sql = sql & "select HSXLTSPI "
    sql = sql & "from TBCME019 "
    sql = sql & "where HINBAN = '" & tHinInf.hinban & "' "
    sql = sql & "and MNOREVNO = '" & tHinInf.mnorevno & "' "
    sql = sql & "and FACTORY = '" & tHinInf.factory & "' "
    sql = sql & "and OPECOND = '" & tHinInf.opecond & "' "
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    If rs Is Nothing Then
        Exit Function
    End If

    ' �Y������f�[�^�������ꍇ
    If rs.EOF Then Exit Function

    ''�f�[�^�i�[
    With siyou
        .HSXLTSPI = Trim(CStr(NulltoStr(rs.Fields("HSXLTSPI").Value)))
    End With
    
    rs.Close
    
    DBDRV_KNS_GetSXLData = FUNCTION_RETURN_SUCCESS

End Function

'�T�v      :������R�̎擾�ƂP�O�����Z�l�̎Z�o���s��
'���Ұ�    :�ϐ���        ,IO  ,�^           ,����
'          :tblCrySmpMan  ,I   ,typ_XSDCS    ,�T���v��ID
'          :sKekka        ,I   ,String       ,���茋��
'          :sIncval       ,I   ,String       ,�X��
'          :sCutval       ,I   ,String       ,�ؕ�
'          :sSetval       ,I   ,String       ,�ݒ�l
'          :sJiteiko      ,I   ,String       ,������R
'          :sKansanchi    ,I   ,String       ,�P�O�����Z�l
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,���o�̐���
'����      :
'���l      : �P�O�����Z���̎Z�o���@
'               �`�����C�t�^�C�����茋��
'               �a��������R
'               �b���ؕ� [����=XXX.XX]
'               �c���X�� [����=XXX.XX]
'               �f���ݒ�l [����=XXX.XX]
'               �d�����_�lLT���c�~�a�{�b
'               �e�������ʐ���l���P�^((1�^�`)�\(1�^�d))
'               �P�O�����Z�l���P�^((�P�^�f)�{(�P�^�e)) [����=XXXX]
'����      :�V�K 2005/11/14 M.Makino
Public Function GetKansanchi(tblCrySmpMan As typ_XSDCS, sKekka As String, sIncVal As String, _
        sCutVal As String, sSetVal As String, sJiteiko As String, sKansanchi As String) As Integer
    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim RironchiLT As Double    ' ���_�lLT
    Dim Osenryo As Double       ' �����ʐ���l

    GetKansanchi = FUNCTION_RETURN_FAILURE

    ' SQL���쐬
    sql = ""
    sql = sql & "SELECT MEAS1"
    sql = sql & " FROM  TBCMJ002"
    sql = sql & " WHERE CRYNUM='" & tblCrySmpMan.XTALCS & "'"
    sql = sql & " AND   POSITION=" & tblCrySmpMan.INPOSCS
    sql = sql & " AND   SMPKBN='" & tblCrySmpMan.SMPKBNCS & "'"
    sql = sql & " AND   TRANCOND='0'"
    sql = sql & " ORDER BY TRANCNT DESC"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    If rs Is Nothing Then
        Exit Function
    End If

    If rs.EOF Then
        ' �Y������f�[�^�������ꍇ����͋󕶎�
        sJiteiko = ""
    Else
        sJiteiko = Trim(CStr(NulltoStr(rs.Fields("MEAS1").Value)))
    End If

    ' �P�O�����Z�l�̌v�Z
    If sKekka <> "" And sIncVal <> "" And sCutVal <> "" And _
       sSetVal <> "" And sJiteiko <> "" Then

        '0�̏��Z�΍�
        On Error GoTo ERROR_CALC

        '�P�O�����Z�l���Z�o
        RironchiLT = CDbl(sIncVal) * CDbl(sJiteiko) + CDbl(sCutVal)
        Osenryo = 1 / ((1 / CInt(sKekka)) - (1 / RironchiLT))
        sKansanchi = CStr(Round(1 / ((1 / CDbl(sSetVal)) + (1 / Osenryo)), 0))
    Else
        sKansanchi = ""
    End If
    
    GetKansanchi = FUNCTION_RETURN_SUCCESS
    Exit Function

ERROR_CALC:
    sKansanchi = ""
    GetKansanchi = FUNCTION_RETURN_SUCCESS
End Function

'
' �󕶎���i""�j�ɑ΂��āwnull�x��Ԃ��C���̑��̕�����͉��������ɕԂ�
'
'����      :2005/11/14�ǉ��@�q��
Public Function LZeroToNull(ByVal sTmp As String) As String
    If "" = sTmp Then
        LZeroToNull = "null"
    Else
        LZeroToNull = sTmp
    End If
End Function



