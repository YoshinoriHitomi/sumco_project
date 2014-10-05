Attribute VB_Name = "s_cmzcDope"
Type DopeData
    DopKind As String * 7       '�h�[�v����(CODE)
    IonDensity As Double        '�C�I���Z�x(OutPut)
    CoreCoeff As Integer        '�␳�W��
End Type
Type CodeList
    RCLSCODE As String * 3      '������������(�ϋ敪CODE)
    WEIGHT As Long              '�d��
    IonDensity As Double        '�C�I���Z�x(OutPut)
End Type
Type type_DopeCal
    NTYPE As String * 1         ' �_���^�C�v
    res As Double               ' �_����R
    CHARGE As Long              ' �`���[�W��
    Dope As DopeData
    CryList() As CodeList
    FixNumA As Double           '�萔A
    FixNumB As Double           '�萔B
End Type

'*ADD* TCS)K.Kunori 2004.11.29 START >>>
'���h�[�p���g�ʃf�[�^
Public Type typ_DpData
    DPWEIGHT    As Double           ' �K�v�h�[�p���g��
    ZanDp       As Double           ' �c�t�h�[�p���g��
    AddDp       As Double           ' �ǉ��h�[�p���g��
    InpDp       As Double           ' ���̓h�[�p���g��
    '*ADD* TCS)K.Kunori 2004.10.14
    PutDp       As Double           ' �����h�[�p���g��
End Type

Public sDpData As typ_DpData

Public strRes1  As String           '���������Z�x���v(�~f�~��)
Public strRes2  As String           '�ް���Ċ�ߗ�(f)
'*ADD* TCS)K.Kunori 2004.11.29 END <<<
'*ADD* �␳�W���ǉ� TCS)K.Kunori 2004.12.16
Public strRes3  As String           '�␳�W��(��)

Option Explicit

Public Function Log10(x)
   Log10 = Log(x) / Log(10#)
End Function

Public Function Exp10(x)
   Exp10 = Exp(x) ^ Log(10#)
End Function

Public Function DopeCalculation(CC As type_DopeCal) As Double
    Dim Ion As Double
    Dim temp As Double
    Dim c0 As Integer
    
    DopeCalculation = -9999
    '�h�[�p���g�ʌv�Z�p�f�[�^���W�B
    If GetDopeCalData(CC) = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If
    
    On Error GoTo Err
    temp = 0
    For c0 = 1 To UBound(CC.CryList())
        temp = temp + CC.CryList(c0).IonDensity * CC.CryList(c0).WEIGHT * 10 ^ 14
    Next
    
    Ion = Exp10((Log10(CC.res) - CC.FixNumB) / CC.FixNumA) / 2.34
    '�v�Z���̕ύX   2008/09/08 Kameda
    'DopeCalculation = ((Ion * CC.CHARGE - temp) / (CC.Dope.IonDensity * 10 ^ 14)) / CC.Dope.CoreCoeff
    DopeCalculation = ((Ion * CC.CHARGE - temp) / (CC.Dope.IonDensity * 10 ^ 14)) * CC.Dope.CoreCoeff
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    DopeCalculation = -9999
End Function

Public Function GetDopeCalData(CC As type_DopeCal) As FUNCTION_RETURN
    Dim sql As String       'SQL�S��
    Dim sql1 As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim c0 As Integer
    Dim c1 As Integer
    Dim MaxRec As Integer
    Dim MaxRec1 As Integer
    Dim temp() As CodeList
    Dim sFactor As Single
    Dim sHenseki As Single
    
    GetDopeCalData = FUNCTION_RETURN_FAILURE
    
    MaxRec = UBound(CC.CryList())
    
    '���������̃C�I���Z�x�����߂�B
    If MaxRec > 0 Then
        sql1 = ""
        For c0 = 1 To MaxRec
            sql1 = sql1 & "'" & CC.CryList(c0).RCLSCODE & "',"
        Next
        sql1 = Left(sql1, Len(sql1) - 1)
        
        sql = "select RCLSCODE, IonDensity from TBCMB007 where "
        sql = sql & "RCLSCODE in (" & sql1 & ")"
    
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        If rs.RecordCount = 0 Then
            Exit Function
        End If
        
        MaxRec1 = rs.RecordCount
        ReDim temp(1 To MaxRec1) As CodeList
        For c0 = 1 To MaxRec1
            temp(c0).RCLSCODE = rs("RCLSCODE")
            temp(c0).IonDensity = rs("IonDensity")
            rs.MoveNext
        Next
        rs.Close
        For c0 = 1 To MaxRec
            For c1 = 1 To MaxRec1
                If CC.CryList(c0).RCLSCODE = temp(c1).RCLSCODE Then
                    CC.CryList(c0).IonDensity = temp(c1).IonDensity
                End If
            Next
        Next
    End If
    
    '�w��h�[�v�̃C�I���Z�x�����߂�B 2011/05/31 kameda
    'sql = "select IonDensity,CoreCoeff from TBCMB009 where "
    'sql = sql & "DopKind = '" & CC.Dope.DopKind & "' "

    'Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    'If rs.RecordCount = 0 Then
    '    Exit Function
    'End If
    
    'CC.Dope.IonDensity = rs("IonDensity")
    CC.Dope.IonDensity = (Getrnoudo(CC.Dope.DopKind, CC.res)) * 100000
    'CC.Dope.CoreCoeff = rs("CoreCoeff")
    If GetFactor(CC.NTYPE, sFactor, sHenseki) = False Then
        Call MsgOut(0, "�t�@�N�^�[�E�ΐ͌W���擾�G���[", ERR_DISP)
        Exit Function
    End If
    CC.Dope.CoreCoeff = sFactor
    'rs.Close
    
    '�萔A�AB�����߂�B
    sql = "select FIXNUMA, FIZNUMB from TBCMB010 where "
    sql = sql & "TYPE = '" & CC.NTYPE & "' "
    sql = sql & "and RESFROM >= " & CC.res & " "
    sql = sql & "and RESTO <= " & CC.res

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        Exit Function
    End If
    CC.FixNumA = rs("FixNumA")
    CC.FixNumB = rs("FizNumB")
    rs.Close
    
    GetDopeCalData = FUNCTION_RETURN_SUCCESS
End Function

'*ADD* �v���ް�ߗʌv�Z���@�ύX�Ή� ��Ӽޭ�ٓ��� TCS)K.Kunori 2004.11.29 START >>>
'=============================================================================
'@(f)       DopeCalculationPart
'
'�@�\       �v���ް�ߗʂ����߂�
'
'�߂�l     �v���ް�ߗʌv�Z����(dpVal as Double)
'
'����       �_������                    strprmType As String
'           �_����R                    dblprmRes As Double
'           ����ޗ�                     dblprmChrg As Double
'           �w���ް���ė�               dblprmSijiDp As Double
'           ����އ�                     strprmChrgNo As String(txtChargeNo.Text)
'
'�@�\�T�v   //�v���ް�ߗʌv�Z��//
'�@�@�@�@�@ �v���ް�ߗ� = �ް���ėʎw��(g) - �ް���Ċ�ߗ�(f) �~ �����ް���ėʍ��v(mg)/1000 - �ް���ėʎ���(g)
'
'���l       �v�Z�ȏ㔭�����́A'-9999'��Ԃ�
'=============================================================================
Public Function DopeCalculationPart(strprmType As String, _
                                    dblprmRes As Double, _
                                    dblprmChrg As Double, _
                                    dblprmSijiDp As Double, _
                                    strprmChrgNo As String) As Variant
    
    Dim strType         As String       '�_������
    Dim dblRes          As Double       '�_����R
    Dim dblChrg         As Double       '����ޗ�
    Dim dblTgResBtm     As Double       '�_����R(����)
    Dim dblTgResTop     As Double       '�_����R(���)
    Dim dblData(0 To 2) As Double       'EE,EZ,EC
    Dim varmo           As Variant      'mo
    Dim varDilTmp       As Variant      '�ް���Ċ�ߗ��Z�o�p�ϐ�(mo*W*��)
    Dim varDilution     As Variant      '�ް���Ċ�ߗ�
    Dim varCoefficient  As Variant      '���W��
    Dim strErrData      As String       '�װ�ް�
    '*ADD* TCS)K.Kunori 2004.12.16
    Dim dblSupplCoefficient As Double   '�␳�W��
    
    DopeCalculationPart = -9999
    
    On Error GoTo ErrHand
    
    '������������������������������
    '  �K�v�ް���ė�(�l)�Z�o����
    '������������������������������
    
    '���K�v�ް���ėʌv�Z����
    '�v���ް���ė� = �ް���ėʎw�� - �ް���Ċ�ߗ�(f) �~ ���������ܗL�ް���ėʍ��v - �ް���ėʎ���
    '>>> �l = dblprmSijiDp - f * sDpData.PutDp - sDpData.InpDp
    '�ް���Ċ�ߗ� = (mo * ����ޗ� * ���W��) / �ް���ėʎ���
    '>>> f = (mo * dblChrg * ��) / sDpData.InpDp
    '>>> mo = ((EC * log10(dblRes)) ^ 2) - (EE * log10(dblRes)) - EZ)
    
    '///۰�ٕϐ��Ɋi�[
    strType = strprmType    '�_������
    dblRes = dblprmRes      '�_����R(��)
    dblChrg = dblprmChrg    '����ޗ�
    
    '+++++++++++++++
    '  mo�Z�o����
    '+++++++++++++++
    
    '------------------------------
    '  �_����R���������擾����
    '------------------------------
    '///�_����R(����)�ƃς��r���āA���ƂȂ����(�_����R��������)������
    '///�����F�_������
    '   �@�@�@�_����R(��)
    '   �@�@�@�_����R(����)
    '   �@�@�@�_����R(���)
    If GetResData(strType, dblRes, dblTgResBtm, dblTgResTop) = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If
    
    '--------------
    '  mo�Z�o����
    '--------------
    '���_����R(����) <> 0 ���� �_����R(���) <> 0 �̏ꍇ
    If dblTgResBtm <> 0 Or dblTgResTop <> 0 Then
        
        '///�������ɂ��Amo�Z�o�p�ް��擾(EC,EE,EZ)
        '///�����F�_������
        '   �@�@�@�_����R(����)
        '   �@�@�@�_����R(���)
        '   �@�@�@EE,EZ,EC�i�[�p�ϐ�
        If GetmoCalData(strType, dblTgResBtm, dblTgResTop, dblData()) = FUNCTION_RETURN_FAILURE Then
            Exit Function
        End If
        
        '///mo�Z�o ��Log�ɂP�O��������
        varmo = 10 ^ ((dblData(2) * (Log10(dblRes)) ^ 2) - (dblData(0) * (Log10(dblRes))) - dblData(1))
        
    '���ς�10.0�ȏ�̏ꍇ
    Else
        '///mo�Z�o
        varmo = 0.501 * (dblRes ^ -1.0185)
    End If
        
    '++++++++++++++++++++++++++++
    '  �ް���Ċ�ߗ�(f)�Z�o����
    '++++++++++++++++++++++++++++
    
    '------------------
    '  ���W���擾����
    '------------------
    '///�������ɂ��A���W�����擾
    '///�����F�_������
    '   �@�@�@��
    '   �@�@�@���W��
    If GetCoefficientData(strType, varCoefficient, strprmChrgNo) = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If
    
    '----------------------------
    '  �ް���Ċ�ߗ�(f)�Z�o����
    '----------------------------
    '///mo*W*���Z�o ������ޗʂ�۸��тɊ��Z����ׂ�1000�Ŋ���
    varDilTmp = varmo * (dblChrg / 1000) * varCoefficient
    '��(mo*W*��)���O�ȊO�̏ꍇ
    If varDilTmp <> 0 Then
        '///�ް���Ċ�ߗ�(f)�Z�o ���w����ް���ėʂ��ظ��тɊ��Z����ׂ�1000��������
        '*CHG* �w����ް���ėʂ��g�p����悤�C�� TCS)K.Kunori 2004.11.24
'''        varDilution = (sDpData.InpDp * 1000) / varDilTmp
        varDilution = (dblprmSijiDp * 1000) / varDilTmp
    '��(mo*W*��)���O�̏ꍇ
    Else
        '///�ް���Ċ�ߗ����O�Ƃ���
        varDilution = 0
    End If
    
    '*ADD* �␳�W���l�擾�����ǉ� TCS)K.Kunori 2004.12.16 START >>>
    '-------------------
    '  �␳�W���l�擾��
    '-------------------
    '///�ް���Ċ�ߗ�(f)�ɂ��A�␳�W���l���擾
    '///�����F�ް���Ċ�ߗ�(f)
'    If GetSupplCoefficientData(CDbl(varDilution), dblSupplCoefficient) = FUNCTION_RETURN_FAILURE Then  '�_����R�ʂɕ␳�W����ύX 2007/03/05 SETsw kubota
    If GetSupplCoefficientData(CDbl(varDilution), dblSupplCoefficient, strType, dblprmRes) = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If
    '*ADD* �␳�W���l�擾�����ǉ� TCS)K.Kunori 2004.12.16 END <<<
    
    '+++++++++++++++++++++++++
    '  �v���ް���ėʎZ�o����
    '+++++++++++++++++++++++++
    '///�v���ް���ė�(g)�Z�o
    '*CHG* �␳�W���ǉ� TCS)K.Kunori 2004.12.16
'''    DopeCalculationPart = CDec(dblprmSijiDp - CDec(varDilution) * sDpData.PutDp - sDpData.InpDp)
    DopeCalculationPart = CDec(dblprmSijiDp - _
                               CDec(varDilution) * dblSupplCoefficient * sDpData.PutDp - _
                               sDpData.InpDp)
    
    '---------------------------------------------------------
    '  ���������Z�x���v(�~f�~��)���ް���Ċ�ߗ�(f)�ޔ�����(�\���p)
    '---------------------------------------------------------
    '///���������Z�x���v(�~f�~��)
    '*CHG* �␳�W���ǉ��̈׌v�Z���ύX TCS)K.Kunori 2004.12.16
'''    strRes1 = CStr(CDec(varDilution) * sDpData.PutDp)
    strRes1 = CStr(CDec(varDilution) * dblSupplCoefficient * sDpData.PutDp)
    '///�ް���Ċ�ߗ�(f)
    strRes2 = CStr(CDec(varDilution))
    '*ADD* �␳�W���ǉ� TCS)K.Kunori 2004.12.16
    strRes3 = CStr(dblSupplCoefficient)
    
    '��ýėp��
    Debug.Print "���߁F" & strType
    Debug.Print "����ޗʁF" & CStr(CDec(dblChrg))
    Debug.Print "�ρF" & CStr(dblRes)
    Debug.Print "EE�F" & dblData(0)
    Debug.Print "EZ�F" & dblData(1)
    Debug.Print "EC�F" & dblData(2)
    Debug.Print "���W���F" & CStr(varCoefficient)
    Debug.Print "m0�F" & CStr(CDec(varmo))
    Debug.Print "�ް���Ċ�ߗ�(f)�F" & CStr(CDec(varDilution))
    Debug.Print "�ް���Ďw���F" & CStr(CDec(dblprmSijiDp))
    Debug.Print "�ް���Ď��сF" & CStr(CDec(sDpData.InpDp))
    Debug.Print "���������ܗL�s�����ʁF" & CStr(CDec(sDpData.PutDp))
    Debug.Print "�v���ް�ߗʁF" & CStr(CDec(DopeCalculationPart))
    
    Exit Function

ErrHand:
    '///�I��
    gErr.Pop
    Exit Function
End Function

'=============================================================================
'@(f)       GetResData
'
'�@�\       �_����R���������擾����
'
'�߂�l     True/False
'
'����       �_������                    strType As String
'�@�@�@     �_����R(��)                dblRes As Double
'�@�@�@     �_����R(����)              dblTgResBtm As Double
'�@�@�@     �_����R(���)              dblTgResTop As Double
'
'�@�\�T�v
'
'���l
'=============================================================================
Public Function GetResData(strType As String, _
                           dblRes As Double, _
                           dblTgResBtm As Double, _
                           dblTgResTop As Double) As FUNCTION_RETURN
    
    GetResData = FUNCTION_RETURN_FAILURE
        
    '���_�����߂��o�̏ꍇ
    If strType = "P" Then
        Select Case dblRes
            '���ς�10.0�ȏ�̏ꍇ
            Case Is >= 10
                dblTgResBtm = 10
                dblTgResTop = 99999
            '���ς�1.0�ȏ�̏ꍇ
            Case Is >= 1
                dblTgResBtm = 1
                dblTgResTop = 10
            '���ς�0.1�ȏ�̏ꍇ
            Case Is >= 0.1
                dblTgResBtm = 0.1
                dblTgResTop = 1
            '���ς�0.0195�ȏ�̏ꍇ
            Case Is >= 0.0195
                dblTgResBtm = 0.0195
                dblTgResTop = 0.1
            '���ς�0.01�ȏ�̏ꍇ
            Case Is >= 0.01
                dblTgResBtm = 0.01
                dblTgResTop = 0.0195
            '���ς�0.005�ȏ�̏ꍇ
            Case Is >= 0.005
                dblTgResBtm = 0.005
                dblTgResTop = 0.01
            '���ς�0.001�ȏ�̏ꍇ
            Case Is >= 0.001
                dblTgResBtm = 0.001
                dblTgResTop = 0.005
        End Select
    '���_�����߂��m�̏ꍇ
    ElseIf strType = "N" Then
        Select Case dblRes
            '���ς�10.0�ȏ�̏ꍇ
            Case Is >= 10
                dblTgResBtm = 0
                dblTgResTop = 0
            '���ς�1.0�ȏ�̏ꍇ
            Case Is >= 1
                dblTgResBtm = 1
                dblTgResTop = 10
            '���ς�0.1�ȏ�̏ꍇ
            Case Is >= 0.1
                dblTgResBtm = 0.1
                dblTgResTop = 1
            '���ς�0.245�ȏ�̏ꍇ
            Case Is >= 0.0245
                dblTgResBtm = 0.0245
                dblTgResTop = 0.1
            '���ς�0.01�ȏ�̏ꍇ
            Case Is >= 0.01
                dblTgResBtm = 0.01
                dblTgResTop = 0.0245
            '���ς�0.01��菬�����ꍇ
            Case Is < 0.01
                dblTgResBtm = 0
                dblTgResTop = 0.01
        End Select
    '���_�����߂�sb�̏ꍇ
    ElseIf strType = "sb" Then
        Select Case dblRes
            '���ς�0.05�ȏ�̏ꍇ
            Case Is >= 0.05
                dblTgResBtm = 0.05
                dblTgResTop = 0
            '���ς�0.01�ȏ�̏ꍇ
            Case Is >= 0.015
                dblTgResBtm = 0.015
                dblTgResTop = 0.05
            '���ς�0.01��菬�����ꍇ
            Case Is < 0.015
                dblTgResBtm = 0
                dblTgResTop = 0.015
        End Select
    End If
    
    GetResData = FUNCTION_RETURN_SUCCESS
    
    Exit Function

End Function

'=============================================================================
'@(f)       GetmoCalData
'
'�@�\       mo�Z�o�p�ް��擾
'
'�߂�l     True/False
'
'����       �_������                    strType As String
'�@�@�@     �_����R(����)              dblTgResBtm As Double
'�@�@�@     �_����R(���)              dblTgResTop As Double
'�@�@�@     EE,EZ,EC                    dblData() As Double
'
'�@�\�T�v
'
'���l
'=============================================================================
Public Function GetmoCalData(strType As String, _
                             dblTgResBtm As Double, _
                             dblTgResTop As Double, _
                             dblData() As Double) As FUNCTION_RETURN
    
    Dim strSql          As String       'SQL
    Dim rs              As OraDynaset   'ں��޾��
    Dim strTgResBtm     As String       '�_����R(����)
    Dim strTgResTop     As String       '�_����R(���)
    
    GetmoCalData = FUNCTION_RETURN_FAILURE
    
    '///�^�ϊ�����
    strTgResBtm = CStr(dblTgResBtm)     '�_����R(����)
    strTgResTop = CStr(dblTgResTop)     '�_����R(���)
    
    '-----------
    '  SQL���s
    '-----------
    '///�����F���ы敪 = 'K'
    '   �@�@�@��ʺ��� = 'A5'
    '   �@�@�@�֘A���� = strType(�_������)
    '�@�@�@ �@�ް�1 �@ = strTgResBtm(�_����R(����))
    '�@�@�@   �ް�2 �@ = strTgResTop(�_����R(���))
    strSql = ""
    strSql = strSql & "SELECT KCODE03A9, KCODE04A9, KCODE05A9"
    strSql = strSql & "  FROM KODA9"
    strSql = strSql & " WHERE SYSCA9 = 'K'"
    strSql = strSql & "   AND SHUCA9 = 'A5'"
    strSql = strSql & "   AND KCODEA9 = '" & LCase$(strType) & "' "
    strSql = strSql & "   AND KCODE01A9 = '" & strTgResBtm & "' "
    strSql = strSql & "   AND KCODE02A9 = '" & strTgResTop & "' "
    
    Set rs = OraDB.DBCreateDynaset(strSql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        Exit Function
    End If
    
    '///�擾�ް��i�[
    dblData(0) = rs("KCODE03A9")        'EE
    dblData(1) = rs("KCODE04A9")        'EZ
    dblData(2) = rs("KCODE05A9")        'EC
    
    rs.Close
    
    GetmoCalData = FUNCTION_RETURN_SUCCESS
    
    Exit Function

End Function

'=============================================================================
'@(f)       GetCoefficientData
'
'�@�\       ���W���擾����
'
'�߂�l     True/False
'
'����       �_������                    strType As String
'�@�@�@     ���W��                      varCoefficient As Variant
'           ����އ�                     strChrgNo As String(txtChargeNo.Text)
'
'�@�\�T�v
'
'���l
'=============================================================================
Public Function GetCoefficientData(strType As String, _
                                   varCoefficient As Variant, _
                                   strChrgNo As String) As FUNCTION_RETURN
    
    Dim strSql          As String       'SQL
    Dim rs              As OraDynaset   'ں��޾��
    Dim strCore         As String       '��
    Dim strEdit         As String       '���ް��ҏW�p�ϐ�
    Dim intInstr        As Integer      '"<"�J�n�ʒu

    GetCoefficientData = FUNCTION_RETURN_FAILURE

    '-----------
    '  SQL���s
    '-----------
    '///�����F���ы敪         = 'SC'(TBCMB005)
    '   �@�@�@��ʺ���         = '2'(TBCMB005)
    '   �@�@�@����             = TBCME018.HSXCDIR(TBCMB005)
    '�@�@�@ �@�i��  �@         = TBCMH001.HINBAN(TBCME018)
    '�@�@�@   ���i�ԍ������ԍ� = TBCMH001.NMNOREVNO(TBCME018)
    '   �@�@�@�H��             = TBCMH001.NFACTORY(TBCME018)
    '         ���Ə���         = TBCMH001.NOPECOND(TBCME018)
    '         �iSX�����ʕ���   = TBCMB005.CODE(TBCME018)
    strSql = ""
    strSql = strSql & "SELECT DA9.KCODEA9 AS JIKU"
    strSql = strSql & "  FROM KODA9 DA9, TBCME018 TE18, TBCMH001 H01"
    strSql = strSql & " WHERE DA9.SYSCA9 = 'K'"
    strSql = strSql & "   AND DA9.SHUCA9 = 'AI'"
    strSql = strSql & "   AND H01.UPINDNO = '" & strChrgNo & "' "
    strSql = strSql & "   AND TE18.HINBAN = H01.HINBAN"
    strSql = strSql & "   AND TE18.MNOREVNO = H01.NMNOREVNO"
    strSql = strSql & "   AND TE18.FACTORY = H01.NFACTORY"
    strSql = strSql & "   AND TE18.OPECOND = H01.NOPECOND"
    '*CHG* �Q�ƶ�ѕύX TCS)K.Kunori 2004.11.24
'''    strSql = strSql & "   AND TRIM(DA9.CODEA9) = TRIM(TE18.HSXCDIR)"
    strSql = strSql & "   AND TRIM(DA9.CODEA9) = SUBSTR(TE18.MCNO,2,1)"
    
    Set rs = OraDB.DBCreateDynaset(strSql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        Exit Function
    End If
    
    '///���ް��ҏW���i�[
    strCore = IIf(IsNull(rs("JIKU")), vbNullString, rs("JIKU")) '��
    
    rs.Close
    
    '-----------
    '  SQL���s
    '-----------
    '///�����F���ы敪 = 'K'
    '   �@�@�@��ʺ��� = 'A6'
    '   �@�@�@�֘A���� = strType(�_������)
    '�@�@�@   �ް�2 �@ = strCore(��)
    strSql = ""
    strSql = strSql & "SELECT CTR01A9 AS COEFF"
    strSql = strSql & "  FROM KODA9"
    strSql = strSql & " WHERE SYSCA9 = 'K'"
    strSql = strSql & "   AND SHUCA9 = 'A6'"
    strSql = strSql & "   AND KCODE01A9 = '" & LCase$(strType) & "' "
    strSql = strSql & "   AND KCODE02A9 = '" & "<" & strCore & ">" & "' "
    
    Set rs = OraDB.DBCreateDynaset(strSql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        Exit Function
    End If
    
    '///�擾�ް��i�[
    varCoefficient = CDbl(rs("COEFF"))    '���W��
    
    rs.Close
    
    GetCoefficientData = FUNCTION_RETURN_SUCCESS
    
    Exit Function

End Function
'*ADD* �v���ް�ߗʌv�Z���@�ύX�Ή� ��Ӽޭ�ٓ��� TCS)K.Kunori 2004.11.29 END <<<

'*ADD* �␳�W���擾�����ǉ� TCS)K.Kunori 2004.12.16 START >>>
'=============================================================================
'@(f)       GetSupplCoefficientData
'
'�@�\       �␳�W���擾����
'
'�߂�l     True/False
'
'����       �ް���Ċ�ߗ�(f)            varDilution As Variant
'�@�@�@     �␳�W��                    varSupplCoefficient As Variant
'�@�@�@     �^�C�v                      strType As String
'�@�@�@     �_����R                    dblNerai As String
'
'�@�\�T�v
'
'���l
'=============================================================================
Public Function GetSupplCoefficientData(dblDilution As Double _
                                      , dblSupplCoefficient As Double _
                                      , ByVal strType As String _
                                      , ByVal dblNerai As Double _
                                      ) As FUNCTION_RETURN
    
    Dim strSql          As String       'SQL
    Dim rs              As OraDynaset   'ں��޾��
    Dim dblfBtm         As Double       '���l����
    Dim dblfTop         As Double       '���l���
    Dim dblGanma        As Double       '�␳�W��(�ޔ�p)
    
    GetSupplCoefficientData = FUNCTION_RETURN_FAILURE
    
    '///�ް����̏ꍇ��'1'�Œ�Ƃ���
    dblSupplCoefficient = 1
    
    '-----------
    '  SQL���s
    '-----------
    '///�����F���ы敪 = 'K'
    '   �@�@�@��ʺ��� = 'AO'
    strSql = ""
    strSql = strSql & "SELECT KCODE01A9, KCODE02A9, KCODE03A9" & vbLf
    strSql = strSql & "  FROM KODA9" & vbLf
    strSql = strSql & " WHERE SYSCA9 = 'K'" & vbLf
    strSql = strSql & "   AND SHUCA9 = 'AO'"
    
    Set rs = OraDB.DBCreateDynaset(strSql, ORADYN_DEFAULT)
    
    '���ް��L�̏ꍇ
    If rs.RecordCount <> 0 Then
        
        '///�擾�ް��i�[
        dblfBtm = val(NulltoStr(rs("KCODE01A9")))               '���l����
        dblfTop = val(NulltoStr(rs("KCODE02A9")))               '���l���
        dblGanma = val(NulltoStr(rs("KCODE03A9")))              '�␳�W��
        
        '�����l������O�ȊO�̏ꍇ
        If dblfTop <> 0 Then
            '�����l���͈͓��̏ꍇ(�����l�ȏ����l����)
            If dblfBtm <= dblDilution And dblDilution < dblfTop Then
                dblSupplCoefficient = dblGanma
                '�␳�W���Z�o���@�ύX 2007/03/05�ǉ� SETsw kubota
                If GetHosei_Nerai(strType, CStr(dblNerai), dblSupplCoefficient) = False Then
                    Exit Function
                End If
            End If
        '�����l������O(NULL)�̏ꍇ
        Else
            '�����l�������l�ȏ�̏ꍇ
            If dblfBtm <= dblDilution Then
                dblSupplCoefficient = dblGanma
                '�␳�W���Z�o���@�ύX 2007/03/05�ǉ� SETsw kubota
                If GetHosei_Nerai(strType, CStr(dblNerai), dblSupplCoefficient) = False Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    rs.Close
    
    GetSupplCoefficientData = FUNCTION_RETURN_SUCCESS
    
    Exit Function

End Function
'*ADD* �␳�W���擾�����ǉ� TCS)K.Kunori 2004.12.16 END <<<

'*ADD* �_����R�ʕ␳�W���擾�����ǉ�(200mm�֐��̃R�s�[) SETsw kubota 2007.03.05 START >>>
'�T�v      :�␳�W���̎擾
'���Ұ�(In):�^�C�v
'           �˂炢��R
'          :�߂�l�F����^�ُ�
'����      :
'����      :2007.02.19 �쐬
Public Function GetHosei_Nerai(ByVal sType As String _
                             , ByVal sNerai As String _
                             , ByRef dblSupplCoefficient As Double _
                             ) As Boolean

    Dim sSql        As String
    Dim objDS       As Object
    Dim dblLow      As Double
    Dim dblHigh     As Double
    Dim dblG        As Double
    Dim dblNerai    As Double
    Dim lCnt        As Long
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    GetHosei_Nerai = False
    
    dblNerai = val(sNerai)
    
    '�R�[�h�ݒ�Ȃ����̏ꍇ�A�␳�W����ݒ肹���ɏI������
    '''gdblGanma = 1   '���݂��Ȃ��ꍇ�A�P�Œ�Ƃ���B
    
    '�␳�W���擾
    sSql = ""
    sSql = sSql & "SELECT kcode01a9,kcode02a9,ctr01a9" & vbLf
    sSql = sSql & "  FROM KODA9 " & vbLf
    sSql = sSql & " WHERE SYSCA9 = 'X'" & vbCrLf
    sSql = sSql & "   AND SHUCA9 = 'RS'" & vbCrLf
    sSql = sSql & "   AND CODEA9 LIKE '" & Left$(UCase$(sType), 1) & "%'" & vbCrLf
    sSql = sSql & " ORDER BY CODEA9" & vbCrLf
    
    Set objDS = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    
    For lCnt = 1 To objDS.RecordCount
    
        dblLow = val(NulltoStr(objDS(0)))
        dblHigh = val(NulltoStr(objDS(1)))
        dblG = val(NulltoStr(objDS(2)))
        
        If dblLow = 0 Then
            '�����l��Null(= 0)�̏ꍇ�́A����l�����̎��擾�f�[�^��ݒ肷��B
            If dblHigh > dblNerai Then
                dblSupplCoefficient = dblG
                Exit For
            End If
        ElseIf dblHigh = 0 Then
            '����l��Null(= 0)�̏ꍇ�́A�����l�ȏ�̎��擾�f�[�^��ݒ肷��B
            If dblLow <= dblNerai Then
                dblSupplCoefficient = dblG
                Exit For
            End If
        Else
            '�㉺���l���ݒ肳��Ă���ꍇ�́A�㉺���͈͓��̎��擾�f�[�^��ݒ肷��
            If dblLow <= dblNerai And dblNerai < dblHigh Then
                dblSupplCoefficient = dblG
                Exit For
            End If
        End If
        objDS.MoveNext
    
    Next lCnt
    
    GetHosei_Nerai = True
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit
    
End Function
'*ADD* �_����R�ʕ␳�W���擾�����ǉ�(200mm�֐��̃R�s�[) SETsw kubota 2007.03.05 END <<<

'///////////////////////////////////////////////////
' @(f)
' �@�\    : ��R�Z�x�擾 Kameda
' �Ԃ�l  : True:����
'           False:���s
' ������  : �ް���Ď��
' �@�\����: ��R�Z�x����o�^�ΐ͌W����ǂݍ���
' �C������: �Z�x(CTR01A)�����ȉ��S�����R��   2009/12/17
'///////////////////////////////////////////////////
Public Function Getrnoudo(sDopant As String, sPuro As Double) As Double
    Dim sSqlStmt As String
    Dim objOraDyn As Object
    Dim lCount As Long
    Dim syubetu As String
    Dim sSql As String
    
    Getrnoudo = 0
    
    'SQL�ҏW
    sSqlStmt = "SELECT  NVL(shuca9   , ' ')"                  ' 0:[��ʃR�[�h]
    sSqlStmt = sSqlStmt & ",NVL(kcode01a9, ' ')"                  ' 1:[�f�[�^�P] �c ��R����
    sSqlStmt = sSqlStmt & ",NVL(kcode02a9, ' ')"                  ' 2:[�f�[�^�Q] �c ��R���
    sSqlStmt = sSqlStmt & ",NVL(ctr01a9, 0)"                      ' 3:[������P]�@�c �h�[�p���g�Z�x
    sSqlStmt = sSqlStmt & "  FROM koda9 "
    sSqlStmt = sSqlStmt & " WHERE sysca9 = 'X'"
    sSqlStmt = sSqlStmt & "   AND shuca9 >= 'D0'"
    sSqlStmt = sSqlStmt & "   AND shuca9 <= 'D9'"
    sSqlStmt = sSqlStmt & "   AND codea9 = '" & sDopant & "'"
    sSqlStmt = sSqlStmt & " ORDER BY shuca9"
    
    ''�_�C�i�Z�b�g�쐬
    If DynSet2(objOraDyn, sSqlStmt) = False Then
        Exit Function
    End If
    If objOraDyn.EOF Then
        ''�Y�������ނ���������
        Exit Function
    End If
    
    For lCount = 1 To objOraDyn.RecordCount
    
       If objOraDyn(1) <= sPuro Or objOraDyn(1) = "" Then
          If sPuro < objOraDyn(2) Or objOraDyn(2) = "" Then
                Getrnoudo = objOraDyn(3) / 10
          End If
       End If
       objOraDyn.MoveNext
   Next lCount
      

End Function
'///////////////////////////////////////////////////
' @(f)
' �@�\    : �h�[�p���g��ގ擾 kameda
' �Ԃ�l  : True:����
'           False:���s
' ������  :
' �@�\����:
' �C������:
'///////////////////////////////////////////////////
Public Function GetDopeKind(sDopeKind() As String) As FUNCTION_RETURN
    Dim sSqlStmt As String
    Dim objOraDyn As Object
    Dim lCount As Long
    Dim syubetu As String
    Dim sSql As String
    
    GetDopeKind = FUNCTION_RETURN_FAILURE
    
    'SQL�ҏW
    sSqlStmt = "SELECT  NVL(codea9   , ' ')"
    sSqlStmt = sSqlStmt & "  FROM koda9 "
    sSqlStmt = sSqlStmt & " WHERE sysca9 = 'X'"
    sSqlStmt = sSqlStmt & "   AND shuca9 = 'D0'"
    sSqlStmt = sSqlStmt & " ORDER BY codea9"
    
    ''�_�C�i�Z�b�g�쐬
    If DynSet2(objOraDyn, sSqlStmt) = False Then
        Exit Function
    End If
    If objOraDyn.EOF Then
        ''�Y�������ނ���������
        Exit Function
    End If
    
    ReDim sDopeKind(objOraDyn.RecordCount)
    
    For lCount = 1 To objOraDyn.RecordCount
          sDopeKind(lCount) = objOraDyn(0)
       objOraDyn.MoveNext
   Next lCount
      
      GetDopeKind = FUNCTION_RETURN_SUCCESS

End Function

'///////////////////////////////////////////////////
' @(f)
' �@�\    : �h�[�p���g�v�Z�p�t�@�N�^�[�E�ΐ͌W���擾
'
' �Ԃ�l  : false:���s
'           true:�擾����
'
' ������  : �����敪
'           �^�C�v:"P"/"N"
'           �t�@�N�^�[(OUT)
'           �ΐ͌W��(OUT)
'
' �@�\����: �h�[�p���g�v�Z�p�t�@�N�^�[�E�ΐ͌W���擾
'///////////////////////////////////////////////////
Function GetFactor(ByVal sType As String, ByRef sngFactor As Single, ByRef sngHenseki As Single) As Boolean
    GetFactor = False
    
    Dim sSqlStmt As String
    Dim objOraDyn As Object
    
    
    ''�r�p�k���쐬
    sSqlStmt = "SELECT NVL(kcodea9, ' '),                   "
    sSqlStmt = sSqlStmt & "NVL(kcode01a9, ' ')                "
    sSqlStmt = sSqlStmt & "FROM koda9                       "
    sSqlStmt = sSqlStmt & "WHERE sysca9 = 'X'               "
    sSqlStmt = sSqlStmt & "  AND shuca9 = '36'              "
    sSqlStmt = sSqlStmt & "  AND codea9 = '" & sType & "' "
    
    ''�_�C�i�Z�b�g�쐬
    If DynSet2(objOraDyn, sSqlStmt) = False Then
        ''�_�C�i�Z�b�g�쐬���s
        Call MsgOut(100, sSqlStmt, ERR_DISP_LOG)
        
        GetFactor = False
        Exit Function
    End If
    If objOraDyn.EOF Then
        GetFactor = False
        Exit Function
    End If

    sngHenseki = objOraDyn(0)   ''�ΐ͌W��
    sngFactor = objOraDyn(1)    ''�t�@�N�^�[
   
    GetFactor = True
End Function

