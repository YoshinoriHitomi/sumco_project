Attribute VB_Name = "s_cmbc021_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMJ003 (�n������)
' �Q�Ɓ@�@: 060211_��������
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmjc001b_Disp
   ' CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
   ' TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    factory As String * 1           ' �H��
    opecond As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    OIMEAS1 As Double               ' �n������l�P
    OIMEAS2 As Double               ' �n������l�Q
    OIMEAS3 As Double               ' �n������l�R
    OIMEAS4 As Double               ' �n������l�S
    OIMEAS5 As Double               ' �n������l�T
    ORGRES As Double                ' �n�q�f����
    SETDTM As Date                  ' �ݒ����
    EFFECTTM As Integer             ' �L������
    FTIRMETH As String              ' �e�s�h�q���֎�
    YCOEF As Double                 ' �e�s�h�q���Z���i�x�ؕЁj
    XCOEF As Double                 ' �e�s�h�q���Z���i�w�W���j
    AVE As Double                   ' �`�u�d
    SIGMA As Double                 ' �Ёi�V�O�}�j
    FTIRCONV As Double              ' �e�s�h�q���Z
    INSPECTWAY As String * 2        ' �������@
    JudgData As Double
   ' TSTAFFID As String * 8          ' �o�^�Ј�ID
   ' REGDATE As Date                 ' �o�^���t
   ' KSTAFFID As String * 8          ' �X�V�Ј�ID
   ' UPDDATE As Date                 ' �X�V���t
   ' SENDFLAG As String * 1          ' ���M�t���O
   ' SENDDATE As Date                ' ���M���t
End Type


'                                     2001/06/20
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMJ004 (Cs����)
' �Q�Ɓ@�@: 060211_��������
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmjc001b_Disp2
'    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
'    TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    factory As String * 1           ' �H��
    opecond As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    CSMEAS As Double                ' Cs�����l
    PRE70P As Double                ' �V�O������l
    INSPECTWAY As String * 2        ' �������@
 '   TSTAFFID As String * 8          ' �o�^�Ј�ID
 '   REGDATE As Date                 ' �o�^���t
 '   KSTAFFID As String * 8          ' �X�V�Ј�ID
 '   UPDDATE As Date                 ' �X�V���t
 '   SENDFLAG As String * 1          ' ���M�t���O
 '   SENDDATE As Date                ' ���M���t
End Type
'                                     2001/06/20
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMJ003 (�n������)
' �Q�Ɓ@�@: 060211_��������
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmjc001c_Disp
   ' CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
   ' TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    factory As String * 1           ' �H��
    opecond As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    OIMEAS1 As Double               ' �n������l�P
    OIMEAS2 As Double               ' �n������l�Q
    OIMEAS3 As Double               ' �n������l�R
    OIMEAS4 As Double               ' �n������l�S
    OIMEAS5 As Double               ' �n������l�T
    ORGRES As Double                ' �n�q�f����
    SETDTM As Date                  ' �ݒ����
    EFFECTTM As Integer             ' �L������
    FTIRMETH As String              ' �e�s�h�q���֎�
    YCOEF As Double                 ' �e�s�h�q���Z���i�x�ؕЁj
    XCOEF As Double                 ' �e�s�h�q���Z���i�w�W���j
    AVE As Double                   ' �`�u�d
    SIGMA As Double                 ' �Ёi�V�O�}�j
    FTIRCONV As Double              ' �e�s�h�q���Z
    INSPECTWAY As String * 2        ' �������@
    JudgData As Double              ' �����Ώےl
   ' TSTAFFID As String * 8          ' �o�^�Ј�ID
   ' REGDATE As Date                 ' �o�^���t
   ' KSTAFFID As String * 8          ' �X�V�Ј�ID
   ' UPDDATE As Date                 ' �X�V���t
   ' SENDFLAG As String * 1          ' ���M�t���O
   ' SENDDATE As Date                ' ���M���t
End Type


'(2002/07 s_cmzcF_TBCME019_SQL.bas���ړ�)
'�t�B�[���h�������p
Dim fldNames() As String    '��rs�Ɋ܂܂��t�B�[���h���ێ��z��
Dim fldCnt As Integer       '��rs�Ɋ܂܂��t�B�[���h��



'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ003�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_cmjc001b_Disp ,���o���R�[�h
'          :SPLNUMs()     ,I  ,Integer      ,���o�����z��(�T���v��No)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/20(Wed)�쐬�@����
Public Function DBDRV_Getcmjc001b_Disp(records() As typ_cmjc001b_Disp, SPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere As String  'SQL��WHERE����
Dim sqlGroup As String  'SQL��GROUP����
Dim sqlOrder As String  'SQL��Order����
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long           '���[�v�J�E���g

    DBDRV_Getcmjc001b_Disp = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001b_SQL.bas -- Function DBDRV_Getcmjc001b_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, MAX(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA "
    sqlBase = sqlBase & "From TBCMJ003"
    ''���o����(�����NO)�̎��o��
    sqlWhere = "Where SMPLNO in ("
    For i = 1 To UBound(SPLNUMs)
        sqlWhere = sqlWhere & "'" & SPLNUMs(i) & "'"
        If i < UBound(SPLNUMs) Then
            sqlWhere = sqlWhere & ", "
        End If
    Next
    sqlWhere = sqlWhere & ") "
    sqlGroup = "GROUP BY CRYNUM, POSITION, SMPKBN, TRANCOND "
    sqlOrder = "ORDER BY POSITION"
    sql = sqlBase & sqlWhere & sqlGroup & sqlOrder
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_Getcmjc001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .POSITION = rs("POSITION")       ' �ʒu
            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
            .TRANCOND = rs("TRANCOND")       ' ��������
            .SMPLNO = rs("SMPLNO")           ' �T���v���m��
            .SMPLUMU = rs("SMPLUMU")         ' �T���v���L��
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .GOUKI = rs("GOUKI")             ' ���@
            .OIMEAS1 = rs("OIMEAS1")         ' �n������l�P
            .OIMEAS2 = rs("OIMEAS2")         ' �n������l�Q
            .OIMEAS3 = rs("OIMEAS3")         ' �n������l�R
            .OIMEAS4 = rs("OIMEAS4")         ' �n������l�S
            .OIMEAS5 = rs("OIMEAS5")         ' �n������l�T
            .ORGRES = rs("ORGRES")           ' �n�q�f����
            .SETDTM = rs("SETDTM")           ' �ݒ����
            .EFFECTTM = rs("EFFECTTM")       ' �L������
            .FTIRMETH = rs("FTIRMETH")       ' �e�s�h�q���֎�
            .YCOEF = rs("YCOEF")             ' �e�s�h�q���Z���i�x�ؕЁj
            .XCOEF = rs("XCOEF")             ' �e�s�h�q���Z���i�w�W���j
            .AVE = rs("AVE")                 ' �`�u�d
            .SIGMA = rs("SIGMA")             ' �Ёi�V�O�}�j
            .FTIRCONV = rs("FTIRCONV")       ' �e�s�h�q���Z
            .INSPECTWAY = rs("INSPECTWAY")   ' �������@
            .JudgData = rs("JUDGDATA")       ' �����Ώےl
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001b_Disp = FUNCTION_RETURN_SUCCESS


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



'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------
'�T�v      :�e�[�u���uTBCMJ004�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_cmjc001b_Disp2 ,���o���R�[�h
'          :SPLNUMs()     ,I  ,Integer      ,���o�����z��(�T���v��No)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/20�쐬�@����
Public Function DBDRV_Getcmjc001b_Disp2(records() As typ_cmjc001b_Disp2, SPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere As String  'SQL��WHERE����
Dim sqlGroup As String  'SQL��GROUP����
Dim sqlOrder As String  'SQL��Order����
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    DBDRV_Getcmjc001b_Disp2 = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001b_SQL.bas -- Function DBDRV_Getcmjc001b_Disp2"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, CSMEAS, PRE70P, INSPECTWAY "
    sqlBase = sqlBase & "From TBCMJ004"
    ''���o����(�����NO)�̎��o��
    sqlWhere = "Where SMPLNO in ("
    For i = 1 To UBound(SPLNUMs)
        sqlWhere = sqlWhere & "'" & SPLNUMs(i) & "'"
        If i < UBound(SPLNUMs) Then
            sqlWhere = sqlWhere & ", "
        End If
    Next
    sqlWhere = sqlWhere & ") "
    sqlGroup = "GROUP BY CRYNUM, POSITION, SMPKBN, TRANCOND "
    sqlOrder = "ORDER BY POSITION"
    sql = sqlBase & sqlWhere & sqlGroup & sqlOrder
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_Getcmjc001b_Disp2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .POSITION = rs("POSITION")       ' �ʒu
            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
            .TRANCOND = rs("TRANCOND")       ' ��������
            .SMPLNO = rs("SMPLNO")           ' �T���v���m��
            .SMPLUMU = rs("SMPLUMU")         ' �T���v���L��
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .GOUKI = rs("GOUKI")             ' ���@
            .CSMEAS = rs("CSMEAS")           ' Cs�����l
            .PRE70P = rs("PRE70P")           ' �V�O������l
            .INSPECTWAY = rs("INSPECTWAY")   ' �������@
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001b_Disp2 = FUNCTION_RETURN_SUCCESS

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


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�����œn���ꂽ���R�[�h��TBCMJ003�ɒǉ�����
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :record        ,I  ,typ_TBCMJ003 ,���o���R�[�h
'          :CRYNUM        ,I  ,String       ,�����ԍ�
'          :TSTAFFID      ,I  ,String       ,�o�^�Ј�ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�V�K�ǉ��̍ہA�����񐔂̓e�[�u����̍ő�l+1�Ƃ���B
'����      :2001/06/20(Wed)�쐬�@����

Public Function DBDRV_Getcmjc001b_Exec(record As typ_cmjc001b_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL�x�[�X����
Dim sqlWhere As String  'SQLWhere����
Dim sqlGroup As String  'SQLGroup����
Dim SetDate As Variant  '�ݒ����
Dim rs As OraDynaset    'OracleDynaset

'    CRYNUM             �����ԍ��@�ˈ���
'    TRANCNT         �@ �����񐔁@�ˍő�
'   TSTAFFID            �o�^�Ј�ID�@�ˈ���
 '   REGDATE �@�@�@     �o�^���t�@��SYSDATE
 '   KSTAFFID           �X�V�Ј�ID�@��" "
 '   UPDDATE            �X�V���t�@��SYSDATE
 '   SENDFLAG           ���M�t���O�@��"0"
 '   SENDDATE           ���M���t�@��SYSDATE
    
    DBDRV_Getcmjc001b_Exec = FUNCTION_RETURN_FAILURE
    
    ''�ő�J�E���g�擾

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001b_SQL.bas -- Function DBDRV_Getcmjc001b_Exec"

    sqlBase = "select nvl(MAX(TRANCNT),0) + 1 as w_TRANCNT from TBCMJ003 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
    ''�ݒ�����t�H�[�}�b�g����
    SetDate = Format$(record.SETDTM, "yyyy-mm-dd hh:mm:ss")

    ''SQL��g�ݗ��Ă�
    sql = "Insert into TBCMJ003 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
    sql = sql & "Values( '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', " & rs!w_TRANCNT & ", " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', " & record.OIMEAS1 & ", " & _
               record.OIMEAS2 & ", " & record.OIMEAS3 & ", " & record.OIMEAS4 & ", " & record.OIMEAS5 & ", " & record.ORGRES & ", " & _
               "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss'), " & record.EFFECTTM & ", '" & record.FTIRMETH & "', " & record.YCOEF & ", " & record.XCOEF & ", " & _
               record.AVE & ", " & record.SIGMA & ", " & record.FTIRCONV & ", '" & record.INSPECTWAY & "', " & record.JudgData & ", '" & TSTAFFID & "', " & _
               "SYSDATE, ' ', SYSDATE, '0', SYSDATE) "
''''    '' OI_NULL�Ή��@2005/03/07 TUKU START �R���a����̈˗��ɂ��ύX���~--------------------------------------------------------------------
''''    sql = "Insert into TBCMJ003 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
''''              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
''''              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
''''    sql = sql & "Values( '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', " & rs!w_TRANCNT & ", " & _
''''               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
''''               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', "
''''                If (record.OIMEAS1 = -1) Then sql = sql & " NULL , " Else sql = sql & record.OIMEAS1 & ", "    'OI����l1
''''                If (record.OIMEAS2 = -1) Then sql = sql & " NULL , " Else sql = sql & record.OIMEAS2 & ", "    'OI����l2
''''                If (record.OIMEAS3 = -1) Then sql = sql & " NULL , " Else sql = sql & record.OIMEAS3 & ", "    'OI����l2
''''                If (record.OIMEAS4 = -1) Then sql = sql & " NULL , " Else sql = sql & record.OIMEAS4 & ", "    'OI����l3
''''                If (record.OIMEAS5 = -1) Then sql = sql & " NULL , " Else sql = sql & record.OIMEAS5 & ", "    'OI����l4
''''                If (record.ORGRES = -1) Then sql = sql & " NULL , " Else sql = sql & record.ORGRES & ", "      'ORG
''''    sql = sql & "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss'), " & record.EFFECTTM & ", '" & record.FTIRMETH & "', " & record.YCOEF & ", " & record.XCOEF & ", " & _
''''               record.AVE & ", " & record.SIGMA & ", " & record.FTIRCONV & ", '" & record.INSPECTWAY & "', " & record.JudgData & ", '" & TSTAFFID & "', " & _
''''               "SYSDATE, ' ', SYSDATE, '0', SYSDATE) "
''''    '' OI_NULL�Ή��@2005/03/07 TUKU END   --------------------------------------------------------------------

    ''SQL�̎��s
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001b_Exec = FUNCTION_RETURN_SUCCESS
    

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


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�����œn���ꂽ���R�[�h��TBCMJ004�ɒǉ�����
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :record        ,I  ,typ_TBCMJ004 ,���o���R�[�h
'          :CRYNUM        ,I  ,String       ,�����ԍ�
'          :TSTAFFID      ,I  ,String       ,�o�^�Ј�ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�V�K�ǉ��̍ہA�����񐔂̓e�[�u����̍ő�l+1�Ƃ���B
'����      :2001/06/22(Fri)�쐬�@����

Public Function DBDRV_Getcmjc001b_Exec2(record As typ_cmjc001b_Disp2, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL�x�[�X����
Dim sqlWhere As String  'SQLWhere����
Dim sqlGroup As String  'SQLGroup����

'    CRYNUM             �����ԍ��@�ˈ���
'    TRANCNT         �@ �����񐔁@�ˍő�
'    TSTAFFID           �o�^�Ј�ID�@�ˈ���
'    REGDATE �@�@�@     �o�^���t�@��SYSDATE
'    KSTAFFID           �X�V�Ј�ID�@�ˁh�h
'    UPDDATE            �X�V���t�@��SYSDATE
'    SENDFLAG           ���M�t���O�@��"�O"
'    SENDDATE           ���M���t�@��SYSDATE
    
    DBDRV_Getcmjc001b_Exec2 = FUNCTION_RETURN_FAILURE
    
    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001b_SQL.bas -- Function DBDRV_Getcmjc001b_Exec2"

    sqlBase = "Insert into TBCMJ004 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, CSMEAS, PRE70P, INSPECTWAY, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', " & record.CSMEAS & ", " & _
               record.PRE70P & ", '" & record.INSPECTWAY & "', '" & TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ004 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
''''    '' OI_NULL�Ή��@2005/03/07 TUKU START �R���a����̈˗��ŕύX���~--------------------------------------------------------------------
''''    sqlBase = "Insert into TBCMJ004 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
''''              " PROCCODE, GOUKI, CSMEAS, PRE70P, INSPECTWAY, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
''''    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
''''               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
''''               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', "
''''                If (record.CSMEAS = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.CSMEAS & ", "    'CS����l
''''                If (record.PRE70P = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.PRE70P & ", "    'CS70%
''''    sqlBase = sqlBase & " '" & record.INSPECTWAY & "', '" & TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ004 "
''''    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'''''    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
''''    sql = sqlBase & sqlWhere & sqlGroup
''''    '' OI_NULL�Ή��@2005/03/07 TUKU END   --------------------------------------------------------------------
            
    ''SQL�̎��s
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001b_Exec2 = FUNCTION_RETURN_SUCCESS



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





'�T�v      :�����̃t�B�[���h����fldNames()�z��Ɋ܂܂�Ă��邩�ǂ����̔���B
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :fldName       ,I  ,typ_TBCME018 ,���o���R�[�h
'          :�߂�l        ,O  ,Boolean      ,True:�݂�^False�F����
'����      :
'����      :2001/06/27�쐬�@�쑺 (2002/07 s_cmzcF_TBCME019_SQL.bas���ړ�)

Private Function fldNameExist(fldName As String) As Boolean
    Dim sql         As String           'SQL�S��
    Dim i As Integer                    'ٰ�߶���


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME019_SQL.bas -- Function fldNameExist"

    fldNameExist = False                '�װ�ð���i�����l�j���
    
    For i = 1 To fldCnt                 '̨���ސ���ٰ��
        If fldName = fldNames(i) Then   '������̨���ޖ��ƈ�v������̂��������ꍇ
            fldNameExist = True         '����ð�����
            Exit For                    'ٰ�߂𔲂���
        End If
    Next
    

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




'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ003�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ003 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
Public Function DBDRV_GetTBCMJ003(records() As typ_TBCMJ003, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ003"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ003 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .POSITION = rs("POSITION")       ' �ʒu
            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
            .TRANCOND = rs("TRANCOND")       ' ��������
            .TRANCNT = rs("TRANCNT")         ' ������
            .SMPLNO = rs("SMPLNO")           ' �T���v���m��
            .SMPLUMU = rs("SMPLUMU")         ' �T���v���L��
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .GOUKI = rs("GOUKI")             ' ���@
''''            .OIMEAS1 = rs("OIMEAS1")         ' �n������l�P
''''            .OIMEAS2 = rs("OIMEAS2")         ' �n������l�Q
''''            .OIMEAS3 = rs("OIMEAS3")         ' �n������l�R
''''            .OIMEAS4 = rs("OIMEAS4")         ' �n������l�S
''''            .OIMEAS5 = rs("OIMEAS5")         ' �n������l�T
''''            .ORGRES = rs("ORGRES")           ' �n�q�f����
'OI_NULL�Ή��@2005/03/03 TUKU START --------------------------------------------------
            If IsNull(rs("OIMEAS1")) = False Then .OIMEAS1 = rs("OIMEAS1") Else .OIMEAS1 = -1  '�n������l1
            If IsNull(rs("OIMEAS2")) = False Then .OIMEAS2 = rs("OIMEAS2") Else .OIMEAS2 = -1  '�n������l2
            If IsNull(rs("OIMEAS3")) = False Then .OIMEAS3 = rs("OIMEAS3") Else .OIMEAS3 = -1  '�n������l3
            If IsNull(rs("OIMEAS4")) = False Then .OIMEAS4 = rs("OIMEAS4") Else .OIMEAS4 = -1  '�n������l4
            If IsNull(rs("OIMEAS5")) = False Then .OIMEAS5 = rs("OIMEAS5") Else .OIMEAS5 = -1  '�n������l5
            If IsNull(rs("ORGRES")) = False Then .ORGRES = rs("ORGRES") Else .ORGRES = -1    ' �n�q�f����
'OI_NULL�Ή��@2005/03/03 TUKU END   --------------------------------------------------
            
            .SETDTM = rs("SETDTM")           ' �ݒ����
            .EFFECTTM = rs("EFFECTTM")       ' �L������
            .FTIRMETH = rs("FTIRMETH")       ' �e�s�h�q���֎�
            .YCOEF = rs("YCOEF")             ' �e�s�h�q���Z���i�x�ؕЁj
            .XCOEF = rs("XCOEF")             ' �e�s�h�q���Z���i�w�W���j
            .AVE = rs("AVE")                 ' �`�u�d
            .SIGMA = rs("SIGMA")             ' �Ёi�V�O�}�j
            .FTIRCONV = rs("FTIRCONV")       ' �e�s�h�q���Z
            .INSPECTWAY = rs("INSPECTWAY")   ' �������@
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
    rs.Close

    DBDRV_GetTBCMJ003 = FUNCTION_RETURN_SUCCESS
End Function

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ004�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ004 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
Public Function DBDRV_GetTBCMJ004(records() As typ_TBCMJ004, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, CSMEAS, PRE70P, INSPECTWAY, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ004"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ004 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .POSITION = rs("POSITION")       ' �ʒu
            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
            .TRANCOND = rs("TRANCOND")       ' ��������
            .TRANCNT = rs("TRANCNT")         ' ������
            .SMPLNO = rs("SMPLNO")           ' �T���v���m��
            .SMPLUMU = rs("SMPLUMU")         ' �T���v���L��
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .GOUKI = rs("GOUKI")             ' ���@
''''            .CSMEAS = rs("CSMEAS")           ' Cs�����l
''''            .PRE70P = rs("PRE70P")           ' �V�O������l
'OI_NULL�Ή��@2005/03/03 TUKU START --------------------------------------------------
            If IsNull(rs("CSMEAS")) = False Then .CSMEAS = rs("CSMEAS") Else .CSMEAS = -1  ' Cs�����l
            If IsNull(rs("PRE70P")) = False Then .PRE70P = rs("PRE70P") Else .PRE70P = -1  ' �V�O������l
'OI_NULL�Ή��@2005/03/03 TUKU START --------------------------------------------------
            .INSPECTWAY = rs("INSPECTWAY")   ' �������@
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

    DBDRV_GetTBCMJ004 = FUNCTION_RETURN_SUCCESS
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
'����      :2001/08/24�쐬�@�쑺
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

'�T�v      :�����œn���ꂽ���R�[�h��TBCMJ003�ɒǉ�����
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :record        ,I  ,typ_cmjc001c_Disp ,���o���R�[�h
'          :CRYNUM        ,I  ,String       ,�����ԍ�
'          :TSTAFFID      ,I  ,String       ,�o�^�Ј�ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�����񐔂̓e�[�u����̍ő�l+1�Ƃ���B
'����      :2001/06/22(Fri)�쐬�@����

Public Function DBDRV_Getcmjc001c_Exec(record As typ_cmjc001c_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN

Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL�x�[�X����
Dim sqlWhere As String  'SQLWhere����
Dim sqlGroup As String  'SQLGroup����
Dim SetDate As Variant  '�ݒ����

'    CRYNUM             �����ԍ��@�ˈ���
'    TRANCNT         �@ �����񐔁@�ˍő�
'   TSTAFFID            �o�^�Ј�ID�@�ˈ���
 '   REGDATE �@�@�@     �o�^���t�@��SYSDATE
 '   KSTAFFID           �X�V�Ј�ID�@��" "
 '   UPDDATE            �X�V���t�@��SYSDATE
 '   SENDFLAG           ���M�t���O�@��"0"
 '   SENDDATE           ���M���t�@��SYSDATE
    
    DBDRV_Getcmjc001c_Exec = FUNCTION_RETURN_FAILURE

    ''�ݒ����

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001c_SQL.bas -- Function DBDRV_Getcmjc001c_Exec"

    SetDate = Format$(record.SETDTM, "yyyy-mm-dd hh:mm:ss")

    ''SQL��g�ݗ��Ă�
    sqlBase = "Insert into TBCMJ003 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', " & record.OIMEAS1 & ", " & _
               record.OIMEAS2 & ", " & record.OIMEAS3 & ", " & record.OIMEAS4 & ", " & record.OIMEAS5 & ", " & record.ORGRES & ", " & _
               "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss'), " & record.EFFECTTM & ", '" & record.FTIRMETH & "', " & record.YCOEF & ", " & record.XCOEF & ", " & _
               record.AVE & ", " & record.SIGMA & ", " & record.FTIRCONV & ", '" & record.INSPECTWAY & "', " & record.JudgData & ", '" & TSTAFFID & "', " & _
               "SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ003 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
    
    ''SQL�̎��s
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001c_Exec = FUNCTION_RETURN_SUCCESS
    

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
