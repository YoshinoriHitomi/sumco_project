Attribute VB_Name = "s_cmbc022_SQL"
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
Public Type typ_cmjc001c_Disp
   ' CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
   ' TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long 6���Ή� 2007/05/28 SETsw kubota
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

'(2002/07 DBDRV_GetTBCME018���ړ�)
'�t�B�[���h�������p
Dim fldNames() As String    '��rs�Ɋ܂܂��t�B�[���h���ێ��z��
Dim fldCnt As Integer       '��rs�Ɋ܂܂��t�B�[���h��

'�g�p���Ă��Ȃ����ߍ폜 2011/08/23 SETsw kubota
''------------------------------------------------
'' DB�A�N�Z�X�֐�
''------------------------------------------------
'
''�T�v      :�e�[�u���uTBCMJ003�v��������ɂ��������R�[�h�𒊏o����
''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''          :records()     ,O  ,typ_cmjc001c_Disp ,���o���R�[�h
''          :SPLNUMs()     ,I  ,Integer      ,���o�����z��(�T���v��No)
''          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
''����      :
''����      :2001/06/20(Wed)�쐬�@����
'Public Function DBDRV_Getcmjc001c_Disp(records() As typ_cmjc001c_Disp, SPLNUMs() As Integer) As FUNCTION_RETURN
'Dim sql As String       'SQL�S��
'Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
'Dim sqlWhere As String  'SQL��WHERE����
'Dim sqlGroup As String  'SQL��GROUP����
'Dim sqlOrder As String  'SQL��Order����
'Dim rs As OraDynaset    'RecordSet
'Dim recCnt As Long      '���R�[�h��
'Dim i As Long           '���[�v�J�E���g
'
'    DBDRV_Getcmjc001c_Disp = FUNCTION_RETURN_FAILURE
'
'    ''SQL��g�ݗ��Ă�
'
'    '�G���[�n���h���̐ݒ�
'    On Error GoTo proc_err
'    gErr.Push "s_cmzcF_cmjc001c_SQL.bas -- Function DBDRV_Getcmjc001c_Disp"
'
'    sqlBase = "Select POSITION, SMPKBN, TRANCOND, MAX(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
'              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
'              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA "
'    sqlBase = sqlBase & "From TBCMJ003"
'    ''���o����(�����NO)�̎��o��
'     sqlWhere = "Where SMPLNO in ("
'    For i = 1 To UBound(SPLNUMs)
'        sqlWhere = sqlWhere & "'" & SPLNUMs(i) & "'"
'        If i < UBound(SPLNUMs) Then
'            sqlWhere = sqlWhere & ", "
'        End If
'    Next
'    sqlWhere = sqlWhere & ") "
'    sqlGroup = "GROUP BY CRYNUM, POSITION, SMPKBN, TRANCOND "
'    sqlOrder = "ORDER BY POSITION"
'    sql = sqlBase & sqlWhere & sqlGroup & sqlOrder
'
'    ''�f�[�^�𒊏o����
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'    If rs Is Nothing Then
'        ReDim records(0)
'        DBDRV_Getcmjc001c_Disp = FUNCTION_RETURN_FAILURE
'        GoTo proc_exit
'    End If
'
'    ''���o���ʂ��i�[����
'    recCnt = rs.RecordCount
'    ReDim records(recCnt)
'    For i = 1 To recCnt
'        With records(i)
'            .POSITION = rs("POSITION")       ' �ʒu
'            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
'            .TRANCOND = rs("TRANCOND")       ' ��������
'            .SMPLNO = rs("SMPLNO")           ' �T���v���m��
'            .SMPLUMU = rs("SMPLUMU")         ' �T���v���L��
'            .hinban = rs("HINBAN")           ' �i��
'            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
'            .factory = rs("FACTORY")         ' �H��
'            .opecond = rs("OPECOND")         ' ���Ə���
'            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
'            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
'            .GOUKI = rs("GOUKI")             ' ���@
'            .OIMEAS1 = rs("OIMEAS1")         ' �n������l�P
'            .OIMEAS2 = rs("OIMEAS2")         ' �n������l�Q
'            .OIMEAS3 = rs("OIMEAS3")         ' �n������l�R
'            .OIMEAS4 = rs("OIMEAS4")         ' �n������l�S
'            .OIMEAS5 = rs("OIMEAS5")         ' �n������l�T
'            .ORGRES = rs("ORGRES")           ' �n�q�f����
'            .SETDTM = rs("SETDTM")           ' �ݒ����
'            .EFFECTTM = rs("EFFECTTM")       ' �L������
'            .FTIRMETH = rs("FTIRMETH")       ' �e�s�h�q���֎�
'            .YCOEF = rs("YCOEF")             ' �e�s�h�q���Z���i�x�ؕЁj
'            .XCOEF = rs("XCOEF")             ' �e�s�h�q���Z���i�w�W���j
'            .AVE = rs("AVE")                 ' �`�u�d
'            .SIGMA = rs("SIGMA")             ' �Ёi�V�O�}�j
'            .FTIRCONV = rs("FTIRCONV")       ' �e�s�h�q���Z
'            .INSPECTWAY = rs("INSPECTWAY")   ' �������@
'            .JudgData = rs("JUDGDATA")       ' �����Ώےl
'        End With
'        rs.MoveNext
'    Next
'    rs.Close
'
'    DBDRV_Getcmjc001c_Disp = FUNCTION_RETURN_SUCCESS
'
'proc_exit:
'    '�I��
'    gErr.Pop
'    Exit Function
'
'proc_err:
'    '�G���[�n���h��
'    Debug.Print "====== Error SQL ======"
'    Debug.Print sql
'    gErr.HandleError
'    Resume proc_exit
'End Function

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
    '' OI_NULL�Ή��@2005/03/07 TUKU START �R���a����̈˗��ŕύX���~--------------------------------------------------------------------
''''    sqlBase = "Insert into TBCMJ003 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
''''              " PROCCODE, GOUKI, OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, SETDTM, EFFECTTM, FTIRMETH, YCOEF, XCOEF," & _
''''              " AVE, SIGMA, FTIRCONV, INSPECTWAY, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
''''    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
''''               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.factory & "', '" & _
''''               record.opecond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', "
''''                If (record.OIMEAS1 = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.OIMEAS1 & ", "    'OI����l1
''''                If (record.OIMEAS2 = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.OIMEAS2 & ", "    'OI����l2
''''                If (record.OIMEAS3 = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.OIMEAS3 & ", "    'OI����l2
''''                If (record.OIMEAS4 = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.OIMEAS4 & ", "    'OI����l3
''''                If (record.OIMEAS5 = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.OIMEAS5 & ", "    'OI����l4
''''                If (record.ORGRES = -1) Then sqlBase = sqlBase & " NULL , " Else sqlBase = sqlBase & record.ORGRES & ", "      'ORG
''''    sqlBase = sqlBase & "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss'), " & record.EFFECTTM & ", '" & record.FTIRMETH & "', " & record.YCOEF & ", " & record.XCOEF & ", " & _
''''               record.AVE & ", " & record.SIGMA & ", " & record.FTIRCONV & ", '" & record.INSPECTWAY & "', " & record.JudgData & ", '" & TSTAFFID & "', " & _
''''               "SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ003 "
''''    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'''''    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
''''    sql = sqlBase & sqlWhere & sqlGroup
    '' OI_NULL�Ή��@2005/03/07 TUKU END   --------------------------------------------------------------------
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


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMB014�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB014 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMB014_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMB014(records() As typ_TBCMB014, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select GOUKI, INPDATE, FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG," & _
              " SENDDATE "
    sqlBase = sqlBase & "From TBCMB014"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB014 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .GOUKI = rs("GOUKI")             ' ���@
            .INPDATE = rs("INPDATE")         ' ���t
            .FTIRFZI = rs("FTIRFZI")         ' FTIR�iFZ)
            .FTIRCZH = rs("FTIRCZH")         ' FTIR�iCZ���j
            .FTIRCZC = rs("FTIRCZC")         ' FTIR�iCZ���j
            .MS1FZ = rs("MS1FZ")             ' ����T���v��1�iFZ)
            .MS1CZ1 = rs("MS1CZ1")           ' ����T���v��1�iCZ-1)
            .MS1CZ2 = rs("MS1CZ2")           ' ����T���v��1�iCZ-2)
            .MS2FZ = rs("MS2FZ")             ' ����T���v��2�iFZ)
            .MS2CZ1 = rs("MS2CZ1")           ' ����T���v��2�iCZ-1)
            .MS2CZ2 = rs("MS2CZ2")           ' ����T���v��2�iCZ-2)
            .MS3FZ = rs("MS3FZ")             ' ����T���v��3�iFZ)
            .MS3CZ1 = rs("MS3CZ1")           ' ����T���v��3�iCZ-1)
            .MS3CZ2 = rs("MS3CZ2")           ' ����T���v��3�iCZ-2)
            .MS4FZ = rs("MS4FZ")             ' ����T���v��4�iFZ)
            .MS4CZ1 = rs("MS4CZ1")           ' ����T���v��4�iCZ-1)
            .MS4CZ2 = rs("MS4CZ2")           ' ����T���v��4�iCZ-2)
            .MS5FZ = rs("MS5FZ")             ' ����T���v��5�iFZ)
            .MS5CZ1 = rs("MS5CZ1")           ' ����T���v��5�iCZ-1)
            .MS5CZ2 = rs("MS5CZ2")           ' ����T���v��5�iCZ-2)
            .MSAVEFZ = rs("MSAVEFZ")         ' ���蕽�ρiFZ�j
            .MSAVECZ1 = rs("MSAVECZ1")       ' ���蕽�ρiCZ-1�j
            .MSAVECZ2 = rs("MSAVECZ2")       ' ���蕽�ρiCZ-2�j
            .MSSGFZ = rs("MSSGFZ")           ' ����ЁiFZ�j
            .MSSGCZ1 = rs("MSSGCZ1")         ' ����ЁiCZ-1�j
            .MSSGCZ2 = rs("MSSGCZ2")         ' ����ЁiCZ-2�j
            .MSPSGFZ = rs("MSPSGFZ")         ' ����AVE+�ЁiFZ�j
            .MSPSGCZ1 = rs("MSPSGCZ1")       ' ����AVE+�ЁiCZ-1�j
            .MSPSGCZ2 = rs("MSPSGCZ2")       ' ����AVE+�ЁiCZ-2�j
            .MSNSGFZ = rs("MSNSGFZ")         ' ����AVE-�ЁiFZ�j
            .MSNSGCZ1 = rs("MSNSGCZ1")       ' ����AVE-�ЁiCZ-1�j
            .MSNSGCZ2 = rs("MSNSGCZ2")       ' ����AVE-�ЁiCZ-2�j
            .MINFZ = rs("MINFZ")             ' MIN�iFZ�j
            .MINCZ1 = rs("MINCZ1")           ' MIN�iCZ-1�j
            .MINCZ2 = rs("MINCZ2")           ' MIN�iCZ-2�j
            .MAXFZ = rs("MAXFZ")             ' MAX�iFZ�j
            .MAXCZ1 = rs("MAXCZ1")           ' MAX�iCZ-1�j
            .MAXCZ2 = rs("MAXCZ2")           ' MAX�iCZ-2�j
            .SGCK1FZ = rs("SGCK1FZ")         ' ��ck�T���v��1�iFZ)
            .SGCK1CZ1 = rs("SGCK1CZ1")       ' ��ck�T���v��1�iCZ-1)
            .SGCK1CZ2 = rs("SGCK1CZ2")       ' ��ck�T���v��1�iCZ-2)
            .SGCK2FZ = rs("SGCK2FZ")         ' ��ck�T���v��2�iFZ)
            .SGCK2CZ1 = rs("SGCK2CZ1")       ' ��ck�T���v��2�iCZ-1)
            .SGCK2CZ2 = rs("SGCK2CZ2")       ' ��ck�T���v��2�iCZ-2)
            .SGCK3FZ = rs("SGCK3FZ")         ' ��ck�T���v��3�iFZ)
            .SGCK3CZ1 = rs("SGCK3CZ1")       ' ��ck�T���v��3�iCZ-1)
            .SGCK3CZ2 = rs("SGCK3CZ2")       ' ��ck�T���v��3�iCZ-2)
            .SGCK4FZ = rs("SGCK4FZ")         ' ��ck�T���v��4�iFZ)
            .SGCK4CZ1 = rs("SGCK4CZ1")       ' ��ck�T���v��4�iCZ-1)
            .SGCK4CZ2 = rs("SGCK4CZ2")       ' ��ck�T���v��4�iCZ-2)
            .SGCK5FZ = rs("SGCK5FZ")         ' ��ck�T���v��5�iFZ)
            .SGCK5CZ1 = rs("SGCK5CZ1")       ' ��ck�T���v��5�iCZ-1)
            .SGCK5CZ2 = rs("SGCK5CZ2")       ' ��ck�T���v��5�iCZ-2)
            .SGCKDFZ = rs("SGCKDFZ")         ' ��ck�f�[�^���iFZ�j
            .SGCKDCZ1 = rs("SGCKDCZ1")       ' ��ck�f�[�^���iCZ-1�j
            .SGCKDCZ2 = rs("SGCKDCZ2")       ' ��ck�f�[�^���iCZ-2�j
            .SGCKAFZ = rs("SGCKAFZ")         ' ��ck���ρiFZ�j
            .SGCKAACZ1 = rs("SGCKAACZ1")     ' ��ck���ρiCZ-1�j
            .SGCKACZ2 = rs("SGCKACZ2")       ' ��ck���ρiCZ-2�j
            .SGNFZ = rs("SGNFZ")             ' ��ck�ЁiFZ�j
            .SGNCZ1 = rs("SGNCZ1")           ' ��ck�� CZ-1�j
            .SGNCZ2 = rs("SGNCZ2")           ' ��ck�ЁiCZ-2�j
            .FTIRFZ = rs("FTIRFZ")           ' FTIR���Z�iFZ�j
            .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR���Z�iCZ-1�j
            .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR���Z�iCZ-2�j
            .EFFECTTM = rs("EFFECTTM")       ' �L������
            .YCOEF = rs("YCOEF")             ' �e�s�h�q���Z���i�x�ؕЁj
            .XCOEF = rs("XCOEF")             ' �e�s�h�q���Z���i�w�W���j
            .RSQUARE = rs("RSQUARE")         ' �q�Q��
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

    DBDRV_GetTBCMB014 = FUNCTION_RETURN_SUCCESS
End Function


'�T�v      :�����̃t�B�[���h����fldNames()�z��Ɋ܂܂�Ă��邩�ǂ����̔���B
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :fldName       ,I  ,typ_TBCME018 ,���o���R�[�h
'          :�߂�l        ,O  ,Boolean      ,True:�݂�^False�F����
'����      :
'����      :2001/06/27�쐬�@�쑺  (2002/07 DBDRV_GetTBCME018���ړ�)

Private Function fldNameExist(fldName As String) As Boolean
    Dim sql         As String           'SQL�S��
    Dim i As Integer                    'ٰ�߶���


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME018_SQL.bas -- Function fldNameExist"

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

'�T�v      :�e�[�u���uTBCME036�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME036 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCME036_SQL.bas���ړ�)
'          :06/04/11 ooba�@�֐����ύX <DBDRV_GetTBCME036> �� <DBDRV_GetTBCME036_cmbc022>
Public Function DBDRV_GetTBCME036_cmbc022(records() As typ_TBCME036, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, EPDSETCH, EPDUP, CUTUNIT, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO," & _
              " STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME036"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME036_cmbc022 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
'NULL�Ή� ----- START ----- 2003/12/10
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .hinban = rs("HINBAN")           ' �i��
            .mnorevno = rs("MNOREVNO")       ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .EPDSETCH = rs("EPDSETCH")       ' EPD�@�I���G�b�`
            .EPDUP = fncNullCheck(rs("EPDUP"))             ' EPD�@���
            .CUTUNIT = fncNullCheck(rs("CUTUNIT"))         ' �J�b�g�P��
            .IFKBN = rs("IFKBN")             ' �h�^�e�敪
            .SYORIKBN = rs("SYORIKBN")       ' �����敪
            .SPECRRNO = rs("SPECRRNO")       ' �d�l�o�^�˗��ԍ�
            .SXLMCNO = rs("SXLMCNO")         ' �r�w�k��������ԍ�
            .WFMCNO = rs("WFMCNO")           ' �v�e��������ԍ�
            .StaffID = rs("STAFFID")         ' �Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close
'NULL�Ή� -----  END  ----- 2003/12/10

    DBDRV_GetTBCME036_cmbc022 = FUNCTION_RETURN_SUCCESS
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

'�T�v      :�e�[�u���uTBCMJ005�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ005 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMJ005_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMJ005(records() As typ_TBCMJ005, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, CALCMAX, CALCAVE, MEAS1, MEAS2, MEAS3, MEAS4," & _
              " MEAS5, MEAS6, MEAS7, MEAS8, MEAS9, MEAS10, MEAS11, MEAS12, MEAS13, MEAS14, MEAS15, MEAS16, MEAS17, MEAS18," & _
              " MEAS19, MEAS20, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ005"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ005 = FUNCTION_RETURN_FAILURE
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
            .MEASMETH = rs("MEASMETH")       ' ������@
            .MEASSPOT = rs("MEASSPOT")       ' ����_
            .MAG = rs("MAG")                 ' �{��
            .HTPRC = rs("HTPRC")             ' �M�������@
            .KKSP = rs("KKSP")               ' �������ב���ʒu
            .KKSET = rs("KKSET")             ' �������ב�������{�I��ET��@�@char(1)�{number(2)
            .CALCMAX = rs("CALCMAX")         ' �v�Z���� Max
            .CALCAVE = rs("CALCAVE")         ' �v�Z���� Ave
            .MEAS1 = rs("MEAS1")             ' ����l�P
            .MEAS2 = rs("MEAS2")             ' ����l�Q
            .MEAS3 = rs("MEAS3")             ' ����l�R
            .MEAS4 = rs("MEAS4")             ' ����l�S
            .MEAS5 = rs("MEAS5")             ' ����l�T
            .MEAS6 = rs("MEAS6")             ' ����l�U
            .MEAS7 = rs("MEAS7")             ' ����l�V
            .MEAS8 = rs("MEAS8")             ' ����l�W
            .MEAS9 = rs("MEAS9")             ' ����l�X
            .MEAS10 = rs("MEAS10")           ' ����l�P�O
            .MEAS11 = rs("MEAS11")           ' ����l�P�P
            .MEAS12 = rs("MEAS12")           ' ����l�P�Q
            .MEAS13 = rs("MEAS13")           ' ����l�P�R
            .MEAS14 = rs("MEAS14")           ' ����l�P�S
            .MEAS15 = rs("MEAS15")           ' ����l�P�T
            .MEAS16 = rs("MEAS16")           ' ����l�P�U
            .MEAS17 = rs("MEAS17")           ' ����l�P�V
            .MEAS18 = rs("MEAS18")           ' ����l�P�W
            .MEAS19 = rs("MEAS19")           ' ����l�P�X
            .MEAS20 = rs("MEAS20")           ' ����l�Q�O
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

    DBDRV_GetTBCMJ005 = FUNCTION_RETURN_SUCCESS
End Function


'�T�v      :�e�[�u���uTBCMJ006�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ006 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCMJ006_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMJ006(records() As typ_TBCMJ006, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MSRSDEN, MSRSLDL, MSRSDVD2, MS01LDL1, MS01LDL2, MS01LDL3, MS01LDL4, MS01LDL5, MS01DEN1, MS01DEN2," & _
              " MS01DEN3, MS01DEN4, MS01DEN5, MS02LDL1, MS02LDL2, MS02LDL3, MS02LDL4, MS02LDL5, MS02DEN1, MS02DEN2, MS02DEN3," & _
              " MS02DEN4, MS02DEN5, MS03LDL1, MS03LDL2, MS03LDL3, MS03LDL4, MS03LDL5, MS03DEN1, MS03DEN2, MS03DEN3, MS03DEN4," & _
              " MS03DEN5, MS04LDL1, MS04LDL2, MS04LDL3, MS04LDL4, MS04LDL5, MS04DEN1, MS04DEN2, MS04DEN3, MS04DEN4, MS04DEN5," & _
              " MS05LDL1, MS05LDL2, MS05LDL3, MS05LDL4, MS05LDL5, MS05DEN1, MS05DEN2, MS05DEN3, MS05DEN4, MS05DEN5, MS06LDL1," & _
              " MS06LDL2, MS06LDL3, MS06LDL4, MS06LDL5, MS06DEN1, MS06DEN2, MS06DEN3, MS06DEN4, MS06DEN5, MS07LDL1, MS07LDL2," & _
              " MS07LDL3, MS07LDL4, MS07LDL5, MS07DEN1, MS07DEN2, MS07DEN3, MS07DEN4, MS07DEN5, MS08LDL1, MS08LDL2, MS08LDL3," & _
              " MS08LDL4, MS08LDL5, MS08DEN1, MS08DEN2, MS08DEN3, MS08DEN4, MS08DEN5, MS09LDL1, MS09LDL2, MS09LDL3, MS09LDL4," & _
              " MS09LDL5, MS09DEN1, MS09DEN2, MS09DEN3, MS09DEN4, MS09DEN5, MS10LDL1, MS10LDL2, MS10LDL3, MS10LDL4, MS10LDL5," & _
              " MS10DEN1, MS10DEN2, MS10DEN3, MS10DEN4, MS10DEN5, MS11LDL1, MS11LDL2, MS11LDL3, MS11LDL4, MS11LDL5, MS11DEN1," & _
              " MS11DEN2, MS11DEN3, MS11DEN4, MS11DEN5, MS12LDL1, MS12LDL2, MS12LDL3, MS12LDL4, MS12LDL5, MS12DEN1, MS12DEN2," & _
              " MS12DEN3, MS12DEN4, MS12DEN5, MS13LDL1, MS13LDL2, MS13LDL3, MS13LDL4, MS13LDL5, MS13DEN1, MS13DEN2, MS13DEN3," & _
              " MS13DEN4, MS13DEN5, MS14LDL1, MS14LDL2, MS14LDL3, MS14LDL4, MS14LDL5, MS14DEN1, MS14DEN2, MS14DEN3, MS14DEN4," & _
              " MS14DEN5, MS15LDL1, MS15LDL2, MS15LDL3, MS15LDL4, MS15LDL5, MS15DEN1, MS15DEN2, MS15DEN3, MS15DEN4, MS15DEN5," & _
              " MS01DVD2,  MS02DVD2 , MS03DVD2 , MS04DVD2 , MS05DVD2 , TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ006"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ006 = FUNCTION_RETURN_FAILURE
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
            .MSRSDEN = rs("MSRSDEN")         ' ���茋�� Den
            .MSRSLDL = rs("MSRSLDL")         ' ���茋�� L/DL
            .MSRSDVD2 = rs("MSRSDVD2")       ' ���茋�� DVD2
            .MS01LDL1 = rs("MS01LDL1")       ' ����l01 L/DL1
            .MS01LDL2 = rs("MS01LDL2")       ' ����l01 L/DL2
            .MS01LDL3 = rs("MS01LDL3")       ' ����l01 L/DL3
            .MS01LDL4 = rs("MS01LDL4")       ' ����l01 L/DL4
            .MS01LDL5 = rs("MS01LDL5")       ' ����l01 L/DL5
            .MS01DEN1 = rs("MS01DEN1")       ' ����l01 Den1
            .MS01DEN2 = rs("MS01DEN2")       ' ����l01 Den2
            .MS01DEN3 = rs("MS01DEN3")       ' ����l01 Den3
            .MS01DEN4 = rs("MS01DEN4")       ' ����l01 Den4
            .MS01DEN5 = rs("MS01DEN5")       ' ����l01 Den5
            .MS02LDL1 = rs("MS02LDL1")       ' ����l02 L/DL1
            .MS02LDL2 = rs("MS02LDL2")       ' ����l02 L/DL2
            .MS02LDL3 = rs("MS02LDL3")       ' ����l02 L/DL3
            .MS02LDL4 = rs("MS02LDL4")       ' ����l02 L/DL4
            .MS02LDL5 = rs("MS02LDL5")       ' ����l02 L/DL5
            .MS02DEN1 = rs("MS02DEN1")       ' ����l02 Den1
            .MS02DEN2 = rs("MS02DEN2")       ' ����l02 Den2
            .MS02DEN3 = rs("MS02DEN3")       ' ����l02 Den3
            .MS02DEN4 = rs("MS02DEN4")       ' ����l02 Den4
            .MS02DEN5 = rs("MS02DEN5")       ' ����l02 Den5
            .MS03LDL1 = rs("MS03LDL1")       ' ����l03 L/DL1
            .MS03LDL2 = rs("MS03LDL2")       ' ����l03 L/DL2
            .MS03LDL3 = rs("MS03LDL3")       ' ����l03 L/DL3
            .MS03LDL4 = rs("MS03LDL4")       ' ����l03 L/DL4
            .MS03LDL5 = rs("MS03LDL5")       ' ����l03 L/DL5
            .MS03DEN1 = rs("MS03DEN1")       ' ����l03 Den1
            .MS03DEN2 = rs("MS03DEN2")       ' ����l03 Den2
            .MS03DEN3 = rs("MS03DEN3")       ' ����l03 Den3
            .MS03DEN4 = rs("MS03DEN4")       ' ����l03 Den4
            .MS03DEN5 = rs("MS03DEN5")       ' ����l03 Den5
            .MS04LDL1 = rs("MS04LDL1")       ' ����l04 L/DL1
            .MS04LDL2 = rs("MS04LDL2")       ' ����l04 L/DL2
            .MS04LDL3 = rs("MS04LDL3")       ' ����l04 L/DL3
            .MS04LDL4 = rs("MS04LDL4")       ' ����l04 L/DL4
            .MS04LDL5 = rs("MS04LDL5")       ' ����l04 L/DL5
            .MS04DEN1 = rs("MS04DEN1")       ' ����l04 Den1
            .MS04DEN2 = rs("MS04DEN2")       ' ����l04 Den2
            .MS04DEN3 = rs("MS04DEN3")       ' ����l04 Den3
            .MS04DEN4 = rs("MS04DEN4")       ' ����l04 Den4
            .MS04DEN5 = rs("MS04DEN5")       ' ����l04 Den5
            .MS05LDL1 = rs("MS05LDL1")       ' ����l05 L/DL1
            .MS05LDL2 = rs("MS05LDL2")       ' ����l05 L/DL2
            .MS05LDL3 = rs("MS05LDL3")       ' ����l05 L/DL3
            .MS05LDL4 = rs("MS05LDL4")       ' ����l05 L/DL4
            .MS05LDL5 = rs("MS05LDL5")       ' ����l05 L/DL5
            .MS05DEN1 = rs("MS05DEN1")       ' ����l05 Den1
            .MS05DEN2 = rs("MS05DEN2")       ' ����l05 Den2
            .MS05DEN3 = rs("MS05DEN3")       ' ����l05 Den3
            .MS05DEN4 = rs("MS05DEN4")       ' ����l05 Den4
            .MS05DEN5 = rs("MS05DEN5")       ' ����l05 Den5
            .MS06LDL1 = rs("MS06LDL1")       ' ����l06 L/DL1
            .MS06LDL2 = rs("MS06LDL2")       ' ����l06 L/DL2
            .MS06LDL3 = rs("MS06LDL3")       ' ����l06 L/DL3
            .MS06LDL4 = rs("MS06LDL4")       ' ����l06 L/DL4
            .MS06LDL5 = rs("MS06LDL5")       ' ����l06 L/DL5
            .MS06DEN1 = rs("MS06DEN1")       ' ����l06 Den1
            .MS06DEN2 = rs("MS06DEN2")       ' ����l06 Den2
            .MS06DEN3 = rs("MS06DEN3")       ' ����l06 Den3
            .MS06DEN4 = rs("MS06DEN4")       ' ����l06 Den4
            .MS06DEN5 = rs("MS06DEN5")       ' ����l06 Den5
            .MS07LDL1 = rs("MS07LDL1")       ' ����l07 L/DL1
            .MS07LDL2 = rs("MS07LDL2")       ' ����l07 L/DL2
            .MS07LDL3 = rs("MS07LDL3")       ' ����l07 L/DL3
            .MS07LDL4 = rs("MS07LDL4")       ' ����l07 L/DL4
            .MS07LDL5 = rs("MS07LDL5")       ' ����l07 L/DL5
            .MS07DEN1 = rs("MS07DEN1")       ' ����l07 Den1
            .MS07DEN2 = rs("MS07DEN2")       ' ����l07 Den2
            .MS07DEN3 = rs("MS07DEN3")       ' ����l07 Den3
            .MS07DEN4 = rs("MS07DEN4")       ' ����l07 Den4
            .MS07DEN5 = rs("MS07DEN5")       ' ����l07 Den5
            .MS08LDL1 = rs("MS08LDL1")       ' ����l08 L/DL1
            .MS08LDL2 = rs("MS08LDL2")       ' ����l08 L/DL2
            .MS08LDL3 = rs("MS08LDL3")       ' ����l08 L/DL3
            .MS08LDL4 = rs("MS08LDL4")       ' ����l08 L/DL4
            .MS08LDL5 = rs("MS08LDL5")       ' ����l08 L/DL5
            .MS08DEN1 = rs("MS08DEN1")       ' ����l08 Den1
            .MS08DEN2 = rs("MS08DEN2")       ' ����l08 Den2
            .MS08DEN3 = rs("MS08DEN3")       ' ����l08 Den3
            .MS08DEN4 = rs("MS08DEN4")       ' ����l08 Den4
            .MS08DEN5 = rs("MS08DEN5")       ' ����l08 Den5
            .MS09LDL1 = rs("MS09LDL1")       ' ����l09 L/DL1
            .MS09LDL2 = rs("MS09LDL2")       ' ����l09 L/DL2
            .MS09LDL3 = rs("MS09LDL3")       ' ����l09 L/DL3
            .MS09LDL4 = rs("MS09LDL4")       ' ����l09 L/DL4
            .MS09LDL5 = rs("MS09LDL5")       ' ����l09 L/DL5
            .MS09DEN1 = rs("MS09DEN1")       ' ����l09 Den1
            .MS09DEN2 = rs("MS09DEN2")       ' ����l09 Den2
            .MS09DEN3 = rs("MS09DEN3")       ' ����l09 Den3
            .MS09DEN4 = rs("MS09DEN4")       ' ����l09 Den4
            .MS09DEN5 = rs("MS09DEN5")       ' ����l09 Den5
            .MS10LDL1 = rs("MS10LDL1")       ' ����l10 L/DL1
            .MS10LDL2 = rs("MS10LDL2")       ' ����l10 L/DL2
            .MS10LDL3 = rs("MS10LDL3")       ' ����l10 L/DL3
            .MS10LDL4 = rs("MS10LDL4")       ' ����l10 L/DL4
            .MS10LDL5 = rs("MS10LDL5")       ' ����l10 L/DL5
            .MS10DEN1 = rs("MS10DEN1")       ' ����l10 Den1
            .MS10DEN2 = rs("MS10DEN2")       ' ����l10 Den2
            .MS10DEN3 = rs("MS10DEN3")       ' ����l10 Den3
            .MS10DEN4 = rs("MS10DEN4")       ' ����l10 Den4
            .MS10DEN5 = rs("MS10DEN5")       ' ����l10 Den5
            .MS11LDL1 = rs("MS11LDL1")       ' ����l11 L/DL1
            .MS11LDL2 = rs("MS11LDL2")       ' ����l11 L/DL2
            .MS11LDL3 = rs("MS11LDL3")       ' ����l11 L/DL3
            .MS11LDL4 = rs("MS11LDL4")       ' ����l11 L/DL4
            .MS11LDL5 = rs("MS11LDL5")       ' ����l11 L/DL5
            .MS11DEN1 = rs("MS11DEN1")       ' ����l11 Den1
            .MS11DEN2 = rs("MS11DEN2")       ' ����l11 Den2
            .MS11DEN3 = rs("MS11DEN3")       ' ����l11 Den3
            .MS11DEN4 = rs("MS11DEN4")       ' ����l11 Den4
            .MS11DEN5 = rs("MS11DEN5")       ' ����l11 Den5
            .MS12LDL1 = rs("MS12LDL1")       ' ����l12 L/DL1
            .MS12LDL2 = rs("MS12LDL2")       ' ����l12 L/DL2
            .MS12LDL3 = rs("MS12LDL3")       ' ����l12 L/DL3
            .MS12LDL4 = rs("MS12LDL4")       ' ����l12 L/DL4
            .MS12LDL5 = rs("MS12LDL5")       ' ����l12 L/DL5
            .MS12DEN1 = rs("MS12DEN1")       ' ����l12 Den1
            .MS12DEN2 = rs("MS12DEN2")       ' ����l12 Den2
            .MS12DEN3 = rs("MS12DEN3")       ' ����l12 Den3
            .MS12DEN4 = rs("MS12DEN4")       ' ����l12 Den4
            .MS12DEN5 = rs("MS12DEN5")       ' ����l12 Den5
            .MS13LDL1 = rs("MS13LDL1")       ' ����l13 L/DL1
            .MS13LDL2 = rs("MS13LDL2")       ' ����l13 L/DL2
            .MS13LDL3 = rs("MS13LDL3")       ' ����l13 L/DL3
            .MS13LDL4 = rs("MS13LDL4")       ' ����l13 L/DL4
            .MS13LDL5 = rs("MS13LDL5")       ' ����l13 L/DL5
            .MS13DEN1 = rs("MS13DEN1")       ' ����l13 Den1
            .MS13DEN2 = rs("MS13DEN2")       ' ����l13 Den2
            .MS13DEN3 = rs("MS13DEN3")       ' ����l13 Den3
            .MS13DEN4 = rs("MS13DEN4")       ' ����l13 Den4
            .MS13DEN5 = rs("MS13DEN5")       ' ����l13 Den5
            .MS14LDL1 = rs("MS14LDL1")       ' ����l14 L/DL1
            .MS14LDL2 = rs("MS14LDL2")       ' ����l14 L/DL2
            .MS14LDL3 = rs("MS14LDL3")       ' ����l14 L/DL3
            .MS14LDL4 = rs("MS14LDL4")       ' ����l14 L/DL4
            .MS14LDL5 = rs("MS14LDL5")       ' ����l14 L/DL5
            .MS14DEN1 = rs("MS14DEN1")       ' ����l14 Den1
            .MS14DEN2 = rs("MS14DEN2")       ' ����l14 Den2
            .MS14DEN3 = rs("MS14DEN3")       ' ����l14 Den3
            .MS14DEN4 = rs("MS14DEN4")       ' ����l14 Den4
            .MS14DEN5 = rs("MS14DEN5")       ' ����l14 Den5
            .MS15LDL1 = rs("MS15LDL1")       ' ����l15 L/DL1
            .MS15LDL2 = rs("MS15LDL2")       ' ����l15 L/DL2
            .MS15LDL3 = rs("MS15LDL3")       ' ����l15 L/DL3
            .MS15LDL4 = rs("MS15LDL4")       ' ����l15 L/DL4
            .MS15LDL5 = rs("MS15LDL5")       ' ����l15 L/DL5
            .MS15DEN1 = rs("MS15DEN1")       ' ����l15 Den1
            .MS15DEN2 = rs("MS15DEN2")       ' ����l15 Den2
            .MS15DEN3 = rs("MS15DEN3")       ' ����l15 Den3
            .MS15DEN4 = rs("MS15DEN4")       ' ����l15 Den4
            .MS15DEN5 = rs("MS15DEN5")       ' ����l15 Den5
            'NULL �`�F�b�N
            If IsNull(rs("MS01DVD2")) = False Then
                .MS01DVD2 = rs("MS01DVD2")       ' ����l01 DVD2   2002/7/02 tuku
            Else
                .MS01DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS02DVD2")) = False Then
                .MS02DVD2 = rs("MS02DVD2")       ' ����l01 DVD2   2002/7/02 tuku
            Else
                .MS02DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS03DVD2")) = False Then
                .MS03DVD2 = rs("MS03DVD2")       ' ����l01 DVD2   2002/7/02 tuku
            Else
                .MS03DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS04DVD2")) = False Then
                .MS04DVD2 = rs("MS04DVD2")       ' ����l01 DVD2   2002/7/02 tuku
            Else
                .MS04DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS05DVD2")) = False Then
                .MS05DVD2 = rs("MS05DVD2")       ' ����l01 DVD2   2002/7/02 tuku
            Else
                .MS05DVD2 = DEF_PARAM_VALUE
            End If
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

    DBDRV_GetTBCMJ006 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ007�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ007 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCMJ007_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMJ007(records() As typ_TBCMJ007, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
              " SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ007"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ007 = FUNCTION_RETURN_FAILURE
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
            .MEAS1 = rs("MEAS1")             ' ����l�P
            .MEAS2 = rs("MEAS2")             ' ����l�Q
            .MEAS3 = rs("MEAS3")             ' ����l�R
            .MEAS4 = rs("MEAS4")             ' ����l�S
            .MEAS5 = rs("MEAS5")             ' ����l�T
            .MEASPEAK = rs("MEASPEAK")       ' ����l �s�[�N�l
            .CALCMEAS = rs("CALCMEAS")       ' �v�Z����
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

    DBDRV_GetTBCMJ007 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ008�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ008 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCMJ008_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMJ008(records() As typ_TBCMJ008, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN," & _
              " MEASMAX, MEASAVE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ008"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ008 = FUNCTION_RETURN_FAILURE
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
            .MEASMETH = rs("MEASMETH")       ' ������@
            .MEASSPOT = rs("MEASSPOT")       ' ����_
            .MAG = rs("MAG")                 ' �{��
            .HTPRC = rs("HTPRC")             ' �M�������@
            .KKSP = rs("KKSP")               ' �������ב���ʒu
            .KKSET = rs("KKSET")             ' �������ב�������{�I��ET��@�@char(1)�{number(2)
            .MEAS1 = rs("MEAS1")             ' ����l�P
            .MEAS2 = rs("MEAS2")             ' ����l�Q
            .MEAS3 = rs("MEAS3")             ' ����l�R
            .MEAS4 = rs("MEAS4")             ' ����l�S
            .MEAS5 = rs("MEAS5")             ' ����l�T
            .MEASMIN = rs("MEASMIN")         ' MIN
            .MEASMAX = rs("MEASMAX")         ' MAX
            .MEASAVE = rs("MEASAVE")         ' AVE
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

    DBDRV_GetTBCMJ008 = FUNCTION_RETURN_SUCCESS
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
            .Length = rs("LENGTH")           ' ����
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

'�T�v      :�e�[�u���uTBCMJ003�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ003 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMJ003_SQL.bas���ړ�)
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
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMJ004_SQL.bas���ړ�)
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
            .CSMEAS = rs("CSMEAS")           ' Cs�����l
            .PRE70P = rs("PRE70P")           ' �V�O������l
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

'�T�v      :�e�[�u���uTBCMJ002�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ002 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCMJ002_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMJ002(records() As typ_TBCMJ002, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
              " OPECOND, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID," & _
              " UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ002"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ002 = FUNCTION_RETURN_FAILURE
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
    rs.Close

    DBDRV_GetTBCMJ002 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ001�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ001 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCMJ001_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMJ001(records() As typ_TBCMJ001, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
              " OPECOND, GOUKI, MEASURE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ001"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ001 = FUNCTION_RETURN_FAILURE
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
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .GOUKI = rs("GOUKI")             ' ���@
            .MEASURE = rs("MEASURE")         ' ����l
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

    DBDRV_GetTBCMJ001 = FUNCTION_RETURN_SUCCESS
End Function


'///////////////////////////////////////////////////
' @(f)
' �@�\    : �Д����擾
'
' �Ԃ�l  : True  - ����
' �@�@�@    False - ���s
'
' ������  : sSigCode  - �Д���
'
' �@�\����:
'2006/05/23�ǉ�
'///////////////////////////////////////////////////
Public Function GetSigChkCode(Optional ByRef sSigCode As String) As Boolean
    Dim dbIsMine    As Boolean
    Dim sSQL        As String
    Dim objRs       As Object
    
    GetSigChkCode = False
    sSigCode = ""
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc022_SQL.bas -- Function GetSigChkCode"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�r�p�k���쐬
    sSQL = ""
    sSQL = sSQL & "SELECT NVL(kcode01a9, ' ')"   '0:�Д���
    sSQL = sSQL & "  FROM koda9"
    sSQL = sSQL & " WHERE sysca9 = 'X'"
    sSQL = sSQL & "   AND shuca9 = '19'"
    sSQL = sSQL & "   AND codea9 = 'GFA'"
    
    Set objRs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)
    
    If objRs.EOF Then
        Call MsgOut(0, "�Д����̃R�[�h���o�^����Ă��܂���", ERR_DISP)
        Exit Function
    End If

    sSigCode = objRs(0)     ''�Д���
    
    objRs.Close
    
    ''�Д���
    If IsNumeric(sSigCode) = False Then
        Call MsgOut(0, "�Д����̃R�[�h������������܂���", ERR_DISP)
        Exit Function
    End If
    
    If dbIsMine Then
        OraDBClose
    End If

    GetSigChkCode = True        ''����������Ԃ�

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
    
End Function


