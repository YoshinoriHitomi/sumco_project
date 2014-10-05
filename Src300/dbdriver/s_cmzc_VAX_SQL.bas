Attribute VB_Name = "s_cmzcF_VAX_SQL"
Option Explicit


'                                     2001/10/03
'================================================
' DB�A�N�Z�X�֐�
' VAX_DB�A�N�Z�X�p
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------

Public Type typ_VAX_DR_CNDS
    PG_ID   As String * 6      ' �v���O����ID
    DR_CHRG As Long            ' �`���[�W��
    DR_CPOS As Integer         ' ���c�{�ʒu
    DR_CSIZ As Integer         ' ���c�{�T�C�Y
    DR_DIA  As Integer         ' ���a
    DR_LEN0 As Integer         ' ���㒷(1�{����/R0)
    DR_LEN1 As Integer         ' ���㒷(R1)
    DR_SR   As String * 9      ' �㎲��]��
    DR_CR   As String * 9      ' ������]��
    DR_GAP  As Integer         ' �M���b�v
    DR_PRES7 As String * 8       ' �F����            '2003/05/16 osawa
    'DR_PRES7 As Integer         ' �F����
    DR_AR7   As String * 7       ' �A���S������      '2003/05/16 osawa
    'DR_AR7   As Integer         ' �A���S������
    UPD_DATE  As Date          ' �X�V�������t
    EXT_DATE  As Date          ' ���o�������t
    DR_AR3   As Integer        ' �A���S�����R����
    DR_DOP   As Integer        ' �h�[�v
End Type

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :VAX���e�[�u���uDR_CNDS�v��������ɂ��������R�[�h�𒊏o����i�������R�[�h�j
'���Ұ�    :�ϐ���        ,IO ,�^              ,����
'          :records()     ,O  ,typ_VAX_DR_CNDS ,���o���R�[�h
'          :sqlWhere      ,I  ,String          ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String          ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/10/03 �쐬�@���{
Public Function DBDRV_VAX_DR_CNDS(records() As typ_VAX_DR_CNDS _
                                   , Optional sqlWhere$ = vbNullString _
                                   , Optional sqlOrder$ = vbNullString _
                                   ) As FUNCTION_RETURN
                                   
    Dim sql As String       'SQL�S��
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset


    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_VAX_SQL.bas -- Function DBDRV_VAX_DR_CNDS"
    
    DBDRV_VAX_DR_CNDS = FUNCTION_RETURN_FAILURE
    
    sql = "select "
    sql = sql & "PG_ID, "       ' �v���O����ID
    sql = sql & "DR_CHRG, "     ' �`���[�W��
    sql = sql & "DR_CPOS, "     ' ���c�{�ʒu
    sql = sql & "DR_CSIZ, "     ' ���c�{�T�C�Y
    sql = sql & "DR_DIA, "      ' ���a
    sql = sql & "DR_LEN0, "     ' ���㒷(1�{����/R0)
    sql = sql & "DR_LEN1, "     ' ���㒷(R1)
    sql = sql & "DR_SR, "       ' �㎲��]��
    sql = sql & "DR_CR, "       ' ������]��
    sql = sql & "DR_GAP, "      ' �M���b�v
    sql = sql & "DR_PRES7, "     ' �F����
    sql = sql & "DR_AR7,  "      ' �A���S������
    sql = sql & "DR_AR3,  "     ' �A���S�����R����
    sql = sql & "DR_DOP,  "     ' �h�[�v
    sql = sql & "UPD_DATE, "    ' �X�V�������t
    sql = sql & "EXT_DATE "     ' ���o�������t
    sql = sql & "from DR_CNDS "
    
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & sqlWhere & sqlOrder
    End If
Debug.Print sql

    Set db = DBEngine.Workspaces(0).OpenDatabase("VAX", dbDriverComplete, True, "ODBC;DATABASE=attach 'filename disk$xtal:[usr.rdb]xtal';UID=xtal;PWD=crystal;DSN=VAX")
    Set rs = db.OpenRecordset(sql)
        
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_VAX_DR_CNDS = FUNCTION_RETURN_FAILURE
        rs.Close
        db.Close
        Set rs = Nothing
        Set db = Nothing
        GoTo proc_exit
    End If
    
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .PG_ID = rs("PG_ID")         ' �v���O����ID
            .DR_CHRG = IIf(rs("DR_CHRG") <> "", rs("DR_CHRG"), 0) ' �`���[�W��
            .DR_CPOS = IIf(rs("DR_CPOS") <> "", rs("DR_CPOS"), 0) ' ���c�{�ʒu
            .DR_CSIZ = IIf(rs("DR_CSIZ") <> "", rs("DR_CSIZ"), 0)  ' ���c�{�T�C�Y
            .DR_DIA = IIf(rs("DR_DIA") <> "", rs("DR_DIA"), 0)     ' ���a
            .DR_LEN0 = IIf(rs("DR_LEN0") <> "", rs("DR_LEN0"), 0)  ' ���㒷(1�{����/R0)
            .DR_LEN1 = IIf(rs("DR_LEN1") <> "", rs("DR_LEN1"), 0)  ' ���㒷(R1)
            .DR_SR = IIf(rs("DR_SR") <> "", Trim(rs("DR_SR")), "0") ' �㎲��]��
            .DR_CR = IIf(rs("DR_CR") <> "", Trim(rs("DR_CR")), "0") ' ������]��
            .DR_GAP = IIf(rs("DR_GAP") <> "", rs("DR_GAP"), 0)             ' �M���b�v
            .DR_PRES7 = IIf(rs("DR_PRES7") <> "", Trim(rs("DR_PRES7")), "0")      ' �F����
            '.DR_PRES7 = IIf(rs("DR_PRES7") <> "", rs("DR_PRES7"), 0)      ' �F����
            .DR_AR7 = IIf(rs("DR_AR7") <> "", Trim(rs("DR_AR7")), "0")             ' �A���S������
            .DR_AR3 = IIf(rs("DR_AR3") <> "", rs("DR_AR3"), " ")       ' �A���S�����R���ʁ@4/30
            .DR_DOP = IIf(rs("DR_DOP") <> "", rs("DR_DOP"), " ")       ' �h�[�v
            .UPD_DATE = IIf(rs("UPD_DATE") <> "", rs("UPD_DATE"), Now) ' �X�V�������t
            .EXT_DATE = IIf(rs("EXT_DATE") <> "", rs("EXT_DATE"), Now) ' ���o�������t
            
            '�����ӂꂵ�Ȃ��悤�Ƀ`�F�b�N
            If .DR_CHRG < -9999 Or .DR_CHRG > 9999 Then
                .DR_CHRG = 9999
            End If
            If .DR_CPOS < -999 Or .DR_CPOS > 999 Then
                .DR_CPOS = 999
            End If
            If .DR_CSIZ < -99 Or .DR_CSIZ > 99 Then
                .DR_CSIZ = 99
            End If
            If .DR_DIA < -999 Or .DR_DIA > 999 Then
                .DR_DIA = 999
            End If
            If .DR_LEN0 < -9999 Or .DR_LEN0 > 9999 Then
                .DR_LEN0 = 9999
            End If
            If .DR_LEN1 < -9999 Or .DR_LEN1 > 9999 Then
                .DR_LEN1 = 9999
            End If
            If .DR_GAP < -999 Or .DR_GAP > 999 Then
                .DR_GAP = 999
            End If
        End With
        rs.MoveNext
    Next

'    Do While Not rs.EOF
'        Debug.Print rs.Fields("PG_ID")
'        Debug.Print rs.Fields(3)
'    Loop
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
    
    DBDRV_VAX_DR_CNDS = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_VAX_DR_CNDS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :VAX���e�[�u���uDR_CNDS�v��������ɂ��������R�[�h�𒊏o����(1���R�[�h)
'���Ұ�    :�ϐ���        ,IO ,�^              ,����
'          :record        ,O  ,typ_VAX_DR_CNDS ,���o���R�[�h
'          :sqlWhere      ,I  ,String          ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String          ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/10/03 �쐬�@���{
Public Function DBDRV_VAX_DR_CNDS1(record As typ_VAX_DR_CNDS _
                                   , Optional sqlWhere$ = vbNullString _
                                   , Optional sqlOrder$ = vbNullString _
                                   ) As FUNCTION_RETURN
                                   
    Dim sql As String       'SQL�S��
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_VAX_SQL.bas -- Function DBDRV_VAX_DR_CNDS1"
    
    DBDRV_VAX_DR_CNDS1 = FUNCTION_RETURN_FAILURE
    
    sql = "select "
    sql = sql & "PG_ID, "       ' �v���O����ID
    sql = sql & "DR_CHRG, "     ' �`���[�W��
    sql = sql & "DR_CPOS, "     ' ���c�{�ʒu
    sql = sql & "DR_CSIZ, "     ' ���c�{�T�C�Y
    sql = sql & "DR_DIA, "      ' ���a
    sql = sql & "DR_LEN0, "     ' ���㒷(1�{����/R0)
    sql = sql & "DR_LEN1, "     ' ���㒷(R1)
    sql = sql & "DR_SR, "       ' �㎲��]��
    sql = sql & "DR_CR, "       ' ������]��
    sql = sql & "DR_GAP, "      ' �M���b�v
    sql = sql & "DR_PRES7, "     ' �F����
    sql = sql & "DR_AR7,  "      ' �A���S������
    sql = sql & "DR_AR3,  "      ' �A���S�����R����   4/30
    sql = sql & "DR_DOP,  "      ' �h�[�v
    sql = sql & "UPD_DATE, "    ' �X�V�������t
    sql = sql & "EXT_DATE "     ' ���o�������t
    sql = sql & "from DR_CNDS "
    
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & sqlWhere & sqlOrder
    End If

    Set db = DBEngine.Workspaces(0).OpenDatabase("VAX", dbDriverComplete, True, "ODBC;DATABASE=attach 'filename disk$xtal:[usr.rdb]xtal';UID=xtal;PWD=crystal;DSN=VAX")
    Set rs = db.OpenRecordset(sql)
        
    If rs Is Nothing Then
        rs.Close
        db.Close
        Set rs = Nothing
        Set db = Nothing
        DBDRV_VAX_DR_CNDS1 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    '���R�[�h�O������PGID��Null�ɂ��ĕԂ��B
    If rs.RecordCount = 0 Then
        record.PG_ID = vbNullString
    Else
        With record
            .PG_ID = rs("PG_ID")         ' �v���O����ID
            .DR_CHRG = IIf(rs("DR_CHRG") <> "", rs("DR_CHRG"), 0) ' �`���[�W��
            .DR_CPOS = IIf(rs("DR_CPOS") <> "", rs("DR_CPOS"), 0) ' ���c�{�ʒu
            .DR_CSIZ = IIf(rs("DR_CSIZ") <> "", rs("DR_CSIZ"), 0)  ' ���c�{�T�C�Y
            .DR_DIA = IIf(rs("DR_DIA") <> "", rs("DR_DIA"), 0)     ' ���a
            .DR_LEN0 = IIf(rs("DR_LEN0") <> "", rs("DR_LEN0"), 0)  ' ���㒷(1�{����/R0)
            .DR_LEN1 = IIf(rs("DR_LEN1") <> "", rs("DR_LEN1"), 0)  ' ���㒷(R1)
            .DR_SR = IIf(rs("DR_SR") <> "", Trim(rs("DR_SR")), "") ' �㎲��]��
            .DR_CR = IIf(rs("DR_CR") <> "", Trim(rs("DR_CR")), "") ' ������]��
            .DR_GAP = IIf(rs("DR_GAP") <> "", rs("DR_GAP"), 0)             ' �M���b�v
            .DR_PRES7 = IIf(rs("DR_PRES7") <> "", rs("DR_PRES7"), 0)      ' �F����
            .DR_AR7 = IIf(rs("DR_AR7") <> "", rs("DR_AR7"), 0)            ' �A���S������
            .DR_AR3 = IIf(rs("DR_AR3") <> "", rs("DR_AR3"), " ")         ' �A���S�����R���ʁ@�@4/30
            .DR_DOP = IIf(rs("DR_DOP") <> "", rs("DR_DOP"), " ")         ' �h�[�v
            .UPD_DATE = IIf(rs("UPD_DATE") <> "", rs("UPD_DATE"), Now) ' �X�V�������t
            .EXT_DATE = IIf(rs("EXT_DATE") <> "", rs("EXT_DATE"), Now) ' ���o�������t
            
            '�����ӂꂵ�Ȃ��悤�Ƀ`�F�b�N
            If .DR_CHRG < -9999 Or .DR_CHRG > 9999 Then
                .DR_CHRG = 9999
            End If
            If .DR_CPOS < -999 Or .DR_CPOS > 999 Then
                .DR_CPOS = 999
            End If
            If .DR_CSIZ < -99 Or .DR_CSIZ > 99 Then
                .DR_CSIZ = 99
            End If
            If .DR_DIA < -999 Or .DR_DIA > 999 Then
                .DR_DIA = 999
            End If
            If .DR_LEN0 < -9999 Or .DR_LEN0 > 9999 Then
                .DR_LEN0 = 9999
            End If
            If .DR_LEN1 < -9999 Or .DR_LEN1 > 9999 Then
                .DR_LEN1 = 9999
            End If
            If .DR_GAP < -999 Or .DR_GAP > 999 Then
                .DR_GAP = 999
            End If

        End With
    End If

    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing

    DBDRV_VAX_DR_CNDS1 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_VAX_DR_CNDS1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


