Attribute VB_Name = "s_cmbc023_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMJ002 (������R����)
' �Q�Ɓ@�@: 060211_��������
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmjc001d_Disp
  '  CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
  '  TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��      Integer��Long  �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    GOUKI As String * 3             ' ���@
    TYPE As String * 1              ' �^�C�v
    MEAS1 As Double                 ' ����l�P
    MEAS2 As Double                 ' ����l�Q
    MEAS3 As Double                 ' ����l�R
    MEAS4 As Double                 ' ����l�S
    MEAS5 As Double                 ' ����l�T
    EFEHS As Double                 ' �����ΐ�
    RRG As Double                   ' �q�q�f
    JudgData As Double              ' �����Ώےl
    KANSANCHI As String             '10�����Z�l�@��
  '  TSTAFFID As String * 8          ' �o�^�Ј�ID
  '  REGDATE As Date                 ' �o�^���t
  '  KSTAFFID As String * 8          ' �X�V�Ј�ID
  '  UPDDATE As Date                 ' �X�V���t
  '  SENDFLAG As String * 1          ' ���M�t���O
  '  SENDDATE As Date                ' ���M���t



End Type

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ002�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_cmjc001d_Disp ,���o���R�[�h
'          :SPLNUMs()     ,I  ,Integer      ,���o�����z��(�T���v��No)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/20�쐬�@����
Public Function DBDRV_Getcmjc001d_Disp(records() As typ_cmjc001d_Disp, SPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere As String  'SQL��(WHERE��)
Dim sqlGroup As String  'SQL��(GROUP��)
Dim sqlOrder As String  'SQL��(ORDER��)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    DBDRV_Getcmjc001d_Disp = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmjc001d_SQL.bas -- Function DBDRV_Getcmjc001d_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
              " OPECOND, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, JUDGDATA "
    sqlBase = sqlBase & "From TBCMJ002"
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
        DBDRV_Getcmjc001d_Disp = FUNCTION_RETURN_FAILURE
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
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .FACTORY = rs("FACTORY")         ' �H��
            .OPECOND = rs("OPECOND")         ' ���Ə���
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
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001d_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

PROC_ERR:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�����œn���ꂽ���R�[�h��TBCMJ002�ɒǉ�����
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :record        ,I  ,typ_cmjc001d_Disp ,���o���R�[�h
'          :CRYNUM        ,I  ,String       ,�����ԍ�
'          :TSTAFFID      ,I  ,String       ,�o�^�Ј�ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�����񐔂̓e�[�u����̍ő�l+1�Ƃ���B
'����      :2001/06/22(Fri)�쐬�@����

Public Function DBDRV_Getcmjc001d_Exec(record As typ_cmjc001d_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL�x�[�X����
Dim sqlWhere As String  'SQLWhere����
Dim sqlGroup As String  'SQLGroup����

'    CRYNUM             �����ԍ� �@�ˈ���
'    TRANCNT         �@ �����񐔁@ �ˍő�
'    TSTAFFID           �o�^�Ј�ID �ˈ���
'    REGDATE �@�@�@     �o�^���t�@ ��SYSDATE
'    KSTAFFID           �X�V�Ј�ID ��" "
'    UPDDATE            �X�V���t�@ ��SYSDATE
'    SENDFLAG           ���M�t���O ��"0"
'    SENDDATE           ���M���t�@ ��SYSDATE

    DBDRV_Getcmjc001d_Exec = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmjc001d_SQL.bas -- Function DBDRV_Getcmjc001d_Exec"

    sqlBase = "Insert into TBCMJ002 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY, " & _
              "OPECOND, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.hinban & "', " & record.REVNUM & ", '" & _
               record.FACTORY & "', '" & record.OPECOND & "', '" & record.GOUKI & "', '" & record.TYPE & "', " & _
               record.MEAS1 & ", " & record.MEAS2 & ", " & record.MEAS3 & ", " & record.MEAS4 & ", " & record.MEAS5 & ", " & record.EFEHS & ", " & _
               record.RRG & ", " & record.JudgData & ", '" & TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ002 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
            
    ''SQL�̎��s
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001d_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

PROC_ERR:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�e�[�u���uXSDCS�v�̏����ɂ��������R�[�h���X�V����(�������ق̎����׸�)
'���Ұ�    :�ϐ���        ,IO ,�^                       ,����
'          :tblCrySmpMan  ,I   ,typ_XSDCS               ,�V�T���v���Ǘ��i�u���b�N�j�e�[�u���X�V�p�����[�^
'          :strCryNum     ,I   ,String                  ,�����ԍ�
'          :iSmpNo        ,I   ,Long                    ,�T���v��No.    Integer��Long 6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O   ,Integer                 ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :

Public Function UpdateTbl_CrySmpSuitei(tblCrySmpMan As typ_XSDCS, strCryNum As String, iSmpNo As Long) As Integer
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    
    UpdateTbl_CrySmpSuitei = FUNCTION_RETURN_FAILURE

    ' ��������ID1�̍X�V
    With tblCrySmpMan
        sql = "SELECT CRYNUMCS FROM XSDCS "
        sql = sql & "WHERE XTALCS = '" & strCryNum & "' and "
        sql = sql & "      CRYSMPLIDRS1CS = " & iSmpNo
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount <> 0 Then
            
            sql = "update XSDCS set "
            sql = sql & "CRYRESRS1CS='" & .CRYRESRS1CS & "', "          ' �����������сiRs-1)
            sql = sql & "KDAYCS=sysdate, "                              ' �X�V���t
            sql = sql & "SNDKCS='0' "                                   ' ���M�t���O
            sql = sql & "WHERE XTALCS = '" & strCryNum & "' and "
            sql = sql & "      CRYSMPLIDRS1CS = " & iSmpNo
Debug.Print sql
            If OraDB.ExecuteSQL(sql) <= 0 Then
                rs.Close
                Exit Function
            End If
        End If
        rs.Close
    End With

    ' ��������ID2�̍X�V (���F�X�V���e�́uCRYRESRS1CS�v�ɂ̂ݐݒ肳��Ă���B)
    With tblCrySmpMan
        sql = "SELECT CRYNUMCS FROM XSDCS "
        sql = sql & "WHERE XTALCS = '" & strCryNum & "' and "
        sql = sql & "      CRYSMPLIDRS2CS = " & iSmpNo
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount <> 0 Then
        
            sql = "update XSDCS set "
            sql = sql & "CRYRESRS2CS='" & .CRYRESRS1CS & "', "          ' �����������сiRs-2)
            sql = sql & "KDAYCS=sysdate, "                              ' �X�V���t
            sql = sql & "SNDKCS='0' "                                   ' ���M�t���O
            sql = sql & "WHERE XTALCS = '" & strCryNum & "' and "
            sql = sql & "      CRYSMPLIDRS2CS = " & iSmpNo
Debug.Print sql
            If OraDB.ExecuteSQL(sql) <= 0 Then
                rs.Close
                Exit Function
            End If
        End If
        rs.Close
    End With


    UpdateTbl_CrySmpSuitei = FUNCTION_RETURN_SUCCESS

End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------
'�T�v      :TBCMJ002�ɓo�^�f�[�^(TRANCNT=0)�ƂȂ�f�[�^�����݂��邩�m�F����
'���Ұ�    :�ϐ���        ,I/O ,�^             ,����
'          :getCryNum     ,I   ,String         ,�����ԍ�
'
'          :�߂�l        ,O   ,Integer        ,���������FFUNCTION_RETURN_SUCCESS
'                                              ,�������s�FFUNCTION_RETURN_FAILURE
'����      :TBCMJ002�ɓo�^�f�[�^(TRANCNT=0)�ƂȂ�f�[�^�����݂��邩�m�F����
'����      :2011/08/09 Akizuki


Public Function CheckTRANCNT0_UMU(record As typ_cmjc001d_Disp, CRYNUM$, TSTAFFID$) As Boolean

    Dim sql         As String   '���sSQL
    Dim sqlBase     As String   'SQL�x�[�X����
    Dim sqlWhere    As String   'SQLWhere����
    
    Dim rs As OraDynaset    'RecordSet
    Dim cnt As Integer      '�擾���� �ۑ��p


    CheckTRANCNT0_UMU = False
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmec066_SQL.bas -- Function getSIRDInfo"

    sqlBase = "select CRYNUM from TBCMJ002" & vbCrLf
   
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and "
    sqlWhere = sqlWhere & "(POSITION=" & record.POSITION & ") and "
    sqlWhere = sqlWhere & "(SMPKBN='" & record.SMPKBN & "') and "
    sqlWhere = sqlWhere & "(TRANCOND='" & record.TRANCOND & "') and "
    sqlWhere = sqlWhere & "(TRANCNT = '0') "
    
    sql = sqlBase & sqlWhere

    '��R����(TRANCNT=0)���擾����B
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
 
    '�擾�Ɏ��s�����ꍇ
    If rs Is Nothing Then
        CheckTRANCNT0_UMU = False
        Exit Function
    End If

    'SXLID���̌������擾����B
    cnt = rs.RecordCount
    
    If cnt >= 1 Then
        CheckTRANCNT0_UMU = True
    Else
        CheckTRANCNT0_UMU = False
    End If
    
    Exit Function
    
PROC_ERR:
    '�G���[�n���h��
    gErr.HandleError

End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------
'�T�v      :��R���уe�[�u��TBCMJ002(TRANCNT=0)���쐬����

'���Ұ�    :�ϐ���        ,IO   ,�^                   ,����
'          :record        ,I    ,typ_CMJC022i_Disp    ,���o���R�[�h
'
'          :�߂�l        ,O   ,Integer                 ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :TARNCNT=MAX �Ɠ����f�[�^��TRANCNT=0�ō쐬����
'����      :2011/08/09 SUMCO Akizuki TRANCNT=0�Ή�


 Public Function DBDRV_InsTBCMJ002_TRANCNT0_Exec(record As typ_cmjc001d_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
    Dim sql         As String   '���sSQL
    Dim sqlBase     As String   'SQL�x�[�X����
    Dim sqlWhere    As String   'SQLWhere����
        
    
    DBDRV_InsTBCMJ002_TRANCNT0_Exec = FUNCTION_RETURN_FAILURE


    ''SQL��g�ݗ��Ă�
    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmbc023_SQL.bas - Function DBDRV_InsTBCMJ002_TRANCNT0_Exec"
    
    ''SQL��g�ݗ��Ă�
    sqlBase = "Insert into TBCMJ002 ("
    sqlBase = sqlBase & "CRYNUM,"           '�����ԍ�
    sqlBase = sqlBase & "POSITION,"         '�ʒu
    sqlBase = sqlBase & "SMPKBN,"           '�T���v���敪
    sqlBase = sqlBase & "TRANCOND,"         '��������
    sqlBase = sqlBase & "TRANCNT,"          '������
    sqlBase = sqlBase & "SMPLNO,"           '�T���v��No
    sqlBase = sqlBase & "SMPLUMU,"          '�T���v���L��
    sqlBase = sqlBase & "KRPROCCD,"         '�Ǘ��H���R�[�h
    sqlBase = sqlBase & "PROCCODE,"         '�H���R�[�h
    sqlBase = sqlBase & "HINBAN,"           '�i��
    sqlBase = sqlBase & "REVNUM,"           '���i�ԍ������ԍ�
    sqlBase = sqlBase & "FACTORY,"          '�H��
    sqlBase = sqlBase & "OPECOND," & vbLf   '���Ə���
    sqlBase = sqlBase & "GOUKI,"            '���@
    sqlBase = sqlBase & "TYPE,"             '�^�C�v
    sqlBase = sqlBase & "MEAS1,"            '����l1
    sqlBase = sqlBase & "MEAS2,"            '����l2
    sqlBase = sqlBase & "MEAS3,"            '����l3
    sqlBase = sqlBase & "MEAS4,"            '����l4
    sqlBase = sqlBase & "MEAS5,"            '����l5
    sqlBase = sqlBase & "EFEHS,"            '�����ΐ�
    sqlBase = sqlBase & "RRG,"              'RRG
    sqlBase = sqlBase & "JUDGDATA,"         '�����Ώےl
    sqlBase = sqlBase & "TSTAFFID,"         '�o�^�Ј�ID
    sqlBase = sqlBase & "REGDATE,"          '�o�^���t
    sqlBase = sqlBase & "KSTAFFID,"         '�X�V�Ј�ID
    sqlBase = sqlBase & "UPDDATE,"          '�X�V���t
    sqlBase = sqlBase & "SENDFLAG,"         '���M�t���O
    sqlBase = sqlBase & "SENDDATE)" & vbLf  '���M���t
    
    
    'Select SQL�őΏۃf�[�^���擾���ăZ�b�g
    sqlBase = sqlBase & "VALUES(" & vbLf
    sqlBase = sqlBase & "'" & CRYNUM & "',"                     '�����ԍ�
    sqlBase = sqlBase & "'" & record.POSITION & "',"            '�ʒu
    sqlBase = sqlBase & "'" & record.SMPKBN & "',"              '�T���v���敪
    sqlBase = sqlBase & "'" & record.TRANCOND & "',"            '��������
    sqlBase = sqlBase & "'0'," & vbLf                           '�����񐔁@��TRANCNT=0�ō쐬
    sqlBase = sqlBase & "'" & record.SMPLNO & "',"              '�T���v��No
    sqlBase = sqlBase & "'" & record.SMPLUMU & "',"             '�T���v���L��
    sqlBase = sqlBase & "'" & record.KRPROCCD & "',"            '�Ǘ��H���R�[�h
    sqlBase = sqlBase & "'" & record.PROCCODE & "',"            '�H���R�[�h
    sqlBase = sqlBase & "'" & record.hinban & "',"              '�i��
    sqlBase = sqlBase & "'" & record.REVNUM & "',"              '���i�ԍ������ԍ�
    sqlBase = sqlBase & "'" & record.FACTORY & "',"             '�H��
    sqlBase = sqlBase & "'" & record.OPECOND & "',"             '���Ə���
    sqlBase = sqlBase & "'" & record.GOUKI & "'," & vbLf        '���@
    sqlBase = sqlBase & "'" & record.TYPE & "',"                '�^�C�v
    sqlBase = sqlBase & "'" & record.MEAS1 & "',"               '����l1
    sqlBase = sqlBase & "'" & record.MEAS2 & "',"               '����l2
    sqlBase = sqlBase & "'" & record.MEAS3 & "',"               '����l3
    sqlBase = sqlBase & "'" & record.MEAS4 & "',"               '����l4
    sqlBase = sqlBase & "'" & record.MEAS5 & "',"               '����l5
    sqlBase = sqlBase & "'" & record.EFEHS & "',"               '�����ΐ�
    sqlBase = sqlBase & "'" & record.RRG & "',"                 'RRG
    sqlBase = sqlBase & "'" & record.JudgData & "'," & vbLf     '�����Ώےl
    sqlBase = sqlBase & "'" & TSTAFFID & "',"                   '�o�^�Ј�ID
    sqlBase = sqlBase & "SYSDATE,"                              '�o�^���t
    sqlBase = sqlBase & "'" & TSTAFFID & "',"                   '�X�V�Ј�ID
    sqlBase = sqlBase & "SYSDATE,"                              '�X�V���t
    sqlBase = sqlBase & "'0',"                                  '���M�t���O
    sqlBase = sqlBase & "SYSDATE)" & vbLf                       '���M���t
    
    sql = sqlBase & sqlWhere
    
    ''SQL�̎��s
    OraDB.ExecuteSQL (sql)

    DBDRV_InsTBCMJ002_TRANCNT0_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

PROC_ERR:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
    

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------
'�T�v      :��R���уe�[�u��TBCMJ002(TRANCNT=0)���X�V����

'���Ұ�    :�ϐ���        ,IO   ,�^                   ,����
'          :record        ,I    ,typ_CMJC022i_Disp    ,���o���R�[�h
'
'          :�߂�l        ,O    ,Integer              ,���������FFUNCTION_RETURN_SUCCESS
'                                                     ,�������s�FFUNCTION_RETURN_FAILURE
'����      :TARNCNT=MAX �Ɠ����f�[�^��TRANCNT=0�ɍX�V����
'����      :2011/08/09 SUMCO Akizuki TRANCNT=0�Ή�


 Public Function DBDRV_UpdTBCMJ002_TRANCNT0_Exec(record As typ_cmjc001d_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
 
    Dim sql         As String   'SQL
    Dim sqlBase     As String   'SQL�x�[�X����
    Dim sqlWhere    As String   'SQLWhere����
    
    
    DBDRV_UpdTBCMJ002_TRANCNT0_Exec = FUNCTION_RETURN_FAILURE


    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmbc023_SQL.bas - Function DBDRV_UpdTBCMJ002_TRANCNT0_Exec"
  
    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmjc001d_SQL.bas -- Function DBDRV_Getcmjc001d_Exec"
    
    
    ''SQL��g�ݗ��Ă�
    sqlBase = "Update TBCMJ002 set "
    sqlBase = sqlBase & "CRYNUM = '" & CRYNUM & "',"                    '�����ԍ�
    sqlBase = sqlBase & "POSITION = '" & record.POSITION & "',"         '�ʒu
    sqlBase = sqlBase & "SMPKBN = '" & record.SMPKBN & "',"             '�T���v���敪
    sqlBase = sqlBase & "TRANCOND = '" & record.TRANCOND & "',"         '��������
    sqlBase = sqlBase & "TRANCNT = 0,"                                  '������ ��TRANCNT=0�ō쐬
    sqlBase = sqlBase & "SMPLNO = '" & record.SMPLNO & "',"             '�T���v��No
    sqlBase = sqlBase & "SMPLUMU = '" & record.SMPLUMU & "',"           '�T���v���L��
    sqlBase = sqlBase & "KRPROCCD = '" & record.KRPROCCD & "',"         '�Ǘ��H���R�[�h
    sqlBase = sqlBase & "PROCCODE = '" & record.PROCCODE & "',"         '�H���R�[�h
    sqlBase = sqlBase & "HINBAN = '" & record.hinban & "',"             '�i��
    sqlBase = sqlBase & "REVNUM = '" & record.REVNUM & "',"             '���i�ԍ������ԍ�
    sqlBase = sqlBase & "FACTORY = '" & record.FACTORY & "',"           '�H��
    sqlBase = sqlBase & "OPECOND = '" & record.OPECOND & "',"           '���Ə���
    sqlBase = sqlBase & "GOUKI = '" & record.GOUKI & "'," & vbLf        '���@
    sqlBase = sqlBase & "TYPE = '" & record.TYPE & "',"                 '�^�C�v
    sqlBase = sqlBase & "MEAS1 = '" & record.MEAS1 & "',"               '����l1
    sqlBase = sqlBase & "MEAS2 = '" & record.MEAS2 & "',"               '����l2
    sqlBase = sqlBase & "MEAS3 = '" & record.MEAS3 & "',"               '����l3
    sqlBase = sqlBase & "MEAS4 = '" & record.MEAS4 & "',"               '����l4
    sqlBase = sqlBase & "MEAS5 = '" & record.MEAS5 & "',"               '����l5
    sqlBase = sqlBase & "EFEHS = '" & record.EFEHS & "',"               '�����ΐ�
    sqlBase = sqlBase & "RRG = '" & record.RRG & "',"                   'RRG
    sqlBase = sqlBase & "JUDGDATA = '" & record.JudgData & "'," & vbLf  '�����Ώےl
    sqlBase = sqlBase & "KSTAFFID = '" & TSTAFFID & "',"                '�X�V�Ј�ID
    sqlBase = sqlBase & "UPDDATE = SYSDATE, "                           '�X�V���t
    sqlBase = sqlBase & "SENDFLAG = '0' " & vbLf                        '���M�t���O
    
    '�X�V�ΏۂƂȂ�f�[�^�̎w�� (TRANCNT=0�̓��f�[�^)
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and "
    sqlWhere = sqlWhere & "(POSITION=" & record.POSITION & ") and "
    sqlWhere = sqlWhere & "(SMPKBN='" & record.SMPKBN & "') and "
    sqlWhere = sqlWhere & "(TRANCOND=" & record.TRANCOND & ") and"
    sqlWhere = sqlWhere & "(TRANCNT = 0) "

    sql = sqlBase & sqlWhere
    
    ''SQL�̎��s
    OraDB.ExecuteSQL (sql)
    
    DBDRV_UpdTBCMJ002_TRANCNT0_Exec = FUNCTION_RETURN_SUCCESS


proc_exit:
    '�I��

    Exit Function

PROC_ERR:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
