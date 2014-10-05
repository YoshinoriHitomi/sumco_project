Attribute VB_Name = "s_cmbc028_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMJ001 (EPD����)
' �Q�Ɓ@�@: 060211_��������
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmjc001i_Disp
  '  CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
  '  TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    factory As String * 1           ' �H��
    opecond As String * 1           ' ���Ə���
    GOUKI As String * 3             ' ���@
    MEASURE As Integer              ' ����l
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

'�T�v      :�e�[�u���uTBCMJ001�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_cmjc001i_Disp ,���o���R�[�h
'          :SPLNUMs()     ,I  ,Integer      ,���o�����z��(�T���v��No)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/20�쐬�@����
Public Function DBDRV_Getcmjc001i_Disp(records() As typ_cmjc001i_Disp, SMPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere As String  'SQL��WHERE����
Dim sqlGroup As String  'SQL��GROUP����
Dim sqlOrder As String  'SQL��Order����
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    DBDRV_Getcmjc001i_Disp = FUNCTION_RETURN_FAILURE
    
    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001i_SQL.bas -- Function DBDRV_Getcmjc001i_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
              " OPECOND, GOUKI, MEASURE "
    sqlBase = sqlBase & "From TBCMJ001"
   ''���o����(�����NO)�̎��o��
    sqlWhere = "Where SMPLNO in ("
    For i = 1 To UBound(SMPLNUMs)
        sqlWhere = sqlWhere & "'" & SMPLNUMs(i) & "'"
        If i < UBound(SMPLNUMs) Then
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
        DBDRV_Getcmjc001i_Disp = FUNCTION_RETURN_FAILURE
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
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .GOUKI = rs("GOUKI")             ' ���@
            .MEASURE = rs("MEASURE")         ' ����l
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001i_Disp = FUNCTION_RETURN_SUCCESS

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

'�T�v      :�����œn���ꂽ���R�[�h��TBCMJ001�ɒǉ�����
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :record        ,I  ,typ_cmjc001i_Disp ,���o���R�[�h
'          :CRYNUM        ,I  ,String       ,�����ԍ�
'          :TSTAFFID      ,I  ,String       ,�o�^�Ј�ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�����񐔂̓e�[�u����̍ő�l+1�Ƃ���B
'����      :2001/06/25(Mon)�쐬�@����

Public Function DBDRV_Getcmjc001i_Exec(record As typ_cmjc001i_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
Dim sql         As String   'SQL
Dim sqlBase     As String   'SQL�x�[�X����
Dim sqlWhere    As String  'SQLWhere����
Dim sqlGroup    As String  'SQLGroup����


'    CRYNUM             �����ԍ��@�ˈ���
'    TRANCNT         �@ �����񐔁@�ˍő�
'   TSTAFFID            �o�^�Ј�ID�@�ˈ���
 '   REGDATE �@�@�@     �o�^���t�@��SYSDATE
 '   KSTAFFID           �X�V�Ј�ID�@��" "
 '   UPDDATE            �X�V���t�@��SYSDATE
 '   SENDFLAG           ���M�t���O�@��"0"
 '   SENDDATE           ���M���t�@��SYSDATE

    DBDRV_Getcmjc001i_Exec = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001i_SQL.bas -- Function DBDRV_Getcmjc001i_Exec"

    sqlBase = "Insert into TBCMJ001 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, " & _
              "HINBAN, REVNUM, FACTORY, OPECOND, GOUKI, MEASURE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) "
    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.hinban & "', " & record.REVNUM & ", '" & _
               record.factory & "', '" & record.opecond & "', '" & record.GOUKI & "', " & record.MEASURE & ", '" & _
               TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ001 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
            
    ''SQL�̎��s
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001i_Exec = FUNCTION_RETURN_SUCCESS


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

'�T�v      :�e�[�u���uTBCMJ001�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ001 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
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
'����      :2001/08/24�쐬�@�쑺
'          :06/04/11 ooba�@�֐����ύX <DBDRV_GetTBCME036> �� <DBDRV_GetTBCME036_cmbc028>
Public Function DBDRV_GetTBCME036_cmbc028(records() As typ_TBCME036, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
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
        DBDRV_GetTBCME036_cmbc028 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
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

    DBDRV_GetTBCME036_cmbc028 = FUNCTION_RETURN_SUCCESS
End Function

