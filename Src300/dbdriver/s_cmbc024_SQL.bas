Attribute VB_Name = "s_cmbc024_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMJ008 (�a�l�c����)
' �Q�Ɓ@�@: 060211_��������
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmjc001e_Disp
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
    MEASMETH As String * 1          ' ������@
    MEASSPOT As Integer             ' ����_
    MAG As String * 4               ' �{��
    HTPRC As String * 2             ' �M�������@
    KKSP As String * 3              ' �������ב���ʒu
    KKSET As String * 3             ' �������ב�������{�I��ET��@�@char(1)�{number(2)
    MEAS1 As Double                 ' ����l�P
    MEAS2 As Double                 ' ����l�Q
    MEAS3 As Double                 ' ����l�R
    MEAS4 As Double                 ' ����l�S
    MEAS5 As Double                 ' ����l�T
    MEASMIN As Double               ' MIN
    MEASMAX As Double               ' MAX
    MEASAVE As Double               ' AVE
    BMDMNBUNP As Double             ' BMD�ʓ����z
   ' TSTAFFID As String * 8          ' �o�^�Ј�ID
   ' KSTAFFID As String * 8          ' �X�V�Ј�ID
   ' UPDDATE As Date                 ' �X�V���t
   ' SENDFLAG As String * 1          ' ���M�t���O
   ' SENDDATE As Date                ' ���M���t

End Type



'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ008�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_cmjc001e_Disp ,���o���R�[�h
'          :SPLNUMs()     ,I  ,Integer      ,���o�����z��(�T���v��No)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/20�쐬�@����
'          :2003/04/02    �@yakimura  ���ڒǉ��Ή�
Public Function DBDRV_Getcmjc001e_Disp(records() As typ_cmjc001e_Disp, SMPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere As String  'SQL��WHERE����
Dim sqlGroup As String  'SQL��GROUP����
Dim sqlOrder As String  'SQL��Order����
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    DBDRV_Getcmjc001e_Disp = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001e_SQL.bas -- Function DBDRV_Getcmjc001e_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN, MEASMAX, MEASAVE, BMDMNBUNP "
    sqlBase = sqlBase & "From TBCMJ008"
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
        DBDRV_Getcmjc001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .POSITION = rs("POSITION")      ' �ʒu
            .SMPKBN = rs("SMPKBN")          ' �T���v���敪
            .TRANCOND = rs("TRANCOND")      ' ��������
            .SMPLNO = rs("SMPLNO")          ' �T���v���m��
            .SMPLUMU = rs("SMPLUMU")        ' �T���v���L��
            .hinban = rs("HINBAN")          ' �i��
            .REVNUM = rs("REVNUM")          ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")        ' �H��
            .opecond = rs("OPECOND")        ' ���Ə���
            .KRPROCCD = rs("KRPROCCD")      ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")      ' �H���R�[�h
            .GOUKI = rs("GOUKI")            ' ���@
            .MEASMETH = rs("MEASMETH")      ' ������@
            .MEASSPOT = rs("MEASSPOT")      ' ����_
            .MAG = rs("MAG")                ' �{��
            .HTPRC = rs("HTPRC")            ' �M�������@
            .KKSP = rs("KKSP")              ' �������ב���ʒu
            .KKSET = rs("KKSET")            ' �������ב�������{�I��ET��
            .MEAS1 = rs("MEAS1")            ' ����l�P
            .MEAS2 = rs("MEAS2")            ' ����l�Q
            .MEAS3 = rs("MEAS3")            ' ����l�R
            .MEAS4 = rs("MEAS4")            ' ����l�S
            .MEAS5 = rs("MEAS5")            ' ����l�T
            .MEASMIN = rs("MEASMIN")            ' MIN
            .MEASMAX = rs("MEASMAX")            ' MAX
            .MEASAVE = rs("MEASAVE")            ' AVE
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
            If rs("BMDMNBUNP") <> vbNullString Then
               .BMDMNBUNP = rs("BMDMNBUNP")     ' �a�l�c�ʓ����z
            Else
               .BMDMNBUNP = 0
            End If
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001e_Disp = FUNCTION_RETURN_SUCCESS

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

'�T�v      :�����œn���ꂽ���R�[�h��TBCMJ008�ɒǉ�����
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :record        ,I  ,typ_cmjc001e_Disp ,���o���R�[�h
'          :CRYNUM        ,I  ,String       ,�����ԍ�
'          :TSTAFFID      ,I  ,String       ,�o�^�Ј�ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�����񐔂̓e�[�u����̍ő�l+1�Ƃ���B
'����      :2001/06/22(Fri)�쐬�@����
'          :2003/04/02    �@yakimura  ���ڒǉ��Ή�

Public Function DBDRV_Getcmjc001e_Exec(record As typ_cmjc001e_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL�x�[�X����
Dim sqlWhere As String  'SQLWhere����
Dim sqlGroup As String  'SQLGroup����

'    CRYNUM             �����ԍ��@�ˈ���
'    TRANCNT         �@ �����񐔁@�ˍő�
'   TSTAFFID            �o�^�Ј�ID�@�ˈ���
 '   REGDATE �@�@�@     �o�^���t�@��SYSDATE
 '   KSTAFFID           �X�V�Ј�ID�@��" "
 '   UPDDATE            �X�V���t�@��SYSDATE
 '   SENDFLAG           ���M�t���O�@��"0"
 '   SENDDATE           ���M���t�@��SYSDATE

    DBDRV_Getcmjc001e_Exec = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001e_SQL.bas -- Function DBDRV_Getcmjc001e_Exec"

    sqlBase = "Insert into TBCMJ008 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, " & _
              "KRPROCCD, PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN, MEASMAX, MEASAVE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, BMDMNBUNP) " & vbCrLf
    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "',  '" & record.hinban & "', " & record.REVNUM & ",'" & record.factory & "', '" & record.opecond & "', '" & _
               record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', '" & record.MEASMETH & "', " & record.MEASSPOT & ", '" & record.MAG & "', '" & _
               record.HTPRC & "', '" & record.KKSP & "', '" & record.KKSET & "', " & record.MEAS1 & ", " & record.MEAS2 & ", " & record.MEAS3 & ", " & record.MEAS4 & ", " & _
               record.MEAS5 & ", " & record.MEASMIN & ", " & record.MEASMAX & ", " & record.MEASAVE & ", '" & TSTAFFID & "', SYSDATE,' ', SYSDATE, '0', SYSDATE, " & record.BMDMNBUNP & " from TBCMJ008 "
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') " & vbCrLf
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
            
Debug.Print sql
    ''SQL�̎��s
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001e_Exec = FUNCTION_RETURN_SUCCESS



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

'�T�v      :�e�[�u���uTBCMJ008�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ008 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
'          :2003/04/02    �@yakimura  ���ڒǉ��Ή�
Public Function DBDRV_GetTBCMJ008(records() As typ_TBCMJ008, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN," & _
              " MEASMAX, MEASAVE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, BMDMNBUNP "
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
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
            If rs("BMDMNBUNP") <> vbNullString Then
               .BMDMNBUNP = rs("BMDMNBUNP")  ' �a�l�c�ʓ����z
            Else
               .BMDMNBUNP = 0
            End If
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMJ008 = FUNCTION_RETURN_SUCCESS
End Function



