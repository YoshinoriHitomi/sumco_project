Attribute VB_Name = "s_cmbc027_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMJ007 (���C�t�^�C��)
' �Q�Ɓ@�@: 060211_��������
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmjc001h_Disp
   ' CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
   ' TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    factory As String * 1           ' �H��
    opecond As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    MEAS1 As Integer                ' ����l�P
    MEAS2 As Integer                ' ����l�Q
    MEAS3 As Integer                ' ����l�R
    MEAS4 As Integer                ' ����l�S
    MEAS5 As Integer                ' ����l�T
    MEASPEAK As Integer             ' ����l �s�[�N�l
    CALCMEAS As Integer             ' �v�Z����
   ' TSTAFFID As String * 8          ' �o�^�Ј�ID
   ' REGDATE As Date                 ' �o�^���t
   ' KSTAFFID As String * 8          ' �X�V�Ј�ID
   ' UPDDATE As Date                 ' �X�V���t
   ' SENDFLAG As String * 1          ' ���M�t���O
   ' SENDDATE As Date                ' ���M���t
' Add Start 2005/11/14 M.Makino
    MEAS6 As Integer                ' ����l�U
    MEAS7 As Integer                ' ����l�V
    MEAS8 As Integer                ' ����l�W
    MEAS9 As Integer                ' ����l�X
    MEAS10 As Integer               ' ����l�P�O
    MEASFILE As String              ' ����f�[�^�t�@�C����
    RESVAL As String                ' ������R
    INCVAL As String                ' �X��
    CUTVAL As String                ' �ؕ�
    SETVAL As String                ' �ݒ�l
    CONVAL As String                ' 10�����Z�l
    MEAS1DAT1 As String            ' ����l�P�@���f�[�^�P
    MEAS1DAT2 As String            ' ����l�P�@���f�[�^�Q
    MEAS1DAT3 As String            ' ����l�P�@���f�[�^�R
    MEAS1DAT4 As String            ' ����l�P�@���f�[�^�S
    MEAS1DAT5 As String            ' ����l�P�@���f�[�^�T
    MEAS2DAT1 As String            ' ����l�Q�@���f�[�^�P
    MEAS2DAT2 As String            ' ����l�Q�@���f�[�^�Q
    MEAS2DAT3 As String            ' ����l�Q�@���f�[�^�R
    MEAS2DAT4 As String            ' ����l�Q�@���f�[�^�S
    MEAS2DAT5 As String            ' ����l�Q�@���f�[�^�T
    MEAS3DAT1 As String            ' ����l�R�@���f�[�^�P
    MEAS3DAT2 As String            ' ����l�R�@���f�[�^�Q
    MEAS3DAT3 As String            ' ����l�R�@���f�[�^�R
    MEAS3DAT4 As String            ' ����l�R�@���f�[�^�S
    MEAS3DAT5 As String            ' ����l�R�@���f�[�^�T
    MEAS4DAT1 As String            ' ����l�S�@���f�[�^�P
    MEAS4DAT2 As String            ' ����l�S�@���f�[�^�Q
    MEAS4DAT3 As String            ' ����l�S�@���f�[�^�R
    MEAS4DAT4 As String            ' ����l�S�@���f�[�^�S
    MEAS4DAT5 As String            ' ����l�S�@���f�[�^�T
    MEAS5DAT1 As String            ' ����l�T�@���f�[�^�P
    MEAS5DAT2 As String            ' ����l�T�@���f�[�^�Q
    MEAS5DAT3 As String            ' ����l�T�@���f�[�^�R
    MEAS5DAT4 As String            ' ����l�T�@���f�[�^�S
    MEAS5DAT5 As String            ' ����l�T�@���f�[�^�T
    MEAS6DAT1 As String            ' ����l�U�@���f�[�^�P
    MEAS6DAT2 As String            ' ����l�U�@���f�[�^�Q
    MEAS6DAT3 As String            ' ����l�U�@���f�[�^�R
    MEAS6DAT4 As String            ' ����l�U�@���f�[�^�S
    MEAS6DAT5 As String            ' ����l�U�@���f�[�^�T
    MEAS7DAT1 As String            ' ����l�V�@���f�[�^�P
    MEAS7DAT2 As String            ' ����l�V�@���f�[�^�Q
    MEAS7DAT3 As String            ' ����l�V�@���f�[�^�R
    MEAS7DAT4 As String            ' ����l�V�@���f�[�^�S
    MEAS7DAT5 As String            ' ����l�V�@���f�[�^�T
    MEAS8DAT1 As String            ' ����l�W�@���f�[�^�P
    MEAS8DAT2 As String            ' ����l�W�@���f�[�^�Q
    MEAS8DAT3 As String            ' ����l�W�@���f�[�^�R
    MEAS8DAT4 As String            ' ����l�W�@���f�[�^�S
    MEAS8DAT5 As String            ' ����l�W�@���f�[�^�T
    MEAS9DAT1 As String            ' ����l�X�@���f�[�^�P
    MEAS9DAT2 As String            ' ����l�X�@���f�[�^�Q
    MEAS9DAT3 As String            ' ����l�X�@���f�[�^�R
    MEAS9DAT4 As String            ' ����l�X�@���f�[�^�S
    MEAS9DAT5 As String            ' ����l�X�@���f�[�^�T
    MEAS10DAT1 As String           ' ����l�P�O�@���f�[�^�P
    MEAS10DAT2 As String           ' ����l�P�O�@���f�[�^�Q
    MEAS10DAT3 As String           ' ����l�P�O�@���f�[�^�R
    MEAS10DAT4 As String           ' ����l�P�O�@���f�[�^�S
    MEAS10DAT5 As String           ' ����l�P�O�@���f�[�^�T
    LTSPIFLG As String             ' ����ʒu����t���O
' Add End   2005/11/14 M.Makino
End Type

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ007�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_cmjc001h_Disp ,���o���R�[�h
'          :SPLNUMs()     ,I  ,Integer      ,���o�����z��(�T���v��No)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/20�쐬�@����
Public Function DBDRV_Getcmjc001h_Disp(records() As typ_cmjc001h_Disp, SMPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere As String  'SQL��WHERE����
Dim sqlGroup As String  'SQL��GROUP����
Dim sqlOrder As String  'SQL��Order����
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    DBDRV_Getcmjc001h_Disp = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001h_SQL.bas -- Function DBDRV_Getcmjc001h_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT) , SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS "
    sqlBase = sqlBase & "From TBCMJ007"
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
        DBDRV_Getcmjc001h_Disp = FUNCTION_RETURN_FAILURE
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
            .MEAS1 = rs("MEAS1")             ' ����l�P
            .MEAS2 = rs("MEAS2")             ' ����l�Q
            .MEAS3 = rs("MEAS3")             ' ����l�R
            .MEAS4 = rs("MEAS4")             ' ����l�S
            .MEAS5 = rs("MEAS5")             ' ����l�T
            .MEASPEAK = rs("MEASPEAK")       ' ����l �s�[�N�l
            .CALCMEAS = rs("CALCMEAS")       ' �v�Z����
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001h_Disp = FUNCTION_RETURN_SUCCESS

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

'�T�v      :�����œn���ꂽ���R�[�h��TBCMJ007�ɒǉ�����
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :record        ,I  ,typ_cmjc001h_Disp ,���o���R�[�h
'          :CRYNUM        ,I  ,String       ,�����ԍ�
'          :TSTAFFID      ,I  ,String       ,�o�^�Ј�ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�����񐔂̓e�[�u����̍ő�l+1�Ƃ���B
'����      :2001/06/25(mon)�쐬�@����

Public Function DBDRV_Getcmjc001h_Exec(record As typ_cmjc001h_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN
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

    DBDRV_Getcmjc001h_Exec = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001h_SQL.bas -- Function DBDRV_Getcmjc001h_Exec"

' Mod Start 2005/11/14 M.Makino
'    sqlBase = "Insert into TBCMJ007 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, " & _
'              "KRPROCCD, PROCCODE, GOUKI, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE) " & vbCrLf
'    sqlBase = sqlBase & "select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
'               record.SMPLNO & ", '" & record.SMPLUMU & "',  '" & record.HINBAN & "', " & record.REVNUM & ",'" & record.FACTORY & "', '" & record.OPECOND & "', '" & _
'               record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', " & record.MEAS1 & ", " & record.MEAS2 & ", " & record.MEAS3 & ", " & _
'               record.MEAS4 & ", " & record.MEAS5 & ", " & record.MEASPEAK & ", " & record.CALCMEAS & ", '" & TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ007 "
'    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') " & vbCrLf

    sqlBase = "Insert into TBCMJ007"
    sqlBase = sqlBase & " (CRYNUM"          ' [�����ԍ�]
    sqlBase = sqlBase & ", POSITION"        ' [�ʒu]
    sqlBase = sqlBase & ", SMPKBN"          ' [�T���v���敪]
    sqlBase = sqlBase & ", TRANCOND"        ' [��������]
    sqlBase = sqlBase & ", TRANCNT"         ' [������]
    sqlBase = sqlBase & ", SMPLNO"          ' [�T���v��No]
    sqlBase = sqlBase & ", SMPLUMU"         ' [�T���v���L��]
    sqlBase = sqlBase & ", HINBAN"          ' [�i��]
    sqlBase = sqlBase & ", REVNUM"          ' [���i�ԍ������ԍ�]
    sqlBase = sqlBase & ", FACTORY"         ' [�H��]
    sqlBase = sqlBase & ", OPECOND"         ' [���Ə���]
    sqlBase = sqlBase & ", KRPROCCD"        ' [�Ǘ��H���R�[�h]
    sqlBase = sqlBase & ", PROCCODE"        ' [�H���R�[�h]
    sqlBase = sqlBase & ", GOUKI"           ' [���@]
    sqlBase = sqlBase & ", MEAS1"           ' [����l1]
    sqlBase = sqlBase & ", MEAS2"           ' [����l2]
    sqlBase = sqlBase & ", MEAS3"           ' [����l3]
    sqlBase = sqlBase & ", MEAS4"           ' [����l4]
    sqlBase = sqlBase & ", MEAS5"           ' [����l5]
    sqlBase = sqlBase & ", MEASPEAK"        ' [����l �s�[�N�l]
    sqlBase = sqlBase & ", CALCMEAS"        ' [�v�Z����]
    sqlBase = sqlBase & ", TSTAFFID"        ' [�o�^�Ј�ID]
    sqlBase = sqlBase & ", REGDATE"         ' [�o�^���t]
    sqlBase = sqlBase & ", KSTAFFID"        ' [�X�V�Ј�ID]
    sqlBase = sqlBase & ", UPDDATE"         ' [�X�V���t]
    sqlBase = sqlBase & ", SENDFLAG"        ' [���M�t���O]
    sqlBase = sqlBase & ", SENDDATE"        ' [���M���t]
    sqlBase = sqlBase & ", MEAS6"           ' [����l�U]
    sqlBase = sqlBase & ", MEAS7"           ' [����l�V]
    sqlBase = sqlBase & ", MEAS8"           ' [����l�W]
    sqlBase = sqlBase & ", MEAS9"           ' [����l�X]
    sqlBase = sqlBase & ", MEAS10"          ' [����l�P�O]
    sqlBase = sqlBase & ", MEASFILE"        ' [����f�[�^�t�@�C����]
    sqlBase = sqlBase & ", RESVAL"          ' [������R]
    sqlBase = sqlBase & ", INCVAL"          ' [�X��]
    sqlBase = sqlBase & ", CUTVAL"          ' [�ؕ�]
    sqlBase = sqlBase & ", SETVAL"          ' [�ݒ�l]
    sqlBase = sqlBase & ", CONVAL"          ' [�P�O�����Z�l]
    sqlBase = sqlBase & ", MEAS1DAT1"       ' [����l�P�@���f�[�^�P]
    sqlBase = sqlBase & ", MEAS1DAT2"       ' [����l�P�@���f�[�^�Q]
    sqlBase = sqlBase & ", MEAS1DAT3"       ' [����l�P�@���f�[�^�R]
    sqlBase = sqlBase & ", MEAS1DAT4"       ' [����l�P�@���f�[�^�S]
    sqlBase = sqlBase & ", MEAS1DAT5"       ' [����l�P�@���f�[�^�T]
    sqlBase = sqlBase & ", MEAS2DAT1"       ' [����l�Q�@���f�[�^�P]
    sqlBase = sqlBase & ", MEAS2DAT2"       ' [����l�Q�@���f�[�^�Q]
    sqlBase = sqlBase & ", MEAS2DAT3"       ' [����l�Q�@���f�[�^�R]
    sqlBase = sqlBase & ", MEAS2DAT4"       ' [����l�Q�@���f�[�^�S]
    sqlBase = sqlBase & ", MEAS2DAT5"       ' [����l�Q�@���f�[�^�T]
    sqlBase = sqlBase & ", MEAS3DAT1"       ' [����l�R�@���f�[�^�P]
    sqlBase = sqlBase & ", MEAS3DAT2"       ' [����l�R�@���f�[�^�Q]
    sqlBase = sqlBase & ", MEAS3DAT3"       ' [����l�R�@���f�[�^�R]
    sqlBase = sqlBase & ", MEAS3DAT4"       ' [����l�R�@���f�[�^�S]
    sqlBase = sqlBase & ", MEAS3DAT5"       ' [����l�R�@���f�[�^�T]
    sqlBase = sqlBase & ", MEAS4DAT1"       ' [����l�S�@���f�[�^�P]
    sqlBase = sqlBase & ", MEAS4DAT2"       ' [����l�S�@���f�[�^�Q]
    sqlBase = sqlBase & ", MEAS4DAT3"       ' [����l�S�@���f�[�^�R]
    sqlBase = sqlBase & ", MEAS4DAT4"       ' [����l�S�@���f�[�^�S]
    sqlBase = sqlBase & ", MEAS4DAT5"       ' [����l�S�@���f�[�^�T]
    sqlBase = sqlBase & ", MEAS5DAT1"       ' [����l�T�@���f�[�^�P]
    sqlBase = sqlBase & ", MEAS5DAT2"       ' [����l�T�@���f�[�^�Q]
    sqlBase = sqlBase & ", MEAS5DAT3"       ' [����l�T�@���f�[�^�R]
    sqlBase = sqlBase & ", MEAS5DAT4"       ' [����l�T�@���f�[�^�S]
    sqlBase = sqlBase & ", MEAS5DAT5"       ' [����l�T�@���f�[�^�T]
    sqlBase = sqlBase & ", MEAS6DAT1"       ' [����l�U�@���f�[�^�P]
    sqlBase = sqlBase & ", MEAS6DAT2"       ' [����l�U�@���f�[�^�Q]
    sqlBase = sqlBase & ", MEAS6DAT3"       ' [����l�U�@���f�[�^�R]
    sqlBase = sqlBase & ", MEAS6DAT4"       ' [����l�U�@���f�[�^�S]
    sqlBase = sqlBase & ", MEAS6DAT5"       ' [����l�U�@���f�[�^�T]
    sqlBase = sqlBase & ", MEAS7DAT1"       ' [����l�V�@���f�[�^�P]
    sqlBase = sqlBase & ", MEAS7DAT2"       ' [����l�V�@���f�[�^�Q]
    sqlBase = sqlBase & ", MEAS7DAT3"       ' [����l�V�@���f�[�^�R]
    sqlBase = sqlBase & ", MEAS7DAT4"       ' [����l�V�@���f�[�^�S]
    sqlBase = sqlBase & ", MEAS7DAT5"       ' [����l�V�@���f�[�^�T]
    sqlBase = sqlBase & ", MEAS8DAT1"       ' [����l�W�@���f�[�^�P]
    sqlBase = sqlBase & ", MEAS8DAT2"       ' [����l�W�@���f�[�^�Q]
    sqlBase = sqlBase & ", MEAS8DAT3"       ' [����l�W�@���f�[�^�R]
    sqlBase = sqlBase & ", MEAS8DAT4"       ' [����l�W�@���f�[�^�S]
    sqlBase = sqlBase & ", MEAS8DAT5"       ' [����l�W�@���f�[�^�T]
    sqlBase = sqlBase & ", MEAS9DAT1"       ' [����l�X�@���f�[�^�P]
    sqlBase = sqlBase & ", MEAS9DAT2"       ' [����l�X�@���f�[�^�Q]
    sqlBase = sqlBase & ", MEAS9DAT3"       ' [����l�X�@���f�[�^�R]
    sqlBase = sqlBase & ", MEAS9DAT4"       ' [����l�X�@���f�[�^�S]
    sqlBase = sqlBase & ", MEAS9DAT5"       ' [����l�X�@���f�[�^�T]
    sqlBase = sqlBase & ", MEAS10DAT1"      ' [����l�P�O�@���f�[�^�P]
    sqlBase = sqlBase & ", MEAS10DAT2"      ' [����l�P�O�@���f�[�^�Q]
    sqlBase = sqlBase & ", MEAS10DAT3"      ' [����l�P�O�@���f�[�^�R]
    sqlBase = sqlBase & ", MEAS10DAT4"      ' [����l�P�O�@���f�[�^�S]
    sqlBase = sqlBase & ", MEAS10DAT5"      ' [����l�P�O�@���f�[�^�T]
    sqlBase = sqlBase & ", LTSPIFLG"        ' [����ʒu����t���O]
    sqlBase = sqlBase & ") select"
    sqlBase = sqlBase & "  '" & CRYNUM & "'"                    ' [�����ԍ�]
    sqlBase = sqlBase & ", " & record.POSITION                  ' [�ʒu]
    sqlBase = sqlBase & ", '" & record.SMPKBN & "'"             ' [�T���v���敪]
    sqlBase = sqlBase & ", '" & record.TRANCOND & "'"           ' [��������]
    sqlBase = sqlBase & ", nvl(MAX(TRANCNT),0) + 1"             ' [������]
    sqlBase = sqlBase & ", " & record.SMPLNO                    ' [�T���v��No]
    sqlBase = sqlBase & ", '" & record.SMPLUMU & "'"            ' [�T���v���L��]
    sqlBase = sqlBase & ", '" & record.hinban & "'"             ' [�i��]
    sqlBase = sqlBase & ", " & record.REVNUM                    ' [���i�ԍ������ԍ�]
    sqlBase = sqlBase & ", '" & record.factory & "'"            ' [�H��]
    sqlBase = sqlBase & ", '" & record.opecond & "'"            ' [���Ə���]
    sqlBase = sqlBase & ", '" & record.KRPROCCD & "'"           ' [�Ǘ��H���R�[�h]
    sqlBase = sqlBase & ", '" & record.PROCCODE & "'"           ' [�H���R�[�h]
    sqlBase = sqlBase & ", '" & record.GOUKI & "'"              ' [���@]
    sqlBase = sqlBase & ", " & record.MEAS1                     ' [����l1]
    sqlBase = sqlBase & ", " & record.MEAS2                     ' [����l2]
    sqlBase = sqlBase & ", " & record.MEAS3                     ' [����l3]
    sqlBase = sqlBase & ", " & record.MEAS4                     ' [����l4]
    sqlBase = sqlBase & ", " & record.MEAS5                     ' [����l5]
    sqlBase = sqlBase & ", " & record.MEASPEAK                  ' [����l �s�[�N�l]
    sqlBase = sqlBase & ", " & record.CALCMEAS                  ' [�v�Z����]
    sqlBase = sqlBase & ", '" & TSTAFFID & "'"                  ' [�o�^�Ј�ID]
    sqlBase = sqlBase & ", SYSDATE"                             ' [�o�^���t]
    sqlBase = sqlBase & ", ' '"                                 ' [�X�V�Ј�ID]
    sqlBase = sqlBase & ", SYSDATE"                             ' [�X�V���t]
    sqlBase = sqlBase & ", '0'"                                 ' [���M�t���O]
    sqlBase = sqlBase & ", SYSDATE"                             ' [���M���t]
    sqlBase = sqlBase & ", " & record.MEAS6                     ' [����l�U]
    sqlBase = sqlBase & ", " & record.MEAS7                     ' [����l�V]
    sqlBase = sqlBase & ", " & record.MEAS8                     ' [����l�W]
    sqlBase = sqlBase & ", " & record.MEAS9                     ' [����l�X]
    sqlBase = sqlBase & ", " & record.MEAS10                    ' [����l�P�O]
    sqlBase = sqlBase & ", '" & record.MEASFILE & "'"           ' [����f�[�^�t�@�C����]
    sqlBase = sqlBase & ", " & LZeroToNull(record.RESVAL)       ' [������R]
    sqlBase = sqlBase & ", " & LZeroToNull(record.INCVAL)       ' [�X��]
    sqlBase = sqlBase & ", " & LZeroToNull(record.CUTVAL)       ' [�ؕ�]
    sqlBase = sqlBase & ", " & LZeroToNull(record.SETVAL)       ' [�ݒ�l]
    sqlBase = sqlBase & ", " & LZeroToNull(record.CONVAL)       ' [10�����Z�l]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS1DAT1)    ' [����l�P�@���f�[�^�P]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS1DAT2)    ' [����l�P�@���f�[�^�Q]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS1DAT3)    ' [����l�P�@���f�[�^�R]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS1DAT4)    ' [����l�P�@���f�[�^�S]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS1DAT5)    ' [����l�P�@���f�[�^�T]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS2DAT1)    ' [����l�Q�@���f�[�^�P]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS2DAT2)    ' [����l�Q�@���f�[�^�Q]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS2DAT3)    ' [����l�Q�@���f�[�^�R]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS2DAT4)    ' [����l�Q�@���f�[�^�S]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS2DAT5)    ' [����l�Q�@���f�[�^�T]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS3DAT1)    ' [����l�R�@���f�[�^�P]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS3DAT2)    ' [����l�R�@���f�[�^�Q]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS3DAT3)    ' [����l�R�@���f�[�^�R]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS3DAT4)    ' [����l�R�@���f�[�^�S]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS3DAT5)    ' [����l�R�@���f�[�^�T]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS4DAT1)    ' [����l�S�@���f�[�^�P]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS4DAT2)    ' [����l�S�@���f�[�^�Q]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS4DAT3)    ' [����l�S�@���f�[�^�R]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS4DAT4)    ' [����l�S�@���f�[�^�S]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS4DAT5)    ' [����l�S�@���f�[�^�T]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS5DAT1)    ' [����l�T�@���f�[�^�P]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS5DAT2)    ' [����l�T�@���f�[�^�Q]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS5DAT3)    ' [����l�T�@���f�[�^�R]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS5DAT4)    ' [����l�T�@���f�[�^�S]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS5DAT5)    ' [����l�T�@���f�[�^�T]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS6DAT1)    ' [����l�U�@���f�[�^�P]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS6DAT2)    ' [����l�U�@���f�[�^�Q]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS6DAT3)    ' [����l�U�@���f�[�^�R]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS6DAT4)    ' [����l�U�@���f�[�^�S]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS6DAT5)    ' [����l�U�@���f�[�^�T]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS7DAT1)    ' [����l�V�@���f�[�^�P]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS7DAT2)    ' [����l�V�@���f�[�^�Q]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS7DAT3)    ' [����l�V�@���f�[�^�R]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS7DAT4)    ' [����l�V�@���f�[�^�S]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS7DAT5)    ' [����l�V�@���f�[�^�T]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS8DAT1)    ' [����l�W�@���f�[�^�P]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS8DAT2)    ' [����l�W�@���f�[�^�Q]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS8DAT3)    ' [����l�W�@���f�[�^�R]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS8DAT4)    ' [����l�W�@���f�[�^�S]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS8DAT5)    ' [����l�W�@���f�[�^�T]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS9DAT1)    ' [����l�X�@���f�[�^�P]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS9DAT2)    ' [����l�X�@���f�[�^�Q]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS9DAT3)    ' [����l�X�@���f�[�^�R]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS9DAT4)    ' [����l�X�@���f�[�^�S]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS9DAT5)    ' [����l�X�@���f�[�^�T]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS10DAT1)   ' [����l�P�O�@���f�[�^�P]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS10DAT2)   ' [����l�P�O�@���f�[�^�Q]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS10DAT3)   ' [����l�P�O�@���f�[�^�R]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS10DAT4)   ' [����l�P�O�@���f�[�^�S]
    sqlBase = sqlBase & ", " & LZeroToNull(record.MEAS10DAT5)   ' [����l�P�O�@���f�[�^�T]
    sqlBase = sqlBase & ", '" & record.LTSPIFLG & "'"           ' [����ʒu����t���O]
    sqlBase = sqlBase & " from TBCMJ007"

    sqlWhere = sqlWhere & " where"
    sqlWhere = sqlWhere & " (CRYNUM='" & CRYNUM & "')"
    sqlWhere = sqlWhere & " and (POSITION=" & record.POSITION & ")"
    sqlWhere = sqlWhere & " and (SMPKBN='" & record.SMPKBN & "')"
    sqlWhere = sqlWhere & " and (TRANCOND='" & record.TRANCOND & "')"
' Mod End   2005/11/14 M.Makino

'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup

    ''SQL�̎��s
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001h_Exec = FUNCTION_RETURN_SUCCESS

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

'�T�v      :�e�[�u���uTBCMJ007�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ007 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
Public Function DBDRV_GetTBCMJ007(records() As typ_TBCMJ007, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
' Mod Start 2005/11/14 M.Makino
'    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
'              " PROCCODE, GOUKI, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
'              " SENDFLAG, SENDDATE "
    sqlBase = ""
    sqlBase = sqlBase & "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT"
    sqlBase = sqlBase & ", SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD"
    sqlBase = sqlBase & ", PROCCODE, GOUKI"
    sqlBase = sqlBase & ", nvl(MEAS1, -1) MEAS1"
    sqlBase = sqlBase & ", nvl(MEAS2, -1) MEAS2"
    sqlBase = sqlBase & ", nvl(MEAS3, -1) MEAS3"
    sqlBase = sqlBase & ", nvl(MEAS4, -1) MEAS4"
    sqlBase = sqlBase & ", nvl(MEAS5, -1) MEAS5"
    sqlBase = sqlBase & ", MEASPEAK, CALCMEAS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE"
    sqlBase = sqlBase & ", SENDFLAG, SENDDATE"
    sqlBase = sqlBase & ", nvl(MEAS6, -1) MEAS6"
    sqlBase = sqlBase & ", nvl(MEAS7, -1) MEAS7"
    sqlBase = sqlBase & ", nvl(MEAS8, -1) MEAS8"
    sqlBase = sqlBase & ", nvl(MEAS9, -1) MEAS9"
    sqlBase = sqlBase & ", nvl(MEAS10, -1) MEAS10"
'    sqlBase = sqlBase & ", nvl(MEASFILE, ' ') MEASFILE"
'    sqlBase = sqlBase & ", nvl(RESVAL, -1) RESVAL"
'    sqlBase = sqlBase & ", nvl(INCVAL, -1) INCVAL"
'    sqlBase = sqlBase & ", nvl(CUTVAL, -1) CUTVAL"
'    sqlBase = sqlBase & ", nvl(SETVAL, -1) SETVAL"
'    sqlBase = sqlBase & ", nvl(CONVAL, -1) CONVAL"
'    sqlBase = sqlBase & ", nvl(MEAS1DAT1, -1) MEAS1DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS1DAT2, -1) MEAS1DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS1DAT3, -1) MEAS1DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS1DAT4, -1) MEAS1DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS1DAT5, -1) MEAS1DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS2DAT1, -1) MEAS2DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS2DAT2, -1) MEAS2DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS2DAT3, -1) MEAS2DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS2DAT4, -1) MEAS2DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS2DAT5, -1) MEAS2DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS3DAT1, -1) MEAS3DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS3DAT2, -1) MEAS3DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS3DAT3, -1) MEAS3DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS3DAT4, -1) MEAS3DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS3DAT5, -1) MEAS3DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS4DAT1, -1) MEAS4DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS4DAT2, -1) MEAS4DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS4DAT3, -1) MEAS4DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS4DAT4, -1) MEAS4DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS4DAT5, -1) MEAS4DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS5DAT1, -1) MEAS5DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS5DAT2, -1) MEAS5DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS5DAT3, -1) MEAS5DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS5DAT4, -1) MEAS5DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS5DAT5, -1) MEAS5DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS6DAT1, -1) MEAS6DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS6DAT2, -1) MEAS6DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS6DAT3, -1) MEAS6DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS6DAT4, -1) MEAS6DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS6DAT5, -1) MEAS6DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS7DAT1, -1) MEAS7DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS7DAT2, -1) MEAS7DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS7DAT3, -1) MEAS7DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS7DAT4, -1) MEAS7DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS7DAT5, -1) MEAS7DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS8DAT1, -1) MEAS8DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS8DAT2, -1) MEAS8DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS8DAT3, -1) MEAS8DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS8DAT4, -1) MEAS8DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS8DAT5, -1) MEAS8DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS9DAT1, -1) MEAS9DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS9DAT2, -1) MEAS9DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS9DAT3, -1) MEAS9DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS9DAT4, -1) MEAS9DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS9DAT5, -1) MEAS9DAT5"
'    sqlBase = sqlBase & ", nvl(MEAS10DAT1, -1) MEAS10DAT1"
'    sqlBase = sqlBase & ", nvl(MEAS10DAT2, -1) MEAS10DAT2"
'    sqlBase = sqlBase & ", nvl(MEAS10DAT3, -1) MEAS10DAT3"
'    sqlBase = sqlBase & ", nvl(MEAS10DAT4, -1) MEAS10DAT4"
'    sqlBase = sqlBase & ", nvl(MEAS10DAT5, -1) MEAS10DAT5"
'    sqlBase = sqlBase & ", nvl(LTSPIFLG, -1) LTSPIFLG"
    sqlBase = sqlBase & ", MEASFILE"
    sqlBase = sqlBase & ", RESVAL"
    sqlBase = sqlBase & ", INCVAL"
    sqlBase = sqlBase & ", CUTVAL"
    sqlBase = sqlBase & ", SETVAL"
    sqlBase = sqlBase & ", CONVAL"
    sqlBase = sqlBase & ", MEAS1DAT1"
    sqlBase = sqlBase & ", MEAS1DAT2"
    sqlBase = sqlBase & ", MEAS1DAT3"
    sqlBase = sqlBase & ", MEAS1DAT4"
    sqlBase = sqlBase & ", MEAS1DAT5"
    sqlBase = sqlBase & ", MEAS2DAT1"
    sqlBase = sqlBase & ", MEAS2DAT2"
    sqlBase = sqlBase & ", MEAS2DAT3"
    sqlBase = sqlBase & ", MEAS2DAT4"
    sqlBase = sqlBase & ", MEAS2DAT5"
    sqlBase = sqlBase & ", MEAS3DAT1"
    sqlBase = sqlBase & ", MEAS3DAT2"
    sqlBase = sqlBase & ", MEAS3DAT3"
    sqlBase = sqlBase & ", MEAS3DAT4"
    sqlBase = sqlBase & ", MEAS3DAT5"
    sqlBase = sqlBase & ", MEAS4DAT1"
    sqlBase = sqlBase & ", MEAS4DAT2"
    sqlBase = sqlBase & ", MEAS4DAT3"
    sqlBase = sqlBase & ", MEAS4DAT4"
    sqlBase = sqlBase & ", MEAS4DAT5"
    sqlBase = sqlBase & ", MEAS5DAT1"
    sqlBase = sqlBase & ", MEAS5DAT2"
    sqlBase = sqlBase & ", MEAS5DAT3"
    sqlBase = sqlBase & ", MEAS5DAT4"
    sqlBase = sqlBase & ", MEAS5DAT5"
    sqlBase = sqlBase & ", MEAS6DAT1"
    sqlBase = sqlBase & ", MEAS6DAT2"
    sqlBase = sqlBase & ", MEAS6DAT3"
    sqlBase = sqlBase & ", MEAS6DAT4"
    sqlBase = sqlBase & ", MEAS6DAT5"
    sqlBase = sqlBase & ", MEAS7DAT1"
    sqlBase = sqlBase & ", MEAS7DAT2"
    sqlBase = sqlBase & ", MEAS7DAT3"
    sqlBase = sqlBase & ", MEAS7DAT4"
    sqlBase = sqlBase & ", MEAS7DAT5"
    sqlBase = sqlBase & ", MEAS8DAT1"
    sqlBase = sqlBase & ", MEAS8DAT2"
    sqlBase = sqlBase & ", MEAS8DAT3"
    sqlBase = sqlBase & ", MEAS8DAT4"
    sqlBase = sqlBase & ", MEAS8DAT5"
    sqlBase = sqlBase & ", MEAS9DAT1"
    sqlBase = sqlBase & ", MEAS9DAT2"
    sqlBase = sqlBase & ", MEAS9DAT3"
    sqlBase = sqlBase & ", MEAS9DAT4"
    sqlBase = sqlBase & ", MEAS9DAT5"
    sqlBase = sqlBase & ", MEAS10DAT1"
    sqlBase = sqlBase & ", MEAS10DAT2"
    sqlBase = sqlBase & ", MEAS10DAT3"
    sqlBase = sqlBase & ", MEAS10DAT4"
    sqlBase = sqlBase & ", MEAS10DAT5"
    sqlBase = sqlBase & ", LTSPIFLG"
' Mod End   2005/11/14 M.Makino
    sqlBase = sqlBase & " From TBCMJ007"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
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
' Add Start 2005/11/14 M.Makino
            .MEAS6 = rs("MEAS6")             ' ����l�U
            .MEAS7 = rs("MEAS7")             ' ����l�V
            .MEAS8 = rs("MEAS8")             ' ����l�W
            .MEAS9 = rs("MEAS9")             ' ����l�X
            .MEAS10 = rs("MEAS10")           ' ����l�P�O
'            .MEASFILE = rs("MEASFILE")       ' ����f�[�^�t�@�C����
'            .RESVAL = rs("RESVAL")           ' ������R
'            .INCVAL = rs("INCVAL")           ' �X��
'            .CUTVAL = rs("CUTVAL")           ' �ؕ�
'            .SETVAL = rs("SETVAL")           ' �ݒ�l
'            .CONVAL = rs("RESVAL")           ' 10�����Z�l
'            .MEAS1DAT1 = rs("MEAS1DAT1")     ' ����l�P�@���f�[�^�P
'            .MEAS1DAT2 = rs("MEAS1DAT2")     ' ����l�P�@���f�[�^�Q
'            .MEAS1DAT3 = rs("MEAS1DAT3")     ' ����l�P�@���f�[�^�R
'            .MEAS1DAT4 = rs("MEAS1DAT4")     ' ����l�P�@���f�[�^�S
'            .MEAS1DAT5 = rs("MEAS1DAT5")     ' ����l�P�@���f�[�^�T
'            .MEAS2DAT1 = rs("MEAS2DAT1")     ' ����l�Q�@���f�[�^�P
'            .MEAS2DAT2 = rs("MEAS2DAT2")     ' ����l�Q�@���f�[�^�Q
'            .MEAS2DAT3 = rs("MEAS2DAT3")     ' ����l�Q�@���f�[�^�R
'            .MEAS2DAT4 = rs("MEAS2DAT4")     ' ����l�Q�@���f�[�^�S
'            .MEAS2DAT5 = rs("MEAS2DAT5")     ' ����l�Q�@���f�[�^�T
'            .MEAS3DAT1 = rs("MEAS3DAT1")     ' ����l�R�@���f�[�^�P
'            .MEAS3DAT2 = rs("MEAS3DAT2")     ' ����l�R�@���f�[�^�Q
'            .MEAS3DAT3 = rs("MEAS3DAT3")     ' ����l�R�@���f�[�^�R
'            .MEAS3DAT4 = rs("MEAS3DAT4")     ' ����l�R�@���f�[�^�S
'            .MEAS3DAT5 = rs("MEAS3DAT5")     ' ����l�R�@���f�[�^�T
'            .MEAS4DAT1 = rs("MEAS4DAT1")     ' ����l�S�@���f�[�^�P
'            .MEAS4DAT2 = rs("MEAS4DAT2")     ' ����l�S�@���f�[�^�Q
'            .MEAS4DAT3 = rs("MEAS4DAT3")     ' ����l�S�@���f�[�^�R
'            .MEAS4DAT4 = rs("MEAS4DAT4")     ' ����l�S�@���f�[�^�S
'            .MEAS4DAT5 = rs("MEAS4DAT5")     ' ����l�S�@���f�[�^�T
'            .MEAS5DAT1 = rs("MEAS5DAT1")     ' ����l�T�@���f�[�^�P
'            .MEAS5DAT2 = rs("MEAS5DAT2")     ' ����l�T�@���f�[�^�Q
'            .MEAS5DAT3 = rs("MEAS5DAT3")     ' ����l�T�@���f�[�^�R
'            .MEAS5DAT4 = rs("MEAS5DAT4")     ' ����l�T�@���f�[�^�S
'            .MEAS5DAT5 = rs("MEAS5DAT5")     ' ����l�T�@���f�[�^�T
'            .MEAS6DAT1 = rs("MEAS6DAT1")     ' ����l�U�@���f�[�^�P
'            .MEAS6DAT2 = rs("MEAS6DAT2")     ' ����l�U�@���f�[�^�Q
'            .MEAS6DAT3 = rs("MEAS6DAT3")     ' ����l�U�@���f�[�^�R
'            .MEAS6DAT4 = rs("MEAS6DAT4")     ' ����l�U�@���f�[�^�S
'            .MEAS6DAT5 = rs("MEAS6DAT5")     ' ����l�U�@���f�[�^�T
'            .MEAS7DAT1 = rs("MEAS7DAT1")     ' ����l�V�@���f�[�^�P
'            .MEAS7DAT2 = rs("MEAS7DAT2")     ' ����l�V�@���f�[�^�Q
'            .MEAS7DAT3 = rs("MEAS7DAT3")     ' ����l�V�@���f�[�^�R
'            .MEAS7DAT4 = rs("MEAS7DAT4")     ' ����l�V�@���f�[�^�S
'            .MEAS7DAT5 = rs("MEAS7DAT5")     ' ����l�V�@���f�[�^�T
'            .MEAS8DAT1 = rs("MEAS8DAT1")     ' ����l�W�@���f�[�^�P
'            .MEAS8DAT2 = rs("MEAS8DAT2")     ' ����l�W�@���f�[�^�Q
'            .MEAS8DAT3 = rs("MEAS8DAT3")     ' ����l�W�@���f�[�^�R
'            .MEAS8DAT4 = rs("MEAS8DAT4")     ' ����l�W�@���f�[�^�S
'            .MEAS8DAT5 = rs("MEAS8DAT5")     ' ����l�W�@���f�[�^�T
'            .MEAS9DAT1 = rs("MEAS9DAT1")     ' ����l�X�@���f�[�^�P
'            .MEAS9DAT2 = rs("MEAS9DAT2")     ' ����l�X�@���f�[�^�Q
'            .MEAS9DAT3 = rs("MEAS9DAT3")     ' ����l�X�@���f�[�^�R
'            .MEAS9DAT4 = rs("MEAS9DAT4")     ' ����l�X�@���f�[�^�S
'            .MEAS9DAT5 = rs("MEAS9DAT5")     ' ����l�X�@���f�[�^�T
'            .MEAS10DAT1 = rs("MEAS10DAT1")   ' ����l�P�O�@���f�[�^�P
'            .MEAS10DAT2 = rs("MEAS10DAT2")   ' ����l�P�O�@���f�[�^�Q
'            .MEAS10DAT3 = rs("MEAS10DAT3")   ' ����l�P�O�@���f�[�^�R
'            .MEAS10DAT4 = rs("MEAS10DAT4")   ' ����l�P�O�@���f�[�^�S
'            .MEAS10DAT5 = rs("MEAS10DAT5")   ' ����l�P�O�@���f�[�^�T
'            .LTSPIFLG = rs("LTSPIFLG")       ' ����ʒu����t���O
            .MEASFILE = NulltoStr(rs("MEASFILE"))       ' ����f�[�^�t�@�C����
            .RESVAL = NulltoStr(rs("RESVAL"))           ' ������R
            .INCVAL = NulltoStr(rs("INCVAL"))           ' �X��
            .CUTVAL = NulltoStr(rs("CUTVAL"))           ' �ؕ�
            .SETVAL = NulltoStr(rs("SETVAL"))           ' �ݒ�l
            .CONVAL = NulltoStr(rs("CONVAL"))           ' 10�����Z�l
            .MEAS1DAT1 = NulltoStr(rs("MEAS1DAT1"))     ' ����l�P�@���f�[�^�P
            .MEAS1DAT2 = NulltoStr(rs("MEAS1DAT2"))     ' ����l�P�@���f�[�^�Q
            .MEAS1DAT3 = NulltoStr(rs("MEAS1DAT3"))     ' ����l�P�@���f�[�^�R
            .MEAS1DAT4 = NulltoStr(rs("MEAS1DAT4"))     ' ����l�P�@���f�[�^�S
            .MEAS1DAT5 = NulltoStr(rs("MEAS1DAT5"))     ' ����l�P�@���f�[�^�T
            .MEAS2DAT1 = NulltoStr(rs("MEAS2DAT1"))     ' ����l�Q�@���f�[�^�P
            .MEAS2DAT2 = NulltoStr(rs("MEAS2DAT2"))     ' ����l�Q�@���f�[�^�Q
            .MEAS2DAT3 = NulltoStr(rs("MEAS2DAT3"))     ' ����l�Q�@���f�[�^�R
            .MEAS2DAT4 = NulltoStr(rs("MEAS2DAT4"))     ' ����l�Q�@���f�[�^�S
            .MEAS2DAT5 = NulltoStr(rs("MEAS2DAT5"))     ' ����l�Q�@���f�[�^�T
            .MEAS3DAT1 = NulltoStr(rs("MEAS3DAT1"))     ' ����l�R�@���f�[�^�P
            .MEAS3DAT2 = NulltoStr(rs("MEAS3DAT2"))     ' ����l�R�@���f�[�^�Q
            .MEAS3DAT3 = NulltoStr(rs("MEAS3DAT3"))     ' ����l�R�@���f�[�^�R
            .MEAS3DAT4 = NulltoStr(rs("MEAS3DAT4"))     ' ����l�R�@���f�[�^�S
            .MEAS3DAT5 = NulltoStr(rs("MEAS3DAT5"))     ' ����l�R�@���f�[�^�T
            .MEAS4DAT1 = NulltoStr(rs("MEAS4DAT1"))     ' ����l�S�@���f�[�^�P
            .MEAS4DAT2 = NulltoStr(rs("MEAS4DAT2"))     ' ����l�S�@���f�[�^�Q
            .MEAS4DAT3 = NulltoStr(rs("MEAS4DAT3"))     ' ����l�S�@���f�[�^�R
            .MEAS4DAT4 = NulltoStr(rs("MEAS4DAT4"))     ' ����l�S�@���f�[�^�S
            .MEAS4DAT5 = NulltoStr(rs("MEAS4DAT5"))     ' ����l�S�@���f�[�^�T
            .MEAS5DAT1 = NulltoStr(rs("MEAS5DAT1"))     ' ����l�T�@���f�[�^�P
            .MEAS5DAT2 = NulltoStr(rs("MEAS5DAT2"))     ' ����l�T�@���f�[�^�Q
            .MEAS5DAT3 = NulltoStr(rs("MEAS5DAT3"))     ' ����l�T�@���f�[�^�R
            .MEAS5DAT4 = NulltoStr(rs("MEAS5DAT4"))     ' ����l�T�@���f�[�^�S
            .MEAS5DAT5 = NulltoStr(rs("MEAS5DAT5"))     ' ����l�T�@���f�[�^�T
            .MEAS6DAT1 = NulltoStr(rs("MEAS6DAT1"))     ' ����l�U�@���f�[�^�P
            .MEAS6DAT2 = NulltoStr(rs("MEAS6DAT2"))     ' ����l�U�@���f�[�^�Q
            .MEAS6DAT3 = NulltoStr(rs("MEAS6DAT3"))     ' ����l�U�@���f�[�^�R
            .MEAS6DAT4 = NulltoStr(rs("MEAS6DAT4"))     ' ����l�U�@���f�[�^�S
            .MEAS6DAT5 = NulltoStr(rs("MEAS6DAT5"))     ' ����l�U�@���f�[�^�T
            .MEAS7DAT1 = NulltoStr(rs("MEAS7DAT1"))     ' ����l�V�@���f�[�^�P
            .MEAS7DAT2 = NulltoStr(rs("MEAS7DAT2"))     ' ����l�V�@���f�[�^�Q
            .MEAS7DAT3 = NulltoStr(rs("MEAS7DAT3"))     ' ����l�V�@���f�[�^�R
            .MEAS7DAT4 = NulltoStr(rs("MEAS7DAT4"))     ' ����l�V�@���f�[�^�S
            .MEAS7DAT5 = NulltoStr(rs("MEAS7DAT5"))     ' ����l�V�@���f�[�^�T
            .MEAS8DAT1 = NulltoStr(rs("MEAS8DAT1"))     ' ����l�W�@���f�[�^�P
            .MEAS8DAT2 = NulltoStr(rs("MEAS8DAT2"))     ' ����l�W�@���f�[�^�Q
            .MEAS8DAT3 = NulltoStr(rs("MEAS8DAT3"))     ' ����l�W�@���f�[�^�R
            .MEAS8DAT4 = NulltoStr(rs("MEAS8DAT4"))     ' ����l�W�@���f�[�^�S
            .MEAS8DAT5 = NulltoStr(rs("MEAS8DAT5"))     ' ����l�W�@���f�[�^�T
            .MEAS9DAT1 = NulltoStr(rs("MEAS9DAT1"))     ' ����l�X�@���f�[�^�P
            .MEAS9DAT2 = NulltoStr(rs("MEAS9DAT2"))     ' ����l�X�@���f�[�^�Q
            .MEAS9DAT3 = NulltoStr(rs("MEAS9DAT3"))     ' ����l�X�@���f�[�^�R
            .MEAS9DAT4 = NulltoStr(rs("MEAS9DAT4"))     ' ����l�X�@���f�[�^�S
            .MEAS9DAT5 = NulltoStr(rs("MEAS9DAT5"))     ' ����l�X�@���f�[�^�T
            .MEAS10DAT1 = NulltoStr(rs("MEAS10DAT1"))   ' ����l�P�O�@���f�[�^�P
            .MEAS10DAT2 = NulltoStr(rs("MEAS10DAT2"))   ' ����l�P�O�@���f�[�^�Q
            .MEAS10DAT3 = NulltoStr(rs("MEAS10DAT3"))   ' ����l�P�O�@���f�[�^�R
            .MEAS10DAT4 = NulltoStr(rs("MEAS10DAT4"))   ' ����l�P�O�@���f�[�^�S
            .MEAS10DAT5 = NulltoStr(rs("MEAS10DAT5"))   ' ����l�P�O�@���f�[�^�T
            .LTSPIFLG = Trim(NulltoStr(rs("LTSPIFLG"))) ' ����ʒu����t���O
' Mod End   2005/11/14 M.Makino
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMJ007 = FUNCTION_RETURN_SUCCESS
End Function

'�T�v      :�e�[�u���uKODA9�v��������ɂ������P�O�����Z���ݒ背�R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tbl_OumConv   ,O  ,typ_OumConvSet   ,10�����Z�l�擾�\����
'          :sType         ,I  ,String           ,�^�C�v
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,���o�̐���
'����      :
'����      :2005/11/11�쐬�@�q��
Public Function DBDRV_OumConvGet(record As typ_OumConvSet, sType As String) As Integer

    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet

    DBDRV_OumConvGet = FUNCTION_RETURN_FAILURE

    ' SQL���쐬
    sql = ""
    sql = sql & "SELECT CTR01A9, CTR02A9, CTR03A9"
    sql = sql & " FROM  KODA9"
    sql = sql & " WHERE SYSCA9 = 'X'"
    sql = sql & " AND   SHUCA9 = '19'"
    sql = sql & " AND   CODEA9 = '" & sType & "'"

    ' �P�O�����Z���ݒ���擾����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    If rs Is Nothing Then
        Exit Function
    End If

    ' �Y������f�[�^�������ꍇ����̓G���[
    If rs.EOF Then
        Exit Function
    End If

    ' ���o���ʂ��i�[����
    With record
        ' [�X��]
        .CTR01A9 = Trim(CStr(NulltoStr(rs.Fields("CTR01A9").Value)))
        ' [�ؕ�]
        .CTR02A9 = Trim(CStr(NulltoStr(rs.Fields("CTR02A9").Value)))
        ' [�ݒ�l]
        .CTR03A9 = Trim(CStr(NulltoStr(rs.Fields("CTR03A9").Value)))
    End With
    
    DBDRV_OumConvGet = FUNCTION_RETURN_SUCCESS

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
''Public Function GetKansanchi(tblCrySmpMan As typ_XSDCS, sKekka As String, sIncVal As String, _
''        sCutVal As String, sSetVal As String, sJiteiko As String, sKansanchi As String) As Integer
''    Dim sql As String       'SQL�S��
''    Dim rs As OraDynaset    'RecordSet
''    Dim RironchiLT As Double    ' ���_�lLT
''    Dim Osenryo As Double       ' �����ʐ���l

''    GetKansanchi = FUNCTION_RETURN_FAILURE

    ' SQL���쐬
''    sql = ""
''    sql = sql & "SELECT MEAS1"
''    sql = sql & " FROM  TBCMJ002"
''    sql = sql & " WHERE CRYNUM='" & tblCrySmpMan.XTALCS & "'"
''    sql = sql & " AND   POSITION=" & tblCrySmpMan.INPOSCS
''    sql = sql & " AND   SMPKBN='" & tblCrySmpMan.SMPKBNCS & "'"
''    sql = sql & " AND   TRANCOND='0'"
''    sql = sql & " ORDER BY TRANCNT DESC"

''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
''    If rs Is Nothing Then
''        Exit Function
''    End If

''    If rs.EOF Then
        ' �Y������f�[�^�������ꍇ����͋󕶎�
''        sJiteiko = ""
''    Else
''        sJiteiko = Trim(CStr(NulltoStr(rs.Fields("MEAS1").Value)))
''    End If

    ' �P�O�����Z�l�̌v�Z
''    If sKekka <> "" And sIncVal <> "" And sCutVal <> "" And _
''       sSetVal <> "" And sJiteiko <> "" Then

        '0�̏��Z�΍�
''        On Error GoTo ERROR_CALC

        '�P�O�����Z�l���Z�o
''        RironchiLT = CDbl(sIncVal) * CDbl(sJiteiko) + CDbl(sCutVal)
''        Osenryo = 1 / ((1 / CInt(sKekka)) - (1 / RironchiLT))
''        sKansanchi = CStr(Round(1 / ((1 / CDbl(sSetVal)) + (1 / Osenryo)), 0))
''    Else
''        sKansanchi = ""
''    End If
    
''    GetKansanchi = FUNCTION_RETURN_SUCCESS
''    Exit Function

''ERROR_CALC:
''    sKansanchi = ""
''    GetKansanchi = FUNCTION_RETURN_SUCCESS
''End Function

'
' �󕶎���i""�j�ɑ΂��āwnull�x��Ԃ��C���̑��̕�����͉��������ɕԂ�
'
'����      :2005/11/14�ǉ��@�q��
''Private Function LZeroToNull(ByVal sTmp As String) As String
''    If "" = sTmp Then
''        LZeroToNull = "null"
''    Else
''        LZeroToNull = sTmp
''    End If
''End Function

