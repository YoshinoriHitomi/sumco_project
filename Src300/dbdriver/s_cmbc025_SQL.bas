Attribute VB_Name = "s_cmbc025_SQL"
Option Explicit
'
'                                     2001/06/20
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMJ005 (�n�r�e����)
' �Q�Ɓ@�@: 060211_��������
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmjc001f_Disp
    
   ' CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
   ' TRANCNT As Integer              ' ������
    SMPLNO As Integer               ' �T���v���m��
    SMPLUMU As String * 1           ' �T���v���L��
    HINBAN As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    MEASMETH As String * 1          ' ������@
    MEASSPOT As Integer             ' ����_
    MAG As String * 4               ' �{��
    HTPRC As String * 2             ' �M�������@
    KKSP As String * 3              ' �������ב���ʒu
    KKSET As String * 3             ' �������ב�������{�I��ET��@�@char(1)�{number(2)
    CALCMAX As Double               ' �v�Z���� Max
    CALCAVE As Double               ' �v�Z���� Ave
    MEAS1 As Double                 ' ����l�P
    MEAS2 As Double                 ' ����l�Q
    MEAS3 As Double                 ' ����l�R
    MEAS4 As Double                 ' ����l�S
    MEAS5 As Double                 ' ����l�T
    MEAS6 As Double                 ' ����l�U
    MEAS7 As Double                 ' ����l�V
    MEAS8 As Double                 ' ����l�W
    MEAS9 As Double                 ' ����l�X
    MEAS10 As Double                ' ����l�P�O
    MEAS11 As Double                ' ����l�P�P
    MEAS12 As Double                ' ����l�P�Q
    MEAS13 As Double                ' ����l�P�R
    MEAS14 As Double                ' ����l�P�S
    MEAS15 As Double                ' ����l�P�T
    MEAS16 As Double                ' ����l�P�U
    MEAS17 As Double                ' ����l�P�V
    MEAS18 As Double                ' ����l�P�W
    MEAS19 As Double                ' ����l�P�X
    MEAS20 As Double                ' ����l�Q�O
   ' TSTAFFID As String * 8          ' �o�^�Ј�ID
   ' REGDATE As Date                 ' �o�^���t
   ' KSTAFFID As String * 8          ' �X�V�Ј�ID
   ' UPDDATE As Date                 ' �X�V���t
   ' SENDFLAG As String * 1          ' ���M�t���O
   ' SENDDATE As Date                ' ���M���t
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
    OSFPOS1 As Double               ' ����݋敪�P�ʒu
    OSFWID1 As Double               ' ����݋敪�P��
    OSFRD1  As String               ' ����݋敪�PR/D
    OSFPOS2 As Double               ' ����݋敪�Q�ʒu
    OSFWID2 As Double               ' ����݋敪�Q��
    OSFRD2  As String               ' ����݋敪�QR/D
    OSFPOS3 As Double               ' ����݋敪�R�ʒu
    OSFWID3 As Double               ' ����݋敪�R��
    OSFRD3  As String               ' ����݋敪�RR/D
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
  
End Type

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ005�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_cmjc001f_Disp ,���o���R�[�h
'          :SPLNUMs()     ,I  ,Integer      ,���o�����z��(�T���v��No)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/20�쐬�@����
Public Function DBDRV_Getcmjc001f_Disp(records() As typ_cmjc001f_Disp, SMPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere As String  'SQL��WHERE����
Dim sqlGroup As String  'SQL��GROUP����
Dim sqlOrder As String  'SQL��Order����
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
 
  
    DBDRV_Getcmjc001f_Disp = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmjc001f_SQL.bas -- Function DBDRV_Getcmjc001f_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, CALCMAX, CALCAVE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEAS6," & _
              " MEAS7, MEAS8, MEAS9, MEAS10, MEAS11, MEAS12, MEAS13, MEAS14, MEAS15, MEAS16, MEAS17, MEAS18, MEAS19, MEAS20, " & _
              " OSFPOS1, OSFWID1, OSFRD1, OSFPOS2, OSFWID2, OSFRD2, OSFPOS3, OSFWID3, OSFRD3 "
    sqlBase = sqlBase & "From TBCMJ005"
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
        DBDRV_Getcmjc001f_Disp = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
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
            .HINBAN = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .FACTORY = rs("FACTORY")         ' �H��
            .OPECOND = rs("OPECOND")         ' ���Ə���
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .GOUKI = rs("GOUKI")             ' ���@
            .MEASMETH = rs("MEASMETH")       ' ������@
            .MEASSPOT = rs("MEASSPOT")       ' ����_
            .MAG = rs("MAG")                 ' �{��
            .HTPRC = rs("HTPRC")             ' �M�������@
            .KKSP = rs("KKSP")               ' �������ב���ʒu
            .KKSET = rs("KKSET")             ' �������ב�������{�I��ET��
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
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
            .OSFPOS1 = rs("OSFPOS1")         ' ����݋敪�P�ʒu
            .OSFWID1 = rs("OSFWID1")         ' ����݋敪�P��
            .OSFRD1 = rs("OSFRD1")           ' ����݋敪�PR/D
            .OSFPOS2 = rs("OSFPOS2")         ' ����݋敪�Q�ʒu
            .OSFWID2 = rs("OSFWID2")         ' ����݋敪�Q��
            .OSFRD2 = rs("OSFRD2")           ' ����݋敪�QR/D
            .OSFPOS3 = rs("OSFPOS3")         ' ����݋敪�R�ʒu
            .OSFWID3 = rs("OSFWID3")         ' ����݋敪�R��
            .OSFRD3 = rs("OSFRD3")           ' ����݋敪�RR/D
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001f_Disp = FUNCTION_RETURN_SUCCESS

PROC_EXIT:
    '�I��
    gErr.Pop
    Exit Function

PROC_ERR:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume PROC_EXIT
End Function

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�����œn���ꂽ���R�[�h��TBCMJ005�ɒǉ�����
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :record        ,I  ,typ_cmjc001f_Disp ,���o���R�[�h
'          :CRYNUM        ,I  ,String       ,�����ԍ�
'          :TSTAFFID      ,I  ,String       ,�o�^�Ј�ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�����񐔂̓e�[�u����̍ő�l+1�Ƃ���B
'����      :2001/06/22(Fri)�쐬�@����

Public Function DBDRV_Getcmjc001f_Exec(record As typ_cmjc001f_Disp, Crynum$, TSTAFFID$) As FUNCTION_RETURN

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
     
     
    DBDRV_Getcmjc001f_Exec = FUNCTION_RETURN_FAILURE
    
    If Left(record.OSFRD1, 1) = "R" Then
       record.OSFRD1 = "R"
    ElseIf Left(record.OSFRD1, 1) = "D" Then
       record.OSFRD1 = "D"
    Else
       record.OSFRD1 = "-"
    End If
    
    If Left(record.OSFRD2, 1) = "R" Then
       record.OSFRD2 = "R"
    ElseIf Left(record.OSFRD2, 1) = "D" Then
       record.OSFRD2 = "D"
    Else
       record.OSFRD2 = "-"
    End If
    
    If Left(record.OSFRD3, 1) = "R" Then
       record.OSFRD3 = "R"
    ElseIf Left(record.OSFRD3, 1) = "D" Then
       record.OSFRD3 = "D"
    Else
       record.OSFRD3 = "-"
    End If
    
    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmjc001f_SQL.bas -- Function DBDRV_Getcmjc001f_Exec"

    sqlBase = "Insert into TBCMJ005 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MEASMETH, MEASSPOT, MAG, HTPRC, KKSP, KKSET, CALCMAX, CALCAVE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEAS6, " & _
              "MEAS7, MEAS8, MEAS9, MEAS10, MEAS11, MEAS12, MEAS13, MEAS14, MEAS15, MEAS16, MEAS17, MEAS18, MEAS19, MEAS20, " & _
              "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, OSFPOS1, OSFWID1, OSFRD1, OSFPOS2, OSFWID2, OSFRD2, OSFPOS3, OSFWID3, OSFRD3 ) "
    sqlBase = sqlBase & "select '" & Crynum & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.HINBAN & "', " & record.REVNUM & ", '" & record.FACTORY & "', '" & _
               record.OPECOND & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', '" & record.MEASMETH & "', " & _
               record.MEASSPOT & ", '" & record.MAG & "', '" & record.HTPRC & "', '" & record.KKSP & "', '" & record.KKSET & "', " & _
               record.CALCMAX & ", " & record.CALCAVE & ", " & record.MEAS1 & ", " & record.MEAS2 & ", " & record.MEAS3 & ", " & record.MEAS4 & ", " & _
               record.MEAS5 & ", " & record.MEAS6 & ", " & record.MEAS7 & ", " & record.MEAS8 & ", " & record.MEAS9 & ", " & record.MEAS10 & ", " & _
               record.MEAS11 & ", " & record.MEAS12 & ", " & record.MEAS13 & ", " & record.MEAS14 & ", " & record.MEAS15 & ", " & record.MEAS16 & ", " & _
               record.MEAS17 & ", " & record.MEAS18 & ", " & record.MEAS19 & ", " & record.MEAS20 & ", '" & _
               TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE , " & _
               record.OSFPOS1 & ", " & record.OSFWID1 & ", '" & record.OSFRD1 & "', " & _
               record.OSFPOS2 & ", " & record.OSFWID2 & ", '" & record.OSFRD2 & "', " & _
               record.OSFPOS3 & ", " & record.OSFWID3 & ", '" & record.OSFRD3 & "'" & " From TBCMJ005 "
    sqlWhere = "where (CRYNUM='" & Crynum & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
'yaz
Debug.Print sql
    
    ''SQL�̎��s
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001f_Exec = FUNCTION_RETURN_SUCCESS

PROC_EXIT:
    '�I��
    gErr.Pop
    Exit Function

PROC_ERR:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume PROC_EXIT
End Function




'�T�v      :�f�[�^�ϊ����s��
'���Ұ�    :�ϐ���        ,IO ,�^                ,����
'          :tblLeft       ,IO   ,typ_TBCMJ005      ,�e�[�u���f�[�^�P
'          :tblRight      ,IO   ,typ_cmjc001f_Disp ,�e�[�u���f�[�^�Q
'          :bFlg          ,I   ,Boolean           ,TRUE:�����P�f�[�^�������Q�f�[�^�ւ̕ϊ�  FALSE:�����P�f�[�^�������Q�f�[�^�ւ̕ϊ�
'����      :
Public Sub ConvDate_F_cmjc001f_a(tblLeft As typ_TBCMJ005, tblRight As typ_cmjc001f_Disp, bFlg As Boolean)
    If bFlg = True Then
        With tblRight
            .POSITION = tblLeft.POSITION
            .SMPKBN = tblLeft.SMPKBN
            .TRANCOND = tblLeft.TRANCOND
            .SMPLNO = tblLeft.SMPLNO
            .SMPLUMU = tblLeft.SMPLUMU
            .HINBAN = tblLeft.HINBAN
            .REVNUM = tblLeft.REVNUM
            .FACTORY = tblLeft.FACTORY
            .OPECOND = tblLeft.OPECOND
            .KRPROCCD = tblLeft.KRPROCCD
            .PROCCODE = tblLeft.PROCCODE
            .GOUKI = tblLeft.GOUKI
            .MEASMETH = tblLeft.MEASMETH
            .MEASSPOT = tblLeft.MEASSPOT
            .MAG = tblLeft.MAG
            .HTPRC = tblLeft.HTPRC
            .KKSP = tblLeft.KKSP
            .KKSET = tblLeft.KKSET
            .CALCMAX = tblLeft.CALCMAX
            .CALCAVE = tblLeft.CALCAVE
            .MEAS1 = tblLeft.MEAS1
            .MEAS2 = tblLeft.MEAS2
            .MEAS3 = tblLeft.MEAS3
            .MEAS4 = tblLeft.MEAS4
            .MEAS5 = tblLeft.MEAS5
            .MEAS6 = tblLeft.MEAS6
            .MEAS7 = tblLeft.MEAS7
            .MEAS8 = tblLeft.MEAS8
            .MEAS9 = tblLeft.MEAS9
            .MEAS10 = tblLeft.MEAS10
            .MEAS11 = tblLeft.MEAS11
            .MEAS12 = tblLeft.MEAS12
            .MEAS13 = tblLeft.MEAS13
            .MEAS14 = tblLeft.MEAS14
            .MEAS15 = tblLeft.MEAS15
            .MEAS16 = tblLeft.MEAS16
            .MEAS17 = tblLeft.MEAS17
            .MEAS18 = tblLeft.MEAS18
            .MEAS19 = tblLeft.MEAS19
            .MEAS20 = tblLeft.MEAS20
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
            .OSFPOS1 = tblLeft.OSFPOS1
            .OSFWID1 = tblLeft.OSFWID1
            .OSFRD1 = tblLeft.OSFRD1
            .OSFPOS2 = tblLeft.OSFPOS2
            .OSFWID2 = tblLeft.OSFWID2
            .OSFRD2 = tblLeft.OSFRD2
            .OSFPOS3 = tblLeft.OSFPOS3
            .OSFWID3 = tblLeft.OSFWID3
            .OSFRD3 = tblLeft.OSFRD3
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
        End With
    Else
        With tblLeft
            .POSITION = tblRight.POSITION
            .SMPKBN = tblRight.SMPKBN
            .TRANCOND = tblRight.TRANCOND
            .SMPLNO = tblRight.SMPLNO
            .SMPLUMU = tblRight.SMPLUMU
            .HINBAN = tblRight.HINBAN
            .REVNUM = tblRight.REVNUM
            .FACTORY = tblRight.FACTORY
            .OPECOND = tblRight.OPECOND
            .KRPROCCD = tblRight.KRPROCCD
            .PROCCODE = tblRight.PROCCODE
            .GOUKI = tblRight.GOUKI
            .MEASMETH = tblRight.MEASMETH
            .MEASSPOT = tblRight.MEASSPOT
            .MAG = tblRight.MAG
            .HTPRC = tblRight.HTPRC
            .KKSP = tblRight.KKSP
            .KKSET = tblRight.KKSET
            .CALCMAX = tblRight.CALCMAX
            .CALCAVE = tblRight.CALCAVE
            .MEAS1 = tblRight.MEAS1
            .MEAS2 = tblRight.MEAS2
            .MEAS3 = tblRight.MEAS3
            .MEAS4 = tblRight.MEAS4
            .MEAS5 = tblRight.MEAS5
            .MEAS6 = tblRight.MEAS6
            .MEAS7 = tblRight.MEAS7
            .MEAS8 = tblRight.MEAS8
            .MEAS9 = tblRight.MEAS9
            .MEAS10 = tblRight.MEAS10
            .MEAS11 = tblRight.MEAS11
            .MEAS12 = tblRight.MEAS12
            .MEAS13 = tblRight.MEAS13
            .MEAS14 = tblRight.MEAS14
            .MEAS15 = tblRight.MEAS15
            .MEAS16 = tblRight.MEAS16
            .MEAS17 = tblRight.MEAS17
            .MEAS18 = tblRight.MEAS18
            .MEAS19 = tblRight.MEAS19
            .MEAS20 = tblRight.MEAS20
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
            .OSFPOS1 = tblRight.OSFPOS1
            .OSFWID1 = tblRight.OSFWID1
            .OSFRD1 = tblRight.OSFRD1
            .OSFPOS2 = tblRight.OSFPOS2
            .OSFWID2 = tblRight.OSFWID2
            .OSFRD2 = tblRight.OSFRD2
            .OSFPOS3 = tblRight.OSFPOS3
            .OSFWID3 = tblRight.OSFWID3
            .OSFRD3 = tblRight.OSFRD3
'OSF�CBMD���ڒǉ��Ή�  2002.04.02 yakimura
        End With
    End If

End Sub

