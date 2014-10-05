Attribute VB_Name = "s_cmbc001e_SQL"
Option Explicit

' 2007/08/30 SPK Tsutsumi Add Start
Public Type typ_Mukesaki
    sMukeCode As String     '' ����R�[�h
    sMukeName As String     '' ���於
End Type

Public s_MukesakiBase() As typ_Mukesaki
' 2007/08/30 SPK Tsutsumi Add End

' ������������e�i���X

'�T�v      :������������e�i���X ��������X�V�^�}���p�c�a�h���C�o
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:sMkCondNo�@�@�@,I  ,String         �@,���������
'      �@�@:pMkOld   �@�@�@,I  ,typ_TBCMB012   �@,��������I���W�i��
'      �@�@:pMkNew   �@�@�@,I  ,typ_TBCMB012   �@,�������
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������ݐ���
'����      :
'����      :2001/07/30 ���{ �쐬
Public Function DBDRV_scmzc_fcmbc001e_UpdInsMkCond(sMkCondNo As String, pMkOld() As typ_TBCMB012, pMkNew As typ_TBCMB012) As FUNCTION_RETURN

    Dim sql As String
    Dim bFlag As Boolean
    Dim i As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_UpdInsMkCond"

    With pMkNew
        bFlag = False
        For i = 1 To UBound(pMkOld)
            If RTrim$(pMkOld(i).MKCONDNO) = RTrim$(sMkCondNo) Then
                bFlag = True
                Exit For
            End If
        Next i

        If bFlag = True Then
            '' ��������̍X�V
            sql = "update TBCMB012 set "
            sql = sql & "MKCONDNO='" & .MKCONDNO & "', "    ' �������No.
            sql = sql & "MODEL='" & .MODEL & "', "          ' �@��
            sql = sql & "RTBSIZE='" & .RTBSIZE & "', "      ' ���c�{�T�C�Y
            sql = sql & "CHARGE='" & .CHARGE & "', "        ' �`���[�W��
            sql = sql & "HZTYPE='" & .HZTYPE & "', "        ' HZ�^�C�v
            sql = sql & "UPSPDTYP='" & .UPSPDTYP & "', "    ' ���グ���x�^�C�v
            sql = sql & "MAGTYPE='" & .MAGTYPE & "', "      ' ����^�C�v
            sql = sql & "USECLS='0', "                      ' �g�p�敪
            sql = sql & "TSTAFFID='" & .TSTAFFID & "', "    ' �o�^�Ј�ID
            sql = sql & "REGDATE=sysdate, "                 ' �o�^���t
            sql = sql & "KSTAFFID='" & .KSTAFFID & "', "    ' �X�V�Ј�ID
            sql = sql & "UPDDATE=sysdate, "                 ' �X�V���t
            sql = sql & "SENDFLAG='0', "                    ' ���M�t���O
            sql = sql & "SENDDATE=sysdate"                  ' ���M����
            sql = sql & " where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
        Else
            '' ��������̑}��
            sql = "insert into TBCMB012 ("
            sql = sql & "MKCONDNO, "        ' �������No.
            sql = sql & "MODEL, "           ' �@��
            sql = sql & "RTBSIZE, "         ' ���c�{�T�C�Y
            sql = sql & "CHARGE, "          ' �`���[�W��
            sql = sql & "HZTYPE, "          ' HZ�^�C�v
            sql = sql & "UPSPDTYP, "        ' ���グ���x�^�C�v
            sql = sql & "MAGTYPE, "         ' ����^�C�v
            sql = sql & "USECLS, "          ' �g�p�敪
            sql = sql & "TSTAFFID, "        ' �o�^�Ј�ID
            sql = sql & "REGDATE, "         ' �o�^���t
            sql = sql & "KSTAFFID, "        ' �X�V�Ј�ID
            sql = sql & "UPDDATE, "         ' �X�V���t
            sql = sql & "SENDFLAG, "        ' ���M�t���O
            sql = sql & "SENDDATE)"         ' ���M����
            sql = sql & " values ('"
            sql = sql & .MKCONDNO & "', '"  ' �������No.
            sql = sql & .MODEL & "', '"     ' �@��
            sql = sql & .RTBSIZE & "', '"   ' ���c�{�T�C�Y
            sql = sql & .CHARGE & "', '"    ' �`���[�W��
            sql = sql & .HZTYPE & "', '"    ' HZ�^�C�v
            sql = sql & .UPSPDTYP & "', '"  ' ���グ���x�^�C�v
            sql = sql & .MAGTYPE & "', "    ' ����^�C�v
            sql = sql & "'0', '"            ' �g�p�敪
            sql = sql & .TSTAFFID & "', "   ' �o�^�Ј�ID
            sql = sql & "sysdate, '"        ' �o�^���t
            sql = sql & .KSTAFFID & "', "   ' �X�V�Ј�ID
            sql = sql & "sysdate, "         ' �X�V���t
            sql = sql & "'0', "             ' ���M�t���O
            sql = sql & "sysdate)"          ' ���M����
        End If
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :������������e�i���X ��������폜�p�c�a�h���C�o
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:sMkCondNo�@�@�@,I  ,String         �@,���������
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������ݐ���
'����      :
'����      :2001/07/30 ���{ �쐬
Public Function DBDRV_scmzc_fcmbc001e_DelMkCond(sMkCondNo As String) As FUNCTION_RETURN

    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_DelMkCond"

    '' ��������̍폜
    sql = "delete TBCMB012 where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :������������e�i���X �������PG-ID�Ή��X�V�^�}���p�c�a�h���C�o
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:sMkCondNo�@�@�@,I  ,String         �@,���������
'      �@�@:pPGIDOld �@�@�@,I  ,typ_TBCMB013   �@,�������PG-ID�Ή��I���W�i��
'      �@�@:pPGIDNew �@�@�@,I  ,typ_TBCMB013   �@,�������PG-ID�Ή�
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������ݐ���
'����      :
'����      :2001/07/30 ���{ �쐬
Public Function DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng(sMkCondNo As String, pPGIDOld() As typ_TBCMB013, pPGIDNew() As typ_TBCMB013) As FUNCTION_RETURN

    Dim sql As String
    Dim bFlag As Boolean
    Dim i As Long
    Dim j As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_InsPGIDMng"

    For i = 1 To UBound(pPGIDNew)
        With pPGIDNew(i)
            bFlag = False
            For j = 1 To UBound(pPGIDOld)
                If RTrim$(pPGIDOld(j).MKCONDNO) = RTrim$(sMkCondNo) And _
                   RTrim$(pPGIDOld(j).PGIDNO) = RTrim$(.PGIDNO) Then
                    bFlag = True
                    Exit For
                End If
            Next j

            If bFlag = True Then
                '' �������PG-ID�Ή��̍X�V
                sql = "update TBCMB013 set "
                sql = sql & "MKCONDNO='" & .MKCONDNO & "', "    ' �������No.
                sql = sql & "PGIDNO='" & .PGIDNO & "', "        ' PG-IDNo
                sql = sql & "TSTAFFID='" & .TSTAFFID & "', "    ' �o�^�Ј�ID
                sql = sql & "REGDATE=sysdate, "                 ' �o�^���t
                sql = sql & "KSTAFFID='" & .KSTAFFID & "', "    ' �X�V�Ј�ID
                sql = sql & "UPDDATE=sysdate, "                 ' �X�V���t
                sql = sql & "SENDFLAG='0', "                    ' ���M�t���O
                sql = sql & "SENDDATE=sysdate"                  ' ���M���t
                sql = sql & " where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
                sql = sql & " and rtrim(PGIDNO)='" & RTrim$(.PGIDNO) & "'"
            Else
                '' �������PG-ID�Ή��̑}��
                sql = "insert into TBCMB013 ("
                sql = sql & "MKCONDNO, "        ' �������No.
                sql = sql & "PGIDNO, "          ' PG-IDNo
                sql = sql & "TSTAFFID, "        ' �o�^�Ј�ID
                sql = sql & "REGDATE, "         ' �o�^���t
                sql = sql & "KSTAFFID, "        ' �X�V�Ј�ID
                sql = sql & "UPDDATE, "         ' �X�V���t
                sql = sql & "SENDFLAG, "        ' ���M�t���O
                sql = sql & "SENDDATE)"         ' ���M���t
                sql = sql & " values ('"
                sql = sql & .MKCONDNO & "', '"  ' �������No.
                sql = sql & .PGIDNO & "', '"    ' PG-IDNo
                sql = sql & .TSTAFFID & "', "   ' �o�^�Ј�ID
                sql = sql & "sysdate, '"        ' �o�^���t
                sql = sql & .KSTAFFID & "', "   ' �X�V�Ј�ID
                sql = sql & "sysdate, "         ' �X�V���t
                sql = sql & "'0', "             ' ���M�t���O
                sql = sql & "sysdate)"          ' ���M���t
            End If
        End With
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :������������e�i���X �������PG-ID�Ή��폜�p�c�a�h���C�o
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:sMkCondNo�@�@�@,I  ,String         �@,���������
'      �@�@:sPGIDNo  �@�@�@,I  ,String         �@,PG-ID��
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������ݐ���
'����      :
'����      :2001/07/30 ���{ �쐬
Public Function DBDRV_scmzc_fcmbc001e_DelPGIDMng(sMkCondNo As String, sPGIDNo As String) As FUNCTION_RETURN

    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_DelPGIDMng"

    '' �������PG-ID�Ή��̍폜
    sql = "delete TBCMB013 where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
    If RTrim$(sPGIDNo) <> "" Then
        sql = sql & " and rtrim(PGIDNO)='" & RTrim$(sPGIDNo) & "'"
    End If
    If OraDB.ExecuteSQL(sql) < 0 Then
        DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMB011�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB011 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMB011_SQL.bas���ړ�)
'           2010/ 1/ 5 SUMCO Akizuki �Q�Ɛ�FROM TBCMB011��TBCMB011_CCV�ɕύX

Public Function DBDRV_GetTBCMB011(records() As typ_TBCMB011, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select PGID, HZPART, HZPTRN, SPACER, UPRING, CHARGE, RTBPOS, RTBSIZE, GAP, UPDM, UPLENGTH, UPRC, RFRNEED, UPSPIN," & _
              " DOWNSPIN, ROPRESS, ARUGON, AIMOIMIN, AIMOIMAX, HCCLASS, HC, AVEUPSPD, UPCNTL, BTMSHAPE, MAGSTR, MAGPOS, CONDGRT," & _
              " MODEL, UPMETHOD, UPCLASS, UPNUM, OPETIME, WTRCOOL, PGID2, RCPT1, RCPT2, RCPT3, RCPT4, RCPT5, CNTL1, CNTL2," & _
              " CNTL3, CNTL4, CNTL5, CNTL6, CNTL7, CNTL8, CNTL9, CNTL10, CNTL11, CNTL12, CNTL13, CNTL14, CNTL15, RUNCOND1," & _
              " RUNCOND2, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCMB011"
    sqlBase = sqlBase & "From TBCMB011_CCV"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB011 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .PGID = rs("PGID")               ' PG-ID
            .HZPART = rs("HZPART")           ' HZ�p�[�c
            .HZPTRN = rs("HZPTRN")           ' HZ�p�^�[��
            .SPACER = rs("SPACER")           ' �X�y�[�T
            .UPRING = rs("UPRING")           ' �A�b�p�[�����O
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .RTBPOS = rs("RTBPOS")           ' ���c�{�ʒu
            .RTBSIZE = rs("RTBSIZE")         ' ���c�{�T�C�Y
            .GAP = rs("GAP")                 ' �M���b�v
            .UPDM = rs("UPDM")               ' ���㒼�a
            .UPLENGTH = rs("UPLENGTH")       ' ���㒷�i�S���j
            .UPRC = rs("UPRC")               ' ����iRC�j
            .RFRNEED = rs("RFRNEED")         ' ���t���N�^�v��
            .UPSPIN = rs("UPSPIN")           ' �㎲��]��
            .DOWNSPIN = rs("DOWNSPIN")       ' ������]��
            .ROPRESS = rs("ROPRESS")         ' �F����
            .ARUGON = rs("ARUGON")           ' �A���S����
            .AIMOIMIN = rs("AIMOIMIN")       ' �˂炢Oi�iMIN)
            .AIMOIMAX = rs("AIMOIMAX")       ' �˂炢Oi�iMAX)
            .HCCLASS = rs("HCCLASS")         ' HC���
            .HC = rs("HC")                   ' HC
            .AVEUPSPD = rs("AVEUPSPD")       ' ���ψ��㑬�x
            .UPCNTL = rs("UPCNTL")           ' ���㐧��
            .BTMSHAPE = rs("BTMSHAPE")       ' �{�g���`��
            .MAGSTR = rs("MAGSTR")           ' ���ꋭ�x
            .MAGPOS = rs("MAGPOS")           ' ����ʒu
            .CONDGRT = rs("CONDGRT")         ' �����ۏؓo�^
            .MODEL = rs("MODEL")             ' �@��
            .UPMETHOD = rs("UPMETHOD")       ' ������@
            .UPCLASS = rs("UPCLASS")         ' ����敪
            .UPNUM = rs("UPNUM")             ' ����{��
            .OPETIME = rs("OPETIME")         ' �^�]����
            .WTRCOOL = rs("WTRCOOL")         ' ����Ǘv��
            .PGID2 = rs("PGID2")             ' PG-ID�i��{���j
            .RCPT1 = rs("RCPT1")             ' �Ή����V�sNo�iT1)
            .RCPT2 = rs("RCPT2")             ' �Ή����V�sNo�iT2)
            .RCPT3 = rs("RCPT3")             ' �Ή����V�sNo�iT3)
            .RCPT4 = rs("RCPT4")             ' �Ή����V�sNo�iT4)
            .RCPT5 = rs("RCPT5")             ' �Ή����V�sNo�iT5)
            .CNTL1 = rs("CNTL1")             ' �������ځi1�j
            .CNTL2 = rs("CNTL2")             ' �������ځi2�j
            .CNTL3 = rs("CNTL3")             ' �������ځi3�j
            .CNTL4 = rs("CNTL4")             ' �������ځi4�j
            .CNTL5 = rs("CNTL5")             ' �������ځi5�j
            .CNTL6 = rs("CNTL6")             ' �������ځi6�j
            .CNTL7 = rs("CNTL7")             ' �������ځi7�j
            .CNTL8 = rs("CNTL8")             ' �������ځi8�j
            .CNTL9 = rs("CNTL9")             ' �������ځi9�j
            .CNTL10 = rs("CNTL10")           ' �������ځi10�j
            .CNTL11 = rs("CNTL11")           ' �������ځi11�j
            .CNTL12 = rs("CNTL12")           ' �������ځi12�j
            .CNTL13 = rs("CNTL13")           ' �������ځi13�j
            .CNTL14 = rs("CNTL14")           ' �������ځi14�j
            .CNTL15 = rs("CNTL15")           ' �������ځi15�j
            .RUNCOND1 = rs("RUNCOND1")       ' �^�]�����P
            .RUNCOND2 = rs("RUNCOND2")       ' �^�]�����Q
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

    DBDRV_GetTBCMB011 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMB012�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB012 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMB012_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMB012(records() As typ_TBCMB012, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select MKCONDNO, MODEL, RTBSIZE, CHARGE, HZTYPE, UPSPDTYP, MAGTYPE, USECLS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
              " SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB012"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB012 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .MKCONDNO = rs("MKCONDNO")       ' �������No.
            .MODEL = rs("MODEL")             ' �@��
            .RTBSIZE = rs("RTBSIZE")         ' ���c�{�T�C�Y
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .HZTYPE = rs("HZTYPE")           ' HZ�^�C�v
            .UPSPDTYP = rs("UPSPDTYP")       ' ���グ���x�^�C�v
            .MAGTYPE = rs("MAGTYPE")         ' ����^�C�v
            .USECLS = rs("USECLS")           ' �g�p�敪
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M����
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB012 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMB013�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB013 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMB013_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMB013(records() As typ_TBCMB013, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select MKCONDNO, PGIDNO, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB013"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB013 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .MKCONDNO = rs("MKCONDNO")       ' �������No.
            .PGIDNO = rs("PGIDNO")           ' PG-IDNo
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

    DBDRV_GetTBCMB013 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMB005�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB005 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMB005_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMB005(records() As typ_TBCMB005, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select SYSCLASS, CLASS, CODE, INFO1, INFO2, INFO3, INFO4, INFO5, INFO6, INFO7, INFO8, INFO9, NOTE, TSTAFFID," & _
              " REGDATE, KSTAFFID, UPDDATE "
    sqlBase = sqlBase & "From TBCMB005"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB005 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .SYSCLASS = rs("SYSCLASS")       ' �V�X�e���敪
            .Class = rs("CLASS")             ' �敪
            .CODE = rs("CODE")               ' �R�[�h
            .INFO1 = rs("INFO1")             ' ���P
            .INFO2 = rs("INFO2")             ' ���Q
            .INFO3 = rs("INFO3")             ' ���R
            .INFO4 = rs("INFO4")             ' ���S
            .INFO5 = rs("INFO5")             ' ���T
            .INFO6 = rs("INFO6")             ' ���U
            .INFO7 = rs("INFO7")             ' ���V
            .INFO8 = rs("INFO8")             ' ���W
            .INFO9 = rs("INFO9")             ' ���X
            .NOTE = rs("NOTE")               ' ���l
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB005 = FUNCTION_RETURN_SUCCESS
End Function

' 2007/09/04 SPK Tsutsumi Add Start
Public Function GetMukeCode() As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Long      '���R�[�h��
    Dim i  As Long
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    GetMukeCode = FUNCTION_RETURN_FAILURE
    
    sql = "Select CODEA9,NAMEJA9 "
    sql = sql & "from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '20' "
    sql = sql & "and (CODEA9 = '14' "
    sql = sql & "or CODEA9 = '15' "
    sql = sql & "or CODEA9 = '16' "
    sql = sql & "or CODEA9 = 'ALL') "

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If
    
    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim s_MukesakiBase(recCnt)
    
    If recCnt = 0 Then
        Exit Function
    End If
    
    For i = 1 To recCnt
        With s_MukesakiBase(i)
            If IsNull(rs.Fields("CODEA9")) = False Then
                .sMukeCode = rs.Fields("CODEA9")    ' ����R�[�h
            End If
            
            If IsNull(rs.Fields("NAMEJA9")) = False Then
                .sMukeName = rs.Fields("NAMEJA9")  ' ���於
'                f_cmbc061_0.cmbMukesaki.AddItem .sMukeName
            End If
        End With
        rs.MoveNext
    Next
    rs.Close

    GetMukeCode = FUNCTION_RETURN_SUCCESS
proc_exit:
    '�I��
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
' 2007/09/04 SPK Tsutsumi Add End
