Attribute VB_Name = "s_cmcc001_SQL"
Option Explicit
'                                     2001/06/11
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMB011 (PG-ID�Ǘ�)
' �Q�Ɓ@�@: 060200_�S�e�[�u��
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmbc001c_Disp
    PGID As String * 10             ' PG-ID
    HZPART As String * 4            ' HZ�p�[�c
    AIMOIMIN As Double              ' �˂炢Oi�iMIN)
    AIMOIMAX As Double              ' �˂炢Oi�iMAX)
    AVEUPSPD As Double              ' ���ψ��㑬�x
    MODEL As String * 4             ' �@��
    UPMETHOD As String * 1          ' ������@
    UPCLASS As String * 2           ' ����敪
    REGDATE As Date                 ' �o�^���t
    UPDDATE As Date                 ' �X�V���t
End Type



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
'����      :2001/06/11�쐬�@�쑺
Public Function DBDRV_cmbc001c_Disp(records() As typ_cmbc001c_Disp, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001c_SQL.bas -- Function DBDRV_cmbc001c_Disp"
    
    DBDRV_cmbc001c_Disp = FUNCTION_RETURN_FAILURE
    
    sql = "Select PGID, HZPART, AIMOIMIN, AIMOIMAX, AVEUPSPD, MODEL, UPMETHOD, UPCLASS, REGDATE, UPDDATE "
    sql = sql & "From TBCMB011 "
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & sqlWhere & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_cmbc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .PGID = rs("PGID")               ' PG-ID
            .HZPART = rs("HZPART")           ' HZ�p�[�c
'            .HZPTRN = rs("HZPTRN")           ' HZ�p�^�[��
            .AIMOIMIN = rs("AIMOIMIN")       ' �˂炢Oi�iMIN)
            .AIMOIMAX = rs("AIMOIMAX")       ' �˂炢Oi�iMAX)
            .AVEUPSPD = rs("AVEUPSPD")       ' ���ψ��㑬�x
            .MODEL = rs("MODEL")             ' �@��
            .UPMETHOD = rs("UPMETHOD")       ' ������@
            .UPCLASS = rs("UPCLASS")         ' ����敪
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_cmbc001c_Disp = FUNCTION_RETURN_SUCCESS

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


'�T�v      :�e�[�u���uTBCMB011�v�֑}��
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB011 ,���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/10/04 �쐬�@���{
Public Function DBDRV_cmbc001c_Ins(staff As String, records() As typ_VAX_DR_CNDS) As FUNCTION_RETURN

    Dim i As Long
    Dim sql As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001c_SQL.bas -- Function DBDRV_cmbc001c_Ins"

    ' PG-ID�Ǘ�
    For i = 1 To UBound(records)
        sql = "insert into TBCMB011 ( "
        sql = sql & "PGID, "             ' PG-ID
        sql = sql & "HZPART, "           ' HZ�p�[�c
        sql = sql & "HZPTRN, "           ' HZ�p�^�[��
        sql = sql & "SPACER, "           ' �X�y�[�T
        sql = sql & "UPRING, "           ' �A�b�p�[�����O
        sql = sql & "CHARGE, "           ' �`���[�W��
        sql = sql & "RTBPOS, "           ' ���c�{�ʒu
        sql = sql & "RTBSIZE, "          ' ���c�{�T�C�Y
        sql = sql & "GAP, "              ' �M���b�v
        sql = sql & "UPDM, "             ' ���㒼�a
        sql = sql & "UPLENGTH, "         ' ���㒷�i�S���j
        sql = sql & "UPRC, "             ' ����iRC�j
        sql = sql & "RFRNEED, "          ' ���t���N�^�v��
        sql = sql & "UPSPIN, "           ' �㎲��]��
        sql = sql & "DOWNSPIN, "         ' ������]��
        sql = sql & "ROPRESS, "          ' �F����
        sql = sql & "ARUGON, "           ' �A���S����
        sql = sql & "DRDOP,"
        sql = sql & "DRAR3,"
        sql = sql & "AIMOIMIN, "         ' �˂炢Oi�iMIN)
        sql = sql & "AIMOIMAX, "         ' �˂炢Oi�iMAX)
        sql = sql & "HCCLASS, "          ' HC���
        sql = sql & "HC, "               ' HC
        sql = sql & "AVEUPSPD, "         ' ���ψ��㑬�x
        sql = sql & "UPCNTL, "           ' ���㐧��
        sql = sql & "BTMSHAPE, "         ' �{�g���`��
        sql = sql & "MAGSTR, "           ' ���ꋭ�x
        sql = sql & "MAGPOS, "           ' ����ʒu
        sql = sql & "CONDGRT, "          ' �����ۏؓo�^
        sql = sql & "MODEL, "            ' �@��
        sql = sql & "UPMETHOD, "         ' ������@
        sql = sql & "UPCLASS, "          ' ����敪
        sql = sql & "UPNUM, "            ' ����{��
        sql = sql & "OPETIME, "          ' �^�]����
        sql = sql & "WTRCOOL, "          ' ����Ǘv��
        sql = sql & "PGID2, "            ' PG-ID�i��{���j
        sql = sql & "RCPT1, "            ' �Ή����V�sNo�iT1)
        sql = sql & "RCPT2, "            ' �Ή����V�sNo�iT2)
        sql = sql & "RCPT3, "            ' �Ή����V�sNo�iT3)
        sql = sql & "RCPT4, "            ' �Ή����V�sNo�iT4)
        sql = sql & "RCPT5, "            ' �Ή����V�sNo�iT5)
        sql = sql & "RCPT6, "            ' �Ή����V�sNo�iT6) 8/30 Yam
        sql = sql & "CNTL1, "            ' �������ځi1�j
        sql = sql & "CNTL2, "            ' �������ځi2�j
        sql = sql & "CNTL3, "            ' �������ځi3�j
        sql = sql & "CNTL4, "            ' �������ځi4�j
        sql = sql & "CNTL5, "            ' �������ځi5�j
        sql = sql & "CNTL6, "            ' �������ځi6�j
        sql = sql & "CNTL7, "            ' �������ځi7�j
        sql = sql & "CNTL8, "            ' �������ځi8�j
        sql = sql & "CNTL9, "            ' �������ځi9�j
        sql = sql & "CNTL10, "           ' �������ځi10�j
        sql = sql & "CNTL11, "           ' �������ځi11�j
        sql = sql & "CNTL12, "           ' �������ځi12�j
        sql = sql & "CNTL13, "           ' �������ځi13�j
        sql = sql & "CNTL14, "           ' �������ځi14�j
        sql = sql & "CNTL15, "           ' �������ځi15�j
        sql = sql & "RUNCOND1, "         ' �^�]�����P
        sql = sql & "RUNCOND2, "         ' �^�]�����Q
        sql = sql & "TSTAFFID, "         ' �o�^�Ј�ID
        sql = sql & "REGDATE, "          ' �o�^���t
        sql = sql & "KSTAFFID, "         ' �X�V�Ј�ID
        sql = sql & "UPDDATE, "          ' �X�V���t
        sql = sql & "SENDFLAG, "         ' ���M�t���O
        sql = sql & "SENDDATE )"         ' ���M���t
        With records(i)
            sql = sql & " values ("
            sql = sql & " '" & .PG_ID & "', "            ' PG-ID
            sql = sql & " ' ', "                      ' HZ�p�[�c
            sql = sql & " ' ', "                        ' HZ�p�^�[��
            sql = sql & " ' ', "                     ' �X�y�[�T
            sql = sql & " ' ', "                     ' �A�b�p�[�����O
            sql = sql & " " & .DR_CHRG * 100 & ", "      ' �`���[�W��
            sql = sql & " " & .DR_CPOS & ", "            ' ���c�{�ʒu
            sql = sql & " '" & .DR_CSIZ & "', "          ' ���c�{�T�C�Y
            sql = sql & " " & .DR_GAP & ", "             ' �M���b�v
            sql = sql & " " & .DR_DIA & ", "             ' ���㒼�a
            sql = sql & " " & .DR_LEN0 & ", "            ' ���㒷�i�S���j
            sql = sql & " " & .DR_LEN1 & ", "            ' ����iRC�j
            sql = sql & " ' ', "                         ' ���t���N�^�v��
            sql = sql & " '" & Trim(.DR_SR) & "', "      ' �㎲��]��
            sql = sql & " '" & Trim(.DR_CR) & "', "      ' ������]��
            sql = sql & " '" & Trim(.DR_PRES7) & "', "   ' �F����             '2003/05/16 osawa
            sql = sql & " '" & Trim(.DR_AR7) & "', "     ' �A���S����         '2003/05/16 osawa
            'sql = sql & " '" & .DR_PRES7 & "', "          ' �F����
            'sql = sql & " '" & .DR_AR7 & "', "            ' �A���S����
            sql = sql & " '" & .DR_DOP & "', "
            sql = sql & " '" & .DR_AR3 & "', "
            sql = sql & " 0, "                           ' �˂炢Oi�iMIN)
            sql = sql & " 0, "                           ' �˂炢Oi�iMAX)
            sql = sql & " ' ', "                         ' HC���
            sql = sql & " ' ', "                         ' HC
            sql = sql & " 0, "                           ' ���ψ��㑬�x
            sql = sql & " ' ', "                         ' ���㐧��
            sql = sql & " ' ', "                         ' �{�g���`��
            sql = sql & " 0, "                           ' ���ꋭ�x
            sql = sql & " 0, "                           ' ����ʒu
            sql = sql & " ' ', "                         ' �����ۏؓo�^
            sql = sql & " ' ', "                         ' �@��
            sql = sql & " ' ', "                         ' ������@
            sql = sql & " ' ', "                         ' ����敪
            sql = sql & " ' ', "                         ' ����{��
            sql = sql & " 0, "                           ' �^�]����
            sql = sql & " ' ', "                         ' ����Ǘv��
            sql = sql & " ' ', "                         ' PG-ID�i��{���j
            sql = sql & " ' ', "                         ' �Ή����V�sNo�iT1)
            sql = sql & " ' ', "                         ' �Ή����V�sNo�iT2)
            sql = sql & " ' ', "                         ' �Ή����V�sNo�iT3)
            sql = sql & " ' ', "                         ' �Ή����V�sNo�iT4)
            sql = sql & " ' ', "                         ' �Ή����V�sNo�iT5)
            sql = sql & " ' ', "                         ' �Ή����V�sNo�iT6)
            sql = sql & " ' ', "                         ' �������ځi1�j
            sql = sql & " ' ', "                         ' �������ځi2�j
            sql = sql & " ' ', "                         ' �������ځi3�j
            sql = sql & " ' ', "                         ' �������ځi4�j
            sql = sql & " ' ', "                         ' �������ځi5�j
            sql = sql & " ' ', "                         ' �������ځi6�j
            sql = sql & " ' ', "                         ' �������ځi7�j
            sql = sql & " ' ', "                         ' �������ځi8�j
            sql = sql & " ' ', "                         ' �������ځi9�j
            sql = sql & " ' ', "                         ' �������ځi10�j
            sql = sql & " ' ', "                         ' �������ځi11�j
            sql = sql & " ' ', "                         ' �������ځi12�j
            sql = sql & " ' ', "                         ' �������ځi13�j
            sql = sql & " ' ', "                         ' �������ځi14�j
            sql = sql & " ' ', "                         ' �������ځi15�j
            sql = sql & " ' ', "                         ' �^�]�����P
            sql = sql & " ' ', "                         ' �^�]�����Q
            sql = sql & " '" & staff & "', "             ' �o�^�Ј�ID
            sql = sql & " sysdate, "                     ' �o�^���t
            sql = sql & " '" & staff & "', "             ' �X�V�Ј�ID
            sql = sql & " sysdate, "                     ' �X�V���t
            sql = sql & " '0', "                         ' ���M�t���O
            sql = sql & " sysdate ) "                    ' ���M���t
        End With

        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_cmbc001c_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    
        DBDRV_cmbc001c_Ins = FUNCTION_RETURN_SUCCESS
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
    DBDRV_cmbc001c_Ins = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'7/30�@�⑫
'------------------------------------------------
' DB�A�N�Z�X�֐��i���o�ҁj
'------------------------------------------------
'�T�v      :�e�[�u���uTBCMB011�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB011 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/11�쐬�@����
Public Function DBDRV_cmbc001d_Disp(records() As typ_TBCMB011, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String                   'SQL�S��
Dim sqlBase As String                   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs      As OraDynaset               'RecordSet
Dim recCnt  As Long                     '���R�[�h��
Dim i       As Long                     'ٰ�߶���

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Disp"

    sqlBase = "Select PGID, HZPART, HZPTRN, SPACER, UPRING, CHARGE, RTBPOS, RTBSIZE, GAP, UPDM, UPLENGTH, UPRC, RFRNEED, UPSPIN," & _
              " DOWNSPIN, ROPRESS, ARUGON, AIMOIMIN, AIMOIMAX, HCCLASS, HC, AVEUPSPD, UPCNTL, BTMSHAPE, MAGSTR, MAGPOS, CONDGRT," & _
              " MODEL, UPMETHOD, UPCLASS, UPNUM, OPETIME, WTRCOOL, PGID2, RCPT1, RCPT2, RCPT3, RCPT4, RCPT5, RCPT6, CNTL1, CNTL2," & _
              " CNTL3, CNTL4, CNTL5, CNTL6, CNTL7, CNTL8, CNTL9, CNTL10, CNTL11, CNTL12, CNTL13, CNTL14, CNTL15, RUNCOND1," & _
              " RUNCOND2, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, DRDOP, DRAR3 "
    sqlBase = sqlBase & "From TBCMB011 "
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_cmbc001d_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
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
            .RCPT6 = rs("RCPT6")             ' �Ή����V�sNo�iT6)
            .CNTL1 = rs("CNTL1")             ' ���䍀�ځi1�j
            .CNTL2 = rs("CNTL2")             ' ���䍀�ځi2�j
            .CNTL3 = rs("CNTL3")             ' ���䍀�ځi3�j
            .CNTL4 = rs("CNTL4")             ' ���䍀�ځi4�j
            .CNTL5 = rs("CNTL5")             ' ���䍀�ځi5�j
            .CNTL6 = rs("CNTL6")             ' ���䍀�ځi6�j
            .CNTL7 = rs("CNTL7")             ' ���䍀�ځi7�j
            .CNTL8 = rs("CNTL8")             ' ���䍀�ځi8�j
            .CNTL9 = rs("CNTL9")             ' ���䍀�ځi9�j
            .CNTL10 = rs("CNTL10")           ' ���䍀�ځi10�j
            .CNTL11 = rs("CNTL11")           ' ���䍀�ځi11�j
            .CNTL12 = rs("CNTL12")           ' ���䍀�ځi12�j
            .CNTL13 = rs("CNTL13")           ' ���䍀�ځi13�j
            .CNTL14 = rs("CNTL14")           ' ���䍀�ځi14�j
            .CNTL15 = rs("CNTL15")           ' ���䍀�ځi15�j
            .RUNCOND1 = rs("RUNCOND1")       ' �^�]�����P
            .RUNCOND2 = rs("RUNCOND2")       ' �^�]�����Q
'            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
'            .REGDATE = rs("REGDATE")         ' �o�^���t
'            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
'            .UPDDATE = rs("UPDDATE")         ' �X�V���t
'            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
'            .SENDDATE = rs("SENDDATE")       ' ���M���t
            .DRDOP = IIf(rs("DRDOP") <> "", rs("DRDOP"), " ") ' �h�[�v      4/30
            .DRAR3 = IIf(rs("DRAR3") <> "", rs("DRAR3"), " ") ' �A���S�����R����
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_cmbc001d_Disp = FUNCTION_RETURN_SUCCESS

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

'7/30 �⑫
Public Function DBDRV_cmbc001d_Exec(records As typ_TBCMB011) As FUNCTION_RETURN
'------------------------------------------------
' DB�A�N�Z�X�֐��i�X�V�ҁj
'------------------------------------------------
'�T�v      :�e�[�u���uTBCMB011�v�̏����ɂ��������R�[�h�ɍX�V��������
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records�@     ,O  ,typ_TBCMB011 ,���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :���o�ް��̌����E����������"�ς�"�Ƃ���
'����      :2001/06/19(TUE)�쐬�@����
Dim sql     As String                   'SQL�S��
Dim rs      As OraDynaset               'RecordSet
Dim UpdID   As String                   '�X�V�Ώ�PGID


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Exec"

    UpdID = records.PGID

'2001/09/05 S.Sano Start �X�V�������Z�b�g����Ă��Ȃ��B
'2001/09/05 S.Sano Start ���̃��[�h��sysdate�̃Z�b�g���@���s���B
'    sql = "SELECT * FROM TBCMB011 WHERE(PGID = '" & UpdID & "')"
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'
'    rs.Edit
'    With records
'        rs("HZPART") = StrNoNull(.HZPART)          ' HZ�߰�
'        rs("HZPTRN") = StrNoNull(.HZPTRN)          ' HZ�����
'        rs("SPACER") = StrNoNull(.SPACER)          ' ��߰�
'        rs("UPRING") = StrNoNull(.UPRING)          ' ���߰�ݸ�
'        rs("CHARGE") = .CHARGE          ' ����ޗ�
'        rs("RTBPOS") = .RTBPOS          ' ���ވʒu
'        rs("RTBSIZE") = StrNoNull(.RTBSIZE)        ' ���޻���
'        rs("GAP") = .GAP                ' �ެ���
'        rs("UPDM") = .UPDM              ' ���㒼�a
'        rs("UPLENGTH") = .UPLENGTH      ' ���㒷
'        rs("UPRC") = .UPRC              ' ����RC
'        rs("RFRNEED") = StrNoNull(.RFRNEED)        ' ��׸��v��
'        rs("UPSPIN") = StrNoNull(.UPSPIN)          ' �㎲��]��
'        rs("DOWNSPIN") = StrNoNull(.DOWNSPIN)      ' ������]��
'        rs("ROPRESS") = StrNoNull(.ROPRESS)        ' �F����
'        rs("ARUGON") = StrNoNull(.ARUGON)          ' �ٺ�ݗ�
'        rs("AIMOIMIN") = .AIMOIMIN      ' �˂炢iO(MIN)
'        rs("AIMOIMAX") = .AIMOIMAX      ' �˂炢iO(MAX)
'        rs("HCCLASS") = StrNoNull(.HCCLASS)        ' HC���
'        rs("HC") = StrNoNull(.HC)                  ' HC
'        rs("AVEUPSPD") = .AVEUPSPD      ' ���ψ��㑬�x
'        rs("UPCNTL") = StrNoNull(.UPCNTL)          ' ���㐧��
'        rs("BTMSHAPE") = StrNoNull(.BTMSHAPE)      ' ���ь`��
'        rs("MAGSTR") = .MAGSTR          ' ���ꋭ�x
'        rs("MAGPOS") = .MAGPOS          ' ����ʒu
'        rs("CONDGRT") = StrNoNull(.CONDGRT)        ' �����ۏؓo�^
'        rs("MODEL") = StrNoNull(.MODEL)            ' �@��
'        rs("UPMETHOD") = StrNoNull(.UPMETHOD)      ' ������@
'        rs("UPCLASS") = StrNoNull(.UPCLASS)        ' ����敪
'        rs("UPNUM") = StrNoNull(.UPNUM)            ' ����{��
'        rs("OPETIME") = .OPETIME        ' �^�]����
'        rs("WTRCOOL") = StrNoNull(.WTRCOOL)        ' ����Ǘv��
'        rs("PGID2") = StrNoNull(.PGID2)            ' PG-ID�i��{���j
'        rs("RCPT1") = StrNoNull(.RCPT1)            ' �Ή�ڼ��No�iT1�j
'        rs("RCPT2") = StrNoNull(.RCPT2)            ' �Ή�ڼ��No�iT2�j
'        rs("RCPT3") = StrNoNull(.RCPT3)            ' �Ή�ڼ��No�iT3�j
'        rs("RCPT4") = StrNoNull(.RCPT4)            ' �Ή�ڼ��No�iT4�j
'        rs("RCPT5") = StrNoNull(.RCPT5)            ' �Ή�ڼ��No�iT5�j
'        rs("CNTL1") = StrNoNull(.CNTL1)            ' ���䍀��(1)
'        rs("CNTL2") = StrNoNull(.CNTL2)            ' ���䍀��(2)
'        rs("CNTL3") = StrNoNull(.CNTL3)            ' ���䍀��(3)
'        rs("CNTL4") = StrNoNull(.CNTL4)            ' ���䍀��(4)
'        rs("CNTL5") = StrNoNull(.CNTL5)            ' ���䍀��(5)
'        rs("CNTL6") = StrNoNull(.CNTL6)            ' ���䍀��(6)
'        rs("CNTL7") = StrNoNull(.CNTL7)            ' ���䍀��(7)
'        rs("CNTL8") = StrNoNull(.CNTL8)            ' ���䍀��(8)
'        rs("CNTL9") = StrNoNull(.CNTL9)            ' ���䍀��(9)
'        rs("CNTL10") = StrNoNull(.CNTL10)          ' ���䍀��(10)
'        rs("CNTL11") = StrNoNull(.CNTL11)          ' ���䍀��(11)
'        rs("CNTL12") = StrNoNull(.CNTL12)          ' ���䍀��(12)
'        rs("CNTL13") = StrNoNull(.CNTL13)          ' ���䍀��(13)
'        rs("CNTL14") = StrNoNull(.CNTL14)          ' ���䍀��(14)
'        rs("CNTL15") = StrNoNull(.CNTL15)          ' ���䍀��(15)
'        rs("RUNCOND1") = StrNoNull(.RUNCOND1)      ' �^�]����1
'        rs("RUNCOND2") = StrNoNull(.RUNCOND2)      ' �^�]����2
'    End With
'    rs.Update
'
'    rs.Close
    
'2001/09/05 S.Sano Start
    With records
    sql = "update TBCMB011 set "
    sql = sql & "HZPART = '" & StrNoNull(.HZPART) & "', "       ' HZ�߰�
    sql = sql & "HZPTRN = '" & StrNoNull(.HZPTRN) & "', "       ' HZ�����
    sql = sql & "SPACER = '" & StrNoNull(.SPACER) & "', "       ' ��߰�
    sql = sql & "UPRING = '" & StrNoNull(.UPRING) & "', "       ' ���߰�ݸ�
    sql = sql & "CHARGE = " & .CHARGE & ", "                    ' ����ޗ�
    sql = sql & "RTBPOS = " & .RTBPOS & ", "                    ' ���ވʒu
    sql = sql & "RTBSIZE = '" & StrNoNull(.RTBSIZE) & "', "     ' ���޻���
    sql = sql & "GAP = " & .GAP & ", "                          ' �ެ���
    sql = sql & "UPDM = " & .UPDM & ", "                        ' ���㒼�a
    sql = sql & "UPLENGTH = " & .UPLENGTH & ", "                ' ���㒷
    sql = sql & "UPRC = " & .UPRC & ", "                        ' ����RC
    sql = sql & "RFRNEED = '" & StrNoNull(.RFRNEED) & "', "     ' ��׸��v��
    sql = sql & "UPSPIN = '" & StrNoNull(.UPSPIN) & "', "       ' �㎲��]��
    sql = sql & "DOWNSPIN = '" & StrNoNull(.DOWNSPIN) & "', "   ' ������]��
    sql = sql & "ROPRESS = '" & StrNoNull(.ROPRESS) & "', "     ' �F����
    sql = sql & "ARUGON = '" & StrNoNull(.ARUGON) & "', "       ' �ٺ�ݗ�
    sql = sql & "DRAR3 ='" & StrNoNull(.DRAR3) & "', "          ' �A���S�����R���� ' 4/30 YAM
    sql = sql & "DRDOP ='" & StrNoNull(.DRDOP) & "', "          ' �h�[�v
    sql = sql & "AIMOIMAX = " & .AIMOIMAX & ", "                ' �˂炢iO(MAX)
    sql = sql & "HCCLASS = '" & StrNoNull(.HCCLASS) & "', "     ' HC���
    sql = sql & "HC = '" & StrNoNull(.HC) & "', "               ' HC
    sql = sql & "AVEUPSPD = " & .AVEUPSPD & ", "                ' ���ψ��㑬�x
    sql = sql & "UPCNTL = '" & StrNoNull(.UPCNTL) & "', "       ' ���㐧��
    sql = sql & "BTMSHAPE = '" & StrNoNull(.BTMSHAPE) & "', "   ' ���ь`��
    sql = sql & "MAGSTR = " & .MAGSTR & ", "                    ' ���ꋭ�x
    sql = sql & "MAGPOS = " & .MAGPOS & ", "                    ' ����ʒu
    sql = sql & "CONDGRT = '" & StrNoNull(.CONDGRT) & "', "     ' �����ۏؓo�^
    sql = sql & "MODEL = '" & StrNoNull(.MODEL) & "', "         ' �@��
    sql = sql & "UPMETHOD = '" & StrNoNull(.UPMETHOD) & "', "   ' ������@
    sql = sql & "UPCLASS = '" & StrNoNull(.UPCLASS) & "', "     ' ����敪
    sql = sql & "UPNUM = '" & StrNoNull(.UPNUM) & "', "         ' ����{��
    sql = sql & "OPETIME = " & .OPETIME & ", "                  ' �^�]����
    sql = sql & "WTRCOOL = '" & StrNoNull(.WTRCOOL) & "', "     ' ����Ǘv��
    sql = sql & "PGID2 = '" & StrNoNull(.PGID2) & "', "         ' PG-ID�i��{���j
    sql = sql & "RCPT1 = '" & StrNoNull(.RCPT1) & "', "         ' �Ή�ڼ��No�iT1�j
    sql = sql & "RCPT2 = '" & StrNoNull(.RCPT2) & "', "         ' �Ή�ڼ��No�iT2�j
    sql = sql & "RCPT3 = '" & StrNoNull(.RCPT3) & "', "         ' �Ή�ڼ��No�iT3�j
    sql = sql & "RCPT4 = '" & StrNoNull(.RCPT4) & "', "         ' �Ή�ڼ��No�iT4�j
    sql = sql & "RCPT5 = '" & StrNoNull(.RCPT5) & "', "         ' �Ή�ڼ��No�iT5�j
    sql = sql & "RCPT6 = '" & StrNoNull(.RCPT6) & "', "         ' �Ή�ڼ��No�iT6�j
    sql = sql & "CNTL1 = '" & StrNoNull(.CNTL1) & "', "         ' ���䍀��(1)
    sql = sql & "CNTL2 = '" & StrNoNull(.CNTL2) & "', "         ' ���䍀��(2)
    sql = sql & "CNTL3 = '" & StrNoNull(.CNTL3) & "', "         ' ���䍀��(3)
    sql = sql & "CNTL4 = '" & StrNoNull(.CNTL4) & "', "         ' ���䍀��(4)
    sql = sql & "CNTL5 = '" & StrNoNull(.CNTL5) & "', "         ' ���䍀��(5)
    sql = sql & "CNTL6 = '" & StrNoNull(.CNTL6) & "', "         ' ���䍀��(6)
    sql = sql & "CNTL7 = '" & StrNoNull(.CNTL7) & "', "         ' ���䍀��(7)
    sql = sql & "CNTL8 = '" & StrNoNull(.CNTL8) & "', "         ' ���䍀��(8)
    sql = sql & "CNTL9 = '" & StrNoNull(.CNTL9) & "', "         ' ���䍀��(9)
    sql = sql & "CNTL10 = '" & StrNoNull(.CNTL10) & "', "       ' ���䍀��(10)
    sql = sql & "CNTL11 = '" & StrNoNull(.CNTL11) & "', "       ' ���䍀��(11)
    sql = sql & "CNTL12 = '" & StrNoNull(.CNTL12) & "', "       ' ���䍀��(12)
    sql = sql & "CNTL13 = '" & StrNoNull(.CNTL13) & "', "       ' ���䍀��(13)
    sql = sql & "CNTL14 = '" & StrNoNull(.CNTL14) & "', "       ' ���䍀��(14)
    sql = sql & "CNTL15 = '" & StrNoNull(.CNTL15) & "', "       ' ���䍀��(15)
    sql = sql & "RUNCOND1 = '" & StrNoNull(.RUNCOND1) & "', "   ' �^�]����1
    sql = sql & "RUNCOND2 = '" & StrNoNull(.RUNCOND2) & "', "   ' �^�]����2
    sql = sql & "KSTAFFID = '" & .KSTAFFID & "', "              ' �X�V�Ј�ID
    sql = sql & "UPDDATE = sysdate "                            ' �X�V���t
    sql = sql & "where PGID = '" & UpdID & "'"
    End With
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_cmbc001d_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_cmbc001d_Exec = FUNCTION_RETURN_SUCCESS
'2001/09/05 S.Sano End

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
'8/6 �⑫
'------------------------------------------------
' DB�A�N�Z�X�֐��i�폜�ҁj
'------------------------------------------------
'�T�v      :�e�[�u���uTBCMB011�v�̏����ɂ��������R�[�h���폜
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :PGID�@        ,O  ,String       ,�폜PG-ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/10/05 �쐬�@���{
Public Function DBDRV_cmbc001d_Del(PGID As String) As FUNCTION_RETURN

    Dim sql     As String                   'SQL�S��

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Del"
    
    sql = "delete "
    sql = sql & "from "
    sql = sql & "TBCMB011 "
    sql = sql & "where "
    sql = sql & "trim(PGID)='" & Trim(PGID) & "'"
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_cmbc001d_Del = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_cmbc001d_Del = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_cmbc001d_Del = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'8/6 �⑫
Private Function StrNoNull(s$) As String
    If Trim$(s) = vbNullString Then
        StrNoNull = " "
    Else
        StrNoNull = Trim$(s)
    End If
End Function
