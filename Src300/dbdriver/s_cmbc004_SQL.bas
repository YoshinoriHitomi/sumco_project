Attribute VB_Name = "s_cmbc004_SQL"
Option Explicit
'
'' ����������I������



Public Type type_DBDRV_scmzc_fcmgc001b_Exec
    ' ������������ё}���p
    KRPROCCD As String * 5      ' �Ǘ��H���R�[�h
    PROCCODE As String * 5      ' �H���R�[�h
    TSTAFFID As String * 8      ' �o�^�Ј�ID
    MTRLTYPE As String * 3      ' �������
    MAKERNO As String * 6       ' ���[�J�Ǘ�No
    RVWEIGHT As Double          ' ����w���d��
    CRYCOMMENT As String        ' �R�����g
    WEIGHT    As Double         ' �{���̎����
    
End Type

Public Type type_DBDRV_scmzc_fcmgc001b_Weight
    ' ����������d�|��d�ʒ��o�p
    MTRL As String '* 10     ' �������
    WEIGHT As Double            ' ����w���d��
End Type


'�T�v    :����������I������ �X�V�A�}���p�c�a�h���C�o
'���Ұ�  :�ϐ���       ,IO  ,�^                                    ,����
'        :record       ,I   ,type_DBDRV_scmzc_fcmgc001b_Exec       ,������������ё}���p
'        :��ؒl        ,O   ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'����    :
'����    :2001/06/18 ���{ �쐬
Public Function DBDRV_scmzc_fcmgc001b_Exec(record As type_DBDRV_scmzc_fcmgc001b_Exec) As FUNCTION_RETURN

    Dim sql As String
    Dim MTRLNUM As String
    Dim rs As OraDynaset


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmgc001b_SQL.bas -- Function DBDRV_scmzc_fcmgc001b_Exec"

    DBDRV_scmzc_fcmgc001b_Exec = FUNCTION_RETURN_SUCCESS


    '�����ԍ�
    MTRLNUM = record.MTRLTYPE & record.MAKERNO & "0"

    '������������уe�[�u���֑}��
    sql = " insert into TBCMG001 ( "
    sql = sql & "MTRLNUM, "          ' �����ԍ�
    sql = sql & "JDATE, "            ' ���t
'    sql = sql & "TRANCNT, "          ' ������
    sql = sql & "KRPROCCD, "         ' �Ǘ��H���R�[�h
    sql = sql & "PROCCODE, "         ' �H���R�[�h
    sql = sql & "MTRLTYPE, "         ' �������
    sql = sql & "MAKERNO, "          ' ���[�J�Ǘ�No
    sql = sql & "RVWEIGHT, "         ' ����w���d��
    sql = sql & "CRYCOMMENT, "       ' �R�����g
    sql = sql & "TSTAFFID, "         ' �o�^�Ј�ID
    sql = sql & "REGDATE, "          ' �o�^���t
    sql = sql & "KSTAFFID, "         ' �X�V�Ј��h�c
    sql = sql & "UPDDATE, "          ' �X�V���t
    sql = sql & "SENDFLAG, "         ' ���M�t���O
    sql = sql & "SENDDATE) "         ' ���M���t
    With record
        sql = sql & " values ( "
        sql = sql & " '" & MTRLNUM & "', "           ' �����ԍ�
        sql = sql & " sysdate, "                     ' ���t       sysdate�ɕύX�\��#kk#
'        sql = sql & " nvl(max(TRANCNT),0)+1, "       ' ������      �͂Ȃ��Ȃ�#kk#
        sql = sql & " '" & .KRPROCCD & "', "         ' �Ǘ��H���R�[�h
        sql = sql & " '" & .PROCCODE & "', "         ' �H���R�[�h
        sql = sql & " '" & .MTRLTYPE & "', "         ' �������
        sql = sql & " '" & .MAKERNO & "', "          ' ���[�J�Ǘ�No
        sql = sql & " " & .WEIGHT & ", "             ' ����w���d��
        sql = sql & " '" & .CRYCOMMENT & "', "       ' �R�����g
        sql = sql & " '" & .TSTAFFID & "', "         ' �o�^�Ј�ID
        sql = sql & " sysdate, "                      ' �o�^���t
        sql = sql & " '" & .TSTAFFID & "', "         ' �X�V�Ј��h�c
        sql = sql & " sysdate, "                      ' �X�V���t
        sql = sql & " '0', "                         ' ���M�t���O
        sql = sql & " sysdate ) "                      ' ���M���t
'        sql = sql & " from TBCMG001 "
'        sql = sql & " where MTRLNUM='" & MTRLNUM & "' "
    End With
    

    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmgc001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    '�����݌ɊǗ��̍X�V or �}��
    sql = " select "
    sql = sql & "count(MTRLNUM) as C "
    sql = sql & "from TBCMG005 "
    sql = sql & "where MTRLNUM='" & MTRLNUM & "' "
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '���R�[�h������������}��
    If rs("C") = 0 Then
        '�����݌ɊǗ��e�[�u���ւ̑}��
        sql = "insert into TBCMG005 ( "
        sql = sql & "MTRLNUM, "          ' �����ԍ�
        sql = sql & "USABLCLS, "         ' �g�p�\�敪
        sql = sql & "WEIGHT, "           ' �d��
        sql = sql & "TSTAFFID, "         ' �o�^�Ј�ID
        sql = sql & "REGDATE, "          ' �o�^���t
        sql = sql & "KSTAFFID, "         ' �X�V�Ј�ID
        sql = sql & "UPDDATE ) "           ' �X�V���t
        
        sql = sql & " values ( "
        sql = sql & " '" & MTRLNUM & "', "
        sql = sql & " '1', "
        sql = sql & " " & record.RVWEIGHT & ", "
        sql = sql & " '" & record.TSTAFFID & "', "   ' �o�^�Ј�ID
        sql = sql & " sysdate, "                      ' �o�^���t
        sql = sql & " '" & record.TSTAFFID & "', "   ' �X�V�Ј��h�c
        sql = sql & " sysdate )"                      ' �X�V���t
    
    Else
    
        '�����݌ɊǗ��e�[�u���̍X�V
        sql = "update TBCMG005 set "
        sql = sql & "WEIGHT=" & record.RVWEIGHT & ", "           ' �d��
        sql = sql & "KSTAFFID='" & record.TSTAFFID & "', "         ' �X�V�Ј�ID
        sql = sql & "UPDDATE=sysdate "                                     ' �X�V���t
        sql = sql & "where MTRLNUM='" & MTRLNUM & "' "
    
    End If
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
         DBDRV_scmzc_fcmgc001b_Exec = FUNCTION_RETURN_FAILURE
         GoTo proc_exit
    End If
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmgc001b_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v    :����������I������ �d�|��d�ʒ��o�c�a�h���C�o
'���Ұ�  :�ϐ���       ,IO  ,�^                                    ,����
'        :record       ,I   ,type_DBDRV_scmzc_fcmgc001b_Weight     ,����������d�|��d�ʒ��o�p
'        :��ؒl        ,O   ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'����    :
'����    :2001/07/17 Sano �쐬
Public Function DBDRV_scmzc_fcmgc001b_Weight(record As type_DBDRV_scmzc_fcmgc001b_Weight) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmgc001b_SQL.bas -- Function DBDRV_scmzc_fcmgc001b_Weight"

    DBDRV_scmzc_fcmgc001b_Weight = FUNCTION_RETURN_SUCCESS

    '�����݌ɊǗ��̍X�V or �}��
    sql = " select "
    sql = sql & "nvl(sum(nvl(WEIGHT,0)),0) as W "
    sql = sql & "from TBCMG005 "
    sql = sql & "where MTRLNUM like'" & record.MTRL & "%' "
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    record.WEIGHT = rs("W")
    rs.Close

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmgc001b_Weight = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

