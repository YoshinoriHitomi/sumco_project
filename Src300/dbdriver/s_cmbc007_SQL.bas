Attribute VB_Name = "s_cmbc007_SQL"
Option Explicit

'�����݌ɏC��


'�����݌Ɏ擾�p
Public Type type_DBDRV_scmzc_fcmgc001e_Disp
    '�����݌ɊǗ�
    MTRLNUM As String * 10      ' �����ԍ�
    WEIGHT As Long              ' �d��
End Type


'�����݌ɍX�V�p
Public Type type_DBDRV_scmzc_fcmgc001e_Exec
    '�����݌ɊǗ�
    MTRLNUM As String * 10      ' �����ԍ�
    USABLCLS As String * 1      ' �g�p�\�敪
    KRPROCCD As String * 5      ' �Ǘ��H���R�[�h
    PROCCODE As String * 5      ' �H���R�[�h
    KSTAFFID As String * 8      ' �X�V�Ј�ID
    WEIGHT As Long              ' �V�d��
    SYORIW As Long              ' ������

End Type



'�����\��
'�T�v    :�����݌ɏC�� �\���p�c�a�h���C�o
'���Ұ�  :�ϐ���       ,IO  ,�^                                    ,����
'        :record       ,I   ,type_DBDRV_scmzc_fcmgc001e_Disp       ,�����݌Ɏ擾�p
'        :��ؒl        ,O   ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'����    :
'����    :2001/06/18 ���{ �쐬
Public Function DBDRV_scmzc_fcmgc001e_Disp(records() As type_DBDRV_scmzc_fcmgc001e_Disp) As FUNCTION_RETURN
    
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Integer
    Dim i As Long
    
    '�����Ǘ��e�[�u���Ŏg�p�\�敪�P��select�i�����ԍ��A�d�ʁj
    

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmgc001e_SQL.bas -- Function DBDRV_scmzc_fcmgc001e_Disp"

    sql = "Select MTRLNUM, USABLCLS, WEIGHT, TSTAFFID, REGDATE, KSTAFFID, UPDDATE "
    sql = sql & "From TBCMG005"
    
        
        sql = "select MTRLNUM, WEIGHT"
        sql = sql & " from ( "
        sql = sql & " select MTRLNUM, WEIGHT"
        sql = sql & " from TBCMG005"
        sql = sql & " where USABLCLS='1'"
        sql = sql & " and WEIGHT > 0 "
        sql = sql & " and substr(MTRLNUM,1,1) not in ('P','N')"
        sql = sql & " order by MTRLNUM ) "
        sql = sql & " union all "
        sql = sql & " select MTRLNUM, WEIGHT"
        sql = sql & " from ( "
        sql = sql & " select MTRLNUM, WEIGHT"
        sql = sql & " from TBCMG005"
        sql = sql & " where USABLCLS='1'"
        sql = sql & " and WEIGHT > 0 "
        sql = sql & " and substr(MTRLNUM,1,1) in ('P','N')"
        sql = sql & " order by MTRLNUM )"


    '   order by �����ԍ�
    '   substr(�����ԍ�,1,1) not in ('P','N')
    '   union all
    
    'select ...
    '   order by �����ԍ�
    '   substr(�����ԍ�,1,1) in ('P','N')
    
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    recCnt = rs.RecordCount
    ReDim records(recCnt)
    If recCnt = 0 Then ''2001/07/17 Sano
'2001/07/17 Sano    If rs.RecordCount = 0 Then
        DBDRV_scmzc_fcmgc001e_Disp = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If
    
'2001/07/17 Sano    recCnt = rs.RecordCount
'2001/07/17 Sano    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .MTRLNUM = rs("MTRLNUM")          ' �����ԍ�
            .WEIGHT = rs("WEIGHT")            ' �d��
        End With
        rs.MoveNext
    Next i
    rs.Close

    DBDRV_scmzc_fcmgc001e_Disp = FUNCTION_RETURN_SUCCESS
   

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmgc001e_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'���s��
'�T�v    :�����݌ɏC�� �X�V�A�}���p�c�a�h���C�o
'���Ұ�  :�ϐ���       ,IO  ,�^                                    ,����
'        :record       ,I   ,type_DBDRV_scmzc_fcmgc001e_Exec       ,�����݌ɑ}���p
'        :��ؒl        ,O   ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'����    :
'����    :2001/06/18 ���{ �쐬
Public Function DBDRV_scmzc_fcmgc001e_Exec(record As type_DBDRV_scmzc_fcmgc001e_Exec) As FUNCTION_RETURN
    
    Dim sql As String

    
    '�����Ǘ��e�[�u����V�d�ʂɍX�V
        

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmgc001e_SQL.bas -- Function DBDRV_scmzc_fcmgc001e_Exec"

    DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_SUCCESS
    
    sql = "update TBCMG005 set "
    With record
        sql = sql & "WEIGHT=" & .WEIGHT & ", "               ' �d��
        sql = sql & "KSTAFFID='" & .KSTAFFID & "', "         ' �X�V�Ј�ID
        sql = sql & "UPDDATE=sysdate "                       ' �X�V���t
        sql = sql & "where MTRLNUM='" & .MTRLNUM & "' "
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    
    '�����݌Ɏ��тɑ}��
    sql = " insert into TBCMG006 ( "
    sql = sql & "MTRLNUM, "          ' �����ԍ�
    sql = sql & "TRANCNT, "          ' ������
    sql = sql & "KRPROCCD, "         ' �Ǘ��H���R�[�h
    sql = sql & "PROCCODE, "         ' �H���R�[�h
    sql = sql & "CLASS, "            ' �敪
    sql = sql & "INWEIGHT, "         ' ���͏d��
    sql = sql & "TSTAFFID, "         ' �o�^�Ј�ID
    sql = sql & "REGDATE, "          ' �o�^���t
    sql = sql & "KSTAFFID, "         ' �X�V�Ј�ID
    sql = sql & "UPDDATE, "          ' �X�V���t
    sql = sql & "SENDFLAG, "         ' ���M�t���O
    sql = sql & "SENDDATE ) "        ' ���M���t
    With record
        sql = sql & " select "
        sql = sql & " '" & .MTRLNUM & "', "          ' �����ԍ�
        sql = sql & " nvl(max(TRANCNT),0)+1, "       ' ������
        sql = sql & " '" & .KRPROCCD & "', "         ' �Ǘ��H���R�[�h
        sql = sql & " '" & .PROCCODE & "', "         ' �H���R�[�h
        sql = sql & " '" & .USABLCLS & "', "         ' �敪
        sql = sql & " '" & .SYORIW & "', "           ' ���͏d��
        sql = sql & " '" & .KSTAFFID & "', "         ' �o�^�Ј�ID
        sql = sql & " sysdate, "                     ' �o�^���t
        sql = sql & " '" & .KSTAFFID & "', "         ' �X�V�Ј�ID
        sql = sql & " sysdate, "                     ' �X�V���t
        sql = sql & " '0', "                         ' ���M�t���O
        sql = sql & " sysdate "                      ' ���M���t
        sql = sql & " from TBCMG006 "
        sql = sql & " where MTRLNUM='" & .MTRLNUM & "' "
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_FAILURE
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
    DBDRV_scmzc_fcmgc001e_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

