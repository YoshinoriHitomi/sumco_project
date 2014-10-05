Attribute VB_Name = "s_cmbc001b_SQL"
Option Explicit

' TBCME017 (���i�d�l�Ǘ�)���
Public Type s_cmzcF_cmfc001b_Disp
    '���i�d�l�Ǘ�
    Hinban12 As String * 12         ' �i��
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    REGDATE As Date                 ' �o�^���t
End Type


Public Function DBDRV_s_cmzcF_cmfc001b_Disp(records() As s_cmzcF_cmfc001b_Disp) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001b_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001b_Disp"
    
    ''���i�d�l�Ǘ���������SXL����������Ȃ����R�[�h�擾�i�i�ԁA�d�l�o�^�˗��ԍ��A�o�^���t�j
    ''�������A��������t�^����ɂ��郌�R�[�h�͏���
    sql = "select hinban||ltrim(to_char(mnorevno,'00'))||factory||opecond as hinban12, HMGSTRRNO, REGDATE " & _
          "From tbcme018 " & _
          "where (opecond='1') and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme030) and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme031)"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_s_cmzcF_cmfc001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Hinban12 = rs("HINBAN12")        ' �i��
            .HMGSTRRNO = rs("HMGSTRRNO")    ' �i�Ǘ��d�l�o�^�˗��ԍ�
            .REGDATE = rs("REGDATE")        ' �o�^���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_s_cmzcF_cmfc001b_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'���̉�ʂł�Exec()�͂���Ȃ�
'Public Function DBDRV_s_cmzcF_cmgc001d_Exec(s_cmzcF_cmfc001a_Disp As type_DBDRV_s_cmzcF_cmgc001d_Exec) As FUNCTION_RETURN
'    s_cmzcF_cmgc001c_Exec = FUNCTION_RETURN_SUCCESS
'
'    '�������g��򕥏o���уe�[�u���Ɍ����ԍ�()�A�Ǘ��H���R�[�h()�A�H���R�[�h()�A������d��()�A���X�d�ʁA�Ј��h�c���C���T�[�g
'
'End Function

