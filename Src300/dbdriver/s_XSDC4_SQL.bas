Attribute VB_Name = "s_XSDC4_SQL"
'�s�Ǔ��� (XSDC4) �����֐�


'***�e�[�u���uXSDC4�v�ւ̃f�[�^�A�N�Z�X�֐�***
'������ ���Ұ��ɒl��Ă��鎞�A�܂��S�ď��������邱��

Option Explicit

'���s�Ǔ���
Public Type typ_XSDC4
    XTALC4 As String * 12      ' ��ۯ�ID������ԍ�
    INPOSC4 As Integer         ' �������J�n�ʒu
    KCKNTC4 As Integer         ' �H���A��
    HINBC4 As String * 8       ' �i��
    REVNUMC4 As Integer        ' ���i�ԍ������ԍ�
    FACTORYC4 As String * 1    ' �H��
    OPEC4 As String * 1        ' ���Ə���
    KNKTC4 As String * 5       ' �Ǘ��H��
    WKKTC4 As String * 5       ' �H��
    WKKDC4 As String * 2       ' ��Ƌ敪
    MACOC4 As Integer          ' ������
    SXLIDC4 As String * 13     ' SXLID
    FCODEC4 As String * 3      ' ����R�[�h
    PUCUTLC4 As Integer        ' �s�ǒ���
    PUCUTWC4 As Long           ' �s�Ǐd��
    PUCUTMC4 As Integer        ' �s�ǖ���
    FKUBC4 As String * 1       ' �s�ǋ敪
    TDAYC4 As Date             ' �o�^���t
    KDAYC4 As Date             ' �X�V���t
    SUMITBC3 As String * 1     ' SUMMIT���M�t���O
    SNDKC3 As String * 1       ' ���M�t���O
    SNDDAYC3 As Date           ' ���M���t
End Type

'�X�V�p
Public Type typ_XSDC4_Update
    XTALC4 As String           ' ��ۯ�ID������ԍ�
    INPOSC4 As String          ' �������J�n�ʒu
    KCKNTC4 As String          ' �H���A��
    HINBC4 As String           ' �i��
    REVNUMC4 As String         ' ���i�ԍ������ԍ�
    FACTORYC4 As String        ' �H��
    OPEC4 As String            ' ���Ə���
    KNKTC4 As String           ' �Ǘ��H��
    WKKTC4 As String           ' �H��
    WKKDC4 As String           ' ��Ƌ敪
    MACOC4 As String           ' ������
    SXLIDC4 As String          ' SXLID
    FCODEC4 As String          ' ����R�[�h
    PUCUTLC4 As String         ' �s�ǒ���
    PUCUTWC4 As String         ' �s�Ǐd��
    PUCUTMC4 As String         ' �s�ǖ���
    FKUBC4 As String           ' �s�ǋ敪
    TDAYC4 As String           ' �o�^���t
    KDAYC4 As String           ' �X�V���t
    SUMITBC3 As String         ' SUMMIT���M�t���O
    SNDKC3 As String           ' ���M�t���O
    SNDDAYC3 As String         ' ���M���t
End Type

'��SELECT��

'�T�v      :�e�[�u���uXSDC4�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO   ,�^               ,����
'          :records()     ,O    ,typ_XSDC4     ,���o���R�[�h
'          :sqlWhere      ,I    ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I    ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :

Public Function DBDRV_GetXSDC4(records() As typ_XSDC4, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim sql As String       'SQL�S��
    Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select * From XSDC4"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDC4 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    If recCnt = 0 Then
        Exit Function
    End If
    For i = 1 To recCnt
        With records(i)
            If IsNull(rs.Fields("XTALC4")) = False Then .XTALC4 = rs.Fields("XTALC4")           ' ��ۯ�ID������ԍ�
            If IsNull(rs.Fields("INPOSC4")) = False Then .INPOSC4 = rs.Fields("INPOSC4")         ' �������J�n�ʒu
            If IsNull(rs.Fields("KCKNTC4")) = False Then .KCKNTC4 = rs.Fields("KCKNTC4")         ' �H���A��
            If IsNull(rs.Fields("HINBC4")) = False Then .HINBC4 = rs.Fields("HINBC4")           ' �i��
            If IsNull(rs.Fields("REVNUMC4")) = False Then .REVNUMC4 = rs.Fields("REVNUMC4")       ' ���i�ԍ������ԍ�
            If IsNull(rs.Fields("FACTORYC4")) = False Then .FACTORYC4 = rs.Fields("FACTORYC4")     ' �H��
            If IsNull(rs.Fields("OPEC4")) = False Then .OPEC4 = rs.Fields("OPEC4")             ' ���Ə���
            If IsNull(rs.Fields("KNKTC4")) = False Then .KNKTC4 = rs.Fields("KNKTC4")           ' �Ǘ��H��
            If IsNull(rs.Fields("WKKTC4")) = False Then .WKKTC4 = rs.Fields("WKKTC4")           ' �H��
            If IsNull(rs.Fields("WKKDC4")) = False Then .WKKDC4 = rs.Fields("WKKDC4")           ' ��Ƌ敪
            If IsNull(rs.Fields("MACOC4")) = False Then .MACOC4 = rs.Fields("MACOC4")           ' ������
            If IsNull(rs.Fields("SXLIDC4")) = False Then .SXLIDC4 = rs.Fields("SXLIDC4")         ' SXLID
            If IsNull(rs.Fields("FCODEC4")) = False Then .FCODEC4 = rs.Fields("FCODEC4")         ' ����R�[�h
            If IsNull(rs.Fields("PUCUTLC4")) = False Then .PUCUTLC4 = rs.Fields("PUCUTLC4")       ' �s�ǒ���
            If IsNull(rs.Fields("PUCUTWC4")) = False Then .PUCUTWC4 = rs.Fields("PUCUTWC4")       ' �s�Ǐd��
            If IsNull(rs.Fields("PUCUTMC4")) = False Then .PUCUTMC4 = rs.Fields("PUCUTMC4")       ' �s�ǖ���
            If IsNull(rs.Fields("FKUBC4")) = False Then .FKUBC4 = rs.Fields("FKUBC4")           ' �s�ǋ敪
            If IsNull(rs.Fields("TDAYC4")) = False Then .TDAYC4 = rs.Fields("TDAYC4")           ' �o�^���t
            If IsNull(rs.Fields("KDAYC4")) = False Then .KDAYC4 = rs.Fields("KDAYC4")           ' �X�V���t
            If IsNull(rs.Fields("SUMITBC3")) = False Then .SUMITBC3 = rs.Fields("SUMITBC3")       ' SUMMIT���M�t���O
            If IsNull(rs.Fields("SNDKC3")) = False Then .SNDKC3 = rs.Fields("SNDKC3")           ' ���M�t���O
            If IsNull(rs.Fields("SNDDAYC3")) = False Then .SNDDAYC3 = rs.Fields("SNDDAYC3")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDC4 = FUNCTION_RETURN_SUCCESS
End Function


'��INSERT��

'�T�v      :�e�[�u���uXSDC4�v�Ƀ��R�[�h��}������
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pXSDC4 �@�@  ,I  ,typ_XSDC4_Update   ,XSDC4�X�V�p�ް�
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function CreateXSDC4(pXSDC4 As typ_XSDC4_Update, sErrMsg As String) As FUNCTION_RETURN


    Dim sql As String
    Dim sDbName As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim rs2 As OraDynaset    'RecordSet
'    Dim recCnt As Long      '���R�[�h��
    Dim nowtime As Date
    Dim nowtime_sql     As String   '�T�[�o����(SQL��)
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDC4_SQL.bas -- Function CreateXSDC4"
    sErrMsg = ""
    sDbName = "XSDC4"
    'nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾����悤�ɕύX 2003/6/4 tuku
   
'>>>>> .AddNew��SQL(INSERT)���ɕύX�@2009/06/29 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With pXSDC4
        
        sql = "INSERT INTO XSDC4 ("
        sql = sql & " XTALC4"           ' 1:��ۯ�ID�E�����ԍ�
        sql = sql & ",INPOSC4"          ' 2:�������J�n�ʒu
        sql = sql & ",KCKNTC4"          ' 3:�H���A��
        sql = sql & ",HINBC4"           ' 4:�i��
        sql = sql & ",REVNUMC4"         ' 5:���i�ԍ������ԍ�
        sql = sql & ",FACTORYC4"        ' 6:�H��
        sql = sql & ",OPEC4"            ' 7:���Ə���
        sql = sql & ",KNKTC4"           ' 8:�Ǘ��H��
        sql = sql & ",WKKTC4"           ' 9:�H��
        sql = sql & ",WKKDC4"           '10:��Ƌ敪
        sql = sql & ",MACOC4"           '11:������
        sql = sql & ",SXLIDC4"          '12:SXLID
        sql = sql & ",FCODEC4"          '13:���躰��
        sql = sql & ",PUCUTLC4"         '14:�s�ǒ���
        sql = sql & ",PUCUTWC4"         '15:�s�Ǐd��
        sql = sql & ",PUCUTMC4"         '16:�s�ǖ���
        sql = sql & ",FKUBC4"           '17:�s�ǋ敪
        sql = sql & ",TDAYC4"           '18:�o�^���t
        sql = sql & ",KDAYC4"           '19:�X�V���t
        sql = sql & ",SUMITBC3"         '20:SUMMIT���M�׸�
        sql = sql & ",SNDKC3"           '21:���M�׸�
        sql = sql & ",SNDDAYC3"         '22:���M���t
        sql = sql & ")"
        sql = sql & "VALUES (" & vbLf

        ' 1:��ۯ�ID�E�����ԍ�
        If .XTALC4 <> "" Then
            sql = sql & " '" & .XTALC4 & "'" & vbLf
        Else
            sql = sql & " '" & Space(12) & "'" & vbLf
        End If

        ' 2:�������J�n�ʒu
        If .INPOSC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.INPOSC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 3:�H���A��
        If .KCKNTC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.KCKNTC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 4:�i��
        If .HINBC4 <> "" Then
            sql = sql & ",'" & .HINBC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        ' 5:���i�ԍ������ԍ�
        If .REVNUMC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.REVNUMC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 6:�H��
        If .FACTORYC4 <> "" Then
            sql = sql & ",'" & .FACTORYC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 7:���Ə���
        If .OPEC4 <> "" Then
            sql = sql & ",'" & .OPEC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 8:�Ǘ��H��
        If .KNKTC4 <> "" Then
            sql = sql & ",'" & .KNKTC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        ' 9:�H��
        If .WKKTC4 <> "" Then
            sql = sql & ",'" & .WKKTC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '10:��Ƌ敪
        If .WKKDC4 <> "" Then
            sql = sql & ",'" & .WKKDC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '11:������
        If .MACOC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.MACOC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '12:SXLID
        If .SXLIDC4 <> "" Then
            sql = sql & ",'" & .SXLIDC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(13) & "'" & vbLf
        End If

        '13:���躰��
        If .FCODEC4 <> "" Then
            sql = sql & ",'" & .FCODEC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '14:�s�ǒ���
        If .PUCUTLC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.PUCUTLC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '15:�s�Ǐd��
        If .PUCUTWC4 <> "" Then
            sql = sql & ",'" & CStr(CLng(.PUCUTWC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '16:�s�ǖ���
        If .PUCUTMC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.PUCUTMC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '17:�s�ǋ敪
        If (.FKUBC4 <> "") Then
            sql = sql & ",'" & .FKUBC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '18:�o�^���t
        sql = sql & "," & nowtime_sql & vbLf

        '19:�X�V���t
        sql = sql & "," & nowtime_sql & vbLf

        '20:SUMMIT���M�׸�
        sql = sql & ",'0'" & vbLf

        '21:���M�׸�
        sql = sql & ",'0'" & vbLf

        '22:���M���t
        sql = sql & ",NULL" & vbLf

        sql = sql & ")" & vbLf
    
        'SQL�����s
        If OraDB.ExecuteSQL(sql) < 1 Then
            GoTo proc_err
        End If

    End With
'<<<<< .AddNew��SQL(INSERT)���ɕύX�@2009/06/29 SETsw kubota ------------------

    CreateXSDC4 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", "DB", sDbName)
    CreateXSDC4 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function




