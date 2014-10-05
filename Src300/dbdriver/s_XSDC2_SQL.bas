Attribute VB_Name = "s_XSDC2_SQL"
'��������(��ۯ�) (XSDC2) �����֐�

'***�e�[�u���uXSDC2�v�ւ̃f�[�^�A�N�Z�X�֐�***
'������ ���Ұ��ɒl��Ă��鎞�A�܂��S�ď��������邱��

Option Explicit

'����������(��ۯ�)
Public Type typ_XSDC2
    CRYNUMC2 As String * 12      ' ��ۯ�ID������ԍ�
    KCNTC2 As Integer            ' �H���A��
    XTALC2 As String * 12        ' �����ԍ�
    INPOSC2 As Integer           ' �������J�n�ʒu
    NEKKNTC2 As String * 5       ' �ŏI�ʉߊǗ��H��
    NEWKNTC2 As String * 5       ' �ŏI�ʉߍH��
    NEWKKBC2 As String * 2       ' �ŏI�ʉߍ�Ƌ敪
    NEMACOC2 As Integer          ' �ŏI�ʉߏ�����
    GNKKNTC2 As String * 5       ' ���݊Ǘ��H��
    GNWKNTC2 As String * 5       ' ���ݍH��
    GNWKKBC2 As String * 2       ' ���ݍ�Ƌ敪
    GNMACOC2 As Integer          ' ���ݏ�����
    GNDAYC2 As Date              ' ���ݏ������t
    GNLC2 As Integer             ' ���ݒ���
    GNWC2 As Long                ' ���ݏd��
    GNMC2 As Integer             ' ���ݖ���
    SUMITLC2 As Integer          ' SUMMIT����
    SUMITWC2 As Long             ' SUMMIT�d��
    SUMITMC2 As Integer          ' SUMMIT����
    CHGC2 As Long                ' ����ޗ�
    KAKOUBC2 As String * 1       ' ���H�敪
    KEIDAYC2 As Date             ' �v����t
    GNTKUBC2 As String * 3       ' �I�敪
    GNTNOC2 As String * 4        ' �I�ԍ�
    XTWORKC2 As String * 2       ' �����H��
    WFWORKC2 As String * 2       ' ���ʐ���
    LSTATBC2 As String * 1       ' �ŏI��ԋ敪
    RSTATBC2 As String * 1       ' ������ԋ敪
    LUFRCC2 As String * 3        ' �i�㺰��
    LUFRBC2 As String * 1        ' �i��敪
    LDFRCC2 As String * 3        ' �i������
    LDFRBC2 As String * 1        ' �i���敪
    HOLDCC2 As String * 3        ' ΰ��޺���
    HOLDBC2 As String * 1        ' �z�[���h�敪
    EXKUBC2 As String * 1        ' ��O�敪
    HENPKC2 As String * 1        ' �ԕi�敪
    LIVKC2 As String * 1         ' �����敪
    KANKC2 As String * 1         ' �����敪
    NFC2 As String * 1           ' ���ɋ敪
    SAKJC2 As String * 1         ' �폜�敪
    TDAYC2 As Date               ' �o�^���t
    KDAYC2 As Date               ' �X�V���t
    SUMITBC2 As String * 1       ' SUMMIT���M�t���O
    SNDKC2 As String * 1         ' ���M�t���O
    SNDDAYC2 As Date             ' ���M���t
' 2003.06.11 Y.KATABAMI tuika
    PRIORITYC2 As String * 1     ' �D��x
    CUTCNTC2 As String * 1       ' �V�K�^�Đ؋敪
    '2005/07
    HOLDKTC2 As String * 5
    RPCRYNUMC2 As String * 12    ' �e��ۯ�ID�@05/09/20 ooba
    BDCAUSC2 As String * 3       ' �s�Ǘ��R�@05/12/01 ooba
''���ǉ� START SPT�p���э쐬���@�ύX 2006/06/05 SMP-OKAMOTO
    REALLC2 As Integer           ' ������
    REALWC2 As Long              ' ���d��
''���ǉ� END   SPT�p���э쐬���@�ύX 2006/06/05 SMP-OKAMOTO
    KBLKFLGC2 As String * 1      ' �֘A��ۯ��׸ށ@06/10/31 ooba
    KIKBNC2 As String            ' �����ʋ敪   2006/11/10 SETsw kubota
    PLANTCATC2 As String         ' ���� 2007/08/22 SPK Tsutsumi Add
    STCIDC2 As String            ' STC��ۯ�ID�@08/06/16 ooba
End Type

'�X�V�p
Public Type typ_XSDC2_Update
    CRYNUMC2 As String        ' ��ۯ�ID������ԍ�
    KCNTC2 As String          ' �H���A��
    XTALC2 As String          ' �����ԍ�
    INPOSC2 As String         ' �������J�n�ʒu
    NEKKNTC2 As String        ' �ŏI�ʉߊǗ��H��
    NEWKNTC2 As String        ' �ŏI�ʉߍH��
    NEWKKBC2 As String        ' �ŏI�ʉߍ�Ƌ敪
    NEMACOC2 As String        ' �ŏI�ʉߏ�����
    GNKKNTC2 As String        ' ���݊Ǘ��H��
    GNWKNTC2 As String        ' ���ݍH��
    GNWKKBC2 As String        ' ���ݍ�Ƌ敪
    GNMACOC2 As String        ' ���ݏ�����
    GNDAYC2 As String         ' ���ݏ������t
    GNLC2 As String           ' ���ݒ���
    GNWC2 As String           ' ���ݏd��
    GNMC2 As String           ' ���ݖ���
    SUMITLC2 As String        ' SUMMIT����
    SUMITWC2 As String        ' SUMMIT�d��
    SUMITMC2 As String        ' SUMMIT����
    CHGC2 As String           ' ����ޗ�
    KAKOUBC2 As String        ' ���H�敪
    KEIDAYC2 As String        ' �v����t
    GNTKUBC2 As String        ' �I�敪
    GNTNOC2 As String         ' �I�ԍ�
    XTWORKC2 As String        ' �����H��
    WFWORKC2 As String        ' ���ʐ���
    LSTATBC2 As String        ' �ŏI��ԋ敪
    RSTATBC2 As String        ' ������ԋ敪
    LUFRCC2 As String         ' �i�㺰��
    LUFRBC2 As String         ' �i��敪
    LDFRCC2 As String         ' �i������
    LDFRBC2 As String         ' �i���敪
    HOLDCC2 As String         ' ΰ��޺���
    HOLDBC2 As String         ' �z�[���h�敪
    EXKUBC2 As String         ' ��O�敪
    HENPKC2 As String         ' �ԕi�敪
    LIVKC2 As String          ' �����敪
    KANKC2 As String          ' �����敪
    NFC2 As String            ' ���ɋ敪
    SAKJC2 As String          ' �폜�敪
    TDAYC2 As String          ' �o�^���t
    KDAYC2 As String          ' �X�V���t
    SUMITBC2 As String        ' SUMMIT���M�t���O
    SNDKC2 As String          ' ���M�t���O
    SNDDAYC2 As String        ' ���M���t
' 2003.06.11 Y.KATABAMI tuika
    PRIORITYC2 As String * 1     ' �D��x
    CUTCNTC2 As String * 1       ' �V�K�^�Đ؋敪
    '2005/07
    HOLDKTC2 As String * 5
    RPCRYNUMC2 As String * 12    ' �e��ۯ�ID�@05/09/20 ooba
    BDCAUSC2 As String * 3       ' �s�Ǘ��R�@05/12/01 ooba
''���ǉ� START SPT�p���э쐬���@�ύX 2006/06/05 SMP-OKAMOTO
    REALLC2 As String           ' ������
    REALWC2 As String           ' ���d��
''���ǉ� END   SPT�p���э쐬���@�ύX 2006/06/05 SMP-OKAMOTO
    KBLKFLGC2 As String * 1      ' �֘A��ۯ��׸ށ@06/10/31 ooba
    KIKBNC2   As String          ' �����ʋ敪   2006/11/10 SETsw kubota
    PLANTCATC2 As String         ' ���� 2007/08/22 SPK Tsutsumi Add
    STCIDC2 As String            ' STC��ۯ�ID�@08/06/16 ooba
End Type

'��SELECT��

'�T�v      :�e�[�u���uXSDC2�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO   ,�^               ,����
'          :records()     ,O    ,typ_XSDC2     ,���o���R�[�h
'          :sqlWhere      ,I    ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I    ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :

Public Function DBDRV_GetXSDC2(records() As typ_XSDC2, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim sql As String       'SQL�S��
    Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long


    ''SQL��g�ݗ��Ă�
    sqlBase = "Select * From XSDC2"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDC2 = FUNCTION_RETURN_FAILURE
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
            If IsNull(rs.Fields("CRYNUMC2")) = False Then .CRYNUMC2 = rs.Fields("CRYNUMC2")
            If IsNull(rs.Fields("KCNTC2")) = False Then .KCNTC2 = rs.Fields("KCNTC2")
            If IsNull(rs.Fields("XTALC2")) = False Then .XTALC2 = rs.Fields("XTALC2")
            If IsNull(rs.Fields("INPOSC2")) = False Then .INPOSC2 = rs.Fields("INPOSC2")
            If IsNull(rs.Fields("NEKKNTC2")) = False Then .NEKKNTC2 = rs.Fields("NEKKNTC2")
            If IsNull(rs.Fields("NEWKNTC2")) = False Then .NEWKNTC2 = rs.Fields("NEWKNTC2")
            If IsNull(rs.Fields("NEWKKBC2")) = False Then .NEWKKBC2 = rs.Fields("NEWKKBC2")
            If IsNull(rs.Fields("NEMACOC2")) = False Then .NEMACOC2 = rs.Fields("NEMACOC2")
            If IsNull(rs.Fields("GNKKNTC2")) = False Then .GNKKNTC2 = rs.Fields("GNKKNTC2")
            If IsNull(rs.Fields("GNWKNTC2")) = False Then .GNWKNTC2 = rs.Fields("GNWKNTC2")
            If IsNull(rs.Fields("GNWKKBC2")) = False Then .GNWKKBC2 = rs.Fields("GNWKKBC2")
            If IsNull(rs.Fields("GNMACOC2")) = False Then .GNMACOC2 = rs.Fields("GNMACOC2")
            If IsNull(rs.Fields("GNDAYC2")) = False Then .GNDAYC2 = rs.Fields("GNDAYC2")
            If IsNull(rs.Fields("GNLC2")) = False Then .GNLC2 = rs.Fields("GNLC2")
            If IsNull(rs.Fields("GNWC2")) = False Then .GNWC2 = rs.Fields("GNWC2")
            If IsNull(rs.Fields("GNMC2")) = False Then .GNMC2 = rs.Fields("GNMC2")
            If IsNull(rs.Fields("SUMITLC2")) = False Then .SUMITLC2 = rs.Fields("SUMITLC2")
            If IsNull(rs.Fields("SUMITWC2")) = False Then .SUMITWC2 = rs.Fields("SUMITWC2")
            If IsNull(rs.Fields("SUMITMC2")) = False Then .SUMITMC2 = rs.Fields("SUMITMC2")
            If IsNull(rs.Fields("CHGC2")) = False Then .CHGC2 = rs.Fields("CHGC2")
            If IsNull(rs.Fields("KAKOUBC2")) = False Then .KAKOUBC2 = rs.Fields("KAKOUBC2")
            If IsNull(rs.Fields("KEIDAYC2")) = False Then .KEIDAYC2 = rs.Fields("KEIDAYC2")
            If IsNull(rs.Fields("GNTKUBC2")) = False Then .GNTKUBC2 = rs.Fields("GNTKUBC2")
            If IsNull(rs.Fields("GNTNOC2")) = False Then .GNTNOC2 = rs.Fields("GNTNOC2")
            If IsNull(rs.Fields("XTWORKC2")) = False Then .XTWORKC2 = rs.Fields("XTWORKC2")
            If IsNull(rs.Fields("WFWORKC2")) = False Then .WFWORKC2 = rs.Fields("WFWORKC2")
            If IsNull(rs.Fields("LSTATBC2")) = False Then .LSTATBC2 = rs.Fields("LSTATBC2")
            If IsNull(rs.Fields("RSTATBC2")) = False Then .RSTATBC2 = rs.Fields("RSTATBC2")
            If IsNull(rs.Fields("LUFRCC2")) = False Then .LUFRCC2 = rs.Fields("LUFRCC2")
            If IsNull(rs.Fields("LUFRBC2")) = False Then .LUFRBC2 = rs.Fields("LUFRBC2")
            If IsNull(rs.Fields("LDFRCC2")) = False Then .LDFRCC2 = rs.Fields("LDFRCC2")
            If IsNull(rs.Fields("LDFRBC2")) = False Then .LDFRBC2 = rs.Fields("LDFRBC2")
            If IsNull(rs.Fields("HOLDCC2")) = False Then .HOLDCC2 = rs.Fields("HOLDCC2")
            If IsNull(rs.Fields("HOLDBC2")) = False Then .HOLDBC2 = rs.Fields("HOLDBC2")
            If IsNull(rs.Fields("EXKUBC2")) = False Then .EXKUBC2 = rs.Fields("EXKUBC2")
            If IsNull(rs.Fields("HENPKC2")) = False Then .HENPKC2 = rs.Fields("HENPKC2")
            If IsNull(rs.Fields("LIVKC2")) = False Then .LIVKC2 = rs.Fields("LIVKC2")
            If IsNull(rs.Fields("KANKC2")) = False Then .KANKC2 = rs.Fields("KANKC2")
            If IsNull(rs.Fields("NFC2")) = False Then .NFC2 = rs.Fields("NFC2")
            If IsNull(rs.Fields("SAKJC2")) = False Then .SAKJC2 = rs.Fields("SAKJC2")
            If IsNull(rs.Fields("TDAYC2")) = False Then .TDAYC2 = rs.Fields("TDAYC2")
            If IsNull(rs.Fields("KDAYC2")) = False Then .KDAYC2 = rs.Fields("KDAYC2")
            If IsNull(rs.Fields("SUMITBC2")) = False Then .SUMITBC2 = rs.Fields("SUMITBC2")
            If IsNull(rs.Fields("SNDKC2")) = False Then .SNDKC2 = rs.Fields("SNDKC2")
            If IsNull(rs.Fields("SNDDAYC2")) = False Then .SNDDAYC2 = rs.Fields("SNDDAYC2")
            '2003.06.11 Y.Katabami tuika
            If IsNull(rs.Fields("PRIORITYC2")) = False Then .PRIORITYC2 = rs.Fields("PRIORITYC2")
            If IsNull(rs.Fields("CUTCNTC2")) = False Then .CUTCNTC2 = rs.Fields("CUTCNTC2")
            '2005/07
            If IsNull(rs.Fields("HOLDKTC2")) = False Then .HOLDKTC2 = rs.Fields("HOLDKTC2")
            If IsNull(rs.Fields("RPCRYNUMC2")) = False Then .RPCRYNUMC2 = rs.Fields("RPCRYNUMC2")   '05/09/20 ooba
            If IsNull(rs.Fields("BDCAUSC2")) = False Then .BDCAUSC2 = rs.Fields("BDCAUSC2")         '05/12/01 ooba
            ''���ǉ� START SPT�p���э쐬���@�ύX 2006/06/05 SMP-OKAMOTO
            If IsNull(rs.Fields("REALLC2")) = False Then .REALLC2 = rs.Fields("REALLC2")        ''������
            If IsNull(rs.Fields("REALWC2")) = False Then .REALWC2 = rs.Fields("REALWC2")        ''���d��
            ''���ǉ� END   SPT�p���э쐬���@�ύX 2006/06/05 SMP-OKAMOTO
            If IsNull(rs.Fields("KBLKFLGC2")) = False Then .KBLKFLGC2 = rs.Fields("KBLKFLGC2")      '06/10/31 ooba
            If IsNull(rs.Fields("KIKBNC2")) = False Then .KIKBNC2 = rs.Fields("KIKBNC2")            '06/11/10 SETsw kubota
            If IsNull(rs.Fields("PLANTCATC2")) = False Then .PLANTCATC2 = rs.Fields("PLANTCATC2")            '07/08/22 SPK Tsutsumi Add
            If IsNull(rs.Fields("STCIDC2")) = False Then .STCIDC2 = rs.Fields("STCIDC2")            '08/06/16 ooba
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDC2 = FUNCTION_RETURN_SUCCESS
End Function

'��UPDATE��

'���X�V���ڂ��\���̂ɃZ�b�g���Ĉ����n��

'�T�v      :�e�[�u���uXSDC2�v���X�V���� ptrn1
'���Ұ�    :�ϐ���        ,IO  ,�^               ,����
'          :records()     ,O   ,typ_XSDC2     ,�X�V���R�[�h
'          :sqlWhere      ,I   ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I   ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O   ,FUNCTION_RETURN  ,���o�̐���
'����      :

Public Function UpdateXSDC2(records As typ_XSDC2_Update, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
On Error GoTo proc_err
    gErr.Push "s_XSDC2_SQL.bas -- Function UpdateXSDC2"

    Dim sql As String       'SQL�S��
'    Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
'    Dim rs As OraDynaset    'RecordSet
'    Dim rs2 As OraDynaset
    Dim recCnt As Long      '���R�[�h��
'    Dim i As Long
    Dim nowtime As Date
    Dim nowtime_sql As String   '�T�[�o����(SQL��)
    
    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾����悤�ɕύX 2003/6/4 tuku
    
'>>>>> .Edit��SQL(UPDATE)���ɕύX�@2009/06/22 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"

    With records
        
        ''SQL��g�ݗ��Ă�
        sql = "UPDATE XSDC2 SET" & vbLf
        
        ''�X�V���t
        sql = sql & " KDAYC2 = " & nowtime_sql & vbLf
        
        ''�u���b�NID�E�����ԍ�
        If .CRYNUMC2 <> "" And Left(.CRYNUMC2, 1) <> vbNullChar Then
            sql = sql & ",CRYNUMC2 = '" & .CRYNUMC2 & "'" & vbLf
        End If
        
        ''�H���A��
        If .KCNTC2 <> "" Then
            sql = sql & ",KCNTC2 = '" & CStr(CInt(.KCNTC2)) & "'" & vbLf
        End If
        
        ''�����ԍ�
        If .XTALC2 <> "" And Left(.XTALC2, 1) <> vbNullChar Then
            sql = sql & ",XTALC2 = '" & .XTALC2 & "'" & vbLf
        End If
        
        ''�������J�n�ʒu
        If .INPOSC2 <> "" Then
            sql = sql & ",INPOSC2 = '" & CStr(CInt(.INPOSC2)) & "'" & vbLf
        End If
        
        ''�ŏI�ʉߊǗ��H��
        If .NEKKNTC2 <> "" And Left(.NEKKNTC2, 1) <> vbNullChar Then
            sql = sql & ",NEKKNTC2 = '" & .NEKKNTC2 & "'" & vbLf
        End If
        
        ''�ŏI�ʉߍH��
        If .NEWKNTC2 <> "" And Left(.NEWKNTC2, 1) <> vbNullChar Then
            sql = sql & ",NEWKNTC2 = '" & .NEWKNTC2 & "'" & vbLf
        End If
        
        ''�ŏI�ʉߍ�Ƌ敪
        If .NEWKKBC2 <> "" And Left(.NEWKKBC2, 1) <> vbNullChar Then
            sql = sql & ",NEWKKBC2 = '" & .NEWKKBC2 & "'" & vbLf
        End If
        
        ''�ŏI�ʉߏ�����
        If .NEMACOC2 <> "" Then
            sql = sql & ",NEMACOC2 = '" & CStr(CInt(.NEMACOC2)) & "'" & vbLf
        End If
        
        ''���݊Ǘ��H��
        If .GNKKNTC2 <> "" And Left(.GNKKNTC2, 1) <> vbNullChar Then
            sql = sql & ",GNKKNTC2 = '" & .GNKKNTC2 & "'" & vbLf
        End If
        
        ''���ݍH��
        If .GNWKNTC2 <> "" And Left(.GNWKNTC2, 1) <> vbNullChar Then
            sql = sql & ",GNWKNTC2 = '" & .GNWKNTC2 & "'" & vbLf
        End If
        
        ''���ݍ�Ƌ敪
        If .GNWKKBC2 <> "" And Left(.GNWKKBC2, 1) <> vbNullChar Then
            sql = sql & ",GNWKKBC2 = '" & .GNWKKBC2 & "'" & vbLf
        End If

        ''���ݏ�����
        If .GNMACOC2 <> "" Then
            sql = sql & ",GNMACOC2 = '" & CStr(CInt(.GNMACOC2)) & "'" & vbLf
        End If

        ''���ݏ������t
        If .GNDAYC2 <> "" Then
            sql = sql & ",GNDAYC2 = TO_DATE('" & Format$(CDate(.GNDAYC2), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''���ݒ���
        If .GNLC2 <> "" Then
            sql = sql & ",GNLC2 = '" & CStr(CInt(.GNLC2)) & "'" & vbLf
        End If
        
        ''���ݏd��
        If .GNWC2 <> "" Then
            sql = sql & ",GNWC2 = '" & CStr(CLng(.GNWC2)) & "'" & vbLf
        End If
        
        ''���ݖ���
        If .GNMC2 <> "" Then
            sql = sql & ",GNMC2 = '" & CStr(CInt(.GNMC2)) & "'" & vbLf
        End If
        
        ''SUMIT����
        If .SUMITLC2 <> "" Then
            sql = sql & ",SUMITLC2 = '" & CStr(CInt(.SUMITLC2)) & "'" & vbLf
        End If
        
        ''SUMIT�d��
        If .SUMITWC2 <> "" Then
            sql = sql & ",SUMITWC2 = '" & CStr(CLng(.SUMITWC2)) & "'" & vbLf
        End If
        
        ''SUMIT����
        If .SUMITMC2 <> "" Then
            sql = sql & ",SUMITMC2 = '" & CStr(CInt(.SUMITMC2)) & "'" & vbLf
        End If
        
        ''�`���[�W��
        If .CHGC2 <> "" Then
            sql = sql & ",CHGC2 = '" & CStr(CLng(.CHGC2)) & "'" & vbLf
        End If
        
        ''���H�敪
        If .KAKOUBC2 <> "" And Left(.KAKOUBC2, 1) <> vbNullChar Then
            sql = sql & ",KAKOUBC2 = '" & .KAKOUBC2 & "'" & vbLf
        End If
        
        ''�v����t
        If .KEIDAYC2 <> "" Then
            sql = sql & ",KEIDAYC2 = TO_DATE('" & Format$(CDate(.KEIDAYC2), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''�I�敪
        If .GNTKUBC2 <> "" And Left(.GNTKUBC2, 1) <> vbNullChar Then
            sql = sql & ",GNTKUBC2 = '" & .GNTKUBC2 & "'" & vbLf
        End If
        
        ''�I�ԍ�
        If .GNTNOC2 <> "" And Left(.GNTNOC2, 1) <> vbNullChar Then
            sql = sql & ",GNTNOC2 = '" & .GNTNOC2 & "'" & vbLf
        End If
        
        ''�����H��
        If .XTWORKC2 <> "" And Left(.XTWORKC2, 1) <> vbNullChar Then
            sql = sql & ",XTWORKC2 = '" & .XTWORKC2 & "'" & vbLf
        End If
        
        ''�E�F�[�n����
        If .WFWORKC2 <> "" And Left(.WFWORKC2, 1) <> vbNullChar Then
            sql = sql & ",WFWORKC2 = '" & .WFWORKC2 & "'" & vbLf
        End If
        
        ''�ŏI��ԋ敪
        If .LSTATBC2 <> "" And Left(.LSTATBC2, 1) <> vbNullChar Then
            sql = sql & ",LSTATBC2 = '" & .LSTATBC2 & "'" & vbLf
        End If
        
        ''������ԋ敪
        If .RSTATBC2 <> "" And Left(.RSTATBC2, 1) <> vbNullChar Then
            sql = sql & ",RSTATBC2 = '" & .RSTATBC2 & "'" & vbLf
        End If
        
        ''�i��R�[�h
        If .LUFRCC2 <> "" And Left(.LUFRCC2, 1) <> vbNullChar Then
            sql = sql & ",LUFRCC2 = '" & .LUFRCC2 & "'" & vbLf
        End If
        
        ''�i��敪
        If .LUFRBC2 <> "" And Left(.LUFRBC2, 1) <> vbNullChar Then
            sql = sql & ",LUFRBC2 = '" & .LUFRBC2 & "'" & vbLf
        End If
        
        ''�i���R�[�h
        If .LDFRCC2 <> "" And Left(.LDFRCC2, 1) <> vbNullChar Then
            sql = sql & ",LDFRCC2 = '" & .LDFRCC2 & "'" & vbLf
        End If
        
        ''�i���敪
        If .LDFRBC2 <> "" And Left(.LDFRBC2, 1) <> vbNullChar Then
            sql = sql & ",LDFRBC2 = '" & .LDFRBC2 & "'" & vbLf
        End If
        
        ''�z�[���h�R�[�h
        If .HOLDCC2 <> "" And Left(.HOLDCC2, 1) <> vbNullChar Then
            sql = sql & ",HOLDCC2 = '" & .HOLDCC2 & "'" & vbLf
        End If
        
        ''�z�[���h�敪
        If .HOLDBC2 <> "" And Left(.HOLDBC2, 1) <> vbNullChar Then
            sql = sql & ",HOLDBC2 = '" & .HOLDBC2 & "'" & vbLf
        End If
        
        ''��O�敪
        If .EXKUBC2 <> "" And Left(.EXKUBC2, 1) <> vbNullChar Then
            sql = sql & ",EXKUBC2 = '" & .EXKUBC2 & "'" & vbLf
        End If
        
        ''�ԕi�敪
        If .HENPKC2 <> "" And Left(.HENPKC2, 1) <> vbNullChar Then
            sql = sql & ",HENPKC2 = '" & .HENPKC2 & "'" & vbLf
        End If
        
        ''�����敪
        If .LIVKC2 <> "" And Left(.LIVKC2, 1) <> vbNullChar Then
            sql = sql & ",LIVKC2 = '" & .LIVKC2 & "'" & vbLf
        End If
        
        ''�����敪
        If .KANKC2 <> "" And Left(.KANKC2, 1) <> vbNullChar Then
            sql = sql & ",KANKC2 = '" & .KANKC2 & "'" & vbLf
        End If
        
        ''���ɋ敪
        If .NFC2 <> "" And Left(.NFC2, 1) <> vbNullChar Then
            sql = sql & ",NFC2 = '" & .NFC2 & "'" & vbLf
        End If
        
        ''�폜�敪
        If .SAKJC2 <> "" And Left(.SAKJC2, 1) <> vbNullChar Then
            sql = sql & ",SAKJC2 = '" & .SAKJC2 & "'" & vbLf
        End If
        
        ''�o�^���t
        If .TDAYC2 <> "" Then
            sql = sql & ",TDAYC2 = TO_DATE('" & Format$(CDate(.TDAYC2), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''SUMIT���M�t���O
        If .SUMITBC2 <> "" And Left(.SUMITBC2, 1) <> vbNullChar Then
            sql = sql & ",SUMITBC2 = '" & .SUMITBC2 & "'" & vbLf
        End If
        
        ''���M�t���O
        If .SNDKC2 <> "" And Left(.SNDKC2, 1) <> vbNullChar Then
            sql = sql & ",SNDKC2 = '" & .SNDKC2 & "'" & vbLf
        End If
        
        ''���M���t
        If .SNDDAYC2 <> "" Then
            sql = sql & ",SNDDAYC2 = TO_DATE('" & Format$(CDate(.SNDDAYC2), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''�D��x
        If .PRIORITYC2 <> "" And Left(.PRIORITYC2, 1) <> vbNullChar Then
            sql = sql & ",PRIORITYC2 = '" & .PRIORITYC2 & "'" & vbLf
        End If
        
        ''�ؒf�����敪
        If .CUTCNTC2 <> "" And Left(.CUTCNTC2, 1) <> vbNullChar Then
            sql = sql & ",CUTCNTC2 = '" & .CUTCNTC2 & "'" & vbLf
        End If
        
        ''ΰ��ލH��
        If .HOLDKTC2 <> "" And Left(.HOLDKTC2, 1) <> vbNullChar Then
            sql = sql & ",HOLDKTC2 = '" & .HOLDKTC2 & "'" & vbLf
        End If
        
        ''�e�u���b�NID
        If .RPCRYNUMC2 <> "" And Left(.RPCRYNUMC2, 1) <> vbNullChar Then
            sql = sql & ",RPCRYNUMC2 = '" & .RPCRYNUMC2 & "'" & vbLf
        End If
        
        ''�s�Ǘ��R
        If .BDCAUSC2 <> "" And Left(.BDCAUSC2, 1) <> vbNullChar Then
            sql = sql & ",BDCAUSC2 = '" & .BDCAUSC2 & "'" & vbLf
        End If
        
        ''������
        If .REALLC2 <> "" Then
            sql = sql & ",REALLC2 = '" & CStr(CInt(.REALLC2)) & "'" & vbLf
        End If
        
        ''���d��
        If .REALWC2 <> "" Then
            sql = sql & ",REALWC2 = '" & CStr(CLng(.REALWC2)) & "'" & vbLf
        End If
        
        ''�֘A�u���b�N�t���O
        If .KBLKFLGC2 <> "" And Left(.KBLKFLGC2, 1) <> vbNullChar Then
            sql = sql & ",KBLKFLGC2 = '" & .KBLKFLGC2 & "'" & vbLf
        End If
        
        ''���Ə��敪
        If .PLANTCATC2 <> "" And Left(.PLANTCATC2, 2) <> vbNullChar Then
            sql = sql & ",PLANTCATC2 = '" & .PLANTCATC2 & "'" & vbLf
        End If
        
        ''STC�u���b�NID
        If .STCIDC2 <> "" And Left(.STCIDC2, 1) <> vbNullChar Then
            sql = sql & ",STCIDC2 = '" & .STCIDC2 & "'" & vbLf
        End If
    
        sql = sql & " " & sqlWhere & vbLf
    
        'SQL�����s
        recCnt = OraDB.ExecuteSQL(sql)
        
        '�Ԃ�l��1�ȊO�̓G���[
        If recCnt < 0 Then
            GoTo proc_err
        ElseIf recCnt = 0 Then
            '0���X�V�c�G���[(�����ʂ�)
            UpdateXSDC2 = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf recCnt > 1 Then
            '�������X�V�c�G���[(�����͕���SELECT�����ŏ��̈ꌏ�̂ݍX�V)
            UpdateXSDC2 = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    
    End With
'<<<<< .Edit��SQL(UPDATE)���ɕύX�@2009/06/22 SETsw kubota ------------------

    UpdateXSDC2 = FUNCTION_RETURN_SUCCESS
    

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    UpdateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'��INSERT��  NULL�̏ꍇ�Achar�Ȃ�X�y�[�X�ANumber�Ȃ�NULL������

'�T�v      :�e�[�u���uXSDC2�v�Ƀ��R�[�h��}������
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pXSDC2 �@�@  ,I  ,typ_XSDC2_Update   ,XSDC2�X�V�p�ް�
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function CreateXSDC2(pXSDC2 As typ_XSDC2_Update, sErrMsg As String) As FUNCTION_RETURN


    Dim sql As String
    Dim sDbName As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim rs2 As OraDynaset
'    Dim recCnt As Long      '���R�[�h��
    Dim nowtime As Date
    Dim nowtime_sql As String   '�T�[�o����(SQL��)
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDC2_SQL.bas -- Function CreateXSDC2"
    sErrMsg = ""
    sDbName = "XSDC2"
     'nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾����悤�ɕύX 2003/6/4 tuku

'>>>>> .AddNew��SQL(INSERT)���ɕύX�@2009/06/22 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With pXSDC2
        
        sql = "INSERT INTO XSDC2 ("
        sql = sql & " CRYNUMC2"         ' 1:��ۯ�ID�E�����ԍ�
        sql = sql & ",KCNTC2"           ' 2:�H���A�Ԏ擾
        sql = sql & ",XTALC2"           ' 3:�����ԍ�
        sql = sql & ",INPOSC2"          ' 4:�������J�n�ʒu
        sql = sql & ",NEKKNTC2"         ' 5:�ŏI�ʉߊǗ��H��
        sql = sql & ",NEWKNTC2"         ' 6:�ŏI�ʉߍH��
        sql = sql & ",NEWKKBC2"         ' 7:�ŏI�ʉߍ�Ƌ敪
        sql = sql & ",NEMACOC2"         ' 8:�ŏI�ʉߏ�����
        sql = sql & ",GNKKNTC2"         ' 9:���݊Ǘ��H��
        sql = sql & ",GNWKNTC2"         '10:���ݍH��
        sql = sql & ",GNWKKBC2"         '11:���ݍ�Ƌ敪
        sql = sql & ",GNMACOC2"         '12:���ݏ�����
        sql = sql & ",GNDAYC2"          '13:���ݏ������t(�o�^����)
        sql = sql & ",GNLC2"            '14:���ݒ���
        sql = sql & ",GNWC2"            '15:���ݏd��
        sql = sql & ",GNMC2"            '16:���ݖ���
        sql = sql & ",SUMITLC2"         '17:SUMMIT����
        sql = sql & ",SUMITWC2"         '18:SUMMIT�d��
        sql = sql & ",SUMITMC2"         '19:SUMMIT����
        sql = sql & ",CHGC2"            '20:����ޗ�
        sql = sql & ",KAKOUBC2"         '21:���H�敪
        If .KEIDAYC2 <> "" Then
            sql = sql & ",KEIDAYC2"         '22:�v����t
        End If
        sql = sql & ",GNTKUBC2"         '23:�I�敪
        sql = sql & ",GNTNOC2"          '24:�I�ԍ�
        sql = sql & ",XTWORKC2"         '25:�����H��
        sql = sql & ",WFWORKC2"         '26:���ʐ���
        sql = sql & ",LSTATBC2"         '27:�ŏI��ԋ敪
        sql = sql & ",RSTATBC2"         '28:������ԋ敪
        sql = sql & ",LUFRCC2"          '29:�i�㺰��
        sql = sql & ",LUFRBC2"          '30:�i��敪
        sql = sql & ",LDFRCC2"          '31:�i������
        sql = sql & ",LDFRBC2"          '32:�i���敪
        sql = sql & ",HOLDCC2"          '33:ΰ��޺���
        sql = sql & ",HOLDBC2"          '34:ΰ��ދ敪
        sql = sql & ",EXKUBC2"          '35:��O�敪
        sql = sql & ",HENPKC2"          '36:�ԕi�敪
        sql = sql & ",LIVKC2"           '37:�����敪
        sql = sql & ",KANKC2"           '38:�����敪
        sql = sql & ",NFC2"             '39:���ɋ敪
        sql = sql & ",SAKJC2"           '40:�폜�敪
        sql = sql & ",TDAYC2"           '41:�o�^���t
        sql = sql & ",KDAYC2"           '42:�X�V���t
        sql = sql & ",SUMITBC2"         '43:SUMMIT���M�׸�
        sql = sql & ",SNDKC2"           '44:���M�׸�
        sql = sql & ",SNDDAYC2"         '45:���M���t
        sql = sql & ",PRIORITYC2"       '46:�D��x
        sql = sql & ",CUTCNTC2"         '47:�V�K�^�Đ؋敪
        sql = sql & ",HOLDKTC2"         '48:ΰ��ލH��
        sql = sql & ",RPCRYNUMC2"       '49:�e��ۯ�ID
        sql = sql & ",BDCAUSC2"         '50:�s�Ǘ��R
        sql = sql & ",REALLC2"          '51:������
        sql = sql & ",REALWC2"          '52:���d��
        sql = sql & ",KBLKFLGC2"        '53:�֘A��ۯ��׸�
        sql = sql & ",PLANTCATC2"       '54:����
        sql = sql & ",STCIDC2"          '55:STC��ۯ�ID
        sql = sql & ")"
        sql = sql & "VALUES (" & vbLf

        ' 1:��ۯ�ID�E�����ԍ�
        If .CRYNUMC2 <> "" And Left(.CRYNUMC2, 1) <> vbNullChar Then
            sql = sql & " '" & .CRYNUMC2 & "'" & vbLf
        Else
            sql = sql & " '" & Space(12) & "'" & vbLf
        End If

        ' 2:�H���A�Ԏ擾
        If .KCNTC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.KCNTC2)) & "'" & vbLf
        Else
            sql = sql & ",1" & vbLf
        End If

        ' 3:�����ԍ�
        If .XTALC2 <> "" And Left(.XTALC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .XTALC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        ' 4:�������J�n�ʒu
        If .INPOSC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.INPOSC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If
        
        ' 5:�ŏI�ʉߊǗ��H��
        If .NEKKNTC2 <> "" And Left(.NEKKNTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEKKNTC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        ' 6:�ŏI�ʉߍH��
        If .NEWKNTC2 <> "" And Left(.NEWKNTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEWKNTC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        ' 7:�ŏI�ʉߍ�Ƌ敪
        If .NEWKKBC2 <> "" And Left(.NEWKKBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEWKKBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        ' 8:�ŏI�ʉߏ�����
        If .NEMACOC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.NEMACOC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 9:���݊Ǘ��H��
        If .GNKKNTC2 <> "" And Left(.GNKKNTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNKKNTC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '10:���ݍH��
        If .GNWKNTC2 <> "" And Left(.GNWKNTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNWKNTC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '11:���ݍ�Ƌ敪
        If .GNWKKBC2 <> "" And Left(.GNWKKBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNWKKBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '12:���ݏ�����
        If .GNMACOC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNMACOC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '13:���ݏ������t(�o�^����)
        sql = sql & "," & nowtime_sql & vbLf
        
        '14:���ݒ���
        If .GNLC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNLC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '15:���ݏd��
        If .GNWC2 <> "" Then
            sql = sql & ",'" & CStr(CLng(.GNWC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '16:���ݖ���
        If .GNMC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNMC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '17:SUMMIT����
        If .SUMITLC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.SUMITLC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '18:SUMMIT�d��
        If .SUMITWC2 <> "" Then
            sql = sql & ",'" & CStr(CLng(.SUMITWC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '19:SUMMIT����
        If .SUMITMC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.SUMITMC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '20:����ޗ�
        If .CHGC2 <> "" Then
            sql = sql & ",'" & CStr(CLng(.CHGC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '21:���H�敪
        If .KAKOUBC2 <> "" And Left(.KAKOUBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .KAKOUBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '22:�v����t
        If .KEIDAYC2 <> "" Then
            sql = sql & ",TO_DATE('" & Format$(CDate(.KEIDAYC2), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If

        '23:�I�敪
        If .GNTKUBC2 <> "" And Left(.GNTKUBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNTKUBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '24:�I�ԍ�
        If .GNTNOC2 <> "" And Left(.GNTNOC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNTNOC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(4) & "'" & vbLf
        End If

        '25:�����H��
        sql = sql & ",'" & FACTORYCD & "'" & vbLf

        '26:���ʐ���
        If .WFWORKC2 <> "" And Left(.WFWORKC2, 1) <> vbNullChar Then
            'sql = sql & ",'" & .WFWORKC2 & "'"
            sql = sql & ",'" & .XTWORKC2 & "'" & vbLf   '�����ʂ��
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '27:�ŏI��ԋ敪
        If .LSTATBC2 <> "" And Left(.LSTATBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LSTATBC2 & "'" & vbLf
        Else
            sql = sql & ",'T'" & vbLf
        End If

        '28:������ԋ敪
        If .RSTATBC2 <> "" And Left(.RSTATBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .RSTATBC2 & "'" & vbLf
        Else
            sql = sql & ",'T'" & vbLf
        End If

        '29:�i�㺰��
        If .LUFRCC2 <> "" And Left(.LUFRCC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LUFRCC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '30:�i��敪
        If .LUFRBC2 <> "" And Left(.LUFRBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LUFRBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '31:�i������
        If .LDFRCC2 <> "" And Left(.LDFRCC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LDFRCC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '32:�i���敪
        If .LDFRBC2 <> "" And Left(.LDFRBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LDFRBC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '33:ΰ��޺���
        If .HOLDCC2 <> "" And Left(.HOLDCC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDCC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '34:ΰ��ދ敪
        If .HOLDBC2 <> "" And Left(.HOLDBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDBC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '35:��O�敪
        If .EXKUBC2 <> "" And Left(.EXKUBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .EXKUBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '36:�ԕi�敪
        If .HENPKC2 <> "" And Left(.HENPKC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .HENPKC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '37:�����敪
        If .LIVKC2 <> "" And Left(.LIVKC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LIVKC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '38:�����敪
        If .KANKC2 <> "" And Left(.KANKC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .KANKC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '39:���ɋ敪
        If .NFC2 <> "" And Left(.NFC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .NFC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '40:�폜�敪
        If .SAKJC2 <> "" And Left(.SAKJC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .SAKJC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '41:�o�^���t
        sql = sql & "," & nowtime_sql & vbLf
        
        '42:�X�V���t
        sql = sql & "," & nowtime_sql & vbLf

        '43:SUMMIT���M�׸�
        sql = sql & ",'0'" & vbLf

        '44:���M�׸�
        sql = sql & ",'0'" & vbLf

        '45:���M���t
        sql = sql & ",NULL" & vbLf

        '46:�D��x
        If .PRIORITYC2 <> "" And Left(.PRIORITYC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .PRIORITYC2 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '47:�V�K�^�Đ؋敪
        If .CUTCNTC2 <> "" And Left(.CUTCNTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .CUTCNTC2 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '48:ΰ��ލH��
        If .HOLDKTC2 <> "" And Left(.HOLDKTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDKTC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '49:�e��ۯ�ID
        If .RPCRYNUMC2 <> "" And Left(.RPCRYNUMC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .RPCRYNUMC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '50:�s�Ǘ��R
        If .BDCAUSC2 <> "" And Left(.BDCAUSC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .BDCAUSC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '51:������
        If .REALLC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.REALLC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '52:���d��
        If .REALWC2 <> "" Then
            sql = sql & ",'" & CStr(CLng(.REALWC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '53:�֘A��ۯ��׸�
        If .KBLKFLGC2 <> "" And Left(.KBLKFLGC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .KBLKFLGC2 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '54:����
        If .PLANTCATC2 <> "" And Left(.PLANTCATC2, 2) <> vbNullChar Then
            sql = sql & ",'" & .PLANTCATC2 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '55:STC��ۯ�ID
        If .STCIDC2 <> "" And Left(.STCIDC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .STCIDC2 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If
        
        sql = sql & ")" & vbLf
    
        'SQL�����s
        If OraDB.ExecuteSQL(sql) < 1 Then
            GoTo proc_err
        End If
        
    End With
'<<<<< .AddNew��SQL(INSERT)���ɕύX�@2009/06/22 SETsw kubota ------------------

    CreateXSDC2 = FUNCTION_RETURN_SUCCESS

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
    CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'�T�v      :���ݏ����񐔂��擾����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_sGenKotei  ,I  ,String           ,���ݍH��
'      �@�@:�߂�l       ,O  ,Integer        �@,������
Public Function GetGNMACOC(p_sCrynum As String, p_sGenKotei As String) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
''    sql = "SELECT COUNT(WKKTC3) FROM XSDC3 WHERE WKKTC3 = '" & p_sGenKotei & "'"
'    sql = "SELECT COUNT(DISTINCT(WKKTC3)) FROM XSDC3 WHERE CRYNUMC3 = '" & p_sCrynum
'    sql = sql & "' AND WKKTC3 = '" & p_sGenKotei & "'"
''    sql = sql & "' AND KCNTC3 = (SELECT MAX(KCNTC3) FROM XSDC3 WHERE CRYNUMC3 = '" & p_sCrynum
''    sql = sql & "' AND WKKTC3 = '" & p_sGenKotei & "')"
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'    GetGNMACOC = rs.Fields("COUNT(DISTINCT(WKKTC3))") + 1
    
    sql = "SELECT MACOC3 FROM XSDC3 WHERE CRYNUMC3 = '" & p_sCrynum
    sql = sql & "' AND WKKTC3 = '" & p_sGenKotei
    sql = sql & "' AND KCNTC3 = (SELECT MAX(KCNTC3) FROM XSDC3 WHERE CRYNUMC3 = '" & p_sCrynum
    sql = sql & "' AND WKKTC3 = '" & p_sGenKotei & "')"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("MACOC3")) Then
        GetGNMACOC = 1
    Else
        GetGNMACOC = CInt(rs.Fields("MACOC3")) + 1
    End If
    
End Function


'�T�v      :�ŏI�ʉߏ����񐔂��擾����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_sCrynum    ,I  ,String           ,�u���b�NID
'      �@�@:p_iInpos     ,I  ,Integer          ,�J�n�ʒu
'      �@�@:�߂�l       ,O  ,Integer        �@,������
'          :�쐬�ҁ@�@2002/11/21�@tuku
Public Function GetNEMACOC2(p_sCrynum As String) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    If Left(p_sCrynum, 1) = vbNullChar Then
        GetNEMACOC2 = 1
        Exit Function
    End If
    
    sql = "SELECT GNMACOC2 FROM XSDC2 WHERE CRYNUMC2 = '" & p_sCrynum & "'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        GetNEMACOC2 = 1
    Else
        GetNEMACOC2 = CInt(rs.Fields("GNMACOC2"))
    End If

End Function

