Attribute VB_Name = "s_XSDC1_SQL"
'��������ð���(XSDC1) �����֐�

'��TEST�p

'***�e�[�u���uXSDC1�v�ւ̃f�[�^�A�N�Z�X�֐�***
'������ ���Ұ��ɒl��Ă��鎞�A�܂��S�ď��������邱��

Option Explicit

'����������
Public Type typ_XSDC1
    XTALC1 As String * 12        ' �����ԍ�
    KNKTC1 As String * 5         ' �Ǘ��H������
    WKKTC1 As String * 5         ' �H������
    LENTOC1 As Integer           ' �����iTOP�j
    LENTKC1 As Integer           ' �����i�����j
    LENTAC1 As Integer           ' �����iTAIL�j
    PUFRELC1 As Integer          ' �ذ����
    DIA1C1 As Integer            ' �������a1
    DIA2C1 As Integer            ' �������a2
    DIA3C1 As Integer            ' �������a3
    WGHTTOC1 As Long             ' �d�ʁiTOP�j
    WGHTTKC1 As Long             ' �d�ʁi�����j
    WGHTTAC1 As Long             ' �d�ʁiTAIL�j
    WGHTFRC1 As Long             ' �d�ʁi�ذ�����j
    PUTCUTWC1 As Long            ' į�߶�ďd��
    PUWC1 As Long                ' ���グ�d��
    PUHINBC1 As String * 8       ' �_���i��
    PUCHAGC1 As Long             ' ����ޗ�
    KAKOUBC1 As String * 1       ' ���H�敪
    KEIDAYC1 As Date             ' �v����t
    SEEDC1 As String * 4         ' ����
    PUBTKBC1 As String * 3       ' BOT�󋵋敪
    JDGECC1 As String * 3        ' ���躰��
    PWTIMEC1 As Double           ' ��ܰ����
    ADDOPPC1 As Integer          ' �ǉ��ް�߈ʒu
    ADDOPCC1 As String * 7       ' �ǉ��ް���Ď��
    ADDOPVC1 As Long             ' �ǉ��ް�ߗ�
    ADDOPNC1 As String * 20      ' �ǉ��ް�ߖ�
    TSTAFFC1 As String * 8       ' �o�^�Ј�ID
    TDAYC1 As Date               ' �o�^���t
    KSTAFFC1 As String * 8       ' �X�V�Ј�ID
    KDAYC1 As Date               ' �X�V���t
    SUMITBC1 As String * 1       ' SUMMIT���M�׸�
    SNDKC1 As String * 1         ' ���M�׸�
    SNDDAYC1 As Date             ' ���M���t
    SUIFLG As String * 1         ' ����FLG     2003/10/27 tuku
    PUREVNUMC1 As Integer        ' �_���i�Ԑ����ԍ������ԍ��@04/09/28 ooba
    PUFACTORYC1 As String * 1    ' �_���i�ԍH��@04/09/28 ooba
    PUOPEC1 As String * 1        ' �_���i�ԑ��Ə����@04/09/28 ooba
    LENPUFRC1 As Integer         ' �����ذ�����@05/07/19 ooba
    WGHTPUFRC1 As Long           ' �����ذ�d�ʁ@05/07/19 ooba
    WGHTCUTRHC1 As Long          ' �ؒf��Ǖi�d�ʁ@05/07/19 ooba
''09/02/16 FAE)akiyama start
    SUICHARGE As Long                           '����`���[�W��
''09/02/16 FAE)akiyama end
End Type
    
    
'�X�V�p
Public Type typ_XSDC1_Update
    XTALC1 As String            ' �����ԍ�
    KNKTC1 As String            ' �Ǘ��H������
    WKKTC1 As String            ' �H������
    LENTOC1 As String           ' �����iTOP�j
    LENTKC1 As String           ' �����i�����j
    LENTAC1 As String           ' �����iTAIL�j
    PUFRELC1 As String          ' �ذ����
    DIA1C1 As String            ' �������a1
    DIA2C1 As String            ' �������a2
    DIA3C1 As String            ' �������a3
    WGHTTOC1 As String          ' �d�ʁiTOP�j
    WGHTTKC1 As String          ' �d�ʁi�����j
    WGHTTAC1 As String          ' �d�ʁiTAIL�j
    WGHTFRC1 As String          ' �d�ʁi�ذ�����j
    PUTCUTWC1 As String         ' į�߶�ďd��
    PUWC1 As String             ' ���グ�d��
    PUHINBC1 As String          ' �_���i��
    PUCHAGC1 As String          ' ����ޗ�
    KAKOUBC1 As String          ' ���H�敪
    KEIDAYC1 As String          ' �v����t
    SEEDC1 As String            ' ����
    PUBTKBC1 As String          ' BOT�󋵋敪
    JDGECC1 As String           ' ���躰��
    PWTIMEC1 As String          ' ��ܰ����
    ADDOPPC1 As String          ' �ǉ��ް�߈ʒu
    ADDOPCC1 As String          ' �ǉ��ް���Ď��
    ADDOPVC1 As String          ' �ǉ��ް�ߗ�
    ADDOPNC1 As String          ' �ǉ��ް�ߖ�
    TSTAFFC1 As String          ' �o�^�Ј�ID
    TDAYC1 As String            ' �o�^���t
    KSTAFFC1 As String          ' �X�V�Ј�ID
    KDAYC1 As String            ' �X�V���t
    SUMITBC1 As String          ' SUMMIT���M�׸�
    SNDKC1 As String            ' ���M�׸�
    SNDDAYC1 As String          ' ���M���t
    SUIFLG As String            ' ����FLG�@2003/10/27 tuku
    PUREVNUMC1 As String        ' �_���i�Ԑ����ԍ������ԍ��@04/09/28 ooba
    PUFACTORYC1 As String       ' �_���i�ԍH��@04/09/28 ooba
    PUOPEC1 As String           ' �_���i�ԑ��Ə����@04/09/28 ooba
    LENPUFRC1 As String         ' �����ذ�����@05/07/19 ooba
    WGHTPUFRC1 As String        ' �����ذ�d�ʁ@05/07/19 ooba
    WGHTCUTRHC1 As String       ' �ؒf��Ǖi�d�ʁ@05/07/19 ooba
'C�|OSF3����@�\�ǉ� 2007/05/11 M.Kaga START  ---
    JDGEIDC1    As String       'C-OSF3����ID
'C�|OSF3����@�\�ǉ� 2007/05/11 M.Kaga END    ---
End Type

'''�@���M�t���O�ASUMMIT���M�t���O
'Public Const SNDKC_NOTSEND = 0     '' �����M
'Public Const SNDKC_SENDING = 1     '' ���M��
'Public Const SNDKC_SENDED = 2      '' ���M�ς�
'Public Const SNDKC_WAITING = 3     '' ���M�҂�


'��SELECT��

'�T�v      :�e�[�u���uXSDC1�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO   ,�^               ,����
'          :records()     ,O    ,typ_XSDC1        ,���o���R�[�h
'          :sqlWhere      ,I    ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I    ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :

Public Function DBDRV_GetXSDC1(records() As typ_XSDC1, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim sql As String       'SQL�S��
    Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select * From XSDC1"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDC1 = FUNCTION_RETURN_FAILURE
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
            If IsNull(rs.Fields("XTALC1")) = False Then .XTALC1 = rs.Fields("XTALC1")             ' �����ԍ�
            If IsNull(rs.Fields("KNKTC1")) = False Then .KNKTC1 = rs.Fields("KNKTC1")             ' �Ǘ��H������
            If IsNull(rs.Fields("WKKTC1")) = False Then .WKKTC1 = rs.Fields("WKKTC1")             ' �H������
            If IsNull(rs.Fields("LENTOC1")) = False Then .LENTOC1 = rs.Fields("LENTOC1")           ' �����iTOP�j
            If IsNull(rs.Fields("LENTKC1")) = False Then .LENTKC1 = rs.Fields("LENTKC1")           ' �����i�����j
            If IsNull(rs.Fields("LENTAC1")) = False Then .LENTAC1 = rs.Fields("LENTAC1")           ' �����iTAIL�j
            If IsNull(rs.Fields("PUFRELC1")) = False Then .PUFRELC1 = rs.Fields("PUFRELC1")         ' �ذ����
            If IsNull(rs.Fields("DIA1C1")) = False Then .DIA1C1 = rs.Fields("DIA1C1")             ' �������a1
            If IsNull(rs.Fields("DIA2C1")) = False Then .DIA2C1 = rs.Fields("DIA2C1")             ' �������a2
            If IsNull(rs.Fields("DIA3C1")) = False Then .DIA3C1 = rs.Fields("DIA3C1")             ' �������a3
            If IsNull(rs.Fields("WGHTTOC1")) = False Then .WGHTTOC1 = rs.Fields("WGHTTOC1")         ' �d�ʁiTOP�j
            If IsNull(rs.Fields("WGHTTKC1")) = False Then .WGHTTKC1 = rs.Fields("WGHTTKC1")         ' �d�ʁi�����j
            If IsNull(rs.Fields("WGHTTAC1")) = False Then .WGHTTAC1 = rs.Fields("WGHTTAC1")         ' �d�ʁiTAIL�j
            If IsNull(rs.Fields("WGHTFRC1")) = False Then .WGHTFRC1 = rs.Fields("WGHTFRC1")         ' �d�ʁi�ذ�����j
            If IsNull(rs.Fields("PUTCUTWC1")) = False Then .PUTCUTWC1 = rs.Fields("PUTCUTWC1")       ' į�߶�ďd��
            If IsNull(rs.Fields("PUWC1")) = False Then .PUWC1 = rs.Fields("PUWC1")               ' ���グ�d��
            If IsNull(rs.Fields("PUHINBC1")) = False Then .PUHINBC1 = rs.Fields("PUHINBC1")         ' �_���i��
            If IsNull(rs.Fields("PUCHAGC1")) = False Then .PUCHAGC1 = rs.Fields("PUCHAGC1")         ' ����ޗ�
            If IsNull(rs.Fields("KAKOUBC1")) = False Then .KAKOUBC1 = rs.Fields("KAKOUBC1")         ' ���H�敪
            If IsNull(rs.Fields("KEIDAYC1")) = False Then .KEIDAYC1 = rs.Fields("KEIDAYC1")         ' �v����t
            If IsNull(rs.Fields("SEEDC1")) = False Then .SEEDC1 = rs.Fields("SEEDC1")             ' ����
            If IsNull(rs.Fields("PUBTKBC1")) = False Then .PUBTKBC1 = rs.Fields("PUBTKBC1")         ' BOT�󋵋敪
            If IsNull(rs.Fields("JDGECC1")) = False Then .JDGECC1 = rs.Fields("JDGECC1")           ' ���躰��
            If IsNull(rs.Fields("PWTIMEC1")) = False Then .PWTIMEC1 = rs.Fields("PWTIMEC1")         ' ��ܰ����
            If IsNull(rs.Fields("ADDOPPC1")) = False Then .ADDOPPC1 = rs.Fields("ADDOPPC1")         ' �ǉ��ް�߈ʒu
            If IsNull(rs.Fields("ADDOPCC1")) = False Then .ADDOPCC1 = rs.Fields("ADDOPCC1")         ' �ǉ��ް���Ď��
            If IsNull(rs.Fields("ADDOPVC1")) = False Then .ADDOPVC1 = rs.Fields("ADDOPVC1")         ' �ǉ��ް�ߗ�
            If IsNull(rs.Fields("ADDOPNC1")) = False Then .ADDOPNC1 = rs.Fields("ADDOPNC1")         ' �ǉ��ް�ߖ�
            If IsNull(rs.Fields("TSTAFFC1")) = False Then .TSTAFFC1 = rs.Fields("TSTAFFC1")         ' �o�^�Ј�ID
            If IsNull(rs.Fields("TDAYC1")) = False Then .TDAYC1 = rs.Fields("TDAYC1")             ' �o�^���t
            If IsNull(rs.Fields("KSTAFFC1")) = False Then .KSTAFFC1 = rs.Fields("KSTAFFC1")         ' �X�V�Ј�ID
            If IsNull(rs.Fields("KDAYC1")) = False Then .KDAYC1 = rs.Fields("KDAYC1")             ' �X�V���t
            If IsNull(rs.Fields("SUMITBC1")) = False Then .SUMITBC1 = rs.Fields("SUMITBC1")         ' SUMMIT���M�׸�
            If IsNull(rs.Fields("SNDKC1")) = False Then .SNDKC1 = rs.Fields("SNDKC1")             ' ���M�׸�
            If IsNull(rs.Fields("SNDDAYC1")) = False Then .SNDDAYC1 = rs.Fields("SNDDAYC1")         ' ���M���t
            If IsNull(rs.Fields("SUIFLG")) = False Then .SUIFLG = rs.Fields("SUIFLG")         ' ����FLG
            If IsNull(rs.Fields("PUREVNUMC1")) = False Then .PUREVNUMC1 = rs.Fields("PUREVNUMC1")       '�_���i�Ԑ����ԍ������ԍ��@04/09/28 ooba
            If IsNull(rs.Fields("PUFACTORYC1")) = False Then .PUFACTORYC1 = rs.Fields("PUFACTORYC1")    '�_���i�ԍH��@04/09/28 ooba
            If IsNull(rs.Fields("PUOPEC1")) = False Then .PUOPEC1 = rs.Fields("PUOPEC1")                '�_���i�ԑ��Ə����@04/09/28 ooba
            If IsNull(rs.Fields("LENPUFRC1")) = False Then .LENPUFRC1 = rs.Fields("LENPUFRC1")      '�����ذ�����@05/07/19 ooba
            If IsNull(rs.Fields("WGHTPUFRC1")) = False Then .WGHTPUFRC1 = rs.Fields("WGHTPUFRC1")   '�����ذ�d�ʁ@05/07/19 ooba
            If IsNull(rs.Fields("WGHTCUTRHC1")) = False Then .WGHTCUTRHC1 = rs.Fields("WGHTCUTRHC1") '�ؒf��Ǖi�d�ʁ@05/07/19 ooba
''09/02/16 FAE)akiyama start
            If IsNull(rs.Fields("SUICHARGE")) = False Then .SUICHARGE = rs.Fields("SUICHARGE") '����`���[�W��
''09/02/16 FAE)akiyama end
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDC1 = FUNCTION_RETURN_SUCCESS
End Function


'��UPDATE��

'���X�V���ڂ��\���̂ɃZ�b�g���Ĉ����n��

'�T�v      :�e�[�u���uXSDC1�v���X�V���� ptrn1
'���Ұ�    :�ϐ���        ,IO  ,�^               ,����
'          :records       ,O   ,typ_XSDC1_Update ,�X�V���R�[�h
'          :sqlWhere      ,I   ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I   ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O   ,FUNCTION_RETURN  ,���o�̐���
'����      :

Public Function UpdateXSDC1(records As typ_XSDC1_Update, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
On Error GoTo proc_err

    Dim sql As String       'SQL�S��
'    Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
'    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
'    Dim i As Long
    Dim nowtime As Date
    Dim nowtime_sql As String   '�T�[�o����(SQL��)

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDC1_SQL.bas -- Function UpdateXSDC1"
    
    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾����悤�ɕύX 2003/6/4 tuku

'>>>>> .Edit��SQL(UPDATE)���ɕύX�@2009/06/16 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"

    With records
        
        ''SQL��g�ݗ��Ă�
        sql = "UPDATE XSDC1 SET" & vbLf
        
        ''�X�V���t
        sql = sql & " KDAYC1 = " & nowtime_sql & vbLf
        
        ''�����ԍ�
        If .XTALC1 <> "" And Left(.XTALC1, 1) <> vbNullChar Then
            sql = sql & ",XTALC1 = '" & .XTALC1 & "'" & vbLf
        End If

        ''�Ǘ��H������
        If .KNKTC1 <> "" And Left(.KNKTC1, 1) <> vbNullChar Then
            sql = sql & ",KNKTC1 = '" & .KNKTC1 & "'" & vbLf
        End If
        
        ''�H������
        If .WKKTC1 <> "" And Left(.WKKTC1, 1) <> vbNullChar Then
            sql = sql & ",WKKTC1 = '" & .WKKTC1 & "'" & vbLf
        End If
        
        ''�����iTOP�j
        If .LENTOC1 <> "" Then
            sql = sql & ",LENTOC1 = '" & CStr(CInt(.LENTOC1)) & "'" & vbLf
        End If
        
        ''�����i�����j
        If .LENTKC1 <> "" Then
            sql = sql & ",LENTKC1 = '" & CStr(CInt(.LENTKC1)) & "'" & vbLf
        End If
        
        ''�����iTAIL�j
        If .LENTAC1 <> "" Then
            sql = sql & ",LENTAC1 = '" & CStr(CInt(.LENTAC1)) & "'" & vbLf
        End If
        
        ''�ذ����
        If .PUFRELC1 <> "" Then
            sql = sql & ",PUFRELC1 = '" & CStr(CInt(.PUFRELC1)) & "'" & vbLf
        End If
        
        ''�������a1
        If .DIA1C1 <> "" Then
            sql = sql & ",DIA1C1 = '" & .DIA1C1 & "'" & vbLf
        End If
        
        ''�������a2
        If .DIA2C1 <> "" Then
            sql = sql & ",DIA2C1 = '" & .DIA2C1 & "'" & vbLf
        End If
        
        ''�������a3
        If .DIA3C1 <> "" Then
            sql = sql & ",DIA3C1 = '" & .DIA3C1 & "'" & vbLf
        End If
        
        ''�d�ʁiTOP�j
        If .WGHTTOC1 <> "" Then
            sql = sql & ",WGHTTOC1 = '" & CStr(CLng(.WGHTTOC1)) & "'" & vbLf
        End If
        
        ''�d�ʁi�����j
        If .WGHTTKC1 <> "" Then
            sql = sql & ",WGHTTKC1 = '" & CStr(CLng(.WGHTTKC1)) & "'" & vbLf
        End If
        
        ''�d�ʁiTAIL�j
        If .WGHTTAC1 <> "" Then
            sql = sql & ",WGHTTAC1 = '" & CStr(CLng(.WGHTTAC1)) & "'" & vbLf
        End If
        
        ''�d�ʁi�ذ�����j
        If .WGHTFRC1 <> "" Then
            sql = sql & ",WGHTFRC1 = '" & CStr(CLng(.WGHTFRC1)) & "'" & vbLf
        End If
        
        ''į�߶�ďd��
        If .PUTCUTWC1 <> "" Then
            sql = sql & ",PUTCUTWC1 = '" & CStr(CLng(.PUTCUTWC1)) & "'" & vbLf
        End If
        
        ''���グ�d��
        If .PUWC1 <> "" Then
            sql = sql & ",PUWC1 = '" & CStr(CLng(.PUWC1)) & "'" & vbLf
        End If
        
        ''�_���i��
        If .PUHINBC1 <> "" And Left(.PUHINBC1, 1) <> vbNullChar Then
            sql = sql & ",PUHINBC1 = '" & .PUHINBC1 & "'" & vbLf
        End If
        
        ''����ޗ�
        If .PUCHAGC1 <> "" Then
            sql = sql & ",PUCHAGC1 = '" & CStr(CLng(.PUCHAGC1)) & "'" & vbLf
        End If
        
        ''���H�敪
        If .KAKOUBC1 <> "" And Left(.KAKOUBC1, 1) <> vbNullChar Then
            sql = sql & ",KAKOUBC1 = '" & .KAKOUBC1 & "'" & vbLf
        End If
        
        ''�v����t
        If .KEIDAYC1 <> "" Then
            sql = sql & ",KEIDAYC1 = TO_DATE('" & Format$(CDate(.KEIDAYC1), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''����
        If .SEEDC1 <> "" And Left(.SEEDC1, 1) <> vbNullChar Then
            sql = sql & ",SEEDC1 = '" & .SEEDC1 & "'" & vbLf
        End If
        
        ''BOT�󋵋敪
        If .PUBTKBC1 <> "" And Left(.PUBTKBC1, 1) <> vbNullChar Then
            sql = sql & ",PUBTKBC1 = '" & .PUBTKBC1 & "'" & vbLf
        End If
        
        ''���躰��
        If .JDGECC1 <> "" And Left(.JDGECC1, 1) <> vbNullChar Then
            sql = sql & ",JDGECC1 = '" & .JDGECC1 & "'" & vbLf
        End If
        
        ''��ܰ����
        If .PWTIMEC1 <> "" Then
            sql = sql & ",PWTIMEC1 = '" & .PWTIMEC1 & "'" & vbLf
        End If
        
        ''�ǉ��ް�߈ʒu
        If .ADDOPPC1 <> "" Then
            sql = sql & ",ADDOPPC1 = '" & CStr(CInt(.ADDOPPC1)) & "'" & vbLf
        End If
        
        ''�ǉ��ް���Ď��
        If .ADDOPCC1 <> "" And Left(.ADDOPCC1, 1) <> vbNullChar Then
            sql = sql & ",ADDOPCC1 = '" & .ADDOPCC1 & "'" & vbLf
        End If
        
        ''�ǉ��ް�ߗ�
        If .ADDOPVC1 <> "" Then
            sql = sql & ",ADDOPVC1 = '" & CStr(CLng(.ADDOPVC1)) & "'" & vbLf
        End If
        
        ''�ǉ��ް�ߖ�
        If .ADDOPNC1 <> "" And Left(.ADDOPNC1, 1) <> vbNullChar Then
            sql = sql & ",ADDOPNC1 = '" & .ADDOPNC1 & "'" & vbLf
        End If
        
        ''�o�^�Ј�ID
        If .TSTAFFC1 <> "" And Left(.TSTAFFC1, 1) <> vbNullChar Then
            sql = sql & ",TSTAFFC1 = '" & .TSTAFFC1 & "'" & vbLf
        End If
        
        ''�o�^���t
        If .TDAYC1 <> "" Then
            sql = sql & ",TDAYC1 = TO_DATE('" & Format$(CDate(.TDAYC1), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''�X�V�Ј�ID
        If .KSTAFFC1 <> "" And Left(.KSTAFFC1, 1) <> vbNullChar Then
            sql = sql & ",KSTAFFC1 = '" & .KSTAFFC1 & "'" & vbLf
        End If
        
        ''SUMMIT���M�׸�
        If .SUMITBC1 <> "" And Left(.SUMITBC1, 1) <> vbNullChar Then
            sql = sql & ",SUMITBC1 = '" & .SUMITBC1 & "'" & vbLf
        End If
        
        ''���M�׸�
        If .SNDKC1 <> "" And Left(.SNDKC1, 1) <> vbNullChar Then
            sql = sql & ",SNDKC1 = '" & .SNDKC1 & "'" & vbLf
        End If
        
        ''���M���t
        If .SNDDAYC1 <> "" Then
            sql = sql & ",SNDDAYC1 = TO_DATE('" & Format$(CDate(.SNDDAYC1), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''����FLG
        If .SUIFLG <> "" Then
            sql = sql & ",SUIFLG = '" & .SUIFLG & "'" & vbLf
        End If
        
        ''�_���i�Ԑ����ԍ������ԍ�
        If .PUREVNUMC1 <> "" Then
            sql = sql & ",PUREVNUMC1 = '" & .PUREVNUMC1 & "'" & vbLf
        End If
        
        ''�_���i�ԍH��
        If .PUFACTORYC1 <> "" Then
            sql = sql & ",PUFACTORYC1 = '" & .PUFACTORYC1 & "'" & vbLf
        End If
        
        ''�_���i�ԑ��Ə���
        If .PUOPEC1 <> "" Then
            sql = sql & ",PUOPEC1 = '" & .PUOPEC1 & "'" & vbLf
        End If
        
        ''�����ذ����
        If .LENPUFRC1 <> "" Then
            sql = sql & ",LENPUFRC1 = '" & CStr(CInt(.LENPUFRC1)) & "'" & vbLf
        End If
        
        ''�����ذ�d��
        If .WGHTPUFRC1 <> "" Then
            sql = sql & ",WGHTPUFRC1 = '" & CStr(CLng(.WGHTPUFRC1)) & "'" & vbLf
        End If
        
        ''�ؒf��Ǖi�d��
        If .WGHTCUTRHC1 <> "" Then
            sql = sql & ",WGHTCUTRHC1 = '" & CStr(CLng(.WGHTCUTRHC1)) & "'" & vbLf
        End If
        
        ''C-OSF3����ID
        If .JDGEIDC1 <> "" Then
            sql = sql & ",JDGEIDC1 = '" & .JDGEIDC1 & "'" & vbLf
        End If

        sql = sql & " " & sqlWhere & vbLf
    
        'SQL�����s
        recCnt = OraDB.ExecuteSQL(sql)
        
        '�Ԃ�l��1�ȊO�̓G���[
        If recCnt < 0 Then
            GoTo proc_err
        ElseIf recCnt = 0 Then
            '0���X�V�c�G���[(�����ʂ�)
            UpdateXSDC1 = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf recCnt > 1 Then
            '�������X�V�c�G���[(�����͕���SELECT�����ŏ��̈ꌏ�̂ݍX�V)
            UpdateXSDC1 = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        
    End With
'<<<<< .Edit��SQL(UPDATE)���ɕύX�@2009/06/16 SETsw kubota ------------------

    UpdateXSDC1 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    UpdateXSDC1 = FUNCTION_RETURN_FAILURE
    Debug.Print "==== ERROR SQL ===="
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :����d�ʂ��擾����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_sXtal      ,I  ,String           ,�����ԍ�
'      �@�@:�߂�l       ,O  ,Long           �@,������
'          :�쐬�ҁ@04/09/30 ooba
Public Function GetPutWeight(p_sXtal As String) As Long

    Dim sSQL As String
    Dim rs As OraDynaset
    
    If Left(p_sXtal, 1) = vbNullChar Then
        GetPutWeight = 0
        Exit Function
    End If
    
    sSQL = "SELECT PUWC1 FROM XSDC1 WHERE XTALC1 = '" & p_sXtal & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        GetPutWeight = 0
    Else
        If IsNull(rs.Fields("PUWC1")) = False Then GetPutWeight = CLng(rs.Fields("PUWC1")) Else GetPutWeight = 0
    End If
    
End Function
