Attribute VB_Name = "s_XSDC3_SQL"
'�H������ (XSDC3) �����֐�


'***�e�[�u���uXSDC3�v�ւ̃f�[�^�A�N�Z�X�֐�***
'������ ���Ұ��ɒl��Ă��鎞�A�܂��S�ď��������邱��

Option Explicit

'���H������
Public Type typ_XSDC3
    CRYNUMC3 As String * 12      ' ��ۯ�ID������ԍ�
    INPOSC3 As Integer           ' �������J�n�ʒu
    KCNTC3 As Integer            ' �H���A��
    HINBC3 As String * 8         ' �i��
    REVNUMC3 As Integer          ' ���i�ԍ������ԍ�
    FACTORYC3 As String * 1       ' �H��
    OPEC3 As String * 1          ' ���Ə���
    LENC3 As Integer             ' ����
    XTALC3 As String * 12        ' �����ԍ�
    SXLIDC3 As String * 13        ' SXLID
    KNKTC3 As String * 5         ' �Ǘ��H��
    WKKTC3 As String * 5         ' �H��
    WKKBC3 As String * 2         ' ��Ƌ敪
    MACOC3 As Integer            ' ������
    MODKBC3 As String * 1        ' �ԍ��敪
    SUMKBC3 As String * 1        ' �W�v�敪
    FRKNKTC3 As String * 5       ' (���)�Ǘ��H��
    FRWKKTC3 As String * 5       ' (���)�H��
    FRWKKBC3 As String * 2       ' (���)��Ƌ敪
    FRMACOC3 As Integer          ' (���)������
    TOWNKTC3 As String * 5       ' (���o)�Ǘ��H��
    TOWKKTC3 As String * 5       ' (���o)�H��
    TOMACOC3 As Integer          ' (���o)������
    FRLC3 As Integer             ' �������
'    FRWC3 As Integer             ' ����d�ʁ��f�[�^�^��Long�ɕύX 2003/09/24 �I�[�o�[�t���[���邽��
    FRWC3 As Long                ' ����d��
    FRMC3 As Integer             ' �������
    FULC3 As Integer             ' �s�ǒ���
    FUWC3 As Integer             ' �s�Ǐd��
    FUMC3 As Integer             ' �s�ǖ���
    LOSWC3 As Integer            ' ���X����
    LOSLC3 As Integer            ' ���X�d��
    LOSMC3 As Integer            ' ���X����
    TOLC3 As Integer             ' ���o����
'    TOWC3 As Integer            ' ���o�d�� ���f�[�^�^��Long�ɕύX 2003/09/22 �I�[�o�[�t���[���邽��
    TOWC3 As Long                ' ���o�d��
    TOMC3 As Integer             ' ���o����
    SUMITLC3 As Integer          ' SUMIT����
'    SUMITWC3 As Integer          ' SUMIT�d�ʁ��f�[�^�^��Long�ɕύX 2003/09/23 �I�[�o�[�t���[���邽��
    SUMITWC3 As Long             ' SUMIT�d��
    SUMITMC3 As Integer          ' SUMIT����
    MOTHINC3 As String * 12      ' �U�֕i��(��)
    XTWORKC3 As String * 2       ' �����H��
    WFWORKC3 As String * 2       ' ���ʐ���
    STATIMEC3 As Date            ' �������ԊJ�n
    STOTIMEC3 As Date            ' �������ԏI��
    ETIMEC3 As Date              ' ���ю���
    HOLDCC3 As String * 3        ' �z�[���h�R�[�h
    HOLDBC3 As String * 1        ' �z�[���h�敪
    LDFRCC3 As String * 3        ' �i���R�[�h
    LDFRBC3 As String * 1        ' �i���敪
    TSTAFFC3 As String * 8       ' �o�^�Ј�ID
    TDAYC3 As Date               ' �o�^���t
    KSTAFFC3 As String * 8       ' �X�V�Ј�ID
    KDAYC3 As Date               ' �X�V���t
    SUMITBC3 As String * 1       ' SUMIT���M�t���O
    SNDKC3 As String * 1         ' ���M�t���O
    SNDDAYC3 As Date             ' �o�^���t
    'add start 2003/03/25 hitec)matsumoto ----
    SUMDAYC3 As Date             ' SUMCO����
    PAYCLASSC3 As String * 1     ' �]����H��t���O
    'add end 2003/03/25 hitec)matsumoto ----
    '2003.06.11 (SPK)Y.Katabami tuika
    CUTCNTC3 As String * 1       ' �V�K�^�Đ؋敪 '1':�Đ�
    HINBFLGC3 As String * 1      ' ��\�i�ԃt���O '1'�F��\�i��
    '2005/11
    RPCRYNUMC3 As String
''>>>>> �p���[ON���Ԓǉ��Ή� SETsw H.Iwamoto 2005/11/28
    PROTMC3     As Integer       ' �p���[ON����(��)
    PROMNC3     As Integer       ' �p���[ON����(��)
    PROTM2C3    As Integer       ' (�݌v)�p���[ON����(��)
    PROMN2C3    As Integer       ' (�݌v)�p���[ON����(��)
''<<<<< �p���[ON���Ԓǉ��Ή� SETsw H.Iwamoto 2005/11/28
    PLANTCATC3 As String * 2     ' ����@2007/08/15 SPK Tsutsumi
End Type

'�X�V�p
Public Type typ_XSDC3_Update
    CRYNUMC3 As String           ' ��ۯ�ID������ԍ�
    INPOSC3 As String            ' �������J�n�ʒu
    KCNTC3 As String             ' �H���A��
    HINBC3 As String             ' �i��
    REVNUMC3 As String           ' ���i�ԍ������ԍ�
    FACTORYC3 As String          ' �H��
    OPEC3 As String              ' ���Ə���
    LENC3 As String              ' ����
    XTALC3 As String             ' �����ԍ�
    SXLIDC3 As String             ' SXLID
    KNKTC3 As String             ' �Ǘ��H��
    WKKTC3 As String             ' �H��
    WKKBC3 As String             ' ��Ƌ敪
    MACOC3 As String             ' ������
    MODKBC3 As String            ' �ԍ��敪
    SUMKBC3 As String            ' �W�v�敪
    FRKNKTC3 As String           ' (���)�Ǘ��H��
    FRWKKTC3 As String           ' (���)�H��
    FRWKKBC3 As String           ' (���)��Ƌ敪
    FRMACOC3 As String           ' (���)������
    TOWNKTC3 As String           ' (���o)�Ǘ��H��
    TOWKKTC3 As String           ' (���o)�H��
    TOMACOC3 As String           ' (���o)������
    FRLC3 As String              ' �������
    FRWC3 As String              ' ����d��
    FRMC3 As String              ' �������
    FULC3 As String              ' �s�ǒ���
    FUWC3 As String              ' �s�Ǐd��
    FUMC3 As String              ' �s�ǖ���
    LOSWC3 As String             ' ���X����
    LOSLC3 As String             ' ���X�d��
    LOSMC3 As String             ' ���X����
    TOLC3 As String              ' ���o����
    TOWC3 As String              ' ���o�d��
    TOMC3 As String              ' ���o����
    SUMITLC3 As String           ' SUMIT����
    SUMITWC3 As String           ' SUMIT�d��
    SUMITMC3 As String           ' SUMIT����
    MOTHINC3 As String           ' �U�֕i��(��)
    XTWORKC3 As String           ' �����H��
    WFWORKC3 As String           ' ���ʐ���
    STATIMEC3 As Date            ' �������ԊJ�n
    STOTIMEC3 As Date            ' �������ԏI��
    ETIMEC3 As Date              ' ���ю���
    HOLDCC3 As String            ' �z�[���h�R�[�h
    HOLDBC3 As String            ' �z�[���h�敪
    LDFRCC3 As String            ' �i���R�[�h
    LDFRBC3 As String            ' �i���敪
    TSTAFFC3 As String           ' �o�^�Ј�ID
    TDAYC3 As Date               ' �o�^���t
    KSTAFFC3 As String           ' �X�V�Ј�ID
    KDAYC3 As Date               ' �X�V���t
    SUMITBC3 As String           ' SUMIT���M�t���O
    SNDKC3 As String             ' ���M�t���O
    SNDDAYC3 As Date             ' �o�^���t
    MODMACOC3 As String * 2       ' �ԍ�������
    KAKUCC3 As String * 5         ' �m��R�[�h
    'add start 2003/03/25 hitec)matsumoto ----
    SUMDAYC3 As Date             ' SUMCO����
    PAYCLASSC3 As String         ' �]����H��t���O
    'add end 2003/03/25 hitec)matsumoto ----
    '2003.06.11 (SPK)Y.Katabami tuika
    CUTCNTC3 As String * 1       ' �V�K�^�Đ؋敪 '1':�Đ�
    HINBFLGC3 As String * 1      ' ��\�i�ԃt���O '1'�F��\�i��
    '2005/11
    RPCRYNUMC3 As String
''>>>>> �p���[ON���Ԓǉ��Ή� SETsw H.Iwamoto 2005/11/28
    PROTMC3     As Integer       ' �p���[ON����(��)
    PROMNC3     As Integer       ' �p���[ON����(��)
    PROTM2C3    As Integer       ' (�݌v)�p���[ON����(��)
    PROMN2C3    As Integer       ' (�݌v)�p���[ON����(��)
''<<<<< �p���[ON���Ԓǉ��Ή� SETsw H.Iwamoto 2005/11/28
    PLANTCATC3 As String * 2     ' ����@2007/08/15 SPK Tsutsumi
End Type

'��SELECT��

'�T�v      :�e�[�u���uXSDC3�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO   ,�^               ,����
'          :records()     ,O    ,typ_XSDC3     ,���o���R�[�h
'          :sqlWhere      ,I    ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I    ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :

Public Function DBDRV_GetXSDC3(records() As typ_XSDC3, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim sql As String       'SQL�S��
    Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long


    ''SQL��g�ݗ��Ă�
    sqlBase = "Select * From XSDC3"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDC3 = FUNCTION_RETURN_FAILURE
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
            If IsNull(rs.Fields("CRYNUMC3")) = False Then .CRYNUMC3 = rs.Fields("CRYNUMC3")
            If IsNull(rs.Fields("INPOSC3")) = False Then .INPOSC3 = rs.Fields("INPOSC3")
            If IsNull(rs.Fields("KCNTC3")) = False Then .KCNTC3 = rs.Fields("KCNTC3")
            If IsNull(rs.Fields("HINBC3")) = False Then .HINBC3 = rs.Fields("HINBC3")
            If IsNull(rs.Fields("REVNUMC3")) = False Then .REVNUMC3 = rs.Fields("REVNUMC3")
            If IsNull(rs.Fields("FACTORYC3")) = False Then .FACTORYC3 = rs.Fields("FACTORYC3")
            If IsNull(rs.Fields("OPEC3")) = False Then .OPEC3 = rs.Fields("OPEC3")
            If IsNull(rs.Fields("LENC3")) = False Then .LENC3 = rs.Fields("LENC3")
            If IsNull(rs.Fields("XTALC3")) = False Then .XTALC3 = rs.Fields("XTALC3")
            If IsNull(rs.Fields("SXLIDC3")) = False Then .SXLIDC3 = rs.Fields("SXLIDC3")
            If IsNull(rs.Fields("KNKTC3")) = False Then .KNKTC3 = rs.Fields("KNKTC3")
            If IsNull(rs.Fields("WKKTC3")) = False Then .WKKTC3 = rs.Fields("WKKTC3")
            If IsNull(rs.Fields("WKKBC3")) = False Then .WKKBC3 = rs.Fields("WKKBC3")
            If IsNull(rs.Fields("MACOC3")) = False Then .MACOC3 = rs.Fields("MACOC3")
            If IsNull(rs.Fields("MODKBC3")) = False Then .MODKBC3 = rs.Fields("MODKBC3")
            If IsNull(rs.Fields("SUMKBC3")) = False Then .SUMKBC3 = rs.Fields("SUMKBC3")
            If IsNull(rs.Fields("FRKNKTC3")) = False Then .FRKNKTC3 = rs.Fields("FRKNKTC3")
            If IsNull(rs.Fields("FRWKKTC3")) = False Then .FRWKKTC3 = rs.Fields("FRWKKTC3")
            If IsNull(rs.Fields("FRWKKBC3")) = False Then .FRWKKBC3 = rs.Fields("FRWKKBC3")
            If IsNull(rs.Fields("FRMACOC3")) = False Then .FRMACOC3 = rs.Fields("FRMACOC3")
            If IsNull(rs.Fields("TOWNKTC3")) = False Then .TOWNKTC3 = rs.Fields("TOWNKTC3")
            If IsNull(rs.Fields("TOWKKTC3")) = False Then .TOWKKTC3 = rs.Fields("TOWKKTC3")
            If IsNull(rs.Fields("TOMACOC3")) = False Then .TOMACOC3 = rs.Fields("TOMACOC3")
            If IsNull(rs.Fields("FRLC3")) = False Then .FRLC3 = rs.Fields("FRLC3")
            If IsNull(rs.Fields("FRWC3")) = False Then .FRWC3 = rs.Fields("FRWC3")
            If IsNull(rs.Fields("FRMC3")) = False Then .FRMC3 = rs.Fields("FRMC3")
            If IsNull(rs.Fields("FULC3")) = False Then .FULC3 = rs.Fields("FULC3")
            If IsNull(rs.Fields("FUWC3")) = False Then .FUWC3 = rs.Fields("FUWC3")
            If IsNull(rs.Fields("FUMC3")) = False Then .FUMC3 = rs.Fields("FUMC3")
            If IsNull(rs.Fields("LOSWC3")) = False Then .LOSWC3 = rs.Fields("LOSWC3")
            If IsNull(rs.Fields("LOSLC3")) = False Then .LOSLC3 = rs.Fields("LOSLC3")
            If IsNull(rs.Fields("LOSMC3")) = False Then .LOSMC3 = rs.Fields("LOSMC3")
            If IsNull(rs.Fields("TOLC3")) = False Then .TOLC3 = rs.Fields("TOLC3")
            If IsNull(rs.Fields("TOWC3")) = False Then .TOWC3 = rs.Fields("TOWC3")
            If IsNull(rs.Fields("TOMC3")) = False Then .TOMC3 = rs.Fields("TOMC3")
            If IsNull(rs.Fields("SUMITLC3")) = False Then .SUMITLC3 = rs.Fields("SUMITLC3")
            If IsNull(rs.Fields("SUMITWC3")) = False Then .SUMITWC3 = rs.Fields("SUMITWC3")
            If IsNull(rs.Fields("SUMITMC3")) = False Then .SUMITMC3 = rs.Fields("SUMITMC3")
            If IsNull(rs.Fields("MOTHINC3")) = False Then .MOTHINC3 = rs.Fields("MOTHINC3")
            If IsNull(rs.Fields("XTWORKC3")) = False Then .XTWORKC3 = rs.Fields("XTWORKC3")
            If IsNull(rs.Fields("WFWORKC3")) = False Then .WFWORKC3 = rs.Fields("WFWORKC3")
            If IsNull(rs.Fields("STATIMEC3")) = False Then .STATIMEC3 = rs.Fields("STATIMEC3")
            If IsNull(rs.Fields("STOTIMEC3")) = False Then .STOTIMEC3 = rs.Fields("STOTIMEC3")
            If IsNull(rs.Fields("ETIMEC3")) = False Then .ETIMEC3 = rs.Fields("ETIMEC3")
            If IsNull(rs.Fields("HOLDCC3")) = False Then .HOLDCC3 = rs.Fields("HOLDCC3")
            If IsNull(rs.Fields("HOLDBC3")) = False Then .HOLDBC3 = rs.Fields("HOLDBC3")
            If IsNull(rs.Fields("LDFRCC3")) = False Then .LDFRCC3 = rs.Fields("LDFRCC3")
            If IsNull(rs.Fields("LDFRBC3")) = False Then .LDFRBC3 = rs.Fields("LDFRBC3")
            If IsNull(rs.Fields("TSTAFFC3")) = False Then .TSTAFFC3 = rs.Fields("TSTAFFC3")
            If IsNull(rs.Fields("TDAYC3")) = False Then .TDAYC3 = rs.Fields("TDAYC3")
            If IsNull(rs.Fields("KSTAFFC3")) = False Then .KSTAFFC3 = rs.Fields("KSTAFFC3")
            If IsNull(rs.Fields("KDAYC3")) = False Then .KDAYC3 = rs.Fields("KDAYC3")
            If IsNull(rs.Fields("SUMITBC3")) = False Then .SUMITBC3 = rs.Fields("SUMITBC3")
            If IsNull(rs.Fields("SNDKC3")) = False Then .SNDKC3 = rs.Fields("SNDKC3")
            If IsNull(rs.Fields("SNDDAYC3")) = False Then .SNDDAYC3 = rs.Fields("SNDDAYC3")
            'add start 2003/03/25 hitec)matsumoto ------
            If IsNull(rs.Fields("SUMDAYC3")) = False Then .SUMDAYC3 = rs.Fields("SUMDAYC3") 'SUMCO����
            If IsNull(rs.Fields("PAYCLASSC3")) = False Then .PAYCLASSC3 = rs.Fields("PAYCLASSC3") '�]����t���O
           'add end 2003/03/25 hitec)matsumoto ------
            '2003.06.11 (SPK)Y.katabami tuika
            If IsNull(rs.Fields("CUTCNTC3")) = False Then .CUTCNTC3 = rs.Fields("CUTCNTC3")
            If IsNull(rs.Fields("HINBFLGC3")) = False Then .HINBFLGC3 = rs.Fields("HINBFLGC3")
            '2005/11
            If IsNull(rs.Fields("RPCRYNUMC3")) = False Then .RPCRYNUMC3 = rs.Fields("RPCRYNUMC3")
            If IsNull(rs.Fields("PLANTCATC3")) = False Then .PLANTCATC3 = rs.Fields("PLANTCATC3")   ' 2007/09/04 SPK Tsutsumi Add
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDC3 = FUNCTION_RETURN_SUCCESS
End Function


'��INSERT��  NULL�̏ꍇ�Achar�Ȃ�X�y�[�X�ANumber�Ȃ�NULL������

'�T�v      :�e�[�u���uXSDC3�v�Ƀ��R�[�h��}������
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pXSDC3 �@�@  ,I  ,typ_XSDC3_Update   ,XSDC3�X�V�p�ް�
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function CreateXSDC3(pXSDC3 As typ_XSDC3_Update, sErrMsg As String) As FUNCTION_RETURN


    Dim sql As String
    Dim sql2 As String
    Dim sDbName As String
'    Dim rs As OraDynaset    'RecordSet
    Dim rs2 As OraDynaset    'RecordSet
'    Dim recCnt As Long      '���R�[�h��
    Dim nowtime         As Date
    Dim nowtime_sql     As String   '�T�[�o����(SQL��)
    Dim justNowTime     As Date
    Dim justNowTime_sql As String   '�T�[�o����(SQL��)

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDC3_SQL.bas -- Function CreateXSDC3"
    sErrMsg = ""
    sDbName = "XSDC3"
    'nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    'justNowTime = Format(Time, "hh:mm:ss")
    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾����悤�ɕύX 2003/6/4 tuku
    justNowTime = Format(nowtime, "hh:mm:ss")
    
'>>>>> .AddNew��SQL(INSERT)���ɕύX�@2009/06/29 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    justNowTime_sql = "TO_DATE('" & Format$(justNowTime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With pXSDC3
        
        sql = "INSERT INTO XSDC3 ("
        sql = sql & " CRYNUMC3"         ' 1:��ۯ�ID�E�����ԍ�
        sql = sql & ",INPOSC3"          ' 2:�������J�n�ʒu
        sql = sql & ",KCNTC3"           ' 3:�H���A��
        sql = sql & ",HINBC3"           ' 4:�i��
        sql = sql & ",REVNUMC3"         ' 5:���i�ԍ������ԍ�
        sql = sql & ",FACTORYC3"        ' 6:�H��
        sql = sql & ",OPEC3"            ' 7:���Ə���
        sql = sql & ",LENC3"            ' 8:����
        sql = sql & ",XTALC3"           ' 9:�����ԍ�
        sql = sql & ",SXLIDC3"          '10:SXLID
        sql = sql & ",KNKTC3"           '11:�Ǘ��H��
        sql = sql & ",WKKTC3"           '12:�H��
        sql = sql & ",WKKBC3"           '13:��Ƌ敪
        sql = sql & ",MACOC3"           '14:������
        sql = sql & ",MODKBC3"          '15:�ԍ��敪
        sql = sql & ",SUMKBC3"          '16:�W�v�敪
        sql = sql & ",FRKNKTC3"         '17:(���)�Ǘ��H��
        sql = sql & ",FRWKKTC3"         '18:(���)�H��
        sql = sql & ",FRWKKBC3"         '19:(���)��Ƌ敪
        sql = sql & ",FRMACOC3"         '20:(���)������
        sql = sql & ",TOWNKTC3"         '21:(���o)�Ǘ��H��
        sql = sql & ",TOWKKTC3"         '22:(���o)�H��
        sql = sql & ",TOMACOC3"         '23:(���o)������
        sql = sql & ",FRLC3"            '24:�������
        sql = sql & ",FRWC3"            '25:����d��
        sql = sql & ",FRMC3"            '26:�������
        sql = sql & ",FULC3"            '27:�s�ǒ���
        sql = sql & ",FUWC3"            '28:�s�Ǐd��
        sql = sql & ",FUMC3"            '29:�s�ǖ���
        sql = sql & ",LOSWC3"           '30:���X����
        sql = sql & ",LOSLC3"           '31:���X�d��
        sql = sql & ",LOSMC3"           '32:���X����
        sql = sql & ",TOLC3"            '33:���o����
        sql = sql & ",TOWC3"            '34:���o�d��
        sql = sql & ",TOMC3"            '35:���o����
        sql = sql & ",SUMITLC3"         '36:SUMMIT����
        sql = sql & ",SUMITWC3"         '37:SUMMIT�d��
        sql = sql & ",SUMITMC3"         '38:SUMMIT����
        sql = sql & ",MOTHINC3"         '39:�U�֕i��(��)
        sql = sql & ",XTWORKC3"         '40:�����H��
        sql = sql & ",WFWORKC3"         '41:���ʐ���
        sql = sql & ",STATIMEC3"        '42:�������ԊJ�n
        sql = sql & ",STOTIMEC3"        '43:�������ԏI��
        sql = sql & ",ETIMEC3"          '44:���ю���
        sql = sql & ",HOLDCC3"          '45:ΰ��޺���
        sql = sql & ",HOLDBC3"          '46:ΰ��ދ敪
        sql = sql & ",LDFRCC3"          '47:�i������
        sql = sql & ",LDFRBC3"          '48:�i���敪
        sql = sql & ",TSTAFFC3"         '49:�o�^�Ј�ID
        sql = sql & ",TDAYC3"           '50:�o�^���t
        sql = sql & ",KSTAFFC3"         '51:�X�V�Ј�ID
        sql = sql & ",KDAYC3"           '52:�X�V���t
        sql = sql & ",SUMITBC3"         '53:SUMMIT���M�׸�
        sql = sql & ",SNDKC3"           '54:���M�׸�
        sql = sql & ",SNDDAYC3"         '55:���M���t
        sql = sql & ",SUMDAYC3"         '56:SUMCO����
        sql = sql & ",PAYCLASSC3"       '57:���o�敪
        sql = sql & ",CUTCNTC3"         '58:�V�K�^�Đ؋敪
        sql = sql & ",HINBFLGC3"        '59:��\�i�ԃt���O
        sql = sql & ",RPCRYNUMC3"       '60:�e��ۯ�ID
        sql = sql & ",PROTMC3"          '61:�p���[ON����(��)
        sql = sql & ",PROMNC3"          '62:�p���[ON����(��)
        sql = sql & ",PROTM2C3"         '63:(�݌v)�p���[ON����(��)
        sql = sql & ",PROMN2C3"         '64:(�݌v)�p���[ON����(��)
        sql = sql & ",PLANTCATC3"       '65:����
        sql = sql & ")"
        sql = sql & "VALUES (" & vbLf
        
        ' 1:��ۯ�ID�E�����ԍ�
        If .CRYNUMC3 <> "" Then
            sql = sql & " '" & .CRYNUMC3 & "'" & vbLf
        Else
            sql = sql & " '" & Space(12) & "'" & vbLf
        End If

        ' 2:�������J�n�ʒu
        If .INPOSC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.INPOSC3)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 3:�H���A��
        If .KCNTC3 <> "" Then
            sql = sql & ",'" & .KCNTC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 4:�i��
        If .HINBC3 <> "" Then
            sql = sql & ",'" & .HINBC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        ' 5:���i�ԍ������ԍ�
        If .REVNUMC3 <> "" Then
            sql = sql & ",'" & .REVNUMC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 6:�H��
        If .FACTORYC3 <> "" Then
            sql = sql & ",'" & .FACTORYC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 7:���Ə���
        If .OPEC3 <> "" Then
            sql = sql & ",'" & .OPEC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 8:����
        If .LENC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.LENC3)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 9:�����ԍ�
        If .XTALC3 <> "" Then
            sql = sql & ",'" & .XTALC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '10:SXLID
        If .SXLIDC3 <> "" And Left(.SXLIDC3, 1) <> vbNullChar Then
            sql = sql & ",'" & .SXLIDC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(13) & "'" & vbLf
        End If

        '11:�Ǘ��H��
        If .KNKTC3 <> "" Then
            sql = sql & ",'" & .KNKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '12:�H��
        If .WKKTC3 <> "" Then
            sql = sql & ",'" & .WKKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '13:��Ƌ敪
        If .WKKBC3 <> "" Then
            sql = sql & ",'" & .WKKBC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '14:������
        If .MACOC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.MACOC3)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '15:�ԍ��敪
        If .MODKBC3 <> "" Then
            sql = sql & ",'" & .MODKBC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '16:�W�v�敪
        If .SUMKBC3 <> "" Then
            sql = sql & ",'" & .SUMKBC3 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '17:(���)�Ǘ��H��
        If .FRKNKTC3 <> "" Then
            sql = sql & ",'" & .FRKNKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '18:(���)�H��
        If .FRWKKTC3 <> "" Then
            sql = sql & ",'" & .FRWKKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If
        
        '19:(���)��Ƌ敪
        If .FRWKKBC3 <> "" Then
            sql = sql & ",'" & .FRWKKBC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '20:(���)������
        If .FRMACOC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.FRMACOC3)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '21:(���o)�Ǘ��H��
        If .TOWNKTC3 <> "" Then
            sql = sql & ",'" & .TOWNKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '22:(���o)�H��
        If .TOWKKTC3 <> "" Then
            sql = sql & ",'" & .TOWKKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '23:(���o)������
        If .TOMACOC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.TOMACOC3)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        'ں��ނ̑O��̒l���擾
        sql2 = "SELECT TOLC3, TOWC3, TOMC3 FROM XSDC3 WHERE CRYNUMC3 = '" & .CRYNUMC3
        sql2 = sql2 & "' AND INPOSC3 = " & .INPOSC3
        sql2 = sql2 & " AND KCNTC3 = (SELECT MAX(KCNTC3) FROM XSDC3 WHERE CRYNUMC3 = '" & .CRYNUMC3
        sql2 = sql2 & "' AND MODKBC3 != '1"   ' �ԏ������R�[�h�ȊO
        sql2 = sql2 & "' AND INPOSC3 = " & .INPOSC3 & ")"
        Set rs2 = OraDB.DBCreateDynaset(sql2, ORADYN_DEFAULT)

        '24:�������
        If .FRLC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.FRLC3)) & "'" & vbLf
        Else
            If rs2.RecordCount = 0 Or .WKKTC3 = "CC705" Then
                sql = sql & ",0" & vbLf
            Else
                If Not IsNull(rs2.Fields("TOLC3")) Then
                    .FRLC3 = rs2.Fields("TOLC3")
                    sql = sql & ",'" & CStr(CInt(.FRLC3)) & "'" & vbLf
                Else
                    sql = sql & ",0" & vbLf
                End If
            End If
        End If

        '25:����d��
        If .FRWC3 <> "" Then
            sql = sql & ",'" & CStr(CLng(.FRWC3)) & "'" & vbLf
        Else
            If rs2.RecordCount = 0 Or .WKKTC3 = "CC705" Then
                sql = sql & ",0" & vbLf
            Else
                If Not IsNull(rs2.Fields("TOWC3")) Then
                    .FRWC3 = rs2.Fields("TOWC3")
                    sql = sql & ",'" & CStr(CLng(.FRWC3)) & "'" & vbLf
                Else
                    sql = sql & ",0" & vbLf
                End If
            End If
        End If

        '26:�������
        If .FRMC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.FRMC3)) & "'" & vbLf
        Else
            If rs2.RecordCount = 0 Or .WKKTC3 = "CC705" Then
                sql = sql & ",0" & vbLf
            Else
                If Not IsNull(rs2.Fields("TOMC3")) Then
                    .FRMC3 = rs2.Fields("TOMC3")
                    sql = sql & ",'" & CStr(CInt(.FRMC3)) & "'" & vbLf
                Else
                    sql = sql & ",0" & vbLf
                End If
            End If
        End If

        '27:�s�ǒ���
        If .FULC3 <> "" Then
            sql = sql & ",'" & .FULC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '28:�s�Ǐd��
        If .FUWC3 <> "" Then
            sql = sql & ",'" & .FUWC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '29:�s�ǖ���
        If .FUMC3 <> "" Then
            sql = sql & ",'" & .FUMC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '30:���X����
        If .LOSWC3 <> "" Then
            sql = sql & ",'" & .LOSWC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '31:���X����
        If .LOSLC3 <> "" Then
            sql = sql & ",'" & .LOSLC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '32:���X����
        If .LOSMC3 <> "" Then
            sql = sql & ",'" & .LOSMC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '33:���o����
        If .TOLC3 <> "" Then
            sql = sql & ",'" & .TOLC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '34:���o�d��
        If .TOWC3 <> "" Then
            sql = sql & ",'" & .TOWC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '35:���o����
        If .TOMC3 <> "" Then
            sql = sql & ",'" & .TOMC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '36:SUMMIT����
        If .SUMITLC3 <> "" Then
            sql = sql & ",'" & .SUMITLC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '37:SUMMIT�d��
        If .SUMITWC3 <> "" Then
            sql = sql & ",'" & .SUMITWC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '38:SUMMIT����
        If .SUMITMC3 <> "" Then
            sql = sql & ",'" & .SUMITMC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '39:�U�֕i��(��)
        If .MOTHINC3 <> "" Then
            sql = sql & ",'" & .MOTHINC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '40:�����H��
        sql = sql & ",'" & FACTORYCD & "'" & vbLf

        '41:���ʐ���
        If .WFWORKC3 <> "" Then
            sql = sql & ",'" & .WFWORKC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '42:�������ԊJ�n
        sql = sql & "," & justNowTime_sql & vbLf

        '43:�������ԏI��
        sql = sql & "," & justNowTime_sql & vbLf

        '44:���ю���
        sql = sql & ",0" & vbLf

        '45:ΰ��޺���
        If .HOLDCC3 <> "" Then
            sql = sql & ",'" & .HOLDCC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '46:ΰ��ދ敪
        If .HOLDBC3 <> "" Then
            sql = sql & ",'" & .HOLDBC3 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '47:�i������
        If .LDFRCC3 <> "" Then
            sql = sql & ",'" & .LDFRCC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '48:�i���敪
        If .LDFRBC3 <> "" Then
            sql = sql & ",'" & .LDFRBC3 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '49:�o�^�Ј�ID
        If XSDC3_StaffID <> "" Then
            sql = sql & ",'" & XSDC3_StaffID & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        '50:�o�^���t
        sql = sql & "," & nowtime_sql & vbLf

        '51:�X�V�Ј�ID
        If .KSTAFFC3 <> "" Then
            sql = sql & ",'" & .KSTAFFC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        '52:�X�V���t
        sql = sql & "," & nowtime_sql & vbLf

        '53:SUMMIT���M�׸�
        If .SUMITBC3 <> "" Then
            sql = sql & ",'" & .SUMITBC3 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '54:���M�׸�
        sql = sql & ",'0'" & vbLf

        '55:���M���t
        sql = sql & ",NULL" & vbLf

        '56:SUMCO����
        sql = sql & ",TO_DATE('" & Format$(CalcSumcoTime(nowtime), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf

        '57:���o�敪
        If .PAYCLASSC3 <> "" Then
            sql = sql & ",'" & .PAYCLASSC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If
        
        '58:�V�K�^�Đ؋敪
        If .CUTCNTC3 <> "" And Left(.CUTCNTC3, 1) <> vbNullChar Then
            sql = sql & ",'" & .CUTCNTC3 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '59:��\�i�ԃt���O
        If .HINBFLGC3 <> "" And Left(.HINBFLGC3, 1) <> vbNullChar Then
            sql = sql & ",'" & .HINBFLGC3 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '60:�e��ۯ�ID
        If .RPCRYNUMC3 <> "" And Left(.RPCRYNUMC3, 1) <> vbNullChar Then
            sql = sql & ",'" & .RPCRYNUMC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '61:�p���[ON����(��)
        '62:�p���[ON����(��)
        If .PROTMC3 = 0 And .PROMNC3 = 0 Then
            sql = sql & ",NULL" & vbLf
            sql = sql & ",NULL" & vbLf
        Else
            sql = sql & ",'" & CStr(.PROTMC3) & "'" & vbLf
            sql = sql & ",'" & CStr(.PROMNC3) & "'" & vbLf
        End If
        
        '63:(�݌v)�p���[ON����(��)
        '64:(�݌v)�p���[ON����(��)
        If .PROTM2C3 = 0 And .PROMN2C3 = 0 Then
            sql = sql & ",NULL" & vbLf
            sql = sql & ",NULL" & vbLf
        Else
            sql = sql & ",'" & CStr(.PROTM2C3) & "'" & vbLf
            sql = sql & ",'" & CStr(.PROMN2C3) & "'" & vbLf
        End If

        '65:����
        If .PLANTCATC3 <> "" And Left(.PLANTCATC3, 1) <> vbNullChar Then
            sql = sql & ",'" & .PLANTCATC3 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If
        
        sql = sql & ")" & vbLf
    
        'SQL�����s
        If OraDB.ExecuteSQL(sql) < 1 Then
            GoTo proc_err
        End If
        
    End With
'<<<<< .AddNew��SQL(INSERT)���ɕύX�@2009/06/29 SETsw kubota ------------------

    CreateXSDC3 = FUNCTION_RETURN_SUCCESS

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
    CreateXSDC3 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
