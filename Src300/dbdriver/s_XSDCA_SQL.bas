Attribute VB_Name = "s_XSDCA_SQL"
'��������(�i��) (XSDCA) �����֐�


'***�e�[�u���uXSDCA�v�ւ̃f�[�^�A�N�Z�X�֐�***
'������ ���Ұ��ɒl��Ă��鎞�A�܂��S�ď��������邱��

Option Explicit

'����������(�i��)
Public Type typ_XSDCA
    CRYNUMCA As String * 12      ' ��ۯ�ID������ԍ�
    HINBCA As String * 8
    INPOSCA As Integer         ' �������J�n�ʒu
    REVNUMCA As Integer
    FACTORYCA As String * 1
    OPECA As String * 1
    KCKNTCA As Integer            ' �H���A��
    SXLIDCA As String * 13
    XTALCA As String * 12
    NEKKNTCA As String * 5       ' �ŏI�ʉߊǗ��H��
    NEWKNTCA As String * 5       ' �ŏI�ʉߍH��
    NEWKKBCA As String * 2       ' �ŏI�ʉߍ�Ƌ敪
    NEMACOCA As Integer          ' �ŏI�ʉߏ�����
    GNKKNTCA As String * 5       ' ���݊Ǘ��H��
    GNWKNTCA As String * 5       ' ���ݍH��
    GNWKKBCA As String * 2       ' ���ݍ�Ƌ敪
    GNMACOCA As Integer          ' ���ݏ�����
    GNDAYCA As Date              ' ���ݏ������t
    GNLCA As Integer             ' ���ݒ���
    GNWCA As Long                ' ���ݏd��
    GNMCA As Integer             ' ���ݖ���
    SUMITLCA As Integer          ' SUMMIT����
    SUMITWCA As Long             ' SUMMIT�d��
    SUMITMCA As Integer          ' SUMMIT����
    CHGCA As Long                ' ����ޗ�
    KAKOUBCA As String * 1       ' ���H�敪
    KEIDAYCA As Date             ' �v����t
    GNTKUBCA As String * 3       ' �I�敪
    GNTNOCA As String * 4        ' �I�ԍ�
    XTWORKCA As String * 2       ' �����H��
    WFWORKCA As String * 2       ' ���ʐ���
    LSTATBCA As String * 1       ' �ŏI��ԋ敪
    RSTATBCA As String * 1       ' ������ԋ敪
    LUFRCCA As String * 3        ' �i�㺰��
    LUFRBCA As String * 1        ' �i��敪
    LDFRCCA As String * 3        ' �i������
    LDFRBCA As String * 1        ' �i���敪
    HOLDCCA As String * 3        ' ΰ��޺���
    HOLDBCA As String * 1        ' �z�[���h�敪
    EXKUBCA As String * 1        ' ��O�敪
    HENPKCA As String * 1        ' �ԕi�敪
    LIVKCA As String * 1         ' �����敪
    KANKCA As String * 1         ' �����敪
    NFCA As String * 1           ' ���ɋ敪
    SAKJCA As String * 1         ' �폜�敪
    TDAYCA As Date               ' �o�^���t
    KDAYCA As Date               ' �X�V���t
    SUMITBCA As String * 1       ' SUMMIT���M�t���O
    SNDKCA As String * 1         ' ���M�t���O
    SNDDAYCA As Date             ' ���M���t
    '2003.06.11 (SPK)Y.Katabami tuika
    CUTCNTCA As String * 1       ' �V�K�^�Đ؋敪 '1':�Đ�
    HINBFLGCA As String * 1      ' ��\�i�ԃt���O '1'�F��\�i��
    WFHOLDFLGCA As String * 1    ' �z�[���h�敪(WF) 09/02/13 ooba
    HOLDKTCA As String * 5
    RPCRYNUMCA As String * 12    ' �e��ۯ�ID�@05/09/20 ooba
    KBLKFLGCA As String * 1      ' �֘A��ۯ��׸ށ@06/10/31 ooba
    BLKPOST As Integer            ' �u���b�N���ʒu(XSDCA�⏕����) 07/07/25 shindo
    BLKPOSB As Integer            ' �u���b�N���ʒu(XSDCA�⏕����) 07/07/25 shindo
    PLANTCATCA As String         ' ����@2007/08/15 SPK Tsutsumi
End Type

'�X�V�p
Public Type typ_XSDCA_Update
    CRYNUMCA As String      ' ��ۯ�ID������ԍ�
    HINBCA As String
    INPOSCA As String         ' �������J�n�ʒu
    REVNUMCA As String
    FACTORYCA As String
    OPECA As String
    KCKNTCA As String            ' �H���A��
    SXLIDCA As String
    XTALCA As String
    NEKKNTCA As String       ' �ŏI�ʉߊǗ��H��
    NEWKNTCA As String      ' �ŏI�ʉߍH��
    NEWKKBCA As String       ' �ŏI�ʉߍ�Ƌ敪
    NEMACOCA As String          ' �ŏI�ʉߏ�����
    GNKKNTCA As String      ' ���݊Ǘ��H��
    GNWKNTCA As String        ' ���ݍH��
    GNWKKBCA As String        ' ���ݍ�Ƌ敪
    GNMACOCA As String        ' ���ݏ�����
    GNDAYCA As String              ' ���ݏ������t
    GNLCA As String             ' ���ݒ���
    GNWCA As String                ' ���ݏd��
    GNMCA As String             ' ���ݖ���
    SUMITLCA As String          ' SUMMIT����
    SUMITWCA As String             ' SUMMIT�d��
    SUMITMCA As String          ' SUMMIT����
    CHGCA As String              ' ����ޗ�
    KAKOUBCA As String        ' ���H�敪
    KEIDAYCA As String             ' �v����t
    GNTKUBCA As String        ' �I�敪
    GNTNOCA As String        ' �I�ԍ�
    XTWORKCA As String        ' �����H��
    WFWORKCA As String        ' ���ʐ���
    LSTATBCA As String       ' �ŏI��ԋ敪
    RSTATBCA As String        ' ������ԋ敪
    LUFRCCA As String         ' �i�㺰��
    LUFRBCA As String         ' �i��敪
    LDFRCCA As String         ' �i������
    LDFRBCA As String         ' �i���敪
    HOLDCCA As String        ' ΰ��޺���
    HOLDBCA As String         ' �z�[���h�敪
    EXKUBCA As String         ' ��O�敪
    HENPKCA As String         ' �ԕi�敪
    LIVKCA As String          ' �����敪
    KANKCA As String          ' �����敪
    NFCA As String            ' ���ɋ敪
    SAKJCA As String          ' �폜�敪
    TDAYCA As String               ' �o�^���t
    KDAYCA As String               ' �X�V���t
    SUMITBCA As String        ' SUMMIT���M�t���O
    SNDKCA As String         ' ���M�t���O
    SNDDAYCA As String             ' ���M���t
    '2003.06.11 (SPK)Y.Katabami tuika
    CUTCNTCA As String * 1       ' �V�K�^�Đ؋敪 '1':�Đ�
    HINBFLGCA As String * 1      ' ��\�i�ԃt���O '1'�F��\�i��
    HOLDKTCA As String * 5
    RPCRYNUMCA As String * 12    ' �e��ۯ�ID�@05/09/20 ooba
    KBLKFLGCA As String * 1      ' �֘A��ۯ��׸ށ@06/10/31 ooba
    PLANTCATCA As String         ' ����@2007/08/15 SPK Tsutsumi
End Type

'��SELECT��

'�T�v      :�e�[�u���uXSDCA�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO   ,�^               ,����
'          :records()     ,O    ,typ_XSDCA     ,���o���R�[�h
'          :sqlWhere      ,I    ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I    ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :

Public Function DBDRV_GetXSDCA(records() As typ_XSDCA, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN

    Dim sql As String       'SQL�S��
    Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select * From XSDCA"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDCA = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    If recCnt = 0 Then
        Exit Function
    End If
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
            If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
            If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
            If IsNull(rs.Fields("REVNUMCA")) = False Then .REVNUMCA = rs.Fields("REVNUMCA")
            If IsNull(rs.Fields("FACTORYCA")) = False Then .FACTORYCA = rs.Fields("FACTORYCA")
            If IsNull(rs.Fields("OPECA")) = False Then .OPECA = rs.Fields("OPECA")
            If IsNull(rs.Fields("KCKNTCA")) = False Then .KCKNTCA = rs.Fields("KCKNTCA")
            If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLIDCA = rs.Fields("SXLIDCA")
            If IsNull(rs.Fields("XTALCA")) = False Then .XTALCA = rs.Fields("XTALCA")
            If IsNull(rs.Fields("NEKKNTCA")) = False Then .NEKKNTCA = rs.Fields("NEKKNTCA")
            If IsNull(rs.Fields("NEWKNTCA")) = False Then .NEWKNTCA = rs.Fields("NEWKNTCA")
            If IsNull(rs.Fields("NEWKKBCA")) = False Then .NEWKKBCA = rs.Fields("NEWKKBCA")
            If IsNull(rs.Fields("NEMACOCA")) = False Then .NEMACOCA = rs.Fields("NEMACOCA")
            If IsNull(rs.Fields("GNKKNTCA")) = False Then .GNKKNTCA = rs.Fields("GNKKNTCA")
            If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
            If IsNull(rs.Fields("GNWKKBCA")) = False Then .GNWKKBCA = rs.Fields("GNWKKBCA")
            If IsNull(rs.Fields("GNMACOCA")) = False Then .GNMACOCA = rs.Fields("GNMACOCA")
            If IsNull(rs.Fields("GNDAYCA")) = False Then .GNDAYCA = rs.Fields("GNDAYCA")
            If IsNull(rs.Fields("GNLCA")) = False Then .GNLCA = rs.Fields("GNLCA")
            If IsNull(rs.Fields("GNWCA")) = False Then .GNWCA = rs.Fields("GNWCA")
            If IsNull(rs.Fields("GNMCA")) = False Then .GNMCA = rs.Fields("GNMCA")
            If IsNull(rs.Fields("SUMITLCA")) = False Then .SUMITLCA = rs.Fields("SUMITLCA")
            If IsNull(rs.Fields("SUMITWCA")) = False Then .SUMITWCA = rs.Fields("SUMITWCA")
            If IsNull(rs.Fields("SUMITMCA")) = False Then .SUMITMCA = rs.Fields("SUMITMCA")
            If IsNull(rs.Fields("CHGCA")) = False Then .CHGCA = rs.Fields("CHGCA")
            If IsNull(rs.Fields("KAKOUBCA")) = False Then .KAKOUBCA = rs.Fields("KAKOUBCA")
            If IsNull(rs.Fields("KEIDAYCA")) = False Then .KEIDAYCA = rs.Fields("KEIDAYCA")
            If IsNull(rs.Fields("GNTKUBCA")) = False Then .GNTKUBCA = rs.Fields("GNTKUBCA")
            If IsNull(rs.Fields("GNTNOCA")) = False Then .GNTNOCA = rs.Fields("GNTNOCA")
            If IsNull(rs.Fields("XTWORKCA")) = False Then .XTWORKCA = rs.Fields("XTWORKCA")
            If IsNull(rs.Fields("WFWORKCA")) = False Then .WFWORKCA = rs.Fields("WFWORKCA")
            If IsNull(rs.Fields("LSTATBCA")) = False Then .LSTATBCA = rs.Fields("LSTATBCA")
            If IsNull(rs.Fields("RSTATBCA")) = False Then .RSTATBCA = rs.Fields("RSTATBCA")
            If IsNull(rs.Fields("LUFRCCA")) = False Then .LUFRCCA = rs.Fields("LUFRCCA")
            If IsNull(rs.Fields("LUFRBCA")) = False Then .LUFRBCA = rs.Fields("LUFRBCA")
            If IsNull(rs.Fields("LDFRCCA")) = False Then .LDFRCCA = rs.Fields("LDFRCCA")
            If IsNull(rs.Fields("LDFRBCA")) = False Then .LDFRBCA = rs.Fields("LDFRBCA")
            If IsNull(rs.Fields("HOLDCCA")) = False Then .HOLDCCA = rs.Fields("HOLDCCA")
            If IsNull(rs.Fields("HOLDBCA")) = False Then .HOLDBCA = rs.Fields("HOLDBCA")
            If IsNull(rs.Fields("EXKUBCA")) = False Then .EXKUBCA = rs.Fields("EXKUBCA")
            If IsNull(rs.Fields("HENPKCA")) = False Then .HENPKCA = rs.Fields("HENPKCA")
            If IsNull(rs.Fields("LIVKCA")) = False Then .LIVKCA = rs.Fields("LIVKCA")
            If IsNull(rs.Fields("KANKCA")) = False Then .KANKCA = rs.Fields("KANKCA")
            If IsNull(rs.Fields("NFCA")) = False Then .NFCA = rs.Fields("NFCA")
            If IsNull(rs.Fields("SAKJCA")) = False Then .SAKJCA = rs.Fields("SAKJCA")
            If IsNull(rs.Fields("TDAYCA")) = False Then .TDAYCA = rs.Fields("TDAYCA")
            If IsNull(rs.Fields("KDAYCA")) = False Then .KDAYCA = rs.Fields("KDAYCA")
            If IsNull(rs.Fields("SUMITBCA")) = False Then .SUMITBCA = rs.Fields("SUMITBCA")
            If IsNull(rs.Fields("SNDKCA")) = False Then .SNDKCA = rs.Fields("SNDKCA")
            If IsNull(rs.Fields("SNDDAYCA")) = False Then .SNDDAYCA = rs.Fields("SNDDAYCA")
            '2003.06.11 (SPK)Y.katabami tuika
            If IsNull(rs.Fields("CUTCNTCA")) = False Then .CUTCNTCA = rs.Fields("CUTCNTCA")
            If IsNull(rs.Fields("HINBFLGCA")) = False Then .HINBFLGCA = rs.Fields("HINBFLGCA")
            '2005/07
            If IsNull(rs.Fields("HOLDKTCA")) = False Then .HOLDKTCA = rs.Fields("HOLDKTCA")
            If IsNull(rs.Fields("RPCRYNUMCA")) = False Then .RPCRYNUMCA = rs.Fields("RPCRYNUMCA")   '05/09/20 ooba
            If IsNull(rs.Fields("KBLKFLGCA")) = False Then .KBLKFLGCA = rs.Fields("KBLKFLGCA")      '06/10/31 ooba
            If IsNull(rs.Fields("PLANTCATCA")) = False Then .PLANTCATCA = rs.Fields("PLANTCATCA")      '07/08/22 SPK Tsutsumi Add
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDCA = FUNCTION_RETURN_SUCCESS
End Function

'��UPDATE��

'���X�V���ڂ��\���̂ɃZ�b�g���Ĉ����n��

'�T�v      :�e�[�u���uXSDCA�v���X�V���� ptrn1
'���Ұ�    :�ϐ���        ,IO  ,�^               ,����
'          :records()     ,O   ,typ_XSDCA     ,�X�V���R�[�h
'          :sqlWhere      ,I   ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I   ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O   ,FUNCTION_RETURN  ,���o�̐���
'����      :

Public Function UpdateXSDCA(records As typ_XSDCA_Update, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
On Error GoTo proc_err
    gErr.Push "s_XSDCA_SQL.bas -- Function UpdateXSDCA"

    Dim sql As String       'SQL�S��
    Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim nowtime As Date
    Dim nowtime_sql As String   '�T�[�o����(SQL��)

    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾����悤�ɕύX 2003/6/4 tuku

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select * From XSDCA"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs Is Nothing Then
        UpdateXSDCA = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''�w��ӏ����X�V����
    recCnt = rs.RecordCount

    If recCnt = 0 Then
        UpdateXSDCA = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    'XSDCA��UPDATE�O�ɌĂяo��
    #If Y3_CREATE = 1 Then
        ''XODY3�쐬  add 2009/01/08 SETmiyatake
        Call CreateOrUpdateXODY3(records.CRYNUMCA, records.SXLIDCA, records.LIVKCA, sqlWhere)
    #End If

'>>>>> .Edit��SQL(UPDATE)���ɕύX�@2009/06/29 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"

    With records
        
        ''SQL��g�ݗ��Ă�
        sql = "UPDATE XSDCA SET" & vbLf
        
        ''�X�V���t
        sql = sql & " KDAYCA = " & nowtime_sql & vbLf
        
        ''�u���b�NID�E�����ԍ�
        If .CRYNUMCA <> "" And left(.CRYNUMCA, 1) <> vbNullChar Then
            sql = sql & ",CRYNUMCA = '" & .CRYNUMCA & "'" & vbLf
        End If
        
        ''�i��
        If .HINBCA <> "" And left(.HINBCA, 1) <> vbNullChar Then
            sql = sql & ",HINBCA = '" & .HINBCA & "'" & vbLf
        End If
        
        ''�������J�n�ʒu
        If .INPOSCA <> "" Then
            sql = sql & ",INPOSCA = '" & CStr(CInt(.INPOSCA)) & "'" & vbLf
        End If
        
        ''���i�ԍ������ԍ�
        If .REVNUMCA <> "" Then
            sql = sql & ",REVNUMCA = '" & CStr(CInt(.REVNUMCA)) & "'" & vbLf
        End If
        
        ''�H��
        If .FACTORYCA <> "" And left(.FACTORYCA, 1) <> vbNullChar Then
            sql = sql & ",FACTORYCA = '" & .FACTORYCA & "'" & vbLf
        End If
        
        ''���Ə���
        If .OPECA <> "" And left(.OPECA, 1) <> vbNullChar Then
            sql = sql & ",OPECA = '" & .OPECA & "'" & vbLf
        End If
        
        ''�H���A��
        If .KCKNTCA <> "" Then
            sql = sql & ",KCKNTCA = '" & CStr(CInt(.KCKNTCA)) & "'" & vbLf
        End If
        
        ''SXLID
        If .SXLIDCA <> "" And left(.SXLIDCA, 1) <> vbNullChar Then
            sql = sql & ",SXLIDCA = '" & .SXLIDCA & "'" & vbLf
        End If
        
        ''�����ԍ�
        If .XTALCA <> "" And left(.XTALCA, 1) <> vbNullChar Then
            sql = sql & ",XTALCA = '" & .XTALCA & "'" & vbLf
        End If
        
        ''�ŏI�ʉߊǗ��H��
        If .NEKKNTCA <> "" And left(.NEKKNTCA, 1) <> vbNullChar Then
            sql = sql & ",NEKKNTCA = '" & .NEKKNTCA & "'" & vbLf
        End If
        
        ''�ŏI�ʉߍH��
        If .NEWKNTCA <> "" And left(.NEWKNTCA, 1) <> vbNullChar Then
            sql = sql & ",NEWKNTCA = '" & .NEWKNTCA & "'" & vbLf
        End If
        
        ''�ŏI�ʉߍ�Ƌ敪
        If .NEWKKBCA <> "" And left(.NEWKKBCA, 1) <> vbNullChar Then
            sql = sql & ",NEWKKBCA = '" & .NEWKKBCA & "'" & vbLf
        End If
        
        ''�ŏI�ʉߏ�����
        If .NEMACOCA <> "" Then
            sql = sql & ",NEMACOCA = '" & CStr(CInt(.NEMACOCA)) & "'" & vbLf
        End If
        
        ''���݊Ǘ��H��
        If .GNKKNTCA <> "" And left(.GNKKNTCA, 1) <> vbNullChar Then
            sql = sql & ",GNKKNTCA = '" & .GNKKNTCA & "'" & vbLf
        End If
        
        ''���ݍH��
        If .GNWKNTCA <> "" And left(.GNWKNTCA, 1) <> vbNullChar Then
            sql = sql & ",GNWKNTCA = '" & .GNWKNTCA & "'" & vbLf
        End If
        
        ''���ݍ�Ƌ敪
        If .GNWKKBCA <> "" And left(.GNWKKBCA, 1) <> vbNullChar Then
            sql = sql & ",GNWKKBCA = '" & .GNWKKBCA & "'" & vbLf
        End If
        
        ''���ݏ�����
        If .GNMACOCA <> "" Then
            sql = sql & ",GNMACOCA = '" & CStr(CInt(.GNMACOCA)) & "'" & vbLf
        End If
        
        ''���ݏ������t
        If .GNDAYCA <> "" Then
            sql = sql & ",GNDAYCA = TO_DATE('" & Format$(CDate(.GNDAYCA), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''���ݒ���
        If .GNLCA <> "" Then
            sql = sql & ",GNLCA = '" & CStr(CInt(.GNLCA)) & "'" & vbLf
        End If
        
        ''���ݏd��
        If .GNWCA <> "" Then
            sql = sql & ",GNWCA = '" & CStr(CLng(.GNWCA)) & "'" & vbLf
        End If
        
        ''���ݖ���
        If .GNMCA <> "" Then
            sql = sql & ",GNMCA = '" & CStr(CInt(.GNMCA)) & "'" & vbLf
        End If
        
        ''SUMIT����
        If .SUMITLCA <> "" Then
            sql = sql & ",SUMITLCA = '" & CStr(CInt(.SUMITLCA)) & "'" & vbLf
        End If
        
        ''SUMIT�d��
        If .SUMITWCA <> "" Then
            sql = sql & ",SUMITWCA = '" & CStr(CLng(.SUMITWCA)) & "'" & vbLf
        End If
        
        ''SUMIT����
        If .SUMITMCA <> "" Then
            sql = sql & ",SUMITMCA = '" & CStr(CInt(.SUMITMCA)) & "'" & vbLf
        End If
        
        ''�`���[�W��
        If .CHGCA <> "" Then
            sql = sql & ",CHGCA = '" & CStr(CLng(.CHGCA)) & "'" & vbLf
        End If
        
        ''���H�敪
        If .KAKOUBCA <> "" And left(.KAKOUBCA, 1) <> vbNullChar Then
            sql = sql & ",KAKOUBCA = '" & .KAKOUBCA & "'" & vbLf
        End If
        
        ''�v����t
        If .KEIDAYCA <> "" Then
            sql = sql & ",KEIDAYCA = TO_DATE('" & Format$(CDate(.KEIDAYCA), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''�I�敪
        If .GNTKUBCA <> "" And left(.GNTKUBCA, 1) <> vbNullChar Then
            sql = sql & ",GNTKUBCA = '" & .GNTKUBCA & "'" & vbLf
        End If
        
        ''�I�ԍ�
        If .GNTNOCA <> "" And left(.GNTNOCA, 1) <> vbNullChar Then
            sql = sql & ",GNTNOCA = '" & .GNTNOCA & "'" & vbLf
        End If
        
        ''�����H��
        If .XTWORKCA <> "" And left(.XTWORKCA, 1) <> vbNullChar Then
            sql = sql & ",XTWORKCA = '" & .XTWORKCA & "'" & vbLf
        End If
        
        ''�E�F�[�n����
        If .WFWORKCA <> "" And left(.WFWORKCA, 1) <> vbNullChar Then
            sql = sql & ",WFWORKCA = '" & .WFWORKCA & "'" & vbLf
        End If
        
        ''�ŏI��ԋ敪
        If .LSTATBCA <> "" And left(.LSTATBCA, 1) <> vbNullChar Then
            sql = sql & ",LSTATBCA = '" & .LSTATBCA & "'" & vbLf
        End If
        
        ''������ԋ敪
        If .RSTATBCA <> "" And left(.RSTATBCA, 1) <> vbNullChar Then
            sql = sql & ",RSTATBCA = '" & .RSTATBCA & "'" & vbLf
        End If
        
        ''�i��R�[�h
        If .LUFRCCA <> "" And left(.LUFRCCA, 1) <> vbNullChar Then
            sql = sql & ",LUFRCCA = '" & .LUFRCCA & "'" & vbLf
        End If
        
        ''�i��敪
        If .LUFRBCA <> "" And left(.LUFRBCA, 1) <> vbNullChar Then
            sql = sql & ",LUFRBCA = '" & .LUFRBCA & "'" & vbLf
        End If
        
        ''�i���R�[�h
        If .LDFRCCA <> "" And left(.LDFRCCA, 1) <> vbNullChar Then
            sql = sql & ",LDFRCCA = '" & .LDFRCCA & "'" & vbLf
        End If
        
        ''�i���敪
        If .LDFRBCA <> "" And left(.LDFRBCA, 1) <> vbNullChar Then
            sql = sql & ",LDFRBCA = '" & .LDFRBCA & "'" & vbLf
        End If
        
        ''�z�[���h�R�[�h
        If .HOLDCCA <> "" And left(.HOLDCCA, 1) <> vbNullChar Then
            sql = sql & ",HOLDCCA = '" & .HOLDCCA & "'" & vbLf
        End If
        
        ''�z�[���h�敪
        If .HOLDBCA <> "" And left(.HOLDBCA, 1) <> vbNullChar Then
            sql = sql & ",HOLDBCA = '" & .HOLDBCA & "'" & vbLf
        End If
        
        ''��O�敪
        If .EXKUBCA <> "" And left(.EXKUBCA, 1) <> vbNullChar Then
            sql = sql & ",EXKUBCA = '" & .EXKUBCA & "'" & vbLf
        End If
        
        ''�ԕi�敪
        If .HENPKCA <> "" And left(.HENPKCA, 1) <> vbNullChar Then
            sql = sql & ",HENPKCA = '" & .HENPKCA & "'" & vbLf
        End If
        
        ''�����敪
        If .LIVKCA <> "" And left(.LIVKCA, 1) <> vbNullChar Then
            sql = sql & ",LIVKCA = '" & .LIVKCA & "'" & vbLf
        End If
        
        ''�����敪
        If .KANKCA <> "" And left(.KANKCA, 1) <> vbNullChar Then
            sql = sql & ",KANKCA = '" & .KANKCA & "'" & vbLf
        End If
        
        ''���ɋ敪
        If .NFCA <> "" And left(.NFCA, 1) <> vbNullChar Then
            sql = sql & ",NFCA = '" & .NFCA & "'" & vbLf
        End If
        
        ''�폜�敪
        If .SAKJCA <> "" And left(.SAKJCA, 1) <> vbNullChar Then
            sql = sql & ",SAKJCA = '" & .SAKJCA & "'" & vbLf
        End If
        
        ''�o�^���t
        If .TDAYCA <> "" Then
            sql = sql & ",TDAYCA = TO_DATE('" & Format$(CDate(.TDAYCA), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''SUMIT���M�t���O
        If .SUMITBCA <> "" And left(.SUMITBCA, 1) <> vbNullChar Then
            sql = sql & ",SUMITBCA = '" & .SUMITBCA & "'" & vbLf
        End If
        
        ''���M�t���O
        If .SNDKCA <> "" And left(.SNDKCA, 1) <> vbNullChar Then
            sql = sql & ",SNDKCA = '" & .SNDKCA & "'" & vbLf
        End If
        
        ''���M���t
        If .SNDDAYCA <> "" Then
            sql = sql & ",SNDDAYCA = TO_DATE('" & Format$(CDate(.SNDDAYCA), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''�ؒf�����敪
        If .CUTCNTCA <> "" And left(.CUTCNTCA, 1) <> vbNullChar Then
            sql = sql & ",CUTCNTCA = '" & .CUTCNTCA & "'" & vbLf
        End If
        
        ''��\�i�ԃt���O
        If .HINBFLGCA <> "" And left(.HINBFLGCA, 1) <> vbNullChar Then
            sql = sql & ",HINBFLGCA = '" & .HINBFLGCA & "'" & vbLf
        End If
        
        ''�z�[���h�H��
        If .HOLDKTCA <> "" And left(.HOLDKTCA, 1) <> vbNullChar Then
            sql = sql & ",HOLDKTCA = '" & .HOLDKTCA & "'" & vbLf
        End If
        
        ''�e�u���b�NID
        If .RPCRYNUMCA <> "" And left(.RPCRYNUMCA, 1) <> vbNullChar Then
            sql = sql & ",RPCRYNUMCA = '" & .RPCRYNUMCA & "'" & vbLf
        End If
        
        ''�֘A�u���b�N�t���O
        If .KBLKFLGCA <> "" And left(.KBLKFLGCA, 1) <> vbNullChar Then
            sql = sql & ",KBLKFLGCA = '" & .KBLKFLGCA & "'" & vbLf
        End If
        
        ''����
        If .PLANTCATCA <> "" And left(.PLANTCATCA, 1) <> vbNullChar Then
            sql = sql & ",PLANTCATCA = '" & .PLANTCATCA & "'" & vbLf
        End If
        
        sql = sql & " " & sqlWhere & vbLf
    
        'SQL�����s
        recCnt = OraDB.ExecuteSQL(sql)
        
        '�Ԃ�l��1�ȊO�̓G���[
        If recCnt < 0 Then
            GoTo proc_err
        ElseIf recCnt = 0 Then
            '0���X�V�c�G���[(�����ʂ�)
            UpdateXSDCA = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf recCnt > 1 Then
            '�������X�V�c�G���[(�����͕���SELECT�����ŏ��̈ꌏ�̂ݍX�V)
            UpdateXSDCA = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        
    End With
'<<<<< .Edit��SQL(UPDATE)���ɕύX�@2009/06/29 SETsw kubota ------------------

    UpdateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    UpdateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'��INSERT��

'�T�v      :�e�[�u���uXSDCA�v�Ƀ��R�[�h��}������
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pXSDCA �@�@  ,I  ,typ_XSDCA_Update   ,XSDCA�X�V�p�ް�
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function CreateXSDCA(pXSDCA As typ_XSDCA_Update, sErrMsg As String) As FUNCTION_RETURN


    Dim sql As String
    Dim sDBName As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim rs2 As OraDynaset    'RecordSet
'    Dim recCnt As Long      '���R�[�h��
    Dim nowtime As Date
    Dim nowtime_sql As String   '�T�[�o����(SQL��)

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDCA_SQL.bas -- Function CreateXSDCA"
    sErrMsg = ""
    sDBName = "XSDCA"
    'nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾����悤�ɕύX 2003/6/4 tuku

'>>>>> .AddNew��SQL(INSERT)���ɕύX�@2009/06/29 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With pXSDCA
        
        sql = "INSERT INTO XSDCA ("
        sql = sql & " CRYNUMCA"         ' 1:��ۯ�ID�E�����ԍ�
        sql = sql & ",HINBCA"           ' 2:�i��
        sql = sql & ",INPOSCA"          ' 3:�������J�n�ʒu
        sql = sql & ",REVNUMCA"         ' 4:���i�ԍ������ԍ�
        sql = sql & ",FACTORYCA"        ' 5:�H��
        sql = sql & ",OPECA"            ' 6:���Ə���
        sql = sql & ",KCKNTCA"          ' 7:�H���A��
        sql = sql & ",SXLIDCA"          ' 8:SXLID
        sql = sql & ",XTALCA"           ' 9:�����ԍ�
        sql = sql & ",NEKKNTCA"         '10:�ŏI�ʉߊǗ��H��
        sql = sql & ",NEWKNTCA"         '11:�ŏI�ʉߍH��
        sql = sql & ",NEWKKBCA"         '12:�ŏI�ʉߍ�Ƌ敪
        sql = sql & ",NEMACOCA"         '13:�ŏI�ʉߏ�����
        sql = sql & ",GNKKNTCA"         '14:���݊Ǘ��H��
        sql = sql & ",GNWKNTCA"         '15:���ݍH��
        sql = sql & ",GNWKKBCA"         '16:���ݍ�Ƌ敪
        sql = sql & ",GNMACOCA"         '17:���ݏ�����
        sql = sql & ",GNDAYCA"          '18:���ݏ������t
        sql = sql & ",GNLCA"            '19:���ݒ���
        sql = sql & ",GNWCA"            '20:���ݏd��
        sql = sql & ",GNMCA"            '21:���ݖ���
        sql = sql & ",SUMITLCA"         '22:SUMMIT����
        sql = sql & ",SUMITWCA"         '23:SUMMIT�d��
        sql = sql & ",SUMITMCA"         '24:SUMMIT����
        sql = sql & ",CHGCA"            '25:����ޗ�
        sql = sql & ",KAKOUBCA"         '26:���H�敪
        If .KEIDAYCA <> "" Then
            sql = sql & ",KEIDAYCA"         '27:�v����t
        End If
        sql = sql & ",GNTKUBCA"         '28:�I�敪
        sql = sql & ",GNTNOCA"          '29:�I�ԍ�
        sql = sql & ",XTWORKCA"         '30:�����H��
        sql = sql & ",WFWORKCA"         '31:���ʐ���
        sql = sql & ",LSTATBCA"         '32:�ŏI��ԋ敪
        sql = sql & ",RSTATBCA"         '33:������ԋ敪
        sql = sql & ",LUFRCCA"          '34:�i�㺰��
        sql = sql & ",LUFRBCA"          '35:�i��敪
        sql = sql & ",LDFRCCA"          '36:�i������
        sql = sql & ",LDFRBCA"          '37:�i���敪
        sql = sql & ",HOLDCCA"          '38:ΰ��޺���
        sql = sql & ",HOLDBCA"          '39:ΰ��ދ敪
        sql = sql & ",EXKUBCA"          '40:��O�敪
        sql = sql & ",HENPKCA"          '41:�ԕi�敪
        sql = sql & ",LIVKCA"           '42:�����敪
        sql = sql & ",KANKCA"           '43:�����敪
        sql = sql & ",NFCA"             '44:���ɋ敪
        sql = sql & ",SAKJCA"           '45:�폜�敪
        sql = sql & ",TDAYCA"           '46:�o�^���t
        sql = sql & ",KDAYCA"           '47:�X�V���t
        sql = sql & ",SUMITBCA"         '48:SUMMIT���M�׸�
        sql = sql & ",SNDKCA"           '49:���M�׸�
        sql = sql & ",SNDDAYCA"         '50:���M���t
        sql = sql & ",CUTCNTCA"         '51:�V�K�^�Đ؋敪
        sql = sql & ",HINBFLGCA"        '52:��\�i�ԃt���O
        sql = sql & ",HOLDKTCA"         '53:ΰ��ލH��
        sql = sql & ",RPCRYNUMCA"       '54:�e��ۯ�ID
        sql = sql & ",KBLKFLGCA"        '55:�֘A��ۯ��׸�
        sql = sql & ",PLANTCATCA"       '56:����
        sql = sql & ")"
        sql = sql & "VALUES (" & vbLf

        ' 1:��ۯ�ID�E�����ԍ�
        If .CRYNUMCA <> "" And left(.CRYNUMCA, 1) <> vbNullChar Then
            sql = sql & " '" & .CRYNUMCA & "'" & vbLf
        Else
            sql = sql & " '" & Space(12) & "'" & vbLf
        End If

        ' 2:�i��
        If .HINBCA <> "" And left(.HINBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HINBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        ' 3:�������J�n�ʒu
        If .INPOSCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.INPOSCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 4:���i�ԍ������ԍ�
        If .REVNUMCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.REVNUMCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 5:�H��
        If .FACTORYCA <> "" And left(.FACTORYCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .FACTORYCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 6:���Ə���
        If .OPECA <> "" And left(.OPECA, 1) <> vbNullChar Then
            sql = sql & ",'" & .OPECA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 7:�H���A��
        If .KCKNTCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.KCKNTCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 8:SXLID
        If .SXLIDCA <> "" And left(.SXLIDCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .SXLIDCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(13) & "'" & vbLf
        End If

        ' 9:�����ԍ�
        If .XTALCA <> "" And left(.XTALCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .XTALCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '10:�ŏI�ʉߊǗ��H��
        If .NEKKNTCA <> "" And left(.NEKKNTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEKKNTCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '11:�ŏI�ʉߍH��
        If .NEWKNTCA <> "" And left(.NEWKNTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEWKNTCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '12:�ŏI�ʉߍ�Ƌ敪
        If .NEWKKBCA <> "" And left(.NEWKKBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEWKKBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '13:�ŏI�ʉߏ�����
        If .NEMACOCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.NEMACOCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '14:���݊Ǘ��H��
        If .GNKKNTCA <> "" And left(.GNKKNTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNKKNTCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '15:���ݍH��
        If .GNWKNTCA <> "" And left(.GNWKNTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNWKNTCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '16:���ݍ�Ƌ敪
        If .GNWKKBCA <> "" And left(.GNWKKBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNWKKBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '17:���ݏ�����
        If .GNMACOCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNMACOCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '18:���ݏ������t
        sql = sql & "," & nowtime_sql & vbLf

        '19:���ݒ���
        If .GNLCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNLCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '20:���ݏd��
        If .GNWCA <> "" Then
            sql = sql & ",'" & CStr(CLng(.GNWCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '21:���ݖ���
        If .GNMCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNMCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '22:SUMMIT����
        If .SUMITLCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.SUMITLCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '23:SUMMIT�d��
        If .SUMITWCA <> "" Then
            sql = sql & ",'" & CStr(CLng(.SUMITWCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '24:SUMMIT����
        If .SUMITMCA <> "" Then
            sql = sql & ",'" & CStr(CInt(.SUMITMCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '25:����ޗ�
        If .CHGCA <> "" Then
            sql = sql & ",'" & CStr(CLng(.CHGCA)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '26:���H�敪
        If .KAKOUBCA <> "" And left(.KAKOUBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .KAKOUBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '27:�v����t
        If .KEIDAYCA <> "" Then
            sql = sql & ",TO_DATE('" & Format$(CDate(.KEIDAYCA), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If

        '28:�I�敪
        If .GNTKUBCA <> "" And left(.GNTKUBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNTKUBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '29:�I�ԍ�
        If .GNTNOCA <> "" And left(.GNTNOCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNTNOCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(4) & "'" & vbLf
        End If

        '30:�����H��
        sql = sql & ",'" & FACTORYCD & "'" & vbLf

        '31:���ʐ���
        If .WFWORKCA <> "" And left(.WFWORKCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .WFWORKCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '32:�ŏI��ԋ敪
        If .LSTATBCA <> "" And left(.LSTATBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LSTATBCA & "'" & vbLf
        Else
            sql = sql & ",'T'" & vbLf
        End If

        '33:������ԋ敪
        If .RSTATBCA <> "" And left(.RSTATBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .RSTATBCA & "'" & vbLf
        Else
            sql = sql & ",'T'" & vbLf
        End If

        '34:�i�㺰��
        If .LUFRCCA <> "" And left(.LUFRCCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LUFRCCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '35:�i��敪
        If .LUFRBCA <> "" And left(.LUFRBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LUFRBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '36:�i������
        If .LDFRCCA <> "" And left(.LDFRCCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LDFRCCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '37:�i���敪
        If .LDFRBCA <> "" And left(.LDFRBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LDFRBCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '38:ΰ��޺���
        If .HOLDCCA <> "" And left(.HOLDCCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDCCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '39:ΰ��ދ敪
        If .HOLDBCA <> "" And left(.HOLDBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDBCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '40:��O�敪
        If .EXKUBCA <> "" And left(.EXKUBCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .EXKUBCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '41:�ԕi�敪
        If .HENPKCA <> "" And left(.HENPKCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HENPKCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '42:�����敪
        If .LIVKCA <> "" And left(.LIVKCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .LIVKCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '43:�����敪
        If .KANKCA <> "" And left(.KANKCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .KANKCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '44:���ɋ敪
        If .NFCA <> "" And left(.NFCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .NFCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '45:�폜�敪
        If .SAKJCA <> "" And left(.SAKJCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .SAKJCA & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '46:�o�^���t
        sql = sql & "," & nowtime_sql & vbLf

        '47:�X�V���t
        sql = sql & "," & nowtime_sql & vbLf

        '48:SUMMIT���M�׸�
        sql = sql & ",'0'" & vbLf

        '49:���M�׸�
        sql = sql & ",'0'" & vbLf

        '50:���M���t
        sql = sql & ",NULL" & vbLf

        '51:�V�K�^�Đ؋敪
        If .CUTCNTCA <> "" And left(.CUTCNTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .CUTCNTCA & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '52:��\�i�ԃt���O
        If .HINBFLGCA <> "" And left(.HINBFLGCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HINBFLGCA & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '53:ΰ��ލH��
        If .HOLDKTCA <> "" And left(.HOLDKTCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDKTCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '54:�e��ۯ�ID
        If .RPCRYNUMCA <> "" And left(.RPCRYNUMCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .RPCRYNUMCA & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '55:�֘A��ۯ��׸�
        If .KBLKFLGCA <> "" And left(.KBLKFLGCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .KBLKFLGCA & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '56:����
        If .PLANTCATCA <> "" And left(.PLANTCATCA, 1) <> vbNullChar Then
            sql = sql & ",'" & .PLANTCATCA & "'" & vbLf
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
    
    #If Y3_CREATE = 1 Then
        ''XODY3�쐬  add 2009/01/08 SETmiyatake
        Call CreateOrUpdateXODY3(pXSDCA.CRYNUMCA, pXSDCA.SXLIDCA, pXSDCA.LIVKCA)  'upd SETkimizuka
    #End If

    CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", "DB", sDBName)
    CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�Y������ں��ޗL��������
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_BlockID    ,I  ,String           ,SXLID
'      �@�@:p_Hinban     ,O  ,String           ,����
'      �@�@:p_Inpos      ,O  ,Integer          ,����
'      �@�@:�߂�l       ,O  ,Boolean        �@,ں��ނȂ�(TRUE)/����(FALSE)
'�����@�@�@�F�i�ԐU�ցA�ؽ�يi��Ȃǎ�ں��ނƓ��i�Ԃւ̕ύX�ɑΉ�
'�����@�@�@�F2002/08/29 ohno
Public Function CheckUniqueRecord(p_BlockID As String, p_Hinban As String, p_Inpos As Integer) As Boolean
    Dim sql As String
    Dim rs As OraDynaset

    sql = "SELECT * FROM XSDCA WHERE CRYNUMCA = '" & p_BlockID
    sql = sql & "' AND HINBCA = '" & p_Hinban
    sql = sql & "' AND INPOSCA = " & p_Inpos

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        CheckUniqueRecord = True
    Else
        CheckUniqueRecord = False
    End If

End Function


'�T�v      :�ŏI�ʉߏ����񐔂��擾����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_sCrynum    ,I  ,String           ,�u���b�NID
'      �@�@:p_iInpos     ,I  ,Integer          ,�J�n�ʒu
'      �@�@:�߂�l       ,O  ,Integer        �@,������
Public Function GetNEMACOC(p_sCrynum As String, p_iInpos As Integer) As Integer
    Dim sql As String
    Dim rs As OraDynaset


    sql = "SELECT GNMACOCA FROM XSDCA WHERE CRYNUMCA = '" & p_sCrynum
    sql = sql & "' AND INPOSCA = " & p_iInpos
'    sql = sql & " AND KCKNTCA = (SELECT MAX(KCKNTCA) FROM XSDCA WHERE CRYNUMCA = '" & p_sCrynum
'    sql = sql & "' AND INPOSCA = " & p_iInpos & ")"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        GetNEMACOC = 1
    Else
        GetNEMACOC = CInt(rs.Fields("GNMACOCA"))
    End If

End Function


'�T�v      :�e�[�u���uXSDCA�v�̍H�����`�F�b�N����
'���Ұ�    :�ϐ���        ,IO   ,�^               ,����
'          :tXSDCA()      ,I    ,typ_XSDCA        ,�`�F�b�N���R�[�h
'          :sNowCode      ,I    ,String           ,�`�F�b�N�H��
'          :sErrMsg       ,O    ,String           ,�G���[���b�Z�[�W
'          :�߂�l        ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :�w��f�[�^�̍H���������œn���ꂽ�H���Ɠ������`�F�b�N����
'          :�����u���b�N�̃`�F�b�N�͑Ή����Ă��Ȃ��B
'           2006/03/10 �V�K�쐬�@�d�|�H���ă`�F�b�N�@�\�ǉ�

Public Function DBDRV_CheckCodeXSDCA(tXSDCA() As typ_XSDCA, sNowCode As String, sErrMsg As String) As FUNCTION_RETURN

    Dim lsSql As String             'SQL�S��
    Dim rs As OraDynaset            'RecordSet
    Dim tReadXSDCA() As typ_XSDCA   '�擾�f�[�^
    Dim llLoopCnt   As Long
    Dim llBlockSt  As Long          '�u���b�N�̔z����J�n�ʒu
    Dim llBlockEd  As Long          '�u���b�N�̔z����I���ʒu
    Dim i As Long
    Dim j As Long
    Dim sDBName As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDCA_SQL.bas -- Function DBDRV_CheckCodeXSDCA"
    sErrMsg = ""
    sDBName = "XSDCA"
    llBlockSt = 1
    llBlockEd = UBound(tXSDCA)

'    For llLoopCnt = 1 To UBound(tXSDCA)
'        '�u���b�N�̔z����J�n�ʒu��ێ�
'        llBlockSt = llLoopCnt
'        '�u���b�N�̔z����I���ʒu��ێ�
'        llBlockEd = llLoopCnt
'        For i = llLoopCnt + 1 To UBound(tXSDCA)
'            If tXSDCA(i).CRYNUMCA <> tXSDCA(i - 1).CRYNUMCA Then
'                llBlockEd = i - 1
'                llLoopCnt = llBlockEd
'                Exit For
'            End If
'        Next i

    i = 0
    ReDim tReadXSDCA(0) As typ_XSDCA
    For llLoopCnt = 1 To llBlockEd
        If llLoopCnt = 1 Or _
           (llLoopCnt > 1 And tXSDCA(llLoopCnt).CRYNUMCA <> tXSDCA(llLoopCnt - 1).CRYNUMCA) Then
            ''SQL��g�ݗ��Ă�
            lsSql = ""
            lsSql = lsSql & " SELECT"
            lsSql = lsSql & "   CRYNUMCA"
            lsSql = lsSql & "  ,HINBCA"
            lsSql = lsSql & "  ,INPOSCA"
            lsSql = lsSql & "  ,GNWKNTCA"
            lsSql = lsSql & " FROM"
            lsSql = lsSql & "   XSDCA"
            lsSql = lsSql & " WHERE CRYNUMCA = '" & tXSDCA(llLoopCnt).CRYNUMCA & "'"
            lsSql = lsSql & "   AND LIVKCA = '0' "


            ''�f�[�^�𒊏o����
            Set rs = OraDB.DBCreateDynaset(lsSql, ORADYN_DEFAULT)
            If rs Is Nothing Then
'                ReDim records(0)
                DBDRV_CheckCodeXSDCA = FUNCTION_RETURN_FAILURE
                Exit Function
            End If

            ''���o���ʂ��i�[����
'            i = 0
'            ReDim tReadXSDCA(0) As typ_XSDCA
            Do Until rs.EOF '�f�[�^���Ȃ��Ȃ�܂Ŏ擾
                i = i + 1
                ReDim Preserve tReadXSDCA(i) As typ_XSDCA
                With tReadXSDCA(i)
                    If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
                    If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
                    If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
                    If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
                End With
                rs.MoveNext
            Loop
            rs.Close
        End If
    Next llLoopCnt

        '�����u���b�N�͈̔͂Ń��[�v����
        For i = llBlockSt To llBlockEd
            For j = 1 To UBound(tReadXSDCA)
                '�u���b�N�A�i�ԁA�������J�n�ʒu����������T��
                If Trim(tXSDCA(i).CRYNUMCA) = Trim(tReadXSDCA(j).CRYNUMCA) And _
                   Trim(tXSDCA(i).HINBCA) = Trim(tReadXSDCA(j).HINBCA) And _
                   Trim(tXSDCA(i).INPOSCA) = Trim(tReadXSDCA(j).INPOSCA) Then

                    '���ݍH�����������`�F�b�N����
                    If Trim(tReadXSDCA(j).GNWKNTCA) <> Trim(sNowCode) Then
                        '���ݍH�����Ⴄ = �u���b�N�����łɓ����Ă���ꍇ�A�G���[�I��
                        sErrMsg = GetMsgStr("EBLK6")
                        DBDRV_CheckCodeXSDCA = FUNCTION_RETURN_FAILURE
                        Exit Function
                    Else
                        '�����ꍇ�A���̕i�Ԃ�
                        Exit For
                    End If

                End If
            Next j
        Next i

'    Next llLoopCnt

    DBDRV_CheckCodeXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print lsSql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", "DB", sDBName)
    DBDRV_CheckCodeXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit


End Function


